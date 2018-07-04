unit ULibSaisiePiece;

interface

uses
  Hctrls,            // THGrid
  SysUtils,          // TSearchRec
  Classes,           // TNotifyEvent
  Graphics,          // TCanvas
  Forms,             // TForm
  Dialogs,           // TFindDialog
  Controls,          // pour le mrYes / TKeyEvent
  grids,             // TGridDrawState
  Menus,             // TPopupMenu

  SHDocVw_TLB,       // TWebBrowser_V1
  hpanel,            // THpanel
  HTB97,             // TToolBarButton97 , TDock97

  UTob,
  HEnt1,             // TActionFiche
  Ent1,              // TExoDate
  ZLibAuto,           // TZLibAuto
  uLibEcriture,      // TRechercheGrille
  HFLabel,           // TFlashingLabel
  uLibSaisieDoc,    // gestionnaire des documents    
  ExternalViewer, // LanceExternalViewer
  ULibPieceCompta    // TPieceCompta
  ;

Type

TStatutSaisie = ( ssEditEntete, ssEditGrille ) ;

TSaisiePiece = Class ;

TSaisieGuide = Class
  private

    FTSPiece        : TSaisiePiece ;
    FTobGuide       : Tob ;

    FBoGuideActif   : Boolean ;
    FBoFinGuide     : Boolean ;
    FInDeb          : integer ;
    FInFin          : integer ;
    FTabGuideCol    : Array[1..7] of Integer ;

    FStLastSQL      : string ;
    FBoLastRech     : boolean ;
    FOldModeEche    : TModeEche ;

    FFlashLabel     : TFlashingLabel ;

    function    GetPiece         : TPieceCompta ;
    function    GetFListe        : THGrid ;
    function    GetRechCompte    : TRechercheGrille ;

    procedure   InitFlashGuide ( vControl : TControl ) ;
    procedure   InitTabGuide ;

  public

    property    Guide            : Tob                  read FTobGuide ;
    property    TSP              : TSaisiePiece         read FTSPiece ;

    property    GuideActif       : Boolean              read FBoGuideActif ;
    property    FinGuide         : Boolean              read FBoFinGuide ;

    property    Piece            : TPieceCompta         read GetPiece ;
    property    FListe           : THGrid               read GetFListe ;
    property    RechCompte       : TRechercheGrille     read GetRechCompte ;

    property    FlashLabel       : TFlashingLabel       read FFlashLabel     ;


    constructor Create           ( vTSP : TSaisiePiece ; vControl : TControl = nil ) ;
    destructor  Destroy          ; override ;

    procedure GuideCellExit      ( Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure GuideRowExit       ( Sender: TObject; Ou: Integer;  var Cancel: Boolean; Chg: Boolean );
    procedure GuideKeyDown       ( Sender: TObject; var Key: Word; Shift: TShiftState) ;

    Function  EstColEditable     ( vInRow : integer ; vStColName : String ) : Boolean ;
    function  GetNomCol          ( vIdxArret : integer ) : string ;
    function  GetIdxArret        ( vCol : integer ; vStColName : string = '') : integer ;
    function  GetArretPourCol    ( vRow : integer ; vStName : string ) : boolean ;

    procedure SetFlashGuide      ( vControl : TControl ) ;

    procedure ActiveGuide        ( vBoOn : boolean ) ;
    procedure ModeSilence        ( vBoOn : boolean ) ;

    // Recherche des guides en base...
    function  GetSQL           ( vRow : integer ) : string ;
    function  ExisteGuide      ( vRow : integer ) : boolean ;
    function  RechercherGuide  ( vRow : integer ) : boolean ;
    function  Load             ( const Values : array of string ) : boolean ;
    function  GetLastGuide     : string ;

    // Execution / abondon
    procedure DemarrerGuide    ( vRow : integer ) ;
    procedure FinirGuide       ( vBoAnnul : boolean = False );
    procedure AbandonnerGuide  ;
    function  GetColSuivante   ( var Acol, ARow : integer ) : boolean ;
    function  GetColPrecedente ( vcol, vRow : integer ) : boolean ;

    // Test
    function  EstGuideCorrect  : boolean ;
    function  IsValidLigne     ( vRow : integer ) : boolean ;
    procedure CompleteLigne    ( vRow : integer ) ;

    // interaction utilisateur
    function  AOuvrirAnal        ( vRow : integer  ) : boolean ;
    function  AOuvrirEche        ( vRow : integer  ) : boolean ;

    // Calculs
    function  GetGuideLigne    ( vRow : integer ) : Tob ;
    function  CalculGuide      ( vRow : integer ) : boolean ;
    function  CalculLigne      ( vRow : integer ; lTobEcrGui : Tob ) : boolean ;
    function  RecalculGuide    ( vRow : integer ) : boolean ;
    function  Estformule       ( vSt : string ) : boolean ;

    // Affichage
    procedure AfficheLignes ;

  end ;


TSaisiePiece = Class( TObject )
  private

      // Données de la grille
      FTobPiece          : TPieceCompta;      // Contenu de la Grille

      // Interface
      FEcran            : TForm;             // interface
      FListe            : THGrid;            // Grille de saisie

      // Gestion des formules
      FGrille         : String;
      FTob            : Tob;
      
      // Paramétrage de la grille
      FGridChamps       : String ;           // Liste des champs pour maj par putGriddetail
      FGridColSize      : String ;           // Largeur des colonnes de la grille
      FTabColAvecLib    : Array of Boolean ; // Colonne affichée en libellé complet
      FTobParamLib      : TOB ;              // Gestion de lapersonnalisation des champs libres
      FGrilleCompl      : THGrid ;           // Pointeur sur la grille d'informations complémentaires
      FTabColAvecLibC   : Array of Boolean ; // Colonne affichée en libellé complet

      // Gestion du masque de saisie
      FListeMasques     : Tob ;              // Liste des masques utilisables par la grille
      FTobMasque        : Tob ;              // Pointeur sur le masque de saisie sélectionnée
      FStMasqueCrit1    : string ;           // 1er Critère de sélection du masque
      FStMasqueCrit2    : string ;           // 2eme Critère de sélection du masque
      FStMasqueCrit3    : string ;           // 3eme Critère de sélection du masque

      // Evènements
      FBoEvtOn            : Boolean ;
      FUserKeyDown        : TKeyEvent;
      FUserRowEnter       : TChangeRC ;
      FUserRowExit        : TChangeRC ;
      FUserCellEnter      : TChangeCell ;
      FUserCellExit       : TChangeCell ;
      FUserDblClick       : TNotifyEvent ;
      FInheritedDblClick  : TNotifyEvent ;
      FUserGridEnter      : TNotifyEvent ;
//      FBoFirstEnter       : Boolean ;

      // gestion de la recherche dans la liste
      FFindDialog         : TFindDialog ;       // Objet de recherche
      FirstFind           : Boolean ;
      FRechCompte         : TRechercheGrille ;
      FMessageCompta      : TMessageCompta ;
      FZlibAuto           : TZLibAuto ;            // objet de gestion des libelles automatique

      // Indicateurs
      FLastRef            : string ;
      FInLigneCourante    : Integer ;             // Gestion des comptes autos
      FInIndexComptaAuto  : Integer ;             // Gestion des comptes autos

      // Accélérateur de saisie
      FBoAccActif   : Boolean ;          // Permet la désactivation de l'accélérateur de saisie
      FBoAccArret   : Boolean ;
      FStAccCptHT   : string ;

      // Gestion des guides
      FTSGuide      : TSaisieGuide ;

      // Gestion du visualiseur de document
      FTSDoc           : TSaisieDoc ;
      FTSDocInt        : TSaisieDoc ;
      FTSDocExt        : TSaisieDoc ;
      FBoDocMAJEnCours : Boolean ;
      FViewer          : String ;
      FViewerExt       : TViewerExt ;

      procedure SetGrille      (const Value: THGrid);

      function  GetCount       :  Integer ;
      function  GetRow         :  Integer ;
      procedure SetRow         ( vIndex :  Integer  ) ;
      function  GetCol         :  Integer ;
      procedure SetCol         ( vIndex :  Integer  ) ;
      function  GetEnabled     :  boolean ;
      procedure SetEnabled     ( vBoEna : boolean ) ;

      procedure InitEvtGrid ;
      procedure InitFormatGrid ;
      procedure InitParamLib ;
//      function  ComputeCell    ( vStArg : hstring ) : Variant ;

      Function  GetChampDataType( vStColName : String ) : String ;
//      Function  GetFormatCombo( vStColName : String ) : String ;
      Function  GetFullSize( vGrille : THGrid = nil ) : Integer ;
      Function  GetDossier : String ;
      Function  GetInfo : TInfoEcriture ;

      procedure DupliquerZone( vBoFromFirst : boolean = False ; vBoOnlyRef : boolean = False ) ;

  //////////////////////////////////////////////////////////////////////////////////////////
  Public

      // ======================
      // ===== PROPRIETES =====
      // ======================

      property Dossier:     String          read  GetDossier ;

      // Données
      property Piece:       TPieceCompta     read  FTobPiece ;
      property Info:        TInfoEcriture    read  GetInfo ;

      // Interface
      property Ecran:       TForm            read  FEcran         write  FEcran ;
      property LaGrille:    THGrid           read  FListe         write  SetGrille ;
      property TSGuide:     TSaisieGuide     read  FTSGuide ;
      property TSDoc    :   TSaisieDoc       read  FTSDoc ;
      property TSDocInt :   TSaisieDoc       read  FTSDocInt ;
      property TSDocExt :   TSaisieDoc       read  FTSDocExt ;
      property ViewerExt:   TViewerExt       read  FViewerExt     write FViewerExt ;
      property Viewer   :   String           read  FViewer        write FViewer ;

      // Indicateurs
      property Count:       Integer          read  GetCount ;
      property Row:         Integer          read  GetRow         write  SetRow ;
      property Col:         Integer          read  GetCol         write  SetCol ;
      property Enabled:     Boolean          read  GetEnabled     write  SetEnabled ;

      // Accélérateur de saisie
      property AccArret:    Boolean          read  FBoAccArret    write  FBoAccArret ;
      property AccActif:    Boolean          read  FBoAccActif    write  FBoAccActif ;
      property AccCptHT:    string           read  FStAccCptHT    write  FStAccCptHT ;

      // Evènements
      property OnUserKeyDown:       TKeyEvent    read  FUserKeyDown        write FUserKeyDown;
      property OnUserRowEnter:      TChangeRC    read  FUserRowEnter       write FUserRowEnter;
      property OnUserRowExit:       TChangeRC    read  FUserRowExit        write FUserRowExit;
      property OnUserCellEnter:     TChangeCell  read  FUserCellEnter      write FUserCellEnter;
      property OnUserCellExit:      TChangeCell  read  FUserCellExit       write FUserCellExit;
      property OnUserDblClick:      TNotifyEvent read  FUserDblClick       write FUserDblClick;
      property OnUserGridEnter:     TNotifyEvent read  FUserGridEnter      write FUserGridEnter;

      // Recherche
      property RechCompte:      TRechercheGrille  read FRechCompte ;
      property MessageCompta:   TMessageCompta    read FMessageCompta ;

      // ====================
      // ===== METHODES =====
      // ====================

      // Instanciation
      constructor Create( vEcran: TForm ; vGrille : THGrid ; vPieceCpt : TPieceCompta ) ;
      constructor CreerPourMasque( vTobMasque: Tob ; vGrille : THGrid ; vPieceCpt : TPieceCompta ; vEcran : TForm ) ;
      Destructor  Destroy ; override ;

      // Actions sur les lignes
      procedure InsertRow( vRow : Integer = -1 ; vBoForce : Boolean = False );
      procedure DeleteRow( vRow : Integer = -1 ; vBoForce : Boolean = False );
      procedure videLignes ;
      procedure GereNouvelleLigne( vFrom, vIn : Integer ) ;

      // Actions annexes
      procedure ProratisePiece ;
      procedure SoldePiece ;
      procedure GetCompteAuto ( vBoSuivant : Boolean ; vBoAuto : Boolean = false ) ;
      procedure ExecuteLibelleAuto( vRow : integer ) ;

      // Modification des Tob
      function  GetTOBCourante : TOB ;
      function  ValideCellule( vChamp : String ; vRow : Integer ; vBoRefresh : boolean = true ) : boolean ;
      procedure OnChangeGeneral( Sender : TObject ) ;
      procedure OnChangeAuxi( Sender : TObject ) ;

      // Paramétrage de la grille / liste
      procedure SwapModeSaisie      ( vMode : TModeSaisie ) ;
      procedure SetColCount         ( vStChamps : String ) ;
      procedure SetFormatMontant    ;
      procedure GereOptionsGrid     ;
      procedure EditEnabled         ( vBoAvec : Boolean );
      function  GetColName          ( vCol : Integer = -1 ) : String ;
      function  GetColIndex         ( vStColName : String ) : Integer ;
      Function  EstColEditable      ( vInRow : integer ; vStColName : String ) : Boolean ;
      Function  EstColLookup        ( vCol : Integer = -1 ) : Boolean ;
      Function  EstColCombo         ( vCol : Integer = -1 ) : Boolean ;
      procedure GetParamLibre       ( vStChamp: String ; var vStLibelle : String ; var vBoVisible : Boolean ) ;
      procedure ActiveCol           ( vInCol : Integer ; vBoEditable : boolean = True ) ;
      procedure GetColSuivante      ( var Acol, ARow : integer ; vInSens : integer = 0 );
      function  GetSens             ( ACol, ARow : integer ) : integer ;
      function  ValideSortieLigne   ( vBoCellule : boolean = True ) : boolean ;
      procedure ProchaineLigne      ( vRow : integer ; vStColName : string = '' ; vBoEvtOn : boolean = True ) ;
      procedure ResizeGrille        ;

      // gestion des masques de saisie
      procedure ParamMasqueSaisie   ;
      procedure UpdateMasqueSaisie  ; // Paramétrage d'une grille avec un masque de saisie
      procedure ChoisirMasque       ;
      procedure RechercheMasque     ( vBoInit : boolean = False ) ;
      function  GetTobMasque        ( vBoInit : boolean = False ) : Tob ;
      function  GetTobCol           ( vStColName : string ) : Tob ;  overload ;
      function  GetTobCol           ( vInColIdx  : integer ) : Tob ; overload ;

      // Gestion des grilles supplémentaires
      procedure UpdateMasqueCompl ;
      procedure SetGrilleComplement ( vGrille : THGrid ) ;
      procedure AfficheComplement   ( vRow : integer ) ;
      procedure GComplDrawCell      ( ACol, ARow : Longint; Canvas : TCanvas ; AState : TGridDrawState ) ;
      function  CVariantToStGrid    ( vValeur : Variant ; vGrille : THGrid ; vInCol : integer ) : string ;

      // Raffraichissement des lignes
      function  GetChampsListe      ( vCol : Integer = - 1) : String ;
      procedure AfficheGroupe       ( vRow : Integer = 0 );
      procedure AfficheLignes       ( vRow : Integer = 0 ; vBoForce : boolean = False );
      procedure CAfficheLigne       ( vRow : integer ; vTob : Tob = nil ; vBoForce : boolean = False );
      procedure CAfficheCol         ( vStNom : string ; vInCol : integer = 0 );
      procedure PutTobToGrid        ( vRow : integer ; vStChamps : string ) ;
      function  GetValueGrid        ( aCol, aRow : integer ) : Variant ;
//      procedure PutGridToTob( aCol, aRow : integer );
      procedure InitSaisie          ;
      function  GetValeurPourGrille ( vStNom: string ; vTob : Tob ; vInCol : integer = -1 ) : string ;
      function  CalculSeparateur    (ACol, ARow: Integer ; var vBoHaut: boolean ; vBo : boolean = false  ) : boolean ;
      procedure SetEvtOff           ;
      procedure SetEvtOn            ;
      procedure SetPos              ( vCol, vRow : integer ; vBoEvtOn : Boolean = True ; vBoOnEdit : Boolean = True ; vInSens : integer = 0 );
      procedure SetPosFin           ( vBoTotal : boolean = false ) ;
      procedure SetPosDebut         ( vBoTotal : boolean = false ) ;
      function  GetFirstCol         : integer ;
      function  GetFormule          ( vStFormule: HString ): Variant;
      function  GetGoodValeurGrille ( vStChamp : String ; vTobEcr : TOB ; vStGrille : string = 'SAISIE' ) : Variant;
      function  RecupChampGrille    ( vStChamp : String ; vStGrille : String ) : TOB;
      function  RecupSpecif         ( vStChamp : String ; vStGrille : String ; vStType : String ) : String;
      // Paramétrage de la fiche parente
      procedure FindDialogFind(Sender: TObject);
      procedure Recherche ;
      function  GetActionFiche : TActionFiche ;

      // navigation dans la grille
      procedure SetFocus ;
      procedure RowEnter      ( Sender: TObject; Ou: Integer;  var Cancel: Boolean; Chg: Boolean );
      procedure RowExit       ( Sender: TObject; Ou: Integer;  var Cancel: Boolean; Chg: Boolean );
      procedure CellEnter     ( Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
      procedure CellExit      ( Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
      procedure KeyDown       ( Sender: TObject; var Key: Word; Shift: TShiftState) ;
      procedure KeyPress      ( Sender: TObject; var Key: Char ) ;
      procedure PostDrawCell  ( ACol, ARow : Longint; Canvas : TCanvas ; AState : TGridDrawState ) ;
      procedure ElipsisClick  ( Sender: TObject);
      procedure CanClose      ( var CanClose : Boolean ) ;
      procedure GridEnter     ( Sender: TObject ) ;
      procedure GridExit      ( Sender: TObject ) ;
      procedure MouseDown     ( Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
      procedure DblClick       ( Sender: TObject ) ;
      procedure GetCellCanvas ( ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);

      // Gestion de la recherche
      procedure OnError (sender : TObject; Error : TRecError ) ; // fonction d'affichage des erreurs pour le FRechCompte
      procedure InitRechCompte ;
      function  RecherchePourChamps( vStColName : String ) : Boolean ;

      // Gestion des messages
      procedure InitMessageCompta ;

      // Autres...
      procedure SetInitCell       ;
      procedure SetGridSep        ( ACol, ARow : Integer ;  Canvas : TCanvas ; bHaut : Boolean ) ;
{$IFDEF SBO}
      procedure AddEvenement      ( value : string );
{$ENDIF SBO}

      // Formules
      function  GetFormulePourCalc(Formule: hstring): Variant;
      procedure IncrementerRef( vRow : integer ; vBoPlus : boolean = True ) ;

      // Gestion de la devise
      procedure SaisieDevise     ( vRow : integer ) ;
      function  EstDeviseEditable( vRow : integer ) : boolean ;

      // gestion de l'accélérateur de saisie
      function  GereAcc          ( vRow : integer ; vBoForce : boolean = False ) : integer;
      function  IsActiveAcc      : boolean ;

      // Gestion des guide
      procedure GereGuide        ( vRow : integer ) ;
      procedure LanceGuide       ( vRow : integer ) ;
      function  GuideActif       : boolean ;
      function  GetLastGuide     : string ;

      // gestion du viewer
      procedure paramViewer           ( vBoReinit : Boolean = false ) ;
      procedure ChargeDocument        ( vBoSuivant : boolean = true ) ;
      procedure InitViewer            ( vPanelB, vPanelR, vPanelL, vPanelT : THPanel ;
                                        vSplitB, vSplitR, vSplitL, vSplitT : THSplitter ) ;
      procedure InitPopupViewer       ( vPopup : TPopupMenu ; vBouton : TToolBarButton97 ; vBoutonMove : TToolBarButton97 ) ;
      procedure InitBDoc              ( vBExt : TToolBarButton97 ; vBInt : TToolBarButton97 ) ;
      function  ViewerActif           : boolean ;
      function  ViewerVisible         : Boolean ;
      procedure OnSplitterMoved       ( Sender : TObject ) ;
      function  GetTobMasqueParam     : Tob ;
      function  MajDocEnCours         : Boolean ;
      procedure ViewerKeyDown         ( Sender: TObject; var Key: Word; Shift: TShiftState) ;
      procedure SetDocExterne         ( pDoc : TSaisieDoc );
      procedure SetDocInterne         ( pDoc : TSaisieDoc );
      procedure SetDocActif           ( pDoc : TSaisieDoc );
      procedure CreateExternalViewer;

  end ;

// =============================================================================
// extraction du Paramètrage des tables libres
procedure CGetParamLibre( vTobParamLib : Tob ; vStChamp: String ; var vStLibelle : String ; var vBoVisible : Boolean ) ;

// Chargement de tous les masques correspondant aux critères
function  CChargeMasqueListe  ( vStCrit1, vStCrit2, vStCrit3 : string ; var vTobListe : Tob ; vPiece : TPieceCompta = nil ) : boolean ; overload ;
function  CChargeMasqueListe  ( vInNumero : integer ; var vTobListe : Tob ; vStType : string = 'SAI' ) : boolean ; overload ;
// Recherche du masque le plus appropriée pour les critères depuis une TOB
function  CSelectMasqueFromTob ( vStCrit1, vStCrit2, vStCrit3 : string ; vTobListe : Tob ; vBoExact : Boolean = False; vPiece : TPieceCompta = nil ) : Tob ; overload ;
function  CSelectMasqueFromTob ( vInNumero : integer ; vTobListe : Tob ; vStType : string = 'SAI' ) : Tob ; overload ;

// =============================================================================

// Récupération des valeur de critères contextuels
function  CGetMasqueCritN( vN : integer ; vPiece : TPieceCompta ) : string ;
function  CGetMasqueCrit1( vPiece : TPieceCompta ) : string ;
function  CGetMasqueCrit2( vPiece : TPieceCompta ) : string ;
function  CGetMasqueCrit3( vPiece : TPieceCompta ) : string ;


implementation

uses
  SaisUtil,       // RDevise, trouveAuto, EstInterdit
  CPMasqueSAISIE_TOF,
  ZDevise,        // SaisieZdevise, RSDev
  {$IFDEF EAGLCLIENT}
  {$ELSE}
   {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF EAGLCLIENT}
  {$IFDEF VER150}
   Variants,
  {$ENDIF}
  Windows,           // VK_INSERT, ...
  Messages,          // WM_KEYDOWN
  hmsgbox,           // pour MessageAlerte
  ParamDat,          // ParamDate

  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  CPVersion,
  {$ENDIF MODENT1}

  Lookup,
  UtilPGI,           // OpenSelect
  Formule,           // GFormule
  Choix,             // choisir
  HSysMenu,          // THSystemMenu
  utilSoc,
  uLibWindows,
  ParamSoc,
  rtfCounter,
  hdebug,            // Debug
  CPCHANCELL_TOF,    // FicheChancell
  UTOFLOOKUPTOB,
  ULibAnalytique;

Const
 _COLORTITRE = clBtnFace ;

// =============================================================================
// extraction du Paramètrage des tables libres
procedure CGetParamLibre( vTobParamLib : Tob ; vStChamp: String ; var vStLibelle : String ; var vBoVisible : Boolean ) ;
var lTOBParam : TOB;
begin
  if not Assigned ( vTobParamLib ) then Exit ;
  lTOBParam := vTobParamLib.FindFirst( [ 'PL_CHAMP' ] , [ vStChamp ] , False ) ;
  if lTOBParam <> nil then
    begin
    vStLibelle := lTOBParam.GetValue('PL_LIBELLE');
    vBoVisible := lTOBParam.GetValue('PL_VISIBLE')='X' ;
    end ;
end ;
// =============================================================================

function  CChargeMasqueListe  ( vStCrit1, vStCrit2, vStCrit3 : string ; var vTobListe : Tob ; vPiece : TPieceCompta = nil ) : boolean ;
var lQMasque  : TQuery ;
    lStSQL    : String ;
    lTobM     : Tob ;
    lNumero   : integer ;
    lStTypeC1 : string ;
    lStTypeC2 : string ;
    lStTypeC3 : string ;
    lBoVireC1 : boolean ;
    lBoVireC2 : boolean ;
    lBoVireC3 : boolean ;
    lStCrit1  : string ;
    lStCrit2  : string ;
    lStCrit3  : string ;
begin

  if not Assigned( vTobListe ) then
    vTobListe := Tob.Create('_LISTEMASQUES', nil, -1 ) ;

  // type des critères
  lStTypeC1 := GetParamSocSecur('SO_CPMASQUECRIT1', '000') ;
  lStTypeC2 := GetParamSocSecur('SO_CPMASQUECRIT2', '000') ;
  lStTypeC3 := GetParamSocSecur('SO_CPMASQUECRIT3', '000') ;

  // Gestion de la nature de pièce à virer en BOR / LIB
  lBoVireC1 := False ;
  lBoVireC2 := False ;
  lBoVireC3 := False ;
  if Assigned( vPiece ) then
    begin
    if lStTypeC1 = '4NA' then
      lBoVireC1 := vPiece.ModeSaisie <> msPiece
    else if lStTypeC2 = '4NA' then
      lBoVireC2 := vPiece.ModeSaisie <> msPiece
    else if lStTypeC3 = '4NA' then
      lBoVireC3 := vPiece.ModeSaisie <> msPiece ;
    end ;

  // Valeur Exacte
  lStSQL := 'SELECT * FROM CMASQUECRITERES WHERE CMC_TYPE="SAI" ' ;
  if not lBoVireC1 then
    if vStCrit1<>''
      then lStSQL := lStSQL + ' AND (CMC_CRITERE1="' + vStCrit1 + '" OR CMC_CRITERE1 IS NULL OR CMC_CRITERE1="") '
      else lStSQL := lStSQL + ' AND (CMC_CRITERE1 IS NULL OR CMC_CRITERE1="") ';
  if not lBoVireC2 and (lStTypeC2 <> '000') then
    if vStCrit2<>''
      then lStSQL := lStSQL + ' AND (CMC_CRITERE2="' + vStCrit2 + '" OR CMC_CRITERE2 IS NULL OR CMC_CRITERE2="") '
      else lStSQL := lStSQL + ' AND (CMC_CRITERE2 IS NULL OR CMC_CRITERE2="") ';
  if not lBoVireC3 and (lStTypeC3 <> '000') then
    if vStCrit3<>''
      then lStSQL := lStSQL + ' AND (CMC_CRITERE3="' + vStCrit3 + '" OR CMC_CRITERE3 IS NULL OR CMC_CRITERE3="") '
      else lStSQL := lStSQL + ' AND (CMC_CRITERE3 IS NULL OR CMC_CRITERE3="") ';

  lQMasque := OpenSQL( lStSQL, True ) ;
  result   := not lQMasque.Eof ;

  if result then
    begin

    // Recup des masques trouvés
    while not lQMasque.Eof do
      begin
      lNumero   := lQMasque.FindField('CMC_NUMERO').AsInteger ;
      lStCrit1  := lQMasque.FindField('CMC_CRITERE1').Asstring ;
      lStCrit2  := lQMasque.FindField('CMC_CRITERE2').Asstring ;
      lStCrit3  := lQMasque.FindField('CMC_CRITERE3').Asstring ;

      if not Assigned( vTobListe.FindFirst(['CMS_NUMERO','CMS_CRITERE1','CMS_CRITERE2','CMS_CRITERE3'], [lNumero,lStCrit1,lStCrit2,lStCrit3], True) ) then
        begin
        lTobM := TOB.Create('CMASQUESAISIE', vTobListe, -1 ) ;
        CChargeMasque( lNumero, lTobM ) ;
        lTobM.PutValue('CMS_CRITERE1', lStCrit1 );
        lTobM.PutValue('CMS_CRITERE2', lStCrit2 );
        lTobM.PutValue('CMS_CRITERE3', lStCrit3 );
        end ;
      lQMasque.Next ;
      end ;

    end ;

  Ferme( lQMasque ) ;

end ;
// =============================================================================


// =============================================================================
function  CChargeMasqueListe  ( vInNumero : integer ; var vTobListe : Tob ; vStType : string ) : boolean ; overload ;
var lTobM     : Tob ;
begin

  if not Assigned( vTobListe ) then
    vTobListe := Tob.Create('_LISTEMASQUES', nil, -1 ) ;

  lTobM := TOB.Create('CMASQUESAISIE', vTobListe, -1 ) ;
  result := CChargeMasque( vInNumero, lTobM, vStType ) ;

end ;
// =============================================================================

// =============================================================================
// Recherche du masque le plus appropriée pour les critères depuis une TOB
function  CSelectMasqueFromTob ( vStCrit1, vStCrit2, vStCrit3 : string ; vTobListe : Tob ; vBoExact : Boolean ; vPiece : TPieceCompta ) : Tob ;
var lStTypeC1 : string ;
    lStTypeC2 : string ;
    lStTypeC3 : string ;
    lBoVireC1 : boolean ;
    lBoVireC2 : boolean ;
    lBoVireC3 : boolean ;

    function _FindFirst( vVal1, vVal2, vVal3 : string ) : Tob ;
      begin
      if lBoVireC1 then
        result := vTobListe.FindFirst( ['CMS_TYPE','CMS_CRITERE2','CMS_CRITERE3'], ['SAI',vVal2,vVal3], True )
      else if lBoVireC2 then
        result := vTobListe.FindFirst( ['CMS_TYPE','CMS_CRITERE1','CMS_CRITERE3'], ['SAI',vVal1,vVal3], True )
      else if lBoVireC3 then
        result := vTobListe.FindFirst( ['CMS_TYPE','CMS_CRITERE1','CMS_CRITERE2'], ['SAI',vVal1,vVal2], True )
      else
        result := vTobListe.FindFirst( ['CMS_TYPE','CMS_CRITERE1','CMS_CRITERE2','CMS_CRITERE3'], ['SAI',vVal1,vVal2,vVal3], True ) ;
      end ;
begin

  result := Nil ;
  if not Assigned( vTobListe ) then Exit ;
  if vTobListe.Detail.Count = 0 then Exit ;

  // type des critères
  lStTypeC1 := GetParamSocSecur('SO_CPMASQUECRIT1', '000') ;
  lStTypeC2 := GetParamSocSecur('SO_CPMASQUECRIT2', '000') ;
  lStTypeC3 := GetParamSocSecur('SO_CPMASQUECRIT3', '000') ;
  // Gestion de la nature de pièce à virer en BOR / LIB
  lBoVireC1 := False ;
  lBoVireC2 := False ;
  lBoVireC3 := False ;
  if Assigned( vPiece ) then
    begin
    if lStTypeC1 = '4NA' then
      lBoVireC1 := vPiece.ModeSaisie <> msPiece
    else if lStTypeC2 = '4NA' then
      lBoVireC2 := vPiece.ModeSaisie <> msPiece
    else if lStTypeC3 = '4NA' then
      lBoVireC3 := vPiece.ModeSaisie <> msPiece ;
    end ;

  // recherche exacte
  result := _FindFirst( vStCrit1, vStCrit2, vStCrit3 ) ;

  // On ne demande que la valeur exacte --> on sort
  if vBoExact then Exit ;
  if (vStCrit1='') and (vStCrit2='') and (vStCrit3='') then Exit ;

  // recherche 1,2,tous
  if not Assigned(result) and (vStCrit1 <> '') then
    begin
    if (vStCrit2 <> '') then
      result := _FindFirst( vStCrit1, vStCrit2, '' ) ;
    // recherche 1,tous,3
    if not Assigned(result) and (vStCrit3 <> '') then
      result := _FindFirst( vStCrit1, '', vStCrit3 ) ;
    // recherche 1,tous,tous
    if not Assigned(result) then
      result := _FindFirst( vStCrit1, '', '') ;
    end ;

  // recherche tous,2,3
  if not Assigned(result) and ( vStCrit2 <> '' ) then
    begin
    if ( vStCrit3 <> '' )  then
      result := _FindFirst( '', vStCrit2, vStCrit3 ) ;
    // recherche tous,2,tous
    if not Assigned(result) then
      result := _FindFirst( '', vStCrit2, '' ) ;
    end ;

  // recherche tous,tous,3
  if not Assigned(result) and ( vStCrit3 <> '' )  then
    result := _FindFirst( '', '', vStCrit3 ) ;

  // recherche tous,tous,tous
  if not Assigned(result) then
    result := _FindFirst( '', '', '') ;

end ;
// =============================================================================


// =============================================================================
function  CSelectMasqueFromTob ( vInNumero : integer ; vTobListe : Tob; vStType : string = 'SAI' ) : Tob ; overload ;
begin
  result := Nil ;
  if not Assigned( vTobListe ) then Exit ;
  if vTobListe.Detail.Count = 0 then Exit ;
  // recherche exacte
  result := vTobListe.FindFirst( ['CMS_NUMERO','CMS_TYPE'], [ vInNumero, vStType ], True ) ;
end ;
// =============================================================================

// =============================================================================
function  CGetMasqueCritN( vN : integer ; vPiece : TPieceCompta ) : string ;
var lStQuelleCrit : string ;
begin
  result := '' ;
  lStQuelleCrit := GetParamSocSecur('SO_CPMASQUECRIT' + IntToStr(vN), '000') ;
  // Journal
  if lStQuelleCrit = '1JO' then
    begin
    if Assigned( vPiece )
      then result := vPiece.GetEnteteS('E_JOURNAL') ;
    end
  // utilisateur
  else if lStQuelleCrit = '2UT' then
    begin
    result := V_PGI.User ;
    end
  // Etablissement
  else if lStQuelleCrit = '3ET' then
    begin
    if Assigned( vPiece )
      then result := vPiece.GetEnteteS('E_ETABLISSEMENT')
      else result := GetParamSocSecur('SO_ETABLISDEFAUT','') ;
    end
  // Nature de pièce
  else if lStQuelleCrit = '4NA' then
    begin
    if Assigned( vPiece ) then
      begin
      if vPiece.ModeSaisie = msPiece
        then result := vPiece.GetEnteteS('E_NATUREPIECE')
        else result := ''
      end
    else result := 'OD' ;
    end
  // Devise
  else if lStQuelleCrit = '5DE' then
    begin
    if Assigned( vPiece )
      then result := vPiece.GetEnteteS('E_DEVISE')
      else result := V_PGI.DevisePivot ;
    end
  // Groupe utilisateur
  else if lStQuelleCrit = '6GR' then
    begin
    result := V_PGI.Groupe ;
    end
  // Nature de journal
  else if lStQuelleCrit = '7NJ' then
    begin
    if Assigned( vPiece ) and vPiece.Info.LoadJournal( vPiece.GetEnteteS('E_JOURNAL') )
      then result := vPiece.Info.Journal.GetString('J_NATUREJAL')
      else result := '' ;
    end


end ;
// =============================================================================

// =============================================================================
function  CGetMasqueCrit1( vPiece : TPieceCompta ) : string ;
begin
  result := CGetMasqueCritN( 1, vPiece ) ;
end ;
// =============================================================================

// =============================================================================
function  CGetMasqueCrit2( vPiece : TPieceCompta ) : string ;
begin
  result := CGetMasqueCritN( 2, vPiece ) ;
end ;
// =============================================================================

// =============================================================================
function  CGetMasqueCrit3( vPiece : TPieceCompta ) : string ;
begin
  result := CGetMasqueCritN( 3, vPiece ) ;
end ;
// =============================================================================


 { TSaisiePiece }

procedure TSaisiePiece.DeleteRow( vRow : Integer ; vBoForce : Boolean );
var lInGrp : integer ;
    lRow   : integer ;
    lCol   : integer ;
begin
  if GetActionFiche = taConsult then Exit ;

  if vRow = -1 then vRow := Row ;
  if Piece.isOut( vRow ) then Exit ;

  if not vBoForce then
    begin
    if GetCount <= 1 then Exit ;
    if ( vRow = GetCount ) and not Piece.EstRemplit( vRow ) then Exit ;
    end ;

  lRow   := vRow ;
  lCol   := Col ;
  lInGrp := 0 ;

  if Piece.ModeSaisie = msBor then
    begin
    lInGrp := Piece.GetNumGroupe( lRow ) ;
    end ;

  Piece.DeleteRecord( lRow ) ;

  if (( lRow >= GetCount ) and Piece.EstRemplit( GetCount ))
    or ( vBoForce and (GetCount=0))
    then Piece.NewRecord ;

  AfficheLignes ;

// Repositionnement...
  if lRow > GetCount then
    Dec(lRow)
  else if Piece.ModeSaisie = msBor then
    begin
    if Piece.GetNumGroupe( vRow ) > lInGrp
      then Dec(lRow) ;
    end ;

  if not Piece.EstRemplit( vRow )
    then lCol := GetColIndex('E_GENERAL') ;

  SetPos( lcol, lRow, False) ;


end;

procedure TSaisiePiece.InsertRow( vRow : Integer = -1 ; vBoForce : Boolean = False );
begin

  if GetActionFiche = taConsult then Exit ;

  if vRow = -1 then vRow := Row ;

  // Si la ligne cible est déjà non remplie, on ne fait rien
  if (not vBoForce) and (not FTobPiece.EstRemplit( vRow ) ) then Exit ;

  // Gestion des evènements de sortie
  if (not vBoForce) and not ValideSortieLigne then Exit ;

  // Ajout d'une ligne d'écriture dans la pièce
  Piece.NewRecord( vRow ) ;

  // MAJ affichage grille
  AfficheLignes ;

end;


procedure TSaisiePiece.FindDialogFind(Sender: TObject);
var Cancel: Boolean;
begin
  Rechercher (FListe, FFindDialog, FirstFind ) ;
  RowEnter   (FListe, FListe.Row,   Cancel, True ) ;
end;


function TSaisiePiece.GetRow: Integer;
begin
  Result := FListe.Row ;
end;

procedure TSaisiePiece.AfficheLignes( vRow : Integer ; vBoForce : boolean );
var  //HMT : TComponent ;
     i    : Integer ;
     j    : Integer ;
     lTob : Tob ;
begin

{$IFDEF SBO}
AddEvenement('TSP_AfficheLignes vRow:=' + intToStr(vRow) ) ;
{$ENDIF}

  if not assigned( FTobMasque ) then Exit ;

  try

    // Mise en place de la liste des champs
    Fliste.BeginUpdate ;

    if FListe.RowCount > GetCount + 2 then
      For j:=GetCount+2 to FListe.RowCount - 1 do
        For i:=0 to FListe.colCount - 1 do
          FListe.Cells[ i, j ] := '' ;

    if FListe.RowCount <> GetCount + 2 then
      FListe.RowCount := GetCount + 2 ;

    // ====  AFFICHAGE 1 LIGNE ====
    if not Piece.isOut( vRow ) then
      begin

      CAfficheLigne( vRow, FTobPiece.Detail[vRow-1], vBoForce ) ;
      if (Piece.modeEche = meMulti) and Piece.EstMultiEche( vRow ) then
        begin
        if Piece.GetDouble(vRow, 'E_NUMECHE')=1 then
          for i := vRow + 1 to (vRow + Piece.GetNbEche( vRow ) - 1) do
            CAfficheLigne( i, Piece.Detail[i-1], vBoForce ) ;
        end ;

      end
    else
    // ====  AFFICHAGE TOUTES LIGNES ====
      begin

//      if not ( vBoForce or Piece.IsOneModifie(True) ) then Exit ;

      For i:=0 to FTobPiece.Detail.Count-1 do
        begin
        lTob := FTobPiece.Detail[i] ;
        CAfficheLigne( i+1, lTob, vBoForce ) ;
        end ;

      // Dernière ligne vierge
      For i:=0 to FListe.colCount - 1 do
        FListe.Cells[ i, FListe.RowCount - 1 ] := '' ;

      end ;

  finally
    // Fin des modifs
    FListe.EndUpdate ;
//    FListe.Invalidate;
{
    if Piece.IsOut( vRow )
      then Piece.SetAllModifie(False)
      else Piece.Detail[vRow-1].SetAllModifie(False) ;
}
    end ;

end;

procedure TSaisiePiece.SetGrille(const Value: THGrid);
begin
   FListe := Value;
end;


constructor TSaisiePiece.Create( vEcran: TForm ; vGrille : THGrid ; vPieceCpt : TPieceCompta );
begin

  FEcran         := vEcran;
  FListe         := vGrille;
  FStMasqueCrit1 := '' ;
  FStMasqueCrit2 := '' ;
  FStMasqueCrit3 := '' ;
  FGrilleCompl   := nil ;
  FBoAccActif    := vPieceCpt.AccActif ;

  FTobPiece      := vPieceCpt ;

  // Création des objets
  FFindDialog        := TFindDialog.Create( Ecran );
  FFindDialog.name   := 'FindDialog';
  FFindDialog.OnFind := FindDialogFind;

  // var de gestion
  FInLigneCourante     := -1 ;
  FInIndexComptaAuto   :=  1 ;

  // Gestion de la personnalisation des champs libre
  InitParamLib ;

  // Réaffectation des evts de la grille
  InitEvtGrid ;
  InitFormatGrid ;

  // Gestion des lookup
  InitRechCompte ;

  // gestion des messages
  InitMessageCompta ;

  // gestionnaire de guides
  FTSGuide := TSaisieGuide.Create( self ) ;

  // gestionnaire du viewer
  FViewer   := GetSynRegKey('PosVisu'   ,'INT'   ,true);
  FTSDocExt := TSaisieDoc.Create( FEcran, FTobPiece ) ;
  FTSDocInt := TSaisieDoc.Create( FEcran, FTobPiece ) ;
  CreateExternalViewer;
  if FViewer = 'EXT'
    then SetDocActif(FTSDocExt)
    else SetDocActif(FTSDocInt);
  FBoDocMAJEnCours  := false ;

{$IFDEF SBO}
 AddEvenement('=============================================');
 AddEvenement('TSP_Create >> Début : ' + DateToStr(Now) );
{$ENDIF SBO}

end;

destructor TSaisiePiece.Destroy;
begin

  // Paramétrage champs libre
  if assigned(FTobParamLib) then
    FreeAndNil( FTobParamLib ) ;

  // Composants
  FreeAndNil( FFindDialog ) ;
  FreeAndNil( FRechCompte ) ;
  FreeAndNil( FMessageCompta ) ;

  if Assigned( FListeMasques ) then
    FreeAndnil( FListeMasques ) ;

  if Assigned( FTSGuide ) then
    FreeAndnil( FTSGuide ) ;

  if Assigned(FZlibAuto) then
    FreeAndNil( FZlibAuto ) ;

  if Assigned( FTSDocInt ) then begin
                             FTSDocInt.DesactiveViewer ;
                             FreeAndNil( FTSDocInt ) ;
                             end ;
  if Assigned( FTSDocExt ) then begin
                             FTSDocExt.DesactiveViewer ;
                             FreeAndNil( FTSDocExt ) ;
                             end ;
  if Assigned( FViewerExt ) then begin
                             FViewerExt.TSDocInt := nil;
                             FreeAndNil(FViewerExt);
                             end ;
  FTSDoc    := nil;

{$IFDEF SBO}
  AddEvenement('TSP_Destroy >> Fin : ' + DateToStr(Now) );
  AddEvenement('=============================================');
{$ENDIF SBO}

  inherited Destroy;

end;


function TSaisiePiece.GetCount: Integer;
begin
  Result := Piece.Detail.count ;
end;

function TSaisiePiece.GetTOBCourante: TOB;
begin
  Result := nil ;
  if not FTobPiece.IsOut( GetRow ) then
    Result := FTobPiece.GetTob( GetRow ) ;
end;

procedure TSaisiePiece.CellEnter( Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean );
var lInSens   : integer ;
    lStChamps : String ;
    lCol      : integer;
    lRow      : integer;
    lStQuelDC : String ;
    lTobEcr   : Tob ;
    lBoEdit   : Boolean;
begin

  if ( csDestroying in Ecran.ComponentState) then Exit ;
  if not FBoEvtOn then Exit ;
  if GetActionFiche = taConsult then Exit ;
  if (GoRowSelect in FListe.Options) then Exit ;

{$IFDEF SBO}
AddEvenement('TSP_CellEnter ACol:='+intToStr(ACol)+' ARow:='+intToStr(ARow));
{$ENDIF}

try

  FListe.ElipsisButton := False ;

  if FTobPiece.IsOut( FListe.Row ) then
    begin
    Cancel := True ;
    Exit ;
    end ;

  // Détermination du sens de navigation dans la grille (fn uLibWindows)
  lCol      := FListe.col ;
  lRow      := FListe.Row ;
  lInSens   := CGetGridSens(FListe , ACol, ARow) ; // 1 : de gauche à droite //  -1 : de droite à gauche
  lStChamps := FListe.ColNames[ lCol ] ;

  // Détermination de la prochaine cellule valide
  if not EstColEditable( lRow, lStChamps ) then
    begin
    // Gestion de la saisie en devise
    if (Piece.ModeSaisie = msBOR) and (Piece.GetEnteteS('E_DEVISE')<>Piece.GetValue( lRow, 'E_DEVISE')) then
      if ( lStChamps = 'E_DEBITDEV' ) or ( lStChamps = 'E_CREDITDEV' ) then
          if ( Piece.GetMontantDev(lRow) = 0 ) then
            SaisieDevise( lRow ) ;

    lBoEdit := (GoEditing in FListe.OPtions) ;
    if (ACol = 0) and lBoEdit then
      begin
      ACol := GetColIndex( 'E_GENERAL') ;
      ARow := lRow ;
      end
    else GetColSuivante( ACol, ARow) ;
    Cancel := true ;
    Exit ;
    end ;

  // Faire apparaître les boutons ellipsis bouton sur les champs approprié
  FListe.ElipsisButton := EstColLookup ;

  // =======================================
  // ===== GESTION SAISIE DES MONTANTS =====
  // =======================================
  // Gestion de la saisie en devise
  if (Piece.ModeSaisie<>msPiece) and (( lStChamps = 'E_DEBITDEV' ) or ( lStChamps = 'E_CREDITDEV' )) then
    if (Piece.GetEnteteS('E_DEVISE')<>V_PGI.DEvisePivot) and (Piece.GetEnteteS('E_DEVISE')=Piece.GetValue( lRow, 'E_DEVISE')) then
      if (Piece.GetString( lRow, 'PBTAUX' ) <> 'X') and ( Piece.GetDebutGroupe( lRow ) = lRow ) and ( Piece.GetMontantDev(lRow) = 0 ) then
        if not Piece.IsValidTaux( lRow ) then
          begin
          if PgiAsk( TraduireMemoire('Voulez-vous saisir ce taux dans la table de chancellerie ?')) = mrYes then
            begin
            FicheChancel( Piece.GetString(lRow,'E_DEVISE'), True, Piece.GetDateTime(lRow, 'E_DATECOMPTABLE'), taCreat, True ) ; //( lDtDateC >= V_PGI.DateDebutEuro ) ) ;
            Piece.MajDeviseTaux( lRow, True ) ;
            end ;
          lTobEcr := Piece.GetTob(lRow) ;
          lTobEcr.PutValue('PBTAUX', 'X') ;
          end ;

  // -----------------
  // ----- DEBIT -----
  // -----------------
  if ( lStChamps = 'E_DEBITDEV' ) then
    begin
    // ---> aide à la saisie
    if ( FTobPiece.GetValue( lRow , 'E_NATUREPIECE') <> 'OD' )
            and ( FTobPiece.GetValue( lRow , 'E_DEBITDEV') = 0 )
            and ( lInSens = RC_GAUCHE_DROITE ) then
      begin
        lStQuelDC := Piece.QuelChampsMontant( lRow ) ;
        if lStQuelDC='' then lStQuelDC := lStChamps ;
        if (lStQuelDC<>lStChamps) then
          begin
          ACol      := GetColIndex( lStQuelDC ) ; // on force le positionnement au credit
          ARow      := lRow ;
          Cancel    := true ;
          exit ;
          end;
      end;
    end
  // ------------------
  // ----- CREDIT -----
  // ------------------
  else if ( lStChamps = 'E_CREDITDEV') then
    begin
    // ---> aide à la saisie
    if ( FTobPiece.GetValue( lRow , 'E_DEBITDEV') = 0 )
            and ( FTobPiece.GetValue( lRow , 'E_CREDITDEV') = 0 )
            and ( lInSens = RC_DROITE_GAUCHE ) then
      begin
      lStQuelDC := FTobPiece.QuelChampsMontant( lRow ) ;
      if lStQuelDC='' then lStQuelDC := lStChamps ;
      if (lStQuelDC<>lStChamps) then
        begin
          begin
          ACol    := GetColIndex( lStQuelDC ) ; // on force le positionnement au debit
          ARow    := lRow ;
          Cancel  := true ;
          exit ;
          end;
        end;
      end;
    end
    ;


finally

  AfficheComplement( Row ) ;

  // Appel cellenter client
  if Assigned( FUserCellEnter ) then
    FUserCellEnter( Sender, Acol, Arow, Cancel );


end ; // finally

end;

procedure TSaisiePiece.CellExit( Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean );
var lStChamp : String ;
    lTobLigne : TOB ;
    lInOldCol : Integer ;
    lInOldRow : Integer ;
    lStVal    : String ;
    lStValG   : String ;
begin

  if csDestroying in Ecran.ComponentState then Exit ;
  if not FBoEvtOn then Exit ;
  if GetActionFiche = taConsult then Exit ;
  if (GoRowSelect in FListe.Options) then Exit ;

{$IFDEF SBO}
AddEvenement('TSP_CellExit ACol:='+intToStr(ACol)+' ARow:='+intToStr(ARow));
{$ENDIF}

 lInOldCol := ACol ; // sauvegarde des colonnes car le TRechCompte manipule la colonne
 lInOldRow := ARow ; // sauvegarde des colonnes car le TRechCompte manipule la colonne

 Fliste.BeginUpdate ;
 try

  if FTobPiece.IsOut( ARow ) then Exit ;

  // récupération du nom du champs
  lStChamp := FListe.ColNames[ ACol ] ;
  if lStChamp = '' then Exit ;
  if Copy( Trim( lStChamp ), 1, 1 ) = '@' then Exit ; // Champ formule

  // champs éditables ?
  if not EstColEditable( lInOldRow, lStChamp ) then Exit ;

  // Tests avant traitement sur tob courante
  lTobLigne := FTobPiece.GetTob( ARow ) ;
  if lTobLigne = nil then Exit ;
//  if lTobLigne.GetNumChamp( lStChamp ) < 0 then Exit ;

  // -----------------------------------
  // -----  Gestion des recherches -----
  // -----------------------------------
  if lStChamp = 'E_GENERAL' then
    begin
    RechCompte.Ecr := lTobLigne ;
    if ( FListe.Cells[ACol, ARow] <> '' ) or
       ( ( FListe.Cells[ACol, ARow] ='' ) and ( FListe.Cells[ GetColIndex('E_AUXILIAIRE') , ARow]<>'' ) ) then
      RechCompte.CellExitGen( FListe, ACol, ARow, Cancel ) ;
    if Cancel then
      Exit ;
{     // Le Focus control n'est plus utilisé dans le noyau
      if (RechCompte.FocusControl = ControlGen ) //and ( FListe.Cells[ lInOldCol, lInOldRow ] = '' )  //( lTobLigne.GetString('E_GENERAL') = '' )
        then exit
        else cancel := not GuideActif ;}
    end
  else if lStChamp = 'E_AUXILIAIRE' then
    begin
    RechCompte.Ecr := lTobLigne ;
    RechCompte.CellExitAux( FListe, ACol, ARow, Cancel ) ;
    if Cancel then
      Exit ;
{     // Le Focus control n'est plus utilisé dans le noyau
      if (RechCompte.FocusControl = ControlAux )//and ( FListe.Cells[ lInOldCol, lInOldRow ] = '' )  //( lTobLigne.GetString('E_AUXILIAIRE') = '' )
        then exit
        else cancel := not GuideActif ;}
    end
  else if EstColCombo ( ACol ) then
    begin
    lStValG := GetValueGrid( ACol,ARow ) ;
    if lStValG <> (FListe.Cells[ACol, ARow]) then
      FListe.Cells[ACol, ARow] := lStValG ;

    lStVal  := RechDom( GetChampDataType( lStChamp ), lStValG,  False ) ;
    if ( (lStChamp = 'E_NATUREPIECE') or ( lStValG <> '')) and
       ( ( lStVal = 'Error' ) or ( lStVal = '' ) )  then
      begin
      Cancel := True ;
      PostMessage( FListe.Handle, WM_KEYDOWN, VK_F5, 0) ;
      Exit ;
      end ;
    end
   else if EstColLookup ( ACol ) then
     begin
     lStValG := GetValueGrid( ACol,ARow ) ;
     if ( (FListe.ColTypes[ACol]<>'D') and ( Trim(lStValG) <> '' ) and ( lStValG <> lTobLigne.GetString( lStChamp ) ) )
        or
        ( (FListe.ColTypes[ACol]='D') and ( Trim(FListe.CellValues[ACol,ARow]) <> '' ) and (not HEnt1.IsValidDate( FListe.CellValues[ACol,ARow] ) ) ) then
       begin
       Cancel := True ;
       PostMessage( FListe.Handle, WM_KEYDOWN, VK_F5, 0) ;
       Exit ;
{      Cancel := not RecherchePourChamps( lStChamp ) ;
       if cancel then
         Exit ;
}       end ;
     end ;

  // -------------------------------------------------------------------
  // ----- MAJ des données de la tob courante si changement  -----------
  // -------------------------------------------------------------------
  ValideCellule( lStChamp, lInOldRow ) ;

 finally

  if Assigned( FUserCellExit ) then
    FUserCellExit( Sender, lInOldCol, Arow, Cancel );

  if GuideActif then
    begin
    ACol := lInOldCol ;
    ARow := lInOldRow ;
    TSGuide.GuideCellExit( Sender, ACol, ARow, Cancel ) ;
    end ;

  FListe.EndUpdate ;
 end;

end;

procedure TSaisiePiece.KeyDown( Sender: TObject ; var Key: Word ; Shift: TShiftState ) ;
var Vide      : Boolean ;
    lKeyParam : Word ;
begin

  if ( csDestroying in Ecran.ComponentState) then Exit ;

{$IFDEF SBO}
AddEvenement('KeyDown     >> FBoEvtOn : ' + BoolToStr_( FBoEvtOn ) ) ;
{$ENDIF}

  //  if not FBoEvtOn then Exit ;
  if not FBoEvtOn then
    SetEvtOn ;

  if Not FListe.SynEnabled then
    begin
    Key:=0 ;
    Exit ;
    end ;

  // Récup des touches pour les options de recherche
  lKeyParam        := Key ; // on recupère le Key pour appel du TSP_KeyDown
  RechCompte.InKey := Key ;
  RechCompte.Shift := Shift ;

  try

    if GuideActif then
      TSGuide.GuideKeyDown(Sender, Key, Shift);

    Vide:=(Shift=[]) ;

    case Key of
      VK_RETURN  : if Vide then
                      begin
                      Key := VK_TAB  ;
                      RechCompte.InKey := Key ;
                      PostMessage( FListe.Handle, WM_KEYDOWN, VK_TAB, 0) ;
                      end ;
      VK_F5 : if Shift<>[ssShift] then
                     begin
                     Key:=0 ;
                     ElipsisClick( Sender ) ;
                     end ;
      VK_INSERT : if (Vide) then
                    begin
                    Key:=0;
                    InsertRow( GetRow ) ;
                    end ;
      VK_DELETE : if Shift=[ssCtrl] then
                    begin
                    Key:=0 ;
                    DeleteRow( GetRow ) ;
                    end ;
      VK_ESCAPE : if Vide then
                    begin
                    Key:=0 ;
  //                  AbandonClick(True) ;
                    end ;
      VK_END :   begin
                 Key:=0 ;
                 SetPosFin( Shift=[ssCtrl] ) ;
                 end ;

      VK_HOME :  begin
                 Key:=0 ;
                 SetPosDebut( Shift=[ssCtrl] ) ;
                 end ;

      VK_F4 : begin
              Key:=0 ;
              GetCompteAuto( True ) ;
              end ;
            
      VK_F3 : begin
              Key:=0 ;
              GetCompteAuto( False ) ;
              end ;

      VK_F6  : if Vide then
                 begin
                 Key:=0 ;
                 SoldePiece ;
                 end ;
       VK_F7  : begin
                if Shift<>[ssCtrl] then
                  begin
                  Key:=0 ;
                  DupliquerZone( Shift=[ssShift], Shift=[ssAlt] ) ;
                  end ;
                end ;

  {CTRL +}107 : if Shift=[ssCtrl] then
                  begin
                  Key:=0 ;
                  IncrementerRef( Row ) ;
                  end ;

  {CTRL -}109 : if Shift=[ssCtrl] then
                  begin
                  Key:=0 ;
                  IncrementerRef( Row, False ) ;
                  end ;

  {^D}     68 : if Shift=[ssCtrl] then
                   begin
                   Key:=0 ;
                   if ( GetParamsocSecur('SO_CPLIENGAMME','') = 'S1' ) then
                     begin
                     if ( Row = 1 )
                       then SaisieDevise( 1 ) ;
                     end
                   else SaisieDevise( Row ) ;
                   end ;

  {AG}     71 : if Shift=[ssAlt]  then
                   begin
                   Key:=0 ;
                   GereGuide( Row ) ;
                   end
                 else if Shift=[ssCtrl] then
                   begin
                   Key:=0 ;
                   LanceGuide( Row ) ;
                   //CGuideRun ;
                   end ;

  {^K}     75 : if Shift=[ssCtrl] then
                  begin
                  Key:=0 ;
                  ProratisePiece ;
                  end ;

  {AL}     76 : if Shift=[ssAlt]  then
                  begin
                  Key:=0 ;
                  ExecuteLibelleAuto( Row );
                  end ;

  {^M}     77 : if Shift=[ssCtrl]  then
                  begin
                  Key:=0 ;
                  if EstSpecif('51212') then
                    ChoisirMasque ;
                  end ;

  {^S}     83 : if Shift=[ssCtrl]       then
                  begin
                  Key:=0 ;
                  Piece.InitLibelle( Row, True ) ;
                  AfficheLignes ;
                  end ;

  {^W}     87 : if Shift=[ssCtrl] then
                  begin
                  Key:=0 ;
                  GereAcc( Row, True ) ;
                  end ; // necessecaire de remettre Key:=0, gere le cas ou l'on a ppuie plusieur de fois de suite sur ctrl+w en sur un auxi en milieu de grille ( sinon generation de ligne vide qd on tabule )

  {^Y}     89 : if Shift=[ssCtrl] then
                  begin
                  Key:=0 ;
                  videLignes ;
                  end ;

  {^Z}     90 : if Shift=[ssCtrl] then // Stocke un compte HT pour l'accélérateur de saisie
                  begin
                  Key:=0 ;
                  FStAccCptHT := Piece.GetString( Row, 'E_GENERAL') ;
                  end ;

  {=}     187 : if Vide then
                 begin
                 Key:=0 ;
                 SoldePiece ;
                 end ;

    end ;

  finally

    if (lKeyParam<>0) and Assigned( FUserKeyDown ) then
      FUserKeyDown( Sender, lKeyParam, Shift );

    end ;

end;

procedure TSaisiePiece.RowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var lInIdx  : integer ;
    lInGrp  : integer ;
    lInDe   : integer ;
    lInA    : integer ;
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  if not FBoEvtOn then Exit ;

{$IFDEF SBO}
AddEvenement('TSP_RowEnter ou:='+intToStr(Ou)) ;
{$ENDIF}

  if (GetActionFiche <> taConsult) and not Piece.IsOut( FListe.row ) and not GuideActif then
    begin

    // Accès en modification
    GereOptionsGrid ;

    // Si on arrive dans une ligne vide, on se place sur la saisie du compte
    if not Piece.EstRemplit( FListe.row )
          and ( GetColName <> 'E_GENERAL' )
          and ( GetColName <> 'E_AUXILIAIRE') then
      begin
      // Sur changement de groupe en BOR / libre, on revient sur le jour, sinon le compte général
      if ( Piece.ModeSaisie in [msBOR,msLibre] )
          and ( FListe.row > 1 )
          and ( Piece.GetValue( FListe.row-1, 'E_NUMGROUPEECR') <> Piece.GetValue( FListe.row, 'E_NUMGROUPEECR') )
        then begin
             lInIdx := GetColIndex( 'E_DATECOMPTABLE' ) ;
             if lInIdx > 0
                then FListe.Col := lInIdx
                else FListe.col := GetColIndex( 'E_GENERAL')
             end
        else FListe.col := GetColIndex( 'E_GENERAL') ;
      end ;

    end ;

  // chargement Devise si besoin
  if not Piece.IsOut( FListe.row ) and ( Piece.ModeSaisie = msBor ) then
    begin
    Piece.Info_LoadDevise( Piece.GetString( Row, 'E_DEVISE' ) ) ;
    end ;

  // MAJ doc ged si besoin
  if (Piece.ModeSaisie=msBor) and ViewerVisible and (not MajDocEnCours)  then
    begin
    lInGrp := Piece.GetInteger( ou, 'E_NUMGROUPEECR' ) ;
    if lInGrp > 1 then
      begin
      Piece.GetBornesGroupe( TSDoc.GroupeActif, lInDe, lInA ) ;
      if lInDe <> lInA then
        // Si groupe lié au doc soldé, on passe au suivant
        if ( lInGrp > TSDoc.GroupeActif ) and ( Piece.GetSoldeGroupe( TSDoc.GroupeActif ) = 0 ) then
          begin
          ChargeDocument ; // on passe au doc suivant
//          TSDoc.GroupeActif := TSDoc.GroupeActif + 1 ;
          end ;
      end ;
    end ;

  // Affichage des infos de la grille
  AfficheComplement ( Row ) ;

  // Appel rowenter client
  if Assigned( FUserRowEnter ) then
    FUserRowEnter( Sender, Ou, Cancel, Chg );

end;

procedure TSaisiePiece.RowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var lStChamps    : String ;
    lInNumGroupe : integer ;
    lInCol       : integer ;
begin

  if ( csDestroying in Ecran.ComponentState) then Exit ;
  if not FBoEvtOn then Exit ;

{$IFDEF SBO}
AddEvenement('TSP_RowExit Ou:='+intToStr(Ou)) ;
{$ENDIF}

try

  if GuideActif then
    begin
    TSGuide.GuideRowExit( Sender, Ou, Cancel, Chg ) ;
    Exit ;
    end ;

  if GetActionFiche = taConsult then Exit ;

  if FTobPiece.EstRemplit( ou ) then
    begin
    // Gestion des erreurs de validation de la ligne
    Cancel := not FTobPiece.IsValidLigne( Ou ) ;
    if Cancel then
      begin
      Case FTobPiece.LastError.RC_Error of
        // Erreur sur le compte général
        RC_COMPTEINEXISTANT,
          RC_CINTERDIT,
          RC_CFERMESOLDE,
          RC_CFERME,
          RC_CNATURE,
          RC_COMPTECONFIDENT  : FListe.Col := GetColIndex('E_GENERAL') ;
        // Erreur sur le compte auxiliaire
        RC_AUXINEXISTANT,
          RC_AUXOBLIG,
          RC_NATAUX           : FListe.Col := GetColIndex('E_AUXILIAIRE') ;
        // Erreur sur les montants
        RC_MONTANTINEXISTANT  : begin
                                lStChamps := FTobPiece.QuelChampsMontant ( ou ) ;
                                if lStChamps = '' then
                                  lStChamps := 'E_DEBITDEV' ;
                                FListe.Col := GetColIndex( lStChamps ) ;
                                end ;
        RC_MODEPAIEINCORRECT :  begin
                                lInCol := GetColIndex('E_MODEPAIE') ;
                                if lInCol > 0 then
                                  FListe.Col := lInCol ;
                                end ;
        RC_DATEECHEINCORRECTE :  begin
                                 lInCol := GetColIndex('E_DATEECHEANCE') ;
                                 if lInCol > 0 then
                                   FListe.Col := lInCol ;
                                 end ;
      end ;//case
      FListe.Invalidate;
      Exit ;
      end ;  // if cancel

    end ; // if FTobPiece.EstRemplit( ou ) then

    if Piece.IsOut( FListe.Row ) then
      GereNouvelleLigne( ou, FListe.row ) ;

    if Piece.IsOut(Fliste.Row) then
      begin
      Cancel := True ;
      Exit ;
      end ;

    // En mode bordereau, il faut vérifier l'équilibre du groupe
    if (Piece.ModeSaisie=msBor) and (not Piece.IsOut( Ou ) and (Piece.count > 1) ) then
      begin
      lInNumGroupe := Piece.GetInteger(Ou, 'E_NUMGROUPEECR') ;

      if ( lInNumGroupe <> Piece.GetInteger( FListe.row, 'E_NUMGROUPEECR') )
          and ( Piece.GetSoldeGroupe( lInNumGroupe ) <> 0 ) then
         begin
         Cancel := True ;
         Exit ;
         end ;

      end ;


finally
  // Appel rowexit client
  if Assigned( FUserRowExit ) then
    FUserRowExit( Sender, Ou, Cancel, Chg );

  if not Cancel then
    GereAcc( Ou ) ;

  // Affichage mono-section // FQ 21687
  if Piece.GetString(Ou, 'E_ANA')='X' then
    begin
    Piece.SynchroMonoSection( Ou ) ;
    AfficheLignes(Ou) ;
    end ;


end ; // try


end;


procedure TSaisiePiece.InitEvtGrid;
begin

  // Récupération des evt ancêtres
  if Assigned( FListe.OnDblClick )
    then FInheritedDblClick := FListe.OnDblClick
    else FInheritedDblClick := nil ;

  FBoEvtOn               := True ;
//  FBoFirstEnter          := True ;

  // Réaffectation des évènements
  FListe.OnRowEnter      := RowEnter ;
  FListe.OnRowExit       := RowExit ;
  FListe.OnCellEnter     := CellEnter ;
  FListe.OnCellExit      := CellExit ;
  FListe.OnKeyDown       := KeyDown ;
  FListe.OnKeyPress      := KeyPress ;
  FListe.PostDrawCell    := PostDrawCell ;
  FListe.OnElipsisClick  := ElipsisClick ;
  FListe.OnEnter         := GridEnter ;
  FListe.OnExit          := GridExit ;
  FListe.OnMouseDown     := MouseDown ;
  FListe.OnDblClick      := DblClick ;
  FListe.GetCellCanvas   := GetCellCanvas;
end;

procedure TSaisiePiece.InitFormatGrid;
begin
  FListe.RowCount      := 2 ;
  FListe.ElipsisButton := False ;
  FListe.FlipBool      := True ;
  FListe.CalcInCell    := True ;
//  FListe.onComputeCell := ComputeCell ;
//  FListe.TraiteMontantNegatif := True ;
  SetLength( FTabColAvecLib, FListe.ColCount + 1 ) ;
  FListe.RowHeights[0] := 18 ; // hauteur de la ligne de titre
end;


procedure TSaisiePiece.SetRow(vIndex: Integer);
begin
  FListe.Row := vIndex ;
end;


procedure TSaisiePiece.GereOptionsGrid;
begin

  Case getActionFiche of
    taCreat,taModif :   EditEnabled( FTobPiece.EstLigneModifiable( GetRow ) ) ;
//    taCreat :   EditEnabled( True ) ;
    taConsult : EditEnabled( False ) ;
    end ;

end;

function TSaisiePiece.GetActionFiche: TActionFiche;
begin
  result := taConsult ;
  if Piece<>nil then
    Result := Piece.Action ;
end;

procedure TSaisiePiece.EditEnabled( vBoAvec : Boolean );
var lBoEdit : Boolean ;
begin

  // Option de départ
  lBoEdit := (GoRowSelect in FListe.Options) ;

  // Accès cellules
  if vBoAvec then
      begin
      FListe.Options := FListe.Options - [ GoRowSelect ] ;
      FListe.Options := FListe.Options + [ GoEditing, GoTabs, GoAlwaysShowEditor ] ;
      FListe.MontreEdit ;
      end
    else
      begin
      FListe.CacheEdit ;
      FListe.Options := FListe.Options - [ GoEditing, GoTabs, GoAlwaysShowEditor ] ;
      FListe.Options := FListe.Options + [ GoRowSelect ] ;
      end ;

  // Rafraichissement si nécessaire
  if lBoEdit <> (GoRowSelect in FListe.Options) then
      FListe.Refresh ;   //    FListe.Invalidate ;

end;

function TSaisiePiece.GetChampsListe( vCol : Integer = - 1): String;
var i : Integer ;
begin
  result := '' ;
  for i := 0 to FListe.ColCount - 1 do
    begin
    if ( vCol < 0 ) or ( i = vCol )
      then result := result + FListe.ColNames[ i ] + ';'
      else result := result + ';' ;
    end ;
end;


function TSaisiePiece.GetValueGrid(aCol, aRow: integer) : Variant ;
var lStChamp   : String ;
    lStCellVal : Variant ;
    lInType    : Integer ;
    lInJour    : integer ;
    lDtDate    : TDateTime ;
    lInDays    : Extended ;
    lInYear    : Word ;
    lInMonth   : Word ;
    lInDay     : Word ;
begin
  lStCellVal := FListe.CellValues[aCol, ARow] ;
  lStChamp   := FListe.ColNames[aCol] ;

  // cas spécifique de la saisie du jour en mode bordereau
  if (lStChamp = 'E_DATECOMPTABLE') {and (Piece.ModeSaisie<>msPiece) }then
    begin
    lDtDate    := Piece.GetEnteteDt('E_DATECOMPTABLE') ;
    lInDays    := FinDeMois( lDtDate ) - DebutDeMois( lDtDate ) + 1 ;
    DecodeDate( lDtDate, lInYear, lInMonth, lInDay ) ;
    lInJour    := ValeurI( lStCellVal ) ;
    if (lInJour >= 1) and (lInJour <= lInDays )
      then result := EncodeDate( lInYear, lInMonth, lInJour )
      else result := FinDeMois( Piece.GetEnteteDt('E_DATECOMPTABLE') ) ;
    Exit ;
    end
  else
    lInType := VarType( FTobPiece.Detail[ aRow - 1 ].GetValue(lStChamp) ) ;

  Case lInType of
    varNull    : result := '';
    varDate    : if HEnt1.IsValidDate( lStCellVal )
                   then result := StrToDate( lStCellVal )
                   else result := iDate1900 ;
    varDouble  : result := Valeur(    lStCellVal ) ;
    varInteger : result := ValeurI(   lStCellVal ) ;
    varBoolean : result := uppercase( lStCellVal ) = 'X' ;
    varString  : if EstColCombo ( ACol )
                   then result := AnsiUpperCase( Trim( lStCellVal ) )
                   else result := Trim( lStCellVal ) ;
    else         result := '';
    end ;

end;
{
procedure TSaisiePiece.PutGridToTob(aCol, aRow: integer);
var lStChamp : String ;
    lValue   : Variant ;
begin

  lValue     := GetValueGrid(aCol, aRow) ;
  lStChamp   := FListe.ColNames[aCol] ;

  if VarType( lValue ) = VarType( FTobPiece.Detail[ aRow - 1 ].GetValue(lStChamp) ) then
    FTobPiece.Detail[ aRow - 1 ].PutValue( lStChamp, lValue ) ;

end;
}

procedure TSaisiePiece.PutTobToGrid( vRow : integer ; vStChamps : string ) ;
var lStChamp : String ;
    lStListe : String ;
    lStVal   : String ;
    lTob     : Tob ;
    lInCol   : integer ;
begin

  if Piece.IsOut( vRow ) then Exit ;
  if Trim(vStChamps) = '' then Exit ;

  lStListe   := vStChamps ;
  lTob       := Piece.Detail[ vRow-1] ;

  while (lStListe <> '') do
    begin
    lStChamp := ReadTokenSt( lStListe ) ;
    lInCol   := GetColIndex( lStChamp ) ;
    lStVal   := GetValeurPourGrille ( lStChamp, lTob, lInCol ) ;
    FListe.CellValues[lInCol, vRow] := lStVal ;
    end ;

end;

procedure TSaisiePiece.PostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var lBoGrise          : boolean ;
    lStCol            : String ;
    lRect             : TRect ;
    lValue            : String ;
    lText             : String ;
    lBoSep            : Boolean ;
    lBoHaut           : Boolean ;
    lBoReappliquer    : Boolean ;
    lInPosX           : Integer ;
    lTobEcr           : Tob ;
    lInAxe            : Integer ;
    lStChp            : string ;
    lStBlocNote       : Variant ;
    T1,T2,T3          : TPoint ;
    lTobCompl         : Tob ;
begin

  if csDestroying in Ecran.ComponentState then Exit ;

  If ARow<=0 Then Exit ;

  if ACol<1 then Exit ;
  if FTobPiece.Detail.count = 0 then Exit ;

  lStCol   := GetColName( ACol ) ;

  if not Piece.isOut( ARow ) then
    begin

    // gestion des cases debit et crédit - une seul des deux doit être renseignée
    lBoReappliquer := False ;
    lText          := '' ;

    lRect    := FListe.CellRect(ACol,ARow) ;

    // Cas de grisage suivant les colonnes
    if lStCol = 'E_DEBITDEV' then
      lBoGrise := ( FTobPiece.GetValue( ARow , 'E_CREDITDEV') <> 0 ) and  ( ARow > 0 )
    else if lStCol = 'E_CREDITDEV' then
      lBoGrise := ( FTobPiece.GetValue( ARow , 'E_DEBITDEV')  <> 0 ) and  ( ARow > 0 )
    else if ( lStCol = 'E_AUXILIAIRE' ) and FTobPiece.EstRemplit( ARow ) then
      begin
      lBoGrise := FTobPiece.Info.LoadCompte( FTobPiece.GetValue( ARow , 'E_GENERAL') ) and
                  ( FTobPiece.Compte_GetValue('G_COLLECTIF') <> 'X' ) ;
      end
    else lBoGrise := False ;

    // Modification du TEXTE
    if pos( 'SECTION', lStCol) > 0 then
      begin
      lTobEcr := Piece.GetTob( ARow ) ;
      if (lTobEcr.GetString('E_ANA')='X') and (lTobEcr.Detail.count > 0) then
        begin
        lInAxe := ValeurI( Copy( lStCol, length(lStCol), 1 ) ) ;
        if lTobEcr.Detail[ lInAxe - 1 ].Detail.count > 1 then
          begin
          Canvas.Font.Color := clGraytext ;
          lText             := TraduireMemoire('<< Ventilation... >>');
          lBoReappliquer    := True ;
          end ;
        end ;
      end
    // Cas de grisage suivant les colonnes
    else if (( lStCol = 'E_DEBITDEV' ) or ( lStCol = 'E_CREDITDEV' )) and (Piece.ModeSaisie=msBOR) then
      begin
      lTobEcr := Piece.CGetTob( ARow ) ;
      if (Piece.GetEnteteS('E_DEVISE')<>lTobEcr.GetString('E_DEVISE')) then
        begin
        if ( lStCol = 'E_DEBITDEV' )
          then lStchp := 'E_DEBIT'
          else lStchp := 'E_CREDIT' ;
        lText  := StrFMontant( lTobEcr.GetDouble(lStChp), 15, V_PGI.OkDecV, '', True ) ;
        Canvas.Font.Color := clNavy ;
        lBoReappliquer := True ;
        end ;
      end
    else
    //Affichage du libellé pour les champs combo :
    if ( FListe.Cells[ACol, ARow] <> '' ) and ( FTabColAvecLib[ ACol ] ) and ( GetChampDataType( lStCol ) <> '' ) then
      begin
      lValue := FListe.Cells[ACol, ARow] ;
      if lValue <> '' then
        begin
        lText := RechDom( GetChampDataType( lStCol ), lValue, False ) ;
        if (lText = '') or (lText = 'Error')
          then lText          := ''
          else lBoReappliquer := True ;
        end ;
      end ;

    // Gestion multieche
    if FTobPiece.Detail[ ARow - 1 ].GetInteger('E_NUMECHE') > 1 then
      begin
      if (lStCol = 'E_NUMLIGNE') or (lStCol = 'E_GENERAL') or (lStCol = 'E_AUXILIAIRE')  then
        begin
        lText := '' ;
        lBoReappliquer := True ;
        end ;
      end ;


    if lBoReappliquer then
      begin
      Canvas.FillRect( lRect );
      case FListe.FColAligns[ACol] of
        taRightJustify :  lInPosX := lRect.Right - Canvas.TextWidth(lText) - 3 ;
        taCenter       :  lInPosX := lRect.Left  + ((lRect.Right - lRect.Left - canvas.TextWidth( lText ) ) div 2 ) ;
        else              lInPosX := (lRect.Left + 3 ) ;
        end ;
      Canvas.TextRect( lRect, lInPosX, (lRect.Top+lRect.Bottom) div 2 -5 , lText );
      end ;

    if lBoGrise then
      begin
      FListe.PostDrawCell  := nil ;          // on debranche l'évènement lors du dessin de la grille
      Canvas.TextRect(lRect, lRect.Left, lRect.Top, '');
      SetGridGrise(ACol, ARow, FListe) ;     // uLibWindows...
      FListe.PostDrawCell  := PostDrawCell ;
      end;

  end ; // if not Piece.isOut( ARow ) then...


  // Gestion de la ligne séparatrice
  if (piece.ModeSaisie = msBor) then
    begin
    lBoHaut := False ;
    lBoSep  := CalculSeparateur ( ACol, ARow , lBoHaut ) ;
    if lBoSep then
      begin
      FListe.PostDrawCell  := nil ; // on debranche l'évènement lors du dessin de la grille
      CSetGridSep(ACol, ARow, FListe, Canvas , lBoHaut ) ;
      FListe.PostDrawCell  := PostDrawCell ;
      end ;
    end ;

  // Cut-off
  if not Piece.isOut( ARow ) and ((lStCol = 'E_NUMLIGNE') or (lStCol = 'E_GENERAL')) then
    begin
    lTobEcr   := Piece.Detail[ ARow-1 ] ;
    lTobCompl := CGetTOBCompl( lTobEcr ) ;
    if (lTobCompl<>nil) then
      begin
      if (lStCol = 'E_NUMLIGNE') and ( lTobCompl.GetDateTime('EC_CUTOFFDEB') > iDate1900 ) then
        begin
        Canvas.Brush.Color := clBlue ;
        Canvas.Brush.Style := bsSolid ;
        Canvas.Pen.Color   := clBlue ;
        Canvas.Pen.Mode    := pmCopy ;
        Canvas.Pen.Width   := 1 ;
        T1.X               := lRect.Right-5 ;
        T1.Y               := lRect.Top ;//+1 ;
        T2.X               := T1.X+4       ;
        T2.Y               := T1.Y ;
        T3.X               := T2.X         ;
        T3.Y               := T2.Y+4 ;
        Canvas.Polygon( [ T1, T2, T3 ] ) ;
        end
      else if (lStCol = 'E_GENERAL') and ( Length( lTOBCompl.GetValue('EC_DOCGUID') ) > 1 ) then
        begin
        Canvas.Brush.Color := clGreen ;
        Canvas.Brush.Style := bsSolid ;
        Canvas.Pen.Color   := clGreen ;
        Canvas.Pen.Mode    := pmCopy ;
        Canvas.Pen.Width   := 1 ;
        T1.X               := lRect.BottomRight.X - 1 ;
        T1.Y               := lRect.BottomRight.y - 1 ;
        T2.X               := T1.X ;   //lRect.Right-5 ;
        T2.Y               := T1.Y - 5 ;//+1 ;
        T3.X               := T1.X - 5   ;
        T3.Y               := T1.Y ;
        Canvas.Polygon( [ T1, T2, T3 ] ) ;
        end ;
      end ;
    end ;   // Fin Cut-off


  // commentaire
  if not Piece.isOut( ARow ) and (lStCol = 'E_GENERAL') then
    begin
    lTobEcr   := Piece.Detail[ ARow-1 ] ;
    lStBlocNote := lTobEcr.GetString('E_BLOCNOTE') ;
    if ( Trim(lStBlocNote) <> '' ) then
      begin
      lRect              := FListe.CellRect(ACol,ARow) ;
      Canvas.Brush.Color := clRed ;
      Canvas.Brush.Style := bsSolid ;
      Canvas.Pen.Color   := clRed ;
      Canvas.Pen.Mode    := pmCopy ;
      Canvas.Pen.Width   := 1 ;
      T1.X               := lRect.Right-5 ;
      T1.Y               := lRect.Top ; //+1 ;
      T2.X               := T1.X+4       ;
      T2.Y               := T1.Y ;
      T3.X               := T2.X         ;
      T3.Y               := T2.Y+4 ;
      Canvas.Polygon([T1,T2,T3]) ;
      end ;
    end ; // Fin Commentaire

end;

function TSaisiePiece.GetColName( vCol : Integer = -1 ) : String;
begin
  if vCol < 0
    then result := FListe.ColNames[ FListe.Col ]
    else if vCol < FListe.ColCount
         then result := FListe.ColNames[ vcol ] ;
end;

procedure TSaisiePiece.ElipsisClick(Sender: TObject);
var lTobLigne  : TOB ;
    lStColName : String ;
    lCancel    : Boolean ;
    lChg       : Boolean ;
    lRechRow   : Integer ;
    lRechCol   : Integer ;
    lRechOk    : Boolean ;
begin

  if ( csDestroying in Ecran.ComponentState) then Exit ;

  // col éditable ?
  lstColName := GetColName ;
  if not EstColEditable( Row, lStcolName ) then Exit ;

{$IFDEF SBO}
AddEvenement('TSP_ElipsisClick vCol:='+intToStr(FListe.Col)+' vRow:='+intToStr(FListe.Row)) ;
{$ENDIF}

  // Tests avant traitement sur tob courante
  lTobLigne := FTobPiece.GetTob( FListe.Row ) ;
  if lTobLigne = nil then Exit ;

  // Execution recherche des comptes
  if (lStColName = 'E_GENERAL') or (lStColName = 'E_AUXILIAIRE') then
    begin
    RechCompte.Ecr := lTobLigne ;
    lRechRow := FListe.Row ;
    lRechCol := FListe.Col ;
    if RechCompte.ElipsisClick( Sender ) then
      begin
      lCancel  := False ;
      lChg     := False ;

      ValideCellule( lstColName, lRechRow ) ;

      if FListe.row <> lRechRow then
        RowEnter    ( Sender, lRechRow, lCancel, lChg ) ;
      if (not lCancel) and ( Col <> lRechCol ) then
        CellEnter    ( Sender, lRechCol, lRechRow , lCancel );
      if lCancel then
        begin
        FListe.Row := lRechRow ;
        FListe.Col := lRechCol ;
        end ;

      end ;
    end
  // Autre lookup
  else
    begin
    lRechOk := RecherchePourChamps( lStColName ) ;
    ValideCellule( lStColName, Row ) ;
   // passage automatique à la colonne suivante
    if lRechOk then
      begin
      PostMessage( FListe.Handle, WM_KEYDOWN, VK_TAB, 0) ;
      end ;
    end ;


end;

procedure TSaisiePiece.OnError(sender: TObject; Error: TRecError);
begin
 if trim( Error.RC_Message )<>''
   then PGIInfo( Error.RC_Message, FEcran.Caption )
   else if ( Error.RC_Error <> RC_PASERREUR )
        then FMessageCompta.Execute( Error.RC_Error )
end;

procedure TSaisiePiece.InitRechCompte;
var lInfo : TInfoEcriture ;
begin
  // initialisation de l'objet de recherche ( avant le readFolio qui l'utilise )
  lInfo := Info ;
  FRechCompte         := TRechercheGrille.Create( lInfo ) ;
  FRechCompte.OnError := OnError ;
  FRechCompte.COL_GEN := GetColIndex( 'E_GENERAL' ) ;
  FRechCompte.COL_AUX := GetColIndex( 'E_AUXILIAIRE' ) ;
  FRechCompte.OnChangeGeneral := OnChangeGeneral ;
  FRechCompte.OnChangeAux     := OnChangeAuxi ;

  // intialisation de l'objet de recherche des libellés auto
  FZlibAuto := TZLibAuto.Create ;
  FZlibAuto.Load ;

end;

function TSaisiePiece.GetColIndex(vStColName: String): Integer;
var i : Integer ;
begin
  result := -1 ;
  for i := 0 to FListe.ColCount - 1 do
    begin
    if FListe.ColNames[i] = vStColName then
      begin
      result := i ;
      Exit ;
      end ;
    end ;
end;

procedure TSaisiePiece.InitMessageCompta;
begin
  FMessageCompta := TMessageCompta.Create( FEcran.Caption, msgSaisiePiece ) ;
  FTobPiece.OnError := OnError ;
end;

procedure TSaisiePiece.CanClose(var CanClose: Boolean);
begin
 if CanClose and Piece.ModifEnCours then
   CanClose := PGIAskCancel( FMessageCompta.GetMessage( RC_ABANDON ), FEcran.Caption ) = mrYes
end;


procedure TSaisiePiece.InitSaisie;
begin
  // Libération de l'éventuelle pièce en cours
  FTobPiece.InitSaisie ;
  EditEnabled( True ) ;
  FListe.Row := 1 ;
//  FBoFirstEnter := True ;
  // 1ère ligne vierge
//  InsertRow( 1, True ) ;
end;

procedure TSaisiePiece.Recherche;
begin
  FirstFind := true ;
  FFindDialog.Execute ;
end;

procedure TSaisiePiece.GridEnter(Sender: TObject);
begin

  if ( csDestroying in Ecran.ComponentState) then Exit ;
  if not FBoEvtOn then Exit ;
{  if not FBoFirstEnter then
    begin
    if (GoEditing in FListe.Options) then
      LaGrille.MontreEdit;
    Exit ;
    end ;
}
{$IFDEF SBO}
AddEvenement('TSP_GridEnter');
{$ENDIF}

  // Appel cellenter client
  if Assigned( FUserGridEnter ) then
    FUserGridEnter( Sender );

  // On doit être sur d'avoir vider la pile d'evts
  Application.ProcessMessages ;

  // placement initiale
  SetInitCell ;

  // propriété d'édition de la ligne de la grille
  GereOptionsGrid ;

  if GetActionFiche <> taConsult then
    begin
    PostMessage(LaGrille.Handle, WM_KEYDOWN, VK_TAB,  0) ;
    PostMessage(LaGrille.Handle, WM_KEYDOWN, VK_LEFT, 0) ;
    end ;

//  FBoFirstEnter := False ;

end;

procedure TSaisiePiece.GridExit(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  if not FBoEvtOn then Exit ;

{$IFDEF SBO}
AddEvenement('TSP_GridExit' ) ;
{$ENDIF}
{
  if (GoEditing in LaGrille.Options)
    then FBoFirstEnter := True ;
}
end;

procedure TSaisiePiece.MouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;

{$IFDEF SBO}
AddEvenement('TSP_MouseDown     >> FBoEvtOn : ' + BoolToStr_( FBoEvtOn ) ) ;
{$ENDIF}

  if not FBoEvtOn then
    SetEvtOn ;
end;

procedure TSaisiePiece.videLignes;
begin
  if GetActionFiche = taConsult then Exit ;
  if not Piece.PieceModifiable then Exit ;

  Piece.videPiece ;

  if Piece.Detail.Count = 0
    then InsertRow( 1, True )
    else AfficheLignes ;

  SetPos( 1, 1 ) ;

end;

procedure TSaisiePiece.ProratisePiece;
var lNewMontant   : Double ;
    lOldMontant   : Double ;
    lStChamps     : String ;
begin
  if GetActionFiche = taConsult then Exit ;
  lStChamps := GetColName ;
  if not ( (lStChamps='E_DEBITDEV') or (lStChamps='E_CREDITDEV') ) then Exit ;
  if Not ExJaiLeDroitConcept(TConcept(ccSaisProrata),True) then Exit ;

  // Recup champ + valeur
  lNewMontant := GetValueGrid( FListe.col, FListe.row ) ;
  lOldMontant := Piece.GetValue( FListe.row , lStChamps ) ;

  if (lOldMontant = 0 ) or (lOldMontant = lNewMontant) or (lNewMontant = 0) then Exit ;

  Piece.ProratisePiece( FListe.Row , lStChamps, lNewMontant ) ;

  AfficheLignes ;

end;

procedure TSaisiePiece.GereNouvelleLigne(vFrom, vIn: Integer);
begin
  if vIn <= vFrom then Exit ;
//  if not FTobPiece.EstRemplit( vFrom ) then Exit ;
  if ( vIn > 1 ) and ( FTobPiece.GetTob( vIn -1 ) <> nil )
                 and not FTobPiece.EstRemplit( vIn -1 ) then Exit ;
  if FTobPiece.GetTob( vIn ) = nil then
    InsertRow( vIn, True ) ;
end;

function TSaisiePiece.GetCol: Integer;
begin
  Result := FListe.Col ;
end;

procedure TSaisiePiece.SetCol(vIndex: Integer);
begin
  FListe.Col := vIndex ;
end;


procedure TSaisiePiece.DblClick(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  // Appel ancêtre
  if Assigned( FInheritedDblClick ) then
    FInheritedDblClick( Sender ) ;

  // Appel DblClick client
  if Assigned( FUserDblClick ) then
    FUserDblClick( Sender );
end;

function TSaisiePiece.EstColEditable( vInRow : integer ; vStColName: String ): Boolean;
var
  Condition : String;
begin

  result := False ;

  // Si on est en sélection à la ligne, pas de test à faire
  if (GoRowSelect in FListe.Options) then Exit ;

  // Paramétrage grille
  if not FListe.ColEditables[ GetColIndex( vStColName ) ] then Exit ;

  // En attendant le paramétrage de la devise à la ligne
//  if ( vStColName = 'E_DEVISE' ) then
//    Exit ;

  // Cas du multi-echéances
  if Piece.GetInteger( vInRow, 'E_NUMECHE') > 1 then
    begin
    // Sur un détail multi-échéance, seule quelques champs sont éditables
    if pos(vStColName, 'E_MODEPAIE;E_DATEVALEUR;E_DATEECHEANCE;E_DEBITDEV;E_CREDITDEV')<=0 then
      Exit ;
    end ;

  // Cas spécifique au mode pièce
  if ( Piece.ModeSaisie = msPiece) then
    begin
    // saisie multi-établissement :
    if ( vStColName = 'E_ETABLISSEMENT' ) then
      begin
      result := (EtabForce = '') and ( GetActionFiche = taCreat ) ;
      exit ;
      end
    // Champs modifiable à la ligne uniquement en mode bordereau
    else if {( vStColName = 'E_DEVISE' ) or} ( vStColName = 'E_NATUREPIECE') or (vStColName = 'E_DATECOMPTABLE' ) then
      exit ;
    end
  else if ( Piece.ModeSaisie = msBor) then
    begin
    // montant en devise <> entête
    if ( vStColName = 'E_DEBITDEV' ) or ( vStColName = 'E_CREDITDEV' ) then
      begin
      if Piece.GetValue( vInRow, 'E_DEVISE') <> Piece.GetEnteteS('E_DEVISE') then
        Exit ;
      end
    // Qte1
    else if Piece.Contexte.CPPCLSAISIEQTE then
      begin
      if ( vStColName = 'E_QTE1' ) then
        begin
        if not Piece.Info.LoadCompte( Piece.GetValue( vInRow , 'E_GENERAL') )
             or ( Piece.Compte_GetValue('G_QUALIFQTE1') = '' )
             or ( Piece.Compte_GetValue('G_QUALIFQTE1') = 'AUC' ) then Exit ;
        end
      // Qte2
      else if ( vStColName = 'E_QTE2' ) then
        begin
        if not Piece.Info.LoadCompte( Piece.GetValue( vInRow , 'E_GENERAL') )
             or ( Piece.Compte_GetValue('G_QUALIFQTE2') = '' )
             or ( Piece.Compte_GetValue('G_QUALIFQTE2') = 'AUC' ) then Exit ;
        end
      else if ( vStColName = 'E_QUALIFQTE1' ) or ( vStColName = 'E_QUALIFQTE2' ) then
        Exit ;
      end ;
    end
  else if ( Piece.ModeSaisie = msLibre) then
    begin
    end ;

  // --------------------------------
  // ----- GESTION DES MONTANTS -----
  // --------------------------------
  // ----- DEBIT -----
  if ( vStColName = 'E_DEBITDEV' ) then
    begin
    // ---> interdit de saisir un débit et un crédit
    if ( Piece.GetValue( vInRow , 'E_CREDITDEV') <> 0 ) then
      Exit ;
    end
  // ----- CREDIT -----
  else if ( vStColName = 'E_CREDITDEV') then
    begin
    // ---> interdit de saisir un débit et un crédit
    if ( Piece.GetValue( vInRow , 'E_DEBITDEV') <> 0 ) then
      Exit ;
    end
  // ------ Données d'échéances : MODE DE PAIEMENT / DATE ECHEANCE / DATE VALEUR -----
  else if ( vStColName = 'E_MODEPAIE') or ( vStColName = 'E_DATEECHEANCE') or ( vStColName = 'E_DATEVALEUR') then
    begin
    if ( Piece.GetValue( vInRow , 'E_ECHE') <> 'X' ) and not Piece.EstDivLett( vInRow ) then
      Exit ;
    end

  // ---------------------------------------------------------
  // ----- SAISIE AUXI SUR COMPTES COLLECTIFS UNIQUEMENT -----
  // ---------------------------------------------------------
  else if vStColName = 'E_AUXILIAIRE' then
    if Piece.Info.LoadCompte( Piece.GetValue( vInRow , 'E_GENERAL') )
         and not ( Piece.Compte_GetValue('G_COLLECTIF') = 'X' ) then Exit ;

  // ------------------------------------------------
  // --------- cas fonctionnel sur la ligne ---------
  // ------------------------------------------------
  if not Piece.EstChampsModifiable ( vInRow, vStColName ) then Exit ;

  
  // ------------------------------------------------
  // ------------------ Condition -------------------
  // ------------------------------------------------
  Condition := RecupSpecif(vStColName,'SAISIE','CONDITION');
  if Condition <> '' then
  begin
     FGrille   := 'SAISIE';
     FTob      := Piece.GetTob(GetRow);
     if not GFormuleTest(Condition, GetFormule, nil, 1 ) then
     begin
        Condition := GFormule(Condition, GetFormule, nil, 1 );
        if not(Condition = 'X') then
           Exit;
     end
  end;

  // ----------------------------------
  // --------- MODIF POSSIBLE ---------
  // ----------------------------------
  result := True ;

end;

function TSaisiePiece.EstColLookup(vCol: Integer): Boolean;
var lstColName : String ;
begin
  if vCol < 0
   then lStColName := GetColName( FListe.Col )
   else lStColName := GetColName( vCol ) ;

  result := (lStColName = 'E_GENERAL')    or
            (lStColName = 'E_AUXILIAIRE') or
//            (lStColName = 'E_TABLE0')     or
//            (lStColName = 'E_TABLE1')     or
//            (lStColName = 'E_TABLE2')     or
//            (lStColName = 'E_TABLE3')     or
            (lStColName = 'E_DEVISE')     or
            (lStColName = 'E_LIBREDATE' ) or
            (lStColName = 'E_DATEREFEXTERNE') or
            (lStColName = 'E_BANQUEPREVI') or

            (lStColName = 'SECTIONA1') or
            (lStColName = 'SECTIONA2') or
            (lStColName = 'SECTIONA3') or
            (lStColName = 'SECTIONA4') or
            (lStColName = 'SECTIONA5') or

            EstColCombo( vCol ) ;

end;

function TSaisiePiece.EstColCombo(vCol: Integer): Boolean;
var lstColName : String ;
begin
  if vCol < 0
   then lStColName := GetColName( FListe.Col )
   else lStColName := GetColName( vCol ) ;

  result := (lStColName = 'E_QUALIFQTE1')     or
            (lStColName = 'E_MODEPAIE')       or
            (lStColName = 'E_ETABLISSEMENT')  or
            (lStColName = 'E_NATUREPIECE')    or
            (lStColName = 'E_QUALIFQTE2')     or
            (lStColName = 'E_TABLE0')         or
            (lStColName = 'E_TABLE1')         or
            (lStColName = 'E_TABLE2')         or
            (lStColName = 'E_TABLE3') ;

end;

function TSaisiePiece.GetChampDataType(vStColName: String): String;
begin
  if vStColName = 'E_QUALIFQTE1' then
    result := 'TTQUALUNITMESURE'
  else if vStColName = 'E_QUALIFQTE2' then
    result := 'TTQUALUNITMESURE'
  else if vStColName = 'E_MODEPAIE' then
    result := 'TTMODEPAIE'
  else if vStColName = 'E_QUALIFPIECE' then
    result := 'TTQUALPIECE'
  else if vStColName = 'E_ETABLISSEMENT' then
    result  := 'TTETABLISSEMENT'
  else if vStColName = 'E_REGIMETVA' then
    result := 'TTREGIMETVA'
  else if vStColName = 'E_DEVISE' then
    result := 'TTDEVISE'
  else if vStColName = 'E_EXERCICE' then
    result := 'TTEXERCICE'
  else if vStColName = 'E_JOURNAL' then
    result := 'TTJOURNAUX'
  else if vStColName = 'E_NATUREPIECE' then
    result := 'TTNATUREPIECE'
  else if vStColName = 'E_TABLE0' then
    result := 'TZNATECR0'
  else if vStColName = 'E_TABLE1' then
    result := 'TZNATECR1'
  else if vStColName = 'E_TABLE2' then
    result := 'TZNATECR2'
  else if vStColName = 'E_TABLE3' then
    result := 'TZNATECR3'
  else result := '' ;
end;


function TSaisiePiece.GetFullSize( vGrille : THGrid ) : Integer;
var //lStColSize : String ;
    lTobGrille : Tob ;
    i          : integer ;
begin

  result     := 0 ;

  if vGrille = FGrilleCompl
    then lTobGrille := FTobMasque.FindFirst( ['TYPE'], ['COMPLEMENTS'], True )
    else lTobGrille := FTobMasque.FindFirst( ['TYPE'], ['SAISIE'], True ) ;

{  lStColSize := FGridColSize ;

  if Assigned( FTobMasque ) then
    while lStColSize <> '' do
      result := result + ( ReadTokenI( lStColSize ) )
  else // Cas par les listes
    while lStColSize <> '' do
      result := result + ( ReadTokenI( lStColSize ) * 8 ) ;}

  for i:=0 to lTobGrille.Detail.count - 1 do
    result := result + lTobGrille.Detail[i].GetInteger('LARGEUR') ;

end;

procedure TSaisiePiece.GetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var lStCol            : String ;
begin
  if csDestroying in Ecran.ComponentState then Exit ;

  // gestion des cases debit et crédit - une seul des deux doit être renseignée
  lStCol   := GetColName( ACol ) ;

  // Font de numéro de ligne pas assez claire
  if lStCol = 'E_NUMLIGNE' then
    begin
    Canvas.Font.Color  := clWindowText ;
    Canvas.Brush.Color := FListe.FixedColor ;
    if Piece.EstMultiEche( ARow ) then
      Canvas.Brush.Color := clGrayText ; //clLtGray ;//clSilver ;
    end ;

end;

procedure TSaisiePiece.SetFocus;
begin
{$IFDEF SBO}
AddEvenement('SetFocus' ) ;
{$ENDIF}

  if not Enabled then
    Enabled := True ;

  if FListe.CanFocus then
    begin
    FListe.SetFocus ;
    GridEnter(nil);
    end ;

end;


function TSaisiePiece.ValideCellule( vChamp : String ; vRow : Integer ; vBoRefresh : boolean ) : boolean ;
var  lValGrid  : Variant ;
     lValTob   : Variant ;
     lTobLigne : TOB ;
begin
  lTobLigne := FTobPiece.GetTob( vRow ) ;

  lValTob   := lTobLigne.GetValue( vChamp ) ;                   // Ancienne valeur
  lValGrid  := GetValueGrid( GetColIndex( vChamp ) , vRow ) ;   // Nouvelle valeur

  result := lValGrid <> lValTob ;
  if result then
    begin
    FTobPiece.PutValue( vRow , vChamp, lValGrid ) ;

    if vBoRefresh then
      if (Piece.ModeSaisie=msBor) and Piece.EstChampsGroupe( vChamp )
        then AfficheGroupe( vRow )
        else AfficheLignes( vRow ) ;
    end ;
end;

procedure TSaisiePiece.SetColCount(vStChamps: String);
var lInColCount  : Integer ;
    lStListeCol  : String ;
begin
  lInColCount := 0 ;
  lStListeCol := vStChamps ;

  While ReadTokenSt(lStListeCol)<>'' do
    Inc( lInColCount ) ;

  if FListe.DbIndicator then
    Inc ( lInColCount ) ;

  if FListe.ColCount <> lInColCount then
    FListe.ColCount := lInColCount ;

end;

procedure TSaisiePiece.KeyPress(Sender: TObject; var Key: Char);
var lStColName : String ;
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;

{$IFDEF SBO}
AddEvenement('TSP_KeyPress     >> FBoEvtOn : ' + BoolToStr_( FBoEvtOn ) ) ;
{$ENDIF}

  if not FBoEvtOn then
    SetEvtOn ;

  lstColName := GetColName ;
  if pos( 'DATE', lStColName) > 0 then
    ParamDate( FEcran, Sender, Key ) ;
  if FListe.SynEnabled then
    begin
    if (lstColName='E_NATUREPIECE') then
      if (Key=' ') then Key:=#0 ;
    end ;

end;

procedure TSaisiePiece.GetParamLibre( vStChamp: String ; var vStLibelle : String ; var vBoVisible : Boolean ) ;
begin
  if not Assigned ( FTobParamLib ) then Exit ;
  CGetParamLibre( FTobParamLib , vStChamp, vStLibelle, vBoVisible ) ;
end;

procedure TSaisiePiece.InitParamLib;
var lQParam : TQuery ;
begin
  FTobParamLib := TOB.Create('V_PERSOLIB', nil, -1);
  lQParam := OpenSelect('SELECT * FROM PARAMLIB WHERE PL_TABLE="E"', Dossier);
  if not lQParam.Eof then
    FTobParamLib.loadDetailDB('PARAMLIB','','',lQParam,False);
  Ferme(lQParam) ;
end;

{
function TSaisiePiece.GetFormatCombo(vStColName: String): String;
var lStPlus     : String ;
    lStVide     : String ;
    lStDataType : String ;
begin
  // par défaut :
  result := '' ;

  // tablette
  lStDataType := GetChampDataType( vStColName ) ;
  if lStDataType = '' then Exit ;

  // Condition Plus...
  lStPlus := '' ;

  // Choix <<Tous>>...
  lStVide := '<<Aucun>>' ;
  if vStColName = 'E_ETABLISSEMENT' then
    lStVide := '' ;

  // Composition du format de la colonne Combo :
  result := 'CB=' + lStDataType + '|' + lStPlus  + '|' + lStVide ;

end;
}

function TSaisiePiece.RecherchePourChamps( vStColName : String ) : Boolean ;
var lStTitre : String ;
    lBoVisu  : Boolean ;
    lStAxe   : string ;
    lStVal   : string ;
begin
  if pos( 'E_TABLE', vStColName) > 0 then
    begin
    GetParamLibre( vStColName, lStTitre, lBoVisu ) ;
    if not lBoVisu then
      begin
      result := False ;
      Exit ;
      end ;
    if lStTitre = '' then
      lStTitre := TraduireMemoire('Tables libres n°' + copy( vStColName, 8, 1)) ;
    result := LookupList(  FListe,
                 lStTitre,
                 'NATCPTE',
                 'NT_NATURE',
                 'NT_LIBELLE',
                 'NT_TYPECPTE="E0' + copy( vStColName, 8, 1) + '"',
                 'NT_NATURE',
                 true, 0) ;
    end
  // Tables libres
  else if vStColName = 'E_BANQUEPREVI' then
    begin
    result := LookupList(  FListe,
                 TraduireMemoire('Banque prévisionnelle'),
                 'GENERAUX',
                 'G_GENERAL',
                 'G_LIBELLE',
                 'G_NATUREGENE = "BQE" AND G_FERME="-"',
                 'G_GENERAL',
                 true, 1) ;
    end
  else if (vStColName = 'E_QUALIFQTE1') or (vStColName = 'E_QUALIFQTE2') then
    begin
    result := LookupList(  FListe,
                 TraduireMemoire('Qualifiant de quantité'),
                 'CHOIXCOD',
                 'CC_CODE',
                 'CC_LIBELLE',
                 'CC_TYPE="QME"',
                 'CC_CODE',
                 true, 0) ;
    end
  else if vStColName = 'E_MODEPAIE' then
    begin
    result := LookupList(  FListe,
                 TraduireMemoire('Mode de paiement'),
                 'MODEPAIE',
                 'MP_MODEPAIE',
                 'MP_LIBELLE',
                 '',
                 'MP_MODEPAIE',
                 true, 0) ;
    end
  else if vStColName = 'E_DEVISE' then
    begin
    result := LookupList(  FListe,
                 TraduireMemoire('Devise'),
                 'DEVISE',
                 'D_DEVISE',
                 'D_LIBELLE',
                 '',
                 'D_DEVISE',
                 true, 0) ;
    end
  else if vStColName = 'E_NATUREPIECE' then
    begin
    result := LookupList(  FListe,
                 TraduireMemoire('Nature de pièce'),
                 'COMMUN',
                 'CO_CODE',
                 'CO_LIBELLE',
                 Piece.GetWhereNatPiece,
                 'CO_CODE',
                 true, 0) ;
    end
  else if vStColName = 'E_ETABLISSEMENT' then
    begin
    result := LookupList(  FListe,
                 TraduireMemoire('Etablissement'),
                 'ETABLISS',
                 'ET_ETABLISSEMENT',
                 'ET_LIBELLE',
                 '',
                 'ET_ETABLISSEMENT',
                 true, 0) ;
    end
  // Champs Date
  else if pos( 'DATE', vStColName) > 0 then
    begin
    if (vStColName = 'E_DATECOMPTABLE') {and (Piece.ModeSaisie<>msPiece)} then
      begin
      result := False ;
      Exit ;
      end ;
    V_PGI.ParamDateproc( FListe );
    result := HEnt1.IsValidDate( FListe.CellValues[Col,Row] ) ;
    end
  // Champs Spécif SECTION
  else if pos( 'SECTION', vStColName) > 0 then
    begin
    lStAxe := Copy( vStColName, 8, 2 ) ;
    lStVal := FListe.Cells[Col,Row] ;
    if Piece.Info.LoadSection( lStVal, lStAxe) then
      begin
      FListe.Cells[Col,Row] := Piece.Info.StSection ;
      result := True ;
      end
    else
      result := LookupList( FListe,
                   TraduireMemoire('Section'),
                   'SECTION',
                   'S_SECTION',
                   'S_LIBELLE',
                   'S_AXE="'+lStAxe+'" AND S_FERME="-"',
                   'S_SECTION',
                   true, 0) ;
    end
  else result := False ;

end;

function TSaisiePiece.GetDossier: String;
begin
  if Assigned( Piece )
    then result := Piece.Dossier
    else result := V_PGI.SchemaName ;
end;

function TSaisiePiece.GetInfo: TInfoEcriture;
begin
  result := nil ;
  if Assigned(FTobPiece) then
    result := FTobPiece.Info ;
end;

function TSaisiePiece.GetFormulePourCalc(Formule: hstring): Variant;
var lCalcRow : integer ;
    lEcr     : Tob ;
begin
  Result  := #0 ;
  Formule := AnsiUpperCase(Trim(Formule)) ;
  if Formule='' then Exit ;
  if pos('E_DEBIT',Formule)>0
    then lCalcRow := FListe.Row
    else lCalcRow := FListe.Row-1 ;

  if not Piece.EstRemplit( lCalcRow ) then Exit ;
  lEcr := Piece.GetTob( lCalcRow ) ;
  if lEcr=nil then Exit ;
  if Pos('E_', Formule)>0 then
    Result := lEcr.GetValue(Formule) ;
end;

procedure TSaisiePiece.SetFormatMontant;
var lInCol : Integer ;
    lInDec : Integer ;
    lStFormatDev : string ;
begin

  lInDec := 0 ;
  if ( Assigned(Piece) ) and ( Piece.Devise.Code <> '' )
    then lInDec := Piece.Devise.Decimale ;
  lStFormatDev := '#,##0';
  if lInDec > 0
    then lStFormatDev := lStFormatDev + '.' + Copy('000000000000000000000', 1, lInDec) ;

  lInCol := GetColIndex( 'E_CREDITDEV' ) ;
  if lStFormatDev <> FListe.ColFormats[lInCol] then
    begin
    FListe.ColFormats[lInCol] := lStFormatDev ;
    lInCol := GetColIndex( 'E_DEBITDEV' ) ;
    FListe.ColFormats[lInCol] := lStFormatDev ;
    AfficheLignes(0, True) ;
    end ;

end;

procedure TSaisiePiece.CAfficheLigne(vRow: integer; vTob: Tob ; vBoForce : boolean );
var lStListe  : string ;
    lStChp    : string ;
    lInCol    : integer ;
    lStVal    : HString ;
    lTob      : Tob ;
begin

{$IFDEF SBO}
//AddEvenement('  -------> TSP_CAfficheLigne vRow:=' + intToStr(vRow) ) ;
{$ENDIF}

  lStListe := FGridChamps ;
  lInCol   := -1 ;
  if Assigned( vTob )
    then  lTob := vTob
    else if not Piece.IsOut( vRow )
           then lTob := FTobPiece.Detail[ vRow - 1 ]
           else Exit ;

//  if not ( vBoForce or lTob.IsOneModifie(True)) then Exit ;

  // DB indicateur
  if (FListe.DBIndicator) then
    begin
    inc(lInCol);
    FListe.CellValues[lInCol, vRow] := '';
    end;

  while lStListe <> '' do
    begin
    inc(lInCol);
    lStChp := ReadTokenSt(lStListe);
    lStVal := GetValeurPourGrille ( lStChp, lTob, lInCol );
    // MAJ de la cellule
    FListe.CellValues[lInCol, vRow] := lStVal ;
    end;

   // MAJ de l'objet rattaché à la ligne
   Fliste.Objects[0, vRow] := lTob ;

//   lTob.SetAllModifie( False ) ;

end;

procedure TSaisiePiece.SetGridSep(ACol, ARow: Integer; Canvas: TCanvas; bHaut: Boolean);
var R : TRect ;
begin
  Canvas.Brush.Color := clRed ;
  Canvas.Brush.Style := bsSolid ;
  Canvas.Pen.Color   := clRed ;
  Canvas.Pen.Mode    := pmCopy ;
  Canvas.Pen.Style   := psSolid ;
  Canvas.Pen.Width   := 1 ;

  R := FListe.CellRect(ACol, ARow) ;
  if bHaut then
    begin
    Canvas.MoveTo(R.Left, R.Top) ;
    Canvas.LineTo(R.Right+1, R.Top) ;
    end
  else
    begin
    Canvas.MoveTo(R.Left, R.Bottom-1) ;
    Canvas.LineTo(R.Right+1, R.Bottom-1) ;
    end ;

end;

procedure TSaisiePiece.SwapModeSaisie(vMode: TModeSaisie);
var lInCol : integer ;
begin

  Case vMode of
    msPiece : begin
              // saisie multi-établissement :
              lInCol := GetColIndex( 'E_ETABLISSEMENT' ) ;
              if lInCol > 0 then
                begin
                ActiveCol( lInCol, (EtabForce = '') and ( GetActionFiche = taCreat ) ) ;
                Info.AjouteErrIgnoree([RC_ETABLISSEMENT]);
                end ;
              // Date non modifiable + Format différent
              lInCol := GetColIndex( 'E_DATECOMPTABLE' ) ;
              if lInCol > 0 then
                ActiveCol( lInCol, False ) ;
              // Nature non modifiable
              lInCol := GetColIndex( 'E_NATUREPIECE' ) ;
              if lInCol > 0 then
                ActiveCol( lInCol, False ) ;
              end ;

    else begin
         // saisie multi-établissement :
         lInCol := GetColIndex( 'E_ETABLISSEMENT' ) ;
         if lInCol > 0 then
           begin
           ActiveCol( lInCol, False ) ;
           Info.AjouteErrIgnoree([]);
           end ;
         // Date non modifiable + Format différent
         lInCol := GetColIndex( 'E_DATECOMPTABLE' ) ;
         if lInCol > 0 then
           ActiveCol( lInCol, True ) ;
         // Nature non modifiable
         lInCol := GetColIndex( 'E_NATUREPIECE' ) ;
         if lInCol > 0 then
           ActiveCol( lInCol, True ) ;
         end ;
    end ;

  // Raffraichissement de la liste
//  AfficheLignes ;

end;

procedure TSaisiePiece.CAfficheCol( vStNom : string ; vInCol : integer = 0 );
var lInCol    : integer ;
    lStVal    : HString ;
    i         : integer ;
    lTob      : Tob ;
begin
   Fliste.BeginUpdate ;

  if FListe.RowCount < Piece.Detail.count then
    FListe.RowCount := Piece.Detail.count ;

  if vInCol <> 0
      then lInCol := vInCol
      else lInCol := GetColIndex( vStNom ) ;

  for i := 1 to Piece.Detail.count do
    begin
    lTob := Piece.Detail[i-1] ;
    lStVal := GetValeurPourGrille ( vStNom, lTob, lInCol ) ;
    // MAJ de la cellule
    FListe.CellValues[lInCol, i] := lStVal ;
    end ;

  FListe.EndUpdate ;

end;

function TSaisiePiece.GetValeurPourGrille( vStNom: string; vTob : Tob ; vInCol : integer ): string ;
var lInCol    : integer ;
    lStVal    : HString ;
    lValeur   : Variant ;
    lDtDate   : TDateTime ;
    lInYear    : Word ;
    lInMonth   : Word ;
    lInDay     : Word ;
begin

    if vInCol > 0
      then lInCol := vInCol
      else lInCol := GetColIndex( vStNom ) ;

    if lInCol < 0 then Exit ;

    lValeur := GetGoodValeurGrille(vStNom,vTob,'SAISIE');

    if (vStNom = 'E_DATECOMPTABLE') {and (Piece.ModeSaisie<>msPiece) } then
      begin
      lDtDate := VarAsType(lValeur, vardate);
      DecodeDate( lDtdate, lInYear, lInMonth, lInDay ) ;
      lStVal  := IntToStr( lInDay );
      end
    else if ((vStNom = 'E_CREDITDEV') or (vStNom = 'E_DEBITDEV') or
           (vStNom = 'E_CREDIT') or (vStNom = 'E_DEBIT'))
           and ( vTob.GetDouble( vStNom ) = 0 )
        then lStVal := ''
        else lStVal  := CVariantToStGrid( lValeur, FListe, lInCol ) ;

(*
    if V_PGI.TraduireLesDatas and ChampATRaduire(vStNom, FListe.FTradChampSup) then
      TraduireData(lStVal, true);
*)
    result := lStVal ;

end;

function  TSaisiePiece.GetFormule(vStFormule: HString): Variant;
begin
  result := GetGoodValeurGrille(vStFormule,FTob,FGrille) ;
end;

procedure TSaisiePiece.ActiveCol(vInCol: Integer; vBoEditable : boolean );
begin
  if (vInCol < 0) or (vInCol > (FListe.ColCount-1) ) then Exit ;

  FListe.ColEditables[ vInCol ] := vBoEditable ;
  if vBoEditable
    then FListe.ColColors[ vInCol ]    := clWindowtext
    else FListe.ColColors[ vInCol ]    := clGraytext ;
end;

function TSaisiePiece.CalculSeparateur(ACol, ARow: Integer; var vBoHaut: boolean; vBo: boolean): boolean;
var lInNum            : integer ;
    lInNumPrec        : integer ;
    lInNumSuiv        : integer ;
begin

 lInNum        := -1 ;
 lInNumPrec    := -1 ;
 lInNumSuiv    := -1 ;
 vBoHaut       := false ;
 result        := false ;

//if (ARow = Piece.Detail.count) and not (FListe.Row=ARow) then Exit ;

 if not Piece.IsOut( ARow ) then
   lInNum := Piece.GetInteger( ARow, 'E_NUMGROUPEECR') ;

 if not Piece.IsOut( ARow - 1 ) then
   lInNumPrec := Piece.GetInteger( ARow - 1, 'E_NUMGROUPEECR') ;

 if not Piece.IsOut( ARow + 1 ) then
   lInNumSuiv := Piece.GetInteger( ARow + 1, 'E_NUMGROUPEECR') ;

 if vBo then
  begin
   if ( ACol > 0 ) and ( lInNum <> lInNumPrec )  then
     result := true ;
  end
   else
    begin
     if ( ACol > 0 ) and ( lInNum <> lInNumSuiv ) and (FListe.Row <> ARow) then
       result := true ;
     if ( ACol > 0)  and ( lInNum <> lInNumPrec ) and ( FListe.Row = ARow - 1) then
       begin
       result  := true ;
       vBoHaut := true ;
       end ;
     if ( ACol > 0)  and ( ARow > Piece.Detail.count ) and ( FListe.Row = ARow - 2) then
       begin
       if not Piece.IsOut( ARow - 2 ) then
         lInNum := Piece.GetInteger(ARow - 2, 'E_NUMGROUPEECR') ;
       result  := lInNum<>lInNumPrec ;
       vBoHaut := result ;
       end ;

    end ;

end;

procedure TSaisiePiece.GetColSuivante( var Acol, ARow : integer ; vInSens : integer = 0 );
var lInSens     : integer ;
    lBoFin      : boolean ;
    lInCol      : integer ;
    lInRow      : integer ;
    lStColName  : string ;

    procedure _PasseColSuivante ;
      begin
        // Déplacement vers l'avant
        if lInSens > 0 then
          begin
          if (lInCol + 1) < FListe.ColCount then  // Fin de ligne ?
            Inc( lInCol )
          else
            begin
            if ( lInRow + 1 ) < FListe.RowCount then // Fin de la grille ?
              begin
              // passage à la ligne suivante, 1ère colonne
              Inc( lInRow ) ;
              lInCol := 1 ;
              end
            else
              lBoFin := True ; // on s'arrête
            end ;
          end
        // Déplacement vers l'arrière
        else
          begin
          if (lInCol - 1) >= FListe.FixedCols then  // debut de ligne ?
            Dec( lInCol )
          else
            begin
            if ( lInRow - 1 ) >= 1 then // Début de la grille ?
              begin
              // passage à la ligne suivante, 1ère colonne
              Dec( lInRow ) ;
              lInCol := FListe.ColCount - 1 ;
              end
            else
              lBoFin := True ; // on s'arrête
            end ;
          end ;
      end ;
begin

  lBoFin := false ;

  if vInSens <> 0 then
    begin
    lInSens := vInSens ;
    lInCol := ACol ;
    lInRow := ARow ;
    end
  else
    begin
    lInCol := FListe.col ;
    lInRow := FListe.row ;
    // déterminatin du sens...
    lInSens := GetSens( ACol, ARow ) ;
    end ;

  // Désactivation des évènements ??
  _PasseColSuivante ;

  lStColName := GetColName( lInCol ) ;

  // Détermination de la cellule de destination
  While (not lBoFin) and ( not EstColEditable( lInRow, lStColName) ) do
    begin
    _PasseColSuivante ;
    lStColName := GetColName( lInCol ) ;
    end ;

  // résultat
  ARow := lInRow ;
  Acol := lInCol ;

end;

function TSaisiePiece.GetSens(ACol, ARow: integer): integer;
begin
  result := CGetGridSens ( FListe, ACol, Arow ) ;
end;

procedure TSaisiePiece.ResizeGrille;
var 
  HMT          : TComponent ;
begin
  if GetFullSize <= FListe.Width then
  begin
     HMT := TForm( Ecran ).FindComponent('HMTrad');
     if (HMT<>nil) then
        THSystemMenu(HMT).ResizeGridColumns( FListe );
  end;
end;

procedure TSaisiePiece.ProchaineLigne( vRow : integer ; vStColName : string ; vBoEvtOn : boolean ) ;
var lInRow     : integer ;
    lInCol     : integer ;
    lInEch     : integer ;
    lMultiEche : TMultiEche ;
begin
  // On est déjà sur la dernière ligne ?
  if vRow = (Fliste.RowCount-1) then Exit ;

  if GuideActif and not TSGuide.FinGuide then Exit ;

  // détermination de la colonne
  if vStColName <> ''
    then lInCol := GetColIndex( vStColName )
    else begin
         lInCol := GetColIndex( 'E_DATECOMPTABLE' ) ;
         if lInCol < 1 then
           lInCol := GetColIndex('E_GENERAL') ;
         end ;

  // détermination de la ligne
  lInRow := vRow ;
  lInEch := 1 ;
  if Piece.GetString(lInRow, 'E_ECHE')='X' then
    begin
    lMultiEche := Piece.GetMultiEche( lInRow ) ;
    if lMultiEche <> nil then
      lInEch := lMultiEche.NbEche ;
    end ;
  lInRow := lInRow + lInEch ;

  if Piece.IsOut( lInRow ) then
    GereNouvelleLigne( vRow, lInRow) ;

  // Mise ne place
  SetPos( lIncol, lInRow, vBoEvtOn ) ;

end;


procedure TSaisiePiece.AfficheGroupe(vRow: Integer);
var lInDe : integer ;
    lInA  : integer ;
    i     : integer ;
begin

  if Piece.ModeSaisie <> msBOR then Exit ;
  if Piece.IsOut( vRow ) then Exit ;

  Piece.GetBornesGroupe( Piece.GetValue( vRow, 'E_NUMGROUPEECR'), lInDe, lInA ) ;

  if lInDe > 0 then
    begin
    for i := lInDe to lInA do
      AfficheLignes( i ) ;
    FListe.Invalidate;
    end ;


end;

procedure TSaisiePiece.SetEvtOff;
begin
  // Réaffectation des évènements
  FBoEvtOn               := False ;
end;

procedure TSaisiePiece.SetEvtOn;
begin
  // Réaffectation des évènements
  FBoEvtOn               := True ;
end;

procedure TSaisiePiece.SetPos(vCol, vRow: integer ; vBoEvtOn, vBoOnEdit : Boolean ; vInSens : integer );
var lSens      : integer ;
    lInCol     : integer ;
    lInRow     : integer ;
    lStColName : string ;
    lCancel    : boolean ;
    lChg       : boolean ;
    lOldRow    : integer ;
begin

{$IFDEF SBO}
AddEvenement('TSP_SetPos vCol:=' + intToStr(vCol) + 'vRow:=' + intToStr(vRow) ) ;
{$ENDIF}

  if vInSens <> 0 then
    lSens := vInSens
  else
    begin
    lSens := GetSens( vCol, vRow ) ;
    lSens := lSens * (-1) ; // le calcul se fait à l'invers car on a par encore bougé dans la grille...
    end ;

  lOldRow := FListe.row ;
  try

    if not vBoEvtOn then
      SetEvtOff ;

  // calcul ligne courante
  if vRow >= FListe.RowCount
    then lInRow := FListe.RowCount - 1
    else lInRow := vRow ;

  // calcul colonne si édition possible
  if (GoEditing in FListe.Options) then
    begin
    // Détermination de la prochaine cellule valide
    if vCol >= FListe.ColCount
      then lInCol := FListe.ColCount - 1
      else if vCol < 1
       then lInCol := 1
       else lInCol := vCol ;
    lStColName := GetColName( lInCol ) ;
    if vBoOnEdit and not EstColEditable( lInRow, lStColName ) then
      GetColSuivante( lInCol, lInRow, lSens ) ;
    FListe.Col := lInCol ;
    end ;

  // placement
  Fliste.Row := lInRow ;

  // Faire apparaître les boutons ellipsis bouton sur les champs approprié
  FListe.ElipsisButton := EstColLookup and not Piece.IsOut( lInRow );

  finally

//    if not vBoEvtOn then
//      SetEvtOn
//      ;

    if Assigned( FGrilleCompl ) then
      AfficheComplement( lInRow ) ;

    SetEvtOn ;

    // Changement de ligne : déclenchements des évènements
    if (lOldRow <> lInRow) then //and vBoEvtOn then //if Assigned( FUserRowEnter ) then
      begin
      lCancel    := False ;
      lChg       := False ;
      RowEnter( FListe, lInRow, lCancel, lChg ) ;
      //FUserRowEnter( nil, lInRow, lCancel, lChg );
      end ;

  end ;

end;

function TSaisiePiece.ValideSortieLigne( vBoCellule : boolean ) : boolean;
var Cancel : boolean ;
    Chg    : boolean ;
    ACol   : integer ;
    ARow   : integer ;
begin

{$IFDEF SBO}
AddEvenement('TSP_ValideSortieLigne');
{$ENDIF}

  SetEvtOn ;

  // Gestion des evènements : Sortie de cellule
  ACol   := FListe.Col ;
  ARow   := FListe.Row ;
  Cancel := False  ;
  if vBoCellule then
    CellExit( FListe, ACol, ARow, Cancel ) ;

  // Gestion des evènements : Sortie de ligne
  if not Cancel then
    begin
    SetEvtOn ;
    Chg    := False ;
    RowExit( FListe, FListe.Row, Cancel, Chg ) ;
    end ;

  // sortie ok si pas d'annulation dans les EVts
  result := not Cancel ;

end;

{$IFDEF SBO}
procedure TSaisiePiece.AddEvenement(value: string);
begin
  Debug( Value ) ;
end;
{$ENDIF SBO}

procedure TSaisiePiece.UpdateMasqueSaisie;
Var i            : Integer ;
    n            : Integer;
    lStNomChamp  : String ;
    lStType      : String ;
    lStNomF      : String ;
    lInDec       : integer ;
    lBoVisu      : boolean;
    lStFormat    : String ;
    lStTitreCol  : String ;
    lStFormatDev : string ;
    // NEW
    lTobGrille   : Tob ;
    lTobCol      : Tob ;
begin

  if not Assigned( FTobMasque ) then
    begin
    FTobMasque := GetTobMasque( True ) ;
    if not assigned( FTobMasque ) then Exit ;
    end ;

  lTobGrille := FTobMasque.FindFirst( ['TYPE'], ['SAISIE'], True ) ;
  if not Assigned(lTobGrille) then Exit ;

  // Calcul su nombre de colonne :
  FListe.ColCount := lTobGrille.Detail.count + FListe.FixedCols ;

  // init variables de gestion
  FGridChamps       := '' ;
  FGridColSize      := '' ;
  SetLength( FTabColAvecLib, FListe.ColCount + 1 ) ;

  if FListe.DBIndicator then n := 1
                        else n := 0;

  // Parcours des lignes pour formatage spécifique des colonnes
  for i := n to lTobGrille.Detail.count - 1 + n do
    begin
    lTobCol := lTobGrille.Detail[ i-n ];

    lStNomChamp := lTobCol.GetString('NOM') ;
    lStType     := lTobCol.GetString('TYPE') ;

    // variable gestion
    FGridChamps     := FGridChamps + lStNomChamp + ';' ;
    FGridColSize    := FGridColSize + lTobCol.GetString('LARGEUR') + ';' ;


    // Nom des colonnes
    FListe.ColNames[ i ] := lStNomChamp ;

    // Affichage du libellé complet :
    FTabColAvecLib[ i ] := lTobCol.GetString('AVECLIBELLE')='X' ;

    // Colonne éditable ?
    ActiveCol( i , CEstColEditable( lStNomChamp ) and (lTobCol.GetString('READONLY')='-') ) ;

    // Alignements
    // Numéro de ligne
    if (lStNomChamp = 'E_NUMLIGNE') or (lTobCol.GetString('TYPE') = 'DATE')
      then FListe.ColAligns[i]  := taCenter
    // Entiers / Réels
    else if (lTobCol.GetString('TYPE') = 'INTEGER') or (lTobCol.GetString('TYPE') = 'DOUBLE')
      then FListe.ColAligns[i]  := taRightJustify
    // Text
    else if (lTobCol.GetString('TYPE') = 'COMBO') and (not FTabColAvecLib[ i ])
      then FListe.ColAligns[i]  := taCenter
    // Tous les autres cas
    else FListe.ColAligns[i]  := taLeftJustify ;

    // Titres + Perso table libre
    lStTitreCol := lTobCol.GetString('TITRE') ;
    if ( pos('LIBRE', lStNomChamp ) > 0 ) or ( pos('E_TABLE', lStNomChamp ) > 0 ) then
      begin
      CGetParamLibre( FTobParamLib, lStNomChamp, lStTitreCol, lBoVisu ) ;
      if FListe.ColEditables[ i ] and not lBoVisu then
        ActiveCol( i, False ) ;
      end ;
    FListe.Cells[i,0]    := lStTitreCol ;

    // largeur des colonnes
    FListe.ColWidths[i]  := lTobCol.GetInteger('LARGEUR') ;

    // Formats des colonnes...
    if ( lStType = 'BLOB' ) then
      begin
      FListe.ColTypes[i]   := 'M';
      FListe.ColFormats[i] := IntToStr( Byte( msBook ) );
      end
    else if ( lStType = 'BOOLEAN' ) then
      begin
      FListe.ColTypes[i]       := 'B';
      FListe.ColFormats[i]     := IntToStr( Byte( csCheckBox ) );
      ActiveCol( i , False ) ;
      end
    else if ( lStType = 'DATE' ) then
      begin
      if (lStNomChamp = 'E_DATECOMPTABLE') then
        begin
        FListe.ColTypes[i]     := 'I';
        FListe.ColLengths[ i ] := 2 ;
        FListe.ColFormats[i]   := '' ;
        end
      else
        begin
        FListe.ColTypes[i]   := 'D';
        FListe.ColFormats[i] := '';  //sinon effet de bord // ??
        end
      end
    else if Copy( Trim( lStNomChamp ), 1, 1 ) = '@' then // Champ formule ( NON MODIFIABLE ! )
      begin
      FListe.ColFormats[i] := 'FORMULE=' + lStNomF + 'FORMAT=' + lStFormat;
      ActiveCol( i , False ) ;
      end
    else if Copy( Trim( lStNomChamp ), 1, 1 ) = '(' then // Champ formule ( NON MODIFIABLE ! )
      begin
      lStNomF := copy( lStNomChamp, pos('(',lStNomChamp) + 1, pos(')',lStNomChamp) - 2 );
      FListe.ColNames[ i ] := lStNomF ;
      end
    else if ( lStType = 'DOUBLE' ) then
      begin
      FListe.ColTypes[i]  := 'K';
      lInDec              := V_PGI.OkDecV ;
      if (( lStNomChamp = 'E_CREDITDEV' ) or ( lStNomChamp = 'E_DEBITDEV' ))
         and ( Assigned(Piece) ) and ( Piece.Devise.Code <> '' )
        then lInDec := Piece.Devise.Decimale
      else if (( lStNomChamp = 'E_QTE1' ) or ( lStNomChamp = 'E_QTE2' ))
        then lInDec := V_PGI.OkDecQ
      else if (( lStNomChamp = 'E_CREDIT' ) or ( lStNomChamp = 'E_DEBIT' ))
        then lInDec := V_PGI.OkDecV ;
      lStFormatDev := '#,##0';
      if lInDec > 0
        then lStFormatDev := lStFormatDev + '.' + Copy('000000000000000000000', 1, lInDec) ;
      FListe.ColFormats[i] := lStFormatDev ;
      end
    else if ( lStType = 'INTEGER' ) then
      begin
      FListe.ColTypes[i]   := 'R';
      FListe.ColFormats[i] := '#,##0';
      end
    else if ( pos( 'CHAR', lStType ) > 0 ) then
      begin
      FListe.ColTypes[i]   := #0 ;
      FListe.ColFormats[i] := '';
      // Réglage du nombre de caractères saisissable
      if pos( '17', lStType ) > 0 then
        FListe.ColLengths[i]:=17
      else if pos( '35', lStType ) > 0 then
        FListe.ColLengths[i]:=35
      else
        FListe.ColLengths[i]:=3 ;
      end
    else if ( lStType = 'SPECIF' ) then
      begin
      end ;

    // Cas particuliers....
    // 1. nb caractères maxi pour les généraux
    if lStNomChamp = 'E_GENERAL' then
      begin
      FListe.ColLengths[ i ] := GetInfoCpta( fbGene ).Lg ;
      end
    // 2. nb caractères maxi pour les auxiliaires
    else if lStNomChamp = 'E_AUXILIAIRE' then
      begin
      if VH^.CPCodeAuxiOnly
        then FListe.ColLengths[i] := GetInfoCpta( fbAux ).Lg
        else FListe.ColLengths[i] := 35 ;
      end
    // 3. saisie mono section
    else if pos( 'SECTION', lStNomChamp ) > 0 then
      begin
      FListe.ColLengths[ i ] := GetInfoCpta( AxeToFb( Copy( lStNomChamp, 8, 2) )).Lg ;
      end
    else
      begin end ;

    end ;

  if Assigned( RechCompte ) then
    begin
    RechCompte.COL_GEN := GetColIndex( 'E_GENERAL' ) ;
    RechCompte.COL_AUX := GetColIndex( 'E_AUXILIAIRE' ) ;
    end ;

  swapModeSaisie( Piece.ModeSaisie ) ;

//  if Assigned(TSDoc) then
//    paramViewer ;

  // Resize de la grille
  ResizeGrille;

  if assigned(FGrilleCompl) then
    begin
    updateMasqueCompl ;
    AfficheComplement( Row ) ;
    end ;

  if not Piece.EstRemplit( Row )
    then SetPos( GetColIndex('E_GENERAL'), Row, False, True )
    else SetPos( Col, Row, False, True ) ;

end ;

procedure TSaisiePiece.ParamMasqueSaisie;
begin
  if not assigned( FTobMasque ) then Exit ;
  if FTobMasque.GetString('CMS_TYPE')<>'SAI' then Exit ;

  CPLanceFiche_MasqueSaisie( '', FTobMasque.GetString('CMS_NUMERO'), 'ACTION=MODIFICATION' ) ;

  // Rechargement masque
  FTobMasque.ClearDetail ;
  FTobMasque.LoadDB ;
  CRemplitMasqueFromXML( FTobMasque ) ;

  // Maj grille
  UpdateMasqueSaisie ;

  // Réaffichage des données
  AfficheLignes(0, True) ;

end;

constructor  TSaisiePiece.CreerPourMasque(vTobMasque: Tob; vGrille: THGrid; vPieceCpt: TPieceCompta; vEcran: TForm)  ;
begin

  // Initialisation des paramètres
  FEcran         := vEcran;
  FListe         := vGrille;
  FTobPiece      := vPieceCpt ;
  FTobMasque     := vTobMasque ;

  // Création des objets
  FFindDialog        := TFindDialog.Create( Ecran );
  FFindDialog.name   := 'FindDialog';
  FFindDialog.OnFind := FindDialogFind;

  // Gestion de la personnalisation des champs libre
  InitParamLib ;

  // Réaffectation des evts de la grille
  InitEvtGrid ;
  InitFormatGrid ;

  // Gestion des lookup
  InitRechCompte ;

  // gestion des messages
  InitMessageCompta ;

{$IFDEF SBO}
 AddEvenement('=============================================');
 AddEvenement('TSP_Create >> Début : ' + DateToStr(Now) );
{$ENDIF SBO}

end;

procedure TSaisiePiece.ChoisirMasque;
var lStNum     : String ;
    lTobM      : Tob ;
begin

  // Sélection
  lStNum := Choisir( 'Sélection d''un masque de saisie',
                     'CMASQUESAISIE','CMS_LIBELLE','CMS_NUMERO','CMS_TYPE="SAI"','', False, False, 0) ;

  if ( lStNum<>'' ) then
    begin
    // Chargement
    lTObM := CSelectMasqueFromTob( StrToInt( lStNum ), FListeMasques ) ;
    if not assigned( lTObM ) then
      begin
      CChargeMasqueListe( StrToInt( lStNum ), FListeMasques, 'CEG' ) ;
      lTobM := CSelectMasqueFromTob( StrToInt( lStNum ), FListeMasques, 'CEG' ) ;
      end ;
    // mise en place
    if Assigned( lTObM ) then
      begin
      FTobMasque := lTobM ;
      UpdateMasqueSaisie ;
      AfficheLignes(0, True) ;
      end ;
    end ;

end;

procedure TSaisiePiece.RechercheMasque( vBoInit : boolean ) ;
var lTobM     : Tob ;
begin

  lTobM := GetTobMasque( vBoInit ) ;

  if not Assigned(FTobMasque) or
      (   ( lTobM.GetInteger('CMS_NUMERO')<>FTobMasque.GetInteger('CMS_NUMERO') )
       or ( lTobM.GetString('CMS_TYPE')<>FTobMasque.GetString('CMS_TYPE') ) ) then
    begin
    FTobMasque := lTobM ;
    UpdateMasqueSaisie ;
    end ;
{  else if (ctxPCL in V_PGI.PGIContexte) then
  begin
    paramViewer;
    end ;}
end;

function TSaisiePiece.GetTobCol(vStColName: string): Tob;
begin

  // Recherche accélérée...
  result := GetTobCol( GetColIndex( vStColName ) ) ;

  // Si résultat erroné, recherche détaillée
  if Assigned(Result) and ( result.GetString('NOM')<>vStColName ) then
    result := FTobMasque.Detail[0].FindFirst(['NOM'],[vStColName],True) ;

end;

function TSaisiePiece.GetTobCol(vInColIdx: integer): Tob;
begin
  if (vInColIdx >= 1) and (vInColIdx <= FTobMasque.Detail[0].DEtail.count)
    then result := FTobMasque.Detail[0].Detail[vInColIdx - 1]
    else result := nil ;
end;

procedure TSaisiePiece.OnChangeAuxi(Sender: TObject);
var lTobLigne : Tob ;
    lRow      : integer ;
begin
  if not Assigned(Sender) then Exit ;
  if GetActionFiche = taConsult then Exit ;
  lTobLigne := TOB( Sender ) ;
  lRow      := (lTobLigne.GetIndex + 1) ;

  Piece.PutValue( lRow, 'E_AUXILIAIRE', lTobLigne.GetValue('E_AUXILIAIRE'), True) ;

  if GuideActif then
    TSGuide.RecalculGuide( lRow )
  else
    if (Piece.ModeSaisie=msBor)
      then AfficheGroupe( lRow )
      else AfficheLignes ;

end;

procedure TSaisiePiece.OnChangeGeneral(Sender: TObject);
var lTobLigne : Tob ;
    lRow      : integer ;
begin
  if not Assigned(Sender) then Exit ;
  if GetActionFiche = taConsult then Exit ;

  lTobLigne := TOB( Sender ) ;
  lRow      := (lTobLigne.GetIndex + 1) ;

  Piece.PutValue( lRow, 'E_GENERAL', lTobLigne.GetValue('E_GENERAL'), True ) ;

  if GuideActif then
    TSGuide.RecalculGuide( lRow )
  else
    if (Piece.ModeSaisie=msBor)
      then AfficheGroupe( lRow )
      else AfficheLignes ;

end;

procedure TSaisiePiece.UpdateMasqueCompl;
Var i            : Integer ;
    n            : Integer ;
    lStNomChamp  : String ;
    lStType      : String ;
    lStNomF      : String ;
    lInDec       : integer ;
    lBoVisu      : boolean;
    lStFormat    : String ;
    lStTitreCol  : String ;
    lStFormatDev : string ;
    lTobGrille   : Tob ;
    lTobCol      : Tob ;
    HMT          : TComponent ;
begin

  if not Assigned( FGrilleCompl ) then Exit ;

  lTobGrille := FTobMasque.FindFirst( ['TYPE'], ['COMPLEMENTS'], True ) ;

  if ( not Assigned(lTobGrille) ) or ( lTobGrille.Detail.count = 0 ) then
    begin
    FGrilleCompl.visible := False ;
    FGrilleCompl.Refresh ;
    Exit ;
    end
  else FGrilleCompl.visible := True ;

  // Calcul su nombre de colonne :
  FGrilleCompl.ColCount := lTobGrille.Detail.count + FGrilleCompl.FixedCols ;
  SetLength( FTabColAvecLibC, FListe.ColCount + 1 ) ;

  // hauteur de la ligne de titre
  FGrilleCompl.RowHeights[0] := 18 ;

  if FGrilleCompl.DBIndicator
    then n := 1
    else n := 0 ;

  // Parcours des lignes pour formatage spécifique des colonnes
  for i := n to lTobGrille.Detail.count - 1 + n do
    begin

    lTobCol := lTobGrille.Detail[ i-n ];

    lStNomChamp := lTobCol.GetString('NOM') ;
    lStType     := lTobCol.GetString('TYPE') ;

    // Nom des colonnes
    FGrilleCompl.ColNames[ i ] := lStNomChamp ;

    // Affichage du libellé complet :
    FTabColAvecLibC[ i ] := lTobCol.GetString('AVECLIBELLE')='X' ;

    // largeur des colonnes
    FGrilleCompl.ColWidths[i]  := lTobCol.GetInteger('LARGEUR') ;

    // Alignements
    // Numéro de ligne
    if (lStNomChamp = 'E_NUMLIGNE') or (lTobCol.GetString('TYPE') = 'DATE')
      then FGrilleCompl.ColAligns[i]  := taCenter
    // Entiers / Réels
    else if (lTobCol.GetString('TYPE') = 'INTEGER') or (lTobCol.GetString('TYPE') = 'DOUBLE')
      then FGrilleCompl.ColAligns[i]  := taRightJustify
    // Text
    else if (lTobCol.GetString('TYPE') = 'COMBO') and not ( FTabColAvecLibC[ i ] )
      then FGrilleCompl.ColAligns[i]  := taCenter
    // Tous les autres cas
    else FGrilleCompl.ColAligns[i]  := taLeftJustify ;

    // Titres + Perso table libre
    lStTitreCol := lTobCol.GetString('TITRE') ;
    if ( pos('LIBRE', lStNomChamp ) > 0 ) or ( pos('E_TABLE', lStNomChamp ) > 0 ) then
      CGetParamLibre( FTobParamLib, lStNomChamp, lStTitreCol, lBoVisu ) ;
    FGrilleCompl.Cells[i,0]    := lStTitreCol ;

    // Formats des colonnes...
    if ( lStType = 'BLOB' ) then
      begin
      FGrilleCompl.ColTypes[i]   := 'M';
      FGrilleCompl.ColFormats[i] := IntToStr( Byte( msBook ) );
      end
    else if ( lStType = 'BOOLEAN' ) then
      begin
      FGrilleCompl.ColTypes[i]       := 'B';
      FGrilleCompl.ColFormats[i]     := IntToStr( Byte( csCheckBox ) );
      end
    else if ( lStType = 'DATE' ) then
      begin
      if (lStNomChamp = 'E_DATECOMPTABLE') then
        begin
        FGrilleCompl.ColTypes[i]     := 'I';
        FGrilleCompl.ColLengths[ i ] := 2 ;
        FGrilleCompl.ColFormats[i]   := '' ;
        end
      else
        begin
        FGrilleCompl.ColTypes[i]   := 'D';
        FGrilleCompl.ColFormats[i] := '';  //sinon effet de bord // ??
        end
      end
    else if Copy( Trim( lStNomChamp ), 1, 1 ) = '@' then // Champ formule ( NON MODIFIABLE ! )
      begin
      FGrilleCompl.ColFormats[i] := 'FORMULE=' + lStNomF + 'FORMAT=' + lStFormat;
      end
    else if Copy( Trim( lStNomChamp ), 1, 1 ) = '(' then // Champ formule ( NON MODIFIABLE ! )
      begin
      lStNomF := copy( lStNomChamp, pos('(',lStNomChamp) + 1, pos(')',lStNomChamp) - 2 );
      FGrilleCompl.ColNames[ i ] := lStNomF ;
      end
    else if ( lStType = 'DOUBLE' ) then
      begin
      FGrilleCompl.ColTypes[i]  := 'K';
      lInDec              := V_PGI.OkDecV ;
      if (( lStNomChamp = 'E_CREDITDEV' ) or ( lStNomChamp = 'E_DEBITDEV' ))
         and ( Assigned(Piece) ) and ( Piece.Devise.Code <> '' )
        then lInDec := Piece.Devise.Decimale
      else if (( lStNomChamp = 'E_QTE1' ) or ( lStNomChamp = 'E_QTE2' ))
        then lInDec := V_PGI.OkDecQ
      else if (( lStNomChamp = 'E_CREDIT' ) or ( lStNomChamp = 'E_DEBIT' ))
        then lInDec := V_PGI.OkDecV ;
      lStFormatDev := '#,##0';
      if lInDec > 0
        then lStFormatDev := lStFormatDev + '.' + Copy('000000000000000000000', 1, lInDec) ;
      FGrilleCompl.ColFormats[i] := lStFormatDev ;
      end
    else if ( lStType = 'INTEGER' ) then
      begin
      FGrilleCompl.ColTypes[i]   := 'R';
      FGrilleCompl.ColFormats[i] := '#,##0';
      end
    else if ( pos( 'CHAR', lStType ) > 0 ) then
      begin
      FGrilleCompl.ColTypes[i]   := #0 ;
      FGrilleCompl.ColFormats[i] := '';
      end ;

    end ;

  // Resize de la grille si la taille total des colonnes dépasse l'affichage
  if GetFullSize( FGrilleCompl ) >= (FGrilleCompl.Width/2) then
    begin
    HMT := TForm( Ecran ).FindComponent('HMTrad');
    if (HMT<>nil) then
      THSystemMenu(HMT).ResizeGridColumns( FGrilleCompl );
    end;

end;

procedure TSaisiePiece.SetGrilleComplement(vGrille: THGrid);
begin
  FGrilleCompl := vGrille ;
  if Assigned( vGrille ) then
    begin

    SetLength( FTabColAvecLibC, vGrille.ColCount + 1 ) ;

    vGrille.PostDrawCell  := GComplDrawCell ;
//    vGrille.GridLineWidth := 0 ;
    vGrille.Height        := 38 ;
    vGrille.RowHeights[0] := 18 ; // hauteur de la ligne de titre
    vGrille.Enabled       := False ;
    vGrille.CacheEdit ;
    vGrille.Invalidate ;
    end ;
end;

procedure TSaisiePiece.AfficheComplement(vRow: integer);
Var i            : Integer ;
    lStVal       : String ;
    lTobGrille   : Tob ;
    lTobCol      : Tob ;
    lTobEcr      : Tob ;
begin

  if not Assigned( FGrilleCompl ) then Exit ;
  if not FGrilleCompl.Visible     then Exit ;

  try

      FGrilleCompl.BeginUpdate ;

      // Init des zones
      For i := 0 to FGrilleCompl.ColCount - 1 do
        FGrilleCompl.Cells[ i, 1 ] := '' ;

      // Si pas de données corerspondantes, on arrête
      if Piece.IsOut(vRow) then  Exit ;

      // Recherche paramétrage
      lTobGrille := FTobMasque.FindFirst( ['TYPE'], ['COMPLEMENTS'], True ) ;
      if not Assigned(lTobGrille) then Exit ;

      // Parcours des colonnes pour restitutino des informations
      lTobEcr := Piece.GetTob( vRow ) ;
      for i := 0 to lTobGrille.Detail.count-1 do
        begin
        lTobCol     := lTobGrille.Detail[ i ];

        // Formatage pour la grille
        lStVal := CVariantToStGrid( GetGoodValeurGrille(lTobCol.GetString('NOM'),lTobEcr,'COMPLEMENTS'), FGrilleCompl, i ) ;

         // MAJ de la cellule
        FGrilleCompl.Cells[ i , 1 ] := lStVal ;

        end ;

  finally

  // Resize de la grille si la taille total des colonnes dépasse l'affichage
{  if GetFullSize( FGrilleCompl ) >= FGrilleCompl.Width then
    begin
    HMT := TForm( Ecran ).FindComponent('HMTrad');
    if (HMT<>nil) then
      THSystemMenu(HMT).ResizeGridColumns( FGrilleCompl );
    end;
}
    FGrilleCompl.EndUpdate ;

    end ;

end;

function TSaisiePiece.GetGoodValeurGrille(vStChamp : String ; vTobEcr : TOB ; vStGrille : string) : Variant;
var
  lStPref      : String ;
  lStCal       : String ;
  lTobEC       : Tob ;
begin
  Result     := #0 ;

  if CEstColCalcul(vStChamp) then
  begin
     lStCal := RecupSpecif(vStChamp,vStGrille,'CALCUL');
     if lStCal <> '' then
     begin
        FGrille:= vStGrille;
        FTob   := vTobEcr;
        Result := GFormule(lStCal, GetFormule, nil, 1 );
//        ValideCellule( vStChamp, GetRow, False ) ;
        FTobPiece.PutValue(GetRow, vStChamp  , Result ) ;
        Exit ;
     end;
  end;

  if pos( '_', vStChamp) > 0
    then lStPref := Copy( vStChamp, 1, pos( '_', vStChamp) - 1 )
    else lStPref := '' ;

  // ===== Zones Ecriture =====
  if lStPref = 'E' then
    Result := vTobEcr.GetValue( vStChamp )

  // ===== Zones Général =====
  else if lStPref = 'G' then
    begin
    if Piece.Info_LoadCompte( vTobEcr.GetValue( 'E_GENERAL' ) ) then
      Result := Piece.Info.Compte.GetValue( vStChamp ) ;
    end

  // ===== Zones Auxiliaire =====
  else if lStPref = 'T' then
    begin
    if Piece.Info_LoadAux( vTobEcr.GetValue( 'E_AUXILIAIRE' ) ) then
      Result := Piece.Info.Aux.Item.GetValue( vStChamp) ;
    end

  // ===== Zones Journal =====
  else if lStPref = 'J' then
    begin
    if Piece.Info_LoadJournal( vTobEcr.GetValue( 'E_JOURNAL' ) ) then
      Result := Piece.Info.Journal.Item.GetValue( vStChamp ) ;
    end

  // ===== Zones ECRCompl =====
  else if lStPref = 'EC' then
    begin
    lTobEC := CGetTOBCompl( vTobEcr ) ;
    if Assigned( lTobEC )
      then Result := lTobEC.GetValue( vStChamp )
      else Result := #0 ;
    end

  else if (lStPref = '') and (pos('FORMULE',vStChamp) = 1 ) then
    begin
    FGrille := vStGrille;
    FTob    := vTobEcr;
    Result  := GFormule( RecupSpecif(vStChamp,vStGrille, 'FORMULE'), GetFormule, nil, 1 );
    end
  else
    result := vTobEcr.GetValue( vStChamp ) ;

end;

function TSaisiePiece.RecupChampGrille(vStChamp : String ; vStGrille : String ) : TOB;
var
  TobGrille  : TOB;
begin
  Result := nil;
  TobGrille := FTobMasque.FindFirst(['TYPE'],[vStGrille],True) ;
  if TobGrille = nil then Exit;
  Result := TobGrille.FindFirst(['NOM'],[vStChamp],True) ;
end;

function TSaisiePiece.RecupSpecif(vStChamp : String ; vStGrille : String ; vStType : String) : String;
var
  maTob : TOB;
begin
  Result := '';
  maTob := RecupChampGrille(vStChamp,vStGrille);
  if maTob = nil then Exit;
  if maTob.FieldExists(vStType) then
     Result := maTob.GetString(vStType);
end;

procedure TSaisiePiece.GComplDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var lRect          : TRect ;
    lInPosX        : integer ;
    lInPosY        : integer ;
    lText          : String ;
    lBoReappliquer : boolean ;
    lStVal         : string ;
    lStDT          : string ;
    lStChamp       : string ;
begin
  lRect    := FGrilleCompl.CellRect(ACol,ARow) ;

  // Titres
  if (ARow = 0) then
    begin
    Canvas.Font.Style  := [] ;
    Canvas.Font.Color  := clWindowText	;
    Canvas.Brush.Color := _COLORTITRE	;
    lText   := FGrilleCompl.Cells[Acol, 0] ;
    lInPosX  := lRect.Left  + ((lRect.Right - lRect.Left - canvas.TextWidth( lText ) ) div 2 ) ;
    lInPosY  := 1 ;
    Canvas.TextRect( lRect, lInPosX, lInPosY  , lText );
    end
  // Gestion des champs COMBO affiché en libellé
  else if (ARow = 1 ) then
    begin
    //Affichage du libellé pour les champs combo :
    lStChamp := FListe.ColNames[ ACol ] ;
    if lStChamp = '' then Exit ;
    lBoReappliquer := False ;
    if ( FGrilleCompl.Cells[ACol, ARow] <> '' ) then
      begin
      if FTabColAvecLib[ ACol ] then
        begin
        lStDT := GetChampDataType( lStChamp ) ;
        if lStDT = '' then
          lStDT := ChampToTT( lStChamp ) ;
        if lStDT <> '' then
          begin
          lStVal := FGrilleCompl.Cells[ACol, ARow] ;
          if lStVal <> '' then
            begin
            lText := RechDom( lStDT, lStVal, False ) ;
            if (lText = '') or (lText = 'Error')
              then lText          := ''
              else lBoReappliquer := True ;
            end ;
          end ;
        end ;
      end ;

    if lBoReappliquer then
      begin
      Canvas.FillRect( lRect );
      case FGrilleCompl.FColAligns[ACol] of
        taRightJustify :  lInPosX := lRect.Right - Canvas.TextWidth(lText) - 3 ;
        taCenter       :  lInPosX := lRect.Left  + ((lRect.Right - lRect.Left - canvas.TextWidth( lText ) ) div 2 ) ;
        else              lInPosX := (lRect.Left + 3 ) ;
        end ;
      Canvas.TextRect( lRect, lInPosX, (lRect.Top+lRect.Bottom) div 2 -5 , lText );
      end ;

    end ;

end;

function TSaisiePiece.CVariantToStGrid(vValeur: Variant; vGrille: THGrid; vInCol: integer): string;
var lStVal     : string ;
    lStMask    : string ;
    lDtDate    : TDateTime ;
   {$IFDEF EAGLCLIENT}
    lTobCol    : Tob ;
   {$ENDIF}
begin

  case VarType( vValeur ) of

     varNull: lStVal := '';

     varDate:   begin
                lDtDate := VarAsType(vValeur, vardate);
                lStMask := vGrille.ColFormats[ vInCol ];
                if (vGrille.ColFormats[vInCol] = #0) or (Pos('#', vGrille.ColFormats[vInCol]) > 0) then
                  lStMask := '';
                if lStMask = '' then
                  begin
                  if Int(lDtDate) <> 0 then // Si Il y a bien un jour dans la date
                    begin
                    // Si une heure aussi ?
                    if Frac(lDtDate) <> 0
                      then lStMask := StDateTimeFormat // Format comprend le jour et l'heure
                      else lStMask := ShortDateFormat; // uniquement le jour
                    end
                  else // XP 27-06-2005 : Cas d'une heure sans jour
                    begin
                    if Frac(lDtDate) <> 0
                      then lStMask := ShortTimeFormat // Format heure seulement
                      else lStMask := ''; // Pas de format : Valeur à 0
                    end;
                  end;
                if (lDtDate <= iDate1900) and not ((lDtDate > 0) and (lDtDate < 1))
                  then lStVal := ''
                  else lStVal := FormatDatetime(lStMask, lDtDate);
              end;

      varDouble: begin
                 if vGrille.ColFormats[vInCol] = #0
                   then lStMask := ''
                   else lStMask := vGrille.ColFormats[vInCol];
                 if lStMask = '' then
                   lStMask := '#,##0.00';
                  lStVal := FormatFloat(lStMask, VarAsType(vValeur, varDouble));
                 end;

      varInteger: begin
                  if vGrille.ColFormats[vInCol] = #0
                    then lStMask := ''
                    else lStMask := vGrille.ColFormats[vInCol];
                  if lStMask = '' then
                    lStMask := '#,##0';
                  lStVal := FormatFloat(lStMask, VarAsType(vValeur, varDouble));
                  end;
      {$IFNDEF EAGLCLIENT}
        {$IFDEF VER150}
        varBoolean: lStVal := BoolToStr_(vValeur);
        {$ELSE VER150}
        varBoolean: lStVal := BoolToStr(vValeur);
        {$ENDIF VER150}
      {$ENDIF}

      else
        begin
        {$IFDEF EAGLCLIENT}
        lTobCol := TOB( vGrille.Objects[vInCol, 0] ) ;
        if (Assigned(lTobCol) and (lTobCol.GetString('TYPE') = 'BOOLEAN')) then
          vGrille.ColFormats[ vInCol ] := IntToStr( Byte( csCoche ) ) ;
        {$ENDIF}
        {$IFDEF UNICODE}
        if ( Vartype( vValeur ) = varOleStr ) and V_PGI.DBUnicode
          then lStVal := VarAsType(vValeur, varOleStr)
          else
        {$ENDIF}
        lStVal := VarAsType(vValeur, varString);

        if (Copy(lStVal, 1, 5) = '{\rtf') then
          begin
          vGrille.Coltypes[vInCol] := 'M';
          if (Length(lStVal) > 300) or (CountRTFString(lStVal) > 0)
            then lStVal := #1
            else lStVal := '';
          end;

        end; // else

    end;  // case

  result := lStVal ;

end;

procedure TSaisiePiece.SaisieDevise ( vRow : integer ) ;
var lTobEcr       : Tob ;
    lCumulD       : Double ;
    lCumulC       : Double ;
    lInDebG       : integer ;
    lInFinG       : integer ;
    lAction       : TActionFiche ;
    lMonoDev      : Boolean ;
    lSensDef      : string ;
    lCodeDev      : string ;
    lParams       : RSDev ;
    i             : integer ;
    lDevPivot     : RDevise ;
    lSolde        : Double ;
    lOldDebit     : double ;
    lOldCredit    : double ;
begin

  if not EstDeviseEditable( vRow ) then Exit ;

  // InitVariable
  lTobEcr  := Piece.GetTOB( vRow ) ;
  lMonoDev := False ;

  // Calcul du solde
  Piece.GetBornesGroupe(lTobECr.GetInteger('E_NUMGROUPEECR'), lInDebG, lInFinG);
  lCumulD := 0 ;
  lCumulC := 0 ;
  for i:= lInDebG to lInFinG do
    begin
    lCumulD := lCumulD + Piece.Detail[i-1].GetDouble('E_DEBITDEV') ;
    lCumulC := lCumulC + Piece.Detail[i-1].GetDouble('E_CREDITDEV') ;
    end ;
  lSolde := Arrondi( lCumulD-lCumulC, Piece.Devise.Decimale) ;

  // Détermination de l'action
  if (Piece.Action=taConsult) or not ( Piece.EstChampsModifiable(vRow,'E_DEBITDEV') and Piece.EstChampsModifiable(vRow,'E_CREDITDEV') ) then
    lAction := taConsult
    // Si montant renseigné sur le groupe, modif de la devise impossible
  else if (lCumulD=0) and (lCumulC=0) then lAction := taCreat
  else lAction := taModif ;

  // Quel champs
  if Piece.QuelChampsMontant( vRow ) = 'E_CREDITDEV'
    then lSensDef := 'C'
    else lSensDef := 'D' ;

  // détermination de la devise par défaut
  if (lTobEcr.GetString('E_DEVISE')=Piece.GetEnteteS('E_DEVISE')) then
    begin
    if (lTobEcr.GetString('E_AUXILIAIRE')<>'') and (Piece.Info_LoadAux(lTobEcr.GetString('E_AUXILIAIRE'))) then
      begin
      lCodeDev := Piece.Info.Aux.Item.GetString('T_DEVISE') ;
      if lCodeDev<>'' then
        begin
        lTobEcr.PutValue('E_DEVISE', lCodeDev) ;
        lMonoDev := Piece.Info.Aux.Item.GetString('T_MULTIDEVISE') = '-'  ;
        end ;
      end
    else
      begin
      lCodeDev := GetDeviseBanque( lTobEcr.GetString('E_GENERAL') ) ; // saisUtil (req BANQUECP !)
      if (lCodeDev<>'') then
        begin
        lTobEcr.PutValue('E_DEVISE', lCodeDev) ;
        lMonoDev:=(lCodeDev<>V_PGI.DevisePivot) ;
        end ;
      end ;
    end ;

  lDevPivot.Code := V_PGI.DevisePivot ;
  GetInfosDevise(lDevPivot) ;

  lParams.Ecr               := lTobEcr ;
  lParams.DPivot            := lDevPivot ;
  lParams.TD                := lCumulD ;
  lParams.TC                := lCumulC ;
  lParams.Solde             := lSolde ;
  lParams.SensDef           := lSensDef ;
  lParams.bMonoDev          := lMonoDev ;
  lParams.bDoSolde          := False ;
  lParams.Action            := lAction ;

  lOldDebit                 := lTobEcr.GetDouble('E_DEBITDEV') ;
  lOldCredit                := lTobEcr.GetDouble('E_CREDITDEV') ;

  if Piece.EstTauxVolatil( vRow ) then
    begin
    lParams.volatile    := True ;
    lParams.DeviseLigne := Piece.GetRDevise(vRow) ;
    end ;

  // ---
  if SaisieZDevise( lParams ) then
    begin
    // Recup modif éventuel sur la devise
    if lTobEcr.GetString('E_DEVISE') <> Piece.GetEnteteS('E_DEVISE') then
      Piece.Info_LoadDevise( lTobEcr.GetString('E_DEVISE') ) ;
    // Récup du taux
    if Piece.GetRDevise( vRow ).Taux <> lParams.DeviseLigne.Taux then
      Piece.SetTauxVolatil( vRow, lParams.DeviseLigne ) ;

    // Revalidation des modifs de la ligne
    if (lOldDebit <> lTobEcr.GetDouble('E_DEBITDEV')) or (lOldCredit<>lTobEcr.GetDouble('E_CREDITDEV')) then
      if lTobEcr.GetDouble('E_DEBITDEV') <> 0
        then Piece.PutValue ( vRow, 'E_DEBITDEV', lTobEcr.GetDouble('E_DEBITDEV'), True )
        else Piece.PutValue ( vRow, 'E_CREDITDEV', lTobEcr.GetDouble('E_CREDITDEV'), True ) ;

    // Si analytique non encore affecté alors on le fait maintenant
{    Piece.AffecteAna( vRow ) ;
    // MAJ des échéances
    Piece.AffecteEche( vRow ) ;
    // Recalcul du tableau des montants par tva
    Piece.CalculTvaEnc( vRow ) ;
}
    // equilibre
    Piece.VerifieEquilibre( vRow ) ;
    end ;

  AfficheGroupe( vRow ) ;

end;

function TSaisiePiece.EstDeviseEditable(vRow: integer): boolean;
var lTobEcr       : Tob ;
begin
  result := False ;

  if Piece.ModeSaisie <> msBor then Exit ;
  if not Piece.EstRemplit( vRow ) then Exit ;

  // Si le journal n'est pas multi-devise on sort
  if not (Piece.Info.Journal.GetString('J_MULTIDEVISE') = 'X') then Exit ;
  if ( Piece.GetEnteteS('E_DEVISE') <> V_PGI.DevisePivot ) then Exit ;

  // InitVariable
  lTobEcr  := Piece.GetTOB( vRow ) ;

  // Si présence d'une immo, on sort
  if lTobEcr.GetString('E_IMMO')<>'' then Exit ;

  // Auxiliaire non saisi ou monodevise pivot, on sort
  if Piece.Info.Compte.IsCollectif then
    if (lTobEcr.GetString('E_AUXILIAIRE')='') or not (Piece.Info_LoadAux(lTobEcr.GetString('E_AUXILIAIRE')))
      then Exit
      else if (Piece.Aux_GetValue('T_DEVISE')=V_PGI.DevisePivot) and
              (Piece.Aux_GetValue('T_MULTIDEVISE')='-')
             then Exit ;

  // Si pas la 1ère ligne d'un groupe, on sort
  if ( lTobEcr.GetString('E_DEVISE') = Piece.GetEnteteS('E_DEVISE') ) and
     ( vRow > 1 ) and
     ( Piece.GetValue(Row-1, 'E_NUMGROUPEECR') = lTobEcr.GetValue('E_NUMGROUPEECR') )
    then exit;

  // Si montant renseigné en EURO, on sort
  if ( lTobEcr.GetString('E_DEVISE') = Piece.GetEnteteS('E_DEVISE') ) and
     ( ( lTobEcr.GetDouble('E_DEBIT')<>0 ) or ( lTobEcr.GetDouble('E_DEBIT')<>0 ) )
    then Exit ;

  result := True ;

end;

procedure TSaisiePiece.SetPosFin(vBoTotal: boolean);
var lInCol : integer ;
    lInRow : integer ;
begin

  // On va sur la dernière ligne
  if vBoTotal then
    begin
    // denière ligne
    lInRow := Piece.Count ;
    // 1ère col dispo
    lInCol := GetFirstCol ;
    end

  else // On va à la fin de la ligne
    begin
    // ligne courante
    lInRow := Row ;
    // calcul dernière col dispo en partant de la fin
    if (GoEditing in FListe.Options) then
      begin
      lInCol := FListe.colCount - 1 ;
      while (lInCol > Col) and not (EstColEditable( lInRow, GetColName( lInCol ) ) ) do
        Dec( lInCol ) ;
      end
    else
      lInCol := Col ;

    end ;

  // Positionnement
  SetPos( lInCol, lInRow , True, True, -1) ;

end;

procedure TSaisiePiece.SetPosDebut(vBoTotal: boolean);
var lInCol : integer ;
    lInRow : integer ;
begin

  // ligne
  if vBoTotal
    then lInRow := 1  // On revient sur la 1ère ligne
    else lInRow := Row ;

  // 1ère col dispo
  lInCol := GetFirstCol ;

  // Positionnement
  SetPos(lInCol, lInRow, True, True, 1) ;

end;

function TSaisiePiece.GetFirstCol: integer;
begin

  if (GoEditing in FListe.Options) then
    begin
    result := 1 ;
    while (result < FListe.colCount-1) and not (EstColEditable( result, GetColName( result ) )) do
      Inc( result ) ;
    end
  else
    result := Col ;

end;


procedure TSaisiePiece.DupliquerZone(vBoFromFirst, vBoOnlyRef: boolean);
var lStChp : string ;
    lInRow : integer ;
    lBoMAJ : boolean ;
    function _MAJChamp( vStChamp : string ) : boolean ;
      begin
      result := false ;
      if EstColEditable( Row, vStChamp ) then
        begin
        Piece.PutValue( Row, vStChamp, Piece.GetValue( lInRow, vStChamp ) ) ;
        result := True ;
        end ;
      end ;
begin

  if csDestroying in Ecran.ComponentState then Exit ;
  if GetActionFiche = taConsult then Exit ;
  if (GoRowSelect in FListe.Options) then Exit ;
  if Piece.isOut( Row ) then Exit ;

  // recherche ligne de reférence
  if vBoFromFirst then
    Case Piece.ModeSaisie of
      msBor : lInRow := Piece.GetDebutGroupe( Row ) ;
      else    lInRow := 1 ;
      end
  else
    lInRow := Row - 1 ;
  if Piece.isOut( lInRow ) then Exit ;

  // recherche champ cible
  lBoMAJ := False ;
  if vBoOnlyRef then
    begin
    if _MAJChamp( 'E_REFINTERNE' ) then
      lBoMAJ := True ;
    if _MAJChamp( 'E_LIBELLE' ) then
      lBoMAJ := True ;
    end
  else
    begin
    // rech du champ
    lStChp := GetColName( Col ) ;
    if lStChp = '' then Exit ;
    lBoMAJ := _MAJChamp( lStChp ) ;
    end ;

  if lBoMAJ then
    begin
    AfficheLignes( Row ) ;
    if not vBoOnlyRef then // se place à la fin du mot en recopie de cellule
      FListe.InplaceEditor.SelStart := Length( FListe.Cells[Col,Row] ) ;
    end ;

end;

procedure TSaisiePiece.SoldePiece;
var lInRow  : integer ;
    lInCol  : integer ;
    lCancel : boolean ;
begin
  if GetActionFiche = taConsult   then Exit ;
  if Piece.Count < 2              then Exit ;
  lInRow := Row ;
  if Piece.EstPieceSoldee( lInRow )  then Exit ;

  // Validation de la cellule en cours de saisie...
  lInCol  := Col ;
  lCancel := False ; // FQ 21102 : Pb sur touche F6 / menu pop calcul du solde / ....
  CellExit( LaGrille, lInCol, lInRow, lCancel ) ;
  if lCancel then Exit ;

  // Ligne non remplit - on essaie de renseigner le compte auto / contrepartie
  if not Piece.EstRemplit( lInRow )  then
    begin
    GetCompteAuto( True, True ) ;
    PutTobToGrid( lInRow, 'E_GENERAL' ) ;
    end ;

  // si toujours absence de compte, on sort
  if not Piece.EstRemplit( lInRow )  then Exit ;

  // Maj du solde
  Piece.AttribSolde( lInRow ) ;

  // rafraichissement des zones
  PutTobToGrid( lInRow, 'E_DEBITDEV;E_CREDITDEV' ) ;

  // Gestion des evènements de sortie
  if not ValideSortieLigne then Exit ;

  // Ajout prochaine lignes ?
  if lInRow = Piece.Count then
    Piece.NewRecord ;

  // Raffraichissement grille
  AfficheLignes ;

  ProchaineLigne(lInRow, '', False) ;

end;

procedure TSaisiePiece.GetCompteAuto( vBoSuivant: Boolean ; vBoAuto : Boolean );
var lStCompte  : String ;
    lInCol     : Integer ;
    lInRow     : Integer ;
begin

  if GetActionFiche = taConsult then Exit ;

  lInCol  := GetColIndex('E_GENERAL') ;
  lInRow  := Row ;
  if (not vBoAuto) and (Col <> lInCol) then exit ;

  if Piece.EstRemplit( lInRow ) then Exit ;

  // Utilisé dans 2 cas : journaux banque/Cai/effet et autres...
  Piece.Info_LoadJournal( Piece.GetEnteteS('E_JOURNAL') ) ;
  if Piece.EstJalEffet or Piece.EstJalBqe then
    begin
    lStCompte := Info.Journal.Item.GetString('J_CONTREPARTIE') ;
    end
  else
    begin
    if Row = FInLigneCourante then
      if vBoSuivant then
        begin
        Inc( FInIndexComptaAuto ) ;
        if FInIndexComptaAuto > Info.Journal.NombreDeCompteAuto then
          FInIndexComptaAuto:=1 ;
        end
      else
        begin
        Dec( FInIndexComptaAuto ) ;
        if FInIndexComptaAuto < 1 then
          FInIndexComptaAuto := Info.Journal.NombreDeCompteAuto ;
        end
    else
      begin
      FInLigneCourante   := Row ;
      FInIndexComptaAuto := 1 ;
      end;

    Info.LoadJournal( Piece.GetEnteteS('E_JOURNAL') ) ;
    if Info.Journal.CompteAuto = '' then exit ;
    lStCompte := TrouveAuto( Info.Journal.CompteAuto , FInIndexComptaAuto ) ;

    end ;

  if lStCompte='' then Exit ;
  if EstInterdit( Info.Journal.GetValue('J_COMPTEINTERDIT'), lStCompte, 0) > 0 then Exit ;

  if vBoAuto then
    Piece.PutValue( lInRow, 'E_GENERAL', lStCompte )
  else
    FListe.Cells[ lInCol, lInRow ] := lStCompte ;

end;

procedure TSaisiePiece.ExecuteLibelleAuto( vRow : integer );
var lStSaisi      : string ;
    lInCol        : integer ;
    lTobLibAuto   : Tob ;
    lTobEcr       : Tob ;
begin
  if (GetActionFiche=taConsult) then Exit ;
  lTobEcr := Piece.Gettob( vRow ) ;
  if lTobEcr = nil then Exit ;

  if not EstColEditable( vRow, 'E_LIBELLE') then Exit ;

  lInCol   := GetColIndex('E_LIBELLE') ;
  lStSaisi := FListe.Cells[ lInCol, vRow ] ;

  lTOBLibAuto := FZlibAuto.RechercheEnBase( nil, lStSaisi, lTobEcr.GetString('E_JOURNAL'), lTobEcr.GetString('E_NATUREPIECE'), lTobEcr ) ;
  if assigned(lTOBLibAuto) then
    begin
    // maj de la ref interne si
    if ( Piece.GetString( vRow, 'E_REFINTERNE' ) = '') and ( EstColEditable( vRow, 'E_REFINTERNE') ) then
      Piece.ExecuteFormule( vRow, 'E_REFINTERNE', lTOBLibAuto.GetString('RA_FORMULEREF') ) ;
    Piece.ExecuteFormule( vRow, 'E_LIBELLE', lTOBLibAuto.GetString('RA_FORMULELIB') ) ;
    end ;

  AfficheLignes( vRow ) ;

end;


procedure TSaisiePiece.IncrementerRef(vRow: integer; vBoPlus: boolean);
begin

  if GetActionFiche = taConsult then Exit ;
  if not EstColEditable( vRow, 'E_REFINTERNE' ) then Exit ;

  if Col = GetColIndex('E_REFINTERNE') then
    ValideCellule( 'E_REFINTERNE', vRow ) ;

  Piece.IncrementRef( vRow, FLastRef, vBoPlus ) ;

  AfficheLignes( vRow ) ;

end;


function TSaisiePiece.GetTobMasque(vBoInit: boolean) : Tob ;
var lstCrit1  : string ;
    lstCrit2  : string ;
    lstCrit3  : string ;
    lInMDef   : integer ;
begin

  result := FTobMasque ;

  lStCrit1 := CGetMasqueCrit1( Piece ) ;
  lStCrit2 := CGetMasqueCrit2( Piece ) ;
  lStCrit3 := CGetMasqueCrit3( Piece ) ;

  // si aucune modif on sort...
  if (not vBoInit) and
     not ((FStMasqueCrit1 <> lStCrit1) or (FStMasqueCrit2 <> lStCrit2) or (FStMasqueCrit3 <> lStCrit3)) then
    Exit ;

  // Le masque est-il chargé ?
  result := CSelectMasqueFromTob( lStCrit1, lStCrit2, lStCrit3, FListeMasques, True ) ;

  // sinon, on recherche les nouveaux masques
  if not Assigned( result ) then
    begin
    CChargeMasqueListe( lStCrit1, lStCrit2, lStCrit3, FListeMasques, Piece ) ;
    // Recherche élargie
    result := CSelectMasqueFromTob( lStCrit1, lStCrit2, lStCrit3, FListeMasques, False, Piece ) ;
    end ;

  // Trouvé...
  if not Assigned( result ) then
    begin
    lInMDef := GetParamSocSecur( 'SO_CPMASQUEDEFAUT', 0 ) ;
    if ( lInMDef <> 0 )
      then begin
           result := CSelectMasqueFromTob( lInMDef, FListeMasques, 'CEG' ) ;
           if not Assigned( result ) then
             begin
             CChargeMasqueListe( lInMDef, FListeMasques, 'CEG' ) ;
             result := CSelectMasqueFromTob( lInMDef, FListeMasques, 'CEG' ) ;
             end ;
           end ;
    end ;

  // MAJ des critères...
  FStMasqueCrit1 := lStCrit1 ;
  FStMasqueCrit2 := lStCrit2 ;
  FStMasqueCrit3 := lStCrit3 ;

end;

function TSaisiePiece.GereAcc( vRow : integer ; vBoForce : boolean ) : integer;
var lStAux      : String ;
    lStCptHT    : String ;
    lInRow      : integer ;
    lStCol      : string ;
begin

  result := -1 ;

  // Acc Actif ?
  if GuideActif then exit ;
  if not FBoAccActif then Exit ;
  if FBoAccArret then Exit ;

  if not ( (vRow = Piece.Detail.count) or
           ( (vRow = Piece.Detail.count-1) and not Piece.EstRemplit(vRow+1) )
         )
   then Exit ;

  // Vérif Journal
  if Info.GetString('J_ACCELERATEUR') <> 'X' then Exit ;

  // Vérif saisie ligne complète
  if not Piece.EstRemplit( vRow ) then Exit ;

  // en mode BOR, uniquement sur la 1ère ligne d'un folio
  if Piece.ModeSaisie=msBOR then
    if Piece.GetDebutGroupe( vRow ) <> vRow then Exit ;

  // Déjà traité ?
  if Piece.GetString( vRow, 'ACCELERATEUR') = 'X' then Exit ;

  // Vérif Auxiliaire
  lStAux := Piece.GetString( vRow, 'E_AUXILIAIRE' ) ;
  if lStAux = '' then Exit ;
  if not Info.LoadAux( lStAux ) then Exit ;

  if (Info.GetString('YTC_ACCELERATEUR') <> 'X') and (not vBoForce) then exit ;

  // Vérif du montant
  lStCol := GetColName( Col ) ;
  if (lStCol = 'E_DEBITDEV') or (lStCol = 'E_CREDITDEV') then
    ValideCellule( lStCol, vRow ) ;
  if not Piece.IsValidDebitCredit( vRow ) then Exit ;

  // Récupération compte de contrepartie
  lStCptHt := Piece.GetCompteAcc( vRow, FStAccCptHT ) ;
  if lStCptHT = '' then Exit ;
  if Info.LoadCompte(lStCptHT) and (info.GetString('G_FERME')='X') then
    begin
    PGIInfo( TraduireMemoire('Le compte général paramétré en accélérateur de saisie est fermé. L''accélérateur de saisie est temporairement désactivé pour ce compte'));
    if Info.LoadAux( lStAux ) then
      Info.Aux.Item.PutValue('YTC_ACCELERATEUR', '-' ) ;
    Exit ;
    end ;

  // Pour éviter les appels récursifs
  FBoAccActif := False ;

  // Insertion de la ligne de HT :
  Piece.InsereLigneHT( vRow, lStCptHT, 0 ) ;
  lInRow := Piece.CurIdx ;
  SetPos( GetColIndex('E_GENERAL'), lInRow, False ) ;
  AfficheLignes( lInRow ) ;

  // Gestion des évènements
  if ValideSortieLigne( False ) then
    begin

    Piece.InsereLigneTva( lInRow, vRow ) ;
    lInRow := Piece.CurIdx ;
    SetPos( GetColIndex('E_GENERAL'), lInRow, False ) ;
    AfficheLignes( lInRow ) ;

    ValideSortieLigne(False) ;

    end ;

  Piece.PutValue( vRow, 'ACCELERATEUR', 'X' );

  // repositionnement : fin de la ligne si Viewer activer
  if ViewerVisible
    then SetPos( GetColIndex('E_CREDITDEV'), lInRow, False,True,-1)
    // ligne suivante sinon
    else ProchaineLigne( lInRow, '', False ) ;

  AfficheLignes ;

  result      := lInRow + 1 ;
  FBoAccActif := True ;

end;

function TSaisiePiece.IsActiveAcc: boolean;
begin
  result := FBoAccActif and ( Info.GetString('J_ACCELERATEUR') = 'X' ) ;
end;

procedure TSaisiePiece.SetEnabled(vBoEna: Boolean);
begin
  FListe.Enabled := vBoEna ;
end;

function TSaisiePiece.GetEnabled: boolean;
begin
  result := FListe.Enabled ;
end;


procedure TSaisiePiece.GereGuide(vRow: integer);
begin

  if GetActionFiche = taConsult then Exit ;
  if Piece.IsOut( vRow )        then Exit ;
  if GuideActif                 then Exit ;
  if Piece.EstRemplit( vRow )   then Exit ;
  if (Piece.GetSolde <> 0)      then exit ;

  if TSGuide.RechercherGuide( vRow ) then
    TSGuide.DemarrerGuide( vRow ) ;

end;

function TSaisiePiece.GuideActif: boolean;
begin
  result := Assigned(TSGuide) and TSGuide.GuideActif ;
end;

procedure TSaisiePiece.LanceGuide(vRow: integer);
begin
  if GetActionFiche = taConsult then Exit ;
  if Piece.IsOut( vRow )        then Exit ;
  if GuideActif                 then Exit ;
  if GetLastGuide=''            then Exit ;
  if Piece.EstRemplit( vRow )   then Exit ;
  if (Piece.GetSolde <> 0)      then exit ;

  TSGuide.DemarrerGuide( vRow ) ;

end;

function TSaisiePiece.GetLastGuide: string;
begin
  result := '' ;
  if not Assigned(TSGuide) then Exit ;
  result := TSGuide.GetLastGuide ;
end;
{
function TSaisiePiece.ComputeCell(vStArg: hstring): Variant;
begin
  result := '' ;
end;
}
procedure TSaisiePiece.SetInitCell;
var lRow : integer ;
    lCol : integer ;
begin
  {JP 20/07/07 : FQ 20601 : on se positionne éventuellement sur une ligne déterminée}
  if Piece.NumLigneAppel > 0 then
    lRow := Piece.NumLigneAppel

  else if GetActionFiche = taConsult
    then lRow := 1
    else lRow := Piece.Detail.count ;

  if (Piece.ModeSaisie = msPiece)
    then lCol := GetColIndex('E_GENERAL')
    else begin
         lCol := GetColIndex('E_DATECOMPTABLE') ;
         if lCol < 1 then
           lCol := GetColIndex('E_GENERAL') ;
         end ;

  SetPos( lCol, lRow, False ) ;

end;


procedure TSaisiePiece.paramViewer( vBoReinit : Boolean ) ;
var
  TobViewer : TOB;
begin

  TobViewer := GetTobMasqueParam;

  if Assigned(TSDoc) then
    begin

    TSDocInt.ActiverViewer( TobViewer , TSDoc = TSDocInt ) ;
    TSDocExt.ActiverViewer( TobViewer , TSDoc = TSDocExt ) ;

    if assigned(ViewerExt) then
       ViewerExt.Visible := TSDocExt.IsVisible and (FViewer = 'EXT');

    // Synchronisation TSDoc avec Pièce en cours
    TSDoc.Piece := Piece ;
    // Affectation du groupe en cours d'édition pour gestion des docs
    if Piece.ModeSaisie = msBor
      then TSDoc.GroupeActif := Piece.GetInteger( Piece.Count, 'E_NUMGROUPEECR')
      else TSDoc.GroupeActif := 0 ;

    // repositionnement du web browser initiale
    if TSDoc.IsActif and TSDoc.IsVisible then
      Case TSDoc.Position Of
        vpBas, vpHaut       :  TSDoc.PanelActif.Height := (LaGrille.Height + TSDoc.PanelActif.Height) div 2  ;
        vpDroite, vpGauche  :  TSDoc.PanelActif.Width  := (LaGrille.Width + TSDoc.PanelActif.Width) div 2  ;
        end; // case

    if vBoReinit then
      TSDoc.DesactiveViewer
    else
      begin
      if (FViewer = 'EXT') and TSDocExt.IsVisible then
        begin
        TSDocExt.ChargeDocument(false,true);
        SetDocActif(TSDocExt);
        FViewerExt.TSDocInt := TSDocInt;
        FViewerExt.TSDocExt := TSDocExt;
        FViewerExt.Show;
        end
      else
        ChargeDocument( False ) ;
      end ;

    end ;

  // maj graphique
  LaGrille.refresh;
  ResizeGrille ;

end;


procedure TSaisiePiece.InitViewer( vPanelB, vPanelR, vPanelL, vPanelT: THPanel;
                                   vSplitB, vSplitR, vSplitL, vSplitT: THSplitter);
begin

  if not assigned(TSDocInt) then
  begin
    FViewer := GetSynRegKey('PosVisu'   ,'INT'   ,true);
    if not assigned(FTSDocInt) then
       FTSDocInt := TSaisieDoc.Create( FEcran, FTobPiece ) ;
    if not assigned(FTSDocExt) then
       FTSDocExt := TSaisieDoc.Create( FEcran, FTobPiece ) ;
    if FViewer = 'EXT' then
    begin
       SetDocActif(FTSDocExt);
       CreateExternalViewer;
    end
    else
    begin
       SetDocActif(FTSDocInt);
    end;
  end;
  
  // Référencement des composants pour le gestionnaire du viewer
  TSDocInt.InitViewer( vPanelB, vPanelR, vPanelL, vPanelT, vSplitB, vSplitR, vSplitL, vSplitT ) ;
  // évènement mixte grille / viewer
  vSplitL.OnMoved := OnSplitterMoved ;
  vSplitR.OnMoved := OnSplitterMoved ;
end;

procedure TSaisiePiece.OnSplitterMoved(Sender: TObject);
begin
  // Si on déplace un splitter, il faut gérer le redimensionnement d
  // es cellules de la grille et l'apparition d'un éventuel ascenceur
  if THSPlitter(Sender).Visible then
    begin
    if (TSDoc.Position in [ vpDroite, vpGauche]) then
      begin
      // pour forcer le recalcul de l'affichage d'un éventuel ascenceur horizontal
      LaGrille.ColCount := LaGrille.ColCount + 1 ;
      LaGrille.ColCount := LaGrille.ColCount - 1 ;
      end
    else if (TSDoc.Position in [ vpHaut, vpBas]) then
      begin
      // pour forcer le recalcul de l'affichage d'un éventuel ascenceur vertical
      LaGrille.RowCount := LaGrille.RowCount + 1 ;
      LaGrille.RowCount := LaGrille.RowCount - 1 ;
      end ;
    LaGrille.Invalidate ;
    end ;
  ResizeGrille;
end;

function TSaisiePiece.GetTobMasqueParam: Tob;
begin
  result := FTobMasque.FindFirst( ['TYPE'], ['PARAMETRES'], True ) ;
  if not Assigned(result) then
    result := CInitTobMasqueParam( FTobMasque ) ;
end;

function TSaisiePiece.ViewerActif: boolean;
begin
  result := Assigned(TSDoc) and TSDoc.IsActif ;
end;

procedure TSaisiePiece.InitPopupViewer(vPopup: TPopupMenu ; vBouton : TToolBarButton97 ; vBoutonMove : TToolBarButton97 ) ;
begin
  TSDocInt.SetMenuPop( vPopup, vBouton , vBoutonMove) ;
  vPopup.OnPopup := TSDocInt.Onpopupviewer ;
  TSDocInt.InitPopupViewer(False);
end;

procedure TSaisiePiece.InitBDoc( vBExt : TToolBarButton97 ; vBInt : TToolBarButton97 ) ;
begin
  TSDocInt.SetBDoc( vBExt, vBInt);
end;


function TSaisiePiece.ViewerVisible: Boolean;
begin
  result := ViewerActif and TSDoc.IsVisible ;
end;

procedure TSaisiePiece.ChargeDocument(vBoSuivant: boolean) ;
var lBoDansGrille : boolean ;
    lControlFocus : TWinControl ;
    lBoViewerChg  : boolean ;
begin
  if FBoDocMAJEnCours then Exit ;
  FBoDocMAJEnCours := True ;
  lControlFocus    := Screen.ActiveControl ;
  lBoDansGrille    := ( lControlFocus.Name = LaGrille.Name ) ;
  lBoViewerChg     := ( ViewerActif and ViewerVisible ) ;
  // chargement du document
  TSDoc.ChargeDocument( vBoSuivant );
  // gestion de l'affichage
  if ViewerActif then
     if ViewerVisible then
     begin
        Case TSDoc.Position Of
            vpDroite, vpGauche  :  if TSDoc.FitToLargeOk then
                                 TSDoc.ViewerFitLargeClick(nil) ;
     end ;
        //    TSDoc.EnBasADroite ;
  end
  else
  begin
     if lBoViewerChg then // changement d'état du viewer, il faut faire un resize de la grille
        Case TSDoc.Position Of
           vpDroite, vpGauche : resizeGrille ;
        end ;
        //    TSDoc.EnBasADroite ;
  end ;
  // reprise du focus
  if lBoDansGrille then
  begin
     SetEvtOff ;
     LaGrille.SetFocus ;
     SetEvtOn ;
 end
 else if (lControlFocus is THValComboBox) then
    lControlFocus.SetFocus ;
 FBoDocMAJEnCours := False ;
end;

function TSaisiePiece.MajDocEnCours: Boolean;
begin
  result := FBoDocMAJEnCours ;
end;

procedure TSaisiePiece.SetDocExterne         ( pDoc : TSaisieDoc );
begin
  FTSDocExt := pDoc;
end;

procedure TSaisiePiece.SetDocInterne         ( pDoc : TSaisieDoc );
begin     
  FTSDocInt := pDoc;
end;

procedure TSaisiePiece.SetDocActif           ( pDoc : TSaisieDoc );
begin
  // transfert des informations de gestion avant permutation des objets d'affichage
  if Assigned( TSDoc ) then
    begin
    pDoc.GroupeActif := TSDoc.GroupeActif ;
    pDoc.Fichier     := TSDoc.Fichier ;
    pDoc.Chemin      := TSDoc.Chemin ;
    end ;
//  pDoc.FichierOk   := TSDoc.FichierOk ;
  // Transfert de l'objet actif
  FTSDoc           := pDoc;
end;

procedure TSaisiePiece.ViewerKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var lBoGereFocus : Boolean ;
begin                  
  if ViewerActif and ViewerVisible and (Shift=[ssAlt]) then
    begin

    lBoGereFocus := ( Screen.ActiveControl = FListe ) ;

    if lBoGereFocus then
      SetEvtOff ;

    Case Key Of
      VK_END : begin
               TSDoc.EnBasADroite ;
               Key:=0 ;
               end ;
      VK_HOME : begin
               TSDoc.EnHautAGauche ;
               Key:=0 ;
               end ;
      // Document précédent
      33       : begin
                 TSDoc.ViewerPrevDocClick(Sender) ;
                 Key:=0 ;
                 end ;
      // Document suivant
      34       : begin
                 TSDoc.ViewerNextDocClick(Sender) ;
                 Key:=0 ;
                 end ;
    // flèche de déplacement
      37       : begin
                 TSDoc.AGauche( False ) ;
                 Key:=0 ;
                 end ;
      38       : begin
                 TSDoc.EnHaut( False ) ;
                 Key:=0 ;
                 end ;
      39       : begin
                 TSDoc.ADroite( False ) ;
                 Key:=0 ;
                 end ;
      40       : begin
                 TSDoc.EnBas( False ) ;
                 Key:=0 ;
                 end ;
    // Zoom : "+" = zoom avant
      109      : begin
                 TSDoc.ZoomImage( True ) ;
                 Key:=0 ;
                 end ;
    // Zoom : "-" = zoom arrière
      107      : begin
                 TSDoc.ZoomImage( False ) ;
                 Key:=0 ;
                 end ;
    // * = retourner le document
      106      : begin
                 TSDoc.RotateImage ;
                 Key:=0 ;
                 end ;
    end ;
    if lBoGereFocus then
      begin
      if (Key=0) then
        FListe.SetFocus ;
      SetEvtOn ;
      end ;

    end ;

end;

procedure TSaisiePiece.CreateExternalViewer;
begin
  if not assigned(FViewerExt) then
     FViewerExt := TViewerExt.Create(Application);
  FViewerExt.TSDocExt := FTSDocExt;
  FViewerExt.TSDocInt := FTSDocInt;
  FViewerExt.TSP      := Self;
  FViewerExt.InitViewer;
end;
{ TSaisieGuide }

constructor TSaisieGuide.Create( vTSP : TSaisiePiece ; vControl : TControl = nil ) ;
begin
  inherited Create ;

  FTSPiece   := vTSP ;
  InitFlashGuide ( vControl ) ;
  
end;

destructor TSaisieGuide.Destroy;
begin

  if Assigned( FTobGuide ) then
    FreeAndNil( FTobGuide ) ;

  if Assigned( FFlashLabel ) then
    FreeAndNil( FFlashLabel ) ;

  inherited;

end;

function TSaisieGuide.GetFListe: THGrid;
begin
  result := TSP.LaGrille ;
end;

function TSaisieGuide.GetRechCompte: TRechercheGrille;
begin
  result := TSP.RechCompte ;
end;

function TSaisieGuide.GetPiece: TPieceCompta;
begin
  result := TSP.Piece ;
end;

function TSaisieGuide.EstColEditable(vInRow: integer; vStColName: String): Boolean;
begin
  result := False ;

  if ( vStColName = 'E_AUXILIAIRE' ) and
     Piece.Info.LoadCompte( Piece.GetString( vInRow , 'E_GENERAL') ) and
     not ( Piece.Compte_GetValue('G_COLLECTIF') = 'X' ) then Exit ;

  result := ( vStColName = 'E_GENERAL' )   or
            ( vStColName = 'E_AUXILIAIRE' ) or
            ( vStColName = 'E_REFINTERNE' ) or
            ( vStColName = 'E_LIBELLE' )    or
            ( ( vStColName = 'E_DEBITDEV' ) and (Piece.GetDouble(vInRow, 'E_CREDITDEV')=0))   or
            ( ( vStColName = 'E_CREDITDEV' ) and (Piece.GetDouble(vInRow, 'E_DEBITDEV')=0))   ;
  result := result and GetArretPourCol( vInRow, vStColName ) ;

end;

function TSaisieGuide.GetColSuivante(var Acol, ARow: integer ) : boolean ;
var i,j         : integer ;
    lStArret    : string ;
    lTOBGuide   : TOB  ;
    lCancel     : boolean ;
    lChg        : boolean ;
begin

  result := true ;

  // il faut partir de l'index correspondant aux colonnes de la grille des guides
  Acol := GetIdxArret( Acol ) ;
  Inc( Acol ) ;

  // parcours de la ligne courante à la dernière
  for i := ARow to FInFin do
    begin
    lTOBGuide := GetGuideLigne( i ) ;
    if lTOBGuide = nil then exit ;

    // parcours des indicateurs d'arrêt...
    for j := ACol to 6 do
      begin
      lStArret := copy( lTOBGuide.GetString('EG_ARRET') ,1, 6); // la grille de saisie ne comporte que 6 cases
      if ( length(lStArret) >= ACol ) and ( lStArret <> '' ) and ( trim(lStArret[j]) = 'X' ) then
        begin
        ARow      := i ;
        ACol      := FTabGuideCol[ j ] ;
        exit ;
        end ;
      end ;

    // Pas trouvé sur la ligne en cours on s'apprête a recommencé au début de la ligne suivante
    // --> déclenchement des évt de fin de ligne
    lCancel := False ;
    lChg    := False ;
    TSP.RowExit( FListe, i, lCancel, lChg ) ;

    // --> reinit du numéro de colonne
    ACol := 1 ;

    end;

   // il n'y a plus de point d'arret, on arrête le guide
   FBoFinGuide := True ;
   ACol        := FTabGuideCol[ 1 ];
   ARow        := FInDeb ;

end;

function TSaisieGuide.GetNomCol(vIdxArret: integer): string;
begin
  Case vIdxArret of
    1 : result := 'E_GENERAL' ;
    2 : result := 'E_AUXILIAIRE' ;
    3 : result := 'E_REFINTERNE' ;
    4 : result := 'E_LIBELLE' ;
    5 : result := 'E_DEBITDEV' ;
    6 : result := 'E_CREDITDEV' ;
    else result := '' ;
    end ;
end;

procedure TSaisieGuide.ActiveGuide( vBoOn : boolean ) ;
begin
  if vBoOn then
    begin
    // paramétrage particulier du TPieceCompta
    FOldModeEche           := Piece.ModeEche ;
    Piece.SetMultiEcheAucun ;
    Piece.ModeGroupe := mgStatic ;
    end
  else
    begin
    // paramétrage particulier du TPieceCompta
    Case FOldModeEche of
       meMulti   : Piece.SetMultiEcheMulti ;
       meMono    : Piece.SetMultiEcheOff ;
       meDeporte : Piece.SetMultiEcheDeporte ;
       else      Piece.SetMultiEcheAucun ;
     end ;
    Piece.ModeGroupe := mgDynamic ;
    end ;

  if Assigned( FlashLabel ) then
    begin
    FlashLabel.Flashing  := vBoOn ;
    FlashLabel.Visible   := vBoOn ;
    end ;

  FBoGuideActif        := vBoOn ;

end;

procedure TSaisieGuide.InitFlashGuide( vControl : TControl ) ;
begin
  if Assigned( vControl ) then
    begin
    if not Assigned( FlashLabel ) then
      FFlashLabel               := TFlashingLabel.Create( TSP.Ecran ) ;
    FlashLabel.Parent           := vControl.Parent ;
    FlashLabel.Left             := vControl.Left ;
    FlashLabel.Height           := vControl.Height ;
    FlashLabel.Top              := vControl.Top ; //+ 18 ;
    FlashLabel.Width            := vControl.Width ; //+ 18 ;
    FlashLabel.Caption          := 'GUIDE EN COURS' ;
    FlashLabel.Transparent      := true ;
    FlashLabel.Color            := clRed ;
    FlashLabel.Font.Style       := [fsBold] ;
    FlashLabel.Visible          := false ;
    end ;
end;

procedure TSaisieGuide.SetFlashGuide( vControl: TControl );
begin
  InitFlashGuide ( vControl ) ;
end;

procedure TSaisieGuide.DemarrerGuide( vRow : integer );
var lInCol : integer ;
    lInRow : integer ;
begin

  FBoFinGuide := False ;
  InitTabGuide ;

  // Modif du comportement de la grille
  ActiveGuide( True ) ;

  // Calcul des lignes
  CalculGuide( vRow ) ;

  // Affichage
  AfficheLignes ;

  // pour déplacement des ascenceur de la grille
  TSP.SetPos( 1, FInFin, False, False, 1);
  // Détermination 1ère zone
  lInCol := 0 ; // TSP.GetColIndex('E_GENERAL') ;
  lInRow := vRow ;
  GetColSuivante( lInCol, lInRow ) ;
  TSP.SetPos( lInCol, lInRow, False, False, 1);

//  FListe.Col       := lACol ;
//  FListe.Row       := lInRowRef ;


end;

function TSaisieGuide.RechercherGuide(vRow: integer): boolean;
var lTOBLookUp       : TOB ;
    lTOBListeGuide   : TOB;
    lStSql           : string ;
    lStCol           : string ;
    lStLib           : string ;
    lStGuide         : string ;
    lStTitre         : string ;
begin
  result := False ;
  if Piece.IsOut( vRow ) then exit ;

  lStSql     := GetSQL( vRow ) ;
  if (lStSql = FStLastSQL) and not FBoLastRech then Exit ;

  lStCol := 'EG_GUIDE;GU_LIBELLE;' ;
  lStLib := 'Code;Libellé;' ;

  lTOBListeGuide := TOB.Create('LISTE_GUIDE', nil,-1) ;
  lTOBListeGuide.LoadDetailFromSQL(lStSQL) ;

  FStLastSQL  := lStSQL ;
  FBoLastRech := (lTOBListeGuide.Detail.Count>0) ;

  if not FBoLastRech then
    exit ;

  if Piece.GetString( vRow, 'E_GENERAL')<>''
    then lStTitre := 'Guide pour : ' + Piece.GetString( vRow, 'E_GENERAL')
    else lStTitre := 'Guide de saisie' ;

  if lTOBListeGuide.Detail.Count = 1
    then lStGUIDE := lTOBListeGuide.Detail[0].GetString('EG_GUIDE')
    else begin
         lTOBLookUp := LookUpTob ( FListe ,
                                   lTOBListeGuide ,
                                   lStTitre,
                                   lStCol ,
                                   lStLib );

         if assigned(lTOBLookUp) then
           lStGUIDE := lTOBLookUp.GetString('EG_GUIDE') ;
      end; // if

 lTOBListeGuide.Free ;

 if (lStGUIDE <> '') then
   result := Load( ['NOR',lStGUIDE] ) ;

end;

procedure  TSaisieGuide.FinirGuide( vBoAnnul : boolean );
var i    : integer ;
    lRow : integer ;
begin

  if not FBoGuideActif then Exit ;

  lRow := FInFin ;
  ModeSilence( True ) ;

  // Affectation de la Tva
  Piece.FinaliseRegimeEtTva( FInDeb ) ;

  for i := FInFin downto FInDeb do
    begin
    if Piece.GetInteger( i, 'E_NUMECHE') > 1 then Continue ;
    if vBoAnnul or not Piece.IsValidLigne(i) then
      begin
      TSP.DeleteRow(i, True) ;
      Dec(lRow) ;
      end ;
    end ;
  ModeSilence( False ) ;

  // TMultiEche en mode modif ;
  if Piece.Echeances.detail.count=1 then
    begin
    TMultiEche( Piece.Echeances.detail[0] ).Action := taModif ;
    end ;

  // Retour en mode normal
  ActiveGuide( False ) ;

  // Replacement du curseur
  TSP.ProchaineLigne( lRow, '', False ) ;

end;


function TSaisieGuide.CalculGuide ( vRow : integer ) : boolean ;
var lTobEcrGui : TOB ;
    lInIndex   : integer ;
    i          : integer ;
    lTob       : TOB ;
    lInGrp     : integer ;
begin

  result       := false;

  Guide.Detail.Sort('EG_NUMLIGNE') ;

  FInDeb   := -1 ;
  lInIndex := -1 ;

  lInGrp := Piece.GetInteger( vRow, 'E_NUMGROUPEECR' ) ;

  for i := 0 to Guide.Detail.count - 1 do
    begin

    // Récup ligne de guide
    lTobEcrGui := Guide.Detail[i] ;

    // préparation du TPieceCompta
    lInIndex := i + vRow  ;
    if Piece.isOut( lInIndex ) or Piece.EstRemplit( lInIndex )
      then lTob := Piece.NewRecord( lInIndex )
      else lTob := Piece.GetTob( lInIndex ) ;

    // Affectation manuelle du groupe
    lTob.PutValue('E_NUMGROUPEECR', lInGrp ) ;

    lInIndex := Piece.CurIdx ;
    if FInDeb < 0 then
      FInDeb := lInIndex ;

    // Calcul de la ligne
    CalculLigne( lInIndex , lTobEcrGui ) ;

    lTOB.AddChampSupValeur('EG_GUIDE'          , lTobEcrGui.GetString('EG_GUIDE') ) ;
    lTOB.AddChampSupValeur('EG_NUMLIGNE'       , lTobEcrGui.GetInteger('EG_NUMLIGNE') ) ;

    end;

  FInFin := lInIndex ;

end;

procedure TSaisieGuide.AbandonnerGuide;
begin
 if PGIAskCancel( TraduireMemoire('Voulez-vous annuler le guide de saisie ?'), TSP.Ecran.Caption) = mrYes then
   begin
   // suppression des lignes du guide en cours
   FinirGuide( True ) ;
   end ;
end;

function TSaisieGuide.ExisteGuide(vRow: integer): boolean;
var lStSQL : string ;
begin
  lStSQL := GetSQL( vRow ) ;
  if lStSQL = FStLastSQL
    then result := FBoLastRech
    else result := ExisteSQL( lStSQL ) ;
  FStLastSQL  := lStSQL ;
  FBoLastRech := result ;
end;

function TSaisieGuide.GetSQL(vRow: integer): string;
var lTob : Tob ;
begin
  result := '' ;
  if Piece.IsOut( vRow ) then exit ;

  lTob := Piece.GetTob( vRow ) ;

  result := 'SELECT GU_JOURNAL,GU_LIBELLE,GU_NATUREPIECE,EG_GUIDE,EG_GENERAL '
           + ' FROM ECRGUI, GUIDE '
           + ' WHERE EG_NUMLIGNE=1 AND GU_TYPE="NOR" AND GU_TRESORERIE<>"X" AND GU_GUIDE=EG_GUIDE '
               + ' AND GU_JOURNAL="' + lTOB.GetString('E_JOURNAL') + '" '
               + ' AND ( ( GU_NATUREPIECE="' + lTOB.GetString('E_NATUREPIECE') + '" ) OR '
                     + ' ( GU_NATUREPIECE="" ) OR ( GU_NATUREPIECE IS NULL ) )'
               + ' AND ( ( GU_ETABLISSEMENT="' + lTOB.GetString('E_ETABLISSEMENT') + '") OR '
                     + ' ( GU_ETABLISSEMENT="" ) OR ( GU_ETABLISSEMENT IS NULL ) )' ;

  if lTOB.GetString('E_GENERAL') <> '' then
    result := result + ' AND EG_GENERAL="' + lTOB.GetString('E_GENERAL') + '" ' ;

end;

function TSaisieGuide.CalculLigne(vRow: integer; lTobEcrGui : Tob) : boolean ;
var lSt      : string ;
    lStTob   : string ;
    lMontant : double ;
    lTOB     : Tob ;
begin

  result := False ;

  // Compte Général
  lTob   := Piece.GetTob( vRow ) ;
  lStTob := lTob.GetString('E_GENERAL') ;
  if not Piece.Info.LoadCompte( lStTob ) then
    begin
    lSt:=lTobEcrGui.GetString('EG_GENERAL') ;
    if (lSt<>'') then
      begin
      if EstFormule(lSt) then
        lSt := GFormule(lSt,Piece.GetFormule,Nil,1) ;
      if Piece.Info.LoadCompte( lSt )
        then Piece.PutValue(vRow, 'E_GENERAL', Piece.Info.StCompte, True)
        else Piece.Detail[vRow-1].PutValue('E_GENERAL', lSt);
      result := True ;
      end ;
    end ;

  // Compte auxiliaire
  lTob   := Piece.GetTob( vRow ) ;
  lStTob := lTob.GetString('E_AUXILIAIRE') ;
  if not Piece.Info.LoadAux( lStTob ) then
    begin
    lSt:=lTobEcrGui.GetString('EG_AUXILIAIRE') ;
    if (lSt<>'')  then
      begin
      if EstFormule(lSt) then
        lSt := GFormule( lSt, Piece.GetFormule, Nil, 1 ) ;
      if Piece.Info.LoadCompte( lSt )
        then Piece.PutValue(vRow, 'E_AUXILIAIRE', Piece.Info.StAux )
        else Piece.Detail[ vRow - 1 ].PutValue('E_AUXILIAIRE', lSt) ;
      result := True ;
      end ;
    end ;

  // Référence interne
  Piece.GetTob( vRow ) ;
  lSt    := lTobEcrGui.GetString('EG_REFINTERNE') ;
  if lSt<>'' then
    begin
    if ( (EstFormule(lSt) ) or ( Piece.GetString(vRow, 'E_REFINTERNE') = '' ) ) then
      begin
      lSt := GFormule( lSt, Piece.GetFormule, Nil, 1) ;
      lSt := Trim( lSt ) ;
      if (lSt<>'') and (lSt <> Piece.GetString(vRow, 'E_REFINTERNE') ) then
        begin
        Piece.PutValue( vRow, 'E_REFINTERNE', lSt ) ;
        result := True ;
        end ;
      end ;
    end ;

  // Libellé
  Piece.GetTob( vRow ) ;
  lSt    := lTobEcrGui.GetString('EG_LIBELLE') ;
  if lSt<>'' then
    begin
    if ( (EstFormule(lSt) ) or ( Piece.GetString(vRow, 'E_LIBELLE') = '' ) ) then
      begin
      lSt := GFormule( lSt, Piece.GetFormule, Nil, 1) ;
      lSt := Trim( lSt ) ;
      if (lSt<>'') and (lSt <> Piece.GetString(vRow, 'E_LIBELLE') ) then
        begin
        Piece.PutValue( vRow, 'E_LIBELLE', lSt ) ;
        result := True ;
        end ;
      end ;
    end ;

  // DEBIT
  lTOB   := Piece.GetTob( vRow ) ;
  lSt    := lTobEcrGui.GetString('EG_DEBITDEV') ;
  if lSt<>'' then
    begin
    lSt := GFormule( lSt, Piece.GetFormule, Nil, 1 ) ;
    if (lSt = 'SOLDE') then
      begin
      if Piece.GetSoldePartiel(FInDeb, FInFin) <> 0 then
        begin
        Piece.AttribSolde( vRow ) ;
        result := True ;
        end ;
      end
    else
      begin
      lMontant := Valeur( lSt ) ;
      if (lMontant <> 0) and ( Arrondi(lTOB.GetDouble('E_DEBITDEV')-lMontant, Piece.GetRDevise(vRow).Decimale ) <> 0 ) then
        begin
        Piece.PutValue(vRow, 'E_DEBITDEV', Valeur( lSt ), True ) ;
        result := True ;
        end ;
      end
    end ;

  // CREDIT
  lTOB := Piece.GetTob( vRow ) ;
  lSt  := lTobEcrGui.GetString('EG_CREDITDEV') ;
  if lSt<>'' then
    begin
    lSt := GFormule( lSt, Piece.GetFormule, Nil, 1 ) ;
    if (lSt = 'SOLDE') then
      begin
      if Piece.GetSoldePartiel(FInDeb, FInFin) <> 0 then
        begin
        Piece.AttribSolde( vRow ) ;
        result := True ;
        end ;
      end
    else
      begin
      lMontant := Valeur( lSt ) ;
      if (lMontant <> 0) and ( Arrondi(lTOB.GetDouble('E_CREDITDEV')-lMontant, Piece.GetRDevise(vRow).Decimale ) <> 0 ) then
        begin
        Piece.PutValue(vRow, 'E_CREDITDEV', Valeur( lSt ), True ) ;
        result := True ;
        end ;
      end ;
    end ;

  // TVA
  lSt  := lTobEcrGui.GetString('EG_TVA') ;
  if lSt <> '' then
    if Piece.GetString( vRow, 'E_TVA' ) <> lTobEcrGui.GetString('EG_TVA') then
      begin
      Piece.PutValue( vRow, 'E_TVA',    lTobEcrGui.GetString('EG_TVA')     ) ;
      result := True ;
      end ;

  lSt  := lTobEcrGui.GetString('EG_TVAENCAIS') ;
  if lSt <> '' then
    if Piece.GetString( vRow, 'E_TVAENCAISSEMENT' ) <> lTobEcrGui.GetString('EG_TVAENCAIS') then
      begin
      Piece.PutValue( vRow, 'E_TVAENCAISSEMENT',    lTobEcrGui.GetString('EG_TVAENCAIS')     ) ;
      result := True ;
      end ;

  if ( Piece.GetString( vRow, 'E_ECHE' ) = 'X') and (Piece.GetString( vRow, 'E_MODEPAIE')='') then
    begin
    // tests préalable
    if //    compte collectif + tiers ok
       (( Piece.Info.Compte.IsCollectif and Piece.Info.LoadAux( Piece.GetString( vRow, 'E_AUXILIAIRE') ) ) or
       // OU comte non collectif ok
       ( not Piece.Info.Compte.IsCollectif and Piece.Info.LoadCompte( Piece.GetString( vRow, 'E_GENERAL') ) ) )
       // ET montant ok
       and Piece.IsValidDebitCredit( vRow ) then
      begin
      Piece.SetMultiEcheOff ; // uniquement mode mono supporté en guide pour le moment
      if ( lTobEcrGui.GetString('EG_MODEREGLE') <> '' )
        then Piece.CalculEche( vRow, lTobEcrGui.GetString('EG_MODEREGLE') )
        else if Piece.GetString( vRow, 'E_MODEPAIE')='' then
          Piece.CalculEche( vRow ) ;
      result := result or ( Piece.GetString( vRow, 'E_MODEPAIE' ) <> '' ) ;
      Piece.SetMultiEcheAucun ; // retour en mode sans calcul pour guide
      end ;
    end ;

  // Gestion de l'analytique
  lTOB   := Piece.GetTob( vRow ) ;
  lStTob := lTOB.GetString( 'E_GENERAL') ;
  if Piece.Info.LoadCompte( lStTob ) and ( Piece.GetString(vRow, 'E_ANA') = 'X' ) then
    if Piece.IsValidDebitCredit( vRow ) then
      if not VentilationExiste( lTob ) then
        begin
        // 1er essai sur Guide ANA
        VentilerTob( lTOb,
                     Guide.GetString('GU_GUIDE'),
                     lTob.GetInteger('EG_NUMLIGNE'),
                     Piece.GetRDevise(vRow).Decimale,
                     False,
                     Piece.Dossier,
                     False ) ;
        // 2ème essai sur ventil par defaut
        if not VentilationExiste( lTob ) then
          VentilerTob( lTOb,
                       '',
                       0,
                       Piece.GetRDevise(vRow).Decimale,
                       False,
                       Piece.Dossier,
                       False ) ;
        result := True ;
        end ;

end;

function TSaisieGuide.RecalculGuide(vRow: integer): boolean;
var i        : integer ;
begin
  result := False ;
  for i := FInDeb to FInFin do
    begin
    if Piece.GetInteger( i, 'E_NUMECHE' )>1 then Continue ;
    if CalculLigne( i, GetGuideLigne( i ) ) then
      result := True ;
    end ;

  if result then
    AfficheLignes ;

end;

function TSaisieGuide.Estformule(vSt: string): boolean;
begin
  Result := Pos('{',vSt)>0 ;
  if Not Result then Result:=Pos('}',vSt)>0 ;
  if Not Result then Result:=Pos('[',vSt)>0 ;
  if Not Result then Result:=Pos(']',vSt)>0 ;
end;

function TSaisieGuide.GetIdxArret(vCol: integer; vStColName : string = ''): integer;
var lStChp : string ;
begin

  result := 0 ;
  if vStColName<>''
    then lStchp := vStColName
    else lStChp := TSP.GetColName( vCol ) ;

  if lStChp= 'E_GENERAL'
    then result := 1
  else if lStchp = 'E_AUXILIAIRE'
    then result := 2
  else if lStchp = 'E_REFINTERNE'
    then result := 3
  else if lStchp = 'E_LIBELLE'
    then result := 4
  else if lStchp = 'E_DEBITDEV'
    then result := 5
  else if lStchp = 'E_CREDITDEV'
    then result := 6 ;
    
end;

function TSaisieGuide.EstGuideCorrect: boolean;
var i : integer ;
begin
  result := False ;
  ModeSilence( True ) ;
  for i:=FInDeb to FInFin do
    result := result and Piece.IsValidLigne(i) ;
  ModeSilence( False ) ;
end;

procedure TSaisieGuide.ModeSilence(vBoOn: boolean);
begin
  if vBoOn
    then Piece.OnError := nil
    else Piece.OnError := TSP.OnError ;
end;

function TSaisieGuide.Load(const Values: array of string): boolean;
var lStType : string ;
    lStCode : string ;
    lStSQL  : string ;
begin

  result := False ;

  if (SizeOf(Values)=0) then exit ;

  lStType := Values[0] ;
  lStCode := Values[1] ;

  if not Assigned( FTobGuide ) then
    FTobGuide := Tob.Create('GUIDE', nil, -1);

  if ( FTobGuide.GetString('GU_TYPE') = lStType ) and
     ( FTobGuide.GetString('GU_GUIDE') = lStCode ) then
     begin
     result := True ;
     Exit ;
     end ;

  FTobGuide.PutValue('GU_TYPE', lStType);
  FTobGuide.PutValue('GU_GUIDE', lStCode);

  result := FTobGuide.LoadDB ;
  if result then
    begin
    lStSQL := 'SELECT * FROM ECRGUI WHERE EG_TYPE="' + lStType + '" AND EG_GUIDE="' + lStCode + '"' ;
    FTobGuide.LoadDetailFromSQL(lStSQL) ;
    FTobGuide.Detail.sort('EG_NUMLIGNE') ;
    end ;

end;

function TSaisieGuide.GetGuideLigne(vRow: integer): Tob;
var lInVal : integer ;
begin
  result := nil ;
  if Guide = nil then Exit ;
  lInVal := Piece.GetInteger( vRow, 'EG_NUMLIGNE' ) ;
  if lInVal > 0 then
    result := Guide.FindFirst( ['EG_NUMLIGNE'], [ lInVal ], False ) ;
end;

function TSaisieGuide.GetArretPourCol(vRow : integer ; vStName: string): boolean;
var lTob     : Tob ;
    lIdx     : integer ;
    lStArret : string ;
begin
  result := False ;
  lTob := GetGuideLigne( vRow ) ;
  if Assigned( lTob ) then
    begin
    lStArret := lTob.GetString('EG_ARRET') ;
    lIdx     := GetIdxArret( 0, vStName ) ;
    if (length(lStArret) >= lIdx) and (lStArret<>'') and ( trim(lStArret[lIdx]) = 'X' )  then
      result := True ;
    end ;
end;

procedure TSaisieGuide.AfficheLignes;
var i : integer ;
begin
  Case Piece.ModeSaisie of
    msBOR :   TSP.AfficheGroupe( TSP.Row ) ;
    msLibre : for i := FInDeb to FInFin do
                TSP.AfficheLignes( i ) ;
    msPiece : TSP.AfficheLignes ;
    end ;
end;

function TSaisieGuide.IsValidLigne(vRow: integer): boolean;
begin
  ModeSilence( True ) ;
  result := Piece.IsValidLigne( vRow ) ;
  ModeSilence( False ) ;
end;

procedure TSaisieGuide.CompleteLigne(vRow: integer);
begin
  CalculLigne( vRow, GetGuideLigne( vRow ) ) ;
end;

function TSaisieGuide.AOuvrirAnal(vRow: integer): boolean;
begin
  result := False ;
//  if not Piece.EstRemplit(vRow) then Exit ;
  if Piece.GetInteger(vRow, 'E_NUMECHE') > 1 then Exit ;
  if Piece.GetString(vRow, 'E_ANA') <> 'X' then Exit ;
//  if not (Piece.GetMontantDev( vRow ) <> 0) then Exit ;
  result := CIsValidVentil( Piece.GetTob( vRow ), Piece.Info ).RC_Error <> RC_PASERREUR ;
end;

function TSaisieGuide.AOuvrirEche(vRow: integer): boolean;
var lTobGui  : Tob ;
    lStArret : string ;
begin
  result := False ;
  if not Piece.EstRemplit(vRow) then Exit ;
  if Piece.GetString(vRow, 'E_ECHE') <> 'X' then Exit ;
  if Piece.GetInteger(vRow, 'E_NUMECHE') > 1 then Exit ;
  if not (Piece.GetMontantDev( vRow ) <> 0) then Exit ;

  lTobGui := GetGuideLigne( vRow ) ;

  if Piece.GetString(vRow, 'E_MODEPAIE')='' then
    if lTobGui.GetString('EG_MODEREGLE') <> ''
      then Piece.CalculEche( vRow, lTobGui.GetString('EG_MODEREGLE') )
      else Piece.CalculEche( vRow ) ;

  lStArret  := lTobGui.GetString('EG_ARRET') ;
  result    := (length(lStArret) >= 7) and (lStArret<>'') and ( trim(lStArret[7]) = 'X' ) ;

end;

procedure TSaisieGuide.InitTabGuide;
begin
  FillChar( FTabGuideCol, Sizeof(FTabGuideCol), #0 ) ;

  FTabGuideCol[1] := TSP.GetColIndex('E_GENERAL') ;
  FTabGuideCol[2] := TSP.GetColIndex('E_AUXILIAIRE') ;
  FTabGuideCol[3] := TSP.GetColIndex('E_REFINTERNE') ;
  FTabGuideCol[4] := TSP.GetColIndex('E_LIBELLE') ;
  FTabGuideCol[5] := TSP.GetColIndex('E_DEBITDEV') ;
  FTabGuideCol[6] := TSP.GetColIndex('E_CREDITDEV') ;
  FTabGuideCol[7] := TSP.GetColIndex('E_MODEPAIE') ;

end;

procedure TSaisieGuide.GuideCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var lInOldRow : integer ;
begin
  if not GuideActif then Exit ;

  if Cancel then Exit ;

  lInOldRow := ARow ;

  GetColSuivante( ACol, ARow ) ;
  Cancel := not FBoFinGuide ;

  if not Cancel and ( ARow <> FListe.Row ) and (not Piece.IsValidLigne(TSP.Row)) then
    Cancel := true
  else if Cancel then
    begin
    TSP.Col := ACol ;
    TSP.Row := ARow ;
    end ;

  RecalculGuide ( lInOldRow ) ;

  if FBoFinGuide then
    FinirGuide ;

end;

procedure TSaisieGuide.GuideRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

  Cancel := ( FListe.Row < FInDeb ) or ( FListe.Row > FInFin ) ;
  if not Cancel then
    begin
    // Appel rowexit client
    if not IsValidLigne( Ou ) then
      CompleteLigne( Ou ) ;
    end ;

end;

procedure TSaisieGuide.GuideKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var Vide : Boolean ;
begin

  if Not FListe.SynEnabled then
    begin
    Key:=0 ;
    Exit ;
    end ;

  Vide:=(Shift=[]) ;

  case Key of
    VK_ESCAPE : if Vide then
                  begin
                  Key:=0 ;
                  AbandonnerGuide ;
                  end ;
    VK_TAB : if (Shift=[ssShift]) then
                  begin
                  Key:=0 ;
                  GetColPrecedente( TSP.Col, TSP.Row ) ;
                  end ;
// ===== Anulation de certaines Action en mode guide
    VK_INSERT : if (Vide) then Key:=0;
    VK_DELETE : if Shift=[ssCtrl] then Key:=0 ;
    VK_END,VK_HOME : Key:=0 ;
{^D,K,M,W,Y,Z}
    68,75,77,87,89,90 : if Shift=[ssCtrl] then Key:=0 ;

{Alt-G,L}
    71,76 : if Shift=[ssAlt]  then Key:=0 ;
  end ;

end;

function TSaisieGuide.GetColPrecedente( vCol, vRow : integer ): boolean;
var i,j         : integer ;
    lStArret    : string ;
    lTOBGuide   : TOB  ;
    lInCol      : integer ;
    lInRow      : integer ;
    ACol        : integer ;
    ARow        : integer ;
begin

  result := False ;

  if (vCol = FTabGuideCol[ 1 ]) and (vRow = FInDeb) then Exit ;

  ACol   := vCol ;
  ARow   := vRow ;
  lInCol := ACol ;
  lInRow := ARow ;

  try

    // il faut partir de l'index correspondant aux colonnes de la grille des guides
    Acol := GetIdxArret( Acol ) ;
    Dec( Acol ) ;

    // parcours de la ligne courante à la dernière
    for i := ARow downto FInDeb do
      begin
      lTOBGuide := GetGuideLigne( i ) ;
      if lTOBGuide = nil then exit ;

      // parcours des indicateurs d'arrêt...
      for j := ACol downto 1 do
        begin
        lStArret := copy( lTOBGuide.GetString('EG_ARRET') ,1, 6); // la grille de saisie ne comporte que 6 cases
        if ( length(lStArret) >= ACol ) and ( lStArret <> '' ) and ( trim(lStArret[j]) = 'X' ) then
          begin
          ARow      := i ;
          ACol      := FTabGuideCol[ j ] ;
          exit ;
          end ;
        end ;

      // --> reinit du numéro de colonne
      ACol := 6 ;

      end;

  finally

    if (ACol <> lInCol) or (ARow<>lInRow) then
      begin
      result := true ;
      TSP.SetPos( ACol, ARow, False, False, -1) ;
      end ;

    end ;

end;

function TSaisieGuide.GetLastGuide: string;
begin
  result := '' ;
  if Assigned( FTobGuide ) then
    result := FTobGuide.GetString('GU_GUIDE') ;
end;

end.


