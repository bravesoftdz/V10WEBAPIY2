{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 26/10/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPSAISIEPIECE ()
Mots clefs ... : TOF;CPSAISIEPIECE
*****************************************************************}
Unit CPSAISIEPIECE_TOF ;

Interface

Uses Controls,
     Windows,           // VK_INSERT, ...
     Classes,           // TList
     Menus,             // TPopupMenu
     forms,             // TForm
     ComCtrls,          // TPageControl

{$IFDEF SAISIEDOC}
     uLibSaisieDoc,    // gestionnaire des documents
{$ENDIF SAISIEDOC}

     HCtrls,            // THGrid
     HTB97,             // TToolBarButton97
     HEnt1,             //  TActionFiche
     uTob,
     uTof,
{$IFDEF EAGLCLIENT}
{$ELSE}
     DB, {FQ 17754 : TBookMark}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF EAGLCLIENT}
     uLibEcriture,             // TInfoEcriture
     ULibSaisiePiece,          // TSaisiePiece
     ULibPieceCompta,          // TPieceCompta
     SaisUtil,                 // RMVT
{$IFNDEF GCGC}
     SaisComp,                 // SaisieComplement  , R_COMP
{$ELSE !GCGC}
{$IFDEF CMPGIS35}
     SaisComp,                // SaisieComplement  , R_COMP
{$ENDIF CMPGIS35}
{$ENDIF}
     CPOBJENGAGE ;             //  TPieceEngagement fb 10/03/2006

var
  Engagement : boolean;   //fb 02/05/2006

function CPLanceFiche_SaisiePiece(vStRange, vStLequel, vStArgs : string) : string;


Function  TrouveSaisieParam(Q : TQuery ; Var M : RMVT ; Simul : String ; vDossier : String = '') : Boolean ;
Function  TrouveEtLanceSaisieParam(Q : TQuery ; TypeAction : TActionFiche ; Simul : String ; ModLess : Boolean = FALSE ; vDossier : String = '') : boolean ;

// fonction standard équivalent du lanceSaisie
Function  LanceSaisieParam ( QListe : TQuery ; Action : TActionFiche ; Var M : RMVT ; ModLess : Boolean = FALSE ; vDossier : String = '') : boolean ;
// Moification d'un Tpiececompta en mémoire
Function  ModifiePieceCompta ( vPieceCpt : TPieceCompta ; Action : TActionFiche ; vBoInMemory : boolean = True ) : boolean ;
// Ouverture de la saisie param
Function  SaisiePieceCompta  ( vStQualifPiece : string = 'N'; vAction : TActionFiche = taCreat ; vStJournal : string = '' ; vDtDateC : TDateTime = 0 ; vInNumeroPiece : integer = 0 ; vBoMiniCompta : boolean = False  ; vDossier : string = '' ) : boolean ;
{JP 20/07/07 : FQ 20601 : Création d'une fonction d'appel à partir d'une TOB}
function  SaisiePieceComptaTOB(vParam : TOB; vAction : TActionFiche = taCreat) : Boolean ;
{JP 20/07/07 : FQ 20601 : Fonction renvoyant la Tob Paramètre de SaisiePieceComptaTOB avec les champs voulus}
function  CreateTobParamSaisiePiece : TOB;

{$IFDEF SCANGED}
procedure SaisieParamMyAfterImport (Sender : TObject; FileGuid: string; var Cancel: Boolean) ;
{$ENDIF SCANGED}

Function  RmvtToPiece( M : RMVT ; vDossier : string = '' ) : TPieceCompta ;


Type
  TOF_CPSAISIEPIECE = Class (TOF)

    FListe           : THGrid ;                  {FP 02/05/2006 Rend public certains objets}

    FGCompl          : THGrid ;                  {FP 02/05/2006 Rend public certains objets}

    // Boutons
    BParamListe      : TToolBarButton97 ;
    BSolde           : TToolBarButton97 ;
    BValider         : TToolBarButton97 ;
    BEche            : TToolBarButton97 ;
    BVentil          : TToolBarButton97 ;
    BGenereTVA       : TToolBarButton97 ;
    BControlTVA      : TToolBarButton97 ;
    BModifTVA        : TToolBarButton97 ;
    BComplement      : TToolBarButton97 ;
    BRechercher      : TToolBarButton97 ;
    BProrata         : TToolBarButton97 ;
    BMenuTva         : TToolBarButton97 ;
    BScan            : TToolBarButton97 ;
    BDevise          : TToolBarButton97 ;
    BZoomImmo        : TToolBarButton97 ;

    // zones d'entête
    E_JOURNAL        : THValComboBox ;
    E_NATUREPIECE    : THValComboBox ;
    E_DATECOMPTABLE  : THEdit ;
    E_DEVISE         : THValComboBox ;
    E_ETABLISSEMENT  : THValComboBox ;
    E_NUMEROPIECE    : THNumEdit ;
    PERIODE          : THValComboBox ;
    FOLIO            : THValComboBox ;

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure OnAfterFormShow          ;
  private

    // ================================
    // ========== COMPOSANTS ==========
    // ================================

    // TSaisiePiece
    TSP                 : TSaisiePiece ;

    // Eléments interface
    FQMul               : TQuery ;
    PCEntete            : TPageControl ;
    PCPied              : TPageControl ;

    // Menus POPUP
    POPS                : TPopUpMenu;
    POPUPZOOM           : TPopUpMenu;
    POPUPGUIDE          : TPopUpMenu;
    POPUPLIGNES         : TPopUpMenu;
    POPUPMODIFS         : TPopUpMenu;
    POPUPTVA            : TPopUpMenu;

    // ===============================
    // ========== VARIABLES ==========
    // ===============================

    // Données de l'écriture
    FPieceCpt           : TPieceCompta ;          // Objet de gestion de la pièce
    PieceEngage         : TPieceEngagement;       //fb 02/05/2006

    // Stockage des pièces saisies
    FHistorique         : TList ;                 // Liste des pièces précédemment saisies
    FTobReglt           : Tob ;

    // Message
    FMessCompta         : TMessageCompta ;        // affichage des messages

    // indicateurs
    FAction             : TActionFiche ;
    FBoDetruirePiece    : Boolean ;               // Indicateur de libération de la mémoire du TPieceCompta
    FBoMajDB            : Boolean ;               // Indicateur de maj en base
    FModalResult        : TModalResult ;
    FBoJalSoldeP        : boolean ;
    FBoCptSoldeP        : boolean ;

    // accélérateur
    FStGenHT            : string ;

    FBoMouseDown        : boolean ;
    FBoChargePiece      : boolean ;
    FStatutSaisie       : TStatutSaisie ;


    // ======= Objets de gestion des données
    function  Piece     : TPieceCompta ;
    function  InfoEcr   : TInfoEcriture ;
    function  PContexte : TPieceContexte ;

    // ======= Initialisations
    procedure AnalyseParametre( vStArg : String ) ;
    procedure InitVariables ;
    procedure InitComposants ;
    procedure InitEvenements ;
    procedure InitPresentation ;
    function  GetTitre : string ;
    procedure SetTitre ;
    procedure CompletePiece ;

    // indicateurs
    function ANouveau : Boolean ;
    function Simul    : string ;
    function Dossier  : string ;

    // ======= Gestion du pied de page
    procedure RefreshPCPied             ;
    Procedure MajSoldePiece             ;
    Procedure InitInfosLigne            ;
    Procedure MajInfosLigne             ( vRow : Integer ) ;
    Procedure MajMessageLigne           ( vRow : Integer ) ;
    Procedure AfficherSolde             ( vStControl : String ; vSolde : Double ; vBoPer : boolean = False ) ;

    // =======  Gestion des soldes par période
    procedure ReinitSoldes              ;
    procedure CalculSoldeJal            ;
    procedure CalculSoldeGene           ( vStCpt : string ) ;
    procedure CalculSoldeAuxi           ( vStCpt : string ) ;
    procedure CalculSoldeGenePer        ( vStCpt : string ) ;
    procedure CalculSoldeAuxiPer        ( vStCpt : string ) ;
    procedure AfficherSoldeJal          ;
    procedure AfficherSoldeGene         ( vStCpt : string ) ;
    procedure AfficherSoldeAuxi         ( vStCpt : string ) ;

    // ====== Gestion accès des zones écrans
    procedure AccesEntete               ;
    procedure AccesDevise               ;
    procedure AccesListe                (StatutEng : string); //fb 02/05/2006

    // ======= Paramétrage des zones d'entête tous modes
    procedure AccesBoutons              ;
    procedure ZappePieceFolio           ( vStJal : String = '');
    procedure InitEtablissement         ;
    procedure majTypeExo                ;

    // ======= Gestion du journal
    procedure ParamJournal              ;
    procedure SetJournal                ( vStJal : string ) ;
    procedure JOURNALExit               ( Sender : TObject ) ;
    procedure JOURNALChange             ( Sender : TObject ) ;

    // BOR ======= Gestion de la période
    procedure ParamPeriode              ;
    procedure InitPeriode               ( vBoFocus : Boolean = False ) ;
    procedure PeriodeExit               ( Sender : TObject ) ;

    // BOR ======= Gestion du Folio
    procedure ParamFolio                ;
    procedure InitFolio                 ( vBoFocus : Boolean = False ; vBoForce : boolean = False ) ;
    procedure FolioExit                 ( Sender : TObject ) ;

    // PIECE ======= Gestion de la date comptable
    procedure InitDateComptable         ;
    procedure DATECOMPTABLEExit         ( Sender : TObject ) ;

    // PIECE ======= Gestion de la devise
    procedure InitDevise                ;
    procedure DEVISEExit                ( Sender : TObject ) ;

    // PIECE ======= Gestion de la nature de pièce
    procedure ParamNaturePiece          ( vStJal : string = '' ) ;
    procedure InitNaturePiece           ;
    procedure NaturePieceExit           ( Sender : TObject ) ;
    procedure NextNature                ( vRow : integer ) ;

    // PIECE ======= Gestion du numéro de la pièce
    procedure InitNumeroPiece           ;

    // ======= Chargement de la pièce
    procedure ChargeEntete              ;
    procedure ChargePieceFromMul        ( AvecRefresh : Boolean = True ) ;
    procedure ChargePiece               ( vBoReinit : Boolean = False) ;
    procedure RefreshModeTva            ;
    procedure SetInitFocus              ( vBoReinit : Boolean = False ) ;

    // ======= EVT entête
    procedure ETABLISSEMENTExit         ( Sender : TObject ) ;

    // ======= EVT GRILLE
    procedure TSP_RowEnter              ( Sender: TObject; Ou: Integer;  var Cancel: Boolean; Chg: Boolean );
    procedure TSP_RowExit               ( Sender: TObject; Ou: Integer;  var Cancel: Boolean; Chg: Boolean );
    procedure TSP_CellEnter             ( Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure TSP_CellExit              ( Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure TSP_DblClick              ( Sender : TObject ) ;
    procedure TSP_GridEnter             ( Sender : TObject ) ;
    procedure TSP_KeyDown               ( Sender: TObject; var Key: Word; Shift: TShiftState) ;

    // EVT FORM
    procedure FormKeyDown               ( Sender: TObject; var Key: Word; Shift: TShiftState) ;
    procedure FormCloseQuery            ( Sender: TObject; var CanClose: Boolean);
    procedure EcranClose                ;
    procedure FormMouseDown             (Sender: TObject;Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure PCEnteteEnter             (Sender: TObject);
    // ===========================
    // ===== GESTION TOOLBAR =====
    // ===========================

    // ======= Boutons de navigation
    procedure BFirstClick               ( Sender : TObject ) ;
    procedure BPrevClick                ( Sender : TObject ) ;
    procedure BNextClick                ( Sender : TObject ) ;
    procedure BLastClick                ( Sender : TObject ) ;
    procedure majBoutonNav ;
    {JP 31/05/07 : FQ 17754 : déplacement en multi sociétés}
    procedure ChangePiece(Sender : TObject);

    // ======= Actions annexes
    procedure BParamListeClick          ( Sender : TObject ) ;
    procedure BSoldeClick               ( Sender : TObject ) ;
    procedure BEcheClick                ( Sender : TObject ) ;
    procedure BVentilClick              ( Sender : TObject ) ;
    procedure BGenereTVAClick           ( Sender : TObject ) ;
    procedure BControlTVAClick          ( Sender : TObject ) ;
    procedure BModifTVAClick            ( Sender : TObject ) ;
    procedure BComplementClick          ( Sender : TObject ) ;
    procedure BRechercherClick          ( Sender : TObject ) ;
    procedure BProrataClick             ( Sender : TObject ) ;
    procedure BVisuRegltClick           ( Sender : TObject ) ;
    procedure BScanClick                ( Sender : TObject ) ;
    procedure BDeviseClick              ( Sender : TObject ) ;
    procedure BValiderClick             ( Sender : TObject ) ;

{$IFDEF SAISIEDOC}
    // Agrandissement de la fenêtre
    procedure BAgrandirClick            ( Sender : TObject ) ;
    procedure BReduireClick             ( Sender : TObject ) ;
{$ENDIF SAISIEDOC}

    // ======= Action annexes
    procedure ClickValide               ;
    procedure ClickAbandon              ( vBoFromGrille : Boolean = False ) ;
    procedure ZoomCompte                ( vCol : Integer ) ;
    procedure ZoomGeneral               ;
    procedure ZoomTiers                 ;
    procedure LettrageEnSaisie          ;
    procedure EditionSaisie             ;

    // ==========================
    // ===== GESTION POPUPS =====
    // ==========================
    procedure OnPopUpPOPS        ( Sender : TObject ) ;
    procedure EnCoursDev         ( Sender : TObject ) ;
    // ---> Popup Guides
    procedure OnPopUpZooms       ( Sender : TObject ) ;
    procedure InitPopupZooms     ( vActivation : Boolean = True ) ;
    procedure ZoomJournalClick   ( Sender : TObject ) ;
    procedure ZoomDeviseClick    ( Sender : TObject ) ;
    procedure ZoomEtabClick      ( Sender : TObject ) ;
    procedure ZoomScenarioClick  ( Sender : TObject ) ;
    procedure ZoomPiecesClick    ( Sender : TObject ) ;
    procedure ZoomTauxClick      ( Sender : TObject ) ;
    procedure ZoomTiersClick     ( Sender : TObject ) ;
    procedure ZoomGeneralClick   ( Sender : TObject ) ;
    procedure ZoomImmoClick      ( Sender : TObject ) ;
    // ---> Popup Guides
    procedure OnPopUpGuide       ( Sender : TObject ) ;
    procedure InitPopupGuide     ( vActivation : Boolean = True  ) ;
    procedure GuideOuvrirClick   ( Sender : TObject ) ;
    procedure GuideCreerClick    ( Sender : TObject ) ;
    // ---> Popup Actions Lignes
    procedure OnPopUpLignes      ( Sender : TObject ) ;
    procedure InitPopupLignes    ( vActivation : Boolean = True  ) ;
    procedure LignesInsertClick  ( Sender : TObject ) ;
    procedure LignesSupprClick   ( Sender : TObject ) ;
    procedure LignesLibelleClick ( Sender : TObject ) ;
    procedure LignesRAZClick     ( Sender : TObject ) ;
    procedure AssocierPdf        ( Sender : TObject ) ;
    // ---> Popup Modifs
    procedure OnPopUpModifs      ( Sender : TObject ) ;
    procedure InitPopupModifs    ( vActivation : Boolean = True  ) ;
    procedure ModifsSerieClick   ( Sender : TObject ) ;
    procedure ModifsRegimeClick  ( Sender : TObject ) ;
    procedure ModifsRIBClick     ( Sender : TObject ) ;
    // ---> Popup Tva
    procedure InitPopupTva ;
    procedure TvaDebitClick      ( Sender : TObject ) ;
    procedure TvaEncClick        ( Sender : TObject ) ;
    procedure TvaMixteClick      ( Sender : TObject ) ;
    // ---> Popup Acc
    procedure AccActifClick      ( Sender : TObject ) ;
    procedure AccTiersClick      ( Sender : TObject ) ;

    // compléments
    function  GereComplements    ( vRow   : integer ; lBoEntete : Boolean = False ) : Boolean ;
    function  OuvreComplements   ( vRow   : integer ; vRComp : R_COMP ; vAction : TActionfiche ) : Boolean ;

    // Cutoff
    procedure GereCutoff         ( vRow : integer) ;

    // Analytique
    procedure OuvreAnal          ( vRow : Integer ; vInAxe : Integer = 0 ) ;
    procedure GereAnal           ( vRow : Integer ) ;
    function  TestQteAna         ( vRow : integer ) : boolean ;

    // Echéances
    procedure GereEche           ( vRow : Integer ) ;
    procedure OuvreEche          ( vRow : Integer ; vBoDeplace : boolean = false ) ;

    // Les immos
    procedure GereLesImmos       ( vRow : Integer ) ;

    // Tiers payeurs
    function GereTiersPayeurs : boolean ;

    // Gestion de la devise
    procedure AvertirPbTaux      ( vRow : Integer = 0 ) ;

    // Gestion dernières pièces
    procedure StockeLaPiece ;

    // Gestion des erreurs
    procedure TraiteErreurTva ( vErrId : Integer ) ;

    // Controle des doulons
    function  TraiteDoublon : boolean ;
    function  TrouveDoublonLigne( vRow : integer ; vStChamp : string ; var vTobRech : Tob ) : Boolean ;
    function  GetAction : TActionFiche;
    procedure SetAction(Value : TActionFiche);

   {$IFDEF SCANGED}
    procedure SetGuidID ( vGuidId : string ) ;
    {$ENDIF}


    property Action : TActionFiche read GetAction write SetAction default taConsult;

    public
    property SaisiePiece: TSaisiePiece  read TSP;      {FP 02/05/2006}

   {$IFDEF SCANGED}
    function GetInfoLigne : string ;
    property GuidID : string  write SetGuidID ;
    {$ENDIF}

{$IFDEF SAISIEDOC}
    // gestion de la GED
    procedure InitViewer        ;
{$ENDIF SAISIEDOC}

  end ;

  {b FP 02/05/2006}
  TBeforeValidePieceCompta = function(Tof: TOF_CPSAISIEPIECE; Ecran: TForm; Piece: TPieceCompta): Boolean of object;
  PBeforeValidePieceCompta = ^TBeforeValidePieceCompta;

  TBeforeSuppLigne = function(Tof: TOF_CPSAISIEPIECE; Piece: TPieceCompta; NumLigne: Integer): Boolean of object;
  PBeforeSuppLigne = ^TBeforeSuppLigne;

  TBeforeChargePiece = procedure(Tof: TOF_CPSAISIEPIECE; Ecran: TForm; Piece: TPieceCompta; var Action: TActionFiche) of object;
  PBeforeChargePiece = ^TBeforeChargePiece;

  TAfterChargePiece = procedure(Tof: TOF_CPSAISIEPIECE; Ecran: TForm; Piece: TPieceCompta) of object;
  PAfterChargePiece = ^TAfterChargePiece;
  {e FP 02/05/2006}
Implementation

uses
     sysutils, // DateToStr
     Grids,    // GoEditing

     // --------------> AGL
     Vierge,
     Choix,                     // choisir
     ParamSoc,
     AglInit,                  // ActionToString, theTob, theData
     ed_tools,                 // videliste
     messages,                  // WM_KEYDOWN
     uTobDebug,
     hDebug,
     HMsgBox,                  // PGIAskCancel
     {$IFDEF EAGLCLIENT}
       MaineAGL,          // AGLLanceFiche
       UtileAGL,          // LanceEtat
     {$ELSE}
       EdtREtat,          // LanceEtat
       FE_Main,          // AGLLanceFiche
     {$ENDIF EAGLCLIENT}
     {$IFDEF VER150}
       Variants,
     {$ENDIF}
     {$IFDEF SCANGED}
       UtilGed,
       UGedFiles ,
       UGedViewer,
       AnnOutils,
       Dialogs ,
       cbpPath ,
     {$ENDIF}

     // --------------> Produit
     Ent1,                     // VH^
     UtilPGI,                  // _Blocage
     uLibAnalytique,           // AlloueAxe
     SaisComm,                 // MvtToIdent
     CPOuvertureGEP,           // pour OnAfterValidePieceCompta et OnBeforeValidePieceCompta
     ULibBonAPayer,            //
     ULibExercice,             //
     CPGENEREBAP_TOF,          //
     CPMASQUESAISIE_TOF,

     // --------------> ZOOMS
     Devise_TOM,	             // Devise,
     FichComm,                 // Etablissement
     Saistaux1,                // Taux
     CPGeneraux_TOM,           // Généraux
     CPTiers_TOM,              // Tiers
     CPCHANCELL_TOF,           // Chancellerie
     SaisVisu,                 // Visualisation des dernieres pieces saisies
     CPJOURNAL_TOM,            // Journal
     SUIVCPTA_TOM,             // Scénario cwas
     CPVisuPieces_Tof,
     CPMULSAISLETT_TOF,        // pour LettrerEnSaisie
     // --------------> Actions complémentaires
     SaisTVA,                  // Contrôle de TVA
     {$IFNDEF GCGC}
       SaisBase,                 // Modification des bases TVA
       SaisEnc,                  // Visu des info de TVA
     {$ENDIF}
     eSaisAnal,                // Saisie analytiqque
     Zdevise,                  // Saisie en devise des folio
  {$IFDEF COMPTAAVECSERVANT}
    EntImo, FicheImo,
  {$ENDIF}
  {$IFDEF AMORTISSEMENT}
    IMMO_TOM,	// pour FicheImmobilisation
  {$ENDIF}
{$IFDEF SAISIEDOC}
  hpanel,            // THpanel
{$ENDIF SAISIEDOC}
     CPEcheMono_TOF ,          // Mono-échéance
     CPEcheMulti_TOF ;         // Multi-échéances

//=======================================================
procedure _LiberePiece( vPieceCpt : TPieceCompta ) ;
Var lInfo : TInfoEcriture ;
begin
  if Assigned( vPieceCpt ) then
    begin
    lInfo := vPieceCpt.Info ;
    FreeAndNil( vPieceCpt ) ;
    end ;
  if Assigned( lInfo ) then
    FreeAndNil( lInfo ) ;
end ;
//=======================================================
Function SaisieParamLancee : Boolean ;
Var i : Integer ;
begin
  result := False ;
  for i:=0 to Screen.FormCount-1 do
    if UpperCase(Screen.Forms[i].Name)='CPSAISIEPIECE' then
      begin
      result:=True ;
      break ;
      end ;
end ;
//=======================================================
function CanOpenSaisie( var Action : TActionFiche ; Contexte : TPieceContexte; vDtDateSaisie : TDateTime = 0 ) : Boolean ;
var Q         : TQuery ;
    sExo      : string ;
    lDtDateC  : TDateTime ;
    lDtCloPer : TDateTime ;
    lInNbJAv  : integer ;
    lInNbJAp  : integer ;

    function _TestJalExist : boolean ;
      begin
        Q := OpenSelect('SELECT  ##TOP 1## J_JOURNAL FROM JOURNAL WHERE J_FERME="-" AND J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"ANA" AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ODA"', contexte.Dossier) ;
        result := not Q.Eof ;
        Ferme(Q) ;
        if not result then
          PgiBox( TraduireMemoire('Veuillez créer au moins un journal de type Pièce, Bordereau ou Libre'), TitreHalley ) ;
      end ;

begin

  // Masque de saisie ok ?
  if not CPresenceMasque then
    begin
    Result := False ;
    Exit ;
    end ;

  // TEST fenêtre de saisie déjà ouverte
  if SaisieParamLancee then
    Action := taConsult ;

  if Action <> taConsult then
    begin

    // TEST date d'entrée (ou de saisie)
    if vDtDateSaisie <= iDate1900
      then lDtDateC := V_PGI.DateEntree
      else lDtDateC := vDtDateSaisie ;

    if (not (ctxPCL in V_PGI.PGIContexte)) then
      begin
      // Test Date dans Exo ouvert
      sExo := Contexte.GetExoDT( lDtDateC ) ;
      if (sExo = Contexte.Exercices.EnCours.Code) and
         ( (lDtDateC > Contexte.Exercices.Encours.Fin) or (lDtDateC < Contexte.Exercices.Encours.Deb) ) then
        sExo:='' ;
      if not ((sExo=Contexte.Exercices.EnCours.Code) or (sExo=Contexte.Exercices.Suivant.Code)) then
        if ( Action = taCreat ) then
          begin
          PgiBox( TraduireMemoire('La date d''entrée est incompatible avec la saisie.'), TitreHalley ) ;
          Result := False ;
          Exit ;
          end
        else Action := taConsult ;
      end ;

      // test date sur période non close
      lDtCloPer := Contexte.DateCloturePer ;
      if ( (lDtCloPer>0) and (lDtDateC<=lDtCloPer) ) then
        if ( Action = taCreat ) then
          begin
          PgiBox( TraduireMemoire('La date que vous avez renseignée est antérieure à une clôture.'), TitreHalley ) ;
          Result := False ;
          Exit ;
          end
        else Action := taConsult ;

      // test plage de date en mode création
      if (not (ctxPCL in V_PGI.PGIContexte)) and ( Action = taCreat ) then
        begin
        lInNbJAv  := Contexte.NbJEcrAvant ;
        lInNbJAp  := Contexte.NbJEcrApres ;
        if ( (lInNbJAv > 0) and (lDtDateC<V_PGI.DateEntree) and (V_PGI.DateEntree-lDtDateC > lInNbJAv) )
           or
           ( (lInNbJAp > 0) and (lDtDateC>V_PGI.DateEntree) and (lDtDateC-V_PGI.DateEntree > lInNbJAp) ) then
              begin
              PgiBox( TraduireMemoire('La date que vous avez renseignée est en dehors des limites autorisées.'), TitreHalley ) ;
              Result := False ;
              Exit ;
              end ;
        end ;

    // TEST établissement par défaut
    if Contexte.EtablisDefaut = '' then
      begin
      PgiBox( TraduireMemoire('L''établissement par défaut n''est pas renseigné.'), TitreHalley ) ;
      Result:=FALSE ;
      Exit ;
      end ;

    // TEST droit de création d'écritures
    if not ExJaiLeDroitConcept( TConcept(ccSaisEcritures), True) then
      begin
      Result := False ;
      Exit ;
      end ;

    // TEST Journaux autorisés
    if (Contexte.HalleyUser.JalAutorises<>'') then
      begin
      Result := False ;
      Q := OpenSelect('SELECT J_JOURNAL FROM JOURNAL WHERE J_FERME="-" AND J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"ANA" AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ODA"', contexte.Dossier) ;
      while not Q.EOF do
        begin
        if (Pos( ';' + Q.FindField('J_JOURNAL').AsString + ';', Contexte.HalleyUser.JalAutorises)>0) then
          begin
          result := True ;
          break ;
          end ;
        Q.Next ;
        end ;
      Ferme(Q) ;
      // Existe-t-il un journal de type Folio autorisé pour l'utilisateur courant ?
      if not Result then
        begin
        PgiBox( TraduireMemoire('Aucun journal autorisé pour cette saisie.'), TitreHalley ) ;
        Exit ;
        end ;
      end
    else
      begin
      if not _TestJalExist then
        begin
        result := FALSE ;
        exit;
        end;
      end ;

    // Existe-t-il une période de saisie au delà de la date de révision ?
    if Contexte.ModeRevisionActive( lDtDateC ) or
       Contexte.ModeRevisionActive( FinDeMois( Contexte.Exercices.Entree.Fin ) ) then
      begin
      PgiBox( TraduireMemoire('Le mode révision est actif. Seul un contrôleur peut créer ou modifier des écritures avant le '+ DateToStr(Contexte.DateRevision)+ ' .'), TitreHalley ) ;
      result := FALSE ;
      exit;
      end;

    end
  else
    begin
    if not _TestJalExist then
      begin
      result := FALSE ;
      exit;
      end;
    end ;

  result:=TRUE ;

end ;

// =======================================================================

function CPLanceFiche_SaisiePiece(vStRange, vStLequel, vStArgs : string) : string;
begin
  result := AGLLanceFiche('CP', 'CPSAISIEPIECE', vStRange, vStLequel, vStArgs);
end ;

// =======================================================================

Function TrouveEtLanceSaisieParam(Q : TQuery ; TypeAction : TActionFiche ; Simul : String ; ModLess : Boolean = FALSE ; vDossier : String = '') : Boolean ;
Var M : RMVT ;
begin
{b fb 02/05/2006}
  Engagement := (Simul ='p');
{e fb 02/05/2006}
  Result := TrouveSaisieParam( Q, M, Simul, vDossier ) ;
  if Engagement then
    M.Simul:='p';
  if Result then
    LanceSaisieParam(Q,TypeAction,M,ModLess,vDossier) ;
end ;
// =======================================================================

Function TrouveSaisieParam(Q : TQuery ; Var M : RMVT ; Simul : String ; vDossier : String ) : Boolean ;
Var Q1     : TQuery ;
    lStSQL :  String ;
begin

  Result := FALSE ;
  if (Q.EOF) And (Q.Bof) then Exit ;

  if vDossier = '' then
    vDossier := V_PGI.SchemaName ;

  // Requête
  lStSQL := 'SELECT ##TOP 1## E_JOURNAL, E_EXERCICE, E_QUALIFPIECE, E_DATECOMPTABLE, E_NUMEROPIECE, '
                           + 'E_VALIDE, E_ETABLISSEMENT, E_MODESAISIE, E_DEVISE, E_NATUREPIECE, '
                           + 'E_TAUXDEV, E_DATETAUXDEV, E_NUMLIGNE, E_NUMECHE, E_GENERAL, E_ETATREVISION '
             + ' FROM ECRITURE '
             + ' WHERE E_JOURNAL="'     + Q.FindField('E_JOURNAL').AsString + '"'
               + ' AND E_EXERCICE="'    + QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) + '"'
               + ' AND E_NUMEROPIECE='  + Q.FindField('E_NUMEROPIECE').AsString ;

  M.ModeSaisieJal := Q.FindField('E_MODESAISIE').AsString ;
  if ((M.ModeSaisieJal<>'-') and (M.ModeSaisieJal<>'')) then
    lStSQL := lStSql   + ' AND E_DATECOMPTABLE>="' + USDATETIME(DebutDeMois(Q.FindField('E_DATECOMPTABLE').AsDateTime)) + '"'
                       + ' AND E_DATECOMPTABLE<="' + USDATETIME(FinDeMois(Q.FindField('E_DATECOMPTABLE').AsDateTime)) + '"'
  else
    lStSQL := lStSql   + ' AND E_DATECOMPTABLE="' + USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime) + '"' ;

  If Simul<>''
    then lStSQL := lStSQL + ' AND E_QUALIFPIECE="' + Simul + '" '
    else lStSQL := lStSQL + ' AND E_QUALIFPIECE="N" ' ;

  // Recherche
  Q1 := OpenSelect( lStSQL, vDossier ) ;
  if Not Q1.EOF then
    begin
    result := True ;
    M := MvtToIdent(Q1,fbGene,False) ;
    end ;
  Ferme(Q1) ;

END ;

// =======================================================================
Function  SaisiePieceCompta  ( vStQualifPiece : string ; vAction : TActionFiche; vStJournal : string ; vDtDateC : TDateTime ; vInNumeroPiece : integer ; vBoMiniCompta : boolean ; vDossier : string ) : boolean ;
var
  T : TOB;
begin
  T := CreateTobParamSaisiePiece;
  T.SetString('E_QUALIFPIECE', vStQualifPiece);
  T.SetString('E_JOURNAL', vStJournal);
  T.SetString('DOSSIER', vDossier);
  if vBoMiniCompta then T.SetString('MINICOMPTA', 'X');
  T.SetDateTime('E_DATECOMPTABLE', vDtDateC);
  T.SetInteger('E_NUMEROPIECE', vInNumeroPiece);
  {JP 20/07/07 : FQ 20601 : Code de lancement de la saisie paramétrable déporté dans SaisiePieceComptaTOB}
  Result := SaisiePieceComptaTOB(T, vAction);
  FreeAndNil( T ) ;
end;

{JP 20/07/07 : FQ 20601 : Création d'une fonction d'appel à partir d'une TOB}
{---------------------------------------------------------------------------------------}
function SaisiePieceComptaTOB(vParam : TOB; vAction : TActionFiche = taCreat) : Boolean ;
{---------------------------------------------------------------------------------------}
var
  lStArgs      : String ;
  lDtDatePiece : TDateTime ;
  lPieceCpt    : TPieceCompta ;
  lInfoEcr     : TInfoEcriture ;
  lJournal     : string;
begin
  Result := False;

  // init date de saisie
  if vParam.GetDateTime('E_DATECOMPTABLE') > iDate1900
    then lDtDatePiece := vParam.GetDateTime('E_DATECOMPTABLE')
    else lDtDatePiece := V_PGI.DateEntree ;

  // passage de la pièce
  lInfoEcr  := TInfoEcriture.Create(vParam.GetString('DOSSIER'));
  lPieceCpt := TPieceCompta.CreerPiece( lInfoEcr ) ;

  lJournal := vParam.GetString('E_JOURNAL');
  // init PCL
  if (ctxPCL in V_PGI.PGIContexte) and ( lJournal = '' ) then
    begin
    lPieceCpt.contexte.GetZLastJal( lJournal ) ;
    lPieceCpt.contexte.GetZLastDate( lDtDatePiece ) ;
    if ( lJournal <> '') then
      if lPieceCpt.contexte.GetExoDT( lDtDatePiece ) <> lPieceCPt.Contexte.Exercices.CPExoRef.code then
        lDtDatePiece := lPieceCPt.Contexte.Exercices.CPExoRef.Deb ;
    end ;

  // Blocage
  if not CanOpenSaisie( vAction, lPieceCpt.contexte, lDtDatePiece) then
    begin
    _LiberePiece( lPieceCpt ) ;
    Exit ;
    end ;

  if vAction <> taConsult then
    if _Blocage(['nrCloture'],True,'nrSaisieCreat') then
      begin
      _LiberePiece( lPieceCpt ) ;
      Exit ;
      end ;

  // Paramétrage de la saisie
  if vParam.GetString('MINICOMPTA') = 'X' then
    begin
    lPieceCpt.SetMultiEcheOff ;
    lPieceCpt.SetScenarioOff ;
    end
  else
    begin
    lPieceCpt.ActiveAcc ;
    lPieceCpt.ActiveTP ;
    lPieceCpt.SetMultiEcheMulti ;
    lPieceCpt.SetScenarioOn ;
    end ;

  // Initialisation de l'entête
  lPieceCpt.InitPiece( lJournal, lDtDatePiece, '', '', '', vParam.GetString('E_QUALIFPIECE') );
  if vParam.GetInteger('E_NUMEROPIECE') <> 0 then
    lPieceCpt.PutEntete('E_NUMEROPIECE', vParam.GetInteger('E_NUMEROPIECE')) ;

  {JP 20/07/07 : Ajout de nouveau champs ....}
  if vParam.GetString('SAISIEPIECECOMPTA') = '-' then begin
    if vParam.GetInteger('E_NUMLIGNE') > 0 then
      lPieceCpt.NumLigneAppel := vParam.GetInteger('E_NUMLIGNE'); 
    {... éventuelle gestion d'autres champs ...}
  end;

  TheTob    := TOB( lPieceCpt ) ;
  TheData   := nil ;

  // Préparation des paramètres
  lStArgs := ActionToString( vAction ) + ';LIBEREMEMOIRE=X' ;

  // Ouverture de la fiche
  result := CPLanceFiche_SaisiePiece( '', '', lStArgs ) <> 'CANCEL' ;

  if vAction <> taConsult then
    _Bloqueur('nrSaisieCreat',False) ;

//  Result := True;
end;

{JP 20/07/07 : FQ 20601 : Fonction renvoyant la Tob Paramètre de SaisiePieceComptaTOB avec les champs voulus}
{---------------------------------------------------------------------------------------}
function CreateTobParamSaisiePiece : TOB;
{---------------------------------------------------------------------------------------}
begin
  Result := TOB.Create('_SaisieParametrable', nil, -1);
  {Champs de la fonction SaisiePieceCompta}
  Result.AddChampSupValeur('E_JOURNAL', '');
  Result.AddChampSupValeur('E_DATECOMPTABLE', iDate1900);
  Result.AddChampSupValeur('E_NUMEROPIECE', 0);
  Result.AddChampSupValeur('E_QUALIFPIECE', 'N');
  Result.AddChampSupValeur('MINICOMPTA', '-');
  Result.AddChampSupValeur('DOSSIER', '');
  {Champ à mettre à -, si l'appel n'est pas standard}
  Result.AddChampSupValeur('SAISIEPIECECOMPTA', 'X');
  {Champs nouveaux}
  Result.AddChampSupValeur('E_EXERCICE', '');
  Result.AddChampSupValeur('E_NUMLIGNE', 0);
  Result.AddChampSupValeur('E_NUMECHE', -1);
end;

// =======================================================================
Function  ModifiePieceCompta ( vPieceCpt : TPieceCompta ; Action : TActionFiche ; vBoInMemory : boolean = True ) : boolean ;
Var lStArgs   : String ;
begin
  Result  := False ;

  // Blocage
  if not CanOpenSaisie( Action, vPieceCPT.contexte, vPieceCpt.GetEnteteDt('E_DATECOMPTABLE') ) then Exit ;
  if Action = taModif then
    if _Blocage(['nrCloture','nrBatch'],True,'nrSaisieModif') then Exit ;

  // passage de la pièce
  TheTob  := TOB( vPieceCpt ) ;
  TheData := nil ;

  // Préparation des paramètres
  lStArgs := ActionToString( Action ) ;
  if vBoInMemory then
    lStArgs := lStArgs + ';MAJDB=OFF' ;
  result  := CPLanceFiche_SaisiePiece( '', '', lStArgs ) <> 'CANCEL';

  // Déblocage
  Case Action of
    taModif : _Bloqueur('nrSaisieModif',False) ;
    end ;

end ;

// =======================================================================

Function  LanceSaisieParam ( QListe : TQuery ; Action : TActionFiche ; Var M : RMVT ; ModLess : Boolean ; vDossier : String ) : boolean ;
Var lStRange  : String ;
    lStLequel : String ;
    lStArgs   : String ;
    DateSais  : TDateTime ;
    lPieceCpt : TPieceCompta ;
begin
  Result         := False ;

{b fb 02/05/2006}
  Engagement := (M.Simul ='p');
  if Engagement then
    M.Simul := 'P';
{e fb 02/05/2006}

  M.LastNumCreat := -1 ;
  DateSais       := V_PGI.DateEntree ;
  lPieceCpt      := RmvtToPiece( M, vDossier ) ;

  // Blocage
  if not CanOpenSaisie( Action, lPieceCpt.contexte, lPieceCpt.GetEnteteDt('E_DATECOMPTABLE') ) then
    begin
    _LiberePiece( lPieceCpt ) ;
    Exit ;
    end ;

  Case Action of
     taCreat : begin
               if PasCreerDate(DateSais) then
                  begin
                  _LiberePiece( lPieceCpt ) ;
                  Exit ;
                  end ;
               if _Blocage(['nrCloture'],True,'nrSaisieCreat') then
                  begin
                  _LiberePiece( lPieceCpt ) ;
                  Exit ;
                  end ;
               end ;
     taModif : begin
               if RevisionActive(M.DateC) then
                 begin
                 _LiberePiece( lPieceCpt ) ;
                 Exit ;
                 end ;
               if _Blocage(['nrCloture','nrBatch'],True,'nrSaisieModif') then
                 begin
                 _LiberePiece( lPieceCpt ) ;
                 Exit ;
                 end ;
               end ;
    end ;


{$IFNDEF EAGLCLIENT}
    if Assigned( QListe ) then
      TheData := QListe ;
{$ENDIF EAGLCLIENT}

  // paramétrage du TPieceCompta pour la saisie classique
  lPieceCpt.ActiveAcc ;
  lPieceCpt.ActiveTP ;
  lPieceCpt.SetMultiEcheMulti ;
  lPieceCpt.SetScenarioOn ;

  // Préparation des paramètres
  TheTob    := TOB( lPieceCpt ) ;
  lStArgs := ActionToString( Action ) + ';' + 'DATESAIS=' + DateToStr( DateSais )
                                      + ';' + 'LIBEREMEMOIRE=X' ;

  // Ouverture de la fiche
  CPLanceFiche_SaisiePiece( lStRange, lStLequel, lStArgs );

  // Déblocage
  Case Action of
    taCreat : _Bloqueur('nrSaisieCreat',False) ;
    taModif : _Bloqueur('nrSaisieModif',False) ;
    end ;

end ;

// =======================================================================

Function  RmvtToPiece( M : RMVT ; vDossier : string = '' ) : TPieceCompta ;
var lInfoEcr : TInfoEcriture ;
begin
  lInfoEcr := TInfoEcriture.Create( vDossier ) ;
  Result   := TPieceCompta.CreerPiece( lInfoEcr ) ;
  with M do
    begin
    // Données d'entête
    Result.InitPiece( Jal, DateC, Nature, CodeD, Etabl, Simul ) ;
    Result.PutEntete('E_NUMEROPIECE',   Num ) ;
    if ANouveau then
      Result.PutEntete('E_ECRANOUVEAU', 'H' ) ;
    end ;
end ;


{$IFDEF SCANGED}
procedure SaisieParamMyAfterImport (Sender : TObject; FileGuid: string; var Cancel: Boolean) ;
var LastError  : String;
    lTobDoc    : TOB;
    lTobDocGed : TOB;
    lStDocId   : string;
begin

 if (ctxPCL in V_PGI.PGIContexte) and not JaileDroitConceptBureau (187315) then exit ;

 lTobDoc := Tob.Create('YDOCUMENTS', nil, -1) ;
 lTobDoc.LoadDb ; // Pas de clé = Charge une clé à 0
 lTobDocGed := Tob.Create('DPDOCUMENT', nil, -1) ;
 lTobDocGed.LoadDb ;

 try

  // TobDoc :    principaux champs à renseigner
  //             DOCID, ANNEE, MOIS, DATEDEB, DATEFIN, LIBELLEDOC,
  //             NATDOCUMENT, AUTEUR, MOTSCLES, BLOCNOTE
  lTobDoc.PutValue('YDO_LIBELLEDOC'    , TOF_CPSAISIEPIECE(TFVierge(Sender).LaTOF).GetInfoLigne) ; // PAR EXEMPLE !!
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
     TOF_CPSAISIEPIECE(TFVierge(Sender).LaTOF).GuidID := lStDocId ;
//   TFSaisie(Sender).GuidID := lStDocId ; // pour traitement par l'intéressé


end ;

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

 SaisieParamMyAfterImport(Sender,lStFileGuid,lBoCancel) ;

end ;




{$ENDIF SCANGED}


// =======================================================================

procedure TOF_CPSAISIEPIECE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPSAISIEPIECE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPSAISIEPIECE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPSAISIEPIECE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPSAISIEPIECE.OnArgument (S : String ) ;
begin
  Inherited ;

  // Récupération TheData
{$IFNDEF EAGLCLIENT}
  If Assigned(TheData) then
    begin
    FQMul := TQuery( TheData ) ;
    TheData := nil ;
    end ;
{$ENDIF EAGLCLIENT}

  // Récupère les arguments
  AnalyseParametre( S ) ;
  // Initialisation Variables
  InitVariables ;
  // Initialisation fiche
  InitComposants ;
  InitEvenements ;
  // Initialisation présentation
  InitPresentation ;

{$IFDEF SCANGED}
{
  CA - 20/02/2007 - Si la GED n'est pas sérialisée, on autorise quand même
                    l'insertion de documents en saisie
}
  if not VH^.OkModGed then
  begin
    if V_PGI.RunFromLanceur then
      InitializeGedFiles(V_PGI.DefaultSectionDbName, SaisieParamMyAfterImport)
    else
      InitializeGedFiles(V_PGI.DbName, SaisieParamMyAfterImport);
  end;
{$ENDIF}


end ;

procedure TOF_CPSAISIEPIECE.OnClose ;
begin
  Inherited ;

{$IFDEF SCANGED}
  // Ajout CA - 20/02/2007 - Dans le cas où la GED n'est pas sérialisée
  if not VH^.OkModGed then FinalizeGedFiles;
{$ENDIF}

  // Edition de fin de saisie
  if ( GetAction = taCreat ) and ( FHistorique.Count > 0 ) then
    EditionSaisie ;

  if FModalResult = mrOk
    then TFVierge(Ecran).retour := 'OK'
    else TFVierge(Ecran).retour := 'CANCEL' ;

  // Deblocage du folio
  if (Piece.ModeSaisie <> msPiece) and FBoMajDB then // and (Piece.Action <> taConsult)
    Piece.UnLockFolio( True ) ;

 // Libération des objets
 try

  if assigned(TSP)               then TSP.Free;
  if assigned(FMessCompta)       then FMessCompta.Free;
  if assigned(FHistorique)       then begin
                                      VideListe(FHistorique) ;
                                      FHistorique.Free ;
                                      end ;
  if Assigned( FTobReglt )       then FTobReglt.Free ;

  // TPieceCompta libéré uniquement sur demande
  if FBoDetruirePiece and Assigned(FPieceCpt) then
    begin
    InfoEcr.Free ;
    FPieceCpt.Free ;
    end ;

  if Assigned(PieceEngage) then PieceEngage.Free ;

 finally
  TSP                       := nil;
  FMessCompta               := nil;
  FHistorique               := nil;
  FPieceCpt                 := nil;
  FTobReglt                 := nil;
  PieceEngage               := nil;
  ECran.caption := '' ;
  UpdateCaption( Ecran ) ;

 end;

end ;

procedure TOF_CPSAISIEPIECE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPSAISIEPIECE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPSAISIEPIECE.AnalyseParametre(vStArg: String);
var lStChaine : string;
begin

  // Initialisation des paramètres d'appels
  Action               := taConsult ;
  FBoMajDB             := True ;
  FBoDetruirePiece     := False ;

  while vStArg <> '' do
    begin
    lStChaine := ReadTokenSt(vStArg);

    // Stocke le XXX de "MODE=XXX" dans la variable
    if (Pos('ACTION=', Uppercase(lStChaine)) > 0) then
      begin
      Action            := StringToAction( lStChaine ) ;
      end

    // libération de la piece
    else if (Pos('LIBEREMEMOIRE=', Uppercase(lStChaine)) > 0) then
      begin
      FBoDetruirePiece := Copy( lStChaine, 15, length(lStChaine) ) = 'X' ;
      end

    else if (Pos('MAJDB=', Uppercase(lStChaine)) > 0) then
      begin
      FBoMajDB := not ( Copy( lStChaine, 7, length(lStChaine) ) = 'OFF' ) ;
      end ;

  end;

end;


procedure TOF_CPSAISIEPIECE.InitComposants;
begin

  // Eléments interface
  FListe          := THGrid( GetControl( 'FLISTE', True ) ) ;
  PCEntete        := TPageControl( GetControl('PCEntete', True) );
  PCPied          := TPageControl( GetControl('PCPied', True) );
  PCPied.Enabled  := False ;
  FGCompl    := THGrid( GetControl( 'GCOMPL', True ) ) ;

  // Menus popup
  POPS        := TPopUpMenu( GetControl('POPS', True) );
  POPUPZOOM   := TPopUpMenu( GetControl('POPUPZOOM', True) );
  POPUPGUIDE  := TPopUpMenu( GetControl('POPUPGUIDE', True) );
  POPUPLIGNES := TPopUpMenu( GetControl('POPUPLIGNES', True) );
  POPUPMODIFS := TPopUpMenu( GetControl('POPUPMODIFS', True) );
  POPUPTVA    := TPopUpMenu( GetControl('POPUPTVA', True) );

  // zones d'entête
  E_JOURNAL        := THValComboBox( GetControl( 'E_JOURNAL',       True ) )  ;
  E_NATUREPIECE    := THValComboBox( GetControl( 'E_NATUREPIECE',   True ) )  ;
  E_DATECOMPTABLE  := THEdit(        GetControl( 'E_DATECOMPTABLE', True ) )  ;
  E_DEVISE         := THValComboBox( GetControl( 'E_DEVISE',        True ) )  ;
  E_ETABLISSEMENT  := THValComboBox( GetControl( 'E_ETABLISSEMENT', True ) )  ;
  E_NUMEROPIECE    := THNumEdit(     GetControl( 'E_NUMEROPIECE',   True ) )  ;
  PERIODE          := THValComboBox( GetControl( 'PERIODE',         True ) )  ;
  FOLIO            := THValComboBox( GetControl( 'FOLIO',           True ) )  ;

  // boutons de navigation invisible en EAGL
{$IFNDEF EAGLCLIENT}
  if not Assigned( FQMul ) then
{$ENDIF EAGLCLIENT}
    begin
    SetControlVisible('BFIRST', False) ;
    SetControlVisible('BPREV',  False) ;
    SetControlVisible('BNEXT',  False) ;
    SetControlVisible('BLAST',  False) ;
    end ;

  // Creation de l'objet de gestion de saisie en grille
  TSP := TSaisiePiece.Create( Ecran, FListe, FPieceCpt ) ;
  TSP.SetGrilleComplement(FGCompl);
  TSP.TSGuide.SetFlashGuide( GetControl('SA_MESSAGE') ) ;

{$IFDEF SAISIEDOC}
  // Déclaration des composants de gestion du viewer de documents
  InitViewer ;
{$ENDIF SAISIEDOC}

  // Liste des boutons à gérer : (TToolBarButton97)
  BParamListe := TToolBarButton97( GetControl( 'BParamListe', True ) ) ;
  BValider    := TToolBarButton97( GetControl( 'BValider', True ) ) ;
  BSolde      := TToolBarButton97( GetControl( 'BSolde', True ) ) ;
  BEche       := TToolBarButton97( GetControl( 'BEche', True ) ) ;
  BVentil     := TToolBarButton97( GetControl( 'BVentil', True ) ) ;
  BGenereTVA  := TToolBarButton97( GetControl( 'BGenereTVA', True ) ) ;
  BControlTVA := TToolBarButton97( GetControl( 'BControleTVA', True ) ) ;
  BModifTVA   := TToolBarButton97( GetControl( 'BModifTVA', True ) ) ;
  BComplement := TToolBarButton97( GetControl( 'BComplement', True ) ) ;
  BRechercher := TToolBarButton97( GetControl( 'BRechercher', True ) ) ;
  BProrata    := TToolBarButton97( GetControl( 'BProrata', True ) ) ;
  BMenuTva    := TToolBarButton97( GetControl( 'BMenuTva', True) );
  BScan       := TToolBarButton97( GetControl( 'BScan', True) ) ;
  BDevise     := TToolBarButton97( GetControl( 'BDevise', True) );

  BZoomImmo   := TToolBarButton97( GetControl( 'BZoomImmo', True) );

  // Affichage du bouton de scan
//  SetControlVisible('BSCAN', ( VH^.OkModGed ) ) ;

  TFVierge(Ecran).BFerme.Cancel    := False ;
  TFVierge(Ecran).BValider.Default := False ;

end;

procedure TOF_CPSAISIEPIECE.InitEvenements;
begin

  BValider.OnClick := BValiderClick ;

  // Bouton de navigation
{$IFNDEF EAGLCLIENT}
  if Assigned( FQMul ) then
{$ENDIF EAGLCLIENT}
    begin
    TToolBarButton97( GetControl('BFIRST', True) ).OnClick := BFirstClick ;
    TToolBarButton97( GetControl('BPREV', True) ).OnClick  := BPrevClick ;
    TToolBarButton97( GetControl('BNEXT', True) ).OnClick  := BNextClick ;
    TToolBarButton97( GetControl('BLAST', True) ).OnClick  := BLastClick ;
    end ;

  // EVT Entete
  E_JOURNAL.OnChange     := JOURNALChange ;
  E_JOURNAL.OnExit       := JOURNALExit ;
  E_NATUREPIECE.OnExit   := NATUREPIECEExit ;
  E_DATECOMPTABLE.OnExit := DATECOMPTABLEExit ;
  E_DEVISE.OnExit        := DEVISEExit ;
  E_ETABLISSEMENT.OnExit := ETABLISSEMENTExit ;
  Periode.OnExit         := PeriodeExit ;
  Folio.OnExit           := FolioExit ;

  // EVT GRille
  TSP.OnUserRowEnter         := TSP_RowEnter ;
  TSP.OnUserRowExit          := TSP_RowExit ;
  TSP.OnUserCellEnter        := TSP_CellEnter ;
  TSP.OnUserCellExit         := TSP_CellExit ;
  TSP.OnUserDblClick         := TSP_DblClick ;
  TSP.OnUserGridEnter        := TSP_GridEnter ;
  TSP.OnUserKeyDown          := TSP_KeyDown ;

  // EVT FORM
  TFVierge(Ecran).OnKeyDown       := FormKeyDown ;
  TFVierge(Ecran).OnCloseQuery    := FormCloseQuery ;
  TFVierge(Ecran).OnAfterFormShow := OnAfterFormShow ;
  TFVierge(Ecran).OnMouseDown     := FormMouseDown ;
  PCEntete.OnEnter                := PCEnteteEnter ;

  // EVT MENUS POPUPS
  PopS.OnPopup        := OnPopupPOPS ;
  PopupZoom.OnPopup   := OnPopupZooms ;
  PopupGuide.OnPopup  := OnPopupGuide ;
  PopupLignes.OnPopup := OnPopupLignes ;
  PopupModifs.OnPopup := OnPopupModifs ;
  PopupTva.OnPopup    := OnPopupModifs ;
  InitPopupZooms  ( False ) ;
  InitPopupGuide  ( False ) ;
  InitPopupLignes ( False ) ;
  InitPopupModifs ( False ) ;
  InitPopupTva ;

  // EVT BOUTON
  BParamListe.OnClick := BParamListeClick ;
  BSolde.OnClick      := BSoldeClick ;
  BVentil.OnClick     := BVentilClick ;
  BEche.OnClick       := BEcheClick ;
  BGenereTVA.OnClick  := BGenereTVAClick ;
  BControlTVA.OnClick := BControlTVAClick ;
  BModifTVA.OnClick   := BModifTVAClick ;
  BComplement.OnClick := BComplementClick ;
  BRechercher.OnClick := BRechercherClick ;
  BProrata.OnClick    := BProrataClick ;
  BScan.OnClick       := BScanClick ;
  BDevise.OnClick     := BDeviseClick ;
  BZoomImmo.OnClick   := ZoomImmoClick ;

{$IFDEF SAISIEDOC}
  TToolBarButton97( GetControl('BAGRANDIR', True) ).OnClick := BAgrandirClick ;
  TToolBarButton97( GetControl('BREDUIRE', True) ).OnClick := BReduireClick ;
  SetControlVisible('BAGRANDIR', True ) ;
  SetControlVisible('BREDUIRE',  False ) ;
{$ELSE SAISIEDOC}
  SetcontrolVisible( 'BAGRANDIR', False ) ;
  SetcontrolVisible( 'BREDUIRE',  False ) ;
{$ENDIF SAISIEDOC}


end;

procedure TOF_CPSAISIEPIECE.InitPresentation;
begin

  // Contexte d'aide + Titre
  Ecran.HelpContext := 7246100 ;
  SetTitre ;

  // Raccourci clavier sur TLabel:
  SetControlText('TE_DEVISE', '&Devise');
//  SetControlText('TE_ETABLISSEMENT', '&Etablissement');

  // Bouton de paramétrage des masque en saisie sur code spécif
  if not EstSpecif('51212') then
    BParamListe.Visible := False ;

  // Seule la 1ère page de l'onglet est affichée
//  SetActiveTabSheet('TSENTETE1') ;

  // Paramétrage des zones d'entête
  ParamJournal ;
  ParamPeriode ;
  ParamNaturePiece ;
  ParamFolio ;

  // Paramétrage des zones de pieds
  if not FBoCptSoldeP then
    begin
    SetControlVisible( 'SA_SOLDEGPER', False ) ;
    SetControlVisible( 'SA_SOLDETPER', False ) ;
    end ;

  if not FBoJalSoldeP then
    begin
    SetControlVisible( 'SA_JALSOLDE', False ) ;
    SetControlVisible( 'SA_JALDEBIT', False ) ;
    SetControlVisible( 'TSA_JALSOLDE', False ) ;
    SetControlVisible( 'SA_DEVISESOLDE', False ) ;
//    SetControlText( 'SA_JALCREDIT', False ) ;
    SetControlText( 'TSA_JALTOTAL', 'Solde Piece' ) ;
    end ;

  // Affichage des devises
  SetControlText('SA_DEVISEJAL',   V_PGI.SymbolePivot ) ;
  SetControlText('SA_DEVISESOLDE', V_PGI.SymbolePivot ) ;

{$IFDEF AMORTISSEMENT}
  BZoomImmo.Visible         := ((VH^.OkModImmo) or (V_PGI.VersionDemo)) ;
{$ELSE}
  BZoomImmo.Visible         := FALSE ;
{$ENDIF}

end;

procedure TOF_CPSAISIEPIECE.InitVariables;
var lInfoEcr  : TInfoEcriture ;
begin

  // Variables initalisées par les paramètres d'appels...
  FModalResult         := mrNone ;

  // Var de la pièce d'engagement
  PieceEngage    := Nil ;
  FBoChargePiece := False ;
  FStatutSaisie  := ssEditEntete ;

  // Objet d'affichage des messages
  FMessCompta := TMessageCompta.Create( Ecran.Caption, msgSaisiePiece ) ;

  // Hitoerique des pièce dejà saisie
  FHistorique          := TList.Create ;

  // Affectation des objets de gestion
  FPieceCpt := nil ;
  if Assigned( TFVierge(Ecran).laTof.laTob ) then
    FPieceCpt := TPieceCompta( TFVierge(Ecran).laTof.laTob ) ;
  if not Assigned( FPieceCpt ) then
    begin
    lInfoEcr         := TInfoEcriture.Create( Dossier ) ;
    FPieceCpt        := TPieceCompta.CreerPiece( lInfoEcr ) ;
    FBoDetruirePiece := True ;
    end ;
  FPieceCpt.SetVentilDefautOff ;

  FTobReglt := Tob.Create('$PIECESREGLT', nil, -1) ;

  // Indicateurs de présentation des soldes
  FBoJalSoldeP := GetParamSocSecur('SO_CPJALSOLDEPER', False ) ;
  FBoCptSoldeP := GetParamSocSecur('SO_CPCPTSOLDEPER', False ) ;

end;


procedure TOF_CPSAISIEPIECE.AccesEntete;
var lBoEntete : Boolean ;
begin

  // Gestion des zones d'entête
  if Piece.ModeSaisie = msPiece then
    begin
    lBoEntete := (GetAction=taCreat) (*and (Piece.Action = taCreat)*) and (Piece.Count<=1) and (not Piece.EstRemplit(1)) ;

    E_JOURNAL.Enabled         := lBoEntete ;
    E_NATUREPIECE.Enabled     := (Piece.Action = taCreat) ;
    E_DATECOMPTABLE.Enabled   := lBoEntete ;

    Case Piece.Action of
      taCreat : SetControlText( 'TE_NUMEROPIECE', 'N° Provisoire' ) ;
      else      SetControlText( 'TE_NUMEROPIECE', 'N° Définitif' ) ;
      end ;

    end
  else
    begin
    lBoEntete           := (GetAction=taCreat) and (Piece.Count<=1) and (not Piece.EstRemplit(1)) ; //and not Piece.ModifEnCours ;
    E_JOURNAL.Enabled   := lBoEntete ;
    Periode.Enabled     := lBoEntete ;
    Folio.Enabled       := lBoEntete ;
    end ;

  // Devise
  AccesDevise ;
  
  // Gestion de l'établissement :
  SetControlEnabled('E_ETABLISSEMENT', (Piece.Action = taCreat) and (PContexte.EtablisCpta) and (EtabForce='') ) ;

  // Modification du titre
  SetTitre ;

end;

procedure TOF_CPSAISIEPIECE.BFirstClick(Sender: TObject);
begin
  if FQMul.Bof then Exit ;
  {JP 31/05/07 : FQ 17754 : déplacement en multi sociétés}
  ChangePiece(Sender);
end;

procedure TOF_CPSAISIEPIECE.BLastClick(Sender: TObject);
begin
  if FQMul.Eof then Exit ;
  {JP 31/05/07 : FQ 17754 : déplacement en multi sociétés}
  ChangePiece(Sender);
end;

procedure TOF_CPSAISIEPIECE.BNextClick(Sender: TObject);
begin
  if FQMul.Eof then Exit ;
  {JP 31/05/07 : FQ 17754 : déplacement en multi sociétés}
  ChangePiece(Sender);
end;

procedure TOF_CPSAISIEPIECE.BPrevClick(Sender: TObject);
begin
  if FQMul.Bof then Exit ;
  {JP 31/05/07 : FQ 17754 : déplacement en multi sociétés}
  ChangePiece(Sender);
end;

procedure TOF_CPSAISIEPIECE.ChargePieceFromMul( AvecRefresh : Boolean ) ;
//var lStSQL : String ;
begin

  // Ujniquement depuis un Mul...
  if not assigned(FQMul) then Exit ;
{
  lStSQL :=   'E_JOURNAL="'         + FQMul.FindField('E_JOURNAL').AsString                       + '" AND '
              + 'E_EXERCICE="'      + FQMul.FindField('E_EXERCICE').AsString                      + '" AND '
              + 'E_DATECOMPTABLE="' + USDATETIME( FQMul.FindField('E_DATECOMPTABLE').AsDateTime ) + '" AND '
              + 'E_NUMEROPIECE='    + IntToStr( FQMul.FindField('E_NUMEROPIECE').AsInteger )      + ' AND '
              + 'E_QUALIFPIECE="'   + FQMul.FindField('E_QUALIFPIECE').AsString                   + '"' ;

  // Chargement graphique
  ChargePiece( lStSQL ) ;
}


  // TEST si on est sur la même pièce !
  if ( Piece.GetEnteteS('E_JOURNAL') = FQMul.FindField('E_JOURNAL').AsString ) and
     ( Piece.GetEnteteI('E_PERIODE') = GetPeriode( FQMul.FindField('E_DATECOMPTABLE').AsDateTime ) ) and
     ( Piece.GetEnteteS('E_QUALIFPIECE') = FQMul.FindField('E_QUALIFPIECE').AsString ) and
     ( Piece.GetEnteteI('E_NUMEROPIECE') = FQMul.FindField('E_NUMEROPIECE').AsInteger )
     then Exit ;

  if Piece.ModifEnCours then
    if PGIAskCancel( TSP.MessageCompta.GetMessage( RC_ABANDON ), Ecran.Caption ) <> mrYes then
      Exit ;

  if Piece.EstLockFolio then
    Piece.unlockFolio( True ) ;

  Piece.videPiece ;
  Piece.InitPiece( FQMul.FindField('E_JOURNAL').ASstring,             // journal
                   FQMul.FindField('E_DATECOMPTABLE').AsDateTime,     // Date
                   '',                                                // Nature
                   '',                                                // devise
                   '',                                                // Etablissement
                   FQMul.FindField('E_QUALIFPIECE').AsString          // qualification
                    );
  Piece.PutEntete('E_NUMEROPIECE',   FQMul.FindField('E_NUMEROPIECE').AsInteger );
  ChargePiece ;

end;

procedure TOF_CPSAISIEPIECE.majBoutonNav;
begin
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  if Not Assigned (FQMul) then Exit ;
  SetControlEnabled('BFIRST', not FQMul.Bof) ;
  SetControlEnabled('BPREV',  not FQMul.Bof) ;
  SetControlEnabled('BNEXT',  not FQMul.Eof) ;
  SetControlEnabled('BLAST',  not FQMul.Eof) ;
  {$ENDIF EAGLCLIENT}
end;



procedure TOF_CPSAISIEPIECE.MajSoldePiece ;
Var lTotalC  : Double ;
    lTotalD  : Double ;
    lInDec   : integer ;
    lSolde   : Double ;
    lStSymb  : string ;
begin

  // Infos totaux / solde de la pièce :
  if Piece.ModeSaisie = msPiece
    then SetControlText( 'TSA_TOTAL', 'Total Pièce' )
    else SetControlText( 'TSA_TOTAL', 'Total Folio' ) ;

  SetControlText( 'SA_TOTALCREDIT' , '0' ) ;
  SetControlText( 'SA_TOTALDEBIT' ,  '0' ) ;

  if not FBoJalSoldeP then
    SetControlText( 'SA_JALCREDIT' ,   '0' ) ;


  if ( Piece = nil ) then Exit ;

  lTotalC  := Piece.GetTotalCredit( (Piece.GetEnteteS('E_DEVISE')<>V_PGI.DevisePivot) ) ;
  lTotalD  := Piece.GetTotalDebit( (Piece.GetEnteteS('E_DEVISE')<>V_PGI.DevisePivot) ) ;

  if (Piece.GetEnteteS('E_DEVISE')=V_PGI.DevisePivot) then
    begin
    lInDec  := V_PGI.OkDecV ;
    lStSymb := V_PGI.SymbolePivot ;
    end
  else
    begin
    lInDec  := Piece.Devise.Decimale ;
    lStSymb := Piece.Devise.Symbole ;
    end ;

  // Maj de la tob Pied
  SetControlText('SA_TOTALDEBIT',  StrFMontant( lTotalD,  15, lInDec, '', True ) ) ;
  SetControlText('SA_TOTALCREDIT', StrFMontant( lTotalC,  15, lInDec, '', True ) ) ;
  SetControlText('SA_DEVISE',      lStSymb ) ;

  if not FBoJalSoldeP then
    begin
    lSolde  := Arrondi( lTotalD- lTotalC, lInDec ) ;
    SetControlText( 'SA_JALCREDIT', StrFMontant( lSolde,   15, lInDec, '', True ) ) ;
    SetControlText( 'SA_DEVISEJAL', lStSymb ) ;
    end ;

end;


procedure TOF_CPSAISIEPIECE.DATECOMPTABLEExit(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;

  // Vérification de la date
  if ( not IsValidDate( GetControlText('E_DATECOMPTABLE') ) ) or
     ( not Piece.IsValidDateCreation( StrToDate( GetControlText('E_DATECOMPTABLE') ) ) )
    then SetControlText( 'E_DATECOMPTABLE', Piece.GetEnteteS('E_DATECOMPTABLE') )
    else begin
         Piece.PutEntete('E_DATECOMPTABLE', StrToDate( GetControlText('E_DATECOMPTABLE') ) ) ;
         majTypeExo ; // MAJ du type d'exo (TYPEEXO)
         TSP.AfficheLignes ;
         end ;

end;

procedure TOF_CPSAISIEPIECE.DEVISEExit(Sender: TObject);
//var lDtDateC : TDateTime ;
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  Piece.PutEntete('E_DEVISE', E_DEVISE.Value ) ;

  // raffraichissement format devise
  // MAJ grille suivant paramètrage de la liste
  TSP.RechercheMasque ;
  TSP.SetFormatMontant ;
  TSP.AfficheLignes ;

  // Test taux devise
//  lDtDateC := Piece.GetEnteteDt('E_DATECOMPTABLE') ;
  if (E_DEVISE.Enabled) and (E_DEVISE.Value<>'') and (E_DEVISE.Value<>V_PGI.DevisePivot)
     //and ( ( Piece.Devise.DateTaux <> lDtDateC ) or ( not Piece.IsValidTaux ) )
    then   AvertirPbTaux ;

  // raffraichissement de la devise
  SetControlText( 'SA_DEVISE',    Piece.Devise.Symbole ) ;
  if not FBoJalSoldeP then
    SetControlText( 'SA_DEVISEJAL', Piece.Devise.Symbole ) ;

end;

procedure TOF_CPSAISIEPIECE.ETABLISSEMENTExit(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  Piece.PutEntete('E_ETABLISSEMENT', E_ETABLISSEMENT.Value ) ;

  // MAJ grille suivant paramètrage
  TSP.RechercheMasque ;
  TSP.AfficheLignes ;

  if (Screen.ActiveControl = TSP.LaGrille) or (Screen.ActiveControl = PCEntete)
{$IFDEF SAISIEDOC}
     or ( (TSP.ViewerActif) and (Screen.ActiveControl = TSP.TSDoc.WebBrowser ) )
{$ENDIF SAISIEDOC}
    then
      begin
      // Chargement du folio
//      if Piece.ModeSaisie in [msBor,msLibre] then
//        ChargePiece ;
      TSP.SetFocus ;
      end ;
//  TSP.LaGrille.InplaceEditor.Setfocus ;
//  TSP.LaGrille.InplaceEditor.Invalidate ;
//  TSP.LaGrille.ShowEditor ;

end;

procedure TOF_CPSAISIEPIECE.JOURNALExit(Sender: TObject);
begin

  if ( csDestroying in Ecran.ComponentState) then Exit ;

  // Test mini
  if GetControlText('E_JOURNAL')='' then
    begin
    PGIBox('La saisie du journal est obligatoire', Ecran.Caption) ;
    E_Journal.SetFocus ;
    Exit ;
    end ;

end;

procedure TOF_CPSAISIEPIECE.NATUREPIECEExit(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;

  Piece.PutEntete('E_NATUREPIECE', E_NATUREPIECE.Value ) ;

  // MAJ grille suivant paramètrage de la liste
  TSP.RechercheMasque ;
  TSP.AfficheLignes ;

end;

procedure TOF_CPSAISIEPIECE.AccesDevise;
begin

  // Le journal doit être connu
  if ( GetControlText('E_JOURNAL')<>'' ) and
     ( InfoEcr.LoadJournal( GetControlText('E_JOURNAL') ) )

    then SetControlEnabled('E_DEVISE' , (Piece.Action = taCreat)
                                         and not Piece.EstRemplit(1)
                                         and ( InfoEcr.Journal.GetValue('J_MULTIDEVISE') = 'X' ) ) 
    else SetControlEnabled('E_DEVISE' , False ) ;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Compta
Créé le ...... : 06/12/2004
Modifié le ... :   /  /
Description .. : Mise à jour de l'affichage des zones d'entête
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSAISIEPIECE.ChargeEntete;
begin
{$IFDEF SBO}
//TSP.AddEvenement('ECRAN >> ChargeEntete' ) ;
{$ENDIF}

  if Piece.GetEnteteS('E_JOURNAL') = '' then
    begin
    PgiBox( TraduireMemoire('Le journal n''est pas renseigné !'), Ecran.Caption ) ;
    E_JOURNAL.SetFocus;
    Exit ;
    end ;

  SetJournal( Piece.GetEnteteS('E_JOURNAL') ) ;

  ZappePieceFolio ;
  TSP.SwapModeSaisie ( Piece.ModeSaisie ) ;

  ParamNaturePiece ;

  // Zones spécifiques
  Case Piece.ModeSaisie of
    msPiece : begin
              // Nature / Date / Numéro
              SetControlText( 'E_NATUREPIECE',   Piece.GetEnteteS('E_NATUREPIECE')                 ) ;
              SetControlText( 'E_DATECOMPTABLE', DateToStr( Piece.GetEnteteDt('E_DATECOMPTABLE') ) ) ;
              SetControlText( 'E_NUMEROPIECE',   Piece.GetEnteteS('E_NUMEROPIECE')                 ) ;
              end ;

    else // Bor et libre
              begin
              // Période du folio
              Periode.ItemIndex := Periode.Values.IndexOf( DateToStr( DebutDeMois( Piece.GetEnteteDt('E_DATECOMPTABLE') ) ) ) ;
              // Numéro de folio
              ParamFolio ;
              Folio.Text := Piece.GetEnteteS('E_NUMEROPIECE') ;
              end ;

    end ;

  SetControlText( 'E_DEVISE',        Piece.GetEnteteS('E_DEVISE') ) ;        // Devise
  SetControlText( 'E_ETABLISSEMENT', Piece.GetEnteteS('E_ETABLISSEMENT') ) ; // Etablissement

  majTypeExo ;
  RefreshModeTva ;


end;

{***********A.G.L.***********************************************
Auteur  ...... : Compta
Créé le ...... : 06/12/2004
Modifié le ... :   /  /
Description .. : Mise à jour de l'affichage des zones de pied
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSAISIEPIECE.refreshPCPied;
begin
  MajSoldePiece ;
  MajInfosLigne( TSP.Row ) ;
end;

procedure TOF_CPSAISIEPIECE.ChargePiece( vBoReinit : Boolean = False);
var lStReq: string ;
begin

{$IFDEF SBO}
//TSP.AddEvenement('ECRAN >> ChargePiece     Reinit : ' + BoolToStr_( vBoReinit ) ) ;
{$ENDIF}

  if FBoChargePiece then Exit ;
  if FStatutSaisie = ssEditGrille then Exit ;

  FBoChargePiece := True ;

  // ------------------------------------------
  // --- En création, démarage de la saisie ---
  // ------------------------------------------
  if (vBoReinit) then
    begin

    // Si l'entête n'est pas complète, il faut la finaliser...
    CompletePiece;

    if Piece.EstLockFolio then
      Piece.unlockFolio( True ) ;

//    Piece.InitSaisie ;
    TSP.InitSaisie ;
{    // On vide la pièce
    Piece.videPiece ;
    // Valeurs par défaut
    if Piece.ModeSaisie = msPiece then
      Piece.AttribNumeroTemp ;
    // init scénario
    Piece.InitEnteteScenario ;
}    // ré-init solde
    ReinitSoldes ;
    end
  else
  // ----------------------------------------
  // --- Sinon chargement données en base ---
  // ----------------------------------------
    begin

    // === chargement de la pièce ===
    if FBoMajDB then
      begin
      Piece.LoadFromSQL ;
      InfoEcr.Compte.GetCompteParFolio( Piece.GetWhereSQL ) ;

      // === Spécifiques au mode bordereau ===
      if (Piece.ModeSaisie <> msPiece) and (Piece.Action=taModif) and (Action<>taConsult)  then
        begin
        if not Piece.LockFolio( True ) then
          Piece.Action := taConsult ;
        end ;

      // Test enreg modifiable
      if not Piece.EstEnregModifiable then
        Piece.Action := taConsult ;

      end ;

    // JP 12/05/06 : Si le TPieceCompta est en lecture seule, on l'impose à la fiche
    if Action = taConsult
      then Piece.Action := taConsult ;

    // === Spécifiques au mode pièce ===
    if ( Piece.ModeSaisie = msPiece ) then
    begin

      // --- Recherche d'un éventuel lot de reglt multi-soc / multi-étab
      CChargeLiensReglt ( Piece, FTobReglt ) ;

      // --- Gestion des BAP
      if Piece.GestionBapOk and (Piece.Action = taModif) then
        if ExisteSQL('SELECT BAP_STATUTBAP FROM ' + GetTableDossier(Dossier, 'CPBONSAPAYER') + ' WHERE ' +
                     'BAP_STATUTBAP <> "' + sbap_Definitif + '" AND ' +
                     'BAP_JOURNAL = "' + Piece.GetEnteteS('E_JOURNAL') + '" AND ' +
                     'BAP_EXERCICE = "' + Piece.GetEnteteS('E_EXERCICE') + '" AND ' +
                     'BAP_DATECOMPTABLE = "' + UsDateTime( Piece.GetEnteteDt('E_DATECOMPTABLE')) + '" AND ' +
                     'BAP_NUMEROPIECE = ' + Piece.GetEnteteS('E_NUMEROPIECE') + ' ')
        and

          (HShowMessage('0;' + Ecran.Caption + ';En modifiant un autre champ que le libellé ou'#13 +
                       'la référence interne ou en changeant la ventilation analytique'#13 +
                       'vous détruirez le bon à payer associé.'#13#13 +
                       'Voulez-vous passer en lecture seule ?;W;YN;Y;Y;', '', '') = mrYes) then
          Piece.Action := taConsult;

    end;
  end ;

  // ----- Si mode modif/création, ajout d'une ligne vierge pour édition -----
  if (Action <> taConsult) and ( ( Piece.Count = 0 ) or Piece.EstRemplit( Piece.Count ) ) then
    Piece.NewRecord ;

  // ----- Affichage Entete / lignes -----
  if Piece.GetEnteteI('E_NUMEROPIECE')=0 then
    vBoReinit := True ;
  ChargeEntete ;

  // MAJ grille suivant paramètrage
  TSP.RechercheMasque( True ) ;

  // Affichage grille
  TSP.AfficheLignes(0, True) ;

  // ----- Maj pied -----
  refreshPCPied ;

  // ----- Maj accès fonctions
  AccesEntete ;

{$IFDEF SAISIEDOC}
  // Affectation du groupe en cours d'édition pour gestion des docs
  if Piece.ModeSaisie = msBor
    then TSP.TSDoc.GroupeActif := Piece.GetInteger( Piece.Count, 'E_NUMGROUPEECR')
    else TSP.TSDoc.GroupeActif := 0 ;
{$ENDIF SAISIEDOC}

  // Focus
  SetInitFocus ( vBoReinit ) ;
  AccesBoutons ;

  if Engagement and ( Piece.ModeSaisie = msPiece ) then begin
    PieceEngage:=TPieceEngagement.CreerPieceEng;
    if Piece.GetEnteteI('E_NUMEROPIECE') > 0 then
      begin
      lStReq :=   ' E_JOURNAL="'    + Piece.GetEnteteS('E_JOURNAL') + '"' +
              ' AND E_EXERCICE="'   + Piece.GetEnteteS('E_EXERCICE') + '"' +
//              ' AND E_EXERCICE="'   + Piece.Contexte.GetExoDT( Piece.GetDateTime('E_DATECOMPTABLE') ) + '"' +
              ' AND E_NUMEROPIECE=' + Piece.GetEnteteS('E_NUMEROPIECE') ;
      PieceEngage.LoadFromSQL(lStReq);
      end ;
    end ;

  if FBoMajDB then
    Piece.ModifEnCours := False ;
  FBoChargePiece     := False ;


end;

procedure TOF_CPSAISIEPIECE.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var Vide         : Boolean ;
begin

{$IFDEF SAISIEDOC}
  if TSP.ViewerActif and TSP.ViewerVisible then
    TSP.ViewerKeyDown(Sender,Key,Shift ) ;
{$ENDIF SAISIEDOC}

  if not ( Screen.ActiveControl = FListe ) then
    begin
    if ( Key = VK_ESCAPE )  then
      begin
      FModalResult := mrCancel ;
      Key          := 0 ;         // SBO 10/07/2007 : FQ 20986
      ClickAbandon( False ) ;
      end ;
    exit;
   end ;

  Vide:=(Shift=[]) ;

  case Key of
    VK_F5     : if Shift=[ssShift] then
                   begin
                   Key:=0 ;
                   ZoomCompte( TSP.Col ) ;
                   exit ;
                   end ;
    VK_ESCAPE : if Vide then
                  begin
                  Key:=0 ;
                  ClickAbandon( True ) ;
                  end ;
     VK_F8     : if Shift=[ssAlt] then
                    begin
                    Key:=0 ;
//                        ClickSwapPivot ;
                    end
                     else
                      if Shift=[ssCtrl] then
                       begin
                        {$IFDEF SCANGED}
                         Key := 0 ;
                         AjouterFichierDansGed(Ecran) ;
                        {$ENDIF}
                       end ;
     VK_F9     : if Vide then
                   begin
                   Key:=0 ;
                   BGenereTvaClick(nil) ;
                   end ;
     VK_F10    : if Vide then
                   begin
                   Key:=0 ;
                   ClickValide ;
                   end ;

    VK_F11     : begin
                  Key:=0 ;
                  POPS.PopUp( Mouse.CursorPos.X, Mouse.CursorPos.Y ) ;
                  end ;
{    VK_DELETE : if Shift=[ssCtrl] then
                  begin
                  Key:=0 ;
                  LignesSupprClick ( nil ) ;
                  end ;
}    VK_SPACE  : if Vide then
                  begin
                  if TSP.GetColName( TSP.Col ) = 'E_NATUREPIECE' then
                    begin
                    Key:=0 ;
                    NextNature( TSP.Row ) ;
                    end ;
                  end ;
(*
{A1.9}49..57: if Shift=[ssAlt] then BEGIN AppelAuto(Key-48) ; Key:=0 ; END ;
{AM}     77 : if Shift=[ssAlt]  then BEGIN Key:=0 ; ClickPieceGC ; END ;
{^O}     79 : if Shift=[ssctrl] then BEGIN Key:=0 ; ClickImmo ; END ;
*)

{^A}     65 : if Shift=[ssCtrl] then    begin Key:=0 ; BVentilClick(nil) ;            end ;
{AC}     67 : if Shift=[ssAlt]  then    begin Key:=0 ; BComplementClick(nil) ;        end ;
{AD}     68 : if Shift=[ssAlt]  then    begin Key:=0 ; if E_DEVISE.CanFocus then E_DEVISE.SetFocus ; end ;
{AE}     69 : if Shift=[ssAlt]  then    begin Key:=0 ; BEcheClick(nil) ;              end ;
{AH}     72 : if Shift=[ssAlt]  then    begin Key:=0 ; ZoomTauxClick(nil) ;           end ;
{AI}     73 : if Shift=[ssCtrl]  then   begin Key:=0 ; ZoomDeviseClick(nil) ;        end ;
{AJ}     74 : if Shift=[ssAlt]  then    begin Key:=0 ; ZoomJournalClick(nil) ;        end ;
{AM}     77 : if Shift=[ssAlt]  then    begin Key:=0 ; BVisuRegltClick(nil) ;         end ;
{^O}     79 : if Shift=[ssCtrl]  then   begin Key:=0 ; ZoomImmoClick(nil) ;           end ;
{AP}     80 : if Shift=[ssAlt]  then    begin Key:=0 ; ZoomPiecesClick(nil) ;         end ;
{^R}     82 : if Shift=[ssCtrl]  then   begin Key:=0 ; ModifsRegimeClick(nil) ;       end else
{AR}            if Shift=[ssAlt]  then  begin Key:=0 ; ModifsRIBClick(nil) ;          end ;
{AS}     83 : if Shift=[ssAlt]  then    begin Key:=0 ; ZoomScenarioClick(nil) ;       end ;
{^T}     84 : if Shift=[ssCtrl] then    begin Key:=0 ; BControlTVAClick(nil) ;        end else
{AT}            if Shift=[ssAlt]  then    begin Key:=0 ; ZoomEtabClick(nil) ;           end ;
{AV}     86 : if Shift=[ssAlt]  then    begin Key:=0 ; BModifTVAClick(nil) ;          end ;

  end ;

end;

procedure TOF_CPSAISIEPIECE.TSP_CellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  MajInfosLigne( TSP.Row ) ;
end;

procedure TOF_CPSAISIEPIECE.TSP_CellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var lStColName : String ;
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;

  lStColName := TSP.GetColName( ACol ) ;

  if ( lStColName = 'E_GENERAL' ) then
    begin
    if Piece.EstRemplit( ARow ) then
      begin
      GereCutoff ( ARow ) ;
      if Piece.EstHT(ARow) then
        FStGenHT := Piece.GetString(ARow, 'E_GENERAL') ;
      end ;
    RefreshModeTva ;
    AccesEntete ;
    end
  else if ( lStColName = 'E_AUXILIAIRE' ) then
     begin
     RefreshModeTva ;
     AccesEntete ;
     end
  else if ( lStColName = 'E_DEBITDEV' ) or ( lStColName = 'E_CREDITDEV' ) then
     begin
     refreshPCPied;
     TSP.AfficheLignes ;
     end
  else if ( lStColName = 'E_LIBELLE' ) then
     begin
     if Piece.EstRemplit( ARow ) and (Piece.GetMontantDev( ARow ) = 0) then
       begin
       Piece.InitLibelle( ARow ) ;
       TSP.AfficheLignes( ARow ) ;
       end ;
     end
  else if ( lStColName = 'E_DATECOMPTABLE' ) or ( lStColName = 'E_NATUREPIECE' ) then
     begin
//     if Piece.ModeSaisie = msBor then
//       TSP.AfficheLignes ;
     end
     ;

  AccesBoutons ;

end;

procedure TOF_CPSAISIEPIECE.TSP_RowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;

  // Gestion des accès boutons
  AccesBoutons ;

  // ----- Maj du pied -----
  MajInfosLigne( TSP.Row ) ;

end;

procedure TOF_CPSAISIEPIECE.TSP_RowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;

  if Action=taConsult then Exit ;

  // Ouverture auto de la fenêtre infos complémentaires
  GereComplements( Ou ) ;

  // Ouverture auto de la fenêtre infos complémentaires
  GereEche( Ou ) ;

  // Ouverture auto de la fenêtre infos complémentaires
  GereAnal( Ou ) ;

  // Gestion des immos
  GereLesImmos(Ou) ;

  // Ouverture auto du Cutoff
  GereCutoff ( Ou ) ;

//  TSP.AfficheLignes ;

end;

procedure TOF_CPSAISIEPIECE.MajInfosLigne( vRow: Integer );
var lTobEcr : TOB ;
begin
{$IFDEF SBO}
//TSP.AddEvenement('ECRAN >> MajInfosLigne        Row : ' + IntToStr( vRow) ) ;
{$ENDIF}

  // Intialisation des zones
  InitInfosLigne ;

  // Récupération TOb sur ligne courante :
  lTobEcr := Piece.GetTob( vRow ) ;
  if lTobEcr = nil then exit ;

  // ----------------------------
  // --- Infos Journal ---
  // ----------------------------
  AfficherSoldeJal ;

  // ----------------------------
  // --- Infos Compte général ---
  // ----------------------------
  if (lTobEcr.GetValue('E_GENERAL') <> '') and InfoEcr.LoadCompte( lTobEcr.GetValue('E_GENERAL') )   then
    begin
    // Infos du général
    InfoEcr.Compte.Item.PutEcran( Ecran, PCPied ) ;
    // Affichage du solde
    AfficherSoldeGene( lTobEcr.GetValue('E_GENERAL') ) ;
    end ;

  // -------------------------------
  // --- Infos Compte Auxiliaire ---
  // -------------------------------
  if (lTobEcr.GetValue('E_AUXILIAIRE') <> '') and InfoEcr.LoadAux( lTobEcr.GetValue('E_AUXILIAIRE'))  then
    begin
    // Infos de l'auxiliaire
    InfoEcr.Aux.Item.PutEcran( Ecran, PCPied ) ;
    // Calcul du solde
    AfficherSoldeAuxi( lTobEcr.GetValue('E_AUXILIAIRE') ) ;
    end
  else ;

  // -------------------------------------
  // -- Affichage message avertissement --
  // -------------------------------------
  MajMessageLigne( vRow ) ;

end;

procedure TOF_CPSAISIEPIECE.InitInfosLigne;
var lTobVide : Tob ;
begin
  // Info général de la ligne courante
  SetControlText('SA_SOLDEG',    '');
  SetControlText('SA_SOLDEGPER', '');
  SetControlText('SA_DEVISEG',   '');

  // Info Tiers de la ligne courante
  SetControlText('SA_SOLDET',    '');
  SetControlText('SA_SOLDETPER', '');
  SetControlText('SA_DEVISET',   '');

  // Message divers (en rouge)
  SetControlText( 'SA_MESSAGE' , ''  ) ;

  // MAJ des zones personnalisées Auxi et Généraux et journal...
  lTobVide := Tob.Create('GENERAUX', nil, -1) ;
  lTobVide.InitValeurs ;
  lTobVide.PutEcran( Ecran, PCPied ) ;
  FreeAndNil( lTobVide ) ;

  lTobVide := Tob.Create('TIERS', nil, -1) ;
  lTobVide.InitValeurs ;
  lTobVide.PutEcran( Ecran, PCPied ) ;
  FreeAndNil( lTobVide ) ;

  lTobVide := Tob.Create('JOURNAL', nil, -1) ;
  lTobVide.InitValeurs ;
  lTobVide.PutEcran( Ecran, PCPied ) ;
  FreeAndNil( lTobVide ) ;

end;

procedure TOF_CPSAISIEPIECE.BParamListeClick(Sender: TObject);
begin
  if not EstSpecif('51212') then Exit ;
  if TSP.GuideActif then Exit ;
  TSP.ParamMasqueSaisie ;
end;

procedure TOF_CPSAISIEPIECE.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if FBoMajDB
    then TSP.CanClose( CanClose ) ;

  if CanClose then
    TFVierge(Ecran).FormCloseQuery( Sender, CanClose ) ;
end;

procedure TOF_CPSAISIEPIECE.BComplementClick(Sender: TObject);
{$IFNDEF GCGC}
var lTobEcr    : TOB ;
    lRComp     : R_COMP ;
    lAction    : TActionFiche ;
{$ENDIF GCGC}
begin
{$IFNDEF GCGC}

  if not BComplement.Enabled then Exit ;

  if not PiecE.EstRemplit( TSP.Row ) then Exit ;

  lTobEcr := Piece.GetTob( TSP.Row ) ;
  if lTobEcr = nil then Exit ;

  lRComp.StLibre    := 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' ;
  lRComp.StComporte := '--XXXXXXXX' ; // Pas de ModeLC
  lRComp.Conso      := True ;
  lRComp.DateC      := lTobEcr.GetValue('E_DATECOMPTABLE') ;
  lRComp.Attributs  := False ;
  if Piece.GetScenario <> nil
    then lRComp.MemoComp   := Piece.GetMemoScenario( TSP.Row )
    else lRComp.MemoComp   := nil ;
  lRComp.Origine    := -1 ;
  lRComp.TOBCompl   := CGetTobCompl(lTobEcr) ;

  lRComp.AvecCalcul  := False ;
  if Piece.Info_LoadCompte( lTobEcr.GetString('E_GENERAL') ) then
    begin
    lRComp.CutOffPer   := Piece.Info.Compte.GetValue('G_CUTOFFPERIODE') ;
    lRComp.CutOffEchue := Piece.Info.Compte.GetValue('G_CUTOFFECHUE') ;
    end ;

  if (Piece.Action = taCreat)
    then lAction := taModif
    else lAction := Piece.Action ;

  OuvreComplements   (  TSP.Row , lRComp, lAction ) ;

{$ENDIF GCGC}

end;


procedure TOF_CPSAISIEPIECE.BControlTVAClick(Sender: TObject);
begin
{$IFNDEF GCGC}

//  BControlTVA.Enabled  := ( Action <> taConsult ) and lBoLigneOk and (not ANouveau) and ( Piece.EstHT( TSP.Row ) ) ;
  if Not GetControlEnabled('BControleTVA') then Exit ;
  if Not Piece.EstRemplit( TSP.Row ) then Exit ;

  if Not Piece.EstHT( TSP.Row ) then
    begin
    TraiteErreurTva( TVA_ERRLIGNEHT ) ;
    Exit ;
    end ;

  if Piece.TrouveLigneTiers = nil then
    begin
    TraiteErreurTva( TVA_ERRNOTIERS ) ;
    Exit ;
    end ;

  CPLanceFiche_ControleTva( Piece, TSP.Row ) ;

{$ENDIF GCGC}

end;

procedure TOF_CPSAISIEPIECE.BEcheClick(Sender: TObject);
begin
  if not GetControlEnabled('BECHE') then Exit ;
  if Piece.GetValue( TSP.row, 'E_ECHE' ) <> 'X' then Exit;
  if not Piece.EstRemplit( TSP.row ) then Exit;
  if not (Piece.GetMontantDev( TSP.row ) <> 0) then Exit ;

  OuvreEche( TSP.Row ) ; //( TSP.row, False, True ) ;
end;

procedure TOF_CPSAISIEPIECE.BGenereTVAClick(Sender: TObject);
Var lInErr : Integer ;
begin
  if GetAction = taConsult then Exit ;
  if Not GetControlEnabled('BGENERETVA') then Exit ;

  // Validatino de la ligne en cours de saisie
  if Piece.EstRemplit( TSP.Row ) then
    if not TSP.ValideSortieLigne then Exit ;

  // ----- Traitement de la pièce -----
  lInErr := Piece.GenereTVA( TSP.Row ) ;
  if lInErr >= 0 then
    TraiteErreurTva ( lInErr ) ;

  // ----- Maj de la grille -----
  TSP.AfficheLignes ;

  // ----- Maj accès fonctions
  AccesBoutons ;

  // ----- Maj entête et pied -----
  refreshPCPied ;

end;

procedure TOF_CPSAISIEPIECE.BModifTVAClick(Sender: TObject);
{$IFNDEF GCGC}
Var lTob         : TOB ;
    lTobTmp      : TOB ;
    lMOD         : TMOD ;
    lAction      : TActionFiche ;
    RTVA         : Enr_Base ;
    lBoEstClient : Boolean ;
{$ENDIF}
BEGIN
// ??????? {#TVAENC} //??????
{$IFNDEF GCGC}
  // Gestion de la TVA ?
  if Not ( ((Simul='N') or (Simul='S') or (Simul='R')) and ( PContexte.OuiTvaEnc ) ) then Exit ;

  // Accès autorisé ?
  if Not GetControlEnabled('BMODIFTVA') then Exit ;

  if Piece.IsOut( TSP.Row ) then Exit ;

  lTob := Piece.GetTob( TSP.Row ) ;

  if not Piece.EstLigneModifiable( TSP.Row )
    then lAction := taConsult
    else lAction := Piece.Action ;

  // pour les lignes de HT
  if Piece.EstHT( TSP.Row ) then
    begin
    if lTob.GetValue('E_REGIMETVA')='' then
      Piece.AffecteTVA( TSP.Row ) ;

    lTobTmp := TOB.Create( 'ECRITURE', nil, -1 ) ;
    lTobTmp.Dupliquer( lTob, True, True ) ;
//    lOBM.PutValue( 'E_BLOCNOTE', lTob.GetValue('E_BLOCNOTE') ) ;

    if InfosTvaEnc( lTobTmp, lAction ) then
      begin
//      lTob.Dupliquer( lTobTmp, True, True ) ;
      if lTob.GetString('E_TVA')<>lTobTmp.GetString('E_TVA') then
        Piece.PutValue( TSP.Row, 'E_TVA', lTobTmp.GetString('E_TVA') ) ;
      if lTob.GetString('E_TVAENCAISSEMENT')<>lTobTmp.GetString('E_TVAENCAISSEMENT') then
        begin
        if lTobTmp.GetString('E_TVAENCAISSEMENT') = 'X'
          then Piece.SetTvaMode( TSP.Row, tvaEncais, True)
          else Piece.SetTvaMode( TSP.Row, tvaDebit, True) ;
        end ;
      FreeAndNil( lTobTmp ) ;
      TSP.AfficheLignes ;
      AccesBoutons ;
      end ;

    end
  // pour les lignes de Tiers
  else if Piece.EstTiers( TSP.Row ) then
    begin
    // TVA DIV
    if Piece.GetSorteTVA( TSP.Row ) <> stvDivers then
      begin
      lAction := taConsult ;
      // Y'a-t-il un seul tiers ?
      if not Piece.UnSeulTiers( TSP.Row ) then Exit ;
      end ;

      // Détermination client / fournisseur...
      lBoEstClient := Piece.EstCli( TSP.Row ) ;
      if (not lBoEstClient) and (not Piece.EstFou( TSP.Row )) then Exit ;

    // Vérif général collectif
    if (lTob.GetValue('E_GENERAL')='') or (Not EstCollFact( lTob.GetValue('E_GENERAL') ) ) then Exit ;

    // Echéances
    lMOD := TMOD.Create ;
    lMOD.ModR := Piece.GetModeRegl( TSP.Row ) ;

    // Paramètre TVA
    RTVA.Regime  := lTob.GetValue('E_REGIMETVA') ;
    RTVA.Nature  := Piece.GetEnteteS('E_NATUREPIECE') ;
    RTVA.Client  := lBoEstClient ;
    RTVA.Action  := lAction ;
    RTVA.CodeTva := lTob.GetValue('E_TVA') ;

    // Saisie
    if SaisieBasesHT( lMOD, RTVA ) then
      begin
      Piece.PutValue( TSP.Row, 'E_TVA', RTVA.CodeTva ) ;
      // MAJ des infos tva sur les écheances
      Piece.SetModeRegl( TSP.Row, lMOD.ModR ) ;
      TSP.AfficheLignes ;
      AccesBoutons ;      
      end ;

    end ;

{$ENDIF}
end;

procedure TOF_CPSAISIEPIECE.BRechercherClick(Sender: TObject);
begin
  if not BRechercher.Enabled then Exit ;
  TSP.Recherche ;
end;

procedure TOF_CPSAISIEPIECE.BSoldeClick(Sender: TObject);
begin
//  if Action = taConsult                   then Exit ;
  if not GetControlEnabled('BSOLDE')      then Exit ;

  TSP.SoldePiece ;

  // Totaux
  refreshPCPied;
  // Maj accès fonctions
  AccesBoutons ;

end;

procedure TOF_CPSAISIEPIECE.BVentilClick(Sender: TObject);
begin

  if not GetControlEnabled('BVENTIL') then Exit ;
  if Piece.GetValue( TSP.row, 'E_ANA' ) <> 'X' then Exit;
  if not Piece.EstRemplit( TSP.row ) then Exit;
  if not (Piece.GetMontantDev( TSP.row ) <> 0) then Exit ;

  OuvreAnal( TSP.row ) ;
end;


procedure TOF_CPSAISIEPIECE.OuvreAnal(vRow: Integer; vInAxe : Integer ) ;
var
 lRecA     : ARG_ANA ;
 lTob      : TOB ;
 lBoVisu   : Boolean ;
begin

 lTob         := nil ;

 try

  lTob                  := Piece.GetTob( vRow ) ;

  FillChar(lRecA, sizeof(ARG_ANA), #0);             {FP 02/05/2006}
  lBoVisu := (Action = taConsult) or ( not Piece.EstLigneModifiable( TSP.row ) ) ;
  if lBoVisu then
    lRecA.Action	        := taConsult
    else lRecA.Action     := Piece.Action ;

  lRecA.CC              := nil ; //CZompteVersTGGeneral ( InfoEcr.Compte ) ;
  lRecA.GuideA          := '' ;
  lRecA.VerifVentil     := True ;
  lRecA.ControleBudget  := false ;
  lRecA.DernVentilType  := '' ;
  lRecA.DEV             := InfoEcr.Devise.Dev ;
  lRecA.QuelEcr         := EcrGen ;
  lReCA.NumGeneAxe      := vInAxe ;
  // new dev 3946
  lRecA.VerifQte        := ( Action <> taConsult ) and TestQteAna(vRow);
  lReCA.Info            := InfoEcr ;
  lReCA.MessageCompta   := FMessCompta ;
  lReCA.EnMontant       := Piece.Contexte.CPAnaSurMontant ;

  AlloueAxe( lTOB ) ;

  // Appel de la saisie analytique
  eSaisieAnal( lTOB , lRecA ) ;

  // Test modifications ?
  if lTob.IsOneModifie and not lBoVisu then
    Piece.ModifEnCours := True ;

  // synchronisation saisie mono section
  Piece.SynchroMonoSection( vRow ) ;
  TSP.AfficheLignes( vRow ) ;

  // Maj accès fonctions
  AccesBoutons ;

 finally
  if assigned(lRecA.CC) then lRecA.CC.Free;
 end;

end;

procedure TOF_CPSAISIEPIECE.ClickValide;
var lreq      : string; // fb 02/05/2006
    Pointeur  : PBeforeValidePieceCompta; // fb 02/05/2006
    BeforeValidePieceCompta: TBeforeValidePieceCompta;// fb 02/05/2006
    procedure _replaceListe ;
      begin
      TSP.Enabled := True ; // Acces grille
      TSP.SetFocus ;
      end ;

begin
  if Action = taConsult then Exit ;
  if not BValider.Enabled then Exit ;
  if TSP.GuideActif then Exit ;

  // Gestion des evènements de sortie de grille
  if ( Screen.ActiveControl.Name = TSP.LaGrille.Name ) then
    if not TSP.ValideSortieLigne then Exit ;

  TSP.Enabled := False ; // Acces grille

  if not Piece.IsValidPiece then
    begin
    _replaceListe ;
    Exit ;
    end ;

  if not Piece.IsValidTva(1) then
    begin
    // Vérification TVA :
    if PgiAsk('La TVA de la pièce est incorrecte. Confirmez-vous la validation de cette pièce ? ', Ecran.Caption ) <> mrYes then
      begin
      _replaceListe ;
      Exit ;
      end ;
    end ;

  // b FP 02/05/2006
  if Engagement then
    begin
    if not Piece.IsValidEngagement then
      begin
      _replaceListe ;
      Exit ;
      end ;
    if Piece.FieldExists('BeforeValidePieceCompta') then
      begin
      Pointeur := PBeforeValidePieceCompta(Integer(Piece.GetValue('BeforeValidePieceCompta')));
      if Assigned(Pointeur) then
        begin
        BeforeValidePieceCompta := Pointeur^;
        if not BeforeValidePieceCompta(Self, Ecran, Piece) then
          begin
          _replaceListe ;
          Exit;
          end ;
        end;
      end;
    end ;
  // e FP 02/05/2006

  // Modification de la pièce ne mémoire
  FModalResult := mrOk ;
  if not FBoMajDB then
    begin
    EcranClose ;
    Exit ;
    end ;

  // b FP 14/11/2005 FQ15575
  if (not TraiteDoublon) then // Uniqement en création !! sinon ce test soi-même....
      begin
      _replaceListe ;
      Exit;
      end ;
  // e FP 14/11/2005 FQ15575

  // Before Validation GEP
  if not V_PGI.CGEPDisabled then
    if not OnBeforeValidePieceCompta (Ecran, Piece) then
      begin
      _replaceListe ;
      Exit;
      end ;

  SourisSablier ;

  {$IFDEF SAISIEDOC}
  if Piece.ModeSaisie = msPiece then
    TSP.ChargeDocument ;
  {$ENDIF SAISIEDOC}

  if Piece.Save then
    begin

    // Deblocage du folio
    if (Piece.ModeSaisie <> msPiece) and (Piece.Action <> taConsult) then
      Piece.UnLockFolio( True )  ;

    // Tiers Payeur / Lettrage en saisie
    if GereTiersPayeurs then
      begin
      if Transactions( Piece.GenerePiecesTP, 1 ) <> oeOk then
         PGIError( 'ATTENTION : L''écriture est validée sans génération de pièces sur le tiers payeur.', Ecran.caption ) ;
      end
    else
      begin
      if Transactions( Piece.SupprimePiecesTP, 1 ) <> oeOk then
         PGIError( 'ATTENTION : L''écriture est validée sans suppression des pièces sur le tiers payeur.', Ecran.caption ) ;
      end ;

    // Lettrage à la validation
    LettrageEnSaisie ;

    if ( Piece.ModeSaisie = msPiece ) then
      try
          {On ne lance la création des BAP que si on est en création ou
           en modifiaction avec modification d'un autre champ que E_LIBELLE et E_REFINTERNE}
          if (not Piece.ConserverBap and (Piece.Action = taModif)) or
             (Piece.Action = taCreat) then begin
            {JP 05/06/06 : En création, on ne passe pas TPieceCompta.LoadFromSQL, donc GestionBapOk est à False !!}
            if (Action = taCreat) then Piece.IsPieceABap;

            {S'il y a un type de visa de paramétrer que l'on est sur un journal d'achat ...}
            if Piece.GestionBapOk then begin
              CpLanceFiche_GenereBap('N;' + Piece.GetEnteteS('E_JOURNAL') + ',' +
                                            Piece.GetEnteteS('E_EXERCICE') + ',' +
                                            DateToStr( Piece.GetEnteteDt('E_DATECOMPTABLE') ) + ',' +
                                            Piece.GetEnteteS('E_NUMEROPIECE') + ',;');
            end;
          end;
        except
          on E : Exception do
            PGIError('Impossible de créer le bon à payer avec le message : '#13#13 + E.Message);
        end;

    // b fb 02/05/2006
    if Engagement then begin
      PieceEngage.ClearDetail;
      lReq :=' E_JOURNAL="' + E_JOURNAL.value + '"' +
             ' AND E_EXERCICE="' + QuelExo(E_DATECOMPTABLE.Text) + '"' +
             ' AND E_NUMEROPIECE=' + E_NUMEROPIECE.Text;
      PieceEngage.LoadFromSQL( lReq ) ;
      PieceEngage.SaveEngage;
      FreeAndNil(PieceEngage);
      end;
    // e fb 02/05/2006

    // After Validation GEP
    if not V_PGI.CGEPDisabled then
      OnAfterValidePieceCompta (Ecran, Piece) ;

    // Stockage pour zoom pièces saisies ou édition
    StockeLaPiece ;

    // Indicateur PCL pour dernière date saisie
    if (ctxPCL in V_PGI.PGIContexte) and ( Piece.modeSaisie <> msPiece ) then
      Piece.Contexte.SetZLastInfo( Piece.GetEnteteS('E_JOURNAL'), Piece.GetEnteteDt('E_DATECOMPTABLE') ) ;

    // Réinitialisation de la saisie Mode Piece
    FStatutSaisie := ssEditEntete ;
    if GetAction=taCreat //( ( (Piece.ModeSaisie=msPiece)  and ( Action = taCreat ) ) or (Piece.ModeSaisie<>msPiece) )
      then ChargePiece( True )
      else EcranClose;

  end ;

  SourisNormale ;
  TSP.Enabled := True ; // Acces grille


end;

procedure TOF_CPSAISIEPIECE.EcranClose;
begin
  TFVierge(Ecran).BFermeClick(nil);
end;

procedure TOF_CPSAISIEPIECE.StockeLaPiece;
Var X    : RMVT ;
    P    : P_MV ;
    i    : Integer ;
    lPME : TPieceCompta ;
begin

  if ( Action <> taCreat ) then Exit  ;

  FillChar( X, Sizeof(X), #0 ) ;
  X   := Piece.GetRMVT ;
  P   := P_MV.Create ;
  P.R := X ;

  FHistorique.Add( P ) ;

  // En cas de génération du multi-établissement, stockage des pièces surl es autres établissements
  if Piece.PiecesEtab.Count > 0 then
    For i := 0 to Piece.PiecesEtab.Count - 1 do
      begin
      lPME := TPieceCompta( Piece.PiecesEtab.Objects[ i ] ) ;

      FillChar( X, Sizeof(X), #0 ) ;
      X   := lPME.GetRMVT ;
      P   := P_MV.Create ;
      P.R := X ;

      FHistorique.Add( P ) ;
      end

end;

procedure TOF_CPSAISIEPIECE.OnPopUpGuide(Sender: TObject);
begin
  InitPopupGuide ;
end;

procedure TOF_CPSAISIEPIECE.OnPopUpLignes(Sender: TObject);
begin
  InitPopupLignes ;
end;

procedure TOF_CPSAISIEPIECE.OnPopUpModifs(Sender: TObject);
begin
  InitPopupModifs ;
end;

procedure TOF_CPSAISIEPIECE.OnPopUpPOPS(Sender: TObject);
var
  lMenu  : TMenuItem ;
  lStAux : string ;
  i      : Integer ;

  function _AjoutePopup( vStNom : string ; vStTitre : string ) : TMenuItem ;
    begin
    result          := TMenuItem.Create(PopS) ;
    result.Name     := 'P' + vStNom ;
    result.Caption  := vStTitre ;
    result.Tag      := 1 ;
    end ;

begin
  InitPopup(Ecran);

  if (Action<>taConsult) and Piece.IsActiveAcc then
    begin

    //Arrêt de l'accélérateur de saisie
    lMenu := _AjoutePopup( 'ACCACTIF', 'Arrêt de l''accélérateur de saisie') ;
    lMenu.OnClick   := AccActifClick ;
    lMenu.Checked   := TSP.AccArret ;
    PopS.Items.Insert(0,lMenu) ;

    // Activation et Enregistrement des paramètres contextuels de l'accélérateur de saisi
    lStAux := Piece.GetString( TSP.Row, 'E_AUXILIAIRE')  ;
    if InfoEcr.LoadAux( lStAux ) and (InfoEcr.getString('YTC_ACCELERATEUR')='-') then
      begin
      lMenu := _AjoutePopup( 'ACCTIERS', 'Activer l''accélérateur de saisie') ;
      lMenu.OnClick := AccTiersClick ;
      PopS.Items.Insert(1,lMenu) ;
      end ;

    end ;

  for i := 0 to PopS.Items.Count - 1 do
    begin
    // Zoom Journal
    if PopS.Items[i].Name = 'ZOOMDEVISE' then
      PopS.Items[i].Caption := TraduireMemoire('Voir la devise') ;
    end ;


end;

procedure TOF_CPSAISIEPIECE.OnPopUpZooms(Sender: TObject);
begin
  InitPopupZooms ;
end;

procedure TOF_CPSAISIEPIECE.InitPopupGuide(vActivation: Boolean);
var i : Integer ;
begin

  if vActivation then AccesBoutons ;

  for i := 0 to PopUpGuide.Items.Count - 1 do
    begin
    // Appel d'un guide
    if PopUpGuide.Items[i].Name = 'GUIDEOUVRIR' then
      begin
      if vActivation
        then PopUpGuide.Items[i].Enabled := GetControlEnabled('BGUIDEOUVRIR')
        else PopUpGuide.Items[i].OnClick := GuideOuvrirClick;
      end
    // Création d'un guide
    else if PopUpGuide.Items[i].Name = 'GUIDECREER' then
      begin
      if vActivation
        then PopUpGuide.Items[i].Enabled := GetControlEnabled('BGUIDECREER')
        else PopUpGuide.Items[i].OnClick := GuideCreerClick;
      PopUpGuide.Items[i].Visible := False ;
      end;

    end ;

  // En initialisation, Mise en place des evt bouttons
  if not vActivation then
    begin
    TToolBarButton97( GetControl('BGUIDEOUVRIR', True)).OnClick := GuideOuvrirClick ;
    TToolBarButton97( GetControl('BGUIDECREER',  True)).OnClick := GuideCreerClick ;
    end ;

end;

procedure TOF_CPSAISIEPIECE.InitPopupLignes(vActivation: Boolean);
var i : Integer ;
begin

  if vActivation then AccesBoutons ;

  for i := 0 to PopUpLignes.Items.Count - 1 do
    begin
    // Insérer une ligne
    if PopUpLignes.Items[i].Name = 'LIGNESINSERT' then
      begin
      if vActivation
        then PopUpLignes.Items[i].Enabled := GetControlEnabled('BLIGNESINSERT')
        else PopUpLignes.Items[i].OnClick := LignesInsertClick;
      end
    // Supprimer une ligne
    else if PopUpLignes.Items[i].Name = 'LIGNESSUPPR' then
      begin
      if vActivation
        then PopUpLignes.Items[i].Enabled := GetControlEnabled('BLIGNESSUPPR')
        else PopUpLignes.Items[i].OnClick := LignesSupprClick;
      end
    // Libellés automatiques
    else if PopUpLignes.Items[i].Name = 'LIGNESLIBELLE' then
      begin
      if vActivation
        then PopUpLignes.Items[i].Enabled := GetControlEnabled('BLIGNESLIBELLE')
        else PopUpLignes.Items[i].OnClick := LignesLibelleClick;
      end
    // Supprimer toutes les lignes
    else if PopUpLignes.Items[i].Name = 'LIGNESRAZ' then
      begin
      if vActivation
        then PopUpLignes.Items[i].Enabled := GetControlEnabled('BLIGNESRAZ')
        else PopUpLignes.Items[i].OnClick := LignesRAZClick;
      end
       // Associer un doc pdf
       else if PopUpLignes.Items[i].Name = 'BSCANPDF' then
         begin
         if vActivation
           then PopUpLignes.Items[i].Enabled := GetControlEnabled('BSCANPDF')
           else PopUpLignes.Items[i].OnClick := AssocierPdf;
         end;

    end ;

  // En initialisation, Mise en place des evt bouttons
  if not vActivation then
    begin
    TToolBarButton97( GetControl('BLIGNESINSERT',  True)).OnClick := LignesInsertClick ;
    TToolBarButton97( GetControl('BLIGNESSUPPR',   True)).OnClick := LignesSupprClick ;
    TToolBarButton97( GetControl('BLIGNESLIBELLE', True)).OnClick := LignesLibelleClick ;
    TToolBarButton97( GetControl('BLIGNESRAZ',     True)).OnClick := LignesRAZClick ;
    TToolBarButton97( GetControl('BSCANPDF',       True)).OnClick := AssocierPdf ;
    end ;

end;

procedure TOF_CPSAISIEPIECE.InitPopupModifs(vActivation: Boolean);
var i : Integer ;
begin

  if vActivation then AccesBoutons ;

  for i := 0 to PopUpModifs.Items.Count - 1 do
    begin

    // Modif. en série libellé / réf.
    if PopUpModifs.Items[i].Name = 'MODIFSSERIE' then
      begin
      if vActivation
        then PopUpModifs.Items[i].Enabled := GetControlEnabled('BMODIFSSERIE')
        else PopUpModifs.Items[i].OnClick := ModifsSerieClick;
      end
    // Changement de régime fiscal
    else if PopUpModifs.Items[i].Name = 'MODIFSREGIME' then
      begin
      if vActivation
        then PopUpModifs.Items[i].Enabled := GetControlEnabled('BMODIFSREGIME')
        else PopUpModifs.Items[i].OnClick := ModifsRegimeClick;
      end
    // Modification du RIB
    else if PopUpModifs.Items[i].Name = 'MODIFSRIB' then
      begin
      if vActivation
        then PopUpModifs.Items[i].Enabled := GetControlEnabled('BMODIFSRIB')
        else PopUpModifs.Items[i].OnClick := ModifsRIBClick ;
      end ;
    end ;

  // En initialisation, Mise en place des evt bouttons
  if not vActivation then
    begin
    TToolBarButton97( GetControl('BMODIFSSERIE',  True)).OnClick := ModifsSerieClick ;
    TToolBarButton97( GetControl('BMODIFSREGIME', True)).OnClick := ModifsRegimeClick ;
    TToolBarButton97( GetControl('BMODIFSRIB',    True)).OnClick := ModifsRIBClick ;
    end ;

end;

procedure TOF_CPSAISIEPIECE.InitPopupZooms(vActivation: Boolean);
var i : Integer ;
begin

  if vActivation then AccesBoutons ;

  for i := 0 to PopUpZoom.Items.Count - 1 do
    begin

    // Zoom Journal
    if PopUpZoom.Items[i].Name = 'ZOOMJOURNAL' then
      begin
      if vActivation
        then PopUpZoom.Items[i].Enabled := GetControlEnabled('BZOOMJOURNAL' )
        else PopUpZoom.Items[i].OnClick := ZoomJournalClick;
      end

    // Zoom Devise
    else if PopUpZoom.Items[i].Name = 'ZOOMDEVISE' then
      begin
      if vActivation
        then PopUpZoom.Items[i].Enabled := GetControlEnabled('BZOOMDEVISE')
        else begin
             PopUpZoom.Items[i].Caption := TraduireMemoire('Voir la devise') ;
             PopUpZoom.Items[i].OnClick := ZoomDeviseClick;
             end ;
      end

    // Eoom Etablissement
    else if PopUpZoom.Items[i].Name = 'ZOOMETAB' then
      begin
      if vActivation
        then PopUpZoom.Items[i].Enabled := GetControlEnabled('BZOOMETAB')
        else PopUpZoom.Items[i].OnClick := ZoomEtabClick;
      end

    // Eoom Scénario
    else if PopUpZoom.Items[i].Name = 'ZOOMSCENARIO' then
      begin
      if vActivation
        then PopUpZoom.Items[i].Enabled := GetControlEnabled('BZOOMSCENARIO')
        else PopUpZoom.Items[i].OnClick := ZoomScenarioClick;
      end

    // Dernière écritures
    else if PopUpZoom.Items[i].Name = 'ZOOMPIECES' then
      begin
      if vActivation
        then PopUpZoom.Items[i].Enabled := GetControlEnabled('BZOOMPIECES')
        else PopUpZoom.Items[i].OnClick := ZoomPiecesClick;
      end

    // Historique des taux
    else if PopUpZoom.Items[i].Name = 'ZOOMTAUX' then
      begin
      if vActivation
        then PopUpZoom.Items[i].Enabled := GetControlEnabled('BZOOMTAUX')
        else PopUpZoom.Items[i].OnClick := ZoomTauxClick;
      end

    // Zoom lot réglement BZOOMREGLT
    else if PopUpZoom.Items[i].Name = 'ZOOMREGLT' then
      begin
      if vActivation
        then PopUpZoom.Items[i].Enabled := GetControlEnabled('BZOOMREGLT')
        else PopUpZoom.Items[i].OnClick := BVisuRegltClick;
      end

    // Zoom Généraux
    else if PopUpZoom.Items[i].Name = 'ZOOMGENE' then
      begin
      if vActivation
        then PopUpZoom.Items[i].Enabled := GetControlEnabled('BZOOMGENE')
        else PopUpZoom.Items[i].OnClick := ZoomGeneralClick;
      end

    // Zoom Tiers
    else if PopUpZoom.Items[i].Name = 'ZOOMTIERS' then
      begin
      if vActivation
        then PopUpZoom.Items[i].Enabled := GetControlEnabled('BZOOMTIERS')
        else PopUpZoom.Items[i].OnClick := ZoomTiersClick;
      end

    // Zoom Immo
    else if PopUpZoom.Items[i].Name = 'ZOOMIMMO' then
      begin
{$IFDEF AMORTISSEMENT}
      if vActivation
        then PopUpZoom.Items[i].Enabled := GetControlEnabled('BZOOMIMMO')
        else PopUpZoom.Items[i].OnClick := ZoomImmoClick;
{$ELSE  AMORTISSEMENT}
        PopUpZoom.Items[i].Visible := False ;
{$ENDIF AMORTISSEMENT}
      end

    ;

    end ;


  // Premier passage - Mise en place des evt button
  if not vActivation then
    begin
    TToolBarButton97( GetControl('BZOOMJOURNAL',  True)).OnClick := ZoomJournalClick ;
    TToolBarButton97( GetControl('BZOOMDEVISE',   True)).OnClick := ZoomDeviseClick ;
    TToolBarButton97( GetControl('BZOOMETAB',     True)).OnClick := ZoomEtabClick ;
    TToolBarButton97( GetControl('BZOOMSCENARIO', True)).OnClick := ZoomScenarioClick ;
    TToolBarButton97( GetControl('BZOOMPIECES',   True)).OnClick := ZoomPiecesClick ;
    TToolBarButton97( GetControl('BZOOMTAUX',     True)).OnClick := ZoomTauxClick ;
    TToolBarButton97( GetControl('BZOOMREGLT',    True)).OnClick := BVisuRegltClick ;
    TToolBarButton97( GetControl('BZOOMTIERS',    True)).OnClick := ZoomTiersClick ;
    TToolBarButton97( GetControl('BZOOMGENE',     True)).OnClick := ZoomGeneralClick ;
    end ;

end;

procedure TOF_CPSAISIEPIECE.EnCoursDev(Sender: TObject);
begin
  PGIInfo('Ce module est en cours de développement !', Ecran.Caption ) ;
end;

procedure TOF_CPSAISIEPIECE.GuideCreerClick(Sender: TObject);
begin
  if not GetControlEnabled('BGUIDECREER') then Exit ;
  EnCoursDev ( nil ) ;
end;

procedure TOF_CPSAISIEPIECE.GuideOuvrirClick(Sender: TObject);
begin
  if GetAction = taConsult then Exit ;
  if not ( Screen.ActiveControl = FListe ) then Exit ;
  if not GetControlEnabled('BGUIDEOUVRIR') then Exit ;

  TSP.GereGuide( TSP.Row ) ;
end;

procedure TOF_CPSAISIEPIECE.LignesInsertClick(Sender: TObject);
begin
  if Action = taConsult then Exit ;
  if not GetControlEnabled('BLIGNESINSERT') then Exit ;

  TSP.InsertRow ;
  // Maj accès fonctions
  AccesBoutons ;
  // Maj infos pied
  refreshPCPied ;
end;

procedure TOF_CPSAISIEPIECE.LignesLibelleClick(Sender: TObject);
begin
//  if ModeLC then Exit ;
  if Piece.IsOut( TSP.Row ) then Exit ;
  if (Not GetControlEnabled('BLignesLibelle')) or (Action=taConsult) then Exit ;

  TSP.ExecuteLibelleAuto( TSP.Row ) ;

end;

procedure TOF_CPSAISIEPIECE.LignesRAZClick(Sender: TObject);
begin
  if Action = taConsult then Exit ;
  if not GetControlEnabled('BLIGNESRAZ') then Exit ;

  TSP.videLignes ;

  // Maj accès fonctions
  AccesBoutons ;
  // Maj infos pied
  refreshPCPied ;
end;

procedure TOF_CPSAISIEPIECE.AssocierPdf(Sender: TObject);
begin
  if Action = taConsult then Exit ;
  if not GetControlEnabled('BSCANPDF') then Exit ;
  {$IFDEF SCANGED}
  AjouterFichierDansGed(Ecran) ;
  {$ENDIF}
end ;

procedure TOF_CPSAISIEPIECE.LignesSupprClick(Sender: TObject);
var
  {b FP 02/05/2006}
  Pointeur:        PBeforeSuppLigne;
  BeforeSuppLigne: TBeforeSuppLigne;
  {e FP 02/05/2006}
begin
  if Action = taConsult then Exit ;
  if not GetControlEnabled('BLIGNESSUPPR') then Exit ;

  {b FP 02/05/2006}
  if Piece.FieldExists('BeforeSuppLigne') then
    begin
    Pointeur := PBeforeSuppLigne(Integer(Piece.GetValue('BeforeSuppLigne')));
    if Assigned(Pointeur) then
      begin
      BeforeSuppLigne := Pointeur^;
      if not BeforeSuppLigne(Self, Piece, TSP.Row) then
        Exit;
      end;
    end;
  {e FP 02/05/2006}

{$IFDEF SCANGED}
  if ( Piece.RechGuidId( TSP.Row ) <> '' ) then
    begin
    Case PGIAskCancel('Il existe une image ou un fichier associé , voulez-vous le supprimer ?') of
      mrYes :    Piece.SupprimeLeDocGuid( TSP.Row ) ;
      mrCancel : Exit ;
      end ;
    end ;
{$ENDIF}

  // suppression ligne
  TSP.DeleteRow ;

  // Maj accès fonctions
  AccesBoutons ;

  // Maj infos pied
  refreshPCPied ;

end;

procedure TOF_CPSAISIEPIECE.ModifsRegimeClick(Sender: TObject);
var lStNewReg : String ;
    lStCurReg : String ;
    lStTitre  : String ;
begin
//if Not BChoixRegime.Enabled then Exit ;
  if not GetControlEnabled('BMODIFSREGIME') then Exit ;
  if Action = taConsult then Exit ;
  if Piece.ModeSaisie<>msPiece then Exit ;

  if lStCurReg = '' then
    lStCurReg := PContexte.RegimeDefaut ;

  lStTitre := TraduireMemoire('Choix du régime (Actuel ' + RechDom('TTREGIMETVA', lStCurReg, False ) + ')' ) ;

  lStNewReg := Choisir( lStTitre, 'CHOIXCOD', 'CC_LIBELLE', 'CC_CODE', 'CC_TYPE="RTV"', '' ) ;

  if ( lStNewReg = '' ) or ( lStNewReg = lStCurReg ) then Exit ;

//  PGIAsk('Les TVA générées ou calculées peuvent devenir incorrectes. Confirmez-vous l''opération ?')
//  if HPiece.Execute(42,Caption,'')<>mrYes then Exit ;

//  GeneRegTVA:=NewReg ;
//  RegimeForce:=True ;
  Piece.PutValueAll( 'E_REGIMETVA', lStNewReg ) ;

end;

procedure TOF_CPSAISIEPIECE.ModifsSerieClick(Sender: TObject);
{$IFNDEF GCGC}
var lRComp         : R_COMP ;
    lAction        : TActionFiche ;
    i              : integer ;
    lBoToutModifie : boolean ;
{$ENDIF GCGC}
begin
{$IFNDEF GCGC}

  if TSP.GuideActif then Exit ;
  if Piece.ModeSaisie <> msPiece then Exit ;
  if Action=taConsult then Exit ;

  lRComp.StLibre    := 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' ;
  lRComp.StComporte := 'XXXXXXXXXX' ; // Pas de ModeLC
  lRComp.Conso      := True ;
  lRComp.DateC      := Piece.GetEnteteDt('E_DATECOMPTABLE') ;
  lRComp.Attributs  := False ;
  lRComp.MemoComp   := nil ;
  lRComp.Origine    := -1 ;
  lRComp.TOBCompl   := nil ;
  lRComp.AvecCalcul  := False ;

  lAction := taModif ;

  if not OuvreComplements   ( -1, lRComp, lAction ) then Exit ;

  if PgiAskCancel( TraduireMemoire('Confirmez-vous la modification de toutes les lignes de la pièce ?') ) <> mrYes then Exit ;

  lBoToutModifie := True ;
  // parcours des lignes
  for i := 1 to Piece.count do
    begin
    // On ne traite pas les eche > 1, maj par TPieceCompta sur eche=1
    if Piece.GetInteger(i,'E_NUMECHE')>1 then
      Continue ;

    // on ne traite que les lgnes modifiables
    if Piece.EstLigneModifiable( i )
      then begin
           Piece.AffecteEnteteScenario( i ) ;
           // E_CONSO non géré dans les maj d'entete
           if Piece.Entete.GetString('E_CONSO')<>'' then
             Piece.PutValue(i, 'E_CONSO', Piece.Entete.GetString('E_CONSO') );
           end
      else lBoToutModifie := False ;

    end ;

  if not lBoToutModifie then
    PgiInfo( TraduireMemoire('ATTENTION : certaines écritures pointées ou lettrées n''ont pas été modifiées') ) ;

{$ENDIF GCGC}

end;

procedure TOF_CPSAISIEPIECE.ZoomDeviseClick(Sender: TObject);
var lStCode : string ;
begin

  if (Piece.modeSaisie = msBOR) and not Piece.isOut( TSP.Row )
    then lStCode := Piece.GetValue( TSP.Row, 'E_DEVISE' )
    else lStCode := Piece.GetEnteteS( 'E_DEVISE' ) ;

  FicheDevise( lStCode, taConsult, False ) ;
  Piece.Info_LoadDevise( lStCode, True) ;

end;

procedure TOF_CPSAISIEPIECE.ZoomEtabClick(Sender: TObject);
Var Arguments : string;
begin
//  FicheEtablissement_AGL(taConsult) ;
if _Blocage(['nrCloture','nrBatch'],True,'nrAucun') then Exit ;
Arguments:=ActionToString(taConsult);
AGLLanceFiche('YY','YYETABLISS','',Piece.GetEnteteS('E_ETABLISSEMENT'),Arguments) ;
end;

procedure TOF_CPSAISIEPIECE.ZoomJournalClick(Sender: TObject);
{$IFNDEF GCGC}
Var lAction    : TActionFiche ;
    lStCodeJal : String ;
{$ENDIF}
begin
{$IFNDEF GCGC}
  lStCodeJal := Piece.GetEnteteS('E_JOURNAL') ;
  if ( lStCodeJal = '' ) then
    if GetControlText('E_JOURNAL') <> ''
      then lStCodeJal := GetControlText('E_JOURNAL')
      else Exit ;

  if Action=taConsult
    then lAction := taConsult
    else if ExJaiLeDroitConcept( TConcept(ccJalModif), False )
          then lAction := taModif  // if ModeLC then a:=taConsult
          else lAction := taConsult ;

  FicheJournal( Nil, '', lStCodeJal, lAction, 0 ) ;

  InfoEcr.LoadJournal( lStCodeJal, true ) ;
  MajInfosLigne( TSP.Row ) ;

{$ENDIF}
end;

procedure TOF_CPSAISIEPIECE.ZoomPiecesClick(Sender: TObject);
begin
{$IFNDEF GCGC}
  VisuDernieresPieces( FHistorique ) ;
{$ENDIF}
end;

procedure TOF_CPSAISIEPIECE.ZoomScenarioClick(Sender: TObject);
{$IFNDEF GCGC}
Var lTobScenario : Tob ;
    lStArg : string ;
    lStLequel : string ;
{$ENDIF}
begin
{$IFNDEF GCGC}
  lTobScenario := Piece.GetScenario( TSP.Row ) ;
  lStLequel    := '' ;
  lStArg       := ActionToString(taModif);
  if lTobScenario<>nil then
    begin
    // clé : SC_USERGRP, SC_JOURNAL, SC_NATUREPIECE, SC_QUALIFPIECE, SC_ETABLISSEMENT
    lStLequel := lTObSCenario.GetString('SC_USERGRP') + ';' +
                 lTObSCenario.GetString('SC_JOURNAL') + ';' +
                 lTObSCenario.GetString('SC_NATUREPIECE') + ';' +
                 lTObSCenario.GetString('SC_QUALIFPIECE') + ';' +
                 lTObSCenario.GetString('SC_ETABLISSEMENT') ;
    end ;
    AGLLanceFiche('CP','CPSCENARIO',lStLequel,'',lStArg);
  // rechargement du scenario en cours.
  Piece.Contexte.VideScenario ;
//  Piece.GetScenario( TSP.Row ) ;
{$ENDIF}
end;

procedure TOF_CPSAISIEPIECE.ZoomTauxClick(Sender: TObject);
var lDev   : RDevise ;
    lNumL  : integer ;
    lDateC : TDateTime ;
begin
  if Action = taConsult then Exit ;
  if Not GetControlEnabled ('BZOOMTAUX') then Exit ;

  if Piece.ModeSaisie = msBor
    then lNumL := TSP.Row
    else lNumL := 0 ;

  lDev := Piece.GetRDevise( lNumL ) ;
  if lDev.Code = V_PGI.DevisePivot then Exit ;
  if lDev.Code = '' then Exit ;

  lDateC := Piece.GetDateTime( TSP.Row, 'E_DATECOMPTABLE') ;

  if SaisieNewTaux2000( lDev, lDateC  ) then
    begin
//    Volatile:=True ;
    Piece.SetTauxVolatil  ( TSP.Row, lDev );
//    Piece.ChangeTauxPiece ( TSP.Row ) ;
    end ;
end;

procedure TOF_CPSAISIEPIECE.AccesBoutons;
var lBoLigneOk      : Boolean ;
    lBoTvaLoc       : Boolean ;
    lTob            : TOB ;
    lMontantOk      : Boolean ;
//    lreq            : string;
    lBoSolde        : Boolean ;
    lBoOkRib        : boolean ;
    lBoAction       : boolean ;
    lTobScenario    : Tob ;
    lBoGuide        : boolean ;
begin

  lBoAction    := (Action <> taConsult) and (not TSP.GuideActif) ;
  lBoLigneOk   := Piece.EstRemplit( TSP.row ) and (not TSP.GuideActif) ;
  lBoTvaLoc    := ((Simul='N') or (Simul='S') or (Simul='R')) and ( PContexte.OuiTvaEnc ) ;
  lTob         := Piece.GetTob( TSP.row ) ;
  lMontantOk   := lBoLigneOk and ( Piece.GetMontantDev( TSP.row ) <> 0 ) ;
  lBoSolde     := Piece.EstPieceSoldee ;

  lTobScenario := Piece.GetScenario( TSP.Row ) ;
  lBoOkRib     := (lTobScenario <> nil) and
                  ( ( lTobScenario.GetString('SC_RIB') = 'MAN' ) or
                    ( lTobScenario.GetString('SC_RIB') = 'PRI' )
                  ) or
                  ( (lTobScenario = nil) and Piece.Contexte.AttribRIBAuto );

  // Options générales
  BValider.Enabled     := lBoAction and Piece.ModifEnCours and lBoSolde ;
  BSolde.Enabled       := lBoAction and (Piece.Count >= 2 ) and (not lBoSolde) and lBoLigneOk ;
  BEche.Enabled        := lBoLigneOk and (lTob.GetValue('E_ECHE')='X') and lMontantOk;
  BVentil.Enabled      := lBoLigneOk and (lTob.GetValue('E_ANA')='X') and lMontantOk;
  BGenereTVA.Enabled   := lBoAction and ( Piece.ModeSaisie = msPiece ) and ( InfoEcr.LoadJournal( Piece.GetEnteteS('E_JOURNAL') ) )
                                    and ( ( InfoEcr.Journal.GetValue('J_NATUREJAL') = 'ACH' ) OR ( InfoEcr.Journal.GetValue('J_NATUREJAL') = 'VTE' ) ) ;
  if BGenereTVA.Enabled and ( lTobScenario <> nil ) then
    BGenereTVA.Enabled := ( lTobScenario.GetString('SC_CONTROLETVA') = 'GTD' ) or ( lTobScenario.GetString('SC_CONTROLETVA') = 'GTG' ) ;

  BControlTVA.Enabled  := lBoAction and lBoLigneOk and (not ANouveau) and ( Piece.ModeSaisie = msPiece ) and ( Piece.EstHT( TSP.Row ) ) ;
  BModifTVA.Enabled    := lBoAction and lBoLigneOk and lBoTvaLoc  and ( Piece.ModeSaisie = msPiece ) and ( Piece.EstHT( TSP.Row ) or Piece.EstTiers( TSP.Row ) ) ;
  BComplement.Enabled  := lBoLigneOk and (not ANouveau) ;
  BProrata.Enabled     := lBoAction and (Piece.Count >= 2 ) and (Not ANouveau)  ;
  BRechercher.Enabled  := not TSP.GuideActif  ;

  // popup Zooms
  SetControlEnabled('BMENUZOOM',      True ) ;
  SetControlEnabled('BZOOMJOURNAL',   True ) ;
  SetControlEnabled('BZOOMDEVISE',    True ) ;
  SetControlEnabled('BZOOMETAB',      True ) ;
  SetControlEnabled('BZOOMSCENARIO',  True ) ;
  SetControlEnabled('BZOOMPIECES',    Assigned( FHistorique ) and (FHistorique.Count > 0) ) ;
  SetControlEnabled('BZOOMTAUX',      lBoAction and {( Piece.Count <= 1 ) and (not Piece.EstRemplit(1)) and}
                                          (not EstMonnaieIN( Piece.GetString( TSP.Row, 'E_DEVISE' ) ) ) ) ;
  SetControlEnabled('BZOOMREGLT',     Assigned( FTobReglt ) and (FTobReglt.Detail.Count > 0) ) ;
  SetControlEnabled('BZOOMGENE',      lBoLigneOk ) ;
  SetControlEnabled('BZOOMTIERS',     lBoLigneOk and (lTob.GetString('E_AUXILIAIRE')<>'') ) ;

  // popup Actions lignes
  SetControlEnabled('BLIGNESINSERT',  lBoAction and lBoLigneOk ) ;
  SetControlEnabled('BLIGNESSUPPR',   lBoAction and (Piece.Count >= 2 ) and Piece.EstLigneModifiable( TSP.Row ) ) ;
  SetControlEnabled('BLIGNESLIBELLE', lBoAction and lBoLigneOk ) ;
  SetControlEnabled('BLIGNESRAZ',     lBoAction and (Piece.Count >= 2) ) ;
  SetControlEnabled('BMENULIGNES',    GetControlEnabled('BLIGNESINSERT') or GetControlEnabled('BLIGNESSUPPR') or GetControlEnabled('BLIGNESRAZ') ) ;

  // popup Guide
  lBoGuide := lBoAction and (Piece.Getsolde=0) and
              ( ( ( Screen.ActiveControl = FListe ) and
                  (TSP.Row = Piece.Detail.Count) and
                  TSP.TSGuide.ExisteGuide( TSP.Row )
                 ) ) ;
{                 //Activation depuis l'entête désactivé pour le moment
                  or
                ( (Piece.ModeSaisie=msPiece) and
                  (Piece.Detail.count=1) and
                  (not Piece.EstRemplit(1)) and
                  TSP.TSGuide.ExisteGuide( 1 )
                )
              ) ; }
  SetControlEnabled('BGUIDEOUVRIR',   lBoGuide ) ; // A Dev   // ( Action = taCreat ) and not ANouveau ;
  SetControlEnabled('BGUIDECREER',    False ) ; // A Dev   // ( Action = taCreat ) and not ANouveau ;
  SetControlVisible('BGUIDECREER',    False ) ; // A Dev   // ( Action = taCreat ) and not ANouveau ;
  SetControlEnabled('BMENUGUIDE',     GetControlEnabled('BGUIDEOUVRIR') or GetControlEnabled('BGUIDECREER') ) ;

  // popup modifs lignes
  SetControlEnabled('BMODIFSREGIME',  lBoAction and ( Piece.ModeSaisie = msPiece ) ) ;
  SetControlEnabled('BMODIFSSERIE',   lBoAction and ( Piece.ModeSaisie = msPiece ) and (Piece.Count >= 2 ) ) ;
  SetControlEnabled('BMODIFSRIB',     lBoAction and lBoLigneOk and lBoOkRib ) ;
  SetControlEnabled('BMENUMODIFS',    GetControlEnabled('BMODIFSREGIME') or GetControlEnabled('BMODIFSSERIE') or GetControlEnabled('BMODIFSRIB') ) ;

  {$IFDEF SCANGED}
  SetControlEnabled('BSCAN', {( VH^.OkModGed ) and} lBoLigneOk and ( Piece.RechGuidId( TSP.Row ) <> '' ) ) ;
  SetControlEnabled('BSCANPDF', lBoLigneOk ) ;
  {$ENDIF}

  SetControlEnabled('BDEVISE',        ( Piece.ModeSaisie = msBor ) and TSP.EstDeviseEditable( TSP.Row )  ) ;

  // Bouton Tva de l'entête
  BMenuTva.Visible          := ((Simul='N') or (Simul='S') or (Simul='R')) and ( PContexte.OuiTvaEnc ) and Piece.GereTvaMode( TSP.Row ) and (Piece.ModeSaisie = msPiece);
  BMenuTva.Enabled          := lBoAction ;

  // gestion des immos
{$IFDEF AMORTISSEMENT}
  BZoomImmo.Enabled         := ((VH^.OkModImmo) or (V_PGI.VersionDemo)) and (Piece.GetString( TSP.Row, 'E_IMMO') <> '' ) ;
{$ELSE}
  BZoomImmo.Enabled         := FALSE ;
{$ENDIF}

  // Boutons de navigations
  majBoutonNav ;

// b fb 02/05/2006
    if Engagement and assigned(PieceEngage) then begin
{      PieceEngage:=TPieceEngagement.CreerPieceEng;
      lReq :=' E_JOURNAL="' + E_JOURNAL.value + '"' +
             ' AND E_EXERCICE="' + QuelExo(E_DATECOMPTABLE.Text) + '"' +
             ' AND E_NUMEROPIECE=' + E_NUMEROPIECE.Text;

      PieceEngage.LoadFromSQL(lReq);
}      if Piece.Action = tamodif then begin
        if (PieceEngage.GetValue(TSP.row,'CEN_STATUTENG')='P') then begin
          AccesListe('P');
          SetControlEnabled( 'BValider', false);
          SetControlEnabled( 'BSolde', false);
          SetControlEnabled( 'BEche', false);
          SetControlEnabled( 'BVentil', false);
          SetControlEnabled( 'BGenereTVA', false);
          SetControlEnabled( 'BControleTVA', false);
          SetControlEnabled( 'BModifTVA', false);
          SetControlEnabled( 'BComplement', true);
          SetControlEnabled( 'BRechercher', false);
          SetControlEnabled( 'BProrata', false);
          SetControlEnabled( 'BMenuTva', false);
          SetControlEnabled('BMENUZOOM', false);
          SetControlEnabled('BZOOMJOURNAL', false );
          SetControlEnabled('BZOOMDEVISE', false );
          SetControlEnabled('BZOOMETAB', false );
          SetControlEnabled('BZOOMSCENARIO', false );
          SetControlEnabled('BZOOMPIECES', false );
          SetControlEnabled('BZOOMTAUX', false );
          SetControlEnabled('BZOOMREGLT', false );
          SetControlEnabled('BLIGNESINSERT', false );
          SetControlEnabled('BLIGNESSUPPR', false );
          SetControlEnabled('BLIGNESLIBELLE', false );
          SetControlEnabled('BLIGNESRAZ', false );
          SetControlEnabled('BMENULIGNES', false );
          SetControlEnabled('BGUIDEOUVRIR', false );
          SetControlEnabled('BGUIDECREER', false );
          SetControlEnabled('BMENUGUIDE', false );
          SetControlEnabled('BMODIFSREGIME', false );
          SetControlEnabled('BMODIFSSERIE', false );
          SetControlEnabled('BMODIFSRIB', false );
          SetControlEnabled('BMENUMODIFS', false );
          SetControlEnabled('BVALIDER', true );
         end
        else if (PieceEngage.GetValue(TSP.row,'CEN_STATUTENG')='E') then begin
          SetControlEnabled('BVALIDER', true );
          end;
          //PieceEngage.Free;
        end;
     end;
//e fb 02/05/2006

end;


function TOF_CPSAISIEPIECE.InfoEcr: TInfoEcriture;
begin
  result := nil ;
  if Piece <> nil then
    result := Piece.Info ;
end;

function TOF_CPSAISIEPIECE.Piece: TPieceCompta;
begin
  result := nil ;
  if assigned( FPieceCpt ) then
    result := FPieceCpt ;
end;

procedure TOF_CPSAISIEPIECE.ModifsRIBClick(Sender: TObject);
var lBoEstAuxi : Boolean ;
    lTobEcr    : TOB ;
begin
  if Not GetControlEnabled('BMODIFSRIB') then Exit ;
  if Piece.IsOut( TSP.Row ) then Exit ;
  lTobEcr := Piece.GetTOB( TSP.Row  ) ;

  lBoEstAuxi := lTobEcr.GetValue('E_AUXILIAIRE') <> '' ;
  ModifRIBOBM( lTobEcr, False, False, '', lBoEstAuxi ) ;

  TSP.AfficheLignes( TSP.Row ) ;
  MajInfosLigne( TSP.Row ) ;

end;

procedure TOF_CPSAISIEPIECE.BProrataClick(Sender: TObject);
begin
  if Action = taConsult then Exit ;
  if TSP.GuideActif then Exit ;

  TSP.ProratisePiece ;
  // Maj accès fonctions
  AccesBoutons ;
  // Maj infos pied
  refreshPCPied ;
end;

procedure TOF_CPSAISIEPIECE.GereEche( vRow : Integer ) ;
begin
{$IFDEF SBO}
//TSP.AddEvenement('ECRAN >> GereEche       Row : ' + IntToStr( vRow ) ) ;
{$ENDIF}

//  if not GetControlEnabled('BECHE') then Exit ;
  if not Piece.EstRemplit( vRow ) then Exit ;
  if not (Piece.GetMontantDev( vRow ) <> 0) then Exit ;
  if Piece.GetInteger(vRow, 'E_NUMECHE') > 1 then Exit ;

  // Cas de non ouverture de la fenêtre uniquement si E_MODEPAIE deja renseigné
  if ( Piece.GetString( vRow, 'E_MODEPAIE') <> '' ) then
    begin

    // Compte lettrable divers
    if Piece.EstDivLett( vRow ) then exit

    // Deja saisie (mode déporte <> mode multi)     // or ( (Piece.ModeEche<>meDeporte) and ( Piece.GetString('OUVREECHE')<>'X' ) )
    else if ( ( Piece.GetMultiEche( vRow ) <> nil ) and ( Piece.GetMultiEche( vRow ).Action <> taCreat ) )
       then Exit

    // Guide avec arrêt demandé
    else if ( TSP.GuideActif ) then
         begin
         if not TSP.TSGuide.AOuvrirEche( vRow ) then Exit ;
         end

    // OU scénario chargé sans ouverture des echéances
    else if ( (Piece.GetScenario(vRow) <> nil) and ( Piece.GetScenario(vRow).GetString('SC_OUVREECHE')<>'X') )
         then Exit

    // OU pas de scénario chargé ET ouverture des echéances non demandé dans les paramètres sociétés
    else if (Piece.GetScenario(vRow) = nil) and ( not Piece.Contexte.CPEcheOuvrir )
         then Exit

    else begin end ;
    
    end ;

  OuvreEche( vRow, True ) ;

{
  if AOuvrirEche(Lig,Cpte,MR,OuvreAuto,RempliAuto,t) then
   BEGIN
   if ((OkScenario) and (FromLigne) and (t=MultiEche)) then
      if Scenario.GetMvt('SC_OUVREECHE')<>'X' then RempliAuto:=True ;
   OuvreEche(Lig,Cpte,MR,RempliAuto,t) ;
   END ;
}

end;

procedure TOF_CPSAISIEPIECE.OuvreEche( vRow : Integer; vBoDeplace : boolean ) ;
Var lTobEcr     : TOB ;
    lMultiEche  : TMultiEche ;
    lBoMulti    : Boolean ;
    lAction     : TActionFiche ;
    lStArg      : String ;
    lBoVisu     : Boolean ;
    lTobTmp     : Tob ;
begin

  if not Piece.EstRemplit( vRow ) then Exit ;
  lTobEcr := Piece.GetTob( vRow ) ;
  if lTobEcr.GetValue( 'E_ECHE' ) <> 'X' then Exit ;
  if not (Piece.GetMontantDev( vRow ) <> 0) then Exit ;
  if Piece.EstDivLett( vRow ) then Exit ;

  // Mono ou Multi ?
  lBoMulti := Piece.EstOkMultiEche( vRow ) ;

  // Type d'action
  lBoVisu := ( Action = taConsult ) or ( not Piece.EstLigneModifiable( vRow ) ) ;
  if lBoVisu
    then lAction  := taConsult
    else lAction  := Piece.Action ;
  lStArg := ActionToString( lAction ) ;


  // La tob de l'écheance
  lMultiEche := Piece.GetMultiEche( vRow ) ;
  if lMultiEche = nil then Exit ;

  //=========================
  //===== MULTI ECHEANCE ====
  //=========================
  if lBoMulti then
    begin
    TheTob  := lMultiEche ;
    // Saisie
    CPLanceFiche_EcheMulti ( lStArg ) ;
    if ( lMultiEche.ModifEnCours ) and (not lBoVisu) then
      Piece.ModifEnCours := True ;
    end
  //========================
  //===== MONO ECHEANCE ====
  //========================
  else
    begin
    // La tob de l'écheance
    lTobTmp := Piece.GetCopyTob( vRow ) ;
    if lTobTmp = nil then Exit ;
    TheTob := lTobTmp ;
    // Saisie
    if CPLanceFiche_EcheMono( lStArg ) then
      begin
      // Mise à jour de la tob principale si multi déporté
      if Piece.ModeEche = meDeporte then
        lMultiEche.SynchroniseEcr
      else
        begin
        Piece.PutValue( vRow, 'E_MODEPAIE',     lTobTmp.GetString('E_MODEPAIE')       ) ;
        Piece.PutValue( vRow, 'E_DATEECHEANCE', lTobTmp.GetDateTime('E_DATEECHEANCE') ) ;
        Piece.PutValue( vRow, 'E_NUMTRAITECHQ', lTobTmp.GetString('E_NUMTRAITECHQ')   ) ;
        Piece.PutValue( vRow, 'E_DATEVALEUR',   lTobTmp.GetDateTime('E_DATEVALEUR')   ) ;
        end ;
      end ;
    FreeAndNil( lTobTmp ) ;
    end ;

  if lMultiEche.Action = taCreat then
    lMultiEche.Action := taModif ;

  // Refresh lignes
  TSP.AfficheLignes ;

  if vBoDeplace then
    if lMultiEche.NbEche > 1 then
      TSP.ProchaineLigne( vRow ) ;

  // Maj accès fonctions
  AccesBoutons ;


end;


procedure TOF_CPSAISIEPIECE.ClickAbandon( vBofromGrille : Boolean ) ;
var CanCancel : boolean ;
begin

  FModalResult := mrCancel ;

  // Retour dans l'entête si dans la ilste
  CanCancel := True ;
  if FBoMajDB then
    begin
    if TSP.GuideActif then
      // on quitte le guide
      begin
      TSP.TSGuide.AbandonnerGuide ;
      CanCancel := False ;
      end
    else
      begin
      TSP.CanClose( CanCancel ) ;
      if CanCancel and vBofromGrille and (FAction = TaCreat )then
        begin
        FStatutSaisie := ssEditEntete ;
//        TSP.Enabled := False ;
        ChargePiece( True ) ;
        TSP.LaGrille.Invalidate ;
        end ;
      end ;
    end ;

  if CanCancel and ( (GetAction <> TaCreat ) or not ( vBofromGrille and FBoMajDB )) then
    EcranClose ;

end;

procedure TOF_CPSAISIEPIECE.TSP_DblClick(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  if Action = taConsult then
    begin
    ZoomCompte( TSP.GetColIndex('E_GENERAL') ) ;
    end
  else
    begin
    if (TSP.GetColName = 'E_GENERAL') or (TSP.GetColName = 'E_AUXILIAIRE') then
      if not Piece.EstRemplit( TSP.Row )
        then TSP.RechCompte.ElipsisClick( Sender , False )
        else ZoomCompte( TSP.col ) ;
    end ;

end;

procedure TOF_CPSAISIEPIECE.ZoomGeneral;
var lAction : TActionFiche ;
    lStCpt  : String ;
begin

  if Piece.IsOut( TSP.Row ) then Exit ;
  lStCpt := Piece.GetValue( TSP.Row, 'E_GENERAL' ) ;
  if (lStCpt = '') or not ( InfoEcr.LoadCompte( lStCpt ) )  then Exit ;

  if ( Action <> taConsult ) and ExJaiLeDroitConcept( TConcept(ccGenModif), False )
    then lAction := taModif
    else lAction := taConsult ;

  FicheGene( nil, '', lStCpt, lAction, 0) ;
  InfoEcr.LoadCompte(lStCpt,true) ;

  MajInfosLigne( TSP.Row ) ;

end;

procedure TOF_CPSAISIEPIECE.ZoomTiers;
var lAction : TActionFiche ;
    lStCpt  : String ;
begin

  if Piece.IsOut( TSP.Row ) then Exit ;
  lStCpt := Piece.GetValue( TSP.Row, 'E_AUXILIAIRE' ) ;
  if (lStCpt = '') or not ( InfoEcr.LoadAux( lStCpt ) )  then Exit ;

  if ( Action <> taConsult ) and ExJaiLeDroitConcept( TConcept(ccAuxModif), False )
    then lAction := taModif
    else lAction := taConsult ;

  FicheTiers( nil, '', lStCpt, lAction, 1) ;

  InfoEcr.LoadAux(lStCpt,true) ;

  MajInfosLigne( TSP.Row ) ;

end;

procedure TOF_CPSAISIEPIECE.ZoomCompte(vCol: Integer);
var lStColName : String ;
begin
  lStColName := TSP.GetColName( vCol ) ;
  if lStColName = 'E_GENERAL' then
    begin
    ZoomGeneral ;
    end
  else if lStColName = 'E_AUXILIAIRE' then
    begin
    ZoomTiers ;
    end ;
end;



procedure TOF_CPSAISIEPIECE.InitPopupTva;
var i : Integer ;
begin
  for i := 0 to PopUpTva.Items.Count - 1 do
    begin
    // Tva sur les débits
    if PopUpTva.Items[i].Name = 'TVADEBIT'
      then PopUpTva.Items[i].OnClick := TvaDebitClick
    // Tva sur les encaissements
    else if PopUpTva.Items[i].Name = 'TVAENC'
      then PopUpTva.Items[i].OnClick := TvaEncClick
    // Tva mixte
    else if PopUpTva.Items[i].Name = 'TVAMIXTE'
      then PopUpTva.Items[i].OnClick := TvaMixteClick ;
    end ;

end;

procedure TOF_CPSAISIEPIECE.TvaDebitClick(Sender: TObject);
begin
  if not BMenuTva.Enabled then Exit ;
  BMenuTva.Caption := 'D' ;
  Piece.setTvaMode( TSP.Row, tvaDebit, True ) ;
  TSP.AfficheLignes ;
end;

procedure TOF_CPSAISIEPIECE.TvaEncClick(Sender: TObject);
begin
  if not BMenuTva.Enabled then Exit ;
  BMenuTva.Caption := 'E' ;
  Piece.setTvaMode( TSP.Row, tvaEncais, True ) ;
  TSP.AfficheLignes ;
end;

procedure TOF_CPSAISIEPIECE.TvaMixteClick(Sender: TObject);
begin
  if not BMenuTva.Enabled then Exit ;
  BMenuTva.Caption := 'M' ;
  Piece.setTvaMode( TSP.Row, tvaMixte, True ) ;
  TSP.AfficheLignes ;
end;

{b FP 14/11/2005 FQ15575}
function TOF_CPSAISIEPIECE.TrouveDoublonLigne( vRow : integer ; vStChamp : string ; var vTobRech : Tob ) : Boolean ;
var //lInCol     : integer ;
    lStSQL     : String ;
    lStVal     : String ;
    lTobEcr    : TOB ;
    lQZoom     : TQuery ;
    lStChpsEcr : string ;
begin

  Result   := False;

  if not Piece.EstRemplit( vRow ) then Exit ;

  lTobEcr  := Piece.GetTob( vRow ) ;
  if Length( lTobEcr.GetString('E_AUXILIAIRE') ) <> GetInfoCpta( fbAux ).Lg then Exit ;

  // valeur ?
  lStVal   := lTobEcr.GetString( vStChamp ) ;

  if ( lStVal = '' ) then Exit ;

  if Piece.ModeSaisie = msBor then
    begin
    vTobRech := nil ;
    // recherche dans les autres groupes du bordereau
    if Piece.RechercheChampDoublon( vRow, vStChamp, vTobRech ) then
      begin
      result := (vTobRech<>nil) ;
      vTobRech := nil ;
(*      if PgiAsk('Il existe des doublons dans le bordereau. Confirmez-vous la validation de la pièce ?') <> mrYes then
        begin
        lInCol := TSP.GetColIndex( vStChamp ) ;
        if lInCol < 1 then
          lInCol := TSP.GetColIndex( 'E_GENERAL' ) ;
        TSP.SetPos( lInCol , lTobRech.GetIndex + 1 ) ;
        Result:=False ;
        end ;
 *)
      Exit ;
      end ;
    end ;

  // recherche dans les ecritures
  vTobRech   := nil ;
  lStChpsEcr := 'E_JOURNAL, E_ETABLISSEMENT, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NATUREPIECE, '
              + 'E_QUALIFPIECE, E_DEVISE, E_TAUXDEV, E_MODESAISIE, E_DATETAUXDEV, E_VALIDE, E_NUMLIGNE, '
              + 'E_NUMECHE, E_GENERAL, E_PERIODE ' ;

  // En mode modif il faut exlure les lignes du bordereau courant du test...
  if Piece.Action = taModif then
    begin
    lStSQL := 'SELECT ' + lStChpsEcr + ' FROM ECRITURE WHERE E_GENERAL="'    + lTobEcr.GetString('E_GENERAL') + '"'
                                      + ' AND E_AUXILIAIRE="'                + lTobEcr.GetString('E_AUXILIAIRE') + '"'
                                      + ' AND E_QUALIFPIECE="N"'
                                      + ' AND E_NATUREPIECE IN ("AC","AF","FC","FF")'
                                      + ' AND ' + vStChamp + '="' + CheckdblQuote( lStVal ) + '"' ;

    lQZoom := OpenSelect( lStSQL, Dossier ) ;
    if Not lQZoom.EOF then
      begin
      while Not lQZoom.EOF do
        begin
        if not ( ( lQZoom.FindField('E_JOURNAL').AsString = Piece.GetEnteteS('E_JOURNAL') ) and
                 ( lQZoom.FindField('E_PERIODE').AsInteger = Piece.GetEnteteI('E_PERIODE') ) and
                 ( lQZoom.FindField('E_NUMEROPIECE').AsInteger = Piece.GetEnteteI('E_NUMEROPIECE') )
                ) then
           begin
           vTobRech := TOB.Create( 'ECRITURE' , nil , -1 ) ;
           vTobRech.SelectDB( '', lQZoom ) ;
           Break ;
           end ;
        lQZoom.Next ;
        end ;
      end ;
    Ferme(lQZoom) ;

    end
  else
    begin
    lStSQL := 'SELECT ##TOP 1## ' + lStChpsEcr +
                      ' FROM ECRITURE WHERE E_GENERAL="'    + lTobEcr.GetString('E_GENERAL') + '"'
                                    + ' AND E_AUXILIAIRE="'          + lTobEcr.GetString('E_AUXILIAIRE') + '"'
                                    + ' AND E_QUALIFPIECE="N"'
                                    + ' AND E_NATUREPIECE IN ("AC","AF","FC","FF")'
                                    + ' AND ' + vStChamp + '="' + CheckdblQuote( lStVal ) + '"' ;

    lQZoom := OpenSelect( lStSQL, Dossier ) ;
    if Not lQZoom.EOF then
      begin
      vTobRech := TOB.Create( 'ECRITURE' , nil , -1 ) ;
      vTobRech.SelectDB( '', lQZoom ) ;
      end ;
    Ferme(lQZoom) ;
    end ;

  Result:=(vTobRech<>nil) ;
{
  if Assigned( lTobRech ) then
    begin
    if PgiAsk('Il existe des doublons dans la base. Confirmez-vous la validation de la pièce ?') <> mrYes then
      begin
      Result:=False ;
      if PgiAsk('Voulez-vous visualiser le premier doublon détecté ?') = mrYes then
        begin
        M := TOBToIdent( lTobRech, False ) ;
        LanceSaisieParam( Nil, taConsult, M ) ;
        end ;
      end ;
    lTobRech.free ;
    end ;
}
end;


procedure TOF_CPSAISIEPIECE.TraiteErreurTva(vErrId: Integer);
var lMess : TStringList ;
begin
  lMess  := TStringList.Create ;
  CRemplirMessageTVA( lMess ) ;
  HShowMessage( lMess[ vErrId ] , '' , '' ) ;
end;

procedure TOF_CPSAISIEPIECE.RefreshModeTva;
begin
  case Piece.getTvaMode( TSP.Row ) of
    tvaMixte  : BMenuTva.Caption := 'M' ;
    tvaDebit  : BMenuTva.Caption := 'D' ;
    tvaEncais : BMenuTva.Caption := 'E' ;
  end ;
end;

procedure TOF_CPSAISIEPIECE.ParamJournal;
begin

  if Simul='N' then
    begin
    E_JOURNAL.DataType := 'CPJOURNAL' ;
    E_JOURNAL.plus     := ' J_FERME="-" AND J_NATUREJAL not IN ("ANA", "ODA", "ANO", "CLO") ' ;
    ZappePieceFolio ;
    end
  else
    begin
    // A NOUVEAU
    if ((Action=taCreat) and (Simul='N') and (ANouveau))
      then SetControlProperty('E_JOURNAL','DataType','ttJalANouveau')
    // IFRS
    else if Simul='I'
      then SetControlProperty('E_JOURNAL','DataType','CPJOURNALIFRS')
    // saisie par engagement
    else if Engagement then
      begin
      SetControlProperty('E_JOURNAL','DataType','CPJOURNAL');
      E_JOURNAL.plus     := ' J_FERME="-" AND J_MODESAISIE="-" AND J_NATUREJAL IN ("ACH", "VTE") ' ;
      end
    // Autre saisie
    else SetControlProperty('E_JOURNAL','DataType','ttJalSansEcart') ;
    end ;

end;

procedure TOF_CPSAISIEPIECE.OnAfterFormShow;
var
  {b FP 02/05/2006}
  PointeurBefore:    PBeforeChargePiece;
  BeforeChargePiece: TBeforeChargePiece;

  PointeurAfter:     PAfterChargePiece;
  AfterChargePiece:  TAfterChargePiece;
  LocalAction:       TActionFiche;
  {e FP 02/05/2006}
begin
{$IFDEF SBO}
//TSP.AddEvenement('ECRAN >> OnAfterFormShow' ) ;
{$ENDIF}
  {b FP 02/05/2006}

//  Ecran.SetFocus ;

  LocalAction := Action;
  if Assigned(FPieceCpt) and Piece.FieldExists('BeforeChargePiece') then
    begin
    PointeurBefore := PBeforeChargePiece(Integer(Piece.GetValue('BeforeChargePiece')));
    if Assigned(PointeurBefore) then
      begin
      BeforeChargePiece := PointeurBefore^;
      BeforeChargePiece(Self, Ecran, Piece, LocalAction);
      end;
    end;
  if LocalAction <> Action then
    Action := LocalAction ;
  {e FP 02/05/2006}

  // MAJ grille suivant paramètrage de la liste
//  TSP.RechercheMasque(True);

  if Action = taCreat then
    begin
    TSP.EditEnabled( True ) ;             // option de navigation dans la grille
//    TSP.Enabled := False ; // Acces grille
    end ;

  // Chargement des données
{$IFDEF SAISIEDOC}
  if Assigned(TSP.TSDoc) then
    begin
//    TSP.TSDoc.Piece := Piece ;
    TSP.ChargeDocument(False) ; // necessaire ??
    end ;
{$ENDIF SAISIEDOC}
  ChargePiece( Action = taCreat ) ;
  if Action=taCreat then
    begin
    SetFocusControl('E_JOURNAL') ;
    TSP.AfficheLignes( 1, True ) ;
    end ;

//  TSP.LaGrille.Invalidate ;

  {b FP 02/05/2006}
  if Assigned(FPieceCpt) and Piece.FieldExists('AfterChargePiece') then
    begin
    PointeurAfter := PAfterChargePiece(Integer(Piece.GetValue('AfterChargePiece')));
    if Assigned(PointeurAfter) then
      begin
      AfterChargePiece := PointeurAfter^;
      AfterChargePiece(Self, Ecran, Piece);
      end;
    end;
  {e FP 02/05/2006}
end;

procedure TOF_CPSAISIEPIECE.SetInitFocus( vBoReinit : Boolean ) ;
//var lCol : integer ;
begin
{$IFDEF SBO}
//TSP.AddEvenement('ECRAN >> SetInitFocus     Reinit : ' + BoolToStr_( vBoReinit ) ) ;
{$ENDIF}

  // En mode mémoire, on accède directement à la grille
  if not FBoMajDB then
    TSP.SetFocus

  // Gestion du Focus sur abandon des modifs grilles --> retour sur le journal
  else if vBoReinit and (FModalResult=mrCancel) then
    SetFocusControl('E_JOURNAL')

  // Gestion du Focus sur mode pièce
  else if (Piece.ModeSaisie = msPiece) then
    begin
    if vBoReinit then // cas du bouclage en saisie
      begin
      if ( Action = taCreat ) and ( E_DATECOMPTABLE.CanFocus ) and
         ( (Piece.Contexte.CPDateObli) or ( ( Piece.GetScenario <> nil ) and
                                            ( Piece.GetScenario.GetValue('SC_DATEOBLIGEE') = 'X' ) ) )
          then SetFocusControl('E_DATECOMPTABLE') // retour sur la date comptable si obligatoire
          else TSP.SetFocus ;
      end
    else // 1ère entrée sur mode piece
      begin
      if (Action = taCreat)
          then SetFocusControl('E_JOURNAL')       // Focus sur journal à l'entrée dans la fiche
          else TSP.SetFocus ;
      end ;
    end

  // Gestion du Focus sur mode BOR / Libre
  else
    if (Piece.Detail.count = 0) or vBoReinit
      then SetFocusControl('FOLIO')
      else TSP.SetFocus ;

  if not TSP.Enabled then
    TSP.Enabled := True ;

end;

function TOF_CPSAISIEPIECE.GereComplements(vRow: integer ; lBoEntete : Boolean ): Boolean ;
{$IFNDEF GCGC}
Var lStComp   : String ;
    lStLibre  : String ;
    lRComp    : R_COMP ;
    lInNumRad : integer ;
    lTobEcr   : TOB ;
    lStCpt    : String ;
    lMemo     : HTStringList ;
{$ENDIF}
begin
{$IFDEF SBO}
//TSP.AddEvenement('ECRAN >> GereComplements       Row : ' + IntToStr( vRow ) ) ;
{$ENDIF}

 result := False ;

{$IFNDEF GCGC}
//  if not Piece.ModifEnCours then Exit ;

  // ******************************
  // ** 2 cas possible d'appel : **
  // ******************************

  // 1. Infos complémentaires dans l'entête (arrivée dans la grille)

  if lBoEntete then
    begin
    if (Piece.Action<>taCreat) then Exit ; // mdoe entête en mode création uniquement

    // Récup paramétrage
    Piece.GetCompLigne( lStComp, lStLibre, 0 ) ;
    // Si rien à saisir , on quitte
    if ((Pos('X',lStComp)<=0) and (Pos('X',lStLibre)<=0)) then Exit ;
    lRComp.Origine:=0 ;

    lRComp.DateC      := Piece.GetEnteteDT('E_DATECOMPTABLE') ;

    end

  // 2. Infos complémentaires dans les lignes (à charque sortie de ligne)
  else
    begin
    // Recup Tob ligne Ecriture
    lTobEcr := Piece.GetTob( vRow ) ;
    if lTobEcr = nil then Exit ;

    // COMPS Pour eviter Double ouverture ??
    if lTobEcr.GetNumChamp('COMPS') < 0
      then lTobEcr.AddChampSupValeur('COMPS','-')
      else if lTobEcr.GetValue('COMPS') = 'X' then Exit ;

    // Récup paramétrage
    lInNumRad := Piece.GetCompLigne( lStComp, lStLibre, vRow ) ;

    // On a récupérer un comportemant issu du scénario OU la gestion des qte PCL est activée
    if (lInNumRad > 0) or Piece.GereQtePCL( vRow ) then
      begin
      // Si rien à saisir , on quitte
      if ((Pos('X',lStComp)<=0) and (Pos('X',lStLibre)<=0)) then Exit ;
      lRComp.Origine    := lInNumRad ;

      lTobEcr.PutValue('COMPS' , 'X' ) ;     // COMPS Pour eviter Double ouverture ??

      lRComp.DateC      := lTobEcr.GetDateTime('E_DATECOMPTABLE') ;
      lRComp.TOBCompl   := CGetTOBCompl( lTobEcr ) ;
      lStCpt := Piece.GetString( vRow, 'E_GENERAL' ) ;
      if (lStCpt <> '') and ( InfoEcr.LoadCompte( lStCpt ) )  then
        begin
        lRComp.CutOffPer   := InfoEcr.Compte.GetValue('G_CUTOFFPERIODE') ;
        lRComp.CutOffEchue := InfoEcr.Compte.GetValue('G_CUTOFFECHUE') ;
        end ;
      end
    else
    // Si on sort
      Exit ;

    end ;

  // On doit ouvrir la fiche
  result := True ;

  // *****************************
  // ** Paramètres de lancement **
  // *****************************
  lRComp.StComporte := lStComp ;
  lRComp.StLibre    := lStLibre ;
  lRComp.Conso      := False ;
  lRComp.Attributs  := True ;
  // comportement...
  if Piece.GetScenario( vRow ) <> nil then
    begin
    lMemo := Piece.GetMemoScenario( vRow ) ;
    if Assigned( lMemo ) then
      lRComp.MemoComp      := lMemo ;
    end
  else lRComp.MemoComp   := nil ;

  // ** appel fiche **
  OuvreComplements   ( vRow, lRComp , taCreat ) ;

{$ENDIF}
end;

function TOF_CPSAISIEPIECE.OuvreComplements   ( vRow   : integer ; vRComp : R_COMP ; vAction : TActionfiche ) : Boolean ;
{$IFNDEF GCGC}
var lTobEcr    : TOB ;
    lAction    : TActionFiche ;
    lBoModifBN : Boolean ;
    lOldQte1   : Double ;
    lOldQte2   : Double ;
    lNewQte1   : Double ;
    lNewQte2   : Double ;
    lOBM       : TOBM ;
    lBoVisu    : Boolean ;
{$ENDIF GCGC}
begin
  result := False ;

{$IFNDEF GCGC}

  lOldQte1      := 0 ;
  lOldQte2      := 0 ;
  // appel pour modif en série
  if vRow < 0 then
    begin
    if Action=taConsult then Exit ;
    if Piece.ModeSaisie <> msPiece then Exit ;
    lTobEcr := Piece.Entete ;
    end
  // appel en entete
  else if vRow = 0 then
    begin
    if Action<>taCreat then Exit ;
    if (Piece.Count > 1) or Piece.EstRemplit(1) then Exit ;
    lTobEcr := Piece.Entete ;
    end
  // appel pour une ligne
  else
    begin
    lTobEcr := Piece.GetTob( vRow ) ;
    if lTobEcr = nil then Exit ;

    // réouverture du Cut-off
    if ( lTobEcr.GetString('CUTOFF') = '0' ) then
     begin
      lTobEcr.PutValue('CUTOFF','1') ;
      GereCutOff( vRow ) ;
      exit ;
     end ;

    lOldQte1      := lTobEcr.GetValue('E_QTE1') ;
    lOldQte2      := lTobEcr.GetValue('E_QTE2') ;

    end ;

  // Reprise du TOBM
  lOBM := TOBM.Create( EcrGen, '', True ) ;
  lOBM.Dupliquer( lTobEcr, true, true ) ;

  try
    // à la ligne
    if vRow > 0 then
      begin
      lBoVisu := (Action = taConsult) or ( not Piece.EstLigneModifiable( vRow ) ) ;

      if lBoVisu
        then lAction := taConsult
        else lAction := vAction ;
      end
    // à l'entête
    else
      begin
      lAction := vAction ;
      lBoVisu := False ;
      end ;

    result := SaisieComplement( lOBM, EcrGen, lAction, lBoModifBN, vRComp, False, True ) ;
    if not result then Exit ;

    if not lBoVisu then
      begin
      // Reprise des données du TOBM dans la TOB de la pièce ( ou dans l'entête selon pointeur )
      lTobEcr.Dupliquer( lOBM, True, True ) ;

      if vRow = 0 then
        // modif zone d'entête : impact sur la 1ère ligne deja en mémoire
        Piece.AffecteEnteteScenario( 1 )
      else
        begin
        if lTobEcr.GetString('E_ANA')='X' then
          begin
          // Nouvelles quantités
          lNewQte1 := lTobEcr.GetDouble('E_QTE1') ;
          lNewQte2 := lTobEcr.GetDouble('E_QTE2') ;
          // Recalcul au proprata dans les lignes de ventilations
          if ( (lNewQte1<>lOldQte1) and (lOldQte1<>0) ) then
            ProrateQteAnalTOB( lTobEcr, lOldQte1, lNewQte1, '1') ;
          if ( (lNewQte2<>lOldQte2) and (lOldQte2<>0) ) then
            ProrateQteAnalTOB( lTobEcr, lOldQte2, lNewQte2, '2') ;
          // Reaffectation des champs commun E_xxx >> Y_xxx
          Piece.AffecteVentilPourComplement( vRow ) ;
          end ;

        // Test modifications ?
        Piece.ModifEnCours := True ;

        end ;

      end ;


  finally

    if Assigned( lOBM ) then
      FreeAndNil( lOBM ) ;

    TSP.AfficheLignes ;

    // Maj accès fonctions
    AccesBoutons ;

  end;

{$ENDIF GCGC}

end;

procedure TOF_CPSAISIEPIECE.TSP_GridEnter(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;

  if FStatutSaisie = ssEditGrille then Exit ;

  // Appel des info coplémentaire en entête de pièce
  if Piece.ModeSaisie = msPiece then
    begin
    if Piece.Action<>taCreat then Exit ;
    GereComplements ( 0, True ) ;
//    if GereComplements ( 0, True ) then
//      TSP.SetFocus ;
    end
  else if Piece.ModeSaisie in [msBor,msLibre] then
    begin
    ChargePiece ;
    end ;

  FStatutSaisie := ssEditGrille ;

end;

procedure TOF_CPSAISIEPIECE.GereAnal(vRow: Integer);
var lInAxe  : Integer ;
    lTobEcr : Tob ;
begin
{$IFDEF SBO}
//TSP.AddEvenement('ECRAN >> GereAnal       Row : ' + IntToStr( vRow ) ) ;
{$ENDIF}

  // bouton
//  if not GetControlEnabled('BVENTIL')         then Exit ;
  // Compte renseigné
  if not Piece.EstRemplit( vrow )             then Exit ;
  // Compte ventilable
  if not Piece.EstVentilable( vrow )          then Exit ;
  // Montant renseigné
  if not (Piece.GetMontantDev( vrow ) <> 0)   then Exit ;

  // Si scénario charger et ouverture des echéances non prévu alors on ouvre pas !
  lInAxe  := 0 ;
  lTobEcr := Piece.GetTob( vRow ) ;
  if (TSP.GuideActif and TSP.TSGuide.AOuvrirAnal( vRow ) ) or
     CAOuvrirVentil( lTobEcr, Piece.GetScenario(vRow), InfoEcr, lInAxe, True ) then
    begin
    // indicateur d'ouverture
{    if lTobEcr.GetNumChamp('OUVREANA') > 0
      then if lTobEcr.GetString('OUVREANA') = 'X'
             then Exit
             else lTobEcr.PutValue('OUVREANA', 'X')
      else lTobEcr.AddChampSupValeur('OUVREANA', 'X') ;
}    // ouverture
    OuvreAnal ( vRow, lInAxe ) ;
    end ;

end;

procedure TOF_CPSAISIEPIECE.AvertirPbTaux ( vRow : Integer ) ;
Var lInRep    : integer ;
    lDtDateC  : TDateTime ;
begin

  if Action=taConsult then Exit ;

  if Piece.ModeSaisie<>msPiece then Exit ;

  // test validation du taux de change
  if Piece.IsValidTaux( vRow )  then Exit ;

  // Récup date comptable + date taux dev
{  if Piece.IsOut(vRow)
    then lDtDateC  := Piece.GetEnteteDt('E_DATECOMPTABLE')
    else lDtDateC  := Piece.GetDateTime( vRow, 'E_DATECOMPTABLE') ;
}  lDtDateC  := Piece.GetEnteteDt('E_DATECOMPTABLE') ;

  // Si monnaie In et date > mise en place de l'euro
  if EstMonnaieIN( Piece.Devise.Code ) and ( lDtDateC >= V_PGI.DateDebutEuro ) then
    begin
    lInRep := PgiAsk('ATTENTION. La parité est incorrecte. Voulez-vous la saisir ?', Ecran.Caption ) ;
    if lInRep = mrYes then
      begin
      FicheDevise( Piece.GetEnteteS('E_DEVISE'), taModif, False ) ;
      Piece.MajDeviseTaux( vRow, True ) ;
      end ;
    end
  else
    begin
    lInRep := PgiAsk( 'Voulez-vous saisir ce taux dans la table de chancellerie ?', Ecran.Caption ) ;
    if lInRep = mrYes then
      begin
      FicheChancel( Piece.GetEnteteS('E_DEVISE'), True, lDtDateC, taCreat, ( lDtDateC >= V_PGI.DateDebutEuro ) ) ;
      Piece.MajDeviseTaux( vRow, True ) ;
      end ;
    end ;

end;

procedure TOF_CPSAISIEPIECE.LettrageEnSaisie;
Var X         : RMVT ;
    Q         : TQuery ;
    Nb        : integer ;
    lTobScena : Tob ;
begin
{$IFNDEF GCGC}

  if Simul<>'N' then Exit ;

  // En création uniquement en mode pièce
  if (Piece.ModeSaisie= msPiece) and ( Piece.Action<>taCreat ) then Exit ;

  // Uniquement si prévu dans le scénario (du 1er groupe en mode BOR )
  lTobScena := Piece.GetScenario(1) ;
  if ( lTobScena = nil ) or ( lTobScena.GetString('SC_LETTRAGESAISIE') <> 'X' )  then Exit ;

  X := Piece.GetRMVT ;

  Q:=OpenSQL('SELECT COUNT(*) AS NB FROM ECRITURE WHERE ' + Piece.GetWhereSQL +
                            ' AND E_ECHE="X" AND E_NUMECHE>0 AND E_ETATLETTRAGE="AL"', True ) ;
  if Not Q.EOF
    then Nb := Q.FindField('NB').AsInteger
    else Nb := 0 ;
  Ferme(Q) ;

  if Nb > 0 then
    if PgiAsk( 'Voulez-vous lettrer vos échéances ?', Ecran.Caption ) = mrYes then
      LettrerEnSaisie(X,Nb) ;

{$ENDIF}
end;

procedure TOF_CPSAISIEPIECE.EditionSaisie;
var lStWhere : String ;
    i        : integer ;
    RR       : RMVT ;
    lStEtat  : String ;
begin

  if Piece.ModeSaisie <> msPiece then Exit ;
  if Piece.GetScenario = nil then Exit ;

  lStEtat := Piece.GetScenario.GetString('SC_DOCUMENT') ;
  if lStEtat = '' then Exit ;

  if PgiAsk('Voulez-vous éditer le document interne des écritures créées ?', Ecran.Caption) <> mrYes then Exit ;

  lStWhere := '' ;

  for i := 0 to ( FHistorique.Count - 1 ) do
    begin
    RR := P_MV( FHistorique[i] ).R ;
    lStWhere := lStWhere + '(E_JOURNAL="' + RR.Jal + '" AND E_NUMEROPIECE='  + IntToStr(RR.Num)
                                                   +  ' AND E_QUALIFPIECE="' + RR.Simul
                                                   + '" AND E_EXERCICE="'    + RR.Exo
                                                   + '")' ;
    if i < ( FHistorique.Count - 1 ) then
      lStWhere := lStWhere + ' OR ' ;
    end ;

  if FHistorique.Count > 1 then
    lStWhere := '(' + lStWhere + ')' ;

  {$IFNDEF IMP}
    {$IFNDEF GCGC}
    LanceEtat( 'E', 'SAI', lStEtat, True, False, False, Nil, lStWhere, '', False ) ;
    {$ENDIF}
  {$ENDIF}

end;

procedure TOF_CPSAISIEPIECE.GereCutoff(vRow: integer);
Var lTobEcr      : TOB ;
    lRComp       : R_COMP ;
    lTOBCompl    : TOB ;
    lStCpt       : String ;
begin
{$IFDEF SBO}
//TSP.AddEvenement('ECRAN >> GereCutoff       Row : ' + IntToStr( vRow ) ) ;
{$ENDIF}

  if not Piece.ModifEnCours then Exit ;

  lTobEcr := Piece.GetTob( vRow ) ;
  if lTobEcr = Nil then Exit ;

  lStCpt := Piece.GetValue( vRow, 'E_GENERAL' ) ;
  if (lStCpt = '') or not ( Piece.Info.LoadCompte( lStCpt ) )  then Exit ;

  if Piece.Compte_GetValue('G_CUTOFF') <> 'X' then
    begin
    if (lTobEcr.GetNumChamp('CUTOFF') >= 0) and ( lTobEcr.GetString('CUTOFF') = 'X' ) then
      begin
      CFreeTOBCompl( lTobEcr ) ;
      lTobEcr.PutValue( 'CUTOFF', '-' ) ;
      end ;
    exit ;
    end ;

  // Réouverture
  if lTobEcr.GetString('CUTOFF') = '1' then
    lTobEcr.PutValue( 'CUTOFF', '-' ) ;

  // Condition de sortie
  if ( lTobEcr.GetString('CUTOFF') <> '-' ) then Exit ;

  lRComp.StComporte  := '' ;
  lRComp.StLibre     := '' ;
  lRComp.Conso       := False ;
  lRComp.DateC       := Piece.GetEnteteDt('E_DATECOMPTABLE') ;
  lRComp.Attributs   := (Piece.Action=taCreat) ;
  lRComp.MemoComp    := nil ; // MemoComp := FindField('SC_ATTRCOMP')).AsString
  lRComp.Origine     := 0 ;
  lRComp.StLibre     := '---CUTXXXXXXXXXXXXXXXXXXXXXXXX' ;
  lRComp.StComporte  := '--XXXXXXXX' ;
  lRComp.AvecCalcul  := True ;
  lRComp.CutOffPer   := InfoEcr.Compte.GetValue('G_CUTOFFPERIODE') ;
  lRComp.CutOffEchue := InfoEcr.Compte.GetValue('G_CUTOFFECHUE') ;

  lTOBCompl := Piece.GetTobCompl( vRow ) ;
  lTOBCompl.PutValue('EC_DATECOMPTABLE', lTobEcr.GetValue('E_DATECOMPTABLE')) ;

  CCalculDateCutOff( lTOBCompl, Piece.Compte_GetValue('G_CUTOFFPERIODE') , Piece.Compte_GetValue('G_CUTOFFECHUE') ) ;
  lRComp.TOBCompl := lTOBCompl ;

  lTobEcr.PutValue( 'CUTOFF', 'X' ) ;
  if not OuvreComplements( vRow, lRComp, taCreat ) then
    begin
    lTobEcr.PutValue( 'CUTOFF', '0' ) ;
    CFreeTOBCompl( lTobEcr ) ;
    end ;

  if TSP.LaGrille.CanFocus then
    TSP.LaGrille.SetFocus ;

end;

function TOF_CPSAISIEPIECE.PContexte: TPieceContexte;
begin
  result := nil ;
  if Piece <> nil then
    result := Piece.Contexte ;
end;


function TOF_CPSAISIEPIECE.ANouveau: Boolean;
begin
  result := False ;
  if Assigned( FPieceCpt ) then
    result := FPieceCpt.GetEnteteS('E_ECRANOUVEAU') = 'X' ;
end;

function TOF_CPSAISIEPIECE.Dossier: string ;
begin
  result := V_PGI.SchemaName ;
  if Assigned( FPieceCpt ) then
    result := FPieceCpt.Dossier ;
end;

function TOF_CPSAISIEPIECE.Simul: string ;
begin
  result := 'N' ;
  if Assigned( FPieceCpt ) then
    result := FPieceCpt.GetEnteteS('E_QUALIFPIECE') ;
end;

function TOF_CPSAISIEPIECE.GetTitre: string;
begin
  result := 'Saisie' ; //paramétrable' ;
  if Simul = 'N' then
    begin
    result := result + ' des écritures courantes' ;
    if ( GetControlText('E_JOURNAL')<>'' ) then
      Case Piece.ModeSaisie of
        msBOR   : result := result + ' en mode bordereau' ;
        msLibre : result := result + ' en mode libre' ;
        else      result := result + ' en mode pièce' ;
        end ;
   end
  else if Simul = 'S'
    then result := result + ' des écritures de simulation'
  else if Simul = 'I'
    then result := result + ' des écritures IFRS'         // FQ 17694
  else if Simul = 'U'
    then result := result + ' des écritures de situation'
  else if (Simul = 'P') and not Engagement //fb 02/05/2006
    then result := result + ' des écritures de prévision'
  else if Simul = 'R'
    then result := result + ' des écritures de révision'
{b fb 02/05/2006}
  else if Engagement then
    result := result + ' des écritures d''engagement';
{e fb 02/05/2006}
end;

procedure TOF_CPSAISIEPIECE.BVisuRegltClick(Sender: TObject);
var lTobVisu : Tob ;
    i        : integer ;
    lTobLP   : Tob ;
    lTobEcr  : Tob ;
    lStSoc   : string ;
    lStEtab  : string ;
    lStJal   : string ;
    lInNumP  : integer ;
    lDtDateC : TDateTime ;
begin

  if TSP.GuideActif then Exit ;
  if not assigned( FTobReglt ) then Exit ;
  if FTobReglt.Detail.Count = 0 then Exit ;

  lTobVisu := TOB.Create('$ECR', nil, -1) ;

  // parcours des pièces principales
  for i := 0 to FTobReglt.detail.count - 1 do
    begin
    lTobLP   := FTobReglt.Detail[i] ;
    lStSoc   := lTobLP.GetString('CLP_SOCIETE') ;
    lStEtab  := lTobLP.GetString('CLP_ETABLISSEMENT') ;
    lStJal   := lTobLP.GetString('CLP_JOURNAL') ;
    lInNumP  := lTobLP.GetInteger('CLP_NUMEROPIECE') ;
    lDtDateC := lTobLP.GetDateTime('CLP_DATECOMPTABLE') ;
    // Ajout pièces principales
    if lTobVisu.FindFirst( ['E_SOCIETE', 'E_ETABLISSEMENT', 'E_JOURNAL', 'E_NUMEROPIECE', 'E_DATECOMPTABLE'],
                           [ lStSoc, lStEtab, lStJal, lInNumP, lDtDateC ],
                           True ) = nil
      then begin
           lTobEcr := Tob.Create('ECRITURE', lTobVisu, -1 ) ;

           // Champs minimum : 'E_SOCIETE;E_ETABLISSEMENT;E_JOURNAL;E_NUMEROPIECE;E_DATECOMPTABLE;E_REFINTERNE' ;
           lTobEcr.putValue( 'E_SOCIETE',          lStSoc     ) ;
           lTobEcr.putValue( 'E_ETABLISSEMENT',    lStEtab    ) ;
           lTobEcr.putValue( 'E_JOURNAL',          lStJal     ) ;
           lTobEcr.putValue( 'E_NUMEROPIECE',      lInNumP    ) ;
           lTobEcr.putValue( 'E_DATECOMPTABLE',    lDtDateC   ) ;
           lTobEcr.putValue( 'E_EXERCICE',         lTobLP.GetString('CLP_EXERCICE')  ) ;
           lTobEcr.putValue( 'E_QUALIFPIECE',      'N'   ) ;
           lTobEcr.putValue( 'E_REFINTERNE',       lTobLP.GetString('CLP_REFINTERNE')       ) ;
           lTobEcr.AddChampSupValeur( 'SYSDOSSIER',         lTobLP.GetString('CLP_DOSSIER')          ) ;
{
           if lTobLP.GetString('CLP_DOSSIER') <> lTobLP.GetString('CLP_DOSSIERDEST')
             then lTobEcr.AddChampSupValeur( 'STYLE',  'BOLD' )     // pièce multi-soc principale
             else lTobEcr.AddChampSupValeur( 'STYLE',  'ITALIC' ) ; // pièce multi-étab principale
}
           end ;

    end ;

  // parcours des pièces destination
  for i := 0 to FTobReglt.detail.count - 1 do
    begin
    lTobLP := FTobReglt.Detail[i] ;
    lStSoc   := lTobLP.GetString('CLP_SOCIETEDEST') ;
    lStEtab  := lTobLP.GetString('CLP_ETABDEST') ;
    lStJal   := lTobLP.GetString('CLP_JOURNALDEST') ;
    lInNumP  := lTobLP.GetInteger('CLP_NUMPIECEDEST') ;
    lDtDateC := lTobLP.GetDateTime('CLP_DATEDEST') ;
    // Ajout pièces destination si pas elle même pièce principale
    if lTobVisu.FindFirst( ['E_SOCIETE', 'E_ETABLISSEMENT', 'E_JOURNAL', 'E_NUMEROPIECE', 'E_DATECOMPTABLE'],
                           [ lStSoc, lStEtab, lStJal, lInNumP, lDtDateC ],
                           True ) = nil
      then begin
           lTobEcr := Tob.Create('ECRITURE', lTobVisu, -1 ) ;

           // Champs minimum : 'E_SOCIETE;E_ETABLISSEMENT;E_JOURNAL;E_NUMEROPIECE;E_DATECOMPTABLE;E_REFINTERNE' ;
           lTobEcr.putValue( 'E_SOCIETE',          lStSoc     ) ;
           lTobEcr.putValue( 'E_ETABLISSEMENT',    lStEtab    ) ;
           lTobEcr.putValue( 'E_JOURNAL',          lStJal     ) ;
           lTobEcr.putValue( 'E_NUMEROPIECE',      lInNumP    ) ;
           lTobEcr.putValue( 'E_DATECOMPTABLE',    lDtDateC   ) ;
           lTobEcr.putValue( 'E_EXERCICE',         lTobLP.GetString('CLP_EXERCICEDEST')  ) ;
           lTobEcr.putValue( 'E_QUALIFPIECE',      'N'   ) ;
           lTobEcr.putValue( 'E_REFINTERNE',       lTobLP.GetString('CLP_REFINTERNE')           ) ;
           lTobEcr.AddChampSupValeur( 'SYSDOSSIER',         lTobLP.GetString('CLP_DOSSIERDEST')          ) ;
           end ;

    end ;

  // Tri par société / établissement / numéro pièce
  lTobVisu.Detail.Sort('E_SOCIETE;E_ETABLISSEMENT;E_NUMEROPIECE');

  TheTob := lTobVisu ;

  // Préparation des paramètres
  CPLanceFiche_VisuPieces( ActionToString( taConsult ) + ';' + 'TITRE=Consultation des règlements multi-sociétés / multi-établissements' ) ;

  TheTob  := nil ;


end;

{---------------------------------------------------------------------------------------}
function TOF_CPSAISIEPIECE.GetAction : TActionFiche;
{---------------------------------------------------------------------------------------}
begin
  if Piece.Action = taConsult
    then result := taConsult
    else Result := FAction;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPSAISIEPIECE.SetAction(Value : TActionFiche);
{---------------------------------------------------------------------------------------}
begin
  FAction := Value;
  TFVierge(Ecran).TypeAction := Value;
end;

function TOF_CPSAISIEPIECE.GereTiersPayeurs: boolean;
Var lInFirst : integer ;
begin
  Result := False ;

  if Piece.ModeSaisie <> msPiece then Exit ;

  if not Piece.TestParamTP( True ) then Exit ;

  // 1ère ligne traitable ?
  lInFirst := Piece.GetFirstIdxTP ;
  if lInFirst < 1 then Exit ; // aucune --> on sort

  // Recup info Aux
  piece.GetTob( lInFirst ) ;

  // Remboursement sur Avoir ?
  if ((Piece.GetEnteteS('E_NATUREPIECE')='AC') or (Piece.GetEnteteS('E_NATUREPIECE')='AF')) then
    if Piece.Aux_GetValue('T_AVOIRRBT')='X' then Exit ;

  // Avertissement ?
  if Piece.Aux_GetValue('T_DEBRAYEPAYEUR')='X' then
    if PgiAsk('Voulez-vous générer les écritures sur le tiers payeur ?', Ecran.Caption)<>mrYes then Exit ;

  Result := True ;
end;

procedure TOF_CPSAISIEPIECE.BScanClick(Sender: TObject);
begin
  if TSP.GuideActif then Exit ;
{$IFDEF SCANGED}
ShowGedViewer( Piece.RechGuidId( TSP.Row ), True ) ;
{$ENDIF}
end;

{$IFDEF SCANGED}
function TOF_CPSAISIEPIECE.GetInfoLigne: string;
begin
 result := 'Document scanné Date : '     + DateToStr( Piece.GetEnteteDt('E_DATECOMPTABLE') )
                        + ' Journal : '  + Piece.GetEnteteS('E_JOURNAL')
                        + ' N° Pièce : ' + Piece.GetEnteteS('E_NUMEROPIECE') ;
end;

procedure TOF_CPSAISIEPIECE.SetGuidID ( vGuidId : string ) ;
begin

 if ( Piece.RechGuidId(TSP.Row) <> '' ) and
    ( PGIAsk('Il existe une image ou un fichier associé , voulez-vous le supprimer ?') = mrNo ) then exit ;

 Piece.AjouteGuidId(TSP.Row, vGuidId ) ;
 Piece.ModifEnCours := True ;
 AccesBoutons ;

end;
{$ENDIF SCANGED}


{b fb 02/05/2006}
procedure TOF_CPSAISIEPIECE.AccesListe(StatutEng : string);
begin
  if StatutEng='P' then
    TSP.Enabled := False // Acces grille
  else if StatutEng='E' then
    TSP.Enabled := True ; // Acces grille
end;
{e fb 02/05/2006}

procedure TOF_CPSAISIEPIECE.FolioExit(Sender: TObject);
var lInNum : integer ;
begin

  if ( csDestroying in Ecran.ComponentState) then Exit ;

  // Test du numero
  if length(Folio.Text) > 10 then Folio.Text := '1' ;
  if ValeurI(Folio.Text) > 0
    // Valeur ok : affectation au TPieceCompta
    then begin
         lInNum := Trunc(ValeurI(Folio.Text)) ;
         Piece.PutEntete('E_NUMEROPIECE', lInNum ) ;
         end
    // Sinon : reinit du folio, on sort
    else begin
         InitFolio( True, True ) ;
         Exit ;
         end ;

  if Screen.ActiveControl=E_JOURNAL then Exit ;
  if Screen.ActiveControl=Periode then Exit ;

  // Chargement du folio
//  ChargePiece ;

end;

procedure TOF_CPSAISIEPIECE.PeriodeExit(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;

  majTypeExo ;

  if GetPeriode(StrToDate(Periode.Value)) <> GetPeriode(Piece.GetEnteteDt('E_DATECOMPTABLE')) then
    begin
    Piece.PutEntete('E_DATECOMPTABLE', FinDeMois(StrToDate(Periode.Value)) ) ;
    if Sender<>nil then
      begin
      // maj du numéro de Folio
      paramFolio ;
      // Set du numéro et Focus
      InitFolio ;
      end ;
    end ;

end;

procedure TOF_CPSAISIEPIECE.InitPeriode   ( vBoFocus : Boolean ) ;
var lDtDateC : TDateTime ;
begin

  // Init de la valeur
  lDtDateC := DebutDeMois( InfoEcr.Journal.GetValue('J_DATEDERNMVT') ) ;

  if (lDtDateC = iDate1900) or ( lDtDateC < StrToDate( Periode.Values[0] ) )
                            or ( lDtDateC > StrToDate( Periode.Values[Periode.Values.count-1] ) )
    then lDtDateC := StrToDate( Periode.Values[0] ) ;

  Periode.ItemIndex := Periode.Values.IndexOf( DateToStr( lDtDateC ) ) ;
  if Periode.ItemIndex < 0 then
    begin
    Periode.ItemIndex := Periode.Items.Count - 1 ;
    lDtDateC := StrToDate( Periode.Value ) ;
    end ;

  lDtDateC := FinDeMois( lDtDateC ) ;
  Piece.PutEntete('E_DATECOMPTABLE', lDtDateC ) ;

  // gestion Focus
  if vBoFocus then
    Folio.SetFocus ;

end;


procedure TOF_CPSAISIEPIECE.ZappePieceFolio( vStJal : String );
var lBoModePiece : boolean ;
begin

  if vStJal = ''
    then vStJal := Piece.GetEnteteS('E_JOURNAL') ;
  lBoModePiece := ( vStJal = '') or
                  ( Piece.Info_LoadJournal( vStJal ) and
                    not ( ( Piece.Info.Journal.GetValue('J_MODESAISIE') = 'BOR' ) or
                          ( Piece.Info.Journal.GetValue('J_MODESAISIE') = 'LIB' ) ) ) ;

  // zone spécif mode piece
  setControlVisible('E_DATECOMPTABLE',  lBoModePiece ) ;
  setControlVisible('TE_DATECOMPTABLE', lBoModePiece ) ;
  setControlVisible('E_NATUREPIECE',    lBoModePiece ) ;
  setControlVisible('TE_NATUREPIECE',   lBoModePiece ) ;
  setControlVisible('E_NUMEROPIECE',    lBoModePiece ) ;
  setControlVisible('TE_NUMEROPIECE',   lBoModePiece ) ;

  // zones spécif bordereau
  setControlVisible('PERIODE',  not lBoModePiece ) ;
  setControlVisible('TPERIODE', not lBoModePiece ) ;
  setControlVisible('FOLIO',    not lBoModePiece ) ;
  setControlVisible('TFOLIO',   not lBoModePiece ) ;
  setControlVisible('TYPEEXO',  not lBoModePiece ) ;

end;

procedure TOF_CPSAISIEPIECE.ParamFolio;
begin

  if ( E_JOURNAL.Value = '') or ( PERIODE.Value = '' ) then
    begin
    Folio.Items.Clear ;
    Folio.Values.Clear ;
    Exit ;
    end ;

  CRempliComboFolio ( Folio.Items, Folio.Values , E_JOURNAL.Value, Piece.GetEnteteS('E_EXERCICE') , strToDate(PERIODE.Value) ) ;

end;

procedure TOF_CPSAISIEPIECE.InitFolio ( vBoFocus : Boolean ; vBoForce : boolean ) ;
var lDtPer : TDatetime ;
    lDtJal : TDatetime ;
    lInNum : Integer ;
begin

  // Maj du Champ
  lDtPer := Piece.GetEnteteDt('E_DATECOMPTABLE') ; //StrTodate( Periode.value ) ;

//  if vBoForce or (GetPeriode(lDtPer) <> GetPeriode(StrToDate(Periode.Value))) then
    begin
//    Piece.PutEntete('E_DATECOMPTABLE', FinDeMois(StrToDate(Periode.Value)) ) ;
    // Dernier journal ?
    lDtJal := InfoEcr.Journal.GetValue('J_DATEDERNMVT') ;
    if GetPeriode( lDtJal ) = GetPeriode( lDtPer ) then
      begin
      lInNum := InfoEcr.Journal.GetValue('J_NUMDERNMVT') ;
      // spécif BOR paramètre "Incrémenté le bordereau"
      if (InfoEcr.Journal.GetValue('J_INCNUM')='X') or (lInNum=0) then
        Inc( lInNum ) ;
      end
    else lInNum := 1 ;

    Folio.Text := IntTostr( lInNum ) ; // Dernier Numero ou suivant
    Piece.PutEntete('E_NUMEROPIECE', lInNum) ;

    end ;

  // Focus
  if vBoFocus then
    Folio.SetFocus ;

end;

procedure TOF_CPSAISIEPIECE.ParamPeriode;
var lBoAllPer : boolean ;
    lStExo    : string ;
begin

  // Exo en cours
  lStExo := Piece.GetEnteteS('E_EXERCICE') ;
  if (ctxPCL in V_PGI.PGIContexte) and
     ( (ctxExercice.CPExoRef.Code=VH^.Encours.Code) or
       (ctxExercice.CPExoRef.Code=ctxExercice.Suivant.Code)  )
   then lStExo:=ctxExercice.CPExoRef.Code ;

  if lStExo='' then
    lStExo := ctxExercice.Entree.Code ;

  if Periode.Items.Count = 0 then
    begin
    lBoAllPer := Action=taConsult ;
    // 1er chargement en fonction du type d'action
    ListePeriode(lStExo, Periode.Items, Periode.Values, True, lBoAllPer) ;
    // si aucune période trouvé, rechargement avec passage en mode consultation
    if (not lBoAllPer) and (Periode.Items.Count = 0) then
      begin
      ListePeriode(lStExo, Periode.Items, Periode.Values, True, True) ;
      Action:=taConsult ;
      end ;
    end ;

end;

procedure TOF_CPSAISIEPIECE.InitNaturePiece ;
begin
  if E_Journal.Value = '' then
    E_NATUREPIECE.ItemIndex := -1
  else
    begin
    E_NATUREPIECE.Value   := Piece.GetNatureInitiale ;
    Piece.PutEntete('E_NATUREPIECE', E_NATUREPIECE.Value ) ;
    end ;
end;

procedure TOF_CPSAISIEPIECE.ParamNaturePiece( vStJal : string );
begin

//  if GetControlText('E_JOURNAL')='' then
  if vStJal = '' then
    vStJal := Piece.GetEnteteS('E_JOURNAL') ;
  if vStJal = '' then
    begin
    E_NATUREPIECE.DataType := '' ;
    E_NATUREPIECE.Value    := '' ;
    Exit ;
    end ;

  // tablette utilisée
  InfoEcr.LoadJournal( vStJal ) ;
  if Engagement then
    begin
    E_NATUREPIECE.DataType := 'ttNaturePiece' ;
    Case CaseNatJal( InfoEcr.Journal.GetValue('J_NATUREJAL') ) of
      tzJVente  : E_NATUREPIECE.plus := ' AND CO_CODE IN ("FC","AC")' ;
      tzJAchat  : E_NATUREPIECE.plus := ' AND CO_CODE IN ("FF","AF")' ;
      end ;
    end
  else
    begin
    Case CaseNatJal( InfoEcr.Journal.GetValue('J_NATUREJAL') ) of
      tzJVente  :      E_NATUREPIECE.DataType := 'ttNatPieceVente'  ;
      tzJAchat  :      E_NATUREPIECE.DataType := 'ttNatPieceAchat'  ;
      tzJBanque :      E_NATUREPIECE.DataType := 'ttNatPieceBanque' ;
      tzJEcartChange : E_NATUREPIECE.DataType := 'ttNatPieceEcartChange' ;
      tzJOD     :      E_NATUREPIECE.DataType := 'ttNaturePiece' ;
       else            E_NATUREPIECE.DataType := 'ttNaturePiece' ;
      end ;
    end ;

end;

procedure TOF_CPSAISIEPIECE.SetTitre;
begin
  Ecran.Caption := GetTitre ;
//  if Assigned( FPieceCpt ) and Piece.IsActiveAcc then
//    Ecran.Caption := Ecran.Caption + ' (accélérateur de saisie activé)' ;
  UpdateCaption( Ecran ) ;
end;

procedure TOF_CPSAISIEPIECE.InitNumeroPiece ;
begin
  if Piece.ModeSaisie<>msPiece then Exit ;

  // gestion des numéro
  Piece.AttribNumeroTemp ;
  SetControlText('E_NUMEROPIECE', Piece.GetEnteteS('E_NUMEROPIECE') ) ;
end;


procedure TOF_CPSAISIEPIECE.InitDateComptable;
begin
  // Init de la valeur
  SetControlText('E_DateComptable',  DateToStr( V_PGI.DateEntree ) ) ;
  Piece.PutEntete('E_DATECOMPTABLE', StrToDate( GetControlText('E_DATECOMPTABLE') ) ) ;
end;

procedure TOF_CPSAISIEPIECE.InitDevise ;
begin
  // Init de la valeur
  SetControlText( 'E_DEVISE', V_PGI.DevisePivot ) ;
  Piece.PutEntete('E_DEVISE', V_PGI.DevisePivot ) ;
end;


procedure TOF_CPSAISIEPIECE.InitEtablissement;
var lStEtab : string ;
begin
  lStEtab := Piece.Contexte.GetEtablisDefaut ;
  if Trim(lStEtab) = '' then
    lStEtab := E_ETABLISSEMENT.Values[0] ;
  SetControlText( 'E_ETABLISSEMENT', lStEtab ) ;
  Piece.PutEntete('E_ETABLISSEMENT', lStEtab ) ;
end;

procedure TOF_CPSAISIEPIECE.majTypeExo;
var lType : TTypeExo ;
    lNomC : string ;
begin
  if Piece.ModeSaisie = msPiece then
    begin
    lType := Piece.Contexte.GetTypeExo( StrToDate(GetControlText('E_DATECOMPTABLE') ) ) ;
    lNomC := 'TE_EXERCICE' ;
    end
  else
    begin
    if IsValidDate(Periode.Value)
      then lType := Piece.Contexte.GetTypeExo( StrToDate(Periode.Value) )
      else lType := Piece.Contexte.GetTypeExo( StrToDate(GetControlText('E_DATECOMPTABLE') ) ) ;
    lNomC := 'TYPEEXO' ;
    end ;

  Case lType of
    teEncours :    SetControlText( lNomC, '(N)') ;
    tePrecedent :  SetControlText( lNomC, '(N-1)') ;
    teSuivant :    SetControlText( lNomC, '(N+1)') ;
    else           SetControlText( lNomC, '') ;
    end ;
    
end;

procedure TOF_CPSAISIEPIECE.AfficherSolde( vStControl: String; vSolde: Double ; vBoPer : boolean ) ;
var lStSolde : string ;
    lInDec   : integer ;
begin

  if Piece.ModeSaisie = msPiece
    then lInDec := Piece.Devise.Decimale
    else lInDec := V_PGI.OkDecV ;

//  lSolde  := Arrondi( lTotalD- lTotalC, lInDec ) ;

  if vSolde < 0
    then lStSolde := StrFMontant( -1*vSolde, 15, lInDec, 'C', true )     // solde créditeur
    else lStSolde := StrFMontant(    vSolde, 15, lInDec, 'D', true )  ;  // solde débiteur

  if vBoPer then
    lStSolde := ' (M) ' + lStSolde ;

  SetControlText( vStControl, lStSolde ) ;

end;

procedure TOF_CPSAISIEPIECE.CalculSoldeAuxiPer( vStCpt : string ) ;
var lTob    : Tob ;
    lQSum   : TQuery ;
    lStSQL  : string ;
    lDtDate : TDateTime ;
begin

  if not Piece.Info_LoadAux( vStCpt ) then Exit ;

  lTob := Piece.Info.Aux.Item ;
  if lTob.GetNumChamp('_DEBITPERIODE') > 0 then exit ;

  lTob.AddChampSupValeur( '_DEBITPERIODE', 0) ;
  lTob.AddChampSupValeur( '_CREDITPERIODE', 0) ;

  lDtDate := Piece.GetEnteteDt( 'E_DATECOMPTABLE' ) ;

  lStSQL := ' SELECT SUM(E_DEBIT) AS DEB, SUM(E_CREDIT) AS CRE FROM ECRITURE' ;
  lStSQL := lStSQL  + ' WHERE E_JOURNAL="'        + Piece.GetEnteteS('E_JOURNAL')  + '"'
                      + ' AND E_EXERCICE="'       + Piece.GetEnteteS('E_EXERCICE') + '"'
                      + ' AND E_DATECOMPTABLE>="' + USDateTime( DebutDeMois(lDtDate) )+ '"'
                      + ' AND E_DATECOMPTABLE<="' + USDateTime( FinDeMois(lDtDate) )+ '"'
                      + ' AND ( E_QUALIFPIECE="N" OR E_QUALIFPIECE="L" ) '
                      + ' AND E_ETABLISSEMENT="'  + Piece.GetEnteteS('E_ETABLISSEMENT') + '" '
                      + ' AND E_AUXILIAIRE="'     + vStCpt + '"' ;

  if Piece.Action<>taCreat then
     lStSQL := lStSQL + ' AND E_NUMEROPIECE<>' + Piece.GetEnteteS('E_NUMEROPIECE') ;

  lQSum := OpenSQL( lStSQL, True ) ;
  if not lQSum.Eof then
    begin
    lTob.PutValue('_DEBITPERIODE',  lQSum.FindField('DEB').AsFloat );
    lTob.PutValue('_CREDITPERIODE', lQSum.FindField('CRE').AsFloat );
    end ;
  Ferme( lQSum ) ;

end;

procedure TOF_CPSAISIEPIECE.CalculSoldeGenePer( vStCpt : string ) ;
var lTob    : Tob ;
    lQSum   : TQuery ;
    lStSQL  : string ;
    lDtDate : TDateTime ;
begin

  if not Piece.Info_LoadCompte( vStCpt ) then Exit ;

  lTob := Piece.Info.Compte.Item ;
  if lTob.GetNumChamp('_DEBITPERIODE') > 0 then exit ;

  lTob.AddChampSupValeur( '_DEBITPERIODE', 0) ;
  lTob.AddChampSupValeur( '_CREDITPERIODE', 0) ;

  lDtDate := Piece.GetEnteteDt( 'E_DATECOMPTABLE' ) ;

  lStSQL := ' SELECT SUM(E_DEBIT) AS DEB, SUM(E_CREDIT) AS CRE FROM ECRITURE' ;
  lStSQL := lStSQL  + ' WHERE E_JOURNAL="'        + Piece.GetEnteteS('E_JOURNAL')  + '"'
                      + ' AND E_EXERCICE="'       + Piece.GetEnteteS('E_EXERCICE') + '"'
                      + ' AND E_DATECOMPTABLE>="' + USDateTime( DebutDeMois(lDtDate) )+ '"'
                      + ' AND E_DATECOMPTABLE<="' + USDateTime( FinDeMois(lDtDate) )+ '"'
                      + ' AND ( E_QUALIFPIECE="N" OR E_QUALIFPIECE="L" ) '
                      + ' AND E_ETABLISSEMENT="' + Piece.GetEnteteS('E_ETABLISSEMENT') + '" '
                      + ' AND E_GENERAL="'       + vStCpt + '"' ;

  if Piece.Action<>taCreat then
     lStSQL := lStSQL + ' AND E_NUMEROPIECE<>' + Piece.GetEnteteS('E_NUMEROPIECE') ;

  lQSum := OpenSQL( lStSQL, True ) ;
  if not lQSum.Eof then
    begin
    lTob.PutValue('_DEBITPERIODE',  lQSum.FindField('DEB').AsFloat );
    lTob.PutValue('_CREDITPERIODE', lQSum.FindField('CRE').AsFloat );
    end ;
  Ferme( lQSum ) ;

end;

procedure TOF_CPSAISIEPIECE.CalculSoldeJal;
var lTob    : Tob ;
    lQSum   : TQuery ;
    lStSQL  : string ;
    lDtDate : TDateTime ;
begin

  if not Piece.Info_LoadJournal( Piece.GetEnteteS('E_JOURNAL') ) then Exit ;
  lTob := Piece.Info.Journal.Item ;

  if lTob.GetNumChamp('_DEBITPERIODE') > 0 then exit ;

  lTob.AddChampSupValeur( '_DEBITPERIODE', 0) ;
  lTob.AddChampSupValeur( '_CREDITPERIODE', 0) ;

  lDtDate := Piece.GetEnteteDt( 'E_DATECOMPTABLE' ) ;

  lStSQL := ' SELECT SUM(E_DEBIT) AS DEB, SUM(E_CREDIT) AS CRE FROM ECRITURE' ;
  lStSQL := lStSQL  + ' WHERE E_JOURNAL="'        + Piece.GetEnteteS('E_JOURNAL')  + '"'
                      + ' AND E_EXERCICE="'       + Piece.GetEnteteS('E_EXERCICE') + '"'
                      + ' AND E_DATECOMPTABLE>="' + USDateTime( DebutDeMois(lDtDate) )+ '"'
                      + ' AND E_DATECOMPTABLE<="' + USDateTime( FinDeMois(lDtDate) )+ '"'
                      + ' AND ( E_QUALIFPIECE="N" OR E_QUALIFPIECE="L" ) '
                      + ' AND E_ETABLISSEMENT="' + Piece.GetEnteteS('E_ETABLISSEMENT') + '" ' ;

  if Piece.Action<>taCreat then
    lStSQL := lStSQL + ' AND E_NUMEROPIECE<>' + Piece.GetEnteteS('E_NUMEROPIECE') ;

  lQSum := OpenSQL( lStSQL, True ) ;
  if not lQSum.Eof then
    begin
    lTob.PutValue('_DEBITPERIODE',  lQSum.FindField('DEB').AsFloat );
    lTob.PutValue('_CREDITPERIODE', lQSum.FindField('CRE').AsFloat );
    end ;
  Ferme( lQSum ) ;

end;

procedure TOF_CPSAISIEPIECE.AfficherSoldeAuxi( vStCpt : string ) ;
var lCptC   : Double ;   // Total crédit du compte
    lCptD   : Double ;   // Total débit du compte
    lPieceC : Double ;   // Total crédit de la pièce pour ce compte
    lPieceD : Double ;   // Total débit de la pièce pour ce compte
    lTob    : Tob ;
begin
  if not Piece.Info_LoadAux( vStCpt )   then Exit ;

  // Infos de l'auxiliaire
  lTob := Piece.Info.Aux.Item ;
//  lTob.PutEcran( Ecran, PCPied ) ;

  // ---------------------------------
  // --- Calcul du solde classique ---
  // ---------------------------------
  CalculSoldeAuxi( vStCpt ) ;
  Piece.GetTotauxPourChamps( 'E_AUXILIAIRE', vStCpt, lPieceD, lPieceC ) ;
  lPieceD := lTob.GetDouble('_DEBITHISTO') + lPieceD ;
  lPieceC := lTob.GetDouble('_CREDITHISTO') + lPieceC ;

  // Affichage du solde
  AfficherSolde(  'SA_SOLDET',  lPieceD - lPieceC )  ;
  SetControlText( 'SA_DEVISET',  V_PGI.SymbolePivot ) ;

  // -------------------------
  // --- Solde par période ---
  // -------------------------
  if FBoCptSoldeP then
    begin
    CalculSoldeAuxiPer( vStCpt ) ;
    // Affichage du solde classique
    lCptD := Piece.Info.Aux.Item.GetDouble('_DEBITPERIODE') ;
    lCptC := Piece.Info.Aux.Item.GetDouble('_CREDITPERIODE') ;
    AfficherSolde( 'SA_SOLDETPER', (lCptD + lPieceD) - (lCptC + lPieceC), True )  ;
    end ;

end;

procedure TOF_CPSAISIEPIECE.AfficherSoldeGene( vStCpt : string ) ;
var lCptC   : Double ;   // Total crédit du compte
    lCptD   : Double ;   // Total débit du compte
    lPieceC : Double ;   // Total crédit de la pièce pour ce compte
    lPieceD : Double ;   // Total débit de la pièce pour ce compte
    lTob    : Tob ;
begin
  if not Piece.Info_LoadCompte( vStCpt )   then Exit ;

  // Infos du général
  lTob := Piece.Info.Compte.Item ;
//  lTob.PutEcran( Ecran, PCPied ) ;

  // ---------------------------------
  // --- Calcul du solde classique ---
  // ---------------------------------
  CalculSoldeGene( vStCpt ) ;
  Piece.GetTotauxPourChamps( 'E_GENERAL', vStCpt, lPieceD, lPieceC ) ;
  lPieceD := lTob.GetDouble('_DEBITHISTO') + lPieceD ;
  lPieceC := lTob.GetDouble('_CREDITHISTO') + lPieceC ;

  // Affichage du solde classique
  AfficherSolde( 'SA_SOLDEG',  lPieceD - lPieceC )  ;
  SetControlText('SA_DEVISEG',  V_PGI.SymbolePivot ) ;

  // -------------------------
  // --- Solde par période ---
  // -------------------------
  if FBoCptSoldeP then
    begin
    CalculSoldeGenePer( vStCpt ) ;
    // Affichage du solde classique
    lCptD := Piece.Info.Compte.Item.GetDouble('_DEBITPERIODE') ;
    lCptC := Piece.Info.Compte.Item.GetDouble('_CREDITPERIODE') ;
    AfficherSolde( 'SA_SOLDEGPER', (lCptD + lPieceD) - (lCptC + lPieceC), True )  ;
    end ;

end;

procedure TOF_CPSAISIEPIECE.AfficherSoldeJal;
var lCptC   : Double ;   // Total crédit du compte
    lCptD   : Double ;   // Total débit du compte
    lPieceC : Double ;   // Total crédit de la pièce pour ce compte
    lPieceD : Double ;   // Total débit de la pièce pour ce compte
    lInDec  : integer ;
begin

  // Info solde journal
  if FBoJalSoldeP then
    begin
    SetControlText('SA_JALSOLDE',  '');
    SetControlText('SA_JALCREDIT', '');
    SetControlText('SA_JALDEBIT',  '');
    end ;

  if not Piece.Info_LoadJournal( Piece.GetEnteteS('E_JOURNAL') )   then Exit ;

  // Infos du journal
  Piece.Info.Journal.Item.PutEcran( Ecran, PCPied ) ;

  // -------------------------
  // --- Solde par période ---
  // -------------------------
  if FBoJalSoldeP then
    begin
    // calculs solde hors piece en cours
    CalculSoldeJal ;
    lCptD   := Piece.Info.Journal.Item.GetDouble('_DEBITPERIODE') ;
    lCptC   := Piece.Info.Journal.Item.GetDouble('_CREDITPERIODE') ;
    // Calcul solde piece en cours
    if Piece.GetEnteteS('E_DEVISE') = V_PGI.DevisePivot then
      begin
      // Totaux deja calculé en devise pivot, on les récupère
      lPieceD := Valeur( GetControlText( 'SA_TOTALDEBIT' ) ) ;
      lPieceC := Valeur( GetControlText( 'SA_TOTALCREDIT' ) ) ;
      end
    else
      begin
      // Calcul des Totaux en devise pivot
      lPieceD  := Piece.GetTotalDebit( False ) ;
      lPieceC  := Piece.GetTotalCredit( False ) ;
      end ;
    // Affichage
    if Piece.ModeSaisie = msPiece
      then lInDec := Piece.Devise.Decimale
      else lInDec := V_PGI.OkDecV ;
    SetControlText('SA_JALDEBIT',  StrFMontant( lCptD + lPieceD,  15, lInDec, '', True ) ) ;
    SetControlText('SA_JALCREDIT', StrFMontant( lCptC + lPieceC,  15, lInDec, '', True ) ) ;
    AfficherSolde( 'SA_JALSOLDE', (lCptD + lPieceD) - (lCptC + lPieceC) )  ;
    end
  else
    begin
    // Infos totaux / solde de la pièce :
    if Piece.ModeSaisie = msPiece
      then SetControlText( 'TSA_JALTOTAL', 'Solde Pièce' )
      else SetControlText( 'TSA_JALTOTAL', 'Solde Folio' ) ;
    end ;

end;

procedure TOF_CPSAISIEPIECE.ReinitSoldes;
begin
  Piece.Info.Journal.Items.DelChampSup('_DEBITPERIODE', True);
  Piece.Info.Journal.Items.DelChampSup('_CREDITPERIODE', True);

  Piece.Info.Compte.Items.DelChampSup('_DEBITPERIODE', True);
  Piece.Info.Compte.Items.DelChampSup('_CREDITPERIODE', True);

  Piece.Info.Aux.Items.DelChampSup('_DEBITPERIODE', True);
  Piece.Info.Aux.Items.DelChampSup('_CREDITPERIODE', True);

  Piece.Info.Compte.Items.DelChampSup('_DEBITHISTO', True);
  Piece.Info.Compte.Items.DelChampSup('_CREDITHISTO', True);

  Piece.Info.Aux.Items.DelChampSup('_DEBITHISTO', True);
  Piece.Info.Aux.Items.DelChampSup('_CREDITHISTO', True);

end;


procedure TOF_CPSAISIEPIECE.BDeviseClick(Sender: TObject);
begin
//  if TSP.GuideActif then Exit ;
  TSP.SaisieDevise( TSP.Row ) ;
  refreshPCPied ;  
//bStopDevise:=TRUE ;
end;

function TOF_CPSAISIEPIECE.TraiteDoublon: boolean;
var lStCode  : string ;
    lStChamp : string ;
    i        : integer ;
    lTobRech : Tob ;
    M        : RMVT ;

begin

  Result   := True;

  if Action=taConsult then Exit ;
  if ( Piece.GetEnteteS('E_QUALIFPIECE')<>'N' ) then Exit ;
  if (Piece.ModeSaisie=msPiece) and (Piece.action<>taCreat) then Exit ;

  lStChamp := '' ;

  // Quel champ ?
  lStCode := Piece.Contexte.CPChampDoublon ;
  if lStCode='RFI'
    then lStChamp := 'E_REFINTERNE'
  else if lStCode='LIB'
    then lStChamp := 'E_LIBELLE'
  else if lStCode='RFX'
    then lStChamp := 'E_REFEXTERNE'
  else if lStCode='RFL'
    then lStChamp := 'E_REFLIBRE' ;

  if lStChamp = '' then Exit ;

  lTobRech := nil ;
  for i:=1 to Piece.Detail.count do
    if TrouveDoublonLigne( i, lStchamp, lTobRech ) then
      begin
      if PgiAsk('Il existe des doublons sur la référence ' + Piece.Detail[i-1].GetString( lStChamp )
                    + '. Confirmez-vous la validation de la pièce ?') <> mrYes then
        begin
        Result:=False ;
        if Assigned( lTobRech ) then
          begin
          if (PgiAsk('Voulez-vous visualiser le premier doublon détecté ?') = mrYes) then
            begin
            M := TOBToIdent( lTobRech, False ) ;
            LanceSaisieParam( Nil, taConsult, M ) ;
            end ;
          lTobRech.free ;
          end ;
        end ;
      Exit ;
      end ;

end;

procedure TOF_CPSAISIEPIECE.JOURNALChange(Sender: TObject);
var lInErr   : Integer ;
    lModeOld : TModeSaisie ;
begin

  if ( csDestroying in Ecran.ComponentState) then Exit ;

  lModeOld := Piece.ModeSaisie ;

  if GetControlText('E_JOURNAL')='' then Exit ;
{ Ce test est déporté en sortie de zone pour ne pas faire appraître le message inopinément
  if GetControlText('E_JOURNAL')='' then
    begin
    PGIBox('La saisie du journal est obligatoire', Ecran.Caption) ;
    E_Journal.SetFocus ;
    Exit ;
    end ;
}
  // Si inchangé on sort, sur evt utilisateur
  if (Sender<>nil) and ( GetControlText('E_JOURNAL') = Piece.GetEnteteS('E_JOURNAL') ) then Exit ;

  // Chargement du E_JOURNAL dans le TInfo
  if not InfoEcr.LoadJournal( GetControlText('E_JOURNAL') ) then Exit ;

  // vérification du journal
  lInErr := CIsValidJal( GetControlText('E_JOURNAL'), GetControlText('E_DEVISE'), InfoEcr ) ;
  if lInErr <> RC_PASERREUR then
    begin
    if lInErr = RC_JALNONMULTIDEVISE then
      InitDevise
    else
      begin
      FMessCompta.Execute( lInErr ) ;
      SetJournal( Piece.GetEnteteS('E_JOURNAL') ) ;
      SetFocusControl('E_JOURNAL') ;
      Exit ;
      end ;
    end ;

  // Maj pièce
  Piece.PutEntete('E_JOURNAL', E_JOURNAL.Value ) ;

  // MAJ de l'interface
  if (lModeOld <> Piece.ModeSaisie) then
    begin
    ZappePieceFolio ;
    TSP.SwapModeSaisie ( Piece.ModeSaisie ) ;
    if Piece.Count = 1 then
      begin
      Piece.ClearDetail ;
      Piece.newRecord ;
      end ;
    end ;

  // === Zones communes ===
  ParamNaturePiece ;   // Initialisation de la liste des natures de pièce
  InitDevise ;         // Devise
  InitEtablissement ;  // Etablissement

  // === Zone spécifique ===
  Case Piece.ModeSaisie of
    // Mode pièce
    msPiece : begin
              // Nature de pièce
              InitNaturePiece ;
              // Date comptable
              InitDateComptable ;
              // Numero de pièce
              InitNumeroPiece ;
              // Acces Grille
              TSP.Enabled := True ;
              end ;
    // Mode Bor et libre
    else
              begin
              // Acces Grille
//              TSP.Enabled := False ;
              // Période du folio
              InitPeriode ;
              // Numéro de folio
              ParamFolio ;
              // Set du numéro et Focus
              InitFolio( false, True);
              end ;

    end ;

  // Maj boutons et zones saisissables
  AccesEntete ;
  AccesBoutons ;
  MajMessageLigne( TSP.Row ) ;

  // MAJ grille suivant paramètrage
  TSP.RechercheMasque ;
  TSP.AfficheLignes ;

end;

procedure TOF_CPSAISIEPIECE.ZoomGeneralClick(Sender: TObject);
begin
  ZoomGeneral ;
end;

procedure TOF_CPSAISIEPIECE.ZoomTiersClick(Sender: TObject);
begin
  ZoomTiers ;
end;

procedure TOF_CPSAISIEPIECE.NextNature( vRow : integer );
var lStNat : string ;
    lIdx   : Integer ;
begin

  if not ( Piece.ModeSaisie in [msBOR, msLibre] ) then Exit ;
  if Piece.IsOut( vRow ) then Exit ;

  lIdx := E_NATUREPIECE.Values.IndexOf( Piece.GetValue( vRow, 'E_NATUREPIECE') ) ;
  Inc( lIdx ) ;
  if lIdx > E_NATUREPIECE.Values.Count - 1 then
    lIdx := 0 ;
  lStNat := E_NATUREPIECE.Values[lIdx] ;

  Piece.PutValue( vRow, 'E_NATUREPIECE', lStNat ) ;

  if Piece.ModeSaisie=msBOR
    then TSP.AfficheGroupe( vRow )
    else TSP.AfficheLignes( vRow ) ;

end;



procedure TOF_CPSAISIEPIECE.CompletePiece;
begin

  if Piece.GetEnteteS('E_JOURNAL') = '' then
    begin
    if (E_JOURNAL.values.count > 0) then
       E_JOURNAL.value := E_JOURNAL.values[0] ;
    end ;

  if (Piece.ModeSaisie in [msBor, msLibre]) and (Piece.GetEnteteI('E_NUMEROPIECE') = 0) then
      InitFolio( False, True ) ;

end;

procedure TOF_CPSAISIEPIECE.SetJournal(vStJal: string);
begin
  E_JOURNAL.OnChange   := nil ;
  E_JOURNAL.Value      := vStJal ;
  E_JOURNAL.OnChange   := JOURNALChange ;
end;

procedure TOF_CPSAISIEPIECE.TSP_KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin

  Case Key of

    VK_F6  : if Shift=[] then
               begin
               AccesBoutons ;
               end ;

    VK_DELETE : if Shift=[ssCtrl] then
                  begin
                  // Maj accès fonctions
                  AccesBoutons ;
                  // Maj infos pied
                  refreshPCPied ;
                  end ;
    VK_INSERT : if Shift=[] then
                  begin
                  // Maj accès fonctions
                  AccesBoutons ;
                  // Maj infos pied
                  refreshPCPied ;
                  end ;

{AG}     71 : if Shift=[ssAlt]  then
                begin
                AccesBoutons ;
                MajInfosLigne( TSP.Row ) ;
                end ;
  end ;

end;

procedure TOF_CPSAISIEPIECE.CalculSoldeGene(vStCpt: string);
var lTob    : Tob ;
    lDebG   : Double ;
    lCreG   : Double ;
    lDebH   : Double ;
    lCreH   : Double ;
begin

  if not Piece.Info_LoadCompte( vStCpt ) then Exit ;

  lTob := Piece.Info.Compte.Item ;
  if lTob.GetNumChamp('_DEBITHISTO') > 0 then exit ;

  lTob.AddChampSupValeur( '_DEBITHISTO', 0) ;
  lTob.AddChampSupValeur( '_CREDITHISTO', 0) ;

  Piece.Info.Compte.solde( lDebG, lCreG, Piece.Contexte.GetTypeExo( Piece.GetEnteteDt('E_DATECOMPTABLE') ) ) ;
  Piece.GetHistoPourChamps( 'E_GENERAL', vStCpt, lDebH, lCreH ) ;

  lTob.PutValue('_DEBITHISTO',  lDebG - lDebH ) ;
  lTob.PutValue('_CREDITHISTO', lCreG - lCreH ) ;

end;

procedure TOF_CPSAISIEPIECE.CalculSoldeAuxi(vStCpt: string);
var lTob    : Tob ;
    lDebG   : Double ;
    lCreG   : Double ;
    lDebH   : Double ;
    lCreH   : Double ;
begin

  if not Piece.Info_LoadAux( vStCpt ) then Exit ;

  lTob := Piece.Info.Aux.Item ;
  if lTob.GetNumChamp('_DEBITHISTO') > 0 then exit ;

  lTob.AddChampSupValeur( '_DEBITHISTO', 0) ;
  lTob.AddChampSupValeur( '_CREDITHISTO', 0) ;

  Piece.Info.Aux.solde( lDebG, lCreG, Piece.Contexte.GetTypeExo( Piece.GetEnteteDt('E_DATECOMPTABLE') ) ) ;
  Piece.GetHistoPourChamps( 'E_AUXILIAIRE', vStCpt, lDebH, lCreH ) ;

  lTob.PutValue('_DEBITHISTO',  lDebG - lDebH ) ;
  lTob.PutValue('_CREDITHISTO', lCreG - lCreH ) ;

end;

procedure TOF_CPSAISIEPIECE.AccActifClick(Sender: TObject);
begin
  if TSP.GuideActif then Exit ;
 TSP.AccArret          := not TSP.AccArret ;
 PopS.Items[0].Checked := TSP.AccArret ;
 MajMessageLigne( TSP.Row ) ;
end;

procedure TOF_CPSAISIEPIECE.AccTiersClick(Sender: TObject);
var lStCpt : string ;
begin

  if TSP.GuideActif then Exit ;
  if ( Action = taConsult ) or
     not ExJaiLeDroitConcept( TConcept(ccAuxModif), False ) then Exit ;

  lStCpt := Piece.GetString( TSP.Row, 'E_AUXILIAIRE' ) ;
  if (lStCpt = '') or not ( InfoEcr.LoadAux( lStCpt ) )  then Exit ;

  FicheTiersMZS( lStCpt,
                 taModifEnSerie,
                 7,
                 'YTC_SCHEMAGEN="'+Piece.GetCompteAcc(TSP.Row,TSP.AccCptHT)+'";YTC_ACCELERATEUR="X"'
                 ) ;

  InfoEcr.LoadAux(lStCpt,true) ;
  MajInfosLigne( TSP.Row ) ;

end;

procedure TOF_CPSAISIEPIECE.MajMessageLigne(vRow: Integer);
var lTexte : string ;
    lStAux : string ;
    lStCpt : string ;
begin
  // Restriction sur la ligne ?
  lTexte := '' ;
  // Guide en cours
  if TSP.GuideActif then
    lTexte := ''
  // Données sur dossier externe
  else if not PContexte.local then
    lTexte := lTexte + #13 + #10 + 'Pièce de la société : ' + PContexte.Libelle
  // Mvt non modifiable
  else if not Piece.EstLigneModifiable( vRow ) then
    lTexte := 'Mouvement non modifiable. ( lettré, pointé, compte ou journal fermé... )'
  // Accélérateur de saisie
  else if TSP.IsActiveAcc then
    begin
    if TSP.AccArret then
      lTexte := 'Accélérateur de saisie désactivé'
    else
      begin
      lTexte := 'Accélérateur de saisie activé' ;
      // détermination du compte
      lStAux := Piece.GetString( vRow,'E_AUXILIAIRE') ;
      if ( lStAux <> '' ) and InfoEcr.loadAux( lStAux ) then
        begin
        if ( InfoEcr.getString('YTC_ACCELERATEUR') = 'X') then
          begin
          lStCpt := InfoEcr.getString('YTC_SCHEMAGEN') ;
//          else lStCpt := Piece.GetCompteAcc( vRow, FStGenHT ) ;
          if (lStCpt <> '') and InfoEcr.LoadCompte(lStCpt) then
            begin
            lTexte := lTexte + ' : ' + lStCpt + ' - ' + InfoEcr.GetString('G_LIBELLE') ;
//          if ( InfoEcr.getString('YTC_ACCELERATEUR') = 'X')
            lTexte := lTexte + ' (AUTOMATIQUE)' ;
            end ;
          // repositionnement sur info de la ligne
          InfoEcr.LoadCompte(Piece.GetString( vRow,'E_GENERAL') );
          end ;
        end ;
      end ;

    end ;

  // Affichage
  SetControlText('SA_MESSAGE', lTexte ) ;

end;

procedure TOF_CPSAISIEPIECE.FormMouseDown(Sender: TObject;Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var lRect : TRect ;
begin

  if FBoMouseDown then Exit ;

  FBoMouseDown := true ;
  try

    if (not TSP.Enabled) then
      begin
      lRect := Rect( TSP.LaGrille.Left,
                     TSP.LaGrille.Top,
                     TSP.LaGrille.Left + TSP.LaGrille.Width,
                     TSP.LaGrille.Top  + TSP.LaGrille.Height) ;

      // On a clické sur la grille
      if PtInRect( lRect, Point(X, Y) ) then
        begin
//        if ( Screen.ActiveControl.Parent is TTabSheet ) then
//          nextPrevControl(Ecran, True, true)  ;

        if Piece.ModeSaisie in [msBor,msLibre] then
          begin
          nextPrevControl(Ecran, True, true)  ;
          ChargePiece ;
          end ;

        TSP.Setfocus ;
//        if GetAction<>taconsult then
//          TSP.LaGrille.MontreEdit;
//          TSP.LaGrille.InplaceEditor.Setfocus ;
        end ;

//    if GetAction=taconsult then
//      PostMessage(TSP.LaGrille.Handle, WM_LBUTTONUP, 0, 0) ;

    end ;

  finally
      FBoMouseDown := False ;
    end ;


end;

{JP 31/05/07 : FQ 17754 : déplacement en multi sociétés}
{---------------------------------------------------------------------------------------}
procedure TOF_CPSAISIEPIECE.ChangePiece(Sender : TObject);
{---------------------------------------------------------------------------------------}

    {---------------------------------------------------------------------}
    function _Deplace(aSender : TObject) : Boolean;
    {---------------------------------------------------------------------}
    begin
      Result := True;

           if (aSender as TToolBarButton97).Name = 'BFIRST' then FQMul.First
      else if (aSender as TToolBarButton97).Name = 'BPREV'  then FQMul.Prior
      else if (aSender as TToolBarButton97).Name = 'BNEXT'  then FQMul.Next
      else if (aSender as TToolBarButton97).Name = 'BLAST'  then FQMul.Last;

      if Assigned(FQMul.FindField('SYSDOSSIER')) and
         (Dossier <> FQMul.FindField('SYSDOSSIER').AsString) then begin

        if (aSender as TToolBarButton97).Name = 'BFIRST' then begin
          {Théoriquement, Result ne peut être que True, sinon cela signifie qu'il n'y a pas
           d'enregsitrement dans FQMul, donc on n'aurait pas pu entrer dans la fiche : en effet
           si on passe ici, c'est que le premier enregistrement a un dossier différent, donc on
           redescend le Query}
          if not FQMul.EOF then Result := _Deplace(GetControl('BNEXT'))
                           else Result := False;
        end
        else if (aSender as TToolBarButton97).Name = 'BPREV' then begin
          if not FQMul.BOF then Result := _Deplace(aSender)
                           else Result := False;
        end
        else if (aSender as TToolBarButton97).Name = 'BNEXT' then begin
          if not FQMul.EOF then Result := _Deplace(aSender)
                           else Result := False;
        end
        else if (aSender as TToolBarButton97).Name = 'BLAST' then begin
          {Théoriquement, Result ne peut être que True, sinon cela signifie qu'il n'y a pas
           d'enregsitrement dans FQMul, donc on n'aurait pas pu entrer dans la fiche : en effet
           si on passe ici, c'est que le dernier enregistrement a un dossier différent, donc on
           remonte le Query}
          if not FQMul.BOF then Result := _Deplace(GetControl('BPREV'))
                           else Result := False;
        end
        else
          Result := False;
      end;
    end;

var
  Res : Boolean;
begin
  FQMul.DisableControls;
  try
    {Déplacement récursif, jusqu'a ce que l'on trouve une pièce ayant le même dossier que celui
     de la pièce lors de l'ouverture de la fiche}
    Res := _Deplace(Sender);
    if not Res then begin
      {Dans le cas des boutons "suivant" et "précédent", Résult = False, lorsque l'on est sur le
       dernier ou le premier enregistrement : on inverse la commande pour se replacer sur l'enregistrement
       d'origine. Je n'utilise pas de BookMark, car GotoBookMark n'a pas d'impact sur la grille du MUL,
       ce qui fait que le mul est positionné sur la première ou la dernière ligne alors que le dossier
       peut être différent ; Trop de risque que cela génère une FQ}
      if (Sender as TToolBarButton97).Name = 'BNEXT' then _Deplace(GetControl('BPREV'));
      if (Sender as TToolBarButton97).Name = 'BPREV' then _Deplace(GetControl('BNEXT'));
      {On ne peut plus faire de Next, on désactive les deux boutons}
      SetControlEnabled('BNEXT' , ((Sender as TToolBarButton97).Name <> 'BNEXT'));
      SetControlEnabled('BLAST' , ((Sender as TToolBarButton97).Name <> 'BNEXT'));
      {On ne peut plus faire de Prev, on désactive les deux boutons}
      SetControlEnabled('BPREV' , ((Sender as TToolBarButton97).Name <> 'BPREV'));
      SetControlEnabled('BFIRST', ((Sender as TToolBarButton97).Name <> 'BPREV'));
    end
    else begin
      {Chargement de la pièce}
      ChargePieceFromMul;
      {Gestion des Boutons}
      SetControlEnabled('BFIRST', ((Sender as TToolBarButton97).Name <> 'BFIRST') and not FQMul.BOF);
      SetControlEnabled('BLAST' , ((Sender as TToolBarButton97).Name <> 'BLAST' ) and not FQMul.EOF);
      SetControlEnabled('BPREV' , ((Sender as TToolBarButton97).Name <> 'BFIRST') and not FQMul.BOF);
      SetControlEnabled('BNEXT' , ((Sender as TToolBarButton97).Name <> 'BLAST' ) and not FQMul.EOF);
    end;
  finally
    FQMul.EnableControls;
  end;
end;


procedure TOF_CPSAISIEPIECE.BValiderClick(Sender: TObject);
begin
  ClickValide ;
end;

procedure TOF_CPSAISIEPIECE.GereLesImmos(vRow: Integer);
{$IFDEF AMORTISSEMENT}
Var lTobEcr     : TOB ;
    lTobTiers   : TOB ;
    lMontantHT  : Double ;
    lMontantTVA : Double ;
    lStRegime   : String ;
    lBoTpf      : boolean ;
    lStCodeTva  : String ;
    lStCodeTpf  : String ;
    lTobImmo    : TOB ;
{$IFDEF COMPTAAVECSERVANT}
    lRecImmo    : TBaseImmo1 ;
{$ENDIF}
    lStImmo     : String ;
    lStLibAuxi  : String ;
    lStAuxi     : String ;
{$ENDIF}
BEGIN
{$IFDEF AMORTISSEMENT}
  if Action<>taCreat then Exit ;
  if TSP.GuideActif then Exit ;

  lTobEcr := Piece.GetTob( vRow ) ;
  if not Assigned( lTobEcr )  then Exit ;

  if Piece.GetString( vRow, 'E_DEVISE') <> V_PGI.DevisePivot then Exit ;
  if Piece.Info.GetString('J_NATUREJAL')<>'ACH' then
    if ( Piece.Info.GetString('J_NATUREJAL') <> 'OD' ) or
       ( Piece.GetString( vRow, 'E_NATUREPIECE') <> 'FF' ) then Exit ;
  if Piece.GetString( vRow, 'E_NATUREPIECE') = 'AF' then Exit ;

  if Piece.Info.GetString('G_NATUREGENE') <> 'IMO' then Exit ;

  if lTobEcr.GetDouble('E_DEBITDEV') = 0 then exit ;
  if lTobEcr.GetInteger('E_NUMEROIMMO')=666
    then Exit
    else lTobEcr.PutValue('E_NUMEROIMMO',666) ;

  if lTobEcr.GetString('E_IMMO')<>'' then Exit ;
  if PgiAsk( 'Voulez-vous créer une fiche d''immobilisation ?' )<>mrYes then Exit ;

  {$IFDEF AMORTISSEMENT}
  lMontantHT  := lTobEcr.GetDouble('E_DEBIT') + lTobEcr.GetDouble('E_CREDIT') ;

  lTobTiers   := Piece.TrouveLigneTiers( vRow ) ;

  lStLibAuxi   := '' ;
  lStAuxi      := '' ;
  lBoTpf       := False ;

  if ( lTobTiers <> nil ) and ( lTobTiers.GetString('E_AUXILIAIRE')<>'' ) then
    begin
    if Piece.Info.LoadAux( lTobTiers.GetString('E_AUXILIAIRE') ) then
      begin
      lStLibAuxi := Piece.Info.GetString('T_LIBELLE') ;
      lStAuxi    := Piece.Info.GetString('T_AUXILIAIRE') ;
      end ;
    lBoTpf     := Piece.EstSoumisTPF( lTobTiers.GetIndex + 1 ) ;
    end ;

  lStRegime  := Piece.GetRegimeTva( vRow ) ;
  lStCodeTva := Piece.GetCodeTva( vRow ) ;
  lStCodeTpf := Piece.GetCodeTpf( vRow ) ;

  lMontantTVA := HT2TVA(  lMontantHT, lStRegime, lBoTpf, lStCodeTva, lStCodeTpf, True, V_PGI.OkDecV ) ;
//HT2TVA( X, GeneRegTVA, GeneSoumisTPF, O.GetMvt('E_TVA'), O.GetMvt('E_TPF'), True, DECDEV)

  lTobImmo := TOB.Create('IMMO',nil,-1);
  lTobImmo.PutValue('I_DATEAMORT',      Piece.GetDateTime( vRow, 'E_DATECOMPTABLE') ) ;
  lTobImmo.PutValue('I_DATEPIECEA',     Piece.GetDateTime( vRow, 'E_DATECOMPTABLE') ) ;
  lTobImmo.PutValue('I_COMPTEIMMO',     Piece.GetString( vRow, 'E_GENERAL') ) ;
  lTobImmo.PutValue('I_MONTANTHT',      lMontantHT ) ;
  lTobImmo.PutValue('I_TVARECUPERABLE', lMontantTVA );
  lTobImmo.PutValue('I_TIERSA',         lStLibAuxi );
  lTobImmo.PutValue('I_REFINTERNEA',    Piece.GetString( vRow, 'E_REFINTERNE' ) );
  lTobImmo.PutValue('I_VALEURACHAT',    lMontantHT );
  lTobImmo.PutValue('I_QUANTITE',       Piece.GetDouble(vRow, 'E_QTE1') );
  if lTobImmo.GetInteger('I_QUANTITE') = 0 then
    lTobImmo.PutValue('I_QUANTITE',1);
  lTobImmo.PutValue('I_ETABLISSEMENT',  Piece.GetString( vRow, 'E_ETABLISSEMENT' ) ) ;
  lTobImmo.PutValue('I_LIBELLE',        Piece.GetString( vRow, 'E_LIBELLE' ) );
{$ENDIF}

  {$IFDEF COMPTAAVECSERVANT}
    if VHIM^.ServantPGI then
      begin
      lRecImmo.CompteImmo   := lTobImmo.GetString('I_COMPTEIMMO') ;
      lRecImmo.Fournisseur  := lStAuxi ;
      lRecImmo.Reference    := lTobImmo.GetString('I_REFINTERNEA') ;
      lRecImmo.Libelle      := lTobImmo.GetString('I_LIBELLE') ;
      lRecImmo.CodeEtab     := lTobImmo.GetString('I_ETABLISSEMENT') ;
      lRecImmo.DateAchat    := lTobImmo.GetDateTime('I_DATEPIECEA') ;
      lRecImmo.MontantHT    := lTobImmo.GetDouble('I_MONTANTHT') ;
      lRecImmo.MontantTVA   := lTobImmo.GetDouble('I_TVARECUPERABLE') ;
      lRecImmo.Quantite     := lTobImmo.GetDouble('I_QUANTITE') ;
      lStImmo              := SaisieExterne(taCreat,0,lRecImmo) ;
      end
    else if VH^.OkModImmo then
      lStImmo := AMLanceFiche_FicheImmobilisationEnSaisie ( lTobImmo );
    if Assigned( lTobimmo ) then
      FreeAndNil( lTobImmo ) ;
  {$ELSE}
    {$IFDEF AMORTISSEMENT}
      if VH^.OkModImmo then lStImmo:=AMLanceFiche_FicheImmobilisationEnSaisie ( lTobImmo );
      if Assigned( lTobimmo ) then
        FreeAndNil( lTobImmo ) ;
    {$ENDIF}
  {$ENDIF}
  lTobEcr.PutValue('E_IMMO', lStImmo ) ;
{$ENDIF}
end;

procedure TOF_CPSAISIEPIECE.ZoomImmoClick ( Sender: TObject ) ;
var lTobEcr : Tob ;
    lStImmo : string ;
{$IFDEF COMPTAAVECSERVANT}
    i : Integer ;
{$ENDIF}
begin

  if not BZoomImmo.Enabled then Exit ;

  lTobEcr := Piece.GetTob( TSP.Row ) ;
  if not Assigned( lTobEcr ) then Exit ;
  lStImmo := lTobEcr.GetString('E_IMMO') ;
  if lStImmo='' then Exit ;

  //  CA - 09/06/2006 - FQ 18340
  if Assigned(TheMulQ) then
    TheMulQ := nil;

  {$IFDEF COMPTAAVECSERVANT}
  if VHIM^.ServantPGI then
    begin
    for i:=1 To Length(lStImmo) do
      If lStImmo[i] in ['0'..'9']=FALSE Then Exit ;
    if not Presence('IMOREF','IRF_NUM',lStImmo)
      then TSP.MessageCompta.Execute( RC_NOEXISTIMMO )
      else VoirFicheImo( taConsult, ValeurI(lStImmo) ) ;
    end
  else
    begin
    {$IFDEF AMORTISSEMENT}
    if not Presence('IMMO', 'I_IMMO', lStImmo)
      then TSP.MessageCompta.Execute( RC_NOEXISTIMMO )
      else AMLanceFiche_FicheImmobilisation(lStImmo, taConsult, '') ;
    {$ENDIF}
    end ;
  {$ENDIF}
end;

procedure TOF_CPSAISIEPIECE.PCEnteteEnter(Sender: TObject);
begin

  if ( csDestroying in Ecran.ComponentState) then Exit ;

  if not (GoEditing in TSP.LaGrille.Options) and (FStatutSaisie=ssEditGrille) then
    begin
    if TSP.LaGrille.CanFocus then
      TSP.LaGrille.SetFocus ;
    Exit ;
    end ;

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
function TOF_CPSAISIEPIECE.TestQteAna ( vRow : integer ) : boolean;
begin
  if ctxPCL in V_PGI.PGIContexte then
    result := False
  else if Piece.GetScenario(vRow) <> nil then
    result := Piece.GetScenario(vRow).GetString('SC_CONTROLEQTE')='X'
  else
    result := Piece.Contexte.ZCtrlQte ;
end;

{$IFDEF SAISIEDOC}
procedure TOF_CPSAISIEPIECE.InitViewer;
var lPanelR, lPanelB : THPanel ;
    lPanelL, lPanelT : THPanel ;
    lSpR, lSpB       : THSplitter ;
    lSpL, lSpT       : THSplitter ;
begin

  // récup des coposant déjà poser sur la fiche
  lPanelR := THPanel( GetControl('PRIGHT', True ) );
  lSpR    := THSplitter( GetControl('SPRIGHT', True ) );
  lPanelB := THPanel( GetControl('PBOTTOM', True ) );
  lSpB    := THSplitter( GetControl('SPBOTTOM', True ) );
  lPanelL := THPanel( GetControl('PLEFT', True ) );
  lSpL    := THSplitter( GetControl('SPLEFT', True ) );
  lPanelT := THPanel( GetControl('PTOP', True ) );
  lSpT    := THSplitter( GetControl('SPTOP', True ) );

  TSP.InitViewer( lPanelB, lPanelR, lPanelL, lPanelT, lSpB, lSpR, lSpL, lSpT ) ;
  TSP.InitPopupViewer( TPopUpMenu( GetControl('POPUPVIEWER') ) , TToolBarButton97(GetControl('BMENUVIEWER')) ) ;
//  TSP.ActiverViewer ;
end;
{$ENDIF SAISIEDOC}

{$IFDEF SAISIEDOC}
procedure TOF_CPSAISIEPIECE.BAgrandirClick(Sender: TObject);
begin
  PCPied.Visible := False ;
  SetControlVisible('BAGRANDIR', False ) ;
  SetControlVisible('BREDUIRE',  True ) ;
  // resize du viewer de doc s'il est en bas
  if TSP.TSDoc.IsActif and (TSP.TSDoc.Position=vpBas) then
    TSP.TSDoc.PanelActif.Height := TSP.TSDoc.PanelActif.Height + PCPied.Height ;
  Ecran.Refresh ;
end;
{$ENDIF SAISIEDOC}

{$IFDEF SAISIEDOC}
procedure TOF_CPSAISIEPIECE.BReduireClick(Sender: TObject);
begin
  PCPied.Visible := True ;
  SetControlVisible('BAGRANDIR', True ) ;
  SetControlVisible('BREDUIRE',  False ) ;
  // resize du viewer de doc s'il est en bas
  if TSP.TSDoc.IsActif and (TSP.TSDoc.Position=vpBas) then
    begin
    if TSP.TSDoc.PanelActif.Height > PCPied.Height then
      TSP.TSDoc.PanelActif.Height := TSP.TSDoc.PanelActif.Height - PCPied.Height ;
    end ;
  Ecran.Refresh ;
end;
{$ENDIF SAISIEDOC}


Initialization
  registerclasses ( [ TOF_CPSAISIEPIECE ] ) ;
end.



