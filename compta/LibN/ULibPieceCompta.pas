unit ULibPieceCompta;

interface

uses
  Classes,           // TStringList
  UTob,              // Tob
  HEnt1,             // TActionFiche
  SaisUtil,          // RDevise
  Ent1,
  {$IFDEF MODENT1}
  CPObjetGen,
  CPTypeCons,
  {$ENDIF MODENT1}
  uLibEcriture,      // TInfoEcriture
  uLibExercice,      // TZExercice
  SaisComm           // TTypeExo
  ;

Type

TModeVentil  = ( mvAucun, mvAttente, mvDefaut ) ;
TModeEche    = ( meAucun, meMono, meMulti, meDeporte ) ;
TModeSaisie  = ( msPiece, msBor, msLibre ) ;
TModeFonc    = ( mfPiece, mfAlone ) ;
TModeGroupe  = ( mgDynamic, mgStatic ) ;

TPieceCompta = Class ;

TDossierContexte = Class
    protected
      // Nom de la base
      FDossier          : String ;
    public
      // Informations
      function  local : Boolean ;
      // Propriétés
      property Dossier       : String     read  FDossier ;

  end ;


TPieceContexte = Class( TDossierContexte )
    private

      FTobEntete        : TOB ;          // Entête de la piece

      FExercices        : TZExercice ;   // Liste des exercices du dossier cible

      FZScenario        : TZScenario ;   // Liste des scénarii

      FLiensEtab        : TOB ;          // Paramètrage multi-Etab contenant les liaisons inter-établissement

      FZHalleyUser      : TZHalleyUser ;

      // ==========================
      // === Paramètres Société ===
      // ==========================
      // Société
      FSociete          : string ;
      FLibelle          : string ;

      // Paramètre Tva
      FOuiTvaEnc        : boolean ;
      FCollCliEnc       : String ;
      FCollFouEnc       : String ;
      FTvaEncSociete    : String ;
      FRegimeDefaut     : String ;
      FCodeTvaDevaut    : string ;

      // Autres paramètres
      FCPChampDoublon   : String ;
      FAttribRIBAuto    : Boolean ;
      FPaysLocalisation : String ;
      FEtablisDefaut    : String ;
      FEtablisCpta      : Boolean ;
      FJalMultiEtab     : string ;
      FTypeEngagement   : String ;
      FCPPCLSAISIEQTE   : boolean ;

      // Paramètres Tiers Payeurs
      FOuiTP            : Boolean ;
      FJalATP           : string ;
      FJalVTP           : string ;

      // Mode révision
      FDateRevision     : TDateTime ;

      // Ano dynamiques
      FCPAnoDyna        : boolean ;

      // paramètres analytiques
      FCPAnaSurMontant  : boolean ;
      FZCtrlQTE         : boolean ;

      // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
      FBoInsertSpecif  : boolean ;
      FBoAttribRIBOff  : boolean ;

      // interactions utilisateur
      FCPDateObli           : boolean ;
      FCPEcheOuvrir         : boolean ;

      // Paramétres de dates
      FDateCloturePer  : TDateTime ;
      FNbJEcrAvant     : integer ;
      FNbJEcrApres     : integer ;

    procedure ChargeTva ;
    procedure ChargeLiensEtab ;
    procedure Initialisation( vDossier : string = '' ) ;
    function  GetScenario : TOB ;
    function  GetExercices: TZExercice;

    function GetSociete             : string;       // ok
    function GetLibelle             : string;       // ok
    function GetOuiTvaEnc           : boolean;      // ok
    function GetCollCliEnc          : String;       // ok
    function GetCollFouEnc          : String;       // ok
    function GetTvaEncSociete       : String;       // ok
    function GetRegimeDefaut        : String;       // ok
    function GetCodeTvaDevaut       : string;       // ok
    function GetCPChampDoublon      : String;       // ok
    function GetAttribRIBAuto       : Boolean;      // ok
    function GetPaysLocalisation    : String;       // ok
    function GetFEtablisDefaut      : String;       // ok
    function GetEtablisCpta         : Boolean;      // ok
    function GetJalMultiEtab        : string;       // ok
    function GetTypeEngagement      : String;       // ok
    function GetOuiTP               : Boolean;      // ok
    function GetJalATP              : string;       // ok
    function GetJalVTP              : string;       // ok
    function GetDateRevision        : TDateTime;    // ok
    function GetCPAnoDyna           : boolean;      // ok
    function GetCPAnaSurMontant     : boolean;      // ok
    function GetZCtrlQTE            : boolean;      // ok
    function GetCPPCLSAISIEQTE      : boolean;      // ok
    function GetDateCloturePer      : TDateTime;
    function GetNbJEcrAvant         : integer ;
    function GetNbJEcrApres         : integer ;
    function GetCPDateObli          : boolean;      // ok
    function GetCPEcheOuvrir        : boolean;      // ok

    public

      NumCodeBase      : Array[1..4] of String[3] ;

      // Construteurs / Destructeurs
      Class function CreerContexte( vDossier : String ) : TPieceContexte ;

      // Construteurs / Destructeurs
      Destructor Destroy ; override ;

      // Initialisations
      procedure ChargeInfos ;
      function  ChargeScenario       ( vStJal, vStNatP, vStQualP, vStEtab : String ) : Boolean ;
      procedure VideScenario         ;
      function  InitEntetePiece      : TOB ;

      // Infos de saisie PCL
      procedure GetZLastJal          ( var vStJournal : string ) ;
      procedure GetZLastDate         ( var vDtDateC : TDateTime ) ;
      procedure SetZLastInfo         ( vStJournal : string ; vDtDateC : TDateTime ) ;

      // Informations
      function  GetEntete         ( vChamps : String )     : Variant ;
      function  GetLiensEtab      ( vEtabDest : String )   : TOB ;
      function  GetTypeExo        ( vDate : TDateTime )    : TTypeExo ;
      function  GetExoDT          ( vDate : TDateTime )    : string ;
      function  GetLibelleEtab    ( vStCodeEtab : string ) : string ;
      function  GetEtablisDefaut                          : string ;
      function  EstEngagementHT                            : Boolean ;

      // Enregistrements
      function  InsertDB( vTobEcr : TOb ) : Boolean ;

      // Test saisie
      function ModeRevisionActive ( vDtDate : TDateTime ) : boolean ;

      // Propriétés
      property Scenario        : TOB           read  GetScenario  ;
      property LiensEtab       : TOB           read  FLiensEtab   ;
      property EntetePiece     : TOB           read  FTobEntete       write  FTobEntete ;
      property Exercices       : TZExercice    read  GetExercices ;
      property HalleyUser      : TZHalleyUser  read  FZHalleyUser ;
      // paramètres
      property Societe          : string      read  GetSociete ;
      property Libelle          : string      read  GetLibelle ;
      property OuiTvaEnc        : boolean     read  GetOuiTvaEnc ;
      property CollCliEnc       : String      read  GetCollCliEnc ;
      property CollFouEnc       : String      read  GetCollFouEnc ;
      property TvaEncSociete    : String      read  GetTvaEncSociete ;
      property RegimeDefaut     : String      read  GetRegimeDefaut ;
      property CodeTvaDevaut    : string      read  GetCodeTvaDevaut ;
      property CPChampDoublon   : String      read  GetCPChampDoublon ;
      property AttribRIBAuto    : Boolean     read  GetAttribRIBAuto ;
      property PaysLocalisation : String      read  GetPaysLocalisation ;
      property EtablisDefaut    : String      read  GetFEtablisDefaut ;
      property EtablisCpta      : Boolean     read  GetEtablisCpta ;
      property JalMultiEtab     : string      read  GetJalMultiEtab ;
      property TypeEngagement   : String      read  GetTypeEngagement ;
      property OuiTP            : Boolean     read  GetOuiTP ;
      property JalATP           : string      read  GetJalATP ;
      property JalVTP           : string      read  GetJalVTP ;
      property DateRevision     : TDateTime   read  GetDateRevision ;
      property CPAnoDyna        : boolean     read  GetCPAnoDyna ;
      property CPAnaSurMontant  : boolean     read  GetCPAnaSurMontant ;
      property ZCtrlQTE         : boolean     read  GetZCtrlQTE ;
      property InsertSpecif     : Boolean     read  FBoInsertSpecif ;
      property AttribRIBOff     : Boolean     read  FBoAttribRIBOff  write FBoAttribRIBOff ;
      property CPPCLSAISIEQTE   : boolean     read GetCPPCLSAISIEQTE ;  // utilisé
      property DateCloturePer   : TDateTime   read  GetDateCloturePer ;
      property NbJEcrAvant      : integer     read  GetNbJEcrAvant ;
      property NbJEcrApres      : integer     read  GetNbJEcrApres ;
      property CPDateObli       : boolean     read GetCPDateObli ;      // utilisé
      property CPEcheOuvrir     : boolean     read GetCPEcheOuvrir ;    // utilisé

  end ;
//////////////////////////////////////////////////////////////////////////////////////////


TMultiEche = Class(TOB)
  //////////////////////////////////////////////////////////////////////////////////////////
  private
    FModeFonc       : TModeFonc ;
    FStModeInit     : String ;
    FStModeFinal    : String ;

    FPiece          : TPieceCompta ;
    FTobEcr         : TOB ;
    FBoModif        : Boolean ;
    FAction         : TActionFiche ;
    FMontantSaisi   : Double ;

    function  GetNumLigne     : Integer ;
    function  GetNbEche       : Integer ;
    function  GetDevise       : RDevise ;
    function  GetInfo         : TInfoEcriture ;
    function  GetEcrIndex     : integer ;
    function  GetDetail       : TTobList ;
    function  GetMontantSaisi : Double ;

    function  GetModifTVA : Boolean ;
    procedure SetModifTVA ( vBoModif : Boolean  ) ;

    function  EstDebit       : Boolean ;

  //////////////////////////////////////////////////////////////////////////////////////////
  Public

    // ======================
    // ===== PROPRIETES =====
    // ======================
    property Devise       : RDevise       read  GetDevise ;
    property ModeInit     : String        read  FStModeInit   write  FStModeInit ;
    property ModeFinal    : String        read  FStModeFinal  write  FStModeFinal ;
    property ModifTVA     : Boolean       read  GetModifTVA   write  SetModifTVA ;
    property ModifEnCours : Boolean       read  FBoModif      write  FBoModif ;

    property Piece        : TPieceCompta  read  FPiece ;      // Référence à la pièce comptable
    property TobEcr       : TOB           read  FTobEcr ;     // 1ère ligne de l'échéance (ou ligne de référence de la pièce)
    property Info         : TInfoEcriture read  GetInfo ;     // référence au TInfoEcriture

    property Action       : TActionFiche  read  FAction       write FAction ;
    property NumLigne     : Integer       read  GetNumLigne ;
    property NbEche       : Integer       read  GetNbEche ;

    property DetailEche   : TTobList      read  GetDetail ;
    property MontantSaisi : Double        read  GetMontantSaisi  ;

    // ====================
    // ===== METHODES =====
    // ====================

    // Constructeur / Destructeur
    constructor Create( vTobParent : TOB ; vPiece : TPieceCompta ; vTobEcr : Tob ; vAction : TActionFiche ; vModeFonc : TModeFonc = mfPiece ) ;  reintroduce; overload ; virtual ;
    destructor  Destroy ; override ;

    // Mode de fonctionnement
    procedure SetModePiece ;
    procedure SetModeAlone ;

    // initialisation
    procedure SetLignePiece    ( vTobEcr : Tob ) ;

    // Indicateurs
    function  IsOutEche        ( vNumEche : Integer )      : Boolean ;
    function  EstLettre        ( vNumEche : Integer )      : Boolean ;
    function  EstModifiable    ( vNumEche : Integer = -1 ) : Boolean ;
    function  GetMontantEche   ( vNumEche : Integer )      : Double ;
    function  GetTotalEche     ( vBoDev : Boolean = True )  : Double ;
    function  EstOkMultiEche   : boolean ;

    // Gestion des lignes
    function  NewEche          ( vNumEche : Integer = 0 ) : TOB ;
    function  GetEche          ( vNumEche : Integer )     : TOB ;
    procedure AddEche          ( vTobEcr : TOB ) ;
    procedure Renumerote       ( vInFrom : integer = -1 ) ;
    procedure ClearEche ;

    // Calculs de l'échéance
    procedure SetMontantSaisie ( vMontant : double ) ;
    procedure SetMontant       ( vNumEche : integer ; vMontant : double ) ;
    Function  GetModeInit : String ;
    procedure CalculEche       ( vStModeForce : String = '' ) ;
    procedure CalculModeFinal ;

    procedure ProratiseEche    ( vMontant : Double ) ;
    procedure UpdateEche       ( vNumEche : Integer ; vModePaie : String ; vMontant : Double ; vDateEche : TDateTime ) ;
    function  GereArrondi      : Boolean ;
    procedure GerePoucentage ;

    // Calcul des champs TVa
    Procedure SetTvaEnc( vTvaEnc1, vTvaEnc2, vTvaEnc3, vTvaEnc4, vTvaDebit : Double ; lCoefTiers : Double = 1 ) ;

    // Compatibilité T_MODEREGL <==> TMULTIECHE
    Function  GetModeRegl  : T_MODEREGL ;
    procedure SetModeRegl  ( vModeRegl : T_MODEREGL ) ;

    //=========================================
    //=== Synchronisme Piece / Echeances... ===
    //=========================================
    procedure CumuleTobEcr ;
    procedure SynchroniseEcheances ;
    procedure SynchroniseEcr ;
    // --> unifromisation des données NUMECHE=1 --> NUMECHE>1
    procedure ReporteEche( vChamps : string ; vValeur : Variant ) ;
    //Recopie
    procedure CopieEcheances ( vSource : TMultiEche ) ;
    // Reprise des info
    Procedure PutTobEche     ( vNumEche : Integer ; vTobSource : TOB ) ;


  end;


TPieceCompta = Class(TOBCompta)
  private

    // Société
    FContexte     : TPieceContexte ;   // Paramètres de la société courante

    // données en base
    FTobOrigine   : TOB ;              // Contient la pièce avant modification

    // Devise en cours
    FDevise       : RDevise ;
    FTobChancell  : Tob ;

    // données d'entête
    FEntete       : TOB ;              // Contient les données d'identification de la pièce + infos diverses

    // données échéances
    FEcheances    : TOB ;              // Contient les données concernant les échéances pour chaque ligne

    // Pièce multi-etab
    FPiecesEtab   : TStringList ;            // Contient les pièces multi-établissement

    // données EcrCompl
    FEcrCompl     : TOB ;

    // Indicateurs divers
    FModeSaisie   : TModeSaisie  ;     // Type de saisie ( msPiece, msBor, msLibre )
    FModeGroupe   : TModeGroupe ;      // RecalculAuto des groupes ( mgDynamic, mgStatic ) ;
    FAction       : TActionFiche ;     // Peut prendre 3 valeurs : taCreat, taModif, taConsult
    FBoModif      : Boolean ;          // Pièce modifiée ?
    FCurIdx       : Integer ;          // Index de la ligne courante référencée
    FFromTob      : TOB ;
    FBoLockFolio  : Boolean ;          // indicateur de blocage du folio en cours

    // multi-échéance
    FModeEche     : TModeEche ;        // Permet la désactivation de la gestion du multi-échéance
    // Analytique
    FModeVentil   : TModeVentil ;      // Type de ventilation auto : ( mvAucun, mvAttente, mvDefaut ) ;
    // Accélérateur de saisie
    FBoAccActif   : Boolean ;          // Permet la désactivation de l'accélérateur de saisie
    // Activation des scénarii dee saisie
    FBoScenActif  : Boolean ;          // Permet la désactivation des évènements sur scénario de saisie

    // Tiers Payeurs
    FBoTPActif    : Boolean ;          // Indicateur de gestion des tiers payeurs
    FYaTP         : T_YATP ;           // indicateur de présence d'une pièce de tiers payeurs :
                                       //   - yaRien : pas de pièce correspondante
                                       //   - yaNL   : présence d'une pièce non lettrée
                                       //   - yaL    : présence d'une pièce lettrée
    FListeTP      : TStringList ;      // Contient la liste des références au pièces de TP au chargement de la pièce

    {JP 20/07/07 : FQ 20601 : Le numéro de pièce, sert en ouverture de bordereau à bien se positionner
                   et non pas en fin de bordereau}
    FNumLigneAppel : Integer;

    // Gestion de la TVA
    FTabTvaEnc    : Array[1..5] of Double ;

    function  _Save : boolean ;

    Procedure CPutValue  ( vIndex : Integer ; vChamps : String ; vValeur : Variant ) ;
    function  GetCurrentTob : TOB ;

    // Calcul des soldes
    procedure CalculDesSoldes ;
    Function  ADevalider : Boolean ;

    // Chargement des données
    procedure ChargeInfo ;
    procedure ChargeEntete ( vTob : TOB = nil ) ;
    procedure ChargeEche ;
    procedure ChargeChampsSup ;
    procedure ChargeVentil ;
    procedure ChargeEcrCompl ; // Gestion du Cut-Off via EcrCompl

    // 12/05/06 : Gestion des BAP et éventuelles mises en lecture seule de la pièce
    procedure GestionDesBAP;

    // Gestion interne des bordereaux
    procedure ReCalculGroupe       ( vIndex : integer = 0 );

    // GEstionnaire d'evts internes suite à la modification d'un champs
    function  OnChangeField        ( vIndex :Integer ; vChamps : String ; vOldValue, vNewValue : Variant ) : Boolean ;

    procedure AffecteConso         ( vIndex : Integer ) ;
    procedure AffecteConf          ( vIndex : Integer ) ;
    procedure AffecteAna           ( vIndex : Integer ) ;
    procedure AffecteRIB           ( vIndex : Integer ) ;
    procedure AffecteEche          ( vIndex : Integer ) ;
    procedure AffecteTypeMvt       ( vIndex : Integer ) ;
    procedure AffecteTvaEnc        ( vIndex : Integer ) ;
    procedure ReinitEche           ( vIndex : Integer ) ;

    // fonctions de lecture des Property
    function  GetDossier : String ;

    // traitements internes
//    procedure CompleteEcriture        ( vIndex : Integer ) ;
    procedure TermineLaPiece        ;

    // fonctions interne de recherche
    Function  TrouveLigneHT           ( vIndex : integer = 0 ; vFromLigne : Integer = 0 ) : Tob ;
    Function  TrouveLigneTVA          ( vIndex : integer = 0 ; vFromLigne : Integer = 0 ) : Tob ;

    // Init des données
    procedure InitNewRecord           ( TobEcr : TOB ; vIndex : Integer ; lBoARenumeroter : Boolean ) ;

  Protected

    procedure Initialisation ; virtual;  // Lek déplacer et ajouter virtual pour classe descendant
    procedure ChargeExtension; virtual;  // Lek créer pour classe descendant
    procedure CFreeExtension   ( vTobEcr : Tob ) ; virtual ;


  //////////////////////////////////////////////////////////////////////////////////////////
  Public
    ConserverBap : Boolean;
    GestionBapOk : Boolean;

    // ======================
    // ===== PROPRIETES =====
    // ======================
    property ModeSaisie    : TModeSaisie    read  FModeSaisie     write  FModeSaisie ;
    property ModeEche      : TModeEche      read  FModeEche       write  FModeEche ;
    property ModeVentil    : TModeVentil    read  FModeVentil     write  FModeVentil ;
    property ModeGroupe    : TModeGroupe    read  FModeGroupe     write  FModeGroupe ;
    property Entete        : TOB            read  FEntete   ;
    property Echeances     : TOB            read  FEcheances ;
    property TobOrigine    : TOB            read  FTobOrigine   ;
    property Action        : TActionFiche   read  FAction         write  FAction ;
    property Devise        : RDevise        read  FDevise         write  FDevise ;
    property ModifEnCours  : Boolean        read  FBoModif        write  FBoModif ;
    property PiecesEtab    : TStringList    read  FPiecesEtab ;
    property Dossier       : String         read  GetDossier ;
    property Contexte      : TPieceContexte read  FContexte ;
    property CurIdx        : integer        read  FCurIdx         write  FCurIdx ;
    property CurTob        : Tob        read  GetCurrentTob ;
    property AccActif      : Boolean        read  FBoAccActif     write  FBoAccActif ;
    {JP 20/07/07 : FQ 20601 : Le numéro de pièce, sert en ouverture de bordereau à bien se positionner
                   et non pas en fin de bordereau}
    property NumLigneAppel : Integer        read  FNumLigneAppel  write  FNumLigneAppel;

    // ====================
    // ===== METHODES =====
    // ====================

    destructor  Destroy ; override ;

    Class function CreerPiece       ( vInfoEcr : TInfoEcriture = nil ) : TPieceCompta ;

    Class function GetChampsGene    : String ;
    Class function GetChampsAuxi    : String ;


    // Chargement
    procedure LoadFromTOB           ( vTobLignes : TOB ; vInfo : TInfoEcriture = nil ) ;
    procedure LoadFromSQL           ( vStCondSQL : String = '') ;
    function  GetWhereSQL           : String ;
    procedure ChargeInfoOpti        ( vTobSource : Tob ) ;
    function  EstEnregModifiable    : boolean ;

    // Optimisation recherche dans TInfoEcriture
    function  Info_LoadCompte       ( vStCode : string ) : Boolean ;
    function  Info_LoadAux          ( vStCode : string ) : Boolean ;
    function  Info_LoadJournal      ( vStCode : string ) : Boolean ;
    function  Info_LoadDevise       ( vStCode : string ; vBoEnBase : boolean = False ) : Boolean ;
    function  Compte_GetValue       ( vStNom : string )  : Variant ;
    function  Aux_GetValue          ( vStNom : string )  : Variant ;
    function  Journal_GetValue      ( vStNom : string )  : Variant ;

    // Gestion du scénario
    function  GetScenario           ( vIndex : integer = 0 ) : TOB ;
    function  GetMemoScenario       ( vIndex : integer = 0 ) : HTStringList ;
    function  ChargeScenario        ( vIndex : integer = 0 ) : Boolean ;
    function  GetCompLigne          ( var vStComporte, vStLibre : string ; vNumLigne : integer = 0 ) : integer ;
    procedure AffecteEnteteScenario ( vIndex : integer ; vBoForce : Boolean = false ) ;
    procedure InitEnteteScenario    ;
    function  GereQtePCL            ( vIndex : integer ) : Boolean ;

    // Calculs totaux / solde de la pièce
    Function  GetTotalDebit         ( vBoDev : Boolean = True ) : Double ;
    Function  GetTotalCredit        ( vBoDev : Boolean = True ) : Double ;
    Function  GetSolde              ( vBoDev : Boolean = True ) : Double ;
    function  GetSoldePartiel       ( vInDe, vInA : integer ; vBoDev : Boolean = True ) : Double ;
    Function  EstPieceSoldee        ( vIndex : integer = 0 ) : Boolean ;
    Function  EstPieceEquilibree    ( vIndex : integer = 0 ) : Boolean ;
    Function  ProratisePiece        ( vIndex : Integer ; vChamps : String ; vNewMontant : Double ) : Boolean ;
    Function  GetMontant            ( vIndex : Integer )     : Double ;
    Function  GetMontantDev         ( vIndex : Integer )     : Double ;
    Procedure GetInfosLigne         ( vIndex : Integer ; var vDebit, vCredit, vTotalDebit, vTotalCredit : Double ) ;
    Procedure GetTotauxPourChamps   ( vChamps : String ; vValeur : String ; var vTotalDebit, vTotalCredit : Double ; vInGroupe : integer = -1) ;
    Procedure GetHistoPourChamps    ( vChamps : String ; vValeur : String ; var vTotalDebit, vTotalCredit : Double ) ;
    function  VerifieEquilibre      ( vIndex : integer = 0 ) : boolean ;

    // Gestion Filles
    function  CGetTob               ( vIndex : Integer ) : Tob ;
    function  GetCopyTob            ( vIndex : Integer ) : Tob ;
    Function  GetTobLigne           ( vNumLigne : Integer ; vNumEche : Integer = - 1 ) : TOB ;
    Function  GetTob                ( vIndex : Integer ) : TOB ;
    Function  GetValue              ( vIndex : Integer ; vChamps : String ) : Variant ;   overload ;virtual;
    Function  GetInteger            ( vIndex : Integer ; vChamps : String ) : Integer ;   overload ;virtual;
    Function  GetString             ( vIndex : Integer ; vChamps : String ) : String ;    overload ;virtual;
    Function  GetDateTime           ( vIndex : Integer ; vChamps : String ) : TDateTime ; overload ;virtual;
    Function  GetDouble             ( vIndex : Integer ; vChamps : String ) : Double ;    overload ;virtual;
    Procedure PutValue              ( vIndex : Integer ; vChamps : String ; vValeur : Variant ; vBoForce : boolean = False ) ; overload ;virtual;
    Procedure PutValueSpecif        ( vIndex : Integer ; vChamps : String ; vValeur : Variant ) ; virtual;
    Procedure PutValueAll           ( vChamps : String ; vValeur : Variant ) ;
    Procedure LoadInfo              ( vIndex : Integer ) ;
    procedure videPiece             ;
    Procedure InitSaisie            ;
    Procedure InitVariables         ;
    Function  Count                 : Integer ;
    Function  GetRMVT               : RMVT ;

    // Gestion données d'entête
    procedure InitPiece            (M: RMVT); overload; {FP 02/05/2006}
    Procedure InitPiece            ( vStJournal : String ; vDtDateComptable : TDateTime ; vStNaturePiece : string = '' ;
                                      vStDevise : string = '' ; vStEtablissement : string = '' ; vStQualifPiece : string = '' ;
                                      vStEcrANO : String = '') ; overload;

    Procedure PutEntete             ( vChamps : String ; vValeur : Variant ) ;
    function  GetEntete             ( vChamps : String ) : Variant ;
    function  GetEnteteS            ( vChamps : String ) : String ;
    function  GetEnteteI            ( vChamps : String ) : Integer ;
    function  GetEnteteDt           ( vChamps : String ) : TDateTime ;
    function  GetNatureInitiale     : string ;
    function  GetWhereNatPiece      : string ;
    procedure AffecteModeSaisie     ;

    // Ajout et Suppression d'écritures
    function  NewRecord             ( vIndex : Integer = -1 ) : TOB ;  virtual ; // ok
    procedure DeleteRecord          ( vIndex : Integer ) ; virtual ;             // ok

    procedure DupliquerPiece        ( PieceSrc: TPieceCompta ); virtual ;                                                    {FP 02/05/2006}
    procedure DupliquerLigne        ( vNumLigneDst : Integer ; PieceSrc : TPieceCompta ; vNumLigneSrc : Integer ) ; virtual; {FP 02/05/2006}
    procedure CumulerMontant        ( vNumLigneDst : Integer ; PieceSrc : TPieceCompta ; vNumLigneSrc : Integer ) ;          {FP 02/05/2006}

    // Enregistrement
    function  Save                  : Boolean ; virtual;    {Lek+FP b ajoute virtual pour class descendant}
    procedure DetruitPiece          ( vBoPourMAJ : boolean = False ) ;

    // Vérification des lignes
    Function  IsValidDateCreation   ( vDtDateC : TDateTime ) : boolean ;
    Function  IsValidPiece          : Boolean ;
    Function  IsValidLigne          ( vIndex : Integer ) : Boolean ; // ok
    Function  IsValidDebitCredit    ( vIndex : Integer ) : Boolean ;
    Function  EstLigneModifiable    ( vIndex : Integer ) : Boolean ;
    Function  EstChampsModifiable   ( vIndex : Integer ; vNomChamp : string ) : Boolean ;
    Function  PieceModifiable       ( vIndex : Integer = 0) : Boolean ;
    Function  ExisteLettrage        ( vIndex : Integer ) : Boolean ;
    Function  ExisteEcheance        ( vIndex : Integer ) : Boolean ;
    Function  RechercheChampDoublon ( vIndex : integer ; vStChp : string ; var vTobRech : Tob ) : Boolean ;
    Function  IsValidTaux           ( vIndex : Integer ) : Boolean ;
    Function  IsValidTva            ( vIndex : Integer ) : Boolean ;
    Function  IsValidEngagement     : Boolean ;
    procedure CSupprimeLigneVide    ;

    // Gestion des devises
    Function  GetRDevise            ( vIndex : integer = 0 ; vBoMajTaux : boolean = False ) : RDevise ;
    procedure AffecteDevise         ( vIndex : Integer     ) ;
    procedure majDeviseTaux         ( vIndex : integer = 0 ; vBoEnBase : boolean = False ) ;
    procedure ChangeTauxPiece       ( vIndex : integer = 0 ) ;
    function  EstMonoDevise         : boolean ;
    procedure SetTauxVolatil        ( vIndex : integer ; vDev: RDevise ) ;
    function  EstTauxVolatil        ( vIndex : Integer     ) : boolean ;

    // Gestion cohérence des données
    procedure AttribNumeroTemp      ;
    procedure AttribNumeroDef       ;
    procedure AttribSolde           ( vIndex : Integer ) ;
    procedure RenumeroteLignes      ( Offset    : Integer = 0) ;

    // Recherche de ligne
    Function  TrouveLigneCompte       ( vStCompte  : String ; vFromLigne : Integer = 0 ) : Tob ;
    Function  TrouveLigneTiers        ( vIndex : integer = 0 ; vFromLigne : Integer = 0 ) : Tob ;
    Function  TrouveIndiceLigneCompte ( vStCompteGene, vStCompteAuxi: String; vFromLigne: Integer = 1): Integer; {FP 02/05/2006}

    // gestion de la TVA
    Function  GenereTVA           ( vIndex : integer ) : Integer ;
    procedure RegroupeTVA         ( vIndex : integer ) ;
    Function  EstAchat            ( vIndex : Integer = 0) : Boolean ;
    Function  EstSoumisTPF        ( vIndex : Integer ) : Boolean;
    Function  GetSorteTVA         ( vIndex : Integer ) : TSorteTva ;
    Function  GetTvaMode          ( vIndex : Integer ) : TExigeTva ;
    Function  GetTvaEnc           ( vIndex : Integer ) : String ;
    Function  GetCompteTva        ( vIndex : Integer ) : String ;
    Function  GetCompteTpf        ( vIndex : Integer ) : String ;
    Function  GetCodeTva          ( vIndex : Integer ) : String ;
    Function  GetCodeTpf          ( vIndex : Integer ) : String ;
    Function  GetTauxTva          ( vIndex : Integer ) : Double ;
    Function  GetTauxTpf          ( vIndex : Integer ) : Double ;
    Function  GetRegimeTva        ( vIndex : Integer = 0 ) : String ;
    procedure SetTvaMode          ( vIndex : integer ; vTvaMode : TExigeTva ; vBoForce : Boolean = False ) ;
    Function  GereTvaMode         ( vIndex : Integer ) : Boolean ;
    procedure CalculTvaEnc        ( vIndex : integer ) ;
    procedure SetTvaTauxDirecteur ( vIndex : Integer ) ;
    procedure SetTvaEnc           ( vIndex : Integer ; vBoForce : boolean = False ) ;
    function  ControleTva         ( vIndex : Integer ) : Integer ;
    procedure SetInfosTva         ( vIndex : integer ; vRegime, vTva, vTpf : string ) ;
    procedure GetInfosTva         ( vIndex : integer ; var vStRegime, vStTva, vStTpf : string ) ;
    procedure FinaliseRegimeEtTva  ( vIndex : Integer = -1 ) ;

    // Affectation infos à la ligne
    Procedure AffecteDateValeur    ( vIndex : Integer ) ;
    procedure AffecteTVA           ( vIndex : Integer ) ;
    procedure InitLibelle          ( vIndex : integer ; vBoForce : boolean = False ) ;
    procedure IncrementRef         ( vIndex : integer ; var vLastRef : string ; vBoPlus : boolean = True) ;

    // Gestion des spécif bordereaux
    procedure GetBornesPiece       ( vNumLigne  : Integer ; var vInDe, vInA : integer ) ;
    procedure GetBornesGroupe       ( vNumGroupe : Integer ; var vInDe, vInA : integer ) ;
    function  GetSoldeGroupe       ( vNumGroupe : Integer; vBoDev : Boolean = True ) : Double ;
    Procedure PutValeurGroupe      ( vNumGroupe : integer ; vChamps : String ; vValeur : Variant ) ;
    function  GetNumGroupe         ( vIndex : Integer ) : integer ;
    function  GetDebutGroupe       ( vIndex : Integer ) : integer ;
    function  LockJournal          ( vBoParle : boolean = false ) : Boolean ;
    function  LockFolio            ( vBoParle : boolean = false ) : Boolean ;
    function  UnLockFolio          ( vBoParle : boolean = false ) : Boolean ;
    function  EstLockFolio         : Boolean ;
    function  EstChampsGroupe      ( vChamps : string ) : boolean ;

    // Gestion de la référence
    function  GetIncRefAuto        ( vStRef : string ; vBoPlus : boolean = True ) : string ;

    // Tests divers lignes / pièce / journal
    Function  IsOut                ( vIndex : Integer ) : Boolean ; // ok
    Function  EstRemplit           ( vIndex : Integer ) : Boolean ; // ok
    Function  EstVentilable        ( vIndex : Integer ; vNumAxe : integer = 0 ) : Boolean ; // ok
    Function  EstTiers             ( vIndex : Integer ) : Boolean ; // ok
    Function  EstHT                ( vIndex : Integer ) : Boolean ; // ok
    Function  EstBQE               ( vIndex : Integer ) : Boolean ; // ok
    Function  EstCli               ( vIndex : Integer ) : Boolean ; // ok
    Function  EstFou               ( vIndex : Integer ) : Boolean ; // ok
    Function  EstDivLett           ( vIndex : Integer ) : Boolean ; // ok
    Function  EstCollFact          ( vIndex : Integer ) : Boolean ; // ok
    Function  EstJalEffet                               : Boolean ;
    Function  EstJalBqe                                 : Boolean ;
    Function  EstJalFact                                : Boolean ;
    Function  EstTvaLoc                                 : Boolean ;
    Function  UnSeulTiers          ( vIndex : integer = 0 ) : Boolean ;
    Function  UnSeulTTC            ( vIndex : integer = 0 ) : Boolean ;
    Function  GetTiers             ( vIndex : integer ) : string ;

    // Pour les Formules
    Procedure ExecuteFormule       ( vIndex : integer ; vChamps,vFormule : string ; vFromTob : Tob = nil ) ;
    Function  CGetFormule          ( vIndex  : Integer ; vStFormule : hString) : Variant ;
    Function  GetFormule           ( vStFormule : HString ) : Variant ;

    // GEstion de l'analytique
    procedure SetVentilDefautOff           ;
    procedure SetVentilSurAttente          ;
    procedure AffecteVentilPourComplement  ( vIndex : Integer ) ;
    Procedure PutVentil                    ( vIndex : Integer ; vChamps : String ; vValeur : Variant ) ;
    procedure SynchroMonoSection           ( vIndex : integer ) ;

    // Aide à la saisie
    Function  QuelChampsMontant            ( vIndex : Integer ) : String ;

    // =============================
    // === Gestion des échéances ===
    // =============================
    // TModeEche    = ( meAucun, meMono, meMulti, meDeporte ) ;
    // --> paramétrage
    procedure SetMultiEcheOff     ;
    procedure SetMultiEcheAucun   ;
    procedure SetMultiEcheMulti   ;
    procedure SetMultiEcheDeporte ;
    // --> Indicateurs
    Function  EstMultiEche          ( vIndex : Integer ) : Boolean ;     // plus d'1 echéance ?
    Function  EstOkMultiEche        ( vIndex : Integer ) : Boolean ;     // multi-echéances possible ?
    Function  EstAvecEche           ( vIndex : Integer ) : Boolean ;  // gestion des échéances sur la ligne
    // --> Recup objet de gestion
    Function  GetMultiEche          ( vIndex : Integer ) : TMultiEche ; overload ;
    Function  GetMultiEche          ( vTobEcr   : Tob )     : TMultiEche ; overload ;
    // --> Synchronisation Ecr --> Eche
    Procedure PutEche               ( vIndex : Integer ; vChamps : string ; vValeur : Variant ) ;
    // --> Recup 1ère ligne d'échéance
    function  GetFirstEche          ( vIndex : integer ) : Tob ;
    Function  GetNbEche             ( vIndex : Integer ) : integer ;
    // Calcul
    procedure CalculEche            ( vindex : integer ; vStModeRegle : string = '' ) ;


    // T_MODEREGL <==> TOB
    Function  GetModeRegl           ( vIndex : Integer ) : T_MODEREGL ;
    procedure SetModeRegl           ( vIndex : Integer ; vModeRegl : T_MODEREGL ) ;

    // Gestion de l'accelerateur de saisie
    procedure ActiveAcc             ;
    function  IsActiveAcc           : boolean ;
    function  GetCompteAcc          ( vIndex : integer ; vStGenDefaut : string = '') : string ;
    procedure InsereLigneHT         ( vIndexTTC : integer ; vStCptHT : string; vMontantHT : double = 0 );
    Function  InsereLigneTVA        ( vIndexHT  : Integer ; vStRegime : String ; vBoTpf : Boolean ) : Boolean ; overload ;
    Function  InsereLigneTVA        ( vIndexHT  : integer ; vIndexTTC : integer )                   : Boolean ; overload ;

    // Scénario
    procedure SetScenarioOn         ;
    procedure SetScenarioOff        ;

    // Activation mode optimisé

    // Gestion du multi-établissement
    Function  EstMultiEtab          : Boolean ;
    Function  IsValidEtab           ( vEtab     : String )  : Boolean ;
    function  GetPieceEtab          ( vEtabDest : String )  : TPieceCompta ;
    procedure TraiteMultiEtab       ( vIndex    : Integer ) ;
    procedure AjustePieceMultiEtab  ;

    // Gestion du Cutoff
    function  GetTobCompl           ( vIndex : Integer ) : TOB ;
    procedure SynchroniseTobCompl   ;
    procedure ReinitCutOff          ( vIndex : Integer ) ;

    // Gestion des Tiers Payeurs
    procedure ActiveTP              ;                                         // permet la gestin des TP
    function  TestParamTP           ( vBoMess : boolean = false ) : boolean ; // retourne vrai si paramètre TP ok
    function  GetFirstIdxTP         : integer ;                               // Retourne l'idx 1ère ligne de tiers suceptible d'être traitée
{$IFDEF COMPTA}
    procedure ChargeTP              ;  // charge dans un TStringList les réf des pièces TP associées
    procedure GenerePiecesTP        ;  // Création des pièces TP associées
    procedure SupprimePiecesTP      ;  // Suppression des pièces TP associées
{$ENDIF COMPTA}

    // ouverture GEP
    procedure CreateFromTOB         ( vTob  : TOB ) ;
    procedure RecopieInfoFromTob    ( vNumLigne : Integer ; vTob  : TOB ) ;

    // Gestion SCANGED
    function  RechGuidId            ( vIndex : Integer ) : string ;
{$IFDEF SCANGED}
    procedure SupprimeLeDocGuid     ( vIndex : Integer ) ;
    procedure AjouteGuidId          ( vIndex : Integer ; vGuidId : string ) ;
{$ENDIF SCANGED}

    // Gestion des BAP // JP 05/06/06 : Est-ce une pièce succeptible d'être rattachée à un BAP
    function IsPieceABap            : Boolean;

    procedure Debug ;

  end;

  // JP 19/09/06 : Tob gérant une liste de TInfosEcriture
  TobPieceCompta = class(Tob)
  private
    lInfoEcriture : TStringList;

    procedure   LibereLaListe (Detruire : Boolean);
  public
    procedure   ClearDetailPC;
    procedure   AddInfoEcriture(NomBase : string; Inf : TInfoEcriture);
    function    GetInfoEcriture(NomBase : string) : TInfoEcriture;
    function    CreateInfoEcr  (NomBase : string) : TInfoEcriture;
    constructor Create(LeNomTable : string; LeParent : TOB; IndiceFils : Integer); override;
    destructor  Destroy; override;
  end;

Function  CGetModeSaisie ( vStJal : String ;  vInfoEcr : TInfoEcriture ) : TModeSaisie ;
Procedure CRemplirMessageTVA   ( vListe : TStringList )                                  ;
Function  CEstCollFact         ( vStCompte : String )                        : Boolean   ;
Procedure CTransfertPiece      ( vPiece : TPieceCompta ; var vTob : TOB )                ;
Function  CChargeLiensReglt    ( vPiece : TPieceCompta ; var vTob : TOB )    : Boolean   ;
function  CCreerLiensSoc       ( vPieceLocal, vPieceDossier : TPieceCompta ; vStGene, vStLib : string ) : Tob ;
procedure CopyInfosOpti        ( vTobNew, vTobSource : Tob ) ;
Function  CCreerMultiEche      ( vParent : Tob ; vPiece : TPieceCompta ; vTobEcr : Tob ; vAction : TActionFiche ; vBoAvecInit : Boolean = False ; vModeFonc : TModeFonc = mfPiece ) : TMultiEche ;



const // === Liste des erreur pour la gestion de la TVA ===
      // Pour le contrôle de TVA
      TVA_PASERREUR              = -1 ;       // ok
      TVA_ERRNATJAL              =  0 ;       // nature du journal Vente ou Achat
      TVA_ERRNOAUXI              =  1 ;       // au - 1 ligne avec un compte auxiliaire
      TVA_ERRNTIERS              =  2 ;       // 1 seule ligne de tiers
      TVA_ERRNOHT                =  3 ;       // pas de ligne de HT
      TVA_WNGVENTIL              =  4 ;       // ventilation sur compte d'attente
      TVA_ERRPARAMCPTHT          =  5 ;       // comptes HT incompatibles
      TVA_ERRNOREGIME            =  6 ;       // Régime non renseigné
      TVA_WNGCPTTVATPF           =  7 ;       // comptes TVA / TPF non renseignés
      // Pour le contrôle de TVA
      TVA_ERRLIGNEHT             =  8 ;       // uniquement pour les lignes de HT
      TVA_ERRNOTIERS             =  9 ;       // Au - 1 lignes de tiers
      TVA_ERRCPTTVATPF           =  10 ;      // comptes TVA / TPF non renseignés
      TVA_ERRTVAINCORRECT        =  11 ;      // TVA incorrect


implementation

uses
  SysUtils,      // uppercase
  {$IFDEF EAGLCLIENT}
  {$ELSE}
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  {$IFDEF VER150}
    Variants,
  {$ENDIF}
  Formule,       // GFormule
  Hctrls,        // ReadTokenSt
  ParamSoc,      //
  {$IFNDEF EAGLSERVER}
  HMsgBox,       // PgiBox
  uTobDebug,     // TobDebug
  {$ENDIF EAGLSERVER}
  ed_tools,      // videliste

{$IFDEF SCANGED}
     UtilGed,
{$ENDIF}
  UtilPGI,       // Resolution
{$IFDEF COMPTA}
  TiersPayeur,   // Gestion des tiers payeurs
{$IFNDEF CMPGIS35}
  uLibBonAPayer,
{$ENDIF}
{$ENDIF COMPTA}
  {$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  CPVersion,
  {$ENDIF MODENT1}

  Constantes,    // QUALIFTRESO
  UProcGen,      // DisposeListe
  uLibAnalytique,// ChargeAnalytique
  UtilSais;      // MajSoldesEcritureTOB


Const

  // Champs à synchroniser Ecr -> Echéance
  _InMaxChampsSynchEche = 37;
  _recChampsSynchEche : array[1.._InMaxChampsSynchEche] of string =
       ('E_REFINTERNE'  ,'E_LIBELLE'         ,'E_REFLIBRE'       ,'E_REFEXTERNE'    ,'E_DATEREFEXTERNE'
       ,'E_AFFAIRE'     ,'E_QTE1'            ,'E_QTE2'           ,'E_QUALIFQTE1'    ,'E_QUALIFQTE2'
       ,'E_LIBRETEXTE0' ,'E_LIBRETEXTE1'     ,'E_LIBRETEXTE2'    ,'E_LIBRETEXTE3'   ,'E_LIBRETEXTE4'
       ,'E_LIBRETEXTE5' ,'E_LIBRETEXTE6'     ,'E_LIBRETEXTE7'    ,'E_LIBRETEXTE8'   ,'E_LIBRETEXTE9'
       ,'E_TABLE0'      ,'E_TABLE1'          ,'E_TABLE2'         ,'E_TABLE3'        ,'E_LIBREDATE'
       ,'E_LIBREBOOL0'  ,'E_LIBREBOOL1'      ,'E_CONSO'          ,'E_LIBREMONTANT0' ,'E_LIBREMONTANT1'
       ,'E_LIBREMONTANT2' ,'E_LIBREMONTANT3' ,'E_BANQUEPREVI'    ,'E_ETABLISSEMENT'
       ,'E_DEBITDEV'      ,'E_CREDITDEV'     ,'E_RIB'
       ) ;
  _InMaxChampsOpti = 48 ;
  _recChampsOpti : array[1.._InMaxChampsOpti] of string =
       ( 'T_AUXILIAIRE'    ,'T_LIBELLE'      ,'T_FERME'       ,'T_COLLECTIF'    ,'T_SOUMISTPF'
        ,'T_CONFIDENTIEL'  ,'T_NATUREAUXI'   ,'T_LETTRABLE'   ,'T_REGIMETVA'    ,'T_TVAENCAISSEMENT'
        ,'T_PAYEUR'        ,'T_ISPAYEUR'     ,'T_CONSO'       ,'T_ESCOMPTE'     ,'T_MODEREGLE'
        ,'T_MULTIDEVISE'   ,'T_DEVISE'       ,'T_TOTALDEBIT'  ,'T_TOTALCREDIT'  ,'T_TOTDEBE'
        ,'T_TOTDEBS'       ,'T_TOTDEBP'      ,'T_TOTCREE'     ,'T_TOTCRES'      ,'T_TOTCREP'
        ,'G_GENERAL'       ,'G_LIBELLE'      ,'G_FERME'       ,'G_SENS'         ,'G_NATUREGENE'
        ,'G_VENTILABLE'    ,'G_VENTILABLE1'  ,'G_VENTILABLE2' ,'G_VENTILABLE3'  ,'G_VENTILABLE4'
        ,'G_VENTILABLE5'   ,'G_CONFIDENTIEL' ,'G_POINTABLE'   ,'G_LETTRABLE'    ,'G_COLLECTIF'
        ,'G_BUDGENE'       ,'G_QUALIFQTE1'   ,'G_QUALIFQTE2'  ,'G_SOUMISTPF'    ,'G_REGIMETVA'
        ,'G_TVA'           ,'G_TPF'          ,'G_TVASURENCAISS'
        );



Function  CGetModeSaisie ( vStJal : String ;  vInfoEcr : TInfoEcriture ) : TModeSaisie ;
begin
  result := msPiece ;
  if not vInfoEcr.LoadJournal( vStJal ) then Exit ;

  if vInfoEcr.Journal.GetValue('J_MODESAISIE') = 'BOR'
    then result := msBor
  else if vInfoEcr.Journal.GetValue('J_MODESAISIE') = 'LIB'
    then result := msLibre ;

end ;


Function  CCreerMultiEche ( vParent : Tob ; vPiece : TPieceCompta ; vTobEcr : Tob ; vAction : TActionFiche ; vBoAvecInit : Boolean = False ; vModeFonc : TModeFonc = mfPiece ) : TMultiEche ;
begin
  result := TMultiEche.Create( vParent, vPiece, vTobEcr, vAction, vModeFonc ) ;

  if vBoAvecInit then
    result.CalculEche ;

//  if result.NbEche = 0 then
//    result.NewEche ;

end ;

procedure CopyInfosOpti ( vTobNew, vTobSource : Tob ) ;
var i : integer ;
    lstChp : string ;
begin
  for i:=1 to _InMaxChampsOpti do
    begin
    lstChp := _recChampsOpti[i] ;
    if vTobSource.GetNumChamp( lstChp ) >= 0 then
      vTobNew.AddChampSupValeur( lStChp, vTobSource.GetValue( lStChp ) ) ;
    end ;
end ;

Procedure CRemplirMessageTVA( vListe : TStringList ) ;
begin
  vListe.Clear;
  vListe.Add('0;?caption?;Génération de TVA impossible : la pièce et le journal doivent être de nature Vente ou Achat;E;O;O;O;');
  vListe.Add('1;?caption?;Génération de TVA impossible : la pièce doit posséder une ligne sur un compte auxiliaire;E;O;O;O;');
  vListe.Add('2;?caption?;Génération de TVA impossible : la pièce ne doit posséder qu''une seule ligne de Tiers;E;O;O;O;');
  vListe.Add('3;?caption?;Génération de TVA impossible : la pièce ne possède pas de ligne sur un compte HT;E;O;O;O;');
  vListe.Add('4;?caption?;Les lignes générées sont ventilées sur les sections d''attente;A;O;O;O;');
  vListe.Add('5;?caption?;Génération de TVA incohérente : certains comptes HT sont incompatibles avec le paramétrage de TVA;E;O;O;O;');
  vListe.Add('6;?caption?;Génération de TVA impossible : le régime n''est pas renseigné pour le tiers ;E;O;O;O;');
  vListe.Add('7;?caption?;Génération de TVA incomplète : certains comptes de TVA ou TPF ne sont pas renseignés;E;O;O;O;');
  vListe.Add('8;?caption?;Contrôle de TVA impossible : vous devez être positionné sur une ligne avec un compte HT;E;O;O;O;');
  vListe.Add('9;?caption?;Contrôle de TVA impossible : la pièce doit posséder une ligne sur un compte de Tiers;E;O;O;O;');
  vListe.Add('10;?caption?;Contrôle de TVA impossible : certains comptes de TVA ou TPF ne sont pas renseignés;E;O;O;O;');
  vListe.Add('11;?caption?;La TVA de la pièce est incorrecte. Confirmez-vous la validation de cette pièce ?;Q;YNC;N;N;');
end ;

Function  CEstCollFact      ( vStCompte : String ) : Boolean ;
var lStListe  : String ;
    lStClient : String ;
begin
  Result := False ;

  if vStCompte = '' then Exit ;
  if Not VH^.OuiTvaEnc then Exit ;

  lStListe  := VH^.CollCliEnc ;

  if lStListe<>'' then
    begin

    if lStListe[ Length(lStListe) ] <> ';' then
      lStListe := lStListe + ';' ;

    repeat
      lStClient := ReadTokenSt( lStListe ) ;
      if lStClient = '' then Break ;
      if Copy( vStCompte, 1, Length( lStClient ) ) = lStListe then
        BEGIN
        result := True ;
        Exit ;
        END ;
    until ( lStListe='' ) ;

    end ;

  if not Result then
    begin
    lStListe := VH^.CollFouEnc ;
    if lStListe <> '' then
      begin

      if lStListe[ Length(lStListe) ]<> ';' then
        lStListe := lStListe + ';' ;

      repeat
        lStClient := ReadTokenSt( lStListe ) ;
        if lStClient = '' then Break ;
        if Copy( vStCompte, 1, Length(lStClient) ) = lStClient then
          begin
          result:=True ;
          Exit ;
          end ;
      until ( lStListe = '' ) ;

      end ;
   end ;

end ;


function CCreerLiensSoc( vPieceLocal, vPieceDossier : TPieceCompta ; vStGene, vStLib : string ) : Tob ;
begin
  result := TOB.Create('CLIENSPIECES', nil, -1) ;

  // référence pièce émettrice
  result.PutValue('CLP_SOCIETE',        vPieceLocal.GetEntete('E_SOCIETE') ) ;
  result.PutValue('CLP_ETABLISSEMENT',  vPieceLocal.GetEntete('E_ETABLISSEMENT') ) ;
  result.PutValue('CLP_DOSSIER',        vPieceLocal.Dossier ) ;
  result.PutValue('CLP_JOURNAL',        vPieceLocal.GetEntete('E_JOURNAL') ) ;
  result.PutValue('CLP_NUMEROPIECE',    vPieceLocal.GetEntete('E_NUMEROPIECE') ) ;
  result.PutValue('CLP_EXERCICE',       vPieceLocal.GetEntete('E_EXERCICE') ) ;
  result.PutValue('CLP_DATECOMPTABLE',  vPieceLocal.GetEntete('E_DATECOMPTABLE') ) ;

  // Compte général + réf.
  result.PutValue('CLP_GENERAL',        vStGene ) ; // ;
  result.PutValue('CLP_REFINTERNE',     vStLib )  ; // ;

  // Référence pièce dossier
  result.PutValue('CLP_SOCIETEDEST',    vPieceDossier.GetEntete('E_SOCIETE') ) ;
  result.PutValue('CLP_ETABDEST',       vPieceDossier.GetEntete('E_ETABLISSEMENT') ) ;
  result.PutValue('CLP_DOSSIERDEST',    vPieceDossier.Dossier ) ;
  result.PutValue('CLP_JOURNALDEST',    vPieceDossier.GetEntete('E_JOURNAL') ) ;
  result.PutValue('CLP_NUMPIECEDEST',   vPieceDossier.GetEntete('E_NUMEROPIECE') ) ;
  result.PutValue('CLP_EXERCICEDEST',   vPieceDossier.GetEntete('E_EXERCICE') ) ;
  result.PutValue('CLP_DATEDEST',       vPieceDossier.GetEntete('E_DATECOMPTABLE') ) ;

end ;

procedure CTransfertPiece( vPiece : TPieceCompta ; var vTob : TOB ) ;
var lNumLigne  : Integer ;
    lNumEche   : Integer ;
    lTobLigne  : Tob ;
    lMultiEche : TMultiEche ;
begin

  vTob.ClearDetail ;

  for lNumLigne := vPiece.count downto 1 do
    begin
    lTobLigne := vPiece.GetTob( lNumLigne ) ;

    // Traitement du multi-échéance
    if vPiece.EstMultiEche( lNumLigne ) then
      begin
      // tob des lignes d'échéances
      lMultiEche  := vPiece.GetMultiEche( lNumLigne ) ;
      for lNumEche := lMultiEche.NbEche downto 1 do
        lMultiEche.Detail[ lNumEche - 1 ].ChangeParent( vTob, 0 ) ;
      end
    else
      lTobLigne.ChangeParent( vTob, 1 ) ;

    end ;

end ;


procedure _CherchePiecePrincipale( var vSoc, vEtab, vJal : string ; var vNumP : integer ; var vDateC : TDateTime ) ;
var lStSQL : string ;
    lQRech : TQuery ;
begin
  lStSQL := 'SELECT CLP_SOCIETE, CLP_ETABLISSEMENT,CLP_JOURNAL,CLP_NUMEROPIECE,CLP_DATECOMPTABLE '
                   + 'FROM CLIENSPIECES WHERE (CLP_SOCIETEDEST = "'   + vSoc + '"'
                                       + ' AND CLP_ETABDEST = "' + vEtab + '"'
                                       + ' AND CLP_JOURNALDEST = "'       + vJal + '"'
                                       + ' AND CLP_NUMPIECEDEST = '    + IntToStr( vNumP )
                                       + ' AND CLP_DATEDEST = "' + USDateTime( vDateC ) + '")' ;
  lQRech := OpenSQL( lStSQL, True ) ;
  if not lQRech.Eof then
    begin
    vSoc   := lQRech.FindField('CLP_SOCIETE').AsString ;
    vEtab  := lQRech.FindField('CLP_ETABLISSEMENT').AsString ;
    vJal   := lQRech.FindField('CLP_JOURNAL').AsString ;
    vNumP  := lQRech.FindField('CLP_NUMEROPIECE').AsInteger ;
    vDateC := lQRech.FindField('CLP_DATECOMPTABLE').AsDateTime ;
    _CherchePiecePrincipale ( vSoc, vEtab, vJal, vNumP, vDateC ) ;
    end ;
  Ferme( lQRech ) ;

end ;

procedure _ChargeDetailLP( vSoc, vEtab, vJal : string ; vNumP : integer ; vDateC : TDateTime ; var vTobListe : Tob ) ;
var lStSQL : string ;
begin
  lStSQL := 'SELECT * FROM CLIENSPIECES WHERE CLP_SOCIETE = "' + vSoc + '"'
                                       + ' AND CLP_ETABLISSEMENT = "' + vEtab + '"'
                                       + ' AND CLP_JOURNAL = "' + vJal + '"'
                                       + ' AND CLP_NUMEROPIECE = '    + IntToStr( vNumP )
                                       + ' AND CLP_DATECOMPTABLE = "' + USDateTime( vDateC ) + '"' ;

  vTobListe.LoadDetailDBFromSQL( 'CLIENSPIECES', lStSQL, True ) ;
end ;


Function CChargeLiensReglt ( vPiece : TPieceCompta ; var vTob : TOB ) : Boolean ;
var lStSQL     : string ;
    lStSoc     : string ;
    lStEtab    : string ;
    lStJal     : string ;
    lInNumP    : Integer ;
    lDtDate    : TDateTime ;
    i          : integer ;
    lTobLP     : Tob ;
begin
{
CLP_SOCIETE, CLP_ETABLISSEMENT, CLP_DOSSIER, CLP_JOURNAL, CLP_NUMEROPIECE, CLP_EXERCICE,
CLP_DATECOMPTABLE, CLP_REFINTERNE, CLP_GENERAL,
CLP_EXERCICEDEST, , CLP_DOSSIERDEST, CLP_ETABDEST,
clé : CLP_SOCIETE,CLP_ETABLISSEMENT,CLP_JOURNAL,CLP_NUMEROPIECE,CLP_SOCIETEDEST,CLP_ETABDEST
}
  result := False ;
  if vPiece.Count = 0 then Exit ;

  if vTob.Detail.count > 0 then
    vTob.ClearDetail ;

  // uniquement en mode pièce
  if vPiece.ModeSaisie <> msPiece then Exit ;

  // Recherche des objets CLIENSPIECES
  lStSoc    := vPiece.GetEnteteS('E_SOCIETE') ;
  lStEtab   := vPiece.GetEnteteS('E_ETABLISSEMENT') ;
  lStJal    := vPiece.GetEnteteS('E_JOURNAL') ;
  lInNumP   := vPiece.GetEnteteI('E_NUMEROPIECE') ;
  lDtDate   := vPiece.GetEnteteDt('E_DATECOMPTABLE') ;

  lStSQL := 'SELECT * FROM CLIENSPIECES WHERE ' ;
  lStSQL := lStSQL + ' (CLP_SOCIETE = "' + lStSoc + '" AND CLP_ETABLISSEMENT = "' + lStEtab + '" AND CLP_JOURNAL = "' + lStJal + '"'
                   + ' AND CLP_NUMEROPIECE = '    + IntToStr( lInNumP )
                   + ' AND CLP_DATECOMPTABLE = "' + USDateTime( lDtDate ) + '")' ;
  lStSQL := lStSQL + 'OR (CLP_SOCIETEDEST = "' + lStSoc + '" AND CLP_ETABDEST = "' + lStEtab + '" AND CLP_JOURNALDEST = "' + lStJal + '"'
                   + ' AND CLP_NUMPIECEDEST = '    + IntToStr( lInNumP )
                   + ' AND CLP_DATEDEST = "' + USDateTime( lDtDate ) + '")' ;

  // La pièce fait-elle partie d'un lot rglt multi-soc / multi-étab
  if not ExisteSQL( lStSQL ) then Exit ;

  // Détermination pièce principale du lot
  _CherchePiecePrincipale( lStSoc, lStEtab, lStJal, lInNumP, lDtDate ) ;

  // Chargement des éléments LiensPieces 1er niveau
  _ChargeDetailLP( lStSoc, lStEtab, lStJal, lInNumP, lDtDate, vTob ) ;

  // Chargement des éléments LiensPieces 2ème niveau
  for i := vTob.Detail.count - 1 downto 0  do
    begin
    lTobLP := vTob.Detail[ i ] ;
    lStSoc    := lTobLP.GetString('CLP_SOCIETEDEST') ;
    lStEtab   := lTobLP.GetString('CLP_ETABDEST') ;
    lStJal    := lTobLP.GetString('CLP_JOURNALDEST') ;
    lInNumP   := lTobLP.GetInteger('CLP_NUMPIECEDEST') ;
    lDtDate   := lTobLP.GetDateTime('CLP_DATEDEST') ;
    _ChargeDetailLP( lStSoc, lStEtab, lStJal, lInNumP, lDtDate, vTob ) ;
    end ;

  result := vTob.Detail.Count > 0 ;

end ;

{ TPieceCompta }


function TPieceCompta.GetValue(vIndex: Integer; vChamps: String): Variant;
begin
  result := #0 ;
  if IsOut( vIndex ) then Exit ;
  result := Detail[ vIndex - 1 ].GetValue( vChamps ) ;
end;

Procedure TPieceCompta.PutValue(vIndex: Integer; vChamps: String; vValeur: Variant ; vBoForce : boolean ) ;
var lOldValue : Variant ;
    lTobLigne : TOB ;
    lInErr    : Integer ;
begin

  // Récup pointeur ligne concernée
  if IsOut( vIndex ) then Exit ;
  lTobLigne := CGetTob( vIndex ) ;

  // Si champ supplémentaire, on sort
  if lTobLigne.GetNumChamp( vChamps ) >= 1000 then
    begin
    PutValueSpecif(vIndex, vChamps, vValeur ) ;
    Exit ;
    end ;

  // Récup ancienne valeur et comparaison
  lOldValue := lTobLigne.GetValue( vChamps ) ;
  if not vBoForce and (lOldValue = vValeur) then Exit ;

  // Affectation de la nouvelle valeur
  lOldValue := lTobLigne.GetValue( vChamps ) ;

  // === MONTANT DEBIT ===
  if (vChamps = 'E_CREDITDEV') then
    begin
    lInErr := CIsValidMontant( vValeur ) ;
    if lInErr <> RC_PASERREUR then
      begin
      NotifyError( lInErr, '' ) ;
      Exit ;
      end ;
    CSetMontants( lTobLigne, 0, VarAsType( vValeur, varDouble) , GetRDevise( vIndex ), True ) ;
    end
  // === MONTANT CREDIT ===
  else if (vChamps = 'E_DEBITDEV') then
    begin
    lInErr := CIsValidMontant( vValeur ) ;
    if lInErr <> RC_PASERREUR then
      begin
      NotifyError( lInErr, '' ) ;
      Exit ;
      end ;
    CSetMontants( lTobLigne, VarAsType( vValeur, varDouble), 0, GetRDevise( vIndex ), True ) ;
    end
  // DATE d'échéance
  else if vChamps = 'E_DATEECHEANCE' then
    begin
    if not NbJoursOK( lTobLigne.GetDateTime('E_DATECOMPTABLE'), vValeur) then
      begin
      NotifyError( RC_DATEECHEINCORRECTE, '' ) ;
      Exit ;
      end ;
    lTobLigne.PutValue( vChamps, vValeur ) ;
    end
  // === Mode de paiement ===
  else if (vChamps = 'E_MODEPAIE') then
    begin
    if ( lTobLigne.GetString('E_ECHE')='X' ) and ( Trim(vValeur) = '' ) then
      begin
      NotifyError( RC_MODEPAIEINCORRECT, '' ) ;
      Exit ;
      end ;
    lTobLigne.PutValue( vChamps, vValeur ) ;
    end
  // === AUTRE CHAMPS ===
  else
    begin
    lTobLigne.PutValue( vChamps, vValeur ) ;
    end ;

  // Appel de l'évènement sur changement du champ pour traitement annexe
  if OnChangeField( vIndex, vChamps, lOldValue, vValeur )
    then begin
         if EstRemplit( vIndex ) then
           FBoModif := True ;
         PutVentil  ( vIndex, vChamps, vValeur ) ;
         PutEche    ( vIndex, vChamps, vValeur ) ;
         // BOR : Gestion des champs géré par groupe
         if (ModeSaisie=msBOR) then
           begin
           if EstChampsGroupe( vChamps ) then
             PutValeurGroupe( GetValue( vIndex, 'E_NUMGROUPEECR'), vChamps, vValeur ) ;
//           if (vChamps = 'E_DEBITDEV') or (vChamps = 'E_CREDITDEV') then
//             ReCalculGroupe ;
           end ;
         {JP 12/06/05 : Gestion des BAP : on détruira le BAP si l'on a modifié un autre champ
                        que E_LIBELLE et E_REFINTERNE}
         if (vChamps <> 'E_LIBELLE') and (vChamps <> 'E_REFINTERNE') then ConserverBap := False;
         end
    else lTobLigne.PutValue( vChamps, lOldValue ) ;

end;


destructor TPieceCompta.Destroy;
begin

  if Assigned(FEntete)     then
    FreeAndNil(FEntete) ;

  if Assigned(FEcheances)  then
    FreeAndNil(FEcheances) ;

  if Assigned(FTobOrigine) then
    FreeAndnil(FTobOrigine) ;

  if Assigned(FTobChancell) then
    FreeAndnil(FTobChancell) ;

  if Assigned(FPiecesEtab)  then
    LibereListe( FPiecesEtab, True ) ;

  if Assigned(FEcrCompl) then
    FreeAndNil(FEcrCompl) ;

  if assigned( FContexte ) then
    FreeAndNil( FContexte ) ;

  if Assigned( FListeTP ) then
    FreeAndNil( FlisteTP ) ;

  inherited;
    
end;


function TPieceCompta.EstLigneModifiable( vIndex : Integer ): Boolean;
var lTobEcr   : TOB ;
begin

  Result := True ;

  // Récupération TOb sur ligne courante :
  lTobEcr := GetTob( vIndex ) ;
  if lTobEcr = nil then Exit ;

  // ==================================
  // La ligne n'est pas modifiable si :
  // ==================================

  //  - une de ses échéances est totalement ou partiellement lettré
  result := ( ExisteLettrage( vIndex ) )
  //  - elle est pointée
         or ( lTobEcr.GetString('E_REFPOINTAGE')<>'' )
  //  - le E_JOURNAL est fermé
         or ( Info.GetString('J_FERME')='X' )
  //  - l'auxiliaire est fermé
         or ( Info.GetString('T_FERME')='X' )
  //  - le général est fermé
         or ( Info.GetString('G_FERME')='X' )
  //  - elle a subit une édition de TVA
         or ( lTobEcr.GetString('E_EDITEETATTVA')='#' )
  //  - une immo est présente
         or ( lTobEcr.GetString('E_IMMO')<>'' )
  //  - le compte général est visé
         or ( Info.GetString('G_VISAREVISION')='X')
  // JP 01/08/05 : FQ 15124 : On ne peut modifier les écritures TTC de Gescom
         or ((lTobEcr.GetString('E_REFGESCOM') <> '') and (lTobEcr.GetString('E_TYPEMVT') = 'TTC'))
  // JP 26/06/07 : FQ TRESO 10491 : on vérouille les flux originaires de la Tréso
         or ( lTobEcr.GetString('E_QUALIFORIGINE') = QUALIFTRESO ) ;

  // Donc la ligne est modifiable si...
  result := not result ;

end;

function TPieceCompta.ExisteLettrage( vIndex : Integer ): Boolean;
var lTobEcr   : TOB ;
begin

  Result := False ;

  // Récupération TOb sur ligne courante :
  lTobEcr := GetTob( vIndex ) ;
  if lTobEcr = nil then Exit ;

  // test de la ligne principale
  result :=   ( lTobEcr.GetValue('E_ETATLETTRAGE')='PL' )
           or ( lTobEcr.GetValue('E_ETATLETTRAGE')='TL' ) ;

  // Manque le test des autres échéances si ligne en multi-échéances
  // ...
end;

procedure TPieceCompta.LoadInfo( vIndex : Integer );
begin
  if IsOut( vIndex ) then Exit ;
  Info_LoadCompte( Detail[ vIndex - 1 ].GetString('E_GENERAL') ) ;
  Info_LoadAux( Detail[ vIndex - 1 ].GetString('E_AUXILIAIRE') ) ;
end;

function TPieceCompta.GetTobLigne( vNumLigne : Integer ; vNumEche : Integer ) : TOB ;
var i : integer ;
begin
  Result := nil ;
  if isOut( vNumLigne ) then Exit ;

  if FModeEche <> meMulti then
    begin
    // En mono échéance, index = numligne-1
    result := Detail[ vNumLigne - 1 ] ;
    end
  else
    begin
    for i := (vNumLigne - 1) to (Detail.Count - 1) do
      begin
      if Detail[ i ].GetInteger('E_NUMLIGNE')>vNumLigne then Break ;

      if Detail[ i ].GetInteger('E_NUMLIGNE')=vNumLigne then
        if (vNumEche <> -1) then
          begin
          // recherche d'une échéance
          if Detail[ i ].GetInteger('E_NUMECHE')=vNumEche then
            begin
            result := Detail[ i ] ;
            break ;
            end ;
          end
        else
          // uniquement la ligne
          begin
          result := Detail[ i ] ;
          break ;
          end ;
      end ;
    end ;

  if result <> nil then
    begin
    Info_LoadCompte( result.GetString('E_GENERAL') ) ;
    Info_LoadAux( result.GetString('E_AUXILIAIRE') ) ;
    end ;

end;


function TPieceCompta.IsValidPiece : Boolean ;
var  lTError : TRecError ;
     lStEtab : String ;
     lBoCPOk : Boolean ;
     lInIdx  : Integer ;
     lTobEcr : Tob     ;  // FQ 22974 29/05/2008 SBO
begin

  // Suppression des éventuelles lignes vides
  CSupprimeLigneVide ;

  // Affectation de la Tva
  FinaliseRegimeEtTva ;

  // ==> Début Modif // FQ 22974 29/05/2008 SBO
  lBoCPOk := True ;
  // Finalisation des lignes
  for lInIdx := 1 to Detail.Count do
    begin
    lTobEcr := CGetTob( lInIdx ) ;
    // --- Test pour Affectation des contreparties
    lBoCPOk := lBOCPOk and ( lTobEcr.GetString( 'E_CONTREPARTIEGEN' ) <> '' ) ;
    // --- Remise à 0 du E_NUMGROUPEECR pour le mode pièce // FQ 22974 27/05/2008 SBO
    if ModeSaisie=msPiece then
      lTobEcr.PutValue('E_NUMGROUPEECR', 0)
    else if ModeSaisie = msLibre then
      lTobEcr.PutValue('E_NUMGROUPEECR', 1) ;
    end ;
  // ==> Fin Modif // FQ 22974 29/05/2008 SBO
  if not lBoCPOk then
    CAffectCompteContrePartie( self, Info ) ;

  // Vérification standards
  lTError    := CIsValidSaisiePiece( self, Info ) ;
  result     := (lTError.RC_Error = RC_PASERREUR) ;
  if not result then
    NotifyError( lTError.RC_Error, lTError.RC_Message, lTError.RC_Methode ) ;
{  if (lTError.RC_Error <> RC_PASERREUR) then
    OnError( nil, lTError) ;
}
  // Vérification multi-établissements
  if result and EstMultiEtab then
    begin
    // Test journaux
    if (EstJalEffet or EstJalBqe) and (Contexte.JalMultiEtab = '' ) then  // FQ17766 SBO 01/05/2006
      begin
      result := False ;
      NotifyError( 0, 'Le journal de substitution des pièces sans contrepartie n''est pas renseigné !' ) ;
      end
    else
      begin
      // test établissement
      lStEtab := GetEnteteS('E_ETABLISSEMENT') ;
      // Au moins une ligne sur l'établissement d'origine
      result := FindFirst( ['E_ETABLISSEMENT'], [ lStEtab ], False ) <> nil ;
      if not result then
        NotifyError( 0, 'La pièce doit contenir au moins une ligne sur l''établissement principal ' ) ;
      end ;
    end ;
end;

procedure TPieceCompta.LoadFromSQL( vStCondSQL : String = '') ;
var lQEcr    : TQuery ;
    lStWhere : String ;
    lStChps  : String ;
begin
  // Init variables
  videPiece ;

  // Chargement depuis la base
  if Trim(vStCondSQL) = '' then
    begin
    if ( GetEnteteS('E_JOURNAL')='' ) or (GetEnteteI('E_NUMEROPIECE')=0) then
      begin
      FAction  := taCreat ;
      Exit ;
      end ;
    lStWhere := GetWhereSQL
    end
  else lStWhere := vStCondSQL ;

  // SBO 05/07/2007 : gestion pb DB2 avec champ blocnote à positionner à la fin
  lStChps := GetSelectAll('E', True ) + ', E_BLOCNOTE' ;
  if not Contexte.local then
    begin
    lQEcr := OpenSelect( 'SELECT ' + lStChps + ' FROM ECRITURE WHERE ' + lStWhere, Dossier ) ;
    if not lQEcr.Eof then
      LoadDetailDB( 'ECRITURE', '', '', lQEcr, False ) ;
    Ferme( lQEcr ) ;
    end
  else
    LoadDetailDBFromSQL( 'ECRITURE', 'SELECT ' + lStChps + ' FROM ECRITURE WHERE ' + lStWhere ) ;

  // Chargement de la base donc modification !
  if Detail.Count > 0
    then FAction  := taModif
    else FAction  := taCreat ;

  // Chargement de l'analytique
  if Action = taModif then
    begin
    // tri
    Detail.Sort('E_NUMLIGNE;E_NUMECHE');

    ChargeVentil ;

    FTobOrigine.Dupliquer( self, True, True ) ;

    // Chargement des infos d'entete
    ChargeEntete ;

    // Chargement du TInfoEcriture
    ChargeInfo ;

    // Infos des tiers payeurs
  {$IFDEF COMPTA}
    ChargeTP;
  {$ENDIF COMPTA}

    // Chargement du multi-écheance
    if FModeEche = meDeporte then
      ChargeEche ;

    // Champs supp de gestion
    ChargeChampsSup ;

    // Gestion du cutoff
    ChargeExtension; {Lek 07/0406}

    // JP 12/05/06 : S'il y a un BAP validé définitivement, la pièce sera en lecture seule
    GestionDesBAP;

    end ;

  FBoModif := False ;

end;

procedure TPieceCompta.LoadFromTOB( vTobLignes : TOB  ; vInfo : TInfoEcriture = nil );
var i : Integer ;
begin
  videPiece ;

  // Récupération des filles de la TOB paramètre
  for i:=(vTobLignes.detail.count - 1) downto 0 do
    vTobLignes.detail[ i ].ChangeParent( self, -1 ) ;

  // Copie SAV pour traitement postérieur
  FTobOrigine.ClearDetail ;
  FTobOrigine.Dupliquer( self, True, True ) ;

  // Chargement des infos d'entete
  ChargeEntete ;

  // Chargement du TInfoEcriture
  if Assigned( vInfo ) then
    begin
    if Assigned( Info ) then Info.Free ;
    Info := vInfo ;
    end ;
  ChargeInfo ;

  // Chargement du multi-écheance
  ChargeEche ;

  // Champs supp de gestion
  ChargeChampsSup ;

  // Gestion du cutoff
  ChargeExtension; {Lek 07/0406}

  // Chargement de la base donc modification !
  FAction  := taModif ;
  FBoModif := False ;

end;

procedure TPieceCompta.ChargeInfo ;
begin
  Info_LoadJournal(    GetEnteteS('E_JOURNAL')           ) ; // Journal
  Info.Etabliss.load(  [ GetEnteteS('E_ETABLISSEMENT') ] ) ; // Etablissement
  Info_LoadDevise(     GetEnteteS('E_DEVISE')            ) ; // Devise
end;

{***********A.G.L.***********************************************
Auteur  ...... : Compta
Créé le ...... : 07/12/2004
Modifié le ... :   /  /
Description .. : Ajoute une ligne d'écriture dans la pièce à l'index passé en
Suite ........ : paramètre (à la fin si l'index est erroné)
Mots clefs ... :
*****************************************************************}
function TPieceCompta.NewRecord ( vIndex : Integer ) : TOB ;
var lBoARenumeroter : Boolean ;
    lMultiEche      : TMultiEche ;
    lInNumEche      : integer ;
    lTob            : Tob ;
begin

  if IsOut( vIndex ) then
    vIndex := Detail.count + 1
  else
    begin
    lTob := CGetTob( vIndex ) ;
    lInNumEche := lTob.GetInteger('E_NUMECHE') ;
    // est-on en mode multi-eche ?
    if (FModeEche = meMulti) and (lInNumEche > 1) then
      begin
      lMultiEche := GetMultiEche( vIndex ) ;
      result     := lMultiEche.NewEche( lInNumEche ) ;
      CurIdx     := vIndex ;
      Exit ;
      end ;
    end ;

  lBoARenumeroter := ( Detail.Count > 0 ) and ( vIndex <= Detail.Count ) ;

  result := Tob.Create('ECRITURE', self, vIndex - 1 ) ;
  result.InitValeurs;
  CPutDefautEcr(Result);
  InitNewRecord(Result, vIndex, lBoARenumeroter);

  CurIdx := vIndex ;

end;

procedure TPieceCompta.DupliquerPiece(PieceSrc: TPieceCompta);
var
  i:       Integer;
begin
  {b FP 02/05/2006}
  InitPiece(PieceSrc.GetRMVT);
  for i:=0 to PieceSrc.Count-1 do
    DupliquerLigne(Count+1, PieceSrc, i+1);
  RenumeroteLignes;
  {e FP 02/05/2006}
end;

procedure TPieceCompta.DupliquerLigne(vNumLigneDst: Integer; PieceSrc: TPieceCompta; vNumLigneSrc: Integer);
var
  TobLigneDst:   TOB;
  TobLigneSrc:   TOB;
begin
  {b FP 02/05/2006}
  if IsOut(vNumLigneDst) then
    begin
    vNumLigneDst := Detail.Count + 1;
    TobLigneDst  := Tob.Create('ECRITURE', Self, vNumLigneDst-1);
    end
  else
    begin
    TobLigneDst  := GetTob(vNumLigneDst);
    TobLigneDst.ClearDetail;
    end;

  TobLigneSrc := PieceSrc.GetTob(vNumLigneSrc);
  TobLigneDst.Dupliquer(TobLigneSrc, True, True);
  InitNewRecord(TobLigneDst, vNumLigneDst, True);
  PutValue(vNumLigneDst, 'E_DATEECHEANCE', TobLigneSrc.GetValue('E_DATEECHEANCE'));
  AffecteEche(vNumLigneDst);

  // Champs supp de gestion
  ChargeChampsSup ;

  // Gestion du cutoff
  if Assigned(CGetTOBCompl(TobLigneSrc)) then
    begin
    TobLigneDst.AddChampSupValeur('CUTOFF','X');
    CCreateTOBCompl(TobLigneDst, FEcrCompl);
    CGetTOBCompl(TobLigneDst).Dupliquer(CGetTOBCompl(TobLigneSrc), True, True);
    CMAJTOBCompl(TobLigneDst);
    end;

  CSynchroVentil(TobLigneDst);

  // Chargement de la base donc modification !
  FAction  := taModif ;
  FBoModif := False ;
  CurIdx   := vNumLigneDst ;
  {e FP 02/05/2006}
end;

function TPieceCompta.GetWhereSQL: String;
begin

  if ModeSaisie=mspiece then
    begin
    Result := 'E_JOURNAL="'         + GetEnteteS('E_JOURNAL')                      + '"' +
           ' AND E_EXERCICE="'      + GetEnteteS('E_EXERCICE')                     + '"' +
           ' AND E_DATECOMPTABLE="' + UsDateTime( GetEnteteDt('E_DATECOMPTABLE') ) + '"' +
           ' AND E_NUMEROPIECE='    + GetEnteteS('E_NUMEROPIECE')                  +
           ' AND E_QUALIFPIECE="'   + GetEnteteS('E_QUALIFPIECE')                  + '"';
    end
  else
    begin
    Result := 'E_JOURNAL="'         + GetEnteteS('E_JOURNAL')              + '"' +
           ' AND E_EXERCICE="'      + GetEnteteS('E_EXERCICE')             + '"' +
           ' AND E_PERIODE='        + GetEnteteS('E_PERIODE')              +
           ' AND E_NUMEROPIECE='    + GetEnteteS('E_NUMEROPIECE')          +
           ' AND E_QUALIFPIECE="'   + GetEnteteS('E_QUALIFPIECE')          + '"';
    end ;

//  result := WhereEcritureTOB( tsGene, self, False ) ;
end;

procedure TPieceCompta.PutValueAll(vChamps: String; vValeur: Variant);
var i : Integer ;
begin

  for i := 1 to Detail.Count do
    PutValue( i, vChamps, vValeur ) ;

end;

function TPieceCompta.IsValidLigne ( vIndex : Integer ) : Boolean ;
var  lTError : TRecError ;
     lTob    : TOB ;
begin
  result := False ;

  lTob    := GetTob( vIndex ) ;
  if lTob = nil then Exit ;

  if IsValidDebitCredit( vIndex ) then
    begin
    // SI Echéance non renseigné
    if not ExisteEcheance( vIndex ) then
       AffecteEche( vIndex ) ;
    end ;

  lTError := CIsValidLigneSaisie( lTob, Info ) ;

  result     := lTError.RC_Error = RC_PASERREUR ;

  if (lTError.RC_Error <> RC_PASERREUR) then
    NotifyError( lTError.RC_Error, '' ) ;

end;

function TPieceCompta.GetSolde( vBoDev : Boolean ) : Double;
begin
  result := GetSoldePartiel( 1, Detail.Count, vBoDev ) ;
end;

function TPieceCompta.GetTotalCredit( vBoDev : Boolean ) : Double;
var i : Integer ;
    lStChp : string ;
begin
  result := 0 ;
  if vBoDev
    then lStChp := 'E_CREDITDEV'
    else lStChp := 'E_CREDIT' ;
  for i := 0 to Detail.Count - 1 do
    result := result + Detail[i].GetDouble(lStChp) ;
  if vBoDev
    then result := Arrondi( result  , Devise.Decimale )   // FQ 17129 : Pb d'arrondi Delphi
    else result := Arrondi( result  , V_PGI.OkDecV ) ;
end;

function TPieceCompta.GetTotalDebit( vBoDev : Boolean ) : Double;
var i : Integer ;
    lStChp : string ;
begin
  result := 0 ;
  if vBoDev
    then lStChp := 'E_DEBITDEV'
    else lStChp := 'E_DEBIT' ;
  for i := 0 to Detail.Count - 1 do
    result := result + Detail[i].GetDouble(lStChp) ;
  if vBoDev
    then result := Arrondi( result , Devise.Decimale )   // FQ 17129 : Pb d'arrondi Delphi
    else result := Arrondi( result , V_PGI.OkDecV ) ;
end;
{
procedure TPieceCompta.CompleteEcriture( vIndex: Integer );
var lTob : TOB ;
begin

  if IsOut( vIndex ) then Exit ;

  lTob := GetTob( vIndex ) ;

  // Renseignement de base sur les échéances
  if not ExisteEcheance( vIndex ) then
    AffecteEche( vIndex ) ;

  // info régime TVA
  AffecteTVA( vIndex ) ;

  // Code conso
  AffecteConso ( vIndex ) ;

  // le RIB
  if GetValue( vIndex, 'E_RIB' ) = '' then
    AffecteRIB ( vIndex ) ;

  // Confidentialité
  AffecteConf ( vIndex ) ;

  // Info Ana
  if ( lTob.GetValue('E_ANA') = 'X' ) and ( lTob.Detail.count = 0 ) then
    AffecteAna ( vIndex ) ;

end;
}
function TPieceCompta.OnChangeField( vIndex :Integer ; vChamps : String ; vOldValue, vNewValue : Variant ) : Boolean ;
var lTob         : TOB ;
    lStCollectif : String ;
    lStVal       : String ;
begin

  result := True ;
//  if vOldValue = vNewValue then Exit ; // A tester dans les fonctions appelantes

  lTob := CGetTob( vIndex ) ;

  // Pas d'evt sur les échéances > 1
  if lTob.GetInteger('E_NUMECHE') > 1 then Exit ;

  // ========================================
  // ==== Modification du compte général ====
  // ========================================
  if vChamps = 'E_GENERAL' then
    begin
    // Pour provoquer recalcul
    if vOldValue <> '' then
      begin
      PutValue( vIndex, 'E_CONTREPARTIEGEN', '' ) ;
      PutValue( vIndex, 'E_CONTREPARTIEAUX', '' ) ;
      // On vire l'auxiliaire si le compte n'est pas collectif mais qu'un auxiliaire est présent !
      if ( Compte_GetValue('G_COLLECTIF') <> 'X' ) and
         ( GetValue( vIndex, 'E_AUXILIAIRE') <> '' )
        then begin
             CPutValue( vIndex, 'E_AUXILIAIRE', '' ) ;
             putValue( vIndex, 'E_RIB', '' ) ;
             ReinitEche( vIndex ) ;
             end ;
      ReinitCutOff( vIndex ) ;
      end ;

    if vNewValue = '' then Exit ;

    result := Info_LoadCompte( vNewValue ) ;
    if result then
      begin

      // E_CONSO
      AffecteConso( vIndex ) ;
      // E_REGIMETVA, E_TVA, E_TPF
      AffecteTVA( vIndex ) ;
      // E_CONFIDENTIEL
      AffecteConf( vIndex ) ;
      // Changement de compte général --> Suppression de l'ancienne ventilation
      if lTob.GetValue( 'E_ANA' ) = 'X' then lTob.ClearDetail ;
      AffecteAna( vIndex ) ;
      // E_RIB
      AffecteRIB ( vIndex ) ;
      // maj echéance (pour les TIC / TID )
      AffecteEche( vIndex ) ;
      // Type de mvt
      AffecteTypeMvt( vIndex ) ;
      // G_BUDGENE
      lTob.PutValue('E_BUDGET', Compte_GetValue('G_BUDGENE') ) ;
      // Infos Tva Enc
      AffecteTvaEnc( vIndex ) ;
      // Qualifiant de quantité

      lStVal := lTob.GetValue('E_QUALIFQTE1') ;
      if ((lStVal='') or (lStVal='...')) then
        lTob.PutValue('E_QUALIFQTE1', Compte_GetValue('G_QUALIFQTE1') ) ;
      lStVal := lTob.GetValue('E_QUALIFQTE2') ;
      if ((lStVal='') or (lStVal='...')) then
        lTob.PutValue('E_QUALIFQTE2', Compte_GetValue('G_QUALIFQTE2') ) ;
      end
    end
  // ===========================================
  // ==== Modification du compte auxiliaire ====
  // ===========================================
  else if vChamps = 'E_AUXILIAIRE' then
    begin
    // Pour provoquer recalcul
    if vOldValue <> '' then
      begin
      PutValue( vIndex, 'E_CONTREPARTIEGEN', '' ) ;
      PutValue( vIndex, 'E_CONTREPARTIEAUX', '' ) ;
      putValue( vIndex, 'E_RIB', '' ) ;
      ReinitEche( vIndex ) ;
      end ;

    if vNewValue = '' then Exit ;

    result := Info_LoadAux( lTob.GetString('E_AUXILIAIRE') ) ;
    if result then
      begin

      // collectif si général non renseigné
      if ( lTob.GetValue('E_GENERAL')='' ) then
        begin
        lStCollectif := Aux_GetValue('T_COLLECTIF') ;
        lTob.PutValue( 'E_GENERAL', lStCollectif ) ;
        // Appel de l'évènement sur changement du champ pour traitement annexe
        if not OnChangeField( vIndex, 'E_GENERAL', '', lStCollectif )
          then
            begin
            lTob.PutValue( 'E_GENERAL', '' ) ;
            result := False ;
            Exit ;
            end ;
        end ;

      // E_REGIMETVA, E_TVA, E_TPF
      AffecteTVA( vIndex ) ;
      // E_CONFIDENTIEL
      AffecteConf( vIndex ) ;
      // Echéances
      AffecteEche( vIndex ) ;
      // E_RIB
      AffecteRIB ( vIndex ) ;
      // Type de mvt
      AffecteTypeMvt( vIndex ) ;

      end ;
    end
  // ===========================================
  // ==== Modification du compte auxiliaire ====
  // ===========================================
  else if vChamps = 'E_DEBITDEV' then
    begin
    // Si analytique non encore affecté alors on le fait maintenant
    AffecteAna( vIndex ) ;
    // MAJ des échéances
    AffecteEche( vIndex ) ;
    // Recalcul du tableau des montants par tva
    CalculTvaEnc( vIndex ) ;
    end
  // ===========================================
  // ==== Modification du compte auxiliaire ====
  // ===========================================
  else if vChamps = 'E_CREDITDEV' then
    begin
    // Si analytique non encore affecté alors on le fait maintenant
    AffecteAna( vIndex ) ;
    // MAJ des échéances
    AffecteEche( vIndex ) ;
    // Recalcul du tableau des montants par tva
    CalculTvaEnc( vIndex ) ;
    end
  // ==============================
  // ==== Autres Modifications ====
  // ==============================
  else if vChamps = 'E_DATECOMPTABLE' then
    begin
    lTob.PutValue( 'E_EXERCICE',      Contexte.GetExoDT( vNewValue ) ) ;
    lTob.PutValue( 'E_PERIODE',       GetPeriode ( vNewValue ) ) ;
    lTob.PutValue( 'E_SEMAINE',       NumSemaine ( vNewValue ) ) ;
    if lTob.GetString('E_DEVISE') <> V_PGI.DevisePivot then
      majDeviseTaux( vIndex ) ;
    end
  else if vChamps = 'E_NATUREPIECE' then
    begin
    // Type de mvt
    AffecteTypeMvt( vIndex ) ;
    end
  else if vChamps = 'E_ETABLISSEMENT' then
    begin
    // Test validité établissement
    result := IsValidEtab( vNewValue ) ;
    if not result then
      NotifyError( 0, 'Le paramétrage multi-établissement n''autorise pas cette saisie' ) ;
    end
  else if vChamps = 'E_BANQUEPREVI' then
    begin
    // Test validité établissement
    if vNewValue<>'' then
      begin
      result := Info.LoadCompte( vNewValue ) and
                ( Info.GetString('G_NATUREGENE')='BQE' ) and
                ( Info.GetString('G_FERME')='-' ) ;
      if not result then
        NotifyError( 0, 'Le compte de banque prévisionnel renseigné n''est pas correct' ) ;
      end
    end
  else
    begin
    end ;

end;

procedure TPieceCompta.ChargeEntete ( vTob : TOB = nil ) ;
var lTobLigne : TOB ;
begin
  if Detail.count = 0 then Exit ;

  if vTob = nil
    then lTobLigne := Detail[0]
    else lTobLigne := vTob ;

  FEntete.putValue('E_JOURNAL'        , lTobLigne.GetValue('E_JOURNAL')  ) ;
  FEntete.putValue('E_NATUREPIECE'    , lTobLigne.GetValue('E_NATUREPIECE')  ) ;
  FEntete.putValue('E_DATECOMPTABLE'  , lTobLigne.GetValue('E_DATECOMPTABLE')  ) ;
  FEntete.putValue('E_DEVISE'         , lTobLigne.GetValue('E_DEVISE')  ) ;
  if not EstMonoDevise // cas Bordereau multi-devise
    then FEntete.putValue('E_DEVISE'  , V_PGI.DevisePivot ) ;
  FEntete.putValue('E_ETABLISSEMENT'  , lTobLigne.GetValue('E_ETABLISSEMENT')  ) ;
  FEntete.putValue('E_NUMEROPIECE'    , lTobLigne.GetValue('E_NUMEROPIECE')  ) ;
  FEntete.putValue('E_EXERCICE'       , lTobLigne.GetValue('E_EXERCICE')  ) ;
  FEntete.putValue('E_QUALIFPIECE'    , lTobLigne.GetValue('E_QUALIFPIECE')  ) ;
  FEntete.putValue('E_MODESAISIE'     , lTobLigne.GetValue('E_MODESAISIE')  ) ;
  FEntete.putValue('E_ECRANOUVEAU'    , lTobLigne.GetValue('E_ECRANOUVEAU')  ) ;
  FEntete.putValue('E_SOCIETE'        , lTobLigne.GetValue('E_SOCIETE')  ) ;
  FEntete.putValue('E_PERIODE'        , lTobLigne.GetValue('E_PERIODE')  ) ;

  Info.LoadJournal( lTobLigne.GetString('E_JOURNAL')  ) ;
  AffecteModeSaisie ;
  if ModeSaisie in [msBor, msLibre] then
    begin
    FEntete.putValue('E_NATUREPIECE',     Info.Journal.NatureParDefaut ) ;
    FEntete.putValue('E_DATECOMPTABLE'  , FinDeMois( lTobLigne.GetDateTime('E_DATECOMPTABLE') )  ) ;
    end ;

end;

procedure TPieceCompta.PutEntete(vChamps: String; vValeur: Variant);
begin

//  if vValeur <> FEntete.GetValue( vChamps ) then Exit ;

  // Affectation de la tob d'entête
  FEntete.PutValue ( vChamps, vValeur ) ;
  // Affectation des lignes
  PutValueAll     ( vChamps, vValeur ) ;

  // Affectation du RDevise
  if vChamps = 'E_DEVISE' then
    Info_LoadDevise( vValeur )

  // Affectation de l'exercice pour la date
  else  if vChamps = 'E_DATECOMPTABLE' then
    begin
    PutEntete( 'E_EXERCICE', Contexte.GetExoDT( VarAsType(vValeur, vardate) ) ) ;
    PutEntete( 'E_PERIODE',  GetPeriode ( VarAsType(vValeur, vardate) )       ) ;
    PutEntete( 'E_SEMAINE',  NumSemaine ( VarAsType(vValeur, vardate) )       ) ;
    end

  // gestion du journal
  else if vChamps = 'E_JOURNAL' then
    begin
    Info.LoadJournal( vValeur ) ;
    AffecteModeSaisie ;
    Case ModeSaisie of
      msBor   : begin
                PutEntete   ('E_NATUREPIECE', Info.Journal.NatureParDefaut ) ;
                PutEntete   ('E_DEVISE',      V_PGI.DevisePivot ) ;
                RecalculGroupe (1);
                end ;
      msLibre : begin
                PutEntete   ('E_NATUREPIECE',   Info.Journal.NatureParDefaut ) ;
                PutEntete   ('E_DEVISE',        V_PGI.DevisePivot ) ;
                PutValueAll ( 'E_NUMGROUPEECR', 1 ) ;
                end ;
      else begin
           PutValueAll ( 'E_NUMGROUPEECR', 0 ) ;
           end ;
      end ;
    end ;

end;

function TPieceCompta.GetEntete(vChamps: String): Variant;
begin
  result := FEntete.GetValue( vChamps ) ;
end;


function TPieceCompta._Save : Boolean ;
var lTobLigne  : TOB ;
    lMultiEche : TMultiEche ;
    lNumLigne  : Integer ;
    lNumPiece  : Integer ;
    lNumEche   : Integer ;
    lPieceEtab : TPieceCompta ;
    lTobLiens  : Tob ;
    lStCpt     : String ;
    lStLib     : string ;
begin

  result := False ;

  try

    // =================================================
    // === Suppression de l'enregistrement si besoin ===
    // =================================================
    if Action <> taCreat then
      DetruitPiece( True ) ;

    // =================================
    // === Enregistrement à la ligne ===
    // =================================
    if Contexte.local then
      begin
      for lNumLigne := 1 to Detail.count do
        begin
        lTobLigne := Detail[ lNumLigne - 1 ] ;
        OnUpdateEcritureTOB( lTobLigne, Action ,[], Info) ;  // [cEcrCompl]
        end ;
      // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
      if Contexte.InsertSpecif
        then CTobInsertDB( self, True )
        else self.InsertDBByNivel(True) ;

      if FEcrCompl.Detail.count > 0 then
        begin
        SynchroniseTobCompl ;
        // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
        if Contexte.InsertSpecif
          then CTobInsertDB( FEcrCompl, True )
          else FEcrCompl.InsertDBByNivel(True) ;
        end ;
      end
    else
      for lNumLigne := 1 to Detail.count do
        begin
        lTobLigne := Detail[ lNumLigne - 1 ] ;

        // Traitement du multi-échéance
        if (FModeEche = meDeporte) and EstMultiEche( lNumLigne ) then
          begin
          // tob des lignes d'échéances
          lMultiEche  := GetMultiEche( lNumLigne ) ;
          lMultiEche.SynchroniseEcheances ;
          // Maj annexes
          for lNumEche := 0 to ( lMultiEche.Detail.Count - 1 ) do
            begin
            OnUpdateEcritureTOB( lMultiEche.Detail[ lNumEche ] , taCreat , [], Info) ;
            if not Contexte.InsertDB( lMultiEche.Detail[ lNumEche ] )
              then Exit ;
            end ;

          end
        else
        // Ligne MonoEchéance ou Autre
          begin
          OnUpdateEcritureTOB( lTobLigne, Action ,[cEcrCompl], Info) ;
          if not Contexte.InsertDB( lTobLigne )
            then Exit ;
          end ;

      end ;

    // ======================
    // === MAJ DES SOLDES ===
    // ======================
    CalculDesSoldes ;

    // Pb maj des soldes
    if V_PGI.IoError <> oeOk then
      Exit ;

    // ===================================
    // === Gestion Multi-Etablissement ===
    // ===================================
    // Enregistrement des pièces sur les autres établissements
    if ( Action = taCreat ) and ( FPiecesEtab.Count > 0 ) then
      For lNumPiece := 0 to ( FPiecesEtab.Count - 1 ) do
        begin
        // sauvegarde de la pièce
        lPieceEtab := TPieceCompta( FPiecesEtab.Objects[ lNumPiece ] ) ;
        lPieceEtab.Save ;

        // récup infos
        if Contexte.GetLiensEtab( lPieceEtab.GetEnteteS('E_ETABLISSEMENT') ) <> nil
          then lStCpt := Contexte.GetLiensEtab( lPieceEtab.GetEnteteS('E_ETABLISSEMENT') ).GetString('CLE_COMPTEORIG')
          else lStCpt := '' ;
        lStLib := Copy('Liaison étab. (' + Contexte.GetLibelleEtab( GetEnteteS('E_ETABLISSEMENT') ) + ')' , 1, 35)  ;

        // Céation lien
        lTobLiens := CCreerLiensSoc( self, lPieceEtab, lStCpt, lStLib ) ;
        lTobLiens.InsertDB( nil ) ;
        lTobLiens.Free ;

        end ;

    // === Tout s'est bien passé ===
    Result := True ;

  finally

  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Compta
Créé le ...... : 20/12/2004
Modifié le ... :   /  /
Description .. : Enregistre la pièce en base.
Suite ........ : Retourne un False et renseigne le TRecError en cas de
Suite ........ : problème.
Suite ........ : Retourne True et passe la pièce en mode modification si
Suite ........ : aucun problème.
Mots clefs ... :
*****************************************************************}
function TPieceCompta.Save: Boolean;
var lNumLigne  : Integer ;
    lPieceEtab : TPieceCompta ;
begin

 result        := false ;
 V_PGI.IoError := oeOk ;

 try

   beginTrans ;

   setAllModifie( True ) ;

   // Génération des pièces multi-établissements
   if ( Action = taCreat ) and EstMultiEtab then
     begin
     // On traite lignes par ligne le multi-établissement
     lNumLigne := 1 ;
     While not isOut( lNumLigne ) do
       begin
//     For lNumLigne := 1 to Detail.Count do
       TraiteMultiEtab( lNumLigne ) ;
       Inc( lNumLigne ) ;
       end ;
     // Finalisation des pièces
     For lNumLigne := 0 to ( FPiecesEtab.Count - 1 ) do
       begin
       lPieceEtab := TPieceCompta( FPiecesEtab.Objects[ lNumLigne ] ) ;
       lPieceEtab.AjustePieceMultiEtab ; // Pièces générées
       end ;
     AjustePieceMultiEtab ; // Pièce principale
     end ;

   // Finalisation de la pièce
   TermineLaPiece ;

   // Enregistrement
   result := _Save ;

 finally
    if result
       // passage mode modif si ok
      then begin
           commitTrans ;
           FAction := taModif ;
           FBoModif := False ;
           end
      // Notification d'erreur sinon
      else begin
           rollback ;
           NotifyError( RC_BADWRITE, '' ) ;
           end ;

  end ;

end;



procedure TPieceCompta.DetruitPiece( vBoPourMAJ : boolean ) ;
var lBoModePiece  : boolean ;
    lTOB          : TOB ;
    i             : integer ;
begin

  if Action = taCreat then Exit ;
  if ( FTobOrigine = nil ) or ( FTobOrigine.Detail = nil ) or ( FTobOrigine.Detail.Count = 0 ) then exit ;

  // Traitement de pré-suppression
  if not vBoPourMAJ then
    for i := 0 to FTobOrigine.Detail.Count-1 do
      OnDeleteEcritureTob( FTobOrigine.Detail[i], taCreat, [cEcrCompl] ) ;

  // suppression des écritures
  if CDetruitAncienPiece( FTobOrigine, Dossier ) then
    begin
    lTOB := Detail[0] ;

    // suppression de l'analytique
    CDetruitAncienAnaPiece( FTobOrigine, Dossier ) ;

    // suppression des BAP
    {$IFDEF COMPTA}
    lBoModePiece := (lTOB.GetString('E_MODESAISIE')='-') or (lTOB.GetString('E_MODESAISIE')='') ;
{$IFNDEF CMPGIS35}
    if lBoModePiece and GestionBapOk and not ConserverBap then
      DetruitBAP( FTobOrigine, Dossier );
{$ENDIF}
    {$ENDIF COMPTA}

    if vBoPourMAJ and contexte.Local then
      begin
      // suppression des ECRCOMPL ==> PRIS EN COMPTE PAR OnDeleteEcritureTOB si mode suppression seule
      //                          ==> Effectué par le onupdateEcritureTob si pas en local
      CSupprimerEcrCompl( lTOB ) ;
      end
    else
      // Maj des solde des comptes uniquement si suppression seule, fait lors de insertion sinon
      MajSoldesEcritureTOB ( FTobOrigine , False , Info ) ;

    end ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Compta
Créé le ...... : 20/12/2004
Modifié le ... :   /  /
Description .. : Détermine un numéro de pièce provisoire à la pièce
Mots clefs ... :
*****************************************************************}
procedure TPieceCompta.AttribNumeroTemp;
var lStMM       : String17 ;    // masquenum non utilisé
    lDtDateC    : TDateTime ;   // date comptable saisie
    lInNumero   : Integer ;     // Numero de la pièce provisoire
    lStCompteur : string ;      // Code du compteur a utilisé
begin

  if Action <> taCreat then Exit ;
//  if (GetEnteteI('E_NUMEROPIECE') <> 0) then Exit ;

  if ( GetEnteteS('E_JOURNAL') = '' ) or
     not Info_LoadJournal( GetEnteteS('E_JOURNAL') ) then
  begin
  PutEntete( 'E_NUMEROPIECE', 0 ) ;
  Exit ;
  end ;

  // Numéro temporaire en mode pièce
  if ModeSaisie=msPiece then
    begin
    lDtDateC := GetEnteteDt( 'E_DATECOMPTABLE' ) ;
    if ( GetEnteteS('E_QUALIFPIECE') <> 'N' ) and ( GetEnteteS('E_QUALIFPIECE') <> 'I' ) // Modif IFRS
      then lStCompteur := Journal_GetValue('J_COMPTEURSIMUL')
      else lStCompteur := Journal_GetValue('J_COMPTEURNORMAL') ;
    lInNumero := GetNum( EcrGen, lStCompteur, lStMM, lDtDateC, Dossier) ;
    PutEntete( 'E_NUMEROPIECE', lInNumero ) ;
    end
  // Numéro temporaire en mode Bor / Libre
  else
//    AttribNumeroDef
    ;


end;

procedure TPieceCompta.AttribSolde( vIndex : Integer ) ;
var lSolde  : Double ;
    lNumL   : Integer ;
    lInDec  : integer ;
begin

  if isOut( vIndex ) then Exit ;

  // Solde de la pièce sans la ligne courante
  lSolde := 0 ;
  lInDec := GetRDevise(vIndex).Decimale ;
  for lNumL := 1 to Detail.count do
    if lNumL <> vIndex then
      lSolde := lSolde + ( GetValue( lNumL, 'E_DEBITDEV') - GetValue( lNumL, 'E_CREDITDEV') ) ;
  lSolde := Arrondi( lSolde, lInDec ) ; // FQ 17129 : Pb d'arrondi Delphi

  // si la ligne courante contient déjà la solde, on sort
  if lSolde = Arrondi( GetValue( vIndex, 'E_CREDITDEV')-GetValue( vIndex, 'E_DEBITDEV'), lInDec )
    then Exit ;

  if lSolde = 0 then
    begin
    if GetValue( vIndex, 'E_CREDITDEV') <> 0
      then PutValue( vIndex, 'E_CREDITDEV',  0 )
      else PutValue( vIndex, 'E_DEBITDEV',   0 )

    end
  else if lSolde > 0
    then PutValue( vIndex, 'E_CREDITDEV',  lSolde )
    else PutValue( vIndex, 'E_DEBITDEV',   -1 * lSolde ) ;

  VerifieEquilibre( vIndex ) ;

//  if ( ModeSaisie = msBor) and ( vIndex < Detail.Count )  then
//    RecalculGroupe ;

end;

Function TPieceCompta.GenereTVA ( vIndex : integer ) : Integer ;
Var lTobAux   : Tob ;
    lTobHT    : Tob ;
    lTobTva   : Tob ;
    lBoTpf    : Boolean ;
    lStCompte : String ;
    lNumLigne : Integer ;
    lStRegime : String ;
    lStTva    : String ;
    lStTpf    : String ;
    lIdxRef   : integer ;
begin

  result := TVA_PASERREUR ;
  if ModeSaisie <> msPiece then Exit ;

  // Uniquement pour les vente et achats
  if ( Journal_GetValue('J_NATUREJAL') <> 'ACH' ) and
     ( Journal_GetValue('J_NATUREJAL') <> 'VTE' ) then
    begin
    result := TVA_ERRNATJAL ;
    Exit ;
    end ;

  // Contrôles de présence 1 seul TTC et au moins 1 HT
  lTobAux := TrouveLigneTiers( vIndex ) ;
  if lTobAux = nil then
    begin
    result := TVA_ERRNOAUXI ;
    exit ;
    end ;
  if not UnSeulTiers( vIndex ) then
    begin
    result := TVA_ERRNTIERS ;
    exit ;
    end ;
  lTobHT := TrouveLigneHT( vIndex ) ;
  if lTobHT = nil then
    begin
    result := TVA_ERRNOHT ;
    exit ;
    end ;

  // Récup info base
  lBoTpf    := False ;
  lStRegime := lTobAux.GetString('E_REGIMETVA') ;
  if lStRegime = '' then
    AffecteTVA( lTobAux.GetIndex+1 ) ;
  lStCompte := lTobAux.GetString('E_AUXILIAIRE') ;
  if lStCompte <> '' then
    begin
    if Info_LoadAux( lStCompte ) then
      lBoTpf := Info.GetString('T_SOUMISTPF') = 'X' ;
    end
  else
    begin
    lStCompte := lTobAux.GetValue('E_GENERAL') ;
    if Info_LoadCompte( lStCompte ) then
      lBoTpf    := Info.GetString('G_SOUMISTPF') = 'X' ;
    end ;

  // Destruction des lignes TVA/TPF concernées déja présentes
  repeat
    lTobTva := TrouveLigneTVA( vIndex ) ;
    if lTobTva<>nil then
      DeleteRecord( lTobTva.GetIndex + 1) ;
  Until ( lTobTva = nil ) ;

  // Génération des lignes TVA/TPF
  lIdxRef := -1 ;
  for lNumLigne := Detail.Count downto 1 do
    if EstHT( lNumLigne ) then
      begin
      lTobHT := GetTob( lNumLigne ) ;
      lTobHT.PutValue( 'E_REGIMETVA', lStRegime ) ;
      if InsereLigneTVA( lNumLigne, lStRegime, lBoTpf ) then
        begin
        if (lIdxRef < 0) then
          begin
          lIdxRef := CurIdx ;
          lStTva  := GetString( lIdxRef, 'E_TVA' ) ;
          lStTpf  := GetString( lIdxRef, 'E_TPF' ) ;
          end ;
        end
      else
        result := TVA_WNGCPTTVATPF ;
      end ;

  // Renumérotation des lignes
//  RenumeroteLignes ;

  // Regroupement selon scénario
  if ChargeScenario and (GetScenario.GetValue('SC_CONTROLETVA') = 'GTG' ) then
    RegroupeTVA( vIndex ) ;

end;

function TPieceCompta.GetSorteTVA( vIndex : Integer ): TSorteTva;
begin
  result := stvDivers ;

  if not Info_LoadJournal( GetEnteteS('E_JOURNAL') ) then Exit ;

  if ( Journal_GetValue('J_NATUREJAL') = 'VTE' ) and
     ( ( GetString( vIndex, 'E_NATUREPIECE') = 'FC') or ( GetString( vIndex, 'E_NATUREPIECE') ='AC') )
     then result := stvVente
     else if ( Journal_GetValue('J_NATUREJAL') = 'ACH' ) and
             ( (GetString( vIndex, 'E_NATUREPIECE')='FF') or (GetString( vIndex, 'E_NATUREPIECE')='AF') )
          then result := stvAchat ;
end;

function TPieceCompta.EstAchat ( vIndex : Integer ) : Boolean;
begin
  result := False ;
  if (vIndex > 0) and Info_LoadCompte( GetString(vIndex, 'E_GENERAL') )
                  and ( ( Info.GetString('G_NATUREGENE') = 'CHA' ) or
                        ( Info.GetString('G_NATUREGENE') = 'PRO' )
                  ) then result := Info.GetString('G_NATUREGENE') = 'CHA'
//    if
  // Pourquoi ne pas aller voir la nature du compte général (la présence d'un compte de charge/produit dans le groupe)
  else if ( (GetEnteteS('E_NATUREPIECE') = 'FF') or (GetEnteteS('E_NATUREPIECE') = 'AF') or
       (GetEnteteS('E_NATUREPIECE') = 'OF') or (GetEnteteS('E_NATUREPIECE') = 'RF') )
    then result := True
    else if ( Info_LoadJournal( GetEnteteS('E_JOURNAL') ) ) and
            ( Journal_GetValue('J_NATUREJAL') = 'ACH' )
           then result := True ;
end;

function TPieceCompta.EstBQE( vIndex : Integer ) : Boolean;
var lTob : tob ;
begin
  result := false ;

  lTob := GetTob( vIndex ) ;
  if lTob = nil then Exit ;

  if (lTob.GetString('E_GENERAL') <> '') and Info_LoadCompte( lTob.GetString('E_GENERAL') ) then
    if ( Compte_GetValue('G_NATUREGENE') = 'BQE' ) or
       ( Compte_GetValue('G_NATUREGENE') = 'CAI' )
      then result := True ;
end;

function TPieceCompta.EstHT( vIndex : Integer ) : Boolean;
var lTob : Tob ;
begin
  result := False ;

  lTob := GetTob( vIndex ) ;
  if lTob = nil then Exit ;

  if (lTob.GetString('E_GENERAL') <> '') and Info_LoadCompte( lTob.GetString('E_GENERAL') ) then
    if ( Compte_GetValue('G_NATUREGENE') = 'CHA' ) or
       ( Compte_GetValue('G_NATUREGENE') = 'PRO' ) or
       ( Compte_GetValue('G_NATUREGENE') = 'IMO' )
      then result := True ;
end;

function TPieceCompta.EstTiers( vIndex : Integer ) : Boolean;
var lTob : TOB ;
begin
  result := False ;

  lTob := GetTob( vIndex ) ;
  if lTob = nil then Exit ;

  // Auxiliaire ?
  if ( lTob.GetString('E_AUXILIAIRE') <> '' ) and
     Info_LoadAux( lTob.GetString('E_AUXILIAIRE') )
    then result := True ;

  // Général de nature TIC ou TID ?
  if ( lTob.GetString('E_GENERAL') <> '' ) and
     ( Info_LoadCompte( lTob.GetString('E_GENERAL') ) ) and
     ( ( Compte_GetValue('G_NATUREGENE') = 'TID' ) or
       ( Compte_GetValue('G_NATUREGENE') = 'TIC' ) )
    then result := True ;

end;

function TPieceCompta.TrouveLigneCompte ( vStCompte : String ; vFromLigne : Integer ) : Tob ;
var i   : integer ;
    lDe : integer ;
    lA  : integer ;
begin
  result := nil ;
  GetBornesPiece( vFromLigne, lDe, lA ) ;
  if  vFromLigne > lA then Exit ;
  if (vFromLigne > lDe) and (vFromLigne <= lA) then
    lDe := vFromLigne ;
  for i := lA downto lDe do
    begin
    if GetInteger(i, 'E_NUMECHE')>1 then Continue ;
    if GetString(i,'E_GENERAL') = vStCompte then
      begin
      Result := getTob( i ) ;
      Exit ;
      end ;
    end ;
end;

Function TPieceCompta.TrouveIndiceLigneCompte(vStCompteGene, vStCompteAuxi: String; vFromLigne: Integer = 1): Integer;
var i : integer ;
begin
  {b FP 02/05/2006}
  result := -1;
  for i := vFromLigne to detail.count do
    begin
    if GetInteger(i, 'E_NUMECHE')>1 then Continue ;
    if (GetString(i,'E_GENERAL') = vStCompteGene) and
       (GetString(i,'E_AUXILIAIRE') = vStCompteAuxi) then
      begin
      Result := i;
      Exit ;
      end ;
    end ;
  {e FP 02/05/2006}
end;

function TPieceCompta.TrouveLigneHT ( vIndex : integer ; vFromLigne : Integer ) : Tob ;
var i    : integer ;
    lDe  : integer ;
    lA   : integer ;
//    lTob : Tob ;
begin
  result := nil ;

  GetBornesPiece( vIndex, lDe, lA ) ;
  if (vFromLigne > lDe) and (vFromLigne <= lA) then
    lDe := vFromLigne ;

  for i := lDe to lA do
    begin
    if EstHT( i ) then
      begin
      result := getTob( i ) ;
      Exit ;
      end ;
    end ;
end;

function TPieceCompta.TrouveLigneTiers ( vIndex : integer ; vFromLigne : Integer ) : Tob ;
var i   : integer ;
    lDe : integer ;
    lA  : integer ;
begin
  result := nil ;

  GetBornesPiece( vIndex, lDe, lA ) ;
  if (vFromLigne > lDe) and (vFromLigne <= lA) then
    lDe := vFromLigne ;

  for i := lDe to lA do
    begin
    if EstTiers( i ) then
      begin
      Result := getTob( i ) ;
      if result.GetInteger('E_NUMECHE')>1 then
        begin
        result := nil ;
        continue ;
        end ;
      Exit ;
      end ;
    end ;

end;

function TPieceCompta.TrouveLigneTVA ( vIndex : integer ; vFromLigne : Integer ) : Tob ;
var i             : integer ;
//    j             : integer ;
    lTob          : TOB ;
//    lStCompte     : String ;
//    lStCptTva     : String ;
//    lStCptTpf     : String ;
    lDe           : integer ;
    lA            : integer ;
begin
  result := nil ;

  GetBornesPiece( vIndex, lDe, lA ) ;

  if vFromLigne > lA then Exit ;
  
  if (vFromLigne > lDe) and (vFromLigne <= lA) then
    lDe := vFromLigne ;

  for i := lDe to lA do
    begin
    lTob := GetTob( i ) ;
    if lTob.GetInteger('E_NUMECHE')>1 then Continue ;
    if lTob.GetString('E_TYPEMVT')='TVA' then
      begin
      result := lTob ;
      break ;
      end ;
{
    if ( lTob.GetValue('E_GENERAL') <> '') and Info_LoadCompte( lTob.GetString('E_GENERAL') ) then
      begin
      lStcompte := lTob.GetValue('E_GENERAL') ;
      for j := 1 to Detail.Count do
        begin
        // Récupération du compte général
        if not Info_LoadCompte( GetString( j, 'E_GENERAL') ) then Continue ;

        lStCptTva := GetCompteTva( j ) ;
        lStCptTpf := GetCompteTpf( j ) ;

        if ( lStCompte = lStCptTva ) or ( lStCompte = lStCptTpf ) then
           begin
           result := lTob ;
           Exit ;
           end ;

        end ;
      end ;
}
    end ;
end;

function TPieceCompta.GetCompteTpf(vIndex: Integer): String;
begin
  result := '' ;
  if IsOut( vIndex ) then Exit ;
  if GetValue( vIndex, 'E_TVAENCAISSEMENT') = 'X'
    then result := Tpf2Encais ( GetRegimeTva(vIndex), GetValue( vIndex, 'E_TPF' ), EstAchat(vIndex) )
    else result := Tpf2Cpte   ( GetRegimeTva(vIndex), GetValue( vIndex, 'E_TPF' ), EstAchat(vIndex) ) ;
end;

function TPieceCompta.GetCompteTva(vIndex: Integer): String;
begin
  result := '' ;
  if IsOut( vIndex ) then Exit ;
  if GetValue( vIndex, 'E_TVAENCAISSEMENT') = 'X'
    then result := Tva2Encais ( GetRegimeTva(vIndex), GetCodeTva( vIndex ), EstAchat(vIndex) )
    else result := Tva2Cpte   ( GetRegimeTva(vIndex), GetCodeTva( vIndex ), EstAchat(vIndex) ) ;
end;

function TPieceCompta.GetTauxTpf(vIndex: Integer): Double;
begin
  result := 0 ;
  if IsOut( vIndex ) then Exit ;
  result := Tpf2Taux ( GetRegimeTva(vIndex), GetValue( vIndex, 'E_TPF' ), EstAchat(vIndex) ) ;
end;

function TPieceCompta.GetTauxTva(vIndex: Integer): Double;
begin
  result := 0 ;
  if IsOut( vIndex ) then Exit ;
  result := Tva2Taux ( GetRegimeTva(vIndex), GetCodeTva( vIndex ), EstAchat(vIndex) ) ;
end;

procedure TPieceCompta.AffecteTVA(vIndex: Integer);
var lTobEcr   : TOB ;
    lStRegime : string ;
begin

  if not EstRemplit( vIndex ) then Exit ;
  lTobEcr := GetTob( vIndex ) ;

  // Affectation Régime
  if EstTiers( vIndex ) then
    begin
    lStRegime := '' ;
    // finalement régime de l'auxi
    if (lTobEcr.GetString('E_AUXILIAIRE') <> '') and Info.LoadAux( lTobEcr.GetString('E_AUXILIAIRE') ) then
      lStRegime := Info.GetString('T_REGIMETVA')
    // finalement régime du général
    else if ( lTobEcr.GetString('E_GENERAL') <> '' ) and Info.LoadCompte( lTobEcr.GetString('E_GENERAL') ) then
      lStRegime := Info.GetString('G_REGIMETVA') ;
    if lStRegime = '' then
      lStRegime := Contexte.RegimeDefaut ;
    if (lStRegime<>'') and (lStRegime<>GetString( vIndex , 'E_REGIMETVA')) then
      PutValue( vIndex, 'E_REGIMETVA', lStRegime ) ;

    lTobEcr := TrouveLigneHT( vIndex ) ;
    if Assigned( lTobEcr ) then
      begin
      PutValue( vIndex, 'E_TVA', lTobEcr.GetString('E_TVA') ) ;
      PutValue( vIndex, 'E_TPF', lTobEcr.GetString('E_TPF') ) ;
      end ;
{      if (ModeSaisie=msBor)
        then PutValue( vIndex, 'E_REGIMETVA', lStRegime )
        else PutValueAll(      'E_REGIMETVA', lStRegime ) ;
}   end

  // Code Tva / Tpf
  else
    begin
    // Ligne HT
    if EstHT( vIndex ) then
      begin
      // Tva
      if Compte_GetValue('G_TVA')=''
        then PutValue( vIndex, 'E_TVA', Contexte.CodeTvaDevaut )
        else PutValue( vIndex, 'E_TVA', Compte_GetValue('G_TVA') ) ;
      // Tpf
      PutValue( vIndex, 'E_TPF', Compte_GetValue('G_TPF') ) ;
      // y-a-t-il une ligne de ttc dans le groupe permettant de déterminé un régime ?
      end ;
    // toutes lignes
    lTobEcr := TrouveLigneTiers( vIndex ) ;
    if Assigned( lTobEcr ) then
      PutValue( vIndex, 'E_REGIMETVA', lTobEcr.GetString('E_REGIMETVA') ) ;
    end ;

end;

procedure TPieceCompta.AffecteConso( vIndex: Integer );
var lTobEcr  : Tob ;
begin
  lTobEcr := CGetTob( vIndex ) ;
  if lTobEcr = nil then Exit ;
  CGetConso( Detail[ vIndex - 1 ], Info );
  if lTobEcr.GetString('E_ECHE')='X' then
    putEche( vIndex, 'E_CONSO', lTobEcr.GetValue('E_CONSO') ) ;
end;

procedure TPieceCompta.CPutValue( vIndex : Integer; vChamps: String; vValeur: Variant);
begin
  if IsOut( vIndex ) then Exit ;
  Detail[ vIndex-1 ].PutValue( vChamps , vValeur ) ;
end;

procedure TPieceCompta.AffecteConf( vIndex : Integer );
var lStConf : String ;
    lTobEcr : Tob ;
begin

  lTobEcr := CGetTob( vIndex ) ;
  if lTobEcr = nil then Exit ;

  // Confidentialité actuelle
  lStConf := lTobEcr.GetValue( 'E_CONFIDENTIEL' ) ;

  // test confidentialité du Général
  if Info_LoadCompte( lTobEcr.GetString( 'E_GENERAL' ) ) then
    if Compte_GetValue('G_CONFIDENTIEL') > lStConf then
      lStconf := Compte_GetValue('G_CONFIDENTIEL') ;

  // test confidentialité du Auxiliaire
  if Info_LoadAux( lTobEcr.GetString( 'E_AUXILIAIRE' ) ) then
    if Aux_GetValue('T_CONFIDENTIEL') > lStConf then
      lStconf := Aux_GetValue('T_CONFIDENTIEL') ;

  // maj ligne
  if lStConf > lTobEcr.GetString( 'E_CONFIDENTIEL' ) then
    CPutValue( lTobEcr.GetIndex+1, 'E_CONFIDENTIEL', lStConf ) ;

end;

procedure TPieceCompta.AffecteAna(vIndex: Integer );
var lTobECr : TOB ;
begin

  lTobEcr := CGetTob( vIndex ) ;
  if lTobEcr = nil then Exit ;

  // TEST COMPTE VENTILABLE
  if EstVentilable( vIndex ) then
    begin
    lTobEcr.PutValue('E_ANA', 'X') ;
    AlloueAxe( lTobEcr ) ;

    // On ventile uniquement si le montant de l'écriture a été saisi :
    if (FModeVentil <> mvAucun) then
      if ( lTobEcr.GetValue('E_DEBIT') + lTobEcr.GetValue('E_CREDIT') ) <> 0 then
        begin
        VentilerTOB( lTobEcr, '', 0, V_PGI.OkDecV , (FModeVentil = mvAttente), Info.Dossier ) ; // Ventilation par defaut
        // synchronisation saisie mono section
        SynchroMonoSection( vIndex ) ;
        end ;

    end
  else
    begin
    lTobEcr.PutValue('E_ANA', '-') ;
    if lTobEcr.Detail.Count > 0 then
      lTobEcr.ClearDetail ;
    end ;

end;

function TPieceCompta.IsOut( vIndex : Integer) : Boolean ;
begin
  result := ( vIndex < 1 ) or ( vIndex > (Detail.Count) ) ;
end;

procedure TPieceCompta.AffecteRIB( vIndex : Integer ) ;
var lStCompte : String ;
    lStRIB    : String ;
    lTobTiers : Tob ;
    lTobEcr   : Tob ;
begin

  // Cas possible de non gestin des RIB
  if ChargeScenario( vIndex ) then
    begin
    if ( GetScenario( vIndex ).GetValue('SC_RIB') <> 'PRI' ) then Exit ;
    end
  else if ( Not Contexte.AttribRIBAuto ) then Exit ;

  lTobEcr := CGetTob( vIndex ) ;

  // Détermination du compte Tiers / TIC / TID...
  lTobTiers := nil ;
  if Info_LoadAux( lTobEcr.GetString( 'E_AUXILIAIRE' ) ) then
    begin
    lStCompte := lTobEcr.GetString( 'E_AUXILIAIRE' ) ;
    lTobTiers := Info.Aux.Item ;
    end
  else if Info_LoadCompte( lTobEcr.GetString( 'E_GENERAL' ) ) then
    begin
    if ( (Compte_GetValue('G_NATUREGENE') = 'TIC') or (Compte_GetValue('G_NATUREGENE') = 'TID')) then
      begin
      lStCompte := lTobEcr.GetString( 'E_GENERAL' ) ;
      lTobTiers := Info.Compte.Item ;
      end
    end
  else
    begin
    NotifyError( RC_RIB, '' ) ;
    Exit ;
    end ;

  if lTobTiers = nil then Exit ;

  if ( lTobTiers.GetNumChamp('RIBPRINCIPAL') < 0 )  or ( Trim(lTobTiers.GetString('RIBPRINCIPAL')) = '' ) then
    begin
    // on récupère le rib
    lStRIB := GetRIBPrincipal( lStCompte ) ;
    // on l'affecte à la tob tiers
    if ( lTobTiers.GetNumChamp('RIBPRINCIPAL') < 0 )
      then lTobTiers.AddChampSupValeur('RIBPRINCIPAL', lStRib )
      else lTobTiers.PutValue('RIBPRINCIPAL', lStRib ) ;
    end
  else
    lStRIB := lTobTiers.GetString('RIBPRINCIPAL') ;

  PutValue( vIndex , 'E_RIB', lStRib ) ;

end ;

procedure TPieceCompta.AffecteEche( vIndex : Integer ) ;
var lTob       : TOB ;
    lMultiEche : TMultiEche ;
    lNbEche    : integer ;
    i          : integer ;
begin

  lTob := CGetTob( vIndex ) ;
  if lTob = nil then Exit ;
  if lTob.GetInteger('E_NUMECHE')>1 then Exit ;

  // Compte lettrable ou pointable
  if EstAvecEche( vIndex ) then
    begin

    // MAJ Indicateurs + init des champs associées
    if GetValue( vIndex, 'E_ECHE' ) <> 'X' then
      begin
      lTob.PutValue( 'E_ECHE', 'X' ) ;
      if EstBqe ( vIndex )
        then begin
             CRemplirInfoPointage( lTob ) ;
             lTob.PutValue('E_MODEPAIE', '') ;
             end
        else begin
             CRemplirInfoLettrage( lTob ) ;
             lTob.PutValue('E_MODEPAIE', '') ;
             end ;
      end ;

    // Echéances
    if IsValidDebitCredit( vIndex ) then
      CalculEche( vIndex ) ;

    end
  // Pas d'échéance
  else
    begin
    if GetValue( vIndex, 'E_ECHE' ) = 'X' then
      // Suppression des infos si besoin
      begin
      Case FModeEche of

        // Gestion dans objet = mono échéance
        meAucun :  begin
                   CSupprimerInfoLettrage( lTob ) ;
                   end ;

        meDeporte : begin
                    CSupprimerInfoLettrage( lTob ) ;
                    lMultiEche := GetMultiEche( vIndex ) ;
                    if lMultiEche <> nil then
                      begin
                      lMultiEche.ClearDetail ;
                      FreeAndNil( lMultiEche ) ;
                      end ;
                    end ;

        // meMono, meMulti : nouvelle gestion
        else        begin
                    // comptage
                    lNbEche   := 0 ;
                    for i := vIndex to Detail.Count-1 do
                      begin
                      if Detail[i].GetInteger('E_NUMLIGNE')<>lTob.GetInteger('E_NUMLIGNE') then Break ;
                      Inc(lNbEche) ;
                      end ;
                    // suppression
                    for i := ( vIndex + lNbEche - 1 ) downto vIndex do
                      Detail[ i - 1 ].Free ;
                    CSupprimerInfoLettrage( lTob ) ;
                    end ;
         end ;
        // E_ENCAISSEMENT
        lTob.PutValue('E_ENCAISSEMENT', 'RIE' ) ;
      end ;
    end ;

end;

function TPieceCompta.ExisteEcheance(vIndex: Integer): Boolean;
begin
  result := False ;
  if isOut( vIndex ) then Exit ;
  result := GetMultiEche( vIndex ) <> nil ;
{
  result := ( GetValue( vNumLigne, 'E_DATECOMPTABLE' ) = iDate1900 ) or
            ( GetValue( vNumLigne, 'E_MODEPAIE'      ) = '' ) ;
}

end;

function TPieceCompta.IsValidDebitCredit(vIndex: Integer): Boolean;
begin
  result := False ;

  if isOut( vIndex ) then Exit ;

  result := ( CIsValidMontant( GetValue( vIndex, 'E_DEBITDEV')  ) = RC_PASERREUR ) and
            ( CIsValidMontant( GetValue( vIndex, 'E_CREDITDEV') ) = RC_PASERREUR ) and
            ( GetMontantDev( vIndex ) <> 0 ) ;
end ;

function TPieceCompta.EstRemplit(vIndex: Integer): Boolean;
var lTobEcr : Tob ;
begin
  result := False ;

  lTobEcr := CGetTob( vIndex ) ;
  if lTobEcr = nil then Exit ;

  result := ( lTobEcr.GetString('E_GENERAL')<>'' ) and Info_LoadCompte( lTobEcr.GetString('E_GENERAL') ) ;
  
end;

procedure TPieceCompta.InitSaisie ;
begin

  // Init lignes + indicateurs gestion
  videPiece ;

  // init info entete
  InitEnteteScenario    ;

  // Numero par défaut
  if ModeSaisie = msPiece then
    AttribNumeroTemp ;

  // indicateurs mode de saisie
  FAction     := taCreat ;
  FBoModif    := False ;

end;

function TPieceCompta.CGetFormule( vIndex : Integer ; vStFormule : hString ) : Variant ;
var lIdx  : integer ;
    lTob  : TOB ;
    lStr  : string ;
begin

  result     := #0 ;
  if isOut( vIndex ) then Exit ;

  vStFormule := uppercase( trim( vStFormule ) ) ;
  if vStFormule='' then Exit ;

  if Assigned( FFromTob ) then
    begin
    lTob := FFromTob ;
    lIdx := vIndex ;
    end
  else
    begin
    lIdx := QuelleLigne( vStFormule ) ;
    if lIdx > -2 then
      if lIdx <= 0 then
        lIdx := vIndex + lIdx
      else
        lIdx := GetDebutGroupe( vIndex ) + lIdx - 1 ;
    if isOut( lIdx ) then lIdx := vIndex ;

    lTob := GetTob( lIdx ) ;
    end ;

  // ===== TVA / TPF ===
  if (vStFormule='TVA') or (vStFormule='E_TVA')         then
    Result := GetTauxTva( lIdx )
//  else if vStFormule='TVANOR' then Result := Load_TVATPF( CGen, TRUE, 'NOR' )
//  else if vStFormule='TVARED' then Result := Load_TVATPF( CGen, TRUE, 'RED' )
  else if (vStFormule='TPF') or (vStFormule='E_TPF')    then
    Result := GetTauxTpf( lIdx )

  // ===== Solde ?? =====
  else if vStFormule='SOLDE' then
    Result := Abs( GetSolde - GetSoldePartiel(vIndex,vIndex) )

  // ===== Zones Ecriture =====
  else if Copy(vStFormule,1,2)='E_' then result := lTob.GetValue( vStFormule )

  // ===== Zones Général =====
  else if Copy(vStFormule,1,2)='G_' then
    begin
    if not Info_LoadCompte( lTob.GetString( 'E_GENERAL' ) ) then Exit ;
    result := Compte_GetValue( vStFormule ) ;
    end

  // ===== Zones Auxiliaire =====
  else if Copy(vStFormule,1,2)='T_' then
    begin
    if not Info_LoadAux( lTob.GetString( 'E_AUXILIAIRE' ) ) then Exit ;
    result := Aux_GetValue( vStFormule ) ;
    end

  // ===== Zones Journal =====
  else if Copy(vStFormule,1,2)='J_' then
    begin
    if not Info_LoadJournal( lTob.GetString( 'E_JOURNAL' ) ) then Exit ;
    result := Journal_GetValue( vStFormule ) ;
    end

  // ===== Zones Devise =====
  else if Copy(vStFormule,1,2)='D_' then
    begin
    lStr := GetString( lIdx, 'E_DEVISE') ;
    if Info.Devise.Load( [lStr] )<>-1 then
      result := Info.Devise.GetValue( vStFormule ) ;
    end

  // ===== Comptes auto =====
  else if Copy(vStFormule,1,4) = 'AUTO'
    then Result := TrouveAuto( Journal_GetValue( 'J_COMPTEAUTOMAT' ),
                               Ord( vStFormule[ Length( vStFormule ) ] ) - 48 )
    else if vStFormule = 'INTITULE' then
           begin
           if Info_LoadAux( lTob.GetString( 'E_AUXILIAIRE' ) )
             then Result := Aux_GetValue( 'T_LIBELLE' )
             else if Info_LoadCompte( lTob.GetString( 'E_GENERAL' ) )
                    then Result := Compte_GetValue( 'G_LIBELLE' ) ;
           end ;


end;

function TPieceCompta.GetFormule(vStFormule: HString): Variant;
begin
  result := CGetFormule( FCurIdx , vSTFormule ) ;
end;

function TPieceCompta.EstPieceEquilibree( vIndex : integer ) : Boolean;
var lInDeb    : integer ;
    lInFin    : integer ;
    i         : integer ;
    lSoldeDev : Double ;
    lSoldePiv : Double ;
    lTobL     : Tob ;
    lRDev     : RDevise ;
    lInGrp    : integer ;
begin

  result := True ;

  if (ModeSaisie=msBOR) and not isOut(vIndex) then
    begin
    lInGrp := GetNumGroupe( vIndex ) ;
    if lInGrp <= 0 then Exit ;
    GetBornesGroupe( lInGrp, lInDeb, lInFin) ;
    end
  else
    begin
    lInDeb := 1 ;
    lInFin := Count ;
    end ;

  if IsOut( lInDeb ) then Exit ;

  lTobL    := Detail[ lInDeb-1 ] ;
  if lTobL.GetString('E_DEVISE')=V_PGI.DevisePivot then Exit ;

  result := False ;

  // Calcul des soldes
  lSoldeDev := 0 ;
  lSoldePiv := 0 ;
  for i:= lInDeb to lInFin do
    begin
    lSoldeDev := lSoldeDev + Detail[i-1].GetDouble('E_DEBITDEV') - Detail[i-1].GetDouble('E_CREDITDEV') ;
    lSoldePiv := lSoldePiv + Detail[i-1].GetDouble('E_DEBIT') - Detail[i-1].GetDouble('E_CREDIT') ;
    end ;

  // Si le solde en devise n'est pas nul rien a faire
  lRDev := GetRDevise( lInDeb ) ;
  if Arrondi(lSoldeDev, lRDev.Decimale) <> 0 then Exit ;

  // Si le solde en devise est nul mais par celui en pivot, il faut équilibré
  result := Arrondi(lSoldePiv, V_PGI.OkDecV) = 0 ;

end;

function TPieceCompta.Count: Integer;
begin
  result := Detail.Count ;
end;


procedure TPieceCompta.AffecteVentilPourComplement( vIndex: Integer );
var lTobEcr : TOB ;
begin

  lTobEcr := GetTob( vIndex ) ;
  if lTobEcr = nil then Exit ;

  if lTobEcr.GetValue('E_ANA') <> 'X' then Exit ;

  // Report des valeurs modifiables dans la saisie complémentaire
  //  sur les éventuelles lignes de ventilation analytique
  PutVentil( vIndex, 'E_REFINTERNE',      lTobEcr.GetValue('E_REFINTERNE') ) ;
  PutVentil( vIndex, 'E_REFEXTERNE',      lTobEcr.GetValue('E_REFEXTERNE') ) ;
  PutVentil( vIndex, 'E_REFLIBRE',        lTobEcr.GetValue('E_REFLIBRE') ) ;
  PutVentil( vIndex, 'E_DATEREFEXTERNE',  lTobEcr.GetValue('E_DATEREFEXTERNE') ) ;
  PutVentil( vIndex, 'E_CONSO',           lTobEcr.GetValue('E_CONSO') ) ;
  PutVentil( vIndex, 'E_AFFAIRE',         lTobEcr.GetValue('E_AFFAIRE') ) ;
  PutVentil( vIndex, 'E_QTE1',            lTobEcr.GetValue('E_QTE1') ) ;
  PutVentil( vIndex, 'E_QTE2',            lTobEcr.GetValue('E_QTE2') ) ;
  PutVentil( vIndex, 'E_QUALIFQTE1',      lTobEcr.GetValue('E_QUALIFQTE1') ) ;
  PutVentil( vIndex, 'E_QUALIFQTE2',      lTobEcr.GetValue('E_QUALIFQTE2') ) ;

end;

function TPieceCompta.EstCli(vIndex: Integer): Boolean;
var lTob : Tob ;
begin
  result := False ;

  lTob := GetTob( vIndex ) ;
  if lTob = nil then Exit ;

  if (lTob.GetString('E_AUXILIAIRE') <> '') and Info_LoadAux( lTob.GetString('E_AUXILIAIRE') ) then
    result := (Aux_GetValue('T_NATUREAUXI') = 'CLI') or (Aux_GetValue('T_NATUREAUXI') = 'AUD')
  else if (lTob.GetString('E_GENERAL') <> '') and Info_LoadCompte( lTob.GetString('E_GENERAL') ) then
    result := (Compte_GetValue('G_NATUREGENE') = 'TID') ;

end;

function TPieceCompta.EstFou(vIndex: Integer): Boolean;
var lTob : Tob ;
begin
  result := False ;

  lTob := GetTob( vIndex ) ;
  if lTob = nil then Exit ;

  if (lTob.GetString('E_AUXILIAIRE') <> '') and Info_LoadAux( lTob.GetString('E_AUXILIAIRE') ) then
    result := (Aux_GetValue('T_NATUREAUXI') = 'FOU') or (Aux_GetValue('T_NATUREAUXI') = 'AUC')
  else if (lTob.GetString('E_GENERAL') <> '') and Info_LoadCompte( lTob.GetString('E_GENERAL') ) then
    result := (Compte_GetValue('G_NATUREGENE') = 'TIC') ;

end;

function TPieceCompta.ProratisePiece ( vIndex : Integer ; vChamps : String ; vNewMontant : Double ) : Boolean ;
var lRatio       : Double ;
    lOldVal      : Double ;
    lNewVal      : Double ;
    i            : Integer ;
    lBoEquilibre : Boolean ;
    lStChampsDC  : String ;
    lTob         : TOB ;
    lInDeb       : integer ;
    lInFin       : integer ;
    lInGrp       : integer ;
begin
  result := False ;
  if Not EstRemplit( vIndex ) then Exit ;
  if Detail.Count < 2 then Exit ;

  lOldVal := GetValue( vIndex, vChamps ) ;
  if (lOldVal = 0 ) or (lOldVal = vNewMontant) or (vNewMontant = 0) then Exit ;

  if ( (vChamps='E_CREDITDEV') and ( GetValue( vIndex, 'E_DEBITDEV')<>0  ) ) or
     ( (vChamps='E_DEBITDEV')  and ( GetValue( vIndex, 'E_CREDITDEV')<>0 ) ) then Exit ;

  if not PieceModifiable( vIndex ) then Exit ;

  if ModeSaisie = msBor then
    begin
    lInGrp := GetNumGroupe( vIndex ) ;
    if lInGrp <= 0 then Exit ;
    GetBornesGroupe( lInGrp, lInDeb, lInFin ) ;
    end
  else
    begin
    lInDeb := 1 ;
    lInFin := Detail.Count ;
    end ;

  lRatio       := Abs( vNewMontant / lOldVal ) ; // Ratio à appliquer
  lBoEquilibre := EstPieceSoldee(vIndex) ;       // indicateur d'équilibrage de la pièce avant traitement

  // Recalcul des montants pour toutes les lignes sauf celle qui initie le traitement
  for i := lInDeb to lInFin do
    begin
    if i = vIndex then
      PutValue( i, vChamps, vNewMontant  )
      else begin
           lTob := GetTob( i ) ;
           if lTob.GetValue('E_DEBITDEV')<>0
             then lStChampsDC := 'E_DEBITDEV'
             else if lTob.GetValue('E_DEBITDEV')<>0
                  then lStChampsDC := 'E_CREDITDEV'
                  else continue ;
           lOldVal := lTob.GetValue( lStChampsDC ) ;
           lNewVal := Arrondi( (lOldVal * lRatio), Devise.Decimale ) ;
           PutValue( i, lStChampsDC, lNewVal  ) ;
           end ;
    end ;

  // Si la pièce était équilibré, alors on rétabli l'équilibre si besoin
  if lBoEquilibre and not EstPieceSoldee(vIndex) then
    begin
    if vIndex = lInFin
      then i := lInDeb
      else i := lInFin ;
    if not IsOut( i ) then
      AttribSolde( i ) ;
    end ;

  result := True ;

end;

function TPieceCompta.GetMontant(vIndex: Integer): Double;
begin
  result := GetValue( vIndex , 'E_CREDIT' )  + GetValue( vIndex , 'E_DEBIT' ) ;
  result := Arrondi( result  , V_PGI.OkDecV ) ;   // FQ 17129 : Pb d'arrondi Delphi
end;

function TPieceCompta.GetMontantDev(vIndex: Integer): Double;
begin
  result := GetValue( vIndex , 'E_CREDITDEV' )  + GetValue( vIndex , 'E_DEBITDEV' ) ;
  result := Arrondi( result  , Devise.Decimale ) ;   // FQ 17129 : Pb d'arrondi Delphi
end;

function TPieceCompta.QuelChampsMontant ( vIndex : Integer ) : String ;
var lInParam : Integer ;
    lStSens  : String ;
    lInSens  : Integer ;
begin

  result := '' ;

  if not EstRemplit( vIndex ) then Exit ;

  if Info_LoadCompte( GetString( vIndex, 'E_GENERAL') ) then
    begin

    // Détermination sens du compte
    lStSens := Compte_GetValue('G_SENS') ;
    if lStSens = '' then lStSens := 'D' ;
    case lStSens[1] of
      'D' : lInParam := 1 ;
      'C' : lInParam := 2 ;
      else  lInParam := 3 ;
      end;

    // Détermination zone à saisir dans la grille    // QuelSens : Dernier paramètre : 1=Débit ; 2=Crédit ; 3=Mixte
    // QuelSens : Dernier paramètre : 1=Débit ; 2=Crédit ; 3=Mixte
    lInSens := QuelSens( GetEnteteS('E_NATUREPIECE'), Compte_GetValue('G_NATUREGENE'), lInParam, EstBQE(vIndex) ) ;

    // Retour Champs
    case lInSens of
      1 :  result := 'E_DEBITDEV' ;
      2 :  result := 'E_CREDITDEV' ;
      else result := 'E_DEBITDEV' ;
      end;

    end;


end;

function TPieceCompta.GetModeRegl( vIndex : Integer ): T_MODEREGL;
var lMultiEche  : TMultiEche ;
begin
  FillChar( result, Sizeof( result ), #0 ) ;

  if IsOut( vIndex ) then Exit ;

  lMultiEche := GetMultiEche( vIndex ) ;
  if lMultiEche = nil then Exit ;
  
  result := lMultiEche.GetModeRegl ;
end;

function TPieceCompta.EstJalEffet : Boolean ;
begin
  result := False ;
  if not Info_LoadJournal( GetEnteteS('E_JOURNAL') ) then Exit ;
  result := Journal_GetValue('J_EFFET') = 'X' ;
end;

function TPieceCompta.EstVentilable(vIndex: Integer ; vNumAxe : integer ): Boolean;
var lTobEcr : Tob ;
begin
  result := False ;

  lTobEcr := GetTob( vIndex ) ;
  if lTobEcr = nil then Exit ;

  if lTobEcr.GetString('E_GENERAL')='' then Exit ;
  if not Info_LoadCompte( lTobEcr.GetString('E_GENERAL') ) then Exit ;

  if vNumAxe = 0
    then result := Compte_GetValue('G_VENTILABLE') = 'X'
    else result := Compte_GetValue('G_VENTILABLE' + IntToStr(vNumAxe) ) = 'X'

end;

procedure TPieceCompta.SetModeRegl( vIndex: Integer ; vModeRegl : T_MODEREGL ) ;
var lMultiEche  : TMultiEche ;
begin

  if IsOut( vIndex ) then Exit ;

  lMultiEche := GetMultiEche( vIndex ) ;
  if lMultiEche = nil then Exit ;

  lMultiEche.SetModeRegl( vModeRegl ) ;

end;

function TPieceCompta.EstTvaLoc: Boolean;
begin
  result := ( ( GetEnteteS('E_QUALIFPIECE') = 'N' ) or
              ( GetEnteteS('E_QUALIFPIECE') = 'S' ) or
              ( GetEnteteS('E_QUALIFPIECE') = 'R' ) )
            and ( Contexte.OuiTvaEnc ) ;
end;

function TPieceCompta.UnSeulTiers( vIndex : integer ) : Boolean;
var i           : Integer ;
    lInDe       : integer ;
    lInA        : integer ;
    lStTiers    : string  ;
    lStRef      : string ;
begin

  // En mode bordereau se limiter au groupe
  if ModeSaisie = msBOR then
    GetBornesGroupe( GetInteger( vIndex, 'E_NUMGROUPEECR') , lInDe, lInA )
  else
    begin
    lInDe := 1 ;
    lInA  := Detail.count ;
    end ;

  lStRef      := '' ;
  result      := True ;

  for i := lInDe to lInA do
    begin
    if Detail[i-1].GetInteger('E_NUMECHE') > 1 then continue ;

    lStTiers := GetTiers( i ) ;
    if lStTiers<>'' then
      begin
      if (lStRef = '') then
        lStRef := lStTiers
      else if lStTiers <> lStRef then
        begin
        result := False ;
        Break ;
        end ;
      end ;
    end ;

end;


function TPieceCompta.EstMultiEche(vIndex: Integer): Boolean;
var lMultiEche : TMultiEche ;
begin
  result := False;
  if FModeEche = meDeporte then
    begin
    lMultiEche := getMultiEche( vIndex ) ;
    if lMultiEche <> nil then
      result := lMultiEche.NbEche > 1 ;
    end
  else
    begin
    if isOut( vIndex ) then Exit ;
    if Detail[vIndex-1].GetString('E_ECHE')='-' then Exit ;
    result :=    ( Detail[vIndex-1].GetInteger('E_NUMECHE') > 1 )  // détail multi...
              or (   ( Detail[vIndex-1].GetInteger('E_NUMECHE') = 1 ) // 1ère ligne d'éché + suivante sur eche > 1
                 and ( not IsOut( vIndex+1 ) and ( Detail[vIndex].GetInteger('E_NUMECHE') > 1 ) ) ) ;
    end ;

end;



function TPieceCompta.EstOkMultiEche(vIndex: Integer): Boolean;
begin
  Result :=    ( FModeEche in [meMulti, meDeporte]) // paramétrage objet métier
           and (ModeSaisie = msPiece)               // mode de saisie du journal
           and EstAvecEche(vIndex)                  // lettrable ou pointable ?
           and not ( EstBQE( vIndex ) or            // pas sur un compte de banque
                     EstVentilable( vIndex ) or     // pas si ventilable
                     EstJalEffet or                 // pas sur un journal d'effet
                     EstJalBqe or                   // pas sur un journal de banque
                     EstDivLett( vIndex ) ) ;       // pas sur un compte divers lettrable // Ajout CDL SBO 24/08/2005
end;

function TPieceCompta.EstAvecEche(vIndex: Integer): Boolean;
var lTobEcr : Tob ;
begin

  result := False ;

  lTobEcr := GetTob( vIndex ) ;
  if lTobEcr = nil then Exit ;

  if Compte_GetValue('G_COLLECTIF') = 'X' then
    begin
    if ( not Info_LoadAux( lTobEcr.GetString('E_AUXILIAIRE') ) )
       and ( Aux_GetValue('T_LETTRABLE') = '-' ) then Exit ;
    end
  else if not ( Compte_GetValue('G_POINTABLE')='X' ) and
          not ( Compte_GetValue('G_LETTRABLE') = 'X' ) then Exit ;

  result := True ;

end;



function TPieceCompta.GetRegimeTva( vIndex : Integer )  : String ;
begin
  result := GetString( vindex, 'E_REGIMETVA') ;
  if result = '' then
    result := Contexte.RegimeDefaut ;
end;

procedure TPieceCompta.DeleteRecord ( vIndex : Integer ) ;
var lBoARenumeroter : Boolean ;
    lMultiEche      : TMultiEche ;
    lTobEcr         : TOB ;
    lInNumEche      : integer ;
    lBoRecGroupe    : boolean ;
    lInGrpLM1       : integer ;
    lInGrpLP1       : integer ;
    lInGrpL         : integer ;
begin

  lTobEcr         := GetTob( vIndex ) ;

  if lTobEcr = nil then Exit ;
  if not EstLigneModifiable( vIndex ) then Exit ;

  lBoARenumeroter := ( vIndex <> Detail.Count ) ;

  lBoRecGroupe := False ;
  if ModeSaisie = msBor then
    begin
    lInGrpLM1     := 0 ;
    lInGrpLP1     := 1 ;
    // Groupe Ligne N-1
    if not isOut(vIndex-1) then
      lInGrpLM1 := GetInteger( vIndex-1, 'E_NUMGROUPEECR') ;
    // Groupe Ligne N+1
    if not isOut(vIndex+1) then
      lInGrpLP1 := GetInteger( vIndex+1, 'E_NUMGROUPEECR') ;
    // Groupe N
    lInGrpL      := GetInteger( vIndex, 'E_NUMGROUPEECR') ;
    lBoRecGroupe := (lInGrpL <> lInGrpLM1) and (lInGrpL <> lInGrpLP1) ;
    end ;

  // Gestion des échéances
  lInNumEche      := lTobEcr.GetInteger('E_NUMECHE') ;
  lMultiEche      := GetMultiEche( vIndex ) ;
  if ( lMultiEche<>nil ) then
    if (FModeEche=meDeporte) then
      begin
      lMultiEche.ClearDetail ;
      lMultiEche.free ;
      end
    else if (lInNumEche=1) and (lMultiEche.NbEche>1) then
      begin
      lMultiEche.ClearEche ;
      end ;

  // libération de l'EcrCompl
  CFreeTOBCompl( lTobEcr ) ;
  lTobEcr.Free ;

  if lBoARenumeroter then
    if (FModeEche = meMulti) and (lInNumEche>1) and assigned( lMultiEche )
      then lMultiEche.Renumerote( lInNumEche )
      else begin
           RenumeroteLignes ;
           if lBoRecGroupe then
             RecalculGroupe(vIndex) ;
           end ;

  if ( Detail.Count = 0 ) or ( ( Detail.count = 1 ) and ( not EstRemplit(1) ) ) then
    InitVariables ;

  FBoModif := True ;

end;

function TPieceCompta.InsereLigneTVA( vIndexHT : Integer ; vStRegime : String ; vBoTpf : Boolean ) : Boolean ;
Var lTobHT        : TOB ;
    lMontantHT    : Double ;
    lStSens       : String ;
    lStCodeTva    : String ;
    lStCodeTpf    : String ;
    lStCptGener   : String ;
    lMontantGener : Double ;
    lTobGener     : TOB ;
    lTauxGener    : Double ;
    lIndexTVA     : integer ;
//    lMontantTTC   : integer ;
begin

  Result := False ;
  if IsOut( vIndexHT ) then Exit ;

  Result := True ;
  lTobHT := GetTob( vIndexHT ) ;

  // Récup info ligne HT
  lStSens    := 'E_DEBITDEV' ;
  lMontantHT := lTobHT.GetDouble( lStSens ) ;
  if lMontantHT = 0 then
    begin
    lStSens := 'E_CREDITDEV' ;
    lMontantHT := lTobHT.GetDouble( lStSens ) ;
    end ;
  lStCodeTva := lTobHT.GetString('E_TVA') ;
  lStCodeTpf := lTobHT.GetString('E_TPF') ;
  lIndexTVA  := vIndexHT ;

  // ===========================================
  // ===== Génération de la TPF si demandé =====
  // ===========================================
  if vBoTpf then
    begin
    // Recup infos géénration
    lStCptGener   := GetCompteTpf( vIndexHT ) ;
    lTauxGener    := GetTauxTpf( vIndexHT ) ;
    if ( lStCodeTpf <> '' ) and ( lStCptGener <> '' ) and ( lTauxGener <> 0 ) and Info_LoadCompte( lStCptGener ) then
      begin
      lMontantGener := HT2TPF( lMontantHT, vStRegime, vBoTpf, lStCodeTva, lStCodeTpf, EstAchat(vIndexHT), Devise.Decimale ) ;
      // Mise en place de la nouvelle ligne
//      Inc(lIndexTVA) ;
{     if (not IsOut( lIndexTVA )) and ( not EstRemplit( lIndexTVA ) )
        then lTobGener := GetTob( lIndexTVA )
        else lTobGener := NewRecord( lIndexTVA ) ;
}     if not EstRemplit( Detail.count )
       then lTobGener := GetTob( Detail.count )
        else lTobGener := NewRecord ;
      lIndexTVA := FCurIdx ;
      PutValue( lIndexTVA, 'E_GENERAL',    lStCptGener ) ;
      PutValue( lIndexTVA, lStSens,        lMontantGener ) ;
      PutValue( lIndexTVA, 'E_CONTREPARTIEGEN', GetString( vIndexHT, 'E_CONTREPARTIEGEN' ) ) ;
      PutValue( lIndexTVA, 'E_CONTREPARTIEAUX', GetString( vIndexHT, 'E_CONTREPARTIEAUX' ) ) ;
      lTobGener.putValue( 'E_REGIMETVA',   vStRegime  ) ;
      lTobGener.putValue( 'E_TVA',         lStCodeTva ) ;
      lTobGener.putValue( 'E_TPF',         lSTCodeTpf ) ;
      end
    else if lTauxGener <> 0
         then Result := False ;
    end ;

  // ================================
  // ===== Génération de la TVA =====
  // ================================
  // Recup infos génération
  lStCptGener   := GetCompteTva( vIndexHT ) ;
  lTauxGener    := GetTauxTva( vIndexHT ) ;
  if ( lStCodeTva <> '' ) and ( lStCptGener <> '' ) and ( lTauxGener <> 0 ) and Info_LoadCompte( lStCptGener ) then
    begin
    // Calculs
    lMontantGener := HT2TVA( lMontantHT, vStRegime, vBoTpf, lStCodeTva, lStCodeTpf, EstAchat(vIndexHT), Devise.Decimale ) ;
    // Mise en place de la nouvelle ligne
{    Inc(lIndexTVA) ;
    if (not IsOut( lIndexTVA )) and ( not EstRemplit( lIndexTVA ) )
      then lTobGener := GetTob( lIndexTVA )
      else lTobGener := NewRecord( lIndexTVA ) ;
}
    if not EstRemplit( Detail.count )
      then lTobGener := GetTob( Detail.count )
      else lTobGener := NewRecord ;
    lIndexTVA := FCurIdx ;
    PutValue( lIndexTVA, 'E_GENERAL', lStCptGener ) ;
    PutValue( lIndexTVA, lStSens,     lMontantGener ) ;
    PutValue( lIndexTVA, 'E_CONTREPARTIEGEN', GetString( vIndexHT, 'E_CONTREPARTIEGEN' ) ) ;
    PutValue( lIndexTVA, 'E_CONTREPARTIEAUX', GetString( vIndexHT, 'E_CONTREPARTIEAUX' ) ) ;
    lTobGener.putValue( 'E_REGIMETVA',   vStRegime  ) ;
    lTobGener.putValue( 'E_TVA',         lStCodeTva ) ;
    lTobGener.putValue( 'E_TPF',         lSTCodeTpf ) ;
    end
  else
    if lTauxGener <> 0
       then Result:=False ;



  FCurIdx := lIndexTVA ;

end;


function TPieceCompta.EstJalBqe: Boolean;
begin
  result := False ;
  if Info_LoadJournal( GetEnteteS('E_JOURNAL') ) then
    result := ( Journal_GetValue('J_NATUREJAL') = 'BQE' ) or ( Journal_GetValue('J_NATUREJAL') = 'CAI' )  ;
end;

procedure TPieceCompta.AffecteTvaEnc(vIndex: Integer);
begin

  if IsOut( vIndex ) then Exit ;

  // Calculs qui vont bien...
  SetTvaMode( vIndex, GetTvaMode( vIndex ) ) ;

end;

procedure TPieceCompta.AffecteTypeMvt(vIndex: Integer);
var lTobEcr : Tob ;
begin
  if not EstRemplit( vIndex ) then Exit ;
  lTobEcr := CGetTob( vIndex ) ;
  CGetTypeMvt( lTobEcr, Info ) ;  // Modif Lettrage des comptes divers SBO 09/09/2005
end;

function TPieceCompta.EstCollFact( vIndex : Integer ) : Boolean;
var lStCompte : String ;
    lStListe  : String ;
    lStClient : String ;
    lTobEcr   : tob ;
begin
  Result := False ;

  if Not Contexte.OuiTvaEnc then Exit ;

  lTobEcr := GetTob( vIndex ) ;
  if lTobEcr = nil then Exit ;

  lStCompte := lTobEcr.GetString('E_GENERAL') ;
  if (lStCompte='') or not Info_LoadCompte( lStCompte ) then Exit ;

  lStListe  := Contexte.CollCliEnc ;
  if lStListe <> '' then
    begin

    if lStListe[ Length(lStListe) ] <> ';' then
      lStListe := lStListe + ';' ;

    repeat
      lStClient := ReadTokenSt( lStListe ) ;
      if lStClient = '' then Break ;
      if Copy( lStCompte, 1, Length( lStClient ) ) = lStClient then  // SBO 21/08/2007 : FQ 21139 : pb affectation ECHENC1..4 sur les ventes 
        begin
        result := True ;
        Exit ;
        end ;
    until ( lStListe='' ) ;

    end ;

  if not Result then
    begin
    lStListe := Contexte.CollFouEnc ;
    if lStListe <> '' then
      begin

      if lStListe[ Length(lStListe) ]<> ';' then
        lStListe := lStListe + ';' ;

      repeat
        lStClient := ReadTokenSt( lStListe ) ;
        if lStClient = '' then Break ;
        if Copy( lStCompte, 1, Length(lStClient) ) = lStClient then
          begin
          result:=True ;
          Exit ;
          end ;
      until ( lStListe = '' ) ;

      end ;
   end ;

end;

procedure TPieceCompta.ChargeEche ;
var lNumLigne  : Integer ;
    lIdxLigne  : Integer ;
    lTobEcr    : TOB ;
    lMultiEChe : TMultiEche ;
begin


  // Libération des éventuels objets déjà chargés
  FEcheances.ClearDetail ;

  if FModeEche<>meDeporte then Exit ;

  // Tri des lignes
  Detail.Sort('E_NUMLIGNE;E_NUMECHE');

  // Initialisation des objets Multi-echéance
  for lIdxLigne := 0 to ( Detail.Count - 1 ) do
    begin
    lTobEcr   := Detail[ lIdxLigne ] ;
    if lTobEcr.GetValue( 'E_NUMECHE' ) = 1 then
      CCreerMultiEche( FEcheances, self, lTobEcr, Action, False, mfAlone )   // 1ère échéance recopiée
    end ;

  // Ajout des lignes d'échéances > 1
  for lIdxLigne := ( Detail.Count - 1 ) downto 0 do
    begin
    lTobEcr   := Detail[ lIdxLigne ] ;
    lNumLigne := lTobEcr.GetInteger('E_NUMLIGNE') ;
    if lTobEcr.GetValue( 'E_NUMECHE' ) > 1 then
      begin
      lMultiEche := GetMultiEche( lNumLigne ) ;
      lMultiEche.AddEche( lTobEcr ) ;                            // 2ème à Nième échéances déplacées
      end ;
    end ;

  // Une fois les échéances chargées, on calcul le cumul des lignes d'écritures restantes
 // TEST MULTIECHE
  for lNumLigne := 0 to FEcheances.Detail.count - 1 do
    TMultiEche( FEcheances.Detail[ lNumLigne ] ).CumuleTobEcr ;

end;

procedure TPieceCompta.AttribNumeroDef;
var lDtDateC    : TDateTime ;   // date comptable saisie
    lInNumero   : Integer ;     // Numero de la pièce provisoire
    lStCompteur : string ;      // Code du compteur a utilisé
    lBoNormal   : Boolean ;
begin

  if Action <> taCreat then Exit ;

  if (ModeSaisie <> msPiece) and (GetEnteteI('E_NUMEROPIECE') > 0) then Exit ;

  if ( GetEnteteS('E_JOURNAL') = '' ) or
     not Info_LoadJournal( GetEnteteS('E_JOURNAL') ) then Exit ;

  lDtDateC  := GetEnteteDt( 'E_DATECOMPTABLE' ) ;
  lBoNormal := ( GetEnteteS('E_QUALIFPIECE') = 'N' ) or ( GetEnteteS('E_QUALIFPIECE') = 'I' ) ;
  if lBoNormal
    then lStCompteur := Journal_GetValue('J_COMPTEURNORMAL')
    else lStCompteur := Journal_GetValue('J_COMPTEURSIMUL') ;

  lInNumero := GetNewNumJal( GetEnteteS('E_JOURNAL'), lBoNormal, lDtDateC, lStCompteur, GetEnteteS('E_MODESAISIE'), Dossier ) ;

  PutEntete( 'E_NUMEROPIECE', lInNumero ) ;

end;

procedure TPieceCompta.SetTvaMode( vIndex : integer ; vTvaMode: TExigeTva ; vBoForce : Boolean );
Var //lTobEcr    : TOB ;
    lMultiEche : TMultiEche ;
    lNumLigne  : integer ;
    lStFlag    : String ;
    lInDe      : integer ;
    lInA       : integer ;
begin
  if not GereTvaMode( vIndex ) then Exit ;

  if ModeSaisie = msBor then
    GetBornesGroupe( GetValue( vIndex, 'E_NUMGROUPEECR' ), lInDe, lInA )
 else
   begin
   lInDe := 1  ;
   lInA  := Detail.count ;
   end ;

  // Attention
{ if vBoForce then
    begin
    FBoTvaFixe := True ;
    FTvaMode   := vTvaMode ;
    end
  else if FBoTvaFixe then
    vTvaMode := FTvaMode ;
}
  // ======================
  // ==== MODE FACTURE ====
  // ======================
  if GetSorteTva( vIndex ) <> stvDivers then
    begin
    lStFlag := '-';
    // Détermination du type de TVA
    for lNumLigne := lInDe to lInA do
      if EstHT( lNumLigne ) then
        begin
        Case vTvaMode of
           tvaDebit  : lStFlag:='-' ;
           tvaEncais : lStFlag:='X' ;
           tvaMixte  : begin
                       Info_LoadCompte( GetString(lNumLigne, 'E_GENERAL' ) ) ;
                       if Info.GetString('G_TVASURENCAISS')='X'
                         then lStFlag:='X'
                         else lStFlag:='-' ;
                       end ;
            else lStFlag := '-' ;
          end ;
        break ;
        end ;
    // Affectation de l'indicateur
    if ModeSaisie = msBor
      then PutValue( lInDe, 'E_TVAENCAISSEMENT', lStFlag )
      else PutValueAll( 'E_TVAENCAISSEMENT', lStFlag ) ;

    // Recalcul du tableau des montants par tva pour le grouope en cours
    CalculTvaEnc( vIndex ) ;

    // Recalcul des montants de tva des lignes
    for lNumLigne := lInDe to lInA do
      SetTvaEnc ( lNumLigne ) ;

    end
  else
  // ====================
  // ==== AUTRE MODE ====
  // ====================
    begin
    for lNumLigne := lInDe to lInA do
      // Pour chaque ligne de tiers, on calcul sur la base du taux de tva principal
      if EstTiers( lNumLigne ) then
        begin
        lMultiEche := GetMultiEche( lNumLigne ) ;
        if lMultiEche <> Nil then
          begin
          lMultiEche.ModifTva := False ;
          SetTvaTauxDirecteur( lNumLigne ) ;
          end ;
        end ;
    end ;

end;

function TPieceCompta.GereTvaMode ( vIndex : Integer ) : Boolean;
var lStNat : String ;
begin

  result := False ;
{$IFDEF ESP}
  if (Contexte.PaysLocalisation=CodeISOES) or (Not EstTvaLoc) then Exit ;
{$ELSE}
  if Not EstTvaLoc then Exit ;
{$ENDIF}

  lStNat := GetString( vIndex, 'E_NATUREPIECE') ;

  if not ( ( lStNat='FC' ) or ( lStNat='AC' ) or ( lStNat='OC' ) or
           ( lStNat='FF' ) or ( lStNat='AF' ) or ( lStNat='OF' ) ) then Exit ;

  Info_LoadJournal( GetEnteteS('E_JOURNAL') ) ;

  // Cas client : détermination dans les paramètres société
  if (((lStNat='OC') or (lStNat='FC') or (lStNat='AC')) and (Contexte.TvaEncSociete='TD')) then Exit ;

  // Cas Facture ou Avoir client mais pa sur un journal de vente
  if (((lStNat='FC') or (lStNat='AC')) and ( Journal_GetValue('J_NATUREJAL') <> 'VTE' )) then Exit ;

  // Cas Facture ou Avoir client mais pa sur un journal de vente
  if (((lStNat='FF') or (lStNat='AF')) and ( Journal_GetValue('J_NATUREJAL') <> 'ACH' )) then Exit ;

  Result := True ;

end;

procedure TPieceCompta.CalculTvaEnc( vIndex : integer );
var lNumL    : Integer ;
    lSolde   : Double ;
    lNumBase : Integer ;
    lInDe    : Integer ;
    lInA     : Integer ;
begin

  FillChar( FTabTvaEnc, Sizeof(FTabTvaEnc), #0 ) ;

  if not EstTvaLoc then Exit ;
  if not GereTvaMode( vIndex ) then Exit ;

  if ModeSaisie = msBOR then
    GetBornesGroupe( GetInteger( vIndex, 'E_NUMGROUPEECR' ), lInDe, lInA )
  else
    begin
    lInDe := 1 ;
    lInA := Detail.count ;
    end ;

  for lNumL := lInDe to lInA do
    begin
    // Uniquement pour les lignes de HT...
    if EstHT( lNumL ) then
      begin
      if GetSorteTva( lNumL ) = stvVente
        then lSolde := GetDouble( lNumL, 'E_CREDIT') - GetDouble( lNumL, 'E_DEBIT')
        else lSolde := GetDouble( lNumL, 'E_DEBIT')  - GetDouble( lNumL, 'E_CREDIT') ;
      lSolde := Arrondi( lSolde, Devise.Decimale ) ; // FQ 17129 : Pb d'arrondi Delphi

      // Tva sur encaissement
      if GetString( lNumL, 'E_TVAENCAISSEMENT' ) = 'X' then
        begin
        lNumBase := Tva2NumBase( GetValue( lNumL, 'E_TVA') ) ;
        if lNumBase > 0 then
          FTabTvaEnc[lNumBase] := FTabTvaEnc[lNumBase] + lSolde;
        end
      else
      // Tva sur débit
          FTabTvaEnc[5]        := FTabTvaEnc[5]        + lSolde ;
      end ;
    end ;

end;

function TPieceCompta.GetTvaEnc(vIndex: Integer): String;
var lTobEcr : Tob ;
begin
  result := '' ;
  if not EstRemplit( vIndex ) then Exit ;

  if EstCli( vIndex )
    then result := Contexte.TvaEncSociete
    else
      begin
      lTobEcr := GetTob( vIndex ) ;
      if (lTobEcr.GetString('E_AUXILIAIRE') <> '') and Info_LoadAux( lTobEcr.GetString('E_AUXILIAIRE') )
        then result := Aux_GetValue('T_TVAENCAISSEMENT')
        else result := Compte_GetValue('G_TVAENCAISSEMENT') ;
      end ;

end;

procedure TPieceCompta.SetTvaTauxDirecteur(vIndex: Integer);
Var lTobEche     : Tob ;
    lMultiEche   : TMultiEche ;
    lMontantTTC  : Double ;
    lMontantHT   : Double ;
    lMontantEche : Double ;
    lTauxTva     : Double ;
    lCodeTva     : String ;
    lNumEche     : Integer ;
begin

  if IsOut( vIndex ) then Exit ;
  if Not EstTvaLoc then Exit ;
  if GetTvaMode( vIndex ) = tvaMixte then Exit ;
  if GetSorteTva( vIndex ) <> stvDivers then Exit ;
  if Not EstTiers( vIndex ) then Exit ;
  if Not ( EstCli( vIndex ) or EstFou( vIndex ) ) then Exit ;

  lMultiEche := GetMultiEche( vIndex ) ;
  if (lMultiEche = nil) or (lMultiEche.nbEche = 0) then Exit ;
  if lMultiEche.ModifTva then Exit ;
  if lMultiEche.GetTotalEche = 0 then Exit ;

  lMontantTTC := GetValue( vIndex, 'E_CREDIT') - GetValue( vIndex, 'E_DEBIT') ;
  lMontantTTC := Arrondi( lMontantTTC, V_PGI.OkDecV ) ; // FQ 17129 : Pb d'arrondi Delphi
  if not EstCli( vIndex ) then lMontantTTC := - lMontantTTC ;

  lCodeTva := Contexte.NumCodeBase[1] ;
  if lCodeTva = '' then Exit ;

  lTauxTva := Tva2Taux( GetRegimeTva( vIndex ), lCodeTva, Not EstCli( vIndex) ) ;
  if lTauxTva = -1 then Exit ;

  // Prorater le HT, calculé sur le taux directeur, sur les échéances
  lMontantHT := Arrondi( lMontantTTC / (1.0 + lTauxTva), V_PGI.OkDecV ) ;
  for lNumEche := 1 to lMultiEche.nbEche do
    begin
    lTobEche := lMultiEche.GetEche( lNumEche ) ;
    lMontantEche := lTobEche.GetDouble('E_DEBIT') + lTobEche.GetDouble('E_CREDIT') ;
    lMontantEche := Arrondi( lMontantHT * ( lMontantEche / abs( lMontantTTC ) ) , V_PGI.OkDecV ) ;
    if GetTvaMode( vIndex ) = tvaDebit
      then lTobEche.PutValue( 'E_ECHEDEBIT', lMontantEche )
      else lTobEche.PutValue( 'E_ECHEENC1',  lMontantEche ) ;
    end ;

  // Eviter la ré-initialisation ultérieure, recalcul par prorata ou manuel
  lMultiEche.ModifTva := True ;

end;

function TPieceCompta.GetTvaMode( vIndex : Integer ) : TExigeTva;
var lStCodeTva : String ;
    lNumLigne  : Integer ;
    lInDe    : Integer ;
    lInA     : Integer ;
begin

  result := tvaMixte ;

  if ModeSaisie = msBOR then
    GetBornesGroupe( GetInteger( vIndex, 'E_NUMGROUPEECR' ), lInDe, lInA )
  else
    begin
    lInDe := 1 ;
    lInA := Detail.count ;
    end ;

  for lNumLigne := lInDe to lInA do
    if EstTiers( lNumLigne ) then
      begin
      // Mode Tva du tiers
      if GetString( lNumLigne, 'E_AUXILIAIRE' ) <> ''
        then lStCodeTva := Info.GetString('T_TVAENCAISSEMENT')        // Tiers
        else lStCodeTva := Info.GetString('G_TVAENCAISSEMENT') ;   // Tic ou Tid
      // TExigeTva résultante
      if lStCodeTva = 'TD'
        then result := tvaDebit
        else if lStCodeTva = 'TE'
             then result := tvaEncais ;
      end ;

end;

procedure TPieceCompta.SetTvaEnc ( vIndex : Integer ; vBoForce : boolean) ;
var lMultiEche   : TMultiEche ;
    lTobEcr      : TOB ;
    lCoefTiers   : double ;
    lTotalDebit  : double ;
    lTotalCredit : double ;
    lStTiers     : string ;
    lStChamp     : string ;
    lTotalEche   : double ;
begin

  lTobEcr := GetTob( vIndex ) ;

  if lTobEcr.GetInteger('E_NUMECHE') > 1 then Exit ;

  lMultiEche  := GetMultiEche( vIndex ) ;
  if lMultiEche = nil then Exit ;
  if lMultiEche.ModifTVA then Exit ;

  lTobEcr.PutValue( 'E_ECHEENC1',    0 ) ;
  lTobEcr.PutValue( 'E_ECHEENC2',    0 ) ;
  lTobEcr.PutValue( 'E_ECHEENC3',    0 ) ;
  lTobEcr.PutValue( 'E_ECHEENC4',    0 ) ;
  lTobEcr.PutValue( 'E_ECHEDEBIT',   0 ) ;
  lTobEcr.PutValue( 'E_EMETTEURTVA', '-' ) ;

  // Condition d'exclusion
  if not GereTvaMode( vIndex ) then Exit ;
  if not EstCollFact( vIndex ) then Exit ;

  if GetString( vIndex, 'E_AUXILIAIRE') <> '' then
    begin
    lStChamp := 'E_AUXILIAIRE' ;
    lStTiers := GetString( vIndex, 'E_AUXILIAIRE') ;
    end
  else
    begin
    lStChamp := 'E_GENERAL' ;
    lStTiers := GetString( vIndex, 'E_GENERAL') ;
    end ;

  if ModeSaisie=msBor
    then GetTotauxPourChamps( lStChamp, lStTiers, lTotalDebit, lTotalCredit, GetInteger(vIndex, 'E_NUMGROUPEECR') )
    else GetTotauxPourChamps( lStChamp, lStTiers, lTotalDebit, lTotalCredit ) ;

  lTotalEche  := lMultiEche.GetTotalEche( False );  // GetDouble( vindex, 'E_DEBIT') ;

  if Arrondi(lTotalDebit - lTotalCredit, V_PGI.OkDecV) = 0 then Exit ;

  if Arrondi( lTotalEche - (lTotalDebit + lTotalCredit), V_PGI.OkDecV) <> 0
    then lCoefTiers := (lTotalEche / (lTotalDebit + lTotalCredit))
    else lCoefTiers := 1 ;

  if vBoForce or UnSeulTiers( vIndex )
    then lMultiEche.SetTvaEnc( FTabTvaEnc[1], FTabTvaEnc[2], FTabTvaEnc[3], FTabTvaEnc[4], FTabTvaEnc[5], lCoefTiers )
    ;
//    else if ( GetSorteTva = stvDivers ) and lMultiEche.modifTva
//         then lMultiEche.CumuleTobEcr ;

end;

procedure TPieceCompta.GetInfosLigne(vIndex: Integer; var vDebit, vCredit, vTotalDebit, vTotalCredit: Double);
var lStCpt : String ;
    lNumL  : Integer ;
begin
  vDebit       := 0 ;
  vCredit      := 0 ;
  vTotalDebit  := 0 ;
  vTotalCredit := 0 ;

  if not EstRemplit( vIndex ) then Exit ;
  lStCpt := GetValue( vIndex, 'E_GENERAL' ) ;

  vDebit       := GetValue( vIndex, 'E_DEBITDEV' ) ;
  vCredit      := GetValue( vIndex, 'E_CREDITDEV' ) ;

  for lNumL := 1 to Detail.Count do
    if GetValue( lNumL, 'E_GENERAL' ) = lStCpt then
      begin
      vTotalDebit  := vTotalDebit + GetValue( lNumL, 'E_DEBITDEV' ) ;
      vTotalCredit := vTotalCredit + GetValue( lNumL, 'E_CREDITDEV' ) ;
      end ;

  vTotalDebit  := Arrondi( vTotalDebit,  Devise.Decimale ) ; // FQ 17129 : Pb d'arrondi Delphi
  vTotalCredit := Arrondi( vTotalCredit, Devise.Decimale ) ; // FQ 17129 : Pb d'arrondi Delphi

end;

function TPieceCompta.EstSoumisTPF( vindex : Integer ) : Boolean;
var lStcompte : String ;
begin

  result := False ;
  if IsOut( vindex ) then Exit ;

  lStCompte := GetValue( vindex, 'E_AUXILIAIRE') ;
  if lStCompte <> '' then
    begin
    if Info_LoadAux( lStCompte ) then
      result := Aux_GetValue('T_SOUMISTPF') = 'X' ;
    end
  else
    begin
    lStCompte := GetValue( vindex, 'E_GENERAL') ;
    if Info_LoadCompte( lStCompte ) then
      result := Compte_GetValue('G_SOUMISTPF') = 'X' ;
    end ;

end;

function TPieceCompta.ControleTva( vIndex : Integer ) : Integer ;
Var // Montants saisis dans les lignes
    lDebitHT         : Double ;
    lCreditHT        : Double ;
    lTotalDebitHT    : Double ;
    lTotalCreditHT   : Double ;
    lDebitTVA        : Double ;
    lCreditTVA       : Double ;
    lTotalDebitTVA   : Double ;
    lTotalCreditTVA  : Double ;
    lDebitTPF        : Double ;
    lCreditTPF       : Double ;
    lTotalDebitTPF   : Double ;
    lTotalCreditTPF  : Double ;
    // Montants calculés
    lCalcDebitTVA        : Double ;
    lCalcCreditTVA       : Double ;
    lCalcTotalDebitTVA   : Double ;
    lCalcTotalCreditTVA  : Double ;
    lCalcDebitTPF        : Double ;
    lCalcCreditTPF       : Double ;
    lCalcTotalDebitTPF   : Double ;
    lCalcTotalCreditTPF  : Double ;
    // Ecarts
    lEcartTVA        : double ;
    lEcartTPF        : double ;
    lTotalEcartTVA   : double ;
    lTotalEcartTPF   : double ;
    // Autres infos
    lStCptTVA        : String ;
    lStCptTPF        : String ;
    lTauxTva         : Double ;
    lTauxTpf         : Double ;
    lResol           : Double ; // ??
    lBoAchat         : Boolean ;
    lBoSoumisTPF     : boolean ;
    lTobEcr          : TOB ;
    lTobTva          : TOB ;
    lTobTpf          : TOB ;
    lStCodeTva       : String ;
    lStCodeTpf       : String ;
    lStRegimeTva     : String ;
    lDecim           : Integer ;
begin
  result := TVA_PASERREUR ;
  if  not EstRemplit( vIndex ) then Exit ;

  // Init des variables
  lDebitHT         := 0 ;
  lCreditHT        := 0 ;
  lTotalDebitHT    := 0 ;
  lTotalCreditHT   := 0 ;
  lDebitTVA        := 0 ;
  lCreditTVA       := 0 ;
  lTotalDebitTVA   := 0 ;
  lTotalCreditTVA  := 0 ;
  lDebitTPF        := 0 ;
  lCreditTPF       := 0 ;
  lTotalDebitTPF   := 0 ;
  lTotalCreditTPF  := 0 ;
  lDecim           := Devise.Decimale ;

  // Recup infos régime, Codes TVA et TPF du compte HT
  lTobEcr      := GetTob( vIndex ) ;
  lStCodeTva   := lTobEcr.GetValue('E_TVA') ;
  lStCodeTpf   := lTobEcr.GetValue('E_TPF') ;

  lStRegimeTva := GetRegimeTVA( vIndex ) ;
  lBoSoumisTPF := EstSoumisTPF( vIndex ) ;
  lBoAchat     := EstAchat(vIndex) ;

  lTauxTva     := Tva2Taux( lStRegimeTva, lStCodeTva, lBoAchat ) ;
  lTauxTpf     := Tpf2Taux( lStRegimeTva, lStCodeTpf, lBoAchat ) ;
  lStCptTVA    := GetCompteTva( vIndex ) ;
  lStCptTPF    := GetCompteTpf( vIndex ) ;

  // Test comptes Tva / Tpf
  if ((lStCptTVA='') and (lTauxTva<>0)) or ((lStCptTPF='') and (lTauxTpf<>0)) then
     begin
     result := TVA_ERRCPTTVATPF ;
     Exit ;
     end ;

  // Calculs et recup de la ligne HT
  GetInfosLigne( vIndex, lDebitHT, lCreditHT, lTotalDebitHT, lTotalCreditHT ) ;

  // Calculs et recup de la ligne TVA
  if lStCptTVA<>'' then
    begin
    lTobTva := TrouveLigneCompte( lStCptTva ) ;
    if lTobTva <> nil then
      GetInfosLigne( lTobTva.GetIndex + 1, lDebitTva, lCreditTVA, lTotalDebitTVA, lTotalCreditTVA ) ;
    end ;
  lCalcDebitTVA       := HT2TVA( lDebitHT,       lStRegimeTva, lBoSoumisTPF, lStCodeTva, lStCodeTpf, lBoAchat, lDecim ) ;
  lCalcCreditTVA      := HT2TVA( lCreditHT,      lStRegimeTva, lBoSoumisTPF, lStCodeTva, lStCodeTpf, lBoAchat, lDecim ) ;
  lCalcTotalDebitTVA  := HT2TVA( lTotalDebitHT,  lStRegimeTva, lBoSoumisTPF, lStCodeTva, lStCodeTpf, lBoAchat, lDecim ) ;
  lCalcTotalCreditTVA := HT2TVA( lTotalCreditHT, lStRegimeTva, lBoSoumisTPF, lStCodeTva, lStCodeTpf, lBoAchat, lDecim ) ;

  // Calculs et recup de la ligne TPF
  if lStCptTPF<>'' then
    BEGIN
    lTobTPF := TrouveLigneCompte( lStCptTPF ) ;
    if lTobTPF <> nil then
      GetInfosLigne( lTobTPF.GetIndex + 1, lDebitTPF, lCreditTPF, lTotalDebitTPF, lTotalCreditTPF ) ;
    END ;
  lCalcDebitTPF       := HT2TPF( lDebitHT,       lStRegimeTva, lBoSoumisTPF, lStCodeTva, lStCodeTpf, lBoAchat, lDecim ) ;
  lCalcCreditTPF      := HT2TPF( lCreditHT,      lStRegimeTva, lBoSoumisTPF, lStCodeTva, lStCodeTpf, lBoAchat, lDecim ) ;
  lCalcTotalDebitTPF  := HT2TPF( lTotalDebitHT,  lStRegimeTva, lBoSoumisTPF, lStCodeTva, lStCodeTpf, lBoAchat, lDecim ) ;
  lCalcTotalCreditTPF := HT2TPF( lTotalCreditHT, lStRegimeTva, lBoSoumisTPF, lStCodeTva, lStCodeTpf, lBoAchat, lDecim ) ;

  // Calculs des écarts
  lEcartTVA      := Arrondi( ( lDebitTva - lCalcDebitTVA ) - ( lCreditTVA - lCalcCreditTVA ) , lDecim ) ;
  lEcartTPF      := Arrondi( ( lDebitTPF - lCalcDebitTPF ) - ( lCreditTPF - lCalcCreditTPF ) , lDecim ) ;
  lTotalEcartTVA := Arrondi( ( lTotalDebitTVA - lCalcTotalDebitTVA ) - ( lTotalCreditTVA - lCalcTotalCreditTVA ), lDecim ) ;
  lTotalEcartTPF := Arrondi( ( lTotalDebitTPF - lCalcTotalDebitTPF ) - ( lTotalCreditTPF - lCalcTotalCreditTPF ), lDecim ) ;

  // Tests tva correcte
  lResol  := Resolution ( lDecim ) ;
  if ( ( Abs(lEcartTVA)      <= lResol )   and ( Abs(lEcartTPF)      <= lResol ) and
       ( Abs(lTotalEcartTVA) <= lResol )   and ( Abs(lTotalEcartTPF) <= lResol ) )
    then Exit ;

  result := TVA_ERRTVAINCORRECT ;

end;

procedure TPieceCompta.FinaliseRegimeEtTva  ( vIndex : Integer ) ;
var lInIndex   : Integer ;
    lBoUnTiers : Boolean ;
    lStRegime  : string ;
    lStTva     : string ;
    lStTPF     : string ;
    lInGroupe  : integer ;
    lInDe      : integer ;
    lInA       : integer ;
begin

  // MAJ de la TVa / TPF uniquement si un seul tiers ?
  lBoUnTiers := False ;
  if (ModeSaisie = msBor) and (vIndex >= 1) then
    GetBornesGroupe( vIndex, lInDe, lInA )
  else
    begin
    if (ModeSaisie = msPiece) then
      lBoUnTiers := UnSeulTiers ;
    lInDe      := 1 ;
    lInA       := Detail.count ;
    end ;

  // Uniformisation du régime
  lInGroupe := -1 ;
  for lInIndex := lInDe to lInA do
    begin
    // ANCIENNE METHODE
{    if ModeSaisie = mspiece then
      begin
      // Valeur du régiome de tva unique pour toute la pièce
      if (ModeSaisie=msPiece) and (FRegimeTva<>'') then
        PutValue( lInIndex, 'E_REGIMETVA', FRegimeTva ) ;

      // Seule les lignes de HT peuvent avoir des valeur tva/tpf distinctes
      if lBoUnTiers and ( not EstHT( lInIndex ) ) then
        begin
        if FCodeTva <> '' then
          PutValue( lInIndex, 'E_TVA', FCodeTva ) ;
        if FCodeTpf <> '' then
          PutValue( lInIndex, 'E_TPF', FCodeTpf ) ;
        end ;

      // Calcul des champs de tva sur encaissement
      SetTvaEnc( lInIndex ) ;
      end
    else
}    // NOUVELLE METHODE
      begin
      if lInGroupe <> GetInteger( lInIndex, 'E_NUMGROUPEECR') then
        begin
        lInGroupe := GetInteger( lInIndex, 'E_NUMGROUPEECR') ;
        GetInfosTva( lInIndex, lStRegime, lStTva, lStTpf ) ;
        CalculTvaEnc( lInIndex ) ;
        if ModeSaisie = msBor then
          lBoUnTiers := UnSeulTiers( lInIndex ) ;
        end ;

      PutValue( lInIndex, 'E_REGIMETVA', lStRegime ) ;
      if not EstHT( lInIndex ) then
        begin
        PutValue( lInIndex, 'E_TVA', lStTva ) ;
        PutValue( lInIndex, 'E_TPF', lStTpf ) ;
        end ;

      if lBoUnTiers and EstTiers( lInIndex ) then
          SetTvaEnc( lInIndex, True ) ;

      end ;
    end ;

end;

procedure TPieceCompta.InitVariables;
begin
//  FBoTvaFixe    := False ;
//  FTvaMode      := tvaMixte ;
//  FRegimeTVA    := '' ;
//  FCodeTVA      := '' ;
//  FCodeTPF      := '' ;
  FCurIdx       := 0 ;
  FFromTob      := nil ;
  FAction       := taCreat ;

  FillChar( FTabTvaEnc, Sizeof(FTabTvaEnc), #0 ) ;

end;


procedure TPieceCompta.videPiece;
begin
  // Init variables
  InitVariables ;

  // Libération de la pièce précédentes
  ClearDetail ;              // Ecritures visu

  // Echéances
  if ModeEche = meDeporte then
    FEcheances.ClearDetail ;

  // Copie SAV pour traitement postérieur
  FTobOrigine.ClearDetail ;

  // Libération des Tobcompl
  FEcrCompl.ClearDetail ;

  // Reinit de la liste des TP
  FListeTP.clear ;

  // Pièce multi-etab
  LibereListe( FPiecesEtab, False ) ;

end;

procedure TPieceCompta.ChargeChampsSup;
var lNumLigne : Integer ;
    lTobEcr   : Tob ;
    iAxe      : Integer ;
begin
  for lNumLigne := 1 to Detail.Count do
    begin
    lTobEcr := Detail[ lNumLigne - 1 ] ;
    lTobEcr.AddChampSupValeur( 'ACCELERATEUR', 'X' ) ;
    lTobEcr.AddChampSupValeur( 'COMPS',        'X' ) ;
    lTobEcr.AddChampSupValeur( 'CUTOFF',       'X' ) ;
    lTobEcr.AddChampSupValeur( 'PBTAUX',       'X' ) ;

    lTobEcr.AddChampSupValeur( 'SECTIONA1', '' ) ;
    lTobEcr.AddChampSupValeur( 'SECTIONA2', '' ) ;
    lTobEcr.AddChampSupValeur( 'SECTIONA3', '' ) ;
    lTobEcr.AddChampSupValeur( 'SECTIONA4', '' ) ;
    lTobEcr.AddChampSupValeur( 'SECTIONA5', '' ) ;
    if lTobEcr.GetString('E_ANA') = 'X' then
      begin
      for iAxe := 1 to MAXAXE do
        if lTobEcr.Detail[iAxe-1].Detail.Count = 1 then
          lTObEcr.PutValue('SECTIONA' + IntToStr(iAxe), lTobEcr.Detail[iAxe-1].Detail[0].GetString('Y_SECTION') ) ;
      end ;

    end ;
end;


function TPieceCompta.GetMultiEche ( vIndex : Integer ) : TMultiEche ;
var i         : Integer ;
    lTob      : Tob ;
    lNumLigne : Integer ;
begin

  result := nil ;
//  if FModeEche = meAucun then Exit ;

  lTob := GetTob( vIndex ) ;
  if lTob=nil then Exit ;

  lNumLigne := lTob.GetInteger('E_NUMLIGNE') ;

  // recherche du TMultiEche correspondant à la ligne d'écriture
  if FModeEche = meDeporte then
    begin
    for i := 0 to FEcheances.Detail.count - 1 do
      begin
      if TMultiEche( FEcheances.Detail[ i ] ).NumLigne = lNumLigne then
        begin
        result := TMultiEche( FEcheances.Detail[ i ] ) ;
        break ;
        end ;
      end
    end
  else
    begin
    if lTob.GetInteger('E_NUMECHE')>1 then
      lTob := GetFirstEche( vIndex ) ;
    result := GetMultiEche( lTob ) ;
    end ;

end;

function TPieceCompta.IsValidEtab(vEtab: String): Boolean;
var lStEtab : String ;
begin

  result := True ;

  // Si etab ligne = étab entête alors on est toujours bon
  if GetEnteteS('E_ETABLISSEMENT') = vEtab then Exit ;

  // Test Etablissement forcé
  lStEtab := EtabForce ;
  if (lStEtab <> '') and ( lStEtab <> vEtab )
    then Result := False

  // Test Lien paramétré
  else if Contexte.GetLiensEtab( vEtab ) = nil
    then result := False ;

end;

function TPieceCompta.EstMultiEtab: Boolean;
var i : Integer ;
begin
  result := False ;
  For i := 0 to Detail.Count - 1 do
    begin
    if Detail[ i ].GetString('E_ETABLISSEMENT') <> getEnteteS('E_ETABLISSEMENT') then
      begin
      result := True ;
      exit ;
      end ;
    end ;
end;

function TPieceCompta.GetRMVT: RMVT;
begin

  result.Axe           := '' ;
  result.Etabl         := GetEnteteS('E_ETABLISSEMENT') ;
  result.Jal           := GetEnteteS('E_JOURNAL') ;
  result.Exo           := Contexte.GetExoDT( GetEnteteDt('E_DATECOMPTABLE' ) ) ;
  result.CodeD         := GetEnteteS('E_DEVISE') ;
  result.Simul         := GetEnteteS('E_QUALIFPIECE') ;
  result.Nature        := GetEnteteS('E_NATUREPIECE') ;
  result.DateC         := GetEnteteDt('E_DATECOMPTABLE') ;
  result.Num           := GetEnteteI('E_NUMEROPIECE') ;
  result.TauxD         := Devise.Taux ;
  result.DateTaux      := Devise.DateTaux ;
  result.Valide        := False ;
  result.ANouveau      := False ; //ANouveau ;
  result.ModeSaisieJal := GetEnteteS('E_MODESAISIE') ;   {FP 02/05/2006}

end;

procedure TPieceCompta.TraiteMultiEtab(vIndex: Integer) ;
var lTobEcr      : TOB ;
    lPieceEtab   : TPieceCompta ;
    lTobEtab     : TOB ;
    lLiensEtab   : TOB ;
    lIdxLigne    : Integer ;
    lStCpt       : string ;
    lStAux       : string ;
    lStLiaison   : string ;
    lMultiEcheEcr  : TMultiEche ;
    lMultiEcheEtab : TMultiEche ;
    lTobEche       : Tob ;
    lNumEche       : integer ;
    lNumLigne      : integer ;
begin

  lTobEcr    := GetTob( vIndex ) ;
  if lTobEcr.GetString('E_ETABLISSEMENT') = GetEnteteS('E_ETABLISSEMENT') then Exit ;
  // ligne de multi échéances > 1 traité avec la 1ère échéance
  if lTobEcr.GetInteger('E_NUMECHE') > 1  then Exit ;

  // Récup du paramétrage de liaison inter-établissement
  lLiensEtab := Contexte.GetLiensEtab( lTobEcr.GetString('E_ETABLISSEMENT') ) ;
  if lLiensEtab = nil then Exit ;

  // -------------------------------------------------
  // --- MAJ de la pièce sur l'autre établissement ---
  // -------------------------------------------------

  // Recherche ou création de la pièce pour l'établissement
  lPieceEtab := GetPieceEtab( lTobEcr.GetString('E_ETABLISSEMENT') ) ;
  if lPieceEtab = nil then Exit ;

  // recup gén et auxi
  lStCpt     := lTobEcr.GetString('E_GENERAL') ;
  lStAux     := lTobEcr.GetString('E_AUXILIAIRE') ;
  lStLiaison := lLiensEtab.GetString('CLE_COMPTEDEST') ;

  // Ligne sur le compte saisie
  lTobEtab   := lPieceEtab.newRecord ;
  lIdxLigne  := lPieceEtab.CurIdx ;
  lNumLigne  := lTobEtab.GetInteger('E_NUMLIGNE') ;
  lTobEtab.Dupliquer( lTobEcr, True, True ) ;
  lTobEtab.PutValue( 'E_NUMLIGNE', lNumLigne ) ;
  lPieceEtab.PutValue( lIdxLigne, 'E_CONTREPARTIEGEN', lStLiaison ) ;
  lPieceEtab.PutValue( lIdxLigne, 'E_CONTREPARTIEAUX', '' ) ;
  if (FModeEche=meDeporte) and ( lTobEtab.GetValue( 'E_ECHE' ) = 'X' ) then
    lPieceEtab.GetMultiEche( lIdxLigne ).CopieEcheances( GetMultiEche( vIndex ) ) ;
  lPieceEtab.PutVentil( lIdxLigne, 'E_ETABLISSEMENT', lTobEtab.GetValue('E_ETABLISSEMENT') );
  lPieceEtab.PutVentil( lIdxLigne, 'E_CONFIDENTIEL',  lTobEtab.GetValue('E_CONFIDENTIEL') ) ;

  // gestion du multi-echéance...
  if EstMultiEche( vIndex ) then
    begin
    lMultiEcheEcr  := GetMultiEche( lTobEcr ) ;
    // recopie du détail si possible
    if Assigned(lMultiEcheEcr) and lPieceEtab.EstOkMultiEche( lIdxLigne ) then
      begin
      lMultiEcheEtab := lPieceEtab.GetMultiEche( lTobEtab ) ;
      if Assigned( lMultiEcheEtab ) then
        for lNumEche := 2 to lMultiEcheEcr.GetNbEche do
          begin
          lTobEche := lMultiEcheEcr.GetEche( lNumEche ) ;
          lMultiEcheEtab.NewEche ;
          lMultiEcheEtab.UpdateEche( lNumEche, lTobEche.GetString('E_MODEPAIE'), lTobEche.GetDouble('E_DEBITDEV')+lTobEche.GetDouble('E_CREDITDEV'), lTobEche.GetDateTime('E_DATEECHEANCE'));
          end ;
      end ;
    // la ligne originale de sera plus multieche,...
    // pas necessaire de la faire à la main puisque le putValue plus bas sur E_GENERAL va réinitialiser les échéances
    end ;

  // Ligne sur le compte de liaison
  lTobEtab   := lPieceEtab.newRecord ;
  lIdxLigne  := lPieceEtab.CurIdx ;
  lNumLigne  := lTobEtab.GetInteger('E_NUMLIGNE') ;
  lTobEtab.Dupliquer( lTobEcr, True, True ) ;
  lTobEtab.PutValue( 'E_NUMLIGNE', lNumLigne ) ;
  lPieceEtab.PutValue( lIdxLigne,  'E_GENERAL', lStLiaison ) ;
  lPieceEtab.PutValue( lIdxLigne, 'E_CONTREPARTIEGEN', lStCpt ) ;
  lPieceEtab.PutValue( lIdxLigne, 'E_CONTREPARTIEAUX', lStAux ) ;
  lPieceEtab.PutVentil( lIdxLigne, 'E_ETABLISSEMENT', lTobEtab.GetValue('E_ETABLISSEMENT') );
  lPieceEtab.PutVentil( lIdxLigne, 'E_CONFIDENTIEL',  lTobEtab.GetValue('E_CONFIDENTIEL') );
  if lTobEtab.GetNumChamp('ESTECHEORIG')>0 then // Ajout pour gestion auto des enca/deca
    lTobEtab.PutValue('ESTECHEORIG', '-') ;

  // Réaffectation des montants
  lPieceEtab.AttribSolde( lIdxLigne ) ;

  // Finition
  lPieceEtab.RenumeroteLignes ;

  // ---------------------------------------------------
  // --- MAJ de la ligne sur établissement d'origine ---
  // ---------------------------------------------------
  PutValue( vIndex,  'E_ETABLISSEMENT',  GetEnteteS('E_ETABLISSEMENT') ) ;
  PutValue( vIndex,  'E_GENERAL',        lLiensEtab.GetString('CLE_COMPTEORIG') ) ;
  if lTobEcr.GetNumChamp('ESTECHEORIG') > 0 then // Ajout pour gestion auto des enca/deca
    lTobEcr.PutValue('ESTECHEORIG', '-') ;

end;

function TPieceCompta.GetPieceEtab( vEtabDest : String ) : TPieceCompta ;
var i : integer ;
begin

  // Recherche dans la TList si une pièce a déjà été créée pour cet établissement
  i := FPiecesEtab.IndexOf( vEtabDest ) ;
{
  for i := 0 to FPiecesEtab.Count - 1 do
    begin
    if TPieceCompta( FPiecesEtab[ i ] ).GetEnteteS('E_ETABLISSEMENT') = vEtabDest then
      begin
      result := TPieceCompta( FPiecesEtab[ i ] ) ;
      Exit ;
      end ;
    end ;
}
  // Sinon création d'une nouvelle pièce
  if i<>-1 then
    result := TPieceCompta( FPiecesEtab.Objects[ i ] ) 
//  if result = nil then
  else
    begin
    result := TPieceCompta.CreerPiece( Info ) ;
    result.SetMultiEcheMulti ;
    result.Entete.Dupliquer( FEntete, False, True ) ;
    result.PutEntete('E_ETABLISSEMENT', vEtabDest ) ;
    result.PutEntete('E_NUMEROPIECE',   0 ) ;
    result.videPiece ;

    FPiecesEtab.AddObject( vEtabDest, result ) ;

    end ;

end;



procedure TPieceCompta.PutEche(vIndex: Integer; vChamps: String; vValeur : Variant );
var lMultiEche  : TMultiEche ;
    lTobEcr     : TOB ;
begin

  if not (FModeEche in [meMulti, meDeporte]) then Exit ;

  lTobEcr := GetTob( vIndex ) ;
  if lTobEcr = nil then Exit ;

  if lTobEcr.GetValue('E_ECHE') <> 'X' then Exit ;

  // Récupération du TMultiEche
  lMultiEche := GetMultiEche ( vIndex ) ;
  if lMultiEche = nil then Exit ;
  if lMultiEche.NbEche < 2 then Exit ;

  // Report des valeurs
  lMultiEche.reporteEche( vChamps, vValeur );

//  if vChamps =

end;

procedure TPieceCompta.PutVentil( vIndex: Integer ; vChamps: String ; vValeur : Variant ) ;
var lTobEcr     : TOB ;
begin

  lTobEcr := GetTob( vIndex ) ;

  if lTobEcr = nil then Exit ;
  if lTobEcr.GetValue('E_ANA') <> 'X' then Exit ;

  if ((vChamps = 'E_DEBITDEV') or (vChamps = 'E_CREDITDEV')) and not (Contexte.CPAnaSurMontant)
    then RecalculProrataAnalNEW('Y',lTobEcr,0,Devise)
    else CReporteVentil( lTobEcr, vChamps, vValeur ) ;

end;

procedure TPieceCompta.GetTotauxPourChamps(vChamps, vValeur: String; var vTotalDebit, vTotalCredit: Double ; vInGroupe : integer) ;
var i      : Integer ;
    lInDe  : integer ;
    lInA   : integer ;
begin
  vTotalDebit  := 0 ;
  vTotalCredit := 0 ;

  if vInGroupe >= 1 then
    GetBornesGroupe( vInGroupe, lInDe, lInA )
  else
    begin
    lInDe  := 1 ;
    lInA   := Detail.Count ;
    end ;

  for i := lInDe to lInA do
    if GetValue( i, vChamps ) = vValeur then
      begin
      vTotalDebit  := vTotalDebit  + GetValue( i, 'E_DEBIT' ) ;
      vTotalCredit := vTotalCredit + GetValue( i, 'E_CREDIT' ) ;
      end ;
  // FQ 17129 : Pb d'arrondi Delphi
  vTotalDebit  := Arrondi( vTotalDebit  , V_PGI.OkDecV ) ;
  vTotalCredit := Arrondi( vTotalCredit , V_PGI.OkDecV ) ;

end;

function TPieceCompta.EstDivLett(vIndex: Integer): Boolean;
var lTob : Tob ;
begin
  result := False ;

  lTob := GetTob( vIndex ) ;
  if lTob = nil then Exit ;

  if (lTob.GetString('E_GENERAL') <> '')  and Info_LoadCompte( lTob.GetString('E_GENERAL') ) then
    result := ( Compte_GetValue('G_NATUREGENE') = 'DIV' ) and ( Compte_GetValue('G_LETTRABLE') = 'X' ) ;

end;

function TPieceCompta.ChargeScenario ( vIndex : integer ) : Boolean ;
var lTob   : Tob ;
    lStNat : string ;
begin

  result := False ;
  if not FBoScenActif then Exit ;

  // Minimum syndical = Journal + nature
  if GetEnteteS('E_JOURNAL') = '' then Exit ;

  if ModeSaisie=msPiece then
    begin
    lStNat := GetEnteteS('E_NATUREPIECE') ;
    end
  else
    begin
    lTob := GetTob( vIndex ) ;
    if lTob = nil then Exit ;
    lStNat := lTob.GetString('E_NATUREPIECE') ;
    end ;

  if lStNat = '' then Exit ;

  result := Contexte.ChargeScenario ( GetEnteteS('E_JOURNAL'), lStNat, GetEnteteS('E_QUALIFPIECE'), GetEnteteS('E_ETABLISSEMENT') ) ;

end;

procedure TPieceCompta.TermineLaPiece;
var lBoCPOk : Boolean ;
    lInIdx  : Integer ;
    lTobEcr : Tob     ;  // FQ 22974 29/05/2008 SBO
begin

  // ==> Début Modif // FQ 22974 29/05/2008 SBO
  lBoCPOk := True ;
  // Finalisation des lignes
  for lInIdx := 1 to Detail.Count do
    begin
    lTobEcr := CGetTob( lInIdx ) ;
    // --- Test pour Affectation des contreparties
    lBoCPOk := lBOCPOk and ( lTobEcr.GetString( 'E_CONTREPARTIEGEN' ) <> '' ) ;
    // --- Remise à 0 du E_NUMGROUPEECR pour le mode pièce // FQ 22974 27/05/2008 SBO
    if ModeSaisie=msPiece then
      lTobEcr.PutValue('E_NUMGROUPEECR', 0)
    else if ModeSaisie = msLibre then
      lTobEcr.PutValue('E_NUMGROUPEECR', 1) ;
    // --- MAJ du E_IO pour les échanges
    lTobEcr.PutValue('E_IO', 'X') ;
    end ;
  // ==> Fin Modif // FQ 22974 29/05/2008 SBO
  if not lBoCPOk then
    CAffectCompteContrePartie( self, Info ) ;

  // Equilibrage hors mode BOR
  if ModeSaisie <> msBOR then
    VerifieEquilibre ;

  //  Traitement mode piece
  if ModeSaisie=msPiece then
    begin
    // Scénario -> Ecriture validée ?
    if ChargeScenario and ( GetScenario.GetValue('SC_VALIDE') = 'X' ) then
      PutValueAll( 'E_VALIDE', 'X' ) ;
    end ;

    // Numero Définitif.
    AttribNumeroDef ;

end;

function TPieceCompta.GetScenario ( vIndex : integer ) : TOB ;
begin
  result := nil ;
  if not FBoScenActif then Exit ;
  if ChargeScenario( vIndex ) then
    result := Contexte.Scenario ;
end;

function TPieceCompta.IsValidTaux ( vIndex : Integer ) : Boolean;
var lDtDateC  : TDateTime ;
    lDtTaux   : TDateTime ;
    lStAlerte : String ;
    lBoPbTaux : boolean ;
begin
  Result := True ;
//  if ModeSaisie <> msPiece then Exit ;

  if ( ( Devise.Code = V_PGI.DevisePivot ) or
       ( Devise.Code = V_PGI.DeviseFongible )  ) then Exit ;

  // Récup date comptable + date taux dev
  if not isOut( vIndex )
    then lDtDateC := GetDateTime( vIndex, 'E_DATECOMPTABLE')
    else lDtDateC := GetEnteteDt('E_DATECOMPTABLE') ;
  lDtTaux   := GetRDevise( vIndex ).DateTaux ;

  // test taux
  lBoPbTaux := ( Devise.Taux = 1 ) ;
{  if ( ( lDtDateC  < V_PGI.DateDebutEuro ) or ( Not EstMonnaieIn( Devise.Code ) ) )
      then lBoPbTaux := ( Devise.Taux = 1 )
      else if EstMonnaieIn( Devise.Code )
           then lBoPbTaux := ( Devise.Taux = V_PGI.TauxEuro ) ;
}

  if (lDtTaux = lDtDateC) and not lBoPbTaux then Exit ;

  // Test alerte scénario
  if ChargeScenario then
    begin
    lStAlerte := GetScenario.GetValue('SC_ALERTEDEV') ;
    if ( ( lStAlerte = 'MOI' ) and ( GetPeriode( lDtTaux ) = GetPeriode( lDtDateC ) ) ) or // Au mois
       ( ( lStAlerte = 'SEM' ) and ( NumSemaine( lDtTaux ) = NumSemaine( lDtDateC ) ) )    // A la semaine
       then exit ;
    end ;

  result := False ;

{
}
end;

procedure TPieceCompta.majDeviseTaux( vIndex : integer = 0 ; vBoEnBase : boolean = False ) ;
var lDtDate : TDateTime ;
begin

  if FDevise.Code = V_PGI.DevisePivot then Exit ;

//  FDevise.Taux := GetTaux( FDevise.Code, FDevise.DateTaux, GetEnteteDt('E_DATECOMPTABLE'), Dossier ) ;
  // MAJ de la devise
  if (modeSaisie <> msPiece) and not isOut( vIndex )
    then lDtDate := GetDateTime( vIndex, 'E_DATECOMPTABLE' )
    else lDtDate := GetEnteteDT( 'E_DATECOMPTABLE' ) ;

  FDevise.Taux   := Info.Devise.GetTauxReel( lDtDate, FDevise.DateTaux, vBoEnBase ) ;
  Info.Devise.Item.PutValue( 'TAUX'     , FDevise.Taux     ) ;
  Info.Devise.Item.PutValue( 'DATETAUX' , FDevise.DateTaux ) ;

//  if vBoForce or ( lDtDate <> Info.Devise.Item.GetDateTime('DATETAUX') ) then
{  if vBoForce or ( lDtDate <> Info.Devise.Item.GetDateTime('DATETAUX') ) then
    begin
    FDevise.Taux := GetTaux( FDevise.Code, FDevise.DateTaux, lDtDate, Contexte.Dossier ) ;
    Info.Devise.Item.PutValue( 'TAUX'     , FDevise.Taux     ) ;
    Info.Devise.Item.PutValue( 'DATETAUX' , lDtDate ) ;
    end ;
}
end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 02/09/2005
Modifié le ... :   /  /
Description .. : Fusionne les lignes de Tva.
Suite ........ : Utilisé après la génération de la TVA détaillée, suivant 
Suite ........ : valeur de la zone "contrôle de TVA" du scénario.
Suite ........ : 
Suite ........ : Reprise de saisie.pas
Suite ........ : En attente de réactualisation
Mots clefs ... :
*****************************************************************}
procedure TPieceCompta.RegroupeTVA( vIndex : integer );
Var LigTVA    : integer ;
    LigNext   : integer ;
    CpteTVA   : String17 ;
    DTVA      : Double ;
    CTVA      : Double ;
    DNext     : Double ;
    CNext     : Double ;
    TobTva    : TOB ;
    TobNext   : Tob ;
begin

  LigTVA := 0 ;

  Repeat
    Inc( LigTva ) ;
    TobTva := TrouveLigneTVA( vIndex, LigTVA ) ;

    if TobTva <> nil then
      begin

      CpteTVA := TobTva.GetString(  'E_GENERAL'   ) ;
      DTVA    := TobTva.GetDouble(  'E_DEBITDEV'  ) ;
      CTVA    := TobTva.GetDouble(  'E_CREDITDEV' ) ;
      repeat
        LigTva  := TobTva.GetIndex + 1 ;
        TobNext := TrouveLigneCompte( CpteTVA, LigTVA + 1 ) ;
        if TobNext <> nil then
          begin
          LigNext := TobNext.GetIndex + 1 ;
          DNext   := TobNext.GetDouble(  'E_DEBITDEV'  ) ;
          CNext   := TobNext.GetDouble(  'E_CREDITDEV' ) ;
          DTVA    := DTVA + DNext ;
          CTVA    := CTVA + CNext ;

          DeleteRecord ( LigNext ) ;

          end ;
      Until ( TobNext = nil ) ;

      // Réaffectation des montants
      CSetMontants( TobTva, 0, 0,  Devise, True ) ;
      if Abs( DTVA ) > Abs( CTVA )
        then PutValue( LigTva, 'E_DEBITDEV',  DTVA-CTVA )
        else PutValue( LigTva, 'E_CREDITDEV', CTVA-DTVA ) ;

      end ;

  Until ( TobTVA = nil ) ;

  // Finition
//  RenumeroteLignes ;

end;

function TPieceCompta.IsValidTva ( vIndex : Integer ) : Boolean;
Var LigHT : integer ;
    NbHT  : integer ;
    Okok  : boolean ;
    TobHT : TOB ;
begin

  Result := True ;

  if not ChargeScenario(vIndex) then Exit ;
  if ( GetScenario(vIndex).GetValue('SC_CONTROLETVA') <> 'OUI' ) then Exit ;

  LigHT := 0 ;
  NbHT  := 0 ; //Okok:=True ;

  Okok:=True ;
  Repeat
    TobHT := TrouveLigneHT( LigHT + 1 ) ;
    if TobHT <> nil then
      begin
      LigHT := TobHT.GetIndex ;
      Inc(NbHT) ;
      Okok := ( ControleTva( LigHt ) = TVA_PASERREUR ) ;
      end ;
  Until ( (TobHT = nil) or (Not Okok) ) ;

  result := (NbHT > 0) and Okok ;

{
  // Vérification TVA :
  if not IsTvaOk then
    NotifyError( 0, 'La TVA de la pièce est incorrecte. Confirmez-vous la validation de cette pièce ? ' ) ;
}

end;

procedure TPieceCompta.ChargeVentil;
begin
  ChargeAnalytique( self, Dossier ) ;
end;

procedure TPieceCompta.InitPiece ( vStJournal : String ; vDtDateComptable : TDateTime ; vStNaturePiece : string = '' ;
                                     vStDevise : string = '' ; vStEtablissement : string = '' ; vStQualifPiece : string = '' ;
                                     vStEcrANO : String = '') ; 
begin

  PutEntete( 'E_JOURNAL',        vStJournal ) ;
  PutEntete( 'E_DATECOMPTABLE',  vDtDateComptable ) ;

  if vStNaturePiece<>'' then
    PutEntete( 'E_NATUREPIECE',    vStNaturePiece ) ;

  if vStDevise <> '' then
    PutEntete( 'E_DEVISE',         vStDevise ) ;

  if vStEtablissement<>'' then
    PutEntete( 'E_ETABLISSEMENT',  vStEtablissement ) ;

  if vStQualifPiece<>'' then
    PutEntete( 'E_QUALIFPIECE',    vStQualifPiece ) ;

  if vStEcrANO<>'' then
    PutEntete( 'E_ECRANOUVEAU',    vStEcrANO ) ;

end;

procedure TPieceCompta.InitPiece(M: RMVT);
begin
  InitPiece( M.Jal, M.DateC, M.Nature, M.CodeD, M.Etabl, M.Simul, 'N');
end;

procedure TPieceCompta.ChargeEcrCompl;
var lQ       : TQuery ;
    lTobEcr  : TOB ;
    lInIndex : integer ;
    lStSQL   : string ;
begin

  if Action = taCreat then Exit ;

  FEcrCompl.ClearDetail ;

  lStSQL := 'SELECT * FROM ECRCOMPL WHERE EC_JOURNAL="'   + GetEnteteS('E_JOURNAL') + '"'
                                  + ' AND EC_EXERCICE="'  + GetEnteteS('E_EXERCICE') + '"' ;

  if ModeSaisie=msPiece
    then  lStSQl := lStSQL + ' AND EC_DATECOMPTABLE="'    + UsDateTime( GetEnteteDt('E_DATECOMPTABLE') ) + '"'
    else  lStSQl := lStSQL + ' AND EC_DATECOMPTABLE>="'   + USDateTime( DebutDeMois( GetEnteteDt('E_DATECOMPTABLE') ) )+'"'
                           + ' AND EC_DATECOMPTABLE<="'   + USDateTime( FinDeMois(   GetEnteteDt('E_DATECOMPTABLE') ) )+'"' ;

  lStSQl := lStSql + ' AND EC_NUMEROPIECE='               + GetEnteteS('E_NUMEROPIECE')
                   + ' AND EC_QUALIFPIECE="'              + GetEnteteS('E_QUALIFPIECE') + '"'
                   + ' ORDER BY EC_JOURNAL, EC_EXERCICE, EC_DATECOMPTABLE, EC_NUMEROPIECE, EC_NUMLIGNE, EC_NUMECHE' ;

  lQ := OpenSelect( lStSQL, Dossier ) ;
  while not lQ.EOF do
    begin
    lInIndex := lQ.FindField('EC_NUMLIGNE').AsInteger ;
    lTobEcr  := GetTob( lInIndex ) ;
    if lTobEcr = nil then break;
    // CUTOFF indicateur de saisie du cutoff
{    if lTobEcr.GetNumChamp('CUTOFF') < 0
      then lTobEcr.AddChampSupValeur('CUTOFF','X')
      else lTobEcr.PutValue('CUTOFF', 'X') ;
}    CCreateDBTOBCompl(lTobEcr, FEcrCompl, lQ) ;
    lQ.Next ;
    end ;

  ferme(lQ) ;
{ VALEUR
Suite ........ : '-' : cutoff a saisir
Suite ........ : 'X' : cutoff saisie ( mais la date peu etre egale a idate1900 )
Suite ........ : '1' : reforce la saisie du cutoff ( qd on fait alt+C sur un
Suite ........ : compte de cutoff )
Suite ........ : '0' : cutoff saisie mais annulé ( ne sera plus reproposé )
Suite ........ : 'XD' : cutoff saisie et date <> iDate1900
}
end;

{JP 05/06/06 : Est-ce une pièce succeptible d'être rattachée à un BAP
{---------------------------------------------------------------------------------------}
function TPieceCompta.IsPieceABap : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := False;
  // SBO 17/12/2007 FQ 19174 maj de l'attribut GestionBapOk si multi-tiers pour éviter l'ouverture de la génération
  GestionBapOk := False;
  if ModeSaisie <> msPiece then Exit ;
  if not UnSeulTTC then Exit ;
{$IFDEF COMPTA}
{$IFNDEF CMPGIS35}
  if (GetEnteteS('E_QUALIFPIECE') = 'N') and
     GestionBap(GetEnteteS('E_NATUREPIECE'), GetEnteteS('E_JOURNAL'), Dossier) then begin
    GestionBapOk := True;
    Result := True;
  end;
{$ENDIF}
{$ENDIF COMPTA}
end;

{---------------------------------------------------------------------------------------}
procedure TPieceCompta.GestionDesBAP;
{---------------------------------------------------------------------------------------}
begin
  {Si on gère les BAP ...}
{$IFDEF COMPTA}
{$IFNDEF CMPGIS35}
  if IsPieceABap then begin
    {... On regarde s'il existe un BAP validé définitivement sur la Pièce ...}
    if ExisteSQL('SELECT BAP_STATUTBAP FROM ' + GetTableDossier(Dossier, 'CPBONSAPAYER') + ' WHERE ' +
                 'BAP_STATUTBAP = "' + sbap_Definitif + '" AND ' +
                 'BAP_JOURNAL = "' + VarToStr(GetEnteteS('E_JOURNAL')) + '" AND ' +
                 'BAP_EXERCICE = "' + VarToStr(GetEnteteS('E_EXERCICE')) + '" AND ' +
                 'BAP_DATECOMPTABLE = "' + UsDateTime( GetEnteteDt('E_DATECOMPTABLE') ) + '" AND ' +
                 'BAP_NUMEROPIECE = ' + GetEnteteS('E_NUMEROPIECE') + ' ') then
      Action := taConsult;
  end;
{$ENDIF}
{$ENDIF COMPTA}
end;


{Lek 10/04/06  b Charge extension écriture}
procedure TPieceCompta.ChargeExtension;
begin
  ChargeEcrCompl;
end;
{Lek 10/04/06  e Charge extension écriture}

function TPieceCompta.GetTobCompl( vIndex : Integer ): TOB;
var lTobEcr : TOB ;
begin
  result := nil ;

  lTobEcr := GetTob( vIndex ) ;
  if lTobEcr = nil then Exit ;

  result := CGetTOBCompl( lTobEcr ) ;
  if result = nil then
    result := CCreateTOBCompl( lTobEcr, FEcrCompl ) ;

end;

procedure TPieceCompta.CreateFromTOB(vTob: TOB);
var lTobLigne : Tob ;
    lTobEcr   : Tob ;
    i         : Integer ;
    lNumLigne : Integer ;
    procedure _OnError ( vStChamps : String ) ;
      begin
      if vTob.GetNumChamp ('ERRCHAMPS') < 0
        then vTob.AddChampSupValeur ('ERRCHAMPS', vStChamps )
        else vTob.SetString         ('ERRCHAMPS', vStChamps ) ;
      end ;
    procedure _AffecteEntete ( vStChamps : String ; vBoBloquant : Boolean ) ;
      begin
      if lTobEcr.GetValue( vStChamps ) <> ''
        then PutEntete( vStChamps, lTobEcr.GetValue( vStChamps ) )
        else if vBoBloquant then
             begin
             _OnError( vStChamps ) ;
             Exit ;
             end ;
      end ;
begin

  // Au moins 2 lignes d'écritures...
  if vTob.Detail.count < 2 then Exit ;

  // Récupération des infos générale (journal, devise, date comptable, ...)
  lTobEcr := vTob.Detail[ 0 ] ;
  _AffecteEntete( 'E_JOURNAL',       True  ) ;
  _AffecteEntete( 'E_DEVISE',        False ) ;
  _AffecteEntete( 'E_DATECOMPTABLE', False ) ;
  _AffecteEntete( 'E_NATUREPIECE',   False ) ;
  _AffecteEntete( 'E_ETABLISSEMENT', False ) ;

  // Intégration des lignes
  ClearDetail ;
  for i := 0 to vTob.Detail.Count - 1 do
    begin
    lTobEcr := vTob.Detail[ i ] ;
    lTobLigne := NewRecord ;
    lNumLigne := lTobLigne.GetInteger('E_NUMLIGNE') ;

    // Général
    if lTobEcr.GetValue( 'E_GENERAL' ) <> ''
      then PutValue( lNumLigne, 'E_GENERAL', lTobEcr.GetValue('E_GENERAL') )
      else begin
           _OnError( 'E_GENERAL' ) ;
           Exit ;
           end ;

    // Auxilaire
    if Compte_GetValue('G_COLLECTIF')='X' then
      begin
      if lTobEcr.GetValue( 'E_AUXILIAIRE' ) <> ''
        then PutValue( lNumLigne, 'E_AUXILIAIRE', lTobEcr.GetValue('E_AUXILIAIRE') )
        else begin
             _OnError( 'E_AUXILIAIRE' ) ;
             Exit ;
             end ;
        end ;

    // Montants
    if lTobEcr.GetDouble('E_DEBITDEV') <> 0
      then PutValue( lNumLigne, 'E_DEBITDEV', lTobEcr.GetValue('E_DEBITDEV') )
      else if lTobEcr.GetDouble('E_CREDITDEV') <> 0
        then PutValue( lNumLigne, 'E_CREDITDEV', lTobEcr.GetValue('E_CREDITDEV') )
        else _OnError( 'E_DEBITDEV' ) ;

    // Recopie Infos complémentaires
    RecopieInfoFromTob ( lNumLigne, lTobEcr ) ;

    end ;

end;

procedure TPieceCompta.RecopieInfoFromTob(vNumLigne: Integer; vTob: TOB);
var lTobLigne : Tob ;
begin
    lTobLigne := GetTob( vNumLigne ) ;
    if lTobLigne = nil then Exit ;

    lTobLigne.PutValue( 'E_REFINTERNE',        vTob.GetValue('E_REFINTERNE') ) ;
    lTobLigne.PutValue( 'E_LIBELLE',           vTob.GetValue('E_LIBELLE') ) ;

    lTobLigne.PutValue( 'E_QTE1',              vTob.GetValue('E_QTE1') ) ;
    lTobLigne.PutValue( 'E_QTE2',              vTob.GetValue('E_QTE2') ) ;
    lTobLigne.PutValue( 'E_QUALIFQTE1',        vTob.GetValue('E_QUALIFQTE1') ) ;
    lTobLigne.PutValue( 'E_QUALIFQTE2',        vTob.GetValue('E_QUALIFQTE2') ) ;

    lTobLigne.PutValue( 'E_REFEXTERNE',        vTob.GetValue('E_REFEXTERNE') ) ;
    lTobLigne.PutValue( 'E_DATEREFEXTERNE',    vTob.GetValue('E_DATEREFEXTERNE') ) ;

    lTobLigne.PutValue( 'E_REFLIBRE',          vTob.GetValue('E_REFLIBRE') ) ;
    lTobLigne.PutValue( 'E_AFFAIRE',           vTob.GetValue('E_AFFAIRE') ) ;
    lTobLigne.PutValue( 'E_LIBREDATE',         vTob.GetValue('E_LIBREDATE') ) ;
    lTobLigne.PutValue( 'E_CONSO',             vTob.GetValue('E_CONSO') ) ;

    lTobLigne.PutValue( 'E_LIBRETEXTE0',       vTob.GetValue('E_LIBRETEXTE0') ) ;
    lTobLigne.PutValue( 'E_LIBRETEXTE1',       vTob.GetValue('E_LIBRETEXTE1') ) ;
    lTobLigne.PutValue( 'E_LIBRETEXTE2',       vTob.GetValue('E_LIBRETEXTE2') ) ;
    lTobLigne.PutValue( 'E_LIBRETEXTE3',       vTob.GetValue('E_LIBRETEXTE3') ) ;
    lTobLigne.PutValue( 'E_LIBRETEXTE4',       vTob.GetValue('E_LIBRETEXTE4') ) ;
    lTobLigne.PutValue( 'E_LIBRETEXTE5',       vTob.GetValue('E_LIBRETEXTE5') ) ;
    lTobLigne.PutValue( 'E_LIBRETEXTE6',       vTob.GetValue('E_LIBRETEXTE6') ) ;
    lTobLigne.PutValue( 'E_LIBRETEXTE7',       vTob.GetValue('E_LIBRETEXTE7') ) ;
    lTobLigne.PutValue( 'E_LIBRETEXTE8',       vTob.GetValue('E_LIBRETEXTE8') ) ;
    lTobLigne.PutValue( 'E_LIBRETEXTE9',       vTob.GetValue('E_LIBRETEXTE9') ) ;

    lTobLigne.PutValue( 'E_TABLE0',            vTob.GetValue('E_TABLE0') ) ;
    lTobLigne.PutValue( 'E_TABLE1',            vTob.GetValue('E_TABLE1') ) ;
    lTobLigne.PutValue( 'E_TABLE2',            vTob.GetValue('E_TABLE2') ) ;
    lTobLigne.PutValue( 'E_TABLE3',            vTob.GetValue('E_TABLE3') ) ;

    lTobLigne.PutValue( 'E_LIBREBOOL0',        vTob.GetValue('E_LIBREBOOL0') ) ;
    lTobLigne.PutValue( 'E_LIBREBOOL1',        vTob.GetValue('E_LIBREBOOL1') ) ;

    lTobLigne.PutValue( 'E_LIBREMONTANT0',     vTob.GetValue('E_LIBREMONTANT0') ) ;
    lTobLigne.PutValue( 'E_LIBREMONTANT1',     vTob.GetValue('E_LIBREMONTANT1') ) ;
    lTobLigne.PutValue( 'E_LIBREMONTANT2',     vTob.GetValue('E_LIBREMONTANT2') ) ;
    lTobLigne.PutValue( 'E_LIBREMONTANT3',     vTob.GetValue('E_LIBREMONTANT3') ) ;

end;


class function TPieceCompta.CreerPiece( vInfoEcr : TInfoEcriture = nil ) : TPieceCompta;
begin
  result := TPieceCompta.Create( 'V_TOBPIECE', nil, -1 ) ;

  {JP 12/05/06 : Par défaut, si on modifie la pièce, on conserve le BAP attaché}
  Result.ConserverBap := True;

  {JP 20/07/07 : FQ 20601 : intialisation du numéro de ligne}
  Result.NumLigneAppel := 0;
  // Création de l'objet TInfoEcriture
  if Assigned( vInfoEcr )
    then result.Info := vInfoEcr
    else result.Info := TInfoEcriture.Create ;

  result.Initialisation ;

end;

procedure TPieceCompta.Initialisation ;
begin

//  Info_LoadDevise( V_PGI.DevisePivot ) ;
  FContexte := TPieceContexte.CreerContexte( Info.Dossier ) ;

  // Création de l'objet contenant les infos d'identiciation de la pièce
  FEntete := Contexte.InitEntetePiece ;

  // Création de l'objet contenant les infos d'identiciation de la pièce
  FEcheances := TOB.Create('V_ECHEANCES', nil, -1) ;

  // Tob de Sauvegarde des données d'entête
  FTobOrigine := TOB.Create('V_PIECEORIGINE', nil, -1 ) ;

  // Tob des données de chancellerie relative à la devise en cours
  FTobChancell := TOB.Create('V_CHANCELL', nil, -1 ) ;

  // Initialisation des infos devises
  Info_LoadDevise( V_PGI.DevisePivot ) ;

  // Par défaut, l'action est la création d'écriture courante
  FAction      := taCreat ;
  FBoModif     := False ;
  FModeSaisie  := msPiece ;
  FModeVentil  := mvDefaut ;
  FModeGroupe  := mgDynamic ;
  FBoScenActif := False ;
  FBoAccActif  := False ;
  FModeEche    := meMono ;
  FBoTPActif   := False;          // Indicateur de gestion des tiers payeurs
  FYaTP        := yaRien ;        // indicateur de présence d'une pièce de tiers payeurs :

  // List des pièces générées dans le cadre du multi-établissement
  FPiecesEtab := TStringList.Create ;
  FPiecesEtab.Duplicates := dupError ;
  FPiecesEtab.sorted     := True ;

  // Tob de données du cutoff
  FEcrCompl := TOB.Create('',nil,-1) ;

  // liste pour les TP
  FListeTP  := TStringList.create ;

end;


function TPieceCompta.GetDossier: String;
begin
  result := '' ;
  if not Assigned( Contexte ) then Exit ;
  result := Contexte.Dossier ;
end;

procedure TPieceCompta.SetVentilDefautOff;
begin
  FModeVentil  := mvAucun ;
end;

procedure TPieceCompta.SetMultiEcheOff;
begin
  FModeEche := meMono ;
end;

procedure TPieceCompta.AffecteDateValeur(vIndex: Integer);
begin

  if not EstRemplit( vIndex ) then Exit ;
  if not Info_LoadCompte( GetValue(vIndex, 'E_GENERAL') ) then Exit ;

  // Comptes pointables --> Date comptable
  if Compte_GetValue('G_POINTABLE')='X' then
    PutValue(vIndex, 'E_DATEVALEUR', GetValue(vIndex, 'E_DATECOMPTABLE') )
   // Comptes lettrable  --> Date d'échéances
  else if (Compte_GetValue('G_LETTRABLE') = 'X') or
          ( ( GetValue(vIndex, 'E_AUXILIAIRE') <> '' ) and
            Info_LoadAux( GetString(vIndex, 'E_AUXILIAIRE') ) and
            ( Aux_GetValue('T_LETTRABLE') = 'X' ) ) then
     PutValue( vIndex, 'E_DATEVALEUR', GetValue( vIndex, 'E_DATEECHEANCE') )
   // Sinon Date 01/01/1900
   else
     PutValue( vIndex, 'E_DATEVALEUR', iDate1900 )

end;

procedure TPieceCompta.ExecuteFormule(vIndex: integer; vChamps, vFormule: string; vFromTob: Tob);
var lStVal  : Variant ;
    lOldIdx : integer ;
begin
  if isOut( vIndex ) then Exit ;
  FFromTob := vFromTob ;
  lOldIdx  := FCurIdx ;
  FCurIdx  := vIndex ;
  lStVal   := GFormule( vFormule, GetFormule, nil, 1 ) ;
  if pos( vChamps, 'E_REFINTERNE, E_REFEXTERNE, E_LIBELLE, E_REFLIBRE') > 0
    then PutValue( vIndex, vChamps, Copy( VarToStr(lStVal), 1, 35 )  )
    else PutValue( vIndex, vChamps, lStVal ) ;
  FFromTob := nil ;
  FCurIdx  := lOldIdx ;
end;


procedure TPieceCompta.SetVentilSurAttente;
begin
  FModeVentil  := mvAttente ;
end;

procedure TPieceCompta.AjustePieceMultiEtab;
var lStJal : string ;
    lStCpt : string ;
begin

  if not ( EstJalEffet or EstJalBqe ) then Exit ;

  // Recherche du compte de contrepartie
  if not Info_LoadJournal( GetEnteteS('E_JOURNAL') ) then Exit ;

  lStCpt := Journal_GetValue('J_CONTREPARTIE') ;
  if not Info_LoadCompte( lStCpt ) then Exit ;

  // Recherche compte de contrepartie
  if TrouveLigneCompte( lStCpt ) = nil then
    begin
    // recherche du journal de substitution
    lStJal := Contexte.JalMultiEtab ;
    if not info_LoadJournal( lStJal ) then Exit ;

    // affectation du nouveau journal
    PutEntete( 'E_JOURNAL', lStJal ) ;
    PutEntete( 'E_NUMEROPIECE', 0 ) ;
    end ;

end;

procedure TPieceCompta.ActiveAcc;
begin
  FBoAccActif := True ;
end;

function TPieceCompta.IsActiveAcc: boolean;
begin
  result := False ;
  if not FBoAccActif then Exit ;
  if not Info_LoadJournal( GetEnteteS('E_JOURNAL') ) then Exit ;
  if Journal_GetValue('J_ACCELERATEUR') <> 'X' then Exit ;
  result := True ;
end;

procedure TPieceCompta.ActiveTP;
begin
  FBoTPActif := True ;
end;

{$IFDEF COMPTA}
procedure TPieceCompta.GenerePiecesTP;
var lRMVT : RMVT ;
begin

  if not FBoTPActif then Exit ;
  lRMVT := GetRMVT ;

  if FYaTP=yaNL then
    SupprimePiecesTP ;

  GenerePiecesPayeur( lRMVT ) ;

end;
{$ENDIF COMPTA}

{$IFDEF COMPTA}
procedure TPieceCompta.ChargeTP;
var i          : integer ;
    lStTP      : string ;
    lStPieceTP : string ;
    lTob       : Tob ;
begin

  if not TestParamTP then Exit ;

  FListeTP.Clear ;

  // Chargement des réf aux pièces des tiers payeurs
  For i := 1 to Detail.Count do
    begin

    if FYaTP = yaL then break ;

    lStTP := GetValue(i, 'E_TIERSPAYEUR') ;
    if lStTP='' then continue ;

    lStPieceTP := GetValue(i, 'E_PIECETP') ;
    if lStPieceTP='' then continue ;

    FListeTP.Add( lStPieceTP ) ;

    // MAJ des indicateurs
    FYaTP := yaNL ;
    if ExisteLettrageSurTP( lStTP, lStPieceTP ) then
      FYaTP := yaL ;

    end ;

  // Maj du lettrage pour rendre la modification possible si la pièce sur TP n'est pas lettré
  if FYaTP = yaNL then
    For i := 1 to Detail.Count do
      begin
      lTob := GetTob( i ) ;
      if lTob.GetString('E_AUXILIAIRE')=''  then Continue ;
      if lTob.GetString('E_ECHE')<>'X'      then Continue ;
      if lTob.GetString('E_TIERSPAYEUR')='' then Continue ;
      if lTob.GetString('E_PIECETP')=''     then Continue ;

      lTob.PutValue('E_LETTRAGE',     '') ;
      lTob.PutValue('E_ETATLETTRAGE', 'AL') ;
      lTob.PutValue('E_COUVERTURE',    0) ;
      lTob.PutValue('E_COUVERTUREDEV', 0) ;
      lTob.PutValue('E_LETTRAGEDEV',   '-') ;
      lTob.PutValue('E_DATEPAQUETMIN', lTob.GetValue('E_DATECOMPTABLE') ) ;
      lTob.PutValue('E_DATEPAQUETMAX', lTob.GetValue('E_DATECOMPTABLE') ) ;

      end ;


  // important pour replacer index du TInfo sur journal de la pièce :
  Info_LoadJournal( GetEnteteS('E_JOURNAL') ) ;

end;
{$ENDIF COMPTA}

{$IFDEF COMPTA}
procedure TPieceCompta.SupprimePiecesTP;
var lStPieceTP : string ;
    i          : integer ;
begin
  for i:=0 to FListeTP.Count-1 do
    begin
    lStPieceTP := FListeTP[i] ;
    SupprimePieceTP( lStPieceTP ) ;
    end ;
end;
{$ENDIF COMPTA}

function TPieceCompta.EstJalFact: Boolean;
begin
  result := False ;
  if Info_LoadJournal( GetEnteteS('E_JOURNAL') ) then
    result := ( Journal_GetValue('J_NATUREJAL') = 'VTE' ) or ( Journal_GetValue('J_NATUREJAL') = 'ACH' )  ;
end;

function TPieceCompta.TestParamTP(vBoMess: boolean): boolean;
var lNature    : string ;
begin
  result := False ;
  if not FBoTPActif then Exit ;
  if FYaTP = yaL then Exit ;
  if FAction <> taModif then Exit ;
  if Not Contexte.OuiTP then Exit ;

  if GetEnteteS('E_QUALIFPIECE') <> 'N' then Exit ;
  if Entete.GetValue('E_ECRANOUVEAU')<>'N' then Exit ;
  if not EstJalFact then Exit ;

  lNature := GetEnteteS('E_NATUREPIECE') ;
  if (lNature<>'FC') and (lNature<>'FF') and (lNature<>'AC') and (lNature<>'AF') then Exit ;

  if ((lNature='FC') or (lNature='AC')) then
    result := ( Contexte.JalVTP <> '' ) and ( Info_LoadJournal( Contexte.JalVTP ) ) ;

  if ((lNature='FF') or (lNature='AF')) then
    result := ( Contexte.JalATP <> '' ) and ( Info_LoadJournal( Contexte.JalATP ) ) ;

  {$IFNDEF EAGLSERVER}
  if (not result) and vBoMess then
    PgiBox('Vous n''avez pas paramétré de journal pour les tiers payeurs, le mécanisme ne sera pas activé') ;
  {$ENDIF EAGLSERVER}

  // important pour replacer index du TInfo sur journal de la pièce
  Info_LoadJournal( GetEnteteS('E_JOURNAL') ) ;

end;

function TPieceCompta.GetFirstIdxTP: integer;
var i : integer ;
begin

  result := 0 ;

  for i := 1 to Detail.Count do
    if GetValue( i, 'E_AUXILIAIRE') <> '' then
      if Info_LoadAux( GetString( i, 'E_AUXILIAIRE') ) then
        if ( Aux_GetValue('T_PAYEUR')<>'' ) and ( Aux_GetValue('T_ISPAYEUR')<>'X' ) then
          begin
          result := i ;
          break ;
          end ;

end;

function TPieceCompta.RechGuidId( vIndex : Integer ) : string;
var lTOBEcr : TOB ;
begin

 result := '' ;

 Case ModeSaisie of
   msBor : lTOBEcr := GetTob( GetDebutGroupe( vIndex ) ) ;
   else    lTOBEcr := GetTob(1) ;
 end ;

 if lTOBEcr = nil then exit ;

 result := VarAsType( CGetValueTOBCompl(lTOBEcr,'EC_DOCGUID'), varString ) ;
 if result = #0 then result := '' ;

end;
{$IFDEF SCANGED}
procedure TPieceCompta.SupprimeLeDocGuid( vIndex : Integer ) ;
var lTOBCompl : TOB ;
    lTOBEcr   : TOB ;
begin

 Case ModeSaisie of
   msBor : lTOBEcr := GetTob( GetDebutGroupe( vIndex ) ) ;
   else    lTOBEcr := GetTob(1) ;
 end ;
 if lTOBEcr = nil then exit ;

 lTOBCompl := CGetTOBCompl(lTOBEcr) ;

 if lTOBCompl <> nil then
  begin
   if Length(lTOBCompl.GetValue('EC_DOCGUID')) > 1 then
    SupprimeDocumentGed(lTOBCompl.GetValue('EC_DOCGUID')) ;
   lTOBCompl.PutValue('EC_DOCGUID', '' ) ;
  end ;

end;

procedure TPieceCompta.AjouteGuidId( vIndex : Integer ; vGuidId: string ) ;
var lTOBCompl : TOB ;
    lTOBEcr   : TOB ;
begin

 Case ModeSaisie of
   msBor : lTOBEcr := GetTob( GetDebutGroupe( vIndex ) ) ;
   else    lTOBEcr := GetTob(1) ;
 end ;
 if lTOBEcr = nil then exit ;

 lTOBCompl := CGetTOBCompl(lTOBEcr) ;

 if lTOBCompl = nil then
  lTOBCompl := CCreateTOBCompl(lTOBEcr,FEcrCompl)
   else
    if Length(lTOBCompl.GetValue('EC_DOCGUID')) > 1 then
     SupprimeDocumentGed(lTOBCompl.GetValue('EC_DOCGUID')) ;

 lTOBCompl.PutValue('EC_DOCGUID', vGuidId) ;

end;
{$ENDIF SCANGED}

procedure TPieceCompta.SetScenarioOff;
begin
  FBoScenActif := False ;
end;

procedure TPieceCompta.SetScenarioOn;
begin
  FBoScenActif := True ;
end;

function TPieceCompta.Info_LoadCompte( vStCode : string ) : Boolean;
begin
  result := Info.LoadCompte( vStCode ) ;
end;

function TPieceCompta.Aux_GetValue(vStNom: string): Variant;
begin
  result := Info.GetValue( vStNom ) ;
end;

function TPieceCompta.Compte_GetValue(vStNom: string): Variant;
begin
  result := Info.GetValue( vStNom ) ;
end;

function TPieceCompta.Info_LoadAux(vStCode: string): Boolean;
begin
  result := Info.LoadAux( vStCode ) ;
end;

function TPieceCompta.Journal_GetValue(vStNom: string): Variant;
begin
  result := Info.Journal.GetValue( vStNom ) ;
end;

function TPieceCompta.Info_LoadJournal(vStCode: string): Boolean;
begin
  result := Info.LoadJournal( vStCode ) ;
end;

class function TPieceCompta.GetChampsAuxi: String;
begin
  result := 'T_AUXILIAIRE, T_LIBELLE, T_FERME, T_COLLECTIF, T_SOUMISTPF, '
          + 'T_CONFIDENTIEL, T_NATUREAUXI, T_LETTRABLE, T_REGIMETVA, '
          + 'T_TVAENCAISSEMENT, T_PAYEUR, T_ISPAYEUR, T_CONSO, '
          + 'T_ESCOMPTE, T_MODEREGLE, T_MULTIDEVISE, T_DEVISE, '
          + 'T_TOTALDEBIT, T_TOTALCREDIT, T_TOTDEBE, T_TOTDEBS, '
          + 'T_TOTDEBP, T_TOTCREE, T_TOTCRES, T_TOTCREP ' ;
end;

class function TPieceCompta.GetChampsGene: String;
begin
  result := 'G_GENERAL, G_LIBELLE, G_FERME, G_SENS, G_NATUREGENE, '
          + 'G_VENTILABLE, G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5, '
          + 'G_CONFIDENTIEL, G_POINTABLE, G_LETTRABLE, G_COLLECTIF, G_BUDGENE, '
          + 'G_QUALIFQTE1, G_QUALIFQTE2, G_SOUMISTPF, '
          + 'G_REGIMETVA, G_TVA, G_TPF, G_TVASURENCAISS ' ;
end;

procedure TPieceCompta.ChargeInfoOpti(vTobSource: Tob);
var i : integer ;
    lTob : Tob ;
begin
  For i := 0 to vTobSource.Detail.count - 1 do
    begin
    lTob := vTobSource.Detail[i];
    Info.LoadOpti( lTob.GetString('E_GENERAL'), lTob.GetString('E_AUXILIAIRE'), lTob ) ;
    end ;
end;


procedure TPieceCompta.InitNewRecord(TobEcr: TOB; vIndex: Integer; lBoARenumeroter: Boolean);
var lInGrp      : integer ;
    lInDeb      : integer ;
    lInFin      : integer ;
    lStNatPiece : string ;
    lDtDateC    : TDateTime ;
    lStDevise   : string ;
    i           : integer ;
    lStRef      : string ;
begin

  CSupprimerInfoLettrage( TobEcr ) ;

  // Données d'entête
  TobEcr.PutValue('E_SOCIETE',       Entete.GetValue('E_SOCIETE')       ) ;
  TobEcr.PutValue('E_JOURNAL',       Entete.GetValue('E_JOURNAL')       ) ;
  TobEcr.PutValue('E_ETABLISSEMENT', Entete.GetValue('E_ETABLISSEMENT') ) ;
  TobEcr.PutValue('E_NUMEROPIECE',   Entete.GetValue('E_NUMEROPIECE')   ) ;
  TobEcr.PutValue('E_QUALIFPIECE',   Entete.GetValue('E_QUALIFPIECE')   ) ;
  TobEcr.PutValue('E_ECRANOUVEAU',   Entete.GetValue('E_ECRANOUVEAU')   ) ;
  TobEcr.PutValue('E_MODESAISIE',    Entete.GetValue('E_MODESAISIE')    ) ;
  TobEcr.PutValue('E_REGIMETVA',     Entete.GetValue('E_REGIMETVA')     ) ;

  // détermination de la date
//lDtDateC := GetEnteteDt('E_DATECOMPTABLE') ;
{
  TobEcr.PutValue('E_DATECOMPTABLE', Entete.GetValue('E_DATECOMPTABLE') ) ;
  TobEcr.PutValue('E_EXERCICE',      Entete.GetValue('E_EXERCICE')      ) ;
  TobEcr.PutValue('E_SEMAINE',       Entete.GetValue('E_SEMAINE')       ) ;
  TobEcr.PutValue('E_PERIODE',       Entete.GetValue('E_PERIODE')       ) ;
}
  // Gestion spécif BOR / LIB / Piece :
  if ModeSaisie = msBOR then
    begin
    // 1ère ligne
    if Detail.count = 1 then
      begin
      PutValue( vIndex, 'E_NATUREPIECE',   GetNatureInitiale                           ) ;
//      PutValue( vIndex, 'E_DATECOMPTABLE', FinDeMois( GetEnteteDt('E_DATECOMPTABLE') )       ) ;
      lDtDateC := FinDeMois( GetEnteteDt('E_DATECOMPTABLE') ) ;
      PutValue( vIndex, 'E_NUMGROUPEECR',  1                                           ) ;
      PutValue( vIndex, 'E_DEVISE',        GetEnteteS('E_DEVISE')                      ) ;
      end
    // Récup des info du groupe en cours
    else
      begin
      lStDevise   :=  GetEnteteS('E_DEVISE') ;
      if vIndex > 1 then
        begin
        lStNatPiece := GetString( vIndex-1, 'E_NATUREPIECE'   ) ;
        lDtDateC    := GetDateTime( vIndex-1, 'E_DATECOMPTABLE' ) ;
       end
      else
        begin
        lStNatPiece := GetString( vIndex+1, 'E_NATUREPIECE'   ) ;
        lDtDateC    := GetDateTime( vIndex+1, 'E_DATECOMPTABLE' ) ;
        end ;

      // détermination du groupe
      if vIndex = Detail.count then
        begin
        lInGrp := GetNumGroupe( vIndex - 1 ) ;
        lInDeb := GetDebutGroupe( vIndex - 1 ) ;
        lInFin := vIndex ;
        if ( GetSoldePartiel( lInDeb, (vIndex - 1) ) = 0 )
          then begin
               Inc( lInGrp ) ;
               lInDeb := vIndex ;
               end ;
        PutValue( vIndex, 'E_NUMGROUPEECR', lInGrp ) ;
        end
      // Si insertion alors recup ligne au dessous
      else
        begin
        lInGrp := GetNumGroupe( vIndex + 1 ) ;
        PutValue( vIndex, 'E_NUMGROUPEECR', lInGrp ) ;
        GetBornesGroupe( lInGrp, lInDeb, lInFin ) ;
        end ;

      // autres champs relatifs aux groupes
      for i:=lInDeb to lInFin do
        begin
        if i = vIndex
          then Continue
          else begin
               lStNatPiece := GetString(i, 'E_NATUREPIECE') ;
               lDtDateC    := GetDateTime(i, 'E_DATECOMPTABLE') ;
               lStDevise   := GetString(i, 'E_DEVISE') ;
               break ;
               end ;
        end ;
      putValue( vIndex, 'E_NATUREPIECE',   lStNatPiece ) ;
//      putValue( vIndex, 'E_DATECOMPTABLE', lDtDateC ) ;
      putValue( vIndex, 'E_DEVISE',        lStDevise ) ;

      end ;
    end
  else if ModeSaisie = msLibre then
    begin
    PutValue( vIndex, 'E_NUMGROUPEECR',  1) ;
    PutValue( vIndex, 'E_DEVISE',        GetEnteteS('E_DEVISE') ) ;
    // 1ère ligne
    if Detail.count = 1 then
      begin
      PutValue(vIndex, 'E_NATUREPIECE',   GetNatureInitiale ) ;
      lDtDateC := FinDeMois( GetEnteteDt('E_DATECOMPTABLE') ) ;
      end
    // insertion intermédiaire
    else if (vIndex < Detail.count) then
      begin
      PutValue(vIndex, 'E_NATUREPIECE',   GetValue( vIndex + 1, 'E_NATUREPIECE')   ) ;
      lDtDateC := GetDateTime( vIndex + 1, 'E_DATECOMPTABLE') ;
      end
    // dernière ligne : Recup valeur precedente
    else
      begin
      PutValue(vIndex, 'E_NATUREPIECE',   GetValue( vIndex - 1, 'E_NATUREPIECE')   ) ;
      lDtDateC := GetDateTime( vIndex - 1, 'E_DATECOMPTABLE') ;
      end ;
//    PutValue( vIndex, 'E_DATECOMPTABLE', lDtDateC ) ;
    end
  else
    begin
    PutValue( vIndex,  'E_NATUREPIECE',   GetEnteteS('E_NATUREPIECE')    ) ;
    PutValue( vIndex,  'E_NUMGROUPEECR',  0                              ) ;
    PutValue( vIndex,  'E_DEVISE',        GetEnteteS('E_DEVISE')         ) ;
//    PutValue( vIndex,  'E_DATECOMPTABLE', GetEnteteDt('E_DATECOMPTABLE') ) ;
    lDtDateC := GetEnteteDt('E_DATECOMPTABLE') ;
    end ;

  // MAJ de ma date comptable !!!
  PutValue( vIndex, 'E_DATECOMPTABLE', lDtDateC, True ) ;

  // Si pas la 1ère ligne, reprendre la libellé et la référence interne de la ligne supérieure
  if vIndex > 1 then
    begin
    // Recopie du libellé si...
    if ( ModeSaisie = msPiece ) or      // mode pièce
       ( ( ModeSaisie = msBor ) and   // Mode bor/lib et même groupe
         ( GetInteger(vIndex, 'E_NUMGROUPEECR') = GetInteger(vIndex-1, 'E_NUMGROUPEECR') ) ) then
       begin
       TobEcr.PutValue('E_REFINTERNE', Detail[ vIndex - 2 ].GetString('E_REFINTERNE') ) ;
       TobEcr.PutValue('E_LIBELLE',    Detail[ vIndex - 2 ].GetString('E_LIBELLE') ) ;
       //test régime
       if GetInteger( vIndex-1, 'E_NUMGROUPEECR') = GetInteger( vIndex, 'E_NUMGROUPEECR') then
         TobEcr.PutValue('E_REGIMETVA',  Detail[ vIndex - 2 ].GetString('E_REGIMETVA') ) ;
       end
    else
      // Cas de l'incrémentation auto
      if ( ModeSaisie = msBor ) and
         ( GetInteger(vIndex, 'E_NUMGROUPEECR') <> GetInteger(vIndex-1, 'E_NUMGROUPEECR') ) then
        begin
        lStRef := GetString(vIndex - 2, 'E_REFINTERNE') ;
        if ( Info.Journal.GetValue('J_INCREF') = 'X' )
           then lStRef := GetIncRefAuto( lStRef )
           else lStRef := '' ;
        TobEcr.PutValue('E_REFINTERNE', lStRef ) ;
        end ;

    end ;

  // Init champs divers
  TobEcr.PutValue( 'E_IO' , 'X' ) ;
  TobEcr.PutValue('E_QUALIFQTE1', '' ) ;
  TobEcr.PutValue('E_QUALIFQTE2', '' ) ;

  // numérotation des lignes
  if lBoARenumeroter
    then RenumeroteLignes
    else if (vIndex = 1) or (Detail.count = 1)
      then TobEcr.PutValue('E_NUMLIGNE', 1 )
      else TobEcr.PutValue('E_NUMLIGNE', ValeurI( GetValue( vIndex-1, 'E_NUMLIGNE') ) + 1 ) ;

  // Champs de gestion
  TobEcr.AddChampSupValeur( 'ACCELERATEUR', '-' ) ;
  TobEcr.AddChampSupValeur( 'COMPS',        '-' ) ;
  TobEcr.AddChampSupValeur( 'CUTOFF',       '-' ) ;
  TobEcr.AddChampSupValeur( 'PBTAUX',       '-' ) ;

  // champs spécifique
  TobEcr.AddChampSupValeur( 'SECTIONA1', '' ) ;
  TobEcr.AddChampSupValeur( 'SECTIONA2', '' ) ;
  TobEcr.AddChampSupValeur( 'SECTIONA3', '' ) ;
  TobEcr.AddChampSupValeur( 'SECTIONA4', '' ) ;
  TobEcr.AddChampSupValeur( 'SECTIONA5', '' ) ;

  // Affectation des données d'entête de l'éventuel scénario
  AffecteEnteteScenario( vIndex ) ;

end;

procedure TPieceCompta.CumulerMontant(vNumLigneDst: Integer; PieceSrc: TPieceCompta; vNumLigneSrc: Integer);
  function IsSectionPresent(TobAxeDst, TobVentilSrc: TOB; var TobVentilDst: TOB): Boolean;
  var
    i:            Integer;
    iVentil:      Integer;
  begin
    Result := False;
    for iVentil:=0 to TobAxeDst.Detail.Count-1 do
      begin
      TobVentilDst := TobAxeDst.Detail[iVentil];
      Result := TobVentilSrc.GetValue('Y_SECTION') = TobVentilDst.GetValue('Y_SECTION');
      for i:= 1 to 6 do
        Result := Result and (TobVentilSrc.GetValue('Y_SOUSPLAN'+IntToStr(i)) = TobVentilDst.GetValue('Y_SOUSPLAN'+IntToStr(i)));
      if Result then
        break;
      end;
    if not Result then
      TobVentilDst := nil;
  end;

  function IsEcheancePresent(MultiEchDst, TobEchSrc: TOB; var TobEchDst: TOB): Boolean;
  var
    iEch:        Integer;
    DateEchSrc:  TDateTime;
    DateEchDst:  TDateTime;
  begin
    Result := False;
    if MultiEchDst = nil then
      Exit;
    for iEch:=0 to MultiEchDst.Detail.Count-1 do
      begin
      TobEchDst  := MultiEchDst.Detail[iEch];
      DateEchSrc := TobEchSrc.GetValue('E_DATEECHEANCE');
      DateEchDst := TobEchDst.GetValue('E_DATEECHEANCE');
      if (DateEchSrc = DateEchDst) and
         (TobEchSrc.GetValue('E_MODEPAIE')     = TobEchDst.GetValue('E_MODEPAIE')) then
        begin
        Result := True;
        break;
        end;
      end;
  end;

  procedure SetSensMontant(TobEcr: TOB; Prefix, SensEcr: String);
  begin
    if SensEcr = 'DEBIT' then
      begin
      TobEcr.PutValue(Prefix+'_DEBITDEV',  Arrondi(TobEcr.GetValue(Prefix+'_DEBITDEV')-TobEcr.GetValue(Prefix+'_CREDITDEV'),  Devise.Decimale));
      TobEcr.PutValue(Prefix+'_CREDITDEV', 0);
      end
    else
      begin
      TobEcr.PutValue(Prefix+'_DEBITDEV',  0);
      TobEcr.PutValue(Prefix+'_CREDITDEV', Arrondi(TobEcr.GetValue(Prefix+'_CREDITDEV')-TobEcr.GetValue(Prefix+'_DEBITDEV'),  Devise.Decimale));
      end;
  end;

  procedure CumulMntMultiEcheance(TobEchSrc: TOB; SensEcr: String);
  var
    MultiEchDst:  TMultiEche;
    TobEchDst:    TOB;
  begin
    MultiEchDst := GetMultiEche(vNumLigneDst);
    if IsEcheancePresent(MultiEchDst, TobEchSrc, TobEchDst) then
      begin
      TobEchDst.PutValue('E_DEBITDEV',  TobEchSrc.GetValue('E_DEBITDEV')+TobEchDst.GetValue('E_DEBITDEV'));
      TobEchDst.PutValue('E_CREDITDEV', TobEchSrc.GetValue('E_CREDITDEV')+TobEchDst.GetValue('E_CREDITDEV'));
      end
    else
      begin
      TobEchDst := Tob.Create('ECRITURE', Self, vNumLigneDst-1);
      TobEchDst.Dupliquer(TobEchSrc, True, True);
      InitNewRecord(TobEchDst, vNumLigneDst, False);
      TobEchDst.PutValue('E_NUMLIGNE', vNumLigneDst);
      TobEchDst.PutValue('E_DATEECHEANCE', TobEchSrc.GetValue('E_DATEECHEANCE'));
      AffecteEche(vNumLigneDst);
      TobEchDst.PutValue('E_NUMECHE', MultiEchDst.NbEche+1);
      MultiEchDst.AddEche(TobEchDst);
      end;
    SetSensMontant(TobEchDst, 'E', SensEcr);
  end;

  procedure CumulAnalytique(NewDebit, NewCredit: Double; SensEcr: String; var EffaceAnaSrc: Boolean);
  var
    Mnt:          Double;
    Pourcent:     Double;
    iAxe:         Integer;
    iVentil:      Integer;
    TobAxeSrc:    TOB;
    TobAxeDst:    TOB;
    TobVentilSrc: TOB;
    TobVentilDst: TOB;
  begin
    if (GetValue(vNumLigneDst, 'E_ANA') = 'X') and (PieceSrc.GetValue(vNumLigneSrc, 'E_ANA') <> 'X') then
      begin
      {Ecriture Dst avec analytique et Ecriture Src sans analytique}
      PieceSrc.PutValue(vNumLigneSrc, 'E_ANA', 'X');
      PieceSrc.SetVentilSurAttente;// Lek 260706 FBoAnaSurAtt := True;
      PieceSrc.AffecteAna(vNumLigneSrc);
      PieceSrc.FModeVentil:=mvDefaut; // Lek 260706 FBoAnaSurAtt := False;
      EffaceAnaSrc := True;
      end
    else if (GetValue(vNumLigneDst, 'E_ANA') <> 'X') and (PieceSrc.GetValue(vNumLigneSrc, 'E_ANA') = 'X') then
      begin
      {Ecriture Dst sans analytique et Ecriture Src avec analytique}
      PutValue(vNumLigneDst, 'E_ANA', 'X');
      SetVentilSurAttente; // Lek 260706 FBoAnaSurAtt := True;
      AffecteAna(vNumLigneDst);
      FModeVentil:=mvDefaut; // Lek 260706 FBoAnaSurAtt := False;
      end;

    if (GetValue(vNumLigneDst, 'E_ANA') = 'X') and (PieceSrc.GetValue(vNumLigneSrc, 'E_ANA') = 'X') then
      begin
      {Ecriture Dst avec analytique et Ecriture Src avec analytique}
      for iAxe:=0 to GetTob(vNumLigneDst).Detail.Count-1 do
        begin
        TobAxeSrc := PieceSrc.GetTob(vNumLigneSrc).Detail[iAxe];
        TobAxeDst := GetTob(vNumLigneDst).Detail[iAxe];

        {Ajoute ou modifie chaque section}
        for iVentil:=0 to TobAxeSrc.Detail.Count-1 do
          begin
          TobVentilSrc := TobAxeSrc.Detail[iVentil];
          if IsSectionPresent(TobAxeDst, TobVentilSrc, TobVentilDst) then
            begin
            TobVentilDst.PutValue('Y_DEBITDEV',  Arrondi(TobVentilSrc.GetValue('Y_DEBITDEV') +TobVentilDst.GetValue('Y_DEBITDEV'),  Devise.Decimale));
            TobVentilDst.PutValue('Y_CREDITDEV', Arrondi(TobVentilSrc.GetValue('Y_CREDITDEV')+TobVentilDst.GetValue('Y_CREDITDEV'), Devise.Decimale));
            end
          else
            begin
            TobVentilDst := CGetNewTOBAna(iAxe+1, TobAxeDst);
            TobVentilDst.Dupliquer(TobVentilSrc, True, True);
            end;
          SetSensMontant(TobVentilDst, 'Y', SensEcr);
          end;

        {recalcule le pourcentage}
        Pourcent := 0;
        for iVentil:=0 to TobAxeDst.Detail.Count-1 do
          begin
          TobVentilDst := TobAxeDst.Detail[iVentil];
          if iVentil = TobAxeDst.Detail.Count-1 then
            Mnt := Arrondi(100-Pourcent, 4)
          else
            begin
            Mnt := Arrondi(TobVentilDst.GetValue('Y_DEBITDEV')+TobVentilDst.GetValue('Y_CREDITDEV'), Devise.Decimale);
            Mnt := Abs(Arrondi(Mnt/(NewDebit+NewCredit)*100, 4));
            end;
          TobVentilDst.PutValue('Y_POURCENTAGE', Mnt);
          Pourcent := Arrondi(Pourcent + Mnt, 4);
          end;
        end;
      end;
  end;

  procedure CumulMultiEcheance(SensEcr: String);
  var
    iEch:         Integer;
    MultiEchSrc:  TMultiEche;
    TobEchSrc:    TOB;
  begin
    if not EstAvecEche(vNumLigneDst) then
      Exit;

    if (PieceSrc.EstAvecEche(vNumLigneSrc)) and Assigned(PieceSrc.GetMultiEche(vNumLigneSrc)) then
      begin
      MultiEchSrc := PieceSrc.GetMultiEche(vNumLigneSrc);
      for iEch:=0 to MultiEchSrc.Detail.Count-1 do
        begin
        TobEchSrc := MultiEchSrc.Detail[iEch];
        CumulMntMultiEcheance(TobEchSrc, SensEcr);
        end;
      end
    else
      begin
      TobEchSrc := PieceSrc.GetTob(vNumLigneSrc);
      CumulMntMultiEcheance(TobEchSrc, SensEcr);
      end;
  end;

var
  NewDebit:     Double;
  NewCredit:    Double;
  EffaceAnaSrc: Boolean;
  SensEcr:      String;
begin                {CumulerMontant}
  {b FP 02/05/2006}
  NewDebit  := Arrondi(GetValue(vNumLigneDst,'E_DEBITDEV') +PieceSrc.GetValue(vNumLigneSrc,'E_DEBITDEV'),  Devise.Decimale);
  NewCredit := Arrondi(GetValue(vNumLigneDst,'E_CREDITDEV')+PieceSrc.GetValue(vNumLigneSrc,'E_CREDITDEV'), Devise.Decimale);

  if (PieceSrc.GetValue(vNumLigneSrc, 'E_GENERAL') <> GetValue(vNumLigneDst, 'E_GENERAL')) or
     (PieceSrc.GetValue(vNumLigneSrc, 'E_AUXILIAIRE') <> GetValue(vNumLigneDst, 'E_AUXILIAIRE')) or
     ((NewDebit-NewCredit) = 0) or
     (PieceSrc.Devise.Code <> Devise.Code) then
    Exit;

  EffaceAnaSrc := False;

  if (NewDebit-NewCredit) > 0 then
    SensEcr := 'DEBIT'
  else
    SensEcr := 'CREDIT';

  if SensEcr = 'DEBIT' then
    begin
    NewDebit  := Abs(Arrondi(NewDebit-NewCredit,  Devise.Decimale));
    NewCredit := 0;
    end
  else
    begin
    NewCredit := Abs(Arrondi(NewDebit-NewCredit,  Devise.Decimale));
    NewDebit  := 0;
    end;

  CumulAnalytique(NewDebit, NewCredit, SensEcr, EffaceAnaSrc);
  CumulMultiEcheance(SensEcr);

  PutValue(vNumLigneDst, 'E_DEBITDEV',  NewDebit);
  PutValue(vNumLigneDst, 'E_CREDITDEV', NewCredit);

  CSynchroVentil(GetTob(vNumLigneDst));

  if EffaceAnaSrc then
    begin
    PieceSrc.PutValue(vNumLigneSrc, 'E_ANA', '-');
    PieceSrc.GetTob(vNumLigneSrc).ClearDetail;
    end;
  {e FP 02/05/2006}
end;


procedure TPieceCompta.SetInfosTva( vIndex : integer ; vRegime, vTva, vTpf : string ) ;
var lInDe    : integer ;
    lInA     : integer ;
    lIndex   : integer ;
    lTob     : Tob ;
begin

  if vRegime='' then Exit ;

  if ModeSaisie = msBor then
    GetBornesGroupe( GetNumGroupe(vIndex), lInDe, lInA )
  else
    begin
    lInDe := 1  ;
    lInA  := Detail.count ;
    end ;

  for lIndex := lInDe to lInA do
    begin
    lTob := CGetTob( lIndex ) ;
    lTob.PutValue('E_REGIMETVA', vRegime ) ;
    lTob.PutValue('E_TVA', vTva ) ;
    lTob.PutValue('E_TPF', vTpf ) ;
    end ;

end;

function TPieceCompta.GetNatureInitiale: string;
begin
  result := 'OD' ;
  if not Info_LoadJournal( GetEnteteS('E_JOURNAL') ) then Exit ;

  result := '' ;
  if (ModeSaisie<>msPiece) and (Info.Journal.NatureParDefaut<>'') then
    result := Info.Journal.NatureParDefaut ;

  if result = '' then
    Case CaseNatJal( Info.Journal.GetValue('J_NATUREJAL') ) of
        tzJVente  : result := 'FC' ;
        tzJAchat  : result := 'FF' ;
        tzJBanque : result := 'RC' ;
        tzJEcartChange : if Info.Journal.GetValue('J_MULTIDEVISE') = 'X'
                           then result := 'ECC' ;
        else      result := 'OD' ;
      end ;
end;

procedure TPieceCompta.AffecteModeSaisie;
begin
  if Info.Journal.GetValue('J_MODESAISIE') = 'BOR' then
    begin
    FModeSaisie := msBor ;
    PutEntete('E_MODESAISIE', 'BOR' ) ;
    end
  else  if Info.Journal.GetValue('J_MODESAISIE') = 'LIB' then
    begin
    FModeSaisie := msLibre ;
    PutEntete('E_MODESAISIE', 'LIB' ) ;
    end
  else
    begin
    FModeSaisie := msPiece ;
    PutEntete('E_MODESAISIE', '-' ) ;
    end ;

end;

function TPieceCompta.GetWhereNatPiece: string;
begin
  result := ' CO_TYPE="NTP" ' ;
  Case CaseNatJal( Journal_GetValue('J_NATUREJAL') ) of
    tzJVente  :      result := result + ' AND (CO_CODE="FC" OR CO_CODE="AC" OR CO_CODE="OD")' ;
    tzJAchat  :      result := result + ' AND (CO_CODE="FF" OR CO_CODE="AF" OR CO_CODE="OD")' ;
    tzJBanque :      result := result + ' AND (CO_CODE="RC" OR CO_CODE="RF" OR CO_CODE="OC" OR CO_CODE="OF" OR CO_CODE="OD")' ;
    tzJEcartChange : result := result + ' AND (CO_CODE="ECC" OR CO_CODE="OD")' ;
    end ;

end;


function TPieceCompta.GetNumGroupe(vIndex: Integer): integer;
begin
  if isOut( vIndex ) then
    begin
    result := - 1 ;
    Exit ;
    end ;

  case ModeSaisie of
    mspiece : Result := 0 ;
    msLibre : Result := 1 ;
    else begin
         if ( vIndex <= 2 )
           then result := 1
           else result := GetValue( vIndex, 'E_NUMGROUPEECR' ) ;
         end ;
    end ;
end;

function TPieceCompta.GetSoldePartiel(vInDe, vInA: integer; vBoDev: Boolean): Double;
var lStChpD : string ;
    lStChpC : string ;
    i       : Integer ;
begin
  result := 0 ;
  // Test des bornes
  if vInDe < 1 then
    vInDe := 1 ;
  if vInA > Detail.count then
    vInA := Detail.count ;
  // Déterminatino des champs
  if vBoDev then
    begin
    lStChpD := 'E_DEBITDEV' ;
    lStChpC := 'E_CREDITDEV' ;
    end
  else
    begin
    lStChpD := 'E_DEBIT' ;
    lStChpC := 'E_CREDIT' ;
    end ;
  // Calcul
  for i := vInDe to vInA do
    result := result + Detail[i-1].GetDouble(lStChpD) - Detail[i-1].GetDouble(lStChpC) ;
  // Arrondi
  if vBoDev
    then result := Arrondi( result  , Devise.Decimale )
    else result := Arrondi( result  , V_PGI.OkDecV ) ;
end;

procedure TPieceCompta.GetBornesGroupe( vNumGroupe: Integer ; var vInDe, vInA : integer );
var i : integer ;
begin

  if vNumGroupe <= 0 then Exit ;

  vInDe := 0 ;
  vInA  := -1 ;

  for i := 1 to Detail.Count do
    begin
    if (vInDe=0) and ( GetValue(i, 'E_NUMGROUPEECR') = vNumGroupe ) then
      vInDe := i ;
    if (vInDe<>0) and ( GetValue(i, 'E_NUMGROUPEECR') <> vNumGroupe ) then
      begin
      vInA := i-1 ;
      break ;
      end ;
    end ;

  if (vInDe <> 0) and (vInA < vInDe) then
    vInA := Detail.Count ;

end;

function TPieceCompta.GetDebutGroupe(vIndex: Integer): integer;
var lNumGroupe : integer ;
    i          : integer ;
begin

  result := 1 ;
  if IsOut( vIndex ) then Exit ;
  lNumGroupe := Detail[vIndex-1].GetInteger('E_NUMGROUPEECR') ;
  if (vIndex<=2) or ( lNumGroupe <= 1 )  then Exit ;

  result := vIndex ;
  for i := (vIndex - 1) downto 1 do
    begin
    if ( Detail[i-1].GetInteger('E_NUMGROUPEECR') <> lNumGroupe ) then
      begin
      result := i+1 ;
      break ;
      end ;
    end ;

end;

function TPieceCompta.GetIncRefAuto(vStRef: string ; vBoPlus : boolean): string;
var lInStart  : Integer ;
    lInEnd    : Integer ;
    i         : Integer ;
    lBoFind   : Boolean ;
    lBoNum    : Boolean ;
    lInVal    : LongInt ;
begin

  Result   := vStRef ;
  if Trim(vStRef)='' then Exit ;

  lBoFind  := False ;
  lInStart := -1 ;
  lInEnd   := -1 ;

  for i := Length(vStRef) downto 1 do
    begin
    lBoNum := (vStRef[i] in ['0'..'9']) ;
    if (lBoFind) and (not lBoNum) then
      begin
      lInStart := i ;
      Break ;
      end ;
    if (not lBoFind) and (lBoNum) then
      begin
      lInEnd   := i ;
      lBoFind  := True ;
      end ;
    end ;

  if (lInEnd>0) and (lInStart<0) then
    lInStart:=0 ;

  if (lInStart>=0) and (lInEnd>0) then
    begin
    Result:='' ;
    if Length( Copy( vStRef, lInStart+1, lInEnd-lInStart ) ) > 9 then
      begin
      Result := vStRef ;
      Exit ;
      end ;

    lInVal := ValeurI( Copy(vStRef, lInStart+1, lInEnd-lInStart) ) ;
    if vBoPlus
      then Inc( lInVal )
      else Dec( lInVal ) ;

    Result := Copy( vStRef, 1, lInStart ) ;
    Result := Result + IntToStr( lInVal ) ;
    Result := Result + Copy( vStRef, lInEnd+1, Length(vStRef)-lInEnd+1 ) ;

    if Length( Result ) > 35 then
      Result := vStRef ;

    end ;

end;

procedure TPieceCompta.AffecteDevise(vIndex: Integer);
begin
  Case ModeSaisie of
    // Modes bor
    msBOR : begin
            // 1ère ligne
            if Detail.count = 1
              then PutValue(1, 'E_DEVISE', GetEnteteS('E_DEVISE') )
            else
              // insertion intermédiaire
              if (vIndex < Detail.count)
                then PutValue(vIndex, 'E_DEVISE', GetValue( vIndex + 1, 'E_DEVISE') )
                // dernière ligne : Recup valeur precedente
                else PutValue(vIndex, 'E_DEVISE', GetValue( vIndex - 1, 'E_DEVISE') ) ;
            end ;

    // mode libre / pièce / ...
    else PutValue( vIndex, 'E_DEVISE', GetEnteteS('E_DEVISE') ) ;

    end ;
end;

function TPieceCompta.GetEnteteI(vChamps: String): integer;
begin
  result := FEntete.GetInteger( vChamps ) ;
end;

function TPieceCompta.GetEnteteS(vChamps: String): string;
begin
  result := FEntete.getString( vChamps ) ;
end;

function TPieceCompta.GetEnteteDt(vChamps: String): TDateTime;
begin
  result := FEntete.GetDateTime( vChamps ) ;
end;

procedure TPieceCompta.PutValeurGroupe( vNumGroupe : integer; vChamps : String; vValeur : Variant );
var lInDe : integer ;
    lInA  : integer ;
    i     : integer ;
    lTob  : Tob ;
begin

  if vNumGroupe < 1 then Exit ;
  if ModeSaisie <> msBor then Exit ;

  lInDe := 1 ;
  lInA  := 0 ;

  GetBornesGroupe( vNumGroupe, lInDe, lInA ) ;
  for i := lInDe to lInA do
    begin
    lTob := Detail[i-1] ;
    if vChamps = 'E_DATECOMPTABLE' then
      PutValue( i, 'E_DATECOMPTABLE', vValeur ) ;
    if lTob.GetValue(vChamps) <> vValeur then
      begin
      lTob.PutValue(vChamps, vValeur);
      if lTob.GetString('E_ANA') = 'X' then
        PutVentil  ( i, vChamps, vValeur ) ;
      if lTob.GetString('E_ECHE') = 'X' then
        PutEche    ( i, vChamps, vValeur ) ;
      end ;
    end ;

end;

function TPieceCompta.EstChampsGroupe( vChamps: string ): boolean;
begin
  if ModeSaisie = msBOR
    then result := (vChamps = 'E_DATECOMPTABLE')   or
                   (vChamps = 'E_NATUREPIECE')     or
                   (vChamps = 'E_DEVISE')          or
                   (vChamps = 'E_TVAENCAISSEMENT') or
                   (vChamps = 'E_REGIMETVA')
    else result := False ;
end;

procedure TPieceCompta.InitLibelle(vIndex: integer;vBoForce:boolean);
var lStLib : string ;
begin
  if isOut( vIndex ) then Exit ;
  if (Trim( GetValue( vIndex, 'E_LIBELLE') ) <> '') and not vBoForce  then Exit ;
//  if ValeurI( GetValue( vIndex, 'E_NUMGROUPEECR') ) > 1  then Exit ;

  lStLib := '' ;

  // Recherche du libellé du tiers
  if ( GetValue( vIndex , 'E_AUXILIAIRE') <> '' ) and
     Info_LoadAux( GetString( vIndex , 'E_AUXILIAIRE') ) then
      lStLib := Aux_GetValue('T_LIBELLE')
  // Recherche du libellé du général
  else if ( GetValue( vIndex , 'E_GENERAL' ) <> '' ) and
          Info_LoadCompte( GetString( vIndex , 'E_GENERAL') ) then
      lStLib := Compte_GetValue('G_LIBELLE') ;

  if Trim(lStLib) <> '' then
    PutValue( vIndex, 'E_LIBELLE', lstLib ) ;

end;

function TPieceCompta.GetSoldeGroupe(vNumGroupe: Integer; vBoDev : Boolean ) : Double ;
var lInDe : integer ;
    lInA  : integer ;
begin

  result := 0 ;
  lInDe  := 0 ;
  lInA   := -1 ;
  GetBornesGroupe  ( vNumGroupe, lInDe, lInA) ;
  if lInA > 0 then
    result := GetSoldePartiel( lInDe, lInA, vBoDev ) ;

end;

procedure TPieceCompta.ReCalculGroupe ( vIndex : integer );
var lInNumRef : integer ;
    lInNumL   : integer ;
    lSoldeGrp : Double ;
begin
  if ModeSaisie <> msBor then Exit ;
  if ModeGroupe <> mgDynamic then Exit ;

  if (vIndex = 0) then Exit ;

  lInNumRef  := 1 ;
  lSoldeGrp  := 0 ;

  for lInNumL := 1 to Detail.Count do
    begin
    // calcul du solde
    lSoldeGrp := lSoldeGrp + Detail[lInNumL-1].GetDouble('E_DEBITDEV') - Detail[lInNumL-1].GetDouble('E_CREDITDEV') ;
    lSoldeGrp := Arrondi( lSoldeGrp, GetRDevise(lInNumL).Decimale ) ;

    // Affectation du groupe de la ligne précédente
    Detail[ lInNumL - 1 ].SetInteger('E_NUMGROUPEECR', lInNumRef) ;

    // Calcul de la condition de changement de groupe
    //  --> Solde atteind et condition de groupe similaire ( Nature, Date, Devise )
    if ( (
         (lSoldeGrp = 0) and EstRemplit( lInNumL )
                         and ( (lInNumL>=(Detail.Count-1) ) or ( Detail[ lInNumL+1 ].GetInteger('E_NUMECHE') <= 1) )
       // Cas du changement de devise
         )
         or
         ( ( lInNumL<Detail.Count ) and (  ( Detail[lInNumL-1].GetString('E_DEVISE') <> Detail[lInNumL].GetString('E_DEVISE') )
                                        or ( Detail[lInNumL-1].GetString('E_NATUREPIECE') <> Detail[lInNumL].GetString('E_NATUREPIECE') )
                                        or ( Detail[lInNumL-1].GetDateTime('E_DATECOMPTABLE') <> Detail[lInNumL].GetDateTime('E_DATECOMPTABLE') )
         ) ) )
      then  begin
            Inc (lInNumRef ) ;
            lSoldeGrp := 0 ; // On repart de 0 pour calculer le solde du groupe suivant !
            end ;
    end ;

end;

function TPieceCompta.LockFolio( vBoParle : boolean ): Boolean;
begin
  result := False ;
  if ModeSaisie <> msPiece then
    begin
    if LockJournal( vBoParle ) then
      result := CBlocageBor( GetEnteteS('E_JOURNAL'),
                             DebutDeMois( GetEnteteDt('E_DATECOMPTABLE') ) ,
                             GetEnteteI('E_NUMEROPIECE'),
                             True ) ;
    end ;

  FBoLockFolio := True ;
end;

function TPieceCompta.UnLockFolio( vBoParle : boolean ): Boolean;
begin
  result := True ;
  if not FBoLockFolio then Exit ;

  result       := False ;
  FBoLockFolio := False ;

  if ModeSaisie <> msPiece then
    begin
    result := CBloqueurBor( GetEnteteS('E_JOURNAL'),
                            DebutDeMois( GetEnteteDt('E_DATECOMPTABLE') ) ,
                            GetEnteteI('E_NUMEROPIECE'),
                            False,
                            vBoParle );
    CBloqueurJournal( false, GetEnteteS('E_JOURNAL'), vBoParle ) ;
    end ;

end;

function TPieceCompta.LockJournal( vBoParle : boolean ): Boolean;
begin
  result := False ;
  if ModeSaisie = msPiece then Exit ;
  if not ( CEstBloqueJournal( GetEnteteS('E_JOURNAL'), vBoParle) ) then
    result := CBloqueurJournal( True, GetEnteteS('E_JOURNAL'), vBoParle ) ;
end;

function TPieceCompta.EstLockFolio: Boolean;
begin
  result := FBoLockFolio;
end;

function TPieceCompta.GetMemoScenario(vIndex: integer): HTStringList;
begin
  if GetScenario( vIndex ) <> nil
    then result := Contexte.FZScenario.Memo
    else result := nil ;
end;

function TPieceCompta.GetCompLigne(var vStComporte, vStLibre: string;  vNumLigne: integer): integer;
var lStRad     : string ;
    i          : integer ;
    lStGene    : string ;
    lTobScenar : TOB ;
begin
  result := 0 ;
  vStComporte  :='' ;
  vStLibre     :='' ;

  lTobScenar := GetScenario( vNumLigne ) ;
  if lTobScenar <> Nil then
    begin

    if vNumLigne = 0 then // paramétrage d'entête
      begin
  {$IFNDEF SANSCOMPTA}
   {$IFNDEF PGIIMMO}
      vStComporte  := Scenario2Comp( GetScenario ) ;
   {$ENDIF PGIIMMO}
  {$ENDIF SANSCOMPTA}
      vStLibre     := lTobScenar.GetString('SC_LIBREENTETE') ;
      end
    else  // paramétrage des lignes
      begin
      if not EstRemplit( vNumLigne ) then Exit ;
      lStGene := GetString( vNumLigne, 'E_GENERAL') ;
      for i:=1 to 10 do
        begin
        lStRad := lTobScenar.GetString( 'SC_RADICAL'+IntToStr(i) ) ;
        if ((lStRad<>'') and ( Copy( lStGene, 1, Length(lStRad) )= lStRad ) ) or (lStRad='*') then
          begin
          vStComporte := lTobScenar.GetString( 'SC_COMPLEMENTS' + IntToStr(i) ) ;
          vStLibre    := lTobScenar.GetString( 'SC_COMPLIBRE' + IntToStr(i) ) ;
          result      := i ;
          Break ;
          end ;
        end ;
      end ;

    end ;

  // Gestion de la saisie QTE de PCL
  if (result=0) and GereQtePCL( vNumLigne ) then
    begin
    lStGene     := GetString( vNumLigne, 'E_GENERAL') ;
    vStLibre    := '----------------------------------------' ;
    vStComporte := '----------' ;
    if ( ( Info.GetString('G_QUALIFQTE1') <> '' ) and ( Info.GetString('G_QUALIFQTE1') <> 'AUC' ) )
      then vStComporte[7] := 'X'
      else vStComporte[7] := '-' ;
    if ( ( Info.GetString('G_QUALIFQTE2') <> '' ) and ( Info.GetString('G_QUALIFQTE2') <> 'AUC' ) )
      then vStComporte[8] := 'X'
      else vStComporte[8] := '-' ;
    end ;


end;

procedure TPieceCompta.RenumeroteLignes(Offset: Integer);
var lNumLigne  : integer ;
    lTobEcr    : Tob ;
    lIdx       : integer ;
    lMultiEche : TMultiEche ;
begin

  lNumLigne := OffSet ;

  // Parcours des lignes de la pièces
  for lIdx := 1 to Detail.Count do
    begin

    // MAJ de l'écriture générale
    lTobEcr := Detail[ lIdx-1 ] ;

    // incrémentation du numligne si pas un détail multi-échéances
    if lTobEcr.GetInteger('E_NUMECHE') >= 2 then continue ;

    Inc( lNumLigne ) ;
    lTobEcr.PutValue('E_NUMLIGNE', lNumLigne ) ;

    // MAJ des ventilations
    if lTobEcr.GetString('E_ANA') = 'X' then
      CReporteVentil( lTobEcr, 'E_NUMLIGNE', lNumLigne, True ) ;

    // MAJ des échéances
    if (lTobEcr.GetString('E_ECHE') = 'X') and ( EstMultiEche( lIdx ) ) then
      begin
      lMultiEche := GetMultiEche( lTobEcr ) ;
      if lMultiEche <> nil then
        lMultiEche.Renumerote ;
      end ;

    end ;

  // En saisie bordereau, il faut renuméroter les groupes de pièces
//  if ModeSaisie=msBor then
//    ReCalculGroupe ;

end;

function TPieceCompta.EstChampsModifiable(vIndex: Integer; vNomChamp: string): Boolean;
var lTob : Tob ;
    lAxe : integer ;
begin
  result := True ;
  if IsOut( vIndex ) then Exit ;

  result := False ;
  if not EstLigneModifiable( vIndex ) then Exit ;

  lTob := GetTob( vIndex ) ;

  // Cas spécifique des saisie mono section
  if pos('SECTION', vNomChamp) > 0 then
    begin
    // compte général saisi
    if not EstRemplit( vIndex ) then Exit ;
    // compte ventilable sur l'axe désigné
    lAxe := ValeurI( Copy( vNomChamp, length(vNomChamp), 1) ) ;
    if not EstVentilable( vIndex, lAxe) then Exit ;
    // Ventilation mono section
    if (lTob.Detail.Count > 0) and (lTob.Detail[lAxe-1].Detail.count > 1) then Exit ;
    end ;

  result := True ;

end;


procedure TPieceCompta.PutValueSpecif(vIndex: Integer; vChamps: String; vValeur: Variant);
var lStAxe    : string ;
    lInAxe    : integer ;
    lTobEcr   : Tob ;
    lOldValue : Variant ;
begin

  lTobEcr := getTob( vIndex ) ;
  if lTobEcr = nil then Exit ;

  // Récup ancienne valeur et comparaison
  lOldValue := lTobEcr.GetValue( vChamps ) ;
  if lOldValue = vValeur then Exit ;

  // saisie mono-section
  if pos('SECTION', vChamps) > 0 then
    begin
    lStAxe  := Copy( vChamps, 8, 2 ) ;
    lInAxe  := ValeurI( Copy( vChamps, 9, 1 ) ) ;

    if lTobEcr.GetString('E_ANA') <> 'X' then Exit ;

    if lTobEcr.Detail.count = 0 then
      AlloueAxe( lTobEcr ) ;

    if lTobEcr.Detail[ lInAxe-1 ].Detail.Count > 1 then Exit ;

    if Info.LoadSection( vValeur, lStAxe ) then
      begin
      if lTobEcr.Detail[ lInAxe-1 ].Detail.Count = 0
        then CVentilerAttente( lTobEcr, lInAxe, vValeur )
        else lTobEcr.Detail[ lInAxe - 1 ].Detail[0].PutValue('Y_SECTION', vValeur ) ;
      lTobEcr.PutValue( vChamps, vValeur ) ;
      end ;
    end
  else
    lTobEcr.PutValue( vChamps, vValeur ) ;
    ;
end;

procedure TPieceCompta.SynchroMonoSection( vIndex : integer ) ;
var iAxe  : integer ;
    lTobE : Tob ;
    lTobV : Tob ;
begin

  lTobE := GetTob( vIndex ) ;
  if lTobE = nil then Exit ;
  if lTobE.GetString('E_ANA')<>'X' then Exit ;

  for iAxe := 1 to lTobE.Detail.Count do
    begin

    if (lTobE.Detail[ iAxe - 1 ].Detail.Count = 1)  then
      begin
      lTobV := lTobE.Detail[ iAxe - 1 ].Detail[ 0 ] ;
      if lTobV.GetString('Y_SECTION') <> lTobE.GetString( 'SECTIONA' + IntToStr( iAxe ) )
        then lTobE.PutValue( 'SECTIONA' + IntToStr( iAxe ) , lTobV.GetValue('Y_SECTION') ) ;
      end
    else
      lTobE.PutValue( 'SECTIONA' + IntToStr( iAxe ), '' )

    end ;
end;

function TPieceCompta.RechercheChampDoublon(vIndex: integer; vStChp : string ;  var vTobRech: Tob): Boolean;
var lStVal      : String ;
    lTobEcr     : TOB ;
    lStChpExclu : string ;
    i           : integer ;
begin
  Result   := False ;
  vTobRech := nil ;

  if ( GetEnteteS('E_QUALIFPIECE')<>'N' ) then Exit ;

  if not EstRemplit( vIndex ) then Exit ;

  if ModeSaisie=msBor
    then lStChpExclu := 'E_NUMGROUPEECR'
    else lStChpExclu := 'E_NUMLIGNE' ;

  lTobEcr  := GetTob( vIndex ) ;
  if Length( lTobEcr.GetString('E_AUXILIAIRE') ) <> GetInfoCpta( fbAux ).Lg then Exit ;

  lStVal   := lTobEcr.GetString( vStChp ) ;
  if ( vStChp = '' ) or ( lStVal = '' ) then Exit ;

  for i := 1 to Detail.count do
    begin
    vTobRech := Detail[ i-1 ] ;
    if vTobRech.GetInteger(lStChpExclu) = lTobEcr.GetInteger(lStChpExclu) then continue ;
    if vTobRech.GetString('E_AUXILIAIRE') <> lTobEcr.GetString('E_AUXILIAIRE') then continue ;
    if vTobRech.GetString('E_GENERAL') <> lTobEcr.GetString('E_GENERAL') then continue ;
    if pos( vTobRech.GetString('E_NATUREPIECE'), 'AC;AF;FC;FF') = 0 then continue ;
    if vTobRech.GetString(vStChp) = lStVal then
      begin
      result := True ;
      break ;
      end ;

    end ;

  if not result then
    vTobRech := nil ;

end;

function TPieceCompta.GetRDevise(vIndex: integer; vBoMajTaux : boolean ): RDevise;
var lStCode : string ;
begin

  if (modeSaisie = msBOR) and not isOut( vIndex )
    then lStCode := GetValue( vIndex, 'E_DEVISE' )
    else lStCode := GetEnteteS( 'E_DEVISE' ) ;

  if FDevise.code<>lStCode
    // MAJ de la devise
    then Info_LoadDevise( lStCode )
    // Maj du taux
    else majDeviseTaux( vIndex, vBoMajTaux ) ;

  result := FDevise ;

end;


function TPieceCompta.VerifieEquilibre( vIndex : integer ) : boolean ;
var lInDeb    : integer ;
    lInFin    : integer ;
    i         : integer ;
    lSoldeDev : Double ;
    lSoldePiv : Double ;
    lCreditL  : Double ;
    lDebitL   : Double ;
    lTobL     : Tob ;
    lRDev     : RDevise ;
    lInGrp    : integer ;
begin

  result := False ;

  if (ModeSaisie=msBOR) and not isOut(vIndex) then
    begin
    lInGrp := GetNumGroupe( vIndex ) ;
    if lInGrp <= 0 then Exit ;
    GetBornesGroupe( lInGrp, lInDeb, lInFin) ;
    end
  else
    begin
    lInDeb := 1 ;
    lInFin := Count ;
    end ;

  if IsOut( lInDeb ) then Exit ;

  lTobL    := Detail[ lInDeb-1 ] ;
  if lTobL.GetString('E_DEVISE')=V_PGI.DevisePivot then Exit ;

  // Calcul des soldes
  lSoldeDev := 0 ;
  lSoldePiv := 0 ;
  for i:= lInDeb to lInFin do
    begin
    lSoldeDev := lSoldeDev + Detail[i-1].GetDouble('E_DEBITDEV') - Detail[i-1].GetDouble('E_CREDITDEV') ;
    lSoldePiv := lSoldePiv + Detail[i-1].GetDouble('E_DEBIT') - Detail[i-1].GetDouble('E_CREDIT') ;
    end ;

  // Si le solde en devise n'est pas nul rien a faire
  lRDev := GetRDevise( lInDeb ) ;
  if Arrondi(lSoldeDev, lRDev.Decimale) <> 0 then Exit ;

  // Si le solde en devise est nul mais par celui en pivot, il faut équilibré
  if Arrondi(lSoldePiv, V_PGI.OkDecV) = 0 then Exit ;

  // Mise en place du diff sur la 1ère ligne du groupe
  // Sauf en mode libre ==> sur la dernière ligne
  if (ModeSaisie = msLibre) then
    lInDeb := Detail.count ;
  lTobL    := Detail[ lInDeb-1 ] ;
  lDebitL  := lTobL.GetDouble('E_DEBIT')  ;
  lCreditL := lTobL.GetDouble('E_CREDIT') ;
  if lDebitL <> 0
    then lDebitL  := lDebitL  - lSoldePiv
    else lCreditL := lCreditL + lSoldePiv ;
  lTobL.PutValue('E_DEBIT', lDebitL ) ;
  lTobL.PutValue('E_CREDIT', lCreditL ) ;

  result := True ;

end;

function TPieceCompta.Info_LoadDevise( vStCode: string ; vBoEnBase : boolean ): Boolean;
begin

  result := Info.Devise.Load( [vStCode], vBoEnBase ) <> -1 ;
  if vBoEnBase or (FDevise.Code <> vStCode) then
    begin
    FDevise.Code := vStCode ;
    CTOBVersRDevise( Info.Devise.Item, FDevise ) ;
    // MAJ des taux
    Info.Devise.Item.PutValue( 'TAUX'     , 1     ) ;
    Info.Devise.Item.PutValue( 'DATETAUX' , iDate1900 ) ;
    majDeviseTaux( 0, vBoEnBase ) ;
    end ;

end;

procedure TPieceCompta.ChangeTauxPiece(vIndex: integer);
var lInDe : integer ;
    lInA  : integer ;
    i     : integer ;
    lRDev : RDevise ;
    lTobL : Tob ;
begin

  Case ModeSaisie of

    msBor :   begin
              if IsOut( vIndex ) then Exit ;
              GetBornesGroupe( GetValue( vIndex, 'E_NUMGROUPEECR'), lInDe, lInA ) ;
              end ;

    else      begin
              lInDe := 1 ;
              lInA  := Detail.Count ;
              end ;
    end ;

  if isOut(lInDe) then Exit ;

  lRDev := GetRDevise( vIndex ) ;

  for i := lInDe to lInA do
    begin
    lTobL := Detail[i-1] ;
    if lTobL.GetDouble('E_DEBITDEV')<>0
      then PutValue( i, 'E_DEBITDEV',  lTobL.GetDouble('E_DEBITDEV'),  True )
      else PutValue( i, 'E_CREDITDEV', lTobL.GetDouble('E_CREDITDEV'), True ) ;
//    CSetMontants( lTobL, lTobL.GetDouble('E_DEBITDEV'), lTobL.GetDouble('E_CREDITDEV'), lRDev, True ) ;
    end ;
    
end;

function TPieceCompta.EstPieceSoldee(vIndex: integer): Boolean;
var lInGrp : integer ;
begin
  if (ModeSaisie = msBor) and (vIndex <> 0) then
    begin
    result := False ;
    lInGrp := GetNumGroupe( vIndex ) ;
    if lInGrp <= 0 then Exit ;
    result := GetSoldeGroupe( lInGrp ) = 0 ;
    end
  else
    result := ( Detail.Count >= 2 ) and ( GetSolde = 0 ) ;
end;

function TPieceCompta.PieceModifiable(vIndex: Integer): Boolean;
var i      : integer ;
    lInDeb : integer ;
    lInFin : integer ;
    lInGrp : integer ;
begin
  result := True ;

  if ModeSaisie=msBor then
    begin
    lInGrp := GetNumGroupe( vIndex ) ;
    if lInGrp <= 0 then Exit ;
    GetBornesGroupe( lInGrp, lInDeb, lInFin ) ;
    end
  else
    begin
    lInDeb := 1 ;
    lInFin := Detail.count ;
    end ;

  for i := lInDeb to lInFin do
    begin
    if not EstLigneModifiable( i ) then
      begin
      result := False ;
      Exit ;
      end;
    end ;
end;

function TPieceCompta.EstMonoDevise: boolean;
var i       : integer ;
    lStCode : string ;
begin
  result := True ;

  if ModeSaisie = msBor then
    begin
    lStCode := GetEnteteS('E_DEVISE') ;
    for i:=0 to Detail.count-1 do
      begin
      if lStCode <> Detail[i].GetString('E_DEVISE') then
        begin
        result := False ;
        Exit ;
        end ;
      end ;
    end ;

end;

procedure TPieceCompta.SetMultiEcheAucun;
begin
  FModeEche := meAucun ;
end;

function TPieceCompta.GetMultiEche(vTobEcr: Tob): TMultiEche;
begin

  result := nil ;

//  if FModeEche = meAucun then Exit ;
  if vTobEcr.GetString('E_ECHE')<>'X' then Exit ;

  if vTobEcr.GetInteger('E_NUMECHE')>1 then
    vTobEcr := GetFirstEche( vTobEcr.GetIndex + 1 ) ;

  if FModeEche <> meDeporte then
    begin
    if FEcheances.Detail.count = 0 then
      CCreerMultiEche( FEcheances, self, vTobEcr, taModif ) ;
    result := TMultiEche( FEcheances.Detail[0] ) ;
    result.SetLignePiece( vTobEcr ) ;
    end
  else
    result := GetMultiEche( vTobEcr.GetInteger('E_NUMLIGNE') ) ;

end;


function TPieceCompta.CGetTob(vIndex: Integer): Tob;
begin
  result := Nil ;
  if isOut( vIndex ) then Exit ;
  result := Detail[ vIndex-1 ] ;
end;

procedure TPieceCompta.SetMultiEcheDeporte;
begin
  FModeEche := meDeporte ;
end;

procedure TPieceCompta.SetMultiEcheMulti;
begin
  FModeEche := meMulti ;
end;

function TPieceCompta.GetTob(vIndex: Integer): TOB;
begin
  result := CGetTob( vIndex ) ;
  if Assigned(Result) then
    begin
    Info_LoadCompte( result.GetString('E_GENERAL') ) ;
    Info_LoadAux( result.GetString('E_AUXILIAIRE') ) ;
    Info_LoadDevise( result.getString('E_DEVISE') );
    CurIdx := vIndex ;
    end ;
end;

function TPieceCompta.GetDateTime(vIndex: Integer; vChamps: String): TDateTime;
begin
  result := iDate1900 ;
  if IsOut( vIndex ) then Exit ;
  result := Detail[ vIndex - 1 ].GetDateTime( vChamps ) ;
end;

function TPieceCompta.GetDouble(vIndex: Integer; vChamps: String): Double;
begin
  result := 0.0 ;
  if IsOut( vIndex ) then Exit ;
  result := Detail[ vIndex - 1 ].GetDouble( vChamps ) ;
end;

function TPieceCompta.GetInteger(vIndex: Integer; vChamps: String): Integer;
begin
  result := 0 ;
  if IsOut( vIndex ) then Exit ;
  result := Detail[ vIndex - 1 ].GetInteger( vChamps ) ;
end;

function TPieceCompta.GetString(vIndex: Integer; vChamps: String): String;
begin
  result := '' ;
  if IsOut( vIndex ) then Exit ;
  result := Detail[ vIndex - 1 ].GetString( vChamps ) ;
end;

function TPieceCompta.GetFirstEche(vIndex: integer): Tob;
var i : integer ;
begin
  result := nil ;
  if isOut( vIndex ) then Exit ;
  if GetInteger(vIndex, 'E_NUMECHE')<=1 then
    result := CGetTob( vIndex )
  else
    for i:=vIndex downto 1 do
      if Detail[i-1].GetInteger('E_NUMECHE') <=1 then
        begin
        result := CGetTob( i ) ;
        Break ;
        end ;
end;

procedure TPieceCompta.IncrementRef( vIndex : integer ; var vLastRef : string ; vBoPlus : boolean ) ;
var lTob    : Tob ;
    lStRef  : String ;
begin

  lTob := CGetTob( vindex ) ;
  if lTob = nil then Exit ;

  if lTob.GetString('E_REFINTERNE')<>''
    then lStRef := GetIncRefAuto( lTob.GetString('E_REFINTERNE'), vBoPlus )
    else lStRef := GetIncRefAuto( vLastRef, vBoPlus ) ;

  PutValue( vIndex, 'E_REFINTERNE', lStRef ) ;

  vLastRef := lTob.getString('E_REFINTERNE') ;

end;

function TPieceCompta.GetNbEche(vIndex: Integer): integer;
var i          : integer ;
    lTob       : Tob ;
    lMultiEche : TMultiEche ;
begin

  result := 0 ;

  lTob := CGetTob( vIndex ) ;
  if lTob = nil then Exit ;
  if lTob.GetString('E_ECHE')<>'X' then Exit ;

  result := 1 ;
  if FModeEche = meDeporte then
    begin
    lMultiEche := GetMultiEche( vIndex ) ;
    if Assigned( lMultiEche ) then
      result := lMultiEche.Detail.count ;
    end
  else
    begin
    if lTob.GetInteger('E_NUMECHE')>1 then
      begin
      lTob   := GetFirstEche( vIndex ) ;
      vIndex := lTob.GetIndex + 1 ;
      end ;

    for i := vIndex to (Detail.Count - 1) do
      begin
      if Detail[i].GetInteger('E_NUMLIGNE')<> lTob.GetInteger('E_NUMLIGNE') then Break ;
      Inc(result) ;
      end ;

    end ;

end;

procedure TPieceCompta.Debug;
begin
  {$IFNDEF EAGLSERVER}
  TobDebug( self ) ;
  {$ENDIF EAGLSERVER}
end;

procedure TPieceCompta.CalculDesSoldes;
var lTS   : TList;
    lTSA  : TList ;
    k     : integer ;
    pT    : ^TFRM ;

   procedure _GestionMS( vList : Tlist ; vStDossier : string ) ;
   var j  : integer ;
     begin
     for j := 0 to vList.Count-1 do
       begin
       pT := vList.Items[ j ] ;
       {$IFNDEF PGIIMMO}
         pT.Dossier := vStDossier ;
       {$ENDIF}
       end ;
     end ;

   procedure _CalculListPourPiece( vPiece : Tob ; vBoPlus : boolean ) ;
     var lTOBLigneEcr    : TOB;
         i               : integer;
     begin
       // Parcours de la pièce
       for i := 0 to vPiece.Detail.Count - 1 do
         begin
         lTOBLigneEcr    := vPiece.Detail[i];
         // ajoute les elements à la TList
         AjouteTOB   (lTS,  lTOBLigneEcr , vBoPlus );
         {$IFNDEF PGIIMMO}
           if Contexte.CPAnoDyna then
             begin
             Info.LoadCompte( lTOBLigneEcr.GetString('E_GENERAL') ) ;
             AjouteAno( lTSA, lTOBLigneEcr, Info.GetString('G_NATUREGENE'), not vBoPlus ) ;
             end ;
         {$ENDIF}
         end; // if
     end ;


begin

  // Création des objets
  lTS  := TList.Create ;
  lTSA := TList.Create ;

  // Passage de la tob d'origine si besoin
  if Action=taModif then
    _CalculListPourPiece( FTobOrigine, False ) ;

  // Calculs
  _CalculListPourPiece( self, true ) ;

   // Gestion MultiSoc
   _GestionMS( lTS,  Dossier ) ;
   _GestionMS( lTSA, Dossier ) ;

  try

    // enregistre en base
    ExecMajSoldeTOB ( lTS ) ;
    ADevalider ;

    // ANO dyna
    if Contexte.CPAnoDyna then
      begin
      if not ExecReqMAJAno(lTSA) then
        begin
        V_PGI.IoError := oeSaisie ;
        exit ;
        end ;
      end ;

   finally

      for k:=0 to lTS.Count-1 do
       if assigned(lTS[k]) then
         begin
         pT := lTS[k] ;
         Dispose( pT );
         lTS[k] := nil ;
         end ;
      FreeAndNil( lTS ) ;

      for k:=0 to lTSA.Count-1 do
       if assigned(lTSA[k]) then
         begin
         pT := lTSA[k] ;
         Dispose( pT ) ;
         pT := nil ;
         end ;
      FreeAndNil ( lTSA ) ;

    end ;

end;

function TPieceCompta.ADevalider: Boolean;
Var lDebExo   : TDateTime ;
    lStExo    : String ;
    lStValJ   : String ;
    lStValE   : String ;
    lStChamp  : String ;
    iPer      : Byte ;
    lQSelect  : TQuery ;
    lStJal    : string ;
    lDtDateC  : TDateTime ;
begin
  Result   := False ;
  lStJal   := GetEnteteS('E_JOURNAL') ;
  lDtDateC := GetEnteteDt('E_DATECOMPTABLE') ;
  lStExo   := GetEnteteS('E_EXERCICE') ;

  if (lStExo = contexte.Exercices.EnCours.Code) then
    begin
    lDebExo  := contexte.Exercices.EnCours.Deb ;
    lStChamp := 'J_VALIDEEN' ;
    end
  else if (lStExo = contexte.Exercices.Suivant.Code) then
    begin
    lDebExo  := contexte.Exercices.Suivant.Deb ;
    lStChamp := 'J_VALIDEEN1' ;
    end 
  else Exit ;

  iPer      := QuellePeriode( lDtDateC, lDebExo ) ;
  lQSelect  := OpenSelect( 'SELECT ' + lStChamp + ' FROM ' + GetTableDossier( Dossier, 'JOURNAL')
                                                + ' WHERE J_JOURNAL="'+lStJal+'"'
                            , Dossier ) ;
  lStValJ   := lQSelect.FindField( lStChamp ).AsString ;
  Ferme(lQSelect) ;

  if lStValJ='' then Exit ;
  if lStValJ[iPer]='-' then Exit
    else
      begin
      lQSelect   := OpenSelect('SELECT EX_VALIDEE FROM EXERCICE Where EX_EXERCICE="'+lStExo+'"',Dossier) ;
      lStValE := lQSelect.FindField('EX_VALIDEE').AsString ;
      Ferme(lQSelect) ;
      if ((lStValE<>'') and (lStValE[iPer]='X')) then
        begin
        lStValE[iPer] := '-' ;
        ExecuteSql('UPDATE ' + GetTableDossier( Dossier, 'EXERCICE') + ' SET EX_VALIDEE="' + lStValE + '" WHERE EX_EXERCICE="' + lStExo + '"') ;
        end ;
      lStValJ[iPer] := '-' ;
      ExecuteSql('UPDATE JOURNAL SET '+lStChamp+'="'+lStValJ+'" Where J_JOURNAL="'+lStJal+'"') ;
      end ;

  Result := True ;

end;

procedure TPieceCompta.SynchroniseTobCompl;
var lTobTmp   : Tob ;
    i         : integer ;
    lTobCompl : Tob ;
begin

  if FEcrCompl.Detail.count = 0 then Exit ;

  lTobTmp := Tob.Create('$TEMP', nil, -1);
  // On isole les TobCompl ayant une référence sur les lignes
  for i := 0 to Detail.count-1 do
    begin
    lTobCompl := CGetTobCompl( Detail[i] ) ;
    if Assigned( lTobCompl ) then
      lTobCompl.ChangeParent( lTobTmp, -1 ) ;
    end ;

  // efface les TobCompl isolées
  FEcrCompl.ClearDetail ;

  // réintégration des tobcompl
  for i := lTobTmp.Detail.count-1 downto 0 do
    begin
    lTobCompl := lTobTmp.Detail[i] ;
    lTobCompl.ChangeParent( FEcrCompl, -1 ) ;
    end ;

end;

procedure TPieceCompta.GetHistoPourChamps(vChamps, vValeur: String; var vTotalDebit, vTotalCredit: Double);
var i      : Integer ;
    lTob   : Tob ;
begin
  vTotalDebit  := 0 ;
  vTotalCredit := 0 ;

  if FTobOrigine.Detail.Count = 0 then Exit ;

  for i := 0 to FTobOrigine.Detail.Count-1 do
    begin
    lTob := FTobOrigine.Detail[i] ;
    if lTob.GetString( vChamps ) = vValeur then
      begin
      vTotalDebit  := vTotalDebit  + lTob.GetDouble( 'E_DEBIT' ) ;
      vTotalCredit := vTotalCredit + lTob.GetDouble( 'E_CREDIT' ) ;
      end ;
    end ;
  // FQ 17129 : Pb d'arrondi Delphi
  vTotalDebit  := Arrondi( vTotalDebit  , V_PGI.OkDecV ) ;
  vTotalCredit := Arrondi( vTotalCredit , V_PGI.OkDecV ) ;

end;

function TPieceCompta.GetCompteAcc(vIndex: integer ; vStGenDefaut : string): string;
var lStAux : string ;
    i      : integer ;

   function _GetContre : string ;
     var j    : integer ;
     begin
       result := GetString( vIndex, 'E_CONTREPARTIEGEN') ;
       if result<>'' then Exit ;
       for j := vIndex-1 downto 1 do
          begin
          if lStAux = GetString( j, 'E_AUXILIAIRE') then
            begin
            result := GetString( j, 'E_CONTREPARTIEGEN') ;
            Break ;
            end;
          end ; // for
     end ;

begin

  result := '' ;

  lStAux := GetString( vIndex, 'E_AUXILIAIRE') ;
  if lStAux = '' then Exit ;
  if not Info.LoadAux(lStAux) then Exit ;

  if ( Info.getString('YTC_ACCELERATEUR') = 'X') then
    begin
    result := Info.getString('YTC_SCHEMAGEN') ;
    if result = '' then
      result := _GetContre ;
    end  // if
  else
    begin
    result := _GetContre ;
    if result = '' then
      result:= Info.getString('YTC_SCHEMAGEN') ;
    end ;

 if result = '' then
   begin
   result := vStGenDefaut ;
   if result='' then
     for i := vIndex-1 downto 1 do
       begin
       if EstHT(i) then
         begin
         result := GetString( i, 'E_GENERAL') ;
         Break ;
         end;
       end ;
   end ;

 if result <> '' then
   if ( not Info.LoadCompte(result) ) or Info.Compte.IsCollectif  then
     result := '' ;

end;

procedure TPieceCompta.ReinitCutOff( vIndex : Integer );
var lTob : Tob ;
begin

  lTob := CGetTob( vIndex ) ;
  if lTob = nil then Exit ;

  lTob.PutValue( 'CUTOFF', '-' ) ;

{$IFDEF SCANGED}
  if ( RechGuidId( vIndex ) = '' ) then
  CFreeTOBCompl( lTob )
   else
    begin
     CPutValueTOBCompl( lTob, 'EC_CUTOFFDEB', iDate1900 ) ;
     CPutValueTOBCompl( lTob, 'EC_CUTOFFFIN', iDate1900 ) ;
    end ;
{$ENDIF SCANGED}

end;

procedure TPieceCompta.InsereLigneHT( vIndexTTC: integer; vStCptHT : string ; vMontantHT: double );
var lStAux      : String ;
    lStCptTva   : String ;
    lMontantHT  : Double ;
    lMontantTTC : Double ;
    lStChamps   : String ;
    lInNumL     : Integer ;
begin

  // Vérif Auxiliaire
  lStAux := GetString( vIndexTTC, 'E_AUXILIAIRE' ) ;
  if lStAux = '' then Exit ;
  if not Info_LoadAux( lStAux ) then Exit ;
  if not Info_LoadCompte( vStCptHT ) then Exit ; // pas de compte HT, on sort

  // Sens du montant TTC
  lMontantTTC := GetDouble( vIndexTTC, 'E_DEBITDEV' ) ;
  if lMontantTTC <> 0 then
    lStChamps   := 'E_CREDITDEV'
  else
    begin
    lMontantTTC := GetDouble( vIndexTTC, 'E_CREDITDEV' ) ;
    lStChamps   := 'E_DEBITDEV' ;
    end ;

  if lMontantTTC = 0 then Exit ; // pas de montant TTC on sort

  // ========================================
  // ======= Ajout de la ligne de HT ========
  // ========================================

  // paramétrage Ventilation auto ?
//  lSavVentil   := FModeVentil ;
//  FModeVentil  := mvDefaut ;

  // Mise en place de la ligne HT sous la ligne TTC
  NewRecord( vIndexTTC + 1 ) ;
  lInNumL := CurIdx ;
  PutValue( lInNumL, 'E_GENERAL', vStCptHT ) ;

  // Autres info
  PutValue( lInNumL, 'E_REGIMETVA',       GetRegimeTva( vIndexTTC ) ) ;//GetRegimeTva(lInNumL) ) ;
  PutValue( lInNumL, 'E_CONTREPARTIEGEN', GetString( vIndexTTC, 'E_GENERAL' ) ) ;
  PutValue( lInNumL, 'E_CONTREPARTIEAUX', GetString( vIndexTTC, 'E_AUXILIAIRE' ) ) ;

  // Y'aura-t-il de la Tva ???
  lStCptTva := GetCompteTva( lInNumL ) ;
  if vMontantHT <> 0 then
    lMontantHT := vMontantHt
  else if lStCptTva <> '' then
    lMontantHT := TTC2HT( lMontantTTC,
                          GetRegimeTva( vIndexTTC ),
                          EstSoumisTPF( vIndexTTC ),
                          GetCodeTva( lInNumL ),
                          GetCodeTpf( lInNumL ),
                          EstAchat( lInNumL ),
                          GetRDevise( lInNumL ).Decimale )
  else
    lMontantHT := Arrondi( lMontantTTC ,Devise.Decimale) ;

  // Affectation du montant HT
  PutValue( lInNumL, lStChamps, lMontantHT ) ;

  // === MAJ LIGNE TTC ===
  // MAJ des informations de TVA de la ligne TTC
  PutValue( vIndexTTC, 'E_TVA', GetCodeTva( lInNumL ) ) ;
  PutValue( vIndexTTC, 'E_TPF', GetCodeTpf( lInNumL ) ) ;
  // MAJ des contreparties
  PutValue( vIndexTTC, 'E_CONTREPARTIEGEN', GetString( lInNumL, 'E_GENERAL' ) ) ;
  PutValue( vIndexTTC, 'E_CONTREPARTIEAUX', GetString( lInNumL, 'E_AUXILIAIRE' ) ) ;

  FCurIdx := lInNumL ;

end;

function TPieceCompta.InsereLigneTVA(vIndexHT, vIndexTTC: integer) : boolean ;
Var lTobHT        : TOB ;
    lStRegime     : string ;
    lBoTpf        : boolean ;
    lInIdxTVA     : integer ;
    lSoldeTTC     : double ;
    lSoldeHT      : double ;
//    lStChamp      : string ;
begin

  result := False ;

  lTobHT  := CGetTob( vIndexHT ) ;
  if lTobHT=nil then Exit ;

  if isOut( vIndexTTC ) then Exit ;

  // infos reprises du TTC
  lStRegime := GetString( vIndexTTC, 'E_REGIMETVA') ;
  lBoTpf    := EstSoumisTPF( vIndexTTC ) ;

  // affectation du régime au HT
  if GetString( vIndexHT, 'E_REGIMETVA')='' then
    lTobHT.PutValue('E_REGIMETVA', lStRegime );

  result    := InsereLigneTVA( vIndexHT, lStRegime, lBoTpf ) ;
  lInIdxTVA := CurIdx ;

  // gestoin de l'arrondi sur la dernière ligne de tva
  if result and (lInIdxTVA > vIndexHT) then
    begin
    lSoldeTTC := GetSoldePartiel( vIndexTTC, vIndexTTC ) ;
    lSoldeHT  := GetSoldePartiel( vIndexHT, lInIdxTVA ) ;
    if Arrondi( lSoldeTTC + lSoldeHT, GetRDevise(vIndexTTC).Decimale ) <> 0 then
      begin
      if GetDouble( lInIdxTva, 'E_CREDITDEV' ) > 0
        then PutValue( lInidxTva, 'E_CREDITDEV', GetDouble( lInIdxTva, 'E_CREDITDEV' ) + (lSoldeTTC + lSoldeHT) )
        else PutValue( lInidxTva, 'E_DEBITDEV',  GetDouble( lInIdxTva, 'E_DEBITDEV' ) - (lSoldeTTC + lSoldeHT) ) ;
      end ;
    end ;

end;

function TPieceCompta.GetTiers(vIndex: integer): string;
var lTob : Tob ;
begin

  result := '' ;

  lTob := CGetTob( vIndex ) ;
  if lTob = nil then Exit ;

  // Auxiliaire ?
  if ( lTob.GetString('E_AUXILIAIRE') <> '' ) and
     Info_LoadAux( lTob.GetString('E_AUXILIAIRE') )
    then result := lTob.GetString('E_AUXILIAIRE')
  else
  // Général de nature TIC ou TID ?
  if ( lTob.GetString('E_GENERAL') <> '' ) and
     ( Info_LoadCompte( lTob.GetString('E_GENERAL') ) ) and
     ( ( Compte_GetValue('G_NATUREGENE') = 'TID' ) or
       ( Compte_GetValue('G_NATUREGENE') = 'TIC' ) )
    then result := lTob.GetString('E_GENERAL') ;

end;

procedure TPieceCompta.GetInfosTva(vIndex: integer; var vStRegime, vStTva, vStTpf: string );
var lInDe    : integer ;
    lInA     : integer ;
    lIndex   : integer ;
begin

  if ModeSaisie = msBor then
    GetBornesGroupe( GetNumGroupe(vIndex), lInDe, lInA )
 else
   begin
   lInDe := 1  ;
   lInA  := Detail.count ;
   end ;

  vStRegime := GetRegimeTva( vIndex ) ;
  vStTva    := '' ;
  vStTpf    := '' ;

  if EstHT( vIndex ) then
    begin
    vStTva := GetString( vIndex, 'E_TVA' ) ;
    vStTpf := GetString( vIndex, 'E_TPF' ) ;
    end
  else
    for lIndex := lInDe to lInA do
      if (GetString( lIndex, 'E_TVA' ) <> '') then
        begin
        vStTva := GetString( lIndex, 'E_TVA' ) ;
        vStTpf := GetString( lIndex, 'E_TPF' ) ;
        Exit ;
        end ;

end;

procedure TPieceCompta.ReinitEche(vIndex: Integer);
var lTob       : TOB ;
    lMultiEche : TMultiEche ;
begin

  lTob := CGetTob( vIndex ) ;
  if lTob = nil then Exit ;

  lMultiEche := GetMultiEche( lTob ) ;
  if lMultiEche<>nil then
    begin
    lMultiEche.ClearEche ;
    lMultiEche.ModeFinal := '' ;
    end ;

  CSupprimerInfoLettrage( lTob ) ;
  if lTob.GetNumChamp('MODEFINAL')>0 then
    lTob.PutValue('MODEFINAL', '') ;

end;

procedure TPieceCompta.CalculEche(vindex: integer; vStModeRegle: string);
var lTob       : TOB ;
    lMultiEche : TMultiEche ;
begin

  lTob := CGetTob( vIndex ) ;
  if lTob = nil then Exit ;
  if lTob.GetInteger('E_NUMECHE')>1 then Exit ;

  // Compte lettrable ou pointable
  if not EstAvecEche( vIndex ) then Exit ;
  if not IsValidDebitCredit( vIndex ) then Exit ;

  Case FModeEche of

      // Gestion dans objet = mono échéance
      meAucun :  begin
{                 if lTob.GetString('E_MODEPAIE') = '' then
                   begin
                   CCalculMonoEche( lTOB , Info, vStModeRegle ) ;
                   AffecteDateValeur( vIndex ) ;
                   end ;
}                 end ;

      meDeporte : begin
                  lMultiEche := GetMultiEche( vIndex ) ;
                  if lMultiEche = nil
                    then lMultiEche := CCreerMultiEche( FEcheances, self, lTob, taCreat, True, mfAlone )
                    else lMultiEche.SynchroniseEcheances ;
                  // Maj de la ligne d'affichage avec la 1ère échéance
                  lMultiEche.SynchroniseEcr;
                  end ;

      // meMono, meMulti : nouvelle gestion
      else        begin
                  if lTob.GetString('E_MODEPAIE') = '' then
                    begin
                    lMultiEche := GetMultiEche( lTob ) ;
                    lMultiEche.CalculEche( vStModeRegle ) ;
                    lMultiEche.Action := taCreat ;
                    if (lMultiEche.ModeFinal <> '') then
                      if lTob.GetNumChamp('MODEFINAL') > 0
                        then lTob.putValue('MODEFINAL', lMultiEche.ModeFinal )
                        else lTob.addchampsupValeur('MODEFINAL', lMultiEche.ModeFinal ) ;
                    end ;
                  end ;
      end ;

end;

function TPieceCompta.GetCodeTpf(vIndex: Integer): String;
begin
  result := GetString( vindex, 'E_TPF') ;
end;

function TPieceCompta.GetCodeTva(vIndex: Integer): String;
begin
  result := GetString( vindex, 'E_TVA') ;
  if result = '' then
    result := Contexte.CodeTvaDevaut ;
end;

procedure TPieceCompta.GetBornesPiece(vNumLigne: Integer; var vInDe, vInA: integer);
begin

  vInDe := 1 ;
  vInA  := Detail.count ;

  if isOut( vNumLigne ) then
    Exit ;

  if ModeSaisie = msBor then
    GetBornesGroupe( GetNumGroupe( vNumLigne ), vInDe, vInA ) ;
    
end;

procedure TPieceCompta.SetTauxVolatil(vIndex: integer; vDev: RDevise);
var lStCode : string ;
    lDtDate : TDateTime ;
begin

  if (modeSaisie = msBOR) and not isOut( vIndex )
    then lStCode := GetString( vIndex, 'E_DEVISE' )
    else lStCode := GetEnteteS( 'E_DEVISE' ) ;
  if lStCode = V_PGI.DevisePivot then Exit ;

  if (modeSaisie <> msPiece) and not isOut( vIndex )
    then lDtDate := GetValue( vIndex, 'E_DATECOMPTABLE' )
    else lDtDate := GetEnteteDt( 'E_DATECOMPTABLE' ) ;

  if Info_LoadDevise( lStCode ) then
    begin
    Info.Devise.SetTauxVolatil( lDtDate, vDev );
    FDevise.Taux := vDev.Taux ;
    ChangeTauxPiece( vIndex ) ;
    end ;

{  FDevise.Taux := vTaux ;

  if not isOut( vIndex ) then
    begin
    if Getdouble(vIndex, 'E_TAUXDEV')<>vTaux then
      ChangeTauxPiece( vIndex ) ;
    end ;}
end;

function TPieceCompta.EstTauxVolatil(vIndex: Integer): boolean;
var lStCode : string ;
    lDtDate : TDateTime ;
begin

  result := False ;

  if (modeSaisie = msBOR) and not isOut( vIndex )
    then lStCode := GetString( vIndex, 'E_DEVISE' )
    else lStCode := GetEnteteS( 'E_DEVISE' ) ;
  if lStCode = V_PGI.DevisePivot then Exit ;

  if (modeSaisie <> msPiece) and not isOut( vIndex )
    then lDtDate := GetValue( vIndex, 'E_DATECOMPTABLE' )
    else lDtDate := GetEnteteDt( 'E_DATECOMPTABLE' ) ;

  if Info_LoadDevise( lStCode ) then
    result := Info.Devise.EstTauxVolatil( lDtDate );

end;


function TPieceCompta.GetCopyTob(vIndex: Integer): Tob;
var lTobLigne : Tob ;
begin
  result := nil ;
  lTobLigne := CGetTob( vIndex ) ;
  if not Assigned( lTobLigne ) then Exit ;
  result := Tob.Create('ECRITURE', nil, -1 ) ;
  result.dupliquer( lTobLigne, True, true ) ;
end;

procedure TPieceCompta.AffecteEnteteScenario(vIndex: integer ; vBoForce : Boolean );
var i : integer ;
    procedure _traiteChaine( vStChp : string ) ;
      begin
      if Entete.GetString( vStchp ) <> '' then
        PutValue( vIndex, vStChp, Entete.GetString( vStchp ) ) ;
      end ;

    procedure _traiteDate( vStChp : string ) ;
      begin
      if Entete.GetDateTime( vStchp ) > iDate1900 then
        PutValue( vIndex, vStChp, Entete.GetDateTime( vStchp ) ) ;
      end ;

    procedure _traiteDouble( vStChp : string ) ;
      begin
      if Entete.GetDouble( vStChp ) <> 0 then
        PutValue( vIndex, vStChp, Entete.GetDouble( vStchp ) ) ;
      end ;

begin

  if ModeSaisie<>msPiece then Exit ;
  if (not vBoForce) and ( GetScenario( vIndex ) = nil ) then Exit ;

  // Zones extra-comptables
  _traiteChaine ( 'E_REFINTERNE'     ) ; // ref interne
  _traiteChaine ( 'E_LIBELLE'        ) ; // entête
  _traiteChaine ( 'E_REFEXTERNE'     ) ; // ref externe
  _traiteDate   ( 'E_DATEREFEXTERNE' ) ; // date refexterne
  _traiteChaine ( 'E_AFFAIRE'        ) ; // Affaire
  _traiteChaine ( 'E_REFLIBRE'       ) ; // ref libre
  _traiteDate   ( 'E_LIBREDATE'      ) ; // date libre

  // zones libres
  _traiteChaine( 'E_LIBREBOOL0'      ) ;
  _traiteChaine( 'E_LIBREBOOL0'      ) ;
  _traiteChaine( 'E_TABLE0'          ) ;
  _traiteChaine( 'E_TABLE1'          ) ;
  _traiteChaine( 'E_TABLE2'          ) ;
  _traiteChaine( 'E_TABLE3'          ) ;
  _traiteDouble( 'E_LIBREMONTANT0'   ) ;
  _traiteDouble( 'E_LIBREMONTANT1'   ) ;
  _traiteDouble( 'E_LIBREMONTANT2'   ) ;
  _traiteDouble( 'E_LIBREMONTANT3'   ) ;
  for i:=0 to 9 do
    _traiteChaine( 'E_LIBRETEXTE' + IntToStr(i) ) ;

end;

procedure TPieceCompta.InitEnteteScenario;
var i : integer ;
begin

  // Zones extra-comptables
  putEntete ( 'E_REFINTERNE',       '' ) ;         // ref interne
  putEntete ( 'E_LIBELLE',          '' ) ;        // libellé
  putEntete ( 'E_REFEXTERNE',       '' ) ;        // ref externe
  putEntete ( 'E_DATEREFEXTERNE',   iDate1900 ) ; // date refexterne
  putEntete ( 'E_AFFAIRE',          '' ) ;        // Affaire
  putEntete ( 'E_REFLIBRE',         '' ) ;        // ref libre
  putEntete ( 'E_LIBREDATE',        iDate1900 ) ; // date libre

  // zones libres
  putEntete( 'E_LIBREBOOL0',        '-' ) ;
  putEntete( 'E_LIBREBOOL0',        '-' ) ;
  putEntete( 'E_TABLE0',            '' ) ;
  putEntete( 'E_TABLE1',            '' ) ;
  putEntete( 'E_TABLE2',            '' ) ;
  putEntete( 'E_TABLE3',            '' ) ;
  putEntete( 'E_LIBREMONTANT0',     0 ) ;
  putEntete( 'E_LIBREMONTANT1',     0 ) ;
  putEntete( 'E_LIBREMONTANT2',     0 ) ;
  putEntete( 'E_LIBREMONTANT3',     0 ) ;
  for i:=0 to 9 do
    putEntete( 'E_LIBRETEXTE' + IntToStr(i), '' ) ;

end;

function TPieceCompta.IsValidEngagement: Boolean;
var i       : integer ;
begin

  result := False ;
  if Contexte.EstEngagementHT then
    // Engagement HT : il faut au moins une ligne assimilée HT
    result  := FindFirst( ['E_TYPEMVT'], [ 'HT' ], False ) <> nil
  else
    // Engagement TTC : il faut au moins une ligne avec un auxi + assimilée TTC
    for i:=1 to Detail.count do
      begin
      if ( GetString( i, 'E_AUXILIAIRE') <> '' ) and ( GetString( i, 'E_TYPEMVT') = 'TTC' ) then
          begin
          result := True ;
          break;
          end ;
      end ;

  if not result then
    if Contexte.EstEngagementHT
      then NotifyError( 0, 'La pièce d''engagement doit contenir au moins une ligne avec un compte HT' )
      else NotifyError( 0, 'La pièce d''engagement doit contenir au moins une ligne TTC avec un auxiliaire' ) ;

end;

procedure TPieceCompta.CSupprimeLigneVide;
var lTobEcr : TOB ;
    i    : integer ;
begin
  if (Count=0) then exit;
  for i:=Count downto 1 do
    begin
    lTobEcr := CGetTob( i ) ;
    if (lTobEcr.GetDouble('E_DEBITDEV')=0) and (lTobEcr.GetDouble('E_CREDITDEV')=0) then
      begin
      CFreeExtension( lTobEcr ) ;
      lTobEcr.Free ;
      end ;
    end;
end;

procedure TPieceCompta.CFreeExtension   ( vTobEcr : Tob ) ; 
begin
  CFreeTOBCompl( vTobEcr ) ; // libération de l'EcrCompl
end;

function TPieceCompta.GetCurrentTob: TOB;
begin
  result := nil ;
  if not isOut( FCurIdx ) then
    result := CGetTob( FCurIdx ) ;
end;

function TPieceCompta.EstEnregModifiable: boolean;
var lTobFirst : Tob ;
begin
  result := True ;
  if (Detail.count = 0) then Exit ;

  lTobFirst := CGetTob(1) ;

  // Test si la pièce à subit une suppression
  result := lTobFirst.GetString('E_CREERPAR') <> 'DET' ;

  // Test si la pièce ne fait pas partie d'une période close (cloture périodique)
  if result then
    begin
    result := not ( ( lTobFirst.GetDateTime('E_DATECOMPTABLE') < GetEncours.Deb ) or
                    ( GetParamSocSecur('SO_DATECLOTUREPER',iDate1900,true) >= lTobFirst.GetDateTime('E_DATECOMPTABLE') ) ) ;
    if not result then
      NotifyError(RC_PERIODECLOSE, '') ;
    end ;

  // Test si la pièce n'est pas validé
  if result then
    begin
    result := lTobFirst.GetString('E_VALIDE')<>'X' ;
    if not result then
      NotifyError(RC_BORCLOTURE, '') ;
    end; // if

 // TEST IMPORT COM
  if result then
    begin
    result := lTobFirst.GetString('E_ETATREVISION')<>'X' ;
    if not result then
      NotifyError(RC_BORREVISION, '') ;
    end ;

end;

function TPieceCompta.GereQtePCL ( vIndex : integer ) : Boolean;
begin
  result := EstRemplit( vindex ) and Contexte.CPPCLSAISIEQTE and Info.LoadCompte( GetString( vIndex, 'E_GENERAL') ) and
             ( ( ( Info.GetString('G_QUALIFQTE1') <> '' ) and ( Info.GetString('G_QUALIFQTE1') <> 'AUC' ) )
               or
               ( ( Info.GetString('G_QUALIFQTE2') <> '' ) and ( Info.GetString('G_QUALIFQTE2') <> 'AUC' ) )
             ) ;
end;

function TPieceCompta.IsValidDateCreation ( vDtDateC : TDateTime ) : boolean;
var lDtCloPer : TDateTime ;
    lInNbJAv  : integer ;
    lInNbJAp  : integer ;
    lStExo    : string ;
begin

  result := False ;

  // Test Date dans Exo ouvert
  lStExo := Contexte.GetExoDT( vDtDateC ) ;
  if (lStExo = Contexte.Exercices.EnCours.Code) and
     ( (vDtDateC > Contexte.Exercices.Encours.Fin) or (vDtDateC < Contexte.Exercices.Encours.Deb) ) then
    lStExo:='' ;
  if not ((lStExo=Contexte.Exercices.EnCours.Code) or (lStExo=Contexte.Exercices.Suivant.Code)) then
    begin
    NotifyError(0, TraduireMemoire('La date d''entrée est incompatible avec la saisie.') ) ;
    Exit ;
    end ;

  // test date sur période non close
  lDtCloPer := Contexte.DateCloturePer ;
  if ( (lDtCloPer>0) and (vDtDateC<=lDtCloPer) ) then
      begin
      NotifyError(0, TraduireMemoire('La date que vous avez renseignée est antérieure à une clôture.') ) ;
      Exit ;
      end ;

  // test plage de date en mode création
  lInNbJAv  := Contexte.NbJEcrAvant ;
  lInNbJAp  := Contexte.NbJEcrApres ;
  if ( (lInNbJAv > 0) and (vDtDateC<V_PGI.DateEntree) and (V_PGI.DateEntree-vDtDateC > lInNbJAv) )
     or
     ( (lInNbJAp > 0) and (vDtDateC>V_PGI.DateEntree) and (vDtDateC-V_PGI.DateEntree > lInNbJAp) ) then
    begin
    NotifyError(0, TraduireMemoire('La date que vous avez renseignée est en dehors des limites autorisées.') ) ;
    Exit ;
    end ;

  result := True ;

end;

function TPieceCompta.UnSeulTTC(vIndex: integer): Boolean;
var i           : Integer ;
    lInDe       : integer ;
    lInA        : integer ;
    lStTiers    : string  ;
    lStRef      : string ;
begin

  // En mode bordereau se limiter au groupe
  if ModeSaisie = msBOR then
    GetBornesGroupe( GetInteger( vIndex, 'E_NUMGROUPEECR') , lInDe, lInA )
  else
    begin
    lInDe := 1 ;
    lInA  := Detail.count ;
    end ;

  lStRef      := '' ;

  for i := lInDe to lInA do
    begin
    if Detail[i-1].GetInteger('E_NUMECHE') > 1 then continue ;

    lStTiers := GetTiers( i ) ;
    if lStTiers<>'' then
      begin
      // 1er tiers / TTC
      if (lStRef = '') then
        lStRef := lStTiers
      else
      // 2ème tiers / TTC
        begin
        lStRef := '' ;
        break ;
        end ;
      end ;

    end ;

  result := lStRef <> '' ;

end;

{ TPieceContexte }

procedure TPieceContexte.ChargeInfos;
begin

  // chargement des taux de tva
  ChargeTva ;

  if local then
    FBoInsertSpecif  := EstSpecif('51215')   // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
  else
    // en multi, sur une autre base, on charge les paramètres la 1ère fois
    begin
    FSociete          := GetParamsocDossierSecur('SO_SOCIETE',        '', FDossier ) ;
    FLibelle          := GetParamsocDossierSecur('SO_LIBELLE',        '', FDossier ) ;
    FCPChampDoublon   := GetParamsocDossierSecur('SO_CPCHAMPDOUBLON', '', FDossier ) ;
    FAttribRIBAuto    := GetParamsocDossierSecur('SO_ATTRIBRIBAUTO',  False, FDossier ) ;
    FRegimeDefaut     := GetParamsocDossierSecur('SO_REGIMEDEFAUT',   '', FDossier ) ;
    FCodeTvaDevaut    := GetParamsocDossierSecur('SO_CODETVAGENEDEFAULT',   '', FDossier ) ;
    FEtablisDefaut    := GetParamsocDossierSecur('SO_ETABLISDEFAUT',  '', FDossier ) ;
    FEtablisCpta      := GetParamsocDossierSecur('SO_ETABLISCPTA',    '', FDossier ) ;
    FPaysLocalisation := GetParamsocDossierSecur('SO_PAYSLOCALISATION', '', FDossier) ;
    if (FPaysLocalisation<>CodeISOFR) and (FPaysLocalisation<>CodeISOES) then
       FPaysLocalisation:=CodeISOFR ;
    FOuiTvaEnc        := GetParamsocDossierSecur('SO_OUITVAENC', False, FDossier) or (PaysLocalisation=CodeISOES) ;
    FTvaEncSociete    := GetParamsocDossierSecur('SO_TVAENCAISSEMENT',  '', FDossier) ;
    if FOuiTvaEnc then
      begin
      FCollCliEnc     := GetParamsocDossierSecur( 'SO_COLLCLIENC', '', FDossier) ;
      FCollFouEnc     := GetParamsocDossierSecur( 'SO_COLLFOUENC', '', FDossier) ;
      end
    else
      begin
      FCollCliEnc       := '' ;
      FCollFouEnc       := '' ;
      end ;
    FJalMultiEtab     := GetParamsocDossierSecur('SO_CPJALENCADECA',   '',          FDossier) ;  // FQ17766 SBO 01/05/2006
    FOuiTP            := GetParamsocDossierSecur('SO_OUITP',           False,       FDossier ) ; // Param Tiers Payeurs
    FJalATP           := GetParamsocDossierSecur('SO_JALATP',          '',          FDossier ) ;
    FJalVTP           := GetParamsocDossierSecur('SO_JALVTP',          '',          FDossier ) ;
    FDateRevision     := GetParamsocDossierSecur('SO_DATEREVISION',    idate1900,   FDossier) ;
    FCPAnoDyna        := GetParamsocDossierSecur('SO_CPANODYNA',       false,       FDossier) ;
    FCPAnaSurMontant  := GetParamsocDossierSecur('SO_CPAnaSurMontant', False,       FDossier) ;
    FZCtrlQTE         := GetParamsocDossierSecur('SO_ZCTRLQTE',        False,       FDossier) ;
    FTypeEngagement   := GetParamSocDossierSecur('SO_CPGESTENGAGE',    False,       FDossier) ;
    FCPPCLSAISIEQTE   := GetParamsocDossierSecur( 'SO_CPPCLSAISIEQTE', False,       FDossier ) ;
    FDateCloturePer   := GetParamsocDossierSecur( 'SO_DATECLOTUREPER', idate1900,   FDossier ) ;
    FNbJEcrAvant      := GetParamsocDossierSecur( 'SO_NBJECRAVANT',    0,           FDossier ) ;
    FNbJEcrApres      := GetParamsocDossierSecur( 'SO_NBJECRAPRES',    0,           FDossier ) ;
    FCPDateObli       := GetParamsocDossierSecur( 'SO_CPDATEOBLI',     False,       FDossier ) ;
    FCPEcheOuvrir     := GetParamsocDossierSecur( 'SO_ZSAISIEECHE',    False,       FDossier ) ;

    FBoInsertSpecif   := False ; // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
    end ;

end;

procedure TPieceContexte.ChargeLiensEtab;
var lQLiens : TQuery ;
begin

  if Assigned( FLiensEtab ) then
    begin
    FLiensEtab.ClearDetail ;
    FreeAndNil( FLiensEtab ) ;
    end ;

  FLiensEtab := Tob.Create( '_LIENSETAB', nil, -1 ) ;
  lQLiens := OpenSelect('SELECT * FROM CLIENSETAB WHERE CLE_ETABORIG="' + GetEntete('E_ETABLISSEMENT') + '"', FDossier ) ;
  FLiensEtab.LoadDetailDB( 'CLIENSETAB', '', '', lQLiens, False ) ;
  Ferme( lQLiens ) ;

end;

function TPieceContexte.ChargeScenario( vStJal, vStNatP, vStQualP, vStEtab : String ) : Boolean ;
begin
 result := FZScenario.Load( [ vStJal, vStNatP, vStQualP, vStEtab ]) >= 0 ;
end;

procedure TPieceContexte.ChargeTva;
var lQTva   : TQuery ;
    lNbBase : Integer ;
begin
  FillChar( NumCodeBase, Sizeof(NumCodeBase), #0) ;
  lQTva := OpenSelect('Select CC_CODE, CC_LIBRE from CHOIXCOD Where CC_TYPE="TX1" AND CC_LIBRE>"0" AND CC_LIBRE<="4"', FDossier ) ;
  While Not lQTva.EOF do
    begin
    lNbBase                := StrToInt( lQTva.FindField('CC_LIBRE').AsString ) ;
    NumCodeBase[ lNbBase ] := lQTva.FindField('CC_CODE').AsString ;
    lQTva.Next ;
    end ;
  Ferme(lQTva) ;
end;


destructor TPieceContexte.Destroy;
begin
  inherited;

  try

    if Assigned(FLiensEtab)  then
      begin
      FLiensEtab.ClearDetail ;
      FLiensEtab.free ;
      end ;

    if Assigned( FExercices ) then
      FExercices.Free ;

    if Assigned( FZHalleyUser ) then
      FZHalleyUser.Free ;

    if Assigned(FZScenario) then
      FZScenario.Free ;

  finally
    FTobEntete     := nil ;
    FLiensEtab     := nil ;
    FExercices     := nil ;
    FZHalleyUser   := nil ;
    FTobEntete     := nil ;
    FZScenario      := nil ;
  end ;

end;

function TPieceContexte.GetEntete(vChamps: String): Variant;
begin
  result := FTobEntete.GetValue( vChamps ) ;
end;

function TPieceContexte.GetLiensEtab(vEtabDest: String): TOB;
begin
  if not Assigned( FLiensEtab ) then ChargeLiensEtab;
  result := FLiensEtab.FindFirst( ['CLE_ETABDEST'], [ vEtabDest ], False ) ;
end;

function TPieceContexte.InitEntetePiece: TOB;
begin
  // Données entête d'une pièce
  result := TOB.Create('ECRITURE', nil, -1) ;

  result.PutValue('E_JOURNAL',          '' ) ;
  result.PutValue('E_NATUREPIECE',      'OD' ) ;
  result.PutValue('E_DATECOMPTABLE',    V_PGI.DateEntree ) ;
  result.PutValue('E_EXERCICE' ,        GetExoDt   ( V_PGI.DateEntree ) ) ;
  result.PutValue('E_PERIODE' ,         GetPeriode ( V_PGI.DateEntree ) ) ;
  result.PutValue('E_SEMAINE',          NumSemaine ( V_PGI.DateEntree ) ) ;
  result.PutValue('E_DEVISE' ,          V_PGI.DevisePivot ) ;
  result.PutValue('E_NUMEROPIECE' ,     0 ) ;
  result.PutValue('E_QUALIFPIECE' ,     'N' ) ;
  result.PutValue('E_MODESAISIE' ,      '-' ) ;
  result.PutValue('E_ECRANOUVEAU' ,     'N' ) ;
  result.PutValue('E_SOCIETE' ,         Societe ) ;
  result.PutValue('E_ETABLISSEMENT' ,   GetEtablisDefaut ) ;
  result.PutValue('E_REGIMETVA' ,       RegimeDefaut ) ;

  // On conserve un pointeur sur l'objet de gestion de l'entête de la pièce
  FTobEntete := result ;

end;

procedure TPieceContexte.Initialisation( vDossier : string = '' ) ;
begin

  if ( Trim( vDossier ) = '' )
    then FDossier := V_PGI.SchemaName
    else FDossier := vDossier ;

  FZScenario       := TZScenario.Create  ( FDossier ) ;
  FZHalleyUser     := TZHalleyUser.Create( FDossier ) ;
  FBoAttribRIBOff  := False ;
  ChargeInfos ;

end;

function TPieceContexte.InsertDB(vTobEcr: TOb): Boolean;
var lAxe       : Integer ;
    lAna       : Integer ;
    lTobTemp   : TOB ;
    lDBok      : Boolean ;
begin

  lDBOk  := False ;

  Try

    result := False ;

    // Cas classique
    if local then
      begin
      // SBO 01/07/2007 : enregistrement spécifique pour pb CWAS
      if FBoInsertSpecif
        then CTobInsertDB( vTobEcr )
        else vTobEcr.InsertDB( nil ) ;
      lDBOk := True ;
      end
    else
    // Cas spécifique au dossier !
      begin

      lTobTemp := Tob.Create( '$TEMP', nil, -1 ) ;
      try
        // Pour éviter l'insertion en cascade, on change les parents des ventilations analytiques
        for lAxe := (vTobEcr.Detail.Count - 1) downto 0 do
          vTobEcr.Detail[ lAxe ].ChangeParent( lTobTemp, -1 ) ;

        // Enregistrement de l'ecriture
        lDBok := InsertTobMS( vTobEcr, FDossier ) ;

        // Enregistrement de l'analytique
        if lDBok then
          for lAxe := 0 to lTobTemp.Detail.Count - 1 do
            for lAna := 0 to lTobTemp.Detail[ lAxe ].Detail.Count - 1 do
              begin
              lDBOk := InsertTobMS( lTobTemp.Detail[ lAxe ].Detail[ lAna ], FDossier ) ;
              if not lDbOk then break ;
              end ;

        // Remise en place des ventilations dans l'ecriture générale
        for lAxe := lTobTemp.Detail.Count - 1 downto 0 do
          lTobTemp.Detail[ lAxe ].ChangeParent( vTobEcr, -1 ) ;

      finally
        if Assigned(lTobTemp) then FreeAndNil(lTobTemp);
      end;
      end;

    result := True ;

  finally

    if not lDbOk then
      V_PGI.IoError := oeSaisie ;
  end ;

end;


function TPieceContexte.GetTypeExo(vDate: TDateTime): TTypeExo;
var
 lStCodeExo : string;
begin
  if local
    then result := CGetTypeExo( vDate ) // Fonction de Saiscomm.pas
    else begin
         // utilisation du TZExercice
         lStCodeExo := Exercices.QuelExoDT( vDate ) ;
         if lStCodeExo = Exercices.Encours.Code
           then result := teEncours
         else if lStCodeExo = Exercices.Suivant.Code
           then result := teSuivant
         else result := tePrecedent;
         end ;

end;

function TPieceContexte.GetLibelleEtab(vStCodeEtab: string): string;
var lQEtab : TQuery ;
begin
  if local then
    result := RechDom('TTETABLISSEMENT', vStCodeEtab, False )
  else
    begin
    lQEtab := OpenSelect( 'SELECT ET_LIBELLE FROM ETABLISS WHERE ET_ETABLISSEMENT = "' + vStCodeEtab + '"', Dossier ) ;
    if not lQEtab.Eof
      then result := lQEtab.FindField('ET_LIBELLE').AsString
      else result := vStCodeEtab ;
    Ferme( lQEtab ) ;
    end ;
end;

function TPieceContexte.GetExoDT(vDate: TDateTime): string;
begin
  if local
    then result := QuelExoDT( vDate ) // Fonction de Saiscomm.pas
    else result := Exercices.QuelExoDT( vDate ) ;
end;


function TPieceContexte.GetEtablisDefaut: string;
begin
{$IFNDEF NOVH}
  if local then
    begin
    if VH^.ProfilUserC[prEtablissement].Etablissement <> ''
      then result := VH^.ProfilUserC[prEtablissement].Etablissement
      else result := GetParamsocSecur('SO_ETABLISDEFAUT', '') ;
    end
  else
{$ENDIF}
    result := EtablisDefaut ;
end;

function TPieceContexte.GetScenario: TOB;
begin
  result := FZScenario.Item ;
end;


procedure TPieceContexte.GetZLastDate(var vDtDateC: TDateTime);
var lDtDatePCL     : TDateTime ;
    lExoRef        : TExoDate ;
begin

 lDtDatePCL    := GetParamsocDossierSecur('SO_ZLASTDATE', iDate1900, Dossier ) ;

 if (lDtDatePCL <> iDate1900) then
   begin
   lExoRef := Exercices.CPExoRef ;
   if lExoRef.Code <> '' then
     if ( lDtDatePCL >= lExoRef.Deb ) and ( lDtDatePCL <= lExoRef.Fin )
      then vDtDateC := lDtDatePCL ;
   end ;

end;

procedure TPieceContexte.GetZLastJal(var vStJournal: string);
var lStJalPCL     : string ;
begin

 lStJalPCL     := GetParamsocDossierSecur('SO_ZLASTJAL', '', Dossier ) ;

  if (vStJournal='') and (lStJalPCL<>'') then
      vStJournal := lStJalPCL ;

end;


procedure TPieceContexte.SetZLastInfo(vStJournal: string; vDtDateC: TDateTime);
begin
  SetParamSocDossier( 'SO_ZLASTJAL',  vStJournal, Dossier ) ;
  SetParamSocDossier( 'SO_ZLASTDATE', vDtDateC,   Dossier ) ;
end;

function TPieceContexte.ModeRevisionActive(vDtDate: TDateTime): boolean;
begin
  result := ( ( Not V_PGI.Controleur ) and
              ( DateRevision > iDate1900 ) and ( vDtDate <= DateRevision )
            ) ;
end;

procedure TPieceContexte.VideScenario;
begin
  FZScenario.Clear ;
end;

function TPieceContexte.EstEngagementHT: Boolean;
begin
  result := TypeEngagement = '2' ;
end;

function TPieceContexte.GetExercices: TZExercice;
begin
  if local then
    result := ctxExercice
  else
    begin
    if not Assigned( FExercices ) then
      FExercices   := TZExercice.Create( False, FDossier ) ;
    result := FExercices ;
    end ;
end;


function TPieceContexte.GetAttribRIBAuto: Boolean;
begin
  if FBoAttribRIBOff then
    result := False
  else if local then
    result := GetParamsocSecur('SO_ATTRIBRIBAUTO',      False)
  else
    result := FAttribRIBAuto ;
end;

function TPieceContexte.GetCodeTvaDevaut: string;
begin
  if local
    then result := GetParamsocSecur('SO_CODETVAGENEDEFAULT',        '')
    else result := FCodeTvaDevaut ;
end;

function TPieceContexte.GetCollCliEnc: String;
begin
  if local then
    begin
    if OuiTvaEnc then
      result     := GetParamsocSecur( 'SO_COLLCLIENC',      '') ;
    end
  else
    result := FCollCliEnc ;
end;

function TPieceContexte.GetCollFouEnc: String;
begin
  if local then
    begin
    if OuiTvaEnc then
      result     := GetParamsocSecur( 'SO_COLLFOUENC',      '') ;
    end
  else
    result := FCollFouEnc ;
end;

function TPieceContexte.GetCPAnaSurMontant: boolean;
begin
  if local
    then result := GetParamsocSecur('SO_CPAnaSurMontant',     False )
    else result := FCPAnaSurMontant ;
end;

function TPieceContexte.GetCPAnoDyna: boolean;
begin
  if local
    then result := _GetParamsocSecur('SO_CPANODYNA',     False )
    else result := FCPAnoDyna ;
end;

function TPieceContexte.GetCPChampDoublon: String;
begin
  if local
    then result := GetParamsocSecur('SO_CPCHAMPDOUBLON',        '')
    else result := FCPChampDoublon ;
end;

function TPieceContexte.GetDateRevision: TDateTime;
begin
  if local
    then result := GetParamsocSecur('SO_DATEREVISION',  idate1900 )
    else result := FDateRevision ;
end;

function TPieceContexte.GetEtablisCpta: Boolean;
begin
  if local
    then result := GetParamsocSecur('SO_ETABLISCPTA',   False )
    else result := FEtablisCpta ;
end;

function TPieceContexte.GetJalATP: string;
begin
  if local
    then result := GetParamsocSecur('SO_JALATP',        '')
    else result := FJalATP ;
end;

function TPieceContexte.GetJalMultiEtab: string;
begin
  if local
    then result := GetParamsocSecur('SO_CPJALENCADECA',        '')
    else result := FJalMultiEtab ;
end;

function TPieceContexte.GetJalVTP: string;
begin
  if local
    then result := GetParamsocSecur('SO_JALVTP',        '')
    else result := FJalVTP ;
end;

function TPieceContexte.GetLibelle: string;
begin
  if local
    then result := GetParamsocSecur('SO_LIBELLE',   '' )
    else result := FLibelle ;
end;

function TPieceContexte.GetOuiTP: Boolean;
begin
  if local
    then result := GetParamsocSecur('SO_OUITP',     False )
    else result := FOuiTP ;
end;

function TPieceContexte.GetOuiTvaEnc: boolean;
begin
  if local
    then result := GetParamsocSecur('SO_OUITVAENC', False) or (PaysLocalisation=CodeISOES)
    else result := FOuiTvaEnc ;
end;

function TPieceContexte.GetPaysLocalisation: String;
begin
  if local
    then result := GetParamsocSecur('SO_PAYSLOCALISATION',  CodeISOFR )
    else result := FPaysLocalisation ;
end;

function TPieceContexte.GetRegimeDefaut: String;
begin
  if local
    then result := GetParamsocSecur('SO_REGIMEDEFAUT',        '')
    else result := FRegimeDefaut ;
end;

function TPieceContexte.GetSociete: string;
begin
  if local
    then result := GetParamsocSecur('SO_SOCIETE',        '')
    else result := FSociete ;
end;

function TPieceContexte.GetTvaEncSociete: String;
begin
  if local
    then result := GetParamsocSecur('SO_TVAENCAISSEMENT',        '')
    else result := FTvaEncSociete ;
end;

function TPieceContexte.GetTypeEngagement: String;
begin
  if local
    then result := GetParamsocSecur('SO_CPGESTENGAGE',     '')
    else result := FTypeEngagement ;
end;

function TPieceContexte.GetZCtrlQTE: boolean;
begin
  if local
    then result := GetParamsocSecur('SO_ZCTRLQTE',     False )
    else result := FZCtrlQTE ;
end;

function TPieceContexte.GetFEtablisDefaut: String;
begin
  if local
    then result := GetParamsocSecur('SO_ETABLISDEFAUT',        '')
    else result := FEtablisDefaut ;
end;


function TPieceContexte.GetCPPCLSAISIEQTE: boolean;
begin
  if local
    then result := GetParamsocSecur('SO_CPPCLSAISIEQTE',  False )
    else result := CPPCLSAISIEQTE ;
end;

function TPieceContexte.GetDateCloturePer: TDateTime;
begin
  if local
    then result := GetParamsocSecur('SO_DATECLOTUREPER',  0 )
    else result := FDateCloturePer ;
end;

function TPieceContexte.GetNbJEcrApres: integer;
begin
  if local
    then result := GetParamsocSecur('SO_NBJECRAPRES',  0 )
    else result := FNbJEcrApres ;
end;

function TPieceContexte.GetNbJEcrAvant: integer;
begin
  if local
    then result := GetParamsocSecur('SO_NBJECRAVANT',  0 )
    else result := FNbJEcrAvant ;
end;

function TPieceContexte.GetCPDateObli: boolean;
begin
  if local
    then result := GetParamsocSecur('SO_CPDATEOBLI',  False)
    else result := FCPDateObli ;
end;

function TPieceContexte.GetCPEcheOuvrir: boolean;
begin
  if local
    then result := GetParamsocSecur('SO_ZSAISIEECHE',  False)
    else result := FCPEcheOuvrir ;
end;

{ TDossierContexte }


class function TPieceContexte.CreerContexte( vDossier: String ) : TPieceContexte;
begin
  result := TPieceContexte.Create;
  result.Initialisation( vDossier ) ;
end;


function TDossierContexte.local: Boolean;
begin
  result := ( FDossier = V_PGI.SchemaName ) ;
end;

{ TobPieceCompta }

{---------------------------------------------------------------------------------------}
constructor TobPieceCompta.Create(LeNomTable: string; LeParent: TOB; IndiceFils: Integer);
{---------------------------------------------------------------------------------------}
begin
  inherited Create(LeNomTable, LeParent, IndiceFils);
  lInfoEcriture := TStringList.Create;
end;

{---------------------------------------------------------------------------------------}
destructor TobPieceCompta.Destroy;
{---------------------------------------------------------------------------------------}
begin
  LibereLaListe(True);
  inherited Destroy;
end;

{---------------------------------------------------------------------------------------}
procedure TobPieceCompta.AddInfoEcriture(NomBase : string; Inf : TInfoEcriture);
{---------------------------------------------------------------------------------------}
begin
  if Assigned(Inf) then lInfoEcriture.AddObject(NomBase, Inf);
end;

{Se contente de renvoyer le Bon TInfoEcriture s'il existe
{---------------------------------------------------------------------------------------}
function TobPieceCompta.GetInfoEcriture(NomBase : string) : TInfoEcriture;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  Result := nil;
  n := lInfoEcriture.IndexOf(NomBase);
  if n > - 1 then Result := TInfoEcriture(lInfoEcriture.Objects[n]);
end;

{A la différence de GetInfoEcriture, crée le TInfoEcriture s'il n'existe pas
{---------------------------------------------------------------------------------------}
function TobPieceCompta.CreateInfoEcr(NomBase : string) : TInfoEcriture;
{---------------------------------------------------------------------------------------}
begin
  Result := GetInfoEcriture(NomBase);
  {Si le TInfoEcriture n'existe pas ...}
  if not Assigned(Result) then begin
    {... on le créé ...}
    Result  := TInfoEcriture.Create(NomBase);
    {... et on le mémorise}
    if Assigned(Result) then AddInfoEcriture(NomBase, Result);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TobPieceCompta.ClearDetailPC;
{---------------------------------------------------------------------------------------}
begin
  ClearDetail;
  LibereLaListe(False);
end;

{---------------------------------------------------------------------------------------}
procedure TobPieceCompta.LibereLaListe(Detruire: Boolean);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  o : TObject;
begin
  if Assigned(lInfoEcriture)  then begin
    for n := 0 to lInfoEcriture.Count - 1 do
      if Assigned(lInfoEcriture.Objects[n]) then begin
        o := lInfoEcriture.Objects[n];
        FreeAndNil(o);
      end;

    if Detruire then FreeAndNil(lInfoEcriture)
                else lInfoEcriture.Clear;
  end;
end;

{ TMultiEche }

procedure TMultiEche.AddEche( vTobEcr: TOB );
begin

  if FModeFonc = mfPiece then Exit ;

  if vTobEcr.GetInteger('E_NUMECHE') = 0 then Exit ;
  if vTobEcr.GetInteger('E_NUMLIGNE') <> NumLigne then Exit ;

  vTobEcr.ChangeParent( self,  0 ) ;

end;

procedure TMultiEche.CalculEche( vStModeForce: String );
begin
  if FStModeInit = ''
    then FStModeInit := GetModeInit ;
  if vStModeForce<>'' then
    FStModeFinal := vStModeForce ;
  CCalculEche ( Piece, GetEcrIndex, FStModeFinal, Info, Piece.ModeEche=meMono ) ;
end;

procedure TMultiEche.CalculModeFinal;
var lTobMR : Tob ;
    lTotal : Double ;
begin
  if FMontantSaisi <> 0
    then lTotal := FMontantSaisi
    else lTotal := GetTotalEche ;

  if FStModeInit = ''
    then FStModeInit := GetModeInit ;

  lTobMR := Info.ModeRegle.CalculModeFinal( lTotal, FStModeInit ) ;

  FStModeFinal := lTobMR.GetString('MR_MODEREGLE') ;
  
end;

procedure TMultiEche.ClearEche;
var lInIdx : integer ;
    lNbLig : integer ;
    i      : integer ;
begin
  Case FModeFonc of

    mfAlone : Detail.Clear ;
    
    mfPiece : begin
              lInIdx := GetEcrIndex ;
              lNbLig := nbEche ;
              for i := (lNbLig + lInIdx - 1) downto (lInIdx + 1) do
                Piece.Detail[i].Free ;
              SetMontant( 1, FMontantSaisi ) ;
              end ;
  end ;

  if TobEcr.GetString('E_ECHE')<>'X' then
    CSupprimerInfoLettrage( TobEcr ) ;

end;

procedure TMultiEche.CopieEcheances(vSource: TMultiEche);
var i : Integer ;
begin

  if Piece.ModeEche<>meDeporte then Exit ;

  FStModeFinal := vSource.ModeFinal ;
  ModifTVA := vSource.ModifTVA ;
  FBoModif     := vSource.ModifEnCours ;

  ClearDetail ;
  for i := 1 to vSource.Detail.count do
    begin
    NewEche ;
    PutTobEche( i, vSource.GetEche( i )  ) ;
    end ;

end;

constructor TMultiEche.Create( vTobParent : TOB ; vPiece : TPieceCompta ; vTobEcr : Tob ; vAction : TActionFiche ; vModeFonc : TModeFonc ) ;
begin
  inherited Create('V_ECHEANCES', vTobParent, -1) ;

  FPiece        := vPiece ;


  FAction       := vAction ;
  FBoModif      := False ;
  FModeFonc     := vModeFonc ;
  FStModeInit   := '' ;
  FStModeFinal  := '' ;


  FTobEcr       := vTobEcr ;
  if Assigned( vTobEcr ) then
    FStModeInit   := GetModeInit ;

  if vTobEcr.GetString('E_MODEPAIE')=''
    then FMontantSaisi := vTobEcr.GetDouble('E_DEBITDEV') + vTobEcr.GetDouble('E_CREDITDEV')
    else FMontantSaisi := GetTotalEche ;

  if vTobEcr.Parent <> vPiece then
    FModeFonc     := mfAlone ;

  if FModeFonc = mfAlone then
    begin
    NewEche ;
    SetMontant( 1, FMontantSaisi ) ;
    end ;

end;

procedure TMultiEche.CumuleTobEcr;

  function _GetCumulEche ( vStChamps : String ) : Double ;
  var i : Integer ;
  begin
    result := 0 ;
    for i := 0 to ( NbEche - 1 ) do
      result := result + Detail[ i ].GetDouble( vStChamps ) ;
  end ;

begin
    if Piece.ModeEche <> meDeporte then Exit ;

    TobEcr.PutValue( 'E_DEBIT',         _GetCumulEche('E_DEBIT' ) ) ;
    TobEcr.PutValue( 'E_CREDIT',        _GetCumulEche('E_CREDIT') ) ;
    TobEcr.PutValue( 'E_DEBITDEV',      _GetCumulEche('E_DEBITDEV') ) ;
    TobEcr.PutValue( 'E_CREDITDEV',     _GetCumulEche('E_CREDITDEV') ) ;
    TobEcr.PutValue( 'E_COUVERTURE',    _GetCumulEche('E_COUVERTURE') ) ;
    TobEcr.PutValue( 'E_COUVERTUREDEV', _GetCumulEche('E_COUVERTUREDEV') ) ;
    TobEcr.PutValue( 'E_ECHEENC1',      _GetCumulEche('E_ECHEENC1') ) ;
    TobEcr.PutValue( 'E_ECHEENC2',      _GetCumulEche('E_ECHEENC2') ) ;
    TobEcr.PutValue( 'E_ECHEENC3',      _GetCumulEche('E_ECHEENC3') ) ;
    TobEcr.PutValue( 'E_ECHEENC4',      _GetCumulEche('E_ECHEENC4') ) ;
    TobEcr.PutValue( 'E_ECHEDEBIT',     _GetCumulEche('E_ECHEDEBIT') ) ;
end;

destructor TMultiEche.Destroy;
begin
  inherited;
end;

function TMultiEche.EstDebit: Boolean;
begin
  result := TobEcr.GetDouble('E_DEBITDEV')<>0 ;
end;

function TMultiEche.EstLettre(vNumEche: Integer): Boolean;
begin
  result := False ;
  if IsOutEche( vNumEche ) then Exit ;
  result := Trim( GetEche( vNumEche ).GetString('E_LETTRAGE') ) <> '' ;
end;

function TMultiEche.EstModifiable(vNumEche: Integer): Boolean;
var i : Integer ;
begin
  result := True ;

  // Demande pour 1 ligne
  if (vNumEche > 0) then
    begin
    if FModeFonc = mfPiece
      then result := Piece.EstLigneModifiable( GetEcrIndex + vNumEche )
      else result := not EstLettre( vNumEche ) ;
    end
  else
  // demande pour toutes les lignes
     for i := GetEcrIndex to DetailEche.Count-1 do
       begin
       if DetailEche[i].GetInteger('E_NUMLIGNE')<>NumLigne then Break ;
       if FModeFonc = mfPiece
         then result := result and Piece.EstLigneModifiable( i+1 )
         else result := result and not EstLettre( i+1 ) ;
       end ;
end;

function TMultiEche.EstOkMultiEche: boolean;
begin

  result := False ;
  if not Info.LoadCompte( TobEcr.GetString('E_GENERAL') ) then Exit ;

  Result :=  ( Piece.ModeEche <> meMono ) and
               not (   ( Info.GetString('G_NATUREGENE') = 'BQE' )
                    or ( Info.GetString('G_NATUREGENE') = 'CAI' )
                    or   Info.Compte.IsVentilable
                    or ( Info.Journal.GetValue('J_EFFET') = 'X' )
                    or ( Info.Journal.GetValue('J_NATUREJAL') = 'BQE' )
                    or ( Info.Journal.GetValue('J_NATUREJAL') = 'CAI' )
                  ) ;

end;

function TMultiEche.GereArrondi: Boolean;
var lMontant  : Double ;
begin
  result := False ;

  if NbEche < 1 then Exit ;
  if FMontantSaisi = 0 then Exit ;

  // Si le total des échéances n'est pas égale au montant de l'ecriture, on impacte la différence sur la dernière ligne
  if GetTotalEche <> FMontantSaisi then
    begin
    // Recalcul du montant de la ligne
    lMontant :=  GetMontantEche( NbEche - 1 ) + ( FMontantSaisi - GetTotalEche ) ;
    // MAJ du montant
    if EstDebit
      then CSetMontants( GetEche(NbEche-1),   lMontant,    0,           Devise,   True )
      else CSetMontants( GetEche(NbEche-1),   0,           lMontant,    Devise,   True ) ;
    // Indicateur de modification
    result := True ;
    end;
end;

procedure TMultiEche.GerePoucentage;
var lTotalEche : double ;
    lTobEche   : TOB ;
    lPourcent  : Double ;
    i          : Integer ;
begin

  // Montant Total
  if FMontantSaisi <> 0
    then lTotalEche := FMontantSaisi
    else lTotalEche := GetTotalEche ;

  if lTotalEche <> 0 then
    for i := GetEcrIndex to (DetailEche.Count - 1) do
      begin
      if DetailEche[i].GetInteger('E_NUMLIGNE')<>GetNumLigne then Break ;
      lTobEche := DetailEche[i] ;
      lPourcent := Arrondi( 100.0 * ( lTobEche.GetDouble('E_CREDITDEV') + lTobEche.GetDouble('E_DEBITDEV') ) / FMontantSaisi, ADecimP );
      if lTobEche.GetNumChamp('POURCENTAGE') > 0
        then lTobEche.PutValue          ( 'POURCENTAGE', lPourcent )
        else lTobEche.AddChampSupValeur ( 'POURCENTAGE', lPourcent ) ;
      end ;
end;

function TMultiEche.GetDetail: TTobList;
begin
  if FModeFonc = mfPiece
    then result := Piece.Detail
    else result := Detail ;
end;

function TMultiEche.GetDevise: RDevise;
begin
  result := Piece.GetRDevise( NumLigne ) ;
end;

function TMultiEche.GetEche(vNumEche: Integer): TOB;
var i : integer ;
begin
  result := nil ;
  if FModeFonc = mfAlone then
    begin
    if not isOutEche( vNumEche ) then
      result := Detail[ vNumEche - 1 ] ;
    end
  else
    for i := GetEcrIndex to (Piece.Detail.Count - 1) do
      begin
      if Piece.Detail[i].GetInteger('E_NUMLIGNE')<>GetNumLigne then Break ;
      if Piece.Detail[i].GetInteger('E_NUMECHE')=vNumEche then
        begin
        result := Piece.Detail[i] ;
        Break ;
        end
      end ;
end;

function TMultiEche.GetEcrIndex: integer;
begin
  if FModeFonc = mfALone
    then result := 0
    else result := TobEcr.GetIndex ;
end;

function TMultiEche.GetInfo: TInfoEcriture;
begin
  result := Piece.Info ;
end;

function TMultiEche.GetModeInit: String;
begin
  result := CGetModeRegleInit( TobEcr, Info ) ;
end;

function TMultiEche.GetModeRegl: T_MODEREGL;
var lTobEche : TOB ;
    i        : Integer ;
begin

  FillChar( result, Sizeof( result ), #0 ) ;

  if not EstModifiable
    then result.Action := taConsult
    else result.Action := taModif ;

  if Info.Compte.IsCollectif then
    begin
    result.Aux          := TobEcr.GetValue('E_AUXILIAIRE') ;
    result.ModeInitial  := Info.GetString('T_MODEREGLE') ;
    end
  else
    begin
    result.Aux          := TobEcr.GetValue('E_GENERAL') ;
    result.ModeInitial  := Info.GetString('G_MODEREGLE') ;
    end ;

  result.TotalAPayerD := GetTotalEche ;
  result.TotalAPayerP := GetTotalEche( False ) ;

  result.CodeDevise   := Devise.Code ;
  result.Symbole      := Devise.Symbole ;
  result.Quotite      := Devise.Quotite ;
  result.TauxDevise   := Devise.Taux ;
  result.Decimale     := Devise.Decimale ;

  result.DateFact     := TobEcr.GetDateTime('E_DATECOMPTABLE') ;
  result.DateBL       := TobEcr.GetDateTime('E_DATECOMPTABLE') ;
  result.DateFactExt  := TobEcr.GetDateTime('E_DATEREFEXTERNE') ;
  if result.DateFactExt <= IDate1900 then
    result.DateFactExt := TobEcr.GetDateTime('E_DATECOMPTABLE') ;

  result.JourPaiement1 := 0 ;
  result.JourPaiement2 := 0 ;

  result.NbEche       := NbEche ;
  result.ModeFinal    := ModeFinal ;

  // Détails
  for i := 1 to NbEche do
    begin

    lTobEche := GetEche( i ) ;
    if lTobEche = nil then Break ;
    if lTobEche.GetInteger('E_NUMLIGNE') <> NumLigne then Break ;

    result.TabEche[i].MontantD       := lTobEche.GetDouble ('E_CREDITDEV') + lTobEche.GetDouble('E_DEBITDEV') ;
    result.TabEche[i].MontantP       := lTobEche.GetDouble ('E_CREDIT')    + lTobEche.GetDouble('E_DEBIT') ;

    result.TabEche[i].ReadOnly       := Trim( lTobEche.GetString('E_LETTRAGE') ) <> '' ;

    result.TabEche[i].CodeLettre     := lTobEche.GetString   ('E_LETTRAGE') ;
    result.TabEche[i].ModePaie       := lTobEche.GetString   ('E_MODEPAIE') ;
    result.TabEche[i].DateEche       := lTobEche.GetDateTime ('E_DATEECHEANCE') ;
    result.TabEche[i].DateValeur     := lTobEche.GetDateTime ('E_DATEVALEUR') ;
    result.TabEche[i].DateRelance    := lTobEche.GetDateTime ('E_DATERELANCE') ;
    result.TabEche[i].NiveauRelance  := lTobEche.GetInteger  ('E_NIVEAURELANCE') ;
    result.TabEche[i].Couverture     := lTobEche.GetDouble   ('E_COUVERTURE') ;
    result.TabEche[i].CouvertureDev  := lTobEche.GetDouble   ('E_COUVERTUREDEV') ;
    result.TabEche[i].LettrageDev    := lTobEche.GetString   ('E_LETTRAGEDEV') ;
    result.TabEche[i].DatePaquetMax  := lTobEche.GetDateTime ('E_DATEPAQUETMAX') ;
    result.TabEche[i].DatePaquetMin  := lTobEche.GetDateTime ('E_DATEPAQUETMIN') ;
    result.TabEche[i].EtatLettrage   := lTobEche.GetString   ('E_ETATLETTRAGE') ;

    // Calcul du pourcentage...
    if result.TotalAPayerP <> 0 then
      result.TabEche[i].Pourc        := Arrondi( 100.0 * result.TabEche[i].MontantP / result.TotalAPayerP , ADecimP ) ;

    // #TVAENC
    if Piece.EstTvaLoc then
      begin
      result.TabEche[i].TAV[1] := lTobEche.GetDouble( 'E_ECHEENC1' ) ;
      result.TabEche[i].TAV[2] := lTobEche.GetDouble( 'E_ECHEENC2' ) ;
      result.TabEche[i].TAV[3] := lTobEche.GetDouble( 'E_ECHEENC3' ) ;
      result.TabEche[i].TAV[4] := lTobEche.GetDouble( 'E_ECHEENC4' ) ;
      result.TabEche[i].TAV[5] := lTobEche.GetDouble( 'E_ECHEDEBIT' ) ;
      end ;

    end ; // fin For

end;

function TMultiEche.GetModifTVA: Boolean;
begin
  if TOBEcr.GetNumChamp('MODIFTVA') <= 0
    then result := False
    else result := TobEcr.GetString('MODIFTVA')='X' ;
end;

function TMultiEche.GetMontantEche(vNumEche: Integer): Double;
var lTobEche : TOB ;
begin
  result := 0 ;
  lTobEche := GetEche( vNumEche ) ;
  if lTobEche = nil then Exit ;
  result := Arrondi( lTobEche.GetDouble('E_CREDITDEV') + lTobEche.GetDouble('E_DEBITDEV'), Devise.Decimale ) ;
end;

function TMultiEche.GetMontantSaisi: Double;
begin
  if ( Piece.ModeEche = meDeporte )
    then result := TobEcr.GetDouble('E_DEBITDEV') - TobEcr.GetDouble('E_CREDITDEV')
    else result := FMontantSaisi ;
end;

function TMultiEche.GetNbEche: Integer;
var i : integer ;
begin
  result := 0 ;
  if FModeFonc = mfAlone then
    result := Detail.count
  else
    for i := GetEcrIndex to (DetailEche.Count - 1) do
      begin
      if DetailEche[i].GetInteger('E_NUMLIGNE')<>GetNumLigne then Break ;
      Inc(result) ;
      end ;

end;

function TMultiEche.GetNumLigne: Integer;
begin
  result := FTobEcr.GetInteger('E_NUMLIGNE') ;
end;

function TMultiEche.GetTotalEche( vBoDev : Boolean )  : Double ;
var i : integer ;
begin
  result := 0 ;
  for i := GetEcrIndex to (DetailEche.Count - 1) do
    begin
    if DetailEche[i].GetInteger('E_NUMLIGNE')<>GetNumLigne then Break ;
    if vBoDev
      then  result := result + DetailEche[i].GetDouble('E_DEBITDEV') + DetailEche[i].GetDouble('E_CREDITDEV')
      else  result := result + DetailEche[i].GetDouble('E_DEBIT') + DetailEche[i].GetDouble('E_CREDIT') ;
    end ;
  result := Arrondi( result, Devise.Decimale) ;
end;

function TMultiEche.IsOutEche(vNumEche: Integer): Boolean;
begin
  if FModeFonc = mfAlone
    then result := (vNumEche < 1) or (vNumEche > Detail.Count)
    else begin
         result :=    (vNumEche < 1 )
                   or ( ( GetIndex + vNumEche - 1 ) > Piece.Detail.count )
                   or ( Piece.Detail[ GetIndex + vNumEche - 1 ].GetInteger('E_NUMLIGNE')<>NumLigne ) ;
         end ;

end;

function TMultiEche.NewEche(vNumEche: Integer): TOB;
begin

  if (vNumEche < 1) or (vNumEche > NbEche )
    then vNumEche := NbEche + 1;

  case FModeFonc of

    mfPiece : begin
              result := Tob.Create('ECRITURE', Piece, vNumEche + GetEcrIndex - 1 ) ;
              result.Dupliquer( TobEcr, False, True, True ) ;
              end ;

    else      begin // msAlone
              result := Tob.Create('ECRITURE', self, vNumEche - 1 ) ;
              result.Dupliquer( TobEcr, False, True, True ) ;
              end ;

    end ;

  // renumérotation
  result.PutValue('E_NUMECHE', vNumEche ) ;

  if vNumEche <> NbEche then
    Renumerote( vNumEche ) ;

  // Remise à Zero des montants
  SetMontant( vNumEche, 0 ) ;

end;

procedure TMultiEche.ProratiseEche(vMontant: Double);
Var lTotal       : Double ;
    lTaux        : Double ;
    i            : Integer ;
    lMontantEche : Double ;
    lTobEche     : TOB ;
begin

  lTotal := GetTotalEche ;
  If lTotal = 0 then Exit ;

  lTaux := vMontant / lTotal ;
  if lTaux = 1 then Exit ;

  lTotal    := 0 ;

  // Recalcul des montants des échéances
  for i := GetEcrIndex to (DetailEche.Count - 1) do
    begin
    lTobEche := DetailEche[ i ] ;
    if lTobEche.GetInteger('E_NUMLIGNE')<>GetNumLigne then Break ;

    lMontantEche := lTobEche.GetDouble('E_DEBITDEV') + lTobEche.GetDouble('E_CREDITDEV') ;
    lMontantEche := Arrondi( lTaux * lMontantEche , Devise.Decimale ) ;

    if EstDebit
      then CSetMontants( lTobEche, lMontantEche, 0, Devise, True )
      else CSetMontants( lTobEche, 0, lMontantEche, Devise, True ) ;

    // #TVAENC
    if ((VH^.OuiTvaEnc) and (ModifTVA) and (lTaux<>1)) then
      begin
      lTobEche.PutValue('E_ECHEENC1',  Arrondi( lTaux * lTobEche.GetDouble('E_ECHEENC1') ,  V_PGI.OkDecV ) )  ;
      lTobEche.PutValue('E_ECHEENC2',  Arrondi( lTaux * lTobEche.GetDouble('E_ECHEENC2') ,  V_PGI.OkDecV ) )  ;
      lTobEche.PutValue('E_ECHEENC3',  Arrondi( lTaux * lTobEche.GetDouble('E_ECHEENC3') ,  V_PGI.OkDecV ) )  ;
      lTobEche.PutValue('E_ECHEENC4',  Arrondi( lTaux * lTobEche.GetDouble('E_ECHEENC4') ,  V_PGI.OkDecV ) )  ;
      lTobEche.PutValue('E_ECHEDEBIT', Arrondi( lTaux * lTobEche.GetDouble('E_ECHEDEBIT') , V_PGI.OkDecV ) )  ;
      end ;

    lTotal := lTotal + lMontantEche ;

    end ;

  // Gestion des différences d'arrondi
  if lTotal <> vMontant then
    begin
    lTobEche := GetEche( NbEche ) ;
    lMontantEche := lTobEche.GetDouble('E_DEBITDEV') + lTobEche.GetDouble('E_CREDITDEV') ;
    lMontantEche := lMontantEche + vMontant - lTotal ;
    if EstDebit
      then CSetMontants( lTobEche, lMontantEche, 0, Devise, True )
      else CSetMontants( lTobEche, 0, lMontantEche, Devise, True ) ;
    end ;

end;

procedure TMultiEche.PutTobEche(vNumEche: Integer; vTobSource: TOB);
var lTobEche : TOB ;
begin

  lTobEche := GetEche( vNumEche ) ;
  if lTobEche = nil then Exit ;

  lTobEche.PutValue('E_ECHE',            'X'  ) ;

  lTobEche.PutValue('E_NIVEAURELANCE',   vTobSource.GetValue('E_NIVEAURELANCE') ) ;
  lTobEche.PutValue('E_MODEPAIE',        vTobSource.GetValue('E_MODEPAIE')      ) ;
  lTobEche.PutValue('E_DATEECHEANCE',    vTobSource.GetValue('E_DATEECHEANCE')  ) ;
  lTobEche.PutValue('E_DATEVALEUR',      vTobSource.GetValue('E_DATEVALEUR')    ) ;
  lTobEche.PutValue('E_COUVERTURE',      vTobSource.GetValue('E_COUVERTURE')    ) ;
  lTobEche.PutValue('E_COUVERTUREDEV',   vTobSource.GetValue('E_COUVERTUREDEV') ) ;
  lTobEche.PutValue('E_ETATLETTRAGE',    vTobSource.GetValue('E_ETATLETTRAGE')  ) ;
  lTobEche.PutValue('E_LETTRAGE',        vTobSource.GetValue('E_LETTRAGE')      ) ;
  lTobEche.PutValue('E_LETTRAGEDEV',     vTobSource.GetValue('E_LETTRAGEDEV')   ) ;
  lTobEche.PutValue('E_DATEPAQUETMAX',   vTobSource.GetValue('E_DATEPAQUETMAX') ) ;
  lTobEche.PutValue('E_DATEPAQUETMIN',   vTobSource.GetValue('E_DATEPAQUETMIN') ) ;
  lTobEche.PutValue('E_DATERELANCE',     vTobSource.GetValue('E_DATERELANCE')   ) ;
  lTobEche.PutValue('E_DEBITDEV',        vTobSource.GetValue('E_DEBITDEV')      ) ;
  lTobEche.PutValue('E_DEBIT',           vTobSource.GetValue('E_DEBIT')         ) ;
  lTobEche.PutValue('E_CREDITDEV',       vTobSource.GetValue('E_CREDITDEV')     ) ;
  lTobEche.PutValue('E_CREDIT',          vTobSource.GetValue('E_CREDIT')        ) ;
  lTobEche.PutValue('E_ORIGINEPAIEMENT', vTobSource.GetValue('E_DATEECHEANCE')  ) ;
  lTobEche.PutValue('E_ENCAISSEMENT',    vTobSource.GetValue('E_ENCAISSEMENT')  ) ;
  lTobEche.PutValue('E_CODEACCEPT',      MPTOACC ( vTobSource.GetValue('E_MODEPAIE') ) ) ;

  // TVA
  lTobEche.PutValue( 'E_ECHEENC1',       vTobSource.GetValue('E_ECHEENC1') ) ;
  lTobEche.PutValue( 'E_ECHEENC2',       vTobSource.GetValue('E_ECHEENC2') ) ;
  lTobEche.PutValue( 'E_ECHEENC3',       vTobSource.GetValue('E_ECHEENC3') ) ;
  lTobEche.PutValue( 'E_ECHEENC4',       vTobSource.GetValue('E_ECHEENC4') ) ;
  lTobEche.PutValue( 'E_ECHEDEBIT',      vTobSource.GetValue('E_ECHEDEBIT') ) ;
  lTobEche.PutValue( 'E_EMETTEURTVA',    vTobSource.GetValue('E_EMETTEURTVA') ) ;

  // ??
  lTobEche.PutValue('E_NUMTRAITECHQ',    vTobSource.GetValue('E_NUMTRAITECHQ')  ) ;

  // Gestion Cut-Off
  if vTobSource.GetNumChamp('PCOMPL') > 0 then
    begin
    if lTobEche.GetNumChamp('PCOMPL') <= 0 then
      lTobEche.AddChampSup('PCOMPL', False ) ;
    lTobEche.PutValue( 'PCOMPL', lTobEche.GetValue('PCOMPL') ) ;
    end ;

end;

procedure TMultiEche.Renumerote( vInFrom : integer ) ;
var lTobEche : TOB ;
    i        : Integer ;
begin

  if vInFrom <= 0 then
    vInFrom := 1 ;

  for i := vInFrom to NbEche do
    begin
    lTobEche := DetailEche[ GetEcrIndex + i - 1 ] ;
    lTobEche.PutValue('E_NUMLIGNE',        NumLigne   ) ;
    lTobEche.PutValue('E_NUMECHE',         i   ) ;
    end ;

end;

procedure TMultiEche.ReporteEche(vChamps: string; vValeur: Variant);
var i : integer ;
    j : integer ;
begin

  // Montant a proratiser si modif en mode déporté
  if ( vChamps = 'E_CREDITDEV' ) or ( vChamps = 'E_DEBITDEV' ) then
    begin
    if (Piece.ModeEche = meDeporte)
      then ProratiseEche( vValeur ) ;
    end
  // sinon, Report de valeur
  else
    begin
    for i := 1 to _InMaxChampsSynchEche do
      if _recChampsSynchEche[i] = vChamps then
        begin
        // Réaffectation des valeurs
        for j := GetEcrIndex to (NbEche-1) do
          begin
          if DetailEche[ j ].GetInteger('E_NUMLIGNE')<>GetNumLigne then Exit ;
          DetailEche[ j ].PutValue( vChamps, vValeur ) ;
          end ;
        Break ;
        end ;
    end ;

end;

procedure TMultiEche.SetLignePiece(vTobEcr: Tob);
begin

  if FModeFonc <> mfPiece then Exit ;
  if not Assigned( vTobEcr ) then Exit ;
  if vTobEcr.GetInteger('E_NUMECHE')<>1 then Exit ;

  if Piece.Action = taConsult
    then FAction := taConsult
  else if vTobEcr.GetString('E_MODEPAIE')=''
    then FAction := taCreat
  else if vTobEcr<>FTobEcr then
    FAction := taModif ;

  FTobEcr       := vTobEcr ;

  FStModeInit   := GetModeInit ;
  FStModeFinal  := '' ;
  if vTobEcr.GetNumChamp('MODEFINAL')>0 then
    FStModeFinal := vTobEcr.GetString('MODEFINAL') ;

  if vTobEcr.GetString('E_MODEPAIE')=''
    then FMontantSaisi := vTobEcr.GetDouble('E_DEBITDEV') + vTobEcr.GetDouble('E_CREDITDEV')
    else FMontantSaisi := GetTotalEche ;

  FBoModif      := False ;

end;

procedure TMultiEche.SetModeAlone;
begin
  FModeFonc := mfAlone ;
end;

procedure TMultiEche.SetModePiece;
begin
  FModeFonc := mfPiece ;
end;

procedure TMultiEche.SetModeRegl(vModeRegl: T_MODEREGL);
var lTobEche : TOB ;
    lNumEche : Integer ;
begin

  if Piece.ModeEche <> meDeporte then Exit ;

  // Suppression des anciennes échéances
  if NbEche > 0 then
    ClearDetail ;

  // Mode final
  ModeFinal := vModeRegl.ModeFinal ;

  // Parcours des écheances
  for lNumEche := 1 to vModeRegl.NbEche do
    begin
    lTobEche := GetEche( lNumEche ) ;
    if lTobEche = nil then
      lTobEche := NewEche( lNumEche ) ;

    if EstDebit then
      begin
      lTobEche.PutValue('E_DEBITDEV',     vModeRegl.TabEche[lNumEche].MontantD ) ;
      lTobEche.PutValue('E_DEBIT',        vModeRegl.TabEche[lNumEche].MontantP ) ;
      lTobEche.PutValue('E_CREDITDEV',    0 ) ;
      lTobEche.PutValue('E_CREDIT',       0 ) ;
      end
    else
      begin
      lTobEche.PutValue('E_DEBITDEV',  0 ) ;
      lTobEche.PutValue('E_DEBIT',     0 ) ;
      lTobEche.PutValue('E_CREDITDEV', vModeRegl.TabEche[lNumEche].MontantD ) ;
      lTobEche.PutValue('E_CREDIT',    vModeRegl.TabEche[lNumEche].MontantP ) ;
      end ;


    lTobEche.PutValue( 'E_MODEPAIE',        vModeRegl.TabEche[lNumEche].ModePaie      ) ;
    lTobEche.PutValue( 'E_DATEECHEANCE',    vModeRegl.TabEche[lNumEche].DateEche      ) ;
    if vModeRegl.TabEche[lNumEche].DateValeur > iDate1900
      then lTobEche.PutValue( 'E_DATEVALEUR',    vModeRegl.TabEche[lNumEche].DateValeur ) ;
    lTobEche.PutValue( 'E_ORIGINEPAIEMENT', vModeRegl.TabEche[lNumEche].DateEche      ) ;

    // Zone lettrage + relance non utilisée normalement en saisie
    lTobEche.PutValue( 'E_DATERELANCE',   vModeRegl.TabEche[lNumEche].DateRelance   ) ;
    lTobEche.PutValue( 'E_NIVEAURELANCE', vModeRegl.TabEche[lNumEche].NiveauRelance ) ;

    lTobEche.PutValue( 'E_LETTRAGE',      vModeRegl.TabEche[lNumEche].CodeLettre    ) ;
    lTobEche.PutValue( 'E_COUVERTURE',    vModeRegl.TabEche[lNumEche].Couverture    ) ;
    lTobEche.PutValue( 'E_COUVERTUREDEV', vModeRegl.TabEche[lNumEche].CouvertureDev ) ;
    lTobEche.PutValue( 'E_LETTRAGEDEV',   vModeRegl.TabEche[lNumEche].LettrageDev   ) ;
    lTobEche.PutValue( 'E_DATEPAQUETMAX', vModeRegl.TabEche[lNumEche].DatePaquetMax ) ;
    lTobEche.PutValue( 'E_DATEPAQUETMIN', vModeRegl.TabEche[lNumEche].DatePaquetMin ) ;
    lTobEche.PutValue( 'E_ETATLETTRAGE',  vModeRegl.TabEche[lNumEche].EtatLettrage  ) ;

    // Gestion de la tva
    if Piece.EstTvaLoc then
      begin
      lTobEche.PutValue( 'E_ECHEENC1' ,   vModeRegl.TabEche[lNumEche].TAV[1] ) ;
      lTobEche.PutValue( 'E_ECHEENC2' ,   vModeRegl.TabEche[lNumEche].TAV[2] ) ;
      lTobEche.PutValue( 'E_ECHEENC3' ,   vModeRegl.TabEche[lNumEche].TAV[3] ) ;
      lTobEche.PutValue( 'E_ECHEENC4' ,   vModeRegl.TabEche[lNumEche].TAV[4] ) ;
      lTobEche.PutValue( 'E_ECHEDEBIT',   vModeRegl.TabEche[lNumEche].TAV[5] ) ;
      end ;
    end ;
end;

procedure TMultiEche.SetModifTVA(vBoModif: Boolean);
var lStVal : string ;
begin
  if vBoModif
    then lStVal := 'X'
    else lStVal := '-' ;

  if TOBEcr.GetNumChamp('MODIFTVA') <= 0
    then TOBEcr.AddChampSupValeur('MODIFTVA', lStVal)
    else TOBEcr.putValue( 'MODIFTVA', lStVal ) ;

end;

procedure TMultiEche.SetMontant(vNumEche: integer; vMontant: double);
var lTobEche : Tob ;
begin
  lTobEche := GetEche( vNumEche ) ;
  if lTobEche = nil then Exit ;

  if EstDebit
    then CSetMontants( lTobEche, vMontant, 0, Devise, True )
    else CSetMontants( lTobEche, 0, vMontant, Devise, True ) ;

end;

procedure TMultiEche.SetMontantSaisie(vMontant: double);
begin
  if FModeFonc <> mfAlone then Exit ;
  FMontantSaisi := vMontant ;
end;

procedure TMultiEche.SetTvaEnc(vTvaEnc1, vTvaEnc2, vTvaEnc3, vTvaEnc4,  vTvaDebit: Double ; lCoefTiers : Double ) ;
var lTobEche : TOB;
    lNumEche : Integer ;
    lTotal   : Double ;
    lCoef     : Double ;
    lMontant  : Double ;
begin

  if ModifTva then Exit ;

  if FMontantSaisi <> 0
    then lTotal := FMontantSaisi
    else lTotal := GetTotalEche( False ) ;
  if lTotal = 0 then Exit ;

  for lNumEche := 1 to NbEche do
    begin

    lTobEche  := GetEche( lNumEche ) ;

    // Calcul du coefficient de répartition sur l'échéance courante
    lCoef := lCoefTiers * ( (lTobEche.GetDouble('E_DEBIT') + lTobEche.GetDouble('E_CREDIT')) / lTotal ) ;

    // ECHE ENC 1
    lMontant := Arrondi( vTvaEnc1 * lCoef , V_PGI.OkDecV ) ;
    lTobEche.PutValue('E_ECHEENC1' ,     lMontant ) ;
    // ECHE ENC 2
    lMontant := Arrondi( vTvaEnc2 * lCoef , V_PGI.OkDecV ) ;
    lTobEche.PutValue('E_ECHEENC2' ,     lMontant ) ;
    // ECHE ENC 3
    lMontant := Arrondi( vTvaEnc3 * lCoef , V_PGI.OkDecV ) ;
    lTobEche.PutValue('E_ECHEENC3' ,     lMontant ) ;
    // ECHE ENC 4
    lMontant := Arrondi( vTvaEnc4 * lCoef , V_PGI.OkDecV ) ;
    lTobEche.PutValue('E_ECHEENC4' ,     lMontant ) ;
    // ECHE DEBIT
    lMontant := Arrondi( vTvaDebit * lCoef , V_PGI.OkDecV ) ;
    lTobEche.PutValue('E_ECHEDEBIT' ,     lMontant ) ;

    lTobEche.PutValue('E_EMETTEURTVA',   'X' ) ;

    end ;

  // MAJ de l'écriture de référence
  if Piece.ModeEche = meDeporte then
    begin
    TobEcr.PutValue( 'E_EMETTEURTVA',   'X' ) ;
    CumuleTobEcr ;
    end ;

end;

procedure TMultiEche.SynchroniseEcheances;
var lTobEche : TOB ;
    lTobTemp : TOB ;
    i        : Integer ;
begin
  if Piece.ModeEche <> meDeporte then Exit ;

  lTobTemp := TOB.Create('ECRITURE', nil, -1)  ;
  lTobTemp.InitValeurs ;

  for i := 1 to GetNbEche do
    begin
    lTobEche := GetEche( i ) ;

    // Mise en zone tampon des infos de l'échéance
    lTobTemp.Dupliquer ( lTobEche, False, True ) ;

    // Maj de l'échéance avec les données de l'écriture "globale"
    lTobEche.Dupliquer ( TobEcr, False, True ) ;

    // Récupération des infos spécifique à l'échéance :
    lTobEche.PutValue('E_NUMECHE',         i    ) ;
    PutTobEche( i, lTobTemp ) ;

    end ;

  FreeAndNil( lTobTemp ) ;

end;

procedure TMultiEche.SynchroniseEcr;
begin
  if Piece.ModeEche <> meDeporte then Exit ;
  if NbEche < 1 then Exit ;

  TobEcr.PutValue( 'E_MODEPAIE',        Detail[0].GetValue('E_MODEPAIE') ) ;
  TobEcr.PutValue( 'E_DATEECHEANCE',    Detail[0].GetValue('E_DATEECHEANCE') ) ;
  TobEcr.PutValue( 'E_ORIGINEPAIEMENT', Detail[0].GetValue('E_ORIGINEPAIEMENT') ) ;
  TobEcr.PutValue( 'E_CODEACCEPT',      Detail[0].GetValue('E_CODEACCEPT') ) ;
end;

procedure TMultiEche.UpdateEche(vNumEche: Integer; vModePaie: String; vMontant: Double; vDateEche: TDateTime);
var lTobEche : TOB ;
begin

  lTobEche := GetEche( vNumEche ) ;
  if lTobEche = nil then
    lTobEche := NewEche( vNumEche ) ;

  // MAJ mode de paiement
  lTobEche.PutValue( 'E_MODEPAIE',        vModePaie      ) ;
  lTobEche.PutValue( 'E_CODEACCEPT',      MPTOACC ( vModePaie ) ) ;

  // MAJ montant
  if EstDebit
    then CSetMontants( lTobEche,   vMontant,    0,           Devise,   True )
    else CSetMontants( lTobEche,   0,           vMontant,    Devise,   True ) ;

  // MAJ date d'échéance
  lTobEche.PutValue( 'E_DATEECHEANCE',    vDateEche  ) ;
  lTobEche.PutValue( 'E_ORIGINEPAIEMENT', vDateEche  ) ;

  // MAJ de l'écriture
  if Piece.ModeEche = meDeporte then
    SynchroniseEcr ;

  if lTobEche.IsOneModifie then
    FBoModif := True ;

end;

end.
