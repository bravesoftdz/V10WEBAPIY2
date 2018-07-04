unit UTomAffaire;

interface

uses StdCtrls,
  Controls,
  Classes,
  Vierge,
  Messages,                                                        
{$IFDEF EAGLCLIENT}
  eFiche,
  MaineAGL,
  utileAGL,
  eFichlist,
  Spin,
{$ELSE}
  db,
  FE_Main,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  DbCtrls,
  HDB,
  Fiche,
  FichList,
  AglIsoflex,
  dbgrids,
{$ENDIF}

{$IFDEF BTP}
  Etudes,
  EtudesUtil,
  CalcOLEGenericBTP,
  BTPUtil,
{$ENDIF}

  UtilConfid,
  forms,
  sysutils,
  ComCtrls,
  ParamSoc,
  Utob,
  M3FP,
  Menus,
  graphics,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOM,
  Dialogs,
  AffaireUtil,
  AffaireRegroupeUtil,
  SaisUtil,
  AglInit,
  HSysMenu,
  UtilPaieAffaire,
  Facture,
  EntGC,
  FactUtil,
  FactCalc,
  FactComm,
  Dicobtp,
  Ent1,
  HTB97,
  UtilPGI,
  Windows,
  lookup,
  UtilGC,
  TiersUtil,
  TraducAffaire,
  UtofAffaire_Reg, // pour Ok script appel fiche Affaire_REg
  AffEcheanceUtil,
  ConfidentAffaire,
  UtofAfChoixEcheAffair,
  UtofAfChoixEditAffair,
  Utilressource,

{$IFDEF GIGI}
  UtofApprecon_Mul,
{$ENDIF}

  UtofAfTableauBord,
  UtofAfBasePiece_Mul,
  //ONYX
  UAFO_REVPRIXCALCULCOEF,
  utilRevision,
  utofafrevision,
  uTomAfRevParamFormule,
  uTofAfRevSelFormule,
  utofAfRevTraitFormule,

  AGLInitGC,

  FactTOB,
  FactAdresse,uEntCommun,UtilTOBPiece,
  {$IFNDEF UTILS} BTSpigao, {$ENDIF} Edisys_IULM_Core_TLB, StrUtils;


type
  T_TypeimpactTotHT = (TitTypeGener, TitTotLignes, TitTotHT, TitEcheances, TitTypePrevu, TitRecalcul) ;
  // ******* TOM Détail des Echéances de l'affaire *************
type
  TOM_FactAff = class (TOM)
    procedure OnNewRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField (F: TField) ; override;
    procedure OnArgument (stArgument: string) ; override;
  private
  	typeCalcul : string;
    zinterval : integer;
    StatutAffaire	: String;
    DateEche			: String;
    StGenerAuto		: String;
    AncStGenerAuto: String;
    EcheFact			: String;
    LibDevise			: String;
    Tiers					: String;
    RepriseActiv	: String;
    Profil				: String;
    choix					: string;
    MethodeCalcul : string;
    TermeEChe     : string;
    DateFinFac		: TDateTime;
    DateDebFac		: Tdatetime;
    EnSaisie			: Boolean;
    SaisieContre	: Boolean;
    NumecheSto		: Integer;
    DEV						: RDEVISE;
    libEche       : string;

    procedure CacheZones (GenerAuto: string; var zlib: string) ;
    procedure AASetControlEnabled;
    procedure DefiniLibEche;
  end;
(*
  // *********** TOM liste des tiers sur l'affaire *************
type
  TOM_AffTiers = class (TOM)
    procedure OnNewRecord; override;
    procedure OnLoadRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnChangeField (F: TField) ; override;
    procedure OnArgument (stArgument: string) ; override;
  private
    tsArgsReception: TStringList;
    CodeEnc				: string;
  end;
*)
  // ******************** TOM Affaire  *************************
type
  //
  TModeTrait = (TmtAffaire,TmtDocument);
  TTypegestion = (TtgStd,TTgCotrait,TTgSousTrait);
  //
  TOM_Affaire = class (TOM)
    private
    //
    fIdCliSPIGAo : integer;
    fModeTrait  : TmodeTrait;
    fAffaire    : String;
    fModeGestion: TTypeGestion;
    //
    AffaireToDuplic : String;
    DuplicFromMenu  : Boolean;
    Domaine         : THValComboBox;
    //
    BDUPLICATION    : TToolbarButton97;
    BDELETE         : TToolbarButton97;
    //
    BFeuVert        : TToolbarButton97;
    BFeuOrange      : TToolbarButton97;
    BFeuRouge       : TToolbarButton97;
    //
    Risque          : THEdit;
    EtatRisque      : String;
    // Modified by f.vautrain 28/09/2017 16:23:06
    TabSheetDataGestion : TTabSheet;
    //
    CoefFG              : THNumEdit;
    CoefFS              : THNumEdit;
    CoefSAV             : THNumEdit;
    CoefFD              : THNumEdit;
    CoefMARG            : THNumEdit;
    TauxRG              : THNumEdit;
    DateRG              : THDBEdit;
    MtAvance            : THNumEdit;
    DebRestAvance       : THNumEdit;
    FinRestAvance       : THNumEdit;
    //
    MtCaution0          : THNumEdit;
    MtCaution1          : THNumEdit;
    MtCaution2          : THNumEdit;
    MtCaution3          : THNumEdit;
    MtCaution4          : THNumEdit;
    MtCaution5          : THNumEdit;
    MtCaution6          : THNumEdit;
    MtCaution7          : THNumEdit;
    MtCaution8          : THNumEdit;
    MtCaution9          : THNumEdit;
    //
    NumCaution0         : THNumEdit;
    NumCaution1         : THNumEdit;
    NumCaution2         : THNumEdit;
    NumCaution3         : THNumEdit;
    NumCaution4         : THNumEdit;
    NumCaution5         : THNumEdit;
    NumCaution6         : THNumEdit;
    NumCaution7         : THNumEdit;
    NumCaution8         : THNumEdit;
    NumCaution9         : THNumEdit;
    //
    Secteur             : THEdit;
    //
    procedure PositionneEtabUser(NomChamp: string);
    procedure ChangeTiersPiece(TOBPiece: TOB; CodeTiers: string);
    procedure DefiniMenuCotraitance;
    procedure SetActionsCotraitance;
    procedure CoTraitSetInterv (Sender : Tobject);
    procedure CoTraitFrsgestion(Sender : Tobject);
    procedure CoTraitDepenses(Sender : Tobject);
    procedure CotraitRecap(Sender : Tobject);
    function  CoTraitControlAffectation : boolean;
    procedure GestionCheckMandataire;
    procedure EncodeMandataire;
    procedure DecodeMandataire;
    procedure GestionEventCotrait(Active: boolean);
    procedure GereAffichagePeriodicite;
    procedure OnArgumentContrat;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure SuppActionPreventive;
    procedure NewModeleClick (Sender : Tobject);
    procedure ModifModeleClick (Sender : Tobject);
    procedure SuppModeleClick (Sender : Tobject);
    procedure GenereWordClick (Sender : Tobject);
    procedure OuvrirWordClick (Sender : Tobject);
    procedure SelectModele (Sender : TObject);
    procedure SelectDocWord (Sender : TObject);

		procedure ConstitueTOB(OneTOB: TOB; var ListeChamps : string; CodeContrat: string; WithData : boolean=true);
    procedure AjouteChampsClientFac(OneTOB: TOB; ListSqlAFF: string; WithData: Boolean);
    procedure AjouteAdressesInt(OneTOB: TOB; var ListNomAdrInt,ListSqlAFF: string;WithData: Boolean);
    procedure AjouteChampsContrat(OneTOB: TOB; var ListNomAFF,ListSqlAFF: string; WithData: Boolean);
    //
    procedure Duplication (CleDuplic: string);
    procedure ReInitDuplication(TobForm: TOB);
    procedure tobToEcran(TobForm: TOB);
    procedure Duplication_OnClick(Sender: TObject);
    procedure MiseAjourMotifAnnul(Affaire: string; DateResil: Tdatetime;preventives: Boolean=true);
    //
    procedure LoadRecordAFF(var tmp: string);
    procedure LoadRecordPRO(var tmp: string);
    procedure VoirEncours(Sender: TObject);
    procedure GestionFeuEncour;
    procedure ExitLigCaution(Sender: Tobject);
    function GetLibResponsable(CodeResssource: string): string;
    //
    public
    //
  	MnBordereaux 	  : TmenuItem;
    MnAppel         : TMenuItem;
    MnConso         : TMenuItem;
    MnImportDIE     : TMenuItem;
    MntelechargerDIE: TMenuItem;
    MnChiffrage     : TMenuItem;
    MnVoirAffSpigao : TMenuItem;
    MnCodeBarre     : TmenuItem;
    MnActions       : TmenuItem;
    MnExportDIE     : TMenuItem;
    MnPrevFacture   : TMenuItem;
    BtLigne				  : Ttoolbarbutton97;
    BtTache 			  : Ttoolbarbutton97;
    //
    TResponsable    : THDBEdit;
    //
    ThTypePaie      : THvalComboBox;
    // Evenements de la TOM
    procedure BRechResponsable(Sender: TObject);
    //
    procedure OnChangeField (F: TField) ; override;
    procedure OnNewRecord; override;
    procedure OnLoadRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;
    procedure OnArgument (stArgument: string) ; override;
    procedure SetArguments (StSQL: string) ;
    // Spécif CEGID
    procedure OnArgumentCEGID;
    procedure OnUpdateCEGID;
    procedure OnNewRecordCEGID;
    procedure AfterFormShow; overload;

    // Gestion des écheances
    function CreerEcheance (ForceCalcul, Parle, Boutton: string) : integer;
    procedure LanceEcheance (ForceCalcul, Parle, Boutton: string) ;
    function LanceChoixRecalcul (bPourcentage: Boolean) : string;
    function TestTypeModifEcheances (var DateDebut, DateFin, DateLiquid, DateResil: TDateTime; var TypeCalcul: string; var zmont: double; var TypeModif: T_TypeModifAff; Boutton: string; bSilence: boolean) : integer;
    // Lien piece Affaire
    procedure AppelPieceAffaire (TypeAppel: string) ;
    //divers outils
    procedure Egal2Champs (CS, CD: string) ;
    // Gestion du contact
    procedure NomContactAff;
    // Evenements
    procedure CAFF_AFFAIRE1Exit (sender: Tobject) ;
    procedure CAFF_AFFAIRE2Exit (sender: Tobject) ;
    procedure PartiesAffaireExit (Sender: tObject) ;
    procedure GAffTiersDblClick (Sender: TObject) ;
    procedure BTNSOUSTRAITClick (Sender : TObject);
    //Spécifique Line S1
//  procedure OnArgumentLINE;
//  Procedure AppelClient(Sender: TObject);
//  Procedure VerifieAffaireLine;

  private
    FromSPIGAONew : Boolean;
  	nbMoisEngagement : integer;

    StAffaire   : THEdit;
    STTiers     : THDBEdit;
    PlusGenere  : String;
    //
    BTChantier  : TToolbarButton97;
    //
    TermeEche_Sav       : string;
    DateDebGener_sav		: TDateTime;
    DateFinGener_sav		: TDateTime;
    DateLiquidGener_sav	: TDateTime;
    DateFin_sav					: TDateTime;
    DateResil_sav				: TDateTime;
    DateResil_old				: TdateTime;
    //
    AFTypeAction				: TActionFiche;
    //
    MethEcheance    		: string;
    Periodicite     		: string;
    TypeGenerAuto   		: string;
    StatutAffaire  		 	: string;
    EtatAffaire     		: string;
    OldTiers       			: string;
    LibAff          		: string;
    Modele          		: string;
    DerniereCreate  		: string;
    ActiviteReprise 		: string;
    NewTiers  					: string;
    NatureAuxi					: string;
    MultiEche 					: string;
    StProfil						: string;
    EtatAffairePourech	: string; //gm 15/11/02
    fStDateFin					: string;
    StSQL								: string;
    //
    Interval						: Integer;
    NbPiece							: integer;
    MoisCloture					: integer;
    //
    MontantEch					: Double;
    PourcentEch					: Double;
    CalcEcheance        : String;
    //
    EnSaisie            : Boolean;
    PieceChargee        : Boolean;
    SaisieContre        : Boolean;
    EchContre           : Boolean;
    PieceContre					: Boolean;
    ModifLot						: boolean;
    //
    TobPiece						: Tob;
    TobAdresses					: Tob;
    TobSSAff						: Tob;
    TOBAffTiers					: Tob;
    //
    CleDocAffaire				: R_CLEDOC;
    DEV									: RDEVISE;
    TypeAffaire					: T_TypeAffaire;
    GAffTiers						: THGRID;
    OldFacture          : string;
    theAuxiliaire       : string;
    // Modif BTP
    TOBEtude						: TOB;
    Spigao              : Boolean;
    // --
    fResiliation: Boolean;
    //
    // Lien piece Affaire
    procedure IncidenceTiersSurAffaire (CodeTiers: string) ;
    procedure IncidenceLignesSurAffaire (HTBefore, TaxeBefore, PRBefore: Double) ;
    procedure LanceEntetePieceAffaire (CodeTiers, CodeAffaire: string; CleDocAffaire: R_CLEDOC) ;
    procedure IncidenceEntetePieceSurAffaire (Tobpiece: TOB) ;
    procedure LoadTobPieceAffaire (CleDocAffaire: R_CLEDOC; AvecDetail: Boolean) ;
    procedure ChangePieceContre (Libere, bContre: Boolean) ;
    // Codification des affaires
    procedure MAJZonesCodeAffaire;
    // Devise + Euro
    procedure ImpactChangeDevise;
    procedure ConvertirAffaireDevise;
    procedure AfficheSigleEuro;
    function GetContre: Boolean;
    procedure ConvertirAffaireContre;
    procedure EgaliseAffaireSaisie;

    // Interface champs DB / affichage
    procedure ChampsEnabledOnInsert (OnUpdate: Boolean) ;
    procedure GereTotDevise (Entree: boolean) ;
    // Calcul sur les dates
    procedure CalculDatesMission;
    // Calcul sur le montant des écheances
    procedure CalculMontantGlobal (bSilence, bPourcentage: boolean) ;
    procedure CalculEchSurDoc(bSilence, bPourcentage: boolean) ;
    // Calcul du montant global de l'affaire
    procedure ImpactCalculTotalHT (TypeImpact: T_TypeimpactTotHT) ;
    // Calcul du montant global de l'affaire
    procedure CalculMontantEch;
    // Divers outils ...
    procedure SetCleAffaireEnabled (bEnabled: Boolean) ;
    // Fonction publiées pour les scripts
    procedure LanceAffTiers (Code: string) ;
    procedure LanceRegroupeAffaire;
    // Gestion des affaires complètes / light
    procedure AFF_AFFCOMPLETEOnClick (Sender: TObject) ;
    procedure ImpactAffCompleteOuLight;
    procedure OnUpdateAffaireLight;
    // Gestion des affaires de référence
    procedure OnUpdateAffaireRef;
    procedure ImpactAffRefouIndependante (bRecharge: Boolean) ;
    function TiersIdemAffRef: Boolean;
    // Alim du libellé d'état de l'affaire
    procedure AlimLibEtat;
    // Impression de l'affaire (fiche ou choix du modèle)
    procedure ImprimeModele (Statut: string; F: TForm) ;
    procedure TraiteAffTermine;
    // Gestion de lien avec isoflex
    procedure GereIsoflex;
    // Maj prospect / Client
    function ProspectToClient: Boolean;
    procedure ChargeAffTiers(CodeAffaire : String);
    procedure BTLigneClick (Sender: TObject) ;
    procedure BTTacheClick (Sender: TObject) ;
    procedure BordereauxClick(Sender: TObject);
    procedure MnAppelClick(Sender: TObject);
    procedure MnConsoClick(Sender: TObject);
    procedure MnImportDIEClick(Sender: TObject);
    procedure MnChiffrageClick(Sender: TObject);
    procedure MnExportDIEClick(Sender: TObject);
    procedure MnVoirAffSpigaoClick(Sender: TObject);
    procedure MnCodeBarreClick(Sender: TObject);
    Procedure MnPrevFactureClick(Sender: TObject);
    //FV1 : 18/01/2017 - FS#2314 - SES : Ajouter l'accès aux actions depuis la fiche Affaire
    procedure MnActionsClick(Sender: TObject);
    procedure MnTelechargerDIEClick(Sender: TObject);
    procedure VoirChantier(Sender: TOBJect);
    procedure LectureChantier(CodeChantier : String);
    procedure ChangeTiersAffaire (Sender : TObject);

    //ONYX
    procedure FormuleOnChange (SEnder: TObject) ;
    procedure FormuleOnExit (SEnder: TObject) ;
    procedure mnCalculOnClick (Sender: TObject) ;
    procedure mnFormuleOnClick (Sender: TObject) ;
    //procedure mnElementsOnClick(Sender: TObject);
    procedure mnAppliquerCoefOnClick (Sender: TObject) ;
    procedure mnDesappliquerCoefOnClick (Sender: TObject) ;
    procedure mnRetablirCoefOnClick (Sender: TObject) ;
    procedure GestionMenusRevision;
    procedure CoefAppliques;
    function BlocageNextCalcul (pStFormule: string) : Boolean;
    procedure CreerEcheanceDuBudget;
    procedure DefiniLibelle;
    procedure TiersFactureChange (Sender : Tobject);
    procedure Majtaches (DateDebut,DateDeFin : TdateTime);
    procedure CBCOTRAITANCEClick (Sender : Tobject);
    procedure RBMANDATAIREClick(Sender : Tobject);
    procedure RBCOTRAITANTEClick(Sender : Tobject);
		procedure ChangePeriodicite (Sender : TObject);
    procedure RegroupeEches (Sender : TObject);

  end;

const
  // libellés des messages de la TOM Affaire
  TexteMsgAffaire: array [1..50] of string = (
    {1}'Attention l''échéance ne sera pas facturée',
    {2}'Attention l''échéance  sera refacturée',
    {3}'Les Echéances non datées ne seront pas générées',
    {4}'Intervenant obligatoire',
    {5}'Type d''intervention obligatoire',
    {6}'Code Affaire Invalide',
    {7}'Code client non valide',
    {8}'Dates début et fin incohérentes',
    {9}'Vérifez vos dates de génération',
    {10}'Attention Total HT de l''affaire à zéro',
    {11}'Code client non renseigné',
    {12}'Suppression impossible Cette affaire est utilisée sur des lignes de consommations',
    {13}'Suppression impossible Cette affaire est utilisée sur des lignes de pièces',
    {14}'Suppression impossible Cette affaire est utilisée sur des pièces',
    {15}'Suppression impossible Cette affaire est utilisée par un planning.',
    {16}'Affaire modèle en mode création uniquement',
    {17}'Création d''affaire impossible sur un client fermé',
    {18}'saisi n''existe pas (table libre)',
    {19}'Calcul impossible, la date de début ou de fin d''affaire ne sont pas renseignées',
    {20}'Le statut d''affaire n''est pas valide',
    {21}'Le responsable de l''affaire est Incorrect',
    {22}'L''apporteur d''affaire est incorrect',
    {23}'Création de la pièce associée impossible',
    {24}'Facturation regroupée sur l''affaire de référence impossible car le tiers est différent',
    {25}'Ce client a un état comptable rouge, la création d''affaires est impossible',
    {26}'Attention : Ce client a un état comptable orange',
    {27}'Attention : Ce client a un état comptable rouge',
    {28}'Type de génération incorrect en GI',
    {29}'Libellé réduit non renseigné',
    {30}'Le passage du prospect en client a échoué',
    {31}'Plus d''une pièce associée à l''affaire',
    {32}'Date de résiliation inférieure à la date de début',
    {33}'La saisie du champs suivant est obligatoire : ',
    {34}'Voulez vous aligner les dates de fin de tâches sur la nouvelle date de fin d''affaire ?',
    {35}'Code inexistant : ressource libre',
    {36}'Ce client est fermé',
    //ONYX
    {37}'Confirmer-vous le calcul des coefficients pour la révision du %s ?',
    {38}'Fin du calcul des coefficients.',
    {39}'Il n''y a pas de formules dans cette affaire.',
    {40}'Cette formule n''existe pas.',
    {41}'Erreur lors du calcul des coefficients',
    {42}'On ne peut pas avoir une seconde formule de révision sans avoir de première formule.',
    {43}'Toutes les formules de cette affaire ne sont pas paramètrées.',
    {44}'Aucune facture n''a été effectuée pour le coefficient précédemment appliqué du %s. Vous ne pouvez donc pas calculer le prochain coefficient.',
    {45}'La détermination du montant de l''échéance est incorrect',
    {46}'Veuillez renseigner le mode de paiement du cotraitant',
    {47}'Veuillez indiquer si vous êtes mandataire ou cotraitant',
    {48}'Veuillez renseigner le nombre de crédits',
    {49}'Le nombre de crédit ne peut-être inférieur à la quantité déjà consommé',
    {50}'Cette date n''est pas comprise dans la période de facturation'
    ) ;

procedure AFLanceFiche_Affaire (lequel, argument: string) ;
procedure AFLanceFiche_LienAffaireTiers (Range, lequel, argument: string) ;

implementation
uses
  UtofAfRegroupeAffaire
  , Facttiers
  , PiecesRecalculs
  , DateUtils
  , UCotraitance
  , HrichOle
  , UtilWord
  , ShellApi
  , UPrintScreen
  ;

// ******* TOM Détail des Echéances de l'affaire *************

procedure TOM_FactAff.OnNewRecord;
var
  QQ: TQuery;
  IMax: integer;
begin
  inherited;
  EnSaisie := False;
  QQ := OpenSQL ('SELECT MAX(AFA_NUMECHE) FROM FACTAFF WHERE AFA_AFFAIRE="' + GetField ('AFA_AFFAIRE') + '" AND AFA_TYPECHE="NOR"', TRUE) ;
  if not QQ.EOF then
    imax := QQ.Fields [0] .AsInteger + 1
  else
    iMax := 1;
  Ferme (QQ) ;
  SetField ('AFA_NUMECHE', IMax) ;
  SetField ('AFA_NUMECHEBIS', IMax) ; //mcd 15/05/03
  SetField ('AFA_DEVISE', DEV.Code) ;
  SetField ('AFA_TIERS', Tiers) ;
  SetField ('AFA_TYPECHE', 'NOR') ;
  SetField ('AFA_REPRISEACTIV', RepriseActiv) ;
  SetField ('AFA_PROFILGENER', Profil) ; //mcd 03/07/03
  SetField ('AFA_GENERAUTO', StGenerAuto) ; //mcd 03/07/03
  SetControlText ('LIBTIERS', '') ;
  
  EcheFact := '';
  DateEche := '';
  NumecheSto := Imax;
end;

procedure TOM_FactAff.OnArgument (stArgument: string) ;
var
  Critere, Champ, valeur, Lib: string;
  x: integer;
  DateDevise: TDateTime;
begin
  inherited;

  DateDevise    := idate1900;
  LibEche := VH_GC.AFFLibfactAff;
  Critere       := (Trim (ReadTokenSt (stArgument) ) ) ;

  while (Critere <> '') do
  begin
    if Critere <> '' then
    begin
      X := pos (':', Critere) ;
      if x <> 0 then
      begin
        Champ := copy (Critere, 1, X - 1) ;
        Valeur := Copy (Critere, X + 1, length (Critere) - X) ;
      end;
      if Champ = 'AFF_GENERAUTO' then
        StGenerAuto := Valeur
      else
        if Champ = 'AFF_DEVISE' then
        DEV.Code := Valeur
      else
        if Champ = 'AFF_SAISIECONTRE' then
        SaisieContre := (Valeur = 'X')
      else
        if Champ = 'AFF_TIERS' then
        begin
        	Tiers := Valeur;
        end
      else
        if Champ = 'DATEDEBFAC' then
        DateDebFac := StrToDate (Valeur)
      else
        if Champ = 'DATEFINFAC' then
        DateFinFac := StrToDate (Valeur)
      else
        if Champ = 'AFF_REPRISEACTIV' then
        RepriseActiv := Valeur
      else
        if Champ = 'AFF_PROFILGENER' then
        Profil := Valeur
      else
        if Champ = 'CHOIX' then
        Choix := Valeur
      else
        if Champ = 'AFF_DATECREATION' then
        DateDevise := strToDate (Valeur)
      else
      	if champ = 'AFF_PERIODICITE' then typeCalcul := valeur
      else
      	if champ = 'AFF_INTERVALGENER' then zinterval := strtoint(valeur)
      else
        if champ = 'AFF_STATUTAFFAIRE' then StatutAffaire := Valeur
      else
        if champ = 'AFF_METHECHEANCE' then MethodeCalcul := Valeur
      else
        if champ = 'AFF_TERMEECHEANCE' then TermeEChe := Valeur;
    end;
    Critere := (Trim (ReadTokenSt (stArgument) ) ) ;
  end;
  if (DEV.Code = '') then
    DEV.code := V_PGI.DevisePivot;

  if DEV.Code = V_PGI.DevisePivot then
  begin
    if SaisieContre then
    begin
      if VH^.TenueEuro then
      begin
        DEV.Libelle := RechDom ('TTDEVISETOUTES', V_PGI.DeviseFongible, False) ;
        DEV.Decimale := V_PGI.OkDecE;
      end
      else
      begin
        DEV.Decimale := V_PGI.OkDecE;
        DEV.Symbole := 'E';
        DEV.Libelle := 'Euro';
      end;
    end
    else
    begin // mcd 22/03/02
      GetInfosDevise (DEV) ;
      DEV.Taux := GetTaux (DEV.Code, DEV.DateTaux, DateDevise) ;
    end;
  end
  else
  begin
    GetInfosDevise (DEV) ;
    DEV.Taux := GetTaux (DEV.Code, DEV.DateTaux, DateDevise) ;
  end;

  LibDevise := DEV.Libelle;
  SetControlText ('LIBDEVISE', LibDevise) ;
  // mcd 10/07/02
{$IFDEF BTP}
  if StatutAffaire = 'INT' then
  begin
    SetControlProperty ('AFA_GENERAUTO', 'Plus', 'GA" AND CO_CODE="CON') ;
    SetControlEnabled ('AFA_GENERAUTO', False) ;
    SetControlVisible ('AFA_REPRISEACTIV', False) ;
    SetControlVisible ('TAFA_REPRISEACTIV', False) ;
  end else
  begin
    SetControlProperty ('AFA_GENERAUTO', 'Plus', 'BTP') ;
  end;
{$ELSE}
  SetControlProperty ('AFA_GENERAUTO', 'Plus', 'GA') ;
  if ctxScot in V_PGI.PGIContexte then
    SetControlProperty ('AFA_GENERAUTO', 'Plus', 'GA" AND CO_CODE<>"CON') ;
{$ENDIF}
  if StGenerAuto = 'POU' then
  begin
    TFFicheListe (Ecran) .LibreName := 'AFA_POURCENTAGE';
    TFFicheListe (Ecran) .TitreCols := 'Num;Libelle;Pourcentage';
  end
  else
  begin
    TFFicheListe(Ecran).LibreName := 'AFA_MONTANTECHEDEV';
    TFFicheListe(Ecran).TitreCols := 'Num;Libelle;    Montant';
  end;
  CacheZones (StGenerAuto, Lib) ;
  SetControlEnabled('AFA_DATEECHE',false);
  AASetControlEnabled;
end;

procedure TOM_FactAff.AASetControlEnabled;
var facture : Boolean;
begin
  facture := (GetField ('AFA_ECHEFACT')='X');
  SetControlEnabled('AFA_DATEDEBUTFAC',not facture);
  SetControlEnabled('AFA_DATEFINFAC',not facture);
  SetControlEnabled('AFA_LIBELLEECHE',not facture);
  SetControlEnabled('AFA_MONTANTECHEDEV',not facture);
end;

{ TOM_FactAff }

procedure TOM_FactAff.DefiniLibEche;
var Posit : Integer;
    DateDebut,DateFin,DateEche : TDateTime;
    st : string;
begin
  if libEche = '' then Exit;
  st := libeche;
  Posit := Pos ('$$',st);
  DateDebut := GetField('AFA_DATEDEBUTFAC');
  DateFin := GetField('AFA_DATEFINFAC');
  DateEche := GetField('AFA_DATEECHE');

  If (Posit <> 0) then
  Begin
    delete (st,Posit,2);
    if DateFin <> IDate1900 then
    begin
      Insert(' '+DateToStr(DateDebut)+' au '+DateToStr(DateFin),st,Posit);
    end else
    begin
      Insert(DateToStr(DateEche),st,Posit);
    end;
  end;
  SetField('AFA_LIBELLEECHE',st);
end;

procedure TOM_FactAff.OnChangeField (F: TField) ;
var st        : String;
    lib       : string;
    posit     : Integer;
    PositBis  : Integer;
begin
  inherited;

  if (F.FieldName = 'AFA_DATEDEBUTFAC') then
     begin
    (*
    if (GetField('AFA_DATEDEBUTFAC') > DateFinFac) or
       (GetField('AFA_DATEDEBUTFAC') < DateDebFac) then
    begin
      LastError := 50;
      LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
    end;
    *)
    if TermeEChe = 'ACH' then SetField('AFA_DATEECHE',GetField('AFA_DATEDEBUTFAC'));
    DefiniLibEche;
  end else if (F.FieldName = 'AFA_DATEFINFAC') then
  begin
    (*
    if (GetField('AFA_DATEFINFAC') > DateFinFac) or
       (GetField('AFA_DATEFINFAC') < DateDebFac) then
    begin
      LastError := 50;
      LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
    end;
    *)
    if TermeEChe = 'ECH' then SetField('AFA_DATEECHE',GetField('AFA_DATEFINFAC'));
    DefiniLibEche;
  end else if (F.FieldName = 'AFA_ECHEFACT') then
  begin
     st := GetField ('AFA_ECHEFACT') ;
     if ((EcheFact <> GetField ('AFA_ECHEFACT') ) and (EcheFact = '-') ) then
        begin
        LastError := 1;
        LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
        end;
     if ((GetField ('AFA_ECHEFACT') = '-') and (EcheFact = 'X') ) then
        begin
        LastError := 2;
        LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
        end;
    AASetControlEnabled;
  end else if (F.FieldName = 'AFA_GENERAUTO') and (GetField ('AFA_GENERAUTO') <> AncStGenerAuto) then
       begin
       CacheZones (GetField ('AFA_GENERAUTO') , lib) ;
       if (StGenerAuto <> GetField ('AFA_GENERAUTO') ) and (GetField ('AFA_GENERAUTO') <> 'ACT') then
          PgiInfoAf ('Attention, le mode est différent de la fiche affaire', 'Affaire') ;

  end else if (F.FieldName = 'AFA_NUMECHEBIS') then //mcd 23/05/03
       begin
       St := string (GetField ('AFA_LIBELLEECHE') ) ;
       if (St <> '') then
          begin
          Posit := Pos (InttoStr (NumEcheSto) , St) ;
          if Posit = 0 then
             Posit := Pos ('**', St)
          else
             begin
             PositBis := Pos (DateEche, St) ;
             if (Posit >= PositBis) and (Posit <= PositBis + 8) then
                begin // la position trouvé est sur un des N° de la date ..
                if (lenGth (St) - PositBis - 11) > 0 then //mcd 16/10/03 pour cas plus rien après date, il en faut pas continuer
                   Posit := PositBis + 9 + Pos (InttoStr (NumEcheSto) , Copy (St, PositBis + 10, lenGth (St) - PositBis - 9) )
                else
                   posit := 0;
                end
             end;
          if (Posit <> 0) then
             begin
             Delete (St, Posit, length (IntToStr (NumEcheSto) ) ) ;
             Insert (GetField ('AFA_NumEcheBis') , st, Posit) ;
             SetField ('AFA_LIBELLEECHE', St) ;
             end;
    end else
          begin
      if (GetField ('AFA_ECHEFACT') = '-') then St := VH_GC.AFFLibfactAff
                                           else St := GetParamSoc ('SO_AFLIBLIQFACTAF') ;
          Posit := Pos ('**', St) ; // $$  pour reprendre la n°
          if (Posit <> 0) then
             begin
             delete (St, Posit, 2) ;
             Insert (GetField ('AFA_NUMECHEBIS') , St, Posit) ;
             end;
          SetField ('AFA_LIBELLEECHE', St) ;
          end;
       NumEcheSto := GetField ('AFA_NumEcheBis') ; //mcd 16/09/03 changer de place, il faut ds tous les cas stocker la nouvelle valeur

  end else if (F.FieldName = 'AFA_DATEECHE') then
       begin
    (*
       St := string (GetField ('AFA_LIBELLEECHE') ) ;
       if (St <> '') then
          begin
          Posit := Pos (DateEche, St) ;
          if (Posit <> 0) then
             begin
             Delete (St, Posit, length (DateEche) ) ;
             Insert (GetField ('AFA_DATEECHE') , st, Posit) ;
             SetField ('AFA_LIBELLEECHE', St) ;
             DateEche := GetField ('AFA_DATEECHE') ;
             end;
    end else
          begin
      if (GetField ('AFA_ECHEFACT') = '-') then St := VH_GC.AFFLibfactAff
                                           else St := GetParamSoc ('SO_AFLIBLIQFACTAF') ;
          if (string (GetField ('AFA_DATEECHE') ) <> '') then
             begin
             Posit := Pos ('$$', St) ; // $$  pour reprendre la date
             if (Posit <> 0) then
                begin
                delete (St, Posit, 2) ;
                Insert (GetField ('AFA_DATEECHE') , St, Posit) ;
                end;
             SetField ('AFA_LIBELLEECHE', St) ;
             end;
          end;
    *)
       end;

end;

procedure TOM_FactAff.CacheZones (GenerAuto: string; var Zlib: string) ;
begin // fct qui cahce champ % ou mtt en fct de l'enrgt en cours. appeler de OnLoad,Onargume et sur chgmt valeur

  if (GenerAuto = 'POU') then // au pourcentage
     begin
     SetControlVisible ('AFA_POURCENTAGE', TRUE) ;
     SetControlVisible ('TAFA_POURCENTAGE', TRUE) ;
     SetControlVisible ('AFA_MONTANTECHEDEV', FALSE) ;
     SetControlVisible ('LAFA_MONTANTECHEDEV', FALSE) ;
     SetControlVisible ('TAFA_MONTANTECHE', FALSE) ;
     zlib := '%';
     end
  else if (GenerAuto = 'POT') then // sur lignes d'affaires
     begin
     SetControlVisible ('AFA_POURCENTAGE', FALSE) ;
     SetControlVisible ('TAFA_POURCENTAGE', FALSE) ;
     SetControlVisible ('AFA_MONTANTECHEDEV', FALSE) ;
     SetControlVisible ('LAFA_MONTANTECHEDEV', FALSE) ;
     SetControlVisible ('TAFA_MONTANTECHE', FALSE) ;
     zlib := LibDevise;
     end
  else if (GenerAuto = 'CON') then // contrat PA le 10/08/2001
     begin
     SetControlVisible ('AFA_POURCENTAGE', FALSE) ;
     SetControlVisible ('TAFA_POURCENTAGE', FALSE) ;
     SetControlVisible ('AFA_MONTANTECHEDEV', TRUE) ;
     SetControlVisible ('LAFA_MONTANTECHEDEV', TRUE) ;
     SetControlVisible ('TAFA_MONTANTECHE', TRUE) ;
     zlib := LibDevise;
     end
  else // par défaut
     begin
     SetControlVisible ('AFA_POURCENTAGE', FALSE) ;
     SetControlVisible ('TAFA_POURCENTAGE', FALSE) ;
     SetControlVisible ('AFA_MONTANTECHEDEV', TRUE) ;
     SetControlVisible ('LAFA_MONTANTECHEDEV', TRUE) ;
     SetControlVisible ('TAFA_MONTANTECHE', TRUE) ;
     zlib := LibDevise;
     end;

  if (GenerAuto <> 'FOR') then SetControlText ('TAFA_MONTANTECHE', 'Esti&mation de l''échéance') ;

  (* mcd 01/10/03
  if VH^.TenueEuro then
     SetControlVisible ('SIGLEEURO', Not(SaisieContre))
  else
     SetControlVisible ('SIGLEEURO', SaisieContre);*)

  SetControlVisible ('LIBDEVISE', not ((GenerAuto = 'POU') or (GenerAuto = 'POT') ) ) ;
  SetControlVisible ('SIGLEEURO', False) ;

end;

procedure TOM_FactAff.OnUpdateRecord;
var DateEchePriv	: string;
    Montant				: Double;
    DateDebutD    : TDateTime;
begin
  inherited;

  DateEchePriv := GetField ('AFA_DATEECHE') ;

  if (DateEchePriv = '') then
  	 begin
     LastError := 3;
     LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
     end;

  if EnSaisie then
     begin
     	DateDebutD := DateDebfac;
    DateDebutD := GetDateDebutPeriode (TypeCalcul, DateDebutD, zinterval,MethodeCalcul)  ;

     //mcd 15/05/03 on fait un message si date échénace ne fait pas partie dates de fatcuration
     if (StrToDAte (DateEchePriv) > DateFinFac) or (StrToDATe (DateEchePriv) < DateDebutD) then
        begin
        PgiInfoAf ('La date d''échéance n''est pas comprise dans la période de facturation : début :' + DateToStr (DateDebFac) + ' Fin :' + DateToStr (DateFinFac) , traduitGa ('Affaire') ) ;
        end;
     if (DEV.Code = '') then DEV.code := V_PGI.DevisePivot;
     if (DEV.Code <> V_PGI.DevisePivot) then
        begin
        // mcd 13/02/03 if VH^.TenueEuro then BEGIN SufF:='CON' ; SufE:='' ; END else BEGIN SufF:='' ; SufE:='CON' ; END ;
        Montant := GetField ('AFA_MONTANTECHEDEV') ;
        //SetField('AFA_MONTANTECHE'+SufF,DeviseToFranc(Montant,DEV.Taux, DEV.Quotite)) ;
        //SetField('AFA_MONTANTECHE'+SufE,DeviseToEuro (Montant,DEV.Taux, DEV.Quotite)) ;
        SetField ('AFA_MONTANTECHE', DeviseToEuro (Montant, DEV.Taux, DEV.Quotite) ) ;
    end else
        begin
        if SaisieContre then
           begin
           SetField ('AFA_MONTANTECHE', ConvertSaisieEF (GetField ('AFA_MONTANTECHEDEV') , SaisieContre) ) ;
      end else
           begin
           SetField ('AFA_MONTANTECHE', GetField ('AFA_MONTANTECHEDEV') ) ;
           end;
        end;
     end;

end;

procedure TOM_FactAff.OnLoadRecord;
var
  zcum, zpou, zcumAF, zpouAF: double;
  znbr, ZnbrAf: integer;
  zlib, zmont, ZmontAF: string;
  CleDoc: R_CLEDOC;
  QQ: TQuery;
begin
  inherited;

  EnSaisie := True;
  EcheFact := GetField ('AFA_ECHEFACT') ;
  DateEche := GetField ('AFA_DATEECHE') ;
  NumEcheSto := GetField ('AFA_NUMECHEBIS') ;
  AncStGenerAuto := GetField ('AFA_GENERAUTO') ;
  ChargeCleAffaire (nil, THEDIT (GetControl ('AFF_AFFAIRE1') ) , THEDIT (GetControl ('AFF_AFFAIRE2') ) ,
    THEDIT (GetControl ('AFF_AFFAIRE3') ) , THEDIT (GetControl ('AFF_AVENANT') ) , nil, taConsult, GetField ('AFA_AFFAIRE') , False) ;
  CacheZOnes (AncStGenerAuto, Zlib) ;
  SetControlVisible ('LIQUIDATIVE', GetField ('AFA_LIQUIDATIVE') = 'X') ;
  // Calcul des cumuls : nbre d'échéances , montant ou  %
  zcum := 0;
  znbr := 0;
  zpou := 0;

  CumulEcheances (GetField ('AFA_AFFAIRE') , zcum, zpou, znbr) ;

  // PL le 06/03/02 : INDEX 3   mcd 22/03/02 ajout affichage ech à facturer
  QQ := OpenSQL ('SELECT COUNT(AFA_NUMECHE),SUM(AFA_MONTANTECHEDEV),SUM(AFA_POURCENTAGE) FROM FACTAFF WHERE AFA_TYPECHE="NOR" AND AFA_AFFAIRE="' + GetField ('AFA_AFFAIRE') + '"AND AFA_ECHEFACT="-"', TRUE) ;
  if not QQ.EOF then
     begin
     znbrAF := QQ.Fields[0].AsInteger;
     zcumAF := QQ.Fields[1].AsFloat;
     zpouAF := QQ.Fields[2].AsFloat;
  end else
     begin
     znbrAF := 0;
     zcumAF := 0;
     zpouAF := 0;
     end;
  Ferme (QQ) ;

  zmont := StrF00 (zcum, DEV.Decimale) ;
  zmontAF := StrF00 (zcumAF, DEV.Decimale) ;
  if (StGenerAuto = 'POU') then
    begin
    SetControlText ('CUMUL', Format ('%d Echéance(s) soit %8.2n %s ', [znbr, zpou, zlib] ) ) ;
    SetControlText ('CUMUL1', Format ('%d Echéance(s) non facturée(s) soit %8.2n %s ', [znbrAF, zpouAF, zlib] ) ) ;
  end else if (StGenerAuto = 'POT') then
       begin
       SetControlText ('CUMUL', TraduitGA (Format ('%d Echéance(s) sur les lignes d''affaires ', [znbr] ) ) ) ;
       SetControlText ('CUMUL1', TraduitGA (Format ('%d Echéance(s) non facturée(s) ', [ZnbrAF] ) ) ) ;
  end else
       begin
       SetControlText ('CUMUL', Format ('%d Echéance(s)  soit  %s  %s ', [znbr, zmont, zlib] ) ) ;
       SetControlText ('CUMUL1', Format ('%d Echéance(s) non facturée(s) soit  %s  %s ', [znbrAF, zmontAf, zlib] ) ) ;
       end;

  // Affectation en clair du dernier numero de piece
  if GetField ('AFA_NUMPIECE') <> '' then
  begin
    DecodeRefPiece (GetField ('AFA_NUMPIECE') , CleDoc) ;
    SetControlText ('DERNIEREPIECE',
      Format ('le %s   %s n° %d ', [DateToStr (CleDoc.DatePiece) , CleDoc.NaturePiece, CleDoc.NumeroPiece] ) ) ;
  end else
  begin
    SetControlText ('DERNIEREPIECE', '') ;
  end;
end;

{ TOM_AffTiers }
(*
procedure TOM_AffTiers.OnNewRecord;
begin
  inherited;
  Setfocuscontrol ('AFT_TYPEINTERV') ;
end;

procedure TOM_AffTiers.OnLoadRecord;
begin
  inherited;
  if (GetField ('AFT_TIERS') <> '') then
  begin
    SetControltext ('AFT_TYPEINTERV', 'TIE') ;
    CodeEnc := GetField ('AFT_TIERS') ;
  end
  else
  begin
    SetControltext ('AFT_TYPEINTERV', 'RES') ;
    CodeEnc := GetField ('AFT_RESSOURCE') ;
  end;
end;

procedure TOM_AffTiers.OnArgument (stArgument: string) ;
var
  stmp, wargument: string;
begin
  inherited;

{$IFDEF BTP}
  SetControlVisible ('AFT_PROPOINTERV', False) ;
  SetControlVisible ('TAFT_PROPOINTERV', False) ;
{$ENDIF}

  if ctxScot in V_PGI.PGIContexte then
  	 begin
   	 SetControlVisible ('AFT_PROPOINTERV', False) ;
     SetControlVisible ('TAFT_PROPOINTERV', False) ;
	   end;

  if (stArgument = '') then exit;

  sTmp := StringReplace (stArgument, ';', chr (VK_RETURN) , [rfReplaceAll] ) ;
  tsArgsReception := TStringList.Create;
  tsArgsReception.Text := sTmp;
  Wargument := tsArgsReception.Values ['AFFAIRE'] ;

  // fin MCD
  ChargeCleAffaire (nil, THEdit (GetControl ('AFT_AFFAIRE1') ) , THEdit (GetControl ('AFT_AFFAIRE2') ) , THEdit (GetControl ('AFT_AFFAIRE3') ) , THEDIT (GetControl ('AFT_AVENANT') ) , nil, taConsult, WArgument, False) ;

  UpdateCaption (Ecran) ;

end;

procedure TOM_AffTiers.OnUpdateRecord;
var
  QQ: TQuery;
  IMax: integer;
begin
  inherited;
  if (GetField ('AFT_TIERS') <> '') then
  begin
    SetField ('AFT_TYPEINTERV', 'TIE') ;
  end
  else
  begin
    SetField ('AFT_TYPEINTERV', 'RES') ;
  end;

  if (GetField ('AFT_RESSOURCE') = '') and (GetField ('AFT_TIERS') = '') then
  begin
    LastError := 4;
    SetFocusControl ('AFT_TYPEINTERV') ;
    LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
  end
  else
    if (GetField ('AFT_LIENAFFTIERS') = '') then
  begin
    LastError := 5;
    SetFocusControl ('AFT_LIENAFFTIERS') ;
    LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
  end;

  QQ := OpenSQL ('SELECT MAX(AFT_RANG) FROM AFFTIERS WHERE AFT_AFFAIRE="' + GetField ('AFT_AFFAIRE') + '"', TRUE) ;
  if not QQ.EOF then
    imax := QQ.Fields [0] .AsInteger + 1
  else
    iMax := 1;
  Ferme (QQ) ;
  SetField ('AFT_RANG', IMax) ;
end;

procedure TOM_AffTiers.OnChangeField (F: TField) ;
var
  QQ: TQuery;
  CCVisible, CCCache1, CCCache2, Prefixe, Table, Complement: string;
begin
  inherited;
  if (F.FieldName = 'AFT_TIERS') or (F.FieldName = 'AFT_RESSOURCE') then
  begin
    if (GetControlText (F.FieldName) = '') then
      Exit;
    if (CodeEnc = GetControlText (F.FieldName) ) then
      Exit
    else
      CodeEnc := GetControlText (F.FieldName) ;

    if (F.FieldName = 'AFT_TIERS') then
    begin
      Prefixe := 'T_';
      Table := 'TIERS';
      Complement := '';
    end
    else
      if (F.FieldName = 'AFT_RESSOURCE') then
    begin
      Prefixe := 'ARS_';
      Table := 'RESSOURCE';
      Complement := '';
    end;

    QQ := OpenSQL ('SELECT ' + prefixe + 'LIBELLE ' + Complement + ' FROM ' + table + ' WHERE ' + Prefixe + Table + ' ="' + GetControlText (F.FieldName) + '"', TRUE) ;
    if not QQ.EOF then
    begin
      SetField ('AFT_LIBELLE', QQ.Fields [0] .AsString) ;
    end;
    Ferme (QQ) ;
  end
  else
    if (F.FieldName = 'AFT_LIENAFFTIERS') then
  begin
    if THValCOmboBox(GetControl('AFT_TYPEINTERV')).Value = 'TIE' then // Tiers
    begin
      CCVisible := 'AFT_TIERS';
      CCCache1 := 'AFT_RESSOURCE';
      CCCache2 := 'AFT_APPORTEUR';
    end
    else // Ressource
    begin
      CCVisible := 'AFT_RESSOURCE';
      CCCache1 := 'AFT_TIERS';
      CCCache2 := 'AFT_APPORTEUR';
    end;

    SetControlVisible (CCCache1, False) ;
    SetField (CCCache1, '') ;
    SetControlVisible (CCCache2, False) ;
    SetField (CCCache2, '') ;
    SetControlVisible (CCVisible, True) ;
    if GetField (CCVisible) = '' then
      SetControlText ('AFT_LIBELLE', '') ;
  end;
end;
*)
{ TOM_Affaire }

procedure TOM_Affaire.OnChangeField (F: TField) ;
var
  St: string;
//  Memo: TMemo;
  NbMois : integer;
  DateFin : TDateTime;
  DateCal : TdateTime;
//  lng: integer;
  //posit : Integer;
begin
  inherited;

  //Modif FV pour Gestion du recalcul auto des dates
  if StatutAffaire = 'INT' then
  begin
    if (F.FieldName = 'AFF_DATEDEBUT') and (TFFiche(ecran).TypeAction = tacreat) then
  	begin
      DateFin := IncMonth(GetField('AFF_DATEDEBUT'), NbMoisEngagement);
      DateFin := PlusDate (DateFin,-1,'J');
      SetField ('AFF_DATEFIN', DateFin) ;
      SetField ('AFF_DATEFINGENER', DateFin) ;
      if (StrToDate(GetField('AFF_DATEDEBGENER')) <> GetField('AFF_DATEDEBUT'))then
      begin
        SetField('AFF_DATEDEBGENER',GetField('AFF_DATEDEBUT'));
      end;
  	end
    else IF (f.fieldName = 'AFF_DATERESIL') then
    begin
      if (GetField ('AFF_DATERESIL') <> idate1900) and (GetField ('AFF_DATERESIL') <> idate2099) Then
      begin
        if DateToStr(DateResil_Sav) <> DateToStr(GetField('AFF_DATERESIL')) then fresiliation := True;
        if ((GetField ('AFF_DATERESIL') <  GetField('AFF_DATEDEBUT')) and
            (GetField ('AFF_DATEDEBUT') <> idate1900))  or
           ((GetField ('AFF_DATERESIL') <  GetField('AFF_DATEDEBGENER')) and
            (GetField ('AFF_DATEDEBGENER') <> idate1900)) then
        begin
//          LastError := 32;
//          LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
          PGIBoxAF (TexteMsgAffaire [32] , '') ;
          //
          SetField('AFF_DATERESIL', DateToStr(DateResil_old));
          SetFocusControl ('AFF_DATERESIL') ;
          fResiliation := False;
          Exit;
        end;
        SetField('AFF_DATEFINGENER', GetField('AFF_DATERESIL'));
        SetField('AFF_DATEFIN', GetField('AFF_DATERESIL'));
      end;
    end
    else if (F.FieldName = 'AFF_DATEFIN') then
    begin
        DateFin := StrToDate(GetField('AFF_DATEFIN'));
//    	  SetField('AFF_DATEDEBGENER', GetField('AFF_DATEDEBUT'));
        if (StrToDate(GetField('AFF_DATEFINGENER')) <> DateFin) and (AFTypeAction = TaCreat) then
        begin
     	    SetField('AFF_DATEFINGENER', DateToStr(DateFin));
        end;
        //uniquement en création de fiche
//        if AFTypeAction = TaCreat then
      	   Begin
           if datefin <> idate2099 then
              begin
    	        NbMois:= StrToInt(GetParamSoc ('SO_AFALIMGARANTI'));
              DateCal := IncMonth(DateFin, NbMois);
              SetField('AFF_DATEGARANTIE', DateToStr(DateCal));
              NbMois:= StrToInt(GetParamSoc ('SO_AFALIMCLOTURE'));
              DateCal := IncMonth(DateFin, NbMois);
              SetField('AFF_DATECLOTTECH', DateToStr(DateCal));
              //Date Limite d'intervention
              NbMois:= StrToInt(GetParamSoc ('SO_AFCALCULFIN'));
              DateCal := IncMonth(DateFin, NbMois);
              SetField('AFF_DATELIMITE', DateToStr(DateCal));
           		end;
      	   end;
        end
  	 else if (F.FieldName = 'AFF_DATEFINGENER') then
        begin
          DateFin := StrToDate(GetField('AFF_DATEFINGENER'));
          (*
          if StrToDate(getField('AFF_DATEFIN')) <> DateFin then
          begin
            SetField('AFF_DATEFIN', DateToStr(DateFin));
          end;
          *)
    end;
  //
  if (F.FieldName = 'AFF_INTERVALGENER') then
     begin
     if not (ctxscot in V_PGI.PGIContexte) then
        begin
        if (GetField ('AFF_GENERAUTO') = 'CON') then
           //Modif FV problème de calcul de l'échéance (19012007)
           //SetField ('AFF_MONTANTECHEDEV', GetField ('AFF_TOTALHTDEV') * GetField ('AFF_INTERVALGENER') ) ;
           CalculEchSurDoc(true, false);
           //SetField ('AFF_MONTANTECHEDEV', GetField ('AFF_TOTALHTDEV')) ;
        end;
     end;

  if (F.FieldName = 'AFF_METHECHEANCE') then
     begin
     if not (ctxscot in V_PGI.PGIContexte) then
        begin
        if (GetField ('AFF_GENERAUTO') = 'CON') then
           //Modif FV problème de calcul de l'échéance (19012007)
           CalculEchSurDoc(true, false);
        end;
     end;

  if (F.FieldName = 'AFF_DETECHEANCE') then
     begin
     if not (ctxscot in V_PGI.PGIContexte) then
        begin
        if (GetField ('AFF_GENERAUTO') = 'CON') then
           //Modif FV problème de calcul de l'échéance (19012007)
           CalculEchSurDoc(true, false);
        end;
     end;

  if (F.FieldName = 'AFF_GENERAUTO') then
     begin
     St := GetField ('AFF_GENERAUTO') ;
     if ctxScot in V_PGI.PGIContexte then
        begin
        if (st = 'CON') then
           begin
           PGIBoxAf ('Valeur Incorrecte dans la GI', TitreHalley) ;
           SetFocusControl ('AFF_GENERAUTO') ;
           end;
        end
     else
        begin
        if (GetField ('AFF_GENERAUTO') = 'CON') then
           //Modif FV problème de calcul de l'échéance (19012007)
           //SetField ('AFF_MONTANTECHEDEV', GetField ('AFF_TOTALHTDEV') * GetField ('AFF_INTERVALGENER') ) ;
           CalculEchSurDoc(true, false);
           //SetField ('AFF_MONTANTECHEDEV', GetField ('AFF_TOTALHTDEV')) ;
        end;
        // PL le 28/11/02 : on deconnecte le multiecheance de l'alim par le budget
        // PL le 10/10/02 : ouverture pour Scot
        // PL le 20/06/02 : en alimentation des echeances par le budget, on doit toujours voir la case
        // à cocher AFF_MULTIECHE
        // Dans le cas contraire on gere suivant les cas
        if (GetParamSoc ('SO_AFMULTIECHE') = false) then
           begin
           // PL le 18/06/02 : gestion multi echeances
           SetControlVisible ('AFF_MULTIECHE', False) ;
           SetField ('AFF_MULTIECHE', '-') ;
           end
        else
           // On gère le multi echeance
           begin
           if (GetField ('AFF_GENERAUTO') <> 'ACT') and
              (GetField ('AFF_GENERAUTO') <> 'MAN') and
              (GetField ('AFF_REPRISEACTIV') <> 'NON') and
              (GetField ('AFF_REPRISEACTIV') <> '') then
              SetControlVisible ('AFF_MULTIECHE', True)
           else
              begin
              SetControlVisible ('AFF_MULTIECHE', false) ;
              SetField ('AFF_MULTIECHE', '-') ;
              end;
           end;

 //////////////////////////
 /////////////////////////
    if (st <> 'CON') then
    begin
      if ctxScot in V_PGI.PGIContexte then
        SetControlVisible ('GROUPBOXCONTRAT', False)
      else
      begin
        SetControlVisible ('GROUPBOXCONTRAT', true) ;
        SetControlVisible ('AFF_METHECHEANCE', false) ;
        SetControlVisible ('TAFF_METHECHEANCE', false) ;
      end;
      SetControlEnabled ('AFF_MONTANTECHEDEV', TRUE) ;
    end
    else
    begin
      if ctxScot in V_PGI.PGIContexte then
        SetControlVisible ('GROUPBOXCONTRAT', True)
      else
      begin
        SetControlVisible ('GROUPBOXCONTRAT', True) ;
        SetControlVisible ('AFF_METHECHEANCE', True) ;
        SetControlVisible ('TAFF_METHECHEANCE', True) ;
      end;
      SetControlEnabled ('AFF_MONTANTECHEDEV', False) ;
    end;
    if (st <> 'AVA') then
      SetControlVisible ('GROUPBOXSITUATION', False)
    else
      SetControlVisible ('GROUPBOXSITUATION', True) ;
    if ctxscot in V_PGI.PGICOntexte then
      SetControlVisible ('GROUPBOXSITUATION', True) ; // appréciation

    ImpactCalculTotalHT (TitTypeGener) ; // gestion du total en fonction des modifs de type de génération

    if (st = 'FOR') then // Au forfait
       begin
       SetControlVisible ('AFF_POURCENTAGE', FALSE) ;
       SetControlVisible ('TAFF_POURCENTAGE', FALSE) ;
       SetControlVisible ('AFF_MONTANTECHEDEV', TRUE) ;
       SetControlVisible ('TAFF_MONTANTECHE', TRUE) ;
       SetControlVisible ('BRECALCULMTGLOBAL', TRUE) ;
       SetControlText ('TAFF_MONTANTECHE', 'Montant Echéance') ;
       end
    else if (st = 'ACT') then // Sur activité
       begin
       SetControlVisible ('AFF_POURCENTAGE', FALSE) ;
       SetControlVisible ('TAFF_POURCENTAGE', FALSE) ;
       SetControlVisible ('AFF_MONTANTECHEDEV', TRUE) ;
       SetControlVisible ('TAFF_MONTANTECHE', TRUE) ;
       SetControlVisible ('BRECALCULMTGLOBAL', TRUE) ;
       SetControlText ('TAFF_MONTANTECHE', 'Montant Echéance') ;
       end
    else if (st = 'POU') then // Au pourcentage
       begin
       SetField ('AFF_POURCENTAGE', 0) ;
       SetControlVisible ('AFF_MONTANTECHEDEV', FALSE) ;
       SetControlVisible ('TAFF_MONTANTECHE', FALSE) ;
       SetControlVisible ('AFF_POURCENTAGE', TRUE) ;
       SetControlVisible ('TAFF_POURCENTAGE', TRUE) ;
       SetControlVisible ('BRECALCULMTGLOBAL', FALSE) ;
       end
    else if (st = 'POT') then // sur lignes d'affaires
       begin
       SetField ('AFF_POURCENTAGE', 100.00) ;
       SetControlVisible ('AFF_MONTANTECHEDEV', FALSE) ;
       SetControlVisible ('TAFF_MONTANTECHE', FALSE) ;
       SetControlVisible ('AFF_POURCENTAGE', FALSE) ;
       SetControlVisible ('TAFF_POURCENTAGE', FALSE) ;
       SetControlVisible ('BRECALCULMTGLOBAL', FALSE) ;
       end
    else if (st = 'CON') then // contrat
       begin
       SetControlVisible ('AFF_POURCENTAGE', FALSE) ;
       SetControlVisible ('TAFF_POURCENTAGE', FALSE) ;
       SetControlVisible ('AFF_MONTANTECHEDEV', TRUE) ;
       SetControlVisible ('TAFF_MONTANTECHE',  TRUE) ;
       SetControlVisible ('BRECALCULMTGLOBAL', FALSE) ;
       SetControlVisible ('AFF_TOTALHTGLODEV', TRUE);
       SetControlText ('TAFF_MONTANTECHE', 'Montant Echéance') ;
       //gm 23/12/02       SetControlVisible ('AFF_REPRISEACTIV', False);  SetControlVisible ('TAFF_REPRISEACTIV', False);
       end
    else // Par défaut
       begin
       SetControlVisible ('AFF_POURCENTAGE', FALSE) ;
       SetControlVisible ('TAFF_POURCENTAGE', FALSE) ;
       SetControlVisible ('AFF_MONTANTECHEDEV', TRUE) ;
       SetControlVisible ('TAFF_MONTANTECHE', TRUE) ;
       SetControlVisible ('BRECALCULMTGLOBAL', FALSE) ;
       SetControlText ('TAFF_MONTANTECHE', 'Montant Echéance') ;
       end;
    end
  (* surtout pas ---> declanché dans le chargement des zones lors du load de l'enreg
  else if (F.FieldName = 'AFF_PERIODICITE') then //*********** Périodicité *************
  begin
    if (GetField ('AFF_PERIODICITE') = 'S')           //Semaine
    Or (GetField ('AFF_PERIODICITE') = 'M')           //Mensuelle
    Or (GetField ('AFF_PERIODICITE') = 'P') then      //Ponctuelle
      GerePeriodicite(False)
    else if (GetField ('AFF_PERIODICITE') = 'NBI') then // Nombre d'Intervention
      GerePeriodicite(True)
    else
      GerePeriodicite(False);                         //Annuelle
  end
  *)
  else if (F.FieldName = 'AFF_DESCRIPTIF') then
     DefiniLibelle
  else if (F.FieldName = 'AFF_TIERS') then
     incidenceTiersSurAffaire (GetField ('AFF_TIERS') )
  else if (F.FieldName = 'AFF_CALCTOTHTGLO') then
     begin
     if (GetCheckBoxState ('AFF_CALCTOTHTGLO') = cbChecked) then
        begin
        // Si l'on considére que le Total HT est calculé en automatique => recalcul nécessaire.
        SetControlEnabled ('AFF_TOTALHTGLODEV',false);
        ImpactCalculTotalHT (TitRecalcul) ;
        end
     else
        SetControlEnabled ('AFF_TOTALHTGLODEV', True) ;
     end
  else if (F.FieldName = 'AFF_REGROUPEFACT') then
     begin
       if (GetField ('AFF_REGROUPEFACT') = 'AUC') and (GetField ('AFF_REGSURCAF') <> 'X') then //gm 10/10/02           // facturation indépendante
          begin
          SetControlEnabled ('AFF_PRINCIPALE', False) ;
          SetControlVisible ('BUTTONREGROUP', False) ;
          end
       else if GetField ('AFF_REGROUPEFACT') = 'ARE' then // sur affaire de regroupement
          begin
          if not (TiersIdemAffRef) then
             begin
             PGIBoxAF (TexteMsgAffaire [24] , '') ;
             SetControlText ('AFF_REGROUPEFACT', 'AUC') ;
             end
          else
             begin
             if (GetChampsAffaire (GetField ('AFF_AFFAIREREF') , 'AFF_AFFCOMPLETE') <> 'X') then
                SetControlEnabled ('AFF_PRINCIPALE', True) ;
             end;
          end
       else //pa 01/10/2001 - paramétrable... if Copy(GetField('AFF_REGROUPEFACT'),1,2)= 'RE' then // regroupement
          begin
          SetControlEnabled ('AFF_PRINCIPALE', True) ;
          SetControlVisible ('BUTTONREGROUP', True) ;
          end;
     end
  else if (F.FieldName = 'AFF_MODELE') then
     begin
     if (GetField ('AFF_MODELE') = 'X') then
        begin
        if (Modele <> 'X') and not (DS.State in [dsInsert] ) then // Affaire modèle possible uniquement en création
           begin
           PGIBoxAF (TexteMsgAffaire [16] , '') ;
           SetControlChecked ('AFF_MODELE', False) ;
           Exit;
           end
        else
	         //SetControlText ('AFF_TIERS',''); SetControlEnabled( 'AFF_TIERS' , False );
        	 SetControlText ('LIBTIERS', 'Modèle') ;
      	end
// ---MODIFBRL 010606
//   else
//        SetControlEnabled ('AFF_TIERS', DS.State in [dsInsert] ) ;
// ---
     end
  else if (F.FieldName = 'AFF_TOTALHTGLODEV') then
     ImpactCalculTotalHT (TitTotHT)
  else if (F.FieldName = 'AFF_TYPEPREVU') then
     ImpactCalculTotalHT (TitTypePrevu)
  else if (F.FieldName = 'AFF_REPRISEACTIV') then
     begin
    // PL le 28/11/02 : on deconnecte le multiecheance de l'alim par le budget
    // PL le 20/06/02 : en alimentation des echeances par le budget, on doit toujours voir la case
    // à cocher AFF_MULTIECHE
    // Dans le cas contraire on gere suivant les cas
     if (GetParamSoc ('SO_AFMULTIECHE') = false) then
        begin
        // PL le 18/06/02 : gestion multi echeances
        SetControlVisible ('AFF_MULTIECHE', False) ;
        SetField ('AFF_MULTIECHE', '-') ;
        end else
        begin
        // On gère le multi echeance
        if (GetField ('AFF_GENERAUTO') <> 'ACT') and
           (GetField ('AFF_GENERAUTO') <> 'MAN') and
           (GetField ('AFF_REPRISEACTIV') <> 'NON') and
           (GetField ('AFF_REPRISEACTIV') <> '') then
           SetControlVisible ('AFF_MULTIECHE', True)
        else
           begin
           SetControlVisible ('AFF_MULTIECHE', false) ;
           SetField ('AFF_MULTIECHE', '-') ;
           end;
        end;
	   end;

  end else // type AFF ou PRO
  begin
    // Modified by f.vautrain 29/09/2017 10:54:45
    if StatutAffaire = 'AFF' then
    begin
      if (VH_GC.BTCODESPECIF='001') then
      begin
        if (F.FieldName = 'AFF_DATEFIN') then
        begin
          DateFin := StrToDate(GetField('AFF_DATEFIN'));
          if datefin <> idate2099 then
          begin
            NbMois:= StrToInt(GetParamSoc ('SO_AFALIMGARANTI'));
            DateCal := IncMonth(DateFin, NbMois);
            SetField('AFF_DATEGARANTIE', DateToStr(DateCal));
            NbMois:= StrToInt(GetParamSoc ('SO_AFALIMCLOTURE'));
            DateCal := IncMonth(DateFin, NbMois);
            SetField('AFF_DATECLOTTECH', DateToStr(DateCal));
          end;
        end
        else if (F.FieldName = 'AFF_COEFFG')        OR (F.FieldName = 'AFF_COEFFS') OR (F.FieldName = 'AFF_COEFSAV') OR
                (F.FieldName = 'AFF_COEFFD')        OR (F.FieldName = 'AFF_TAUXRG') OR (F.FieldName = 'AFF_DATERG')  OR
                (F.FieldName = 'AFF_COEFMARG')      OR (F.FieldName = 'AFF_MTAVANCE') OR
                (F.FieldName = 'AFF_DEBRESTAVANCE') OR (F.FieldName = 'AFF_FINRESTAVANCE') then
        begin
          DS.edit;
        end;
      end;
    end;

  	if F.FieldName = 'AFF_SSTRAITANCE' then
    begin
      SetControlVisible('AFF_DATESSTRAIT',THDBCheckbox(getControl('AFF_SSTRAITANCE')).Checked );
      SetControlVisible('TAFF_DATESSTRAIT',THDBCheckbox(getControl('AFF_SSTRAITANCE')).Checked );
      if (Not THDBCheckbox(getControl('AFF_SSTRAITANCE')).Checked) then
      begin
        SetField('AFF_DATESSTRAIT',Idate1900);
      end else
      begin
        if GetField('AFF_DATESSTRAIT') = IDate1900 then
        begin
        	SetField('AFF_DATESSTRAIT',V_PGI.DateEntree);
        end;
      end;
    end else if F.FieldName = 'AFF_MANDATAIRE' then
    begin
      {$IFNDEF LINE}
      DefiniMenuCotraitance;
      {$ENDIF LINE}
    end
    else if (F.FieldName = 'AFF_DESCRIPTIF') then
    begin
    	DefiniLibelle;
  	end
    else if (F.FieldName = 'AFF_TIERS') then
    begin
    	incidenceTiersSurAffaire (GetField ('AFF_TIERS') )
    end;
  end;

end;

Procedure TOM_Affaire.GereAffichagePeriodicite;
Var DataCombo : THDBValComboBox;
begin
  SetControlVisible ('AFF_DATEFINGENER', True) ;
  SetControlVisible ('TAFF_DATEFINGENER', True) ;
  SetControlText    ('TAFF_DATEDEBGENER', 'Date &début') ;
  SetControlVisible ('NBMOIS', True) ;
  SetControlVisible ('NBTICKET', False) ;
  SetControlVisible ('TAFF_INTERVALGENER', TRUE) ;
  SetControlVisible ('TAFF_INTERVALGENER2', TRUE) ;
  SetControlVisible ('TAFF_INTERVALCONSO', False) ;
  SetControlVisible ('AFF_INTERVALCONSO', False) ;

  SetControlVisible ('TAFF_CONSOMME', false);
  SetControlVisible ('AFF_CONSOMME', false);

  SetControlEnabled('AFF_DETECHEANCE',True);
  SetControlEnabled('AFF_TERMEECHEANCE',True);
  SetControlEnabled('AFF_METHECHEANCE',true);
  SetControlEnabled('AFF_RECONDUCTION',true);

  if GetField ('AFF_PERIODICITE') = 'S' then
  begin
    SetControlText  ('TAFF_INTERVALGENER', 'Toutes les ') ;
    SetControlText  ('TAFF_INTERVALGENER2', 'Semaines') ;
    SetControlProperty ('AFF_INTERVALGENER', 'MaxValue', '52');
  end
  else if GetField  ('AFF_PERIODICITE') = 'M' then
  begin
    SetControlText  ('TAFF_INTERVALGENER', 'Tous les ') ;
    SetControlText  ('TAFF_INTERVALGENER2', 'Mois') ;
    SetControlProperty ('AFF_INTERVALGENER', 'MaxValue', '12');
  	SetControltext ('NBMOIS',   GetField('AFF_INTERVALGENER'));
  end
  else if GetField  ('AFF_PERIODICITE') = 'P' then
  begin
    SetControlVisible ('AFF_INTERVALGENER', false);
    SetControlVisible ('TAFF_INTERVALGENER', false);
    SetControlVisible ('TAFF_INTERVALGENER2', false);
    SetControlVisible ('AFF_DATEFINGENER', true) ;
    SetControlVisible ('TAFF_DATEFINGENER', true) ;
    SetControlText    ('TAFF_DATEDEBGENER', 'Date &de facturation') ;
  end
  else if GetField('AFF_PERIODICITE') = 'NBI' then
  begin
    SetControlVisible ('NBMOIS', False) ;
    SetControlVisible ('NBTICKET', True) ;
    SetControlVisible ('AFF_CONSOMME', True) ;
  	SetControlVisible ('TAFF_CONSOMME', true);
    SetControlText('TAFF_INTERVALGENER',  'Nombre Total');
    SetControlText('TAFF_INTERVALGENER2', 'de crédits');
    SetControlText('TAFF_DATEDEBGENER',  'Date debut Intervention');
    SetControlText('TAFF_DATEFINGENER',  'Date fin Intervention');
  	SetControltext ('NBTICKET', GetField('AFF_INTERVALGENER'));
    SetControlEnabled('AFF_DETECHEANCE',false);
    SetControlEnabled('AFF_TERMEECHEANCE',false);
    SetControlEnabled('AFF_METHECHEANCE',false);
    SetControlEnabled('AFF_RECONDUCTION',false);
    if ExisteEcheanceAffaire(GetField ('AFF_AFFAIRE'), 'X') then SetControlProperty('NBTICKET', 'Enabled', False);
    SetControltext ('NBTICKET', GetField('AFF_INTERVALGENER'));
    SetControltext ('AFF_CONSOMME', GetField('AFF_CONSOMME'));
    //    if Getfield('AFF_CONSOMME') <> 0 then
    //    begin
    //      SetControlEnabled('NBTICKET', false);
    //      SetControlEnabled('AFF_PERIODICITE', false);
    //    end;
  end
  else if getField('AFF_PERIODICITE') = 'A' then
  begin
    SetControlVisible ('TAFF_INTERVALGENER', false) ;
    SetControlVisible ('TAFF_INTERVALGENER2', false) ;
    SetControlVisible ('AFF_INTERVALGENER', false) ;
    SetControlVisible ('NBMOIS', false) ;
  end;

end;

procedure TOM_Affaire.ImpactCalculTotalHT (TypeImpact: T_TypeimpactTotHT) ;
var
  TotLignes, TotHT: Double;
  CompNum: THNumEdit;
  Nbeche: integer;
  BaseEche, bCalcAuto: Boolean;
begin

  if (GetField ('AFF_AFFCOMPLETE') = '-') then Exit;

(*
  // Affichage ou pas du type de prévu
  if ((GetField ('AFF_GENERAUTO') = 'FOR') or
      (GetField ('AFF_GENERAUTO') = 'ACT') or
      (GetField ('AFF_GENERAUTO') = 'CON') and
      (GetField ('AFF_TOTALHTGLODEV') <> 0)) then
     begin
     SetControlVisible ('AFF_TYPEPREVU', True) ;
     SetControlVisible ('TAFF_TYPEPREVU', True) ;
     end
  else
     begin
     SetControlVisible ('AFF_TYPEPREVU', False) ;
     SetControlVisible ('TAFF_TYPEPREVU', False) ;
     if (GetField ('AFF_GENERAUTO') = 'POT') then
        if (GetField ('AFF_TYPEPREVU') <> 'GEN') then
           SetField ('AFF_TYPEPREVU', 'GEN')
        else
        if GetField ('AFF_TYPEPREVU') = 'GEN' then
           SetField ('AFF_TYPEPREVU', 'GLO') ;
	   end;
*)
{$IFDEF BTP}
  SetControlVisible ('AFF_TYPEPREVU', False) ;
  SetControlVisible ('TAFF_TYPEPREVU', False) ;
{$ENDIF}
  // sortie rapide si l'on ne modifie que le totHt et le type de genération
  // Recalcul uniquement si le total des lignes ou passage par les écheances
  // ne pas recalculer systématiquement / pb fx appelé sur le onload ...
  if TypeImpact = TitTotHT then Exit;
  if TypeImpact = TitTypeGener then Exit;

  (*{$IFDEF EAGLCLIENT}
  TotBox:=TCheckBox(GetControl('AFF_CALCTOTHTGLO'));
  {$ELSE}
  TotBox:=TDBCheckBox(GetControl('AFF_CALCTOTHTGLO'));
  {$ENDIF}
  if TotBox = Nil then Exit; if Not(TotBox.Checked) then Exit; *)
  // Test

  bCalcAuto := (GetField ('AFF_CALCTOTHTGLO') = 'X') ;

  if not (bCalcAuto) then Exit;
  if EnSaisie then
     begin
     CompNum := THNumEdit (GetControl ('LAFF_TOTALHTDEV') ) ;
     TotLignes := CompNum.Value;
     end
  else
     TotLignes := GetField ('AFF_TOTALHTDEV') ;

  // calcul basé sur les lignes d'affaires
  //mcd 30/09/03 ajout type de prévu sinon pas de recalcul sur ligne si type rpévu sur nbre echenace
  if (GetField ('AFF_TOTALHTDEV') <> 0) then
     if (TypeImpact = TitEcheances) and
        (GetField ('AFF_TYPEPREVU') <> 'GEN') then Exit
  else // Calcul basé sur les écheances
     if (TypeImpact = TitTotLignes) or (TypeImpact = TitTypePrevu) then Exit;

  TotHT := CalculTotalAffaire (GetField ('AFF_AFFAIRE') , GetField ('AFF_GENERAUTO') , GetField ('AFF_TYPEPREVU') , TotLignes, nil, NbEche, baseEche, GetField ('AFF_DATEDEBGENER') ) ;

  // mcd 16/12/02 pour forcer modif fiche si le mtt à changer
  if (TotHt <> GetField ('AFF_TOTALHTGLODEV') ) and not (DS.State in [dsInsert, dsEdit] ) then
     DS.edit;

  SetField ('AFF_TOTALHTGLODEV', TotHT) ;

end;
/////////////////////////////////////////////////////////////////////////////

procedure TOM_Affaire.SetArguments (StSQL: string) ;
var
  Critere, ChampMul, ValMul: string;
  x, y: integer;
  Ctrl: TControl;
  Fiche: TFFiche;
begin
  SetControlVisible ('BSTOP', TRUE) ;
  DS.Edit;
  Fiche := TFFiche (ecran) ;
  repeat
    Critere := AnsiUppercase (Trim (ReadTokenPipe (StSQL, '|') ) ) ;
    if Critere <> '' then
    begin
      x := pos ('=', Critere) ;
      if x <> 0 then
      begin
        ChampMul := copy (Critere, 1, x - 1) ;
        ValMul := copy (Critere, x + 1, length (Critere) ) ;
        y := pos (',', ValMul) ;
        if y <> 0 then
          ValMul := copy (ValMul, 1, length (ValMul) - 1) ;
        if copy (ValMul, 1, 1) = '"' then
          ValMul := copy (ValMul, 2, length (ValMul) ) ;
        if copy (ValMul, length (ValMul) , 1) = '"' then
          ValMul := copy (ValMul, 1, length (ValMul) - 1) ;
        SetField (ChampMul, ValMul) ;
        Ctrl := TControl (Fiche.FindComponent (ChampMul) ) ;
        if Ctrl = nil then
          exit;
{$IFDEF EAGLCLIENT}
        if (Ctrl is TCustomCheckBox) or (Ctrl is THValComboBox) or (Ctrl is TCustomEdit) then
          TEdit (Ctrl) .Font.Color := clRed
        else
          if Ctrl is TSpinEdit then
          TSpinEdit (Ctrl) .Font.Color := clRed
        else
          if (Ctrl is TCheckBox) or (Ctrl is THValComboBox) or (Ctrl is THEdit) or (Ctrl is THNumEdit) then
        begin
          TSpinEdit (Ctrl) .Font.Color := clRed;
          SetControlText (ChampMul, ValMul) ;
        end;
{$ELSE}
        if (Ctrl is TDBCheckBox) or (Ctrl is THDBValComboBox) or (Ctrl is THDBEdit) then
          TEdit (Ctrl) .Font.Color := clRed
        else
          if Ctrl is THDBSpinEdit then
          THDBSpinEdit (Ctrl) .Font.Color := clRed
        else
          if (Ctrl is TCheckBox) or (Ctrl is THValComboBox) or (Ctrl is THEdit) or (Ctrl is THNumEdit) then
        begin
          THDBSpinEdit (Ctrl) .Font.Color := clRed;
          SetControlText (ChampMul, ValMul) ;
        end;
{$ENDIF}
      end;
    end;
  until Critere = '';
end;

/////////////////////////////////////////////////////////////////////////////
procedure TOM_Affaire.BTLigneClick (Sender: TObject) ;
var Nomfiche      : string;
    StatutAffaire : string;
    Action        : TActionFiche;
begin

  statutAffaire := GetControlText ('AFF_STATUTAFFAIRE') ;

  if StatutAffaire = 'INT' then
     Begin
     AppelPieceAffaire ('LIGNE');
     //Gestion du bouton calcul d'échéances...
     {Q := OpenSQL('SELECT GL_NUMLIGNE FROM LIGNE WHERE GL_NATUREPIECEG="AFF" AND GL_TYPELIGNE="ART" AND GL_AFFAIRE="' + GetField ('AFF_AFFAIRE') + '"', TRUE);
     if not Q.EOF then
        if Q.RecordCount > 1 then
           SetControlVisible ('BRECALCULMTGLOBAL', False)
        Else
           SetControlVisible ('BRECALCULMTGLOBAL', True);
     Ferme(Q);}
   	 CalculmontantEch;
     end
  else
     Begin
     if StatutAffaire <> 'PRO' then
		    begin
    	  Nomfiche := 'BTDEVIS_MUL';
	      AGLLanceFiche ('BTP', Nomfiche, 'GP_NATUREPIECEG=DBT;GP_AFFAIRE1=' + GetControlText ('AFF_AFFAIRE1') + ';GP_AFFAIRE2=' + GetControlText ('AFF_AFFAIRE2') + ';GP_AFFAIRE3=' + GetControlText ('AFF_AFFAIRE3') + ';GP_AFFAIRE=' + GetControlText ('AFF_AFFAIRE') + ';GP_TIERS=' + GetControlText ('AFF_TIERS') , '', 'MODIFICATION')
	      end
     else
        begin
        if EtatAffaire = 'ACP' then
           Action := taConsult
        else
           Action := TFfiche (Ecran) .typeaction;
        GestionDetailEtude ('','',GetControltext ('AFF_AFFAIRE') , (THDBCheckbox (GetControl('AFF_MANDATAIRE')).Checked), Action, TOBEtude,TatNormal) ;
        end;
     end;

end;

procedure TOM_Affaire.BordereauxClick(Sender: TObject);
var StatutAffaire : string;
    Action        : TActionFiche;
begin
  statutAffaire := GetControlText ('AFF_STATUTAFFAIRE');
  if StatutAffaire = 'AFF' then
  begin
//    if EtatAffaire = 'ACP' then Action := taConsult
//                           else Action := TFfiche(Ecran).typeaction;
		Action := TaModif;
    GestionDetailEtude ('','',Getcontroltext('AFF_AFFAIRE'),false,Action ,TOBEtude,TatBordereaux);   // Bordereaux de prix
  end;

end;

procedure TOM_Affaire.MnAppelClick(Sender: TObject);
begin

   if StatutAffaire = 'INT' then
      Begin
      if GetParamSoc('SO_BTAFFAIRERECHLIEU') then

      else
    	   AglLanceFiche('BTP','BTMULAPPELS','AFF_AFFAIRE0=W;AFF_AFFAIREINIT=' + GetField('AFF_AFFAIRE'),'','ETAT=ECO;STATUT=APP'); // Appels
      end
   else
      Begin
      if GetParamSoc('SO_BTAFFAIRERECHLIEU') then
         AglLanceFiche('BTP', 'BTMULAPPELSINT','AFF_AFFAIRE0=W;AFF_CHANTIER=' + GetField('AFF_AFFAIRE'),'','ETAT=ECO;STATUT=APP')
      else
    	   AglLanceFiche('BTP','BTMULAPPELS','AFF_AFFAIRE0=W;AFF_CHANTIER=' + GetField('AFF_AFFAIRE'),'','ETAT=ECO;STATUT=APP'); // Affaires
      end


end;

procedure TOM_Affaire.MnConsoClick(Sender: TObject);
var DateDeb : TDateTime;
    DateFin : TDateTime;
    Arg     : string;
    LastAutoSearch: boolean;
begin

  SourisSablier;

  if GetField('AFF_AFFAIRE') = '' then
  begin
    PgiError ('Le code Affaire n''est pas renseigné', application.name);
    SourisNormale;
    exit;
  end;

  Arg := 'AFFAIRE=' + GetField('AFF_AFFAIRE');

  DateDeb := StrToDate(GetField('AFF_DATEDEBUT'));
  DateFin := IDate2099;

  Arg := Arg + ';FULLSCREEN';

  LastAutoSearch := V_PGI.AutoSearch;

  V_PGI.AutoSearch := true;

  AGLLanceFiche('BTP','BTJOUCON','','', Arg + ';DATEDEB=' + DateTimeToStr(DateDeb) + ';DATEFIN=' + DateTimeToStr(DateFin)) ;

  V_PGI.AutoSearch := lastautoSearch;

  SourisNormale;

end;

procedure TOM_Affaire.MnImportDIEClick(Sender: TObject);
var NbPiece : Integer;
    ChargeOK : Boolean;
begin
{$IFNDEF UTILS}
  SPIGAOOBJ.ControleAcces;
  NbPiece := SelectPieceAffaire (GetField ('AFF_AFFAIRE') , GetField ('AFF_STATUTAFFAIRE') , CleDocAffaire);
  if NbPiece > 1 then
  begin
    PGIBoxAF (TexteMsgAffaire [31] , '') ;
    Exit;
  end else if NbPiece = -1 then
  begin
    Exit;
  end else if NbPiece = 1 then
  begin
  	if PgiAsk ('Une étude existe déjà pour cette affaire'+ chr (13) +' Voulez-vous l''écraser ?') = mryes then
    begin
      BTPSupprimePiece (CleDocAffaire,True);
      ExecuteSQL('DELETE FROM BDETETUDE Where BDE_AFFAIRE="' + GetField ('AFF_AFFAIRE') + '"');
    end else
    begin
      Exit;
    end;
  end;

  ChargeOK := SPIGAOOBJ.ChargeAO (GetControltext ('AFF_AFFAIRE'), GetControltext ('AFF_TIERS'),GetControltext ('AFF_ETABLISSEMENT'),GetField('AFF_IDSPIGAO') );
  If ChargeOK then AppelPieceAffaire ('SPIGAO');
{$ENDIF}
end;

procedure TOM_Affaire.MnExportDIEClick(Sender: TObject);
var ExportOK : Boolean;
begin
{$IFNDEF UTILS}
  SPIGAOOBJ.ControleAcces;
  ExportOK := SPIGAOOBJ.ExportAO (GetControltext ('AFF_AFFAIRE'));
{$ENDIF}
end;

procedure TOM_Affaire.MnChiffrageClick(Sender: TObject);
begin
  AppelPieceAffaire ('SPIGAO');
end;

procedure TOM_Affaire.MnVoirAffSpigaoClick(Sender: TObject);
begin
{$IFNDEF UTILS}
  SPIGAOOBJ.ControleAcces;
  SPIGAOOBJ.VoirAffaire(GetField('AFF_IDSPIGAO'));
{$ENDIF}
end;

procedure TOM_Affaire.MnTelechargerDIEClick(Sender: TObject);
begin
{$IFNDEF UTILS}
  SPIGAOOBJ.ControleAcces;
  SPIGAOOBJ.TelechargerDIE(GetField('AFF_IDSPIGAO'));
{$ENDIF}
end;

procedure TOM_Affaire.MnCodeBarreClick(Sender: TObject);
var trouve      : Boolean;
    CodeAffaire : string;
    Nature      : String;
    Action      : string;
begin

  CodeAffaire := Getfield('AFF_AFFAIRE');
  Nature      := GetField ('AFF_STATUTAFFAIRE');

  If Nature <> 'AFF' then Exit;

  if not ExisteSQL('SELECT BCB_IDENTIFCAB FROM BTCODEBARRE WHERE BCB_NATURECAB="' + Nature + '" and BCB_IDENTIFCAB="'+ CodeAffaire +'"') then
    Action := 'ACTION=CREATION'
  else
    Action := 'ACTION=MODIFICATION';

  AGLLanceFiche('BTP','BTCODEBARRE','','', Action + ';ID='+Nature+';CODE='+ CodeAffaire);

end;

//FV1 : 18/01/2017 - FS#2314 - SES : Ajouter l'accès aux actions depuis la fiche Affaire
procedure TOM_Affaire.MnActionsClick(Sender: TObject);
Var stRange : string;
begin

  stRange := 'RAC_AFFAIRE=' + GetField('AFF_AFFAIRE');
  stRange := stRange + ';RAC_TIERS=' + GetField('AFF_TIERS');

  AGLLanceFiche('RT','RTACTIONS_AFFAIRE',StRange,'',ActionToString(TFFiche(Ecran).FtypeAction)+';NOCHANGEPROSPECT');

end;

Procedure TOM_AFFAIRE.MnPrevFactureClick(Sender : TObject);
var CodeAffaire : string;
Begin

  CodeAffaire := Getfield('AFF_AFFAIRE');

  AGLLanceFiche ('BTP','BTPREVFAC','','','CODEAFFAIRE=' + CodeAffaire + ';ACTION=MODIFICATION')

end;


/////////////////////////////////////////////////////////////////////////////
//FV1 : 24/07/2013 - duplication d'un affaire sur une autre
Procedure TOM_AFFAIRE.Duplication_OnClick(Sender: TObject);
var  F : TForm ;
     OM : TOM ;
begin
  F:=TForm(ecran) ;

  if (F is TFFiche) then OM:=TFFiche(F).OM else exit;

  if (OM is TOM_AFFAIRE) then TOM_AFFAIRE(OM).Duplication('') else exit;

end;

procedure TOM_AFFAIRE.Duplication (CleDuplic: string);
var TobForm       : TOB;
    StTable       : String;
    StCleDuplic   : String;
    G_CodeAffaire : THCritMaskEdit;
    QQ            : TQuery ;
    stPlus        : string;
    sIncoterm     : String;
begin

  //En remplacant les bonnes zones on a peut-être une chance que ça marche...
  //en même temps allons savoir je suis tellement nul !!!!!
  //FV1 - 13/03/2017 : FS#2409 - SES ETANCH - Duplication affaire ne reprends pas toutes les info
  AffaireToDuplic:=CleDuplic;

  {$IFDEF EAGLCLIENT}
  StTable:='TIERS' ;
  {$ELSE}
  StTable := GetTableNameFromDataSet(DS);
  {$ENDIF}

  TobForm := TOB.Create (StTable, Nil, -1);

  if TobForm <> Nil then
  begin
    if AffaireToDuplic = '' then
    begin
      if (DS.State=dsInsert) then  // si nouvelle fiche  recherche record à dupliquer
      begin
        G_CodeAffaire := THCritMaskEdit.Create (nil);
        DispatchRecherche (G_CodeAffaire, 2, stPlus, '', '');
        AffaireToDuplic := G_CodeAffaire.Text;
        QQ:=OPENSQL('SELECT * From AFFAIRE WHERE AFF_AFFAIRE="'+AffaireToDuplic+'"',True);
        if not TobForm.SelectDB ('', QQ) then AffaireToDuplic := '' ;
        Ferme(QQ);
        G_CodeAffaire.Free;
      end
      else
      begin
        if not TobForm.SelectDB ('',TFFiche(Ecran).QFiche) then AffaireToDuplic :='' ;
        AffaireToDuplic :='*' ;
        TFFiche (Ecran).Bouge (NbInsert);
        AffaireToDuplic := tobForm.GetValue('AFF_AFFAIRE');
      end;
    end
    else
    begin
      if not TobForm.SelectDB ('"' + AffaireToDuplic + '"', Nil) then AffaireToDuplic :='' ;
    end;
    //
    if AffaireToDuplic <> '' then
    begin
      ReInitDuplication(TobForm);
      TobToEcran(TobForm);
      if (VH_GC.BTCODESPECIF='001') and (StatutAffaire = 'AFF') then
      begin
        SetControlText ('AFF_RESPONSABLE1', TobForm.GetValue('AFF_RESPONSABLE'));
        if (GetControlText('AFF_RESPONSABLE1')<>'') then
        begin
          SetControlproperty('LIBAFF_RESPONSABLE1', 'Visible', True);
          SetControlText('LIBAFF_RESPONSABLE1', GetLibResponsable(GetControlText('AFF_RESPONSABLE1')));
        end;
      end;
      //FV1 - 13/03/2017 : FS#2409 - SES ETANCH - Duplication affaire ne reprends pas toutes les info
      SetField ('AFF_NUMEROCONTACT', TobForm.GetValue('AFF_NUMEROCONTACT'));
      NomContactAff;
    end;
  end;

  SetControlvisible('BMEMO', False);
  //
  TobForm.Free;

  SetActiveTabSheet (TFFiche(Ecran).pages.Pages[0].name);
  SetField('AFF_AFFAIRE','');
  SetControlText ('CAFF_AFFAIRE', ''); // DBR - suite fiche 10655 - pour avoir en duplication le message code obligatoire
  SetFocusControl ('CAFF_AFFAIRE0');
  SetControlVisible('BDUPLICATION',FALSE);
  //
end;

//initialisation des zones de tob dans le cas de la duplication...
procedure TOM_AFFAIRE.ReInitDuplication(TobForm : TOB);
begin
  TobForm.PutValue('AFF_DATECREATION', Date);       // ajout PCS 28052002
  TobForm.PutValue('AFF_DATEMODIF', Date);          // ajout PCS 28052002
  TobForm.PutValue('AFF_UTILISATEUR', V_PGI.User);   // ajout PCS 28052002
end;

procedure TOM_AFFAIRE.tobToEcran (TobForm : TOB);
var i : integer ;
    CC : TControl ;
    CC_RO : Boolean ;
begin
        for i := 1 to TobForm.NbChamps do
            begin
            CC:=TControl(Ecran.findcomponent(TobForm.GetNomChamp(i))) ;
{$IFDEF EAGLCLIENT}
            CC_RO:=((CC is TCustomEdit) and (TEdit(CC).ReadOnly)) ;
{$ELSE}
            CC_RO:=((CC is TEdit) and (TEdit(CC).ReadOnly)) or ((CC is TDBEdit) and (TDBEdit(CC).ReadOnly)) ;
{$ENDIF}
            if (CC<>nil) and (CC.enabled and ( not CC_RO) and cc.Visible) then SetField (TobForm.GetNomChamp(i), TobForm.GetValeur (i));
            end;
end;

procedure TOM_Affaire.OnArgument (stArgument: string) ;
var
  i				: Integer;
  X				: integer;
  Tmp			: String;
  champ		: String;
  valeur	:	String;
  st			: string;
  //SPlus		: String;
  CodeChantier : string;
  menuPerso: TMenuItem;
  //bIsoFlex : Boolean;
{$IFDEF EAGLCLIENT}
  CBComplete: TCheckBox;
{$ELSE}
  CBComplete: TDBCheckBox;
{$ENDIF}
  MenuPopBtp: TPopupMenu;
  CC: THvalComboBox;
begin

  inherited;
  //
  fIdCliSPIGAo := -1;
  FromSPIGAONew := false;
  Ecran.OnCloseQuery := FormCloseQuery;
  //
  AppliqueFontDefaut (THRichEditOle(GetControl('AFF_DESCRIPTIF')));
	fModeGestion := TtgStd;
	PlusGenere := '';
  TFFiche(ecran).OnAfterFormShow := AfterFormShow;
  {$IFDEF BTP}
  THLabel (GetControl ('LIBTIERS') ) .caption := '';
  //

  BtLigne := TToolbarButton97 (ecran.FindComponent ('BTLIGNE') ) ;
  BTLigne.OnClick := BTLigneClick;

  BtTache := TToolbarButton97 (ecran.FindComponent ('BTACHES') ) ;
  BTTache.OnClick := BTTacheClick;

  MnBordereaux := TMenuItem (ecran.FindCOmponent('MnBordereaux'));
  if MnBordereaux <> nil then MnBordereaux.OnClick := BordereauxClick;

  MnAppel := TMenuItem (ecran.FindCOmponent('MnListAppel'));
  if MnAppel <> nil then MnAppel.OnClick := MnAppelClick;

  MnConso := TMenuItem (ecran.FindComponent('MnListConso'));
  if MnConso <> nil then MnConso.OnClick := MnConsoClick;

  MnImportDIE := TMenuItem (ecran.FindComponent('MnImportDIE'));
  if MnImportDIE <> nil then MnImportDIE.OnClick := MnImportDIEClick;

  MnExportDIE := TMenuItem (ecran.FindComponent('MnExportDIE'));
  if MnExportDIE <> nil then MnExportDIE.OnClick := MnExportDIEClick;

  MnChiffrage := TMenuItem (ecran.FindComponent('MnChiffrage'));
  if MnChiffrage <> nil then MnChiffrage.OnClick := MnChiffrageClick;

  MnVoirAffSpigao := TMenuItem (ecran.FindComponent('MnVoirAffSpigao'));
  if MnVoirAffSpigao <> nil then
  begin
    MnVoirAffSpigao.OnClick := MnVoirAffSpigaoClick;
    MnVoirAffSpigao.Enabled := False;
  end;

  MnCodeBarre := TMenuItem (ecran.FindComponent('MnCodeBarre'));
  if MnCodeBarre <> nil then
  begin
    MnCodeBarre.OnClick := MnCodeBarreClick;
    MnCodeBarre.Visible := False;          
  end;

{$IFDEF V10}                                                 
  MnPrevFacture := TMenuItem (ecran.FindComponent('MnPrevFacture'));
  if MnPrevFacture <> nil then MnPrevFacture.OnClick := MnPrevFactureClick;
{$ENDIF}
  //FV1 : 18/01/2017 - FS#2314 - SES : Ajouter l'accès aux actions depuis la fiche Affaire
  MnActions := TMenuItem (ecran.FindComponent('MnActions'));
  if MnActions <> nil then
  begin
    MnActions.OnClick := MnActionsClick;
    MnActions.Visible := True;
  end;

  MnTelechargerDIE := TMenuItem (ecran.FindComponent('MnTelechargerDIE'));
  if MnTelechargerDIE <> nil then
  begin
{$IFNDEF UTILS}
	  if SPIGAOOBJ.IsNewGestion then
    begin
      MnTelechargerDIE.Visible := false;
    end else
    begin
      MnTelechargerDIE.OnClick := MnTelechargerDIEClick;
      MnTelechargerDIE.Enabled := False;
    end;
{$ELSE}
      MnTelechargerDIE.Visible := false;
{$ENDIF}
  end;

  if Assigned(GetControl('AFF_TYPEPAIE')) then
    ThTypePaie := THvalComboBox(ecran.FindComponent('AFF_TYPEPAIE'));

  if Assigned(GetControl('BDUPLICATION')) then
  Begin
	   BDuplication 		    := TToolbarButton97(ecran.FindComponent('BDUPLICATION'));
  	 BDuplication.onclick := Duplication_OnClick;
  End;

  //FV1 - 10/01/2018 : FS#2778 - VEODIS GROUP - Les affaires sont supprimables alors que l'annulation est refusé dans les droits
  if Assigned(GetControl('BDelete')) then
  Begin
	   BDelete := TToolbarButton97(ecran.FindComponent('BDelete'));
     //FV1 - 10/01/2018 : FS#2778 - VEODIS GROUP - Les affaires sont supprimables alors que l'annulation est refusé dans les droits
     BDelete.visible := SuppAffaireAutorise;
  end;

  TOBEtude := nil;

	{$ENDIF}

	AffaireToDuplic := '';
	DuplicFromMenu := False;

  Tmp := '';
  StatutAffaire := '';
  NewTiers := '';

  //FV1 - 13/03/2017 : FS#2409 - SES ETANCH - Duplication affaire ne reprends pas toutes les info
  BDUPLICATION.Visible := True;

  i := pos ('MODIFLOT', stArgument) ;
  ModifLot := i <> 0;

  if ModifLot then
  begin
    TFfiche (Ecran) .MonoFiche := true;
    StSQL := copy (stArgument, i + 9, length (stArgument) ) ;
  end
  else
  begin
    AffaireToDuplic := GetArgumentValue(stArgument, 'DUPLICATION');
  	if AffaireToDuplic <> '' then
    begin
      DuplicFromMenu := True;
      BDUPLICATION.Visible := False;
    end;
    fAffaire := GetArgumentValue(StArgument, 'AFF_AFFAIRE');
    ///
    Tmp := (Trim (ReadTokenSt (stArgument) ) ) ;
    while (Tmp <> '') do
    begin
      if Tmp <> '' then
      begin
        X := pos (':', Tmp) ;
        if x = 0 then
          X := pos ('=', Tmp) ;
        if x <> 0 then
           begin
           Champ := copy (Tmp, 1, X - 1) ;
           Valeur := Copy (Tmp, X + 1, length (Tmp) - X) ;
           end
        else
          Champ := Tmp;
        if Champ = 'STATUT'       then StatutAffaire := valeur;
        if Champ = 'ETAT'         then EtatAffaire := valeur;
        if Champ = 'AFF_TIERS'    then NewTiers := valeur;
        if Champ = 'COTRAITANCE'  then fModeGestion := TTgCotrait;
        if Champ = 'SPIGAONEW'    then FromSPIGAONew := true;
      end;
      Tmp := (Trim (ReadTokenSt (stArgument) ) ) ;
    end;
  end;

  if (VH_GC.BTCODESPECIF='001') and (StatutAffaire = 'AFF') then
  begin
    SetControlVisible('AFF_RESPONSABLE',false);
    SetControlVisible('TAFF_RESPONSABLE',false);
    if Assigned(GetControl('AFF_RESPONSABLE1')) Then
    begin
      TResponsable := THDBEdit(GetControl('AFF_RESPONSABLE1'));
      TResponsable.OnElipsisClick := BRechResponsable;
    end;
    if assigned (GetControl('AFF_RESSOURCE1')) then
    begin
      THDbEdit(GetControl('AFF_RESSOURCE1')).OnElipsisClick := BRechResponsable;
    end;
    if assigned (GetControl('AFF_RESSOURCE2')) then
    begin
      THDbEdit(GetControl('AFF_RESSOURCE2')).OnElipsisClick := BRechResponsable;
    end;
    if assigned (GetControl('AFF_RESSOURCE3')) then
    begin
      THDbEdit(GetControl('AFF_RESSOURCE3')).OnElipsisClick := BRechResponsable;
    end;
  end else
  begin
    //FV1 - 14/06/2016 - FS#2023 - En fiche affaire, la recherche sur responsable ne fonctionne plus
    if Assigned(GetControl('AFF_RESPONSABLE')) Then
    begin
      TResponsable := THDBEdit(GetControl('AFF_RESPONSABLE'));
      TResponsable.OnElipsisClick := BRechResponsable;
    end;
  end;

  if EtatAffaire <> 'SPIGAO' Then Spigao := False
  else
  begin
    Spigao := True  ;
    EtatAffaire := '';
  end;

  // Test droit d'accès en création
  if not (CreationAffaireAutorise) then
  begin
    SetControlVisible ('bInsert', False) ;
    SetControlVisible ('BDUPLICATION',CreationAffaireAutorise);
  end;

  if StatutAffaire = '' then StatutAffaire := 'AFF';

  BTChantier := TToolbarButton97(ecran.FindComponent('BTChantier'));
  BTChantier.OnClick := VoirChantier;

  SetControlProperty('TAFF_TYPEAFFAIRE', 'Visible', False);
  SetControlProperty('AFF_TYPEAFFAIRE',  'Visible', False);
  SetControlVisible ('BTNSOUSTRAIT', VH_GC.BTCODESPECIF='001') ;

  if StatutAffaire = 'PRO' then
     begin
{$IFDEF BTP}
    // Traduction des libéllés en fonction si proposition ou affaire
    if spigao then
      Ecran.Caption := TraduitGa ('Appel d''offre SPIGAO')
    else
      Ecran.Caption := TraduitGa ('Appel d''offre') ;
    UpdateCaption (Ecran) ;
    SetControlText ('TAFF_DESCRIPTIF', TraduitGa ('Descriptif de l''appel d''offre') ) ;
    SetControlProperty ('AFF_AFFAIREHT', 'Caption', 'Appel d''offre &HT') ;
    SetControlProperty ('TAFF_AFFAIRE', 'Caption', 'Appel d''&offre') ;
    SetControlProperty ('BTREGROUPE', 'Caption', 'Liste des appels d''offres re&groupées') ;
    SetControlProperty ('AFF_ADMINISTRATIF', 'Caption', 'Appel d''offre &administratif') ;
    SetControlProperty ('AFF_MODELE', 'Caption', 'Appel d''offre &modèle') ;
    SetControlProperty ('BTLIGNE', 'hint', 'Documents associés') ;
    SetControlProperty ('TAFF_REGROUPEFACT1', 'caption', 'Rgpt Etudes sur accept.') ;
    SetControlVisible ('AFF_DETECHEANCE', false) ;
    SetControlVisible ('TAFF_GENERAUTO_', false) ;
    SetControlVisible ('AFF_GENERAUTO', false) ;
    SetControlVisible ('TAFF_PROFILGENER', False) ;
    SetControlVisible ('AFF_PROFILGENER', False) ;
    SetControlVisible ('AFF_DATESIGNE', False) ;
    SetControlVisible ('AFF_DATELIMITE', False) ;
    SetControlVisible ('AFF_DATEGARANTIE', False) ;
    SetControlVisible ('AFF_DATECLOTTECH', False) ;
    SetControlVisible ('TAFF_DATESIGNE', False) ;
    SetControlVisible ('TAFF_DATELIMITE', False) ;
    SetControlVisible ('TAFF_DATEGARANTIE', False) ;
    SetControlVisible ('TAFF_DATECLOTTECH', False) ;
    setControlVisible ('TABSHEETSTAT', True) ;
    setControlVisible ('BTACHES', False) ;
    if TToolBarButton97(GetControl('BPOPMENU')) <> nil then TToolBarButton97(GetControl('BPOPMENU')).visible := false;
{$ELSE}
    Ecran.Caption := TraduitGa ('Proposition d''affaire') ;
    UpdateCaption (Ecran) ;
    SetControlText ('TAFF_DESCRIPTIF', TraduitGa ('Descriptif de la proposition d''affaire') ) ;
    SetControlProperty ('AFF_AFFAIREHT', 'Caption', 'Proposition &HT') ;
    SetControlProperty ('TAFF_AFFAIRE', 'Caption', '&Proposition') ;
    SetControlProperty ('BTREGROUPE', 'Caption', 'Liste des propositions re&groupées') ;
    SetControlProperty ('AFF_ADMINISTRATIF', 'Caption', 'Proposition &administrative') ;
    SetControlProperty ('AFF_MODELE', 'Caption', 'Proposition &modèle') ;
{$ENDIF}
    SetControlVisible ('AFF_ETATAFFAIRE', False) ;
    SetControlVisible ('TAFF_ETATAFFAIRE', False) ;
    If Assigned(GetControl('MnRegEches')) then SetControlVisible ('MnRegEches', False) ;
    SetControlVisible ('MnSep', False) ;
  end
  else if StatutAffaire = 'INT' then
  begin
    OnArgumentContrat;
    If Assigned(GetControl('MnRegEches')) then THMenuItem(GetControl('MnRegEches')).OnClick := RegroupeEches;
  end
  else
     begin
    if assigned(GetControl('AFF_ETATAFFAIRE')) then SetControlProperty ('AFF_ETATAFFAIRE', 'Plus', GetPlusEtatAffaire (True) ) ;
    //
{$IFDEF BTP}
    BTChantier.visible := false;
    StTiers   := THDBEdit(ecran.FindComponent('AFF_TIERS'));
    SetControlVisible ('AFF_ETATAFFAIRE', False) ;
    SetControlVisible ('TAFF_ETATAFFAIRE', False) ;
    SetControlProperty ('AFF_GENERAUTO', 'Plus', 'BTP'+PlusGenere) ;
    SetControlProperty ('AFF_GENERAUTO_', 'Plus', 'BTP'+PlusGenere) ;
    SetControlVisible ('AFF_DATESIGNE', False) ;
    SetControlVisible ('AFF_DATELIMITE', False) ;
    // Modified by f.vautrain 29/09/2017 11:19:15
    if StatutAffaire = 'AFF' then
    Begin
      SetControlVisible ('AFF_DATEGARANTIE',  VH_GC.BTCODESPECIF='001') ;
      SetControlVisible ('AFF_DATECLOTTECH',  VH_GC.BTCODESPECIF='001') ;
      SetControlVisible ('TAFF_DATEGARANTIE', VH_GC.BTCODESPECIF='001') ;
      SetControlVisible ('TAFF_DATECLOTTECH', VH_GC.BTCODESPECIF='001') ;
    end;
    //
    SetControlVisible ('TAFF_DATESIGNE', False) ;
    SetControlVisible ('TAFF_DATELIMITE', False) ;
    SetControlVisible ('AFF_DETECHEANCE', false) ;
    SetControlVisible ('TABSHEETSTAT', True) ;
    SetControlVisible ('BTACHES', False) ;
    if V_PGI.LaSerie = S5                 then SetControlVisible ('BTETUDES', True) ;
    If Assigned(GetControl('MnRegEches')) then SetControlVisible ('MnRegEches', False) ;
    SetControlVisible ('MnSep', False) ;
{$ELSE}
    SetControlProperty ('AFF_GENERAUTO', 'Plus', 'GA') ;
    if ctxScot in V_PGI.PGIContexte then
    begin
      SetControlProperty ('AFF_GENERAUTO', 'Plus', 'GA" AND CO_CODE<>"CON') ;
      SetControlProperty ('AFF_GENERAUTO_', 'Plus', 'GA" AND CO_CODE<>"CON') ;
    end;
{$ENDIF}
  end;

  // Modified by f.vautrain 28/09/2017 16:38:42
  if (StatutAffaire = 'AFF') then
  begin
    if Assigned(GetControl('TABSHEETDATAGESTION'))  then TabSheetDataGestion  := TTabSheet(GetControl('TABSHEETDATAGESTION'));
    if (VH_GC.BTCODESPECIF='001') then
    begin
      if Assigned(GetControl('AFF_COEFFG'))         then CoefFG               := THNumEdit(GetControl('AFF_COEFFG'));
      if Assigned(GetControl('AFF_COEFFS'))         then CoefFS               := THNumEdit(GetControl('AFF_COEFFS'));
      if Assigned(GetControl('AFF_COEFSAV'))        then CoefSav              := THNumEdit(GetControl('AFF_COEFSAV'));
      if Assigned(GetControl('AFF_COEFFD'))         then CoefFD               := THNumEdit(GetControl('AFF_COEFFD'));
      if Assigned(GetControl('AFF_COEFMARG'))       then CoefMarg             := THNumEdit(GetControl('AFF_COEFMARG'));
      if Assigned(GetControl('AFF_TAUXRG'))         then TauxRG               := THNumEdit(GetControl('AFF_TAUXRG'));
      if Assigned(GetControl('AFF_DATERG'))         then DateRG               := THDBEdit (GetControl('AFF_DATERG'));
      if Assigned(GetControl('AFF_MTAVANCE'))       then MtAvance             := THNumEdit(GetControl('AFF_MTAVANCE'));
      if Assigned(GetControl('AFF_DEBRESTAVANCE'))  then DebRestAvance        := THNumEdit(GetControl('AFF_DEBRESTAVANCE'));
      if Assigned(GetControl('AFF_FINRESTAVANCE'))  then FinRestAvance        := THNumEdit(GetControl('AFF_FINRESTAVANCE'));
      //
      if Assigned(GetControl('BAR_MTCAUTION0'))     then MtCaution0           := THNumEdit(GetControl('BAR_MTCAUTION0'));
      if Assigned(GetControl('BAR_MTCAUTION1'))     then MtCaution1           := THNumEdit(GetControl('BAR_MTCAUTION1'));
      if Assigned(GetControl('BAR_MTCAUTION2'))     then MtCaution2           := THNumEdit(GetControl('BAR_MTCAUTION2'));
      if Assigned(GetControl('BAR_MTCAUTION3'))     then MtCaution3           := THNumEdit(GetControl('BAR_MTCAUTION3'));
      if Assigned(GetControl('BAR_MTCAUTION4'))     then MtCaution4           := THNumEdit(GetControl('BAR_MTCAUTION4'));
      if Assigned(GetControl('BAR_MTCAUTION5'))     then MtCaution5           := THNumEdit(GetControl('BAR_MTCAUTION5'));
      if Assigned(GetControl('BAR_MTCAUTION6'))     then MtCaution6           := THNumEdit(GetControl('BAR_MTCAUTION6'));
      if Assigned(GetControl('BAR_MTCAUTION7'))     then MtCaution7           := THNumEdit(GetControl('BAR_MTCAUTION7'));
      if Assigned(GetControl('BAR_MTCAUTION8'))     then MtCaution8           := THNumEdit(GetControl('BAR_MTCAUTION8'));
      if Assigned(GetControl('BAR_MTCAUTION9'))     then MtCaution9           := THNumEdit(GetControl('BAR_MTCAUTION9'));
      //
      if Assigned(GetControl('BAR_NUMCAUTION0'))    then NumCaution0          := THNumEdit(GetControl('BAR_NUMCAUTION0'));
      if Assigned(GetControl('BAR_NUMCAUTION1'))    then NumCaution1          := THNumEdit(GetControl('BAR_NUMCAUTION1'));
      if Assigned(GetControl('BAR_NUMCAUTION2'))    then NumCaution2          := THNumEdit(GetControl('BAR_NUMCAUTION2'));
      if Assigned(GetControl('BAR_NUMCAUTION3'))    then NumCaution3          := THNumEdit(GetControl('BAR_NUMCAUTION3'));
      if Assigned(GetControl('BAR_NUMCAUTION4'))    then NumCaution4          := THNumEdit(GetControl('BAR_NUMCAUTION4'));
      if Assigned(GetControl('BAR_NUMCAUTION5'))    then NumCaution5          := THNumEdit(GetControl('BAR_NUMCAUTION5'));
      if Assigned(GetControl('BAR_NUMCAUTION6'))    then NumCaution6          := THNumEdit(GetControl('BAR_NUMCAUTION6'));
      if Assigned(GetControl('BAR_NUMCAUTION7'))    then NumCaution7          := THNumEdit(GetControl('BAR_NUMCAUTION7'));
      if Assigned(GetControl('BAR_NUMCAUTION8'))    then NumCaution8          := THNumEdit(GetControl('BAR_NUMCAUTION8'));
      if Assigned(GetControl('BAR_NUMCAUTION9'))    then NumCaution9          := THNumEdit(GetControl('BAR_NUMCAUTION9'));
      //
      if TabSheetDataGestion <> nil then TabSheetDataGestion.TabVisible  := True;
      //
      If MtCaution0 <> nil then MtCaution0.OnExit := ExitLigCaution;
      If MtCaution1 <> nil then MtCaution1.OnExit := ExitLigCaution;
      If MtCaution2 <> nil then MtCaution2.OnExit := ExitLigCaution;
      If MtCaution3 <> nil then MtCaution3.OnExit := ExitLigCaution;
      If MtCaution4 <> nil then MtCaution4.OnExit := ExitLigCaution;
      If MtCaution5 <> nil then MtCaution5.OnExit := ExitLigCaution;
      If MtCaution6 <> nil then MtCaution6.OnExit := ExitLigCaution;
      If MtCaution7 <> nil then MtCaution7.OnExit := ExitLigCaution;
      If MtCaution8 <> nil then MtCaution8.OnExit := ExitLigCaution;
      If MtCaution9 <> nil then MtCaution9.OnExit := ExitLigCaution;
      //
    end
    else
    begin
      if TabSheetDataGestion <> nil then TabSheetDataGestion.TabVisible  := False;
    end;
  end;

//  SetControlProperty ('AFF_PERIODICITE', 'Plus', 'AND CO_LIBRE="AFF"') ; // pas de reprise périodicité journalière
  SetControlVisible ('AFF_COMPTAAFFAIRE', GetParamSocSecur('SO_GCVENTCPTAAFF', False)) ;
  SetControlVisible ('TAFF_COMPTAAFFAIRE', GetParamSocSecur('SO_GCVENTCPTAAFF', False)) ;

  if TFfiche (Ecran).typeaction = TaConsult then
  	 begin
     SetControlVisible ('BCALCULDATES', False) ; //MCD
     SetControlVisible ('BDUPLICATION', False) ; //MCD
  	 end;

  if TFfiche (Ecran) .typeaction = TaModif then
    SetCleAffaireEnabled (false)
  else
    SetCleAffaireEnabled (true) ;

  // Gestion des com des apporteurs sur options
  {*
  if not (VH_GC.AFGestionCom) then
  	 begin
     SetControlVisible ('AFF_APPORTEUR', False) ;
     SetControlVisible ('TAFF_APPORTEUR', False) ;
     SetControlVisible ('LIBAFF_APPORTEUR', False) ;
  	 end;
  *}
  
  if (ctxScot in V_PGI.PGIContexte) and (GetParamSoc ('SO_AfFormatExer') = 'AUC') then
  	 begin
     SetControlVisible ('GroupBoxDateExer', False) ;
  	 end;

  if (ctxGCAFF in V_PGI.PGIContexte) then
  	 begin
     SetControlVisible ('BTLIGNE', VH_GC.GASeria) ;
     SetControlVisible ('BCALCULDATES', VH_GC.GASeria) ;
     // conditions de règlement de la pièce à cacher  ...
  	 end;

  //{$IFDEF BTP}
  //SetControlProperty('AFF_GENERAUTO','Plus','BTP');
  //{$ELSE}
  //SetControlProperty('AFF_GENERAUTO','Plus','GA');
  //if ctxScot in V_PGI.PGIContexte then  SetControlProperty('AFF_GENERAUTO','Plus','GA" AND CO_CODE<>"CON');
  //{$ENDIF}

  // Gestion de la date de facture liquidative O ou N
  if (GetParamSoc ('SO_AFGERELIQUIDE') = false) then
  begin
    SetControlVisible ('AFF_DATEFACTLIQUID', False) ;
    SetControlVisible ('TAFF_DATEFACTLIQUID', False) ;
  end;

  // Paramétrage des libellés des tables libres en gescom
  GCMAJChampLibre (TForm (Ecran) , False, 'COMBO', 'AFF_LIBREAFF', 10, '') ;
  GCMAJChampLibre (TForm (Ecran) , False, 'EDIT', 'AFF_VALLIBRE', 3, '') ;
  GCMAJChampLibre (TForm (Ecran) , False, 'EDIT', 'AFF_DATELIBRE', 3, '') ;
  GCMAJChampLibre (TForm (Ecran) , False, 'EDIT', 'AFF_CHARLIBRE', 3, '') ;
  GCMAJChampLibre (TForm (Ecran) , False, 'EDIT', 'AFF_RESSOURCE', 3, '') ;
  GCMAJChampLibre (TForm (Ecran) , False, 'BOOL', 'AFF_BOOLLIBRE', 3, '') ;

  menuPerso := TMenuItem (GetControl ('mnAffRent') ) ;
  if menuPerso <> nil then menuPerso.Visible := false;

  //ONYX
  if TToolBarButton97 (GetControl ('BREVISION') ) <> nil then
  begin
    SetControlVisible ('BREVISION', GetParamSoc ('SO_AFREVISIONPRIX') ) ;
    if TpopUpMenu (GetControl ('POPMENU1') ) <> nil then
    begin
      TpopUpMenu (GetControl ('POPMENU1') ) .Items [0] .Onclick := mnFormuleOnClick;
      TpopUpMenu (GetControl ('POPMENU1') ) .Items [1] .Onclick := mnCalculOnClick;
      //      TpopUpMenu(GetControl('POPMENU1')).Items[2].Onclick := mnElementsOnClick;
      TpopUpMenu (GetControl ('POPMENU1') ) .Items [3] .Onclick := mnAppliquerCoefOnClick;
      TpopUpMenu (GetControl ('POPMENU1') ) .Items [4] .Onclick := mnDesappliquerCoefOnClick;
      TpopUpMenu (GetControl ('POPMENU1') ) .Items [5] .Onclick := mnRetablirCoefOnClick;

      THEdit (GetControl ('AFF_FORCODE1') ) .onChange := FormuleOnChange;
      THEdit (GetControl ('AFF_FORCODE2') ) .onChange := FormuleOnChange;
      THEdit (GetControl ('AFF_FORCODE1') ) .onExit := FormuleOnExit;
      THEdit (GetControl ('AFF_FORCODE2') ) .onExit := FormuleOnExit;
      SetControlText ('LIBELLEFORMULE1', '') ;
      SetControlText ('LIBELLEFORMULE2', '') ;
      TpopUpMenu (GetControl ('POPMENU1') ) .Items [2] .Visible := false;
      TpopUpMenu (GetControl ('POPMENU1') ) .Items [6] .Visible := false;
      SetControlVisible ('TB_REVISIONPRIX', False) ;
    end;
  end;

  // Affichage bouton isoflex
  GereIsoflex;

  // Montrer l'option tâche en fx du module planning ...
  if VH_GC.GAPlanningSeria then
    SetControlVisible ('BTACHES', True) ;
  // Gestion des sous affaires
  if GereSousAffaire then
    SetControlVisible ('BREGROUPE', True) ;
  // Gestion des affaires complètes + Light
{$IFDEF BTP}
  SetControlVisible ('AFF_AFFCOMPLETE', False) ; // en attendant des jours meilleurs
{$ELSE}
  SetControlVisible ('AFF_AFFCOMPLETE', GereAffCompleteEtLight) ;
{$ENDIF}
  // Evenements su checkBox Affaire complète
{$IFDEF EAGLCLIENT}
  CBComplete := TCheckBox (GetControl ('AFF_AFFCOMPLETE') ) ;
{$ELSE}
  CBComplete := TDBCheckBox (GetControl ('AFF_AFFCOMPLETE') ) ;
{$ENDIF}
  CBComplete.OnClick := AFF_AFFCOMPLETEOnClick;
  SetControlVisible ('AFF_GROUPECONF', (GetParamsoc ('SO_AFTYPECONF') = 'AGR') ) ;
  SetControlVisible ('TAFF_GROUPECONF', (GetParamsoc ('SO_AFTYPECONF') = 'AGR') ) ;
  if VH_GC.GCIfDefCEGID then
    OnArgumentCEGID;

  if ctxScot in V_PGI.PGIContexte then
    MenuPopBtp := TPopupMenu (GetControl ('POPUPMENU') )
  else
    MenuPopBtp := TPopupMenu (GetControl ('POPMENU') ) ;
  if MenuPopBtp <> nil then
    for i := 0 to MenuPopBtp.Items.Count - 1 do
    begin
      menuPerso := MenuPopBtp.Items [i] ;
      st := AnsiUppercase (MenuPerso.name) ;
{$IFDEF BTP}
      if (st = 'MNTB') {or (st = 'MNSGED')} or (st = 'MNREGLEMENT') or
        (st = 'MNPIECECOURS') or (st = 'MNTIERS') then
        MenuPerso.Visible := false;
{$ELSE}
      if not (VH_GC.GAAchatSeria) then
      begin
        if (ST = 'MNLISTEPIECEACHAT') then
          MenuPerso.Visible := false;
      end;
      if (ctxScot in V_PGI.PGIContexte) and (GetparamSoc ('So_AfGestionAppreciation') = false)
        and (st = 'MNAPPREC') then
        MenuPerso.Visible := false;
{$IFDEF DIFUSGI}
      if (ST = 'MNBUDGET') then
        MenuPerso.Visible := false;
{$ENDIF}
      if (ST = 'MNAFFAIREPIECE') and (GetparamSoc ('SO_AFPARPIECEAFF') = false) then
        MenuPerso.Visible := false;
      //if (ST = 'MNAFFAIREPIECE') then   MenuPerso.Visible := false;
{$ENDIF}
    end;

{$IFNDEF BTP}
  // Gestion auto de l'initialisation du descriptif de la mission
  if (GetParamSoc ('SO_AFALIMDESCAUTO') <> '') then
  begin
    THEdit (GetControl ('CAFF_AFFAIRE1') ) .OnExit := CAFF_AFFAIRE1Exit;
    THEdit (GetControl ('CAFF_AFFAIRE2') ) .OnExit := CAFF_AFFAIRE2Exit;
  end;
{$ENDIF}

  // PL le 20/06/02 : en alimentation des echeances par le budget, on doit toujours voir la case
  // à cocher AFF_MULTIECHE
  //if (GetParamSoc('SO_AFALIMECHE')= true) and (GetParamSoc('SO_AFMULTIECHE')= true) then
  //    begin
  //    SetControlVisible('AFF_MULTIECHE',true);
  //    end;
  //////////////////////
{$IFDEF CCS3}
  if (getcontrol ('GROUPBOXDATELIBRE') <> nil) then
    SetControlVisible ('GROUPBOXDATELIBRE', False) ;
  if (getcontrol ('GROUPBOXVALLIBRE') <> nil) then
    SetControlVisible ('GROUPBOXVALLIBRE', False) ;
  if (getcontrol ('GROUPBOXDECISION') <> nil) then
    SetControlVisible ('GROUPBOXDECISION', False) ;
  if (getcontrol ('GROUPBOXTEXTLIBRE') <> nil) then
    SetControlVisible ('GROUPBOXTEXTLIBRE', False) ;
  if (getcontrol ('GROUPBOXRESSOURCE') <> nil) then
    SetControlVisible ('GROUPBOXRESSOURCE', False) ;
{$ENDIF}

  PositionneEtabUser ('AFF_ETABLISSEMENT') ;
  if not VH^.EtablisCpta then
  begin
  	if THLabel(GetControl('TAFF_ETABLISSEMENT')) <> nil then THLabel(GetControl('TAFF_ETABLISSEMENT')).Visible := false;
			GetControl('AFF_ETABLISSEMENT').visible := false;
  end;

  Domaine := THValComboBox(GetControl('AFF_DOMAINE')) ;
  if Domaine<>Nil then	PositionneDomaineUser(Domaine) ;


  //Formatage du code chantier
  CodeChantier := GetControlText('AFF_CHANTIER');
  ChargeCleAffaire(THEDIT(GetControl('CAF_CHANTIER0')),THEDIT(GetControl('AFF_CHANTIER1')), THEDIT(GetControl('AFF_CHANTIER2')), THEDIT(GetControl('AFF_CHANTIER3')), THEDIT(GetControl('CHA_AVENANT')), TToolbarButton97 (GetControl('BTCHANTIER')), AfTypeAction, CodeChantier, True);
//  BTChantier.left := 385;

  //
  if ((StatutAffaire ='AFF') OR (StatutAffaire = 'INT')) and (TToolbarButton97 (getControl('TCHANGETIERS')) <> nil) then
  begin
  	if TFFiche (Ecran) .typeaction = tamodif then
    begin
    	TToolbarButton97 (getControl('TCHANGETIERS')).visible := true;
  		TToolbarButton97 (getControl('TCHANGETIERS')).OnClick  := ChangeTiersAffaire;
    end;
  end;
  //
	SetActionsCotraitance;
//uniquement en line
//OnArgumentLINE;
  if Assigned(GetControl('FEUVERT')) then
  begin
    BFeuVert 		     := TToolbarButton97(ecran.FindComponent('FEUVERT'));
    BFeuVert.onclick := VoirEncours;
  end;
  if Assigned(GetControl('FEUORANGE')) then
  begin
    BFeuOrange 		     := TToolbarButton97(ecran.FindComponent('FEUORANGE'));
    BFeuOrange.onclick := VoirEncours;
  end;
  if Assigned(GetControl('FEUROUGE')) then
  begin
    BFeuRouge 		     := TToolbarButton97(ecran.FindComponent('FEUROUGE'));
    BFeuRouge.onclick := VoirEncours;
  end;

  if Assigned(GetControl('T_RISQUE')) then
  begin
    Risque := THEdit(GetControl('T_RISQUE'));
  end;

  if assigned(GetControl('BTNSOUSTRAIT')) and (VH_GC.BTCODESPECIF = '001') and (StatutAffaire = 'AFF') then
  begin
    TToolbarButton97 (GetControl('BTNSOUSTRAIT')).visible := true;
  end;


end;

Procedure TOM_AFFAIRE.OnArgumentContrat;
var SPlus		: String;
		sText : string;
begin
  //
  fResiliation  := False;
  //
  MnBordereaux.Visible := False;
  ThDBCheckBox(GetControl('AFF_MANDATAIRE')).visible := false;
  //
  MnConso.visible := True;
  StAffaire := THEdit(ecran.FindComponent('AFF_AFFAIRE'));
  StTiers   := THDBEdit(ecran.FindComponent('AFF_TIERS'));

  //modif pour gestion du code état dans le cas d'une affaire ou d'un contrat
  sPlus := sPlus + ' AND (CC_LIBRE="BTP")';
  SetControlProperty ('AFF_ETATAFFAIRE','Plus', sPlus);

  // Traduction des libéllés en fonction si proposition ou affaire
  Ecran.Caption := TraduitGa ('Gestion des contrats') ;
  UpdateCaption (Ecran) ;

  SetControlText ('TAFF_DESCRIPTIF', TraduitGa ('Descriptif du contrat') ) ;
  SetControlText ('TAFF_AFFAIRE', TraduitGa ('Contrat :') ) ;
  SetControlProperty ('AFF_AFFAIREHT', 'Caption', 'Contrat &HT') ;
  SetControlProperty ('TAFF_AFFAIRE', 'Caption', 'Contrat') ;
  SetControlProperty ('AFF_PERIODICITE', 'Plus', ' AND CO_CODE IN ("A","M","NBI")') ;

  //Gestion du type d'affaire pour les contrat
  SetControltext('TAFF_TYPEAFFAIRE', TraduitGa ('Type Contrat :')) ;
  SetControlProperty('TAFF_TYPEAFFAIRE', 'Visible', True);
  SetControlProperty('AFF_TYPEAFFAIRE',  'Visible', True);
  SetControlProperty('AFF_TYPEAFFAIRE', 'Plus', 'BTY_AFFAIRE0="I"');

  // -- MODIF LS REGROUPEMENT DE FACTURE
  //SetControlVisible ('TAFF_REGROUPEFACT1', false) ;
  //SetControlVisible ('AFF_REGROUPEFACT', false) ;
  // --
  SetControlVisible ('AFF_DETECHEANCE', True) ;
  SetControlVisible ('BINTERV', True);
  SetControlVisible ('GBTOTLIGNES', True);
  SetControlVisible ('TAFF_GENERAUTO_', False) ;
  SetControlVisible ('AFF_GENERAUTO', False) ;

  SetControlEnabled ('AFF_GENERAUTO__', False) ;

  SetControlVisible ('TAFF_REPRISEACTIV', False) ;
  SetControlVisible ('AFF_REPRISEACTIV', False) ;
  SetControlVisible ('AFF_TYPEPREVU', False) ;
  SetControlVisible ('TAFF_TYPEPREVU', False) ;
  SetControlVisible ('TAFF_PROFILGENER', false) ;
  SetControlVisible ('AFF_PROFILGENER', false) ;
  SetControlVisible ('AFF_TYPEPREVU', False) ;
  SetControlVisible ('TAFF_TYPEPREVU', False) ;
  SetControlVisible ('GROUPBOXREGROUPEMENT', False) ;
  SetcontrolVisible ('GROUPBOXDATES', True) ;
  setControlVisible ('TABSHEETFACTURATION', True) ;
  setControlVisible ('TABSHEETSTAT', True) ;
  setControlVisible ('BTACHES', True) ;
  setControlVisible ('BTCHANTIER', True) ;

  SetControlProperty ('AFF_GENERAUTO', 'Plus', 'GA" AND CO_CODE="CON') ;
  SetControlProperty ('AFF_GENERAUTO_', 'Plus', 'GA" AND CO_CODE="CON') ;

  SetControlVisible ('TAFF_DATEDEBUT', True);
  SetControlVisible ('TAFF_DATEFIN', True);
  SetControlVisible ('AFF_DATEDEBUT', True);
  SetControlVisible ('AFF_DATEFIN', True);

  THEdit(getControl('TX_FACTURE')).OnExit := TiersFactureChange;
  NbMoisEngagement := GetParamSocSecur('SO_BTDUREEENGAGEMENT',12);
  THValCOmboBox(GetCOntrol('AFF_PERIODICITE')).onchange := ChangePeriodicite;

  GereAffichagePeriodicite;
  // gestion des editions de contrat
  If THEdit(GetControl('AFF_MODELEWORD')) <> nil then
  begin
    // menu des modèles
    THDBEdit(GetControl('AFF_MODELEWORD')).OnElipsisClick := SelectModele;
    THEdit(GetControl('DOCGENERE')).OnElipsisClick := SelectDocWord;
    TmenuItem(GetControl('MnnewModele')).onclick := NewModeleClick;
    TmenuItem(GetControl('MnModifModele')).onclick := ModifModeleClick;
    TmenuItem(GetControl('MnSupModele')).onclick := SuppModeleClick;
  end;

  If THEdit(GetControl('DOCGENERE')) <> nil then
  begin
    // gestion des documents
    SetControlText('DOCGENERE','');
    // Menu des documents word
    TmenuItem(GetControl('MnGenerWord')).onclick := GenereWordClick;
    TmenuItem(GetControl('MnOuvreWord')).onclick := OuvrirWordClick;
  end;

end;

//Procedure de gestion de l'écran spécifique à Line
//uniquement en line
{*
Procedure TOM_AFFAIRE.OnArgumentLINE;
Var CC: THvalComboBox;
Begin

	//
  PlusGenere := '" AND CO_CODE <> "DAC';
  //

  //Déclarations et procédures des zones ecran
  StTiers := THDBEdit(ecran.FindComponent('AFF_TIERS'));
  Sttiers.OnElipsisClick := AppelClient;
  //
	if TTabSheet(GetControl('TABSHEETSTAT')) <> nil then TTabSheet(GetControl('TABSHEETSTAT')).TabVisible := false;
  if THLabel(GetControl('TAFF_RESPONSABLE')) <> nil then THLabel(GetControl('TAFF_RESPONSABLE')).visible := false;
  if THDBEdit(GetControl('AFF_RESPONSABLE')) <> nil then THDBEdit(GetControl('AFF_RESPONSABLE')).visible := false;
  if THLabel(GetControl('TAFF_REGROUPEFACT1')) <> nil then THLabel(GetControl('TAFF_REGROUPEFACT1')).visible := false;
  if THLabel(GetControl('TAFF_DETECHEANCE')) <> nil then THLabel(GetControl('TAFF_DETECHEANCE')).visible := false;
  if THDBValComboBox(GetControl('AFF_REGROUPEFACT')) <> nil then THDBValComboBox(GetControl('AFF_REGROUPEFACT')).visible := false;
  if TToolBarButton97(GetControl('BPOPMENU')) <> nil then TToolBarButton97(GetControl('BPOPMENU')).visible := false;
  CC := THValComboBox (GetControl ('AFF_ETABLISSEMENT') ) ; if CC <> nil then CC.visible := false;
  if THLabel(GetControl('TAFF_ETABLISSEMENT')) <> nil then THLabel(GetControl('TAFF_ETABLISSEMENT')).Visible := false;
  if THLabel(GetControl('TAFF_CONTACT')) <> nil then THLabel(GetControl('TAFF_CONTACT')).Visible := false;
  if THEdit(GetControl('LECONTACT')) <> nil then THEdit(GetControl('LECONTACT')).visible := false;

  //

End;

procedure TOM_AFFAIRE.AppelClient(Sender: TObject);
Var EditTiers : THCritMaskEdit;
begin

  StTiers.Text := GetField('AFF_TIERS');

  StTiers.Text := DispatchRecherche(EditTiers,2,'T_NATUREAUXI="CLI"','','');

  if StTiers.Text = '' then
     StTiers.setFocus
  else
     SetField('AFF_TIERS', StTiers.text);

end;
*}

procedure TOM_Affaire.CAFF_AFFAIRE1Exit (Sender: tObject) ;
begin
  PartiesAffaireExit (Sender) ;
end;

procedure TOM_Affaire.CAFF_AFFAIRE2Exit (Sender: tObject) ;
begin
  PartiesAffaireExit (Sender) ;
end;

procedure TOM_Affaire.PartiesAffaireExit (Sender: tObject) ;
var st    : string;
    Posit : integer;
begin
{$IFNDEF BTP}
  if (DS.State in [dsInsert] ) then
  begin
    St := GetParamSoc ('SO_AFALIMDESCAUTO') ;
    if (St <> '') then
    begin
      Posit := Pos ('$partie1$', lowercase (St) ) ; // $Partie1$  pour reprendre la partie1 du code affaire
      if (Posit <> 0) then
      begin
        delete (St, Posit, 9) ;
        Insert (THEdit (GetControl ('CAFF_AFFAIRE1') ) .Text, St, Posit) ;
      end;
      SetField ('AFF_DESCRIPTIF', St) ;

      Posit := Pos ('$libpartie1$', lowercase (St) ) ; // $LibPartie1$  pour reprendre le libellé de la partie 1
      if (Posit <> 0) then
      begin
        delete (St, Posit, 12) ;
        Insert (RechDom ('AFFAIREPART1', THEdit (GetControl ('CAFF_AFFAIRE1') ) .Text, false) , St, Posit) ;
      end;
      SetField ('AFF_DESCRIPTIF', St) ;

      Posit := Pos ('$partie2$', lowercase (St) ) ; // $Partie2$  pour reprendre la partie2 du code affaire
      if (Posit <> 0) then
      begin
        delete (St, Posit, 9) ;
        Insert (THEdit (GetControl ('CAFF_AFFAIRE2') ) .Text, St, Posit) ;
      end;
      SetField ('AFF_DESCRIPTIF', St) ;
    end;
  end;
{$ENDIF}
end;

procedure TOM_Affaire.OnNewRecord;
var
  CC: THValComboBox;
  DateFin : TDateTime;
  QQ : Tquery;
  NbrMois : integer;
  MonAffaire : IDeal;
  fzTiers : string;
begin
  //Inherited;
  EnSaisie := false;
  //
	BFEUROUGE.Visible   := False;
	BFEUORANGE.Visible  := False;
	BFEUVERT.Visible    := False;
  //
  EtatRisque  := '';
  //
  SetField('AFF_SSTRAITANCE','-');
  SetField('AFF_DATESSTRAIT',iDate1900);
//	THDBCheckbox (GetControl('AFF_MANDATAIRE')).State := cbGrayed; // non cotraité
  SetField('AFF_MANDATAIRE','');
  DecodeMandataire;
  SetField ('AFF_INTERVALGENER',	GetParamSocSecur('SO_AFINTERVAL',1));
  SetField ('AFF_PERIODICITE', GetParamSocSecur('SO_AfPeriodicte', 'M')) ;
  SetField ('AFF_GENERAUTO', VH_GC.AFFGenerAuto + Plusgenere) ;

  SetField ('AFF_COEFFREVALO', 1) ;
  SetField ('AFF_RECONDUCTION', VH_GC.AFFReconduction) ;

  SetField ('AFF_DATECREATION', V_PGI.DateEntree) ; // gm 07/01/03 car maintenant l'agl met l'heure

  if not (CtxScot in V_PGI.PGIContexte) then SetField ('AFF_CALCTOTHTGLO', 'X') ; // mcd 11/09/02

  SetField ('AFF_TYPEPREVU', 'GLO') ;
  SetField ('AFF_NATUREPIECEG', 'FPR') ;
  SetField ('AFF_NUMDERGENER', '') ;
  SetField ('AFF_PROFILGENER', GetParamSocSecur('SO_AFPROFILGENER', '')) ;
  SetField ('AFF_TERMEECHEANCE', GetParamSocSecur('SO_AFTERMEECHE', '')) ;
  SetField ('AFF_METHECHEANCE', GetParamSocSecur('SO_AFMETHECHE', '')) ;
  LibAff := '';
  SetField ('AFF_ADMINISTRATIF', '-') ;
  SetField ('AFF_MODELE', '-') ;
  SetField ('AFF_AFFAIREHT', 'X') ;
  SetField ('AFF_DATEDEBUT', V_PGI.DateEntree) ;

  THTypePaie.Text := '';

  if StatutAffaire = 'INT' then
  begin
    (*
    DateFin := idate2099; // par défaut
    QQ:=OpenSQL('SELECT BRE_TYPEACTION,BRE_NBMOIS FROM BRECONDUCTION WHERE BRE_CODE="'+VH_GC.AFFReconduction+'"', TRUE);
    if not QQ.eof then
    begin
      if QQ.findField('BRE_TYPEACTION').asString = '' then
      begin
        nbMois := QQ.findField ('BRE_NBMOIS').asInteger;
        DateFin := IncMonth(GetField('AFF_DATEDEBUT'), NbMois);
        DateFin := PlusDate (DateFin,-1,'J');
      end;
    end;
    ferme (QQ);
    *)
    DateFin := IncMonth(GetField('AFF_DATEDEBUT'), NbMoisEngagement);
    DateFin := PlusDate (DateFin,-1,'J');
    SetField ('AFF_DATEFIN', DateFin) ;
  end else
  begin
    SetField ('AFF_DATEFIN', idate2099) ;
  end;

  SetField ('AFF_REGSURCAF', '-') ; // gm sic 26/09/02
{$IFDEF BTP}
  SetField ('AFF_DATESIGNE', idate2099) ;
{$ELSE}
  if StatutAffaire <> 'PRO' then SetField ('AFF_DATESIGNE', V_PGI.DateEntree) ;
{$ENDIF}

  if VH_GC.AFFGenerAuto = 'CON' then
    SetField ('AFF_DATEDEBGENER', V_PGI.DateEntree)
  else
    if StatutAffaire = 'INT' then
       SetField ('AFF_DATEDEBGENER', V_PGI.DateEntree)
    else
		   SetField ('AFF_DATEDEBGENER', idate1900);

  if StatutAffaire = 'INT' then
  begin
//    SetField ('AFF_DATEFINGENER', DateFin) ;
  end else
  begin
    SetField ('AFF_DATEFINGENER', idate2099) ;
  end;

  SetField ('AFF_DATELIMITE', idate2099) ;
  SetField ('AFF_DATERESIL', idate2099) ;
  SetField ('AFF_DATECLOTTECH', idate2099) ;
  SetField ('AFF_DATEGARANTIE', idate2099) ;
  SetField ('AFF_DATECUTOFF', idate1900) ;

  SetField ('AFF_STATUTAFFAIRE', StatutAffaire) ;
  SetField ('AFF_AFFAIRE0', StatutCompletToReduit (StatutAffaire) ) ;

  
  SetField ('AFF_ETATAFFAIRE', 'ENC');

{ Mis en commentaire par BRL le 31/05/2013 suite pb chez POUCHAIN
  // Spécif BTP ENC ou ACP si passé en argument
  if EtatAffaire = '' then
    SetField ('AFF_ETATAFFAIRE', 'ENC')
  else
    SetField ('AFF_ETATAFFAIRE', EtatAffaire) ;
}

  SetField ('AFF_REGROUPEFACT', 'AUC') ;

  SetField ('AFF_CREATEUR', V_PGI.User) ;
  SetField ('AFF_UTILISATEUR', V_PGI.User) ;
  SetField ('AFF_CREERPAR', 'SAI') ; // saisie

  SetField('AFF_ETABLISSEMENT',  VH^.EtablisDefaut); // MCD 11/12/00 : mis en commentaire => FV1 06/09/13 : remise en activité

  SetField ('AFF_SAISIECONTRE', '-') ;
  SaisieContre := False;
  EchContre := False;
  Piececontre := False;
  SetField ('AFF_DEVISE', V_PGI.DevisePivot) ;
  if (VH_GC.AFFGenerAuto <> 'ACT') then
    SetField ('AFF_REPRISEACTIV', VH_GC.AFRepriseActiv)
  else
    SetField ('AFF_REPRISEACTIV', 'TOU') ;

  // controle des interventions si environnement BTP
{$IFDEF BTP}
  if StatutAffaire = 'INT' then
  begin
    SetField ('AFF_GENERAUTO', 'CON') ;
    SetField ('AFF_GENERAUTO_', 'CON') ;
    SetField ('AFF_REPRISEACTIV', 'TOU') ;
//  end;
//  if StatutAffaire = 'INT' then
//  begin
    SetField ('AFF_DETECHEANCE', 'DME') ;
    GereAffichagePeriodicite;
  end;
{$ENDIF}

  if (GetParamSoc ('SO_AFMULTIECHE') = true) then
    // PL le 29/11/02 : il faut préciser le test, car tout dépend si les conditions sur le GENERAUTO et REPRISEACTIV
    // sont respectées
  begin
    if (GetField ('AFF_GENERAUTO') <> 'ACT') and (GetField ('AFF_GENERAUTO') <> 'MAN')
      and (GetField ('AFF_REPRISEACTIV') <> 'NON') and (GetField ('AFF_REPRISEACTIV') <> '') then
    begin
      SetField ('AFF_MULTIECHE', 'X') ; // gm 08/07/02
      SetControlVisible ('AFF_MULTIECHE', True) ;
    end;
  end;

  DEV.Code := V_PGI.DevisePivot;
  GetInfosDevise (DEV) ;
  if NewTiers <> '' then
     begin
     // Risque client
     //EtatRisque :=GetEtatRisqueClient (NewTiers);
     // if EtatRisque = 'R' then begin PGIBoxAf(TexteMsgAffaire[25],'Risque client'); Newtiers :=''; end else
     // if EtatRisque = 'O' then PGIBoxAf(TexteMsgAffaire[26],'Risque client');
     if NewTiers <> '' then SetField ('AFF_TIERS', NewTiers) ;
     IncidenceTiersSurAffaire (NewTiers) ; // mcd 23/11/01 sinon, on n'a pas les info du tiers ..;
     end;

{$IFDEF BTP}
  SetField ('AFF_AFFCOMPLETE', 'X') ;
  SetField ('MONTANTGLOBAL', 0);
  SetField ('MONTANTDEJAFACT', 0);
  SetField ('MONTANTAFACT', 0);
{$ELSE}
  // Gestion des affaires complètes ou light
  if AffCompleteParDefaut then
    SetField ('AFF_AFFCOMPLETE', 'X')
  else
    SetField ('AFF_AFFCOMPLETE', '-') ;
{$ENDIF}
  //
  TToolbarButton97 (GetControl('MNCOTRAIT')).visible := false;
  if Assigned(GetControl('TS_COTRAITANCE')) then SetControlVisible('TS_COTRAITANCE', false);
  if Assigned(GetControl('AFF_REFEXTERNE')) then SetControlVisible('AFF_REFEXTERNE', True);
  if Assigned(GetControl('TAFF_REFEXTERNE')) then SetControlVisible('TAFF_REFEXTERNE', True);
  //
  if VH_GC.GCIfDefCEGID then OnNewRecordCEGID;
{$IFNDEF UTILS}
  if (StatutAffaire = 'PRO') and (Spigao) then
  begin
    if SPIGAOOBJ.ChargeAffaire(MonAffaire,FromSPIGAONew) then
    begin
      if (SPIGAOOBJ.IsNewGestion) and (MonAffaire.Contractor.Id <> 0) then
      begin
        fIdCliSPIGAo := MonAffaire.Contractor.Id;
        fzTiers := FindTiersFromIDPSIGAO (MonAffaire.Contractor.id);
        if fzTiers <> '' then
        begin
          NewTiers := fzTiers;
        	SetField ('AFF_TIERS',fzTiers);
          IncidenceTiersSurAffaire(NewTiers);
        end;
      end;
      SetField ('AFF_IDSPIGAO',MonAffaire.Id);
      if SPIGAOOBJ.IsNewGestion then
      begin
        SetField ('AFF_DESCRIPTIF',MonAffaire.Caption );
        SetField ('AFF_LIBELLE',MonAffaire.caption);
      end else
      begin
        SetField ('AFF_DESCRIPTIF',MonAffaire.Tender.Subject);
        SetField ('AFF_LIBELLE',MonAffaire.Tender.Subject);
    	end;
      SetField ('AFF_REFEXTERNE',MonAffaire.Tender.Reference);
      SetField ('AFF_DATELIMITE',FormatDateTime('yyyy/mm/dd',MonAffaire.Deadline));
      if MonAffaire.Contractor <> nil then
        THLabel (GetControl ('LIBTIERS') ) .caption := MonAffaire.Contractor.Name;
    end else
    begin
      PostMessage(Ecran.Handle, WM_CLOSE, 0, 0); //on sort de la saisie
      Exit;
    end;
  end;
{$ENDIF}
  end;

procedure TOM_Affaire.OnLoadRecord;
  function isDateModifiable(CodeAffaire : string) : boolean;
  var QQ: TQuery;
  begin
    QQ := OpenSql ('SELECT AFA_AFFAIRE FROM FACTAFF WHERE AFA_AFFAIRE="'+CodeAffaire+'" AND AFA_MONTANTECHEDEV<>0 AND AFA_ECHEFACT="X"',true,-1,'',true);
    result := QQ.eof;
    ferme (QQ);
  end;

var lng         : Integer;
    Tmp         : string;
    EtatRisque  : string;
    CleDoc      : R_CleDoc;
    dDateDebItv : Tdatetime;
    dDateFinItv : TDateTime;
    Marge       : Double;
    LaTob       : TOB;
    FFromMem    : Boolean;
    Fcledoc     : R_CLEDOC;
{$IFDEF BTP}
    Q,QLoc      : Tquery;
{$ENDIF}
		stext       : string;
    //
    // Modified by f.vautrain 03/10/2017 16:02:15 - POC
    TOBCAUTION  : TOB;
    II          : Integer;
    Req         : string;
    QQ          : TQuery;
begin

  if DuplicFromMenu then
  begin
    Duplication(AffaireToDuplic);
    //AffaireToDuplic := '';
  end;

  // Formatage du code affaire
  if (DS.State in [dsInsert] ) then
    AFTypeAction := TaCreat
  else
    AFTypeAction := TaModif;

  fStDateFin := GetField ('AFF_DATEFIN') ;

  //FV1 : 11/10/2011 -> contrôle si cotraitant ou sous traitant sur l'affaire...
  fAffaire := GetcontrolText('AFF_AFFAIRE');
	laTOB := TOB.Create('LES DONNEES',nil,-1);
 	LoadLaTOBPieceAffaireTrait (laTOB,Faffaire, 'COTRAITE');
  if (laTOB= nil) or (laTOB.detail.count = 0) then SetControlvisible('RecapCoTrait',false);
  FreeAndNil(laTob);

  //FV1 : 11/10/2011 -> contrôle si cotraitant ou sous traitant sur le document...
  laTOB := TOB.Create('LES DONNEES',nil,-1);
  StSQL := 'SELECT BDE_PIECEASSOCIEE FROM BDETETUDE WHERE BDE_AFFAIRE = "' + fAffaire + '" AND BDE_SELECTIONNE="X"';
  LaTob.LoadDetailFromSQL(StSQL);
  if (laTOB= nil) or (laTOB.detail.count = 0) then
    SetControlvisible('RecapCoTrait',false)
  else if LaTob.detail[0].GetString('BDE_PIECEASSOCIEE') = '' then
    SetControlvisible('RecapCoTrait',false)
  else
  begin
    DecodeRefPiece(LaTob.detail[0].GetString('BDE_PIECEASSOCIEE'),fCledoc);
    LoadLaTOBPieceTrait (laTOB, fCledoc, 'COTRAITE');
    if (laTOB= nil) or (laTOB.detail.count = 0) then SetControlvisible('RecapCoTrait',false);
  end;
  FreeAndNil(laTob);

  // PL le 31/05/02 : libération de la tob TheTOB
  if (TheTOB <> nil) then
  begin
    TheTob.Free;
    TheTob := nil;
  end;

  TobSSAff := nil;
  NewTiers := ''; // utilisé uniquement sur onnewrecord depuis le mul

  // Gestion du blocage affaire
  if (TFfiche (Ecran) .typeaction = TaModif) then
  begin
    if (BlocageAffaire ('MAF', GetField ('AFF_AFFAIRE') , V_PGI.groupe, V_PGI.DateEntree, true, true, false, dDateDebItv, dDateFinItv, nil) <> tbaAucun)
      or not (ModifAffaireAutorise) then
    begin
      self.ModifAutorisee := false;
      setControlEnabled ('BValider', false) ;
    end
    else
    begin
      self.ModifAutorisee := true;
      setControlEnabled ('BValider', true) ;
{$IFDEF BTP}
      SetControlVisible ('BCALCULDATES', False) ;
{$ELSE}
      SetControlVisible ('BCALCULDATES', true) ;
{$ENDIF}
    end;
  end;

  // stockage de valeurs en entrée pour test de modif
  EnSaisie := True;
  MoisCloture := 0;
  OldTiers := GetField ('AFF_TIERS') ;
  Modele := GetField ('AFF_MODELE') ;
  if GetField ('AFF_SAISIECONTRE') = 'X' then
  begin
    PgiInfoAf ('Modification impossible. En contre valeur', 'Affaire') ;
    self.ModifAutorisee := false; //mcd 20/06/03 on ne put pas modifier des missions qui sont éventuellement encore en Contrevaleur
    setControlEnabled ('BValider', false) ;
    SaisieContre := True;
    EchContre := true;
    Piececontre := True;
  end
  else
  begin
    SaisieContre := False;
    EchContre := false;
    PieceContre := False;
  end;
  FillChar (CleDocAffaire, Sizeof (CleDocAffaire) , #0) ;
  PieceChargee := False;
  NbPiece := -1;

  StatutAffaire := GetField ('AFF_STATUTAFFAIRE') ;
  EtatAffaire := GetField ('AFF_ETATAFFAIRE') ;
  FindTypeAffaire (TypeAffaire, GetField ('AFF_ISAFFAIREREF') , GetField ('AFF_AFFAIREREF') , GetField ('AFF_AFFAIRE') , GetField ('AFF_AFFCOMPLETE') ) ;

  if StatutAffaire = 'PRO' then
     begin
    LoadRecordPRO(Tmp);
  End
  else if StatutAffaire = 'INT' then
     begin
     Ecran.Caption := TraduitGa ('Contrat : ') + GetFiELD ('AFF_TIERS') + ' ' +
     BTPCodeAffaireAffiche (GetField ('AFF_AFFAIRE') , ' ') + ' ' + GetField ('AFF_LIBELLE') ;
     UpdateCaption (ecran) ;
     SetControlText ('TAFF_AFFAIRE', TraduitGa ('Contrat') ) ;
     SetControlText ('TAFF_DESCRIPTIF', TraduitGa ('Descriptif du Contrat') ) ;
     tmp := TraduitGA ('Contrat ') ;
     //
     GereAffichagePeriodicite;
     (*
     if GetField('AFF_PERIODICITE') = 'NBI' then
     begin
        SetControltext ('NBTICKET', GetField('AFF_INTERVALGENER'));
//        SetControltext ('AFF_CONSOMME', GetField('AFF_CONSOMME'));
        if Getfield('AFF_CONSOMME') <> 0 then
        begin
          SetControlEnabled('NBTICKET', false);
          SetControlEnabled('AFF_PERIODICITE', false);
        end;
     end
     else
     *)
     //Vérification si on affiche ou non le bouton de recalcul des echéance
     {Q := OpenSQL('SELECT GL_NUMLIGNE FROM LIGNE WHERE GL_NATUREPIECEG="AFF" AND GL_TYPELIGNE="ART" AND GL_AFFAIRE="' + GetField ('AFF_AFFAIRE') + '"', TRUE);
     if not Q.EOF then
        if Q.RecordCount > 1 then
           SetControlVisible ('BRECALCULMTGLOBAL', False)
        Else
           SetControlVisible ('BRECALCULMTGLOBAL', True);
     Ferme(Q);}
     //Découpage du code affaire pour affichage
     ChargeCleAffaire(THEDIT(GetControl('CAF_CHANTIER0')),THEDIT(GetControl('AFF_CHANTIER1')), THEDIT(GetControl('AFF_CHANTIER2')), THEDIT(GetControl('AFF_CHANTIER3')), THEDIT(GetControl('CHA_AVENANT')), BTChantier, AfTypeAction, GetControlText('AFF_CHANTIER'), True);
//     BTChantier.left := 385;
     //if not V_PGI.SAV then
     // begin
        IF (not isDateModifiable(GetField ('AFF_AFFAIRE'))) and (not V_PGI.SAV) then
        begin
        SetControlEnabled ('AFF_DATEDEBUT',false);
        SetControlEnabled ('AFF_DATEFIN',false);
        SetControlEnabled ('AFF_DATEDEBGENER',false);
        SetControlEnabled ('AFF_DATEFINGENER',false);
        end;
     //  end;

     end
  else
  begin
    LoadRecordAFF(Tmp);
  end;

  AlimLibEtat;
  //if (EtatAffaire<>'ENC') then TFfiche(Ecran).typeaction := taConsult;

  if ((ctxScot in V_PGI.PGIContexte) and (GetField ('AFF_AFFAIRE1') <> '') and (GetParamSoc ('SO_AFFOrmatExer') <> 'AUC') ) then
    tmp := tmp + RechDom ('AFFAIREPART1', GetField ('AFF_AFFAIRE1') , False) ;
  if (GetControl ('LIBENTETEAFFAIRE') <> nil) then
    SetControlText ('LIBENTETEAFFAIRE', tmp) ;

  if (TFfiche (Ecran) .typeaction = TaConsult) or (self.ModifAutorisee = false) then
  begin
    SetControlVisible ('BCALCULDATES', False) ; //MCD
    SetControlEnabled ('BRECALCULMTGLOBAL', False) ; // PA
  end;

  ChampsEnabledOnInsert (false) ;

  // Gestion du mémo
  DefiniLibelle;
  (*
  Memo := TMemo (GetControl ('AFF_DESCRIPTIF') ) ;
  lng := Pos (#$D#$A, Memo.lines [0] ) ;
  if (lng > 70) then
    LibAff := Copy (Memo.lines [0] , 0, 70)
  else
    LibAff := Copy (Memo.lines [0] , 0, lng - 1) ;
  *)
  //  ChargeCleAffaire(Nil,THEDIT(GetControl('CAFF_AFFAIRE1')), THEDIT(GetControl('CAFF_AFFAIRE2')),

  ChargeCleAffaire (THEDIT(GetControl('CAFF_AFFAIRE0')) , THEDIT (GetControl ('CAFF_AFFAIRE1') ) , THEDIT (GetControl ('CAFF_AFFAIRE2') ) ,
                    THEDIT(GetControl('CAFF_AFFAIRE3')) , THEDIT (GetControl ('CAFF_AVENANT') ) , nil, AFTypeAction, GetField ('AFF_AFFAIRE') , True) ;

  if (AFTypeAction <> TaCreat) then // gm 05/03/2001
  begin
    // stockages des valeurs de calcul des écheances , sauf si on est en création
    TermeEche_Sav       := GetField ('AFF_TERMEECHEANCE') ;
    DateDebGener_sav    := StrToDate (GetField ('AFF_DATEDEBGENER') ) ;
    DateFinGener_sav    := StrToDate (GetField ('AFF_DATEFINGENER') ) ;
    DateLiquidGener_sav := StrToDate (GetField ('AFF_DATEFACTLIQUID') ) ;
    Periodicite         := GetField ('AFF_PERIODICITE') ;
    TypeGenerAuto       := GetField ('AFF_GENERAUTO') ;
    Interval            := GetField ('AFF_INTERVALGENER') ;
    MontantEch          := GetField ('AFF_MONTANTECHEDEV') ;
    PourcentEch         := GetField ('AFF_POURCENTAGE') ;
    ActiviteReprise     := GetField ('AFF_REPRISEACTIV') ;
    MethEcheance        := GetField ('AFF_METHECHEANCE') ;
    MultiEche           := GetField ('AFF_MULTIECHE') ;
    DateFin_sav         := StrToDate (GetField ('AFF_DATEFIN') ) ;
    DateResil_sav       := StrToDate (GetField ('AFF_DATERESIL') ) ;
    DateResil_old       := StrToDate (GetField ('AFF_DATERESIL') ) ;
    EtatAffairePourech  := GetField ('AFF_ETATAFFAIRE') ; //gm 15/11/02
    StProfil            := GetField ('AFF_PROFILGENER') ;
    if (GetField ('AFF_FACTURE')<>'') and (StatutAffaire = 'INT') then
    begin
    	Q := OpenSql ('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+GetField('AFF_TIERS')+'" AND T_NATUREAUXI IN ("CLI","PRO")',true,1,'',true);
      if not Q.eof then theAuxiliaire := Q.findField('T_AUXILIAIRE').asstring else theAuxiliaire := '';
      ferme (Q);
      if GetField('AFF_FACTURE') = theAuxiliaire then
      begin
        OldFacture :=GetField('AFF_TIERS');   // a mettre obligatoiremeent avant setcontroltext, car setcontrol passe par onchenge de la zone
        SetCOntrolText('TX_FACTURE',GetField('AFF_TIERS')) ;
      end else
      begin
        QLoc := OpenSql('SELECT T_TIERS FROM TIERS WHERE T_AUXILIAIRE="' +
  							GetField('AFF_FACTURE') + '"', true,1,'',true);
        if not QLoc.Eof then
        begin
          OldFacture:= QLoc.Findfield('T_TIERS').AsString;
          SetCOntrolText('TX_FACTURE', OldFacture);
        end;
        Ferme(QLoc);
      end;
    end;
    //
    //
    // Modified by f.vautrain 03/10/2017 16:11:09
    if (StatutAffaire = 'AFF') AND (VH_GC.BTCODESPECIF='001') then
    begin
      //chargement de la table des retenue Garantie/Affaire
      Req := 'SELECT * FROM AFFAIRERG WHERE BAR_AFFAIRE="' + GetField ('AFF_AFFAIRE') + '" ORDER BY BAR_DATECAUTION DESC';
      QQ := OpenSQL(Req, False);
      If not QQ.Eof then
      begin
        TOBCAUTION := TOB.Create('LESAFFAIRERG', nil, -1);
        TOBCAUTION.LoadDetailDB('AFFAIRERG', '', '', QQ, False);
        FOR II := 0 To TOBCAUTION.Detail.count -1 do
        begin
          if II > 9 then Break;
          //
          SetControlText('BAR_DATECAUTION'+ IntToStr(II), TOBCAUTION.Detail[II].GetValue('BAR_DATECAUTION'));
          SetControlText('BAR_MTCAUTION'  + IntToStr(II), TOBCAUTION.Detail[II].GetValue('BAR_CAUTIONMT'));
          SetControlText('BAR_NUMCAUTION' + IntToStr(II), TOBCAUTION.Detail[II].GetValue('BAR_NUMCAUTION'));
          //
          SetControlVisible('BAR_DATECAUTION' + IntToStr(II),True);
          SetControlVisible('BAR_MTCAUTION'   + IntToStr(II),True);
          SetControlVisible('BAR_NUMCAUTION'  + IntToStr(II),True);
        end;
        If II < 9 then
        begin
          SetControlVisible('BAR_DATECAUTION' + IntToStr(II),True);
          SetControlText   ('BAR_DATECAUTION' + IntToStr(II), DateToStr(Now));
          SetControlVisible('BAR_MTCAUTION'   + IntToStr(II),True);
          SetControlVisible('BAR_NUMCAUTION'  + IntToStr(II),True);
        end;
      end
      else
      Begin
        SetControlVisible('BAR_DATECAUTION0',True);
        SetControlText('BAR_DATECAUTION0', DateToStr(Now));
        SetControlVisible('BAR_MTCAUTION0'  ,True);
        SetControlVisible('BAR_NUMCAUTION0' ,True);
      end;
      Ferme(QQ);
      FreeAndNil(TOBCAUTION);
    end;
  end else
  begin
  	// en mode création uniquement
  	CalculeCodeAffaire (THEDIT(GetControl('CAFF_AFFAIRE0')) , THEDIT (GetControl ('CAFF_AFFAIRE1') ) , THEDIT (GetControl ('CAFF_AFFAIRE2') ) ,
                    		THEDIT(GetControl('CAFF_AFFAIRE3')) , THEDIT (GetControl ('CAFF_AVENANT') ) , nil, AFTypeAction, GetField ('AFF_AFFAIRE') , True) ;
    //
    // Modified by f.vautrain 03/10/2017 16:11:09
    if (StatutAffaire = 'AFF') AND (VH_GC.BTCODESPECIF='001') then
    begin
      SetControlVisible('BAR_DATECAUTION0',True);
      SetControlText('BAR_DATECAUTION0', DateToStr(Now));
      SetControlVisible('BAR_MTCAUTION0'  ,True);
      SetControlVisible('BAR_NUMCAUTION0' ,True);
    End;
  end;

  // Chargement de la devise de l'affaire
  ImpactChangeDevise;
  GereTotDevise (True) ;

  if ctxScot in V_PGI.PGIContexte then
  begin
    if GetField ('AFF_TIERS') <> '' then
      MoisCloture := GetMoisClotureTiers (GetField ('AFF_TIERS') ) ;
    if MoisCloture <> 0 then
      SetControlText ('MOISCLOTURE', LongMonthNames [MoisCloture] ) ;
  end;

  AppliquerConfidentialite (Ecran, StatutAffaire) ; //mcd 21/05/02

  // Modif par lot
  if ModifLot then
    SetArguments (StSQL) ;

  NomContactAff; // Chargement du contact tiers
  // Chargement par défaut des regroupement d'affaires indépendantes
{$IFDEF BTP}
  SetControlProperty ('AFF_REGROUPEFACT', 'plus', ' AND ((CC_LIBRE = "A")  OR ( CC_LIBRE= "BTP" ))') ;
{$ELSE}
  SetControlProperty ('AFF_REGROUPEFACT', 'plus', ' AND ((CC_LIBRE = "A")  OR ( CC_LIBRE= "B" ))') ;
{$ENDIF}

  // Affectation en clair du dernier numero de piece
  if GetField ('AFF_NUMDERGENER') <> '' then
  begin
    DecodeRefPiece (GetField ('AFF_NUMDERGENER') , CleDoc) ;
    SetControlText ('DATEDERNIEREPIECE', DateToStr (CleDoc.DatePiece) ) ;
    SetControlText ('NUMDERNIEREPIECE', Format ('%s n° %d', [CleDoc.NaturePiece, CleDoc.NumeroPiece] ) ) ;
  end
  else
  begin
    SetControlText ('DATEDERNIEREPIECE', '') ;
    SetControlText ('NUMDERNIEREPIECE', '') ;
  end;

  // Modif d'un tiers fermé
  if (IsTiersFerme (GetField ('AFF_TIERS') ) ) and ((TFfiche (Ecran) .TypeAction = TaModif) and (self.ModifAutorisee = true) ) then
  begin
    //mcd 25/04/02 SetControlEnabled('BTLIGNE',False); // mettre com
    PgiBoxAf ('Attention client fermé', 'Saisie d''affaire') ;
    // mcd 25/04/02 , modif seulement autorisée
    self.ModifAutorisee := false;
    setControlEnabled ('BValider', false) ;
  end;

  // Risque client
  if (GetField ('AFF_TIERS') <> '') then
  begin
{$IFDEF BTP}
    Q := OPENSQL ('SELECT T_LIBELLE FROM TIERS WHERE T_TIERS="' + GetField ('AFF_TIERS') + '"', False) ;
    if not Q.EOF then
      THLabel (GetControl ('LIBTIERS') ) .caption := Q.findfield ('T_LIBELLE') .AsString // modif BTP
    else
      THLabel (GetControl ('LIBTIERS') ) .caption := '';
    Ferme (Q) ;
    //
{$ENDIF}

    //if EtatRisque = 'R' then
    //  PGIBoxAf (TexteMsgAffaire [27] , 'Risque client')
    //else
    //  if EtatRisque = 'O' then
    //  PGIBoxAf (TexteMsgAffaire [26] , 'Risque client') ;
  end;

  //Gestion d'affichage des montants
	CalculmontantEch;

  // Gestion des affaires complètes ou light
  ImpactAffCompleteOuLight;
  // Impact Affaire / sous affaires / afaires Ref
  ImpactAffRefouIndependante (False) ;

  // Affichage de la marge
  if not (ctxscot in V_PGI.PGIContexte) then
  begin
    Marge := CalculMargesurAffaire (GetField ('AFF_TOTALHT') , GetField ('AFF_TOTALPR') , GetParamsoc ('SO_AFMARGEAFFAIRE') ) ;
    if THedit (GetControl ('VALEURMARGE') ) <> nil then
      SetControlText ('VALEURMARGE', FloatTostr (Marge) ) ;
  end;

  if not (DS.State in [dsInsert] ) then DerniereCreate := '';
  // Modified by f.vautrain 29/09/2017 10:10:19
  if (StatutAffaire = 'AFF') AND (VH_GC.BTCODESPECIF='001') then
  begin
    if CoefFG         <> nil then CoefFG.text       := GetField('AFF_COEFFG');
    if CoefFS         <> nil then CoefFS.text       := GetField('AFF_COEFFS');
    if CoefSav        <> nil then CoefSAV.text      := GetField('AFF_COEFSAV');
    if CoefFD         <> nil then CoefFD.text       := GetField('AFF_COEFFD');
    if CoefMARG       <> nil then CoefMARG.text     := GetField('AFF_COEFMARG');
    if TauxRG         <> nil then TauxRG.text       := GetField('AFF_TAUXRG');
    if DateRG         <> nil then DateRG.text       := GetField('AFF_DATERG');
    if MtAvance       <> nil then MtAvance.text     := GetField('AFF_MTAVANCE');
    if DebRestAvance  <> nil then DebRestAvance.text:= GetField('AFF_DEBRESTAVANCE');
    if FinRestAvance  <> nil then FinRestAvance.text:= GetField('AFF_FINRESTAVANCE');
  end;

  SetActiveTabSheet ('PGeneral') ;

  // Liste des Intervenants
  //if not(DS.State in [dsInsert]) and
  //if GetControlVisible ('TABSHEETCOMPLEMENT')
  //  and GetControlVisible ('LISTE_AFFTIERS') then
  //begin
  //FV1 - 27/02/2017 -  FS#2409 - SES - Duplication affaire ne reprends pas toutes les info
    //GAffTiers   := THGRID (GetControl ('LISTE_AFFTIERS') ) ;
    //GAffTiers.OnDblClick := GAffTiersDblClick;
    TOBAffTiers := Tob.create ('Liste intervenant', nil, -1) ;
    ChargeAffTiers(GetField ('AFF_AFFAIRE'));
  //end;
  {
  if VH_GC.GCIfDefCEGID then
     SetControlVisible ('BCALCULDATES', False);
  }

{$IFDEF CCCS3}
  SetControlVisible ('AFF_MODELE', False) ;
{$ENDIF}
  // ONYX
  if GetParamSoc ('SO_AFREVISIONPRIX') then
  begin
    GestionMenusRevision;
    CoefAppliques;
  end;
  if EtabForce <> '' then
    THValComboBox (GetControl ('AFF_ETABLISSEMENT') ) .enabled := false;

// MODIF BRL 010606 :
// Ajout contrôle présence pièce pour l'affaire pour autorisation modif client
	if (AFTypeAction = TaModif) and (SelectPieceAffaire(GetControlText ('AFF_AFFAIRE'), StatutAffaire, CleDocAffaire) = 0) then
//uniquement en line
//if StTiers.text <> '' then StTiers.enabled := False;
  	 if StTiers <> nil then StTiers.enabled := True;
  if (StatutAffaire = 'INT') and (GetField('AFF_DETECHEANCE')='') then
  begin
    if not(DS.State in [dsInsert,dsEdit]) then DS.edit;
    Setfield('AFF_DETECHEANCE','DME') ;
  end;

  
  if assigned(GetControl('BTNSOUSTRAIT')) and (VH_GC.BTCODESPECIF = '001') and (StatutAffaire = 'AFF') then
  begin
    TToolbarButton97 (GetControl('BTNSOUSTRAIT')).OnClick := BTNSOUSTRAITClick;
  end;


  if (Pos(EtatAffaire,'TER;CLO')>0) and (not V_PGI.SAV) then
  begin
  	TFfiche (Ecran).typeaction := taconsult;
  end;

  PositionneEtabUser ('AFF_ETABLISSEMENT') ;
  if not VH^.EtablisCpta then
  begin
  	if THLabel(GetControl('TAFF_ETABLISSEMENT')) <> nil then THLabel(GetControl('TAFF_ETABLISSEMENT')).Visible := false;
			GetControl('AFF_ETABLISSEMENT').visible := false;
  end;

  if Domaine<>Nil then	PositionneDomaineUser(Domaine) ;
  
  {$IFNDEF LINE}
  DecodeMandataire;
  DefiniMenuCotraitance;
  {$ENDIF LINE}
  // Fait expres tant que la duplication dnas la fiche ne fait pas le mm chose que dans le menu
  //FV1 - 07/02/2018 : FS#2918 - SCETEC - Le bouton duplication de la fiche affaire a disparu
  //SetCOntrolVisible('BDUPLICATION',false);
  // Fait expres tant que la duplication dnas la fiche ne fait pas le mm chose que dans le menu
  //SetCOntrolVisible('BDUPLICATION',false);
  Risque.text := GetEtatRisqueClient (GetControlText('AFF_TIERS')) ;
  GestionFeuEncour;

end;

Procedure TOM_AFFAIRE.LoadRecordAFF(Var tmp : string);
begin

  Ecran.Caption := TraduitGa ('Chantier : ') + GetFiELD ('AFF_TIERS') + ' ' +
  BTPCodeAffaireAffiche (GetField ('AFF_AFFAIRE') , ' ') + ' ' + GetField ('AFF_LIBELLE') ;
  //
  SetControlText ('TAFF_DESCRIPTIF', TraduitGa ('Descriptif du Chantier') ) ;
  SetControlText ('TAFF_AFFAIRE', TraduitGa ('Chantier') ) ;

  SetControlVisible('GBSSTRAIT',True);

  tmp := TraduitGA ('Chantier ') ;

  UpdateCaption (ecran) ;

  //if (TFfiche(Ecran).typeaction = TaModif) then MnCodeBarre.Visible := True;
  if (VH_GC.BTCODESPECIF='001') then
  begin
    if (GetControlText('AFF_RESPONSABLE1')<>'') then
    begin
      SetControlproperty('LIBAFF_RESPONSABLE1', 'Visible', True);
      SetControlText('LIBAFF_RESPONSABLE1', GetLibResponsable(GetControlText('AFF_RESPONSABLE1')));
    end;
    if (GetControlText('AFF_RESSOURCE1')<>'') then
    begin
      SetControlproperty('LIBAFF_RESSOURCE1', 'Visible', True);
      SetControlText('LIBAFF_RESSOURCE1', GetLibResponsable(GetControlText('AFF_RESSOURCE1')));
    end;
    if (GetControlText('AFF_RESSOURCE2')<>'') then
    begin
      SetControlproperty('LIBAFF_RESSOURCE2', 'Visible', True);
      SetControlText('LIBAFF_RESSOURCE2', GetLibResponsable(GetControlText('AFF_RESSOURCE2')));
    end;
    if (GetControlText('AFF_RESSOURCE3')<>'') then
    begin
      SetControlproperty('LIBAFF_RESSOURCE3', 'Visible', True);
      SetControlText('LIBAFF_RESSOURCE3', GetLibResponsable(GetControlText('AFF_RESSOURCE3')));
    end;
    SetControlVisible('LIBAFF_RESPONSABLE',false);
  end else
  begin
    if (GetControlText('AFF_RESPONSABLE')<>'') then
    begin
      SetControlproperty('LIBAFF_RESPONSABLE', 'Visible', True);
      SetControlText('LIBAFF_RESPONSABLE', GetLibResponsable(GetControlText('AFF_RESPONSABLE')));
    end;
  end;

end;          

Procedure TOM_AFFAIRE.LoadRecordPRO(var tmp : string);
Var Q : TQuery;
begin
{$IFNDEF UTILS}
  if spigao then
    Ecran.Caption := TraduitGa ('Appel d''offre SPIGAO')
  else
    Ecran.Caption := TraduitGa ('Appel d''offre') ;

  Ecran.Caption := Ecran.Caption + GetFiELD ('AFF_TIERS') + ' ' +

  BTPCodeAffaireAffiche (GetField ('AFF_AFFAIRE') , ' ') + ' ' + GetField ('AFF_LIBELLE') ;

  UpdateCaption (ecran) ;

  SetControlText ('TAFF_DESCRIPTIF', TraduitGa ('Descriptif de l''appel d''offre') ) ;

  tmp := TraduitGA ('Appel d''offre ') ;
  //
  //
  TOBEtude := TOB.create ('Les Etudes', nil, -1) ; //mcd 23/01/03 chgmt nom tob
  // SELECT * : sur une seule affaire... pas de pb
  Q := opensql ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE ="' + GetField ('AFF_AFFAIRE') + '"', false) ;
  TOBEtude.SelectDB ('', Q) ;
  ferme (Q) ;

  // SELECT * : nombre de champs et d'enreg restreint
  Q := opensql ('SELECT * FROM BDETETUDE WHERE BDE_AFFAIRE="' + GetField ('AFF_AFFAIRE') + '"', false) ;
  TOBEtude.LoadDetailDB ('BDETETUDE', '', '', Q, false, true) ;
  ferme (Q) ;

  if (SPIGAOOBJ.connected) and (TFfiche (Ecran) .typeaction = TaModif) and (Etataffaire = 'ENC') then
  SetControlVisible ('BSPIGAO', True) ;

  // test sur identifiant affaire SPIGAO
  if GetField ('AFF_IDSPIGAO') <> 0 then
  begin
  Spigao := True;
  if (TFfiche (Ecran) .typeaction = TaModif) then
  begin
   MnVoirAffSpigao.Enabled := True;
   MnTelechargerDIE.Enabled := True;
  end;
  SetControlVisible ('LIBENTETEAFFAIRE', True) ;
  tmp := 'Identifiant SPIGAO : ' + FloatToStr(GetField ('AFF_IDSPIGAO'));
  SetControlProperty ('TAFF_DATELIMITE', 'Caption', 'Limite de réponse :') ;
  SetControlVisible ('TAFF_DATELIMITE', True) ;
  SetControlVisible ('AFF_DATELIMITE', True) ;
  SetControlVisible ('BTLIGNE', False) ;
  end
  else
  begin
  Spigao := False;
  tmp := TraduitGA ('Appel d''offre ') ;
  end;
{$ENDIF}
end;

procedure TOM_Affaire.ChampsEnabledOnInsert (OnUpdate: Boolean) ;
var
  okEnabled: Boolean;
  NoEnabled: Boolean;

begin

  okEnabled := (DS.State in [dsInsert] ) ;

  if OnUpdate then okEnabled := false;

  NoEnabled := not (okEnabled) ;
  SetControlEnabled ('AFF_TIERS', okEnabled) ;
//uniquement en line
//if GetControlText('AFF_TIERS') <> '' then SetControlEnabled('AFF_TIERS', NoEnabled);

  SetControlEnabled ('AFF_MODELE', okEnabled) ;
  SetCleAffaireEnabled (okEnabled) ;

  // Blocage des accès tant que la fiche n'est pas validée
  SetControlEnabled ('BUTTONECHEANCE', NoEnabled) ;
  SetControlEnabled ('BTLIGNE', NoEnabled) ;

  SetControlEnabled ('BPOPMENU', NoEnabled) ;
  SetControlEnabled ('BREGROUPE', NoEnabled) ;
  SetControlEnabled ('BTACHES', NoEnabled) ;

  if ctxScot in V_PGI.PGIContexte then
    SetControlEnabled ('BTENTETE', NoEnabled) ;

  // deplacement du bouton memo
{$IFNDEF BTP}
  if not (ctxScot in V_PGI.PGIContexte) then
    TMenuItem (GetControl ('mnMemo') ) .Enabled := NoEnabled;
{$ENDIF}

  // Bouton AffaireComplète
  SetControlEnabled ('AFF_AFFCOMPLETE', okEnabled) ;
  
end;

procedure TOM_Affaire.OnUpdateRecord;
var NbPiece     : integer;
  	EnHT        : Boolean;
    bAuto       : Boolean;
    bCalculEch  : Boolean;
    ExisteRess  : Boolean;
  	NomChamp    : string;
  	vStDateFin  : String;
    vStDateDeb  : string;
    QQ          : TQuery;
    consomme    : Integer;
    O_NBTicket  : Integer;
    NBTicket    : Integer;
    IdSpigao    : Integer;
    NomFile     : string;
    ind         : Integer;
    // Modified by f.vautrain 03/10/2017 17:00:26
    TOBCAUTION  : TOB;
    TOBCAUTIONLG: TOB;
    II          : Integer;
    Req         : string;
begin

  //uniquement en line
  //VerifieAffaireLine;

	if fModeGestion = TTgCotrait then
  begin
    if (THDBCheckbox (GetControl('AFF_MANDATAIRE')).State = cbGrayed) then
    begin
       LastError := 47;
       LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
       SetFocusControl ('AFF_MANDATAIRE') ;
       exit;
    end;
  end;

	if  (StatutAffaire = 'INT') then
  begin

    if (GetControlText ('TX_FACTURE')= '') then
    begin
      if OldFacture = '' then
      begin
        QQ := OpenSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+GetControlText('AFF_TIERS')+'"', False);
        if not QQ.eof then
        begin
          OldFacture := QQ.findField('T_AUXILIAIRE').AsString;
        end;
        ferme (QQ);
      end;
			if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
      SetField('AFF_FACTURE', oldFacture);
    end else
    begin
      QQ := OpenSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+GetControlText('TX_FACTURE')+'"', False);
      TRY
        if not QQ.Eof then
        begin
					if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
          SetField('T_FACTURE', QQ.Findfield('T_AUXILIAIRE').AsString);
        end else
        begin
          PGierror ('Le client à facturer est incorrect');
          exit ;
        end;
      FINALLY
        Ferme(QQ);
      END;
    end;

    if GetField('AFF_PERIODICITE') = 'NBI' then
    begin
			if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
      O_NBTicket  := StrToInt(GetField('AFF_INTERVALGENER'));
      NBTicket    := StrToInt(GetControltext('NBTICKET'));
      consomme    := GetField('AFF_CONSOMME');
      if (consomme > 0) and (NBTicket < consomme) then
      begin
        LastError := 47;
        LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
        SetFocusControl ('NBTICKET') ;
        SetcontrolText('NBTICKET', IntToStr(O_NBTicket));
        Exit;
      end;
      SetField ('AFF_INTERVALGENER', GetControlText('NBTICKET'))
    end else
    begin
			if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
      SetField ('AFF_INTERVALGENER', GetcontrolText('NBMOIS'));
    end;
    if GetField('AFF_INTERVALGENER')=0 then
    begin
      LastError := 46;
      LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] );
      SetFocusControl ('AFF_INTERVALGENER') ;
      exit;
    end;
    //FV - Date de résiliation modifiée !!!
    if fResiliation then
    begin
      if (GetField('AFF_DATERESIL') <> iDate2099) And (GetField('AFF_DATERESIL') <> iDate1900) then
      begin
//        SetField ('AFF_DATEFIN', GetField ('AFF_DATERESIL') ) ;
        //SetField ('AFF_DATEFINGENER', GetField ('AFF_DATERESIL')) ;
        SetField ('AFF_DATELIMITE', GetField ('AFF_DATERESIL'));
        SetField ('AFF_DATECLOTTECH', GetField ('AFF_DATERESIL'));
        //SetField ('AFF_ETATAFFAIRE', 'TER');
        SuppActionPreventive;
        fResiliation := False;
      end;
    end;
  end;

  // ONYX
  // Controle de la formule2 sans formule1
  if GetParamSoc ('SO_AFREVISIONPRIX') then
     begin
     if THEdit (GetControl ('AFF_FORCODE1') ) <> nil then
     		begin
      If (GetControlText ('AFF_FORCODE1') = '') and (GetControlText ('AFF_FORCODE2') <> '') then
      		 begin
        	 LastError := 42;
        	 LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
        	 SetFocusControl ('AFF_FORCODE1') ;
        	 exit;
      		 end;
    		end;
  	 end;

  bAuto := false;
  bCalculEch := False;

  // Dans le cas où la modification de la fiche est interdite (blocage affaire par exemple), on interdit de valider
  // toute modification : seul moyen de s'en sortir, répondre "non" à la question "voulez-vous valider les modif"
  // car c'est un cas qui ne devrait jamais arriver
  if EnSaisie then
  	 begin
     if (TFfiche (Ecran) .typeaction = TaModif) then
     	  if self.ModifAutorisee = false then
           begin
           LastError := 1;
           exit;
           end;
     end;

  // PL le 20/06/02 : generation des echeances de facturation à partir du budget
  // dans ce cas, et si aucune échéance n'a encore ete creee, on prévient l'utilisateur et on le place
  // sur l'onglet facturation pour qu'il lance le processus de generation des echeances
  // ATTENTION : on ne genere pas automatiquement les echeances car on veut que l'utilisateur positionne
  // lui même la case à cocher concernant la gestion du multi echeance AFF_MULTIECHE...
  if EnSaisie then
     if (GetParamSoc ('SO_AFALIMECHE') = true) and (TFfiche (Ecran) .typeaction = TaModif) then
        if not ExisteEcheanceAffaire (GetField ('AFF_AFFAIRE') ) then
          if (PGIAskAF ('Attention, vous n''avez pas généré les échéances de facturation.' + chr (13) + 'Vous pouvez les générer à partir de l''onglet "Facturation", bouton "Détail des échéances".' + chr (13) + 'Voulez-vous effectuer cette opération ?', Ecran.caption) = mrYes) then
             begin
             TPageControl (GetControl ('PAGES') ) .ActivePage := TTabSheet (GetControl ('TABSHEETFACTURATION') ) ;
             LastError := 1;
             exit;
             end;

  // lorsque l'on n'est pas en saisie directe indicateur EnSaisie = False
  // Attention bAV utilisé pour distinguer la reprise d'affaires / de la création auto d'avenant
  if (FTOB <> nil) then // PL le 09/07/01 champ enant rajouté à la TOB associée à la TOM_Affaire pour gérer l'avenant
     if FTOB.FieldExists ('CreationAuto') then
        if (GetField ('CreationAuto') = 'X') then
           begin
           bAuto := true;
           if FTOB.FieldExists ('CalculEcheances') then
              if (GetField ('CalculEcheances') = 'X') then
                 bCalculEch := True;
           Dev.Code := FTob.GetValue ('AFF_DEVISE') ;
           if Dev.Code <> '' then
              GetInfosDevise (DEV)
           else
              DEV.Code := V_PGI.DevisePivot;
           end;

  //************************** Traitement des erreurs blocantes ***********************////
  // ***** Gestion des erreurs sur le tiers
  if (GetField ('AFF_TIERS') = '') then
     begin
     LastError := 7;
     LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
     SetFocusControl ('AFF_TIERS') ;
     Exit;
     end;

  if EnSaisie then
  	 begin
     if (DS.State in [dsInsert] ) then // non modifiable ensuite ...
   			begin
      	if (IsTiersFerme (GetField ('AFF_TIERS') ) ) then
      		 begin
        	 LastError := 17;
        	 LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
        	 SetFocusControl ('AFF_TIERS') ;
        	 Exit;
      		 end;

      	if not (ExisteSQL ('SELECT T_TIERS From TIERS WHERE T_TIERS="' + GetField ('AFF_TIERS') + '" AND ' + FabriqueWhereNatureAuxiAff (GetField ('AFF_STATUTAFFAIRE') ) ) ) then
      		 begin
        	 LastError := 7;
        	 LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
        	 SetFocusControl ('AFF_TIERS') ;
        	 Exit;
      		 end;
    		end;

     //mcd 19/02/03 pour OK si zone tapée directement
     if (GetField ('AFF_RESSOURCE1') <> '') then
    	  begin
        ExisteRess := ExisteRessource (GetField ('AFF_RESSOURCE1') ) ;
        if not ExisteRess then
           begin
           LastError := 35;
           LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
           SetFocusControl ('AFF_RESSOURCE1') ;
           Exit;
           end;
        end;

    if (GetField ('AFF_RESSOURCE2') <> '') then
    begin
      ExisteRess := ExisteRessource (GetField ('AFF_RESSOURCE2') ) ;
      if not ExisteRess then
      begin
        LastError := 35;
        LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
        SetFocusControl ('AFF_RESSOURCE2') ;
        Exit;
      end;
    end;
    if (GetField ('AFF_RESSOURCE3') <> '') then
    begin
      ExisteRess := ExisteRessource (GetField ('AFF_RESSOURCE3') ) ;
      if not ExisteRess then
      begin
        LastError := 35;
        LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
        SetFocusControl ('AFF_RESSOURCE3') ;
        Exit;
      end;
    end;

    // ***** Test que les valeurs des lookup sont correctes
    if (VH_GC.BTCODESPECIF = '001') and (StatutAffaire ='AFF') then
    begin
      if (GetControlText ('AFF_RESPONSABLE1') <> '') then
      begin
        if not (LookupValueExist (Getcontrol ('AFF_RESPONSABLE1') ) ) then
        begin
          LastError := 21;
          LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
          SetFocusControl ('AFF_RESPONSABLE1') ;
          Exit;
        end;
      end;
    end else
    begin
      if (GetField ('AFF_RESPONSABLE') <> '') then
      begin
        if not (LookupValueExist (Getcontrol ('AFF_RESPONSABLE') ) ) then
        begin
          LastError := 21;
          LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
          SetFocusControl ('AFF_RESPONSABLE') ;
          Exit;
        end;
      end;
    end;
    //
    //
    if (GetField ('AFF_APPORTEUR') <> '') then
    Begin
      if not (LookupValueExist (Getcontrol ('AFF_APPORTEUR') ) ) then
      begin
        LastError := 22;
        LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
        SetFocusControl ('AFF_APPORTEUR') ;
        Exit;
      end;
    end;

    if (GetField ('AFF_GENERAUTO') <> '') then
    begin
      if ctxScot in V_PGI.PGIContexte then
      begin
        if (GetField ('AFF_GENERAUTO') = 'CON') then
        begin
          LastError := 28;
          LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
          SetFocusControl ('AFF_GENERAUTO') ;
          Exit;
        end;
      end;
    end;
    //--- uniquement dans le cadre des contrats / interventions !!!!
    if StatutAffaire = 'INT' then
       Begin
       if (GetField ('AFF_DETECHEANCE') <> '') then
         if not (LookupValueExist (Getcontrol ('AFF_DETECHEANCE') ) ) then
         begin
           LastError := 45;
           LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
           SetFocusControl ('AFF_DETECHEANCE') ;
           Exit;
         end;
       end;
  end;

  if EnSaisie then
  BEGIN
		if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
  END;
  // contrôle de dates
  if GetField ('AFF_AFFCOMPLETE') = '-' then
  begin
    SetField ('AFF_DATEDEBGENER', iDate1900) ;
    SetField ('AFF_DATEFINGENER', iDate2099) ;
  end;

  if (GetField ('AFF_DATEFINGENER') < GetField ('AFF_DATEDEBGENER') ) then
  begin
    LastError := 8;
    LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
    SetFocusControl ('AFF_DATEDEBGENER') ;
    Exit;
  end;
  if (GetField ('AFF_DATEFIN') < GetField ('AFF_DATEDEBUT') ) then
  begin
    LastError := 8;
    LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
    SetFocusControl ('AFF_DATEDEBGENER') ;
    Exit;
  end;

  if (GetField ('AFF_DATERESIL') <> idate1900) and (GetField ('AFF_DATERESIL') <> idate2099)
    and
    (((GetField ('AFF_DATERESIL') < GetField ('AFF_DATEDEBUT') ) and (GetField ('AFF_DATEDEBUT') <> idate1900) )
    or ((GetField ('AFF_DATERESIL') < GetField ('AFF_DATEDEBGENER') ) and (GetField ('AFF_DATEDEBGENER') <> idate1900) ) ) then
  begin
    LastError := 32;
    LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
    SetFocusControl ('AFF_DATEDEBGENER') ;
    Exit;
  end;

  if GetField ('AFF_STATUTAFFAIRE') = '' then
  begin
    LastError := 20;
    LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
    Exit;
  end;

  // ***** Gestion des erreurs sur les zones libres
  if EnSaisie then
  begin
    if GetField ('AFF_LIBREAFF1') <> '' then
      if not LookupValueExist (GetControl ('AFF_LIBREAFF1') ) then
      begin
        LastError := 18;
        LastErrorMsg := TraduitGa (THLabel (GetControl ('TAFF_LIBREAFF1') ) .Caption + ' ' + TexteMsgAffaire [LastError] ) ;
        SetFocusControl ('AFF_LIBREAFF1') ;
        Exit;
      end;
    if GetField ('AFF_LIBREAFF2') <> '' then
      if not LookupValueExist (GetControl ('AFF_LIBREAFF2') ) then
      begin
        LastError := 18;
        LastErrorMsg := TraduitGa (THLabel (GetControl ('TAFF_LIBREAFF2') ) .Caption + ' ' + TexteMsgAffaire [LastError] ) ;
        SetFocusControl ('AFF_LIBREAFF2') ;
        Exit;
      end;
    if GetField ('AFF_LIBREAFF3') <> '' then
      if not LookupValueExist (GetControl ('AFF_LIBREAFF3') ) then
      begin
        LastError := 18;
        LastErrorMsg := TraduitGa (THLabel (GetControl ('TAFF_LIBREAFF3') ) .Caption + ' ' + TexteMsgAffaire [LastError] ) ;
        SetFocusControl ('AFF_LIBREAFF3') ;
        Exit;
      end;
  end;

  // Sortie rapide si des erreurs sont rencontrées
  if Getfield ('AFF_LIBELLE') = '' then
  	 begin
     LastError := 29;
     LastErrorMsg := TexteMsgAffaire [LastError] ;
  	 end;

  if GetControlVisible('TS_COTRAITANCE')=True then
  begin
    if ThTypePaie.Text ='' then
    begin
      LastError := 46;
      LastErrorMsg := TexteMsgAffaire [LastError] ;
    end;
  end;

  if LastError <> 0 then Exit;

  // mcd 21/05/02 gestion zones obligatoire
  if (Ensaisie) then
  begin
    NomChamp := VerifierChampsObligatoires (Ecran, StatutAffaire) ;
    if NomChamp <> '' then
    begin
      if Getfield ('AFF_MODELE') <> 'X' then
      begin
        NomChamp := ReadTokenSt (NomChamp) ;
        SetFocusControl (NomChamp) ;
        LastError := 33;
        LastErrorMsg := TexteMsgAffaire [LastError] + champToLibelle (NomChamp) ;
      end
      else
      begin // mcd 17/03/03 pas blocage si modèle
        PGIBoxAF ('ATTENTION, vous avez des champs obligatoires non renseignés', 'Saisie d''affaire') ;
      end;
    end;
  end;

  if LastError <> 0 then Exit;

  //*************** Fin des erreurs blocantes ************************************
  // on peut valider le compteur d'affaire  si la fiche sera enregistrée ...
  if EnSaisie then MAJZonesCodeAffaire;
  if LastError <> 0 then Exit;

  // maj de zones en automatiques
  SetField ('AFF_UTILISATEUR', V_PGI.User) ;
  GereTotDevise (False) ; // recup totaux pièces affichés

  // Modified by f.vautrain 29/09/2017 10:10:19
  if (StatutAffaire = 'AFF') AND (VH_GC.BTCODESPECIF='001') then
  begin
    SetField('AFF_RESPONSABLE',GetControlText('AFF_RESPONSABLE1'));
    //
    if CoefFG         <> nil then SetField('AFF_COEFFG',        CoefFG.text);
    if CoefFS         <> nil then SetField('AFF_COEFFS',        CoefFS.text);
    if CoefSav        <> nil then SetField('AFF_COEFSAV',       CoefSAV.text);
    if CoefFD         <> nil then SetField('AFF_COEFFD',        CoefFD.text);
    if CoefMARG       <> nil then SetField('AFF_COEFMARG',      CoefMARG.text);
    if TauxRG         <> nil then SetField('AFF_TAUXRG',        TauxRG.text);
    if DateRG         <> nil then SetField('AFF_DATERG',        DateRG.text);
    if MtAvance       <> nil then SetField('AFF_MTAVANCE',      MtAvance.text);
    if DebRestAvance  <> nil then SetField('AFF_DEBRESTAVANCE', DebRestAvance.text);
    if FinRestAvance  <> nil then SetField('AFF_FINRESTAVANCE', FinRestAvance.text);
    //chargement de la table Retenue Garantie/affaire
    //Suppression des enregistrements de la table AffaireRG
    Req := 'DELETE AFFAIRERG WHERE BAR_AFFAIRE="' + GetField ('AFF_AFFAIRE') + '"';
    ExecuteSQL(Req);
    TOBCAUTION    := TOB.Create('LESAFFAIRERG', nil, -1);
    II := 0;
    repeat
      IF StrToFloat(GetControlText('BAR_MTCAUTION'+ IntToStr(II))) <> 0 then
      begin
        TOBCAUTIONLG  := TOB.Create('AFFAIRERG', TOBCAUTION, -1);
        TOBCAUTIONLG.PutValue('BAR_AFFAIRE',     GetField('AFF_AFFAIRE'));
        TOBCAUTIONLG.PutValue('BAR_DATECAUTION', GetControlText('BAR_DATECAUTION'+ IntToStr(II)));
        TOBCAUTIONLG.PutValue('BAR_CAUTIONMT',   GetControlText('BAR_MTCAUTION'  + IntToStr(II)));
        TOBCAUTIONLG.PutValue('BAR_NUMCAUTION',  GetControlText('BAR_NUMCAUTION' + IntToStr(II)));
        TOBCAUTIONLG.SetAllModifie(True);
  end;
      II := II + 1
    until II > 9;
    TOBCAUTION.InsertDB(Nil);
    //
    FreeAndNil(TOBCAUTION);
  end;

  // si affaire light certains champs sont alimentés en auto
  OnUpdateAffaireLight;

  // si gestion affaire ref ou sous affaire certains maj sont nécessaires
  if not (bAuto) then OnUpdateAffaireRef;

  // gestion des devises
  if (EnSaisie) and (SaisieContre <> GetContre) and not (DS.State in [dsInsert] ) then
  begin // passage monnaie pivot <=> Euro
    if ((ctxAffaire in V_PGI.PGIContexte) or (VH_GC.GASeria) ) then
    begin
      ChangePieceContre (True, GetContre) ;
      InverseEcheanceContre (GetField ('AFF_AFFAIRE') , GetContre) ;
    end;
  end;
  if (DEV.Code <> V_PGI.DevisePivot) then
    ConvertirAffaireDevise
  else
    if (EnSaisie) or (bAuto) then
  begin
    EgaliseAffaireSaisie;
    ConvertirAffaireContre;
  end;

  //
  //FV1 - 27/02/2017 -  FS#2409 - SES - Duplication affaire ne reprends pas toutes les info
  //(TFFiche(Ecran).FTypeAction = taCreat) and (DuplicFromMenu) then
  if AffaireToDuplic <> '' then
  begin
    QQ := OpenSQL ('SELECT * FROM AFFTIERS WHERE AFT_AFFAIRE="' + AffaireToDuplic + '" ORDER BY AFT_RANG', True) ;
    if not QQ.EOF then TOBAffTiers.LoadDetailDB ('AFFTIERS', '', '', QQ, False, true);
    Ferme (QQ) ;
    For Ind := 0 TO TobAffTiers.detail.count -1 do
    begin
      TobAffTiers.Detail[Ind].Putvalue('AFT_AFFAIRE', GetField ('AFF_AFFAIRE'));
      if TobAffTiers.Detail[Ind].Getvalue('AFT_TYPEINTERV') = 'CON' then
      begin
        TobAffTiers.Detail[Ind].Putvalue('AFT_TIERS',   GetField ('AFF_TIERS'));
        if theAuxiliaire = '' then
        Begin
          QQ := OpenSql ('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+GetField('AFF_TIERS')+'" AND T_NATUREAUXI IN ("CLI","PRO")',true,1,'',true);
          if not QQ.eof then theAuxiliaire := QQ.findField('T_AUXILIAIRE').asstring else theAuxiliaire := '';
          ferme (QQ);
        end;
        TobAffTiers.Detail[Ind].Putvalue('AFT_AUXILIAIRE', theAuxiliaire);
      end;
    end;
    TOBAffTiers.SetAllModifie(True);
    TOBAffTiers.InsertOrUpdateDB(True);
  end;

  //{$IFDEF CEGID}  // Spécif CEGID
  //OnUpdateCegid;     déplacé le 15/11/02 ( car pbm sur date recalculées cidessous
  // siaffaire CLO)
  //{$ENDIF}

  //Modif du 03 Février 2005

  // Mise à jour des écheances ...

  if (GetField ('AFF_GENERAUTO') = 'CON')  then
  begin
    if (EnSaisie or bAuto) and ((ctxAffaire in V_PGI.PGIContexte) or (VH_GC.GASeria) ) then
    begin
      if bauto then
      begin
        // Pas d'alimentation auto des echeances par le budget :
        // l'utilisateur doit faire la démarche volontairement
        // pour qu'il remplisse lui même la cas à cocher  AFF_MULTIECHE
        if (GetParamSoc ('SO_AFALIMECHE') = false) then
          if bCalculEch then
            CreerEcheance ('non', 'non', 'BValideAuto') ;
      end
      else
        // Pas d'alimentation auto des echeances par le budget :
        // l'utilisateur doit faire la démarche volontairement
        // pour qu'il remplisse lui même la cas à cocher  AFF_MULTIECHE
        if (GetParamSoc ('SO_AFALIMECHE') = false) then
        CreerEcheance ('non', 'oui', 'BValider') ;

      // Création de la pièce si inexistante
      if not PieceChargee then
        NbPiece := SelectPieceAffaire (GetField ('AFF_AFFAIRE') , GetField ('AFF_STATUTAFFAIRE') , CleDocAffaire)
      else
        NbPiece := 1;
      if (NbPiece = 0) then
      begin
        if GetField ('AFF_AFFAIREHT') = 'X' then
          EnHT := True
        else
          EnHT := False;
        if not (CreerPieceAffaire (GetField ('AFF_TIERS') , GetField ('AFF_AFFAIRE') , GetField ('AFF_STATUTAFFAIRE') , GetField ('AFF_ETABLISSEMENT') , CleDocAffaire, EnHT, GetContre) ) then
        begin
          LastError := 23;
          LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
          Exit;
        end;
      end;
      if ctxScot in V_PGI.PGIContexte then
      begin
        if ExisteSQL ('SELECT GPT_SOUCHE FROM PIEDPORT WHERE ' + WherePiece (CleDocAffaire, ttdPorc, False) ) then
          SetField ('AFF_CHANCELLERIE', 'X')
        else
        begin // mcd 25/06/02 ajout test sur article POU
          if ExisteSQL ('SELECT GL_SOUCHE FROM LIGNE WHERE ' + WherePiece (CleDocAffaire, ttdLigne, False) + 'AND GL_TYPEARTICLE="POU"') then
            SetField ('AFF_CHANCELLERIE', 'X')
          else
            SetField ('AFF_CHANCELLERIE', '-') ;
        end;
      end;
    end;
    if EnSaisie then
    begin
      TraiteAffTermine;
      if not (ProspectToClient) then
        PGIBoxAf (TexteMsgAffaire [30] , '') ;
    end;
  end;

  if VH_GC.GCIfDefCEGID then OnUpdateCegid; // Spécif CEGID  déplacé ici  le 14/11/02 gm

{$IFDEF EAGLCLIENT} // fermeture dès la validation sur la saisie E-AGL
    // Attention à modifier si l'eAgl ne sort pas à chaque validation de la fiche
    if (EnSaisie) then
      if (DS.State in [dsInsert] ) then
        TFFiche (ecran) .Retour := GetField ('AFF_TIERS') + ';' + GetField ('AFF_AFFAIRE1') + ';' + GetField ('AFF_AFFAIRE2') + ';' + GetField ('AFF_AFFAIRE3') ;
{$ELSE}
    // Bloquer apres validation de la création l'accés aux champs de la clé
    if (EnSaisie) and (LastError = 0) and (DS.State in [dsInsert] ) then
      ChampsEnabledOnInsert (True) ;
    // Blocage du pb de changement de fiche lié à la création

    if (EnSaisie) then
    begin
      if (DS.State in [dsInsert] ) then
      begin
        DerniereCreate := GetField ('AFF_AFFAIRE') ;
        TFFiche (ecran) .Retour := GetField('AFF_STATUTAFFAIRE')+';'+GetField ('AFF_AFFAIRE');
      end;
      if (DerniereCreate = GetField ('AFF_AFFAIRE') ) and not (DS.State in [dsInsert] ) then
      begin
        TFFiche (ecran) .Retour := GetField ('AFF_TIERS') + ';' + GetField ('AFF_AFFAIRE1') + ';' + GetField ('AFF_AFFAIRE2') + ';' + GetField ('AFF_AFFAIRE3') ;
        Ecran.Close; // le bug arrive on se casse !!!
        TFFiche (ecran) .Retour := GetField ('AFF_TIERS') + ';' + GetField ('AFF_AFFAIRE1') + ';' + GetField ('AFF_AFFAIRE2') + ';' + GetField ('AFF_AFFAIRE3') ;
      end;
    end;

    if TobSSAff <> nil then
      TobSSAff.UpdateDB;


{$ENDIF}

    {if (EnSaisie) then
       BEGIN
       // JP: demande si impression d'une lettre de mission, uniquement si création de mission
       if TFFiche(Ecran).FTypeAction = taCreat then
          if PgiAsk ('Désirez-vous éditer une lettre de mission', Ecran.Caption) = mrYes then
             ImprimeModele ('AFF', Ecran);
       end;}

    // Modif par lot : PL le 10/09/2001
    if ModifLot then TFFiche (ecran) .BFermeClick (nil) ;

    // Sur le changement de la date de fin d'affaire
    // proposition de changement des dates des taches
    vStDateFin := GetField ('AFF_DATEFIN') ;
    vStDateDeb := GetField ('AFF_DATEDEBUT') ;
    if fStDateFin <> '' then
    begin
      if (fStDateFin <> vStDateFin) and
         existeSql ('SELECT ATA_AFFAIRE FROM TACHE WHERE ATA_AFFAIRE = "' + GetField ('AFF_AFFAIRE') + '"') then
      begin
         if PGIAskAF (TexteMsgAffaire [34] , '') = mrYes then
         begin
          Majtaches(strtodate(vStDateDeb),strtodate(vStDateFin));
    //        executeSql ('UPDATE TACHE set ATA_DATEFINPERIOD = "' + UsDateTime (strtodate (vStDateFin) ) + '" WHERE ATA_AFFAIRE = "' + GetField ('AFF_AFFAIRE') + '"') ;
         end;
      end;
    end;
{$IFNDEF UTILS}
    if (Spigao) and (TFfiche (Ecran) .typeaction = TaCreat) then
    begin
      IdSpigao := GetField ('AFF_IDSPIGAO');
      SPIGAOOBJ.MajStatusAffaire(IdSpigao);
      SetControlVisible ('BSPIGAO', True) ;
      if fIdCliSPIGAo <> -1 then
      begin
        AddIDPSIGAOToTiers (GetField('AFF_TIERS'),fIdCliSPIGAo);
      end;
    end;
{$ENDIF}
end;

//uniquement en line
{*
Procedure TOM_AFFAIRE.VerifieAffaireLine;
Begin

	LastError := 0;

  // Dans le cas où la modification de la fiche est interdite (blocage affaire par exemple), on interdit de valider
  // toute modification : seul moyen de s'en sortir, répondre "non" à la question "voulez-vous valider les modif"
  // car c'est un cas qui ne devrait jamais arriver
  if EnSaisie then
  	 begin
     if (TFfiche (Ecran) .typeaction = TaModif) then
     	  if self.ModifAutorisee = false then
           begin
           LastError := 1;
           LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError]);
           end;
     end;

  //Vérification Code Tiers
  if (GetField ('AFF_TIERS') = '') then
     begin
     LastError := 7;
     LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
     SetFocusControl ('AFF_TIERS') ;
     end;

  // V érification du Libellé
  if Getfield ('AFF_LIBELLE') = '' then
  	 begin
     LastError := 29;
     LastErrorMsg := TexteMsgAffaire [LastError] ;
     SetFocusControl ('AFF_LIBELLE') ;
  	 end;

  if (GetField ('AFF_DATEFIN') < GetField ('AFF_DATEDEBUT') ) then
  begin
    LastError := 8;
    LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
    SetFocusControl('AFF_DATEDEBUT') ;
  end;

  if LastError <> 0 then Exit;

End;
*}

procedure TOM_AFFAIRE.MiseAjourMotifAnnul (Affaire : string;DateResil : Tdatetime; preventives : Boolean=true);
var QQ : TQuery;
    TOBOLES,TOBOLE : TOB;
    StSql : string;
begin
  TOBOLES := TOB.create ('LES INTER',nil,-1);
  TRY
    StSql := 'SELECT AFF_AFFAIRE FROM AFFAIRE WHERE AFF_AFFAIREINIT = "' + GetField('AFF_AFFAIRE') + '" '+
             'AND AFF_DATEFIN > "' + UsDateTime(GetField('AFF_DATERESIL'))+'" AND AFF_ETATAFFAIRE<>"ANN"';
    if Not preventives then
    begin
      StSql := StSql + ' AND AFF_CREERPAR ="SAI"';
    end else
    begin
      StSql := StSql + ' AND AFF_CREERPAR ="TAC"';
    end;
    QQ := OpenSQL(StSql,True,-1,'',true);

    while Not QQ.eof do
    begin
      TobOle := Tob.Create('LIENSOLE' ,TOBOles, -1);

      TobOle.PutValue('LO_TABLEBLOB', 'APP');
      TobOle.PutValue('LO_QUALIFIANTBLOB', 'MOT');
      TobOle.PutValue('LO_EMPLOIBLOB', 'ANN');
      TobOle.PutValue('LO_IDENTIFIANT', QQ.Fields[0].AsString);
      TobOle.PutValue('LO_RANGBLOB', 1);

      TobOle.PutValue('LO_LIBELLE', 'Résiliation du contrat ' + Affaire + ' au ' + DateToStr(DateResil));

      TobOle.PutValue('LO_PRIVE', '-');
      TobOle.PutValue('LO_DATEBLOB', DateResil);
      TobOle.PutValue('LO_OBJET', 'Résiliation du contrat au '+DateToStr(DateResil));
      QQ.next
    end;
    ferme (QQ);
    if TOBOLES.detail.count > 0 then
    begin
      TobOles.InsertDB(nil,false);
    end;
  FINALLY
    TOBOLES.free;
  END;
end;

Procedure TOM_AFFAIRE.SuppActionPreventive;
Var StsQL : String;
begin
  StSql := 'SELECT 1 FROM AFFAIRE WHERE AFF_AFFAIREINIT = "' + GetField('AFF_AFFAIRE') + '" '+
           'AND AFF_DATEFIN > "' + UsDateTime(GetField('AFF_DATERESIL')) + '" AND  '+
           'AFF_CREERPAR ="TAC" AND AFF_ETATAFFAIRE<>"ANN"';
  if ExisteSQL(StsQL) then
  begin
    if PGIAsk('Désirez-vous clôturer les interventions préventives postérieures à la date de résiliation ?') = MrYes then
    begin
    If GetField('AFF_AFFAIRE') <> '' then
    begin
        //
        MiseAjourMotifAnnul (GetField('AFF_AFFAIRE'),GetField('AFF_DATERESIL'));
        //
        StSQl := 'UPDATE AFFAIRE SET AFF_ETATAFFAIRE="ANN" '+
                 'WHERE AFF_AFFAIREINIT = "' + GetField('AFF_AFFAIRE') + '" AND '+
                 'AFF_DATEFIN > "' + UsDateTime(GetField('AFF_DATERESIL')) + '" AND '+
                 'AFF_CREERPAR ="TAC" AND AFF_ETATAFFAIRE <> "ANN"';
      if executeSQL(StsQL) <= 0 then
      begin
        V_PGI.IoError:=oeUnknown;
        Exit;
      end;
    end;
  end;
  end;

end;

procedure TOM_Affaire.OnDeleteRecord;
var
  CompWhere: string;
  CleDoc: R_CleDoc;
  statutaffaire: string;
  QQ: TQuery;
begin

  //FV1 - 10/01/2018 : FS#2778 - VEODIS GROUP - Les affaires sont supprimables alors que l'annulation est refusé dans les droits
  If not SuppAffaireAutorise then Exit;

  // Test sur les tables que l'affaire n'est pas utilisée
  // modif BRL 240107 : contrôle effectué sur les consommations
  if SupTablesLiees ('CONSOMMATIONS', 'BCO_AFFAIRE', GetField ('AFF_AFFAIRE') , '', False) then
  begin
    LastError := 12;
    LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
    Exit;
  end;
  //
  //  if (GetField ('AFF_STATUTAFFAIRE') = 'AFF') then
  //    CompWhere := 'AND  GP_NATUREPIECEG<>"' + VH_GC.AFNatAffaire + '"'
  //  else
  //    CompWhere := 'AND  GP_NATUREPIECEG<>"' + VH_GC.AFNatProposition + '"';
  //
  if (GetField ('AFF_STATUTAFFAIRE') = 'AFF') or (GetField ('AFF_STATUTAFFAIRE') = 'INT') then
  begin
    if (GetField ('AFF_STATUTAFFAIRE') = 'INT') then
      CompWhere := 'AND  GP_NATUREPIECEG<>"AFF"';
    if SupTablesLiees ('PIECE', 'GP_AFFAIRE', GetField ('AFF_AFFAIRE') , CompWhere, False) then
    begin
      LastError := 13;
      LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
      Exit;
    end;

    if (GetField ('AFF_STATUTAFFAIRE') = 'AFF') then
      CompWhere := 'AND  GL_NATUREPIECEG<>"' + VH_GC.AFNatAffaire + '"'
    else
      if (GetField ('AFF_STATUTAFFAIRE') = 'INT') then
      CompWhere := 'AND  GL_NATUREPIECEG<>"' + 'AFF' + '"'
    else
      CompWhere := 'AND  GL_NATUREPIECEG<>"' + VH_GC.AFNatProposition + '"';
    if SupTablesLiees ('LIGNE', 'GL_AFFAIRE', GetField ('AFF_AFFAIRE') , CompWhere, False) then
    begin
      LastError := 14;
      LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
      Exit;
    end;

    // C.B 09/09/2002
    // Suppression interdite si planning, mais autorisé si seulement taches et ressources (TacheRessource)
    //If SupTablesLiees ('TACHE', 'ATA_AFFAIRE', GetField('AFF_AFFAIRE'),'', False) Then
    //    Begin  LastError:=15 ; LastErrorMsg:=TraduitGa(TexteMsgAffaire[LastError]) ;   Exit; End;
    if SupTablesLiees ('AFPLANNING', 'APL_AFFAIRE', GetField ('AFF_AFFAIRE') , '', False) then
    begin
      LastError := 15;
      LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
      Exit;
    end
    else
    begin
      SupTablesLiees ('AFPLANNING', 'APL_AFFAIRE', GetField ('AFF_AFFAIRE') , '', True) ;
      SupTablesLiees ('TACHE', 'ATA_AFFAIRE', GetField ('AFF_AFFAIRE') , '', True) ;
      SupTablesLiees ('TACHERESSOURCE', 'ATR_AFFAIRE', GetField ('AFF_AFFAIRE') , '', True) ;
    end;

    // Suppression des enregistrements des tables liées
    SupTablesLiees ('FACTAFF', 'AFA_AFFAIRE', GetField ('AFF_AFFAIRE') , '', True) ;
    SupTablesLiees ('AFFTIERS', 'AFT_AFFAIRE', GetField ('AFF_AFFAIRE') , '', True) ;
    // mcd 10/02/03 on détruit aussi le mémo si existe !!!!
    SupTablesLiees ('LIENSOLE', 'LO_IDENTIFIANT', GetField ('AFF_AFFAIRE') , '', True) ;
    CompWhere := ('AND ADR_TYPEADRESSE="INT"') ;
    SupTablesLiees ('ADRESSES', 'ADR_REFCODE', GetField ('AFF_AFFAIRE') , CompWhere, True) ;
    //mcd 30/09/2003 nouvelle table !!!
    SupTablesLiees ('AFFAIREPIECE', 'API_AFFAIRE', GetField ('AFF_AFFAIRE') , '', True) ;
  end;

  // Suppression de la commande associée
  if (ctxAffaire in V_PGI.PGIContexte) or (VH_GC.GASeria) then
  begin
    statutAffaire := GetControlText ('AFF_STATUTAFFAIRE') ;
    if StatutAffaire = 'PRO' then // Gestion des appels d'offre
    begin
      NettoyageAppelOffre (TOBEtude);
      TOBEtude.free;
    end
    else if (GetField ('AFF_STATUTAFFAIRE') = 'INT') then
    begin
    	SelectPieceAffaire (GetField ('AFF_AFFAIRE') , GetField ('AFF_STATUTAFFAIRE') , CleDoc) ;
      CompWhere := 'AND  GP_NATUREPIECEG="AFF"';
    	SupTablesLiees ('PIECE', 'GP_AFFAIRE', GetField ('AFF_AFFAIRE') , CompWhere, True) ;
      CompWhere := 'AND  GL_NATUREPIECEG="AFF"';
    	SupTablesLiees ('LIGNE', 'GL_AFFAIRE', GetField ('AFF_AFFAIRE') , CompWhere, True) ;
      ExecuteSQL ('DELETE FROM PIEDBASE WHERE ' + WherePiece (CleDoc, ttdPiedBase, False) ) ;
      ExecuteSQL ('DELETE FROM PIEDECHE WHERE ' + WherePiece (CleDoc, ttdEche, False) ) ;
      ExecuteSQL ('DELETE FROM PIEDPORT WHERE ' + WherePiece (CleDoc, ttdPorc, False) ) ;
      ExecuteSQL ('DELETE FROM LIGNECOMPL WHERE ' + WherePiece (CleDoc, ttdLigneCompl, False) ) ;
      ExecuteSQL ('DELETE FROM LIGNEBASE WHERE ' + WherePiece (CleDoc, ttdLigneBase, False) ) ;
      ExecuteSQL ('DELETE FROM LIGNEOUV WHERE ' + WherePiece (CleDoc, ttdOuvrage, False) ) ;
      ExecuteSQL ('DELETE FROM LIGNEOUVPLAT WHERE ' + WherePiece (CleDoc, ttdOuvrageP, False) ) ;
      ExecuteSQL ('DELETE FROM PIECEADRESSE WHERE ' + WherePiece (CleDoc, ttdPieceAdr, False) ) ;
      ExecuteSQL ('DELETE FROM PIECETRAIT WHERE ' + WherePiece (CleDoc, ttdPieceTrait, False) ) ;
    end
(*
    // mcd 11/12/00 OK pour suippression réelle des pièce + destruction table associé
    //SelectPieceAffaire (GetField ('AFF_AFFAIRE') , GetField ('AFF_STATUTAFFAIRE') , CleDoc) ;
    //if (GetField ('AFF_STATUTAFFAIRE') = 'AFF') then
    //  CompWhere := 'AND  GP_NATUREPIECEG="' + VH_GC.AFNatAffaire + '"'
    else if (GetField ('AFF_STATUTAFFAIRE') = 'INT') then
      CompWhere := 'AND  GP_NATUREPIECEG="AFF"'
    else
      CompWhere := 'AND  GP_NATUREPIECEG="' + VH_GC.AFNatProposition + '"';

    SupTablesLiees ('PIECE', 'GP_AFFAIRE', GetField ('AFF_AFFAIRE') , CompWhere, True) ;

    if (GetField ('AFF_STATUTAFFAIRE') = 'AFF') then
      CompWhere := 'AND  GL_NATUREPIECEG="' + VH_GC.AFNatAffaire + '"'
    else
      if (GetField ('AFF_STATUTAFFAIRE') = 'INT') then
      CompWhere := 'AND  GP_NATUREPIECEG="AFF"'
    else
      CompWhere := 'AND  GL_NATUREPIECEG="' + VH_GC.AFNatProposition + '"';

    SupTablesLiees ('LIGNE', 'GL_AFFAIRE', GetField ('AFF_AFFAIRE') , CompWhere, True) ;
    //
    //
    ExecuteSQL ('DELETE FROM PIEDBASE WHERE ' + WherePiece (CleDoc, ttdPiedBase, False) ) ;
    ExecuteSQL ('DELETE FROM PIEDECHE WHERE ' + WherePiece (CleDoc, ttdEche, False) ) ;
    ExecuteSQL ('DELETE FROM PIEDPORT WHERE ' + WherePiece (CleDoc, ttdPorc, False) ) ;

    //onyx
    if GetParamSoc ('SO_AFREVISIONPRIX') then
      ExecuteSQL ('DELETE FROM AFREVISION WHERE ' + WherePiece (CleDoc, ttdRevision, False) ) ;

    if GetParamSoc ('SO_AFVARIABLES') then
      ExecuteSQL ('DELETE FROM AFORMULEVARQTE WHERE ' + WherePiece (CleDoc, ttdVariable, False) ) ;
*)      
  end;
  SupTablesLiees ('AFFAIREFRSGEST', 'BAF_AFFAIRE', GetField ('AFF_AFFAIRE') , '', True) ;
  SupTablesLiees ('AFFAIREINTERV', 'BAI_AFFAIRE', GetField ('AFF_AFFAIRE') , '', True) ;
end;

procedure TOM_Affaire.OnClose;
begin
  inherited;
  if (ctxAffaire in V_PGI.PGIContexte) or (VH_GC.GASeria) then
  begin
    if (SaisieContre <> EchContre) then
      InverseEcheanceContre (GetField ('AFF_AFFAIRE') , SaisieContre) ;
    if (SaisieContre <> PieceContre) then
      ChangePieceContre (True, SaisieContre) ;
    if TobPiece <> nil then
    begin
      TobPiece.Free;
      TobPiece := nil;
    end;
  end;

  // PL le 31/05/02 : libération de la tob TheTOB
  if (TheTOB <> nil) then
  begin
    TheTob.Free;
    TheTob := nil;
  end;
  if (TOBAffTiers <> nil) then
    TOBAffTiers.Free;
  //ClearTobAffaireMem;
end;

procedure TOM_Affaire.OnCancelRecord;
begin
  inherited;
  // on a mis a jour les écheances ou les lignes alors que l'utilisateur annule les modifs de fiches
  if (ctxAffaire in V_PGI.PGIContexte) or (VH_GC.GASeria) then
  begin
    if (SaisieContre <> EchContre) then
      InverseEcheanceContre (GetField ('AFF_AFFAIRE') , SaisieContre) ;
    if (SaisieContre <> PieceContre) then
      ChangePieceContre (True, SaisieContre) ;
  end;
end;

procedure TOM_Affaire.SetCleAffaireEnabled (bEnabled: Boolean) ;
begin

  SetControlEnabled ('CAFF_AFFAIRE1', bEnabled) ;
  SetControlEnabled ('CAFF_AFFAIRE2', bEnabled) ;
  SetControlEnabled ('CAFF_AFFAIRE3', bEnabled) ;
  SetControlEnabled ('CAFF_AVENANT', bEnabled) ;

  SetControlVisible ('BDUPLICATION', Not bEnabled);

//uniquement en line
//if GetControlText('AFF_TIERS') <> '' then
//   SetControlEnabled('AFF_TIERS', false)
//Else
//   SetControlEnabled ('AFF_TIERS', bEnabled) ;

end;

//******************************************************************************
//********************** Calcul des echéances **********************************
// La fonction retourne 1 si on peut accéder au factaff et 2 dans le cas contraire
//******************************************************************************

procedure TOM_Affaire.CreerEcheanceDuBudget;
begin
  UtilCalculEcheancesBudget (GetField ('AFF_PROFILGENER') , GetField ('AFF_AFFAIRE') , GetField ('AFF_REPRISEACTIV') , GetField ('AFF_GENERAUTO') , DEV, GetContre, GetField ('AFF_MULTIECHE') ) ;
end;

function TOM_Affaire.LanceChoixRecalcul (bPourcentage: Boolean) : string;
var
  dMontEch, dMontTot: double;
  param: string;
begin
  if bPourcentage then
  begin
    if GetParamSoc ('SO_AFAUTOECHEANCE') = True then
    begin
      Result := '2';
      Exit;
    end;
    if (PGIAskAF ('Souhaitez-vous recalculer les échéances', '') = mrYes) then
      Result := '2';
  end
  else
  begin
    // si recalcul auto force sur le cas 1) génération sur la base du montant périodique
    if GetParamSoc ('SO_AFAUTOECHEANCE') = True then
    begin
      Result := '1';
      Exit;
    end;
    (*
    // Fenêtre de choix de la méthode de recalcul
    param := 'DEVISE:' + Dev.Code;
    dMontEch := GetField ('AFF_MONTANTECHEDEV') ;
    param := param + ';MTPERIODIQUE:' + Format ('%n', [dMontEch] ) ;
    dMontTot := GetField ('AFF_TOTALHTGLODEV') ;
    param := param + ';MTGLOBAL:' + Format ('%n', [dMontTot] ) ;
    param := param + ';AFF_GENERAUTO:' + GetField ('AFF_GENERAUTO') ;
    Result := AFLanceFiche_ChoixEcheAff (param) ;
    *)
    if (PGIAskAF ('Souhaitez-vous recalculer les échéances', '') = mrYes) then
      Result := '1';
  end;
end;

procedure TOM_Affaire.LanceEcheance (ForceCalcul, Parle, Boutton: string) ;
var
  ActionAffaire: string;
begin

  ActionAffaire := ActionAffaireString (self.ModifAutorisee, TFFiche (Ecran) .typeaction) ;

  // PL le 28/11/2001 : ne pas retirer cette ligne qui permet de ne pas entrer dans les lignes de FACTAFF
  // en CREATION mais en modification car une échéance supplémentaire apparait au 01/01/1900
  if (TFfiche (Ecran) .typeaction <> TaConsult) then //mcd 17/06/02
    if self.ModifAutorisee then
      ActionAffaire := 'ACTION=MODIFICATION'; // sinon pb en entrée dans les factaff en création d'affaire PA

  //
  CreerEcheance (ForceCalcul, Parle, Boutton) ;

  //If (Rep = 1) Then // on permet de rentrer dans tous les cas sur les écheances PA le 13/12/2000
  // pas mis en fct car la tom associé est dans ce source ...
    //mcd 15/05/03 modif appel fiche pour Ok eagl
  //{$IFDEF BTP}
  //AGLLancefiche ('AFF','FACTAFF','NOR;'+GetField('AFF_AFFAIRE'),'',
  //              ActionAffaire+';AFF_GENERAUTO:'+GetField('AFF_GENERAUTO')+';AFF_DEVISE:'+
  //              GetField('AFF_DEVISE')+';AFF_SAISIECONTRE:'+GetField('AFF_SAISIECONTRE')+';AFF_TIERS:'+
  //              GetField('AFF_TIERS')+';AFF_REPRISEACTIV:'+GetField('AFF_REPRISEACTIV')+
  //              ';DATEDEBFAC:'+ DateToStr(GetField('AFF_DATEDEBGENER'))+   //mcd 15/05/03
  //              ';DATEFINFAC:'+ DateToStr(GetField('AFF_DATEFINGENER'))+   //mcd 15/05/03
  //              ';AFF_PROFILGENER:'+ GetField('AFF_PROFILGENER')+   //mcd 03/07/03
  //              ';AFF_DATECREATION:'+ DateToStr(GetField('AFF_DATECREATION'))+ //btp 08/06/04
  //              ';AFF_STATUTAFFAIRE:' + GetControlText ('AFF_STATUTAFFAIRE'));
  //{$ELSE}

  AGLLancefiche ('AFF', 'FACTAFF', 'NOR;' + GetField ('AFF_AFFAIRE') , '',
    ActionAffaire + ';AFF_GENERAUTO:' + GetField ('AFF_GENERAUTO') + ';AFF_DEVISE:' +
    GetField ('AFF_DEVISE') + ';AFF_SAISIECONTRE:' + GetField ('AFF_SAISIECONTRE') + ';AFF_TIERS:' +
    GetField ('AFF_TIERS') + ';AFF_REPRISEACTIV:' + GetField ('AFF_REPRISEACTIV') +
    // correction FQ 12343
    ';AFF_STATUTAFFAIRE:'+getField('AFF_STATUTAFFAIRE') +
    // --
    ';DATEDEBFAC:' + DateToStr (GetField ('AFF_DATEDEBGENER') ) + //mcd 15/05/03
    ';DATEFINFAC:' + DateToStr (GetField ('AFF_DATEFINGENER') ) + //mcd 15/05/03
    ';AFF_PROFILGENER:' + GetField ('AFF_PROFILGENER') + //mcd 03/07/03
    ';AFF_PERIODICITE:'+ GetField ('AFF_PERIODICITE') +
    ';AFF_INTERVALGENER:'+ inttostr(GetField ('AFF_INTERVALGENER')) +
    ';AFF_DATECREATION:' + DateToStr (GetField ('AFF_DATECREATION') )+
                ';AFF_METHECHEANCE:'+ GetField ('AFF_METHECHEANCE')+
                ';AFF_TERMEECHEANCE:'+ getField ('AFF_TERMEECHEANCE')) ;
  //{$ENDIF}

  ImpactCalculTotalHT (TitEcheances) ;

  CalculMontantEch;

end;

function TOM_Affaire.TestTypeModifEcheances (var DateDebut, DateFin, DateLiquid, DateResil: TDateTime; var TypeCalcul: string; var zmont: double; var TypeModif: T_TypeModifAff; Boutton: string; bSilence: boolean) : integer;
var NomChamps : string;
begin
  NoMChamps := 'AFF_PERIODICITE;AFF_MULTIECHE;AFF_GENERAUTO;AFF_INTERVALGENER;AFF_TERMEECHEANCE;AFF_DATEDEBGENER'+
               'AFF_DATEFINGENER;AFF_DATERESIL;AFF_PROFILGENER;AFF_PERIODICITE';
  TypeModif := [] ;
  Result := 1;
  zmont := 0;
  TypeCalcul := '';
  DateDebut := iDate1900;
  DateFin := iDate2099;
  DateLiquid := iDate1900;
  DateResil := iDate1900;
  if (Assigned(Ecran.ActiveControl)) and (Pos(Ecran.ActiveControl.Name,NomChamps) > 0) then
  begin
    SendMessage(Ecran.ActiveControl.Handle,WM_KEYDOWN,VK_TAB,0);
  end;

  // Attention, si modif voir fct ModifAvecTrait de UtofAfAffaire_ModifLot
  // Test si changement d'échéances
  if (ActiviteReprise <> GetField ('AFF_REPRISEACTIV') ) then TypeModif := TypeModif + [tmaRepAct] ;

  if ((Arrondi(MontantEch,V_PGI.okdecV) <> Arrondi(double (GetField ('AFF_MONTANTECHEDEV')),V_PGI.okdecV)) and
     (GetField ('AFF_GENERAUTO') <> 'POU') and
     (GetField ('AFF_GENERAUTO') <> 'POT')) then
     TypeModif := TypeModif + [tmaMntEch] ;

  if ((PourcentEch <> double (GetField ('AFF_POURCENTAGE'))) and
     ((GetField ('AFF_GENERAUTO') = 'POU') or
     (GetField ('AFF_GENERAUTO')  = 'POT'))) then
     TypeModif := TypeModif + [tmaPourcEch] ;

  if ((Periodicite      <> string (GetField ('AFF_PERIODICITE'))) or
      (MultiEche        <> string (GetField ('AFF_MULTIECHE'))) or
      (TypeGenerAuto    <> string (GetField ('AFF_GENERAUTO'))) or
      (Interval         <> Integer (GetField ('AFF_INTERVALGENER'))) or
      (TermeEche_sav    <> getField('AFF_TERMEECHEANCE')) or
      (DateDebGener_sav <> StrToDate (GetField ('AFF_DATEDEBGENER'))) or
      (DateFinGener_sav <> StrToDate (GetField ('AFF_DATEFINGENER'))) or
      (DateResil_sav    <> StrToDate(GetField ('AFF_DATERESIL'))) or
      (MethEcheance     <> string (GetField ('AFF_METHECHEANCE'))) or
      (StProfil         <> string (GetField('AFF_PROFILGENER')))) Then
      TypeModif := TypeModif + [tmaDate];

  // gm 15/11/02 , pour relancer le calcul des echéances si changement d'état
  // sauf si CLO car il y a un traitement particulier qui supprime les échéances...
  if (EtatAffairePourEch <> GetField ('AFF_ETATAFFAIRE') ) and (GetField ('AFF_ETATAFFAIRE') <> 'CLO') then
    TypeModif := TypeModif + [tmaDate] ;

  if (DateLiquidGener_sav <> StrToDate (GetField ('AFF_DATEFACTLIQUID') ) ) then
    TypeModif := TypeModif + [tmaDFacLiq] ;

  TypeCalcul  := GetField ('AFF_PERIODICITE') ;

  DateDebut   := StrToDate (GetField ('AFF_DATEDEBGENER') ) ;
  DateFin     := StrToDate (GetField ('AFF_DATEFINGENER') ) ;
  if (GetField ('AFF_GENERAUTO') = 'CON') and (GetField ('AFF_METHECHEANCE') = 'CIV') then
  begin
    // En mode de génération Contrat et méthode d'échéance civile, on place la date de début au début de la période
    // suivant la périodicité : Annuelle, mensuelle, hebdomadaire
    DateDebut := GetDateDebutPeriode (TypeCalcul, DateDebut, Integer (GetField ('AFF_INTERVALGENER') ),GetField ('AFF_METHECHEANCE') ) ;
  end else if (GetField ('AFF_GENERAUTO') = 'CON') and (GetField ('AFF_METHECHEANCE') = 'AN') then
  begin
    DateDebut   := StrToDate (GetField ('AFF_DATEDEBUT') ) ;
    DateFin     := StrToDate (GetField ('AFF_DATEFIN') ) ;
  end;

  DateLiquid  := StrToDate (GetField ('AFF_DATEFACTLIQUID') ) ;
  DateResil   := StrToDate (GetField ('AFF_DATERESIL') ) ;

  if ((GetField ('AFF_GENERAUTO') = 'POU') or (GetField ('AFF_GENERAUTO') = 'POT') ) then
    zmont := double (GetField ('AFF_POURCENTAGE') )
  else
    zmont := double(GetField ('AFF_MONTANTECHEDEV'));

  // PLGM le 26/06/02 : pb de definition de la date de fin de generation des echeances de facturation
  //If ((DateDebut = idate1900) and  (DateFin = idate2099))  Then TypeModif := [];

  //If (DateFin = idate2099) then
  if (DateFin >= GetParamSoc ('SO_AFDATEFINGENER') ) then
  begin
    if GetParamSoc ('SO_AFDATEFINGENER') <> iDate2099 then
      DateFin := GetParamSoc ('SO_AFDATEFINGENER')
    else
      DateFin := PlusDate (DateDebut, 1, 'A') ;
  end;

  if ((DateDebut = idate1900) or (DateFin = idate2099) ) then
    TypeModif := [] ;
  // fin PLGM le 26/06/02

  if (DateDebut = idate1900) then
  begin
    if GetField ('AFF_STATUTAFFAIRE') = 'PRO' then
      bSilence := True;
    if (ctxTempo in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) then
      bSilence := True;

    if not bSilence then
      if ((Boutton <> 'BValider') and (Boutton <> 'BValideAuto') ) or
        ((Boutton = 'BValider') and (GetField ('AFF_GENERAUTO') <> 'MAN') ) then
        if (GetField ('AFF_ADMINISTRATIF') = '-') and (GetField ('AFF_MODELE') = '-') then
        begin
          ShowMessage ('Attention : aucune échéance de facturation générée en automatique') ;
          SetFocusControl ('AFF_DATEDEBGENER') ;
        end;
    Result := 2;
  end;
end;

procedure TOM_Affaire.MAJZonesCodeAffaire;
var {St,}
  tmp: string;    
  CC: THEdit;
  {Posit,} iPartErreur: integer;
  bProposition: Boolean;
begin

  //if (DS.State in [dsInsert] ) then
  //  AFTypeAction := TaCreat
  //else
  //  AFTypeAction := TaModif;

  bProposition := (GetField ('AFF_STATUTAFFAIRE') = 'PRO') ;
{$IFDEF BTP}
  if GetField ('AFF_STATUTAFFAIRE') = 'INT' then
    tmp := DechargeCleAffaire (THEDIT (GetControl ('AFF_AFFAIRE0')) , THEDIT (GetControl ('CAFF_AFFAIRE1') ) , THEDIT (GetControl ('CAFF_AFFAIRE2') ) ,
           THEDIT (GetControl ('CAFF_AFFAIRE3') ) , THEDIT (GetControl ('CAFF_AVENANT') ) , GetControlText ('AFF_TIERS') , AFTypeAction, True, True, bProposition, iPartErreur)
{$ENDIF}
  else
    //tmp := DechargeCleAffaire (nil, THEDIT (GetControl ('CAFF_AFFAIRE1') ) , THEDIT (GetControl ('CAFF_AFFAIRE2') ) ,
    tmp := DechargeCleAffaire (THEDIT (GetControl ('AFF_AFFAIRE0')), THEDIT (GetControl ('CAFF_AFFAIRE1') ) , THEDIT (GetControl ('CAFF_AFFAIRE2') ) ,
           THEDIT (GetControl ('CAFF_AFFAIRE3') ) , THEDIT (GetControl ('CAFF_AVENANT') ) , GetControlText ('AFF_TIERS') , AFTypeAction, True, True, bProposition, iPartErreur) ;

  if (tmp <> '') then
  begin
    SetField ('AFF_AFFAIRE', tmp) ;
    CC := THEdit (GetControl ('CAFF_AFFAIRE1') ) ;
    SetField ('AFF_AFFAIRE1', CC.Text) ;
    CC := THEdit (GetControl ('CAFF_AFFAIRE2') ) ;
    SetField ('AFF_AFFAIRE2', CC.Text) ;
    CC := THEdit (GetControl ('CAFF_AFFAIRE3') ) ;
    SetField ('AFF_AFFAIRE3', CC.Text) ;
    CC := THEdit (GetControl ('CAFF_AVENANT') ) ;
    SetField ('AFF_AVENANT', CC.Text) ;
    if (GetField ('AFF_DESCRIPTIF') = '') then
    begin
      PartiesAffaireExit (nil) ;
    end;
  end
  else
  begin
    LastError := 6;
    LastErrorMsg := TraduitGa (TexteMsgAffaire [LastError] ) ;
    SetFocusControl ('CAFF_AFFAIRE1') ;
    Exit;
  end;
end;

// *****************************************************************************
// ***************** Gestion devises + Euro  sur les affaires ******************
// *****************************************************************************

procedure TOM_Affaire.ImpactChangeDevise;
var
  CompNum: THNumedit;
  DateDevise: TDateTime;
begin
  DEV.Code := GetField ('AFF_DEVISE') ;
  if DEV.Code = V_PGI.DevisePivot then
  begin
    if GetContre then
    begin
      if VH^.TenueEuro then
      begin
        DEV.Libelle := RechDom ('TTDEVISETOUTES', V_PGI.DeviseFongible, False) ;
        DEV.Decimale := V_PGI.OkDecE;
      end
      else
      begin
        DEV.Decimale := V_PGI.OkDecE;
        DEV.Symbole := 'E';
        DEV.Libelle := 'Euro';
      end;
    end
    else
    begin
      DEV.Decimale := V_PGI.OkDecV;
      DEV.Symbole := V_PGI.SymbolePivot;
      DEV.Libelle := RechDom ('TTDEVISETOUTES', V_PGI.DevisePivot, False) ;
    end;
  end
  else // En devise
  begin
    GetInfosDevise (DEV) ;
    if (DS.State in [dsInsert] ) then
      DateDevise := V_PGI.DateEntree
    else
      DateDevise := GetField ('AFF_DATECREATION') ;
    DEV.Taux := GetTaux (DEV.Code, DEV.DateTaux, DateDevise) ;
  end;
  CompNum := THNumEdit (GetControl ('LAFF_TOTALHTDEV') ) ;
  if CompNum <> nil then
    ChangeMask (CompNum, DEV.Decimale, DEV.Symbole) ;
  CompNum := THNumEdit (GetControl ('LAFF_TOTALTAXEDEV') ) ;
  if CompNum <> nil then
    ChangeMask (CompNum, DEV.Decimale, DEV.Symbole) ;
  CompNum := THNumEdit (GetControl ('LAFF_TOTALTTCDEV') ) ;
  if CompNum <> nil then
    ChangeMask (CompNum, DEV.Decimale, DEV.Symbole) ;

  SetControlText ('LIBDEVISE', 'En ' + DEV.Libelle) ;
{$IFNDEF BTP}
  AfficheSigleEuro;
{$ENDIF}
end;

procedure TOM_Affaire.AfficheSigleEuro;
var
  Aff: boolean;
begin
  Aff := False;
  if (Getfield ('AFF_DEVISE') <> 'EUR') and (Getfield ('AFF_DEVISE') <> 'FRF') then
    aff := false // mcd 17/04/02 n,e pas afficher EURO si devise pas EURO
  else
  begin
    if GetContre and not (VH^.TenueEuro) then
      Aff := true;
    if (not (GetContre) ) and VH^.TenueEuro then
      Aff := true;
  end;
  if Aff then
    SetControlVisible ('SIGLEEURO', true)
  else
    SetControlvisible ('SIGLEEURO', false) ;
end;

function TOM_Affaire.GetContre: Boolean;
begin
  result := (GetField ('AFF_SAISIECONTRE') = 'X') ;
end;

procedure TOM_Affaire.ConvertirAffaireDevise;
var
  X: double;
begin
  // maj des zones en monnaie pivot
  if (DEV.Code <> V_PGI.DevisePivot) then
  begin
    //if VH^.TenueEuro then BEGIN SufF:='CON' ; SufE:='' ; END else BEGIN SufF:='' ; SufE:='CON' ; END ;
    X := GetField ('AFF_TOTALHTDEV') ;
    //SetField('AFF_TOTALHT'+SufF,DeviseToFranc(X,DEV.Taux, DEV.Quotite)) ;
    SetField ('AFF_TOTALHT', DeviseToEuro (X, DEV.Taux, DEV.Quotite) ) ;
    X := GetField ('AFF_TOTALTTCDEV') ;
    //SetField('AFF_TOTALTTC'+SufF,DeviseToFranc(X,DEV.Taux, DEV.Quotite)) ;
    SetField ('AFF_TOTALTTC', DeviseToEuro (X, DEV.Taux, DEV.Quotite) ) ;
    X := GetField ('AFF_TOTALTAXEDEV') ;
    //SetField('AFF_TOTALTAXE'+SufF,DeviseToFranc(X,DEV.Taux, DEV.Quotite)) ;
    SetField ('AFF_TOTALTAXE', DeviseToEuro (X, DEV.Taux, DEV.Quotite) ) ;
    X := GetField ('AFF_TOTALHTGLODEV') ;
    //SetField('AFF_TOTALHTGLO'+SufF,DeviseToFranc(X,DEV.Taux, DEV.Quotite)) ;
    SetField ('AFF_TOTALHTGLO', DeviseToEuro (X, DEV.Taux, DEV.Quotite) ) ;
    X := GetField ('AFF_MONTANTECHEDEV') ;
    //SetField('AFF_MONTANTECHE'+SufF,DeviseToFranc(X,DEV.Taux, DEV.Quotite)) ;
    SetField ('AFF_MONTANTECHE', DeviseToEuro (X, DEV.Taux, DEV.Quotite) ) ;
  end;
end;

procedure TOM_Affaire.ConvertirAffaireContre;
var
  X: double;
  suf: string;
  ModeOppose: Boolean;
begin
  ModeOppose := (GetField ('AFF_SAISIECONTRE') = 'X') ;
  //mcd 13/02/03 if ModeOppose then Suf:='' else Suf:='CON' ;
  if ModeOppose then
  begin
    Suf := '';
    X := GetField ('AFF_TOTALHTDEV') ;
    SetField ('AFF_TOTALHT' + Suf, ConvertSaisieEF (X, ModeOppose) ) ;
    X := GetField ('AFF_TOTALTAXEDEV') ;
    SetField ('AFF_TOTALTAXE' + Suf, ConvertSaisieEF (X, ModeOppose) ) ;
    X := GetField ('AFF_TOTALTTCDEV') ;
    SetField ('AFF_TOTALTTC' + Suf, ConvertSaisieEF (X, ModeOppose) ) ;
    X := GetField ('AFF_TOTALHTGLODEV') ;
    SetField ('AFF_TOTALHTGLO' + Suf, ConvertSaisieEF (X, ModeOppose) ) ;
    X := GetField ('AFF_MONTANTECHEDEV') ;
    SetField ('AFF_MONTANTECHE' + Suf, ConvertSaisieEF (X, ModeOppose) ) ;
  end;
end;

procedure TOM_Affaire.Egal2Champs (CS, CD: string) ;
begin
  SetField (CD, GetField (CS) ) ;
end;

(*Procedure TOM_Affaire.Inverse2Champs (CS,CD : String ) ;
Var XX : variant;
BEGIN
XX := GetField(CD);
SetField(CD,GetField(CS));
SetField(CS,XX);
END ; *)

procedure TOM_Affaire.EgaliseAffaireSaisie;
var
  Suf: string;
begin
  //mcd 13/02/03if GetContre then Suf:='CON' else Suf:='' ;
  if not (GetContre) then
  begin
    Suf := '';
    Egal2Champs ('AFF_TOTALHTDEV', 'AFF_TOTALHT' + Suf) ;
    Egal2Champs ('AFF_TOTALTAXEDEV', 'AFF_TOTALTAXE' + Suf) ;
    Egal2Champs ('AFF_TOTALTTCDEV', 'AFF_TOTALTTC' + Suf) ;
    Egal2Champs ('AFF_TOTALHTGLODEV', 'AFF_TOTALHTGLO' + Suf) ;
    Egal2Champs ('AFF_MONTANTECHEDEV', 'AFF_MONTANTECHE' + Suf) ;
  end;
end;

procedure TOM_Affaire.ChangePieceContre (Libere, bContre: Boolean) ;
var
  NbPiece, i: integer;
  st: string;
  Tobdet: tob;
begin
  if not PieceChargee then
    NbPiece := SelectPieceAffaire (GetField ('AFF_AFFAIRE') , GetField ('AFF_STATUTAFFAIRE') , CleDocAffaire)
  else
    NbPiece := 1;
  if (NbPiece <> 1) then
    Exit;
  LoadTobPieceAffaire (CleDocAffaire, true) ;
  if bContre then
    st := 'X'
  else
    st := '-';
  TobPiece.PutValue ('GP_SAISIECONTRE', st) ;
  for i := 0 to TobPiece.Detail.count - 1 do
  begin
    Tobdet := TobPiece.Detail [i] ;
    TobDet.PutValue ('GL_SAISIECONTRE', st) ;
  end;
  TobPiece.UpdateDB;
  if (libere) and (TobPiece <> nil) then
  begin
    TobPiece.Free;
    TobPiece := nil;
    PieceChargee := False;
  end;
end;

// ******************************************************************************
// Interface entre les zones non DB affichées (ThNumEdit) et les champs de la base
// ******************************************************************************

procedure TOM_Affaire.GereTotDevise (Entree: Boolean) ;
var
  CompNum: THNumedit;
begin
  if not (EnSaisie) then
    Exit;
  if Entree then
  begin // maj des THNumEdit en affichage / zones de la base
    CompNum := THNumEdit (GetControl ('LAFF_TOTALHTDEV') ) ;
    if CompNum <> nil then
      CompNum.value := GetField ('AFF_TOTALHTDEV') ;
    CompNum := THNumEdit (GetControl ('LAFF_TOTALTAXEDEV') ) ;
    if CompNum <> nil then
      CompNum.value := GetField ('AFF_TOTALTAXEDEV') ;
    CompNum := THNumEdit (GetControl ('LAFF_TOTALTTCDEV') ) ;
    if CompNum <> nil then
      CompNum.value := GetField ('AFF_TOTALTTCDEV') ;
  end
  else
  begin // maj des zones de la base / THNumEdit en affichage
    CompNum := THNumEdit (GetControl ('LAFF_TOTALHTDEV') ) ;
    if CompNum <> nil then
      SetField ('AFF_TOTALHTDEV', CompNum.value) ;
    CompNum := THNumEdit (GetControl ('LAFF_TOTALTAXEDEV') ) ;
    if CompNum <> nil then
      SetField ('AFF_TOTALTAXEDEV', CompNum.value) ;
    CompNum := THNumEdit (GetControl ('LAFF_TOTALTTCDEV') ) ;
    if CompNum <> nil then
      SetField ('AFF_TOTALTTCDEV', CompNum.value) ;
  end;
end;

// ****************************************************************************
// ********************** Gestion des pièces sur les affaires *****************
// ******************************************************************************

procedure TOM_Affaire.AppelPieceAffaire (TypeAppel: string) ;
var
  EnHT: Boolean;
  HTBefore, TaxeBefore, PRBefore: Double;
begin

  if GetField ('AFF_AFFAIRE') = '' then
  begin
    PGIBoxAF (TexteMsgAffaire [6] , '') ;
    Exit;
  end;

  if GetField ('AFF_TIERS') = '' then
  begin
    PGIBoxAF (TexteMsgAffaire [11] , '') ;
    Exit;
  end;

  if (NbPiece = -1) or (TypeAppel = 'SPIGAO') then
  begin
    if (DS.State in [dsInsert] ) then
      MAJZonesCodeAffaire;
    if not PieceChargee then
      NbPiece := SelectPieceAffaire (GetField ('AFF_AFFAIRE') , GetField ('AFF_STATUTAFFAIRE') , CleDocAffaire)
    else
      NbPiece := 1;
  end;

  if (NbPiece = 0) then
  begin
    if GetField ('AFF_AFFAIREHT') = 'X' then
      EnHT := True
    else
      EnHT := False;
    if CreerPieceAffaire (GetField ('AFF_TIERS') , GetField ('AFF_AFFAIRE') , GetField ('AFF_STATUTAFFAIRE') , GetField ('AFF_ETABLISSEMENT') , CleDocAffaire, EnHT, GetContre) then
      NbPiece := 1;
  end;

  if (NbPiece = 1) then
  begin
    if GetContre <> SaisieContre then
    begin
      ChangePieceContre (False, Getcontre) ;
      PieceContre := GetContre;
    end;
    if (TypeAppel = 'LIGNE') or (TypeAppel = 'SPIGAO')then
    begin
      HTBefore := GetField ('AFF_TOTALHTDEV') ;
      TaxeBefore := GetField ('AFF_TOTALTAXEDEV') ;
      PRBefore := GetField ('AFF_TOTALPR') ;
      if TheTob <> nil then
      begin
        TheTob.Free;
        TheTob := nil;
      end;

      TheTob := Tob.Create ('AFFAIRE', nil, -1) ;
      TheTob.PutValue ('AFF_GENERAUTO', GetField ('AFF_GENERAUTO') ) ;
      TheTob.PutValue ('AFF_DETECHEANCE', GetField ('AFF_DETECHEANCE') ) ;
      V_PGI.ZoomOLE := true;

      if ((TFfiche (Ecran) .typeaction = taConsult) or (self.ModifAutorisee = false) ) then
        SaisiePiece (CleDocAffaire, taConsult)
      else
        if (TypeAppel = 'SPIGAO') then
        begin
        	SaisiePiece (CleDocAffaire, taModif, '', '', '', False, True)
        end else SaisiePiece (CleDocAffaire, taModif);
      //
      V_PGI.ZoomOLE := false;

      IncidenceLignesSurAffaire (HTBefore, TaxeBefore, PRBefore) ;
      //gm 26/03/2002   car les modifs faites dans entete depuis ligne étaient
      // non affichée quand on allait ensuite dans condition reglement depuis l'affaire
      piecechargee := false;
    end
    else
    begin
      LanceEntetePieceAffaire (GetField ('AFF_TIERS') , GetField ('AFF_AFFAIRE') , CleDocAffaire) ; // Appel Entête à voir
    end;
  end
  else
  begin
    PGIBoxAF (TexteMsgAffaire [31] , '') ;
  end;
end;

procedure TOM_Affaire.LanceEntetePieceAffaire (CodeTiers, CodeAffaire: string; CleDocAffaire: R_CLEDOC) ;
var
  TOBPiece_O: TOB;
  Action: TActionFiche;
begin
  LoadTobPieceaffaire (CleDocAffaire, false) ;
  TOBPiece_O := TOB.Create ('', nil, -1) ;
  // Ajout BTP
  AddLesSupEntete (TOBPiece_O) ;
  //
  TOBAdresses := TOB.Create ('Les Adresses', nil, -1) ; //
  if VH_GC.GCIfDefCEGID then
  begin
    TOB.Create ('PIECEADRESSE', TOBAdresses, -1) ; {Livraison}
    TOB.Create ('PIECEADRESSE', TOBAdresses, -1) ; {Facturation}
  end
  else
  begin
    // MCD 18/04/03 ??? ilf aut passer sur nouvelle adresse
    if GetParamSoc ('SO_GCPIECEADRESSE') then
    begin
      TOB.Create ('PIECEADRESSE', TOBAdresses, -1) ; {Livraison}
      TOB.Create ('PIECEADRESSE', TOBAdresses, -1) ; {Facturation}
    end
    else
    begin
      TOB.Create ('ADRESSES', TOBAdresses, -1) ; {Livraison}
      TOB.Create ('ADRESSES', TOBAdresses, -1) ; {Facturation}
    end;
  end;

  LoadLesAdresses (TobPiece, TobAdresses) ;
  PieceChargee := True;
  TOBPiece_O.Dupliquer (TOBPiece, True, True) ;
  TobPiece.PutValue ('GP_FACTUREHT', GetField ('AFF_AFFAIREHT') ) ;

  if (self.ModifAutorisee = false) then
    Action := taConsult
  else
    Action := TFfiche (Ecran) .TypeAction;
  AccesEntetePieceAffaire (TobPiece, TobAdresses, Action) ;
  IncidenceEntetePieceSurAffaire (TobPiece) ;

  TOBPiece_O.Free;
  TOBAdresses.Free;
end;

procedure TOM_Affaire.LoadTobPieceAffaire (CleDocAffaire: R_CLEDOC; AvecDetail: Boolean) ;
var
  Q: TQuery;
begin
  if (PieceChargee) and (TobPiece <> nil) then
    Exit;
  if TobPiece <> nil then
  begin
    TOBPiece.Free;
    TOBPiece := nil;
  end;
  TOBPiece := TOB.Create ('PIECE', nil, -1) ;
  // Ajout BTP
  AddLesSupEntete (TOBPiece) ;
  // --
  Q := nil;
  try // normal de tout prendre
    Q := OpenSQL ('SELECT * FROM PIECE WHERE ' + WherePiece (CleDocAffaire, ttdPiece, False) , True) ;
    TOBPiece.SelectDB ('', Q) ;
    if AvecDetail then
    begin
      Ferme (Q) ; // normal de tout prendre
      Q := OpenSQL ('SELECT * FROM LIGNE WHERE ' + WherePiece (CleDocAffaire, ttdLigne, False) + ' ORDER BY GL_NUMLIGNE', True) ;
      if not (Q.EOF) then
        TOBPiece.LoadDetailDB ('LIGNE', '', '', Q, False, True) ;
    end;
  finally
    Ferme (Q) ;
  end;
end;

procedure TOM_Affaire.IncidenceTiersSurAffaire (CodeTiers: string) ;
var Q,Qloc: TQuery;
    EuroDef: Boolean;

begin

  if ((CodeTiers <> '') and (OldTiers <> CodeTiers) ) then
  begin
    Q := nil;
    try
      Q := OPENSQL ('SELECT T_FERME,T_DEVISE,T_MULTIDEVISE,T_FACTUREHT,T_MOISCLOTURE,T_EURODEFAUT,T_NATUREAUXI,T_TVAENCAISSEMENT,T_LIBELLE,T_FACTURE FROM TIERS WHERE T_TIERS="' +
        CodeTiers + '"', False) ;
      if not Q.EOF then
      begin
        OldTiers := CodeTiers;
        // Risque client
        Risque.text := GetEtatRisqueClient (CodeTiers) ;
        GestionFeuEncour;
        //if EtatRisque = 'R' then
        //begin
        //  if VH_GC.GCIfDefCEGID then //29/04/02
        //    PGIBoxAf (TexteMsgAffaire [27] , 'Risque client')
        //  else
        //  begin
        //  SetField ('AFF_TIERS', '') ;
        //    OldTiers := '';
        //    SetFocusControl ('AFF_TIERS') ;
        //    PGIBoxAf (TexteMsgAffaire [25] , 'Risque client') ;
        //    exit;
        //  end;
        //end
        //  else
        //  if EtatRisque = 'O' then
        //begin
        //  PGIBoxAf (TexteMsgAffaire [26] , 'Risque client') ;
        //end;
        if (Q.Findfield ('T_FERME') .AsString = 'X') then
        begin //mcd 19/02/03
          PGIBoxAf (TexteMsgAffaire [36] , 'Client fermé') ;
          SetField ('AFF_TIERS', '') ;
          OldTiers := '';
          SetFocusControl ('AFF_TIERS') ;
        end;
        if (Q.FindField ('T_DEVISE') .AsString <> '') then
          SetField ('AFF_DEVISE', Q.FindField ('T_DEVISE') .AsString)
        else
          SetField ('AFF_DEVISE', V_PGI.DevisePivot) ;
        EuroDef := (Q.FindField ('T_EURODEFAUT') .AsString = 'X') ;
        // mcd 11/03/02 pour saisie Aff en Euro si Euro Exclusif, même si client pas en vente EURO
        if (GetParamSoc ('SO_GCTOUTEURO') ) and (EuroDef = False)
          and ((Getfield ('AFF_DEVISE') = 'EUR') or (Getfield ('AFF_DEVISE') = 'FRF') ) // mcd 17/04/02 ne pas afficher EURO si devise pas EURO
        then
        begin
          EuroDef := True;
          SetField ('Aff_Devise', 'EUR')
        end;
        if VH^.TenueEuro then
          EuroDef := not (EuroDef) ;
        if EuroDef then
          SetField ('AFF_SAISIECONTRE', 'X')
        else
          SetField ('AFF_SAISIECONTRE', '-') ;
        // mcd 02/01/01 pas renseigner. Pb si on va ds ech lors de la création
        if GetField ('AFF_SAISIECONTRE') = 'X' then
        begin
          SaisieContre := True;
          EchContre := true;
          Piececontre := True;
        end
        else
        begin
          SaisieContre := False;
          EchContre := false;
          PieceContre := False;
        end;
        // Chargement de la devise de l'affaire
        ImpactChangeDevise;
        SetField ('AFF_AFFAIREHT', Q.FindField ('T_FACTUREHT') .AsString) ;
        // Mois de clôture
        if ctxScot in V_PGI.PGIContexte then
        begin
          MoisCloture := Q.FindField ('T_MOISCLOTURE') .AsInteger;
          if MoisCloture <> 0 then
          begin
            SetControlText ('MOISCLOTURE', LongMonthNames [MoisCloture] ) ;
          end;
        end;
        if (StatutAffaire = 'INT') then
        begin
          // Recup Nature d'auxi pour gestion Prospect / Client
          NatureAuxi := Q.FindField ('T_NATUREAUXI') .AsString;
          SetField('AFF_FACTURE', Q.FindField ('T_FACTURE') .AsString) ;
          QLoc := OpenSql('SELECT T_TIERS FROM TIERS WHERE T_AUXILIAIRE="' +
                  GetField('AFF_FACTURE') + '"', true,1,'',true);
          if not QLoc.Eof then
          begin
            OldFacture:= QLoc.Findfield('T_TIERS').AsString;
            SetControlText ('TX_FACTURE', OldFacture);
            Setfield('TX_FACTURE', OldFacture);
          end;
          Ferme(QLoc);
        end;
      end; // Fin Q.EOF enreg trouvé
    finally
{$IFDEF BTP}
      if not Q.eof then
        THLabel (GetControl ('LIBTIERS') ) .caption := Q.findfield ('T_LIBELLE') .AsString; // modif BTP
{$ENDIF}
      Ferme (Q) ;
    end;
  end;
end;

procedure TOM_Affaire.IncidenceLignesSurAffaire (HTBefore, TaxeBefore, PRBefore: Double) ;
var
  TotTaxes, PRAfter, Marge: double;
  CompNum: THNumedit;
  bMaj: Boolean;

begin
  bMaj := False;
  if TheTob <> nil then
  begin
    if not (TheTob.FieldExists ('GP_TOTALHTDEV') ) then
    begin
      TheTOB.Free;
      TheTob := nil; // mcd 27/05/02 il faut effacer la Thetob idem fin fct ...
      Exit;
    end;

    if not (DS.State in [dsInsert, dsEdit] ) then
      DS.edit;
    // Reprise de la devise modifiable dans les lignes de l'affaire
    SetField ('AFF_DEVISE', TheTob.GetValue ('GP_DEVISE') ) ;
    ImpactChangeDevise;

    TotTaxes := TheTOB.GetValue ('GP_TOTALTTCDEV') - TheTOB.GetValue ('GP_TOTALHTDEV') ;
    // Modif Total HT ou Taxes
    if (HTBefore <> TheTOB.GetValue ('GP_TOTALHTDEV') ) or (TaxeBefore <> TotTaxes) then
    begin
      bMaj := True;
      // MAJ des zones d'affichages,et des zones de la base
      CompNum := THNumEdit (GetControl ('LAFF_TOTALHTDEV') ) ;
      CompNum.Value := Valeur (TheTOB.GetValue ('GP_TOTALHTDEV') ) ;
      SetField ('AFF_TOTALHTDEV', CompNum.Value) ;
      CompNum := THNumEdit (GetControl ('LAFF_TOTALTAXEDEV') ) ;
      CompNum.Value := TotTaxes;
      SetField ('AFF_TOTALTAXEDEV', CompNum.Value) ;
      CompNum := THNumEdit (GetControl ('LAFF_TOTALTTCDEV') ) ;
      CompNum.Value := Valeur (TheTOB.GetValue ('GP_TOTALTTCDEV') ) ;
      SetField ('AFF_TOTALTTCDEV', CompNum.Value) ;
      ImpactCalculTotalHT (TitTotLignes) ;
      if not (ctxscot in V_PGI.PGIContexte) then
      begin
        if (GetField ('AFF_GENERAUTO') = 'CON') then
           //Modif FV problème de calcul de l'échéance (19012007)
           //SetField ('AFF_MONTANTECHEDEV', GetField ('AFF_TOTALHTDEV') * GetField ('AFF_INTERVALGENER') ) ;
           CalculEchSurDoc(true, false);
           //SetField ('AFF_MONTANTECHEDEV', GetField ('AFF_TOTALHTDEV')) ;
      end;
    end;

    SetField ('AFF_TOTALHTGLODEV', TheTOB.GetValue ('GP_TOTALHTDEV'));

    PRAfter := RecupPRPieceAffaire (TheTOB) ;
    if PRBefore <> PRAfter then
    begin
      bMaj := True;
      SetField ('AFF_TOTALPR', PRAfter) ;
    end;

    if bMaj then
    begin
      Marge := CalculMargesurAffaire (Valeur (TheTOB.GetValue ('GP_TOTALHT') ) , PRAfter, GetParamsoc ('SO_AFMARGEAFFAIRE') ) ;
      if THedit (GetControl ('VALEURMARGE') ) <> nil then
        SetControlText ('VALEURMARGE', FloatTostr (Marge) ) ;
      TFFiche (Ecran) .Bouge (nbPost) ; // sauvegarde obligatoire maj des totaux de l'affaire en phase avec la pièce enreg.
    end;
    TheTOB.Free;
    TheTob := nil;
  end;
end;

procedure TOM_Affaire.IncidenceEntetePieceSurAffaire (Tobpiece: TOB) ;
begin

end;

//******************************************************************************
//******************* Fonction de calcul de dates sur la mission ***************
//******************************************************************************

procedure TOM_Affaire.CalculDatesMission;
var
  Exercice: string;
  DebutExer, Finexer, DebutInter, FinInter, DebutFact, FinFact, Liquidative, Garantie, Cloture: TDateTime;
  CC: THEdit;
  iMoisEnPlus: integer;
begin
  if not (DS.State in [dsInsert, dsEdit] ) then
    DS.edit;
  // mcd 18/09/01 if ctxScot in V_PGI.PGIContexte then
  if GetParamSoc ('SO_AfFormatExer') <> 'AUC' then
  begin
    CC := THEdit (GetControl ('CAFF_AFFAIRE2') ) ;
    Exercice := CC.Text;
    if (GetControlText ('AFF_TIERS') = '') and (GetControlText ('AFF_MODELE') <> 'X') then
    begin
      PGIBoxAF (TraduitGa (TexteMsgAffaire [7] ) , '') ;
      Exit;
    end;

    if MoisCloture = 0 then
      MoisCloture := GetMoisClotureTiers (GetField ('AFF_TIERS') ) ;
    if (MoisCloture < 1) or (Moiscloture > 12) then
      MoisCloture := 12;
    CalculDateExercice (VH_GC.AFFORMATEXER, MoisCloture, Exercice, DebutExer, Finexer, DebutInter, FinInter, DebutFact, FinFact, Liquidative, Garantie, Cloture, True) ;
    SetField ('AFF_DATEDEBEXER', DateToStr (DebutExer) ) ;
    SetField ('AFF_DATEFINEXER', DateToStr (finExer) ) ;
    SetField ('AFF_DATEDEBUT', DateToStr (DebutInter) ) ;
    SetField ('AFF_DATEFIN', DateToStr (finInter) ) ;
    if GetField ('AFF_AFFCOMPLETE') <> '-' then // si affaire light pas de maj des dates de facturation
    begin
      // PL 22/01/2001 pour ne pas écraser la date de début ou de fin de facturation si saisie à la main
      if (GetField ('AFF_DATEDEBGENER') = iDate1900) or (GetField ('AFF_DATEDEBGENER') = iDate2099) then
        SetField ('AFF_DATEDEBGENER', DateToStr (DebutFact) ) ;
      if (GetField ('AFF_DATEFINGENER') = iDate1900) or (GetField ('AFF_DATEFINGENER') = iDate2099) then
      begin
        if (TGroupBox (GetControl ('GroupBoxDateExer') ) .visible = false) then
          SetField ('AFF_DATEFINGENER', DateToStr (finFact) )
        else
        begin
          iMoisEnPlus := GetParamSoc ('SO_AFFINFACTURAT') ;
          // mcd 04/10/02 ajout fin de mois
          SetField ('AFF_DATEFINGENER', DateToStr (FinDemois (PlusDate (finExer, iMoisEnPlus, 'M') ) ) ) ;
        end;
      end;
    end;
    SetField ('AFF_DATEFACTLIQUID', DateToStr (Liquidative) ) ;
    SetField ('AFF_DATEGARANTIE', DateToStr (Garantie) ) ;
    SetField ('AFF_DATECLOTTECH', DateToStr (Cloture) ) ; //17/10/01
  end
  else
  begin
    DebutInter := StrTodate (GetField ('AFF_DATEDEBUT') ) ;
    finInter := StrToDate (GetField ('AFF_DATEFIN') ) ;
    if (DebutInter = iDate1900) or (FinInter = iDate2099) then
    begin
      PGIBoxAF (TexteMsgAffaire [19] , 'Calcul des dates de l''affaire') ;
      Exit;
    end;
    CalculDateAffaire (DebutInter, FinInter, DebutFact, FinFact, Liquidative, Garantie, Cloture) ;
    // PL 22/01/2001 pour ne pas écraser la date de début ou de fin de facturation si saisie à la main
    if (GetField ('AFF_DATEDEBGENER') = iDate1900) or (GetField ('AFF_DATEDEBGENER') = iDate2099) then
      SetField ('AFF_DATEDEBGENER', DateToStr (DebutFact) ) ;
    if (GetField ('AFF_DATEFINGENER') = iDate1900) or (GetField ('AFF_DATEFINGENER') = iDate2099) then
    begin
      if (TGroupBox (GetControl ('GroupBoxDateExer') ) .visible = false) then
        SetField ('AFF_DATEFINGENER', DateToStr (finFact) )
      else
      begin
        iMoisEnPlus := GetParamSoc ('SO_AFFINFACTURAT') ;
        // mcd 04/10/02 ajput fion de mois
        SetField ('AFF_DATEFINGENER', DateToStr (FInDeMois (PlusDate (finExer, iMoisEnPlus, 'M') ) ) ) ;
      end;
    end;
    SetField ('AFF_DATECLOTTECH', DateToStr (Cloture) ) ;
    SetField ('AFF_DATEGARANTIE', DateToStr (Garantie) ) ;
  end;
end;

//******************************************************************************
//******************* Fonction de calcul du montant global ***************
//******************************************************************************
procedure TOM_Affaire.CalculEchSurDoc(bSilence, bPourcentage: boolean);
Var NbIt        : Integer;
    RepTest     : Integer;
    TypeCalcul  : string;
    DateDebut   : TDateTime;
    DateFin     : TDateTime;
    DateLiquid  : TDateTime;
    DateResil   : TDateTime;
    TypeModif   : T_TypeModifAff;
    MtDoc       : double;
    STPeriode   : String;
    Interval    : Integer;
Begin

  MtDoc := 0;
  
  // Controle de la coherence des donnees saisies
  RepTest := TestTypeModifEcheances(DateDebut, DateFin, DateLiquid, DateResil, TypeCalcul, MtDoc, TypeModif, 'BGlobal', bSilence);

  // Date de debut incorrecte
  if (RepTest = 2) then exit;

  CalcEcheance := GetField('AFF_DETECHEANCE');
  MtDoc := Double(GetField('AFF_TOTALHTDEV'));

  StPeriode  := GetField('AFF_PERIODICITE');
  Interval   := Integer(GetField('AFF_INTERVALGENER'));

  if TypeCalcul = 'NBI' then
  begin
    NbIt := 1;
    if Interval = 0 then Interval := 1
  end else
  begin
    NbIt := EvaluationNbEcheances (GetField('AFF_AFFAIRE'), TypeCalcul, Interval, DateDebut, DateFin) ;
  end;

  if NbIt = 0 then Exit;
  MtDoc := CalculMtEcheanceContrat (MtDoc,CalcEcheance,stPeriode,Interval,NbIt);
  if (MtDoc <> GetField ('AFF_MONTANTECHEDEV') ) then
  begin
    if not (DS.State in [dsInsert, dsEdit] ) then DS.edit;
    SetField('AFF_MONTANTECHEDEV', Mtdoc);
  end;

end;


procedure TOM_Affaire.CalculMontantGlobal (bSilence, bPourcentage: boolean) ;
var
  QQ: Tquery;
  NbIt, RepTest: Integer;
  MtDejaFact, MtRestant, MtNewEch: currency;
  sReq: string;
  TypeCalcul: string;
  DateDebut, DateFin, DateLiquid, DateResil: TDateTime;
  TypeModif: T_TypeModifAff; // T_TypeModifAff = (tmaMntEch,tmaPourcEch,tmaDate,tmaDFacLiq) ;
  zmont: double;
begin

  MtNewEch := 0;

  // Controle de la coherence des donnees saisies
  RepTest := TestTypeModifEcheances (DateDebut, DateFin, DateLiquid, DateResil, TypeCalcul, zmont, TypeModif, 'BGlobal', bSilence) ;

  // Date de debut incorrecte
  if (RepTest = 2) then exit;

  NbIt := EvaluationNbEcheances (GetField ('AFF_AFFAIRE') , TypeCalcul, Integer (GetField ('AFF_INTERVALGENER') ) , DateDebut, DateFin) ;

  if (NbIt > 0) then
  begin
    MtDejaFact := 0;
    // recherche de la somme des échéances deja facturées
    QQ := nil;
    try
      sReq := 'SELECT AFA_AFFAIRE,SUM(AFA_MONTANTECHEDEV) FROM FACTAFF WHERE AFA_AFFAIRE="' + GetField ('AFF_AFFAIRE') + '" AND AFA_TYPECHE="NOR" AND AFA_ECHEFACT="X" AND AFA_LIQUIDATIVE="-" GROUP BY AFA_AFFAIRE';
      QQ := OpenSQL (sReq, true) ;
      if not QQ.EOF then
        MtDejaFact := QQ.Fields [1] .AsFloat;
    finally
      Ferme (QQ) ;
    end;
    MtRestant := GetField ('AFF_TOTALHTGLODEV') - MtDejaFact;
    if (MtRestant <= 0) then
      exit;
    if (NbIt <> 0) then
      MtNewEch := MtRestant / NbIt;
    if not (DS.State in [dsInsert, dsEdit] ) then
      DS.edit;
    SetField ('AFF_MONTANTECHEDEV', MtNewEch) ;
  end;

end;

procedure TOM_Affaire.ChargeAffTiers(codeAffaire : string);
var
  Q: TQuery;
begin
  TOBAffTiers.cleardetail;

  if GAffTiers <> nil then AffecteGrid (GAffTiers, taConsult) ;

  if not (DS.State in [dsInsert] ) then
  begin
    // SELECT * : ne concerne qu'une affaire......
    Q := OpenSQL ('SELECT * FROM AFFTIERS WHERE AFT_AFFAIRE="' + CodeAffaire + '" ORDER BY AFT_RANG', True) ;
    if not Q.EOF then
      TOBAffTiers.LoadDetailDB ('AFFTIERS', '', '', Q, False, true) ;
    Ferme (Q) ;
  end;
  //
  if GAffTiers <> nil then
  begin
    TOBAffTiers.PutGridDetailOnListe (GAffTiers, 'AFFTIERS') ;
    TFFiche (Ecran) .Hmtrad.ResizeGridColumns (GAffTiers) ;
  end;

end;

procedure TOM_Affaire.GAffTiersDblClick (Sender: TObject) ;
begin
  if (TOBAffTiers.detail.count > 0) then
    LanceAffTiers (GAffTiers.Cells [1, GAffTiers.Row] )
  else
    LanceAffTiers ('') ;
end;

procedure TOM_Affaire.LanceAffTiers (Code: string) ;
var
  StAction: string;
begin
  // Blocage en mode création
  if (DS.State in [dsInsert] ) or ((trim (code) = '') and (TFFiche (Ecran) .typeAction = taConsult) ) then
    exit;
  if (trim (code) <> '') then
  begin
    if (TFFiche (Ecran) .typeaction = taCreat) then
      StAction := 'ACTION=MODIFICATION'
    else
      StAction := ActionAffaireString (self.ModifAutorisee, TFFiche (Ecran) .typeaction) ;
  end
  else
    StAction := 'ACTION=CREATION';
  AFLanceFiche_LienAffaireTiers (GetField ('AFF_AFFAIRE') , code, StAction + ';AFFAIRE=' + GetField ('AFF_AFFAIRE') ) ;
  ChargeAffTiers(GetField ('AFF_AFFAIRE'));
end;

procedure TOM_Affaire.ImprimeModele (Statut: string; F: TForm) ;
var
  reponse: string;
begin
  reponse := AFLanceFiche_ChoixImprAff ('STATUT=' + StatutAffaire + ';AFFAIRE=' + GetField ('AFF_AFFAIRE') ) ;

  if (reponse = 'FICHE') then
    TFFiche (F) .BImprimerClick (nil) ;
  {else
      if reponse='PARSER' then
      begin
           if not(DS.State in [dsInsert,dsEdit]) then
              DS.edit;
           SetField ('AFF_DATELETTREMIS', UsDateTime (V_PGI.DateEntree));
           TFFiche(Ecran).Bouge (nbPost) ;    // sauvegarde obligatoire maj des totaux de l'affaire en phase avec la pièce enreg.
      end;}
  //       DS.Locate ('AFF_AFFAIRE', GetField ('AFF_AFFAIRE'), []);
end;

procedure TOM_Affaire.LanceRegroupeAffaire;
var
  stArg, Tmp, Champ, Valeur, ActionAffaire: string;
  X: integer;
  bValide: Boolean;
begin
  Champ := '';
  Valeur := '';
  bValide := False;
  // Blocage en mode création
  if not (DS.State in [dsInsert] ) then
  begin
    ActionAffaire := ActionAffaireString (self.ModifAutorisee, TFFiche (Ecran) .typeaction) ;
    stArg := AFLanceFiche_RegroupAffaire (ActionAffaire + ';AFF_AFFAIRE:' + GetField ('AFF_AFFAIRE') + ';AFF_TIERS:' + GetField ('AFF_TIERS') + ';AFF_ISAFFAIREREF:' + GetField ('AFF_ISAFFAIREREF') + ';AFF_AFFAIREREF:' + GetField ('AFF_AFFAIREREF') ) ;
    Tmp := (Trim (ReadTokenSt (stArg) ) ) ;
    while (Tmp <> '') do
    begin
      if Tmp <> '' then
      begin
        X := pos (':', Tmp) ;
        if x <> 0 then
        begin
          Champ := copy (Tmp, 1, X - 1) ;
          Valeur := Copy (Tmp, X + 1, length (Tmp) - X) ;
        end;
        if Champ = 'AFF_ISAFFAIREREF' then
        begin
          if not (DS.State in [dsInsert, dsEdit] ) then
            DS.edit;
          SetField (champ, valeur) ;
          bValide := True;
        end;
        if Champ = 'AFF_AFFAIREREF' then
        begin
          if not (DS.State in [dsInsert, dsEdit] ) then
            DS.edit;
          SetField (champ, valeur) ;
          bValide := True;
        end;
      end;
      Tmp := (Trim (ReadTokenSt (stArg) ) ) ;
    end;

    // PL le 31/05/02 : ajout du test sur la tob TheTOB contenant la table AFFAIRE
    if not (bValide) and (TheTob <> nil) and TheTOB.TOBNameExist ('AFFAIRE') then
    begin
      BValide := True;
      if not (DS.State in [dsInsert, dsEdit] ) then
        DS.edit;
    end;

    if bValide then
    begin
      if (TheTob <> nil) then
      begin
        if (TheTob.Detail.count > 0) then
        begin
          TobSSAff := TOB.Create ('liste affaire à valider', nil, -1) ;
          TobSSAff.Dupliquer (TheTOB, True, True, True) ;
        end;
        TheTob.Free;
        TheTob := nil;
      end;
      TFFiche (Ecran) .Bouge (nbPost) ; // sauvegarde obligatoire maj affaire ref ...
      TobSSAff.Free;
      TobSSAff := nil;
      ImpactAffRefouIndependante (True) ;
    end;
  end;
end;
// *****************************************************************************
// ************** Gestion des affaires complètes ou light **********************
// *****************************************************************************

procedure TOM_Affaire.AFF_AFFCOMPLETEOnClick (Sender: TObject) ;
begin
  ImpactAffCompleteOuLight;
end;

procedure TOM_Affaire.ImpactAffCompleteOuLight;
var
  bComplete: BOOLEAN;
begin
  bComplete := (GetControlText ('AFF_AFFCOMPLETE') = 'X') ;
  //if Not(GereAffCompleteEtLight) and bComplete then Exit;
{$IFDEF BTP}
  if StatutAffaire = 'INT' then
  begin
    SetControlVisible ('TABSHEETFACTURATION', bComplete) ;
    SetControlVisible ('TABSHEETCOMPLEMENT', bcomplete) ;
  end
  else
  begin
    SetControlVisible ('TABSHEETFACTURATION', False) ;
    SetControlVisible ('TABSHEETCOMPLEMENT', False) ;
  end;
{$ELSE}
  SetControlVisible ('TABSHEETFACTURATION', bComplete) ;
  SetControlVisible ('TABSHEETCOMPLEMENT', bComplete) ;
{$ENDIF}
end;

procedure TOM_Affaire.OnUpdateAffaireLight;
begin

  if GetField ('AFF_AFFCOMPLETE') = 'X' then Exit;

  // si affaire Light
  SetField ('AFF_GENERAUTO', 'MAN') ;
  SetField ('AFF_DATEDEBGENER', idate1900) ;
  SetField ('AFF_DATEFINGENER', idate2099) ;
  SetField ('AFF_COEFFREVALO', 0) ;
  SetField ('AFF_RECONDUCTION', 'SUP') ;

end;
// *****************************************************************************
// ************** Gestion des affaires de références / ss affaires *************
// *****************************************************************************

procedure TOM_Affaire.OnUpdateAffaireRef;
begin
  if (EnSaisie) then
    if not (DS.State in [dsInsert] ) then
      exit; // mis à jour depuis la fen. complémentaire
  // indépendante par défaut
  SetField ('AFF_AFFAIREREF', GetField ('AFF_AFFAIRE') ) ;
  SetField ('AFF_ISAFFAIREREF', '-') ;
end;

procedure TOM_Affaire.ImpactAffRefouIndependante (bRecharge: Boolean) ;
begin
  if not (GereSousAffaire) then
    Exit;
  if bRecharge then
    FindTypeAffaire (TypeAffaire, GetField ('AFF_ISAFFAIREREF') , GetField ('AFF_AFFAIREREF') , GetField ('AFF_AFFAIRE') , GetField ('AFF_AFFCOMPLETE') ) ;
  if (tIsAffRef in TypeAffaire) or (tIsSSAff in TypeAffaire) then
  begin
{$IFDEF BTP}
    SetControlProperty ('AFF_REGROUPEFACT', 'plus', ' AND ((CC_LIBRE = "A")  OR ( CC_LIBRE= "BTP" ))') ;
{$ELSE}
    SetControlProperty ('AFF_REGROUPEFACT', 'plus', ' AND ((CC_LIBRE = "A")  OR ( CC_LIBRE= "C" ))') ;
{$ENDIF}
    AlimLibEtat;
    if (tIsAffRef in TypeAffaire) then
      SetControlVisible ('GROUPBOXREGROUPEMENT', False) ;
  end
  else // Affaire indépendante
  begin
    SetControlVisible ('GROUPBOXREGROUPEMENT', True) ;
    if StatutAffaire <> 'PRO' then
      SetControlVisible ('LIBETAT', false) ;
{$IFDEF BTP}
    SetControlProperty ('AFF_REGROUPEFACT', 'plus', ' AND ((CC_LIBRE = "A")  OR ( CC_LIBRE= "BTP" ))') ;
{$ELSE}
    SetControlProperty ('AFF_REGROUPEFACT', 'plus', ' AND ((CC_LIBRE = "A")  OR ( CC_LIBRE= "B" ))') ;
{$ENDIF}
  end;
end;

function TOM_Affaire.TiersIdemAffRef: Boolean;
begin
  Result := True;
  if not (tIsSSAff in TypeAffaire) then
    exit;
  if GetField ('AFF_AFFAIREREF') = '' then
    Exit;
  if GetField ('AFF_TIERS') <> GetChampsAffaire (GetField ('AFF_AFFAIREREF') , 'AFF_TIERS') then
    Result := False;
end;

// *****************************************************************************
// ******************* Alim du libellé d'état de l'affaire  ********************
// *****************************************************************************

procedure TOM_AFFAIRE.AlimLibEtat;
var
  stLibEtat: string;
begin
  if (EtatAffaire <> 'ENC') or (tIsAffRef in Typeaffaire) or (tIsSSAff in Typeaffaire) then
  begin
    stLibEtat := '';
    if (tIsSSAff in Typeaffaire) then
      stLibEtat := 'Sous ';
    if StatutAffaire = 'PRO' then
{$IFDEF BTP}
      stLibEtat := stLibEtat + TraduitGA ('Appel d''offre ')
{$ELSE}
      stLibEtat := stLibEtat + TraduitGA ('Proposition ')
{$ENDIF}
    else
      stLibEtat := stLibEtat + TraduitGA ('Affaire ') ;
    if (tIsAffRef in Typeaffaire) then
      stLibEtat := stLibEtat + 'de référence ';

{$IFDEF BTP}
    if (EtatAffaire = 'ACP') then
      stLibEtat := stLibEtat + 'acceptée'
    else //MODIFBTP
{$ELSE}
    if (EtatAffaire = 'ACC') then
      stLibEtat := stLibEtat + 'acceptée'
    else
{$ENDIF}
      if (EtatAffaire = 'REF') then
      stLibEtat := stLibEtat + 'refusée'
    else
      if (EtatAffaire = 'CLO') then
      stLibEtat := stLibEtat + 'terminée'
    else
      if not (tIsSSAff in Typeaffaire) then
      stlibetat := ''; // mcd 10/12/02 pour ne pas afficher Affaire tout seul..
    // Affichage de quelle affaire de référence on est sous affaire
{$IFDEF BTP}
    if (tIsSSAff in Typeaffaire) then
      stLibEtat := stLibEtat + TraduitGa (' de l''affaire ') + BTPCodeAffaireAffiche (GetField ('AFF_AFFAIREREF') ) ;
{$ELSE}
    if (tIsSSAff in Typeaffaire) then
      stLibEtat := stLibEtat + TraduitGa (' de l''affaire ') + CodeAffaireAffiche (GetField ('AFF_AFFAIREREF') ) ;
{$ENDIF}

    SetControlText ('LIBETAT', stLibEtat) ;
    SetControlVisible ('LIBETAT', true) ;
  end
  else
  begin
    SetControlVisible ('LIBETAT', false) ;
  end;
end;

function TOM_AFFAIRE.ProspectToClient: Boolean;
var
  NatPiece: string;
  MajProCli: Boolean;

begin
  Result := True;
  MajProCli := False;
  if (ctxScot in V_PGI.PGIContexte) then
    Exit;
  if (not (EnSaisie) ) or (not (DS.State in [dsInsert] ) )
    or (NatureAuxi <> 'PRO') or (GetField ('AFF_TIERS') = '') then
    Exit;

  if (GetField ('AFF_STATUTAFFAIRE') = 'PRO') then
    NatPiece := VH_GC.AFNatProposition
{$IFDEF BTP}
  else
    if (GetField ('AFF_STATUTAFFAIRE') = 'PRO') then
    NatPiece := 'AFF'
{$ENDIF}
  else
    NatPiece := VH_GC.AFNatAffaire;
{$IFNDEF BTP}
  if (NatPiece <> '') then
    MajProCli := (GetInfoParPiece (NatPiece, 'GPP_PROCLI') = 'X') ;
  if MajProCli then
    result := AFMajProspectClient (GetField ('AFF_TIERS') ) ;
{$ENDIF}
  { attention , si modif voir la fonction  :TransformeProspectClient }
  { il faudrait peut être utilise l'autre }
end;

// *****************************************************************************
// ****************************** Gestion du Contact  **************************
// *****************************************************************************

procedure TOM_AFFAIRE.NomContactAff;
var
  Q: TQuery;
  Nom, Telephone: string;
begin
  if (GetField ('AFF_TIERS') <> '') and (GetField ('AFF_NUMEROCONTACT') <> 0) then
  begin
    Nom := '';
    Telephone := '';
    Q := OpenSql ('SELECT C_NOM,C_TELEPHONE FROM CONTACT WHERE C_TYPECONTACT="T" AND C_AUXILIAIRE="' + TiersAuxiliaire (GetField ('AFF_TIERS') ) + '" AND C_NUMEROCONTACT=' + intToStr (GetField ('AFF_NUMEROCONTACT') ) , TRUE) ;
    if not Q.EOF then
    begin
      Nom := Q.FindField ('C_NOM') .asstring;
      Telephone := Q.FindField ('C_TELEPHONE') .asstring;
    end;
    Ferme (Q) ;
    SetControlText ('LECONTACT', Nom) ;
    if THLabel (GetControl ('TELEPHONECONTACT') ) <> nil then
    begin
      if Telephone <> '' then
        SetControlText ('TELEPHONECONTACT', 'Tél : ' + Telephone) ;
    end;
  end
  else
  begin
    SetControlText ('LECONTACT', '') ;
    if THLabel (GetControl ('TELEPHONECONTACT') ) <> nil then
    begin
      SetControlText ('TELEPHONECONTACT', '') ;
    end;
  end;
end;

procedure TOM_AFFAIRE.TraiteAffTermine;
var
  NewDateFin: TDateTime;
begin
  //NewDateFin := GetField('AFF_DATEFINGENER');

  // *** passage en terminée ***
  if (GetField ('AFF_ETATAFFAIRE') = 'CLO') and (GetField ('AFF_ETATAFFAIRE') <> EtatAffaire) then
  begin
    // *** Attention :si modif voir fx MajAffRemplaceeParAvenant : mis état CLO ...
    // voir aussi fct ModifAvecTrait dans UtofAffair_ModifLot
    SupEcheancesAffaire (GetField ('AFF_AFFAIRE') , False, iDate1900) ;
    NewDateFin := AjusteDateGenerSurEch (GetField ('AFF_AFFAIRE') , False, False, iDate1900, GetField ('AFF_DATEDEBGENER') ) ;
    SetField ('AFF_DATEFINGENER', NewDatefin) ;
    //   if (GetField('AFF_DATEFIN') = idate2099) then SetField ('AFF_DATEFIN',V_PGI.DateEntree);
    // gm le 29/11/02, sinon on avait date fin < date début
    if (GetField ('AFF_DATEFIN') = idate2099) then
    begin
      if (StrToDate (GetField ('AFF_DATEDEBUT') ) > V_PGI.DateEntree) then
        SetField ('AFF_DATEFIN', GetField ('AFF_DATEDEBUT') )
      else
        SetField ('AFF_DATEFIN', V_PGI.DateEntree) ;
    end;

  end;

  // Mise d'une date de résiliation sur un contrat
  if (GetField ('AFF_DATERESIL') <> DateResil_sav) and
     (GetField ('AFF_DATERESIL') <> iDate2099)     and
     (GetField ('AFF_GENERAUTO') = 'CON')         then
  begin
    SetField ('AFF_DATEFIN', GetField ('AFF_DATERESIL') ) ;
    SetField ('AFF_DATEFINGENER', GetField ('AFF_DATERESIL')) ;
    (*
    SupEcheancesAffaire (GetField ('AFF_AFFAIRE') , False, GetField ('AFF_DATERESIL') ) ;
    NewDateFin := AjusteDateGenerSurEch (GetField ('AFF_AFFAIRE') , False, False, GetField ('AFF_DATERESIL') , GetField ('AFF_DATEDEBGENER') ) ;
    SetField ('AFF_DATEFINGENER', NewDatefin) ;
    *)
  end;

  if VH_GC.GCIfDefCEGID then
    //  gm le 17/04/03 afin de pouvoir comptabiliser le nbre de contrats résiliés par jour.
    if (GetField ('AFF_DATERESIL') = idate2099) then
      SetField ('AFF_DATELIBRE1', idate1900)
    else
      if (GetField ('AFF_DATERESIL') <> DateResil_old) then
      SetField ('AFF_DATELIBRE1', V_PGI.DateEntree) ;
end;

// *****************************************************************************
// ************** Fonction Affaire spécif CEGID ...   **************************
// *****************************************************************************

procedure TOM_AFFAIRE.OnArgumentCEGID;
var
  menuPerso: TMenuItem;
begin
  { à gérer avec restruction fiche
  SetControlVisible ('AFF_DATELIMITE', False);   SetControlVisible ('TAFF_DATELIMITE', False);
  SetControlVisible ('AFF_DATECLOTTECH', False); SetControlVisible ('TAFF_DATECLOTTECH', False);
  SetControlVisible ('AFF_DATEGARANTIE', False); SetControlVisible ('TAFF_DATEGARANTIE', False);
  //SetControlVisible ('AFF_DEPARTEMENT', False);  SetControlVisible ('TAFF_DEPARTEMENT', False); gm 05/03/02
  SetControlVisible ('LISTE_AFFTIERS', False);   SetControlVisible ('LISTEINTERVENANT', False);
  SetControlVisible('AFF_RESPONSABLE',False);    SetControlVisible('TAFF_RESPONSABLE',False);
  SetControlVisible('AFF_DATESIGNE',False);      SetControlVisible('TAFF_DATESIGNE',False);
  menuPerso := TMenuItem(GetControl('mnIntervenant'));
  if menuPerso <> Nil then menuPerso.Visible:=false;
  menuPerso := TMenuItem(GetControl('mnAdresseInt'));
  if menuPerso <> Nil then menuPerso.Visible:=false;
  menuPerso := TMenuItem(GetControl('mnTB'));
  if menuPerso <> Nil then menuPerso.Visible:=false;
  SetControlVisible ('BRECALCULMTGLOBAL',False);
  SetControlVisible ('GBMARGE',False);
  SetControlEnabled ('AFF_DATEDEBUT',False); SetControlEnabled ('AFF_DATEFIN',False);
  SetControlVisible ('AFF_TOTALHTGLODEV',False); SetControlVisible ('TAFF_TOTALHTGLODEV',False);
  SetControlVisible ('AFF_CALCTOTHTGLO',False);
  SetControlVisible ('AFF_AFFCOMPLETE',False);
  }
end;

procedure TOM_AFFAIRE.OnNewRecordCEGID;
var stTiers: string;
begin
  SetField ('AFF_DATELIBRE1', idate1900) ;
  SetField ('AFF_INTERVALGENER', 3) ;
  if (OldTiers <> '') then
  begin
    SetField ('AFF_TIERS', OldTiers) ;
    stTiers := OldTiers;
    OldTiers := '';
    IncidenceTiersSurAffaire (stTiers) ;
  end;
  SetField ('AFF_METHECHEANCE', 'CIV') ;
  SetField ('AFF_TERMEECHEANCE', 'ACH') ;
end;

procedure TOM_AFFAIRE.OnUpdateCEGID;
var
  stSQL, Reg, RegRegroupe: string;
  QQ: TQuery;
  i: integer;
  bAucun, bRegroupe, bAutresAff: Boolean;
  TobReg: TOB;
  RegSurCaf: boolean;
  LeCaf: string;
  nbpiece: Integer;
  CleDocTemp: R_CLEDOC;
  StTiers, StRep: string;
begin

  if (GetField ('AFF_REGSURCAF') = 'X') then
    RegsurCaf := true
  else
    RegSurCaf := false;
  Lecaf := GetControlText ('LECAF') ;
  if (RegSurCaf) and (LeCaf = '') then
  begin
    NbPiece := SelectPieceAffaireBis (GetField ('AFF_AFFAIRE') , 'AFF', CleDocTemp, Sttiers, StRep) ;
    if (NbPiece = 1) then
      LeCaf := Sttiers;

  end;

  if (RegSurCaf) then
  begin
    stSQL := 'SELECT DISTINCT(AFF_REGROUPEFACT) FROM AFFAIRE LEFT JOIN PIECE ON GP_NATUREPIECEG="' + VH_GC.AFNatAffaire + '" and GP_AFFAIRE=AFF_AFFAIRE ';
    stSQL := stSQL + ' WHERE GP_TIERSFACTURE ="' + LeCaf + '"'
  end
  else
    stSQL := 'SELECT DISTINCT(AFF_REGROUPEFACT) FROM AFFAIRE WHERE AFF_TIERS ="' + GetField ('AFF_TIERS') + '" AND AFF_AFFAIRE <> "' + GetField ('AFF_AFFAIRE') + '"';

  bAucun := False;
  bRegroupe := False;
  bAutresAff := false;

  if (EnSaisie) then
  begin
    // controle sytématique sur contrat reg simple
    // Par contre Si RegSurCaf, on ne fait le test qu'en modif, car en création, on n'a pas encore saisit
    // le CAF du contrat
    if not (RegSurCaf) or
      ((RegSurCaf) and not (DS.State in [dsInsert] ) ) then
    begin
      TobReg := TOB.Create ('Liste code reg', nil, -1) ;
      QQ := OpenSQL (stSQL, TRUE) ;
      if not QQ.EOF then
      begin
        TobReg.LoadDetailDB ('Regroupe aff', '', '', QQ, False) ;
        bAutresAff := True;
        for i := 0 to TobReg.Detail.Count - 1 do
        begin
          Reg := TobReg.Detail [i] .GetValue ('AFF_REGROUPEFACT') ;
          if (GetField ('AFF_REGROUPEFACT') = 'AUC') and (Reg <> 'AUC') then
          begin
            RegRegroupe := RegRegroupe + Reg + ' - ';
            bRegroupe := true;
          end;
          if Reg = 'AUC' then
            bAucun := true;
        end;
      end;
      Ferme (QQ) ;
      TobReg.Free; //TobReg := NIL; //GISE
      if bRegroupe then
      begin
        if not (Regsurcaf) then
          PGIBoxAF ('Attention contrat non regroupé alors que ce tiers possède des contrats regroupés sur le(s) code(s) : ' + RegRegroupe, 'Gestion des contrats')
        else
          PGIBoxAF ('Attention contrat non regroupé alors que ce Tiers à facturer ' + Lecaf + ' possède des contrats regroupés sur le(s) code(s) : ' + RegRegroupe, 'Gestion des contrats') ;
      end;
      if bAucun then
      begin
        if not (Regsurcaf) then
          PGIBoxAF ('Attention ce tiers possède des contrats non regroupés en facturation', 'Gestion des contrats')
        else
          PGIBoxAF ('Attention ce tiers à facturer ' + LeCaf + ' possède des contrats non regroupés en facturation', 'Gestion des contrats') ;
      end;

      if (DS.State in [dsInsert] ) and not (bAutresAff) then
      begin
        // pas d'autres affaires positionne l'affaire en principale et regroupement
        SetField ('AFF_PRINCIPALE', 'X') ;
        if GetField ('AFF_AFFAIRE1') = 'ASS' then
          SetField ('AFF_REGROUPEFACT', 'AS1') ;
        if GetField ('AFF_AFFAIRE1') = 'MAI' then
          SetField ('AFF_REGROUPEFACT', 'MA1') ;
      end;
    end;
  end;

  if (EnSaisie) then
  begin
    if GetField ('AFF_DATERESIL') <> iDate2099 then
    begin
      SetField ('AFF_DATEFIN', GetField ('AFF_DATERESIL') ) ;
      SetField ('AFF_DATEFINGENER', GetField ('AFF_DATERESIL') ) ;
    end else
    begin
      // valeur du contrat identique à celle de facturation
      SetField ('AFF_DATEDEBUT', GetField ('AFF_DATEDEBGENER') ) ;
      SetField ('AFF_DATEFIN', GetField ('AFF_DATEFINGENER') ) ;
    end;

  end; // En saisie
end;

//------------------------------------------------------------------------------
// ONYX
//------------------------------------------------------------------------------

procedure TOM_AFFAIRE.FormuleOnChange (SEnder: TObject) ;
var
  vSt: string;
begin

{$IFDEF EAGLCLIENT}
  if TEdit (sender) .Name = 'AFF_FORCODE1' then
{$ELSE}
  if THDBEdit (sender) .Name = 'AFF_FORCODE1' then
{$ENDIF}
  begin
    vSt := RechDom ('AFTREVFORMULE', GetControlText ('AFF_FORCODE1') , False) ;
    if vSt <> 'Error' then
      SetControlText ('LIBELLEFORMULE1', vSt)
    else
      SetControlText ('LIBELLEFORMULE1', '') ;
  end
  else
  begin
    vSt := RechDom ('AFTREVFORMULE', GetControlText ('AFF_FORCODE2') , False) ;
    if vSt <> 'Error' then
      SetControlText ('LIBELLEFORMULE2', vSt)
    else
      SetControlText ('LIBELLEFORMULE2', '') ;
  end;
  GestionMenusRevision;
end;

procedure TOM_AFFAIRE.FormuleOnExit (SEnder: TObject) ;
var
  vSt: string;
begin
{$IFDEF EAGLCLIENT}
  if THEdit (sender) .Name = 'AFF_FORCODE1' then
{$ELSE}
  if THDBEdit (sender) .Name = 'AFF_FORCODE1' then
{$ENDIF}
  begin
    vSt := RechDom ('AFTREVFORMULE', GetControlText ('AFF_FORCODE1') , False) ;
    if (GetControlText ('AFF_FORCODE1') <> '') and (vSt = 'Error') then
      PGIBoxAf (TexteMsgAffaire [40] , '') ;
  end
  else
  begin
    vSt := RechDom ('AFTREVFORMULE', GetControlText ('AFF_FORCODE2') , False) ;
    if (GetControlText ('AFF_FORCODE2') <> '') and (vSt = 'Error') then
      PGIBoxAf (TexteMsgAffaire [40] , '') ;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 22/04/2003
Modifié le ... :   /  /
Description .. : Calcul des coefficients
Mots clefs ... :
*****************************************************************}

procedure TOM_AFFAIRE.mnCalculOnClick (Sender: TObject) ;
var
  vSt: string;
  vQR: TQuery;
  vRev: TCALCULCOEF;
  vStFormule: string;
  vDtDate: TDateTime;

begin

  vDtDate := iDate2099;
  if not ToutesFormulesParam (GetControlText ('AFF_AFFAIRE') ) then
    PGIBoxAF (TexteMsgAffaire [43] , '')
  else
  begin
    // controle si il y a une facture sur la première formule de l'affaire
    vSt := 'SELECT AFC_FORCODE, AFC_LASTDATEAPP FROM AFPARAMFORMULE WHERE AFC_AFFAIRE = "' + GetControlText ('AFF_AFFAIRE') + '"';
    vQr := nil;
    try
      vQR := OpenSql (vSt, True) ;
      if not vQR.Eof then
      begin
        vStFormule := vQr.findfield ('AFC_FORCODE') .asString;
        vDtDate := vQr.findfield ('AFC_LASTDATEAPP') .AsDateTime;
      end
      else
        vStFormule := '';
    finally
      Ferme (vQr) ;
    end;

    if (vDtDate = iDate2099) or (not BlocageNextCalcul (vStFormule) ) then
    begin
      vRev := TCALCULCOEF.Create;
      try
        // Appel de l'ecran en passant l'affaire
        AFLanceFiche_CalcFormule ('', '', 'AFFAIRE=' + GetControlText ('AFF_AFFAIRE') ) ;
        try
          if assigned (theTob) then
          begin // recuperation d'une tob
            BeginTrans;
            if vRev.CalculCoefFacturation (GetControlText ('AFF_AFFAIRE') , theTob, nil, nil, False) then
            try
              CommitTrans;
              PgiInfoAf (TexteMsgAffaire [38] , '') ;
            except
              RollBack;
              PGIBoxAF (TexteMsgAffaire [41] , '') ;
            end
            else
              PGIBoxAF (TexteMsgAffaire [41] , '') ;
          end;
        finally
          theTob.Free;
          theTob := nil;
        end;
      finally
        vRev.Free;
      end;
    end
    else
      PGIBoxAF (format (TexteMsgAffaire [44] , [DateToStr (vDtDate) ] ) , Ecran.Caption) ;
  end;
end;

// on autorise les menus en fonction de la présence ou non de formules

procedure TOM_Affaire.GestionMenusRevision;
var
  i: Integer;
  vBoMenus: Boolean;
begin
  if TpopUpMenu (GetControl ('POPMENU1') ) = nil then
    exit;

  vBoMenus := existeSql ('SELECT AFC_AFFAIRE FROM AFPARAMFORMULE where AFC_AFFAIRE ="' + GetControlText ('AFF_AFFAIRE') + '"') ;
  if vBoMenus then
    for i := 1 to 6 do
      TpopUpMenu (GetControl ('POPMENU1') ) .Items [i] .Enabled := true
  else
    for i := 1 to 6 do
      TpopUpMenu (GetControl ('POPMENU1') ) .Items [i] .Enabled := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 21/07/2003
Modifié le ... :   /  /
Description .. : on ne peut pas changer la formule d'une affaire
                 si des coefficients ont été appliqués
Mots clefs ... :
*****************************************************************}

procedure TOM_Affaire.CoefAppliques;
var
  vSt: string;
  vQr: TQuery;
  vTob: Tob;
  i: Integer;
begin

  vSt := 'SELECT DISTINCT AFR_FORCODE FROM AFREVISION ';
  vSt := vSt + ' WHERE AFR_AFFAIRE = "' + GetControlText ('AFF_AFFAIRE') + '"';
  vSt := vSt + ' AND AFR_OKCOEFAPPLIQUE = "X"';

  vQr := nil;
  vTob := TOB.create ('AFEVISION', nil, -1) ;
  try
    vQr := OpenSQL (vSt, True) ;
    if not vQr.EOF then
    begin
      vTob.LoadDetailDB ('AFREVISION', '', '', vQr, false, true) ;
      for i := 0 to vTob.detail.count - 1 do
      begin
        if GetControlText ('AFF_FORCODE1') = vTob.Detail [i] .Getvalue ('AFR_FORCODE') then
          SetControlEnabled ('AFF_FORCODE1', false)
        else
          if GetControlText ('AFF_FORCODE2') = vTob.Detail [i] .Getvalue ('AFR_FORCODE') then
          SetControlEnabled ('AFF_FORCODE2', false) ;
      end;
    end;
  finally
    Ferme (vQr) ;
    vTob.Free;
  end;
end;

{procedure TOM_AFFAIRE.mnElementsOnClick(Sender: TObject);
Var
  Clause_where : string;

begin
  Clause_where:=' WHERE AFR_AFFAIRE ="'+GetControltext('AFF_AFFAIRE')+'" '  ;
  AglLanceFicheAFREVISION('','Clause_where='+Clause_where+';AFFAIRE;ACTION=MODIFICATION') ;
end ;
}

procedure TOM_AFFAIRE.mnFormuleOnClick (Sender: TObject) ;
var
  combien: integer;
begin
  combien := VerifieCoherence (GetControltext ('AFF_AFFAIRE') , false) ;
  if (combien > 0) then
    AFLanceFiche_ParamFormule (GetControltext ('AFF_AFFAIRE') , '', 'ACTION=MODIFICATION')
  else
    PgiinfoAF (TexteMsgAffaire [39] , Ecran.Caption) ;
  GestionMenusRevision;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 03/06/2002
Modifié le ... :   /  /
Description .. : Appliquer les coefficients calculés par formule
Mots clefs ... :
*****************************************************************}

procedure TOM_AFFAIRE.mnAppliquerCoefOnClick (Sender: TObject) ;
begin
  if TestTraitement (GetControltext ('AFF_AFFAIRE') , True) then
    AFLanceFiche_TraitFormule ('', '', 'AFFAIRE=' + GetControltext ('AFF_AFFAIRE') + ';ACTION=APPLIQUER') ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 03/06/2002
Modifié le ... :   /  /
Description .. : Desappliquer les coefficients calculés par formule
Mots clefs ... :
*****************************************************************}

procedure TOM_AFFAIRE.mnDesappliquerCoefOnClick (Sender: TObject) ;
begin
  if TestTraitement (GetControltext ('AFF_AFFAIRE') , False) then
    AFLanceFiche_TraitFormule ('', '', 'AFFAIRE=' + GetControltext ('AFF_AFFAIRE') + ';ACTION=DESAPPLIQUER') ;
end;

procedure TOM_AFFAIRE.mnRetablirCoefOnClick (Sender: TObject) ;
begin
  AFLanceFiche_TraitFormule ('', '', 'AFFAIRE=' + GetControltext ('AFF_AFFAIRE') + ';ACTION=RETABLIR') ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 16/10/2003
Modifié le ... :   /  /
Description .. : blocage du calcul du prochain coefficient si
                 aucune factures pour le précédent coefficient
Mots clefs ... :
*****************************************************************}

function TOM_AFFAIRE.BlocageNextCalcul (pStFormule: string) : Boolean;
var
  vSt: string;
  vQr: TQuery;
  vTob: Tob;
  vBoProchain: Boolean;
  vDateLast: TDateTime;

begin

  // determination si prochain coefficient à calculer ou non
  vSt := 'SELECT MAX(AFR_DATECALCCOEF) AS MADATE, AFR_OKCOEFAPPLIQUE FROM AFREVISION ';
  vSt := vSt + ' WHERE AFR_AFFAIRE = "' + GetControltext ('AFF_AFFAIRE') + '"';
  vSt := vSt + ' AND AFR_FORCODE = "' + pStformule + '" AND AFR_COEFREGUL = "-" ';
  vSt := vSt + ' GROUP BY AFR_OKCOEFAPPLIQUE ';
  vQr := nil;
  vTob := TOB.Create ('AFREVISION', nil, -1) ;
  try
    vQr := OpenSQL (vSt, TRUE) ;
    vTob.LoadDetailDB ('', '', '', vQr, false) ;

    if (vTob.Detail.Count > 1) and
      (vTob.detail [0] .GetValue ('MADATE') = vTob.detail [1] .GetValue ('MADATE') ) then
      // on a les deux cas
      vBoProchain := False
    else
    begin
      if vTob.Detail.Count > 0 then
      begin
        if vTob.Detail [0] .getvalue ('AFR_OKCOEFAPPLIQUE') = 'X' then
          vBoProchain := True
        else
          vBoProchain := False;
      end
      else
        vBoProchain := True;
    end;

  finally
    Ferme (vQr) ;
    vTob.free;
  end;

  // determination si il y a une facture
  result := False;
  if vBoProchain then
  begin

    // lecture de la derniere date
    vSt := 'SELECT AFC_LASTDATEAPP ';
    vSt := vSt + ' FROM AFPARAMFORMULE ';
    vSt := vSt + ' WHERE  AFC_AFFAIRE = "' + GetControltext ('AFF_AFFAIRE') + '"';
    vSt := vSt + ' AND  AFC_FORCODE = "' + pStFormule + '"';

    vQr := nil;
    try
      vQr := OpenSQL (vSt, TRUE) ;
      vDateLast := vQr.FindField ('AFC_LASTDATEAPP') .AsDateTime;
    finally
      Ferme (vQr) ;
    end;

    vSt := ' SELECT COUNT(*) AS NB FROM AFREVISION ';
    vSt := vSt + 'WHERE AFR_NATUREPIECEG = "FAC" ';
    vSt := vSt + 'AND AFR_DATECALCCOEF = "' + usDateTime (vDateLast) + '"';
    vSt := vSt + 'AND AFR_AFFAIRE = "' + GetControltext ('AFF_AFFAIRE') + '"';

    vQr := OpenSQL (vSt, TRUE) ;
    try
      result := vQr.FindField ('NB') .AsInteger = 0;
    finally
      Ferme (vQr) ;
    end;
  end;
end;

// *****************************************************************************
// ************** Fonction appel isoflex spécifique GA *************************
// *****************************************************************************

procedure TOM_AFFAIRE.GereIsoflex;
var
  bIso: Boolean;
  MenuIso: TMenuItem;
begin
  bIso := False;
  MenuIso := TMenuItem (GetControl ('mnSGED') ) ;
{$IFNDEF EAGLCLIENT}
  // mc d24/04/02 if (ctxTempo in v_PGI.PGIContexte) or (VH_GC.GASeria) then bIso := AglIsoflexPresent;
//  if (not (ctxScot in v_PGI.PGIContexte) ) or (VH_GC.GASeria) then
    bIso := AglIsoflexPresent;
{$ENDIF}
  if MenuIso <> nil then
    MenuIso.Visible := bIso;

end;

procedure ImprimeThisScreen (parms : array of Variant; nb : Integer);
var F: TForm;
begin
  F := TForm (Longint (Parms [0] ) ) ;
	PrintThisScreen(F,F.Caption);
end;

procedure AffaireAppelIsoFlex (parms: array of variant; nb: integer) ;
{$IFNDEF EAGLCLIENT}
var
  F: TForm;
  Affaire: string;
{$ENDIF}
begin
{$IFNDEF EAGLCLIENT}
  F := TForm (Longint (Parms [0] ) ) ;
  if (F.Name <> 'BTAFFAIRE') then
    exit;
  Affaire := string (Parms [1] ) ;
  AglIsoflexViewDoc (NomHalley, F.Name, 'AFFAIRE', 'T_CLE1', 'AFF_AFFAIRE', Affaire, '') ;
{$ENDIF}
end;

// *****************************************************************************
// ************** Fonction AGL de la saisie d'affaire **************************
// *****************************************************************************

// ******* Calcul des dates de la mission en fonction de l'exercice ************

procedure AGLSaisiePieceAffaire (parms: array of variant; nb: integer) ;
var
  F: TForm;
  OM: TOM;
begin
  F := TForm (Longint (Parms [0] ) ) ;
  if (F is TFFiche) then
    OM := TFFiche (F) .OM
  else
    exit;
  if (OM is TOM_Affaire) then
  begin
  	V_PGI.ZoomOLE := true;
    TOM_Affaire (OM) .AppelPieceAffaire (Parms [1] );
  	V_PGI.ZoomOLE := false;
  end else
    exit;
end;

// Fiche Affaire calcul automatique des dates d'écheances et lance les echéances

procedure AGLLanceEcheance (parms: array of variant; nb: integer) ;
var
  F: TForm;
  OM: TOM;
  Parle, ForceCalcul, Boutton: string;
begin
  F := TForm (Longint (Parms [0] ) ) ;
  if (F is TFFiche) then
    OM := TFFiche (F) .OM
  else
    exit;
  ForceCalcul := string (Parms [1] ) ;
  Parle := string (Parms [2] ) ;
  Boutton := string (Parms [3] ) ;
  if (OM is TOM_Affaire) then
    TOM_Affaire (OM) .LanceEcheance (ForceCalcul, Parle, Boutton)
  else
    exit;
end;

procedure AGLCalculDatesMission (parms: array of variant; nb: integer) ;
var
  F: TForm;
  OM: TOM;
begin
  F := TForm (Longint (Parms [0] ) ) ;
  if (F is TFFiche) then
    OM := TFFiche (F) .OM
  else
    OM := nil;
  if (OM is TOM_Affaire) then
    TOM_Affaire (OM) .CalculDatesMission;
end;

procedure AGLCalculMontantGlobal (parms: array of variant; nb: integer) ;
var
  F: TForm;
  OM: TOM;
begin
  F := TForm (Longint (Parms [0] ) ) ;
  if (F is TFFiche) then
    OM := TFFiche (F) .OM
  else
    OM := nil;
  if (OM is TOM_Affaire) then
    TOM_Affaire (OM) .CalculMontantGlobal (false, false) ;
end;

procedure AGLLanceAffTiers (parms: array of variant; nb: integer) ;
var
  F: TForm;
  OM: TOM;
begin
  F := TForm (Longint (Parms [0] ) ) ;
  if (F is TFFiche) then
    OM := TFFiche (F) .OM
  else
    OM := nil;
  if (OM is TOM_Affaire) then
    TOM_Affaire (OM) .LanceAffTiers (Parms [1] ) ;
end;

procedure AGLImprimeModele (parms: array of variant; nb: integer) ;
var
  F: TForm;
  OM: TOM;
begin
  F := TForm (Longint (Parms [0] ) ) ;
  if (F is TFFiche) then
    OM := TFFiche (F) .OM
  else
    OM := nil;
  if (OM is TOM_Affaire) then
    TOM_Affaire (OM) .ImprimeModele (Parms [1] , F) ;
end;

procedure AGLLanceRegroupeAffaire (parms: array of variant; nb: integer) ;
var
  F: TForm;
  OM: TOM;
begin
  F := TForm (Longint (Parms [0] ) ) ;
  if (F is TFFiche) then
    OM := TFFiche (F) .OM
  else
    OM := nil;
  if (OM is TOM_Affaire) then
    TOM_Affaire (OM) .LanceRegroupeAffaire;
end;

function AGLModifAutorisee (parms: array of variant; nb: integer) : variant;
var
  F: TForm;
  OM: TOM;
begin
  Result := '';
  F := TForm (Longint (Parms [0] ) ) ;
  if (F is TFFiche) then
    OM := TFFiche (F) .OM
  else
    OM := nil;
  if (OM is TOM_Affaire) then
    Result := ActionAffaireString (TOM_Affaire (OM) .ModifAutorisee, TFFiche (F) .typeaction) ;
end;

procedure AGLNomContactAff (parms: array of variant; nb: integer) ;
var
  F: TForm;
  OM: TOM;
begin
  F := TForm (Longint (Parms [0] ) ) ;
  if (F is TFFiche) then
    OM := TFFiche (F) .OM
  else
    exit;
  if (OM is TOM_AFFAIRE) then
    TOM_AFFAIRE (OM) .NomContactAff
  else
    exit;
end;

procedure AFLanceFiche_Affaire (lequel, argument: string) ;
begin
  if ctxScot in V_PGI.PGIContexte then
    AGLLanceFiche ('AFF', 'MISSION', '', Lequel, Argument)
  else
    AGLLanceFiche ('AFF', 'AFFAIRE', '', Lequel, Argument) ;
end;

procedure AFLanceFiche_LienAffaireTiers (Range, lequel, argument: string) ;
begin
  AGLLanceFiche ('AFF', 'LIENAFFTIERS', Range, lequel, Argument) ;
end;

procedure TOM_Affaire.AfterFormShow;
begin
//
  TToolbarButton97(GetControl('BLast')).visible := false;
  TToolbarButton97(GetControl('BFirst')).visible := false;
  TToolbarButton97(GetControl('BNext')).visible := false;
  TToolbarButton97(GetControl('BPrev')).visible := false;
//
	THCheckBox(GetControl('CBCOTRAITANCE')).OnClick := CBCOTRAITANCEClick;
  TRadioButton (getControl('RBMANDATAIRE')).OnClick := RBMANDATAIREClick;
  TRadioButton (getControl('RBCOTRAITANTE')).OnClick := RBCOTRAITANTEClick;
end;

//******************************************************************************
//********************** Calcul des echéances **********************************
// La fonction retourne 1 si on peut accéder au factaff et 2 dans le cas contraire
//******************************************************************************
function TOM_Affaire.CreerEcheance (ForceCalcul, Parle, Boutton: string) : integer;
var
  TypeCalcul    : String;
  Choix         : string;
  DateDebut     : TDateTime;
  DateFin       : TDateTime;
  DateLiquid    : TDateTime;
  DateResil     : TDateTime;
  DateDebCal    : TDateTime;
  NbIt          : Integer;
  TypeModif     : T_TypeModifAff; // T_TypeModifAff = (tmaMntEch,tmaPourcEch,tmaDate,tmaDFacLiq) ;
  zmont         : double;
  Calcul        : Boolean;
  bMtPeriodique : Boolean;
  bPourcentage  : Boolean;
  bRecalculOK   : Boolean;
  Res           : TIOErr;
	ModeMontantDoc,MethodeCalcul : string;
  DateDebutfac,DateFinfac : TDateTime;
  TermeEche : string;
  mtTotal : double;
  ProfilGener : string;
  Affaire : string;
  Activrepris :string;
  Tiers :string;
  GenerAuto : string;
  Intervalgener : integer;

begin
  bRecalculOK := False;
  result := 0;
  DateDebCal := iDate1900;
  bMtPeriodique := true;
  Choix := '0';

  if (GetField ('AFF_ETATAFFAIRE') = 'CLO') or (GetField ('AFF_ADMINISTRATIF') = 'X') then
    Exit;

  // PLGM le 26/06/02 : pb de definition de la date de fin de generation des echeances de facturation
  if VH_GC.GCIfDefCEGID then // Attention pas de génération des écheances antérieures à la reprise ...
    if GetField ('AFF_DATECREATION') < strToDate ('15/08/2001') then
    begin // si modif voir ModifAvectrait de UtofAffair_ModifLot
      DateDebCal := StrToDate ('15/08/2001') ;
    end;
  // Fin PLGM le 26/06/02

  if (GetParamSoc ('SO_AFALIMECHE') = true) then
    // PL le 19/06/02 : echeances à partir du budget
  begin
    if (TFfiche (Ecran) .typeaction <> TaConsult) then
    begin
      if ExisteEcheanceAffaire (GetField ('AFF_AFFAIRE') ) then
        Calcul := False
      else
        Calcul := True;

      if not ExisteEcheanceBudget (GetField ('AFF_AFFAIRE') ) then
       begin
        PGIInfoAF ('Attention, il n''existe pas de budget pour cette affaire.' + chr (13) + 'Aucune échéance ne sera générée.', Ecran.caption) ;
        exit;
      end;

      if (ForceCalcul = 'oui') then Calcul := true;

      if (Calcul = false) then
        if (PGIAskAF ('Voulez-vous re-générer les échéances à partir du budget ?' + chr (13) + 'Attention, ce traitement va supprimer les anciennes échéances non encore facturées.', Ecran.caption) = mrYes) then
          Calcul := true;

      if Calcul then
      begin
        Res := Transactions (CreerEcheanceDuBudget, 2) ;
        case Res of
          oeOk: bRecalculOK := true;
          oeUnknown, oeSaisie:
            begin
              PGIBoxAf ('Une erreur est survenue lors de la création des échéances.' + chr (13) + 'Elles n''ont pas été générées.', Ecran.Caption) ;
              Exit;
            end;
        end; // case
      end; // Calcul
    end; // (TFfiche(Ecran).typeaction <> TaConsult )
  end // (GetParamSoc('SO_AFALIMECHE')= true)
  else
    // Alimentation normale des echeances (pas par le budget)
  begin
    ////////////////////
    Result := TestTypeModifEcheances(DateDebut, DateFin, DateLiquid, DateResil, TypeCalcul, zmont, TypeModif, Boutton, false);
    if (not ExisteEcheanceAffaire (GetField ('AFF_AFFAIRE'),'-')) then TypeModif := [tmaMntEch] ;

    if (tmaRepAct in TypeModif) or (tmaMntEch in TypeModif) or (tmaPourcEch in TypeModif) or (tmaDate in TypeModif) or (tmaDFacLiq in TypeModif) then
    begin
      if ExisteEcheanceAffaire (GetField ('AFF_AFFAIRE') ) then
        Calcul := False
      else
        Calcul := True;
      if (TypeModif = [tmaRepAct] ) then
        Calcul := True;
      if (TypeModif = [tmaDfacLiq] ) then
        Calcul := True; //mcd 13/11/02
      bPourcentage := ((GetField ('AFF_GENERAUTO') = 'POU') or (GetField ('AFF_GENERAUTO') = 'POT') ) ;

    	if (not ExisteEcheanceAffaire (GetField ('AFF_AFFAIRE'),'-')) then calcul := false;

      if not Calcul then
      begin
        // on veut renvoyer la réponse : modif PL le 26/06/2000
        if (Parle = 'oui') then
        begin
          Choix := LanceChoixRecalcul (bPourcentage) ; // PA externalisé car exception pour cegid ...

          if (Choix <> '0') then
          begin
            if (Choix = '1') then
            begin
              // Montant periodique sur chaque echeance
              bMtPeriodique := true;
              Calcul := true;
            end
            else
              if (Choix = '2') then
            begin
              Calcul := true;
              // "Pourcentage" : montant reparti sur chaque echeance
              if (GetField ('AFF_GENERAUTO') = 'POU') then
              begin
                if (GetField ('AFF_POURCENTAGE') = 0) then
                  // on considere qu'on veut repartir sur chaque echeance avec arrondi sur la derniere
                begin
                  NbIt := EvaluationNbEcheances (GetField ('AFF_AFFAIRE') , TypeCalcul,
                    Integer (GetField ('AFF_INTERVALGENER') ) , DateDebut, DateFin) ;
                  if (NbIt <> 0) then
                    zmont := Arrondi (100 / NbIt, V_PGI.OkDecV) ;

                  bMtPeriodique := false;
                end
                else
                  // on affecte le pourcentage saisi sur chaque echeance
                begin
                  zmont := double (GetField ('AFF_POURCENTAGE') ) ;
                  bMtPeriodique := true;
                end;
              end
              else
                // "Facturation des lignes d'affaires" : meme montant sur toutes les echeances
                if (GetField ('AFF_GENERAUTO') = 'POT') then
              begin
                zmont := double (GetField ('AFF_POURCENTAGE') ) ;
                bMtPeriodique := true;
              end
              else
                // Montant calculé a partir du montant global reparti sur chaque echeance
(*                          if Not bPourcentage then*)
              begin
                CalculMontantGlobal (true, bPourcentage) ;
                zmont := double (GetField ('AFF_MONTANTECHEDEV') ) ;
                bMtPeriodique := false;
              end;
            end
            else
              if (Choix = '3') then
            begin
              Calcul := false;
            end;
          end // (Choix<>'0')
          else
          begin
            Calcul := false;
            Result := 2;
          end;
        end; //(Parle='oui')
      end; // Not Calcul

      if (ForceCalcul = 'oui') and (Result = 1) then
        Calcul := true;

      if Calcul then
      begin
      	TermeEche := THValComboBox(GetControl('AFF_TERMEECHEANCE')).value;
        ModeMontantDoc := THValComboBox(GetControl('AFF_DETECHEANCE')).value;
        MethodeCalcul := THValComboBox(GetControl('AFF_METHECHEANCE')).value;
        DateDebutfac := GetField('AFF_DATEDEBGENER');
        DateFinfac := GetField('AFF_DATEFINGENER');

        Mttotal := GetField ('AFF_TOTALHTGLODEV');
        ProfilGener := GetField ('AFF_PROFILGENER');
        Affaire := GetField ('AFF_AFFAIRE');
        Activrepris := GetField ('AFF_REPRISEACTIV');
        Tiers := GetField ('AFF_TIERS');
        GenerAuto := GetField ('AFF_GENERAUTO');
        Intervalgener := GetField ('AFF_INTERVALGENER');

        UtilCalculEcheances (Mttotal,ProfilGener,TypeCalcul,Affaire ,
                             Activrepris ,
                             Tiers , GenerAuto ,TypeModif,
                             Intervalgener , zmont,
                             DateDebut, DateFin, DateLiquid, DateDebCal,
                             idate2099, DateResil,DEV, GetContre, bMtPeriodique, bPourcentage,
                             GetField ('AFF_MULTIECHE'),ModeMontantDoc,MethodeCalcul,DateDebutFac,DateFinfac,TermeEche ) ;
        bRecalculOK := true;
        // restockages des valeurs de calcul des écheances après génération
    		TermeEche_Sav := GetField ('AFF_TERMEECHEANCE') ;
        DateDebGener_sav := StrToDate (GetField ('AFF_DATEDEBGENER') ) ;
        DateFinGener_sav := StrToDate (GetField ('AFF_DATEFINGENER') ) ;
        DateLiquidGener_sav := StrToDate (GetField ('AFF_DATEFACTLIQUID') ) ;
        DateResil_sav := StrToDate (GetField ('AFF_DATERESIL') ) ;

        Periodicite := GetField ('AFF_PERIODICITE') ;
        TypeGenerAuto := GetField ('AFF_GENERAUTO') ;
        Interval := GetField ('AFF_INTERVALGENER') ;
        MontantEch := GetField ('AFF_MONTANTECHEDEV') ;
        PourcentEch := GetField ('AFF_POURCENTAGE') ;
        MethEcheance := GetField ('AFF_METHECHEANCE') ;
        ActiviteReprise := GetField ('AFF_REPRISEACTIV') ;
        EtatAffairePourech := GetField ('AFF_ETATAFFAIRE') ; // gm 15/11/02
        MultiEche := GetField ('AFF_MULTIECHE') ;
        StProfil := GetField ('AFF_PROFILGENER') ;

        if (Ensaisie) then
          ImpactCalculTotalHT (TitEcheances) ; //mcd 06/10/03 fiche 10420 pour reforcer dans ce cas, le cacul global
      end;
    end; // (tmaRepAct in TypeModif)
  end;

  if not bRecalculOK then
    // si pas de recalcul on vérifie qu'il n'y a pas eu de modif francs <=> Euro
  begin
    if Boutton <> 'BEcheance' then
      ImpactCalculTotalHT (TitEcheances) ;
    if (SaisieContre <> GetContre) then
      InverseEcheanceContre (GetField ('AFF_AFFAIRE') , GetContre) ;
    EchContre := GetContre;
  end;

end;

procedure TOM_Affaire.DefiniLibelle;
var Memo: TMemo;
    lng : integer;
    Indice : integer;
    nbCars : Integer;
begin
  Memo := TMemo (GetControl ('AFF_DESCRIPTIF') ) ;
  if (CompareStr (LibAff, GetField ('AFF_LIBELLE') ) = 0) or (GetField ('AFF_LIBELLE')='') then
  begin
  	libaff := '';
    Setfield('AFF_LIBELLE','');
    // trouve la premiere ligne non vide
    Indice := 0;
    repeat
    	lng := Pos (#$D#$A,Memo.lines [indice] ) ; nbCars := Length(Trim(Memo.lines [indice]));
      if (lng > 1) or (nbCars > 0) then break else inc(indice);
    until indice > Memo.lines.count;

    if Indice > Memo.lines.count then exit;
    lng := length(Trim(Memo.lines [Indice]))+1;
    if (lng > 70) then
    begin
      LibAff := Copy ( Trim(Memo.lines [Indice]) , 0, 70) ;
      SetField ('AFF_LIBELLE', Copy (Trim(Memo.lines [Indice]) , 0, 70) ) ;
    end
    else
    begin
      LibAff := Copy (Trim(Memo.lines [indice]) , 0, lng - 1) ;
      SetField ('AFF_LIBELLE', Copy (Trim(Memo.lines [indice]) , 0, lng - 1) ) ;
    end;
  end;
end;

procedure TOM_Affaire.CalculMontantEch;
Var MtDejaFact  : Double;
    MtTotal			: Double;
    MtRestant		: Double;
    QQ					: TQuery;
    SReq				: String;
begin

    MtDejaFact := 0;
    MtTotal		 := 0;
    MtRestant	 :=0;

    // recherche de la somme Totale des échéances
    QQ := nil;
    try
      sReq := 'SELECT AFA_AFFAIRE, SUM(AFA_MONTANTECHEDEV) FROM FACTAFF WHERE AFA_AFFAIRE="' + GetField ('AFF_AFFAIRE') + '" AND AFA_TYPECHE="NOR" AND AFA_LIQUIDATIVE="-" GROUP BY AFA_AFFAIRE';
      QQ := OpenSQL (sReq, true) ;
      if not QQ.EOF then
        MtTotal := QQ.Fields[1].AsFloat;
    finally
      Ferme (QQ) ;
    end;

    // recherche de la somme des échéances deja facturées
    QQ := nil;
    try
      sReq := 'SELECT AFA_AFFAIRE, SUM(AFA_MONTANTECHEDEV) FROM FACTAFF WHERE AFA_AFFAIRE="' + GetField ('AFF_AFFAIRE') + '" AND AFA_TYPECHE="NOR" AND AFA_ECHEFACT="X" AND AFA_LIQUIDATIVE="-" GROUP BY AFA_AFFAIRE';
      QQ := OpenSQL (sReq, true) ;
      if not QQ.EOF then
        MtDejaFact := QQ.Fields[1].AsFloat;
    finally
      Ferme (QQ) ;
    end;

    MtRestant := MtTotal - MtDejaFact;

    THNumEdit(getControl('MONTANTGLOBAL')).Value   := MtTotal;
	  THNumEdit(getControl('MONTANTDEJAFACT')).value := MtDejaFact;
    THNumEdit(getControl('MONTANTAFACT')).Value    := MtRestant;
		CalculEchSurDoc(true, false);
end;

procedure TOM_Affaire.VoirChantier(Sender: TOBJect);
Var Argument   : string;
		C0, C1, C2, C3 : THCritMaskEdit;
    Avenant, Tiers : THCritMaskEdit;
    CodeChantier   : String;
begin

  CodeChantier := GetControlText('AFF_CHANTIER');

  SetControlText('CAF_CHANTIER0', 'A');

	C0 := THEDIT(GetControl('CAF_CHANTIER0'));
	C1 := THEDIT(GetControl('AFF_CHANTIER1'));
	C2 := THEDIT(GetControl('AFF_CHANTIER2'));
	C3 := THEDIT(GetControl('AFF_CHANTIER3'));
	Avenant := THEDIT(GetControl('CHA_AVENANT'));
	Tiers := THEDIT(GetControl('AFF_TIERS'));

  if not GetAffaireEnteteSt( C0,C1,C2,C3,Avenant, Tiers, CodeChantier,false,false,false,false,false) then
     CodeChantier := '';

  if CodeChantier = '' then exit;

  LectureChantier(CodeChantier);

end;

procedure TOM_Affaire.LectureChantier(CodeChantier : String);
Var TobChantier: TOB;
    Req        : String;
    A1         : String;
    A2         : String;
    A3         : String;
Begin

  Req := '';

  Req := 'SELECT * FROM AFFAIRE ';
  Req := Req + 'WHERE AFF_AFFAIRE ="' + CodeChantier + '"';

  if GetControlText('AFF_TIERS') <> '' then
	  Req := Req + 'AND AFF_TIERS ="' + GetControlText('AFF_TIERS') + '"';

  TobChantier := Tob.Create('LesChantiers',Nil, -1);
  TobChantier.LoadDetailDBFromSQL('AFFAIRE',req,false);

  if TobChantier.Detail.Count <> 0 then
     Begin
     if not (DS.State in [dsInsert, dsEdit] ) then DS.edit;
     
     A1 := TobChantier.Detail[0].GetValue('AFF_AFFAIRE1');
     A2 := TobChantier.Detail[0].GetValue('AFF_AFFAIRE2');
     A3 := TobChantier.Detail[0].GetValue('AFF_AFFAIRE3');
     SetControlText('AFF_CHANTIER', TobChantier.Detail[0].GetValue('AFF_AFFAIRE'));
     SetControlText('CAF_CHANTIER0', TobChantier.Detail[0].GetValue('AFF_AFFAIRE0'));
     SetControlText('AFF_CHANTIER1', A1);
     SetControlText('AFF_CHANTIER2', A2);
     SetControlText('AFF_CHANTIER3', A3);
     SetControlText('CHA_AVENANT', TobChantier.Detail[0].GetValue('AFF_AVENANT'));
     end;

  TobChantier.free;

End;


procedure TOM_Affaire.BTTacheClick(Sender: TObject);
begin

  AGLLanceFiche ('BTP', 'BTTACHES','','', 'BYAFFAIRE;ATA_AFFAIRE:'+StAffaire.text+';ATA_TIERS:'+ StTiers.text);

end;

procedure TOM_Affaire.PositionneEtabUser(NomChamp: string);
Var Etab : String ;
    Forcer : Boolean ;
    CC :THDbValComboBox;
    LastEnabled : boolean;
begin
CC := THDBValCOmboBox (GetCOntrol(NOMChamp));
if CC = nil then exit;
LastEnabled := CC.Enabled ;
if Not VH^.EtablisCpta then Exit ;
if Not CC.Visible then Exit ;
Etab:=VH^.ProfilUserC[prEtablissement].Etablissement ; if Etab='' then Exit ;
Forcer:=VH^.ProfilUserC[prEtablissement].ForceEtab ;
if CC.Values.IndexOf(Etab)<0 then Exit ;
if ((Not CC.Enabled) and (CC.Value<>Etab)) then CC.enabled := true;
SetField(NomChamp,Etab);
if (Forcer) then CC.Enabled:=False ;
end;

procedure TOM_Affaire.ChangeTiersPiece (TOBPiece: TOB; CodeTiers : string);
var TOBTiers,TOBAdresses : TOB;
		QQ : TQuery;
    ReqLig,reqOuv,Exige,stE : string;
    DEV : Rdevise;
    Tcalc : TRecalculPiece;
begin
	V_PGI.IoError := OeOk;

  TOBTiers := TOB.Create ('TIERS',nil,-1);
  TOBAdresses := TOB.Create ('LES ADRESSES',nil,-1);

  QQ := OpenSQL('SELECT * FROM TIERS LEFT JOIN TIERSCOMPL ON YTC_TIERS=T_TIERS WHERE T_TIERS="' + CodeTiers + '"', True);
  TOBTiers.selectDb ('',QQ);
  //
  TiersVersPiece (TOBTiers,TOBPiece);
	TOBPiece.putValue('GP_TVAENCAISSEMENT',PositionneExige(TOBTiers));
  if TOBPiece.GetValue('GP_TVAENCAISSEMENT') = 'TE' then StE := 'X' else StE := '-';

	TOBPiece.putValue('GP_DEVISE',TOBTiers.GetValue('T_DEVISE'));
  DEV.Code := TOBTiers.GetValue('T_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux:= GetTaux(DEV.Code, DEV.DateTaux, TobPiece.GetValue('GP_DATEPIECE'));
  TOBPiece.putValue('GP_TAUXDEV', DEV.Taux);
  TOBPiece.putValue('GP_DATETAUXDEV', DEV.DateTaux);
  //
	LoadLesAdresses (TOBPiece,TOBADRESSES);
  TiersVersAdresses(TOBTiers,TOBAdresses,TOBPiece);
  if not TOBPiece.UpdateDB then V_PGI.IOError := OeUnknown;
  if V_PGI.Ioerror = OeOk then if not TOBAdresses.UpdateDB then V_PGI.IOError := OeUnknown;
  if V_PGI.ioError = OeOk then
  begin
    ReqLig := 'UPDATE LIGNE SET GL_TIERS="'+TOBPIECE.getValue('GP_TIERS')+'", '+
    			 'GL_TIERSFACTURE="'+TOBPIECE.getValue('GP_TIERSFACTURE')+'", '+
           'GL_TIERSPAYEUR="'+TOBPIECE.getValue('GP_TIERSPAYEUR')+'", '+
           'GL_TIERSLIVRE="'+TOBPIECE.getValue('GP_TIERSLIVRE')+'", '+
           'GL_TARIFTIERS="'+TOBPIECE.getValue('GP_TARIFTIERS')+'", '+
           'GL_REGIMETAXE="'+TOBPIECE.getValue('GP_REGIMETAXE')+'", '+
           'GL_FACTUREHT="'+TOBPIECE.getValue('GP_FACTUREHT')+'", '+
           'GL_DEVISE="'+TOBPIECE.getValue('GP_DEVISE')+'", '+
           'GL_TAUXDEV='+IntToStr(TOBPIECE.getValue('GP_TAUXDEV'))+', '+
           'GL_TVAENCAISSEMENT="'+ste+'", '+
           'GL_ESCOMPTE='+INTToStr(TOBPIECE.getValue('GP_ESCOMPTE'))+', '+
           'GL_REMISEPIED='+IntToStr(TOBPIECE.getValue('GP_REMISEPIED'))+''+
           ' WHERE GL_NATUREPIECEG="'+TOBPiece.getValue('GP_NATUREPIECEG')+'"'+
           ' AND GL_SOUCHE="'+TOBPiece.getValue('GP_SOUCHE')+'"'+
           ' AND GL_NUMERO='+InttOsTr(TOBPiece.getValue('GP_NUMERO'))+
           ' AND GL_INDICEG='+IntToStr(TOBPiece.getValue('GP_INDICEG'));
    //
    if ExecuteSql (ReqLig) <= 0 then V_PGI.ioError := OeUnknown;
    if V_PGI.ioError = Oeok then
    begin
      ReqOuv := 'UPDATE LIGNEOUV SET '+
             'BLO_REGIMETAXE="'+TOBPIECE.getValue('GP_REGIMETAXE')+'", '+
             'BLO_FACTUREHT="'+TOBPIECE.getValue('GP_FACTUREHT')+'", '+
             'BLO_DEVISE="'+TOBPIECE.getValue('GP_DEVISE')+'", '+
             'BLO_TAUXDEV='+IntToStr(TOBPIECE.getValue('GP_TAUXDEV'))+', '+
           	 'BLO_TVAENCAISSEMENT="'+ste+'", '+
             'BLO_ESCOMPTE='+INTToStr(TOBPIECE.getValue('GP_ESCOMPTE'))+', '+
             'BLO_REMISEPIED='+IntToStr(TOBPIECE.getValue('GP_REMISEPIED'))+
             ' WHERE BLO_NATUREPIECEG="'+TOBPiece.getValue('GP_NATUREPIECEG')+'"'+
             ' AND BLO_SOUCHE="'+TOBPiece.getValue('GP_SOUCHE')+'"'+
             ' AND BLO_NUMERO='+IntToStr(TOBPiece.getValue('GP_NUMERO'))+
             ' AND BLO_INDICEG='+IntToStr(TOBPiece.getValue('GP_INDICEG'));
			ExecuteSql (ReqOuv);
    end;
    if V_PGI.ioerror = OeOk then
    begin
			if TraitementRecalculPiece (TOBpiece,false,false)<>TrrOk  then V_PGI.ioerror := OeUnknown;
    end;
  end;
  //
  TOBAdresses.free;
  TOBTiers.free;
end;

procedure TOM_Affaire.ChangeTiersAffaire (Sender : TObject);
var TOBR,TOBPIECES : TOB;
		Indice         : integer;
    CodeAffaire    : String;
    NxClient       : String;
    ReqMAJ         : String;
begin
  CodeAffaire := GetField ('AFF_AFFAIRE');

	if ControleDocumentsAffaire (CodeAffaire) then
  begin
		TOBPIECES := TOB.Create ('LES PIECES',nil,-1);
		TOBR := TOB.Create ('UNE TOB',nil,-1);
    TOBR.AddChampSupValeur ('RETOUR','');
    TRY
    	TheTOB := TOBR;
    	AGLLanceFiche('BTP','BTAFFCHANGETIER','','','ACTION=MODIFICATION');
      TheTOB := nil;
      NxClient := TOBR.getValue('RETOUR');

      if NxClient<>'' then
      begin
      	if PgiAsk ('Etes-vous sur(e) de vouloir modifier le client de cette affaire ?') = mryes then
        begin
          TOBPIECES.LoadDetailDBFromSQL ('PIECE','SELECT * FROM PIECE WHERE GP_AFFAIRE="' + CodeAffaire + '" AND GP_VENTEACHAT="VEN"',false,true);
          BEGINTRANS;
          TRY
            for Indice := 0 TO TOBPIECES.Detail.count -1 do
            begin
              ChangeTiersPiece(TOBPIECES.detail[Indice],NxClient);
              if V_PGI.ioerror <> OeOk then break;
            end;
            ///
            ///
            //Traitement de la Mise à jour des échéanciers
            if V_PGI.ioError = OeOk then
            begin
              ReqMAJ := 'UPDATE FACTAFF SET AFA_TIERS="' + NxClient + '"' +
                        ' WHERE AFA_AFFAIRE="'+ CodeAffaire + '"';
              ExecuteSql (ReqMAJ);
            end;
            //Traitement de la Mise à jour des Tâches
            if V_PGI.ioError = OeOk then
            begin
              ReqMAJ := 'UPDATE TACHE SET ATA_TIERS="' + NxClient + '"' +
                        ' WHERE ATA_AFFAIRE="'+ CodeAffaire + '"';
              ExecuteSql (ReqMAJ);
            end;
            ///
          COMMITTRANS;
            if not (DS.State in [dsInsert, dsEdit] ) then DS.edit;
            SetField('AFF_TIERS',TOBR.getValue('RETOUR'));
            TToolbarButton97 (getControl('BVALIDER')).Click ;
          EXCEPT
            ROLLBACK;
          END;
          ///
        end;
      end;
    FINALLY
    	TOBR.free;
      TOBPIECES.free;
    END;
  end;
end;

procedure TOM_Affaire.TiersFactureChange(Sender: Tobject);
var QQ : TQuery;
begin
	if (getControltext('TX_FACTURE')<>OldFacture) then
  begin
    QQ := OpenSql ('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+GetControltext('TX_FACTURE')+'" AND T_NATUREAUXI IN ("CLI","PRO")',true,
    							1,'',true);
    if not QQ.eof then
    begin
      if not(DS.State in [dsInsert,dsEdit]) then
      begin
        DS.edit; // pour passer DS.state en mode dsEdit
      end;
    	SetField('AFF_FACTURE',QQ.findField('T_AUXILIAIRE').asstring);
      OldFacture := getControltext('TX_FACTURE');
    end;
    ferme(QQ);
  end;
end;

procedure TOM_Affaire.Majtaches(DateDebut,DatedeFin : TdateTime);
var TOBTaches : TOB;
		Sql: string;
    Indice : integer;
    DateDebperiod : TDateTime;
    QQ : TQuery;
begin
	TOBTaches := TOB.create ('LES TACHES',nil,-1);
  Sql := 'SELECT * FROM TACHE WHERE ATA_AFFAIRE = "' + GetField ('AFF_AFFAIRE') + '"';
  QQ := OpenSql (Sql,true,-1,'',true);
  TOBTaches.LoadDetailDb('TACHE','','',QQ,false);
  ferme (QQ);
  for Indice := 0 to TOBTaches.detail.count -1 do
  begin
  	if TOBTaches.detail[Indice].getValue('ATA_TERMINE')='X' then
    begin
  		DateDebperiod := incday(Tobtaches.detail[Indice].GetValue('ATA_DATEFINPERIOD'),1);
      if DateDeFin > DateDebperiod then
      begin
        TOBTaches.detail[Indice].putValue('ATA_DATEFINPERIOD',DateDeFin);
        TOBTaches.detail[Indice].putValue('ATA_TERMINE','-');
      end;
    end else
    begin
    	TOBTaches.detail[Indice].putValue('ATA_DATEDEBPERIOD',DateDebut);
    	TOBTaches.detail[Indice].putValue('ATA_DATEFINPERIOD',DateDeFin);
    end;
    TOBTaches.detail[Indice].UpdateDB(false);
  end;
//	executeSql ('UPDATE TACHE set ATA_DATEFINPERIOD = "' + UsDateTime (strtodate (vStDateFin) ) + '" WHERE ATA_AFFAIRE = "' + GetField ('AFF_AFFAIRE') + '"') ;
  TOBTaches.free;
end;

procedure TOM_Affaire.DefiniMenuCotraitance;
begin

	if not VH_GC.SeriaCoTraitance then
  begin
  	THDBCheckbox (GetControl('AFF_MANDATAIRE')).Visible := false;
  	TToolbarButton97  (getControl('MNCOTRAIT')).visible := false;
    if Assigned(GetControl('TS_COTRAITANCE')) then
    begin
    	SetControlVisible('TS_COTRAITANCE', False);
			TTabSheet(GetControl('TS_COTRAITANCE')).TabVisible := false;
    end;
    if Assigned(GetControl('AFF_REFEXTERNE')) then SetControlVisible('AFF_REFEXTERNE', True);
    if Assigned(GetControl('TAFF_REFEXTERNE')) then SetControlVisible('TAFF_REFEXTERNE', True);
    exit;
  end;

	if (StatutAffaire = 'INT') or (AFTypeAction=taCreat) then
  begin
  	//TToolbarButton97  (getControl('MNCOTRAIT')).visible := false;
    GestionCheckMandataire;
    //if Assigned(GetControl('TS_COTRAITANCE')) then SetControlVisible('TS_COTRAITANCE', False);
    //if Assigned(GetControl('AFF_REFEXTERNE')) then SetControlVisible('AFF_REFEXTERNE', True);
    //if Assigned(GetControl('TAFF_REFEXTERNE')) then SetControlVisible('TAFF_REFEXTERNE', True);
  end
  else
    GestionCheckMandataire;

end;

procedure TOM_AFFAIRE.GestionCheckMandataire;
var EtatAffect  : boolean;
    TOBCtrl     : TOB;
begin

  if AFTypeAction=taCreat then
    TToolbarButton97  (getControl('MNCOTRAIT')).visible := False
  else
    TToolbarButton97  (getControl('MNCOTRAIT')).visible := True;

  if EtatAffaire = 'ACP' then
  begin
  	SetControlEnabled('AFF_MANDATAIRE',false);
  	SetControlEnabled('GBCOTRAITANCE',false);
  	SetControlEnabled('CBCOTRAITANCE',false);
  	SetControlEnabled('RBCOTRAITANTE',false);
  	SetControlEnabled('RBMANDATAIRE',false);
  end;

  if THDBCheckbox (GetControl('AFF_MANDATAIRE')).State = cbGrayed then    // Affaire standard (hors co traitance)
  begin
    TToolbarButton97  (getControl('MNCOTRAIT')).visible := false;
    if Assigned(GetControl('TS_COTRAITANCE')) then
    begin
    	SetControlVisible('TS_COTRAITANCE', False);
			TTabSheet(GetControl('TS_COTRAITANCE')).TabVisible := false;
    end;

    if Assigned(GetControl('AFF_REFEXTERNE')) then SetControlVisible('AFF_REFEXTERNE', True);
    if Assigned(GetControl('TAFF_REFEXTERNE')) then SetControlVisible('TAFF_REFEXTERNE', True);
    exit;
  end
  else
  begin
    //FV1 : 14/02/2018 - FS#2906 - MOUTHON : Cotraitants déjà existants lors de la création de l'affaire
    if AFTypeAction=taCreat then
      TToolbarButton97  (getControl('MNCOTRAIT')).visible := False
    else
      TToolbarButton97  (getControl('MNCOTRAIT')).visible := True;
    //
    //TToolbarButton97  (getControl('MNCOTRAIT')).visible := true;
    if Assigned(GetControl('TS_COTRAITANCE')) then
    begin
    	SetControlVisible('TS_COTRAITANCE', true);
			TTabSheet(GetControl('TS_COTRAITANCE')).TabVisible := true;
    end;
    if Assigned(GetControl('AFF_REFEXTERNE')) then SetControlVisible('AFF_REFEXTERNE', False);
    if Assigned(GetControl('TAFF_REFEXTERNE')) then SetControlVisible('TAFF_REFEXTERNE', False);
    SetControlEnabled('DefInterv',true);
  end;
    //
  EtatAffect := false;

  //Vérification si l'on se trouve dans le cas d'une affectation à cotraitant ou sous traitant...
  if (THDBCheckbox (GetControl('AFF_MANDATAIRE')).Checked)  then
  begin
    if CoTraitControlAffectation then
    begin
      EtatAffect := true;
      SetControlEnabled('AFF_MANDATAIRE',false);
    end;
    SetControlvisible('FrsGestion',true);
    SetControlvisible('DepCommunes',False);
    SetControlvisible('RecapCoTrait',true);
    if Assigned(GetControl('AFF_BQMANDATAIRE')) then SetControlVisible('AFF_BQMANDATAIRE', true);
    if Assigned(GetControl('TAFF_CODEBQ1')) then SetControlVisible('TAFF_CODEBQ1', true);
    if Assigned(GetControl('TLIB_BQMANDATAIRE')) then SetControlVisible('TLIB_BQMANDATAIRE', true);
    if Assigned(GetControl('TLIB_CODEBQ')) then SetControlVisible('TLIB_CODEBQ', true);
    if Assigned(GetControl('AFF_CODEBQ')) then SetControlVisible('AFF_CODEBQ', true);
    if Assigned(GetControl('TAFF_CODEBQ')) then SetControlVisible('TAFF_CODEBQ', true);
  end
  else if (not THDBCheckbox (GetControl('AFF_MANDATAIRE')).checked) or
  				(THDBCheckbox (GetControl('AFF_MANDATAIRE')).State = cbUnchecked   )  then
  begin
    if CoTraitControlAffectation then
    begin
      SetControlEnabled('AFF_MANDATAIRE',false);
    end;
    SetControlvisible('FrsGestion',false);
    SetControlvisible('DepCommunes',False);
    SetControlvisible('RecapCoTrait',false);
    if Assigned(GetControl('TS_COTRAITANCE')) then SetControlVisible('TS_COTRAITANCE', true);
    if Assigned(GetControl('AFF_REFEXTERNE')) then SetControlVisible('AFF_REFEXTERNE', false);
    if Assigned(GetControl('TAFF_REFEXTERNE')) then SetControlVisible('TAFF_REFEXTERNE', false);
    if Assigned(GetControl('AFF_BQMANDATAIRE')) then SetControlVisible('AFF_BQMANDATAIRE', false);
    if Assigned(GetControl('TAFF_CODEBQ1')) then SetControlVisible('TAFF_CODEBQ1', false);
    if Assigned(GetControl('AFF_CODEBQ')) then SetControlVisible('AFF_CODEBQ', false);
    if Assigned(GetControl('TAFF_CODEBQ')) then SetControlVisible('TAFF_CODEBQ', false);
    if Assigned(GetControl('TLIB_BQMANDATAIRE')) then SetControlVisible('TLIB_BQMANDATAIRE', false);
    if Assigned(GetControl('TLIB_CODEBQ')) then SetControlVisible('TLIB_CODEBQ', false);
  end;

  SetControlEnabled('FrsGestion',EtatAffect);

  if EtatAffaire = 'ACP' then
  begin
  	SetControlEnabled('AFF_MANDATAIRE',false);
    SetControlvisible('DepCommunes',true)
  end else
  begin
    SetControlvisible('DepCommunes',false);
  end;

  SetControlEnabled('RecapCoTrait',FindRecapCotrait(GetControltext('AFF_AFFAIRE')));
  if (EtatAffaire ='ACP') and (StatutAffaire='PRO') then
  begin
    if assigned(GetControl('POPCOTRAIT')) then
    begin
      SetControlenabled('DefInterv',false);
      SetControlenabled('FrsGestion',false);
      SetControlvisible('DepCommunes',false);
      SetControlvisible('RecapCoTrait',true);
    end;
  end;

end;

procedure TOM_Affaire.SetActionsCotraitance;
begin
	if TMenuItem(getControl('DefInterv')) <> nil then
  begin
    TMenuItem(GetControl('DefInterv')).OnClick := CoTraitSetInterv;
    TMenuItem(GetControl('FrsGestion')).OnClick := CoTraitFrsgestion;
    TMenuItem(GetControl('DepCommunes')).OnClick := CoTraitDepenses;
    TMenuItem(GetControl('RecapCoTrait')).OnClick := CotraitRecap;
  end;
end;

procedure TOM_Affaire.CoTraitDepenses(Sender: Tobject);
begin

end;

procedure TOM_Affaire.CoTraitFrsgestion(Sender: Tobject);
begin
	AglLanceFiche ('BTP','BTAFFAIREFRSGEST','','','AFFAIRE='+GetField ('AFF_AFFAIRE')+';'+
                                               ActionToString (AFTypeAction) );
end;

procedure TOM_Affaire.CotraitRecap(Sender: Tobject);
begin
	AGLLanceFiche('BTP','BTMULAFFECTEXTE','','','COTRAITANCE;AFFAIRE='+GetField ('AFF_AFFAIRE')+';ACTION=MODIFICATION');
end;

procedure TOM_Affaire.CoTraitSetInterv(Sender: Tobject);
begin
	AglLanceFiche ('BTP','BTAFFAIREINTERV','','','AFFAIRE='+GetField ('AFF_AFFAIRE')+
  																						 ';MODEGESTION=COTRAITANCE;MANDATAIRE='+
                                               BoolToStr_(THCheckBox(GetControl('AFF_MANDATAIRE')).checked)+';'+
                                               ActionToString (AFTypeAction) );
  DefiniMenuCotraitance;
end;

function TOM_Affaire.CoTraitControlAffectation: boolean;
begin
	result := false;
  if GetField ('AFF_AFFAIRE') = '' then exit;
  result := CotraitanceAffectationOk (GetField ('AFF_AFFAIRE'));
end;

procedure  TOM_Affaire.GestionEventCotrait(Active : boolean);
begin
  if active then
  begin
    if assigned(GetControl('CBCOTRAITANCE')) then
      THCheckBox(GetControl('CBCOTRAITANCE')).OnClick := CBCOTRAITANCEClick;
    if assigned(GetControl('RBMANDATAIRE')) then
      TRadioButton (getControl('RBMANDATAIRE')).OnClick := RBMANDATAIREClick;
    if assigned(GetControl('RBMANDATAIRE')) then
      TRadioButton (getControl('RBMANDATAIRE')).OnClick := RBCOTRAITANTEClick;
  end else
  begin
    if assigned(GetControl('CBCOTRAITANCE')) then
      THCheckBox(GetControl('CBCOTRAITANCE')).OnClick := nil;
    if assigned(GetControl('RBMANDATAIRE')) then
      TRadioButton (getControl('RBMANDATAIRE')).OnClick := nil;
    if assigned(GetControl('RBMANDATAIRE')) then
      TRadioButton (getControl('RBMANDATAIRE')).OnClick := nil;
  end;
end;

procedure TOM_Affaire.DecodeMandataire;
var CodeMandataire : string;
begin
  GestionEventCotrait(false);
	if (not VH_GC.SeriaCoTraitance) or (StatutAffaire = 'INT') then
  begin
    THRadioGroup (getControl('GBCOTRAITANCE')).Visible := false;
    THCheckbox(GetControl('CBCOTRAITANCE')).checked := false;
    TRadioButton(GetControl('RBMANDATAIRE')).checked := false;
    TRadioButton(GetControl('RBCOTRAITANTE')).checked := false;
    TRadioButton(GetControl('RBMANDATAIRE')).visible := false;
    TRadioButton(GetControl('RBCOTRAITANTE')).visible := false;
    THCheckbox(GetControl('CBCOTRAITANCE')).visible := false;
    exit;
  end;

	CodeMandataire := GetField('AFF_MANDATAIRE');
	if CodeMandataire ='' then
  begin
    THCheckbox(GetControl('CBCOTRAITANCE')).checked := false;
    TRadioButton(GetControl('RBMANDATAIRE')).checked := false;
    TRadioButton(GetControl('RBCOTRAITANTE')).checked := false;
    TRadioButton(GetControl('RBMANDATAIRE')).enabled := false;
    TRadioButton(GetControl('RBCOTRAITANTE')).enabled := false;
  end else if CodeMandataire = 'X' then
  begin
    TRadioButton(GetControl('RBMANDATAIRE')).enabled := true;
    TRadioButton(GetControl('RBCOTRAITANTE')).enabled := true;
    THCheckbox(GetControl('CBCOTRAITANCE')).checked := true;
    TRadioButton(GetControl('RBMANDATAIRE')).checked := true;
    TRadioButton(GetControl('RBCOTRAITANTE')).checked := false;
  end else if CodeMandataire = '-' then
  begin
    TRadioButton(GetControl('RBMANDATAIRE')).enabled := true;
    TRadioButton(GetControl('RBCOTRAITANTE')).enabled := true;
    THCheckbox(GetControl('CBCOTRAITANCE')).checked := true;
    TRadioButton(GetControl('RBMANDATAIRE')).checked := false;
    TRadioButton(GetControl('RBCOTRAITANTE')).checked := true;
  end;
  GestionEventCotrait(true);
end;

procedure TOM_Affaire.EncodeMandataire;
var CodeMandataire : string;
begin
	if (THCheckbox(GetControl('CBCOTRAITANCE')).checked) then
  begin
		if (TRadioButton(GetControl('RBCOTRAITANTE')).checked) then
    begin
    	CodeMandataire := '-';
    end else
    begin
    	CodeMandataire := 'X';
    end;
  end else
  begin
    CodeMandataire := '';
  end;
  SetField ('AFF_MANDATAIRE',CodeMandataire);
end;

procedure TOM_Affaire.CBCOTRAITANCEClick(Sender: Tobject);
begin
	EncodeMandataire;
  DecodeMandataire;
end;

procedure TOM_Affaire.RBMANDATAIREClick(Sender: Tobject);
begin
	EncodeMandataire;
  DecodeMandataire;
end;

procedure TOM_Affaire.RBCOTRAITANTEClick(Sender: Tobject);
begin
	EncodeMandataire;
  DecodeMandataire;
end;

procedure TOM_Affaire.ChangePeriodicite(Sender: TObject);
begin
	if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
  //
  if GetField ('AFF_PERIODICITE') = 'S' then
  begin
    SetField('AFF_DETECHEANCE','DME');
  end else if GetField  ('AFF_PERIODICITE') = 'M' then
  begin
    SetField('AFF_DETECHEANCE','DMM');
    SetField('AFF_INTERVALGENER',1);
  end else if GetField  ('AFF_PERIODICITE') = 'P' then
  begin
    SetField('AFF_DETECHEANCE','DMP');
  end else if GetField('AFF_PERIODICITE') = 'NBI' then
  begin
    SetField('AFF_DETECHEANCE','DME');
    SetField('AFF_TERMEECHEANCE','ACH');
    SetField('AFF_METHECHEANCE','AN');
    SetField('AFF_RECONDUCTION','PDR');
  end else
  begin
    SetField('AFF_DETECHEANCE','DMA');
  end;
	GereAffichagePeriodicite;
end;

procedure TOM_Affaire.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin

end;

procedure TOM_Affaire.GenereWordClick(Sender: Tobject);
var TT : TOB;
		ListeChamps : string;
    NomFile,NomDoc : string;
begin

  NomFile := ExtractFileName(GetControlText('AFF_MODELEWORD'));
  if (NomFile = '') then
  begin
  	PGIError('Vous devez renseigner un nom de modèle');
    Exit;
  end;
	NomFile := IncludeTrailingPathDelimiter(GetParamSocSecur('SO_BTCONREPERTMODELE',''))+NomFile;
  //
  NomDoc := ExtractFileName(GetControlText('DOCGENERE'));
  if (NomDoc = '') or (NomDoc = '*.doc') then
  begin
  	PGIError('Vous devez renseigner un nom de document de sortie');
    Exit;
  end;
  ListeChamps := '';
  TT := TOB.Create ('CONTRAT',nil,-1);
  TRY
    ConstitueTOB(TT,ListeChamps,GetControlText('AFF_AFFAIRE'),true);
  	LancePublipostage('TOB',NomFile,GetControlText('DOCGENERE'),TT,ListeChamps,nil,False);
  finally
    TT.Free;
  end;
end;

procedure TOM_Affaire.ModifModeleClick(Sender: Tobject);
var TT : TOB;
		ListeChamps : string;
    NomFile : string;
begin
  NomFile := ExtractFileName(GetControlText('AFF_MODELEWORD'));
  if (NomFile = '') or (NomFile = '*.dotx') then
  begin
  	PGIError('Vous devez renseigner un nom de modèle');
    Exit;
  end;
	NomFile := IncludeTrailingPathDelimiter(GetParamSocSecur('SO_BTCONREPERTMODELE',''))+NomFile;
  if not FileExists(NomFile) then
  begin
    PgiInfo('Ce document n''existe pas ...');
    exit;
  end;
  ListeChamps := '';
  TT := TOB.Create ('CONTRAT',nil,-1);
  TRY
    ConstitueTOB(TT,ListeChamps,GetControlText('AFF_AFFAIRE'),true);
  	LancePublipostage('OPEN',NomFile,'',TT,ListeChamps,nil,False);
  finally
    TT.Free;
  end;
end;

procedure TOM_Affaire.NewModeleClick(Sender: Tobject);
var TT : TOB;
		ListeChamps,NomFile : string;
begin

  ListeChamps := '';

  TT := TOB.Create ('CONTRAT',nil,-1);
  TRY
		NomFile := IncludeTrailingPathDelimiter(GetParamSocSecur('SO_BTCONREPERTMODELE',''))+'newDoc.doc';
    ConstitueTOB(TT,ListeChamps,GetControlText('AFF_AFFAIRE'),false);
  	LancePublipostage('NEW',NomFile,'',TT,ListeChamps,nil,False);
  finally
    TT.Free;
  end;

end;

procedure TOM_Affaire.OuvrirWordClick(Sender: Tobject);
var NomDoc : string;
begin
//
  NomDoc := ExtractFileName(GetControlText('DOCGENERE'));
  if (NomDoc = '') or (NomDoc = '*.doc') then
  begin
  	PGIError('Vous devez renseigner le nom du document');
    Exit;
  end;
  if not FileExists(GetControlText('DOCGENERE')) then
  begin
    PgiInfo('Ce document n''existe pas ...');
    exit;
  end;
	ShellExecute(0, pchar('open'), pchar(string(GetControlText('DOCGENERE'))), nil, nil, SW_RESTORE)
end;

procedure TOM_Affaire.SuppModeleClick(Sender: Tobject);
var NomFile : string;
begin
  NomFile := ExtractFileName(GetControlText('AFF_MODELEWORD'));
  if (NomFile = '') or (NomFile = '*.dotx') then
  begin
  	PGIError('Vous devez renseigner un nom de modèle');
    Exit;
  end;
	NomFile := IncludeTrailingPathDelimiter(GetParamSocSecur('SO_BTCONREPERTMODELE',''))+NomFile;

  if not FileExists(nomFile) then
  begin
  	PGIError('Ce modèle n''existe pas');
    Exit;
  end;
  if PGIAsk('Désirez-vous réellement supprimer ce modèle ?',ecran.caption) = mryes then
  begin
    DeleteFile(PAnsiChar(nomFile));
		SetControlText('AFF_MODELEWORD','');
  end;
end;

procedure TOM_Affaire.AjouteChampsContrat (OneTOB : TOB; var ListNomAFF,ListSqlAFF : string; WithData: Boolean);

	procedure GetMtDoc(Affaire: string;var MtDocHt,MtDocTTC : double);
  var Sql : string;
  		QQ : TQuery;
  begin
		Sql := 'SELECT GP_TOTALHTDEV,GP_TOTALTTCDEV FROM PIECE WHERE GP_NATUREPIECEG="AFF" AND Gp_AFFAIRE="'+Affaire+'"';
    QQ := OpenSQL(Sql,True,1,'',true);
    if not QQ.eof then
    begin
      MtDocHt := QQ.fields[0].asfloat;
      MtDocTTC := QQ.fields[1].asfloat;
    end;
    ferme (QQ);
  end;

var Sql : String;
		QQ : TQuery;
    ii : Integer;
    MtDocHt,MtDocTTC : double;
begin
  if ListSqlAFF = '' then Exit;
	Sql := 'SELECT '+ListSqlAFF+' FROM AFFAIRE WHERE AFF_AFFAIRE="'+GetField('AFF_AFFAIRE')+'"';
  QQ := OpenSQL(Sql,True,1,'',True);
  for II := 0 to QQ.Fields.Count -1 do
  begin
    if WithData then
      Publipost.SetData(OneTOB, QQ.fields[II])
    else
      OneTOB.AddChampSupvaleur(QQ.Fields[II].FullName,'');
  end;
  ferme(QQ);
  //
  ListNomAFF := ListNomAFF +
  							';Contrat_Client_adresse1;Contrat_Client_adresse2;Contrat_Client_adresse3;Contrat_Client_CodePostal;Contrat_Client_Ville;' +
  							'Contrat_montantdoc_ht;Contrat_montantdoc_ttc;Contrat_montant_global_HT;Contrat_montant_global_TTC7;Contrat_montant_global_TTC196;'+
                'Contrat_Contact_Civilite;Contrat_Contact_Client;Contrat_Contact_tel';
  //              
  OneTOB.AddChampSupValeur('TIE_ADRESSE1','');
  OneTOB.AddChampSupValeur('TIE_ADRESSE2','');
  OneTOB.AddChampSupValeur('TIE_ADRESSE3','');
  OneTOB.AddChampSupValeur('TIE_CODEPOSTAL','');
  OneTOB.AddChampSupValeur('TIE_VILLE','');
  //
  GetMtDoc (GetField('AFF_AFFAIRE'),MtDocHt,MtDocTTC);
  //
  OneTOB.AddChampSupValeur('AFF_MTDOCHT',MtDocHt);
  OneTOB.AddChampSupValeur('AFF_MTDOCTTC',MtDocTTC);
  OneTOB.AddChampSupValeur('AFF_MTGLOBALHT',THNumEdit(GetControl('MONTANTGLOBAL')).Value);
  OneTOB.AddChampSupValeur('AFF_MTGLOBALTTC7',Arrondi(THNumEdit(GetControl('MONTANTGLOBAL')).Value*1.07,2));
  OneTOB.AddChampSupValeur('AFF_MTGLOBALTTC196',Arrondi(THNumEdit(GetControl('MONTANTGLOBAL')).Value*1.196,2));
  //
  OneTOB.AddChampSupValeur('AFF_GO_CIVILITECONTACT','');
  OneTOB.AddChampSupValeur('AFF_GO_NOMCONTACT','');
  OneTOB.AddChampSupValeur('AFF_GO_TELCONTACT','');
  //
  if (GetField ('AFF_TIERS') <> '') then
  begin
    QQ := OpenSql ('SELECT T_ADRESSE1,T_ADRESSE2,T_ADRESSE3,T_CODEPOSTAL,T_VILLE FROM TIERS WHERE T_TIERS="' + GetField ('AFF_TIERS')  + '" AND T_NATUREAUXI="CLI"' , TRUE) ;
    if not QQ.EOF then
    begin
      OneTOB.SetString('TIE_ADRESSE1',QQ.Fields[0].asstring);
      OneTOB.SetString('TIE_ADRESSE2',QQ.Fields[1].asstring);
      OneTOB.SetString('TIE_ADRESSE3',QQ.Fields[2].asstring);
      OneTOB.SetString('TIE_CODEPOSTAL',QQ.Fields[3].asstring);
      OneTOB.SetString('TIE_VILLE',QQ.Fields[4].asstring);
    end;
    ferme(QQ);
  end;
  if (GetField ('AFF_TIERS') <> '') and (GetField ('AFF_NUMEROCONTACT') <> 0) then
  begin
    QQ := OpenSql ('SELECT C_CIVILITE,C_NOM,C_TELEPHONE FROM CONTACT WHERE C_TYPECONTACT="T" AND C_AUXILIAIRE="' + TiersAuxiliaire (GetField ('AFF_TIERS') ) + '" AND C_NUMEROCONTACT=' + intToStr (GetField ('AFF_NUMEROCONTACT') ) , TRUE) ;
    if not QQ.EOF then
    begin
      OneTOB.SetString('AFF_GO_CIVILITECONTACT',QQ.Fields[0].asstring);
      OneTOB.SetString('AFF_GO_NOMCONTACT',QQ.Fields[1].asstring);
      OneTOB.SetString('AFF_GO_TELCONTACT',QQ.Fields[2].asstring);
    end;
    Ferme (QQ) ;
  end;
  //
end;

procedure TOM_Affaire.AjouteAdressesInt (OneTOB : TOB; var ListNomAdrInt,ListSqlAFF : string; WithData: Boolean);
var Sql,SqlT : string;
		QQ : TQuery;
    ii : Integer;
    TiersLiv,NumContact : string;
begin
  if ListSqlAFF = '' then Exit;
  //
	SqlT := 'SELECT GP_NUMERO FROM PIECE WHERE GP_AFFAIRE="'+GetField('AFF_AFFAIRE')+'" AND GP_NATUREPIECEG="AFF"';
	Sql := 'SELECT '+ListSqlAFF+' FROM PIECEADRESSE WHERE GPA_NATUREPIECEG="AFF" AND GPA_NUMERO=('+SQLT+') AND GPA_TYPEPIECEADR="001"';
  QQ := OpenSQL(Sql,True,1,'',True);
  for II := 0 to QQ.Fields.Count -1 do
  begin
    if WithData then
      Publipost.SetData(OneTOB, QQ.fields[II])
    else
      OneTOB.AddChampSupvaleur(QQ.Fields[II].FullName,'');
  end;
  ferme(QQ);
  // Ajout des infos concernant l'interlocuteur de l'adresse d'intervention
	Sql := 'SELECT GP_TIERSLIVRE FROM PIECE WHERE GP_AFFAIRE="'+GetField('AFF_AFFAIRE')+'" AND GP_NATUREPIECEG="AFF"';
  QQ := OpenSQL(Sql,True,1,'',True);
  if not QQ.eof then
  begin
		TiersLiv := TiersAuxiliaire(QQ.fields[0].AsString,false,'CLI');
  end;
  ferme (QQ);
	Sql := 'SELECT GPA_NUMEROCONTACT FROM PIECEADRESSE WHERE GPA_NATUREPIECEG="AFF" AND GPA_NUMERO=('+SQLT+') AND GPA_TYPEPIECEADR="001"';
  QQ := OpenSQL(Sql,True,1,'',True);
  if not QQ.eof then
  begin
    NumContact := QQ.fields[0].AsString;
  end;
  Ferme(QQ);
  ListNomAdrInt := ListNomAdrInt + ';Adrint_Civilite_Contact;AdrInt_Nom_Contact;AdrInt_prenom_Contact;AdrInt_Tel_Contact';
  OneTOB.AddChampSupValeur('GPA_CI_CIVILITE','');
  OneTOB.AddChampSupValeur('GPA_CI_NOMCONTACT','');
  OneTOB.AddChampSupValeur('GPA_CI_PRENOMCONTACT','');
  OneTOB.AddChampSupValeur('GPA_CI_TELCONTACT','');

  if (TiersLiv <> '') and (NumContact <> '') then
  begin
    SqlT := 'SELECT C_CIVILITE,C_NOM,C_PRENOM,C_TELEPHONE,C_FAX,C_TELEX FROM CONTACT WHERE '+
    				'C_TYPECONTACT="T" AND C_AUXILIAIRE="'+TiersLiv+'" AND C_NUMEROCONTACT='+numContact;

    QQ := OpenSQL(SqlT,True,1,'',True);
    if not QQ.eof then
    begin
      OneTOB.PutValue('GPA_CI_CIVILITE',QQ.Fields[0].AsString);
      OneTOB.PutValue('GPA_CI_NOMCONTACT',QQ.Fields[1].AsString);
      OneTOB.PutValue('GPA_CI_PRENOMCONTACT',QQ.Fields[2].AsString);
      OneTOB.PutValue('GPA_CI_TELCONTACT',QQ.Fields[3].AsString);
    end;
    ferme (QQ);
  end;
  //
end;


procedure TOM_Affaire.AjouteChampsClientFac (OneTOB : TOB; ListSqlAFF : string; WithData: Boolean);
var Sql,SqlT : string;
		QQ : TQuery;
    ii : Integer;
begin
  if ListSqlAFF = '' then Exit;
	SqlT := 'SELECT AFF_FACTURE FROM AFFAIRE WHERE AFF_AFFAIRE="'+GetField('AFF_AFFAIRE')+'"';
	Sql := 'SELECT '+ListSqlAFF+' FROM TIERS WHERE T_AUXILIAIRE=('+SQLT+')';
  QQ := OpenSQL(Sql,True,1,'',True);
  for II := 0 to QQ.Fields.Count -1 do
  begin
    if WithData then
      Publipost.SetData(OneTOB, QQ.fields[II])
    else
      OneTOB.AddChampSupvaleur(QQ.Fields[II].FullName,'');
  end;
  ferme(QQ);
end;

procedure TOM_Affaire.ConstitueTOB(OneTOB: TOB; Var ListeChamps : string; CodeContrat: string; WithData : boolean=true);
var
  ListSqlAFF,ListSQlTiersFac,ListSQlTiersint,ListSqlAdrFac,ListSqlAdrInt : string;
	ListnomAFF,ListnomTiersFac,ListNomTiersint,ListNomAdrFac,ListNomAdrInt : string;
begin
  Publipost.SetFieldsList(Publipost.GetPrefixFromTableName('AFFAIRE')     , 'AFFAIRE'     , ListNomAFF     , ListSqlAFF);
  Publipost.SetFieldsList(Publipost.GetPrefixFromTableName('TIERS')       , 'TIERS'       , ListNomTiersFac, ListSQlTiersFac);
  Publipost.SetFieldsList(Publipost.GetPrefixFromTableName('PIECEADRESSE'), 'PIECEADRESSE', ListNomAdrInt  , ListSqlAdrInt);
  //
  AjouteChampsContrat (OneTOB,ListNomAFF,ListSqlAFF,WithData);
  AjouteChampsClientFac (OneTOB,ListSQlTiersFac,WithData);
  AjouteAdressesInt (OneTOB,ListNomAdrInt,ListSqlAdrInt,WithData);
  //
  ListeChamps := ListNomAff+';'+ListNomTiersFac+';'+ListNomAdrInt;
end;


procedure TOM_Affaire.SelectModele(Sender: TObject);
var TT : TOpenDialog;
begin
	TT := TOpenDialog.Create(self.Ecran);
  TRY
    TT.DefaultExt := '.dotx';
    TT.Filter := 'Modèle de document word (*.dotx)|*.dotx';
    TT.InitialDir := GetParamSocSecur('SO_BTCONREPERTMODELE','');
    if TT.Execute then
    begin
      SetField('AFF_MODELEWORD',ExtractFileName(TT.FileName) );
    end;
  FINALLY
  	TT.Free;
  end;
end;

procedure TOM_Affaire.SelectDocWord(Sender: TObject);
var TT : TOpenDialog;
begin
	TT := TOpenDialog.Create(self.Ecran);
  TRY
    TT.DefaultExt := '.doc';
    TT.Filter := 'Document word (*.doc)|*.doc';

    TT.InitialDir := GetParamSocSecur('SO_BCONTRATSTOW','');
    if TT.Execute then
    begin
      SetControltext('DOCGENERE',TT.FileName );
    end;
  FINALLY
  	TT.Free;
  end;
end;

procedure TOM_Affaire.RegroupeEches(Sender: TObject);
begin
  AGLLanceFiche ('BTP', 'BTREGROUPECHES', 'AFA_AFFAIRE='+GetField('AFF_AFFAIRE'),'',';ACTION=MODIFICATION');
end;

function TOM_Affaire.GetLibResponsable(CodeResssource : string) : string;
Var QQ  : TQuery;
begin
  result := '';
  QQ := OpenSQL('SELECT ARS_LIBELLE, ARS_LIBELLE2 from RESSOURCE WHERE ARS_RESSOURCE="' + CodeResssource + '"', False);
  If Not QQ.eof then
  begin
    result :=  QQ.findfield('ARS_LIBELLE2').AsString + ' ' + QQ.findfield('ARS_LIBELLE').AsString;
  end;
  Ferme(QQ);
end;

//FV1 - 14/06/2016 - FS#2023 - En fiche affaire, la recherche sur responsable ne fonctionne plus
procedure TOM_Affaire.BRechResponsable(Sender: TObject);
Var SS  : THCritMaskEdit;
    Zone : THDbEdit;
    Libelle : string;
begin
  Zone := THDbEdit(Sender);
  //
  if Zone.name = 'AFF_RESPONSABLE' then Libelle := 'LIBAFF_RESPONSABLE'
  else if Zone.name = 'AFF_RESPONSABLE1' then Libelle := 'LIBAFF_RESPONSABLE1'
  else if Zone.name = 'AFF_RESSOURCE1' then Libelle := 'LIBAFF_RESSOURCE1'
  else if Zone.name = 'AFF_RESSOURCE2' then Libelle := 'LIBAFF_RESSOURCE2'
  else if Zone.name = 'AFF_RESSOURCE3' then Libelle := 'LIBAFF_RESSOURCE3';
  //
  if GetParamSocSecur('SO_AFRECHRESAV', True) then
  begin
    SS := THCritMaskEdit.Create(nil);
    GetRessourceRecherche(SS,'ARS_RESSOURCE=' + Zone.text + ';TYPERESSOURCE=SAL', '', '');
    if (SS.Text <> Zone.text) then
    begin
      if SS.text = '' then ss.text := Zone.text;
      if not (DS.State in [dsInsert, dsEdit]) then DS.edit;
    end;
    Zone.text  := SS.Text;
    SS.Free;
  end else
  begin
    GetRessourceRecherche(THCritMaskEdit(Zone),'ARS_TYPERESSOURCE="SAL"', '', '');
  end;

  IF THDBEdit(Zone).text <> '' then
  begin
    SetControlproperty(Libelle, 'Visible', true);
    SetControlText(Libelle, GetLibResponsable(Zone.Text));
  end else
  begin
    SetControlproperty(Libelle, 'Visible', False);
  end;

end;

//FV1 : 17/05/2017 - FS#2546 - GUINIER : Dans la fiche affaire, ajouter l'accès à l'encours client.
procedure TOM_Affaire.VoirEncours(Sender : TObject);
var TobTiers  : TOB ;
    QQ        : TQuery;
    Tiers     : string;
    StSQL     : string;
    VOR       : string;
    OkOk      : string ;
    Action    : TActionFiche ;
begin

  Action:=AFTypeAction;

  if STTiers.Text = '' then Exit;

  //Lecture de la table des tiers
  Try
    StSQL := 'Select * from TIERS Where T_TIERS="' + StTiers.Text + '"';
    QQ := OpenSQL(StSQL, False);
    If Not QQ.eof then
    begin
      TOBTiers:=Tob.create('TIERS',Nil,-1);
      TOBTiers.SelectDB('TIERS',QQ,false);

      CalculSoldesAuxi(TOBTiers.GetValue('T_AUXILIAIRE')) ;

      TobTiers.LoadDB(True) ; // Pour recharger les données calculées

      TheTob:= TobTiers;

      OkOk  := AglLanceFiche('GC', 'GCENCOURS','','',ActionToSTring(Action));

      if (OkOk='OK') and (Action<>taConsult) then
      begin
        VOR := Risque.text;
        if VOR <> '' then Risque.text := VOR;
      end;
      TobTiers.free;
    end;
  finally
    Ferme(QQ);
  end;

end;

//Gestion de l'affichage du bouton feu en fonction du risque (0/R/V)
Procedure TOM_Affaire.GestionFeuEncour;
Begin

	bFEUROUGE.Visible := Risque.text = 'R';
	bFEUORANGE.Visible:= Risque.text = 'O';
	bFEUVERT.visible  := Risque.text = 'V';

End;

Procedure TOM_Affaire.ExitLigCaution(Sender : Tobject);
Var II : Integer;
    MtCaution : Double;
    Nomzone   : string;
begin

  Nomzone   := THNumEdit(Sender).name;
  MtCaution := StrToFloat(THNumEdit(Sender).text);

  II :=  StrToInt(RightStr(Nomzone,1));

  if MtCaution= 0 then Exit
  Else
  begin
    II := II + 1;
    if II > 9 Then
    Begin
      Exit;
    End
    else
    Begin
      SetControlVisible('BAR_DATECAUTION' + IntToStr(II),True);
      SetControlVisible('BAR_MTCAUTION'   + IntToStr(II),True);
      SetControlVisible('BAR_NUMCAUTION'  + IntToStr(II),True);
    end;
  end;

end;

procedure TOM_Affaire.BTNSOUSTRAITClick(Sender: TObject);
begin
  AGLLanceFiche('BTP','BTCONTRATST_MUL','BM0_AFFAIRE='+GetField('AFF_AFFAIRE'),'','ACTION=MODIFICATION')
end;

initialization
  registerclasses ([TOM_FactAff, {TOM_AffTiers,} TOM_Affaire] ) ;
  RegisterAglProc ('CalculDatesMission', True, 0, AGLCalculDatesMission) ;
  RegisterAglProc ('SaisiePieceAffaire', True, 1, AGLSaisiePieceAffaire) ;
  RegisterAglProc ('LanceEcheance', True, 3, AGLLanceEcheance) ;
  RegisterAglProc ('CalculMontantGlobal', True, 0, AGLCalculMontantGlobal) ;
  RegisterAglProc ('LanceAfftiers', True, 1, AGLLanceAffTiers) ;
  RegisterAglProc ('LanceRegroupeAffaire', True, 0, AGLLanceRegroupeAffaire) ;
  RegisterAglProc ('ImprimeModele', TRUE, 1, AGLImprimeModele) ;
  RegisterAglFunc ('ActionAffaire', TRUE, 0, AGLModifAutorisee) ;
  RegisterAglProc ('NomContactAff', TRUE, 0, AGLNomContactAff) ;
  RegisterAglProc ('AffaireAppelIsoFlex', TRUE, 1, AffaireAppelIsoFlex) ;
  RegisterAglProc('ImprimeEcran',True,0,ImprimeThisScreen);
end.
