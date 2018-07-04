{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/07/2002
Modifié le ... : 01/08/2005
Description .. : Le 04/01/2004 -> Correction du problème de lancement de
Suite ........ : l'
Suite ........ : écran par le PopUp standard de la compta ( Create/Modif :
Suite ........ : Grisé )
Suite ........ : Source TOF de la FICHE : CONSECR ()
Suite ........ : GCO - 18/02/2004 -> Passage des paramètres au GL avec
Suite ........ : un CritEdt
Suite ........ : GCO - 28/02/2004 -> Modif de la requête SQL(Détail ANO,
Suite ........ : GCO - 26/03/2004 -> Ajout du Test de Gestion d'ICC pour
Suite ........ : éviter
Suite ........ : des SELECT inutiles lors de l'appel de WhatIsOk
Suite ........ : GCO - 17/07/2004 -> Blocage des fonctions si compte
Suite ........ : confidentiel
Mots clefs ... : TOF;CONSECR
*****************************************************************}
unit uTofConsEcr;

interface

uses StdCtrls,
  Controls,
  Classes,
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  uTob,
  Filtre,
  hQry,
  Graphics,
  Vierge,
  Menus,         // MPopUpMenu
  LookUp,        // LookUpList
  Windows,       // VK_
  Grids,         // TGridDrawState
  Htb97,         // TToolBarButton97
  Ent1,          // ExoToDates, TFichierBase
  HRichOle,      // THRichEditOle
  HSysMenu,      // HSystemMenu
  uTofViergeMul, // Fiche Vierge MUL
  uLibEcriture,  // TListJournal
  HPanel ,       // THPanel
  SaisUtil,      // RMVT
  uLibExercice,  // CInitComboExercice, TZExercice

{$IFDEF EAGLCLIENT}
  MainEagl,      // AGLLanceFiche
  UtileAGL,
{$ELSE}
  Db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  FE_main,       // AGLLanceFiche
{$ENDIF}

{$IFDEF VER150}
  Variants,
{$ENDIF}

{$IFDEF MODENT1}
  CPTypeCons,
{$ENDIF MODENT1}
  Dialogs;

type
  TOF_CONSECR = class(TOF_ViergeMul)

    E_General       : THEdit;
    E_Auxiliaire    : THEdit;
    E_DATECOMPTABLE : THEdit;
    E_DATECOMPTABLE_: THEdit;

    LESOLDE         : THEdit;
    LESOLDEN        : THEdit;
    LESOLDENMOINS1  : THEdit;

    ESoldeBancaire  : THNumEdit;
    EMontant        : THEdit;

    CBJustifSolde   : TCheckBox;
    CBCentralisation: TCheckBox;
    CBDetailANO     : TCheckBox;
    CBAvecCollectif : TCheckBox;
    CBDocumentGED   : TCheckBox;

    E_Valide        : TCheckBox;

    E_Journal       : THMultiValComboBox;
    E_NaturePiece   : THMultiValComboBox;
    E_Devise        : THValComboBox;
    EQualifPiece    : THMultiValComboBox;
    E_Exercice      : THValComboBox;
    E_Etablissement : THMultiValComboBox;
    E_EtatLettrage  : THMultiValComboBox;
    E_RefPointage   : THEdit;
    E_RefPointage_  : THEdit;
    E_DatePointage  : THEdit;
    E_Lettrage      : THEdit;
    E_NumeroPiece   : THEdit;
    E_NumeroPiece_  : THEdit;
    ELibelle        : THEdit;
    ERefInterne     : THEdit;

    ComboInfoSup      : THValComboBox;
    ComboAcces        : THValComboBox;
    ComboDefilCpt     : THValComboBox;
    ComboSoldePro     : THValComboBox;
    ComboMontant      : THValComboBox;
    ComboOpMontant    : THValComboBox;
    ComboOpLibelle    : THValComboBox;
    ComboOpRefInterne : THValComboBox;

    BGUP                : TToolBarButton97;
    BGDOWN              : TToolBarButton97;
    BTUP                : TToolBarButton97;
    BTDOWN              : TToolBarButton97;

    BComptable          : TToolBarButton97;
    BSpecifique         : TToolBarButton97;
    BAideRevision       : TToolBarButton97;
    BRevInteg           : TToolBarButton97;
    BParametre          : TToolBarButton97;
    BEdition            : TToolBarButton97;
    BFctMultiSelect     : TToolBarButton97;

    PopUpComptable      : TPopUpMenu;
    PopUpSpecifique     : TPopUpMenu;
    PopUpAideRevision   : TPopUpMenu;
    PopUpRevInteg       : TPopUpMenu;
    PopUpParametre      : TPopUpMenu;
    PopUpMultiSelect    : TPopUpMenu;
    PopUpEdition        : TPopUpMenu;

    PageInfoSup     : TPageControl;
    TabTablesLibres : TTabSheet;

    procedure OnArgument(S: string); override;
    procedure OnClose;               override;
    procedure OnLoad;                override;
    procedure AfterShow;             override;
    procedure OnNew;                 override;
    procedure OnDelete;              override;
    procedure OnUpdate;              override;

    procedure OnKeyDownEcran              ( Sender : TObject; var Key: Word; Shift: TShiftState); override;
    procedure OnFlipSelectionFListe       ( Sender : TObject ); override;

    procedure OnExitE_General             ( Sender : TObject );
    procedure OnExitE_Auxiliaire          ( Sender : TObject );
    procedure OnExitE_RefPointage         ( Sender : TObject );
    procedure OnExitE_DatePointage        ( Sender : TObject );
    procedure OnExitE_DateComptable       ( Sender : TObject );

    procedure OnElipsisClickE_General     ( Sender : TObject );
    procedure OnElipsisClickE_Auxiliaire  ( Sender : TObject );
    procedure OnElipsisClickE_RefPointage ( Sender : TObject );
    procedure OnElipsisClickE_DatePointage( Sender : TObject );

    procedure OnClickCBJustifSolde        ( Sender : TObject );
    procedure OnClickCBDetailAno          ( Sender : TObject );

    procedure OnChangeComboE_Exercice     ( Sender : TObject );
    procedure OnChangeComboInfoSup        ( Sender : TObject );
    procedure OnChangeComboAcces          ( Sender : TObject );

    procedure OnClickBGUP                 ( Sender : TObject );
    procedure OnClickBGDOWN               ( Sender : TObject );
    procedure OnClickBTUP                 ( Sender : TObject );
    procedure OnClickBTDOWN               ( Sender : TObject );

    procedure OnClickBBlocNote            ( Sender : TObject ); override;
    procedure OnClickBValider             ( Sender : TObject ); override;
    procedure OnPopUpPopF11               ( Sender : TObject ); override;

    procedure OnKeyDownFListe  ( Sender : TObject ; var Key: Word; Shift: TShiftState); override;
    procedure OnPostDrawCellFListe (ACol, ARow : Integer; Canvas : TCanvas; AState : TGridDrawState); override;

    procedure OnRowEnterFListe (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean); override;
    procedure GetCellCanvasFListe(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    procedure OnDblClickFListe (Sender: TObject);
    procedure OnBeforeFlipFListe(Sender: TObject; ARow: Integer; var Cancel: Boolean); override;

  private
    GLibelle             : THLabel;
    TLibelle             : THLabel;
    FTobGene             : TOB;           // Tob qui contient les infos du  GENERAL
    FTobTiers            : TOB;           // Tob qui contient les infos du  AUXILIAIRE
    FTobListeICC         : Tob;           // Tob qui contient la liste des comptes présents dans ICC

    FTobEcrSelect        : TOB;           // Tob qui contient les infos des ECRITURES
    FTobListeEcr         : TOB;           // Tob qui contient les ecritures sélectionnés
    FZListJournal        : TZListJournal; //
    FDroitEcritures      : Boolean;
    FFromSaisie          : Boolean;
    FBoLettrageEnSaisie  : Boolean;
    FBoOkControleDate    : Boolean;
    FBoConfidentielGene  : Boolean; // Gestion Confidentiel Général pour blocage des fonctions
    FBoConfidentielAuxi  : Boolean; // Gestion Confidentiel Auxiliaire pour blocage des fonctions
    FBoFaireRequete      : Boolean; //
    FBoOuvreEnJustif     : Boolean; //

    FStArgumentTOF       : string;  // Paramètres passés à la TOF dans le OnArgument
    FStNomBase           : string;  // Nom de la base sur laquelle on fait la requête ( Bureau PGI )

    FOkLettrable         : Boolean; // Active les fonctions de Lettrage
    FOkOngletLettrage    : Boolean; // Autorise l'apparition de l'onglet Lettrage
    FOkPointable         : Boolean; // Active les fonctions de Pointage

    FOkIccGere           : Boolean; // ICC est il géré dans le dossier ?
    FOkIcc               : Boolean; // Active les fonctions d'intérêt de comptes courants

    FOkCreGere           : Boolean; // Calcul de remboursement d'emprunt
    FOkCre               : Boolean;

    FOkPointageJal       : Boolean; // Pointage sur Journal;
    FOkModImmo           : Boolean; // Module Immo sérialisé
    FOkGereQte           : Boolean; // Gestion de la quantité sur le dossier
    FOkAnaCroisaxe       : Boolean; // Analytique Croise Axe
    FOkRevision          : Boolean; // Gestion de la révision
    FOkEdition           : Boolean; // Autorise les éditions

    FStLiasseDossier     : string;

    FExoDate             : TExoDate;

{$IFDEF BUREAU}
    FBoBlocageTotal      : Boolean;    // Verouillage totale des fonctions.
    FZExercice           : TZExercice;
    FZInfoCpta           : TZInfoCpta;
{$ENDIF}

    FOkAnalytique        : Boolean; //
    FOkAnaSurAxe1        : Boolean; //
    FOkAnaSurAxe2        : Boolean; //
    FOkAnaSurAxe3        : Boolean; //
    FOkAnaSurAxe4        : Boolean; //
    FOkAnaSurAxe5        : Boolean; //

    FOkCreateModif       : Boolean; // Test combo mode Création / Modification
    FOkEcr               : Boolean; // Create/Modif et Présence d'écritues
    FOkSaisieCA3         : Boolean; //
    FOkEcheance          : Boolean; //
    FOkTableChancellerie : Boolean; //
    FOkLastRow           : Boolean; // Positionnement sur le dernier mouvement
    FOkCentralisation    : Boolean; // Centralisation des écritures
    FOkJustificatifSolde : Boolean; // Justificatif de solde des Grands Livres
    FOkAfficheRIB        : Boolean; // Autorise l'affichage du RIB dans le MENU
    FOkFctLettrage       : Boolean;


    // Variables contenant le numéro des colonnes imortantes de la grille
    FColGeneral, FColAuxiliaire, FColDebit, FColCredit, FColSoldePro : Integer;
    FColLettrage, FColRefPointage : Integer;

    // Variables d'enregistrement de l'état des combos CBJUSTIFSOLDE,CBDETAILANO
    FBoStateCBJustifSolde : Boolean;
    FBoStateCBDetailANO   : Boolean;

    ////////////////////////////////////////////////////////////////////////////
    procedure AnalyseParametre      ( vStArg : String ) ;
    procedure InitAllPopUP          ( vActivation : Boolean );
    procedure InitPopUpComptable    ( vActivation : Boolean ); //
    procedure InitPopUpSpecifique   ( vActivation : Boolean ); //
    procedure InitPopUpAideRevision ( vActivation : Boolean ); //
    procedure InitPopUpRevInteg     ( vActivation : Boolean ); //
    procedure InitPopUpParametre    ( vActivation : Boolean ); //
    procedure InitPopUpMultiselect  ( vActivation : Boolean );
    procedure InitPopUpEdition      ( vActivation : Boolean );

    procedure GeneSuivant(vBoSuiv : Boolean; vBoREsteSurZone : Boolean);
    procedure AuxSuivant (vBoSuiv : Boolean; vBoREsteSurZone : Boolean);

    function  CGenereSQLNatureAuxi : string;

    procedure ActivationOnglet;
    procedure ActivationOption;
    procedure MiseAJourCaptionEcran;
    procedure AfficheInfoSup;
    procedure AfficheSoldeCompte;
    procedure InitColFListe;
    procedure InitContexte;         // Pour suppression des VH si IFDEF BUREAU

{$IFNDEF EAGLCLIENT}
    procedure TraiteCentralisation; // Centralise les ecritures FTobListe
{$ENDIF}

    procedure IndiceColFListe;
    function ChargeGeneral       (Force : boolean): Boolean;
    function ChargeAuxiliaire    (Force : boolean): Boolean;
    function EstVraiEcr          (vTobEcr : TOB)  : Boolean;
    procedure VerifieLesDates;

    function RecupEcrSelect( TheRow : integer = 0 ) : Boolean;

    function RecupCriteresCommuns : string ;
    function RecupWhereAGL        : string ;
    function RecupWhereEcriture   ( vBoForceDetailAno : boolean = false ) : string ;
    function RecupWhereDetailANo  ( vBoForceDetailAno : boolean = false ) : string ;
    function FaireStWhere         ( vBoForceDetailAno : boolean = false ) : string ; // Balance Général par Général
    function RecupSoldePointage   ( vGeneral : string; vDatePointage: TDateTime; vRefPointage: string): Double;
    function NatureGeneAuxi       ( vStNatGeneral, vStNatAuxiliaire : string ) : Boolean;

    ////////////////////////////////////////////////////////////////////////////
    procedure WhatIsOk; // Determine le traitement possible pour le compte
    ////////////////////////////////////////////////////////////////////////////

    procedure OnPopUpComptable        ( Sender : TObject );
    procedure OnPopUpSpecifique       ( Sender : TObject );
    procedure OnPopUpAideRevision     ( Sender : TObject );
    procedure OnPopUpRevInteg         ( Sender : TObject );
    procedure OnPopUpParametre        ( Sender : TObject );
    procedure OnPopUpEdition          ( Sender : TObject );
    procedure OnPopUpMultiSelect      ( Sender : TObject );
    ////////////////////////////////////////////////////////////////////////////

    procedure OnClickLetGestion       ( Sender : TObject ); // Affichage des PARAMSOC Lettrage Gestion
    procedure OnClickLetParam         ( Sender : TObject ); // Affichage des PARAMSOC Lettrage Paramètre
    procedure OnClickPasseJustifSolde ( Sender : TObject ); // Passage en justificatif de Solde
    procedure AuxiElipsisClick        ( Sender : TObject );

    procedure OnClickCompteRib        ( Sender : TObject ); // Relevé d'identité bancaire

    ////////////////////////////////////////////////////////////////////////////
    // Fonctions complémentaires
    procedure OnClickTableChancellerie( Sender : Tobject ); // Tables de chancelleries ( PAS OK CWAS )
    procedure AnalRevisClick          ( Sender : TObject );

    ////////////////////////////////////////////////////////////////////////////
    // Révision
    procedure OnClickDetailIcc          ( Sender : TObject ); // Détail ICC du compte
    procedure OnClickDetailCapitalICC   ( Sender : TObject ); // Détail ICC du compte de capital
    procedure OnClickCalculIcc          ( Sender : TObject ); // Calucl des intérêts ICC


////////////////////////////////////////////////////////////////////////////////
// Fonctions exclusives appelables par la COMPTA
////////////////////////////////////////////////////////////////////////////////
{$IFDEF COMPTA}
    procedure ChargeFTobListeEcr; // Chargement de la multi-sélection en TOB
    function TestAvecEcrSimuSitu  : Boolean;
    function PositionneExo        : TExoDate;
  {$IFNDEF CCMP}

    procedure HistoCpteClick            ( Sender : TObject ); // Historique du compte
    procedure PointageGClick            ( Sender : TObject ); // Pointage du compte
    procedure PointageEnCoursClick      ( Sender : TObject ); // Depointage du compte

    procedure OnClickSaisieTreso        ( Sender : TObject ); // Saisie tresorerie

    // PopUp de la révision
    procedure OnClickEmpruntCRE         ( Sender : TObject ); // Emprunt CRE
    procedure OnClickRappECRCRE         ( Sender : TObject ); // Rappro CRE/ECR

    // Sous Menu Liasse Fiscale
    procedure OnClickDonneesComp        ( Sender : TObject ); // Données complémentaires
    procedure OnClickInfoLiasse         ( Sender : TObject ); // information de l'affectation du compte dans les liasses
    procedure OnClickTotalCptLiasse     ( Sender : TObject ); // Affectation des montants

    procedure OnClickGCD                ( Sender : TObject ); // Créer une créance dans GCD
    procedure OnClickTableauGCD         ( Sender : TObject ); // Tableau de gestion GCD
    procedure OnClickFeuilleExcel       ( Sender : TObject ); // Feuille Excel

    procedure OnClickNoteTravail        ( Sender : TObject ); // Note de travail
    procedure OnClickTableauVariation   ( Sender : TObject ); // Tableau de variation
    procedure OnClickDRF                ( Sender : TObject ); // DRF
    procedure OnClickVisaCompte         ( Sender : TObject ); // Visa sur le compte

    // Sous-Menu SUIVIREVISION
    procedure OnClickDocPermanent       ( Sender : TObject ); // Document permanent
    procedure OnClickDocAnnuel          ( Sender : TObject ); // Document annuel
    procedure OnClickRechercheDoc       ( Sender : TObject ); // Recherche de document

    procedure OnClickProgrammeTravail   ( Sender : TObject ); // Programme de travail (Réviseur (N2) ou Superviseur (N3 N4))
    procedure OnClickAPG                ( Sender : TObject ); // Appréciation générale du dossier
    procedure OnClickSCY                ( Sender : TObject ); // Synthèse des cycles
    procedure OnClickEXP                ( Sender : TObject ); // Attestation expert

    // Sous-Menu DOCTRAVAUX
    procedure OnClickMemoCycle          ( Sender : TObject ); // Mémo cycle de révision de la documentation des travaux
    procedure OnClickMemoObjectif       ( Sender : TObject ); // Mémo Objectif de révision de la documentation des travaux
    procedure OnClickMemoSynthese       ( Sender : TObject ); // Mémo synthèse
    procedure OnClickMemoMillesime      ( Sender : TObject ); // Mémo millesime
    procedure OnClickMemoCompte         ( Sender : TObject ); // Mémo compte général

    procedure OnClickControleTva        ( Sender : TObject ); // Contrôle de TVA
    procedure OnClickRapproComptaGCD    ( Sender : TObject ); // Rapprochement Comptabilité / GCD
    procedure OnClickDocumentGED      ( Sender : TObject ); // Affiche du document numérisé de l'écriture
{$ENDIF}

    procedure PassageEcrSimplifie       ( Sender : TObject ); // Passage d'écriture  simplifiée

    // Lettrage
    procedure EtudieModeLettrage      (var R: RLETTR);
    procedure LettrageMClick          ( Sender : TObject ); // Lettrage manuel
    procedure LettrageAClick          ( Sender : TObject ); // Lettrage automatique
    procedure DelettreClick           ( Sender : TObject ); // Delettrage manuel
    procedure DelettreMvtExoRef       ( Sender : TObject ); // Delettrage MVT sur EXOREF
    procedure DeletTotal              ( Sender : TObject );

    procedure OnClickCutOff           ( Sender : TObject) ; // ecriture de cutoff

    procedure CreationLienImmoClick   ( Sender : TObject ); // Création du Lien Immo
    procedure SupprimeLienImmoClick   ( Sender : TObject ); // Suppression du lien IMMO

    // Traitements Comptables
    procedure SaisieEcrClick          ( Sender : TObject ); // Saisie d'une Piece
    procedure SaisieBorClick          ( Sender : TObject ); // Saisie d'un Bordereau

    procedure OnClickCompteGeneral    ( Sender : TObject ); // Paramètres du compte général
    procedure OnClickCompteAuxiliaire ( Sender : TObject ); // Paramètres du compte auxiliaire
    procedure RechercheEcrituresClick ( Sender : Tobject ); // Recherche d'ecritures

    procedure AnalytiquesClick        ( Sender : TObject ); // Mouvements analytiques
    procedure RepartAnalytiquesClick  ( Sender : TObject ); // Mouvements analytiques
    procedure VentilAnalytiques1Click ( Sender : TObject );
    procedure VentilAnalytiques2Click ( Sender : TObject );
    procedure VentilAnalytiques3Click ( Sender : TObject );
    procedure VentilAnalytiques4Click ( Sender : TObject );
    procedure VentilAnalytiques5Click ( Sender : TObject );
    procedure VentilAnalytiqueAxe     ( pstAxe : string );

    procedure OnClickReclass          ( Sender : TObject );
    procedure OnErrorTOB              ( Sender : TObject ; Error : TRecError );

    // Utilitaires
    procedure OnClickInfosComp        ( Sender : TObject ); // Informations complementaires de l'ecriture

    procedure CumulsGENEClick         ( Sender : TObject ); // Cumul du compte général
    procedure CumulsAUXClick          ( Sender : TObject ); // Cumul du compte auxiliaire

    procedure PrepareArgumentGL       ( vBoJustifSolde : Boolean);
    procedure PrepareArgumentBAL;
    procedure OnClickJustifSolde      ( Sender : TObject );

    // Editions du Pointage
    procedure OnClickEtatRapprochement( Sender : TObject ); // Etat de Rapprochemment bancaire
    procedure OnClickJustifSoldeBQE   ( Sender : TObject ); // Etat de Justificatif Bancaire

    // Editions des Grands-livres
    procedure OnClickGLGene           ( Sender : TObject ); // Grand Livre Général
    procedure OnClickGLGeneParAuxi    ( Sender : TObject ); // Grand Livre Général par Auxiliaire
    procedure OnClickGLAuxi           ( Sender : TObject ); // Grand Livre Auxiliaire
    procedure OnClickGLAuxiParGene    ( Sender : TObject ); // Grand Livre Auxiliaire par Général
    procedure OnClickGLAuxiPourGene   ( Sender : TObject ); // Grand Livre Auxiliaire pour Général
    procedure OnClickGLGeneParQte     ( Sender : TObject ); // Grand Livre général par Quantité

    // Grand-livre Analytique
    procedure LancementGLANA          ( vFichierBase : TFichierBase );
    procedure OnClickGLAnaSurAxe1     ( Sender : TObject ); // Grand Livre analytique sur Axe 1
    procedure OnClickGLAnaSurAxe2     ( Sender : TObject ); // Grand Livre analytique sur Axe 2
    procedure OnClickGLAnaSurAxe3     ( Sender : TObject ); // Grand Livre analytique sur Axe 3
    procedure OnClickGLAnaSurAxe4     ( Sender : TObject ); // Grand Livre analytique sur Axe 4
    procedure OnClickGLAnaSurAxe5     ( Sender : TObject ); // Grand Livre analytique sur Axe 5

    // Grand-Livre général par analytique
    procedure LancementGLGENEPARANA   ( vFichierBase : TFichierBase );
    procedure OnClickGLGeneParAna1    ( Sender : TObject ); // Grand Livre général par section analytique sur Axe 1
    procedure OnClickGLGeneParAna2    ( Sender : TObject ); // Grand Livre général par section analytique sur Axe 2
    procedure OnClickGLGeneParAna3    ( Sender : TObject ); // Grand Livre général par section analytique sur Axe 3
    procedure OnClickGLGeneParAna4    ( Sender : TObject ); // Grand Livre général par section analytique sur Axe 4
    procedure OnClickGLGeneParAna5    ( Sender : TObject ); // Grand Livre général par section analytique sur Axe 5

   // Balance
    procedure OnClickBalGene          ( Sender : TObject ); // Balance Générale
    procedure OnClickBalGeneParAuxi   ( Sender : TObject ); // Balance Générale par Auxiliaire
    procedure OnClickBalAuxi          ( Sender : TObject ); // Balance Auxiliaire
    procedure OnClickBalAuxiParGene   ( Sender : TObject ); // Balance Auxiliaire par Général

    // Balance générale par analytique
    procedure LancementBalGeneParAna  ( vFichierBase : TFichierBase );
    procedure OnClickBalGeneParAna1   ( Sender : TObject ); // Balance générale par section analytique sur Axe 1
    procedure OnClickBalGeneParAna2   ( Sender : TObject ); // Balance générale par section analytique sur Axe 2
    procedure OnClickBalGeneParAna3   ( Sender : TObject ); // Balance générale par section analytique sur Axe 3
    procedure OnClickBalGeneParAna4   ( Sender : TObject ); // Balance générale par section analytique sur Axe 4
    procedure OnClickBalGeneParAna5   ( Sender : TObject ); // Balance générale par section analytique sur Axe 5

    // Balance analytique
    procedure LancementBalAnaSurAxe   ( vFichierBase : TFichierBase );
    procedure OnClickBalAnaSurAxe1    ( Sender : TObject ); // Balance Analytique sur l'axe 1
    procedure OnClickBalAnaSurAxe2    ( Sender : TObject ); // Balance Analytique sur l'axe 2
    procedure OnClickBalAnaSurAxe3    ( Sender : TObject ); // Balance Analytique sur l'axe 3
    procedure OnClickBalAnaSurAxe4    ( Sender : TObject ); // Balance Analytique sur l'axe 4
    procedure OnClickBalAnaSurAxe5    ( Sender : TObject ); // Balance Analytique sur l'axe 5    procedure OnClickBalAnaSurAxe5    ( Sender : TObject ); // Balance Analytique sur l'axe 1

    procedure OnClickModifSerie       ( Sender : TObject );
    procedure OnClickAffecteAnaEnSerie( Sender : TObject );
    procedure OnClickEcrVersClipboard ( Sender : TObject ); // Copie des écritures dans le presse-papier


{$ENDIF}






{$IFDEF AMORTISSEMENT}
    procedure OnClickRapproComptaImmo ( Sender : TObject ); // Rapprochement Compta / Immo
    procedure ListeImmoClick          ( Sender : TObject ); // Liste des immos du compte
    procedure FicheImmoClick          ( Sender : TObject ); // Fiche Immo du compte
{$ENDIF}

    ////////////////////////////////////////////////////////////////////////////
    function GetPrecedent : TExoDate;
    function GetEnCours   : TExoDate;
    function GetSuivant   : TExoDate;
    function GetExercices( vValue : integer ) : TExoDate;
    function GetExoV8     : TExoDate;
    function GetExoRef    : TExoDate;
    function GetEntree    : TExoDate;
    function GetInfoCpta( vValue : TFichierBase ) : TInfoCpta;

   {JP 26/06/07 : Vérifie si la sélection contient des lignes GC ou Tréso}
    function HasEcrImport : Boolean;

  protected
    procedure InitControl;                         override; // Init des composants de la fiche
    procedure RemplitATobFListe;                   override;
    procedure RefreshFListe( vBoFetch : Boolean ); override; // Recharge la grille avec le PutGridDetail
    function  AjouteATobFListe   ( vTob : Tob ): Boolean; override;
    procedure CalculPourAffichage( vTob : Tob );   override;
    function  BeforeLoad: boolean;                 override;
    function  AfterLoad : boolean;                 override;

    procedure ApresChangementFiltre;               override;
    procedure NouveauFiltre;                       override;

  public
    property BoGesComTTC : Boolean read HasEcrImport;
  end;

 TOF_PARAMCONS = Class (TOF)
 private
  FStGeneral    : string ;
  FStNatureGene : string ;
  procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState) ;
  procedure GeneralElipsisClick(Sender: TObject);
  procedure AuxElipsisClick(Sender: TObject);
  procedure GeneralKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
  procedure AuxKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
  procedure BValiderClick(Sender: TObject);
  procedure GeneralExit(Sender: TObject);
 public

  E_General    : THEdit;
  E_Auxiliaire : THEdit;

  procedure OnArgument (S : String ) ; override ;
  procedure OnClose                  ; override ;
 end ;

procedure OperationsSurComptes ( LeGene,LeExo,ChampsTries : string; LeAux : string = '' ; FromSaisie : Boolean = False; vStNombase : string = ''; vBoJustifSolde : Boolean = False; vStDefilCpt : string = '') ;
function  OperationsSurComptesDetail ( LeGene,LeExo,ChampsTries : string; LeAux : string = '' ; FromSaisie : Boolean = False ) : Variant ;

implementation

uses uTilSoc,       // ChargePageSoc
     uLibWindows,   // AfficheDBCR
     ParamSoc,      // GetParamSocSecur
     LettUtil,      // Fonctions de Delettrage et Lettrage
     SaisComm,      // WhereEcriture
     FichComm,      // FicheRIB_AGL
     AGLInit,       // TheData
     uTilSais,      // QuelZoomTableTNat
     ClipBrd,       // Clipboard
     Constantes,    // JP 26/06/07 : QUALIFTRESO
     CPChancell_TOF,// FicheChancel
     Devise_TOM,    // FicheDevise

{$IFDEF MODENT1}
  CPProcMetier,
  CPVersion,
{$ENDIF MODENT1}

{$IFDEF COMPTA}
     ULibPointage,          // JP 30/07/07 : MsgPointageSurTreso
     CPMULANA_TOF,          // MultiCritereAnaZoom
     ZReclassement,
     Cummens,               // CumulCpteMensuel
     SaisComp,              // R_COMP
     Saisie,                // TrouveSaisie
     SaisODA,               // TrouveEtLanceSaisieODA
     uLibAnalytique,        // CChargeAna
     CPTiers_TOM,           // FicheTiers
     CPGeneraux_TOM,        // FicheGene
     SaisBor,               // LanceSaisieFolio
     RappAuto,              // Rapprochement automatique
     uTofCutOff,            // CPLanceFiche_CPMULCUTOFF
     uTofCPMulMvt,          // MultiCritereMvt
     uTofRechercheEcr,      // CPLanceFiche_CPRechercheEcr
     MZSUtil,               // ModifieEnSerieEcriture
     AffecteAnaEnSerie_TOF, // CPLanceFiche_AffecteAnalytiqueEnSerie
     Lettrage,              // LettrageManuel
     UTOFECRLET,            // Passage ecriture simplifiée

  {$IFNDEF CCMP}
     uTOFHistoCpte,         // CC_LanceFicheHistoCpte;
     CPGESTIONCREANCE,
     uTofPointageEcr,       // CPLanceFiche_Pointage
     CPPOINTAGERAP_TOF,     {JP 12/07/07 : FQ 21045 : nouveau pointage : CPLanceFiche_PointageRappro}
     uTOFPointageMul,       // CPLanceFiche_PointageMul
     uImportBob_TOF,        // CPLanceFiche_CPIMPFICHEXCEL
     uTOFSaisieEteBac,      // CPLanceFiche_SaisieTresorerie
     FPRapproCRE_TOF,       // Rapprochement Compta / CRE
     CPControleLiasse_TOF,  // CPLanceFiche_CPControleLiasse
     CPTotalCptLiasse_TOF,  // CPLancefiche_CPTotalCptLiasse
     uTofCPTableauVar,      // CPLanceFiche_CPTableauVar
     uTofCPNoteTravail,     // CPLanceFiche_CPNoteTravail
     uTofCPResultatfiscal,  // LanceFiche_DRF()
     uLibRevision,          // MiseAJourVisaCompte
     CPREVDocTravaux_TOF,   // CPLanceFiche_CPRevDocTravaux
     CPREVProgTravail_TOF,  // CPLanceFiche_CPProgTravail
     CRevInfoDossier_TOM,   // CPLanceFiche_CPREVINFODOSSIER
     CPControlTva_TOF,      // CPLanceFiche_CPCONTROLTVA
     ComObj,                // CreateOleObject
     uCPGedFileViewer,      // CPLanceFiche_CPGedFileViewer
     utilGed,               // Extraitdocument
     uGedViewer,            // ShowGedViewer
  {$ENDIF}

  {$IFDEF AMORTISSEMENT}
     ImRapCpt,              // EtatRapprochementCompta
     AmListe_TOF,           // AMLanceFiche_ListeDesImmobilisations
     Immo_Tom,              // AMLanceFiche_FicheImmobilisation
  {$ENDIF}

  // Editions
  uTofCPGLGene,     // CPLanceFiche_CPGLGENE
  uTofCPGLAuxi,     // CPLanceFiche_CPGLAUXI
  uTofCPGLAna,      // CPLanceFiche_CPGLGANA
  CPBalGen_Tof,     // CPLanceFiche_BalanceGeneral
  CPBalAuxi_Tof,    // CPLanceFiche_BalanceAuxiliaire
  CPBALGENAUXI_Tof, // CPLanceFiche_BalanceGenAuxi
  CPBALAUXIGEN_Tof, // CPLanceFiche_BalanceAuxiGen
  CPBalAnal_Tof,    // CPLanceFiche_BalanceAnalytique
  CPBalGenAnal_Tof, // CPLanceFiche_BalanceGenAnal
  CPRAPPRODET_Tof,  // CC_LanceFicheEtatRapproDet
  CPJUSTBQ_TOF,     // CC_LanceFicheJustifPointage
  CPREPARTITIONANA_TOF, //CPLanceFiche_RepartitionAnalytique
{$IFNDEF CCMP}
  CPRAPPRODETV7_Tof,    //CC_LanceFicheEtatRapproDetV7
  CPAVERTNEWRAPPRO_TOF, //OkNewPointage
{$ENDIF}
{$ENDIF}

{$IFNDEF IMP}
  BanqueCP_TOM,     // FicheBanqueCP
{$ENDIF}

  uTofMulParamGen;  // YMO - 23/04/07 - F5 sur Auxiliaire

//------------------------------------------------------------------------------
// Clé primaire Table des écritures :
// E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE
//------------------------------------------------------------------------------

var GTobDetailEcr : Tob;

const
  cFI_TABLE = 'CPCONSECR';

  cSelect = ', E_EXERCICE CLE_EXERCICE, E_NUMLIGNE CLE_NUMLIGNE, ' +
            'E_NUMECHE CLE_NUMECHE, E_QUALIFPIECE CLE_QUALIFPIECE, ' + // Champ de la clé obligatoire
            'E_PERIODE CLE_PERIODE, E_ECRANOUVEAU CLE_ECRANOUVEAU, ' + // Champ nécessaire pour la centralisation
            'E_LETTRAGE CLE_LETTRAGE, E_IMMO CLE_IMMO, ' +
            'E_QUALIFORIGINE CLE_QUALIFORIGINE ';

  cSelectGed = ', E2.EC_DOCGUID CLE_DOCGUID ';

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/01/2004
Modifié le ... : 04/02/2004
Description .. : Procédure de lancement compatible avec l'ancien source ConsEcr.Pas
Suite ........ : afin d'éviter toutes modifs dans les sources des autres projets
Mots clefs ... :
*****************************************************************}
procedure OperationsSurComptes ( LeGene,LeExo,ChampsTries : string; LeAux : string = '' ; FromSaisie : Boolean = False; vStNombase : string = ''; vBoJustifSolde : Boolean = False; vStDefilCpt : string = '') ;
var lStArgument : string;
begin
  if FromSaisie then
    lStArgument := LeGene + ';' + LeExo + ';' + ChampsTries + ';' + LeAux + ';' + 'X'
  else
    lStArgument := LeGene + ';' + LeExo + ';' + ChampsTries + ';' + LeAux + ';' + '-';

  // Gestion du paramètre RenvoiDetail ( toujours à false par ce point d'entrée
  lStArgument := lStArgument + ';-';

  // Gestion du nom de la base sur laquelle on doit éxécuter la requête SELECT
  lStArgument := lStArgument + ';' + vStNomBase;

  if vBoJustifSolde then
    lStArgument := lStArgument + ';X'
  else
    lStArgument := lStArgument + ';-';

  lStArgument := lStArgument + ';' + vStDefilCpt;

  AGLLanceFiche('CP', 'CPCONSECR', '', '', lStArgument );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function OperationsSurComptesDetail( LeGene,LeExo,ChampsTries : string; LeAux : string = '' ; FromSaisie : Boolean = False ) : Variant ;
var lStArgument : string;
    lStNomChamp : string;
    lStResult   : string;
    i, j        : integer;
begin
  // A mettre pour Debug depuis la compta
  //V_Pgi.ZoomOle := True;
  try
    lStResult := '';
    if not V_Pgi.ZoomOle then Exit;
    GTobDetailEcr := Tob.Create('', nil, -1);

    if FromSaisie then
      lStArgument := LeGene + ';' + LeExo + ';' + ChampsTries + ';' + LeAux + ';' + 'X'
    else
      lStArgument := LeGene + ';' + LeExo + ';' + ChampsTries + ';' + LeAux + ';' + '-';

    // Renvoi des écritures dans GTobDetailEcr
    lStArgument := lStArgument + ';X';

    AGLLanceFiche('CP', 'CPCONSECR', '', '', lStArgument );

{$IFDEF TT}
    GTobDetailEcr.SaveToFile('C:\GTobDetailEcr.Txt', False, True, True);
{$ENDIF}

    // Conversation de GTobDetailEcr en un Variant contenant les champs séparés
    // par des ;
    lStResult := '';
    if GTobDetailEcr.Detail.Count > 0 then
    begin
      // Ajout de l'entête GENERAL et AUXILIAIRE dans le RESULT
      lStResult := GTobDetailEcr.GetString('GENERAL') + ';' +
                   GTobDetailEcr.GetString('AUXILIAIRE') + ';' + #10;
      // Fin d'ajout de l'entête

      // Envoi de la liste des chanmps dans le RESULT
      for j := 1000 to (1000 + GTobDetailEcr.Detail[0].ChampsSup.Count)-1 do
      begin
        lStNomChamp := GTobDetailEcr.Detail[0].GetNomChamp(j);
        lStResult := lStResult + VarToStr(lStNomChamp) + ';';
      end;
      lStResult := lStResult + #10;

      // Envoi de la liste des ecritures
      for i := 0 to GTobDetailEcr.Detail.Count -1 do
      begin
        for j := 1000 to (1000 + GTobDetailEcr.Detail[i].ChampsSup.Count)-1 do
        begin
          lStNomChamp := GTobDetailEcr.Detail[i].GetNomChamp(j);
          lStResult   := lStResult + VarToStr(GTobDetailEcr.Detail[i].GetValeur(j)) + ';';
        end;
        lStResult := lStResult + #10;
      end;
    end; // if GTobDetailEcr.Detail.Count > 0 then

  finally
    FreeAndNil( GTobDetailEcr );
    Result :=  lStResult ;
    // A mettre pour Debug depuis la compta
    //V_Pgi.ZoomOle := False;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnNew;
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/09/2004
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.OnDelete;
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/09/2004
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnUpdate;
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/09/2004
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.OnClose;
begin
  FreeAndNil(FTobEcrSelect);
  FreeAndNil(FTobListeEcr);
  FreeAndNil(FTobGene);
  FreeAndNil(FTobTiers);
  FreeAndNil(FTobListeICC);
  FreeAndNil(FZListJournal);
{$IFDEF BUREAU}
  FreeAndNil(FZExercice);
  FreeAndNil(FZInfoCpta);
{$ENDIF}
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gamin
Créé le ...... : 26/04/2004
Modifié le ... :   /  /
Description .. : - LG - 26/04/2004 - ajout de la bi compatibilte de socref
Suite ........ : avec le champs EGENERAL et E_GENERAL
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnArgument(S: string);
begin
  FFI_Table := cFI_TABLE;
  FStListeParam := 'CPCONSECR';
  inherited;
  IndiceColfListe;

  // Récupération des arguments
  FStArgumentTOF := S;

  // Elements de la fiche Ancêtre
  FListe.GetCellCanvas   := GetCellCanvasFListe;
  FListe.OnDblClick      := OnDblClickFListe;
  FListe.OnRowEnter      := OnRowEnterFListe;

  FListe.OnKeyDown       := OnKeyDownFListe;
  FListe.OnBeforeFlip    := OnBeForeFlipFListe;
  FListe.OnFlipSelection := OnFlipSelectionFListe;
  FListe.PostDrawCell    := OnPostDrawCellFListe;

  // Elements de la TOF
  FTobEcrSelect    := Tob.Create('ECRITURE', nil, -1);
  FTobListeEcr     := Tob.Create('', nil, -1);
  FTobGene         := Tob.Create('GENERAUX', nil, -1);
  FTobTiers        := Tob.Create('TIERS', nil, -1);
  FZListJournal    := TZListJournal.Create;

  GLibelle         := THLabel(GetControl('GLIBELLE'));
  TLibelle         := THLabel(GetControl('TLIBELLE'));

  BGUP               := TToolBarButton97(GetControl('BGUP', True));
  BGDOWN             := TToolBarButton97(GetControl('BGDOWN', True));
  BTUP               := TToolBarButton97(GetControl('BTUP', True));
  BTDOWN             := TToolBarButton97(GetControl('BTDOWN', True));

  BComptable         := TToolBarButton97(GetControl('BCOMPTABLE', True));
  BSpecifique        := TToolBarButton97(GetControl('BSPECIFIQUE', True));
  BAideRevision      := TToolBarButton97(GetControl('BAIDEREVISION', True));
  BRevInteg          := TToolBarButton97(GetControl('BREVINTEG', True));
  BParametre         := TToolBarButton97(GetControl('BPARAMETRE', True));
  BEdition           := TToolBarButton97(GetControl('BEDITION', True));
  BFctMultiSelect    := TToolBarButton97(GetControl('BFctMultiSelect', True));

  PopUpComptable     := TPopUpMenu(GetControl('PopUpComptable', True));
  PopUpSpecifique    := TPopUpMenu(GetControl('PopUpSpecifique', True));
  PopUpAideRevision  := TPopUpMenu(GetControl('PopUpAideRevision', True));
  PopUpRevInteg      := TPopUpMenu(GetControl('PopUpRevInteg', True));
  PopUpParametre     := TPopUpMenu(GetControl('PopUpParametre', True));
  PopUpMultiSelect   := TPopUpMenu(GetControl('PopUpMultiSelect', True));
  PopUpEdition       := TPopUpMenu(GetControl('PopUpEdition', True));

  E_General        := THEdit(GetControl('EGENERAL', True));
  E_Auxiliaire     := THEdit(GetControl('EAUXILIAIRE', True)) ;
  E_DateComptable  := THEdit(GetControl('E_DATECOMPTABLE', True));
  E_DateComptable_ := THEdit(GetControl('E_DATECOMPTABLE_', True));
  LESOLDE          := THEdit(GetControl('LESOLDE', True));

  // GCO - 26/04/2006 - FQ 12279
  LESOLDEN         := THEdit(GetControl('LESOLDEN', True));
  LESOLDENMOINS1   := THEdit(GetControl('LESOLDENMOINS1', True));
  // FIN GCO

  E_RefPointage    := THEdit(GetControl('E_REFPOINTAGE', True));
  E_RefPointage_   := THEdit(GetControl('E_REFPOINTAGE_', True));
  E_DatePointage   := THEdit(GetControl('EDATEPOINTAGE', True));
  E_Lettrage       := THEdit(GetControl('E_LETTRAGE', True));
  E_NumeroPiece    := THEdit(GetControl('E_NumeroPiece', True));
  E_NumeroPiece_   := THEdit(GetControl('E_NumeroPiece_', True));
  ELibelle         := THEdit(GetControl('ELIBELLE', True));
  ERefInterne      := THEdit(GetControl('EREFINTERNE', True));
  EMontant         := THEdit(GetControl('EMONTANT', True));

  ESoldeBancaire   := THNumEdit(GetControl('ESOLDEBANCAIRE', True));

  CBJustifSolde    := TCheckBox(GetControl('CBJUSTIFSOLDE', True));
  CBCentralisation := TCheckBox(GetControl('CBCENTRALISATION', True));
  CBDetailANO      := TCheckBox(GetControl('CBDETAILANO', True));
  CBAvecCollectif  := TCheckBox(GetControl('CBAVECCOLLECTIF', True));
  CBDocumentGED    := TCheckBox(GetControl('CBDOCUMENTGED', True));

  E_Valide         := TCheckBox(GetControl('E_VALIDE', True));

  E_Journal        := THMultiValComboBox(GetControl('E_Journal', True));
  E_NaturePiece    := THMultiValComboBox(GetControl('E_NaturePiece', True));
  E_Devise         := THValComboBox(GetControl('EDEVISE'));
  E_Exercice       := THValComboBox(GetControl('E_EXERCICE', True));
  CInitComboExercice(E_EXERCICE);

  EQualifPiece    := THMultiValComboBox(GetControl('E_QUALIFPIECE', False));
  if EQualifPiece = nil then
  begin
    THMultiValComboBox(GetControl('E_QUALIFPIECE')).Name := 'EQUALIFPIECE';
    EQualifPiece   := THMultiValComboBox(GetControl('EQUALIFPIECE', False));
  end;
  EQualifPiece.ReadOnly := True;

  E_Etablissement    := THMultiValComboBox(GetControl('EETABLISSEMENT', True));
  E_EtatLettrage     := THMultiValComboBox(GetControl('E_EtatLettrage', True));

  ComboInfoSup       := THValComboBox(GetControl('COMBOINFOSUP', True));
  ComboAcces         := THValComboBox(GetControl('COMBOACCES', True));
  ComboDefilCpt      := THValComboBox(GetControl('COMBODEFILCPT', True));
  ComboSoldePro      := THValComboBox(GetControl('COMBOSOLDEPRO', True));

  ComboMontant       := THValComboBox(GetControl('COMBOMONTANT', True));
  ComboOpMontant     := THValComboBox(GetControl('COMBOOPMONTANT', True));
  ComboOpLibelle     := THValComboBox(GetControl('COMBOOPLIBELLE', True));
  ComboOpRefInterne  := THValComboBox(GetControl('COMBOOPREFINTERNE', True));

  PageInfoSup      := TPageControl(GetControl('PAGEINFOSUP', True));
  TabTablesLibres  := TTabSheet(GetControl('TABTABLESLIBRES', True));

{$IFDEF EAGLCLIENT}
  CBCentralisation.Visible := False;
{$ENDIF}

  //----------------------- Branchement des événements -----------------------//
  Ecran.OnKeyDown     := OnKeyDownEcran;
  E_Exercice.OnChange := OnChangeComboE_Exercice;

  BGUP.OnClick         := OnClickBGUP;
  BGDOWN.OnClick       := OnClickBGDOWN;
  BTUP.OnClick         := OnClickBTUP;
  BTDOWN.OnClick       := OnClickBTDOWN;

  E_General.OnElipsisClick := OnElipsisClickE_General;
  E_General.OnDblClick     := OnElipsisClickE_General;
  E_General.OnExit         := OnExitE_General;

  E_Auxiliaire.OnElipsisClick   := OnElipsisClickE_Auxiliaire;
  E_Auxiliaire.OnDblClick       := OnElipsisClickE_Auxiliaire;
  E_Auxiliaire.OnExit           := OnExitE_Auxiliaire;

  E_DateComptable.OnExit        := OnExitE_DateComptable;

  E_RefPointage.OnElipsisClick  := OnElipsisClickE_RefPointage;
  E_RefPointage.OnDblClick      := OnElipsisClickE_RefPointage;

  E_RefPointage_.OnElipsisClick := OnElipsisClickE_RefPointage;
  E_RefPointage_.OnDblClick     := OnElipsisClickE_RefPointage;

  E_DatePointage.OnElipsisClick := OnElipsisClickE_DatePointage;
  E_DatePointage.OnDblClick     := OnElipsisClickE_DatePointage;
  E_DatePointage.OnExit         := OnExitE_DatePointage;

  CBJustifSolde.OnClick         := OnClickCBJustifSolde;
  CBDetailAno.OnClick           := OnClickCBDetailAno;

  ComboInfoSup.OnChange         := OnChangeComboInfoSup;
  ComboAcces.OnChange           := OnChangeComboAcces;

  PopF11.OnPopup              := OnPopUpPopF11;
  PopUpComptable.OnPopup      := OnPopUpComptable;
  PopUpSpecifique.OnPopup     := OnPopUpSpecifique;
  PopUpAideRevision.OnPopup   := OnPopUpAideRevision;
  PopUpRevInteg.OnPopup       := OnPopUpRevInteg;
  PopUpParametre.OnPopup      := OnPopUpParametre;
  PopUpMultiSelect.OnPopUp    := OnPopUpMultiSelect;
  PopUpEdition.OnPopUp        := OnPopUpEdition;

  // Gestion de la Multi Sélection
  FListe.MultiSelect           := True;
  BFctMultiSelect.Enabled      := False;

{$IFDEF COMPTA}
{$ELSE}
    PopF11.Items.Clear;
    BComptable.Visible         := False;
    BSpecifique.Visible        := False;
    BAideRevision.Visible      := False;
    BRevInteg.Visible          := False;
    BEdition.Visible           := False;
    BFctMultiSelect.Visible    := False;
{$ENDIF}

  THLabel(GetControl('TLESOLDENMOINS1')).Caption := TraduireMemoire('Solde précédent');
  THLabel(GetControl('TLESOLDEN')).Caption       := TraduireMemoire('Solde en cours');

  // GCO - 29/08/2006 - FQ 18715
  THEdit(GetControl('E_TABLE0', True)).DataType := 'TZNATECR0';
  THEdit(GetControl('E_TABLE1', True)).DataType := 'TZNATECR1';
  THEdit(GetControl('E_TABLE2', True)).DataType := 'TZNATECR2';
  THEdit(GetControl('E_TABLE3', True)).DataType := 'TZNATECR3';
  // FIN GCO

  if GetParamSocSecur('SO_CPMULTIERS', false) then
    E_Auxiliaire.OnElipsisClick:=AuxiElipsisClick;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/08/2002
Modifié le ... :   /  /
Description .. : Initialise les contrôles avec leurs valeurs par défaut pour
Suite ........ : l'ouverture de l'écran ou une nouvelle recherche
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.InitControl;
var lSt: string;
begin
  GLibelle.Caption    := '';
  TLibelle.Caption    := '';
  LeSolde.Text        := '0,00';
  LeSoldeN.Text       := '0,00';
  LeSoldeNMoins1.Text := '0,00';

  ESoldeBancaire.Text := '0,00';

  PageControl.ActivePage := PageControl.Pages[0];
  PageControl.Pages[3].TabVisible := False;
  PageControl.Pages[4].TabVisible := False;
  CBJustifSolde.Enabled := False;

  lSt := TraduireMemoire('<<Tous>>');

  E_Journal.SelectAll;
  if E_Journal.Tous then
    E_Journal.Text := lSt;

  E_NaturePiece.SelectAll;
  if E_NaturePiece.Tous then
    E_NaturePiece.Text := lSt;

  EQualifPiece.Text := 'N;';

  E_Etablissement.SelectAll;
  if E_Etablissement.Tous then
    E_Etablissement.Text := lSt;

  E_Devise.ItemIndex := 0;

  E_EtatLettrage.SelectAll;
  if E_EtatLettrage.Tous then
    E_EtatLettrage.Text := lSt;

  CBAvecCollectif.Checked := True;

  ComboInfoSup.Value := '0';

  // Si Enabled = False, on a force la consultation car on vient de la saisie,
  // donc pas touche à la Value := 0;
  if ComboAcces.Enabled then
    ComboAcces.Value := '1';

  if ComboDefilCpt.Value = '' then
    ComboDefilCpt.Value := 'ALL';

  ComboSoldePro.Value := '1';

  OnChangeComboInfoSup(Self);
  OnChangeComboAcces(Self);

  // GCO - 14/02/2006
  ComboMontant.Value   := '(E_DEBIT OR E_CREDIT)';

  ComboOpMontant.Plus  := ' AND CO_CODE <> "C" AND CO_CODE <> "D"' +
                          ' AND CO_CODE <> "L" AND CO_CODE <> "M"' +
                          ' AND CO_CODE <> "V" ANd CO_CODE <> "VV"';

  ComboOpMontant.Value := '=';

  ComboOpLibelle.Plus  := ' AND CO_CODE <> "<" AND CO_CODE <> "<="' +
                          ' AND CO_CODE <> ">" AND CO_CODE <> ">="' +
                          ' AND CO_CODE <> "E" AND CO_CODE <> "G"' +
                          ' AND CO_CODE <> "I" AND CO_CODE <> "J"';
  ComboOpLibelle.Value := '=';

  ComboOpRefInterne.Plus  := ComboOpLibelle.Plus;
  ComboOpRefInterne.Value := '=';
  // FIN GCO

  SetControlVisible( 'EC_CUTOFFDEB' , false ) ;
  SetControlVisible( 'EC_CUTOFFFIN' , false ) ;
  SetControlVisible( 'TZ_AFFAIRE1'  , false ) ;
  SetControlVisible( 'TZ_AFFAIRE2'  , false ) ;

  // GCO - 03/12/2007
  CBDocumentGed.Visible := (CtxPcl in V_Pgi.PGIContexte);

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/08/2002
Modifié le ... :   /  /
Description .. : Mise à jour du Caption de l'écran en fonction de la recherche
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.MiseAJourCaptionEcran;
begin
  Ecran.Caption := TraduireMemoire('Compte') + ' : ';

  if (E_GENERAL.Text <> '') and (E_AUXILIAIRE.Text <> '') then
  begin
    Ecran.Caption := Ecran.Caption + E_GENERAL.Text + '  ' + E_AUXILIAIRE.Text + '  ' +
                     TLIBELLE.Caption + '   ' + LESOLDE.Text + '  du ' +
                     E_DateComptable.Text + ' au ' + E_DateComptable_.Text;
  end
  else
  begin
    if E_AUXILIAIRE.Text <> '' then
      Ecran.Caption := Ecran.Caption + E_AUXILIAIRE.Text + '  ' + TLIBELLE.Caption + '   ' +
                       LESOLDE.Text + '  du ' + E_DateComptable.Text + ' au ' +
                       E_DateComptable_.Text
    else
      if E_GENERAL.Text <> '' then
        Ecran.Caption := Ecran.Caption + E_GENERAL.Text + '  ' + GLIBELLE.Caption + '   ' +
                         LESOLDE.Text + '  du ' + E_DateComptable.Text + ' au ' +
                         E_DateComptable_.Text;
  end;
  (*
  LeSoldeN.Text          := '0,00';
  LeSoldeNMoins1.Text    := '0,00';

  SetControlVisible('LESOLDEN',       (E_General.Text = '') or (E_Auxiliaire.Text = ''));
  SetControlVisible('LESOLDENMOINS1', (E_General.Text = '') or (E_Auxiliaire.Text = ''));
  SetControlVisible('TLESOLDEN',      (E_General.Text = '') or (E_Auxiliaire.Text = ''));
  SetControlVisible('TLESOLDENMOINS1',(E_General.Text = '') or (E_Auxiliaire.Text = ''));

  if (E_General.Text <> '') and (E_Auxiliaire.Text = '') then
  begin
    LeSoldeN.Text       := AfficheDBCR(FTobGene.GetDouble('G_TOTDEBE') - FTobGene.GetDouble('G_TOTCREE'));
    LeSoldeNMoins1.Text := AfficheDBCR(FTobGene.GetDouble('G_TOTDEBP') - FTobGene.GetDouble('G_TOTCREP'));
  end;

  if (E_General.Text = '') and (E_Auxiliaire.Text <> '') then
  begin
    LeSoldeN.Text       := AfficheDBCR(FTobTiers.GetDouble('T_TOTDEBE') - FTobTiers.GetDouble('T_TOTCREE'));
    LeSoldeNMoins1.Text := AfficheDBCR(FTobTiers.GetDouble('T_TOTDEBP') - FTobTiers.GetDouble('T_TOTCREP'));
  end;*)

  UpDateCaption(Ecran);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :   /  /    
Modifié le ... : 28/06/2007
Description .. : - LG - 28/06/2007 - FB 18429 - on applique pas le filtre qd 
Suite ........ : on vient de la consultation des generaux/auxilaires
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.OnLoad;
var
 lStGen : string ;
 lStAux : string ;
begin
  lStGen              := FStArgumentTof ;
  ABoSansFiltreDefaut := ReadTokenSt(lStGen) <> '' ;
  lStAux              := ReadTokenSt(lStGen) ;
  lStAux              := ReadTokenSt(lStGen) ;
  lStAux              := ReadTokenSt(lStGen) ;
  ABoSansFiltreDefaut := ABoSansFiltreDefaut or ( lStAux <> '' ) ;

{$IFDEF BUREAU}
  AFiltreDisabled := True;
{$ENDIF}

  inherited; // InitControl; Chargement du OBJFiltre et init du filtre 'DEFAUT'

  MakeZoomOLE(ECRAN.Handle); // Appel depuis Excel

  LibellesTableLibre(TabTablesLibres, 'TE_TABLE', 'E_TABLE', 'E');

  // Chargement des paramètres du OnArgument
  AnalyseParametre(FStArgumentTof);

  InitContexte;

  InitAllPopUp(False);

  // Mise à blanc des composants Z_
  AfficheInfoSup;

  InitAutoSearch;

  // On laisse charger les dates comptables avant de cocher justif de solde
  if FBoOuvreEnJustif then
    CBJustifSolde.Checked := True;

  // GCO - 30/12/2004 - FQ 15037
  PositionneEtabUser(E_Etablissement);

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/09/2004
Modifié le ... : 20/09/2004
Description .. :
Suite ........ : Prise de Focus de la zone E_GENERAL à l'ouverture
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.AfterShow;
begin
{$IFDEF BUREAU}
  // GCO - 13/09/2007 -
  // PB avec M Desgouttes, 31/12/2099 -> devenait = E_datecomptable.text
  OnChangeComboE_Exercice(nil);
{$ENDIF}

{$IFDEF COMPTA}
  // Pas de filtre par défaut + FQ 20749 + FQ 21388 + FQ 21731 + FQ 21822
  if (FFiltres.Value = '') and (not CBJustifSolde.Checked) then
  begin
    E_Exercice.ItemIndex := -1;
    OnChangeComboE_Exercice(nil);
  end;
{$ENDIF}

  inherited; // BCherche.Click fait par l'ancêtre

  // Mise du focus à la zone général si pas de compte passé en paramètre
  if (Trim(E_General.Text) = '') and (Trim(E_Auxiliaire.Text) = '') then
  begin
    if E_General.CanFocus then
      E_General.SetFocus;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/10/2002
Modifié le ... : 24/11/2006
Description .. : Active les onglets Lettrage et Pointage en Fonction de
Suite ........ : FOkLettrable et FOkPointage
Suite ........ : - LG - FB 19058 - 24/11/2006 - ouverture des fct de
Suite ........ : pointage pour le pointage sur journal
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.ActivationOnglet;
begin
  PageControl.Pages[3].TabVisible := FOkLettrable or FOkOngletLettrage;

  if not FOkLettrable and not FOkOngletLettrage then
    VideControlOnglet(PageControl.Pages[3]);

  if not FOkPointageJal then
   PageControl.Pages[4].TabVisible := FOkPointable
    else
     PageControl.Pages[4].TabVisible := false ;
  if not FOkPointable then
  begin
    VideControlOnglet(PageControl.Pages[4]);
  end;
  ActivationOption;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/10/2003
Modifié le ... : 10/10/2003
Description .. : Grise, Dégrise et Coche ou Décoche le justif de solde, les ANO
Suite ........ : détaillés et la centralisation des journaux
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.ActivationOption;
var lExoDate : TExoDate;
begin
  if (FTobGene = nil) or (FTobTiers = nil) then Exit;

  // Inutile de faire l'activation des options si les comptes n'ont pas été
  // récupérés dans leur tob respective
  if (E_General.Text <> FTobGene.GetString('G_GENERAL')) or
     (E_Auxiliaire.Text <> FTobTiers.GetString('T_AUXILIAIRE')) then Exit;

  // Autorisation de faire un Justificatif de Solde
  CBJustifSolde.Enabled := FOkJustificatifSolde;

  if FOkLettrable then
  begin
    if (E_Exercice.ItemIndex = 0) or (CBJustifSolde.Checked) then
    begin
      CBDetailAno.Enabled := False;
    end
    else
    begin
      //GCO - 11/06/2004 - Test si réelle présence d'écritures "OAN"
      if CQuelExercice( StrToDate(E_DateComptable.Text), lExoDate ) then
      begin
        if E_DateComptable.Text = DateToStr(lExoDate.Deb) then
        begin
          CBDetailAno.Enabled := True;
        end
        else CBDetailAno.Enabled := False;
      end
      else // Exercice Non Trouvé, ne devrait pas arriver
        CBDetailAno.Enabled := False;
    end;

    CBCentralisation.Enabled := (FOkCentralisation) and (not CBJustifSolde.Checked) and (not CBDetailAno.Checked);

    if CBDetailAno.Checked then
      CBJustifSolde.Enabled := False;
  end
  else
  begin
    CBDetailAno.Enabled      := False;
    CBCentralisation.Enabled := FOkCentralisation;
  end;

  if (CBJustifSolde.Enabled) then
  begin
    if (CBJustifSolde.Checked) and (not FBoStateCBJustifSolde) then
      FBoStateCBJustifSolde := True;

    if (not CBJustifSolde.Checked) and (FBoStateCBJustifSolde) then
      CBJustifSolde.Checked := True;
  end
  else
    CBJustifSolde.Checked := False;

  if CBDetailAno.Enabled then
  begin
    if (CBDetailAno.Checked) and (not FBoStateCBDetailANO) then
      FBoStateCBDetailANO := True;

    if (not CBDetailAno.Checked) and (FBoStateCBDetailANO) then
      CBDetailAno.Checked := True;
  end
  else
    CBDetailAno.Checked := False;

  if not CBCentralisation.Enabled then
    CBCentralisation.Checked := False;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/09/2002
Modifié le ... : 24/11/2006
Description .. : Initialisation des colonnes en fonction des critères
Suite ........ : InitColFliste avec uniquement appelée depuis la méthode
Suite ........ : ActivationOnglet
Suite ........ : - LG - FB 19058 - 24/11/2006 - ouverture des fct de
Suite ........ : pointage pour le pointage sur journal - on n'affiche pas
Suite ........ : l'onglet de poinatge pour le pointage sur journal
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.InitColFListe;
var lVisuGene, lVisuAux: Boolean;
begin
  IndiceColFListe;

  lVisuGene := False;
  lVisuAux  := False;

  if FTobGene.GetString('G_GENERAL') <> '' then
  begin
    lVisuGene := False;
    if FTobGene.GetString('G_COLLECTIF') = 'X' then
    begin
      if FTobTiers.GetString('T_AUXILIAIRE') <> '' then
      begin
        lVisuGene := False;
        lVisuAux := False;
      end
      else
        lVisuAux := True;
    end
    else
      lVisuAux := False;
  end
  else
  begin
    if FTobTiers.GetString('T_AUXILIAIRE') <> '' then
      lVisuGene := True;
  end;

  FListe.ColWidths[FColGeneral]     := IIF(lVisuGene, FListe.ColWidths[FColGeneral], -1);
  FListe.ColWidths[FColAuxiliaire]  := IIF(lVisuAux, FListe.ColWidths[FColAuxiliaire], -1);
  FListe.ColWidths[FColLettrage]    := IIF(PageControl.Pages[3].TabVisible, FListe.ColWidths[FColLettrage], -1);
  if FOkPointageJal then
    FListe.ColWidths[FColRefPointage] := -1
  else
    FListe.ColWidths[FColRefPointage] := IIF(PageControl.Pages[4].TabVisible, FListe.ColWidths[FColRefPointage], -1);

  // GCO - 19/09/2007 - FQ 21470
  FListe.ColWidths[FColSoldePro]    := IIF(ComboSoldePro.Value = '0', -1, FListe.ColWidths[FColSoldePro]);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/01/2006
Modifié le ... : 26/06/2007
Description .. : LG - 26/06/2007 - supprimer suite au deplacement du 
Suite ........ : Inherited ds le OnLoad.
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.InitContexte;
begin
{$IFDEF BUREAU}
  E_General.Enabled    := False; // GCO - 03/01/2006 - FQ 17256
  E_Auxiliaire.Enabled := False; // GCO - 03/01/2006 - FQ 17256
  E_Exercice.Enabled   := False; // FQ 17230 - E_Exercice grisé

  BGUp.Enabled   := False;
  BGDown.Enabled := False;
  BTUp.Enabled   := False;
  BTDown.Enabled := False;

  FBoBlocageTotal := True;
  FOkIccGere      := False;
  FOkCreGere      := False;
  FOkPointageJal  := False;
  FOkModImmo      := False;
  FOkGereQte      := False;
  FOkAnaCroisaxe  := False;
  FOkRevision     := False;

{$ELSE}
  // Séria ICC et CRE
  if CtxPcl in V_Pgi.PgiContexte then
  begin
    FOkIccGere := (VH^.OkModIcc) and (ExisteSQl('SELECT ICG_GENERAL FROM ' + FStNomBase + 'ICCGENERAUX ORDER BY ICG_GENERAL'));
    FOkCreGere := VH^.OkModCre;
  end
  else
  begin
    FOkIccGere := (ExisteSQl('SELECT ICG_GENERAL FROM ' + FStNomBase + 'ICCGENERAUX ORDER BY ICG_GENERAL'));
    FOkCreGere := True;
  end;

  if FOkIccGere then
  begin
    FTobListeICC := Tob.Create('ICC', nil, -1);
    FTobListeIcc.LoadDetailFromSQL('SELECT ICG_GENERAL FROM ' + FStNomBase + 'ICCGENERAUX ORDER BY ICG_GENERAL');
  end;

  // Séria IMMO
  FOkModImmo := VH^.OkModImmo;

  FOkPointageJal := VH^.PointageJal;

  // Gestion des Quantités
  FOkGereQte := GetParamSocSecur('SO_CPPCLSAISIEQTE', False);

  FOkAnaCroisaxe := VH^.AnaCroisaxe;

  FStLiasseDossier := GetParamSocSecur('SO_CPCONTROLELIASSE', '');
{$ENDIF}

  FOkRevision := VH^.Revision.Plan <> '';

//  FOkCreateModif  := False;  LG - 26/06/2007 - supprimer suite au deplacement du Inherited ds le OnLoad.

  FDroitEcritures := ExJaiLeDroitConcept(TConcept(ccSaisEcritures), False);

  E_General.MaxLength    := GetInfoCpta(fbGene).Lg;
  E_Auxiliaire.MaxLength := GetInfoCpta(fbAux).Lg;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/06/2004
Modifié le ... :   /  /
Description .. : - LG - 30/06/2004 - FB 12989 12960 - on utilise la même
Suite ........ : requete pour la consultation et le lettrage
Mots clefs ... :
*****************************************************************}
function TOF_CONSECR.FaireStWhere( vBoForceDetailAno : boolean = false ) : string ;
var
  lStWhere           : string ;
  lStWhereDetail     : string ;
  lStLibEtatLettrage : string ;
  lStSqlEtatLettrage : string ;
begin
  lStLibEtatLettrage := '';
  lStSqlEtatLettrage := '';
  lStWhere := '';
  if (E_General.Text <> '') or (E_Auxiliaire.Text <> '') then
  begin
    lStWhere := ' WHERE (';
    if CBJustifSolde.Checked then
    begin
      lStWhere := lStWhere + '(E_DATEPAQUETMIN <= "' + UsDateTime(StrToDate(E_DateComptable_.Text)) + '" ' +
                  'AND ((E_ETATLETTRAGE <> "TL") OR (E_ETATLETTRAGE = "TL" ' +
                  'AND E_DATEPAQUETMAX > "' + UsDateTime(StrToDate(E_DateComptable_.Text)) + '")) ' +
                  'AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")) ' +
                  'AND E_QUALIFPIECE = "N"';

      lStWhere := lStWhere + ' AND ' + RecupWhereAGL;
    end
    else
    begin
      lStWhereDetail := RecupWhereDetailAno(vBoForceDetailAno);
      if lStWhereDetail <> '' then
      begin
        lStWhere := lStWhere + '(' + lStWhereDetail + ') OR ' + RecupWhereEcriture(vBoForceDetailAno);
      end
      else
      begin
        lStWhere := lStWhere + RecupWhereEcriture(vBoForceDetailAno);
      end;
    end;
    // Ajout des critères communs à la requête
    lStWhere := lStwhere + ') AND ' + RecupCriteresCommuns;

    // On envoie le texte SQL de la requête à la fiche ANCETRE pour qu'elle fasse le OPENSQL
    Result := lStWhere;

  end
  else
    Result := '';
end;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/08/2002
Modifié le ... : 22/11/2007
Description .. : Génère la requete SQL en fonction de des crtières
Suite ........ : Charge ATobFListe avec le résultat de la requête
Suite ........ : Modifie l' affichage du solde progressif ( DB ou CR )
Suite ........ : GCO - 10/02/2003 --> Optimisation requete pour justif de
Suite ........ : solde
Suite ........ : GCO - 28/02/2004 --> Ajout des critères communs
Suite ........ : - LG - 30/06/2004 - FB 12989 12960 - on utilise la meme
Suite ........ : requet pour la consultation et le lettrage. Creation d'un fct
Suite ........ : pour retourner la requet a partir des critères de selection
suite ........ : GCO - 22/11/2007 - GED avec EC_CLEECR
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.RemplitATobFListe;
var lStWhere : string ;
begin
  // Traduction des champs dans la LISTE PARAMETRABLE
  AStSqlTobFListe := '' ;

  if FBoFaireRequete then
  begin
    lStWhere  :=  FaireStWhere ;
    if lStWhere <> '' then
    begin
      AStSqlTobFListe := 'SELECT ' + CSqlTextFromList(FStListeChamps) + ' ' +
                          cSelect;

      if (CtxPcl in V_Pgi.PGIContexte) then
        AStSqlTobFListe := AStSqlTobFListe + cSelectGED;

      AStSqlTobFListe := AStSqlTobFListe + ' FROM ' + FStNomBase + 'ECRITURE ';

      // Ajout du LEFT JOIN sur ECRCOMPL
      if (CtxPcl in V_Pgi.PGIContexte) then
      begin
        AStSqlTobFListe := AStSqlTobFListe + 'LEFT JOIN ECRCOMPL ON ' +
                           'EC_JOURNAL = E_JOURNAL AND ' +
                           'EC_EXERCICE = E_EXERCICE AND ' +
                           'EC_DATECOMPTABLE = E_DATECOMPTABLE AND ' +
                           'EC_NUMEROPIECE = E_NUMEROPIECE AND ' +
                           'EC_QUALIFPIECE = E_QUALIFPIECE AND ' +
                           'EC_NUMLIGNE = E_NUMLIGNE ';

        AStSqlTobFListe := AStSqlTobFListe + 'LEFT JOIN ECRCOMPL E2 ON ' +
                           'E2.EC_JOURNAL = E_JOURNAL AND ' +
                           'E2.EC_EXERCICE = E_EXERCICE AND ' +
                           'E2.EC_DATECOMPTABLE = E_DATECOMPTABLE AND ' +
                           'E2.EC_NUMEROPIECE = E_NUMEROPIECE AND ' +
                           'E2.EC_QUALIFPIECE = E_QUALIFPIECE AND ' +
                           'E2.EC_CLEECR = CAST(E_NUMGROUPEECR AS CHAR(6)) AND ' +
                           'E2.EC_DOCGUID <> "" ';
      end;

      AStSqlTobFListe := AStSqlTobFListe + lStWhere +
        ' ORDER BY E_DATECOMPTABLE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE ';
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/01/2005
Modifié le ... : 01/12/2005
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CONSECR.AjouteATobFListe(vTob : Tob) : Boolean;
begin
  Result := True;
  vTob.AddChampSupValeur('SOLDEE', 0, False);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/12/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.CalculPourAffichage( vTob : Tob );
var lSoldePro : Double;
begin
  // Calcul du solde progressif
  if vTob.GetIndex = 0 then
    lSoldePro := 0
  else
    lSoldePro := ATobFListe.Detail[vTob.GetIndex-1].GetDouble('SOLDEE');

  if lSoldePro >= 0 then
    lSoldePro := lSoldePro + vTob.GetDouble('E_DEBIT') - vTob.GetDouble('E_CREDIT')
  else
    lSoldePro := lSoldePro - vTob.GetDouble('E_CREDIT') + vTob.GetDouble('E_DEBIT');

  vTob.SetDouble('SOLDEE', lSoldePro);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/10/2002
Modifié le ... :   /  /
Description .. : Charge la THGRID FLISTE avec le contenu de ATobFListe
Suite ........ : Calcul le solde progressif
Suite ........ : GCO - 21/09/2004 - FQ 14638
Suite ........ : Le traitement de la Centralisation Mensuelle était fait après
Suite ........ : le RefreshFListe d'UtofViergeMul qui fait le Tri de ATOBFliste
Suite ........ : Centralisation Mensuelle non prévu en CWAS, mais probablement
Suite ........ : faisable par la suite..... ( une idée ^^ )
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.RefreshFListe;
{$IFDEF EAGLCLIENT}
var lQuery : TQuery;
{$ENDIF}
begin
  // On force le tri sur un champ supplémentaire afin que le ATobFListe.Detail.Sort
  // renvoi toujours ATobFListe triée dans le même ordre
  AStTriTobFListe := AStTriTobFListe + ' ;CLE_NUMLIGNE';

{$IFDEF EAGLCLIENT}
  // Pas de Centralisation Mensuelle pour l'instant en CWAS
{$ELSE}
  // GCO - 21/09/2004 - FQ 14638
  // Centralisation Mensuelle
  if (ATobFListe <> nil) and (ATobFliste.Detail.Count > 0) and (CBCentralisation.Checked) then
    TraiteCentralisation;
{$ENDIF}

  //if (ATobFListe <> nil) and (ATobFliste.Detail.Count > 0) then
  //begin
  //  ATobFListe.PutValueAllFille('SOLDEE', 0);
  //end;

  inherited; // Appel du inherited de la fiche Ancêtre pour trier ATobFListe et
             // l'afficher dans la grille

  if not vBoFetch then
  begin
    LeSolde.Text := '0,00';
    {$IFDEF EAGLCLIENT}
    if not vBoFetch then
    begin
      if AStSqlWhereTobFListe <> '' then
      begin
        lQuery := OpenSql('SELECT SUM(E_DEBIT)-SUM(E_CREDIT) TOTAL FROM ' + FStNomBase + 'ECRITURE ' + AStSqlWhereTobFListe , True);
        LeSolde.Text := AfficheDBCR( lQuery.FindField('TOTAL').AsFloat );
        Ferme( lQuery );
      end;
    end;
    {$ELSE}
    // Mise à jour des informations du compte affiché
    if (ATobFListe <> nil) and (ATobFListe.Detail.Count <> 0) then
      LeSolde.Text := AfficheDBCR( ATobFListe.Detail[ATobFListe.Detail.Count-1].GetDouble('SOLDEE'));
    {$ENDIF}

    // Affichage du solde Bancaire lorsqu'une date de pointage est saisie
    if E_DatePointage.Text <> DateToStr(iDate1900) then
      ESoldeBancaire.Text := AfficheDBCR(RecupSoldePointage(E_General.Text, StrToDate(E_Datepointage.Text), ''));

    // Mise à Jour du Caption de l' écran
    MiseAJourCaptionEcran;

    //
    ActivationOption;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/08/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOF_CONSECR.RecupWhereAGL : string ;
begin
  Result := RecupWhereCritere(PageControl);

  // Traduction des exercices en relatif
  Result := CMajRequeteExercice(E_EXERCICE.Value, Result);

  // Supression du WHERE ajouter par l'AGL
  Result := Copy(Result, 7, Length(Result));
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/09/2002
Modifié le ... : 03/10/2005
Description .. : Récupère les écritures concernées par les critères
Suite ........ : ((1) or (2)) and (3)
Suite ........ : étape (1) de la requête principale
Suite ........ : - LG - 03/10/2005 - FB 16744 - la consultation depuis la 
Suite ........ : saisei bordereau ne fct plus. Gestion manule du qualifpiece
Mots clefs ... : 
*****************************************************************}

function TOF_CONSECR.RecupWhereEcriture ( vBoForceDetailAno : boolean = false ) : string;
var
  lSt  : string ;
  lSt1 : string ;
  lSt2 : string ;
  lDateComptable : TDateTime;
  lstOldValue    : string ;
begin

  lstOldValue       := EQualifPiece.Text ;
  EQualifPiece.Text := '' ;
  lSt               := RecupWhereAGL;
  EQualifPiece.Text := lstOldValue ;


  // Traitmenement du THMultiValCOmboBox du E_QUALIFPIECE
  // Attention : vVideSiTous doit toujours être à false avec E_QUALIFPIECE
  TraductionTHMultiValComboBox( EQualifPiece, lSt1, lSt2, 'E_QUALIFPIECE', False);
  if FBoLettrageEnSaisie then
    lSt := lSt + ' AND (' + lSt1 + ' OR E_QUALIFPIECE = "L")'
  else
    lSt := lSt + ' AND ' +  lSt1;

  lDateComptable := StrToDate(E_DateComptable.Text);
  if (GetSuivant.Deb > 0) and (lDateComptable > GetSuivant.Deb) then
    lSt := lSt + ' AND E_ECRANOUVEAU="N"'
  else
  begin
    if  CBDetailAno.Checked  or vBoForceDetailAno then
      //lSt := lSt + ' AND E_ECRANOUVEAU="N"';
      lSt := lSt + ' AND (E_ECRANOUVEAU="N" OR (E_ECRANOUVEAU="H" AND E_DATECOMPTABLE="' + UsDateTime(lDateComptable) + '"))'
    else
      lSt := lSt + ' AND (E_ECRANOUVEAU="N" OR ((E_ECRANOUVEAU="H" OR E_ECRANOUVEAU="OAN") AND E_DATECOMPTABLE="' + UsDateTime(lDateComptable) + '"))';
  end;

  if E_DatePointage.Text <> stDate1900 then
  begin
    lSt := lSt + ' AND (E_ECRANOUVEAU <> "OAN") ' + ' AND ( ((E_REFPOINTAGE = "") AND (E_DATECOMPTABLE <= "' +
           UsDateTime(StrToDate(E_DatePointage.Text)) + '")) OR ' + '( (E_REFPOINTAGE <> "") AND ' +
           '(E_DATECOMPTABLE <= "' + UsDateTime(StrToDate(E_DatePointage.Text)) + '") AND ' +
           '(E_DATEPOINTAGE  > "' + UsDateTime(StrToDate(E_DatePointage.Text)) + '"))' + ')';
  end;

  Result := lSt;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/11/2003
Modifié le ... : 16/12/2005
Description .. :  Récupère le détail des ANO des écritures
Suite ........ : ((1) or (2)) and (3)
Suite ........ : étape (2) de la requête principale
Mots clefs ... :
*****************************************************************}
function TOF_CONSECR.RecupWhereDetailANo( vBoForceDetailAno : boolean = false ) : string;
var //lStSql : string;
    //lSt1, lSt2 : string;
    lExoDate : TExoDate;
begin
  Result := '';

  if not vBoForceDetailAno then
  begin
   if  not ((CBDetailAno.Enabled) and ( CBdetailAno.Checked )) then Exit;
  end;

  // Doit on ajouter le détail des A-Nouveaux dans la requête ?
  // Test de l'exercice sélectionné
  if E_Exercice.ItemIndex = 0 then Exit;

  // Test si la date Début correspond à la date de début de l'exercice sélectionné
  lExoDate.Code := CRelatifVersExercice( E_Exercice.Value );
  RempliExoDate(lExoDate);
  if StrToDate(E_DateComptable.Text) <> lExoDate.Deb then Exit;

  // Date Comptable inférieure à la DATE de DEBUT
  if IsValidDate(E_DateComptable.Text) then
    Result := 'E_DATECOMPTABLE < "' + UsDateTime(StrToDate(E_DateComptable.Text)) + '"';

  //
  Result := Result + ' AND E_ECRANOUVEAU <> "OAN"';

  Result := Result + ' AND E_QUALIFPIECE = "N"';

  // Notion de LETTRAGE des écritures
  Result := Result + ' AND (E_ETATLETTRAGE <> "TL" OR ' +
            '(E_ETATLETTRAGE="TL" AND E_DATEPAQUETMAX >= "' + UsDateTime(StrToDate(E_DateComptable.Text)) + '"))';
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/11/2003
Modifié le ... : 16/12/2005
Description .. : Récupère les critères communs aux écritures pour les écritures
Suite ........ : ((1) or (2)) and (3)
Suite ........ : étape (3) de la requête principale
Mots clefs ... :
*****************************************************************}
function TOF_CONSECR.RecupCriteresCommuns : string;
var lSt1, lSt2, lStTemp : string;
    lExoDate : TExoDate;
    lStWhereSqlEtatLettrage : string;
    lStLibEtatLettrage : string;
begin
  // GCO - 27/09/2006 - NB pour supprimer les ecr d'écart de change en lettrage
  Result := ' (E_DEBIT <> 0 OR E_CREDIT <> 0)';

  Result := Result + ' AND E_CREERPAR <> "DET"';

  // Général
  if E_General.Text <> '' then
    Result := Result  + ' AND E_GENERAL = "' + E_General.Text + '"';

  // Auxiliaire
  if E_Auxiliaire.Text <> '' then
    Result := Result + ' AND E_AUXILIAIRE = "' + E_Auxiliaire.Text + '"';

  // Etablissement
  TraductionTHMultiValComboBox( E_Etablissement, lSt1, lSt2, 'E_ETABLISSEMENT');
  if lSt1 <> '' then
    Result := Result + ' AND ' + lSt1;

  // Devise
  if E_Devise.ItemIndex <> 0 then
    Result := Result + ' AND E_DEVISE = "' + E_DEVISE.Value + '"';

  // GCO - 02/12/2003 FB 13073
  // GCO - 10/11/2004 FB 14965
  if GetExoV8.Code <> '' then
  begin
    if E_Exercice.ItemIndex <> 0 then
    begin
      lExoDate.Code := CRelatifVersExercice( E_Exercice.Value );
      RempliExoDate(lExoDate);
      if lExoDate.Deb >= GetExoV8.Deb then
        Result := Result + ' AND E_DATECOMPTABLE >= "' + UsDateTime(GetExoV8.Deb) + '"';
    end;
  end;

  // Recup des criteres pour E_EtatLettrage
  if (not E_EtatLettrage.Tous) and (Trim(E_EtatLettrage.Text) <> '') then
  begin
    TraductionTHMultiValComboBox(E_EtatLettrage, lStWhereSqlEtatLettrage, lStLibEtatLettrage, 'E_ETATLETTRAGE') ;
    // GCO - 14/09/2005 - FQ 15749
    if lStWhereSqlEtatLettrage <> '' then
      Result := Result + ' AND ' + lStWhereSqlEtatLettrage;
  end;

  if Trim(E_Lettrage.Text) <> '' then
   Result := Result + ' AND E_LETTRAGE LIKE "%' + E_Lettrage.Text + '%"';

  if (ComboMontant.Value <> '') and
     (ComboOpMontant.Value <> '') and
     (EMontant.Text <> '') then
  begin
    lStTemp := TraduitOperateur(ComboOpMontant.Value, EMontant.Text, False);
    if lStTemp <> '' then
    begin
      if (ComboMontant.Value = 'E_DEBIT') then
        Result := Result + ' AND ' + ComboMontant.Value + ' ' + lStTemp + ' AND E_CREDIT = 0'
      else
        if (ComboMontant.Value = 'E_CREDIT') then
          Result := Result + ' AND ' + ComboMontant.Value + ' ' + lStTemp + ' AND E_DEBIT = 0'
        else
          Result := Result + ' AND ((E_DEBIT ' + lStTemp + ' AND E_CREDIT = 0) OR (E_CREDIT ' + lStTemp +' AND E_DEBIT = 0))';
    end;
  end;

  if (ELibelle.Text <> '') or (ComboOpLibelle.Value = 'V') or (ComboOpLibelle.Value = 'VV') then
  begin
    lStTemp := TraduitOperateur(ComboOpLibelle.Value, ELibelle.Text, True);
    if lStTemp <> '' then
      Result := Result + ' AND E_LIBELLE ' + lStTemp;
  end;

  if (ERefInterne.Text <> '') or (ComboOpRefInterne.Value = 'V') or (ComboOpRefInterne.Value = 'VV') then
  begin
    lStTemp := TraduitOperateur(ComboOpRefInterne.Value, ERefinterne.Text, True);
    if lStTemp <> '' then
      Result := Result + ' AND E_REFINTERNE ' + lStTemp;
  end;

  // GCO - 03/12/2007
  if (CtxPcl in V_Pgi.PGIContexte) and (CBDocumentGED.State <> cbGrayed) then
  begin
    if CBDocumentGED.Checked then
      Result := Result + ' AND E2.EC_DOCGUID <> ""'
    else
      Result := Result + ' AND (E2.EC_DOCGUID = "" OR E2.EC_DOCGUID IS NULL)';
  end;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/02/2003
Modifié le ... : 11/02/2003
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  inherited;

  case Key of

    VK_DOWN, VK_F4 :
              begin
                if E_GENERAL.Focused then
                begin
                  Key := 0;
                  GeneSuivant(True, ssCtrl in Shift);
                end;

                if E_AUXILIAIRE.Focused then
                begin
                  Key := 0;
                  AuxSuivant(True, ssCtrl in Shift);
                end;
              end;

    VK_UP, VK_F3 :
              begin
                if E_GENERAL.Focused then
                begin
                  Key := 0;
                  GeneSuivant(False, ssCtrl in Shift);
                end;

                if E_AUXILIAIRE.Focused then
                begin
                  Key := 0;
                  AuxSuivant(False, ssCtrl in Shift);
                end;
              end;

    // GCO - 18/09/2007 - FQ 21453
    VK_F5 : begin
              if FListe.Focused and (Shift = [ssShift]) then
                OnClickPasseJustifSolde(nil);
            end;

    // GCO - 05/05/2006 - FQ 16190
    46 : begin // Maj + Suppr -> Justificatif de Soldes + F9
           if (Shift = [ssShift]) then
             OnClickPasseJustifSolde(nil);
         end;

    Ord('C') : begin
                 {$IFDEF COMPTA}
                 if (ABoMultiSelected) and (ssCtrl in Shift) then
                 begin
                   OnClickEcrVersClipboard(nil); // Ctrl + C -> Informations complémentaires
                   Key := 0 ; // FQ 21273 BVE 28.08.07
                 end;

                 if (not ABoMultiSelected) and (ssAlt in Shift) then
                   OnClickInfosComp(nil); // Alt + C -> Informations complémentaires
                 {$ENDIF}
               end;

    Ord('E') : begin // Alt + E
                 if (not ABoMultiSelected) and (ssAlt in Shift) then ;
               end;

    Ord('G') : begin // CTRL + G -> Accès zone EGENERAL
                 if ssCtrl in Shift then
                 begin
                   PageControl.ActivePage := PageControl.Pages[0];
                   if E_General.CanFocus then
                     SetFocusControl('EGENERAL');
                 end;
               end;

    Ord('H') : begin // ALT + H -> Table de Chancellerie
                 if (not ABoMultiSelected) and (ssAlt in Shift) then
                   OnClickTableChancellerie(nil);
               end;

    Ord('M') : begin // CTRL + M -> Onglet Mouvement
                 if ssCtrl in Shift then
                 begin
                   PageControl.ActivePage := PageControl.Pages[5];
                   if EMontant.CanFocus then EMontant.SetFocus;
                 end;
               end;

    Ord('O') : begin // Alt + O -> Commentaire Millésimé
                 if (not ABoMultiselected) and (ssAlt in Shift) then
                 {$IFDEF COMPTA}
                 {$IFNDEF CCMP}
                  OnClickMemoMillesime(nil);
                 {$ENDIF}
                 {$ENDIF}
               end;

    Ord('Q') : begin // CTRL + Q -> Onglet Infos Complémentaires Suite...
                 if (not ABoMultiSelected) and (ssAlt in Shift) then
                 begin
                   PageInfoSup.ActivePage := PageInfoSup.Pages[1];
                   ComboInfoSup.Value := '1';
                 end;
               end;

    Ord('R') : begin // GCO - 20/07/2007 - FQ 20919
                 if (not ABoMultiSelected) and (ssAlt in Shift) then
                  OnClickCompteRib( nil ); // ALT + R -> Accès au RIB
               end;

    Ord('T') : begin // Ctrl + T
                 if ssCtrl in Shift then
                 begin
                   PageControl.ActivePage := PageControl.Pages[0];
                   if E_Auxiliaire.CanFocus then
                     SetFocusControl('EAUXILIAIRE');
                 end;
               end;

    Ord('V') : begin
                 if ssAlt in Shift then // ALT + V -> Viser le compte
                 begin
                 {$IFDEF COMPTA}
                 {$IFNDEF CCMP}
                   // GCO - 10/05/2007 - FQ 20273
                   if (not ABoMultiSelected) and
                      (FExoDate.Code = VH^.Encours.Code) and
                      //(FOkRevision) and
                      (E_General.Text <> '') and
                      AutoriseSuppresionVisaRevision(E_General.Text) then
                       OnClickVisaCompte(nil);
                 {$ENDIF}
                 {$ENDIF}
                 end;
               end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/05/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnKeyDownFListe(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/11/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnPostDrawCellFListe(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var TheRect  : TRect;
    T1,T2,T3 : TPoint;
    lTob     : Tob;
begin
  inherited;

//GP le 18/09/2008 : report debug PCL
  If ARow=0 THen Exit ;
  
  // Dessin du triangle dans la première colonne si DOC dans la GED
  if (CtxPcl in V_Pgi.PGIContexte) and (ACol = 1) then
  begin
    lTob := Tob(FListe.Objects[0, ARow]);
    if (lTob <> nil) and (lTob.GetString('CLE_DOCGUID') <> '') then
    begin
      TheRect:=FListe.CellRect(ACol,ARow);
      Canvas.Brush.Color := clGreen;
      Canvas.Brush.Style := bsSolid;
      Canvas.Pen.Color   := clGreen;
      Canvas.Pen.Mode    := pmCopy;
      Canvas.Pen.Width   := 1;
      T1.X := TheRect.Right - 6;
      T1.Y := TheRect.Bottom - 1;
      T2.X := TheRect.Right - 1;
      T2.Y := TheRect.Bottom - 1;
      T3.X := TheRect.Right - 1;
      T3.Y := TheRect.Bottom - 6;
      Canvas.Polygon([T1,T2,T3]);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/09/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.GetCellCanvasFListe(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  if ARow = 0 then Exit;

  if (ACol = FColSoldePro) and (FListe.Row <> ARow) then
  begin
    if Trim(E_Auxiliaire.Text) <> '' then
    begin
      if FTobTiers.GetString('T_NATUREAUXI') = 'CLI' then
        Canvas.Font.Color := IIF(Pos('C', FListe.Cells[ACol, ARow]) > 0, ClRed, ClGreen)
      else
        if FTobTiers.GetString('T_NATUREAUXI') = 'FOU' then
          Canvas.Font.Color := IIF(Pos('C', FListe.Cells[ACol, ARow]) > 0, ClGreen, ClRed);
    end
    else
    begin
      if Trim(E_General.Text) <> '' then
      begin
        if FTobGene.GetString('G_SENS') = 'D' then
          Canvas.Font.Color := IIF(Pos('C', FListe.Cells[ACol, ARow]) > 0, ClRed, ClGreen)
        else
          if FTobGene.GetString('G_SENS') = 'C' then
            Canvas.Font.Color := IIF(Pos('C', FListe.Cells[ACol, ARow]) > 0, ClGreen, ClRed);
      end;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnChangeComboE_Exercice(Sender: TObject);
begin
  if E_Exercice.ItemIndex = 0 then
  begin
    CbDetailAno.Checked := False;
    if GetExoV8.Code <> '' then
      E_DateComptable.Text := DateToStr(GetExoV8.Deb)
    else
      // Date de début du premier Exercice
      E_DateComptable.Text := DateToStr(GetExercices(1).Deb);
    E_DateComptable_.Text  := DateToStr(iDate2099-1);
  end
  else
  begin
    if E_Exercice.ItemIndex = -1 then
    begin
      if CtxPCl in V_Pgi.PgiContexte then
      begin
        if GetExoRef.Code <> '' then
          E_Exercice.Value := CExerciceVersRelatif ( GetExoRef.Code )
        else
          E_Exercice.Value := CExerciceVersRelatif ( GetEncours.Code );
      end
      else
        E_Exercice.Value := CExerciceVersRelatif( GetEntree.Code );
    end;
    CExoRelatifToDates(E_Exercice.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_);
  end;

  // GCO - 13/07/2007 - FQ 20749
  // Imcompréhensible pourquoi cela corrige, si réaffecte à la même valeur
  E_DateComptable.Text  := E_DateComptable.Text;
  E_DateComptable_.Text := E_DateComptable_.Text;

  ActivationOption; // Pour la gestion des A-Nouveaux détaillés et Justif de solde
  VerifieLesDates;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/01/2003
Modifié le ... : 23/02/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnExitE_DateComptable(Sender: TObject);
begin
  VerifieLesDates;
  if FBoOkControleDate then
    ActivationOption;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_CONSECR.OnChangeComboInfoSup(Sender: TObject);
begin
  if ComboInfoSup.Value = '1' then
  begin
    TPageControl(GetControl('PCUMUL')).Visible := False;
    TPageControl(GetControl('PAGEINFOSUP')).Visible := True;
    // On repasse le PCUMUL a visible True pour qu'il soit au dessus du PAGEINFOSUP
    // Car sinon Problème de placement si on est pas en inside
    TPageControl(GetControl('PCUMUL')).Visible := True;
    AfficheInfoSup;
  end
  else
    TPageControl(GetControl('PAGEINFOSUP')).Visible := False;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CONSECR.OnChangeComboAcces(Sender: TObject);
begin
  FOkCreateModif := (ComboAcces.Value = '1');
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_CONSECR.OnExitE_RefPointage(Sender: TObject);
begin
  E_Refpointage_.Text := E_RefPointage.Text;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_CONSECR.OnElipsisClickE_RefPointage(Sender: TObject);
begin
  LookUpList(THEdit(Sender), 'Référence de pointage', 'EEXBQ', 'EE_REFPOINTAGE',
    'EE_DATEPOINTAGE', 'EE_GENERAL="' + E_General.Text + '"',
    'EE_DATEPOINTAGE DESC', True, 0);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/10/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOF_CONSECR.OnElipsisClickE_DatePointage(Sender: TObject);
begin
  LookUpList(THEdit(Sender), 'Date de pointage', 'EEXBQ', 'EE_DATEPOINTAGE',
    'EE_REFPOINTAGE', 'EE_GENERAL="' + E_General.Text + '"',
    'EE_DATEPOINTAGE DESC', True, 0);
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_CONSECR.OnExitE_DatePointage(Sender: TObject);
begin
  if IsValidDate(E_DatePointage.Text) then
  begin
    // Particularités de Pointage dans les critères
    if E_DatePointage.Text <> DateToStr(iDate1900) then
    begin
      E_Exercice.ItemIndex := 0;
      E_Exercice.OnChange(nil);
      E_DateComptable_.Text := E_DatePointage.Text;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////

function TOF_CONSECR.RecupSoldePointage(vGeneral: string; vDatePointage:
  TDateTime; vRefPointage: string): Double;
var
  lQuery: TQuery;
begin
  Result := 0;
  lQuery := nil;
  try
    lQuery :=
      OpenSql('SELECT (EE_NewSoldeDebEuro - EE_NewSoldeCreEuro) AS SOLDE FROM ' +
              FStNomBase + 'EEXBQ WHERE (EE_GENERAL = "' + vGeneral + '") AND ' +
              '((EE_RefPointage = "' + vRefPointage + '") OR (EE_DatePointage = "' +
              UsdateTime(vDatePointage) + '"))', True);

    if not lQuery.Eof then
      Result := lQuery.FindField('Solde').AsFloat;
  finally
    Ferme(lQuery);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/05/2005
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CONSECR.NatureGeneAuxi ( vStNatGeneral, vStNatAuxiliaire : string ) : Boolean;
begin
  Result := False;
  if (vStNatAuxiliaire = 'FOU') or (vStNatAuxiliaire = 'AUC') then
  begin
    if (vStNatGeneral = 'COF') or (vStNatGeneral = 'COD') then
      Result := True
    else
      if (vStNatGeneral = 'COC') or (vStNatGeneral = 'COS') then
        Result := False;
  end
  else
  if (vStNatAuxiliaire = 'CLI') or (vStNatAuxiliaire = 'AUD') then
  begin
    if (vStNatGeneral = 'COF') or (vStNatGeneral = 'COS') then
      Result := False
    else
      if (vStNatGeneral = 'COC') or (vStNatGeneral = 'COD') then
        Result := True;
  end
  else
  if (vStNatAuxiliaire = 'DIV') then
  begin
    if (vStNatGeneral = 'COF') or (vStNatGeneral = 'COC') or (vStNatGeneral = 'COD') then
      Result := True
    else
      if (vStNatGeneral = 'COS') then
        Result := False;
  end
  else
  if (vStNatAuxiliaire = 'SAL') then
  begin
    if (vStNatGeneral = 'COD') or (vStNatGeneral = 'COS') then
      Result := True
    else
      if (vStNatGeneral = 'COF') or (vStNatGeneral = 'COC') then
        Result := False;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CONSECR.OnElipsisClickE_General(Sender: TObject);
begin
  LookUpList( THEdit(Sender),
              'Comptes généraux',
              'GENERAUX',
              'G_GENERAL' ,
              'G_LIBELLE',
              CGenereSQLConfidentiel('G'),
              'G_GENERAL' ,
              True,
              1 );
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CONSECR.OnExitE_General(Sender: TObject);
begin
  ChargeGeneral(False);
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CONSECR.OnClickBGDOWN(Sender: TObject);
begin
  GeneSuivant(True, True);
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_CONSECR.OnClickBGUP(Sender: TObject);
begin
  GeneSuivant(False, True);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/10/2004
Modifié le ... : 11/05/2006 - FQ 17374    
Description .. : Fonction qui passe au compte général suivant de la sélection
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.GeneSuivant(vBoSuiv : Boolean; vBoResteSurZone : Boolean);
var
  lQuery: TQuery;
  lSQL, lCptGene, lNewCpt, lCptLibelle: string;

  lStPrefixeE : string;
  lStPrefixeP : string;

  lStCodeExoE : string;
  lStCodeExoP : string;
begin
  AFocusFListe := False;
  lCptGene     := E_GENERAL.Text;
  lNewCpt      := lCptGene;

  lStPrefixeE := 'E';
  lStPrefixeP := 'P';
  lStCodeExoE := VH^.Encours.Code;
  lStCodeExoP := VH^.Precedent.Code;

  if (ctxPcl in V_Pgi.PgiContexte) and
     (VH^.CPExoref.Code = VH^.Suivant.Code) then
  begin
    lStPrefixeE := 'S';
    lStPrefixeP := 'E';

    lStCodeExoE := VH^.Suivant.Code;
    lStCodeExoP := VH^.Encours.Code;
  end;

  // GCO - 11/05/2006 - Défilement des comptes
  if ComboDefilCpt.Value = 'EXS' then
  begin
    lSql := 'SELECT DISTINCT(E_GENERAL), G_GENERAL, G_LIBELLE FROM ' + FStNomBase + 'ECRITURE ' +
            'LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL ' +
            'LEFT JOIN JOURNAL ON J_JOURNAL = E_JOURNAL WHERE ' +
            'E_QUALIFPIECE = "N" AND ' +
            '(E_EXERCICE = "' + lStCodeExoE + '" OR E_EXERCICE = "' + lStCodeExoP + '") AND ' +
            'J_NATUREJAL <> "ANO"';
  end
  else
  begin
    lSql := 'SELECT G_GENERAL,G_LIBELLE FROM ' + FStNomBase + 'GENERAUX WHERE ';

    if ComboDefilCpt.Value = 'EXT' then
    begin // Comptes mouvementés sur les deux exercices (avec ANO)
      lSql := lSql + '((G_TOTDEB' + lStPrefixeE + ' <> 0 OR G_TOTCRE' + lStPrefixeE + ' <> 0) OR' +
              ' (G_TOTDEB' + lStPrefixeP + ' <> 0 OR G_TOTCRE' + lStPrefixeP + ' <> 0))';
    end
    else
    if ComboDefilCpt.Value = 'NSL' then
    begin
      // Comptes non soldés sur les deux exercices
      lSql := lSql + '((G_TOTDEB' + lStPrefixeE + ' <> G_TOTCRE' + lStPrefixeE + ') OR' +
              ' (G_TOTDEB' + lStPrefixeP + ' <> G_TOTCRE' + lStPrefixeP + '))';
    end;
  end;
  // FIN GCO - 11/05/2006

  // Gestion du V_Pgi.Confidentiel
  lSql := lSql + ' AND ' + CGenereSQLConfidentiel('G');

  // Précédent ou Suivant + Order By
  if vBoSuiv then
    lSql := lSql + ' AND G_GENERAL > "' + lCptGene + '" ORDER BY G_GENERAL'
  else
    lSql := lSql + ' AND G_GENERAL < "' + lCptGene + '" ORDER BY G_GENERAL DESC';

  lQuery := OpenSQL(lSQL, True);
  if not lQuery.EOF then
  begin
    lNewCpt     := lQuery.FindField('G_GENERAL').AsString;
    lCptLibelle := lQuery.FindField('G_LIBELLE').AsString;
  end;
  Ferme(lQuery);

  if lNewCpt <> lCptGene then
  begin
    AFocusFListe := True;
    E_GENERAL.Text := lNewCpt;
    GLibelle.Caption := lCptLibelle;
    ChargeGeneral(False);
    RefreshPclPge;

    if vBoResteSurZone then
      SetFocusControl('EGENERAL');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/01/2003
Modifié le ... : 27/10/2006
Description .. : Charge FTobGene avec le contenu de E_GENERAL et modifie le
Suite ........ : E_AUXILIAIRE.DataType
Mots clefs ... :
*****************************************************************}
function TOF_CONSECR.ChargeGeneral( Force : boolean ) : Boolean;
var lQuery : TQuery;
    lStGeneral : string; // GCO - 27/10/2006 - 18973
begin
  Result := False;

  if ((FTobGene = nil) or (csDestroying in E_General.ComponentState)) then Exit;

  if E_General.Text = '' then
  begin
    FBoConfidentielGene := False;
    FTobGene.InitValeurs;
    GLibelle.Caption := '';
    E_Auxiliaire.DataType := 'TZTTOUS';
    WhatIsOk;
    Exit;
  end;

  // Placement sur la dernière écriture si changement de compte
  if (not FOkLastRow) and (E_General.Text <> FTobGene.GetString('G_GENERAL')) then
    FOkLastRow := True;

  //if (FTobGene.GetString('G_GENERAL') = E_General.Text) then Exit;

  lStGeneral := BourreEtLess(E_General.Text, fbGene);
  try
    lQuery := OpenSQL('SELECT * FROM ' + FStNomBase + 'GENERAUX WHERE ' +
                      'G_GENERAL = "' + lStGeneral + '"', True);

    if not lQuery.Eof then
    begin // GCO - 15/11/2006 - FQ 18973
      E_General.Text := lStGeneral;
      FTobGene.SelectDB('', lQuery);
      // Test de la confidentialité des comptes
      FBoConfidentielGene := EstConfidentiel(FTobGene.GetString('G_CONFIDENTIEL'));
      if not FBoConfidentielGene then
      begin
        GLibelle.Caption    := FTobGene.GetString('G_LIBELLE');
        AfficheSoldeCompte; // GCO - 29/08/2007 - FQ 20887
        if (FTobGene.GetString('G_COLLECTIF') <> 'X') then
        begin
          E_Auxiliaire.Text := '';
          ChargeAuxiliaire(False);
        end;
        E_AUXILIAIRE.DataType := QuelDataTypeTNat(FTobGene.GetString('G_NATUREGENE'));
        WhatIsOk; // GCO - 31/01/2003
        VerifieLesDates;
      end;
    end
    else
    begin
      E_Auxiliaire.DataType := 'TZTTOUS';
      E_General.ElipsisClick(nil); //FB 11809
    end;
  finally
    Ferme( lQuery );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/10/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.OnClickBTDOWN(Sender: TObject);
begin
  AuxSuivant(True, True);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/10/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.OnClickBTUP(Sender: TObject);
begin
  AuxSuivant(False, True);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/10/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnClickBBlocNote(Sender: TObject);
var lQEcr : TQuery;
begin
  inherited;
  HPB.Visible := BBlocNote.Down;
  if BBlocNote.Down then
  begin
    FBLocNote.Clear;
    if (not RecupEcrSelect) or (not EstVraiEcr(FTobEcrSelect)) then Exit;
    lQEcr := nil;
    try
      lQEcr := OpenSql('SELECT E_BLOCNOTE FROM ' + FStNomBase + 'ECRITURE WHERE ' +
        '(E_JOURNAL = "'       + FTobEcrSelect.GetString('E_JOURNAL') + '") AND ' +
        '(E_EXERCICE = "'      + FTobEcrSelect.GetString('E_EXERCICE') + '") AND ' +
        '(E_DATECOMPTABLE = "' + USDateTime(FTobEcrSelect.GetDateTime('E_DATECOMPTABLE')) + '") AND ' +
        '(E_NUMEROPIECE = '    + IntToStr(FTobEcrSelect.GetInteger('E_NUMEROPIECE')) + ') AND ' +
        '(E_NUMLIGNE = '       + IntToStr(FTobEcrSelect.GetInteger('E_NUMLIGNE')) + ') AND ' +
        '(E_NUMECHE = '        + IntToStr(FTobEcrSelect.GetInteger('E_NUMECHE')) + ') AND ' +
        '(E_QUALIFPIECE = "'   + FTobEcrSelect.GetString('E_QUALIFPIECE') + '")', True);

      StringToRich(FBlocNote, lQEcr.FindField('E_BLOCNOTE').AsString);
    finally
      Ferme(lQEcr);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/10/2005
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.OnClickBValider (Sender : TObject);
var i, lInPos      : integer;
    lStNomChamp    : string;
    lStListeChamps : string;
begin
  // Enregistrement du détail des écritures dans GTobDetailEcr pouur les utiliser
  // avec un verbe OLE
  if GTobDetailEcr <> nil then
  begin
    if ATobFliste.Detail.Count > 0 then
    begin
      lStListeChamps := FStListeChamps;
      GTobDetailEcr.Dupliquer(ATobFListe, True, True, False);

      // Ajout de l'entête GENERAL, AUXILIAIRE au niveau de la MERE
      GTobDetailEcr.AddChampSupValeur('GENERAL', E_General.Text, False);
      GTobDetailEcr.AddChampSupValeur('AUXILIAIRE', E_Auxiliaire.Text, False);
      // Fin d'ajout de l'entête

      for i := 0 to FListe.ColCount-1 do
      begin
        lStNomChamp := ReadTokenSt( lStListeChamps );
        // Suppression dans GTobDetailEcr des colonnes non visibles de la grille
        if FListe.ColWidths[i+1] <= 0 then
        begin
          GTobDetailEcr.Detail[0].DelChampSup( lStNomChamp, True);
        end;
      end;

      // Suppression des champs CLE_
      lStListeChamps := cSelect;
      lStListeChamps := FindEtReplace( lStListeChamps, ',', ';', True);
      while lstListeChamps <> '' do
      begin
        lStNomChamp := ReadTokenSt( lStListeChamps );
        lInPos := Pos( 'CLE_', lStNomChamp );
        if lInPos > 0 then
        begin
          lStNomChamp := Copy( lStNomChamp, lInPos, Length( lStNomChamp ));
          GTobDetailEcr.Detail[0].DelChampSup( lStNomChamp, True);
        end;
      end;

      // Suppression systématique du champ E_BLOCNOTE
      GTobDetailEcr.Detail[0].DelChampSup('E_BLOCNOTE', True);

    end; // FIN if ATobFliste.Detail.Count > 0 then
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/10/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnElipsisClickE_Auxiliaire(Sender: TObject);
var
 lStWhere : string;
begin
  lStWhere := '';
  
  if E_Auxiliaire.DataType = 'TZTTOUTDEBIT' then
    lStWhere := '(T_NATUREAUXI = "AUD" OR T_NATUREAUXI = "CLI" OR T_NATUREAUXI = "DIV")'
  else
    if E_Auxiliaire.DataType = 'TZTTOUTCREDIT' then
      lStWhere := '(T_NATUREAUXI = "AUC" OR T_NATUREAUXI = "FOU" OR T_NATUREAUXI = "DIV")'
    else
      if E_Auxiliaire.DataType = 'TZTSALARIE' then
        lStWhere := 'T_NATUREAUXI = "SAL"';

  {JP 02/11/05 : FQ 16853 : On exclut les tiers spécifiques à la Gescom}
  lStWhere := IIF( lStWhere <> '', ' AND ' + lStWhere, '');
  lStWhere := '(T_NATUREAUXI <> "NCP" AND T_NATUREAUXI <> "CON" AND T_NATUREAUXI <> "PRO" AND T_NATUREAUXI <> "SUS")' + lStWhere;

  LookUpList( THEdit(Sender),
           'Comptes auxiliaires',
           'TIERS',
           'T_AUXILIAIRE' ,
           'T_LIBELLE, T_NATUREAUXI' ,
           CGenereSQLConfidentiel('T') + IIF( lStWhere <> '', ' AND ' + lStWhere, '') ,
           'T_AUXILIAIRE' ,
           True,
           2 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.OnExitE_Auxiliaire(Sender: TObject);
begin
  ChargeAuxiliaire(False);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :   /  /
Modifié le ... : 11/05/2006 - FQ 17374
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.AuxSuivant(vBoSuiv : Boolean; vBoResteSurZone : Boolean);
var
  lQ: TQuery;
  lSQL, lCptAux, lNewCpt: string;
  lCptLibelle : string;

  lStPrefixeE : string;
  lStPrefixeP : string;
  lStCodeExoE : string;
  lStCodeExoP : string;

begin
  AFocusFListe := False;
  lCptAux := E_AUXILIAIRE.Text;
  lNewCpt := lCptAux;

  lStPrefixeE := 'E';
  lStPrefixeP := 'P';
  lStCodeExoE := VH^.Encours.Code;
  lStCodeExoP := VH^.Precedent.Code;

  if (ctxPcl in V_Pgi.PgiContexte) and
     (VH^.CPExoref.Code = VH^.Suivant.Code) then
  begin
    lStPrefixeE := 'S';
    lStPrefixeP := 'E';

    lStCodeExoE := VH^.Suivant.Code;
    lStCodeExoP := VH^.Encours.Code;
  end;

  // GCO - 11/05/2006 - Défilement des comptes
  if ComboDefilCpt.Value = 'EXS' then
  begin
    lSql := 'SELECT DISTINCT(E_AUXILIAIRE), T_AUXILIAIRE, T_LIBELLE FROM ' + FstNomBase + 'ECRITURE ' +
            'LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE ' +
            'LEFT JOIN JOURNAL ON J_JOURNAL = E_JOURNAL WHERE ' +
            'E_QUALIFPIECE = "N" AND ' +
            '(E_EXERCICE = "' + lStCodeExoE + '" OR E_EXERCICE = "' + lStCodeExoP + '") AND ' +
            'J_NATUREJAL <> "ANO" ';
  end
  else
  begin
    lSql := 'SELECT T_AUXILIAIRE, T_LIBELLE FROM ' + FStNomBase + 'TIERS WHERE ';

    if ComboDefilCpt.Value = 'EXT' then
    begin // Comptes mouvementés sur les deux exercices (avec ANO)
      lSql := lSql + '((T_TOTDEB' + lStPrefixeE + ' <> 0 OR T_TOTCRE' + lStPrefixeE + ' <> 0) OR' +
              ' (T_TOTDEB' + lStPrefixeP + ' <> 0 OR T_TOTCRE' + lStPrefixeP + ' <> 0))';

    end
    else
    if ComboDefilCpt.Value = 'NSL' then
    begin
      // Comptes non soldés sur les deux exercices
      lSql := lSql + '((T_TOTDEB' + lStPrefixeE + ' <> T_TOTCRE' + lStPrefixeE + ') OR' +
              ' (T_TOTDEB' + lStPrefixeP + ' <> T_TOTCRE' + lStPrefixeP + '))';
    end;
  end;
  // FIN GCO - 11/05/2006

  // Gestion du V_Pgi.Confidentiel
  lSql := lSql + ' AND ' + CGenereSQLConfidentiel('T');

  // GCO - 29/08/2007 - Nature des comptes
  if (CBAvecCollectif.Checked) then
    lSql := lSql + ' AND ' + CGenereSQLNatureAuxi;

  //if (CBAvecCollectif.Checked) and (FTobTiers.GetString('T_NATUREAUXI') <> '') then
  //  lSQL := lSQL + ' AND T_NATUREAUXI = "' + FTobTiers.GetString('T_NATUREAUXI') + '"';


  if vBoSuiv then
    lSQL := lSql + ' AND T_AUXILIAIRE > "' + lCptAux + '" ORDER BY T_AUXILIAIRE'
  else
    lSQL := lSql + ' AND T_AUXILIAIRE < "' + lCptAux + '" ORDER BY T_AUXILIAIRE DESC';

  lQ := OpenSQL(lSQL, True);
  if not lQ.EOF then
  begin
    lNewCpt := lQ.FindField('T_AUXILIAIRE').AsString;
    lCptLibelle := lQ.FindField('T_LIBELLE').AsString;
  end;
  Ferme(lQ);

  if lNewCpt <> lCptAux then
  begin
    AFocusFListe := True;
    E_AUXILIAIRE.Text := lNewCpt;
    TLibelle.Caption := lCptLibelle;
    ChargeAuxiliaire(True);
    RefreshPclPge;

    if vBoResteSurZone then
      SetFocusControl('EAUXILIAIRE');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/01/2003
Modifié le ... : 23/02/2004
Description .. : Charge FTobTiers avec le contenu de E_AUXILIAIRE et modifie le
Suite ........ : E_GENERAL.TEXT avec le collectif associé au compte auxiliaire
Mots clefs ... :
*****************************************************************}
function TOF_CONSECR.ChargeAuxiliaire(Force: boolean): Boolean;
var lQuery : TQuery;
    lStAuxiliaire : string; // GCO - 27/10/2006 - 18973
begin
  Result := False;

  if (FTobTiers = nil) or (csDestroying in E_Auxiliaire.ComponentState) then Exit;

  if E_Auxiliaire.Text = '' then
  begin
    FBoConfidentielAuxi := False;
    FTobTiers.InitValeurs;
    TLibelle.Caption := '';
    WhatIsOk;
    Exit;
  end;

  // Placement sur la dernière écriture si changement de compte
  if (not FOkLastRow) and (E_Auxiliaire.Text <> FTobTiers.GetString('T_AUXILIAIRE')) then
    FOkLastRow := True;

  //if (FTobTiers.GetValue('T_AUXILIAIRE') = E_Auxiliaire.Text) then Exit;

  lStAuxiliaire := BourreEtLess(E_Auxiliaire.Text, fbAux);
  try
    lQuery := OpenSQL('SELECT * FROM ' + FStNomBase + 'TIERS WHERE ' +
                      'T_AUXILIAIRE = "' + lStAuxiliaire + '"', True);
    if not lQuery.Eof then
    begin // GCO - 15/11/2006 - FQ 18973
      E_Auxiliaire.Text := lStAuxiliaire;
      FTobTiers.SelectDB('', lQuery);
      FBoConfidentielAuxi := EstConfidentiel(FTobTiers.GetString('T_CONFIDENTIEL'));
      if not FBoConfidentielAuxi then
      begin
        TLibelle.Caption := FTobTiers.GetString('T_LIBELLE');
        AfficheSoldeCompte; // GCO - 29/08/2007 - FQ 20887

        if CBAvecCollectif.Checked then
        begin
          if E_General.Text = '' then
            E_General.Text := FTobTiers.GetString('T_COLLECTIF')
          else
          begin
            if (FTobGene.GetString('G_COLLECTIF') = 'X') then
            begin
              if not NatureGeneAuxi(FTobGene.GetString('G_NATUREGENE'), FTobTiers.GetString('T_NATUREAUXI')) then
                E_GENERAL.Text := FTobTiers.GetString('T_COLLECTIF');
            end
            else
              E_GENERAL.Text := FTobTiers.GetString('T_COLLECTIF');
          end;
          ChargeGeneral(True);
        end;
        WhatIsOk; // GCO - 31/01/2003
        VerifieLesDates;
      end;
    end
    else
    begin
      E_Auxiliaire.ElipsisClick(nil);
    end;
  finally
    Ferme( lQuery );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/02/2003
Modifié le ... :   /  /
Description .. : Demande de justification de solde
Suite ........ : GCO - 08/09/2004 - FQ 13083 - Conservation de la date comptable
Suite ........ : supérieure au lieu de mettre 31/12/2099 systématiquement
Mots clefs ... :
******************************************************************}
procedure TOF_CONSECR.OnClickCBJustifSolde(Sender: TObject);
var lDateJustifSolde : TDateTime;
begin
{$IFDEF BUREAU} // GCO - 27/12/2005 - FQ 17230 et FQ 17256
  SetControlEnabled('E_EXERCICE', (not FBoBlocageTotal) and (not CbJustifSolde.Checked));
{$ELSE}
  SetControlEnabled('E_EXERCICE', not CbJustifSolde.Checked);
{$ENDIF}
  SetControlEnabled('E_DATECOMPTABLE', not CbJustifSolde.Checked);

  if FOkLettrable then
    FBoStateCBJustifSolde := CBJustifSolde.Checked;

  if CbJustifSolde.Checked then
  begin
    lDateJustifSolde := StrToDate(E_DateComptable_.Text);
    E_Exercice.ItemIndex := 0;
    E_Exercice.OnChange(Self);
    // GCO - 07/01/2005 - FQ 13083 
    if lDateJustifsolde > StrToDate(E_DateComptable.Text) then
      E_DateComptable_.Text := DateToStr(lDateJustifSolde)
    else
      E_DateComptable_.Text := DateToStr(iDate2099);
      
    E_EtatLettrage.Text   := '<<' + TraduireMemoire('Tous') + '>>';
    E_Lettrage.Text       := '';
    E_NumeroPiece.Text    := '';
    E_NumeroPiece_.Text   := '';
  end
  else
    ActivationOption;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_CONSECR.OnClickCBDetailAno(Sender: TObject);
begin
  if FOkLettrable then
    FBoStateCBDetailANO := CBDetailAno.Checked;
  ActivationOption;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/08/2002
Modifié le ... :   /  /
Description .. : Demande le chargement des infos supplémentaires de l'écriture
Suite ........ : sélectionnée dans la grille
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnRowEnterFListe(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  inherited;
  AfficheInfoSup;
  OnClickBBlocNote(nil);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/02/2003
Modifié le ... : 23/02/2004
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOF_CONSECR.BeforeLoad: boolean;
begin
  inherited BeforeLoad;

  ChargeAuxiliaire(False);
  ChargeGeneral(False);

  // Autorisation de faire la requête, tout les tests sont OK
  FBoFaireRequete := (not FBoConfidentielGene) and
                     (not FBoConfidentielAuxi) and
                     (FBoOkControleDate);

  // Modification des colonnes de la grille en fonction des comptes
  InitColFListe;

  //
  Result := True;
  FTobEcrSelect.InitValeurs ; // Init des champs de la tob
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/02/2003
Modifié le ... : 27/09/2004
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CONSECR.AfterLoad: boolean;
var lCancel : Boolean;
begin
  Result := inherited AfterLoad;

  if FListe.Focused then
  begin
    if FOkLastRow then
    begin
      FListe.Row := FListe.RowCount - 1;
      // Oblige de re-déclencher l'événement à la main
      FListe.OnRowEnter( nil, FListe.Row, lCancel, False );
      FOkLastRow := False;
    end
  end;

  // Confidentialité du GENERAL
  if FBoConfidentielGene then
  begin
    PgiInfo(TraduireMemoire('Le compte') + ' ' + E_General.Text + ' ' +
            TraduireMemoire('est confidentiel.'),
            TraduireMemoire('Consultation des écritures'));
  end;

  // Confidentialité de l' AUXILIAIRE
  if FBoConfidentielAuxi then
  begin
    PGIInfo(TraduireMemoire('Le compte') + ' ' + E_Auxiliaire.Text + ' ' +
            TraduireMemoire('est confidentiel.'),
            TraduireMemoire('Consultation des écritures'));
  end;

  // Affichage des informations complémentaires de l'écriture sélectionnée
  AfficheInfoSup;

  // GCO - 31/07/2007 - FQ 20960 - Ici pour être sur que ATobFListe soit chargé
  FOkEcr := FOkCreateModif and (ATobFListe.Detail.Count <> 0);

  InitPopUpComptable ( True );
  InitPopUpSpecifique( True );
  InitPopUpEdition   ( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/08/2002
Modifié le ... :   /  /
Description .. : Met à jour les informations supplémentaires de l' écriture sélectionné
Suite ........ : dans la grille
Mots clefs ... :
*****************************************************************}

procedure TOF_CONSECR.AfficheInfoSup;
var lPrixUnitaire : Double;
    lQ            : TQuery;

  function Affecte(psz: string): string;
  var v: Variant;
  begin
    v := FTobEcrselect.GetString(psz);
    if VarIsNull(v) then
      Result := ''
    else
      Result := v;
  end;
  
begin
  THLabel(GetControl('Z_NumLigne')).Caption := '';
  THLabel(GetControl('Z_RefExterne')).Caption := '';
  THLabel(GetControl('Z_DateRefExterne')).Caption := '';
  THLabel(GetControl('Z_RefLettrage')).Caption := '';

  THLabel(GetControl('Z_DateEcheance')).Caption := '';
  THLabel(GetControl('Z_ModePaie')).Caption := '';
  THLabel(GetControl('Z_Auxiliaire')).Caption := '';
  THLabel(GetControl('Z_RefPointage')).Caption := '';
  THLabel(GetControl('Z_DatePointage')).Caption := '';
  THLabel(GetControl('Z_Devise')).Caption := '';
  THLabel(GetControl('Z_Debitdev')).Caption := '';

  THLabel(GetControl('Z_Qte1')).Caption := '';
  THLabel(GetControl('Z_Qte2')).Caption := '';
  THLabel(GetControl('Z_QualifQte1')).Caption := '';
  THLabel(GetControl('Z_QualifQte2')).Caption := '';

  THLabel(GetControl('Z_PrixUnitaire1')).Caption := '';
  THLabel(GetControl('Z_PrixUnitaire2')).Caption := '';

  THLabel(GetControl('Z_Affaire')).Caption := '';
  THLabel(GetControl('Z_Table0')).Caption := '';
  THLabel(GetControl('Z_Table1')).Caption := '';
  THLabel(GetControl('Z_Table2')).Caption := '';
  THLabel(GetControl('Z_Table3')).Caption := '';

  SetControlText('EC_CUTOFFDEB' , '' ) ;
  SetControlText('EC_CUTOFFFIN' , '' ) ;

  if (ComboInfoSup.ItemIndex = 0) or (not RecupEcrSelect) then Exit;

  THLabel(GetControl('Z_NumLigne')).Caption       := FTobEcrSelect.GetValue('E_NUMLIGNE');
  THLabel(GetControl('Z_RefExterne')).Caption     := Affecte('E_REFEXTERNE');

  THLabel(GetControl('Z_DateRefExterne')).Caption := FTobEcrSelect.GetValue('E_DateRefExterne');
  THLabel(GetControl('Z_RefLettrage')).Caption    := Affecte('E_REFLETTRAGE');
  THLabel(GetControl('Z_DateEcheance')).Caption   := FTobEcrSelect.GetValue('E_DateEcheance');
  THLabel(GetControl('Z_ModePaie')).Caption       := Affecte('E_ModePaie');
  THLabel(GetControl('Z_Auxiliaire')).Caption     := Affecte('E_Auxiliaire');
  THLabel(GetControl('Z_RefPointage')).Caption    := Affecte('E_RefPointage');
  THLabel(GetControl('Z_DatePointage')).Caption   := FTobEcrSelect.GetValue('E_DatePointage');
  THLabel(GetControl('Z_Devise')).Caption         := Affecte('E_Devise');
  THLabel(GetControl('Z_Debitdev')).Caption       := IIF(FTobEcrSelect.GetValue('E_CreditDev') <> 0, FTobEcrSelect.GetValue('E_CreditDev'), FTobEcrSelect.GetValue('E_DebitDev'));

  THLabel(GetControl('Z_Qte1')).Caption           := StrfMontant(FTobEcrSelect.GetDouble('E_QTE1'), 13, V_Pgi.OkDecQ, '', True);
  THLabel(GetControl('Z_Qte2')).Caption           := StrfMontant(FTobEcrSelect.GetDouble('E_QTE2'), 13, V_Pgi.OkDecQ, '', True);

  THLabel(GetControl('Z_Affaire')).Caption        := Affecte('E_Affaire');
  THLabel(GetControl('Z_Table0')).Caption         := Affecte('E_Table0');
  THLabel(GetControl('Z_Table1')).Caption         := Affecte('E_Table1');
  THLabel(GetControl('Z_Table2')).Caption         := Affecte('E_Table2');
  THLabel(GetControl('Z_Table3')).Caption         := Affecte('E_Table3');

  // GCO - 28/10/2005 - Ajout de la Gestion des quantités
  if PageinfoSup.ActivePage = PageInfoSup.Pages[1] then
  begin
    if FTobEcrSelect.GetDouble('E_QTE1') <> 0 then
    begin
      if FTobEcrSelect.GetString('E_DEVISE') <> V_PGI.DevisePivot then
        lPrixUnitaire := Abs(FTobEcrSelect.GetDouble('E_DEBITDEV') -
                            FTobEcrSelect.GetDouble('E_CREDITDEV')) /
                            FTobEcrSelect.GetDouble('E_QTE1')
      else
        lPrixUnitaire := Abs(FTobEcrSelect.GetDouble('E_DEBIT') -
                            FTobEcrSelect.GetDouble('E_CREDIT')) /
                            FTobEcrSelect.GetDouble('E_QTE1');
    end
    else
      lPrixUnitaire := 0;

    THLabel(GetControl('Z_PrixUnitaire1')).Caption :=
      StrFMontant(lPrixUnitaire, 13, V_Pgi.OkDecP, '', True) + ' ' +
      RechDom('TTDEVISE', FTobEcrSelect.GetString('E_DEVISE'), False);

    if lPrixUnitaire <> 0 then
      THLabel(GetControl('Z_PrixUnitaire1')).Caption :=
      THLabel(GetControl('Z_PrixUnitaire1')).Caption + ' / ' +
      RechDom('TTQUALUNITMESURE', FTobEcrSelect.GetString('E_QUALIFQTE1'), False);

    if FTobEcrSelect.GetDouble('E_QTE2') <> 0 then
    begin
      if FTobEcrSelect.GetString('E_DEVISE') <> V_PGI.DevisePivot then
        lPrixUnitaire := Abs(FTobEcrSelect.GetDouble('E_DEBITDEV') -
                           FTobEcrSelect.GetDouble('E_CREDITDEV')) /
                           FTobEcrSelect.GetDouble('E_QTE2')
      else
        lPrixUnitaire := Abs(FTobEcrSelect.GetDouble('E_DEBIT') -
                           FTobEcrSelect.GetDouble('E_CREDIT')) /
                           FTobEcrSelect.GetDouble('E_QTE2');
    end
    else
      lPrixunitaire := 0;

    THLabel(GetControl('Z_PrixUnitaire2')).Caption :=
    StrFMontant(lPrixUnitaire, 13, V_Pgi.OkDecP, '', True) + ' ' +
    RechDom('TTDEVISE', FTobEcrSelect.GetString('E_DEVISE'), False);

    if lPrixUnitaire <> 0 then
      THLabel(GetControl('Z_PrixUnitaire2')).Caption :=
      THLabel(GetControl('Z_PrixUnitaire2')).Caption + ' / ' +
      RechDom('TTQUALUNITMESURE', FTobEcrSelect.GetString('E_QUALIFQTE2'), False);
  end;


  if FTobGene.GetValue('G_CUTOFF') = 'X' then
   begin
    lQ := OpenSQL('SELECT EC_CUTOFFDEB,EC_CUTOFFFIN FROM ECRCOMPL WHERE EC_JOURNAL="' + FTobEcrSelect.GetString('E_JOURNAL')      + '" ' +
          'AND EC_EXERCICE="'            + FTobEcrSelect.GetString('E_EXERCICE')                          + '" ' +
          'AND EC_DATECOMPTABLE="'       + usDateTime(FTobEcrSelect.GetDateTime('E_DATECOMPTABLE'))                + '" ' +
          'AND EC_EXERCICE="'            + FTobEcrSelect.GetString('E_EXERCICE')                          + '" ' +
          'AND EC_NUMEROPIECE='          + FTobEcrSelect.GetString('E_NUMEROPIECE')                       + ' '  +
          'AND EC_NUMLIGNE='             + FTobEcrSelect.GetString('E_NUMLIGNE')                          + ' '  +
          'AND EC_QUALIFPIECE="'         + FTobEcrSelect.GetString('E_QUALIFPIECE')                       + '" ' , true ) ;

    if not lQ.EOF then
     begin
      SetControlText('EC_CUTOFFDEB' , lQ.FindField('EC_CUTOFFDEB').asString ) ;
      SetControlText('EC_CUTOFFFIN' , lQ.FindField('EC_CUTOFFFIN').asString ) ;
     end ;// if

    Ferme(lQ) ;
  end ; // if

  SetControlVisible( 'EC_CUTOFFDEB' , FTobGene.GetValue('G_CUTOFF') = 'X' ) ;
  SetControlVisible( 'EC_CUTOFFFIN' , FTobGene.GetValue('G_CUTOFF') = 'X' ) ;
  SetControlVisible( 'TZ_AFFAIRE1'  , FTobGene.GetValue('G_CUTOFF') = 'X' ) ;
  SetControlVisible( 'TZ_AFFAIRE2'  , FTobGene.GetValue('G_CUTOFF') = 'X' ) ;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/08/2007
Modifié le ... :   /  /    
Description .. : FQ 20887
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.AfficheSoldeCompte;
begin
  LeSoldeN.Text          := '0,00';
  LeSoldeNMoins1.Text    := '0,00';

  SetControlVisible('LESOLDEN',       (E_General.Text = '') or (E_Auxiliaire.Text = ''));
  SetControlVisible('LESOLDENMOINS1', (E_General.Text = '') or (E_Auxiliaire.Text = ''));
  SetControlVisible('TLESOLDEN',      (E_General.Text = '') or (E_Auxiliaire.Text = ''));
  SetControlVisible('TLESOLDENMOINS1',(E_General.Text = '') or (E_Auxiliaire.Text = ''));

  if (E_General.Text <> '') and (E_Auxiliaire.Text = '') then
  begin
    LeSoldeN.Text       := AfficheDBCR(FTobGene.GetDouble('G_TOTDEBE') - FTobGene.GetDouble('G_TOTCREE'));
    LeSoldeNMoins1.Text := AfficheDBCR(FTobGene.GetDouble('G_TOTDEBP') - FTobGene.GetDouble('G_TOTCREP'));
  end;

  if (E_General.Text = '') and (E_Auxiliaire.Text <> '') then
  begin
    LeSoldeN.Text       := AfficheDBCR(FTobTiers.GetDouble('T_TOTDEBE') - FTobTiers.GetDouble('T_TOTCREE'));
    LeSoldeNMoins1.Text := AfficheDBCR(FTobTiers.GetDouble('T_TOTDEBP') - FTobTiers.GetDouble('T_TOTCREP'));
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/02/2003
Modifié le ... : 14/02/2003
Description .. : Accès à la saisie en pointant sur l'écriture sélectionné
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnDblClickFListe(Sender: TObject);
{$IFDEF COMPTA}
var
  AA: TActionFiche;
  M: RMVT;
  lQEcr: TQuery;
{$ENDIF}
begin
{$IFDEF COMPTA}
  // GCO - 29/07/2005 - FQ 16335 - Blocage des fonctions avec raccourci si MultiSelection
  if ABoMultiSelected then Exit;
  // GCO - 07/07/2004 - FQ 13614
  if (not FListe.Focused) or (not RecupEcrSelect) then Exit;

  lQEcr := nil;
  try
    lQEcr := OpenSql('SELECT * FROM ' + FStNomBase + 'ECRITURE WHERE ' +
      '(E_JOURNAL = "' + FTobEcrSelect.GetValue('E_JOURNAL') + '") AND ' +
      '(E_EXERCICE = "' + FTobEcrSelect.GetValue('E_EXERCICE') + '") AND ' +
      '(E_DATECOMPTABLE = "' + USDateTime(FTobEcrSelect.GetValue('E_DATECOMPTABLE')) + '") AND ' +
      '(E_NUMEROPIECE = ' + IntToStr(FTobEcrSelect.GetValue('E_NUMEROPIECE')) + ') AND ' +
      '(E_NUMLIGNE = ' + IntToStr(FTobEcrSelect.GetValue('E_NUMLIGNE')) + ') AND ' +
      '(E_NumEche = ' + IntToStr(FTobEcrSelect.GetValue('E_NUMECHE')) + ') AND ' +
      '(E_QualifPiece = "' + FTobEcrSelect.GetValue('E_QUALIFPIECE') + '")', True);

    AA := taModif;
    if not FOkCreateModif then
      AA := taConsult;

    if (lQEcr.FindField('E_MODESAISIE').AsString <> '-') and
      (lQEcr.FindField('E_MODESAISIE').AsString <> '') then // Bordereau
      LanceSaisieFolio(lQEcr, AA)
    else
    begin
      // GC - 26/03/2003
      if TrouveSaisie(lQEcr, M, lQEcr.FindField('E_QUALIFPIECE').Asstring) then
      begin
        M.NumLigVisu := lQEcr.FindField('E_NUMLIGNE').AsInteger;
        LanceSaisie(lQEcr, AA, M);
      end;
    end;

  finally
    Ferme(lQEcr);
  end;

  RefreshPclPGE;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/07/2005
Modifié le ... : 19/12/2007
Description .. :
Suite ........ : GCO - 19/12/2007 - FQ 22032
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnBeforeFlipFListe(Sender: TObject; ARow: Integer; var Cancel: Boolean);
var lTobEcr : Tob;
begin
  lTobEcr := GetO(FListe);
  if (lTobEcr = nil) or
     (not EstVraiEcr( lTobEcr )) or
     (lTobEcr.GetString('CLE_ECRANOUVEAU') = 'OAN') then
     Cancel := True;

  if (lTobEcr <> nil) and (lTobEcr.GetString('CLE_ECRANOUVEAU') = 'OAN') then
    PgiInfo(TraduireMemoire('Vous ne pouvez pas faire de traitement ' +
            'sur les écritures d''a-nouveaux PGI.'),
            TraduireMemoire('Sélection d''une écriture'));
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/07/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnFlipSelectionFListe(Sender: TObJect);
var lTobEcr : Tob;
begin
  inherited;
  lTobEcr := GetO(FListe);
  if lTobEcr = nil then Exit;

  BComptable.Enabled      := not ABoMultiSelected;
  BSpecifique.Enabled     := not ABoMultiSelected;
  BAideRevision.Enabled   := not ABoMultiSelected;
  BRevInteg.Enabled       := not ABoMultiSelected;
  BParametre.Enabled      := not ABoMultiSelected;
  BEdition.Enabled        := not ABoMultiSelected;
  BFctMultiSelect.Enabled := ABoMultiSelected;
end;

//--------------------- FONCTIONS DU POPUPREVISION -----------------------------
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/09/2002
Modifié le ... : 02/01/2007
Description .. : Détail du compte ICC
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnClickDetailIcc(Sender: TObject);
begin
  AGLLanceFiche('CP', 'ICCFICHEGENERAUX', '', E_GENERAL.Text,
                'ACTION=MODIFICATION;' +
                E_DATECOMPTABLE.Text + ';' +
                E_DATECOMPTABLE_.Text);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/01/2007
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.OnClickDetailCapitalICC(Sender: TObject);
var lStArgument : string;
begin
  lStArgument := GetParamSocSecur('SO_ICCCOMPTECAPITAL', '');

  if lStArgument <> '' then
    AGLLanceFiche('CP', 'ICCFICHEGENERAUX', '', lStArgument,
                  'ACTION=MODIFICATION;' +
                  E_DATECOMPTABLE.Text + ';' +
                  E_DATECOMPTABLE_.Text);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/01/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.OnClickCalculIcc( Sender : TObject ); // Calucl des intérêts ICC
begin
  AGLLanceFiche('CP', 'ICCPARAMETRE', '', '', E_DATECOMPTABLE.Text + ';' +
                E_DATECOMPTABLE_.Text + ';' + E_GENERAL.Text);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/09/2002
Modifié le ... : 02/01/2007
Description .. : Calcul de reboursement d'emprunt
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickEmpruntCRE(Sender: TObject);
begin
  AGLLanceFiche('FP', 'FMULEMPRUNT', '', '', E_GENERAL.Text);
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 02/03/2007
Modifié le ... :
Description .. : Rapprochement Compta / CRE
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickRappECRCRE(Sender: TObject);
begin
  EtatRapproCRECompta ;
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/02/2007
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickGCD(Sender: TObject);
var
 i         : integer ;
 lTOBEcr   : TOB ;
 lTOBListe : TOB ;
begin
 lTobEcr := GetO(FListe);
 if (lTobEcr = nil) or (not EstVraiEcr( lTobEcr )) then exit ;

 lTOBEcr := TOB.Create('', nil , - 1 ) ;

 try

 if ABoMultiSelected then
  begin

    for i := 1 to FListe.RowCount - 1 do
     if FListe.IsSelected(i) then
      begin
       lTobListe := GetO(FListe) ;
       if ( lTobListe.getValue('CLE_ECRANOUVEAU') = 'OAN' ) then continue ;
       if ( VH^.ExoV8.Code <> '' ) and ( lTobListe.getValue('E_DATECOMPTABLE') < VH^.ExoV8.Deb ) then continue ;
       if not RecupEcrSelect(i) then exit ;
       if ( FTobEcrSelect.getValue('E_ETATLETTRAGE') = 'TL' ) then exit ;
       FTobEcrSelect.ChangeParent(lTOBEcr,-1) ;
       FTobEcrSelect := TOB.Create('ECRITURE', nil, -1) ;
      end; // if
  end
   else
    begin
     RecupEcrSelect ;
     if ( FTobEcrSelect.getValue('E_ETATLETTRAGE') = 'TL' ) then exit ;
     FTobEcrSelect.ChangeParent(lTOBEcr,-1) ;
     FTobEcrSelect := TOB.Create('ECRITURE', nil, -1) ;
    end ;

 if lTOBEcr.Detail.Count > 0 then
  begin
   TheData := lTOBEcr ;
   CPLanceFiche_GCDOperation('','CONSULT') ;
   TheData := nil ;
   RefreshPclPGE ;
  end ;

 finally
  lTOBEcr.Free ;
 end ;
end ;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickFeuilleExcel(Sender: TObject);
begin
  CPLanceFiche_CPIMPFICHEXCEL ('FTS;' + GetControltext('EGENERAL') + ';');
end;
{$ENDIF}
{$ENDIF}


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/01/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickNoteTravail(Sender: TObject);
begin
  CPLanceFiche_CPNoteTravail( E_General.Text, False);
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/01/2007
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickTableauVariation(Sender: TObject);
begin
  CPLanceFiche_CPTableauVar( E_General.Text );
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickDRF(Sender: TObject);
begin
  LanceFiche_DRF( OkLanceDRF );
end;
{$ENDIF}
{$ENDIF}


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/10/2002
Modifié le ... : 02/01/2007
Description .. : Rapprochement Comptabilité / Immo
Mots clefs ... :
*****************************************************************}
{$IFDEF AMORTISSEMENT}
procedure TOF_CONSECR.OnClickRapproComptaImmo(Sender: TObject);
begin
  EtatRapprochementCompta;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/04/2007
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickRapproComptaGCD( Sender : TObject );
begin
 CPLanceFiche_CPMULGCD('GCDCOMPTA') ;
end;

procedure TOF_CONSECR.OnClickTableauGCD( Sender : TObject );
begin
 CPLanceFiche_CPMULGCD() ;
end ;

{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/01/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickDonneesComp(Sender: TObject);
var
  lObj : Variant;
  lstMsg : string;
  lstTableau : string;
  lstGeneral : string;
  i : integer;
begin
  lstGeneral :=  FTobGene.GetString('G_GENERAL');
  if (Copy(lstGeneral,1,2)='70') then lstTableau := 'SCVTILCA'
  else if ((FTobGene.GetValue('G_NATUREGENE')='IMO') and (Copy(lstGeneral,1,2)<>'28')
    and (Copy(lstGeneral,1,2)<>'29')) then lstTableau := 'SC2054P1'
  else if (FTobGene.GetValue('G_NATUREGENE')='IMO') then lstTableau := 'SC2055P1'
  else if (Copy(lstGeneral,1,3)='145') then lstTableau := 'SC2055P2'
  else if (Copy(lstGeneral,1,1)='4') then lstTableau := 'SC2057P2'
  else if ((Copy(lstGeneral,1,2)='67') or (Copy(lstGeneral,1,2)='77')) then lstTableau := 'SC2053P1'
  else lstTableau := 'SCVTILCA';

  lObj := CreateOleObject('CegidPgi.CegidPgi');
  lstMsg := lObj.LanceEtafi (V_PGI.NoDossier);
  if lstMsg <> '' then
  begin
    PGIBox(lstMsg);
    lObj := unassigned;
  end else
  begin
    lObj := unassigned;
    // Temporisation nécessaire pour attendre que Etafi soit correctement lancé
    for i := 0 to 20 do
    begin
      if ExisteSQL ('SELECT 1 FROM COURRIER WHERE MG_TYPE=42 AND MG_EXPEDITEUR="'+V_PGI.User+'"') then
      begin
        lObj := CreateOleObject('CegidPgi.LiasseBIC');
        lstMsg := lObj.SaisieComplementaireEtafi('E', VH^.Encours.Code, lstTableau);
        if lstMsg <> '' then PGIBox(lstMsg);
        lObj := unassigned;
        break;
      end;
      sleep (1000);
    end;
  end;
end;

{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/10/2006
Modifié le ... : 02/01/2007
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickInfoLiasse(Sender: TObject);
begin
  CPLanceFiche_CPControleLiasse(E_General.Text);
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/11/2006
Modifié le ... : 02/01/2007
Description .. : FQ 19161
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickTotalCptLiasse(Sender : TObject);
var lExoDate : TExoDate;
begin
  if CQuelExercice( FTobEcrSelect.GetDateTime('E_DATECOMPTABLE'), lExoDate) then
  begin
    CPLancefiche_CPTotalCptLiasse(FTobEcrSelect.GetString('E_GENERAL') + ';' +
                                  DateToStr(lExoDate.Deb) + ';' +
                                  DateToStr(lExoDate.Fin));
  end;
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/01/2007
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickVisaCompte(Sender: TObject);
var lBoVisaRevision : Boolean;
begin
  // GCO - 23/08/2005 - FQ 16248
  //if (FExoDate.Code = VH^.Encours.Code) and
  //   (not ExisteSql('SELECT MG_EXPEDITEUR FROM COURRIER')) then
  //begin
  lBoVisaRevision := IIF( FTobGene.GetString('G_VISAREVISION') = 'X', False, True);

  if not MiseAJourG_VisaRevision( FTobGene.GetString('G_GENERAL'), lBoVisaRevision )  then
    PgiError('Impossible de mettre à jour les informations de révision du compte.', Ecran.Caption);

  BCherche.Click;
  //end;
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/01/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickDocPermanent(Sender: TObject);
var
  lObj : Variant;
begin
  lObj := CreateOleObject( 'CegidPgi.CegidPgi' );
  lObj.ZoomDP(V_PGI.NoDossier,'Dossier permanent');
  lObj := unassigned;
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/01/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickDocAnnuel(Sender: TObject);
Var
  lobj : Variant;
begin
  lObj := CreateOleObject( 'CegidPgi.CegidPgi' );
  lObj.ZoomDP(V_PGI.NoDossier,'Dossier annuel');
  lObj := unassigned;
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/01/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickRechercheDoc(Sender: TObject);
var
  lObj : Variant;
begin
  lObj := CreateOleObject( 'CegidPgi.CegidPgi' );
  lObj.ZoomDP(V_PGI.NoDossier,'Recherche documentaire');
  lObj := unassigned;
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/02/2007
Modifié le ... : 01/04/2007    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickProgrammeTravail(Sender: TObject);
var lStRacine : HString;
begin
  // GCO - 18/09/2007 - FQ 21451
  lStRacine := '';
  if FTobGene.GetString('G_CYCLEREVISION') <> '' then
    lStRacine := Copy( FTobGene.GetString('G_CYCLEREVISION'), 0, 1);
  CPLanceFiche_CPRevProgTravail( VH^.Encours.Code, lStRacine );
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickAPG(Sender: TObject);
begin
  CPLanceFiche_CPREVINFODOSSIER( V_Pgi.NoDossier, VH^.EnCours.Code, 1, TaModif, 1);
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{$IFDEF COMPTA}
{$IFNDEF CCMP}
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/05/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnClickSCY(Sender: TObject);
begin
  CPLanceFiche_CPREVINFODOSSIER( V_Pgi.NoDossier, VH^.EnCours.Code, 1, TaModif, 2);
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{$IFDEF COMPTA}
{$IFNDEF CCMP}
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/05/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.OnClickEXP(Sender: TObject);
begin
  CPLanceFiche_CPREVINFODOSSIER( V_Pgi.NoDossier, VH^.EnCours.Code, 1, TaModif, 3);
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/01/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickMemoCycle(Sender: TObject);
begin
  CPLanceFiche_CPRevDocTravaux( E_General.Text, '', VH^.EnCours.Code, 0 );
  RefreshPclPge;
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/01/2007
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickMemoObjectif(Sender: TObject);
begin
  CPLanceFiche_CPRevDocTravaux( E_General.Text, '', VH^.EnCours.Code, 1 );
  RefreshPclPge;
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/01/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickMemoSynthese(Sender: TObject);
begin
  CPLanceFiche_CPRevDocTravaux( E_General.Text, '', VH^.EnCours.Code, 2 );
  RefreshPclPge;
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/01/2007
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickMemoMillesime(Sender: TObject);
begin
  CPLanceFiche_CPrevDocTravaux( E_General.Text, '', VH^.EnCours.Code, 3 );
  RefreshPclPge;
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/01/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickMemoCompte(Sender: TObject);
begin
  CPLanceFiche_CPrevDocTravaux( E_General.Text, '', VH^.EnCours.Code, 4 );
  RefreshPclPge;
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/07/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnPopUpEdition(Sender: TObject);
begin
  InitPopUpEdition( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.OnPopUpMultiSelect(Sender: TObject);
begin
  InitPopUpMultiSelect( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/05/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.ChargeFTobListeEcr;
var lTobTemp : Tob;
    i : integer;
begin
  FTobListeEcr.ClearDetail;
//  FBoGesComTTC := False;
  
  for i := 1 to FListe.RowCount - 1 do
  begin
    if FListe.IsSelected(i) then
    begin
      FListe.Row := i ;
      if not RecupEcrSelect then Exit ;
      lTobTemp := Tob.Create('ECRITURE', FTobListeEcr, -1);
      lTobTemp.Dupliquer(FTobEcrSelect, True, True);
                      (*
      if (FBoGescomTTC = false) and
         (lTobTemp.GetString('E_REFGESCOM') <> '') and
         (lTobTemp.Getstring('E_TYPEMVT') = 'TTC') and
         {JP 26/06/07 : FQ TRESO 10491 : on vérouille les flux originaires de la Tréso}
         (lTobTemp.Getstring('E_QUALIFORIGINE') = QUALIFTRESO) then
        FBoGesComTTc := True;
        *)
    end;
  end;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.OnPopUpPopF11(Sender: TObject);
{$IFDEF COMPTA}
var lMenuItem : TMenuItem;
{$ENDIF}
begin
{$IFDEF COMPTA}
  InitAllPopUp( True );
  inherited; // Ajoute les élements visibles des menus de Ecran dans PopF11
  if PopF11.Items.Count > 2 then begin
    lMenuItem := TMenuItem.Create(Ecran);

    lMenuItem.Caption := '-';
    PopF11.Items.Insert( PopF11.Items.Count - 2, lMenuItem);
  end;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/09/2002
Modifié le ... : 24/11/2006
Description .. : Branchement des PopUP avec les procédures et fonctions 
Suite ........ : à appeler
Suite ........ : - LG - 05/07/2006 - deplacement du chargeftoblisteecr ds 
Suite ........ : les fct respectives
Suite ........ : - LG - FB 19058 - 24/11/2006 - ouverture des fct de 
Suite ........ : pointage pour le pointage sur journal
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.InitAllPopUp(vActivation : Boolean);
begin
  // GCO - 02/06/2006 - FQ 13084
  if CritModified then
    BCherche.Click;

  WhatIsOk;
  RecupEcrSelect; // FTobEcrSelect = Récupération en base de l'écriture
  FZListJournal.Load([FTobEcrSelect.GetString('E_JOURNAL')]);

  InitPopUpComptable    ( vActivation );
  InitPopUpSpecifique   ( vActivation );
  InitPopUpAideRevision ( vActivation );
  InitPopUpRevInteg     ( vActivation );
  InitPopUpParametre    ( vActivation );
  InitPopUpMultiselect  ( vActivation );
  InitPopUpEdition      ( vActivation );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.OnPopUpComptable(Sender: TObject);
begin
  InitPopUpComptable( True);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/01/2007
Modifié le ... : 06/04/2007
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.InitPopUpComptable(vActivation : Boolean);
var i : integer;
    lMenuItem : TMenuItem;
    lBoAuMoinsUn : Boolean;
{$IFDEF COMPTA}
{$IFNDEF CCMP}
    lBoOkSaisieTreso: Boolean;
{$ENDIF}
{$ENDIF}
begin
{$IFDEF COMPTA}
{$IFNDEF CCMP}
  lBoOkSaisieTreso := False;
  if FOkCreateModif and (FTobGene.GetString('G_NATUREGENE') = 'BQE') and
     EstVraiEcr(ATobRow) then
  begin
    FZListJournal.Load([ATobRow.GetString('E_JOURNAL')]);
    lBoOkSaisieTreso := FZListJournal.GetValue('J_CONTREPARTIE') = ATobRow.GetString('E_GENERAL');
  end;
{$ENDIF}
{$ENDIF}
  RecupEcrSelect;

  lBoAuMoinsUn := False;
  for i := 0 to PopUpComptable.Items.Count - 1 do
  begin
    lMenuItem := PopUpComptable.Items[i];
    try
      lMenuItem.Enabled := False;
      if (vActivation and ABoMultiSelected) then Continue;

      if lMenuItem.Name = 'MODIFECR' then
      begin
        if vActivation then
        begin
          {JP 26/06/07 : Vérifie si la sélection contient des lignes GC ou Tréso}
          lMenuItem.Enabled := EstVraiEcr(ATobRow) and
                               not (ATobRow.GetString('CLE_QUALIFORIGINE') = QUALIFTRESO);
          if FOkCreateModif then
            lMenuItem.Caption := 'Modifier l''é&criture'
          else
            lMenuItem.Caption := '&Consulter l''écriture';
        end
        else
          lMenuItem.OnClick := OnDblClickFListe;
        Continue;
      end;

      {$IFDEF COMPTA}
      if lMenuItem.Name = 'ECRCPTCPT' then
      begin // Passage d'une écriture compte à compte
        if vActivation then
          lMenuItem.Enabled := FOkEcr and EstVraiEcr(ATobRow)
        else
          lMenuItem.OnClick := PassageEcrSimplifie;
        Continue;
      end;
      {$ENDIf}

      {$IFDEF COMPTA}
      if lMenuItem.Name = 'SAISIEBOR' then
      begin // Saisie Bordereau
        if vActivation then
          lMenuItem.Enabled := FDroitEcritures and FOkCreateModif
        else
          lMenuItem.OnClick := SaisieBorClick;
        Continue;
      end;
      {$ENDIF}

      {$IFDEF COMPTA}
      if lMenuItem.Name = 'SAISIEPIECE' then
      begin // Saisie pièce
        if vActivation then
          lMenuItem.Enabled := FDroitEcritures and FOkCreateModif
        else
          lMenuItem.OnClick := SaisieEcrClick;
        Continue;
      end;
      {$ENDIF}

      {$IFDEF COMPTA}
      {$IFNDEF CCMP}
      if lMenuItem.Name = 'SAISIETRESO' then
      begin // Saisie Trésorerie
        if vActivation then
          lMenuItem.Enabled := lBoOkSaisieTreso
        else
          lMenuItem.OnClick := OnClickSaisieTreso;
        Continue;
      end;
      {$ENDIF}
      {$ENDIF}

    finally
      if vActivation then
      begin
        ActivationMenuItem(lMenuItem);
        if not lBoAuMoinsUn then
          lBoAuMoinsUn := lMenuItem.Visible;
      end;
    end;
  end;

  // GCO - 25/07/2007 - FQ 18489 + FQ 21520
  if vActivation and (not ABoMultiSelected) then
    BComptable.Visible := lBoAuMoinsUn;
end;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.OnPopUpSpecifique(Sender: TObject);
begin
  InitPopUpSpecifique( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/01/2007
Modifié le ... : 05/06/2007
Description .. : Fonction d'Init du PopUpOutil
Suite ........ : - LG - 05/06/2007 - FB 20548 -on ouvre les TIC ou TID que 
Suite ........ : pour els racines de compte 411
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.InitPopUpSpecifique(vActivation : Boolean);
var i,j : integer;
    lMenuItem    : TMenuItem;
    lBoAuMoinsUn : Boolean;
begin
  lBoAuMoinsUn := False;

  RecupEcrSelect ; // LG - 12/09/2007 - FB 21420
  for i := 0 to PopUpSpecifique.Items.Count - 1 do
  begin
    lMenuItem := PopUpSpecifique.Items[i];
    try
      lMenuItem.Enabled := False;
      if vActivation and ABoMultiSelected then Continue;

      if lMenuItem.Name = 'TRAITEMENTSPECIFIQUE' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          if lMenuItem.Items[j].Name = 'CPCONSAUXI' then (**)
          begin
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'COMPTERIB' then
          begin
           if vActivation then
             lMenuItem.Items[j].Enabled := False // GCO - 25/07/2007 - FQ 21096
           else
             lMenuItem.Items[j].OnClick := OnClickCompteRib;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'DETAILECHEANCE' then (**)
          begin
           if vActivation then
             lMenuItem.Items[j].Enabled := False
           else
             lMenuItem.Items[j].OnClick := nil;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'DETAILICC' then
          begin // Interet de comptes courants
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkICC
            else
              lMenuItem.Items[j].OnClick := OnClickDetailIcc;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'DETAILCAPITALICC' then
          begin // Détail du compte de capital ICC
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkICC
            else
              lMenuItem.Items[j].OnClick := OnClickDetailCapitalIcc;
            Continue;
          end;

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'GCD' then
          begin // Créer une créance dans GCD
            if vActivation then
              lMenuItem.Items[j].Enabled :=  VH^.OkModGCD and (Copy(E_General.Text,1,2) = '41')
            else
              lMenuItem.Items[j].OnClick := OnClickGCD;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'EMPRUNTCRE' then
          begin // Calcul de remboursement d'emprunt
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkCre
            else
              lMenuItem.Items[j].OnClick := OnClickEmpruntCRE;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

        end; // FOR
      end; // TRAITEMENTSPECIFIQUE

      if lMenuItem.Name = 'IMMO' then
      begin
        lMenuItem.Enabled := True;

        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          {$IFDEF AMORTISSEMENT}
          if lMenuItem.Items[j].Name = 'LISTEIMMO' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (FOkModImmo) and (FTobGene.GetString('G_NATUREGENE') = 'IMO')
            else
              lMenuItem.Items[j].OnClick := ListeImmoClick;
            Continue;
          end;
          {$ENDIF}

          {$IFDEF AMORTISSEMENT}
          if lMenuItem.Items[j].Name = 'FICHEIMMO' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (FOkModImmo) and
                                            (FTobGene.GetString('G_NATUREGENE') = 'IMO') and
                                            EstVraiEcr(ATobRow) and
                                            (ATobRow.GetString('CLE_IMMO') <> '')
            else
              lMenuItem.Items[j].OnClick := FicheImmoClick;
            Continue;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'SUPPRLIENIMMO' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (FOkModImmo) and
                                            (FTobGene.GetString('G_NATUREGENE') = 'IMO') and
                                            EstVraiEcr(ATobRow) and
                                            (ATobRow.GetString('CLE_IMMO') <> '')
            else
              lMenuItem.Items[j].OnClick := SupprimeLienImmoClick;
            Continue;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'CREATIONLIENIMMO' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (FOkModImmo) and
                                            (FTobGene.GetString('G_NATUREGENE') = 'IMO') and
                                            EstVraiEcr(ATobRow) and
                                            (ATobRow.GetString('CLE_IMMO') = '')
            else
              lMenuItem.Items[j].OnClick := CreationLienImmoClick;
            Continue;
          end;
          {$ENDIF}
        end; // FOR
      end; // IMMO

      if lMenuItem.Name = 'LETTRAGE' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := (lMenuItem.Items[j].IsLine);

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'LETMANUEL' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkLettrable and (not FFromSaisie)
            else
              lMenuItem.Items[j].OnClick := LettrageMClick;
            Continue;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'LETAUTO' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkLettrable and (not FFromSaisie)
            else
              lMenuItem.Items[j].OnClick := LettrageAClick;
            Continue;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'DELETPARTIEL' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkLettrable and (not FFromSaisie) and
                                            EstVraiEcr(ATobRow) and
                                            (ATobRow.GetString('CLE_LETTRAGE') <> '')
            else
              lMenuItem.Items[j].OnClick := DelettreClick;
            Continue;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'DELETEXOREF' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkLettrable and (not FFromSaisie)
            else
              lMenuItem.Items[j].OnClick := DelettreMvtExoRef;
            Continue;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'DELETTOTAL' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkLettrable and (not FFromSaisie)
            else
              lMenuItem.Items[j].OnClick := DeletTotal;
            Continue;
          end;
          {$ENDIF}

          // Passage en Justificatif de Solde
          if lMenuItem.Items[j].Name = 'PASSEJUSTIFSOLDE' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := CBJustifSolde.Enabled
            else
              lMenuItem.Items[j].OnClick := OnClickPasseJustifSolde;
            Continue;
          end;

          {$IFDEF COMPTA}
          // Justificatif de solde pour compte 'Lettrable'
          if lMenuItem.Items[j].Name = 'ETATJUSTIFSOLDE' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            FOkJustificatifSolde and
                                            (ATobFListe.Detail.Count > 0) and
                                            CBJustifSolde.Checked
            else
              lMenuItem.Items[j].OnClick := OnClickJustifSolde;
            Continue;
          end;
          {$ENDIF}
        end; // FOR
      end; // LETTRAGE

      if lMenuItem.Name = 'ANALYTIQUE' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'MVTANALYTIQ' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkAnalytique
            else
              lMenuItem.Items[j].OnClick := AnalytiquesClick;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'VENTILANALYTIQ1' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkAnalytique and FOkAnaSurAxe1
            else
              lMenuItem.Items[j].OnClick := VentilAnalytiques1Click;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'VENTILANALYTIQ2' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkAnalytique and FOkAnaSurAxe2
            else
              lMenuItem.Items[j].OnClick := VentilAnalytiques2Click;
            Continue;
          end;

          // GCO - 22/10/2007 - FQ 21683
          if lMenuItem.Items[j].Name = 'VENTILANALYTIQ3' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkAnalytique and FOkAnaSurAxe3
            else
              lMenuItem.Items[j].OnClick := VentilAnalytiques3Click;
            Continue;
          end;

          // GCO - 22/10/2007 - FQ 21683
          if lMenuItem.Items[j].Name = 'VENTILANALYTIQ4' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkAnalytique and FOkAnaSurAxe4
            else
              lMenuItem.Items[j].OnClick := VentilAnalytiques4Click;
            Continue;
          end;

          // GCO - 22/10/2007 - FQ 21683
          if lMenuItem.Items[j].Name = 'VENTILANALYTIQ5' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkAnalytique and FOkAnaSurAxe5
            else
              lMenuItem.Items[j].OnClick := VentilAnalytiques5Click;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'REPARTANALYTIQ' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkAnalytique
            else
              lMenuItem.Items[j].OnClick := RepartAnalytiquesClick;
            Continue;
          end;

          // Grand livre analytique sur l'axe 1
          if lMenuItem.Items[j].Name = 'GLANASURAXE1' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                             FOkAnalytique and FOkAnaSurAxe1
            else
              lMenuItem.Items[j].OnClick := OnClickGLAnaSurAxe1;
            Continue;
          end;

          // Grand livre analytique sur l'axe 2
          if lMenuItem.Items[j].Name = 'GLANASURAXE2' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe2
            else
              lMenuItem.Items[j].OnClick := OnClickGLAnaSurAxe2;
            Continue;
          end;

          // Grand livre analytique sur l'axe 3
          if lMenuItem.Items[j].Name = 'GLANASURAXE3' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe3
            else
              lMenuItem.Items[j].OnClick := OnClickGLAnaSurAxe3;
            Continue;
          end;

          // Grand livre analytique sur l'axe 4
          if lMenuItem.Items[j].Name = 'GLANASURAXE4' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe4
            else
               lMenuItem.Items[j].OnClick := OnClickGLAnaSurAxe4;
            Continue;
          end;

          // Grand livre analytique sur l'axe 5
          if lMenuItem.Items[j].Name = 'GLANASURAXE5' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe5
            else
              lMenuItem.Items[j].OnClick := OnClickGLAnaSurAxe5;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'GLGENEPARANA1' then
          begin // Grand livre général par anbalytique sur l'axe 1
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe1
            else
              lMenuItem.Items[j].OnClick := OnClickGLGeneParAna1;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'GLGENEPARANA2' then
          begin // Grand livre général par anbalytique sur l'axe 2
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe2
            else
              lMenuItem.Items[j].OnClick := OnClickGLGeneParAna2;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'GLGENEPARANA3' then
          begin // Grand livre général par anbalytique sur l'axe 3
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe3
            else
              lMenuItem.Items[j].OnClick := OnClickGLGeneParAna3;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'GLGENEPARANA4' then
          begin // Grand livre général par anbalytique sur l'axe 4
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe4
            else
              lMenuItem.Items[j].OnClick := OnClickGLGeneParAna4;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'GLGENEPARANA5' then
          begin // Grand livre général par anbalytique sur l'axe 5
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe5
            else
              lMenuItem.Items[j].OnClick := OnClickGLGeneParAna5;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'BALANASURAXE1' then
          begin //  Balance Analytique sur l'axe 1
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe1
            else
              lMenuItem.Items[j].OnClick := OnClickBalAnaSurAxe1;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'BALANASURAXE2' then
          begin //  Balance Analytique sur l'axe 2
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe2
            else
              lMenuItem.Items[j].OnClick := OnClickBalAnaSurAxe2;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'BALANASURAXE3' then
          begin //  Balance Analytique sur l'axe 3
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe3
            else
              lMenuItem.Items[j].OnClick := OnClickBalAnaSurAxe3;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'BALANASURAXE4' then
          begin //  Balance Analytique sur l'axe 4
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe4
            else
              lMenuItem.Items[j].OnClick := OnClickBalAnaSurAxe4;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'BALANASURAXE5' then
          begin //  Balance Analytique sur l'axe 5
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe5
            else
              lMenuItem.Items[j].OnClick := OnClickBalAnaSurAxe5;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'BALGENEPARANA1' then
          begin // Balance Générale par Analytique sur l'axe 1
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe1
            else
              lMenuItem.Items[j].OnClick := OnClickBalGeneParAna1;
          end;

          if lMenuItem.Items[j].Name = 'BALGENEPARANA2' then
          begin // Balance Générale par Analytique sur l'axe 2
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe2
            else
              lMenuItem.Items[j].OnClick := OnClickBalGeneParAna2;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'BALGENEPARANA3' then
          begin // Balance Générale par Analytique sur l'axe 3
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe3
            else
              lMenuItem.Items[j].OnClick := OnClickBalGeneParAna3;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'BALGENEPARANA4' then
          begin // Balance Générale par Analytique sur l'axe 4
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe4
            else
              lMenuItem.Items[j].OnClick := OnClickBalGeneParAna4;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'BALGENEPARANA5' then
          begin // Balance Générale par Analytique sur l'axe 5
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (ATobFListe.Detail.Count > 0) and
                                            FOkAnalytique and FOkAnaSurAxe5
            else
              lMenuItem.Items[j].OnClick := OnClickBalGeneParAna5;
            Continue;
          end;
          {$ENDIF}
        end; // FOR
      end; // ANALYTIQUE

      if lMenuItem.Name = 'POINTAGE' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'POINTAGEMANUEL' then
          begin // Pointage Manuel
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkPointable
            else
              lMenuItem.Items[j].OnClick := PointageGClick;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          // Pointage sur la référence en cours
          if lMenuItem.Items[j].Name = 'POINTAGEENCOURS' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkPointable
            else
              lMenuItem.Items[j].OnClick := PointageEnCoursClick;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          {$IFDEF COMPTA} 
          if lMenuItem.Items[j].Name = 'ETATJUSTIFSOLDEBQE' then
          begin // Etat de Justificatif de Solde Banquaire
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkEdition and
                                            FOkPointable and
                                            (not FOkPointageJal) and
                                            (ATobFListe.Detail.Count > 0) and
                                            (FTobGene.GetString('G_NATUREGENE') = 'BQE')
            else
              lMenuItem.Items[j].OnClick := OnClickJustifSoldeBQE;
            Continue;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'ETATRAPPRO' then
          begin // Etat de rapprochement pour compte 'Pointable' et nature Banque
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            FOkEdition and FOkPointable and
                                            (not FOkPointageJal) and
                                            (ATobFListe.Detail.Count > 0)
            else
              lMenuItem.Items[j].OnClick := OnClickEtatRapprochement;
            Continue;
          end;
          {$ENDIF}

        end; // FOR
      end; // POINTAGE

    finally
      if vActivation then
      begin
        ActivationMenuItem(lMenuItem);
        if not lBoAuMoinsUn then lBoAuMoinsUn := lMenuItem.Visible;
      end;
    end;
  end;

  // GCO - 25/07/2007 - FQ 18489 + FQ 21520
  if vActivation and (not ABoMultiSelected) then
    BSpecifique.Visible := lBoAuMoinsUn;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnPopUpAideRevision(Sender: TObject);
begin
  InitPopUpAideRevision( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/01/2007
Modifié le ... : 06/04/2007
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.InitPopUpAideRevision(vActivation : Boolean);
var i,j : integer;
    lMenuItem : TMenuItem;
{$IFDEF COMPTA}
    lStGen : string;
{$ENDIF}
begin
  // GCO - 30/08/2007
  RecupEcrSelect;

  for i := 0 to PopUpAideRevision.Items.Count - 1 do
  begin
    lMenuItem := PopUpAideRevision.Items[i];
    try
      lMenuItem.Enabled := False;
      if vActivation and ABoMultiSelected then Continue;

      if lMenuItem.Name = 'OUTILREVISION' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'NOTETRAVAIL' then
          begin // Note de travail
            if vActivation then
              lMenuItem.Items[j].Enabled := (E_General.Text <> '') and VH^.OkModRIC and JaiLeRoleCompta( rcReviseur )
            else
              lMenuItem.Items[j].OnClick := OnClickNoteTravail;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'TABLEAUVARIATION' then
          begin // Tableau de variation
            if vActivation then
              lMenuItem.Items[j].Enabled := (E_General.Text <> '') and VH^.OkModRIC and JaiLeRoleCompta( rcReviseur )
            else
              lMenuItem.Items[j].OnClick := OnClickTableauVariation;
            Continue;
           end;
          {$ENDIF}
          {$ENDIF}

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'RECHERCHEECR' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := True (**)
            else
              lMenuItem.Items[j].OnClick := RechercheEcrituresClick;
            Continue;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          // Feuille Excel
          if lMenuItem.Items[j].Name = 'FEUILLEEXCEL' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := VH^.OkModRIC and JaiLeRoleCompta( rcReviseur )
            else
              lMenuItem.Items[j].OnClick := OnClickFeuilleExcel;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'CUTOFF' then
          begin
            lStGen := Copy(FTobGene.GetValue('G_GENERAL'),1,3) ;
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkEcr and EstVraiEcr(FTobEcrSelect) and FDroitEcritures and
                                            ( ( FTobGene.GetString('G_CUTOFF') = 'X' ) or
                                            ( ( lStGen = '486' ) or ( lStGen = '487' ) ) )
            else
              lMenuItem.Items[j].OnClick := OnClickCutOff ;
            Continue;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'CONTROLETVA' then
          begin
            lStGen := Copy(FTobGene.GetValue('G_GENERAL'),1,4) ;
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkEcr and ( ( lStGen = '4456' ) or ( lStGen = '4457' ) ) 
            else
              lMenuItem.Items[j].OnClick := OnClickControleTva;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          if lMenuItem.Items[j].Name = 'CALCULICC' then
          begin // Calcul des ICC
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkIcc
            else
              lMenuItem.Items[j].OnClick := OnClickCalculIcc;
            Continue;
          end;

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'DRF' then
          begin // Détermination du résultat fiscal
            if vActivation then
              lMenuItem.Items[j].Enabled := (FExoDate.Code = VH^.EnCours.Code) and
                                            (VH^.OkModDRF) and 
                                            (OkLanceDRF <> '')
            else
              lMenuItem.Items[j].OnClick := OnClickDRF;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'JUSTIFGCD' then
          begin // Gestion des créances douteuses
            if vActivation then
              lMenuItem.Items[j].Enabled := VH^.OkModGCD and
                       ( E_General.Text = GetParamSocSecur('SO_CPGCDGENERAL','') )   or
                       ( E_General.Text = GetParamSocSecur('SO_CPGCDPERTE','') )     or
                       ( E_General.Text = GetParamSocSecur('SO_CPGCDPROVISION','') ) or
                       ( E_General.Text = GetParamSocSecur('SO_CPGCDDOTPROV','') )   or
                       ( E_General.Text = GetParamSocSecur('SO_CPGCDREPRISE','') )
            else
              lMenuItem.Items[j].OnClick := OnClickTableauGCD ;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'RAPPROCOMPTACRE' then
          begin // Rapprochement Compta / CRE
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkCre
            else
             lMenuItem.Items[j].OnClick := OnClickRappECRCRE;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          {$IFDEF AMORTISSEMENT}
          if lMenuItem.Items[j].Name = 'RAPPROCOMPTAIMMO' then
          begin // Rapprochement Comptabilité / Immobilisation
            if vActivation then
              lMenuItem.Items[j].Enabled := (FOkModImmo) and (FTobGene.GetString('G_NATUREGENE') = 'IMO')
            else
              lMenuItem.Items[j].OnClick := OnClickRapproComptaImmo;
            Continue;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'RAPPROCOMPTAGCD' then
          begin // Rapprochement Comptabilité / GCD
            if vActivation then
              lMenuItem.Items[j].Enabled := VH^.OkModGCD and
                       ( E_General.Text = GetParamSocSecur('SO_CPGCDGENERAL','') )   or
                       ( E_General.Text = GetParamSocSecur('SO_CPGCDPERTE','') )     or
                       ( E_General.Text = GetParamSocSecur('SO_CPGCDPROVISION','') ) or
                       ( E_General.Text = GetParamSocSecur('SO_CPGCDDOTPROV','') )   or
                       ( E_General.Text = GetParamSocSecur('SO_CPGCDREPRISE','') )   
            else
              lMenuItem.Items[j].OnClick := OnClickRapproComptaGCD;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

        end; // FOR
      end;  // OUTILREVISION

      if lMenuItem.Name = 'LIASSEFISCALE' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'INFOLIASSE' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (FStLiasseDossier <> '') and
                                            (E_General.Text <> '') and
                                            (not FBoConfidentielGene)
            else
              lMenuItem.Items[j].OnClick := OnClickInfoLiasse;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'TOTALCPTLIASSE' then
          begin // Montant affecté dans la liasse fiscale
            if vActivation then
              lMenuItem.Items[j].Enabled := (FStLiasseDossier <> '') and
                                            (E_General.Text <> '') and
                                            (not FBoConfidentielGene)
            else
              lMenuItem.Items[j].OnClick := OnClickTotalCptLiasse;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'DONNEESCOMP' then
          begin // Emplacement dans la liasse fiscale
            if vActivation then
              lMenuItem.Items[j].Enabled := (FStLiasseDossier <> '') and
                                            (E_General.Text <> '') and
                                            (not FBoConfidentielGene)
                                            and CPPresenceEtafi
            else
              lMenuItem.Items[j].OnClick := OnClickDonneesComp;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}
        end; // FOR
      end; // LIASSE

      if lMenuItem.Name = 'ANALYSE' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'CUMULCPTE' then
          begin // Cumul du compte général
            if vActivation then
              lMenuItem.Items[j].Enabled := (E_General.Text <> '') and (not FBoConfidentielGene)
            else
              lMenuItem.Items[j].OnClick := CumulsGENEClick;
            Continue;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'CUMULAUXILIAIRE' then
          begin // Cumul du compte auxiliaire
            if vActivation then
              lMenuItem.Items[j].Enabled := (E_Auxiliaire.Text <> '') and (not FBoConfidentielAuxi)
            else
              lMenuItem.Items[j].OnClick := CumulsAUXClick;
            Continue;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'HISTOCPTE' then
          begin // Historique du compte général
            if vActivation then
              lMenuItem.Items[j].Enabled := (E_General.Text <> '') and (not FBoConfidentielGene)
            else
              lMenuItem.Items[j].OnClick := HistoCpteClick;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          if lMenuItem.Items[j].Name = 'ANALYSCOMP' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (E_General.Text <> '') and
                                            (not FBoConfidentielGene) and
                                            (not FBoConfidentielAuxi)
            else
              lMenuItem.Items[j].OnClick := ANALREVISClick;
            Continue;
          end;

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'DOCGUID' then
          begin
            if vActivation then 
              lMenuItem.Items[j].Enabled := (CtxPcl in V_Pgi.PgiContexte) and
                                            (EstVraiEcr(FTobEcrSelect)) and
                                            (ATobRow.GetString('CLE_DOCGUID') <> '')
            else
              lMenuItem.Items[j].OnClick := OnClickDocumentGed;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'INFOCOMP' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := EstVraiEcr(FTobEcrSelect)
            else
              lMenuItem.Items[j].OnClick := OnClickInfosComp;
            Continue;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'MEMOMILLESIME1' then
          begin // Mémo commentaire millésimé
            if vActivation then // GCO - 04/05/2007 - FQ 20244
              lMenuItem.Items[j].Enabled := (E_General.Text <> '') and (not VH^.OkModRIC)
            else
              lMenuItem.Items[j].OnClick := OnClickMemoMillesime;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'MEMOCOMPTE1' then
          begin // Mémo compte général
            if vActivation then // GCO - 04/05/2007 - FQ 20244
              lMenuItem.Items[j].Enabled := (E_General.Text <> '') and (not VH^.OkModRIC)
            else
              lMenuItem.Items[j].OnClick := OnClickMemoCompte;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}
        end; // FOR
      end; // ANALYSE

      {$IFDEF COMPTA}
      {$IFNDEF CCMP}
      if lMenuItem.Name = 'VISACOMPTE' then
      begin // Visa sur le compte
        if vActivation then
        begin //FOkRevision and
          lMenuItem.Enabled := (E_General.Text <> '') and (FExoDate.Code = VH^.Encours.Code);
          if lMenuItem.Enabled then
          begin
            if FTobGene.GetString('G_VISAREVISION') = 'X' then
            begin
              lMenuItem.Caption := TraduireMemoire('Enlever le &visa du compte');
              lMenuItem.Enabled := AutoriseSuppresionVisaRevision(E_General.Text);
            end
            else
              lMenuItem.Caption := TraduireMemoire('&Viser le compte');
          end;
        end
        else
          lMenuItem.OnClick := OnClickVisaCompte;
        Continue;
      end;
      {$ENDIF}
      {$ENDIF}

    finally // GCO - 08/02/2007
      if vActivation then ActivationMenuItem(lMenuItem);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.OnPopUpRevInteg(Sender: TObject);
begin
  InitPopUpRevInteg( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/01/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.InitPopUpRevInteg(vActivation: Boolean);
var i,j : integer;
    lMenuItem : TMenuItem;
begin
  // GCO - 30/08/2007 
  RecupEcrSelect;

  for i := 0 to PopUpRevInteg.Items.Count - 1 do
  begin
    lMenuItem := PopUpRevInteg.Items[i];
    try
      lMenuItem.Enabled := False;
      if vActivation and ABoMultiSelected then Continue;

      {$IFDEF COMPTA}
      {$IFNDEF CCMP}
      if lMenuItem.Name = 'PROGTRAVAIL' then
      begin
        if vActivation then
          lMenuItem.Enabled := VH^.OkModRIC and FOkRevision and
                               ((JaiLeRoleCompta(rcSuperviseur)) or
                                (JaiLeRoleCompta(rcReviseur) and
                                 not JaileRoleCompta(rcSuperviseur))
                               )
        else
          lMenuItem.OnClick := OnClickProgrammeTravail;
        Continue;
      end;
      {$ENDIF}
      {$ENDIF}

      if lMenuItem.Name = 'DOCTRAVAUX' then
      begin
        lMenuItem.Enabled := VH^.OkModRIC;

        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'MEMOCYCLE' then
          begin // Mémo cycle
            if vActivation then
              lMenuItem.Items[j].Enabled := JaiLeRoleCompta( RcReviseur ) and
                                            FOkRevision and
                                            (FTobGene.GetString('G_CYCLEREVISION') <> '')
            else
              lMenuItem.Items[j].OnClick := OnClickMemoCycle;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'MEMOOBJECTIF' then
          begin // Mémo objectif de révision
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl In V_Pgi.PgiContexte) and
                                            JaiLeRoleCompta( RcReviseur ) and
                                            FOkRevision and
                                            (FTobGene.GetString('G_CYCLEREVISION') <> '')
            else
              lMenuItem.Items[j].OnClick := OnClickMemoObjectif;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'MEMOSYNTHESE' then
          begin // Mémo synthèse du cycle
            if vActivation then
              lMenuItem.Items[j].Enabled := JaiLeRoleCompta( RcReviseur ) and
                                            FOkRevision and
                                            (FTobGene.GetString('G_CYCLEREVISION') <> '')
            else
              lMenuItem.Items[j].OnClick := OnClickMemoSynthese;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'MEMOMILLESIME2' then
          begin // Mémo commantire millésimé
            if vActivation then
              lMenuItem.Items[j].Enabled := (E_General.Text <> '')
            else
              lMenuItem.Items[j].OnClick := OnClickMemoMillesime;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'MEMOCOMPTE2' then
          begin // Mémo compte général
            if vActivation then
              lMenuItem.Items[j].Enabled := (E_General.Text <> '')
            else
              lMenuItem.Items[j].OnClick := OnClickMemoCompte;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}

        end; // FOR
      end; // DOCTRAVAUX

      {$IFDEF COMPTA}
      {$IFNDEF CCMP}
      if lMenuItem.Name = 'SUPERVISIONTRAVAUX' then
      begin
        // GCO - 05/09/2007 - FQ 21363
        lMenuItem.Enabled := (CtxPcl In V_Pgi.PgiContexte) and VH^.OkModRic and
                             FOkRevision and JaiLeRoleCompta( rcSuperviseur );

        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := lMenuItem.Enabled; // GCO - 13/09/2007 - FQ 21363
          if lMenuItem.Items[j].Name = 'APG' then
          begin
            if not vActivation then // Appréciation générale du dossier
              lMenuItem.Items[j].OnClick := OnClickAPG;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'SCY' then
          begin
            if not vActivation then // Synthèse des cycles
              lMenuItem.Items[j].OnClick := OnClickSCY;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'EXP' then
          begin
            if not vActivation then // Attestation Expert
              lMenuItem.Items[j].OnClick := OnClickEXP;
            Continue;
          end;
        end;
      end;
      {$ENDIF}
      {$ENDIF}

      if lMenuItem.Name = 'DOSSIERCLIENT' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          {$IFDEF COMPTA}
          {$IFNDEF CCMP}
          if lMenuItem.Items[j].Name = 'DOSPERMANENT' then
          begin // Dossier permanent
            if vActivation then
              lMenuItem.Items[j].Enabled := VH^.OkModGed
            else
              lMenuItem.Items[j].OnClick := OnClickDocPermanent;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'DOSANNUEL' then
          begin // Dossier annuel
            if vActivation then
              lMenuItem.Items[j].Enabled := VH^.OkModGed
            else
              lMenuItem.Items[j].OnClick := OnClickDocAnnuel;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'RECHERCHEDOC' then
          begin // Recherche de document
            if vActivation then
              lMenuItem.Items[j].Enabled := VH^.OkModGed
            else
              lMenuItem.Items[j].OnClick := OnClickRechercheDoc;
            Continue;
          end;
          {$ENDIF}
          {$ENDIF}
        end; // FOR
      end; // DOSSIERCLIENT

    // GCO - 10/10/2005
    finally
      if vActivation then ActivationMenuItem(lMenuItem);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnPopUpParametre(Sender: TObject);
begin
  InitPopUpParametre( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/01/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.InitPopUpParametre(vActivation: Boolean);
var i,j : integer;
    lMenuItem : TMenuItem;
begin
  for i := 0 to PopUpParametre.Items.Count - 1 do
  begin
    lMenuItem := PopUpParametre.Items[i];
    try
      lMenuItem.Enabled := False;
      if (vActivation and ABoMultiSelected) then Continue;

      if lMenuItem.Name = 'PARAMETRES' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False; 

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'PARAMGENERAL' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (E_General.Text <> '') and (not FBoConfidentielGene)
            else
              lMenuItem.Items[j].OnClick := OnClickCompteGeneral;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'PARAMAUXILIAIRE' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (E_Auxiliaire.Text <> '') and (not FBoConfidentielAuxi)
            else
              lMenuItem.Items[j].OnClick := OnClickCompteAuxiliaire;
          end;
          {$ENDIF}

          if lMenuItem.Items[j].Name = 'RIBBQE' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkAfficheRIB
            else
              lMenuItem.Items[j].OnClick := OnClickCompteRib;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'TABLECHANCELLERIE' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkTableChancellerie
            else
              lMenuItem.Items[j].OnClick := OnClickTableChancellerie;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'LETGESTION' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkLettrable and (not FFromSaisie)
            else
              lMenuItem.Items[j].OnClick := OnClickLetGestion;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'LETPARAM' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := FOkLettrable and (not FFromSaisie)
            else
              lMenuItem.Items[j].OnClick := OnClickLetParam;
            Continue;
          end;
        end;
      end;

    finally // GCO - 24/01/2007
      if vActivation then ActivationMenuItem(lMenuItem);
    end;
  end; // For i := 0
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/01/2007
Modifié le ... : 20/07/2007
Description .. : - LG - FB 19707 - 19/02/2007 - on ne peut selectionner une 
Suite ........ : ecriture ds 
Suite ........ : un exo clo
Suite ........ : - LG - 03/07/2007 - FB 20448 - on bloque les comptes 
Suite ........ : ventilables
Suite ........ : - LG - 20/07/2007 - FB 21065 - meme controle en multi 
Suite ........ : selection qu'en monon selection
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.InitPopUpMultiselect(vActivation: Boolean);
var i         : integer;
    lMenuItem : TMenuItem; 
{$IFDEF COMPTA}
    lBoExoClo : Boolean ;
    lTobEcr   : TOB ;
{$ENDIF}
begin
  // Chargement en TOb des écritures sélectionnées.
  // ChargeFTobListeEcr;

  for i := 0 to PopUpMultiSelect.Items.Count - 1 do
  begin
    lMenuItem := PopUpMultiSelect.Items[i];
    try
      lMenuItem.Enabled := False;
      if vActivation and (not ABoMultiSelected) then Continue;

    {$IFDEF COMPTA}
      if lMenuItem.Name = 'MODIFSERIE' then
      begin
        if not vActivation then
          lMenuItem.OnClick := OnClickModifSerie
        else
          lMenuItem.Enabled := ABoMultiSelected and
                                               ((not BoGesComTTC) or EstSpecif('SIC')) and
                                               (ATobFListe.Detail.Count > 0);
      end;
    {$ENDIF}

    {$IFDEF COMPTA}
      if lMenuItem.Name = 'RECLASS' then
      begin
        if not vActivation then
         lMenuItem.OnClick := OnClickReclass
        else
         begin
          lTobEcr   := GetO(FListe);
          lBoExoClo := false ;
          if lTobEcr <> nil then
           lBoExoClo := ctxExercice.EstExoClos(lTOBECr.GetString('CLE_EXERCICE')) ;
          lMenuItem.Enabled := ABoMultiSelected and
                                               ( not lBoExoClo ) and
                                              (not BoGesComTTC) and
                                              (ATobFListe.Detail.Count > 0) and
                                              (FTobGene.GetString('G_VENTILABLE') = '-') and
                                              (FTobGene.GetString('G_VISAREVISION') = '-');
         end ;
      end;

      if lMenuItem.Name = 'LETTSIMPL' then
      begin
        if not vActivation then
         lMenuItem.OnClick := PassageEcrSimplifie
        else
          lMenuItem.Enabled := ABoMultiSelected and (ATobFListe.Detail.Count > 0);
      end;

      {$IFNDEF CCMP}
      // Gestion des créances douteuses
      if lMenuItem.Name = 'GCDM' then
      begin
        if vActivation then
          lMenuItem.Enabled := VH^.OkModGCD and (Copy(E_General.Text,1,2) = '41') // FB 21065 - LG - 20/07/2007 
        else
          lMenuItem.OnClick := OnClickGCD ;
      end;
      {$ENDIF}


      if lMenuItem.Name = 'COPIENOTETRAVAIL' then
      begin
        if vActivation then
          lMenuItem.Enabled := True
        else
          lMenuItem.OnClick := OnClickEcrVersClipboard;
      end;
    {$ENDIF}

    {$IFDEF COMPTA}
      //if lMenuItem.Name = 'LETTRAGE' then
      if lMenuItem.Name = 'LETRAGEMULTI' then   // CA - 02/05/2007 - Pb lors de la refonte du F11 (FQ 20203)
      begin
        if not vActivation then
          lMenuItem.OnClick :=  LETTRAGEMClick
        else
          lMenuItem.Enabled := FOkFctLettrage and
                                               (not FFromSaisie) and
                                               ABoMultiSelected and
                                               (ATobFListe.Detail.Count > 0) and
                                               (FTobGene.GetValue('G_VISAREVISION') = '-');
      end;
    {$ENDIF}

      {$IFDEF COMPTA}
      if lMenuItem.Name = 'AFFECTEANA' then
      begin
        if not vActivation then
          lMenuItem.OnClick := OnClickAffecteAnaEnSerie
        else
          lMenuItem.Enabled := (not FFromSaisie) and
                                               ABoMultiSelected and
                                               (not FOkAnaCroisaxe) and
                                               (ATobFListe.Detail.Count > 0) and
                                               (FTobGene.GetValue('G_VENTILABLE') = 'X');
      end;
      {$ENDIF}

    finally// GCO - 10/10/2005
      if vActivation then ActivationMenuItem(lMenuItem);
    end;  
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/01/2007
Modifié le ... :   /  /
Description .. : GCO - 29/08/2007 - FQ 18489
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.InitPopUpEdition(vActivation: Boolean);
var i : integer;
    lMenuItem : TMenuItem;
    lBoAuMoinsUn : Boolean;
begin
  lBoAuMoinsUn := False;

  for i := 0 to PopUpEdition.Items.Count - 1 do
  begin
    lMenuItem := PopUpEdition.Items[i];
    try
      lMenuItem.Enabled := False;

      // GCO - 24/11/2004 - FQ 13007 Vérouillage des éditions si confidentialité
      if not FOkEdition then Continue;

      // GCO - 23/10/2007 - FQ 21677
      {$IFDEF COMPTA}
      // Justificatif de solde pour compte 'Lettrable' PGE
      if lMenuItem.Name = 'EJUSTIFSOLDE' then
      begin
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               FOkJustificatifSolde and
                               (ATobFListe.Detail.Count > 0) and
                               CBJustifSolde.Checked
        else
          lMenuItem.OnClick := OnClickJustifSolde;
        Continue;
      end;
      {$ENDIF}

      {$IFDEF COMPTA}
      if lMenuItem.Name = 'EETATRAPPRO' then
      begin // Etat de rapprochement pour compte 'Pointable' et nature Banque PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               FOkEdition and FOkPointable and
                               (not FOkPointageJal) and
                               (ATobFListe.Detail.Count > 0)
        else
          lMenuItem.OnClick := OnClickEtatRapprochement;
        Continue;
      end;
      {$ENDIF}

      {$IFDEF COMPTA}
      // Grand Livre Général
      if lMenuItem.Name = 'GLGENE' then
      begin
        if vActivation then
          lMenuItem.Enabled := (FTobGene.GetString('G_GENERAL') <> '') and
                               (ATobFListe.Detail.Count > 0)
        else
          lMenuItem.OnClick := OnClickGLGene;
        Continue;  
       end;
      {$ENDIF}

      {$IFDEF COMPTA}
      // Grand Livre Général par Auxiliaire
      if lMenuItem.Name = 'GLGENEPARAUXI' then
      begin
        if vActivation then
          lMenuItem.Enabled := (ATobFListe.Detail.Count > 0) and
                                                 (FTobTiers.GetString('T_AUXILIAIRE') = '') and
                                                 (FTobGene.GetString('G_COLLECTIF') = 'X')
        else
          lMenuItem.OnClick := OnClickGLGeneParAuxi;
        Continue;
      end;
      {$ENDIF}

      {$IFDEF COMPTA}
      // Grand-Livre Général par Quantité
      if lMenuItem.Name = 'GLGENEPARQTE' then
      begin
        if vActivation then
          lMenuItem.Enabled := FOkGereQte AND
                               (ATobFListe.Detail.Count > 0) AND
                               ((FTobGene.GetString('G_QUALIFQTE1') <> '') OR
                               (FTobGene.GetString('G_QUALIFQTE2') <> ''))
        else
          lMenuItem.OnClick := OnClickGLGeneParQte;
        Continue;  
      end;

      if lMenuItem.Name = 'EGLGENEPARANA1' then
      begin // Grand livre général par anbalytique sur l'axe 1 PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe1
        else
          lMenuItem.OnClick := OnClickGLGeneParAna1;
        Continue;
      end;

      if lMenuItem.Name = 'EGLGENEPARANA2' then
      begin // Grand livre général par anbalytique sur l'axe 2 PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe2
        else
          lMenuItem.OnClick := OnClickGLGeneParAna2;
        Continue;
      end;

      if lMenuItem.Name = 'EGLGENEPARANA3' then
      begin // Grand livre général par anbalytique sur l'axe 3 PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe3
        else
          lMenuItem.OnClick := OnClickGLGeneParAna3;
        Continue;
      end;

      if lMenuItem.Name = 'EGLGENEPARANA4' then
      begin // Grand livre général par anbalytique sur l'axe 4 PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe4
        else
          lMenuItem.OnClick := OnClickGLGeneParAna4;
        Continue;
      end;

      if lMenuItem.Name = 'EGLGENEPARANA5' then
      begin // Grand livre général par anbalytique sur l'axe 5 PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe5
        else
          lMenuItem.OnClick := OnClickGLGeneParAna5;
        Continue;
      end;

      // Grand-Livre Auxiliaire
      if lMenuItem.Name = 'GLAUXI' then
      begin
        if vActivation then
          lMenuItem.Enabled := (FTobTiers.GetString('T_AUXILIAIRE') <> '') and (ATobFListe.Detail.Count > 0)
        else
          lMenuItem.OnClick := OnClickGLAuxi;
        Continue;
      end;
      {$ENDIF}

      {$IFDEF COMPTA}
      // Grand-Livre Auxiliaire par Général
      if lMenuItem.Name = 'GLAUXIPARGENE' then
      begin
        if vActivation then
          lMenuItem.Enabled := (FTobTiers.GetString('T_AUXILIAIRE') <> '') and
                               (FTobGene.GetString('G_GENERAL') = '') and
                               (ATobFListe.Detail.Count > 0)
        else
          lMenuItem.OnClick := OnClickGLAuxiParGene;
      end;
      {$ENDIF}

      {$IFDEF COMPTA}
      // Grand-Livre Auxiliaire pour Général
      if lMenuItem.Name = 'GLAUXIPOURGENE' then
      begin
        if vActivation then
          lMenuItem.Enabled := (FTobTiers.GetString('T_AUXILIAIRE') <> '') and
                               (FTobGene.GetString('G_GENERAL') <> '') and
                               (ATobFListe.Detail.Count > 0)
        else
          lMenuItem.OnClick := OnClickGLAuxiPourGene;
        Continue;  
      end;

      // Grand livre analytique sur l'axe 1 PGE
      if lMenuItem.Name = 'EGLANASURAXE1' then
      begin
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe1
        else
          lMenuItem.OnClick := OnClickGLAnaSurAxe1;
        Continue;
      end;

      // Grand livre analytique sur l'axe 2 PGE
      if lMenuItem.Name = 'EGLANASURAXE2' then
      begin
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe2
        else
          lMenuItem.OnClick := OnClickGLAnaSurAxe2;
        Continue;
      end;

      // Grand livre analytique sur l'axe 3 PGE
      if lMenuItem.Name = 'EGLANASURAXE3' then
      begin
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe3
        else
          lMenuItem.OnClick := OnClickGLAnaSurAxe3;
        Continue;
      end;

      // Grand livre analytique sur l'axe 4 PGE
      if lMenuItem.Name = 'EGLANASURAXE4' then
      begin
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe4
        else
           lMenuItem.OnClick := OnClickGLAnaSurAxe4;
        Continue;
      end;

      // Grand livre analytique sur l'axe 5 PGE
      if lMenuItem.Name = 'EGLANASURAXE5' then
      begin
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe5
        else
          lMenuItem.OnClick := OnClickGLAnaSurAxe5;
        Continue;
      end;

      // Balance Générale 
      if lMenuItem.Name = 'BALGENE' then
      begin
        if not vActivation then
          lMenuItem.OnClick := OnClickBalGene
        else
        begin
          lMenuItem.Enabled := (ATobFListe.Detail.Count > 0) and
                               (FTobGene.GetString('G_GENERAL') <> '') and
                               (FTobTiers.GetString('T_AUXILIAIRE') = '') and
                               (E_Exercice.ItemIndex <> 0);
        end;
      end;

      // Balance Auxiliaire
      if lMenuItem.Name = 'BALAUXI' then
      begin
        if not vActivation then
          lMenuItem.OnClick := OnClickBalAuxi
        else
        begin
          lMenuItem.Enabled := (ATobFListe.Detail.Count > 0) and
                               (FTobTiers.GetString('T_AUXILIAIRE') <> '') and
                               (FTobGene.GetString('G_GENERAL') = '') and
                               (E_Exercice.ItemIndex <> 0);
        end;
      end;

      // Balance Général par auxiliaire
      if lMenuItem.Name = 'BALGENEPARAUXI' then
      begin
        if not vActivation then
          lMenuItem.OnClick := OnClickBalGeneParAuxi
        else
        begin
          lMenuItem.Enabled := (FTobTiers.GetString('T_AUXILIAIRE') = '') and
                               (FTobGene.GetString('G_COLLECTIF') = 'X') and
                               (ATobFListe.Detail.Count > 0) and
                               (E_Exercice.ItemIndex <> 0);
        end;
      end;

      // Balance Auxiliaire par Général
      if lMenuItem.Name = 'BALAUXIPARGENE' then
      begin
        if not vActivation then
          lMenuItem.OnClick := OnClickBalAuxiParGene
        else
        begin
          lMenuItem.Enabled := (FTobTiers.GetString('T_AUXILIAIRE') <> '') and
                               (FTobGene.GetString('G_GENERAL') = '') and
                               (ATobFListe.Detail.Count > 0) and
                               (E_Exercice.ItemIndex <> 0);
        end;
      end;

      if lMenuItem.Name = 'EBALGENEPARANA1' then
      begin // Balance Générale par Analytique sur l'axe 1 PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe1
        else
          lMenuItem.OnClick := OnClickBalGeneParAna1;
      end;

      if lMenuItem.Name = 'EBALGENEPARANA2' then
      begin // Balance Générale par Analytique sur l'axe 2 PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe2
        else
          lMenuItem.OnClick := OnClickBalGeneParAna2;
        Continue;
      end;

      if lMenuItem.Name = 'EBALGENEPARANA3' then
      begin // Balance Générale par Analytique sur l'axe 3 PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe3
        else
          lMenuItem.OnClick := OnClickBalGeneParAna3;
        Continue;
      end;

      if lMenuItem.Name = 'EBALGENEPARANA4' then
      begin // Balance Générale par Analytique sur l'axe 4 PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe4
        else
          lMenuItem.OnClick := OnClickBalGeneParAna4;
        Continue;
      end;

      if lMenuItem.Name = 'EBALGENEPARANA5' then
      begin // Balance Générale par Analytique sur l'axe 5 PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe5
        else
          lMenuItem.OnClick := OnClickBalGeneParAna5;
        Continue;
      end;

      if lMenuItem.Name = 'EBALANASURAXE1' then
      begin //  Balance Analytique sur l'axe 1 PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                              (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe1
        else
          lMenuItem.OnClick := OnClickBalAnaSurAxe1;
        Continue;
      end;

      if lMenuItem.Name = 'EBALANASURAXE2' then
      begin //  Balance Analytique sur l'axe 2 PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe2
        else
          lMenuItem.OnClick := OnClickBalAnaSurAxe2;
        Continue;
      end;

      if lMenuItem.Name = 'EBALANASURAXE3' then
      begin //  Balance Analytique sur l'axe 3 PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe3
        else
          lMenuItem.OnClick := OnClickBalAnaSurAxe3;
        Continue;
      end;

      if lMenuItem.Name = 'EBALANASURAXE4' then
      begin //  Balance Analytique sur l'axe 4 PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe4
        else
          lMenuItem.OnClick := OnClickBalAnaSurAxe4;
        Continue;
      end;

      if lMenuItem.Name = 'EBALANASURAXE5' then
      begin //  Balance Analytique sur l'axe 5 PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and
                               (ATobFListe.Detail.Count > 0) and
                               FOkAnalytique and FOkAnaSurAxe5
        else
          lMenuItem.OnClick := OnClickBalAnaSurAxe5;
        Continue;
      end;
      {$ENDIF}

    finally
      if vActivation then
      begin
        ActivationMenuItem(lMenuItem);
        if not lBoAuMoinsUn then
          lBoAuMoinsUn := lMenuItem.Visible;
      end;  
    end;
  end;

  // GCO - 25/07/2007 - FQ 18489 + FQ 21520
  if vActivation and (not ABoMultiSelected) then
    BEdition.Visible := lBoAuMoinsUn;
end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
////////////////////// FONCTIONS DES DIFFERENTS MENUS //////////////////////////
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
{$IFDEF COMPTA}
procedure TOF_CONSECR.EtudieModeLettrage(var R: RLETTR);
var
  i: integer;
begin
  R.LettrageDevise := False;
  if R.DeviseMvt = V_PGI.DevisePivot then
  begin
    R.CritDev := V_PGI.DevisePivot;
  end
  else
  begin
    if R.DeviseMvt = R.CritDev then
    begin
      {Paquet en devise --> Lettrage devise}
      R.LettrageDevise := True;
    end
    else
    begin
      {Paquet en devise (non explicite) --> poser la question}
      i := PgiAsk('Les mouvements sont en devise. Voulez-vous lettrer en devise ?', 'Lettrage manuel');
      R.LettrageDevise := (i = mrYes);
      if R.LettrageDevise then
        R.CritDev := R.DeviseMvt
      else
        R.Distinguer := False;
    end;
  end;
end;
{$ENDIF}

////////////////////////////////////////é////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/09/2002
Modifié le ... : 12/09/2007
Description .. : LettrageManuel d' un compte
Suite ........ : 
Suite ........ : - LG - 30/06/2003 - on tient compte du parametre de la
Suite ........ : chose Etatlettrage pour ouvrir le lettrage manuel
Suite ........ : - LG - 16/06/2004 - modification des parametre de
Suite ........ : lettragemanuel pour le changement ds la fenetre de l'auxi
Suite ........ : - LG - 30/06/2004 - FB 12989 12960 - on utilise la meme
Suite ........ : requet pour la consultation et le lettrage
Suite ........ : - LG - 10/09/2004 - Si j'utilise l'onglet "LETTRAGE" et que 
Suite ........ : je réduise la liste aux seules écritures AL par exemple,
Suite ........ : lorsque j'appelle la fonction de lettrage toutes les écritures 
Suite ........ : se chargent... même les TL et PL
Suite ........ : - LG - 20/06/2006 - FB 18425 - on passe les ecritures d'ano
Suite ........ : - LG - 12/09/2007 - FB 21420 - ajout d'un recupEcrselect 
Suite ........ : pour recup les info de la devise
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.LETTRAGEMClick(Sender: TObject);
var
  R: RLETTR;
  lAction: TActionFiche;
  i : integer ;
  lTOB : TOBM ;
  lTOBEcr : TOB ;
begin
  if TestAvecEcrSimuSitu then Exit;

  FillChar(R, Sizeof(R), #0);
  lAction := taModif;
  if not FOkCreateModif then
    lAction := taConsult;
    
  R.General              := E_GENERAL.Text;
  R.Auxiliaire           := E_AUXILIAIRE.Text;
  R.Appel                := tlMenu;
  R.CritDev              := E_DEVISE.Value;
  R.DeviseMvt            := FTobEcrSelect.GetValue('E_DEVISE');
  R.GL                   := nil;
  R.CritMvt              := FaireStWhere(True) ;

  EtudieModeLettrage(R);

  if ABoMultiSelected then
   begin

    lTOBEcr := TOB.Create('',nil,-1 ) ;

    try

     for i := 1 to FListe.RowCount - 1 do
      if FListe.IsSelected(i) then
       begin
        if not RecupEcrSelect(i) then exit ;
        if FTobEcrSelect.getValue('E_ECRANOUVEAU') = 'OAN' then continue ;
        if ( VH^.ExoV8.Code <> '' ) and ( FTobEcrSelect.getValue('E_DATECOMPTABLE') < VH^.ExoV8.Deb ) then continue ;
        lTOB := TOBM.Create(ecrGen,'',false,lTOBEcr) ;
        lTOB.Dupliquer(FTobEcrSelect,true,true) ;
        //  le dupliquer a supprimer les champs supp du tob
        lTOB.AddChampSupValeur ('OLDDEBIT'     , 0 );
        lTOB.AddChampSupValeur ('OLDCREDIT'    , 0 );
        lTOB.AddChampSupValeur ('RATIO'        , 0 );
        lTOB.AddChampSupValeur ('CONVERTFRANC' , 0 );
        lTOB.AddChampSupValeur ('CONVERTEURO'  , 0 );
        CAjouteChampsSuppLett(lTOB) ;
       end;

     LettrageEnSaisieTOB  ( R, lTOBEcr,true, lAction ) ;

    finally
     lTOBEcr.Free ;
    end ;

   end
    else
     LettrageManuel(R, true, lAction,nil,FTobTiers.GetValue('T_NATUREAUXI'),ComboDefilCpt.ItemIndex,true);

  RefreshPclPGE;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/09/2002
Modifié le ... :   /  /
Description .. : Lettrage Automatique
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.LettrageAClick(Sender: TObject);
begin
  if TestAvecEcrSimuSitu then Exit;
  RapprochementAuto(E_GENERAL.Text, E_AUXILIAIRE.Text);
  RefreshPclPGE;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/09/2002
Modifié le ... :   /  /
Description .. : DéLettrage manuel
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.DelettreClick(Sender: TObject);
var R : RLETTR;
begin
  if not CEstAutoriseDelettrage((FListe.nbSelected = 0) and (FListe.AllSelected), False) then Exit;
  FillChar(R, Sizeof(R), #0);
  R.General    := FTobEcrSelect.GetValue('E_GENERAL');
  R.Auxiliaire := FTobEcrSelect.GetValue('E_AUXILIAIRE');
  R.CritDev    := FTobEcrSelect.GetValue('E_DEVISE');
  R.DeviseMvt  := R.CritDev;
  R.CodeLettre := FTobEcrSelect.GetValue('E_LETTRAGE');
  R.LettrageDevise := (FTobEcrSelect.GetValue('E_LETTRAGEDEV') = 'X');
  R.Appel      := tlMenu;
  R.ToutSelDel := True;
  LettrageManuel(R, False, taModif);
  RefreshPclPGE;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/01/2003
Modifié le ... :   /  /
Description .. : - 09/01/2003 - utilisation CExec... fct de la fiche vierge
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.DelettreMvtExoRef(Sender: TObject);
var lStSql : string;
    lQuery : TQuery;
begin
  if TestAvecEcrSimuSitu or (not CEstAutoriseDelettrage((FListe.nbSelected = 0) and (FListe.AllSelected), False)) then Exit;

  if PgiAsk('Attention, vous allez supprimer tout le lettrage de l''exercice de référence.' + #13 + #10 +
            'Confirmez vous le traitement ?', 'Delettrage complet pour l''exercice de référence') = MrNo then Exit;

  CExecDelettreMvtExoRef(E_GENERAL.Text, E_AUXILIAIRE.Text);

  // GCO - 15/05/2006 - FQ 16580 - Recalcul du dernier Code Lettrage
  lStSql := 'SELECT MAX(E_LETTRAGE) CODE FROM ECRITURE WHERE';

  if E_Auxiliaire.Text <> '' then
    lStSql := lStSql + ' E_AUXILIAIRE = "' + E_Auxiliaire.Text + '"'
  else
    lStSql := lStSql + ' E_GENERAL = "' + E_General.Text + '"';

  try
    lQuery := OpenSql(lStSql, True);

    if E_Auxiliaire .Text <> '' then
      ExecuteSQL('UPDATE TIERS SET T_DERNLETTRAGE = "' + lQuery.FindField('CODE').AsString + '" WHERE ' +
                 'T_AUXILIAIRE = "' + E_Auxiliaire.Text + '"')
    else
      ExecuteSQL('UPDATE GENERAUX SET G_DERNLETTRAGE = "' + lQuery.FindField('CODE').AsString + '" WHERE ' +
                 'G_GENERAL = "' + E_General.Text + '"');
  finally
    Ferme( lQuery );
  end;

  RefreshPclPGE;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/01/2003
Modifié le ... :   /  /
Description .. : - 09/01/2003 - utilisation CExec... fct de la fiche vierge
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.DeletTotal(Sender: TObject);
begin
  inherited;
  if TestAvecEcrSimuSitu or (not CEstAutoriseDelettrage((FListe.nbSelected = 0) and (FListe.AllSelected), False)) then Exit;

  if PGIAsk('Attention, vous allez supprimer tout le lettrage du compte.' + #13 + #10 +
            'Confirmez-vous le traitement ?', 'Delettrage total du compte') = mrNo then
    Exit;
  CExecDeLettrage(E_GENERAL.Text, E_AUXILIAIRE.Text);
  RefreshPclPGE;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/02/2003
Modifié le ... : 30/08/2005
Description .. : Affichage des PARAMSOC Lettrage Gestion
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnClickLetGestion(Sender: TObject);
begin
  ParamSociete(False, BrancheParamSocAVirer, 'SCO_LETTRAGEDETAIL', '', RechargeParamSoc,
    ChargePageSoc, SauvePageSoc, InterfaceSoc, 1105000);
  RechargeParamSoc;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/02/2003
Modifié le ... : 30/08/2005
Description .. : Affichage des PARAMSOC Lettrage Paramètre
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnClickLetParam(Sender: TObject);
begin
  ParamSociete(False, BrancheParamSocAVirer, 'SCO_LETTRAGE', '', RechargeParamSoc, ChargePageSoc,
    SauvePageSoc, InterfaceSoc, 1105000);
  RechargeParamSoc;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/04/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnClickPasseJustifSolde(Sender: TObject);
begin
  if not CbJustifSolde.Enabled then Exit;

  if not CBJustifSolde.Checked then
    CBJustifSolde.Checked := True;

  BCherche.Click;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 23/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/09/2002
Modifié le ... : 24/11/2006
Description .. : Pointage du Compte
Suite ........ : - LG - FB 19058 - 24/11/2006 - ouverture des fct de
Suite ........ : pointage pour le pointage sur journal
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.PointageGClick(Sender: TObject);
var
  Arg : string;
begin
  {JP 30/07/07 : Gestion de l'appel du nouveau pointage}
  if EstSpecif('51213') and (ctxPcl in V_PGI.PGIContexte) then Arg := ''
  else if MsgPointageSurTreso then Arg := CODENEWPOINTAGE + ';';

 if FOkPointageJal then
  begin
   RecupEcrSelect ;
   CPLanceFiche_PointageMul( Arg + FTobEcrSelect.GetValue('E_JOURNAL') + ';' +  FTobEcrSelect.GetValue('E_REFPOINTAGE') + ';' ) ;
  end
   else
    begin
     if E_General.Text <> '' then
      begin
       CPLanceFiche_PointageMul( Arg + E_General.Text + ';');
       RefreshPclPGE;
     end;
    end ; // if
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/09/2002
Modifié le ... : 24/11/2006
Description .. : Depointage du compte
Suite ........ : - FB 19058 - 24/11/2006 - ouverture du pointage sur journal
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.PointageEnCoursClick(Sender: TObject);
var lQuery : TQuery;
begin
  lQuery := nil;
  try
    try
     {JP 12/07/07 : FQ 21045 : nouveau pointage}
     if E_General.Text <> '' then
     begin
       if FOkPointageJal then
        begin
         if (not RecupEcrSelect) or (not EstVraiEcr(FTobEcrSelect)) then Exit ;
         lQuery := OpenSql('SELECT EE_GENERAL, EE_REFPOINTAGE, EE_DATEPOINTAGE, EE_NUMERO, EE_ORIGINERELEVE ' +
                           'FROM ' + FStNomBase + 'EEXBQ WHERE EE_GENERAL = "' + FTobEcrSelect.GetValue('E_JOURNAL') + '"' +
                           ' ORDER BY EE_DATEPOINTAGE DESC', True) ;
        end
        else
         lQuery := OpenSql('SELECT EE_GENERAL, EE_REFPOINTAGE, EE_DATEPOINTAGE, EE_NUMERO, EE_ORIGINERELEVE ' +
                           'FROM ' + FStNomBase + 'EEXBQ WHERE EE_GENERAL = "' + E_General.Text + '"' +
                           ' ORDER BY EE_DATEPOINTAGE DESC', True);

       if not lQuery.Eof then
       begin
         if EstSpecif('51213') and (ctxPcl in V_PGI.PGIContexte) then begin
           if FOkPointageJal then
             CPLanceFiche_Pointage( lQuery.FindField('EE_GENERAL').AsString + ';' +
                                   lQuery.FindField('EE_DATEPOINTAGE').AsString + ';' +
                                   lQuery.FindField('EE_REFPOINTAGE').AsString)
           else
             CPLanceFiche_Pointage( lQuery.FindField('EE_GENERAL').AsString + ';' +
                                   lQuery.FindField('EE_DATEPOINTAGE').AsString + ';' +
                                   lQuery.FindField('EE_REFPOINTAGE').AsString);
         end
         else
           CPLanceFiche_PointageRappro('ACTION=MODIFICATION;' +
                                       lQuery.FindField('EE_GENERAL').AsString + ';' +
                                       lQuery.FindField('EE_DATEPOINTAGE').AsString  + ';' +
                                       lQuery.FindField('EE_REFPOINTAGE').AsString  + ';' +
                                       lQuery.FindField('EE_NUMERO').AsString + ';' +
                                       lQuery.FindField('EE_ORIGINERELEVE').AsString + ';');
       end
       else
         PgiInfo('Aucune référence de pointage n''a été trouvée.', 'Pointage en cours');

       RefreshPclPGE;
     end;
    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : PointageEnCoursClick');
    end;

  finally
    Ferme( lQuery );
  end;
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/09/2002
Modifié le ... :   /  /
Description .. : Liste des immobilisations du compte
Mots clefs ... :
*****************************************************************}
{$IFDEF AMORTISSEMENT}
procedure TOF_CONSECR.ListeImmoClick(Sender: TObject);
begin
  AMLanceFiche_ListeDesImmobilisations(E_General.Text, False);
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/09/2002
Modifié le ... :   /  /
Description .. : Fiche Immobilisation du compte
Mots clefs ... :
*****************************************************************}
{$IFDEF AMORTISSEMENT}
procedure TOF_CONSECR.FicheImmoClick(Sender: TObject);
var lAction: TActionFiche;
begin
  if FTobEcrSelect.GetValue('E_IMMO') <> '' then
  begin
    lAction := taModif;
    if not FOkCreateModif then
      lAction := taConsult;
    AMLanceFiche_FicheImmobilisation(FTobEcrSelect.GetValue('E_IMMO'), lAction, '');
  end;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.CreationLienImmoClick(Sender: TObject);
var lStCell : string;
    lStImmo : string;
    lStSql  : string;
begin
  lStCell := FListe.Cells[1, FListe.Row];

  // GCO - 02/06/2006 - FQ 18288
  if LookUpList( FListe, 'Liste des immobilisations', 'IMMO', 'I_IMMO',
      'I_LIBELLE', 'I_COMPTEIMMO = "' + E_General.Text + '" AND I_ETAT <> "FER"',
      'I_IMMO', True, 0) then
  begin
    lStImmo := FListe.Cells[1, FListe.Row];

    FListe.Cells[1, FListe.Row] := lStCell;

    if FTobEcrSelect.GetValue('E_IMMO') = '' then
    begin
      lStSql := 'UPDATE ECRITURE SET E_IMMO = "' + lSTImmo + '", ' +
                'E_DATEMODIF = "' + UsTime(NowH) + '" WHERE ' +
                'E_JOURNAL = "' + FTobEcrSelect.GetString('E_JOURNAL') + '" AND ' +
                'E_EXERCICE = "' + FTobEcrSelect.GetString('E_EXERCICE') + '" AND ' +
                'E_DATECOMPTABLE = "' + USDateTime(FTobEcrSelect.GetDateTime('E_DATECOMPTABLE')) + '" AND ' +
                'E_NUMEROPIECE = ' + IntToStr(FTobEcrSelect.GetInteger('E_NUMEROPIECE')) + ' AND ' +
                'E_NUMLIGNE = ' + IntToStr(FTobEcrSelect.GetInteger('E_NUMLIGNE')) + ' AND ' +
                'E_NUMECHE = ' + IntToStr(FTobEcrSelect.GetInteger('E_NUMECHE')) + ' AND ' +
                'E_QUALIFPIECE = "' + FTobEcrSelect.GetString('E_QUALIFPIECE') + '" AND ' +
                'E_DATEMODIF = "' + UsTime(FTobEcrSelect.GetDateTime('E_DATEMODIF')) + '"';

      if ExecuteSql(lStSql) <> 1 then
        PgiInfo('L''écriture a été modifiée par un autre utilisateur');
    end;
    RefreshPclPge;
  end;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/12/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.SupprimeLienImmoClick(Sender: TObject);
var lStSql : string;
begin
  if FTobEcrSelect.GetValue('E_IMMO') <> '' then
  begin
    lStSql := 'UPDATE ECRITURE SET E_IMMO = "", ' +
              'E_DATEMODIF = "' + UsTime(NowH) + '" WHERE ' +
              'E_JOURNAL = "' + FTobEcrSelect.GetString('E_JOURNAL') + '" AND ' +
              'E_EXERCICE = "' + FTobEcrSelect.GetString('E_EXERCICE') + '" AND ' +
              'E_DATECOMPTABLE = "' + USDateTime(FTobEcrSelect.GetDateTime('E_DATECOMPTABLE')) + '" AND ' +
              'E_NUMEROPIECE = ' + IntToStr(FTobEcrSelect.GetInteger('E_NUMEROPIECE')) + ' AND ' +
              'E_NUMLIGNE = ' + IntToStr(FTobEcrSelect.GetInteger('E_NUMLIGNE')) + ' AND ' +
              'E_NUMECHE = ' + IntToStr(FTobEcrSelect.GetInteger('E_NUMECHE')) + ' AND ' +
              'E_QUALIFPIECE = "' + FTobEcrSelect.GetString('E_QUALIFPIECE') + '" AND ' +
              'E_DATEMODIF = "' + UsTime(FTobEcrSelect.GetDateTime('E_DATEMODIF')) + '"';

    if ExecuteSql(lStSql) <> 1 then
      PgiInfo('L''écriture a été modifiée par un autre utilisateur');
  end;
  RefreshPclPge;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 26/04/2005
Modifié le ... : 05/09/2006
Description .. : GCO - 05/09/2006 - FQ 18701
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickCutOff(Sender: TObject);
var lStGen : string ;
begin
  lStGen := Copy(FTobGene.GetValue('G_GENERAL'),1,3) ;
  if ( lStGen = '486' ) or ( lStGen = '487' ) then
    CPLanceFiche_CPMULCUTOFF()
  else
    CPLanceFiche_CPMULCUTOFF(E_GENERAL.Text) ;
  RefreshPclPge;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/04/2007
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickControleTva(Sender: TObject);
var lStGene : string;
begin
  // GCO - 26/07/2007 - FQ 21084
  lStGene := Copy(FTobGene.GetString('G_GENERAL'), 1, 5);

  if (lStGene = '44566') or (lStGene = '44571') then
  begin // Tva Déductible ou Tva collextée
    CPLanceFiche_CPCONTROLTVA( IIF(lStGene = '44566', 'DED;', 'COLL;') );
  end
  else
    CPLanceFiche_CPCONTROLTVA(''); // Tva Déductible
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/09/2002
Modifié le ... :   /  /
Description .. : Saisie d'un Bordereau
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.SaisieBorClick(Sender: TObject);
begin
  SaisieFolio(taModif);
  RefreshPclPGE;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/09/2002
Modifié le ... :   /  /
Description .. : Saisie d'une Piece
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.SaisieEcrClick(Sender: TObject);
var OldSC: Boolean;
begin
  OldSC := VH^.BouclerSaisieCreat;
  VH^.BouclerSaisieCreat := False;
  MultiCritereMvt(taCreat, 'N', False);
  VH^.BouclerSaisieCreat := OldSC;
  RefreshPclPGE;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/09/2002
Modifié le ... :   /  /
Description .. : Saisie de Trésorerie
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickSaisieTreso(Sender: TObject);
begin
  CPLanceFiche_SaisieTresorerie(FTobGene.GetString('G_GENERAL'));
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/06/2005
Modifié le ... :   /  /
Description .. : Modif en Série
Mots clefs ... :
***************************************oncli**************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickModifSerie ( Sender : TObject );
var i : integer;
begin
  ChargeFTobListeEcr ;
  if FTobListeEcr.Detail.Count <> 0 then
  begin
    // GCO - 26/07/2007 - FQ 21084
    for i := 0 to FTobListeEcr.Detail.Count - 1 do
    begin
      if (CtxExercice.QuelExoDate( FTobListeEcr.Detail[i].GetString('E_EXERCICE')).EtatCpta <> 'OUV') then
      begin
        PgiInfo('Traitement impossible. Vous avez sélectionné des écritures d''un exercice cloturé.', 'Modification en série');
        Exit;
      end;
    end;

    ModifieEnSerieEcriture( FTobListeEcr );
  end;
  RefreshPclPGE;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/11/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickAffecteAnaEnSerie(Sender: TObject);
begin
  ChargeFTobListeEcr ;
  if FTobListeEcr.Detail.Count <> 0 then
  begin
    TheTob := FTobListeEcr;
    CPLanceFiche_AffecteAnalytiqueEnSerie;
    TheTob := nil;
  end;
  RefreshPclPge;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{$IFDEF COMPTA}
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/06/2005
Modifié le ... :   /  /
Description .. : Reclassement
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnClickReclass  ( Sender : TObject );
var
 lStResult    : string ;
 i            : integer ;
 lReclass     : TZreclassement ;
 lInfoEcr     : TInfoEcriture ;
 lP           : PInfoTW ;
 lYear,lMonth,lDay : Word ;
begin
 lStResult := AGLLanceFiche('CP','CPPARAMCONS','','',E_General.text) ;

 if lStResult = '' then Exit ;

 lInfoEcr             := TInfoEcriture.Create ;
 lReclass             := TZReclassement.Create(lInfoEcr) ;
 lReclass.OnError     := OnErrorTOB ;
 lReclass.StGenSource := E_General.Text ;
 lReclass.StGenDest   := ReadTokenSt(lStResult) ;
 lReclass.StAuxDest   := ReadTokenSt(lStResult) ;
 lReclass.StAuxSource := trim(E_Auxiliaire.Text) ;

 try

 for i := 1 to FListe.RowCount - 1 do
  begin
   if FListe.IsSelected(i) then
    begin
     System.New(lP) ;
     CPInitPInfoTW( lP ) ;
     if not RecupEcrSelect(i) then exit ;
     lP^.TExo.Code          := FTobEcrSelect.GetValue('E_EXERCICE') ;
     lP^.E_JOURNAL          := FTobEcrSelect.GetValue('E_JOURNAL') ;
     lP^.E_NUMEROPIECE      := FTobEcrSelect.GetValue('E_NUMEROPIECE') ;
     lP^.E_NUMLIGNE         := FTobEcrSelect.GetValue('E_NUMLIGNE') ;
     lP^.E_QUALIFPIECE      := FTobEcrSelect.GetValue('E_QUALIFPIECE') ;
     lP^.E_MODESAISIE       := FTobEcrSelect.GetValue('E_MODESAISIE') ;
     lP^.E_NUMECHE          := FTobEcrSelect.GetValue('E_NUMECHE') ;
     lP^.E_DATECOMPTABLE    := FTobEcrSelect.GetValue('E_DATECOMPTABLE') ;
     lP^.E_ETABLISSEMENT    := FTobEcrSelect.GetValue('E_ETABLISSEMENT') ;
     DecodeDate(FTobEcrSelect.GetValue('E_DATECOMPTABLE'),lYear,lMonth,lDay) ;
     lP^.InAnnee            := lYear ;
     lP^.InMois             := lMonth ;
     lP^.InNiv              := 3 ;
     lP^.E_IMMO             := FTobEcrSelect.GetString('E_IMMO');
     lReclass.Add(lP) ;
     lReclass.StJournalSource   := FTobEcrSelect.GetValue('E_JOURNAL') ; 
    end ;
  end ; // for

 if lReclass.StAuxDest <> '' then
  lReclass.TypeTrait := trAuxLigne
   else
    lReclass.TypeTrait := trGenLigne ;

 lReclass.Execute;

 finally
  lInfoEcr.Free ;
  lReclass.Free ;
 end ;

 RefreshPclPGE;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/06/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnErrorTOB (sender : TObject; Error : TRecError ) ;
var
 lMessCompta  : TMessageCompta ;
{$IFNDEF TT}
 lStJournal : string ;
 lStCompte  : string ;
{$ENDIF}
begin

 lMessCompta          := TMessageCompta.Create(Ecran.Caption) ;

 try

{$IFDEF TT}
 if ( trim(Error.RC_Message) <> '' ) then
  PGIInfo(Error.RC_Message + #10#13 + Error.RC_Methode , Ecran.Caption )
   else
    lMessCompta.Execute(Error.RC_Error) ;
{$ELSE}
 if Error.RC_Error = RC_CINTERDIT then
  begin
   lStJournal := ReadtokenSt(Error.RC_Message) ;
   lStCompte  := ReadtokenSt(Error.RC_Message) ;
   PGIInfo(lMessCompta.GetMessage(Error.RC_Error) +#10#13 +  'Pour le journal ' + lStJournal + ' et le compte ' + lStCompte ) ;
  end
   else
    if ( trim(Error.RC_Message) <> '' ) then
     PGIInfo(Error.RC_Message , Ecran.Caption )
      else
       lMessCompta.Execute(Error.RC_Error) ;
{$ENDIF}

finally
 lMessCompta.Free ;
end ;

end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/09/2002
Modifié le ... : 25/05/2007
Description .. : Informations complémentaires ligne
Suite ........ : - LG - 12/08/2004 - FB 17079 - le champs e_blocnote
Suite ........ : n'etait pas sauvegarder ( la propriete M du TOBM n'est plus
Suite ........ : utilisee)
Suite ........ : - LG - 29/08/2005 - FB 16500 - on avais pas acces au 
Suite ........ : libelle et la ref sur une ecriture de cutoff
Suite ........ : - LG - 25/05/2007 - FB 19687 - l'enregistrement des onfo de 
Suite ........ : cutoof ne fonctionnait pas
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickInfosComp(Sender: TObject);
var
  OBM: TOBM;
  RC: R_COMP;
  AA: TActionFiche;
  ModBN : Boolean;
  lQEcr: TQuery;
begin
  if not RecupEcrselect then Exit;

  lQEcr := nil;
  OBM   := nil;
  try
    lQEcr := OpenSql('SELECT * FROM ' + FStNomBase + 'ECRITURE WHERE ' +
             '(E_JOURNAL = "' + FTobEcrSelect.GetValue('E_JOURNAL') + '") AND ' +
             '(E_EXERCICE = "' + FTobEcrSelect.GetValue('E_EXERCICE') + '") AND ' +
             '(E_DATECOMPTABLE = "' + USDateTime(FTobEcrSelect.GetValue('E_DATECOMPTABLE')) + '") AND ' +
             '(E_NUMEROPIECE = ' + IntToStr(FTobEcrSelect.GetValue('E_NUMEROPIECE')) + ') AND ' +
             '(E_NUMLIGNE = ' + IntToStr(FTobEcrSelect.GetValue('E_NUMLIGNE')) + ') AND ' +
             '(E_NUMECHE = ' + IntToStr(FTobEcrSelect.GetValue('E_NUMECHE')) + ') AND ' +
             '(E_QUALIFPIECE = "' + FTobEcrSelect.GetValue('E_QUALIFPIECE') + '")', False);

    // GCO - 26/07/2007 - FQ 21084
    if (CtxExercice.QuelExoDate( FTobEcrSelect.GetString('E_EXERCICE')).EtatCpta <> 'OUV') or
       (not FOkCreateModif) then
      AA := TaConsult
    else
      AA := taModif;

    OBM := TOBM.Create(EcrGen, '', False);
    OBM.ChargeMvt( lQEcr );
    Ferme( lQEcr );

    RC.TOBCompl := nil ;

    if OBM <> nil then
    begin
      if FTobGene.GetValue('G_CUTOFF') = 'X' then
       begin
        RC.StLibre     := '---CUTXXXXXXXXXXXXXXXXXXXXXXXX' ;
        RC.StComporte  := 'XXXXXXXXXX' ;
        RC.CutOffPer   := FTobGene.GetValue('G_CUTOFFPERIODE') ;
        RC.CutOffEchue := FTobGene.GetValue('G_CUTOFFECHUE') ;
        RC.TOBCompl    := CSelectDBTOBCompl(OBM,nil) ;
       end
        else
         begin
          RC.StLibre := 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX';
          RC.StComporte := 'XXXXXXXXXX';
         end ;
      ModBN := True;
      RC.Conso := True;
      RC.Attributs := False;
      RC.MemoComp := nil;
      RC.Origine := -1;
      if SaisieComplement(TOBM(FTobEcrSelect), EcrGen, AA, ModBN, RC, False, True) then
      begin
        CMAJTOBCompl(OBM) ; // FB 19843
        FTobEcrSelect.SetAllModifie(true) ;
        FTobEcrSelect.UpdateDb;
        if FTobGene.GetValue('G_CUTOFF') = 'X' then
         RC.TOBCompl.InsertOrUpdateDB(false);
        BCherche.Click;
      end;
    end;

  finally
    Ferme(lQEcr);
    FreeAndNil(RC.TOBCompl) ;
    FreeAndNil(OBM);
  end;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/10/2004
Modifié le ... : 22/11/2007
Description .. : Affiche du document numérisé de l'écriture
Suite ........ : GCO - 22/11/2007 - utilisation du EC_CLEECR pour LEFT JOIN
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.OnClickDocumentGED( Sender : TObject );
var lStTemp : HString;
    lStDocGuid : HString;
    lStTitre : String;
begin
  lStDocGuid := ATobRow.GetString('CLE_DOCGUID');
  lStTemp := ExtraitDocument(lStDocGuid, lStTitre);
  lStTitre := lStTitre + ' (lecture seule)';
  if lStTemp = '' then
    PgiInfo('Document numérisé associé à l''écriture non trouvé.', 'Voir le document associé')
  else
  begin
    // GCO - 28/01/2008 - Solution provisoire pour les TIF
    if ExisteSQL('SELECT YFI_FILENAME FROM DPDOCUMENT ' +
                 'INNER JOIN YDOCUMENTS ON DPD_DOCGUID=YDO_DOCGUID ' +
                 'LEFT JOIN YDOCFILES ON YDO_DOCGUID=YDF_DOCGUID ' +
                 'LEFT JOIN YFILES ON YDF_FILEGUID=YFI_FILEGUID WHERE ' +
                 'DPD_DOCGUID="' +lStDocGuid + '" AND ' +
                 'YFI_FILENAME LIKE "%.tif"') then

       ShowGedViewer( lStDocGuid, True)
     else
       CPLanceFiche_CPGedFileViewer(lStTemp, FTobEcrSelect);

    // Pour l'instant, on passe en ForcePGIViewer à False, pour privilégier
    // Acrobat Reader pour tous les pdf non "signés" cegid.
    //ShowGedFileViewer(lStTemp, False, lStTitre, True, False, False, GetParamSocSecur ('SO_MDACROBAT', False));
  end;
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/09/2002
Modifié le ... :   /  /
Description .. : Paramètres du compte général
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickCompteGeneral(Sender: TObject);
var lAction: TActionFiche;
begin
  if E_General.Text = '' then Exit;
  inherited;
  lAction := TaModif;
  if ((not ExJaiLeDroitConcept(TConcept(ccGenModif), False)) or
      (not FOkCreateModif)) then
    lAction := taConsult;

  if not FOkCreateModif then
    lAction := TaConsult;
  FicheGene(nil, '', E_General.Text, lAction, 0);
  RefreshPclPGE;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/09/2002
Modifié le ... :   /  /
Description .. : Paramètres du compte auxiliaire
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickCompteAuxiliaire(Sender: TObject);
var lAction: TActionFiche;
begin
  if E_AUXILIAIRE.Text = '' then Exit;
  inherited;
  lAction := taModif;
  if ((not ExJaiLeDroitConcept(TConcept(ccAuxModif), False)) or (not
    FOkCreateModif)) then
    lAction := taConsult;

  FicheTiers(nil, '', E_Auxiliaire.Text, lAction, 0);
  RefreshPclPGE;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/09/2002
Modifié le ... : 28/12/2006
Description .. : Releve d'identité bancaire
Suite ........ : GCO - 28/12/2006 - FQ 19320
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnClickCompteRib(Sender: TObject);
var lOkExiste : Boolean;
    lAction   : TActionFiche;
    lStCompte : String;
begin
  if not FOkAfficheRIB then Exit;

  lAction := TaModif;
  if not FOkCreateModif then
    lAction := TaConsult;

  if FTobGene.GetValue('G_NATUREGENE') = 'BQE' then
  begin
{$IFNDEF IMP}
      lOkExiste := ExisteSQL('SELECT BQ_GENERAL FROM ' + FStNomBase + 'BANQUECP WHERE BQ_GENERAL="' + E_GENERAL.Text
    + '" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"'); // 24/10/2006 YMO Multisociétés

    if (not lOkExiste) then
    begin
      PgiInfo('Vous devez associer un compte bancaire à ce compte.', 'Consultation des écritures');
      Exit;
    end
    else
      FicheBanqueCP(E_GENERAL.Text, lAction, 0);
{$ENDIF}
    //if (lAction = TaModif) and (not okok) then lAction := taCreatOne;
    //if OkOk and (lAction in [taCreat..taCreatOne]) then Exit;
    //Existe := ExisteSQL('Select BQ_GENERAL from BANQUECP where BQ_GENERAL="' + E_GENERAL.Text + '"');
    // G_GENERAL.Enabled := Not Existe ;
    // if Not Existe then CreerRibDefaut ;
  end
  else
  begin
    if (FTobGene.GetValue('G_NATUREGENE') = 'TIC') or
       (FTobGene.GetValue('G_NATUREGENE') = 'TID') then
      // Rechercher le général dans RIB car les RIBS des 'TIC' ou 'TID' sont
      // stockés dans RIB et non pas dans BANQUECP
      lStCompte := E_General.Text
    else
      // Recherche normale du rib de l'auxiliaire dans la table RIB
      lStCompte := E_Auxiliaire.Text;

    lOkExiste := ExisteSQL('SELECT R_AUXILIAIRE FROM ' + FStNomBase + 'RIB WHERE R_AUXILIAIRE="' + lStCompte + '"');
    if (not lOkExiste) then
    begin
      if lAction = TaConsult then
      begin
        PgiInfo('Le compte n''a pas de RIB associé.', 'Consultation des écritures');
        Exit;
      end
      else
      begin
        if PgiAsk('Le compte n''a pas de RIB associé.', 'Voulez-vous le créer ?') <> MrYes then
          Exit;
        lAction := TaCreatOne;
        FicheRIB_AGL( lStCompte, lAction, False, '', False);
      end;
    end
    else // Ouverture du rib
      FicheRIB_AGL( lStCompte, lAction, False, '', False);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/09/2002
Modifié le ... : 28/08/2008
Description .. : GCO - FQ 21274 - Ajout CRelatifVersExercice
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
function TOF_CONSECR.PositionneExo: TExoDate;
var LeExo : TExoDate;
begin
  LeExo := GetEntree;
  if CRelatifVersExercice(E_Exercice.Value) = GetEncours.Code then
    LeExo := GetEncours
  else
    if GetSuivant.Code <> '' then
    begin
      if CRelatifVersExercice(E_Exercice.Value) = GetSuivant.Code then
        LeExo := GetSuivant;
    end
    else
      if GetPrecedent.Code <> '' then
      begin
        if CRelatifVersExercice(E_Exercice.Value) = GetPrecedent.Code then
          LeExo := GetPrecedent;
      end;

  Result := LeExo;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/10/2004
Modifié le ... :   /  /
Description .. : 
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.RechercheEcrituresClick(Sender: Tobject);
begin
  if CtxPcl in V_Pgi.PgiContexte then
    CPLanceFiche_CPRechercheEcr( not FOkCreateModif )
  else
    MultiCritereMvt(taConsult,'N', not FOkCreateModif); // Consultation des écritures
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/09/2002
Modifié le ... :   /  /
Description .. : Mouvements analytiques du compte
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.AnalytiquesClick(Sender: TObject);
var
    lAction : TActionFiche;
begin
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  ACritEdt.CritEdt.Date1 := StrToDate(E_DATECOMPTABLE.Text);
  ACritEdt.CritEdt.Date2 := StrToDate(E_DATECOMPTABLE_.Text);

  ACritEdt.CritEdt.DateDeb := ACritEdt.CritEdt.Date1;
  ACritEdt.CritEdt.DateFin := ACritEdt.CritEdt.Date2;

  ACritEdt.CritEdt.Exo := PositionneExo;
  ACritEdt.CritEdt.SCpt1 := E_GENERAL.Text;
  lAction := taModif;
  if not FOkCreateModif then
    lAction := taConsult;
  TheData := ACritEdt;
  MultiCritereAnaZoom(lAction ,AcritEdt.CritEdt);
  TheData := nil;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : Cumuls du compte général
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.CumulsGENEClick(Sender: TObject);
var LeExo: TExoDate;
begin
  LeExo := PositionneExo;
  CumulCpteMensuel(fbGene, E_GENERAL.Text, FTobGene.GetValue('G_LIBELLE'), LeExo);
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/09/2002
Modifié le ... :   /  /
Description .. : Historique du compte
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
procedure TOF_CONSECR.HistoCpteClick( Sender : TObject );
begin
  CC_LanceFicheHistoCpte(E_General.Text);
end;
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/09/2002
Modifié le ... :   /  /
Description .. : Cumuls auxiliaire
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_ConsEcr.CumulsAUXClick( Sender : TObject );
var lExo: TExoDate;
begin
  lExo := PositionneExo;
  CumulCpteMensuel(fbAux, E_AUXILIAIRE.Text, FTobTiers.GetValue('T_LIBELLE'),  LExo);
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/10/2002
Modifié le ... : 09/02/2007
Description .. : Tables de Chancelleries
Suite ........ : Suppresion du IFNDEF EAGLCLIENT
Mots clefs ... :
*****************************************************************}
procedure TOF_ConsEcr.OnClickTableChancellerie(Sender: Tobject);
var lAction: TActionFiche;
begin
  if not FOkTableChancellerie then Exit;
  lAction := TaModif;
  if not FOkCreateModif then
    LACtion := TaConsult;

  if (FTobTiers.GetValue('T_MULTIDEVISE') = 'X') then
   //(EstMonnaieIN(FTobTiers.GetValue('T_DEVISE'))) then
    FicheDevise(FTobTiers.GetValue('T_DEVISE'), lAction, False)
  else
    FicheChancel(FTobTiers.GetValue('T_DEVISE'), False, 0, lAction, True);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : Analyse Complémentaire
Mots clefs ... :
*****************************************************************}
procedure TOF_ConsEcr.ANALREVISClick(Sender: TObject);
var lWhereSQL : string;
begin
  lWhereSQL := CMajRequeteExercice ( E_EXERCICE.Value, AStSqlWhereTobFListe);
  AGLLanceFiche('CP', 'CPCONSULTREVIS', '', '', lWhereSQL);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :   /  /
Modifié le ... : 03/11/2004
Description .. : 
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
function TOF_CONSECR.TestAvecEcrSimuSitu : Boolean;
begin
  Result := False;
  if (Pos('S', UpperCase(EQualifPiece.Text)) > 0) or
     (Pos('U', UpperCase(EQualifPiece.Text)) > 0) then
  begin
    PgiInfo('Traitement impossible - Indiquer "Normal" en type d''écriture ' + #10 +
            'sur l''onglet "Compléments".', 'Consultation des écritures');
    Result := True;
  end;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/11/2002
Modifié le ... :   /  /
Description .. : Passage des ecritures simplifiées Compte à Compte
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.PassageEcrSimplifie(Sender: TObject);
var
  lTobEcr     : TOB ;
  lTobEcrSimp : TOB ;
  lTobListe   : TOB ;
  lStRetour   : string ;
  i           : integer ;
  lRdSolde    : double ;

  procedure _CreerTOB ;
   var
    lTOB : TOB ;
   begin
    lTOB  := TOB.Create('ECRITURE',lTOBEcr,-1) ;
    lTOB.Dupliquer(FTOBEcrSelect,false,true) ; // FB 21120 on dupliquait pas la ligne charge en base
    lTOB.AddChampSupValeur('ECARTREGUL',0,true) ;
    lTOB.PutValue( 'ECARTREGUL' , lTOB.GetValue('E_DEBIT') - lTOB.GetValue('E_CREDIT') ) ;
   end ;
begin
  lRdSolde := 0 ;
  lTOBEcr  := TOB.Create('',nil,-1) ;

  try

  if ABoMultiSelected then
   begin

     for i := 1 to FListe.RowCount - 1 do
      if FListe.IsSelected(i) then
       begin
        lTobListe := GetO(FListe) ;
        if lTobListe.getValue('CLE_ECRANOUVEAU') = 'OAN' then continue ;
        if ( VH^.ExoV8.Code <> '' ) and ( lTobListe.getValue('E_DATECOMPTABLE') < VH^.ExoV8.Deb ) then continue ;
        if not RecupEcrSelect(i) then exit ;
        lRdSolde := lRdSolde + ( FTobEcrSelect.GetValue('E_DEBIT') - FTobEcrSelect.GetValue('E_CREDIT') ) ;
        _CreerTOB ;
       end; // if


     if Arrondi(lRdSolde,2) = 0 then
      begin
       PGIInfo('Le paquet d''écriture est soldé !' );
       exit ;
      end ;

     for i := 0 to lTOBEcr.Detail.Count - 1 do
      lTOBEcr.Detail[i].PutValue('ECARTREGUL', lRdSolde ) ;

   end
    else
     begin
       if not (EstVraiEcr( ATobRow )) then Exit;
       _CreerTOB ;
     end ;


   lTobEcrSimp := TOB.Create('', nil, -1);
   lStRetour   := CPLanceFiche_ECRLET(lTOBEcr, lTobEcrSimp, 'AL6');

    if ReadTokenST(lStRetour) = '1' then
     begin
      if FOkLettrable and ( PGIAsk('Voulez vous considérer ces mouvements comme lettrés') = mrYes ) then
       CSaveEcr(lTobEcrSimp,lTOBEcr)
        else
         CSaveEcr(lTobEcrSimp,nil) ;
      RefreshPclPGE;
     end ; // if

  finally
   FreeAndNil(lTOBEcr) ;
   FreeAndNil(lTobEcrSimp) ;
  end ;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
// ------------------------------ Editions -------------------------------------
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/11/2002
Modifié le ... : 18/02/2004
Description .. : Justificatif de solde TID / TIC ou Tiers
Suite ........ : GCO - 18/02/2004
Suite ........ : -> Utilisation du CritEdt pour le GL
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickJustifSolde(Sender: TObject);
begin
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);

  PrepareArgumentGL( True );
  TheData := ACritEdt;

  if (FTobGene.GetValue('G_NATUREGENE') = 'TIC') or (FTobGene.GetValue('G_NATUREGENE') = 'TID') or (FTobGene.GetValue('G_NATUREGENE') = 'DIV') then
  begin
    ACritEdt.CritEdt.Cpt1 := E_General.Text;
    ACritEdt.CritEdt.Cpt2 := E_General.Text;
    CPLanceFiche_CPGLGENE('');
  end
  else
  begin
    ACritEdt.CritEdt.Cpt1 := E_Auxiliaire.Text;
    ACritEdt.CritEdt.Cpt2 := E_Auxiliaire.Text;

    // GCO - 23/11/2004 - FQ 14848
    if Trim(E_General.Text) <> '' then
    begin
      ACritEdt.CritEdt.SCpt1 := E_General.Text;
      ACritEdt.CritEdt.SCpt2 := E_General.Text;
      CPLanceFiche_CPGLAUXIPARGENE('');
    end
    else
      CPLanceFiche_CPGLAUXI('');
  end;
  TheData := nil;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/11/2002
Modifié le ... : 18/02/2004
Description .. : Grand Livre Général
Suite ........ : GCO - 18/02/2004
Suite ........ : -> Utilisation du CritEdt pour le GL
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickGLGene(Sender: TObject);
begin

  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGL( False );
  ACritEdt.CritEdt.Cpt1 := E_General.Text;
  ACritEdt.CritEdt.Cpt2 := E_General.Text;
  TheData := ACritEdt;
  CPLanceFiche_CPGLGene('');
  TheData := nil;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/10/2002
Modifié le ... : 18/02/2004
Description .. : Grand Livre Général par Auxiliaire
Suite ........ : GCO - 18/02/2004
Suite ........ : -> Utilisation du CritEdt pour le GL
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickGLGeneParAuxi(Sender: TObject);
begin
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGL( False );
  ACritEdt.CritEdt.Cpt1  := E_General.Text;
  ACritEdt.CritEdt.Cpt2  := E_General.Text;
  ACritEdt.CritEdt.SCpt1 := E_Auxiliaire.Text;
  ACritEdt.CritEdt.SCpt2 := E_Auxiliaire.Text;
  TheData := ACritEdt;
  CPLanceFiche_CPGLGeneParAuxi('');
  TheData := nil;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/10/2002
Modifié le ... : 18/02/2004
Description .. : Grand-livre auxiliaire
Suite ........ : GCO - 18/02/2004
Suite ........ : -> Utilisation du CritEdt pour le GL
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickGLAuxi(Sender: TObject);
begin
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGL( False );
  ACritEdt.CritEdt.Cpt1  := E_Auxiliaire.Text;
  ACritEdt.CritEdt.Cpt2  := E_Auxiliaire.Text;
  TheData := ACritEdt;
  CPLanceFiche_CPGLAuxi('');
  TheData := nil;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/11/2002
Modifié le ... : 18/02/2004
Description .. : Grand Livre Auxliaire par Général
Suite ........ : GCO - 18/02/2004
Suite ........ : -> Utilisation du CritEdt pour le GL
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickGLAuxiParGene(Sender: TObject);
begin
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGL( False );
  ACritEdt.CritEdt.Cpt1  := E_Auxiliaire.Text;
  ACritEdt.CritEdt.Cpt2  := E_Auxiliaire.Text;
  ACritEdt.CritEdt.SCpt1 := E_General.Text;
  ACritEdt.CritEdt.SCpt2 := E_General.Text;
  TheData := ACritEdt;
  CPLanceFiche_CPGLAuxiParGene('');
  TheData := nil;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/05/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickGLAuxiPourGene(Sender: TObject);
begin
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGL( False );
  ACritEdt.CritEdt.Cpt1  := E_Auxiliaire.Text;
  ACritEdt.CritEdt.Cpt2  := E_Auxiliaire.Text;
  ACritEdt.CritEdt.SCpt1 := E_General.Text;
  ACritEdt.CritEdt.SCpt2 := E_General.Text;
  TheData := ACritEdt;
  CPLanceFiche_CPGLAuxiParGene('');
  TheData := nil;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/11/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickGLGeneParQte(Sender: TObject);
begin
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGL( False );
  ACritEdt.CritEdt.Cpt1 := E_General.Text;
  ACritEdt.CritEdt.Cpt2 := E_General.Text;
  TheData := ACritEdt;
  CPLanceFiche_CPGLGeneParQte('');
  TheData := nil;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/10/2004
Modifié le ... :   /  /
Description .. : Grand Livre Analytique
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.LancementGLAna( vFichierBase : TFichierBase );
begin
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGL( False );
  ACritEdt.CritEdt.GL.Axe := fbToAxe( vFichierBase );
  TheData := ACritEdt;
  CPLanceFiche_CPGLAna;
  TheData := nil;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickGLAnaSurAxe1(Sender: TObject);
begin
  LancementGLAna( fbAxe1 );
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickGLAnaSurAxe2(Sender: TObject);
begin
  LancementGLAna( fbAxe2 );
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickGLAnaSurAxe3(Sender: TObject);
begin
  LancementGLAna( fbAxe3 );
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickGLAnaSurAxe4(Sender: TObject);
begin
  LancementGLAna( fbAxe4 );
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickGLAnaSurAxe5(Sender: TObject);
begin
  LancementGLAna( fbAxe5 );
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.LancementGLGENEPARANA( vFichierBase : TFichierBase );
begin
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGL( False );
  ACritEdt.CritEdt.Cpt1   := E_General.Text;
  ACritEdt.CritEdt.Cpt2   := E_General.Text;
  ACritEdt.CritEdt.GL.Axe := fbToAxe( vFichierBase );
  TheData := ACritEdt;
  CPLanceFiche_CPGLGENEPARANA;
  TheData := nil;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /
Description .. : Grand Livre Général par Analytique sur l'axe 1
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickGLGeneParAna1(Sender: TObject);
begin
  LancementGLGENEPARANA( fbAxe1 );
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /
Description .. : Grand Livre Général par Analytique sur l'axe 1
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickGLGeneParAna2(Sender: TObject);
begin
  LancementGLGENEPARANA( fbAxe2 );
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /
Description .. : Grand Livre Général par Analytique sur l'axe 1
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickGLGeneParAna3(Sender: TObject);
begin
  LancementGLGENEPARANA( fbAxe3 );
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /
Description .. : Grand Livre Général par Analytique sur l'axe 1
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickGLGeneParAna4(Sender: TObject);
begin
  LancementGLGENEPARANA( fbAxe4 );
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /
Description .. : Grand Livre Général par Analytique sur l'axe 1
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickGLGeneParAna5(Sender: TObject);
begin
  LancementGLGENEPARANA( fbAxe5 );
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/03/2004
Modifié le ... : 11/03/2004
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.PrepareArgumentBAL;
begin
  // Exercice
  ACritEdt.CritEdt.Exo.Code := CRelatifVersExercice(E_Exercice.Value);
  // Début de la Date Comptable (1)
  ACritEdt.CritEdt.Date1 := StrToDate(E_DateComptable.Text);
  // Fin de la Datecomptable (2)
  ACritEdt.CritEdt.Date2 := StrToDate(E_DateComptable_.Text);
  // Type d'Ecritures
  ACritEdt.CritEdt.Qualifpiece := EQualifPiece.Text;
  // Etablissement
  ACritEdt.CritEdt.Etab := E_Etablissement.Value;
  // Devise
  ACritEdt.CritEdt.DeviseSelect := E_Devise.Value;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/03/2004
Modifié le ... : 11/03/2004
Description .. : Balance Générale
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickBalGene(Sender: TObject);
begin
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentBAL;
  ACritEdt.CritEdt.Cpt1 := E_General.Text;
  ACritEdt.CritEdt.Cpt2 := E_General.Text;
  TheData := ACritEdt;
  CPLanceFiche_BalanceGeneral;
  TheData := nil;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/03/2004
Modifié le ... : 11/03/2004
Description .. : Balance Auxiliaire
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickBalAuxi(Sender: TObject);
begin
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentBAL;
  ACritEdt.CritEdt.Cpt1 := E_General.Text;
  ACritEdt.CritEdt.Cpt2 := E_General.Text;
  TheData := ACritEdt;
  CPLanceFiche_BalanceAuxiliaire;
  TheData := nil;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/10/2002
Modifié le ... :   /  /
Description .. : Balance général par auxiliaire
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickBalGeneParAuxi(Sender: TObject);
begin
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentBAL;
  ACritEdt.CritEdt.Cpt1  := E_General.Text;
  ACritEdt.CritEdt.Cpt2  := E_General.Text;
  ACritEdt.CritEdt.SCpt1 := E_Auxiliaire.Text;
  ACritEdt.CritEdt.SCpt2 := E_Auxiliaire.Text;
  TheData := ACritEdt;
  CPLanceFiche_BalanceGenAuxi;
  TheData := nil;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/11/2002
Modifié le ... :   /  /
Description .. : Balance Auxiliaire par Général
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickBalAuxiParGene(Sender: TObject);
begin
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentBAL;
  ACritEdt.CritEdt.Cpt1  := E_Auxiliaire.Text;
  ACritEdt.CritEdt.Cpt2  := E_Auxiliaire.Text;
  ACritEdt.CritEdt.SCpt1 := E_General.Text;
  ACritEdt.CritEdt.SCpt2 := E_General.Text;
  TheData := ACritEdt;
  CPLanceFiche_BalanceAuxiGen;
  TheData := nil;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/10/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.LancementBalAnaSurAxe( vFichierBase : TFichierBase );
begin
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentBAL;
  ACritEdt.CritEdt.Bal.Axe := fbToAxe( vFichierBase );
  TheData := ACritEdt;
  CPLanceFiche_BalanceAnalytique;
  TheData := nil;
end;
{$ENDIF}

///////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickBalAnaSurAxe1(Sender: TObject);
begin
  LancementBalAnaSurAxe( fbAxe1 );
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickBalAnaSurAxe2(Sender: TObject);
begin
  LancementBalAnaSurAxe( fbAxe2 );
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickBalAnaSurAxe3(Sender: TObject);
begin
  LancementBalAnaSurAxe( fbAxe3 );
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickBalAnaSurAxe4(Sender: TObject);
begin
  LancementBalAnaSurAxe( fbAxe4 );
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickBalAnaSurAxe5(Sender: TObject);
begin
  LancementBalAnaSurAxe( fbAxe5 );
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.LancementBalGeneParAna( vFichierBase : TFichierBase );
begin
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentBAL;
  ACritEdt.CritEdt.Cpt1 := E_General.Text;
  ACritEdt.CritEdt.Cpt2 := E_General.Text;
  ACritEdt.CritEdt.BAL.Axe := fbToAxe( vFichierBase );
  TheData := ACritEdt;
  CPLanceFiche_BalanceGenAnal;
  TheData := nil;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEf COMPTA}
procedure TOF_CONSECR.OnClickBalGeneParAna1(Sender: TObject);
begin
  LancementBalGeneParAna(FbAxe1);
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEf COMPTA}
procedure TOF_CONSECR.OnClickBalGeneParAna2(Sender: TObject);
begin
  LancementBalGeneParAna(FbAxe2);
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEf COMPTA}
procedure TOF_CONSECR.OnClickBalGeneParAna3(Sender: TObject);
begin
  LancementBalGeneParAna(FbAxe3);
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEf COMPTA}
procedure TOF_CONSECR.OnClickBalGeneParAna4(Sender: TObject);
begin
  LancementBalGeneParAna(FbAxe4);
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEf COMPTA}
procedure TOF_CONSECR.OnClickBalGeneParAna5(Sender: TObject);
begin
  LancementBalGeneParAna(FbAxe5);
end;
{$ENDIF}

{$IFDEF COMPTA}
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/08/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.OnClickEtatRapprochement( Sender : TObject );
begin
{$IFNDEF CCMP}
If OkNewPointage Then CC_LanceFicheEtatRapproDet(FTobGene.GetValue('G_GENERAL'))
                 Else CC_LanceFicheEtatRapproDetV7(FTobGene.GetValue('G_GENERAL')) ;
{$ENDIF}
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/10/2004
Modifié le ... :   /  /
Description .. : Justificatif des soldes bancaires
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickJustifSoldeBQE( Sender : TObject );
begin
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  ACritEdt.CritEdt.Cpt1         := E_General.Text;
  ACritEdt.CritEdt.Cpt2         := E_General.Text;
  ACritEdt.CritEdt.Exo.Code     := CRelatifVersExercice(E_Exercice.Value); // Exercice
  ACritEdt.CritEdt.Date1        := StrToDate(E_DateComptable.Text);        // Date Début
  ACritEdt.CritEdt.Date2        := StrToDate(E_DateComptable_.Text);       // Date Fin
  ACritEdt.CritEdt.Etab         := E_ETABLISSEMENT.Value;                  // Etablissement
  ACritEdt.CritEdt.DeviseSelect := E_DEVISE.Value;                         // Devise
  TheData := ACritEdt;
  CC_LanceFicheJustifPointage;
  TheData := nil;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/10/2002
Modifié le ... : 24/11/2006
Description .. : Détermine les traitements autorisés pour le compte général
Suite ........ : ou
Suite ........ : de l'auxiliaire
Suite ........ : - LG - 13/09/2005 - FB 16348 - on ne peut lettrer des
Suite ........ : ecritures ANO
Suite ........ : - LG - FB 19058 - ouverture des fct pour le lettrage sur
Suite ........ : journal
Suite ........ : GCO - 31/07/2007 - Ne jamais mettre de tests sur les ecritures
Suite ........ : car ATobFListe peut ne pas être encore chargée.
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.WhatIsOk;
var
 lTOBEcr : TOB ;
 lQ : TQuery ;
begin
  // Init par Défaut
  FOkLettrable         := False;
  FOkOngletLettrage    := False;
  FOkPointable         := False;
  FOkAnalytique        := False;
  FOkAnaSurAxe1        := False;
  FOkAnaSurAxe2        := False;
  FOkAnaSurAxe3        := False;
  FOkAnaSurAxe4        := False;
  FOkAnaSurAxe5        := False;
  FOkSaisieCA3         := False;
  FOkEcheance          := False;
  FOkTableChancellerie := False;
  FOKCentralisation    := False;
  FOkJustificatifSolde := False;
  FOkAfficheRIB        := False;
  FOkEdition           := False;
  FOkIcc               := False;
  FOkCre               := False;

  if (FTobGene = nil) or
     (FTobTiers = nil) or
     (ATobFliste = nil) or
     (FBoConfidentielGene) or
     (FBoConfidentielAuxi) then Exit;

   if (FTobTiers.GetString('T_LETTRABLE') = 'X') and
      (FTobGene.GetString('G_GENERAL') <> '') then
     FOkLettrable := True
   else
   begin
     FOkLettrable := ((FTobGene.GetString('G_NATUREGENE') = 'TIC') or
                      (FTobGene.GetString('G_NATUREGENE') = 'TID') or
                      (FTobGene.GetString('G_NATUREGENE') = 'DIV')) and
                      (FTobGene.GetString('G_LETTRABLE') = 'X')
   end;

  lTOBEcr := GetO(FListe) ;
  if lTOBEcr <> nil then
   FOkFctLettrage := FOkLettrable and ( lTOBEcr.GetString('CLE_ECRANOUVEAU') = 'N' ) ;

  FOkJustificatifSolde := (FTobTiers.GetString('T_LETTRABLE') = 'X') OR
                          (((FTobGene.GetString('G_NATUREGENE') = 'TIC') or
                            (FTobGene.GetString('G_NATUREGENE') = 'TID') or
                            (FTobGene.GetString('G_NATUREGENE') = 'DIV')) and
                            (FTobGene.GetString('G_LETTRABLE') = 'X'));

  FOkOngletLettrage := (FTobTiers.GetString('T_LETTRABLE') = 'X') or
                       (FTobGene.GetString('G_COLLECTIF') = 'X');

  // FOkPointable
  if FOkPointageJal then
   begin
    if ( lTOBEcr <> nil ) then
     begin
      FOkPointable := ( lTOBEcr <> nil ) and ( lTOBEcr.GetString('E_REFPOINTAGE') <> '') ;
      if not FOkPointable and ( FZListJournal.Load([lTOBEcr.GetString('E_JOURNAL')]) > - 1 ) and
         ( FZListJournal.GetString('J_NATUREJAL') = 'BQE' ) then
       begin
        lQ := OpenSql('SELECT G_POINTABLE FROM GENERAUX WHERE G_GENERAL="' + FZListJournal.GetString('J_CONTREPARTIE') + '" ' , true) ;
        FOkPointable :=  lQ.FindField('G_POINTABLE').asString = 'X' ;
        Ferme(lQ) ;
       end ;
     end ;
   end
    else
     FOkPointable := (FTobGene.GetString('G_POINTABLE') = 'X');

  // Peut on demander la centralisation des écritures ????
{$IFDEF EAGLCLIENT}
  FOkCentralisation := False;
{$ELSE}
  FOKCentralisation := (E_General.Text <> '') and (FTobGene.GetString('G_CENTRALISABLE') = 'X');
{$ENDIF}

  ActivationOnglet; // Activation des onglets Lettrage et Pointage en fonction de FOkLettrable et FOkPointable

  FOkAnalytique := (GetParamSocSecur('SO_CPPCLSANSANA', True) = False) and
                   (FTobGene.GetString('G_VENTILABLE') = 'X');

  FOkAnaSurAxe1 := FOkAnalytique and (FTobGene.GetString('G_VENTILABLE1') = 'X');
  FOkAnaSurAxe2 := FOkAnalytique and (FTobGene.GetString('G_VENTILABLE2') = 'X');
  FOkAnaSurAxe3 := FOkAnalytique and (FTobGene.GetString('G_VENTILABLE3') = 'X');
  FOkAnaSurAxe4 := FOkAnalytique and (FTobGene.GetString('G_VENTILABLE4') = 'X');
  FOkAnaSurAxe5 := FOkAnalytique and (FTobGene.GetString('G_VENTILABLE5') = 'X');

  // GCO - 31/07/2007 -
  // Pas fiable car la grille n'est pas encpre chargé dans certains cas.
  // Je le met dans le AFTERLOAD pour tester.
  
  FOkSaisieCA3 := (Copy(E_GENERAL.Text, 1, 3) = '445');

  FOkEcheance := (GetParamSocSecur('SO_ZSAISIEECHE', False)) and
                 ((FTobGene.GetString('G_NATUREGENE') = 'TIC') or
                  (FTobGene.GetString('G_NATUREGENE') = 'TID') or
                  (FTobGene.GetString('G_POINTABLE') = 'X') or
                  (E_Auxiliaire.Text <> ''));

  FOkTableChancellerie := (E_Auxiliaire.Text <> '') and
                          ((FTobTiers.GetString('T_DEVISE') <> V_PGI.DevisePivot) or
                           (FTobTiers.GetString('T_MULTIDEVISE') = 'X'));

  // GCO - 25/11/2004 - FQ 14928 - Autorise l'affichage du RIB dans le MENU
  FOkAfficheRIB := (FTobGene.GetString('G_NATUREGENE') = 'BQE') or
                   (FTobGene.GetString('G_NATUREGENE') = 'TIC') or
                   (FTobGene.GetString('G_NATUREGENE') = 'TID') or
                   (FTobTiers.GetString('T_AUXILIAIRE') <> '');

  // GCO - 10/04/2007 Pour info : FQ 13007 Vérouillage des éditions si confidentialité
  FOkEdition := (not ABoMultiSelected) and (not ((FBoConfidentielGene) or (FBoConfidentielAuxi)));

  FOkCre := (FOkCreGere) and ((Copy(FTobGene.GetString('G_GENERAL'), 1, 2) = '16') Or (Copy(FTobGene.GetString('G_GENERAL'), 1, 4) = '6611'));

  // GCO - 05/07/2007 - FQ 20924 + Chargement en Tob au départ pour éviter requête
  if FOkIccGere then
    FOkIcc := (FTobListeICC.FindFirst(['ICG_GENERAL'], [FTobGene.GetString('G_GENERAL')], False)) <> nil;

  // GCO - 22/02/2007
  FillChar( FExoDate, SizeOf(FExoDate), #0);
  if E_Exercice.ItemIndex <> 0 then
  begin
    FExoDate.Code := CRelatifVersExercice( E_Exercice.Value );
    RempliExoDate(FExoDate);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/11/2002
Modifié le ... : 06/11/2002
Description .. : Verifie nature du compte et le type de compte et si présence
suite ........ : du VH^.ExoV8.Code alors on force E_Datecomptable à VH^.ExoV8.Deb
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.VerifieLesDates;
begin
  // GCO - 04/04/2007 - FQ 19535
  if (FTobGene = nil) or (FTobTiers = nil) then Exit; 

  FBoOkControleDate := True;
  if E_EXERCICE.ItemIndex = 0 then
  begin
    if (GetExoV8.Code <> '') then
    begin
      if (FTobTiers.GetString('T_AUXILIAIRE') <> '') or
        ((FTobGene.GetString('G_GENERAL') <> '') and
        ((FTobGene.GetString('G_NATUREGENE') <> 'CHA') and
        (FTobGene.GetString('G_NATUREGENE') <> 'PRO'))) then
      begin
        if StrToDate(E_DateComptable.Text) < GetExoV8.Deb then
        begin
        {$IFNDEF BUREAU}
          // Pas de message bloquant en arrivant dans l'écran depuis le BUREAU
          PgiInfo('Vous ne pouvez pas saisir une date inférieure au  ' + DateToStr(GetExoV8.Deb) + '.', 'Consultation des écritures');
        {$ENDIF}
          E_DateComptable.Text := DateToStr(GetExoV8.Deb);
          FBoOkControleDate := False;
        end;
      end;
    end
    else
    begin
      // GCO - 11/12/2005 - FQ 13365
      if StrToDate(E_DateComptable.Text) < GetExercices(1).Deb then
      begin                          
        if StrToDate(E_DateComptable.Text) <> idate1900 then
          PgiInfo('Vous ne pouvez pas saisir une date inférieure au  ' + DateToStr(GetExercices(1).Deb) + '.', 'Consultation des écritures');
        E_DateComptable.Text := DateToStr(GetExercices(1).Deb);
        FBoOkControleDate := False;
      end;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLCLIENT}
procedure TOF_CONSECR.TraiteCentralisation;
var
  i: integer;
  lTobEcr: Tob;
  lTobEcrCentra: Tob;
  lTobListeCentra: Tob;
  lDateComptable: TDateTime;
  lPeriode: string;
  lJournal: string;
  lDebit, lCredit, lSoldePro : Double;
  lDevise: string;
begin
  lTobListeCentra := Tob.Create('', nil, -1);
  lTobListeCentra.Dupliquer(ATobFListe, False, False, False);

  i := ATobFListe.Detail.Count - 1;
  repeat
    lTobEcr := ATobFListe.Detail[i];
    FZListJournal.Load([lTobEcr.GetString('E_JOURNAL')]);
    if FZListJournal.GetValue('J_CENTRALISABLE') = 'X' then
    begin
      lTobEcrCentra := lTobListeCentra.FindFirst(['CLE_PERIODE', 'E_JOURNAL'],
        [lTobEcr.GetValue('CLE_PERIODE'), lTobEcr.GetValue('E_JOURNAL')], False);
      if lTobEcrCentra = nil then
      begin
        lPeriode       := lTobEcr.GetValue('CLE_PERIODE');
        lDateComptable := lTobEcr.GetValue('E_DATECOMPTABLE');
        lJournal       := lTobEcr.GetValue('E_JOURNAL');
        lDebit         := lTobEcr.GetValue('E_DEBIT');
        lCredit        := lTobEcr.GetValue('E_CREDIT');
        lDevise        := lTobEcr.GetValue('E_DEVISE');
        lSoldePro      := lTobEcr.GetValue('SOLDEE');

        lTobEcr.ChangeParent(lTobListeCentra, -1);
        lTobEcr.InitValeurs;
        lTobEcr.PutValue('CLE_PERIODE', lPeriode);
        lTobEcr.PutValue('E_DATECOMPTABLE', FinDeMois(lDateComptable));
        lTobEcr.PutValue('E_JOURNAL', lJournal);
        lTobEcr.PutValue('E_LIBELLE', TraduireMemoire('Centralisation mensuelle'));
        lTobEcr.PutValue('E_DEVISE', lDevise);
        lTobEcr.PutValue('E_DEBIT', lDebit);
        lTobEcr.PutValue('E_CREDIT', lCredit);
        lTobEcr.PutValue('SOLDEE', lSoldePro);
      end
      else
      begin
        lDebit := lTobEcrCentra.GetValue('E_DEBIT') + lTobEcr.GetValue('E_DEBIT');
        lCredit := lTobEcrCentra.GetValue('E_CREDIT') + lTobEcr.GetValue('E_CREDIT');
        lTobEcrCentra.PutValue('E_DEBIT', lDebit);
        lTobEcrCentra.PutValue('E_CREDIT', lCredit);
        lTobEcr.Free;
      end;
    end;

    i := i - 1;
  until (i < 0);

{$IFDEF TT}
  lTobListeCentra.SaveToFile('C:\Centra.txt', False, True, True);
{$ENDIF}

  // Ajout des centralisations de la Tob de départ
  if lTobListeCentra.Detail.Count <> 0 then
  begin
    i := lTobListeCentra.Detail.Count - 1;
    repeat
      lTobListeCentra.Detail[i].ChangeParent(ATobFListe, -1);
      i := i - 1;
    until (i < 0);
  end;

  lTobListeCentra.ClearDetail;
  lTobListeCentra.Free;

end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/04/2006
Modifié le ... : 23/08/2007
Description .. : - LG - 22/08/2007 - Fb 21264/21266 - on affecte trus si on
Suite ........ : recharge pas l'ecriture
Mots clefs ... :
*****************************************************************}
function TOF_CONSECR.RecupEcrSelect ( TheRow : integer = 0 ) : Boolean;
var lTobEcr : Tob;
    lQuery  : TQuery;
    SQL     : string;
begin
  Result := False;

  if TheRow = 0 then
   lTobEcr := GetO(FListe, FListe.Row)
  else
   lTobEcr := GetO(FListe, TheRow) ;

  if not (EstVraiEcr(lTobEcr)) then
  begin
    FTobEcrSelect.InitValeurs; // Init des champs de la tob
    Exit;
  end;

  if (lTobEcr.GetValue('E_JOURNAL') <> FTobEcrSelect.GetValue('E_JOURNAL')) or
     (lTobEcr.GetValue('CLE_EXERCICE') <> FTobEcrSelect.GetValue('E_EXERCICE')) or
     (lTobEcr.GetValue('E_DATECOMPTABLE') <> FTobEcrSelect.GetValue('E_DATECOMPTABLE')) or
     (lTobEcr.GetValue('E_NUMEROPIECE') <> FTobEcrSelect.GetValue('E_NUMEROPIECE')) or
     (lTobEcr.GetValue('CLE_NUMLIGNE') <> FTobEcrSelect.GetValue('E_NUMLIGNE')) or
     (lTobEcr.GetValue('CLE_NUMECHE') <> FTobEcrSelect.GetValue('E_NUMECHE')) or
     (lTobEcr.GetValue('CLE_QUALIFPIECE') <> FTobEcrSelect.GetValue('E_QUALIFPIECE')) then
  begin
    FTobEcrSelect.InitValeurs; // Init des champs de la tob

    // GCO - 29/05/2006 - FQ 18212
    // GCO - 03/04/2007 - FQ 19625 - Erreur car le document doit s'afficher pour
    // groupe d'écritures, on ne peut utiliser la clé de la table ECRITURE.
    SQL := 'SELECT * FROM ' + FStNomBase + 'ECRITURE ' +
           'LEFT JOIN ' + FStNomBase + 'ECRCOMPL ON ' +
           'E_JOURNAL = EC_JOURNAL AND ' +
           'E_EXERCICE = EC_EXERCICE AND ' +
           'E_DATECOMPTABLE = EC_DATECOMPTABLE AND ' +
           'E_NUMEROPIECE = EC_NUMEROPIECE AND ' +
           'E_NUMLIGNE = EC_NUMLIGNE AND ' +
           'E_NUMECHE = EC_NUMECHE AND ' +
           'E_QUALIFPIECE = EC_QUALIFPIECE WHERE ' +
           'E_JOURNAL= "' + lTobEcr.GetValue('E_JOURNAL') + '" AND ' +
           'E_EXERCICE = "' + lTobEcr.GetValue('CLE_EXERCICE') + '" AND ' +
           'E_DATECOMPTABLE = "' + USDateTime(lTobEcr.GetValue('E_DATECOMPTABLE')) + '" AND ' +
           'E_NUMEROPIECE = ' + IntToStr(lTobEcr.GetValue('E_NUMEROPIECE')) + ' AND ' +
           'E_NUMLIGNE = ' + IntToStr(lTobEcr.GetValue('CLE_NUMLIGNE')) + ' AND ' +
           'E_NUMECHE = ' + IntToStr(lTobEcr.GetValue('CLE_NUMECHE')) + ' AND ' +
           'E_QUALIFPIECE = "' + lTobEcr.GetValue('CLE_QUALIFPIECE') + '"';

    lQuery := OpenSQL(SQL, True);
    try
      if not lQuery.Eof then
        Result := FTobEcrSelect.SelectDB('', lQuery );

    finally
      Ferme( lQuery );
    end;
  end
   else
    result := true ; // Fb 21264/21266

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/02/2003
Modifié le ... :   /  /
Description .. : Test si l'écriture est une écriture de centralisation ou une
Suite ........ : écriture de la base de données
Mots clefs ... :
*****************************************************************}
function TOF_CONSECR.EstVraiEcr(vTobEcr: TOB): boolean;
begin
  Result := (vTobEcr <> nil) and (vTobEcr.GetInteger('E_NUMEROPIECE') <> 0);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/02/2003
Modifié le ... :   /  /
Description .. : Recherche la position de certains champs de la liste pour
Suite ........ : les modifications d'affichage
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.IndiceColFListe;
var
  lStChamps, lStListeChamps: string;
  lInIndex: integer;
begin
  lStListeChamps := FStListeChamps;
  lInIndex := 1;
  while lStListeChamps <> '' do
  begin
    lStChamps := READTOKENST(lStListeChamps);
    if (lStChamps = 'E_DEBIT') then
      FColDebit := lInIndex
    else if (lStChamps = 'E_CREDIT') then
      FColCredit := lInIndex
    else if (lStChamps = 'SOLDEE') then
      FColSoldePro := lInIndex
    else if (lStChamps = 'E_GENERAL') then
      FColGeneral := lInIndex
    else if (lStChamps = 'E_AUXILIAIRE') then
      FColAuxiliaire := lInIndex
    else if (lStChamps = 'E_LETTRAGE') then
      FColLettrage := lInIndex
    else if (lStChamps = 'E_REFPOINTAGE') then
      FColRefPointage := lInIndex;

    Inc(lInIndex);
  end; // while
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/01/2004
Modifié le ... : 18/02/2004
Description .. : Fonction de préparation des paramètres passés au GL
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
Procedure TOF_CONSECR.PrepareArgumentGL(vBoJustifSolde: Boolean);
var lBoOkAL : Boolean;
    lBoOkPL : Boolean;
    lBoOkTL : Boolean;
begin
  // Exercice
  ACritEdt.CritEdt.Exo.Code := CRelatifVersExercice(E_Exercice.Value);

  // Début de la Date Comptable (1)
  ACritEdt.CritEdt.Date1 := StrToDate(E_DateComptable.Text);

  // Fin de la Datecomptable (2)
  ACritEdt.CritEdt.Date2 := StrToDate(E_DateComptable_.Text);

  // Numéro de Pièce (1)
  if E_NumeroPiece.Text <> '' then
    ACritEdt.CritEdt.GL.NumPiece1 := StrToInt(E_NumeroPiece.Text);
  //else
  //  ACritEdt.CritEdt.GL.NumPiece1 := 0;

  // Numéro de Pièce (2)
  if E_NumeroPiece_.Text <> '' then
    ACritEdt.CritEdt.GL.NumPiece2 := StrToInt(E_NumeroPiece_.Text);
  //else
  //  ACritEdt.CritEdt.GL.NumPiece2 := 999999999;

  // Justificatif de Solde
  ACritEdt.CritEdt.GL.EnDateSituation := vBoJustifSolde;

  if not vBoJustifSolde then
  begin
    // A-Nouveaux détaillés, Pas de besoin si justif de solde, c'est automatique
    ACritEdt.CritEdt.GL.DetailAno := CbDetailAno.Checked;
    // GCO - 26/12/2005 - FQ 11773
    ACritEdt.CritEdt.GL.AvecDetailCentralise := CbCentralisation.Checked;
  end;

  // Référence Interne
  //ACritEdt.CritEdt.RefInterne := E_REFINTERNE.Text;

  // Type d'Ecritures
  ACritEdt.CritEdt.Qualifpiece := EQUALIFPIECE.Text;

  // Etablissement
  ACritEdt.CritEdt.Etab := E_ETABLISSEMENT.Value;

  // Devise
  ACritEdt.CritEdt.DeviseSelect := E_DEVISE.Value;

  // Ecritures Valides
  if E_Valide.State = cbGrayed then
    ACritEdt.CritEdt.Valide := ''
  else
    ACritEdt.CritEdt.Valide := IIF( E_Valide.Checked, 'OUI', 'NON' );

  // Ecritures Lettrées
  lBoOkAL := Pos('AL', E_EtatLettrage.Text) > 0;
  lBoOkPL := Pos('PL', E_EtatLettrage.Text) > 0;
  lBoOkTL := Pos('TL', E_EtatLettrage.Text) > 0;

  // Ecritures Lettrées ( 0 = Grayed, 1 = not Checked sinon 2 = Checked
  if (E_EtatLettrage.Text = '') or
     (E_EtatLettrage.Tous) or
     ((lBoOkAL and (lBoOkPL or lBoOkTL))) then
  begin
    ACritEdt.CritEdt.GL.Lettrable := 0;
  end
  else
  begin
    if lBoOkAL then
      ACritEdt.CritEdt.GL.Lettrable := 1
    else
      ACritEdt.CritEdt.GL.Lettrable := 2;
  end;
 //   else
 //     if (Pos('PL;', E_EtatLettrage.Text) > 0) or (Pos('TL;', E_EtatLettrage.Text) > 0) then
 //       ACritEdt.CritEdt.GL.Lettrable := 2;
 // end;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.ApresChangementFiltre;
var lExoDate : TExoDate;
begin
  inherited;
  ChargeGeneral( False );
  ChargeAuxiliaire ( False );
  // GCO - 02/01/2008 - FQ 19533
  if FFiltres.Value = '' then
    CExoRelatifToDates(E_EXERCICE.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_, True)
  else
  begin // GCO - 30/01/2008 - FQ 22318
    lExoDate := CtxExercice.QuelExoDate(CRelatifVersExercice(E_Exercice.Value));
    if StrToDate(E_DateComptable_.Text) < lExoDate.Deb then
      CExoRelatifToDates(E_EXERCICE.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_, True);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/08/2007
Modifié le ... :   /  /    
Description .. : GCO - 29/08/2008 - FQ 21288 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CONSECR.NouveauFiltre;
begin
  inherited;
  // GCO - 26/09/2007 - FQ 21288
  E_Exercice.Itemindex := 0;
  OnChangeComboE_Exercice(nil);
end;

////////////////////////////////////////////////////////////////////////////////
function TOF_CONSECR.GetPrecedent : TExoDate;
begin
{$IFDEF BUREAU}
  Result := FZExercice.Precedent ;
{$ELSE}
  Result := VH^.Precedent;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
function TOF_CONSECR.GetEnCours : TExoDate;
begin
{$IFDEF BUREAU}
  Result := FZExercice.Encours ;
{$ELSE}
  Result := VH^.Encours;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
function TOF_CONSECR.GetSuivant : TExoDate;
begin
{$IFDEF BUREAU}
  Result := FZExercice.Suivant ;
{$ELSE}
  Result := VH^.Suivant;
{$ENDIF}
end ;

////////////////////////////////////////////////////////////////////////////////
function TOF_CONSECR.GetExercices( vValue : integer ) : TExoDate;
begin
{$IFDEF BUREAU}
  Result := FZExercice.Exercices[vValue-1];
{$ELSE}
  Result := VH^.Exercices[vValue];
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
function TOF_CONSECR.GetExoV8 : TExoDate;
begin
{$IFDEF BUREAU}
  Result := FZExercice.ExoV8;
{$ELSE}
  Result := VH^.Exov8;
{$ENDIF}
end ;

////////////////////////////////////////////////////////////////////////////////
function TOF_CONSECR.GetExoRef : TExoDate;
begin
{$IFDEF BUREAU}
  Result := FZExercice.CPExoRef;
{$ELSE}
  Result := VH^.CPExoRef;
{$ENDIF}
end ;

////////////////////////////////////////////////////////////////////////////////
function TOF_CONSECR.GetEntree    : TExoDate;
begin
{$IFDEF BUREAU}
  Result := FZExercice.Entree;
{$ELSE}
  Result := VH^.Entree;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
function TOF_CONSECR.GetInfoCpta( vValue : TFichierBase ) : TInfoCpta;
begin
{$IFDEF BUREAU}
  Result := FZInfoCpta.Cpta[vValue];
{$ELSE}
  Result := VH^.CPta[vValue];
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.AnalyseParametre( vStArg : String );
var lStArg : string;
    lSt : string;
    lStChampsTries : string;
begin
  lStArg := vStArg;

  // Compte Général
  lSt := ReadTokenSt(lStArg);
  if lSt <> '' then
    E_General.Text  := lSt;

  // Exercice
  lSt := ReadTokenSt(lStArg);
  if lSt <> '' then
  begin
    if lSt = '-2' then
    begin
      E_Exercice.ItemIndex := 0;
    end
    else
    begin
      // lSt vaut 0,1 ou -1 si on vient de Excel
      if lSt = '0' then
        lSt := GetEncours.Code
      else
        if lSt = '1' then
          lSt := GetSuivant.Code
        else
          if lSt = '-1' then
            lSt := GetPrecedent.Code;

      E_Exercice.Value := CExerciceVersRelatif( lSt );
    end;
  end;
  //else
  //  E_Exercice.ItemIndex   := -1;

  // Champs de Tri
  lSt := ReadTokenSt(lStArg);
  if lSt <> '' then
    lStChampsTries := lSt;

  // Auxiliaire
  lSt := ReadTokenSt(lStArg);
  if lSt <> '' then
    E_Auxiliaire.Text := lSt;

  // Appel depuis la saisie
  if ReadTokenSt(lStArg) = '-' then
  begin
    FFromSaisie := False;
    FBoLettrageEnSaisie := False;
  end
  else
  begin
    FFromSaisie := True;
    // Autorisation du lettrage en saisie
    FBoLettrageEnSaisie := True;
  end;

{$IFDEF CCSTD}
  // GCO - 08/06/2006 - FQ 18336 - Pas de controle de verrou, car c'est du
  // paramétrage d'écran, de plus on pose un blocagemonoutilisateur à l'ouverture
  // du CCSTD, donc on serait en mode consultation systématiquement.
{$ELSE}
  // On force la consultation si on vient de la saisie Borderau ou si
  // présence d'un verrou de la saisie dans la table COURRIER
  if (FFromSaisie) or CEstSaisieOuverte(False) then
  begin
    ComboAcces.Value   := '0';
    ComboAcces.Enabled := False;
    AFiltreDisabled    := True;
    FOkCreateModif     := False;
  end;
{$ENDIF}

  // Renvoi des écritures dans GTobDetailECr
  if (ReadTokenSt(lStArg) = 'X') then
  begin
    SetControlProperty('BVALIDER', 'MODALRESULT', '1');
    // Réaffectation du BValider car on doit renvoyer la liste des écritures
    BValider.OnClick := OnClickBValider;
  end
  else
    SetControlProperty('BVALIDER', 'MODALRESULT', '0');

  // Nom de la base sur laquelle on doit faire la requête SELECT
  FStNomBase := ReadTokenSt(lStArg);
  if FStNomBase <> '' then
  begin
  {$IFDEF BUREAU}
    FZExercice := TZExercice.Create(False, FStNomBase);
    FZInfoCpta := TZInfoCpta.Create(FStNomBase);
  {$ENDIF}
    FStNomBase := FStNomBase + '.dbo.';
  end;

  // Ouverture en Justificatif de Solde
  lSt := ReadTokenSt(lStArg);
  if lSt <> '' then
    FBoOuvreEnJustif := (lSt = 'X')
  else
    FBoOuvreEnJustif := False;

    // Mode de sélection des comptes
  lSt := ReadTokenSt(lStArg);
  if lSt <> '' then
    ComboDefilCpt.Value := lSt;
end;

////////////////////////////////////////////////////////////////////////////////
{ TOF_PARAMCONS }
{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 25/07/2005
Modifié le ... :   /  /
Description .. : - 25/07/2005 - FB 16252 - suppression de la gestion du
Suite ........ : borderstyle, plante a l'ouverture
Mots clefs ... :
*****************************************************************}
procedure TOF_PARAMCONS.OnArgument(S: String);
begin
  inherited ;

  FStGeneral                                          :=  S ;
  Ecran.OnKeyDown                                     := FormKeyDown ;
  TToolbarButton97(GetControl('BValider')).Onclick    := BValiderClick;

  E_General    := THEdit(GetControl('E_GENERAL', True));
  E_Auxiliaire := THEdit(GetControl('E_AUXILIAIRE', True));

  E_General.OnKeyDown         := GeneralKeyDown;
  E_General.OnElipsisClick    := GeneralElipsisClick;
  E_General.OnExit            := GeneralExit;

  E_Auxiliaire.OnKeyDown      := AuxKeyDown;
  E_Auxiliaire.OnElipsisClick := AuxElipsisClick;

  TTabSheet(GetControl('P1')).TabVisible         := false ;
  TPageControl(GetControl('PG')).ActivePageIndex := 0 ;

  E_General.MaxLength           := GetInfoCpta(fbGene).Lg;
  E_Auxiliaire.MaxLength        := GetInfoCpta(fbAux).Lg;

 SetControlEnabled('E_AUXILIAIRE', true ) ;
 SetControlEnabled('AUXILIAIRE'  , true ) ;


  if E_General.CanFocus then
    E_General.SetFocus;

  Ecran.Caption  := TraduireMemoire('Choix des comptes') ;
  
end;


procedure TOF_PARAMCONS.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 if ( csDestroying in Ecran.ComponentState ) then Exit ;
 case Key of

  VK_F10 : begin
            Key := 0 ;
            TFVierge(Ecran).retour := GetControlText('E_GENERAL') + ';' + GetControlText('E_AUXILIAIRE') ;
            Ecran.Close;
           end;
 end; // case
end;

procedure TOF_PARAMCONS.OnClose;
begin
  inherited;
end;


procedure TOF_PARAMCONS.GeneralKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
begin
 if ( csDestroying in Ecran.ComponentState ) then Exit ;

 case Key of
  VK_F5 : GeneralElipsisClick(Sender);
 end; // case

end;

procedure TOF_PARAMCONS.AuxKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
begin
 if ( csDestroying in Ecran.ComponentState ) then Exit ;

 case Key of
  VK_F5 : AuxElipsisClick(Sender);
 end; // case

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 13/02/2006
Modifié le ... : 23/07/2007
Description .. : - LG - 16/10/2006 - FB 16403 - on autorise la selection du 
Suite ........ : meme compte pour e chg d'auxi
Suite ........ : - LG - 24/10/2006 - FB 16816 - on ne selectionne plus les 
Suite ........ : comptes visés ( on ne tient pas compte de l'exercice en 
Suite ........ : cours )
Suite ........ : - LG - 23/04/2007 - FB 20952 - on enleve le lookUpLocate 
Suite ........ : pour tenir compte du parametre 'toujours ouvrir les lookUp'
Mots clefs ... : 
*****************************************************************}
procedure TOF_PARAMCONS.GeneralElipsisClick(Sender: TObject);
var
 lStSelect     : string ;
 lStColonne    : string ;
 lStOrder      : string ;
 lStWhere      : string ;
// lStGen        : string ;
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit ;

 {if V_PGI.LookUpLocate then
  begin
   lStGen := E_General.Text;
   if Length(lStGen) = VH^.Cpta[fbGene].Lg then exit ;
  end ;}

 CMakeSQLLookupGen(lStWhere,lStColonne,lStOrder,lStSelect) ;
 lStWhere := lStWhere + ' AND G_VENTILABLE="-" AND G_VISAREVISION<>"X" ' ; //AND G_GENERAL<>"' + FStGeneral + '" ' ;
 LookupList(E_General,TraduireMemoire('Comptes'),'GENERAUX',lStColonne,lStSelect,lStWhere,lStOrder,true, 1) ;

end;

procedure TOF_PARAMCONS.AuxElipsisClick(Sender: TObject);
var
 lStSelect     : string ;
 lStColonne    : string ;
 lStOrder      : string ;
 lStWhere      : string ;
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit ;

 CMakeSQLLookupAux(lStWhere,lStColonne,lStOrder,lStSelect,'OD',FStNatureGene) ;
 LookupList(THEdit(GetControl('E_AUXILIAIRE')),TraduireMemoire('Auxiliaire'),'TIERS',lStColonne,lStSelect,lStWhere,lStOrder,true, 2,'',tlLocate) ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : LG
Créé le ...... : 24/05/2007
Modifié le ... :   /  /    
Description .. : - LG - FB 19314 - 24/05/2007 - on empecher de valider 
Suite ........ : avec un compte collectif
Mots clefs ... : 
*****************************************************************}
procedure TOF_PARAMCONS.BValiderClick(Sender: TObject);
begin
 GeneralExit(nil) ;
 if trim(E_General.Text) = '' then
  begin
   Ecran.ModalResult := mrNone ;
   exit ;
  end ;
 TFVierge(Ecran).retour := GetControlText('E_GENERAL') + ';' + GetControlText('E_AUXILIAIRE') ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... :   /  /    
Modifié le ... : 24/05/2007
Description .. : GCO - 14/09/2006 - FQ 18764
Suite ........ : - LG - 16/10/2006 - FQ 16403 - la correction presendente 
Suite ........ : ne fct pas
Suite ........ : - LG - 24/10/2006 - FB 16816 - on ne selectionne plus les 
Suite ........ : comptes visés ( on ne tien pas compte de l'exercice en 
Suite ........ : cours )
Suite ........ : - LG - 23/05/2007 - FB 16816 - ajout d'un msg d'erreur
Suite ........ : - LG - 24/05/2007 - FB 19293 - ouverture d'un lookup en 
Suite ........ : cas d'erreur
Mots clefs ... : 
*****************************************************************}
procedure TOF_PARAMCONS.GeneralExit(Sender: TObject);
var
 lQ            : TQuery ;
 lBoCollectif  : boolean ;
 lStSelect     : string ;
 lStColonne    : string ;
 lStOrder      : string ;
 lStWhere      : string ;
 lStGen        : string ;
 lBoVisa       : boolean ;
begin

 if E_General.Text = '' then Exit;

 lStGen            := E_General.Text ;
 lStGen            := BourreEtLess(lStGen, fbGene);
 lBoCollectif      := false ;

 CMakeSQLLookupGen(lStWhere,lStColonne,lStOrder,lStSelect) ;
 lStWhere := ' AND ' + lStWhere + ' AND G_VENTILABLE="-" ' ; //AND G_GENERAL<>"' + FStGeneral + '" ' ;

 lQ      := OpenSQL('SELECT G_COLLECTIF,G_NATUREGENE,G_VISAREVISION FROM GENERAUX WHERE G_GENERAL="' + lStGen + '" ' + lStWhere , true ) ;

 if not lQ.Eof then
 begin
   lBoVisa        := lQ.FindField('G_VISAREVISION').asString = 'X' ;
   if lBoVisa then
    begin
     E_General.Text := '' ;
     PGIInfo('Reclassement impossible: le compte de destination est visé') ;
     SetFocusControl('E_GENERAL') ;
     GeneralElipsisClick(nil) ;
    end
     else
      begin
       lBoCollectif   := lQ.FindField('G_COLLECTIF').asString = 'X' ;
       FStNatureGene  := lQ.FindField('G_NATUREGENE').asString ;
       E_General.Text := lStGen ;
      end ;
 end
  else
   begin
    E_General.Text := '' ;
    SetFocusControl('E_GENERAL') ;
    GeneralElipsisClick(nil) ;
   end ;


 Ferme(lQ) ;

 SetControlEnabled('E_AUXILIAIRE', lBoCollectif) ;
 SetControlEnabled('AUXILIAIRE'  , lBoCollectif) ;

 if not lBoCollectif then
  SetControlText('E_AUXILIAIRE','')
   else
    SetFocusControl('E_AUXILIAIRE') ;

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 16/11/2006
Modifié le ... :   /  /
Description .. : copie des éléments des lignes d'écriture sélectionnées afin
Suite ........ : de les 'coller' vers la note de travail
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.OnClickEcrVersClipboard( Sender : TObject );
var
  LgFiche : integer;
  StrLigne : String;

  //fonction interne pour création ligne
  Function RecupData(LaLigne : integer) : String;
  var
    SL : String;
  begin
    RecupEcrSelect(LaLigne);
    if Trim(FTobEcrSelect.GetValue('E_EXERCICE')) <> '' then
    begin
      SL := '"' + FTobEcrSelect.GetValue('E_EXERCICE') + '";"' +
            FTobEcrSelect.GetValue('E_GENERAL') + '";"' +
            DateToText(FTobEcrSelect.GetValue('E_DATECOMPTABLE')) + '";"' +
            FTobEcrSelect.GetValue('E_LIBELLE') + '";';
      if Valeur(FTobEcrSelect.GetValue('E_DEBIT')) <> 0 then
        SL := SL + FloatToStr(FTobEcrSelect.GetValue('E_DEBIT')) + '";"D"'
      else
        SL := SL + FloatToStr(FTobEcrSelect.GetValue('E_CREDIT')) + '";"C"';
      SL := SL + chr(13) + chr(10);
    end
    else
      SL := '';

    Result := SL;
 end;

 // gestion du copier
begin
  Clipboard.open;
  Clipboard.Clear;
  StrLigne := '';

  // si aucune sélection
  if not (FListe.AllSelected) and (FListe.nbSelected = 0) then
  begin
    StrLigne := StrLigne + RecupData(FListe.Row);
  end
  else if FListe.nbSelected > 0 then
    begin
      for LgFiche := 0 to FListe.nbSelected -1 do
      begin
       FListe.GotoLeBookmark(LgFiche);
       StrLigne := StrLigne + RecupData(FListe.Row);
      end;
    end
    else
    begin
      for LgFiche := 1 to FListe.RowCount -1 do
      begin
        StrLigne := StrLigne + RecupData(LgFiche);
      end;
    end;

  Clipboard.SetTextBuf(Pchar(StrLigne));
  Clipboard.Close;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 27/03/2007
Modifié le ... :   /  /
Description .. : Répartition analytique
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CONSECR.RepartAnalytiquesClick(Sender: TObject);
var lAction : TActionFiche;
begin
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  ACritEdt.CritEdt.Date1 := StrToDate(E_DATECOMPTABLE.Text);
  ACritEdt.CritEdt.Date2 := StrToDate(E_DATECOMPTABLE_.Text);

  ACritEdt.CritEdt.DateDeb := ACritEdt.CritEdt.Date1;
  ACritEdt.CritEdt.DateFin := ACritEdt.CritEdt.Date2;

  ACritEdt.CritEdt.Exo := PositionneExo;
  ACritEdt.CritEdt.SCpt1 := E_GENERAL.Text;
  lAction := taModif;
  if not FOkCreateModif then
    lAction := taConsult;
  TheData := ACritEdt;
  CPLanceFiche_RepartitionAnalytique(lAction, E_GENERAL.Text);
  TheData := nil;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 11/06/2007
Modifié le ... :   /  /    
Description .. : Modification de la ventilation analytique
Mots clefs ... :
*****************************************************************}
procedure TOF_CONSECR.VentilAnalytiqueAxe (pstAxe : string );
var
  stSQL : string;
  lQAna : TQuery;
  TA : TActionFiche;
begin
  if ABoMultiSelected then Exit;
  if (not FListe.Focused) or (not RecupEcrSelect) then Exit;

  stSQL := 'SELECT * FROM ' + FStNomBase + 'ANALYTIQ WHERE ' +
    '(Y_JOURNAL = "' + FTobEcrSelect.GetValue('E_JOURNAL') + '") AND ' +
    '(Y_EXERCICE = "' + FTobEcrSelect.GetValue('E_EXERCICE') + '") AND ' +
    '(Y_DATECOMPTABLE = "' + USDateTime(FTobEcrSelect.GetValue('E_DATECOMPTABLE')) + '") AND ' +
    '(Y_NUMEROPIECE = ' + IntToStr(FTobEcrSelect.GetValue('E_NUMEROPIECE')) + ') AND ' +
    '(Y_NUMLIGNE = ' + IntToStr(FTobEcrSelect.GetValue('E_NUMLIGNE')) + ') AND ' +
    '(Y_AXE = "'+pstAxe+'") AND ' +
    '(Y_QualifPiece = "' + FTobEcrSelect.GetValue('E_QUALIFPIECE') + '")';

  lQAna := OpenSQL (stSQL,True);
  try
    if (not FOkCreateModif)  then
        TA := taConsult
    else TA := taModif;
    TrouveEtLanceSaisieODA(lQAna, TA);
  finally
    Ferme (lQAna);
  end;
end;

procedure TOF_CONSECR.VentilAnalytiques1Click(Sender: TObject);
begin
  VentilAnalytiqueAxe ('A1');
end;

procedure TOF_CONSECR.VentilAnalytiques2Click(Sender: TObject);
begin
  VentilAnalytiqueAxe ('A2');
end;

procedure TOF_CONSECR.VentilAnalytiques3Click(Sender: TObject);
begin
  VentilAnalytiqueAxe ('A3');
end;

procedure TOF_CONSECR.VentilAnalytiques4Click(Sender: TObject);
begin
  VentilAnalytiqueAxe ('A4');
end;

procedure TOF_CONSECR.VentilAnalytiques5Click(Sender: TObject);
begin
  VentilAnalytiqueAxe ('A5');
end;
{$ENDIF}

{JP 26/06/07 : Vérifie si la sélection contient des lignes GC ou Tréso}
function TOF_CONSECR.HasEcrImport : Boolean;
var
  n : Integer;
  F : TOB;
begin
  Result := False;
  for n := 1 to FListe.RowCount - 1 do begin
    if FListe.IsSelected(n) then begin
      F := GetO(FListe, n);
      if ((F.GetString('E_REFGESCOM') <> '') and
          (F.Getstring('E_TYPEMVT') = 'TTC')) or
          (F.Getstring('E_QUALIFORIGINE') = QUALIFTRESO) then begin
         Result := True;
         Break;
       end;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/08/2007
Modifié le ... :   /  /    
Description .. : GCO - 29/08/2007 - FQ 21305  
Mots clefs ... :
*****************************************************************}
function TOF_CONSECR.CGenereSQLNatureAuxi: string;
begin
  Result := '';

  // JP - 02/11/05 - FQ 16853 : On exclut les tiers spécifiques à la Gescom
  if E_Auxiliaire.DataType = 'TZTTOUS' then
    Result := '(T_NATUREAUXI <> "NCP" AND T_NATUREAUXI <> "CON" AND T_NATUREAUXI <> "PRO" AND T_NATUREAUXI <> "SUS")'
  else
  if E_Auxiliaire.DataType = 'TZTTOUTDEBIT' then
    Result := '(T_NATUREAUXI = "AUD" OR T_NATUREAUXI = "CLI" OR T_NATUREAUXI = "DIV")'
  else
  if E_Auxiliaire.DataType = 'TZTTOUTCREDIT' then
    Result := '(T_NATUREAUXI = "AUC" OR T_NATUREAUXI = "FOU" OR T_NATUREAUXI = "DIV")'
  else
  if E_Auxiliaire.DataType = 'TZTSALARIE' then
    Result := 'T_NATUREAUXI = "SAL"';

end;

////////////////////////////////////////////////////////////////////////////////

initialization
  registerclasses([TOF_CONSECR]);
  registerclasses([TOF_PARAMCONS]);
end.

