{-------------------------------------------------------------------------------------
  Version    |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
08.01.001.009  05/02/07  JP   Création de la nouvelle fiche de pointage
08.00.001.012  25/04/07  JP   FQ 20067 : Ne pas éclater le solde comptable
08.00.001.012  25/04/07  JP   FQ 20063 : Le calcul du reste à pointer prenait les écritures dans le sens opposé.
08.00.001.012  25/04/07  JP   FQ 20065 : Problème à la création d'écritures en pointage Manuel : on sortait de la fonction
08.00.001.012  25/04/07  JP   FQ 20066 : Mauvaise gestion des devises lors de l'affichage de la grille
08.00.001.013  03/05/07  JP   FQ 20197 : Corrections d'une anomalie en combinatoire
08.00.001.013  03/05/07  JP   FQ TRESO 10449 : Il faut ouvrir la possibilité de pointé la TVA.
08.00.001.014  11/05/07  JP   FQ TRESO 10461 : Vocabulaire + Orthographe
08.00.001.015  11/05/07  JP   FQ 20284 : Ajout des boutons export et tout sélectionner
08.00.001.018  01/06/07  JP   FQ TRESO 10471 : Gestion du concept sur la visualisation des opérations bancaires
08.00.001.018  05/06/07  JP   FQ 20066 suite : il restait à règler le problème des soldes
08.00.001.022  26/06/07  JP   FQ TRESO 10492 : Ajout d'une zone de gestion de la colonne référence pour opérations bancaires
08.00.001.024  09/07/07  JP   Amélioration de l'ergonomie de FQ TRESO 10492
08.00.001.025  12/07/07  JP   FQ 21043 : En pointage manuel, on cache la colonne Pointé ('0')
08.00.001.025  12/07/07  JP   FQ 21044 : On donne le focus à la grille
08.00.001.025  12/07/07  JP   FQ 21046 : Gestion de la touche F12
08.00.001.025  12/07/07  JP   FQ 21057 : Ajout du code à côté du libellé dans le caption de la fiche
08.00.001.025  12/07/07  JP   FQ 21054 : Lors de la création d'un mouvement bancaire, on ne le pointe pas par défaut
08.00.001.025  12/07/07  JP   FQ 21052 : accès à la Tom de EEXBQLIG en lecture seule
08.00.001.025  12/07/07  JP   FQ 21048 : on limite la modification d'entête de pièce que pour la saisie normal
08.00.001.025  16/07/07  JP   FQ 21056 : gestion des mvts dont la date de pointage est postérieure à la session en cours
08.00.001.025  16/07/07  JP   FQ 21049 : Refonte de la gestion des boutons sur le modèle du PopF11
08.00.001.025  16/07/07  JP   FQ TRESO 10507 : Commande d'accès à la tom TRECRITURE est fausse
08.00.001.025  17/07/07  JP   FQ 21098 : Faute d'orthographe
08.00.001.025  17/07/07  JP   FQ 21100 : Branchement du Ctrl + F
08.00.001.025  17/07/07  JP   FQ 21107 : Déplacement en boucle dans la grille
08.00.001.025  18/07/07  JP   FQ 21106 : Point dans le DBIndicator si la ligne est pointée
08.00.001.025  18/07/07  JP   FQ 21103 : Dans certains cas les lignes restaient en italique après dépointage
08.00.002.001  18/07/07  JP   FQ 21130 : Raccourci pour le rapprochement automatique (002.001 à cause de la BOB)
08.00.001.025  18/07/07  JP   FQ 21127 : On ne lance pas la suppression sur une pièce bordereau
08.00.001.025  18/07/07  JP   FQ 21126 : Interdiction de créer un mouvement bancaire si on n'est pas en pointage manuel
08.00.001.025  18/07/07  JP   FQ 21121 : Le code rappro centré à gauche plutôt qu'à droite
08.00.001.025  18/07/07  JP   FQ 21132 : On propose l'option Filtre sur les mouvements bancaires en Pointage manuel
08.00.001.025  19/07/07  JP   FQ 21125 : Accès à la saisie simplifiée depuis un mouvement saisi manuellement
08.00.002.001  20/07/07  JP   FQ 20601 : Positionnement dans la saisie paramétrable sur la bonne ligne
08.00.002.001  20/07/07  JP   FQ 21101 : Accès à la saisie ou mouvement est fonction de la ligne en cours, par du raccourci
08.00.002.001  23/07/07  JP   FQ 21053 et 21058 : Gestion des messages à la validation
08.00.002.001  23/07/07  JP   Découplage du reste à pointer avec l'équilibre du pointage cf. #ETAT#
08.00.002.002  30/07/07  JP   FQ 21043 : Mauvaise compréhension de la FQ : il ne s'agit pas de coché la colonne
                              pointé, mais ne rien mettre dans les cellules de montant à 0,00
08.00.002.002  30/07/07  JP   FQ 21192 : on rend actif la saisie simplifiée que si l'on est sur un mouvement bancaire
08.00.002.002  31/07/07  JP   FQ 21125 : sur les mouvements manuels, on crée l'écriture dans le sens du mouvement
08.00.002.002  31/07/07  JP   FQ 21199 : Renforcement des contrôles dans IsTobAPointe en restour de saisie
08.00.002.002  01/08/07  JP   FQ 21195 : Message en pointage sécurisé
08.00.002.002  01/08/07  JP   FQ 21201 : Gestion de l'état des filtres sur les types de mouvements à afficher
08.10.001.004  13/08/07  SBO  FQ 21191 : Ouverture de la saisie paramétrable : chargement préalable des bordereaux + mise en place des verrous
08.10.001.004  16/08/07 JP/SB FQ 21246 : suppression en saisie paramétrable de la saisie en boucle => on peut tester le retour de la fiche  
08.10.001.004  16/08/07  JP   FQ 13017 : En pointage sur journal, possibilité de voir les écritures passées sur un autre journal
08.10.001.005  17/08/07  JP   FQ 20196 : autorisation du rapprochement multi-relevés
08.10.001.005  17/08/07  JP   FQ 21255 : FGestionCibOk : si aucun mode de paiement
08.10.001.010  20/09/07  JP   FQ 21318 : si on ne valide pas le pointage courant, on ouvre la saisie en lecture seule
08.10.001.010  18/09/07  JP   FQ 21443 : On passe à la ligne suivante quand on a pointé manuellement une écriture
08.10.001.010  18/09/07  JP   FQ 21393 : corrections sur la gestion Libellés/code CIB
08.10.001.010  18/09/07  JP   FQ 10518 : l'unidirectionnel est incompatible avec les RecordCount
08.10.001.010  18/09/07  JP   FQ 21333 : Fenêtre de choix d'une date comptable
08.10.005.001  14/11/07  JP   Gestion des comptes pointables qui ne sont pas bancaires
08.10.006.001  28/11/07  JP   SIC : essai d'amélioration des performances en chargement de la session de pointage  
--------------------------------------------------------------------------------------}
unit CPPOINTAGERAP_TOF;

interface

uses
  Controls, Classes, Vierge,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  LettAuto, Spin, ULibPieceCompta, StdCtrls, Graphics, HTB97, Dialogs,
  UObjFiltres, ULibPointage, Grids, SysUtils, HCtrls, UTob, HEnt1, UTOF, Menus;

type
  TOF_CPPOINTAGERAP = class(TOF)
    FListe        : THGrid;
    POPF11        : TPopupMenu;
    AFindDialog   : TFindDialog;

    procedure OnArgument(S : string); override;
    procedure OnClose               ; override;
    procedure OnUpdate              ; override;
    procedure OnLoad                ; override;
  private
    TobReleve : TOB;
    TobCIB    : TOB;
    TobRegles : TOB;
    TobLibCib : TOB;

    ToleranceDate : Integer; {Nombre de jours de tolérance choisi pour les dates}
    NbDecDev      : Integer;
    DecCompta     : Integer;
    DecReleve     : Integer;
    TobBqeEnCours : TOB;
    TobCptEnCours : TOB;
    EtatAffichage : TTypeVisu;
    ObjFiltre     : TObjFiltre;

    PointageAuto   : Boolean;
    FFindFirst     : Boolean;
    FPeutFermer    : Boolean;
    FMajDateValeur : Boolean;
    FPasMessage    : Boolean; {23/07/07 : FQ 21053}
    PointeAvecListeNonComplete : Boolean;
    FDateAno       : TDateTime; {28/11/07}
    FExoAno        : string;    {28/11/07}
    FSaufVVert     : Boolean;   {28/11/07}

    function EcritureSelectionnee   : Boolean;
    function GetDateAno             : TDateTime; {28/11/07}
    function GetExoAno              : string;    {28/11/07}
    function GetMntCoherence        : Double; {#ETAT#}
    function GetModeAffichage       : Boolean;
    function ControleAvantRappro    : Boolean;
    function ControleContrepartie   : Boolean;
    function ControleDevise         : Boolean;
    function ControleRefPointage    : Boolean;
    function MajPointage            : Boolean;
    function IsPointageEncours      : Boolean; {Regarde s'il y a un pointage en cours}
    function IsPointageManuelOk     : Boolean; {En Rappro, on regarde si le pointage manuel est équilibré}
    function IsDateValeurOk         : Boolean; {Si avec Tréso, on avertit si les date de valeur sont <> entre le relevé et les écritures}
    function MessageValidation      (AnnuleOk : Boolean) : Boolean; {Contrôles et confirmations de l'enregistrement du pointage}
    function GetMaxNumPtge          : string;
    function RecupAutreWhere        : string;
    function CalcPointageManuel     : Double;
    function GetLeModeRglt          (var aTob : TOB) : string;
    {$IFNDEF TRESO}
    function GetDateComptable       : TDateTime;
    function GetModePaie            (Cib, Sens : string) : string;
    {$ENDIF TRESO}
    function ConstruitListe         (var LM : T_D; var LP, LB : T_I; MM : Double) : Integer;
    function GetLibCaption          (Chp : string) : string;
    function IsTobAPointe           (TobEcr : TOB) : Boolean; {Vérifie si l'écriture correspond au mouvement bancaire en cours}

    procedure GetPieceAutreJournal   ; {FQ 13017 : en pointage sur journal, affiche les pièces d'un autre journal}
    procedure GereFocusF12           ; {FQ 21046 : Basculement pagecntrol / Grille}
    procedure AnnulePointage         ; {Repasse le champ 'MODIFIE' de la tob à False}
    procedure ReinitialisePointage   ;
    procedure RemetRefPointageABlanc ;
    procedure ChargeGrille           ;
    procedure ChargerMouvements      ;
    procedure PresenteGrille         ;
    procedure ChargeLesCIB           ;
    procedure SelectionJour          ;
    procedure SelectionMois          ;
    procedure SelectionRapide        ;
    procedure LancerSaisie           ;
    procedure LanceLePointageAuto    ;
    procedure MajImage               ; {23/07/07 : Mise à jour de l'image de l'état de pointage #ETAT#}
    procedure MajAvancement          ; {Mise à jour de EE_AVANCEMENT}
    procedure SetTotauxPointes       ; {Mise à jour des totaux pointés}
    procedure InitSoldes             ; {Initialisation des soldes bancaires et des écritures}
    procedure MajTotalisation        ; {Mise à jour des deux pavés de compteurs et totalisations au bas de l'écran}
    procedure RegroupementManuel     ; {En pointage manuel, si la sélection est équilibré, on propose un nouveau Code de regroupement}
    procedure MetEnModeLecture       (SaufVVert : Boolean = False); {Grise les menus et boutons dont l'objet est incompatible avec la lecture}
    procedure RafraichitFiche        (AvecMaj : Byte = 1);{Traitement du BChercheClik}
    procedure MajDtValeur            (aTob : TOB);
    procedure LanceFicheEexBqLig     (Action : TActionfiche);
    procedure MajSoldeManuel         (EnPointage : Boolean = False); {Mise à jour du liebllé de suivi du pointage manuel}
    procedure EcritPointage          (AvecMsg : Boolean = True); {Avec message à la fin du traitement}
    procedure DepointeGroupe         (Value : string);
    procedure AjouteMvt              (var aTob : TOB);
    procedure CocheDecoche           (ARow : Integer; Value : string);
    procedure PointageManuel         (ARow : Integer);
    procedure RemplitLigne           (ARow : Integer; aTob : TOB);
    procedure ShowLignes             (aTyp : TTypeVisu);
    procedure AligneControleSurGrille(vCol : Integer; vNomControl : string; vAvecWidth, RelatifOk : Boolean);
    procedure DrawDBIndicator        (Tip : Byte; ACol, ARow : Integer); {18/07/06 : FQ 21106}
    function  ConvertitTob           (TobEcr : TOB; Code : string) : TOB;
    {Calcul de l'éclatement des soldes entre le pointé et le pointable}
    procedure Eclatement             (var AEcrDeb, AEcrCre, AMvtCre, AMvtDeb : Double;
                                      var PEcrDeb, PEcrCre, PMvtCre, PMvtDeb : Double;
                                      var AMvtNb, AEcrNb, PMvtNb, PEcrNb : Integer; ManuelOk : Boolean = False);
    {Renvoie la ligne suivante / précédente dont le champ Chp a la valeur Value}
    function  RechercheLigne         (Chp, Value : string; Suivant : Boolean = True) : Integer;
    {Renvoie la ligne suivante / précédente dont la date / mois est le suivant par rapport à la ligne en cours}
    procedure RechercheDate          (MoisOk : Boolean);
    {La colonne RefInterne, ne contient pas nécessairement la RefInterne, cela dépend de ComboRef}
    procedure MajColRefInterne       (ARow : Integer = -1; aTob : Tob = nil);
    {$IFDEF TRESO}
    procedure GenereEcriture;
    function  MajRefPointage         (var TOBPiece : TobPieceCompta; TobTreso : TOB = nil) : Boolean;
    {$ELSE}
    function  MajRefPointage         (var TOBPiece : TPieceCompta; TobTreso : TOB = nil) : Boolean;
    {$ENDIF TRESO}
  protected
    procedure AfficheEtatGrille      ;
    procedure InitControles          ;
    procedure InitEvenement          ;
    procedure InitVariables          ;
    procedure InitAffichage          (Arg : string);
    {Prépare la fiche en fonction du type de session : manuel ou non}
    procedure GereTypeSession        ;

    {Procédure pour appeler par le POPF11}
    procedure OnClickSelectionRapide (Sender : TObject);
    procedure OnClickSelectionJour   (Sender : TObject);
    procedure OnClickSelectionMois   (Sender : TObject);
    procedure NextPriorByFieldOnClick(Sender : TObject);
    procedure ResteAPointer          (Sender : TObject); {Détail du reste à pointer}
    procedure CoherenceSession       (Sender : TObject); {#ETAT# : pour les sessions de rappro, détail de la cohérence}
  public
    CkLibExercice : TCheckBox; {28/11/07}
    CNumCheque    : TCheckBox;
    CCombinatoire : TCheckBox;
    CModeReg      : TCheckBox;
    CRegAccro     : TCheckBox;
    CDate         : TCheckBox;
    ckLibCib      : TCheckBox;
    TDate         : TSpinEdit;
    TNiveau       : TSpinEdit;
    TNumCheque    : TSpinEdit;
    TPosCheque    : TSpinEdit;
    ComboRef      : THValCombobox;
    ComboLibMvt   : THValComboBox; {26/06/07 : FQ 10492}
    E_DateComptable_: THEdit;

    FDevise      : string; {Devise du compte général ou du compte de contrepartie du journal}
    FJournal     : string; {Journal si l'on est en pointage par journal}
    FGeneral     : string; {Général ou BQ_CODE si l'on est en pointage sur Tréso}
    FPtgeManuel  : Boolean; {A true, s'il n'y a pas de relevé rattaché à la session de pointage}
    FPtgeNonBqe  : Boolean; {14/11/07 : A True, s'il s'agit de pointage sur un compte non bancaire}
    FBoBordereau : Boolean; {A True, si le journal n'est pas un journal de saisie pièce}
    FStEE_RefPointage  : string;    {Contient la Réference de Pointage à utiliser}
    FDtEE_DatePointage : TDatetime; {Contient la Date de Pointage}
    FItEE_Numero       : Integer;        {Contient le numero interne du relevé}

    FGestionCibOk     : Boolean;   {17/08/07 : FQ 21255 : A True, s'il y a au moins un mode de paiement paramétré dans les CIB}
    FBoAllSelected    : Boolean;
    FBoDroitEcritures : Boolean;   {Droit de créer ou de modifier des écritures}
    FBoDroitMouvement : Boolean;   {Droit de créer des mouvements bancaires}
    FBoConsultation   : Boolean;   {Blocage des fonctions de Pointage ou Dépointage}
    FBoPeutCreer      : Boolean;   {On est en Lecture, mais on peut créer des mouvements}
    FBoSecurise       : Boolean;   {Pointage Securise OUI ou NON}
    FEtatRapproOk     : Boolean;   {Faut-il lancer l'état de rappro en sortie de fiche}
    FSoldeBqeCre      : Double;
    FSoldeBqeDeb      : Double;
    FSoldeEcrCre      : Double;
    FSoldeEcrDeb      : Double;
    FResteAPointe     : Double;
    FEtatPointage     : Double;  {#ETAT#}
    FNbDecimales      : Integer; {Onmbre de décimales de la devise}

    procedure ContrepartieDblClick  (Sender : TObject); {16/08/07 : FQ 13017}
    procedure PopupOnPopup          (Sender : TObject); {20/07/07 : FQ 21101}
    procedure PopupFiltreOnClick    (Sender : TObject);
    procedure LancerCib             (Sender : TObject);
    procedure RedAgrandOnClick      (Sender : TObject);
    procedure LancerMulEexBqLig     (Sender : TObject);
    procedure LancerRegle           (Sender : TObject);
    procedure LancerLibCib          (Sender : TObject);
    procedure BPointageAutoOnClick  (Sender : TObject);
    procedure CCombinatoireOnClick  (Sender : TObject);
    procedure CDateOnClick          (Sender : TObject);
    procedure CModeRegOnClick       (Sender : TObject);
    procedure CNumChequeOnClick     (Sender : TObject);
    procedure ComboRefOnChange      (Sender : TObject);
    procedure CreerEcriture         (Sender : TObject);
    procedure AccesSaisiePara       (Sender : TObject);
    procedure SupprEcriture         (Sender : TObject);
    procedure ModifEntete           (Sender : TObject);
    procedure CreerEexBqLig         (Sender : TObject);
    procedure CRegAccroOnClick      (Sender : TObject);
    procedure PrePointageOnClick    (Sender : TObject);
    procedure BExportOnClick        (Sender : TObject); {11/05/07 : FQ 20284}
    procedure OnClickBSelectAll     (Sender : TObject);
    procedure OnClickBValideEtRappro(Sender : TObject);
    procedure OnClickBValider       (Sender : TObject);
    procedure BFermeOnClick         (Sender : TObject); {23/07/07 : FQ 21058}
    procedure OnExitE_DateComptable_(Sender : TObject);
    procedure OnResizeEcran         (Sender : TObject);
    procedure BChercheOnClick       (Sender : TObject);
    procedure BRechercheOnClick     (Sender : TObject);
    procedure OnFindAFindDialog     (Sender : TObject);
    procedure DebitCreditOnClick    (Sender : TObject);
    procedure FListeOnDblClick      (Sender : TObject);
    procedure AnnulerPointage       (Sender : TObject);
    procedure RAZPointage           (Sender : TObject);
    procedure FormOnDestroy         (Sender : TObject); {27/12/07 : FQ 22118}
    procedure FormOnCloseQuery      (Sender : TObject; var CanClose : Boolean);
    procedure OnRowEnterFListe      (Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
    procedure OnKeyDownEcran        (Sender : TObject; var Key : Word; Shift : TShiftState);
    procedure OnKeyDownFListe       (Sender : TObject; var Key : Word; Shift : TShiftState);
    procedure OnMouseDownFListe     (Sender : TObject; Button : TMouseButton; Shift : TShiftState; X, Y : Integer);
    procedure GetCellCanvasFListe   (ACol, ARow : Integer; Canvas : TCanvas; AState : TGridDrawState);
    procedure PostDrawCell          (ACol, ARow : Integer; Canvas : TCanvas; AState : TGridDrawState);{18/07/06 : FQ 21106}

    property  AfficheDebitCredit : Boolean   read GetModeAffichage;
    property  MntCoherence       : Double    read GetMntCoherence;
    property  DateAno            : TDateTime read GetDateAno;{28/11/07}
    property  ExoAno             : string    read GetExoAno; {28/11/07}
  end;

procedure CPLanceFiche_PointageRappro(Arg : string);

implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  CPProcGen,
  CPTypeCons,
  {$ENDIF MODENT1}
  ULibWindows, UProcGen,
  {$IFDEF EAGLCLIENT}
  UtileAgl, eTablette,
  {$ELSE}
  EdtREtat, Tablette,
  {$ENDIF EAGLCLIENT}
  {$IFDEF TRESO}
  TRSaisieFlux_TOF, UProcEcriture, UProcCommission, UProcSolde,
  TRECRITURE_TOM, TRLIGNEECRCOMPTA_TOF, TRSUPPRECRITURE_TOF,
  {$ELSE}
  CPSAISIEPIECE_TOF, UTOFCPMULMVT, UTOFMODIFENTPIE, CPECRITURESIMPLE_TOF,
  {$ENDIF TRESO}
  CPMULEEXBQLIG_TOF, EEXBQLIG_TOM, REGLEACCRO_TOM, TomCIB, TRMULCIB_TOF,
  AglInit, Constantes, Ent1, Forms, Types, HSysMenu, SaisComm, UtilPgi,
  uTobDebug,HMsgBox, ParamSoc, Windows, ComCtrls, HQry, Commun, ULibExercice,
  uObjEtats,
  ExtCtrls, {TImage} HStatus, {InitMove}
  HXlspas, {ExportGrid : 11/05/07 : FQ 20284}
  CPRapproDet_Tof {CC_LanceFicheEtatRapproDet};

const
  COL_DATE    = 1;
  COL_PIECE   = 2;
  COL_LIGNE   = 3;
  COL_LIBELLE = 4;
  COL_REFINT  = 5;
  COL_CIB     = 6;
  COL_REFPOIN = 7;
  COL_DEVISE  = 8;
  COL_MNTD    = 9;
  COL_MNTC    = 10;
  COL_POINTE  = 11;

  {Si on est en pointage sur TRECRITURE, c'est informations sont nécessaire pour la mise à jour d'ECRITURE}
  TRESO_NODOSSIER  = 'CLE_REFEXTERNE';
  TRESO_USERCOMPTA = 'CLE_NUMTRAITECHQ';
  TRESO_ORIGINE    = 'CLE_QUALIFPIECE';
  TRESO_NUMLIGNE   = 'CLE_REFLIBRE';
  TRESO_NUMTRANSAC = 'CLE_MANUEL';


{---------------------------------------------------------------------------------------}
procedure CPLanceFiche_PointageRappro(Arg : string);
{---------------------------------------------------------------------------------------}
begin //OK@@
  AglLanceFiche('CP', 'CPPOINTAGERAP', '', '', Arg);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.OnArgument(S: string);
{---------------------------------------------------------------------------------------}
begin //OK@@
  inherited;
  {05/06/07 : en attendant le context de la compa}
  Ecran.HelpContext := 50000125;

  {Gestion des concepts}
  {$IFDEF TRESO}
  FBoDroitEcritures := CanValidateBO;
  {$ELSE}
  FBoDroitEcritures := ExJaiLeDroitConcept(TConcept(ccSaisEcritures), False);
  {$ENDIF TRESO}
  FBoDroitMouvement := ExJaiLeDroitConcept(TConcept(ccSaisMvtBqe   ), False);
  {Création / Initialisation des variables}
  InitVariables;
  {Présentation de l'affichage}
  InitAffichage(S);
  {Récupération des contrôles de l'écran}
  InitControles;
  {Rattache les évènements au contrôles}
  InitEvenement;
  {Prépare la fiche en fonction du type de session : manuel ou non}
  MetEnModeLecture;
  {28/12/07 : FQ 22119 : on charge le filtre à la toute fin des initialisation}
  ObjFiltre.Charger;
end;

{Sur le FormShow
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.OnLoad;
{---------------------------------------------------------------------------------------}
begin //OK@@
  inherited;
  (Ecran as TFVierge).Dock971.Height := 34;
  if not FBoConsultation then begin
    if FBoSecurise then
      FBoConsultation := ControleRefPointage
    else
      FBoConsultation := False;
  end;

  {Vérouillage des fonctions de pointage si on est en mode Consultation}
  if FBoConsultation then begin
    SetControlProperty('TCONSULTATION', 'CAPTION', TraduireMemoire('Modification impossible...'));
    SetControlEnabled('BVALIDER', False);
    SetControlEnabled('BVALIDEETRAPPRO', False);
  end
  else
    SetControlProperty('TCONSULTATION', 'CAPTION', '');

  SetControlProperty('TCONTREPARTIE', 'CAPTION', '');
  if (VH^.PointageJal) and (ControleContrepartie) then
    SetControlProperty('TCONTREPARTIE', 'CAPTION', TraduireMemoire('Présence d''écritures sur un autre journal'));

  SetControlProperty('TDEVISE', 'CAPTION', '');

  {03/05/07 : A afficher dans tous les cas, car l'état de rappro sera déséquilibré si des écritures ont
              été saisies dans une devise autre que la devise pivot. La présence du message permettra de
              voir d'être averti d'un décalage entre le pointage et l'état de rappro : le décalage viendra
              des écritures saisies dans une autre devise que la devise pivot. Chercher FQ 13731 dans CPRAPPRODET_TOF.Pas
  if FDevise <> V_PGI.DevisePivot then begin}
    if ControleDevise then
      SetControlProperty('TDEVISE', 'CAPTION', TraduireMemoire('Présence d''écritures saisies dans une autre devise'));
  //end;

  CRegAccroOnClick(CRegAccro);
  CMODEREGOnClick(CMODEREG);
  {Initialisation des soldes}
  InitSoldes;
  {Chargement de la Tob TobReleve contenant les mouvements et les écritures}
  ChargerMouvements;
  {Affichage dans la grille de la Tob TobReleve}
  ChargeGrille;
  {Definition de la grille : largeur, titres ...}
  PresenteGrille;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.OnClose;
{---------------------------------------------------------------------------------------}
begin //OK@@
  if Assigned(TobReleve)   then begin
    //FListe.VidePile(True);
    TobReleve.ClearDetail;
    FreeAndNil(TobReleve);
    FListe.RowCount := 1;
  end;

  if Assigned(TobCIB)      then FreeAndNil(TobCIB);
  if Assigned(TobRegles)   then FreeAndNil(TobRegles);
  if Assigned(TobLibCib)   then FreeAndNil(TobLibCib);
  if Assigned(AFindDialog) then FreeAndNil(AFindDialog);
  {JP 27/12/07 : FQ 22118 : Déplacer dans le FormOnDestroy
  if Assigned(ObjFiltre)   then FreeAndNil(ObjFiltre);}
  inherited;
end;

{Validation du pointage
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.OnUpdate;
{---------------------------------------------------------------------------------------}
begin //OK@@
  inherited;
  if FBoConsultation and not FSaufVVert then Exit;
  if MessageValidation(False) then
    MajPointage //EcritPointage
  else
    AnnulePointage;
end;

{Contrôles et confirmations de l'enregistrement du pointage
 23/07/07 : FQ 21053 et 21058 : Gestion des messages à la validation.
{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.MessageValidation(AnnuleOk : Boolean) : Boolean;
{---------------------------------------------------------------------------------------}
var
  lGene : string;
begin
  Result := True;
  {23/07/07 : FQ 21053 : Gestion des messages de validations}
  FPasMessage := True;

  {23/07/07 : FQ 21053 : Si on vient de la Croix rouge, on demande une confirmation ...}
  if AnnuleOk then
    AnnuleOk := IsPointageEncours and
                (PGIAsk(TraduireMemoire('Voulez-vous enregistrer le pointage en cours ?'), Ecran.Caption) <> mrYes);

  {FQ 21053 : ...On ne veut pas enregistrer les modifications}
  if AnnuleOk then begin
    Result :=  False;
    Exit;
  end;

  if FBoSecurise then begin
    {01/08/07 : FQ 21195 : ajout du contrôle de modifications en cours}
    if IsPointageEncours and (MntCoherence <> 0) then
      Result := PgiAsk(TraduireMemoire('Attention : l''avancement du pointage n''est pas équilibré.') + #13 +
                       TraduireMemoire('Confirmez vous la validation du pointage?'), Ecran.Caption) = mrYes;
  end
  else begin
    if VH^.PointageJal then lGene := FJournal
                       else lGene := FGeneral;

    {23/07/07 : FQ 21058 : message d'avertissement s'il y a des modifications en cours}
    if IsPointageEncours then begin
      {On recherche si des références de pointage postérieures ont déjà été crées}
      if ExisteSql('SELECT EE_DATEPOINTAGE FROM EEXBQ WHERE' +
                   ' EE_GENERAL = "' + lGene + '"' +
                   ' AND EE_DATEPOINTAGE > "' + UsDateTime(FDtEE_DatePointage) + '"') then begin
        {JP 21/12/07 : FQ 10542 : MntCoherence gère si l'on est en manuel ou non. Alors qu'avant on testait systématiquement le reste
                       à pointer, même en automatique, ce qui fait que le second message ne risquait pas de s'afficher}
        if (MntCoherence <> 0) then begin
          Result := PgiAsk(TraduireMemoire('Vous êtes sur le point de valider une session de pointage non équilibrée :') + #13 +
                           TraduireMemoire('toutes les sessions suivantes risquent d''être déséquilibrées.') + #13#13 +
                           TraduireMemoire('Confirmez vous la validation du pointage ?'), Ecran.Caption) = mrYes;
        end
        else begin
          Result := PgiAsk(TraduireMemoire('Vous êtes sur le point de valider une session de pointage équilibrée.') + #13 +
                           TraduireMemoire('Cependant, les modifications en cours risquent de déséquilibrer les session postérieures.') + #13#13 +
                           TraduireMemoire('Confirmez vous la validation du pointage ?'), Ecran.Caption) = mrYes;
        end;
      end;
    end;
  end;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.OnRowEnterFListe(Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin //OK@@
  inherited;
  T := TOB(FListe.Objects[0, Ou]);
  if not Assigned(T) then Exit;
  SetControlCaption('TLIBB', T.GetString('CLE_REFINTERNE'));
  if ctxTreso in V_PGI.PGIContexte then
    SetControlCaption('TDATEVAL', T.GetString('CLE_DATECOMPTABLE'))
  else
    SetControlCaption('TDATEVAL', T.GetString('CLE_DATEVALEUR'));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.ChargerMouvements;
{---------------------------------------------------------------------------------------}
var
  SQL : string;
  Ref : string;
begin //OK@@
  {04/05/07 : la date de fin ne peut être postérieure à la date de pointage}
  if StrToDate(E_DateComptable_.Text) > FDtEE_DatePointage then
    E_DateComptable_.Text := DateToStr(FDtEE_DatePointage);

  ChargeLesCIB;

  {ATTENTION : - L'ordre des champs DOIT être identique dans toutes les requêtes
               - en cas d'ajout de champ dans les requêtes, il FAUT le rajouter dans
                 ULibPointage.AddChampPointage à la même position}

  TobReleve.ClearDetail;
  if EstPointageSurTreso then begin
    Ref := 'IIF(DAY(TE_DATERAPPRO) < 10, "0"||TRIM(STR(DAY(TE_DATERAPPRO))) , TRIM(STR(DAY(TE_DATERAPPRO))))||"/"||';
    Ref := Ref + 'IIF(MONTH(TE_DATERAPPRO) < 10, "0"||TRIM(STR(MONTH(TE_DATERAPPRO))) , TRIM(STR(MONTH(TE_DATERAPPRO))))||"/"||';
    Ref := Ref + 'TRIM(STR(YEAR(TE_DATERAPPRO))) CLE_OLDREF';

    SQL := 'SELECT TE_DATECOMPTABLE CLE_DATECOMPTABLE, TE_NUMEROPIECE CLE_NUMEROPIECE, TE_NUMLIGNE CLE_NUMLIGNE, ' +
           'TE_REFPOINTAGE CLE_REFPOINTAGE, TE_DEVISE CLE_DEVISE, TE_CODECIB CLE_CIB, ';
    SQL := SQL + 'TE_REFINTERNE CLE_REFINTERNE, TE_LIBELLE CLE_LIBELLE, TE_NODOSSIER CLE_REFEXTERNE, TE_CPNUMLIGNE ' +
                 'CLE_REFLIBRE, TE_USERCOMPTABLE CLE_NUMTRAITECHQ, ';
    SQL := SQL + 'IIF(TE_MONTANT < 0, (TE_MONTANT * (-1)), 0) CLE_DEBIT, IIF(TE_MONTANT > 0, TE_MONTANT, 0) CLE_CREDIT, ' +
                 'IIF(TE_MONTANTDEV < 0, (TE_MONTANTDEV * (-1)), 0) CLE_DEBITDEV, IIF(TE_MONTANTDEV > 0, TE_MONTANTDEV, 0) CLE_CREDITDEV, ' +
                 'TE_JOURNAL CLE_JOURNAL, TE_EXERCICE CLE_EXERCICE, TE_DATEVALEUR CLE_DATEVALEUR, TE_DATERAPPRO CLE_DATEPOINTAGE, ';
    SQL := SQL + 'TE_NUMECHE CLE_NUMECHE, TE_QUALIFORIGINE CLE_QUALIFPIECE, TE_GENERAL CLE_GENERAL, ' +
                 'TE_MONTANT CLE_MONTANT, TE_MONTANTDEV CLE_MONTANTDEV, TE_NUMTRANSAC CLE_MANUEL, ';
    SQL := SQL + 'TE_DATECOMPTABLE CLE_DATEECHEANCE, TE_CODERAPPRO CLE_POINTE ,' + Ref + ' FROM TRECRITURE ';
    SQL := SQL + RecupAutreWhere + ' ';
  end
  else begin
    SQL := 'SELECT E_DATECOMPTABLE CLE_DATECOMPTABLE, E_NUMEROPIECE CLE_NUMEROPIECE, E_NUMLIGNE CLE_NUMLIGNE, ' +
           'E_REFPOINTAGE CLE_REFPOINTAGE, E_DEVISE CLE_DEVISE, E_MODEPAIE CLE_CIB, ';
    SQL := SQL + 'E_REFINTERNE CLE_REFINTERNE, E_LIBELLE CLE_LIBELLE, E_REFEXTERNE CLE_REFEXTERNE, E_REFLIBRE ' +
                 'CLE_REFLIBRE, E_NUMTRAITECHQ CLE_NUMTRAITECHQ, ';
    SQL := SQL + 'E_DEBIT CLE_DEBIT, E_CREDIT CLE_CREDIT, E_DEBITDEV CLE_DEBITDEV, E_CREDITDEV CLE_CREDITDEV, ' +
                 'E_JOURNAL CLE_JOURNAL, E_EXERCICE CLE_EXERCICE, E_DATEVALEUR CLE_DATEVALEUR, E_DATEPOINTAGE CLE_DATEPOINTAGE, ';
    SQL := SQL + 'E_NUMECHE CLE_NUMECHE, E_QUALIFPIECE CLE_QUALIFPIECE, E_GENERAL CLE_GENERAL, ';
    if VH^.PointageJal then
      SQL := SQL + '(E_CREDIT - E_DEBIT) CLE_MONTANT, (E_CREDITDEV - E_DEBITDEV) CLE_MONTANTDEV, "" CLE_MANUEL, '
    else
      SQL := SQL + '(E_DEBIT - E_CREDIT) CLE_MONTANT, (E_DEBITDEV - E_CREDITDEV) CLE_MONTANTDEV, "" CLE_MANUEL, ';
    SQL := SQL + 'E_DATEECHEANCE CLE_DATEECHEANCE, E_NATURETRESO CLE_POINTE, E_REFPOINTAGE CLE_OLDREF FROM ECRITURE ';
    SQL := SQL + RecupAutreWhere + ' ';
  end;
  TobReleve.LoadDetailFromSQL(SQL);

  SQL := 'SELECT CEL_DATEOPERATION CLE_DATECOMPTABLE, CEL_NUMRELEVE CLE_NUMEROPIECE, CEL_NUMLIGNE CLE_NUMLIGNE, ' +
         'CEL_REFPOINTAGE CLE_REFPOINTAGE, CEL_DEVISE CLE_DEVISE, CEL_CODEAFB CLE_CIB, ';
  SQL := SQL + 'CEL_LIBELLE1 CLE_REFINTERNE, CEL_LIBELLE CLE_LIBELLE, CEL_LIBELLE2 CLE_REFEXTERNE, CEL_LIBELLE3 ' +
               'CLE_REFLIBRE, CEL_REFORIGINE CLE_NUMTRAITECHQ, ';
  SQL := SQL + 'CEL_DEBITEURO CLE_DEBIT, CEL_CREDITEURO CLE_CREDIT, CEL_DEBITDEV CLE_DEBITDEV, CEL_CREDITDEV CLE_CREDITDEV, ' +
               '"' + LIGNEBANCAIRE + '" CLE_JOURNAL, "" CLE_EXERCICE, CEL_DATEVALEUR CLE_DATEVALEUR, CEL_DATEPOINTAGE CLE_DATEPOINTAGE, ';

  SQL := SQL + '0 CLE_NUMECHE, CEL_REFPIECE CLE_QUALIFPIECE, CEL_GENERAL CLE_GENERAL, ' +
               '(CEL_CREDITEURO - CEL_DEBITEURO) CLE_MONTANT, (CEL_CREDITDEV - CEL_DEBITDEV) CLE_MONTANTDEV, CEL_VALIDE CLE_MANUEL, ';
  SQL := SQL + 'NOW CLE_DATEECHEANCE, CEL_CODEPOINTAGE CLE_POINTE, CEL_REFPOINTAGE CLE_OLDREF FROM EEXBQLIG ';
  //SQL := SQL + 'LEFT JOIN EEXBQ ON EE_GENERAL = CEL_GENERAL AND EE_NUMRELEVE = CEL_NUMRELEVE ';

  if VH^.PointageJal then
    SQL := SQL + 'WHERE CEL_DATEOPERATION <= "' + UsDateTime(FDtEE_DatePointage) + '" AND CEL_GENERAL = "' + FJournal + '" '
  else
    SQL := SQL + 'WHERE CEL_DATEOPERATION <= "' + UsDateTime(FDtEE_DatePointage) + '" AND CEL_GENERAL = "' + FGeneral + '" ';

  {On charge les mouvements pointés lors de cette cession, les mouvements non pointés de sessions
   antérieurs ou encore les mouvements saisis manuellement (CEL_VALIDE = X)}
//  SQL := SQL + 'AND (CEL_DATEPOINTAGE = "' + UsDateTime(FDtEE_DatePointage) + '" OR (CEL_DATEPOINTAGE = "' + UsDateTime(iDate1900) + '" ';
  {16/07/07 : FQ 21056 : Gestion des dates de pointage postérieures}
  SQL := SQL + 'AND (CEL_DATEPOINTAGE >= "' + UsDateTime(FDtEE_DatePointage) + '" OR CEL_DATEPOINTAGE = "' + UsDateTime(iDate1900) + '") ';
  //SQL := SQL + 'AND (EE_DATEPOINTAGE <= "' + UsDateTime(FDtEE_DatePointage) + '" OR CEL_VALIDE = "X")))';

  TobReleve.LoadDetailFromSQL(SQL, True);
  if TobReleve.Detail.Count > 0 then
    TobReleve.Detail[0].AddChampSupValeur('MODIFIE', '-', True);
//  TobDebug(TobReleve);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.ChargeGrille;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin //OK@@
  FListe.TwoColors := False;
  FListe.ColCount := COL_POINTE + 1;
  FListe.RowCount := 2;
  FListe.ColTypes[COL_POINTE] := 'R'; {#0}
  FListe.ColFormats[COL_POINTE] := '0';

  {Il n'y a pas d'écritures ni de mouvements pour les critères}
  if TobReleve.Detail.Count = 0 then begin
    FListe.RowCount := 2;
    MajTotalisation;
    FBoConsultation := True;
    FBoPeutCreer := True;
    MetEnModeLecture(True); {28/11/07 : pour laisser le VVert actif}
  end

  else begin
    FBoPeutCreer := False;
    
    if FDevise <> V_PGI.DevisePivot then begin
      if ctxTreso in V_PGI.PGIContexte then
        TobReleve.Detail.Sort('CLE_REFPOINTAGE;CLE_POINTE;CLE_DATEVALEUR;CLE_MONTANTDEV;CLE_JOURNAL;')
      else
        TobReleve.Detail.Sort('CLE_REFPOINTAGE;CLE_POINTE;CLE_DATECOMPTABLE;CLE_MONTANTDEV;CLE_JOURNAL;');
    end
    else begin
      if ctxTreso in V_PGI.PGIContexte then
        TobReleve.Detail.Sort('CLE_REFPOINTAGE;CLE_POINTE;CLE_DATEVALEUR;CLE_MONTANT;CLE_JOURNAL;')
      else
        TobReleve.Detail.Sort('CLE_REFPOINTAGE;CLE_POINTE;CLE_DATECOMPTABLE;CLE_MONTANT;CLE_JOURNAL;');
    end;

    {26/06/07 : FQ 10492 : paramétrage du champ d'affichage des opérations bancaires}
    if ComboLibMvt.Visible and (ComboLibMvt.ItemIndex = -1) then ComboLibMvt.ItemIndex := 0;

    FListe.RowCount := TobReleve.Detail.Count + 1;
    for n := 0 to TobReleve.Detail.Count - 1 do
      RemplitLigne(n + 1, TobReleve.Detail[n]);
    MajSoldeManuel;
  end;

  {Mise à jour des totaux}
  MajTotalisation;
end;

{La colonne RefInterne, ne contient pas nécessairement la RefInterne, cela dépend de ComboRef
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.MajColRefInterne(ARow : Integer = -1; aTob : Tob = nil);
{---------------------------------------------------------------------------------------}

    {-----------------------------------------------------------------------}
    function _getValeurRefI(lTob : TOB) : string;
    {-----------------------------------------------------------------------}
    begin
      if lTob.GetString('CLE_JOURNAL') = LIGNEBANCAIRE then begin
        {26/06/07 : FQ 10492 : mise en place du paramétrage de la colonne pour les mvt bancaires}
        case ValeurI(ComboLibMvt.Value) of
          0 : Result := lTob.GetString('CLE_QUALIFPIECE'); {CEL_REFPIECE}
          1 : Result := lTob.GetString('CLE_REFINTERNE'); {CEL_LIBELLE1}
          2 : Result := lTob.GetString('CLE_REFEXTERNE'); {CEL_LIBELLE2}
          3 : Result := lTob.GetString('CLE_REFLIBRE');   {CEL_LIBELLE3}
        else
          Result := lTob.GetString('CLE_REFINTERNE');
        end;
      end
      else
        case ValeurI(ComboRef.Value) of
          0 : Result := lTob.GetString('CLE_REFINTERNE'); {Libellé, on laisse la RefInterne dans la grille, car le libellé est déjà affiché}
          1 : Result := lTob.GetString('CLE_REFEXTERNE');
          2 : Result := lTob.GetString('CLE_REFINTERNE');
          3 : Result := lTob.GetString('CLE_REFLIBRE');
          4 : Result := lTob.GetString('CLE_NUMTRAITECHQ');
        else
          Result := lTob.GetString('CLE_REFINTERNE');
        end;
    end;

var
  n : Integer;
begin
  if FPtgeManuel and (ARow <> -1) and Assigned(aTob) then
    FListe.Cells[COL_REFINT, ARow] := aTob.GetString('CLE_REFINTERNE')
  else begin
    if (ARow <> -1) and Assigned(aTob) then
      FListe.Cells[COL_REFINT, ARow] := _getValeurRefI(aTob)
    else begin
      FListe.BeginUpdate;
      for n := 1 to Fliste.RowCount - 1 do
        if Assigned(FListe.Objects[0, n]) then {28/12/07 : petite précaution, si la grille est vide}
          FListe.Cells[COL_REFINT, n] := _getValeurRefI(Tob(FListe.Objects[0, n]));
        
      FListe.EndUpdate;
      FListe.Refresh;
      Application.ProcessMessages;
    end;
  end;
end;


{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.RemplitLigne(ARow : Integer; aTob : TOB);
{---------------------------------------------------------------------------------------}
begin //OK@@
  FListe.BeginUpdate;
  if ctxTreso in V_PGI.PGIContexte then
    FListe.Cells[COL_DATE, ARow] := aTob.GetString('CLE_DATEVALEUR')
  else
    FListe.Cells[COL_DATE, ARow] := aTob.GetString('CLE_DATECOMPTABLE');

  FListe.Cells[COL_PIECE , ARow] := aTob.GetString('CLE_NUMEROPIECE');
  FListe.Cells[COL_LIGNE , ARow] := aTob.GetString('CLE_NUMLIGNE');

  MajColRefInterne(ARow, aTob);

  FListe.Cells[COL_LIBELLE , ARow] := aTob.GetString('CLE_LIBELLE');

  if (aTob.GetString('CLE_JOURNAL') = LIGNEBANCAIRE) then begin
    if EstPointageSurTreso then FListe.Cells[COL_CIB, ARow] := aTob.GetString('CLE_CIB') + ' *'
                           else FListe.Cells[COL_CIB, ARow] := GetLeModeRglt(aTob);
  end
  else
    FListe.Cells[COL_CIB, ARow] := aTob.GetString('CLE_CIB');

  FListe.Cells[COL_REFPOIN , ARow] := aTob.GetString('CLE_REFPOINTAGE');
  FListe.Cells[COL_DEVISE , ARow] := aTob.GetString('CLE_DEVISE');
  {25/04/07 : FQ 20066 : il faut tester la devise du journal et pas à la ligne, sinon en compta on
              risque d'avoir de drôles de mélanges
  if aTob.GetString('CLE_DEVISE') = V_PGI.DevisePivot then begin}
  if FDevise = V_PGI.DevisePivot then begin
    if AfficheDebitCredit then begin
      FListe.Cells[COL_MNTD , ARow] := FormateMontant(aTob.GetDouble('CLE_DEBIT'), FNbDecimales);
      FListe.Cells[COL_MNTC, ARow] := FormateMontant(aTob.GetDouble('CLE_CREDIT'), FNbDecimales);
    end
    else
      FListe.Cells[COL_MNTD , ARow] := FormateMontant(aTob.GetDouble('CLE_MONTANT'), FNbDecimales);
  end else begin
    if AfficheDebitCredit then begin
      FListe.Cells[COL_MNTD , ARow] := FormateMontant(aTob.GetDouble('CLE_DEBITDEV'), FNbDecimales);
      FListe.Cells[COL_MNTC, ARow] := FormateMontant(aTob.GetDouble('CLE_CREDITDEV'), FNbDecimales);
    end
    else
      FListe.Cells[COL_MNTD , ARow] := FormateMontant(aTob.GetDouble('CLE_MONTANTDEV'), FNbDecimales);
  end;
  FListe.Cells[COL_POINTE, ARow] := aTob.GetString('CLE_POINTE'); 
  FListe.Objects[0, ARow] := aTob;
  FListe.EndUpdate;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.PresenteGrille;
{---------------------------------------------------------------------------------------}
begin //OK@@
  {Définition des titres}
  if ctxTreso in V_PGI.PGIContexte then FListe.Cells[COL_DATE, 0] := TraduireMemoire('Date Valeur')
                                   else FListe.Cells[COL_DATE, 0] := TraduireMemoire('Date Compt.');
  FListe.Cells[COL_PIECE  , 0] := TraduireMemoire('Pièce');
  FListe.Cells[COL_LIGNE  , 0] := TraduireMemoire('Ligne');
  FListe.Cells[COL_REFINT , 0] := TraduireMemoire('Référence interne');
  FListe.Cells[COL_LIBELLE, 0] := TraduireMemoire('Libellé');
  FListe.Cells[COL_CIB    , 0] := TraduireMemoire('Cib');
  FListe.Cells[COL_REFPOIN, 0] := TraduireMemoire('Réf. Pointage');
  FListe.Cells[COL_DEVISE , 0] := TraduireMemoire('Devise');
  FListe.Cells[COL_POINTE , 0] := TraduireMemoire('Code');
  if AfficheDebitCredit then begin
    FListe.Cells[COL_MNTD , 0] := TraduireMemoire('Debit');
    FListe.Cells[COL_MNTC , 0] := TraduireMemoire('Crédit');
  end else begin
    FListe.Cells[COL_MNTD , 0] := TraduireMemoire('Montant');
    FListe.Cells[COL_MNTC , 0] := '';
  end;

  {Alignement}
  FListe.ColAligns[COL_DATE   ] := taCenter;
  FListe.ColAligns[COL_CIB    ] := taCenter;
  FListe.ColAligns[COL_DEVISE ] := taCenter;
  FListe.ColAligns[COL_PIECE  ] := taRightJustify;
  FListe.ColAligns[COL_LIGNE  ] := taRightJustify;
  FListe.ColAligns[COL_POINTE ] := taLeftJustify; {18/07/07 : FQ 21121 : à gauche plutôt qu'à droite}
  FListe.ColAligns[COL_MNTC   ] := taRightJustify;
  FListe.ColAligns[COL_MNTD   ] := taRightJustify;
  FListe.ColAligns[COL_REFINT ] := taLeftJustify;
  FListe.ColAligns[COL_LIBELLE] := taLeftJustify;
  FListe.ColAligns[COL_REFPOIN] := taLeftJustify;

  {Largeur}
  FListe.ColWidths[COL_DATE   ] := Round((FListe.Width - 30) * 0.07);
  FListe.ColWidths[COL_CIB    ] := Round((FListe.Width - 30) * 0.05);
  FListe.ColWidths[COL_DEVISE ] := Round((FListe.Width - 30) * 0.04);
  FListe.ColWidths[COL_PIECE  ] := Round((FListe.Width - 30) * 0.08);
  FListe.ColWidths[COL_LIGNE  ] := Round((FListe.Width - 30) * 0.07);
  FListe.ColWidths[COL_REFINT ] := Round((FListe.Width - 30) * 0.15);
  FListe.ColWidths[COL_LIBELLE] := Round((FListe.Width - 30) * 0.20);
  FListe.ColWidths[COL_REFPOIN] := Round((FListe.Width - 30) * 0.10);

  {12/07/07 : FQ 21043 : Dans une session manuelle on cache la colonne Pointé
   30/07/07 : j'annule : j'avais mal compris la FQ
  if FPtgeManuel then begin
    FListe.ColWidths [COL_POINTE] := -1;
    FListe.ColLengths[COL_POINTE] := -1;
  end
  else}
    FListe.ColWidths[COL_POINTE] := Round((FListe.Width - 30) * 0.04);

  if AfficheDebitCredit then begin
    FListe.ColWidths[COL_MNTC ] := Round((FListe.Width - 30) * 0.10);
    FListe.ColWidths[COL_MNTD ] := Round((FListe.Width - 30) * 0.10);
  end else begin
    FListe.ColWidths [COL_MNTC] := -1;
    FListe.ColLengths[COL_MNTC] := -1;
    FListe.ColWidths [COL_MNTD] := Round((FListe.Width - 30) * 0.2);
  end;

  {Format}
  {30/07/07 : FQ 21043 : Ajout de ' ' sur le troisième paramètre concerant à la valeur 0,00}
  FListe.ColFormats[COL_MNTD] := StrFMask(CalcDecimaleDevise(FDevise), '', True) + ';; ';
  FListe.ColFormats[COL_MNTC] := StrFMask(CalcDecimaleDevise(FDevise), '', True) + ';; ';
  FListe.ColFormats[COL_POINTE] := '0';
  {Précision du ColTypes, pour que le tri (par click sur la Colonne) ne soit pas alphanumérique mais numérique}
  FListe.ColTypes[COL_POINTE] := 'R';
  FListe.ColTypes[COL_MNTC  ] := 'R';
  FListe.ColTypes[COL_MNTD  ] := 'R';
  FListe.ColTypes[COL_DATE  ] := 'D';

  {01/08/07 : FQ 21201 : gestion de l'état d'affichage avec filtre sur les données}
  ShowLignes(EtatAffichage);

  {Pour forcer le redessin de l'écran}
  OnResizeEcran(Ecran);

  {12/07/07 : FQ 21044 : on donne le focus à la grille}
  if FListe.CanFocus then FListe.SetFocus
  else if THEdit(GetControl('E_DATECOMPTABLE')).CanFocus then
    SetFocusControl('E_DATECOMPTABLE');
end;

{12/07/07 : FQ 21046 : Gestion du F12
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.GereFocusF12;
{---------------------------------------------------------------------------------------}
begin
  {On est sur la grille ...}
  if Ecran.ActiveControl is THGrid then begin
    {1 / Si la date comptable peut prendre le focus ...}
    if THEdit(GetControl('E_DATECOMPTABLE')).CanFocus then
      SetFocusControl('E_DATECOMPTABLE')

    else begin
      {2 / Sinon, on regarde si le PageControl est sur la première page ...}
      if GetControlVisible('PAGECONTROL') then begin
        (GetControl('PAGECONTROL') as TPageControl).ActivePageIndex := 0;
        {... Pour donner le focus à la date comptable}
        if THEdit(GetControl('E_DATECOMPTABLE')).CanFocus then
          SetFocusControl('E_DATECOMPTABLE');
      end

      else begin
        {3 / Le pagecontrol doit être caché, on le rend visible}
        RedAgrandOnClick(GetControl('BREDUIRE'));
        if THEdit(GetControl('E_DATECOMPTABLE')).CanFocus then
          SetFocusControl('E_DATECOMPTABLE')
      end;
    end;
  end

  {... Sinon, on donne le focus à la grille}
  else
    FListe.SetFocus;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.GetLeModeRglt(var aTob : TOB) : string;
{---------------------------------------------------------------------------------------}
var
  TobMP : TOB;
  CIB   : string;
begin //OK@@
  CIB := aTob.GetString('CLE_CIB');

  Result := CIB;
  if Length(CIB) = 1 then CIB := '0' + CIB;
  TobMP := TobCIB.FindFirst(['TCI_CODECIB'], [CIB], True);

  if TobMP = nil then begin
    Result := CIB + '/...';
    Exit;
  end;

  Result := CIB + '/'+ TobMp.GetString('TCI_MODEPAIE');
  {On mémorise le mode de paiement dans le champ exercice qui est vide pour les mvts bancaires}
  aTob.SetString('CLE_EXERCICE', TobMp.GetString('TCI_MODEPAIE'));
end;

{$IFNDEF TRESO}
{20/09/07 : FQ 21333 : Fenêtre de choix d'une date comptable
{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.GetDateComptable : TDateTime;
{---------------------------------------------------------------------------------------}
begin
  Result := FDtEE_DatePointage;
  if not FBoBordereau then
    Result := RecupDateComptable(Result);
end;

{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.GetModePaie(Cib, Sens : string) : string;
{---------------------------------------------------------------------------------------}
var
  TobMP : TOB;
begin //OK@@
  Result := '';

  if Length(CIB) = 1 then CIB := '0' + Cib;
  TobMP := TobCIB.FindFirst(['TCI_CODECIB', 'TCI_SENS'], [Cib, Sens], True);
  if not Assigned(TobMP) then TobMP := TobCIB.FindFirst(['TCI_CODECIB', 'TCI_SENS'], [Cib, 'MIX'], True);
  if Assigned(TobMP) then Result := TobMP.GetString('TCI_MODEPAIE');
end;
{$ENDIF TRESO}

{Vérifie si l'écriture correspond au mouvement bancaire en cours
 Cette fonction est surtout valable pour les bordereaux
{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.IsTobAPointe(TobEcr : TOB) : Boolean;
{---------------------------------------------------------------------------------------}
var
  TB : TOB;
begin
  Result := False;
  {31/07/07 : FQ 21199 : Avec l'autorisation de créer des écritures à partir de mouvements manuels (FQ 21125)
              Mais même sans cet ajout fonctionnel, sortir pose un problème en saisie bordereau, cf les modifs
              du 27/07/07 ci-dessous
  if FPtgeManuel then begin
    Result := True;
    Exit;
  end;}

  TB := TOB(FListe.Objects[0, FListe.Row]);
  if not Assigned(TB) then Exit;
  if TB.GetString('CLE_JOURNAL') <> LIGNEBANCAIRE then begin
    {27/07/07 : En tréso, quoiqu'il arrive on insère l'écriture si ...}
    if EstPointageSurTreso then begin
      {... On a Même général, pas pointé (ce qui devrait être toujours le cas et la date d'opération est cohérente}
      Result := (TobEcr.GetString('TE_GENERAL') = FGeneral) and
                (TobEcr.GetDateTime('TE_DATECOMPTABLE') <= FDtEE_DatePointage) and
                (TobEcr.GetString('TE_REFPOINTAGE') = '');
      {Par contre, cela entraîne une modification de ConvertitTob pour ne pas la prépointer}
    end
    {27/07/07 : Si on n'est pas en pointage sur journal ou sur un journal en mode bordereau,
                quoiqu'il arrive on insère l'écriture si ...}
    else if not FBoBordereau and not VH^.PointageJal then begin
      {... On a Même général, pas pointé (ce qui devrait être toujours le cas et la date d'opération est cohérente}
      Result := (TobEcr.GetString('E_GENERAL') = FGeneral) and
                (TobEcr.GetDateTime('E_DATECOMPTABLE') <= FDtEE_DatePointage) and
                (TobEcr.GetString('E_REFPOINTAGE') = '');
      {Par contre, cela entraîne une modification de ConvertitTob pour ne pas la prépointer}
    end;
    Exit;
  end;

  if EstPointageSurTreso then begin
    {Pour être sur que l'on est sur la bonne ligne d'écriture, on s'assure :
     1/ que les montants sont identiques
     2/ que les dates comptables correspondent
     3/ que les comptes correspondent
     4/ que l'écriture n'est pas pointée}
    if V_PGI.DevisePivot = TB.GetString('CLE_DEVISE') then
      Result := (Arrondi(TB.GetDouble('CLE_MONTANT'), FNbDecimales) = Arrondi(TobEcr.GetDouble('TE_MONTANT'), FNbDecimales)) and
                (TB.GetString('CLE_DATECOMPTABLE') = TobEcr.GetString('TE_DATECOMPTABLE')) and
                (TobEcr.GetString('TE_REFPOINTAGE') = '') and
                (TB.GetString('CLE_GENERAL') = TobEcr.GetString('TE_GENERAL'))
    else
      Result := (Arrondi(TB.GetDouble('CLE_MONTANTDEV'), FNbDecimales) = Arrondi(TobEcr.GetDouble('TE_MONTANTDEV'), FNbDecimales)) and
                (TB.GetString('CLE_DATECOMPTABLE') = TobEcr.GetString('TE_DATECOMPTABLE')) and
                (TobEcr.GetString('TE_REFPOINTAGE') = '') and
                (TB.GetString('CLE_GENERAL') = TobEcr.GetString('TE_GENERAL'));
  end
  else begin
    {Pour être sur que l'on est sur la bonne ligne d'écriture, on s'assure :
     1/ que les montants sont identiques
     2/ que les dates comptables correspondent
     3/ que les comptes correspondent
     4/ que l'écriture n'est pas pointée}
    if V_PGI.DevisePivot = TB.GetString('CLE_DEVISE') then
      Result := (Arrondi(TB.GetDouble('CLE_CREDIT') + TB.GetDouble('CLE_DEBIT'), FNbDecimales) =
                 Arrondi(TobEcr.GetDouble('E_CREDIT') + TobEcr.GetDouble('E_DEBIT'), FNbDecimales)) and
                (TB.GetString('CLE_DATECOMPTABLE') = TobEcr.GetString('E_DATECOMPTABLE')) and
                (TobEcr.GetString('E_REFPOINTAGE') = '')
    else
      Result := (Arrondi(TB.GetDouble('CLE_CREDITDEV') + TB.GetDouble('CLE_DEBITDEV'), FNbDecimales) =
                 Arrondi(TobEcr.GetDouble('E_CREDITDEV') + TobEcr.GetDouble('E_DEBITDEV'), FNbDecimales)) and
                (TB.GetString('CLE_DATECOMPTABLE') = TobEcr.GetString('E_DATECOMPTABLE')) and
                (TobEcr.GetString('E_REFPOINTAGE') = '');

    if VH^.PointageJal then
      Result := Result and (TB.GetString('CLE_GENERAL') = TobEcr.GetString('E_JOURNAL'))
    else
      Result := Result and (TB.GetString('CLE_GENERAL') = TobEcr.GetString('E_GENERAL'));
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.ChargeLesCIB;
{---------------------------------------------------------------------------------------}
var
  w : string;
  n : Integer;
begin //OK@@
  if not Assigned(TobCib) then
    TobCib := Tob.Create('***', nil ,-1)
  else
    TobCib.ClearDetail;

  if CtxPcl in V_PGI.PGIContexte then
    w := 'WHERE TCI_BANQUE = "' + CODECIBREF + '"'
  else
    w := 'WHERE TCI_BANQUE = (SELECT BQ_BANQUE FROM BANQUECP WHERE BQ_GENERAL = "' + FGeneral + '" ' +
                              'AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '")';
  {Charge le tableau de correspondance CIB-Règle}
  TobCib.LoadDetailFromSQL('SELECT TCI_REGLEACCRO, TCI_MODEPAIE, TCI_CODECIB, TCI_SENS FROM CIB ' +
                           w + ' ORDER BY TCI_CODECIB');
  {17/08/07 : 21255 : On regarde si le client a paramétré au moins un mode de paiement sur les CIB. Dans le cas
              contraires, on peut appliquer les règles d'accrochage sans tester les modes de paiement}
  FGestionCibOk := False;
  for n := 0 to TobCib.Detail.Count - 1 do begin
    if (TobCib.Detail[n].GetString('TCI_MODEPAIE') <> '') and
       (TobCib.Detail[n].GetString('TCI_CODECIB') <> CODECIBTITRE) and
       (TobCib.Detail[n].GetString('TCI_CODECIB') <> CODECIBCOURANT) then begin
      FGestionCibOk := True;
      Break;
    end;
  end;

  {Charge les critères dans une Tob}
  if TobRegles = nil then begin
    TobRegles := TOB.Create('_REGLEACCRO', nil, -1);
    TobRegles.LoadDetailFromSQL('SELECT * FROM REGLEACCRO');
  end;

  {Chargement du paramétrage des libellés sur les CIB}
  if TobLibCib = nil then begin
    TobLibCib := TOB.Create('_LIBCIB', nil, -1);
    TobLibCib.LoadDetailFromSQL('SELECT CC_CODE, CC_ABREGE FROM CHOIXCOD WHERE CC_TYPE = "TCI"');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.PrePointageOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  M : TOB;
  F : TOB;
  O : TOB;
  n : Integer;
  T : string;
begin //OK@@
  M := TOB.Create('µPREPOINTAGE', nil, -1);
  try
    for n := 1 to FListe.RowCount - 1 do begin
      if Trim(FListe.Cells[COL_POINTE, n]) <> '' then begin
        F := TOB.Create('µLIGNE', M, -1);
        F.AddChampSupValeur('REFPOINTAGE', FListe.Cells[COL_REFPOIN , n]);
        F.AddChampSupValeur('REFERENCE'  , FListe.Cells[COL_PIECE   , n]);
        F.AddChampSupValeur('LIGNE'      , FListe.Cells[COL_LIGNE   , n]);
        F.AddChampSupValeur('RGLT'       , FListe.Cells[COL_CIB     , n]);
        F.AddChampSupValeur('ADEVISE'    , FListe.Cells[COL_DEVISE  , n]);
        F.AddChampSupValeur('LIBELLE'    , FListe.Cells[COL_LIBELLE , n]);
        F.AddChampSupValeur('POINTER'    , FListe.Cells[COL_POINTE  , n]);
        F.AddChampSupValeur('LIBELLEENR' , FListe.Cells[COL_REFINT  , n]);

        O := TOB(Fliste.Objects[0, n]);
        if Assigned(O) then begin
          if FListe.Cells[COL_DEVISE, n] = V_PGI.DevisePivot then F.AddChampSupValeur('SOLDE', O.GetDouble('CLE_MONTANT'))
                                                             else F.AddChampSupValeur('SOLDE', O.GetDouble('CLE_MONTANTDEV'));
          F.AddChampSupValeur('DATEVAL'    , O.GetDateTime('CLE_DATEVALEUR'));
          F.AddChampSupValeur('DATE'       , O.GetDateTime('CLE_DATECOMPTABLE'));
          if O.GetString('CLE_JOURNAL') = LIGNEBANCAIRE then F.AddChampSupValeur('TYPEECR', 'GB')
                                                        else F.AddChampSupValeur('TYPEECR', 'GC');
        end;
      end;
    end;

    //T := TraduireMemoire('État de prépointage (Session numéro ') + IntToStr(FItEE_Numero) + ')';
    T := Ecran.Caption;
    
    M.Detail.Sort('REFPOINTAGE;POINTER;TYPEECR;');
    if FPtgeManuel then
      LanceEtatTob('E', 'CPE', 'RAP', M, True, False, False, nil, '', T, False)
    else
      LanceEtatTob('E', 'CPE', 'RA1', M, True, False, False, nil, '', T, False);
  finally
    FreeAndNil(M);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.BPointageAutoOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  if FPtgeManuel or FBoConsultation then Exit; {18/07/07 : FQ 21130}
  if not ControleAvantRappro then Exit;
  if IsPointageEncours then AnnulerPointage(Sender);
  LanceLePointageAuto;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.ControleAvantRappro : Boolean;
{---------------------------------------------------------------------------------------}
begin //OK@@
  Result := False;
  if FBoConsultation then Exit;
  
  if CNumCheque.Checked and (TNumCheque.value > TPosCheque.value) then begin
    HShowMessage('16;' + Ecran.Caption + ';Paramètre de la vérification des numéros de chèque incorrect.;E;O;O;O', '', '');
    Exit;
  end;

  if EtatAffichage <> tv_Tout then begin
    PGIInfo(TraduireMemoire('Veuillez-vous mettre en "Affichage de tous les mouvements et écritures".'));
    Exit;
  end;

//  if CModeReg.Checked and (not ExisteMoyenPaiement) then exit ;
  Result := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.CModeRegOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  CRegAccro.Enabled := CModeReg.Checked;
  ckLibCib .Enabled := CModeReg.Checked;
  CRegAccroOnClick(CRegAccro);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.CRegAccroOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  if UpperCase(TComponent(Sender).Name) = 'CKLIBCIB' then begin
    if ckLibCib.Checked then CRegAccro.Checked := False
  end
  else if CRegAccro.Checked then ckLibCib.Checked := False;

  CDate.Enabled := CRegAccro.Checked and CRegAccro.Enabled;
  CDateOnClick(CDate);
  CNumCheque.Enabled := CDate.Enabled;
  CNumChequeOnClick(CNumCheque);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.ComboRefOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  {26/06/07 : FQ 10492 : mise à jour de la colonne Référence pour les opérations bancaires}
  if TComponent(Sender).Name = ComboLibMvt.Name then
    MajColRefInterne

  else if ComboRef.enabled then begin
    {Dans le cas du libellé ('0'), on ne change rien car le libellé figure déjà dans la grille}
    if ValeurI(ComboRef.Value) > 0 then
      FListe.Cells[COL_REFINT, 0] := ComboRef.Text;
    MajColRefInterne;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.CCombinatoireOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  TNiveau.Enabled := not TNiveau.Enabled;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.CDateOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  TDate.Enabled := CDate.Checked and CDate.Enabled;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.CNumChequeOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  TNumCheque.Enabled := CNumCheque.Checked and CNumCheque.Enabled;
  TPosCheque.Enabled := TNumCheque.Enabled;
  ComboRef  .Enabled := TNumCheque.Enabled;
  SetControlEnabled('LBCARACT', TNumCheque.Enabled); {JP 26/04/05 : FQ TRESO 10241}
  SetControlEnabled('LBSUR', TNumCheque.Enabled); {JP 26/04/05 : FQ TRESO 10241}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.RemetRefPointageABlanc;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin //OK@@
  if FBoConsultation then Exit;

  {On force l'affichage de tous les mouvements et les écritures afin de pouvoir travailler sur la grille}
  if EtatAffichage <> tv_Tout then begin
    PGIInfo(TraduireMemoire('Veuillez-vous mettre en "Affichage de tous les mouvements et écritures".'));
    Exit;
  end;

  {Il y a un pointage en cours ...}
  for n := 1 to FListe.RowCount - 1 do
    if Tob(FListe.Objects[0, n]).GetString('CLE_REFPOINTAGE') = FStEE_RefPointage then
      CocheDecoche(n, '');

  {Mise à jour des totaux}
  MajTotalisation;

  PGIInfo(TraduireMemoire('Tout le pointage de la session en cours a été supprimé'), Ecran.Caption);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.ReinitialisePointage;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  T : TOB;
begin //OK@@
  if FBoConsultation then Exit;
  {On force l'affichage de tous les mouvements et les écritures afin de pouvoir travailler sur la grille}
  if EtatAffichage <> tv_Tout then
    PGIInfo(TraduireMemoire('Veuillez-vous mettre en "Affichage de tous les mouvements et écritures".'))
  else begin
    for n := 1 to FListe.RowCount -1 do begin
      T := TOB(FListe.Objects[0, n]);
      if T.GetString('MODIFIE') = 'X' then CocheDecoche(n, '');
    end;

    {Mise à jour des totaux}
    MajTotalisation;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.LanceLePointageAuto;
{---------------------------------------------------------------------------------------}
var
  LM     : T_D;
  LP, LB : T_I;
  Solde  : Double;
  Infos  : REC_AUTO;
  i, j   : Integer;
  St     : string;
  nk     : Integer;
  nb     : Integer; {FQ 10440}
begin //OK@@
  DecReleve := 2;
  DecCompta := 2;
  NbDecDev  := 2;
  nb := 0;

  InitMove(FListe.RowCount * 5, TraduireMemoire('Rapprochement automatique'));
  try
    if FListe.Cells[1, 1] = '' then Exit;
    PointageAuto := True;
    PointeAvecListeNonComplete := False;
    Infos.Decim := NbDecDev;

    if CDate.Checked then	ToleranceDate := TDate.Value
                     else	ToleranceDate := 0;

    Infos.Decim := NbDecDev ;
    Infos.Temps := 0;

    if CCombinatoire.Checked then begin
      Infos.Nival := TNiveau.Value - 1;
      Infos.Temps := 30;
      Infos.Unique:= False;
    end
    else begin
      Infos.Nival := -1;
      Infos.Temps := MaxTempo;
      Infos.Unique:= True;
    end ;

    St := TraduireMemoire('Traitement ligne n° ');
    {n° de pointage}
    nk := StrToInt(GetMaxNumPtge);

    for i := 1 to FListe.RowCount - 1 do begin
      MoveCur(False);
      Application.ProcessMessages ;
      TobBqeEnCours := TOB(FListe.Objects[0, i]);
      {On ne travaille que sur les lignes bancaires}
      if TobBqeEnCours.GetString('CLE_JOURNAL') <> LIGNEBANCAIRE then Continue;
      {On ne travaille que sur les mouvements de la session de pointage en cours
       17/08/07 : FQ 20196 : Pour permettre le rapprochement de plusieurs relevés dans une même session,  je fais sauter le
                  test. Peut-être faudra-t-il envisagé un concept sur le sujet : Rappro mono session : Oui / Non ?
      SBO 28/08/2007 : Test d'ouverture soumis au paramètre pointage sécurisé  }
      if FBoSecurise then
        if TobBqeEnCours.GetInteger('CLE_NUMEROPIECE') <> FItEE_Numero then Continue;

      {On ne traite pas les mouvements déjà pointés}
      if TobBqeEnCours.GetString('CLE_POINTE') <> '' then Continue;

      if V_PGI.DevisePivot = TobBqeEnCours.GetString('CLE_DEVISE') then
        Solde := Arrondi(TobBqeEnCours.GetDouble('CLE_MONTANT'), DecReleve)
      else
        Solde := Arrondi(TobBqeEnCours.GetDouble('CLE_MONTANTDEV'), DecReleve);

      Infos.NbD := ConstruitListe(LM, LP, LB, Solde);

      {03/05/07 : FQ 20197 : En combinatoire, on va pointer les mouvements bancaires, mais pas les écritures
                  si Infos.NbD > 0. J'ai modifié ConstruitListe pour qu'il renvoie 0 si aucune écriture
                  ne correspond au mouvement courant}
      if Infos.NbD = 0 then Continue;

      if LettrageAuto(Solde, LM, LP, Infos) = 1 then begin
        CocheDecoche(i, IntToStr(nk));
        for j := 1 to Infos.NbD do
          if LP[j - 1] <> 0 then
            CocheDecoche(LB[j - 1], IntToStr(nk));
        Inc(nk);
        Inc(nb);
      end;

    end;

    {Mise à jour des totaux}
    MajTotalisation;

    if PointeAvecListeNonComplete then
      PGIError(TraduireMemoire('Certains montants n''ont pu être pointés correctement.'), Ecran.Caption)
    else begin
      {20/04/07 : FQ TRESO 10440 : affichage d'un message de confirmation}
      if nb = 0 then
        PGIInfo(TraduireMemoire('Aucun mouvement n''a été pointé.'), Ecran.Caption)
      else if nb = 1 then
        PGIInfo(TraduireMemoire('Un mouvement a été pointé.'), Ecran.Caption)
      else
        PGIInfo(IntToStr(nB) + ' ' + TraduireMemoire('mouvements ont été pointés.'), Ecran.Caption);
    end;
  finally
    FiniMove;
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.ConstruitListe(var LM : T_D; var LP, LB : T_I; MM : Double) : Integer;
{---------------------------------------------------------------------------------------}
var
  i, j, k : Integer;
  Mtt : Double;
  s   : string;
  BoucleOk : Boolean;
begin //OK@@
  Result := 0;
  i := 0;
  BoucleOk := False;

  FillChar(LM, Sizeof(LM), #0);
  FillChar(LP, Sizeof(LP), #0);

  for j := 1 to FListe.RowCount - 1 do begin
    MoveCur(False);

    TobCptEnCours := Tob(Fliste.Objects[0, j]);
    {On ne travaille que sur les écritures}
    if TobCptEnCours.GetString('CLE_JOURNAL') = LIGNEBANCAIRE then Continue;
    {On ne traite pas les écritures pointés}
    if TobCptEnCours.GetString('CLE_POINTE') <> '' then Continue;

    s := TobCptEnCours.GetString('CLE_REFPOINTAGE');
    if (TobCptEnCours.GetString('CLE_REFPOINTAGE') = '') and EcritureSelectionnee then begin
      if TobCptEnCours.GetString('CLE_DEVISE') = V_PGI.DevisePivot then
        Mtt := TobCptEnCours.GetDouble('CLE_DEBIT') - TobCptEnCours.GetDouble('CLE_CREDIT')
      else
        Mtt := TobCptEnCours.GetDouble('CLE_DEBITDEV') - TobCptEnCours.GetDouble('CLE_CREDITDEV');
      {En pointage sur journal, comme on affiche les comptes lettrables,
       les montants sont dans le sens bancaires}
      if VH^.PointageJal or EstPointageSurTreso then Mtt := -Mtt;

      {Si les sens sont opposés, on pas à la ligne suivantes}
      if (((MM <= 0) and (Mtt >= 0)) or ((MM >= 0) and (Mtt <= 0))) then Continue;

      k := i;
      if k = MaxDroite - 1 then begin
        PointeAvecListeNonComplete := True;
        Exit;
      end;

      {classement des montants par ordre décroissant}
      while k <> 0 do begin
        if Mtt < LM[k - 1] then begin
          LM[k] := LM[k - 1];
          LP[k] := LP[k - 1];
          LB[k] := LB[k - 1];
        end
        else
          Break;

        System.Dec(k);
      end;
      LM[k] := Mtt;
      LP[k] := 0;
      LB[k] := j;
      Inc(i);
      MoveCur(False);
      BoucleOk := True;
    end;
  end;

  {03/05/07 : FQ 20197 : En combinatoire, on va pointer les mouvements bancaires, mais pas les écritures si Result > 0}
  if BoucleOk then
    Result := i;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.EcritureSelectionnee : Boolean;
{---------------------------------------------------------------------------------------}
var
  NbeCar,PosCar : integer;
  Reference,ModePaieEcr,ModePaieRel,sRef : string;
  TobL,
  TobR    : TOB;
  Regle   : string;

    {Teste l'égalité approximative des deux dates de la colonne fournie}
    {------------------------------------------------------------------------}
    function EgalTolerant(Chp : string) : Boolean;
    {------------------------------------------------------------------------}
    begin
      Result := Trunc(Abs(TobCptEnCours.GetDateTime(Chp) - TobBqeEnCours.GetDateTime(Chp))) <= ToleranceDate;
    end;

    {JP 11/04/05 : FQ 10241 : application de ce traitement à la TRESO
    {------------------------------------------------------------------------}
    function VerifCheque : Boolean;
    {------------------------------------------------------------------------}
    var
      CR : Integer;
    begin
      {En comptabilité, il sera toujours à True si on arrive ici ...}
      if CNumCheque.Checked then begin
        PosCar := TNumCheque.Value;
        NbeCar := TPosCheque.Value;
        NbeCar  :=  NbeCar - PosCar + 1;
        Reference := Copy(TobBqeEnCours.GetString('CLE_QUALIFPIECE'), PosCar, NbeCar);
      end
      {... ce cas ne peut arriver qu'en Tréso}
      else
        Reference := TobBqeEnCours.GetString('CLE_QUALIFPIECE');

      CR := ValeurI(ComboRef.Value);
      case CR of
        0 : sRef := TobCptEnCours.GetString('CLE_LIBELLE');
        1 : sRef := TobCptEnCours.GetString('CLE_REFEXTERNE');
        2 : sRef := TobCptEnCours.GetString('CLE_REFINTERNE');
        3 : sRef := TobCptEnCours.GetString('CLE_REFLIBRE');
        4 : sRef := TobCptEnCours.GetString('CLE_NUMTRAITECHQ')
       else
         sRef := TobCptEnCours.GetString('CLE_LIBELLE');
      end;
      Result := (Pos(Reference, sRef) > 0);
    end;

begin //OK@@
  Result := False;
  {Devise identique}
  if (TobCptEnCours.GetString('CLE_DEVISE') <> TobBqeEnCours.GetString('CLE_DEVISE')) then Exit;

  {Comparaison du Mode de paiement de l'écriture avec le CIB du mouvement bancaire}
  ModePaieEcr := TobCptEnCours.GetString('CLE_CIB');
  {Sur les mouvements bancaires CLE_CIB contient ... le CIB, le mode de paiement est
   stocké dans CLE_EXERCICE qui n'a pas de sens pour ces mouvements}
  ModePaieRel := TobBqeEnCours.GetString('CLE_EXERCICE');
  {17/08/07 : FQ 21255 : Si au moins un cib a un mode de paiement et que le mode de paiement du relevé est vide, on sort}
  if CModeReg.Checked and FGestionCibOk and (ModePaieRel = '') then
    Exit;

  if (CModeReg.Checked and (ModePaieRel <> '') and (ModePaieRel <> ModePaieEcr)) or
     (ckLibCib.Checked and (ModePaieRel <> ModePaieEcr)) then Exit;

  if CModeReg.Checked then begin
    if CRegAccro.Checked then begin
      {Recherche de la règle correspondante}
      TobR := TobBqeEnCours;
      {Récupération de la ligne de paramétrage correspondant au CIB en cours}
      TobL := nil;
      if Assigned(TobR) then
        TobL := TobCIB.FindFirst(['TCI_CODECIB'], [TobR.GetString('CLE_CIB')], True);
      {Récupération de la règle d'accrochage}
      if Assigned(TobL) then
        Regle := TobL.GetString('TCI_REGLEACCRO');
      {Pas de règle associée}
      if Regle = '' then Exit;

      {Critères de la règle}
      TobL := TobRegles.FindFirst(['TRG_REGLEACCRO'], [Regle], True);
      if not Assigned(TobL) then Exit;
      {Date d'opération}
      if (TobL.GetValue('TRG_BDATEOPE') = 'X') and not EgalTolerant('CLE_DATECOMPTABLE') then Exit;
      {Date de valeur}
      if (TobL.GetValue('TRG_BDATEVAL') = 'X') and not EgalTolerant('CLE_DATEVALEUR') then Exit;
      if TobL.GetValue('TRG_BPIECE') = 'X' then begin
        if not VerifCheque then Exit;
      end;
    end

    {Gestion de la correspondance libellés des écritures et libellés paramétrés des CIB}
    else if ckLibCib.Checked then begin
      {On récupère le paramétrage correspondant au CIB en cours}
      TobR := TobLibCib.FindFirst(['CC_CODE'], [TobBqeEnCours.GetString('CLE_CIB')], True);

      if Assigned(TobR) then
        {Si le libellé de l'écriture en cours correspond au paramétrage du CIB, alors on prend}
        Result := CompareAvecJoker(TobR.GetString('CC_ABREGE'), TobCptEnCours.GetString('CLE_LIBELLE'));

      Exit;
    end;
  end;
  
  Result := True;
end;


{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.CocheDecoche(ARow : Integer; Value : string);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  v : string;
begin //OK@@
  T := TOB(FListe.Objects[0, ARow]);
  if not Assigned(T) then Exit;

  {On ne peut pointer lors d'une session manuelle, les mouvements intégrés}
  if FPtgeManuel and (T.GetString('CLE_MANUEL') <> 'X') and (T.GetString('CLE_JOURNAL') = LIGNEBANCAIRE) then Exit;

  {On ne peut modifier que le pointage de la session en cours ou les lignes non pointées
   SBO 28/08/2007 : j'ai débloqué le pointage pour le rapprochement multi-relevé sinon le rapprochement ne se faisait qu'à moitié,
   cad qu'il pointait les lignes d'écritures mais n'arrivait pas à pointer les lignes de relevés
   J'ai soumis cette ouverture au paramètre de pointage sécurisé
   JP 18/09/07 : je reviens en arrière car le problème d'EUROPAQUARTZ ne venait pas de là mais des devises
  if FBoSecurise then
    begin
    if (T.GetString('CLE_REFPOINTAGE') <> FStEE_RefPointage) and
       (T.GetString('CLE_REFPOINTAGE') <> '') then Exit;
    end
  else
    if (T.GetString('CLE_REFPOINTAGE') <> FStEE_RefPointage) and
       (T.GetString('CLE_REFPOINTAGE') <> '') and
       (T.GetDateTime('CLE_DATEPOINTAGE') <> iDate1900) then Exit;
  }
    if (T.GetString('CLE_REFPOINTAGE') <> FStEE_RefPointage) and
       (T.GetString('CLE_REFPOINTAGE') <> '') then Exit;

  FListe.Cells[COL_POINTE, ARow] := Value;

  v := T.GetString('CLE_POINTE');
  T.SetString('CLE_POINTE', Value);
  if Value = '' then begin
    T.SetDateTime('CLE_DATEPOINTAGE', iDate1900);
    T.SetString('CLE_REFPOINTAGE', '');
    FListe.Cells[COL_REFPOIN, ARow] := '';
  end else begin
    T.SetDateTime('CLE_DATEPOINTAGE', FDtEE_DatePointage);
    T.SetString('CLE_REFPOINTAGE', FStEE_RefPointage);
    FListe.Cells[COL_REFPOIN, ARow] := FStEE_RefPointage;
  end;
  T.SetString('MODIFIE', 'X');

  FListe.InvalidateRow(ARow);

  {Si l'ancienne ou la nouvelle valeur indique un pointage manuel, mise à jour du libellé}
  if (v = '0') or (Value = '0') then
    MajSoldeManuel(True);
end;

{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.GetLibCaption(Chp : string) : string;
{---------------------------------------------------------------------------------------}
begin
       if Chp = 'TMVTAFFICHEB'    then Result := TraduireMemoire('Mouvements pointables'   ) + ' : '
  else if Chp = 'TMVTPOINTEB'     then Result := TraduireMemoire('Mouvements pointés'      ) + ' : '
  else if Chp = 'TMVTAFFICHEE'    then Result := TraduireMemoire('Écritures pointables'    ) + ' : '
  else if Chp = 'TMVTPOINTEE'     then Result := TraduireMemoire('Écritures pointées'      ) + ' : ' {FQ 21098}
  else if Chp = 'LBMANU'          then Result := TraduireMemoire('Solde du pointage manuel') + ' : ';
end;

{Calcul de l'éclatement des soldes entre le pointé et le pointable
 #ETAT# : 23/07/07 :
  1/ Nouvelle gestion du reste à pointer : on ne tient plus compte des mouvements bancaires des relevés
  2/ Gestion d'une variable de cohérence de la session
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.Eclatement(var AEcrDeb, AEcrCre, AMvtCre, AMvtDeb : Double;
                                       var PEcrDeb, PEcrCre, PMvtCre, PMvtDeb : Double;
                                       var AMvtNb, AEcrNb, PMvtNb, PEcrNb : Integer; ManuelOk : Boolean = False);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  F : TOB;

    {--------------------------------------------------------------------}
    function _GetMontant(Chp : string) : Double;
    {--------------------------------------------------------------------}
    begin
      if (FDevise <> '') and (FDevise <> V_PGI.DevisePivot) then
        Result := F.GetDouble(Chp + 'DEV')
      else
        Result := F.GetDouble(Chp);
    end;

var
  AManDeb : Double; {#ETAT#}
  AManCre : Double; {#ETAT#}
  PManDeb : Double; {#ETAT#}
  PManCre : Double; {#ETAT#}
begin
  AEcrDeb := 0;
  AEcrCre := 0;
  AMvtCre := 0;
  AMvtDeb := 0;
  AMvtNb  := 0;
  AEcrNb  := 0;
  PEcrDeb := 0;
  PEcrCre := 0;
  PMvtCre := 0;
  PMvtDeb := 0;
  PMvtNb  := 0;
  PEcrNb  := 0;

  AManDeb := 0; {#ETAT#}
  AManCre := 0; {#ETAT#}
  PManDeb := 0; {#ETAT#}
  PManCre := 0; {#ETAT#}

//  if TobReleve.Detail.Count = 0 then Exit;

  for n := 0 to TobReleve.Detail.Count - 1 do begin
    F := TobReleve.Detail[n];
    {On est sur un mouvement bancaire}
    if F.GetString('CLE_JOURNAL') = LIGNEBANCAIRE then begin
      {On incrémente le nombre de mouvements pointables}
      Inc(AMvtNb);

      {#ETAT# : Gestion des mouvements saisis manuellement}
      if F.GetString('CLE_MANUEL') = 'X' then begin
        {Mise à jour des montants des mouvements pointables}
        AManCre := AManCre + _GetMontant('CLE_CREDIT');
        AManDeb := AManDeb + _GetMontant('CLE_DEBIT');
      end;

      {Mise à jour des montants des mouvements pointables}
      AMvtCre := AMvtCre + _GetMontant('CLE_CREDIT');
      AMvtDeb := AMvtDeb + _GetMontant('CLE_DEBIT');

      {Calcul des mouvements pointés}
      if (F.GetString('CLE_REFPOINTAGE') <> '') and (F.GetDateTime('CLE_DATEPOINTAGE') <= FDtEE_DatePointage) then begin
        {On incrémente le nombre de mouvements pointés}
        Inc(PMvtNb);
        {Mise à jour des montants des mouvements pointés}
        PMvtCre := PMvtCre + _GetMontant('CLE_CREDIT');
        PMvtDeb := PMvtDeb + _GetMontant('CLE_DEBIT');
        {#ETAT# : Gestion des mouvements saisis manuellement}
        if F.GetString('CLE_MANUEL') = 'X' then begin
          {Mise à jour des montants des mouvements pointables}
          PManCre := PManCre + _GetMontant('CLE_CREDIT');
          PManDeb := PManDeb + _GetMontant('CLE_DEBIT');
        end;
      end;
    end

    {On est sur une écriture}
    else begin
      {Mise à jour des montants des écritures pointables : cela concerne les écritures
       pointées lors de la session et celles non pointées}
      {24/04/07 : FQ 20063 : Toutes les écritures sont comme pointables, même si ce n'est pas le cas
      if FPtgeManuel or
         ((F.GetString('CLE_REFPOINTAGE') = FStEE_RefPointage) or
          (F.GetString('CLE_REFPOINTAGE') = '')) then begin
        {On incrémente le nombre d'écritures pointables}
        Inc(AEcrNb);

        AEcrCre := AEcrCre + _GetMontant('CLE_CREDIT');
        AEcrDeb := AEcrDeb + _GetMontant('CLE_DEBIT');
      //end;

      {Calcul des écritures pointées}
      if (F.GetString('CLE_REFPOINTAGE') <> '') and (F.GetDateTime('CLE_DATEPOINTAGE') <= FDtEE_DatePointage) then begin
        {On incrémente le nombre d'écritures pointées}
        Inc(PEcrNb);
        {Mise à jour des montants des écritures pointées}
        PEcrCre := PEcrCre + _GetMontant('CLE_CREDIT');
        PEcrDeb := PEcrDeb + _GetMontant('CLE_DEBIT');
      end;
    end;
  end;

  {24/04/07 : FQ 20063 : Calcul du reste à pointer.
   Remarque : pour bien faire, il faudrait trois cas, le pointage sur TRECRITURE demandant un calcul différent :
    FResteAPointe := -FSoldeBqeDeb + FSoldeBqeCre -
                     FSoldeEcrDeb + FSoldeEcrCre +
                     AEcrDeb      - AEcrCre      +
                     AManDeb      - AManCre      +
                     PManCre      - PManDeb      +
                     PEcrCre      - PEcrDeb      ;
   C'est pour cela que dans InitSoldes, j'ai inversé les soldes et je les rétablis dans MajTotalisation
   }
  if VH^.PointageJal or EstPointageSurTreso then begin (*
    FResteAPointe := FSoldeBqeDeb - FSoldeBqeCre +
                     FSoldeEcrDeb - FSoldeEcrCre +
                     AEcrDeb      - AEcrCre      +
                     AMvtDeb      - AMvtCre      +
                     PMvtCre      - PMvtDeb      +
                     PEcrCre      - PEcrDeb      ; *)
    FResteAPointe := FSoldeBqeDeb - FSoldeBqeCre +
                     FSoldeEcrDeb - FSoldeEcrCre +
                     AEcrDeb      - AEcrCre      +
                     AManDeb      - AManCre      +
                     PManCre      - PManDeb      +
                     PEcrCre      - PEcrDeb      ;
    {A vérifier, mais j'ai mis la formule dans le même sens que pour le pointage sur compte}
    FEtatPointage := FResteAPointe -
                     AMvtDeb      + AMvtCre      -
                     PMvtCre      + PMvtDeb      ;{#ETAT#}
  end
  else begin (*
    FResteAPointe := FSoldeBqeDeb - FSoldeBqeCre +
                     FSoldeEcrDeb - FSoldeEcrCre +
                     AEcrCre      - AEcrDeb      +
                     AMvtDeb      - AMvtCre      +
                     PMvtCre      - PMvtDeb      +
                     PEcrDeb      - PEcrCre      ; *)

    FResteAPointe := FSoldeBqeDeb - FSoldeBqeCre +
                     FSoldeEcrDeb - FSoldeEcrCre +
                     AEcrCre      - AEcrDeb      +
                     AManDeb      - AManCre      +
                     PManCre      - PManDeb      +
                     PEcrDeb      - PEcrCre      ;

    FEtatPointage := FResteAPointe -
                     AMvtDeb      + AMvtCre      -
                     PMvtCre      + PMvtDeb      ;{#ETAT#}
  end;

  FResteAPointe := Arrondi(FResteAPointe, FNbDecimales);
  FEtatPointage := Arrondi(FEtatPointage, FNbDecimales);{#ETAT#}
  MajImage;

  {30/07/07 : #ETAT# : Adaptation du reste à pointer}
  if ManuelOk then begin
    AMvtDeb := AManDeb;
    AMvtCre := AManCre;
    PMvtCre := PManCre;
    PMvtDeb := PManDeb;
  end;
end;

{Mise à jour des deux pavés de compteurs et totalisations au bas de l'écran
 Il y a une petite subtilité entre les écritures pointable et les écritures pointées :
 Sont pointables, les écritures non pointées et celles pointées lors de la session en cours
 Sont pointées, toutes les écritures, même celles pointées lors d'une session ultérieure
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.MajTotalisation;
{---------------------------------------------------------------------------------------}
var
  {Pointables}
  AEcrDeb : Double;
  AEcrCre : Double;
  AMvtCre : Double;
  AMvtDeb : Double;
  AMvtNb  : Integer;
  AEcrNb  : Integer;
  {Pointés}
  PEcrDeb : Double;
  PEcrCre : Double;
  PMvtCre : Double;
  PMvtDeb : Double;
  PMvtNb  : Integer;
  PEcrNb  : Integer;
  lSolde  : Double;
begin //OK@@
  Eclatement(AEcrDeb, AEcrCre, AMvtCre, AMvtDeb,
             PEcrDeb, PEcrCre, PMvtCre, PMvtDeb,
             AMvtNb, AEcrNb, PMvtNb, PEcrNb);

  {Mise à jour des contrôles de la fiche}
  SetControlCaption('TMVTAFFICHEB', GetLibCaption('TMVTAFFICHEB') + IntToStr(AMvtNb));
  SetControlCaption('TMVTPOINTEB' , GetLibCaption('TMVTPOINTEB' ) + IntToStr(PMvtNb));
  SetControlCaption('TMVTAFFICHEE', GetLibCaption('TMVTAFFICHEE') + IntToStr(AEcrNb));
  SetControlCaption('TMVTPOINTEE' , GetLibCaption('TMVTPOINTEE' ) + IntToStr(PEcrNb));

  (GetControl('TDLIGAFFICHE') as THNumEdit).Value := AMvtDeb;
  (GetControl('TCLIGAFFICHE') as THNumEdit).Value := AMvtCre;
  (GetControl('TDLIGPOINTE' ) as THNumEdit).Value := PMvtDeb;
  (GetControl('TCLIGPOINTE' ) as THNumEdit).Value := PMvtCre;
  (GetControl('TDMVTAFFICHE') as THNumEdit).Value := AEcrDeb;
  (GetControl('TCMVTAFFICHE') as THNumEdit).Value := AEcrCre;
  (GetControl('TDMVTPOINTE' ) as THNumEdit).Value := PEcrDeb;
  (GetControl('TCMVTPOINTE' ) as THNumEdit).Value := PEcrCre;

  if FResteAPointe >= 0 then begin
    (GetControl('TCAVANCEMENT') as THNumEdit).Value := FResteAPointe;
    (GetControl('TDAVANCEMENT') as THNumEdit).Value := 0;
  end else begin
    (GetControl('TDAVANCEMENT') as THNumEdit).Value := FResteAPointe * -1;
    (GetControl('TCAVANCEMENT') as THNumEdit).Value := 0;
  end;

  {25/04/07 : FQ 20067 : On n'affiche pas le détail, mais la différence entre les deux}
  lSolde := FSoldeEcrCre - FSoldeEcrDeb;
  if lSolde > 0 then begin
    FSoldeEcrCre := lSolde;
    FSoldeEcrDeb := 0;
  end else begin
    FSoldeEcrCre := 0;
    FSoldeEcrDeb := -1 * lSolde;
  end;


  {Mise à jour des soldes comptables / Tréso}
  if EstPointageSurTreso then begin
    {Sur TRECRITURE, pour simplifier les formules de calculs et les rendre identiques au pointage sur journal,
     FSoldeEcrCre est renseigné si le compte est débiteur et vis versa. Au moment de l'affichage, on rétablit}
    (GetControl('TDSOLDECOMPTABLE') as THNumEdit).Value := FSoldeEcrCre;
    (GetControl('TCSOLDECOMPTABLE') as THNumEdit).Value := FSoldeEcrDeb;
  end
  else begin
    (GetControl('TDSOLDECOMPTABLE') as THNumEdit).Value := FSoldeEcrDeb;
    (GetControl('TCSOLDECOMPTABLE') as THNumEdit).Value := FSoldeEcrCre;
  end;

  {Mise à jour des soldes bancaires}
  (GetControl('TDSOLDEBANCAIRE') as THNumEdit).Value := FSoldeBqeDeb;
  (GetControl('TCSOLDEBANCAIRE') as THNumEdit).Value := FSoldeBqeCre;
  Ecran.Refresh;
end;

{Détail du reste à pointer
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.ResteAPointer(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  {Pointables}
  AEcrDeb : Double;
  AEcrCre : Double;
  AMvtCre : Double;
  AMvtDeb : Double;
  AMvtNb  : Integer;
  AEcrNb  : Integer;
  {Pointés}
  PEcrDeb : Double;
  PEcrCre : Double;
  PMvtCre : Double;
  PMvtDeb : Double;
  PMvtNb  : Integer;
  PEcrNb  : Integer;
  Msg     : string;
  Ch, s   : string;
  {31/07/07 : Refonte de l'éclatement}
  ResteA  : Double;
  ResteP  : Double;
  Regul   : Double;
begin //OK@@
  Eclatement(AEcrDeb, AEcrCre, AMvtCre, AMvtDeb,
             PEcrDeb, PEcrCre, PMvtCre, PMvtDeb,
             AMvtNb, AEcrNb, PMvtNb, PEcrNb, True {#ETAT#});

  if VH^.PointageJal or EstPointageSurTreso then begin
    {31/07/07 : Ecritures pointables}
    ResteA := AEcrDeb - AEcrCre;
    {31/07/07 : Ecritures pointées}
    ResteP := PEcrCre - PEcrDeb;
    {31/07/07 : Régul / contestation par mouvements bancaires}
    Regul  := AMvtDeb - AMvtCre +
              PMvtCre - PMvtDeb ;
    ResteA := Arrondi(ResteA, FNbDecimales);
    ResteP := Arrondi(ResteP, FNbDecimales);
    Regul  := Arrondi(Regul , FNbDecimales);
  end
  else begin
    {31/07/07 : Ecritures pointables}
    ResteA := AEcrCre - AEcrDeb;
    {31/07/07 : Ecritures pointées}
    ResteP := PEcrDeb - PEcrCre;
    {31/07/07 : Régul / contestation par mouvements bancaires}
    Regul  := AMvtDeb - AMvtCre +
              PMvtCre - PMvtDeb ;
    ResteA := Arrondi(ResteA, FNbDecimales);
    ResteP := Arrondi(ResteP, FNbDecimales);
    Regul  := Arrondi(Regul , FNbDecimales);
  end;

  if EstPointageSurTreso then
    Ch  := PadR(TraduireMemoire('Solde de trésorerie'), ' ', 39)
  else
    Ch  := PadR(TraduireMemoire('Solde comptable'), ' ', 42);
  s   := PadL(FormateMontant(FSoldeEcrDeb - FSoldeEcrCre, FNbDecimales), ' ', 29);

  Msg := Ch + s + #13;
  Ch  := ' ' + PadR(TraduireMemoire('Solde bancaires'), ' ', 40);
  s   := PadL('-', ' ', 12) + PadL(FormateMontant(FSoldeBqeDeb - FSoldeBqeCre, FNbDecimales), ' ', 21);
  Msg := Msg + Ch + s + #13;

  Ch  := ' ' + PadR(TraduireMemoire('Ecritures pointables'), ' ', 39);
  s   := PadL('+', ' ', 12) + PadL(FormateMontant(ResteA, FNbDecimales), ' ', 21);
  Msg := Msg + Ch + s + #13;

  Ch  := ' ' + PadR(TraduireMemoire('Ecritures pointées'), ' ', 40);
  s   := PadL('+', ' ', 12) + PadL(FormateMontant(ResteP, FNbDecimales), ' ', 21);
  Msg := Msg + Ch + s + #13;

  if Regul <> 0 then begin
    Ch  := ' ' + PadR(TraduireMemoire('Régularisation'), ' ', 40);
    s   := PadL('+', ' ', 12) + PadL(FormateMontant(Regul, FNbDecimales), ' ', 21);
    Msg := Msg + Ch + s + #13;
  end;

  s   := PadL('----------------------------------', ' ', 95);
  Msg := Msg + s + #13;

  Ch  := ' ' + PadR(TraduireMemoire('Éclatement du reste à pointer'), ' ', 33);
  s   := PadL(FormateMontant(FResteAPointe, FNbDecimales), ' ', 33);
  Msg := Msg + Ch + s;

  PgiInfo(Msg, TraduireMemoire('Éclatement du reste à pointer'));
end;

{#ETAT# : pour les sessions de rappro, détail de la cohérence
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.CoherenceSession(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  MvtPble : Double;
  MvtPnte : Double;
  ResteA  : Double;
  ResteP  : Double;
  {Pointables}
  AEcrDeb : Double;
  AEcrCre : Double;
  AMvtCre : Double;
  AMvtDeb : Double;
  AMvtNb  : Integer;
  AEcrNb  : Integer;
  {Pointés}
  PEcrDeb : Double;
  PEcrCre : Double;
  PMvtCre : Double;
  PMvtDeb : Double;
  PMvtNb  : Integer;
  PEcrNb  : Integer;
  Msg     : string;
  Ch, s   : string;
begin //OK@@
  if FPtgeManuel then Exit;
  Eclatement(AEcrDeb, AEcrCre, AMvtCre, AMvtDeb,
             PEcrDeb, PEcrCre, PMvtCre, PMvtDeb,
             AMvtNb, AEcrNb, PMvtNb, PEcrNb);


  if VH^.PointageJal or EstPointageSurTreso then begin
    {31/07/07 : Ecritures pointables}
    ResteA := AEcrDeb - AEcrCre;
    {31/07/07 : Ecritures pointées}
    ResteP := PEcrCre - PEcrDeb;
    ResteA := Arrondi(ResteA, FNbDecimales);
    ResteP := Arrondi(ResteP, FNbDecimales);
  end
  else begin
    {31/07/07 : Ecritures pointables}
    ResteA := AEcrCre - AEcrDeb;
    {31/07/07 : Ecritures pointées}
    ResteP := PEcrDeb - PEcrCre;
    ResteA := Arrondi(ResteA, FNbDecimales);
    ResteP := Arrondi(ResteP, FNbDecimales);
  end;

  {A priori, il n'y a pas de différence entre pointage sur journal et sur compte, car la problèmatique est
   basée sur les mouvements bancaires}
  {Calcul des mouvements pointables}
  MvtPble := AMvtCre - AMvtDeb;
  MvtPble := Arrondi(MvtPble, FNbDecimales);

  {Calcul des mouvements pointés}
  MvtPnte := PMvtCre - PMvtDeb;
  MvtPnte := Arrondi(MvtPnte, FNbDecimales);

  if EstPointageSurTreso then
    Ch  := PadR(TraduireMemoire('Solde de trésorerie'), ' ', 39)
  else
    Ch  := PadR(TraduireMemoire('Solde comptable'), ' ', 43);
  s   := PadL(FormateMontant(FSoldeEcrDeb - FSoldeEcrCre, FNbDecimales), ' ', 29);

  Msg := Ch + s + #13;
  Ch  := ' ' + PadR(TraduireMemoire('Solde bancaires'), ' ', 40);
  s   := PadL('-', ' ', 12) + PadL(FormateMontant(FSoldeBqeDeb - FSoldeBqeCre, FNbDecimales), ' ', 21);
  Msg := Msg + Ch + s + #13;

  Ch  := ' ' + PadR(TraduireMemoire('Ecritures pointables'), ' ', 39);
  s   := PadL('+', ' ', 12) + PadL(FormateMontant(ResteA, FNbDecimales), ' ', 21);
  Msg := Msg + Ch + s + #13;

  Ch  := ' ' + PadR(TraduireMemoire('Ecritures pointées'), ' ', 40);
  s   := PadL('+', ' ', 12) + PadL(FormateMontant(ResteP, FNbDecimales), ' ', 21);
  Msg := Msg + Ch + s + #13;

  Ch  := ' ' + PadR(TraduireMemoire('Mouvements pointables'), ' ', 34);
  if MvtPble >= 0 then
    s   := PadL('+', ' ', 12) + PadL(FormateMontant(MvtPble, FNbDecimales), ' ', 21)
  else
    s   := PadL('-', ' ', 12) + PadL(FormateMontant(Abs(MvtPble), FNbDecimales), ' ', 21);

  Msg := Msg + Ch + s + #13;

  Ch  := ' ' + PadR(TraduireMemoire('Mouvements pointés'), ' ', 36);
  if MvtPnte >= 0 then
    s   := PadL('-', ' ', 12) + PadL(FormateMontant(MvtPnte, FNbDecimales), ' ', 21)
  else
    s   := PadL('+', ' ', 12) + PadL(FormateMontant(Abs(MvtPnte), FNbDecimales), ' ', 21);
  Msg := Msg + Ch + s + #13;

  s   := PadL('----------------------------------', ' ', 95);
  Msg := Msg + s + #13;

  Ch  := ' ' + PadR(TraduireMemoire('Cohérence de la session'), ' ', 38);
  s   := PadL(FormateMontant(FEtatPointage, FNbDecimales), ' ', 33);
  Msg := Msg + Ch + s;

  PgiInfo(Msg, TraduireMemoire('Éclatement de la cohérence de la session'));
end;

{Initialisation des soldes bancaires et des écritures
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.InitSoldes;
{---------------------------------------------------------------------------------------}

    {------------------------------------------------------------------------}
    function _GetRequeteSolde : string;
    {------------------------------------------------------------------------}
    var
      DtC : TDateTime;
      ChC : string;
      ChD : string;
      WhG : string; {14/12/07 : FQ 22058}
    begin
      DtC := StrToDate(E_DateComptable_.Text);

      {On commence par rechercher la Date des a-nouveaux
       28/11/07 : Déplacé dans GetDateAno
      DtA := VH^.ExoV8.Deb;
      if Dtc >= GetEnCours.Deb then
        DtA := GetEnCours.Deb
      else begin
        QuelExoDate(DtC, DtC, Mok, Exo);
        if Exo.Deb > DtA then DtA := Exo.Deb;
      end; }

      {14/12/07 : FQ 22058 : en saisie libre, certains clients ne saisissent qu'une seule écriture bancaire
                  en fin de mois pour solder toutes les opérations : dans ce cas là, ils ne peuvent pointer
                  qu'en journal et on ne peut pas récupérer le solde sur le compte bancaire, sinon cela
                  signifierait de ne faire qu'une session de pointage par mois !!!}
      if VH^.PointageJal then WhG := 'E_GENERAL <> "' + FGeneral + '"'
                         else WhG := 'E_GENERAL = "' + FGeneral + '"';

      {JP 05/06/07 : FQ 20066 : Si le compte / journal est en pivot, on prend toutes les écriture
                     sans faire de filtre sur la devise => il faut tavailler sur E_DEBIT et E_CREDIT}
      if (FDevise = '') or (FDevise = V_PGI.DevisePivot) then begin
        ChC := 'E_CREDIT';
        ChD := 'E_DEBIT';
      end else begin
        ChC := 'E_CREDITDEV';
        ChD := 'E_DEBITDEV';
      end;

      Result := 'SELECT SUM(' + ChC + ') SOLDECRE, SUM(' + ChD + ') SOLDEDEB ' +
                'FROM ECRITURE WHERE ' + WhG + ' AND E_QUALIFPIECE = "N" AND ' +
                'E_EXERCICE >= "' + ExoAno + '" AND ' + {28/11/07 Ajout de l'exercice pour essayer d'améliorer les performances}
                {On ne prend les OAN que s'il s'agit de ceux de la date de début => il faut exclure ceux d'une clôture provisoire}
                '((E_ECRANOUVEAU = "N" OR E_ECRANOUVEAU = "H") OR (E_ECRANOUVEAU="OAN" AND E_DATECOMPTABLE = "' + UsDateTime(DateAno) + '")) AND ' +
                '(E_DATECOMPTABLE BETWEEN "' + UsDateTime(DateAno) + '" AND "' + USDateTime(DtC) + '") ';

      if (VH^.ExoV8.Code <> '') then
        Result := Result + 'AND E_DATECOMPTABLE >= "' + USDateTime(VH^.ExoV8.Deb) + '"';

      if (FDevise <> '') and (FDevise <> V_PGI.DevisePivot) then
        Result := Result + 'AND E_DEVISE = "' + FDevise + '"';
        // RecupAutreWhere
    end;
var
  m : Double;
  Q : TQuery;
begin //OK@@
  {Les soldes bancaires sont mis à jour au chargement de la Fiche}
//  (GetControl('TDSOLDEBANCAIRE') as THNumEdit).Value := FSoldeBqeDeb;
  //(GetControl('TCSOLDEBANCAIRE') as THNumEdit).Value := FSoldeBqeCre;

  {$IFDEF TRESO}
  if EstPointageSurTreso then begin
    m := GetSoldeBancaire(FGeneral, ' AND TE_NATURE = "' + na_Realise + '"', FDtEE_DatePointage, False, FDevise = V_PGI.DevisePivot);
    if m >= 0 then FSoldeEcrDeb := Valeur(StrfMontant(m, 15, FNbDecimales, '', True))
              else FSoldeEcrCre := Valeur(StrfMontant(m * - 1, 15, FNbDecimales, '', True));
  end;
  {$ENDIF TRESO}
  {Calcul des soldes sur ECRITURE / TRECRITURE}
  if not EstPointageSurTreso then begin
    Q := OpenSQL(_GetRequeteSolde, True, -1, '', True);
    try
      if not Q.Eof then begin
        {28/01/08 : FQ 22305 : suite à la FQ 22058 où l'on calcule le solde à partir des écritures qui ne sont
                    pas sur le compte bancaire, j'avais oublié d'inverser les soldes}
        if VH^.PointageJal then begin
          FSoldeEcrDeb := Valeur(StrfMontant(Q.FindField('SOLDECRE').AsFloat, 15, FNbDecimales, '', True));
          FSoldeEcrCre := Valeur(StrfMontant(Q.FindField('SOLDEDEB').AsFloat, 15, FNbDecimales, '', True));
        end else begin
          FSoldeEcrDeb := Valeur(StrfMontant(Q.FindField('SOLDEDEB').AsFloat, 15, FNbDecimales, '', True));
          FSoldeEcrCre := Valeur(StrfMontant(Q.FindField('SOLDECRE').AsFloat, 15, FNbDecimales, '', True));
        end;
      end else begin
        FSoldeEcrDeb := 0.00;
        FSoldeEcrCre := 0.00;
      end;
    finally
      Ferme(Q);
    end;
  end;

  if not EstPointageSurTreso then begin
    {25/04/07 : FQ 20067 : On n'affiche pas le détail, mais la différence entre les deux}
    m := FSoldeEcrCre - FSoldeEcrDeb;
    if m > 0 then begin
      FSoldeEcrCre := m;
      FSoldeEcrDeb := 0;
    end else begin
      FSoldeEcrCre := 0;
      FSoldeEcrDeb := -1 * m;
    end;
  end;

//  (GetControl('TDSOLDECOMPTABLE') as THNumEdit).Value := FSoldeEcrDeb;
  //(GetControl('TCSOLDECOMPTABLE') as THNumEdit).Value := FSoldeEcrCre;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.PointageManuel(ARow : Integer);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  v : string;
begin //OK@@
  if FBoConsultation then Exit;

  T := TOB(FListe.Objects[0, ARow]);
  if not Assigned(T) then Exit;

  {On ne peut pointer lors d'une session manuelle, les mouvements intégrés}
  if FPtgeManuel and (T.GetString('CLE_MANUEL') <> 'X') and (T.GetString('CLE_JOURNAL') = LIGNEBANCAIRE) then Exit;

  v := T.GetString('CLE_POINTE');

  if (EtatAffichage <> tv_Tout) and (v > '0')then Exit;

  {On commence par vider la ligne courantes}
  CocheDecoche(ARow, '');
  {Si l'ancien contenu de la ligne n'était pas '0' ...}
  if v <> '0' then begin
    {... On met '0' sur la ligne courante}
    CocheDecoche(ARow, '0');
    {Si la ligne avait été pointée lors d'un pointage automatique}
    if v <> '' then begin
      {... on dépointe les éventuelles lignes rattachées au même pointage}
      DepointeGroupe(v);
    end;
  end;

  {Mise à jour des totaux}
  MajTotalisation;

  {18/09/07 : FQ 21443 : On passe à la ligne suivante}
  if FPtgeManuel and (ARow < FListe.RowCount - 1) then FListe.Row := FListe.Row + 1; 
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.DepointeGroupe(Value : string);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin //OK@@
  if FPtgeManuel then Exit;
  if FBoConsultation then Exit;

  if EtatAffichage in [tv_Comptable, tv_Bancaire, tv_NonPointe] then Exit;

  for n := 1 to FListe.RowCount -1 do
    {On dépointe toute les lignes qui ont le même code pointage sauf le pointage manuel et
     qui appartienne à la session de pointage en cours}
    if (FListe.Cells[COL_POINTE, n] = Value) and (Value <> '0') and
       (FListe.Cells[COL_REFPOIN, n] = FStEE_RefPointage) then
      CocheDecoche(n, '');

  {Mise à jour des totaux}
  MajTotalisation;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.CalcPointageManuel : Double;
{---------------------------------------------------------------------------------------}
var
  T   : TOB;
  n   : Integer;
  Chp : string;
begin
  Result := 0;
  if FPtgeManuel or (TobReleve.Detail.Count = 0) then Exit;
  if TobReleve.Detail[0].GetString('CLE_DEVISE') = V_PGI.DevisePivot then Chp := 'CLE_MONTANT'
                                                                     else Chp := 'CLE_MONTANTDEV';

  for n := 0 to TobReleve.Detail.Count -1 do begin
    T := TobReleve.Detail[n];
    {On ne prend pas en compte les mouvements des autres sessions de pointage}
    if (T.GetString('CLE_REFPOINTAGE') <> FStEE_RefPointage) and (T.GetString('CLE_REFPOINTAGE') <> '') then Continue;
    if T.GetString('CLE_POINTE') = '0' then begin
      if T.GetString('CLE_JOURNAL') = LIGNEBANCAIRE then Result := Result - Arrondi(T.GetDouble(Chp), FNbDecimales)
                                                    else Result := Result + Arrondi(T.GetDouble(Chp), FNbDecimales);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.MajSoldeManuel(EnPointage : Boolean = False);
{---------------------------------------------------------------------------------------}
var
  M : Double;
  c : string;
begin //OK@@
  if FPtgeManuel then Exit;
  if FBoConsultation then Exit;

  c := GetLibCaption('LBMANU');
  {Calcul du pointage à "0"}
  M := CalcPointageManuel;
  M := Arrondi(M, FNbDecimales);

  if M = 0 then begin
    SetControlCaption('LBMANU', c + '0.00');
    (GetControl('LBMANU') as THLabel).Font.Color := clMaroon;
    {Le pointage manuel est équilibré, on propose un code de regroupement pour les écritures / mouvements à '0'}
    if EnPointage then RegroupementManuel;
  end
  else if M > 0 then begin
    SetControlCaption('LBMANU', c + StrFMontant(M, 15, FNbDecimales, '', True) + ' C');
    (GetControl('LBMANU') as THLabel).Font.Color := clGreen;
  end else begin
    SetControlCaption('LBMANU', c + StrFMontant(-1 * M, 15, FNbDecimales, '', True) + ' D');
    (GetControl('LBMANU') as THLabel).Font.Color := clRed;
  end;
end;

{En pointage manuel, si la sélection est équilibré, on propose un nouveau Code de regroupement}
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.RegroupementManuel;
{---------------------------------------------------------------------------------------}
var
  n  : Integer;
  T  : TOB;
  c  : string;
  Ok : Boolean;
begin //OK@@
  if TobReleve.Detail.Count = 0 then Exit;

  if FBoConsultation then Exit;
  Ok := False;

  {On commence par s'assurer que l'équilibre n'est pas du à l'absence de pointage manuel !}
  for n := 0 to TobReleve.Detail.Count -1 do
    if TobReleve.Detail[n].GetString('CLE_POINTE') = '0' then begin
      Ok := True;
      Break;
    end;


  if Ok and (
     PGIAsk(TraduireMemoire('Le pointage manuel est équilibré.') + #13 +
            TraduireMemoire('Souhaitez-vous regrouper ces mouvements ?'), Ecran.Caption) = mrYes) then begin
    {Récupération du nouveau code de regroupement}
    c := GetMaxNumPtge;

    {Mise à jour du code de regroupement dans la Tob}
    for n := 0 to TobReleve.Detail.Count -1 do begin
      T := TobReleve.Detail[n];
      {On ne touche pas aux autres sessions de pointages}
      if (T.GetString('CLE_REFPOINTAGE') <> FStEE_RefPointage) and
         (T.GetString('CLE_REFPOINTAGE') <> '') then Continue;

      if T.GetString('CLE_POINTE') = '0' then begin
        T.SetString('CLE_POINTE', c);
        T.SetString('MODIFIE', 'X');
      end;
    end;

    {Mise à jour du code de regroupement dans la grille, qui peut-être filtrée}
    for n := 1 to FListe.RowCount - 1 do begin
      {On ne touche pas aux autres sessions de pointages}
      if (FListe.Cells[COL_REFPOIN, n] <> FStEE_RefPointage) and
         (FListe.Cells[COL_REFPOIN, n] <> '') then Continue;

      if FListe.Cells[COL_POINTE, n] = '0' then begin
        FListe.Cells[COL_POINTE, n] := c;
        FListe.InvalidateRow(n);
      end;
    end;

    {Remise à zéro du libellé contenant le cumul du pointage manuel}
    MajSoldeManuel;
  end;
end;

{Méthode appelée avant le lancement d'un autre traitement
{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.MajPointage : Boolean;
{---------------------------------------------------------------------------------------}
begin //OK@@
  if IsPointageEncours then begin
    Result := False;
    {23/07/07 FQ 21058 : gestion des messages d'équilibre}
    if FPasMessage or (PGIAsk(TraduireMemoire('Voulez-vous enregistrer le pointage en cours ?'), Ecran.Caption) = mrYes) then begin
      FPasMessage := False;
      {On s'assure en rapprochement que le pointage manuel est équilibré et que les dates de valeur sont cohérentes}
      FPeutFermer := IsPointageManuelOk and IsDateValeurOk;
      if FPeutFermer then begin
        EcritPointage(False);
        Result := True;
      end;
    end;
  end
  else begin
    {Avec la saisie automatique, les écritures sont pointées en arrivant dans la fiche.
     Si en première ouverture de fiche toutes les écritures sont pointées, on ne passe
     pas dans EcritPointage, il faut dont mettre à jour EE_AVANCEMENT}
    MajAvancement;
    Result := True;
  end;
end;

{Regarde s'il y a un pointage en cours
{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.IsPointageEncours : Boolean;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  Result := False; 

  if TobReleve.Detail.Count = 0 then Exit;

  for n := 0 to TobReleve.Detail.Count - 1 do
    if TobReleve.Detail[n].GetString('MODIFIE') = 'X' then begin
      Result := True;
      Break;
    end;
end;

{Repasse le champ 'MODIFIE' de la tob à False : à utiliser quand on veut que IsPointageEncours
 renvoie définitivement False
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.AnnulePointage;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  if TobReleve.Detail.Count = 0 then Exit;

  for n := 0 to TobReleve.Detail.Count - 1 do
    TobReleve.Detail[n].SetString('MODIFIE', '-');
end;

{En Rappro, on regarde si le pointage manuel est équilibré
{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.IsPointageManuelOk : Boolean;
{---------------------------------------------------------------------------------------}
var
  PtgManOk : Boolean;
  MntMan   : Double;
begin //OK@@
  if not FPtgeManuel then begin
    MntMan := CalcPointageManuel;
    PtgManOk := Arrondi(MntMan, FNbDecimales) = 0;
    {Le pointage n'est pas équilibré}
    if not PtgManOk then
      Result := (PGIAsk(TraduireMemoire('Le pointage manuel n''est pas équilibré, il ne sera pas enregistré.') + #13 +
                       TraduireMemoire('Voulez-vous poursuivre ?'), Ecran.Caption) = mrYes)
    else
      Result := True;
  end
  else
    {Dans une session de pointage manuel, on renvoie toujours vrai}
    Result := True;
end;

{Si avec Tréso, on avertit si les date de valeur sont <> entre le relevé et les écritures}
{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.IsDateValeurOk : Boolean;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  F : TOB;
  B : TOB;
begin
  FMajDateValeur := False;
  Result := True;
  {25/07/07 : Ce traitement n'a aucun intérêt en pointage manuel}
  if FPtgeManuel or not IsPointageEncours then Exit;

  {Ne concerne que ???}
  if {not EstComptaTreso or FPtgeManuel or} VH^.PointageJal then Exit;

  for n := 0 to TobReleve.Detail.Count - 1 do begin
    F := TobReleve.Detail[n];
    {On est sur un mouvement bancaire, sur une ligne non modifiée ou bien dépointée}
    if (F.GetString('MODIFIE') <> 'X') or (F.GetString('CLE_POINTE') = '') or
       (F.GetString('CLE_JOURNAL') = LIGNEBANCAIRE) then Continue;

    {On recherche la ligne de relevé correspondant à l'écriture en cours}
    B := TobReleve.FindFirst(['CLE_JOURNAL', 'CLE_POINTE'], [LIGNEBANCAIRE, F.GetString('CLE_POINTE')], True);
    if Assigned(B) then begin
      {Si la date du relevé est différente de celles des écritures ...}
      if F.GetDateTime('CLE_DATEVALEUR') <> B.GetDateTime('CLE_DATEVALEUR') then begin
        Result := False;
        Break;
      end;
    end;
  end;

  if not Result then begin
    {En compta, on ne pose pas de question de confirmation : Maj automatique}
    if not (CtxTreso in V_PGI.PGIContexte) then begin
      Result := True;
      FMajDateValeur := True;
      Exit;
    end;

    {11/05/07 : FQ TRESO 10461 : Orthographe + vocabulaire}
    case PGIAskCancel(TraduireMemoire('Certaines dates de valeur des écritures sont différentes de celles des relevés.') + #13 +
                      TraduireMemoire('Souhaitez-vous : ') + #13 +
                      TraduireMemoire('- Reprendre la date de valeur du relevé (Oui) ?') + #13 +
                      TraduireMemoire('- Ne pas modifier la date de valeur (Non) ?') + #13 +
                      TraduireMemoire('- Annuler la mise à jour du pointage (Annuler) ?'), Ecran.Caption) of
      mrYes    : begin
                   Result := True;
                   FMajDateValeur := True;
                 end;
      mrNo     : Result := True;
      mrCancel : Result := False;
    end;
  end;
end;

{Traitement du BChercheClik
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.RafraichitFiche(AvecMaj : Byte = 1);
{---------------------------------------------------------------------------------------}
begin
  if AvecMaj > 0 then
   {Si on n'est en rafraîchissement ou en changement d'affichage (Débit / Crédit ou Montant),
    AvecMaj = 2 => même si la réponse est négative, on rafraîchit}
   if not MajPointage and (AvecMaj = 1) then Exit;
   
  {Initialisation des soldes}
  InitSoldes;
  {Chargement de la Tob TobReleve contenant les mouvements et les écritures}
  ChargerMouvements;
  {Affichage dans la grille de la Tob TobReleve}
  ChargeGrille;
  {Definition de la grille : largeur, titres ...}
  PresenteGrille;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.MetEnModeLecture(SaufVVert : Boolean = False);
{---------------------------------------------------------------------------------------}
var
  Ok : Boolean;
begin
  GereTypeSession;
  {Le panel contient le détail des lignes}
  SetControlVisible('PCUMUL' , True);
  SetControlVisible('BSAISIE', not (CtxTreso in V_PGI.PGIContexte));
  SetControlVisible('LBMANU' , not FPtgeManuel);
  SetControlVisible('LBMANU' , not FPtgeManuel);

  SetControlEnabled('BPOINTAGEAUTO'  , not FPtgeManuel and not FBoConsultation);
  SetControlEnabled('BVALIDEETRAPPRO', not FBoConsultation);
  {28/11/07 : vu chez SIC : si le compte est peu mouvementé, il faut pouvoir valider
              une session de pointage même s'il n'y a ni écritures ni mouvements}
  Ok := not FBoConsultation;
  if not Ok and SaufVVert then Ok := True; 
  SetControlEnabled('BVALIDER'       , Ok);
  FSaufVVert := SaufVVert;

  SetControlEnabled('FLISTE'         , not FBoConsultation);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.EcritPointage(AvecMsg : Boolean = True);
{---------------------------------------------------------------------------------------}
type
  lDossierBase = record
    NoDossier : string;
    NomBase   : string;
  end;

var
  T : TOB;
  n : Integer;
  NomBase  : string;
  lDossier : array of lDossierBase;
  RecupOk  : Boolean;

    {-------------------------------------------------------------------}
    procedure _GetNomBase(aNoDossier : string);
    {-------------------------------------------------------------------}
    var
      Q : TQuery;
      n : Integer;
    begin
      n := 0;
      if not IsTresoMultiSoc then Exit;

      try
        if not RecupOk then begin
          {17/09/07 : FQ 10518 : l'unidirectionnel est incompatible avec les RecordCount}
          Q := OpenSQL('SELECT DOS_NOMBASE, DOS_NODOSSIER FROM DOSSIER', True{, -1, '', True});
          SetLength(lDossier, Q.RecordCount);
          while not Q.Eof do begin
            lDossierBase(lDossier[n]).NoDossier := Q.FindField('DOS_NODOSSIER').AsString;
            lDossierBase(lDossier[n]).NomBase   := Q.FindField('DOS_NOMBASE').AsString;
            Inc(n);
            Q.Next;
          end;
          Ferme(Q);
          RecupOk := True;
        end;

        NomBase := '';
        for n := Low(lDossier) to High(lDossier) do begin
          if lDossierBase(lDossier[n]).NoDossier = aNoDossier then begin
            NomBase := lDossierBase(lDossier[n]).NomBase;
            Break;
          end;
        end;
        if NomBase = '' then raise Exception.Create(TraduireMemoire('Le dossier comptable n''a pu être récupéré.'));
      except
        on E : Exception  do begin
          PgiError(E.Message);
        end;
      end;
    end;

    {-------------------------------------------------------------------}
    function _SynchroS1 : string;
    {-------------------------------------------------------------------}
    begin
      Result := 'UPDATE ECRITURE SET E_IO = "X" WHERE ' +
                'E_JOURNAL = "' + T.GetString('CLE_JOURNAL') + '" AND ' +
                'E_EXERCICE = "' + T.GetString('CLE_EXERCICE') + '" AND ' +
                'E_PERIODE = ' + IntToStr(GetPeriode(T.GetValue('CLE_DATECOMPTABLE'))) + ' AND ' +
                'E_NUMEROPIECE = ' + T.GetString('CLE_NUMEROPIECE') + ' AND ' +
                'E_QUALIFPIECE = "N"';
    end;



var
  PtgManOk : Boolean;
begin //OK@@
  if TobReleve.Detail.Count = 0 then Exit;

  if FBoConsultation then Exit;

  {Si on n'est pas pointage manuel, on interdit l'écriture d'un pointage à "0" non équilibré}
  PtgManOk := FPtgeManuel or (Arrondi(CalcPointageManuel, FNbDecimales) = 0);

  RecupOk := False;

  BeginTrans;
  try
    for n := 0 to TobReleve.Detail.Count -1 do begin
      T := TobReleve.Detail[n];
      {En Pointage auto, si le pointage manuel n'est pas équilibré, on n'enregistre pas}
      if not PtgManOk and (T.GetString('CLE_POINTE') = '0') then begin
        T.SetString('MODIFIE', '-');
        Continue;
      end;

      {On ne met à jour que les lignes modifiées}
      if T.GetString('MODIFIE') <> 'X' then Continue
                                       else T.SetString('MODIFIE', '-');

      {Mouvement bancaire}
      if T.GetString('CLE_JOURNAL') = LIGNEBANCAIRE then
        ExecuteSQL('UPDATE EEXBQLIG SET CEL_CODEPOINTAGE = "' + T.GetString('CLE_POINTE') + '", ' +
                                       'CEL_REFPOINTAGE = "'  + T.GetString('CLE_REFPOINTAGE') + '", ' +
                                       'CEL_DATEPOINTAGE = "' + UsDateTime(T.GetDateTime('CLE_DATEPOINTAGE')) + '" WHERE ' +
                   'CEL_GENERAL = "'  + T.GetString('CLE_GENERAL') + '" AND ' +
                   'CEL_NUMRELEVE = ' + T.GetString('CLE_NUMEROPIECE') + ' AND ' +
                   'CEL_NUMLIGNE = '  + T.GetString('CLE_NUMLIGNE'))

      {Pointage sur TRECRITURE}
      else if EstPointageSurTreso then begin
        {Mise à jour des dates de valeurs, si celle du relevé est différente de la pièce}
        MajDtValeur(T);

        ExecuteSQL('UPDATE TRECRITURE SET TE_CODERAPPRO = "' + T.GetString('CLE_POINTE') + '", ' +
                                  'TE_DATEVALEUR = "' + UsDateTime(T.GetDateTime('CLE_DATEVALEUR')) + '", ' +
                                  'TE_DATERAPPRO = "' + UsDateTime(T.GetDateTime('CLE_DATEPOINTAGE')) + '", ' +
                                  'TE_REFPOINTAGE = "' + T.GetString('CLE_REFPOINTAGE') + '" WHERE ' +
                   'TE_NODOSSIER = "'  + T.GetString(TRESO_NODOSSIER) + '" AND ' +
                   'TE_NUMTRANSAC = "'  + T.GetString(TRESO_NUMTRANSAC) + '" AND ' +
                   'TE_NUMEROPIECE = ' + T.GetString('CLE_NUMEROPIECE') + ' AND ' +
                   'TE_NUMLIGNE = ' + T.GetString('CLE_NUMLIGNE'));

        {L'ecriture existe-t-elle en compta ?}
        if T.GetString(TRESO_USERCOMPTA) = '' then Continue;
        {Récupération de la base dans laquelle figure l'écriture en compta}
        _GetNomBase(T.GetString(TRESO_NODOSSIER));

        ExecuteSQL('UPDATE ' + GetTableDossier(NomBase, 'ECRITURE') + ' SET ' +
                           'E_DATEVALEUR = "' + UsDateTime(T.GetDateTime('CLE_DATEVALEUR')) + '", ' +
                           'E_NATURETRESO = "' + T.GetString('CLE_POINTE') + '", ' +
                           'E_DATEPOINTAGE = "' + UsDateTime(T.GetDateTime('CLE_DATEPOINTAGE')) + '", ' +
                           'E_REFPOINTAGE = "' + T.GetString('CLE_REFPOINTAGE') + '" WHERE ' +
                   'E_JOURNAL = "'  + T.GetString('CLE_JOURNAL') + '" AND ' +
                   'E_EXERCICE = "'  + T.GetString('CLE_EXERCICE') + '" AND ' +
                   'E_DATECOMPTABLE = "'  + UsDateTime(T.GetDateTime('CLE_DATECOMPTABLE')) + '" AND ' +
                   'E_QUALIFPIECE = "N" AND ' +
                   'E_NUMEROPIECE = ' + T.GetString('CLE_NUMEROPIECE') + ' AND ' +
                   'E_NUMLIGNE = ' + T.GetString(TRESO_NUMLIGNE) + ' AND ' +
                   'E_NUMECHE = '  + T.GetString('CLE_NUMECHE'));

        {L'Ecriture est modifiée, on met à jour E_IO de la pièce pour Synchro S1}
        ExecuteSQL(_SynchroS1);
      end

      {Pointage sur ECRITURE}
      else begin
        {Mise à jour des dates de valeurs, si celle du relevé est différente de la pièce}
        MajDtValeur(T);

        ExecuteSQL('UPDATE ECRITURE SET E_NATURETRESO = "' + T.GetString('CLE_POINTE') + '", ' +
                                       'E_DATEVALEUR = "' + UsDateTime(T.GetDateTime('CLE_DATEVALEUR')) + '", ' +
                                       'E_DATEPOINTAGE = "' + UsDateTime(T.GetDateTime('CLE_DATEPOINTAGE')) + '", ' +
                                       'E_REFPOINTAGE = "' + T.GetString('CLE_REFPOINTAGE') + '" WHERE ' +
                   'E_JOURNAL = "'  + T.GetString('CLE_JOURNAL') + '" AND ' +
                   'E_EXERCICE = "'  + T.GetString('CLE_EXERCICE') + '" AND ' +
                   'E_DATECOMPTABLE = "'  + UsDateTime(T.GetDateTime('CLE_DATECOMPTABLE')) + '" AND ' +
                   'E_QUALIFPIECE = "'  + T.GetString('CLE_QUALIFPIECE') + '" AND ' +
                   'E_NUMEROPIECE = ' + T.GetString('CLE_NUMEROPIECE') + ' AND ' +
                   'E_NUMLIGNE = ' + T.GetString('CLE_NUMLIGNE') + ' AND ' +
                   'E_NUMECHE = '  + T.GetString('CLE_NUMECHE'));

        {L'Ecriture est modifiée, on met à jour E_IO de la pièce pour Synchro S1}
        ExecuteSQL(_SynchroS1);

        if EstComptaTreso then
          ExecuteSQL('UPDATE TRECRITURE SET ' +
                             'TE_DATEVALEUR = "' + UsDateTime(T.GetDateTime('CLE_DATEVALEUR')) + '", ' +
                             'TE_CODERAPPRO = "' + T.GetString('CLE_POINTE') + '", ' +
                             'TE_REFPOINTAGE = "' + T.GetString('CLE_REFPOINTAGE') + '", ' +
                             'TE_DATERAPPRO = "' + UsDateTime(T.GetDateTime('CLE_DATEPOINTAGE')) + '" WHERE ' +
                   'TE_JOURNAL = "'  + T.GetString('CLE_JOURNAL') + '" AND ' +
                   'TE_EXERCICE = "'  + T.GetString('CLE_EXERCICE') + '" AND ' +
                   'TE_DATECOMPTABLE = "'  + UsDateTime(T.GetDateTime('CLE_DATECOMPTABLE')) + '" AND ' +
                   'TE_NUMEROPIECE = ' + T.GetString('CLE_NUMEROPIECE') + ' AND ' +
                   'TE_CPNUMLIGNE = ' + T.GetString('CLE_NUMLIGNE') + ' AND ' +
                   'TE_NUMECHE = '  + T.GetString('CLE_NUMECHE'));
      end; {if ... else ...}
    end; {for ...}

    {Mise à jour des totaux pointés}
    SetTotauxPointes;

    {Mise à jour de l'avancement du pointage}
    MajAvancement;

    CommitTrans;

    {23/07/07 : FQ 21053 et 21058 : ce message me semble superfétatoire
    if AvecMsg then PGIInfo(TraduireMemoire('La mise à jour s''est correctement effectuée'), Ecran.Caption);}
  except
    on E : Exception do begin
      RollBack;
      PGIError(TraduireMemoire('Une erreur est intervenue lors de l''enregistrement du pointage, avec le message :') + #13#13 + E.Message, Ecran.Caption);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.MajDtValeur(aTob : TOB);
{---------------------------------------------------------------------------------------}
var
  MajOk : Boolean;
  NewDt : TDateTime;
  OldDt : TDateTime;

    {-----------------------------------------------------------------------}
    function IsDateCoherente : Boolean;
    {-----------------------------------------------------------------------}
    var
      F : TOB;
    begin
      Result := True;
      NewDt  := iDate1900;
      {On recherche la ligne de relevé correspondant à l'écriture en cours}
      F := TobReleve.FindFirst(['CLE_JOURNAL', 'CLE_POINTE'], [LIGNEBANCAIRE, aTob.GetString('CLE_POINTE')], True);
      if Assigned(F) then begin
        {Si la date du relevé est différente de celles des écritures ...}
        if F.GetDateTime('CLE_DATEVALEUR') <> aTob.GetDateTime('CLE_DATEVALEUR') then begin
          Result := False;
          {... la date du relevé s'impose}
          NewDt := F.GetDateTime('CLE_DATEVALEUR');
        end;
      end;
    end;

    {$IFDEF TRESO}
    {21/11/06 : recalcul des soldes initiaux
    {-----------------------------------------------------------------------}
    procedure CalcSoldeInit(Mnt : Double);
    {-----------------------------------------------------------------------}
    var
      Q : TQuery;
      D : string;
      T : Double;
    begin
      D := V_PGI.DevisePivot;
      Q := OpenSQL('SELECT BQ_DEVISE FROM BANQUECP WHERE BQ_CODE = "' + GetControlText('CBANQUE') + '"', True);
      if not Q.EOF then D := Q.FindField('BQ_DEVISE').AsString;
      Ferme(Q);
      if D <> V_PGI.DevisePivot then begin
        T := RetPariteEuro(D, NewDt, True);
        Mnt := Mnt * T;
      end;
      GereSoldeInit(OldDt, NewDt, GetControlText('CBANQUE'), Mnt, True);
    end;
    {$ENDIF TRESO}

begin //OK@@
  MajOk := False;
  NewDt := iDate1900;

  if {(not EstComptaTreso) or FPtgeManuel or} VH^.PointageJal then Exit;
  {Si l'écriture a été dépointée, on ne fait rien !!!}
  if aTob.GetString('CLE_POINTE') = '' then Exit;

  {Si les dates sont différentes et qu'on a demandé une mise à jour des dates de valeur}
  if not IsDateCoherente and FMajDateValeur then
    MajOk := True;

  {Mise à jour des tables ECRITURE et TRECRITURE}
  if MajOk and (NewDt > iDate1900) then begin
    OldDt := aTob.GetDateTime('E_DATEVALEUR');
    aTob.SetDateTime('CLE_DATEVALEUR', NewDt);
    {Nécessaire car pour une raison que j'ignore, CLE_DATEVALEUR n'est pas passé à Modifié}
    aTob.SetModifieField('CLE_DATEVALEUR', True);
    if (DebutAnnee(OldDt) <> DebutAnnee(NewDt)) {and not IniOk} then begin
      {$IFDEF TRESO}
      CalcSoldeInit(aTob.GetDouble('CLE_MONTANT'));
      {$ENDIF TRESO}
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.LancerCib(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  if FPtgeManuel then Exit;
  if ctxPcl in V_PGI.PGIContexte then
    TRLanceFiche_MulCIB('TR', 'TRMULCIB', '', '', tc_Reference)
  else
    TRLanceFiche_CIB('TR','TRCIB','','','ACTION=MODIFICATION;' + tc_CIB);
  ChargeLesCIB;
end;

{16/08/07 : FQ 13017 : recherche des pièces sur un autre journal
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.ContrepartieDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if not VH^.PointageJal then Exit;
  GetPieceAutreJournal;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.LancerRegle(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  if FPtgeManuel then Exit;
  TRLanceFiche_ParamRegleAccro('TR', 'TRFICHEACCROCHAGE', '', '', '');
  TobRegles.ClearDetail;
  TobRegles.LoadDetailFromSQL('SELECT * FROM REGLEACCRO');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.LancerLibCib(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if FPtgeManuel then Exit;
  ParamTable('CPRAPPROSURLIB', taCreat, 50000147, nil); {18/09/07 : FQ 21393 : Pour le HelpContext}
  TobLibCib.ClearDetail;
  TobLibCib.LoadDetailFromSQL('SELECT CC_CODE, CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE = "TCI"');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.LancerMulEexBqLig(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  if not FBoDroitMouvement or FPtgeNonBqe then Exit;
  CPLanceFiche_MulEexBqLig('', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.CreerEexBqLig(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  if not FBoDroitMouvement or FPtgeNonBqe then Exit;
  if FPtgeManuel then {18/07/07 : FQ 21126 : qu'en pointage manuel, en particulier pour le raccourci clavier}
    LanceFicheEexBqLig(taCreat);
end;

{01/06/07 : Ajout à l'actionfiche, la fiche d'origine
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.LanceFicheEexBqLig(Action : TActionfiche);
{---------------------------------------------------------------------------------------}
var
  Code : string;
  T    : TOB;
begin //OK@@
  if FPtgeNonBqe then Exit;
  
  if VH^.PointageJal then Code := FJournal
                     else Code := FGeneral;

  {On s'assure que l'on a les droits, notamment avec les raccourcis clavier
   01/06/07 : FQ TRESO 10471 : On affiche quand même la ligne en lecture seule}
  if not FBoDroitMouvement or (FBoConsultation and not FBoPeutCreer) then begin
    {12/07/07 : Ajout de cettre précaution}
    if FListe.Cells[COL_PIECE, FListe.Row] <> '' then
      CPLanceFiche_EEXBQLIG('', Code + ';' + FListe.Cells[COL_PIECE, FListe.Row] + ';' + FListe.Cells[COL_LIGNE, FListe.Row] + ';', ActionToString(taConsult) + ';' + FO_POINTAGE + ';');
    Exit;
  end;

  if Action = taCreat then begin
    TheTob := nil;
    T := TOB.Create('µRELEVE', nil, -1);
    try
      AddChampPointage(T);
      T.PutValue('CLE_JOURNAL'      , LIGNEBANCAIRE);
      T.PutValue('CLE_GENERAL'      , Code);
      T.PutValue('CLE_DEVISE'       , FDevise);
      T.PutValue('CLE_DATECOMPTABLE', FDtEE_DatePointage);
      T.PutValue('CLE_DATEVALEUR'   , FDtEE_DatePointage);
      T.PutValue('CLE_CIB'          , '');
      T.PutValue('CLE_POINTE'       , '');
      T.PutValue('CLE_NUMEROPIECE'  , FItEE_Numero);
      TheTob := T;
      CPLanceFiche_EEXBQLIG('', '', ActionToString(taCreat) + ';' + FO_POINTAGE + ';' + Code + ';' + IntToStr(FItEE_Numero) + ';' + DateToStr(FDtEE_DatePointage) + ';');
      if TheTob <> nil then AjouteMvt(TheTob);
    finally
      TheTob := nil;
      FreeAndNil(T);
    end;
  end

  {Accès en consultation
  {12/07/07 : FQ 21052 : on accède toujours au mouvement en consultation pour les mouvements pointés
   18/07/07 : FQ 21126 : en consultation s'il s'agit d'une session de pointage automatique}
  else begin
    if ExJaiLeDroitConcept(ccSaisMvtBqe, True) and (FListe.Cells[COL_POINTE, FListe.Row] = '') and FPtgeManuel then begin
      {On propose d'enregistrer le pointage}
      if MajPointage then begin
        CPLanceFiche_EEXBQLIG('', Code + ';' + FListe.Cells[COL_PIECE, FListe.Row] + ';' + FListe.Cells[COL_LIGNE, FListe.Row] + ';', ActionToString(taModif) + ';' +
                                  FO_POINTAGE + ';' + Code + ';' + IntToStr(FItEE_Numero) + ';' + DateToStr(FDtEE_DatePointage) + ';');
        RafraichitFiche(0);
      end
      else
        CPLanceFiche_EEXBQLIG('', Code + ';' + FListe.Cells[COL_PIECE, FListe.Row] + ';' + FListe.Cells[COL_LIGNE, FListe.Row] + ';', ActionToString(taConsult) + ';' + FO_POINTAGE + ';')
    end
    else
      CPLanceFiche_EEXBQLIG('', Code + ';' + FListe.Cells[COL_PIECE, FListe.Row] + ';' + FListe.Cells[COL_LIGNE, FListe.Row] + ';', ActionToString(taConsult) + ';' + FO_POINTAGE + ';')
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.LancerSaisie;
{---------------------------------------------------------------------------------------}
var
  TobLig  : TOB;
  F       : TOB;
{$IFDEF TRESO}
  Lequel  : string;
{$ELSE}
  Action  : TActionFiche;
{$ENDIF TRESO}
begin
  TobLig := TOB(FListe.Objects[0, FListe.Row]);

  {$IFDEF TRESO}
  {Affichage de TRECRITURE_TOM}
  if EstPointageSurTreso then begin
    {Recupération de la clef primaire}
    Lequel := TobLig.GetString(TRESO_NODOSSIER) + ';' + TobLig.GetString(TRESO_NUMTRANSAC) + ';' +
              {16/07/07 : FQ TRESO 10507 : pour la TOM, il faut le numéro de ligne de la Tréso, pas celui de la compta !!}
              TobLig.GetString('CLE_NUMEROPIECE') + ';' + TobLig.GetString('CLE_NUMLIGNE') + ';';
    AGLLanceFiche('TR', 'TRFICECRITURE', '', Lequel, na_Realise + ';');
  end

  {Affichage de la Tom simplifiée des écritures de compta}
  else begin
    F := TOB.Create('_LIGNE', nil, -1);
    try
      F.AddChampSupValeur('E_DATECOMPTABLE', TobLig.GetValue('CLE_DATECOMPTABLE'));
      F.AddChampSupValeur('E_EXERCICE'     , TobLig.GetValue('CLE_EXERCICE'));
      F.AddChampSupValeur('E_NUMEROPIECE'  , TobLig.GetValue('CLE_NUMEROPIECE'));
      F.AddChampSupValeur('E_NUMLIGNE'     , TobLig.GetValue('CLE_NUMLIGNE'));
      F.AddChampSupValeur('E_JOURNAL'      , TobLig.GetValue('CLE_JOURNAL'));
      F.AddChampSupValeur('E_GENERAL'      , FGeneral);
      F.AddChampSupValeur('NOMBASE'        , '');
      AffDetailEcr(F, True);
    finally
      FreeAndNil(F);
    end;
  end;

  {$ELSE}
  F := CreateTobParamSaisiePiece;

  if not Assigned(TobLig) then Action := taCreat
  else if TobLig.GetString('CLE_JOURNAL') = LIGNEBANCAIRE then Action := taCreat
  else begin
    Action := taModif;
    {20/07/07 : FQ 20601 : Nouvelle fonction d'appel de la saisie paramétrable pour préciser la ligne}
    F.SetString('E_JOURNAL', TobLig.GetString('CLE_JOURNAL'));
    F.SetInteger('E_NUMEROPIECE', TobLig.GetInteger('CLE_NUMEROPIECE'));
    F.SetDateTime('E_DATECOMPTABLE', TobLig.GetDateTime('CLE_DATECOMPTABLE'));
    F.SetInteger('E_NUMLIGNE', TobLig.GetInteger('CLE_NUMLIGNE'));
    F.SetString('SAISIEPIECECOMPTA', '-');
  end;

  if MajPointage and FBoDroitEcritures and not FPtgeNonBqe then begin
    {16/08/07 : FQ 21246 : Maintenant, si on est en modification, on ne peut plus boucler en saisie
                => on n'a plus accès à l'entête de pièce, ce qui fait que l'on peut tester le retour
                de la fiche. Par ailleurs on ne devrait jamais être dans le cas d'une action à taCreat}
    if SaisiePieceComptaTOB(F, Action) then
      RafraichitFiche(0);
  end
  else begin
    if Action = taModif then
      {20/09/07 : FQ 21318 : si on ne valide pas le pointage courant, on passe en lecture seule}
      SaisiePieceComptaTOB(F, {Action} taConsult)
    else
      PgiBox(TraduireMemoire('Veuillez vous positionner sur une écriture ou valider le pointage'));
  end;
  {$ENDIF TRESO}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.SupprEcriture(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  TF : TOB;
begin
  if FBoConsultation then Exit;

  TF := TOB(Fliste.Objects[0, FListe.Row]);
  if not Assigned(TF) then Exit;
  if TF.GetString('CLE_JOURNAL') = LIGNEBANCAIRE then begin
    PgiBox(TraduireMemoire('Veuillez vous positionner sur une écriture.'), Ecran.Caption);
    Exit;
  end
  else if TF.GetString('CLE_POINTE') <> '' then begin
    PGIError(TraduireMemoire('L''écriture est déjà pointée, elle ne peut être supprimée.'), Ecran.Caption);
    Exit;
  end
  {18/07/07 : FQ 21127 : On n'applique la suppression de pièce qu'en saisie normal}
  else if not (CtxTreso in V_PGI.PGIContexte) and 
          ExisteSQL('SELECT J_MODESAISIE FROM JOURNAL WHERE J_MODESAISIE IN ("BOR", "LIB") AND J_JOURNAL = "' + TF.GetString('CLE_JOURNAL') + '"') then begin
    PGIError(TraduireMemoire('Ce traitement ne peut être appliqué à une écriture de bordereau.'), Ecran.Caption);
    Exit;
  end;


  {$IFDEF TRESO}
  TRLanceFiche_Suppression('TR', 'TRSUPPRECRITURE', '', '', '');
  {$ELSE}
  DetruitSurCritere(TF.GetString('CLE_JOURNAL'    ), TF.GetString('CLE_EXERCICE'), '', 'N',
                    TF.GetString('CLE_NUMEROPIECE'), TF.GetString('CLE_DATECOMPTABLE'), '', False);
  {$ENDIF TRESO}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.ModifEntete(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  TF : TOB;
begin
  if FBoConsultation then Exit;

  TF := TOB(Fliste.Objects[0, FListe.Row]);
  if not Assigned(TF) then Exit;
  if TF.GetString('CLE_JOURNAL') = LIGNEBANCAIRE then begin
    PgiBox(TraduireMemoire('Veuillez vous positionner sur une écriture.'), Ecran.Caption);
    Exit;
  end
  else if TF.GetString('CLE_POINTE') <> '' then begin
    PGIError(TraduireMemoire('L''écriture est déjà pointée, elle ne peut être modifiée.'), Ecran.Caption);
    Exit;
  end
  {12/07/07 : FQ 21048 : On n'applique la modification d'entête de pièce qu'en saisie normal}
  else if ExisteSQL('SELECT J_MODESAISIE FROM JOURNAL WHERE J_MODESAISIE IN ("BOR", "LIB") AND J_JOURNAL = "' + TF.GetString('CLE_JOURNAL') + '"') then begin
    PGIError(TraduireMemoire('Ce traitement ne peut être appliqué à une écriture de bordereau.'), Ecran.Caption);
    Exit;
  end;

  {$IFNDEF TRESO}
  ModifEntPieSurCriteres(TF.GetString('CLE_JOURNAL'), TF.GetString('CLE_EXERCICE'), '',
                         TF.GetString('CLE_DATECOMPTABLE'), '');
  {$ENDIF TRESO}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.AccesSaisiePara(Sender : TObject);
{---------------------------------------------------------------------------------------}
{$IFNDEF TRESO}
var
  PC  : TPieceCompta;
  DT  : TDateTime;{09/10/07 : FQ 21055}
{$ENDIF TRESO}
begin
  {On s'assure que l'on a les droits, notamment pour les raccourcis clavier}
  if FBoConsultation or FPtgeNonBqe then Exit;
  if not FBoDroitEcritures then Exit;

  {$IFDEF TRESO}
  {En Tréso, on lance la saisie de Trésorerie}
  if CtxTreso in V_PGI.PGIContexte then begin
    GenereEcriture;
    Exit;
  end;
  {$ELSE}

  PC := TPieceCompta.CreerPiece;
  try
    PC.SetMultiEcheOff;
    DT := GetDateComptable;
    {09/10/07 : FQ 21055 : On s'assure que l'on peut créer une pièce sur la période}
    if PC.IsValidDateCreation(DT) then begin
      {20/09/07 : FQ 21333 : GetDateComptable : Fenêtre de choix d'une date comptable}
      PC.InitPiece(FJournal, Dt, 'OD', FDevise, VH^.EtablisDefaut, 'N');
      PC.InitSaisie;

      // SBO 13/08/2007 : cas spécifique au journaux bordereaux : il faut préalablement charger le bordereau et mettre en place les verrous
      if PC.ModeSaisie <> msPiece then
        begin
        PC.AttribNumeroDef ;                   // numero de bordereau suivant (à définir sinon chargement inopérant)
        PC.LoadFromSQL ;                       // chargement des lignes
        if not PC.LockFolio( True ) then     // mise en place verrou (avec message éventuel)
          Exit ;
        end ;
      try
        {Si la pièce a été validée depuis la saisie paramétrable}
        if ModifiePieceCompta(PC, taModif) then begin
          {La pièce est sur un autre journal ...}
          if (PC.GetEntete('E_JOURNAL') <> FJournal) then begin   // SBO 13/08/2007 : ne devrait plus arriver car l'utilisation de la fonction ModifiePieceCompta bloque l'entête de la pièce (ou du bordereau)
            {... On se contente de sauver la pièce}
            PC.Save;
          end
          {Sinon on va pointer les écritures}
          else begin
            {Enregistrement de la pièce et mise à jour du pointage}
            MajRefPointage(PC);
          end;
        end;
      finally
        // SBO 13/08/2007 : cas spécifique au journaux bordereau : il ôter les verrous après modif
        if PC.ModeSaisie <> msPiece then
          PC.UnLockFolio ;
      end;
    end;
  finally
    FreeAndNil(PC);
  end;
  Exit;
  {$ENDIF TRESO}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.CreerEcriture(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  TF  : TOB;
  {$IFNDEF TRESO}
  Obj : TObjEcritureSimplif;
  {$ENDIF TRESO}
begin
  if FBoConsultation or FPtgeNonBqe then Exit;

  {On s'assure que l'on a les droits, notamment pour les raccourcis clavier}
  if not FBoDroitEcritures then Exit;
  {FQ 20065 : erreur !! On cache le menu en pointage manuel
  if FPtgeManuel then Exit;}

  TF := TOB(Fliste.Objects[0, FListe.Row]);
  if (TF.GetString('CLE_JOURNAL') <> LIGNEBANCAIRE) then begin
    PGIError(TraduireMemoire('Veuillez vous positionner sur un mouvement bancaire.'), Ecran.Caption);
    Exit;
  end
  else if (TF.GetString('CLE_POINTE') <> '') then begin
    PGIError(TraduireMemoire('Le mouvement est déjà pointé.'), Ecran.Caption);
    Exit;
  end;

  {$IFDEF TRESO}
  {En Tréso, on lance la saisie de Trésorerie}
  if CtxTreso in V_PGI.PGIContexte then begin
    // prépointer l'écriture : trouver une solution !!
    GenereEcriture;
    Exit;
  end;

  {$ELSE}
  {Création de l'objet de gestion de la saisie simplifiée :
   19/07/07 : FQ 21125 : on l'active en pointage manuel => Nlle procédure AccesSaisiePara}
  Obj := TObjEcritureSimplif.Create;
  try
    {Initialisation des valeurs de la première ligne de la future pièce}
    Obj.General  := FGeneral;
    Obj.Etab     := VH^.EtablisDefaut;
    Obj.Devise   := FDevise;
    Obj.NatPiece := 'OD';
    Obj.QualifP  := 'N';
    Obj.Journal  := FJournal;
    Obj.DateCpt := TF.GetDateTime('CLE_DATECOMPTABLE');
    Obj.DateVal := TF.GetDateTime('CLE_DATEVALEUR');
    Obj.Libelle := TF.GetString('CLE_LIBELLE');
    Obj.RefInt  := TF.GetString('CLE_REFINTERNE');
    if TF.GetString('CLE_DEVISE') = V_PGI.DevisePivot then begin
      {31/07/07 : FQ 21125 : sur les mouvements manuels, on crée l'écriture dans le sens du mouvement}
      if TF.GetString('CLE_MANUEL') = 'X' then begin
        Obj.DebitD  := TF.GetDouble('CLE_DEBIT');
        Obj.CreditD := TF.GetDouble('CLE_CREDIT');
      end else begin
        Obj.CreditD := TF.GetDouble('CLE_DEBIT');
        Obj.DebitD  := TF.GetDouble('CLE_CREDIT');
      end;
    end else begin
      {31/07/07 : FQ 21125 : sur les mouvements manuels, on crée l'écriture dans le sens du mouvement}
      if TF.GetString('CLE_MANUEL') = 'X' then begin
        Obj.DebitD  := TF.GetDouble('CLE_DEBITDEV');
        Obj.CreditD := TF.GetDouble('CLE_CREDITDEV');
      end else begin
        Obj.CreditD := TF.GetDouble('CLE_DEBITDEV');
        Obj.DebitD  := TF.GetDouble('CLE_CREDITDEV');
      end;
    end;

    if (Obj.DebitD - Obj.CreditD) >= 0 then
      Obj.ModePaie := GetModePaie(TF.GetString('CLE_CIB'), 'ENC')
    else
      Obj.ModePaie := GetModePaie(TF.GetString('CLE_CIB'), 'DEC');

    {Lancement de la saisie simplifiée}
    if Obj.LanceSaisieSimplifie then begin
      {Récupération du PieceCompta, insertion en base et mise à jour du pointage}
      if not MajRefPointage(Obj.PieceCompta) then
        PgiError(TraduireMemoire('La mise à jour du pointage a échoué'), Ecran.Caption);
    end;
  finally
    FreeAndNil(Obj);
  end;
  {$ENDIF TRESO}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.AjouteMvt(var aTob : TOB);
{---------------------------------------------------------------------------------------}
var
  F : TOB;
  C : string;
begin //OK@@
  if FBoConsultation and not FBoPeutCreer then Exit;

  if (aTob.GetString('CLE_POINTE') = '0') or (aTob.GetString('CLE_POINTE') = '') then begin
    F := aTob.Detail[0];

    if Assigned(F) then begin
      if VH^.PointageJal then c := FJournal
                         else c := FGeneral;
      {On s'assure que le compte / Journal n'a pas été changé}
      if F.GetString('CLE_GENERAL') <> c then Exit;

      if not(EtatAffichage in [tv_Comptable, tv_NonPointe]) then begin
        FListe.InsertRow(FListe.Row + 1);
        F.ChangeParent(TobReleve, -1);
        RemplitLigne(FListe.Row + 1, F);
        {12/07/07 : FQ 21054 : On ne pointe pas l'éciture par défaut
        CocheDecoche(FListe.Row + 1, '0');}
      end;
    end;
  end;

  {Mise à jour des totaux}
  MajTotalisation;
end;

{Récupère le numéro de pointage maximum dans la Grille
{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.GetMaxNumPtge : string;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin //OK@@
  Result := '0';
  for n := 0 to TobReleve.Detail.Count -1 do
    if (ValeurI(TobReleve.Detail[n].GetString('CLE_POINTE')) > ValeurI(Result)) and
       (TobReleve.Detail[n].GetDateTime('CLE_DATEPOINTAGE') = FDtEE_DatePointage) then
      Result := TobReleve.Detail[n].GetString('CLE_POINTE');
  Result := IntToStr(ValeurI(Result) + 1);
end;

{Filtrage du contenu de la grille
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.ShowLignes(aTyp : TTypeVisu);
{---------------------------------------------------------------------------------------}

      {----------------------------------------------------------------}
      function _ATraiter(aTob : TOB) : Boolean;
      {----------------------------------------------------------------}
      begin
        if aTyp = tv_Tout then
          Result := True
        else if aTyp = tv_Bancaire then
          Result := aTob.GetString('CLE_JOURNAL') = LIGNEBANCAIRE
        else if aTyp = tv_Comptable then
          Result := aTob.GetString('CLE_JOURNAL') <> LIGNEBANCAIRE
        else if aTyp = tv_Pointe then
          Result := aTob.GetString('CLE_POINTE') <> ''
        else if aTyp = tv_NonPointe then
          Result := aTob.GetString('CLE_POINTE') = ''
        else if aTyp = tv_EnCours then
          Result := aTob.GetString('MODIFIE') = 'X'
        else if aTyp = tv_Session then
          Result := (aTob.GetString('CLE_REFPOINTAGE') = FStEE_RefPointage) or (aTob.GetString('CLE_REFPOINTAGE') = '')
        else
          Result := False;
      end;

var
  n : Integer;
  p : Integer;
  T : TOB;
begin //OK@@
  FListe.ColTypes[COL_POINTE] := 'R'; {#0}
  FListe.ColFormats[COL_POINTE] := '0';
  p := 1;

  if TobReleve.Detail.Count = 0 then Exit;

  TobReleve.Detail.Sort('CLE_REFPOINTAGE;CLE_POINTE;CLE_DATECOMPTABLE;CLE_DATEVALEUR;CLE_MONTANT;CLE_JOURNAL;');
  {La gestion du RowCount est un peu lourde, mais cela fonctionne}
  {1/ On vide la grille}
  FListe.RowCount := 1;
  {2/ On Insère une ligne vide qui servira de base pour le InsertRow}
  FListe.RowCount := 2;
  for n := 0 to TobReleve.Detail.Count - 1 do begin
    T := TobReleve.Detail[n];
    if _ATraiter(T) then begin
      Inc(p);
      FListe.InsertRow(p);
      RemplitLigne(p - 1, T);
    end;
  end;
  {3/ On supprime une éventuelle ligne vide en fin de grille}
  FListe.RowCount := p;
  {4/ Le nombre de lignes fixes doit être strictement inférieures au nombre de lignes}
  if p > 1 then FListe.FixedRows := 1;

  {Mise à jour des options}
  //SetControlEnabled('BPOINTAGEAUTO', aTyp = tv_Tout);
  {Gérer la variable dans les différents traitements : depointegroupe, pointage manuel ....}
  EtatAffichage := aTyp;
  {Mise à jour du libellé recapitulatif}
  AfficheEtatGrille;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.NextPriorByFieldOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
       if UpperCase(TMenuItem(Sender).Name) = 'POINTESUIVANT'     then Fliste.Row := RechercheLigne('CLE_REFPOINTAGE', FStEE_RefPointage)
  else if UpperCase(TMenuItem(Sender).Name) = 'POINTEPRE'         then Fliste.Row := RechercheLigne('CLE_REFPOINTAGE', FStEE_RefPointage, False)
  else if UpperCase(TMenuItem(Sender).Name) = 'NONPOINTESUIVANT'  then Fliste.Row := RechercheLigne('CLE_REFPOINTAGE', '')
  else if UpperCase(TMenuItem(Sender).Name) = 'NONPOINTEPRE'      then Fliste.Row := RechercheLigne('CLE_REFPOINTAGE', '', False)
  else if UpperCase(TMenuItem(Sender).Name) = 'LIBELLESUIVANT'    then Fliste.Row := RechercheLigne('CLE_LIBELLE', Fliste.Cells[COL_LIBELLE, Fliste.Row])
  else if UpperCase(TMenuItem(Sender).Name) = 'MOISSUIVANT'       then RechercheDate(True)
  else if UpperCase(TMenuItem(Sender).Name) = 'JOURSUIVANT'       then RechercheDate(False)
  {16/07/07 : FQ 21049 : Gestion des boutons sur le modèle du POPF11}
  else if UpperCase(TMenuItem(Sender).Name) = 'POINTESUIVANT1'    then Fliste.Row := RechercheLigne('CLE_REFPOINTAGE', FStEE_RefPointage)
  else if UpperCase(TMenuItem(Sender).Name) = 'POINTEPRE1'        then Fliste.Row := RechercheLigne('CLE_REFPOINTAGE', FStEE_RefPointage, False)
  else if UpperCase(TMenuItem(Sender).Name) = 'NONPOINTESUIVANT1' then Fliste.Row := RechercheLigne('CLE_REFPOINTAGE', '')
  else if UpperCase(TMenuItem(Sender).Name) = 'NONPOINTEPRE1'     then Fliste.Row := RechercheLigne('CLE_REFPOINTAGE', '', False)
  else if UpperCase(TMenuItem(Sender).Name) = 'LIBELLESUIVANT1'   then Fliste.Row := RechercheLigne('CLE_LIBELLE', Fliste.Cells[COL_LIBELLE, Fliste.Row])
  else if UpperCase(TMenuItem(Sender).Name) = 'MOISSUIVANT1'      then RechercheDate(True)
  else if UpperCase(TMenuItem(Sender).Name) = 'JOURSUIVANT1'      then RechercheDate(False)
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.OnClickSelectionRapide(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  SelectionRapide;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.OnClickSelectionJour(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  SelectionJour;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.OnClickSelectionMois(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  SelectionMois;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.AligneControleSurGrille(vCol : Integer; vNomControl : string; vAvecWidth, RelatifOk : Boolean);
{---------------------------------------------------------------------------------------}
var
  lRect  : TRect;
  lLeft  : Integer;
  lRight : Integer;
begin //OK@@

  {si AfficheDebitCredit, les colonnes Débit et Crédit sont visible, sinon, on affiche que la colonne Debit.
   Pour les montants bancaires, on positionne les deux zones sur le Libellé que l'on divise en deux}
  if (not AfficheDebitCredit and (vCol = COL_MNTC)) or (vCol = COL_REFINT) then begin
    if vCol = COL_REFINT then lRect  := FListe.CellRect(COL_LIBELLE, 0)
                         else lRect  := FListe.CellRect(COL_MNTD, 0);

    lLeft  := lRect.Left + Round((lRect.Right - lRect.Left) / 2) + 1;
    lRight := lRect.Right;
  end

  else begin
    lRect  := FListe.CellRect(vCol, 0);
    lLeft  := lRect.Left;

    {Dans ce cas on considère que Right est le milieu de la cellule}
    if (not AfficheDebitCredit and (vCol = COL_MNTD)) or (vCol = COL_LIBELLE) then
      lRight := lRect.Left + Round((lRect.Right - lRect.Left) / 2)
    else
      lRight := lRect.Right;
  end;

  if vAvecWidth then
    SetControlProperty(vNomControl, 'WIDTH', (lRight - lLeft) + 1);

  {Si l'on a affaire aux GroupBox ou les zones "reste à pointer", la position est "absolue" ...}
  if not RelatifOk then
    SetControlProperty(vNomControl, 'LEFT', lLeft + 1)

  {... Sinon, pour les THEdit, les positions sont "relatives" car leur parent est un GroupBox}
  else begin
    lLeft := lLeft - GetControl(vNomControl).Parent.Left;
    SetControlProperty(vNomControl, 'LEFT', lLeft + 1);
  end;
end;


{Alignement des Soldes avec les colonnes de la grille
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.OnResizeEcran(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe);
  {On commence par positionner les deux GoupBox}
  AligneControleSurGrille(COL_CIB , 'GBCOMPTA', False, False);
  AligneControleSurGrille(COL_DATE, 'GBBANQUE', False, False);
  {FQ 21043 : comme on cache la colonne pointé en pointage manuel, le GroupBox doit être un peu plus petit
   30/07/07 : j'annule : j'avais mal compris la FQ
  if FPtgeManuel then
    SetControlProperty('GBCOMPTA', 'WIDTH', Round((GetControl('GBBANQUE') as TGroupBox).Parent.Width / 2.53))
  else}
    SetControlProperty('GBCOMPTA', 'WIDTH', Round((GetControl('GBBANQUE') as TGroupBox).Parent.Width / 2.4));
  SetControlProperty('GBBANQUE', 'WIDTH', Round((GetControl('GBBANQUE') as TGroupBox).Parent.Width / 2.4));
  {On positionne le groupBox sur le milieu de la colonne DATE}
  SetControlProperty('GBBANQUE', 'LEFT' , FListe.CellRect(COL_DATE, 0).Left +
                                           Round((FListe.CellRect(COL_DATE, 0).Right - FListe.CellRect(COL_DATE, 0).Left) / 2));

  {GroupBox sur les écritures}
  AligneControleSurGrille(COL_MNTD, 'TDMVTAFFICHE'    , True, True);
  AligneControleSurGrille(COL_MNTD, 'TDMVTPOINTE'     , True, True);
  AligneControleSurGrille(COL_MNTD, 'TDSOLDECOMPTABLE', True, True);

  AligneControleSurGrille(COL_MNTC, 'TCMVTAFFICHE'    , True, True);
  AligneControleSurGrille(COL_MNTC, 'TCMVTPOINTE'     , True, True);
  AligneControleSurGrille(COL_MNTC, 'TCSOLDECOMPTABLE', True, True);

  if FPtgeManuel then begin
    SetControlProperty('TCMVTAFFICHE'    , 'WIDTH', (GetControl('TCMVTAFFICHE'    ) as THNumEdit).Width - 2);
    SetControlProperty('TCMVTPOINTE'     , 'WIDTH', (GetControl('TCMVTPOINTE'     ) as THNumEdit).Width - 2);
    SetControlProperty('TCSOLDECOMPTABLE', 'WIDTH', (GetControl('TCSOLDECOMPTABLE') as THNumEdit).Width - 2);
  end;

  AligneControleSurGrille(COL_REFPOIN, 'TMVTAFFICHEE'   , False, True);
  AligneControleSurGrille(COL_REFPOIN, 'TMVTPOINTEE'    , False, True);
  AligneControleSurGrille(COL_REFPOIN, 'TSOLDECOMPTABLE', False, True);

  {Reste à pointer}
  AligneControleSurGrille(COL_MNTD   , 'TDAVANCEMENT', True , False);
  AligneControleSurGrille(COL_MNTC   , 'TCAVANCEMENT', True , False);
  AligneControleSurGrille(COL_REFPOIN, 'TAVANCEMENT' , False, False);

  {GroupBox sur les mouvements bancaires}
  AligneControleSurGrille(COL_LIBELLE, 'TDLIGAFFICHE'   , True, True);
  AligneControleSurGrille(COL_LIBELLE, 'TDLIGPOINTE'    , True, True);
  AligneControleSurGrille(COL_LIBELLE, 'TDSOLDEBANCAIRE', True, True);

  AligneControleSurGrille(COL_REFINT, 'TCLIGAFFICHE'   , True, True);
  AligneControleSurGrille(COL_REFINT, 'TCLIGPOINTE'    , True, True);
  AligneControleSurGrille(COL_REFINT, 'TCSOLDEBANCAIRE', True, True);

  AligneControleSurGrille(COL_PIECE, 'TMVTAFFICHEB'   , False, True);
  AligneControleSurGrille(COL_PIECE, 'TMVTPOINTEB'    , False, True);
  AligneControleSurGrille(COL_PIECE, 'TSOLDEBANCAIRE' , False, True);
  AligneControleSurGrille(COL_PIECE, 'TRECAPITULATIF' , False, False);

  {24/07/07 : #ETAT# : ajout d'images pour le suivi de cohérence de la session de pointage}
  SetControlProperty('IROUGE', 'LEFT', (GetControl('TAVANCEMENT') as THLabel).Left - 22);
  SetControlProperty('IVERT' , 'LEFT', (GetControl('TAVANCEMENT') as THLabel).Left - 22);

  Ecran.Refresh;
  Application.ProcessMessages;
end;

{Controle afin d'empecher de saisir une date supérieure à la date de la référence de pointage
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.OnExitE_DateComptable_(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  if not IsValidDate(E_DateComptable_.Text) then
    Exit;

  if StrToDate(E_DateComptable_.Text) > FDtEE_DatePointage then begin
    PgiInfo(TraduireMemoire('La date saisie doit être inférieure au') + ' ' + DateToStr(FDtEE_DatePointage + 1), Ecran.Caption);
    SetFocusControl('E_DATECOMPTABLE_');
    SetControlProperty('E_DATECOMPTABLE_', 'TEXT', FDtEE_DatePointage);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.OnClickBValideEtRappro(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if FBoConsultation then Exit;
  FEtatRapproOk := True;
  OnUpdate;
  Ecran.Close;
end;

{Valide le pointage des écritures
Pas d' héritage de la fiche ancetre car dans celle ci on
faisait le même traitement que le DoubleClique sur la grille
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.OnClickBValider(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  {On s'assure que le pointage est Ok}
  TFVierge(Ecran).BValiderClick(Sender);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.BFermeOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if not FBoConsultation then begin
    if MessageValidation(True) then
      MajPointage
    else
      AnnulePointage;
  end;
  TFVierge(Ecran).BFermeClick(Sender);
end;

{Pointe l'écriture avec la combinaison (Ctrl +) Click => même comportement que la touche espace
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.OnMouseDownFListe(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
{---------------------------------------------------------------------------------------}
var
  Key : Word;
begin //OK@@
  if (FBoConsultation) then Exit;
  Key := VK_SPACE;
  if ssLeft in Shift then begin
    if ssCtrl in Shift then OnKeyDownFListe(Sender, Key, [ssCtrl])
    else if ssShift in Shift then OnKeyDownFListe(Sender, Key, []);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.PopupFiltreOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
       if UpperCase((Sender as TMenuItem).Name) = 'MVTBANCAIRE'  then ShowLignes(tv_Bancaire)
  else if UpperCase((Sender as TMenuItem).Name) = 'MVTCOMPTABLE' then ShowLignes(tv_Comptable)
  else if UpperCase((Sender as TMenuItem).Name) = 'MVTPOINTE'    then ShowLignes(tv_Pointe)
  else if UpperCase((Sender as TMenuItem).Name) = 'MVTNONPOINTE' then ShowLignes(tv_NonPointe)
  else if UpperCase((Sender as TMenuItem).Name) = 'MVTSESSION'   then ShowLignes(tv_Session)
  else if UpperCase((Sender as TMenuItem).Name) = 'MVTENCOURS'   then ShowLignes(tv_EnCours)
  else if UpperCase((Sender as TMenuItem).Name) = 'MVTTOUT'      then ShowLignes(tv_Tout)
  {16/07/07 : FQ 21049 : Gestion des boutons sur le modèle du POPF11}
  else if UpperCase((Sender as TMenuItem).Name) = 'MVTBANCAIRE1'  then ShowLignes(tv_Bancaire)
  else if UpperCase((Sender as TMenuItem).Name) = 'MVTCOMPTABLE1' then ShowLignes(tv_Comptable)
  else if UpperCase((Sender as TMenuItem).Name) = 'MVTPOINTE1'    then ShowLignes(tv_Pointe)
  else if UpperCase((Sender as TMenuItem).Name) = 'MVTNONPOINTE1' then ShowLignes(tv_NonPointe)
  else if UpperCase((Sender as TMenuItem).Name) = 'MVTSESSION1'   then ShowLignes(tv_Session)
  else if UpperCase((Sender as TMenuItem).Name) = 'MVTENCOURS1'   then ShowLignes(tv_EnCours)
  else if UpperCase((Sender as TMenuItem).Name) = 'MVTTOUT1'      then ShowLignes(tv_Tout);
end;

{20/07/07 : FQ 21101 : Gestion des raccourcis d'accès aux mouvements et aux écritures en fonction de la ligne courante}
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.PopupOnPopup(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  TF : TOB;
begin //OK@@
  TF := TOB(Fliste.Objects[0, FListe.Row]);
  {16/08/07 : FQ 13017 : les menus ne sont visibles que s'il y a des pièces sur un autre journal}
  TMenuItem(GetControl('AUTREJAL' )).Visible := (VH^.PointageJal and (GetControlText('TCONTREPARTIE') <> ''));
  TMenuItem(GetControl('AUTREJAL1')).Visible := (VH^.PointageJal and (GetControlText('TCONTREPARTIE') <> ''));

  if not Assigned(TF) then begin
    TMenuItem(GetControl('ACCEDERECR' )).Visible := False;
    TMenuItem(GetControl('ACCEDERMVT' )).Visible := False;
    TMenuItem(GetControl('ACCEDERECR1')).Visible := False;
    TMenuItem(GetControl('ACCEDERMVT1')).Visible := False;
  end
  else if TF.GetString('CLE_JOURNAL') = LIGNEBANCAIRE then begin
    TMenuItem(GetControl('ACCEDERECR'  )).Visible := False;
    TMenuItem(GetControl('ACCEDERMVT'  )).Visible := True;
    TMenuItem(GetControl('ACCEDERECR1' )).Visible := False;
    TMenuItem(GetControl('ACCEDERMVT1' )).Visible := True;
    {30/07/07 : FQ 21192 : on n'affiche la saisie simplifiée que sur les mouvements bancaires}
    TMenuItem(GetControl('SAISIESIMPL' )).Visible := True;
    TMenuItem(GetControl('SAISIESIMPL1')).Visible := True;
  end
  else begin
    TMenuItem(GetControl('ACCEDERECR'  )).Visible := True;
    TMenuItem(GetControl('ACCEDERMVT'  )).Visible := False;
    TMenuItem(GetControl('ACCEDERECR1' )).Visible := True;
    TMenuItem(GetControl('ACCEDERMVT1' )).Visible := False;
    {30/07/07 : FQ 21192 : on n'affiche la saisie simplifiée que sur les mouvements bancaires}
    TMenuItem(GetControl('SAISIESIMPL' )).Visible := False;
    TMenuItem(GetControl('SAISIESIMPL1')).Visible := False;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
{---------------------------------------------------------------------------------------}
begin //OK@@
  inherited;

  case Key of
    VK_F5     : if (Shift = [ssShift]) then ContrepartieDblClick(Sender); {16/08/07 : FQ 13017}

    VK_F6     :      if (Shift = [ssCtrl] ) then LancerRegle(GetControl('MULACC'))
                else if (Shift = [ssShift]) then TMenuItem(GetControl('RESTEAPOINTE')).Click {#ETAT# : le conserve-t-on ?}
                else if (Shift = [ssAlt]  ) then TMenuItem(GetControl('ACCEDERLIBCIB')).Click {18/09/07 : FQ 21393}
                                            else LancerCib  (GetControl('FICCIB'));

    VK_F7     : if (Shift = [ssShift]) then TMenuItem(GetControl('COHERENCE')).Click {#ETAT#}
                                       else BPointageAutoOnClick(GetControl('BPOINTAGEAUTO')); {18/07/07 : FQ 21130}

    VK_F9     : BChercheOnClick(GetControl('BCHERCHE'));

    VK_F10    :      if Shift = [ssAlt] then OnClickBValideEtRappro(GetControl('VALIDERETRAPPRO'))
                else if Shift = []      then TFVierge(Ecran).BValider.Click;

    VK_F11    : PopF11.Popup(Mouse.CursorPos.x, Mouse.CursorPos.y);

    {12/07/07 : FQ 21046 : Basculement entre la grille et le pagecontrol}
    VK_F12    : GereFocusF12;

    VK_INSERT :      if (Shift = [ssCtrl]) then CreerEexBqLig(GetControl('INSMVTBQE'))
                else if (Shift = [ssAlt])  then CreerEcriture(GetControl('SAISIESIMPL'))
                                           else AccesSaisiePara(GetControl('SAISIEPIECE'));

    {A : Tout prépointer}
    Ord('A') : if (Shift = [ssCtrl]) then  TMenuItem(GetControl('TOUTSELECT')).Click;

    {Annule tout le pointage}
    Ord('D') : if Shift = [ssCtrl] then begin
                 RAZPointage(GetControl('DEPOINTE'));
                 FBoAllSelected := False;
               end;

    {JP 17/07/07 : FQ 21110 : Branchement du Ctrl + F}
    Ord('F') : if (Shift = [ssCtrl]) then BRechercheOnClick(GetControl('BRECHERCHER'));

    {G : Réduire / Agrandir l'entête de la fiche}
    Ord('G') :      if Shift = [ssAlt  ] then RedAgrandOnClick(GetControl('BAGRANDIR'))
               else if Shift = [ssShift] then RedAgrandOnClick(GetControl('BREDUIRE'));

    {Ctrl + Z : Annule le pointage en cours}
    Ord('Z') : if Shift = [ssCtrl] then ReinitialisePointage;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.OnKeyDownFListe(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin //OK@@
  inherited;

  case key of
    {16/08/07 : FQ 13017 : gestion du maj dans EcranKeyDown}
    VK_F5     : if not (Shift = [ssShift]) then begin
                   {20/07/07 : FQ 21101 : plus qu'un raccourci pour le FListeOnDblClick}
                   if (ssCtrl in Shift) then LancerMulEexBqLig(FListe)
                                        else FListeOnDblClick(FListe);
                end;

    VK_SPACE  : if (Shift = [ssCtrl]) then DepointeGroupe(FListe.Cells[COL_POINTE, FListe.Row])
                                      else PointageManuel(FListe.Row);

    VK_DELETE : if (Shift = [ssCtrl]) then TMenuItem(GetControl('SUPPRPIECE')).Click;

    Ord('B') : begin
                      if (Shift = [ssShift]) then ShowLignes(tv_Bancaire)
                 else if (Shift = [ssCtrl ]) then ShowLignes(tv_Pointe);
               end;

    {C  : Ligne précédente non pointée : 26/09/07 : FQ 21111}
    Ord('C') : TMenuItem(GetControl('NONPOINTEPRE')).Click;


    Ord('E') : begin {E}
                      if (Shift = [ssShift]) then ShowLignes(tv_Comptable)
                 else if (Shift = [ssCtrl ]) then ShowLignes(tv_NonPointe);
               end;

    {J : Premier mouvement du jour suivant}
    Ord('J') : if (Shift = [ssCtrl]) then SelectionJour
                                     else TMenuItem(GetControl('JOURSUIVANT')).Click;

    {L : Mouvement suivant ayant le même libellé}
    Ord('L') : TMenuItem(GetControl('LIBELLESUIVANT')).Click;

    {M : Premier mouvement du mois suivant}
    Ord('M') :      if (Shift = [ssCtrl] ) then SelectionMois
               else if (Shift = [ssShift]) then TMenuItem(GetControl('MODIFENTETE')).Click
                                           else TMenuItem(GetControl('MOISSUIVANT')).Click;

    {N : Pointe toutes les écritures depuis la ligne courante}
    Ord('N') : if Shift = [ssCtrl] then SelectionRapide;

    {P : Mouvement pointé précédent | 18/07/07 : FQ 21111}
    Ord('P') : if (Shift = [ssShift]) then TMenuItem(GetControl('POINTEPRE')).Click;

    {S : Mouvement pointé suivant | 18/07/07 : FQ 21111}
    Ord('S') : if (Shift = [ssShift]) then TMenuItem(GetControl('POINTESUIVANT')).Click
                                      else ShowLignes(tv_Session);

    {T : Affichage de tous les mouvements | 18/07/07 : FQ 21111}
    Ord('T') : if (Shift = [ssCtrl ]) then ShowLignes(tv_EnCours)
                                      else ShowLignes(tv_Tout);

    {V  : Ligne suivante non pointée : 26/09/07 : FQ 21111}
    Ord('V') : TMenuItem(GetControl('NONPOINTESUIVANT')).Click;
  end;

end;

{18/07/07 : FQ 21106 : Dessein des indicateurs de la colonne DBIndicator
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.DrawDBIndicator(Tip: Byte; ACol, ARow : Integer);
{---------------------------------------------------------------------------------------}
  {dbIndCurrent 0,dbIndSelected 1,dbIndNone 2}
const
  W = 10;
  H = 14;
var
  CT, CL, CH : Integer;
  R          : TRect;
begin
  R := FListe.CellRect(ACol, ARow);

  with FListe.Canvas, R do begin
    CT := Top;
    CL := Left + 3;
    CH := (Bottom - Top);
    Brush.Color := clBlack;
    Pen.Mode    := pmCopy ;
    Pen.Style   := psSolid ;
    Pen.Width   := 1 ;
    Pen.Color   := clBlack;
    case Tip of
      0 : begin {o> dbIndCurrent}
            MoveTo(CL, CT + (CH div 2) - (W div 2));
            LineTo(CL + (W div 2), CT + (CH div 2));
            LineTo(CL, CT + H);
            LineTo(CL, CT + H - 1);
            {simulation epaisseur}
            MoveTo(CL, CT + (CH div 2) - (W div 2) + 1);
            LineTo(CL + (W div 2) - 1, CT + (CH div 2));
            LineTo(CL, CT + H - 1);
            LineTo(CL, CT + H - 1 - 1);
            {rond}
            Rectangle(CL, CT + (CH div 2) - 1, CL + 2, CT + (CH div 2) + 2);
          end;

      1 : begin {o dbIndSelected}
            {rond}
            Ellipse(CL, CT + (CH div 2) - 2, CL + (W div 2), CT + (CH div 2) + 2);
          end;

      2 : begin {> dbIndNone}
            {Triangle fermé}
            Polygon([Point(CL, CT + (CH div 2) - (W div 2)),
                     Point(CL + (W div 2), CT + (CH div 2)),
                     Point(CL, CT + H),
                     Point(CL, CT + (CH div 2) - (W div 2))]);
          end;
    end;
  end;
end;

{18/07/07 : FQ 21106 : Dessein d'un point si la ligne est pointée
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.PostDrawCell(ACol, ARow : Integer; Canvas : TCanvas; AState : TGridDrawState);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  if ARow = 0 then Exit;
  if ACol = 0 then begin
    T := TOB(FListe.Objects[0, ARow]);
    if not Assigned(T) then Exit;
    if T.GetString('CLE_POINTE') <> '' then DrawDBIndicator(1, Acol, ARow);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.GetCellCanvasFListe(ACol, ARow : Integer; Canvas : TCanvas; AState : TGridDrawState);
{---------------------------------------------------------------------------------------}
var
  lOldPen   : TPen;
  aColor    : TColor;
  T         : TOB;
begin //OK@@
  inherited;

  if (ARow = 0) or (ACol = 0) then Exit;

  aColor := clWindow;
  lOldPen := TPen.Create;

  {sauvegarde des valeurs courantes}
  lOldPen.Assign(FListe.Canvas.Pen);
  Canvas.Font.Color := clBlack;
  {18/07/07 : FQ 21103 : Dans certain cas, on était en Italic par défaut !?}
  Canvas.Font.Style  := [];

  try
    T := TOB(FListe.Objects[0, ARow]);
    {Si Mouvement bancaire ...}
    if Assigned(T) and (T.GetString('CLE_JOURNAL') = LIGNEBANCAIRE) then begin
      if V_PGI.NumAltCol = 0 then aColor := $00D6F3F2
                             else aColor := AltColors[V_PGI.NumAltCol];

      {On fonctionne dans le sens bancaire : encaissement = vert, décaissement = rouge}
           if ACol = COL_MNTD then Canvas.Font.Color := clRed
      else if ACol = COL_MNTC then Canvas.Font.Color := clGreen;
    end
    {... Sinon, écriture comptable}
    else begin
      {Si pointage sur journal, les écritures sont dans le même sens que les mouvements}
      if VH^.PointageJal or EstPointageSurTreso then begin
             if ACol = COL_MNTD then Canvas.Font.Color := clRed
        else if ACol = COL_MNTC then Canvas.Font.Color := clGreen;
      end

      {On fonctionne dans le sens bancaire : encaissement = vert, décaissement = rouge}
      else begin
             if ACol = COL_MNTD then Canvas.Font.Color := clGreen
        else if ACol = COL_MNTC then Canvas.Font.Color := clRed
      end;
    end;

    {Si en affichage par montant, les montants sont signés dans le sens bancaire}
    if not AfficheDebitCredit and (ACol = COL_MNTD) then begin
           if Valeur(FListe.Cells[ACol, ARow]) > 0 then Canvas.Font.Color := clGreen
      else if Valeur(FListe.Cells[ACol, ARow]) < 0 then Canvas.Font.Color := clRed
                                                   else Canvas.Font.Color := clBlack;
    end;

    if Assigned(T) and (T.GetString('CLE_POINTE') <> '') then
      Canvas.Font.Style := [fsItalic];

    {Ligne en cours}
    if (gdSelected in AState) then begin
      Canvas.Font.Style  := [fsBold];
      Canvas.Font.Color  := clBlack;
      Canvas.Brush.Color := $00FFF2F4;
    end
    else begin
      Canvas.Brush.Color := aColor;
      Canvas.Font.Style  := Canvas.Font.Style - [fsBold];
    end;

  finally
    {Réaffectation des valeurs du canevas}
    FListe.Canvas.Pen.Assign(lOldPen);
    if assigned(lOldPen) then FreeAndNil(lOldPen);
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.RecupAutreWhere : string;
{---------------------------------------------------------------------------------------}
var
  lSt : string;
begin //OK@@
  Result := RecupWhereCritere((GetControl('PAGECONTROL') as TPageControl));
  {Pointage sur TRECRITURE}
  if EstPointageSurTreso then begin
    Result := FindEtReplace(Result, ' E_', ' TE_', True);
    Result := FindEtReplace(Result, ',E_', ',TE_', True);
    Result := FindEtReplace(Result, '(E_', '(TE_', True);
    lSt := ' AND TE_GENERAL = "' + FGeneral + '"';

    lSt := lSt + ' AND ((TE_DATERAPPRO = "' + UsDateTime(iDate1900) + '") OR ' +
      '(TE_DATERAPPRO = "' + UsDateTime(FDtEE_DatePointage) + '" ) OR ' +
      '(TE_DATECOMPTABLE <= "' + UsDateTime(FDtEE_DatePointage) + '" AND ' +
      'TE_DATERAPPRO  > "' + UsDateTime(FDtEE_DatePointage) + '"))';

    if FDevise <> V_PGI.DevisePivot then
      lSt := lSt + ' AND TE_DEVISE = "' + FDevise + '"';

    lSt := lSt + ' AND (TE_MONTANT <> 0) AND TE_NATURE = "' + na_Realise + '"';
  end

  {Pointage sur ECRITURE}
  else begin
    if not VH^.PointageJal then
      lSt := ' AND E_GENERAL = "' + FGeneral + '"'
    else
      lSt := ' AND E_GENERAL <> "' + FGeneral + '" AND ' +
        ' E_JOURNAL = "' + FJournal + '"';

    lSt := lSt + ' AND E_QUALIFPIECE = "N"';

    // SBO 28/08/2007 : Ouverture sur RefPointage à NULL pour Oracle
    lSt := lSt + ' AND (((E_REFPOINTAGE = "" OR E_REFPOINTAGE IS NULL) AND E_DATEPOINTAGE = "' +
      UsDateTime(iDate1900) + '") OR ' +
      '(E_REFPOINTAGE = "' + FStEE_RefPointage + '" AND ' +
      ' E_DATEPOINTAGE = "' + UsDateTime(FDtEE_DatePointage) + '" ) OR ' +
      '(E_DATECOMPTABLE <= "' + UsDateTime(FDtEE_DatePointage) + '" AND ' +
      'E_DATEPOINTAGE  > "' + UsDateTime(FDtEE_DatePointage) + '"))';

    // FB 12782 - GCO - 13/10/2003
    if FDevise <> V_PGI.DevisePivot then
      lSt := lSt + ' AND E_DEVISE = "' + FDevise + '"';

    lSt := lSt + ' AND (E_ECRANOUVEAU = "N" OR E_ECRANOUVEAU = "H") AND E_TRESOLETTRE = "-"';
    lSt := lSt + ' AND ((E_DEBIT <> 0) OR (E_CREDIT <> 0)) AND E_CREERPAR <> "DET"';

      // Gestion du VH^.ExoV8
    if VH^.ExoV8.Code <> '' then
      lSt := lSt + ' AND E_DATECOMPTABLE >= "' + UsDateTime(VH^.ExoV8.Deb) + '" ';
  end;
  Result := Result + ' ' + lSt;
end;

{Renvoie la ligne suivante / précédente dont le champ Chp a la valeur Value}
{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.RechercheLigne(Chp, Value : string; Suivant : Boolean = True) : Integer;
{---------------------------------------------------------------------------------------}
var
  F : TOB;

    {------------------------------------------------------}
    function _FindNextPrev(Ind : Integer) : Boolean;
    {------------------------------------------------------}
    begin
      Result := False;
      F := TOB(Fliste.Objects[0, Ind]);
      if not Assigned(F) then Exit;

      if F.GetString(Chp) = Value then Result := True;
    end;

    {------------------------------------------------------}
    function _EnAvant(Dep : Integer) : Integer;
    {------------------------------------------------------}
    var
      n : Integer;
    begin
      Result := Dep;
      for n := Dep + 1 to FListe.RowCount - 1 do
        if _FindNextPrev(n) then begin
          Result := n;
          Break;
        end;
    end;

    {------------------------------------------------------}
    function _EnArriere(Dep : Integer) : Integer;
    {------------------------------------------------------}
    var
      n : Integer;
    begin
      Result := Dep;
      for n := Dep - 1 downto 1 do
        if _FindNextPrev(n) then begin
          Result := n;
          Break;
        end;
    end;

var
  lRes : Integer;
begin //OK@@
  Result := FListe.Row;

  if Suivant then begin
    {17/07/07 : FQ 21107 : On part de la ligne courante ...}
    lRes := _EnAvant(FListe.Row);
    {... si on ne trouve rien et qu'on n'est pas parti du début, on repart du début}
    if (FListe.Row > 1) and (lRes = FListe.Row) then
      lRes := _EnAvant(0);
    if (lRes <> FListe.RowCount) and (lRes <> 0) then 
      Result := lRes
  end

  else begin
    {17/07/07 : FQ 21107 : On part de la ligne courante ...}
    lRes := _EnArriere(FListe.Row);
    {... si on ne trouve rien et qu'on n'est pas parti de la fin, on repart de la fin}
    if (FListe.Row < FListe.RowCount - 1) and (lRes = FListe.Row) then
      lRes := _EnArriere(FListe.RowCount);
    if (lRes <> FListe.RowCount) and (lRes <> FListe.Row) then
      Result := lRes
  end;
end;

{Renvoie la ligne suivante / précédente dont la date / mois est le suivant par rapport à la ligne en cours}
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.RechercheDate(MoisOk : Boolean);
{---------------------------------------------------------------------------------------}
var
  sDate   : string;
  aDate   : string;
  a, m, j : Word;
  y, o, d : Word;
  n       : Integer;
begin //OK@@
  if FBoConsultation then Exit;

  sDate := FListe.Cells[COL_DATE, FListe.Row];

  if not IsValidDate(sDate) then Exit;

  DecodeDate(StrToDate(sDate), a, m, j);

  for n := FListe.Row + 1 to FListe.RowCount - 1 do begin
    aDate := FListe.Cells[COL_DATE, n];
    if not IsValidDate(aDate) then Continue;
    DecodeDate(StrToDate(aDate), y, o, d);
    if (MoisOk and ((o > m) or ((o <= m) and (y > a)))) or
       (not MoisOk and (StrToDate(aDate) > StrToDate(sDate))) then begin
      FListe.Row := n;
      Break;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.SelectionMois;
{---------------------------------------------------------------------------------------}
var
  sDate   : string;
  a, m, j : Word;
  m1      : Word;
  n       : Integer;
begin //OK@@
  if FBoConsultation then Exit;
  if not FPtgeManuel then Exit;

  sDate := FListe.Cells[COL_DATE, FListe.Row];

  if not IsValidDate(sDate) then Exit;

  DecodeDate(StrToDate(sDate), a, m, j);

  for n := 1 to FListe.RowCount - 1 do begin
    sDate := FListe.Cells[COL_DATE, n];
    if not IsValidDate(sDate) then Continue;
    DecodeDate(StrToDate(sDate), a, m1, j);
    if m1 = m then CocheDecoche(n, '0');
  end;

  {Mise à jour des totaux}
  MajTotalisation;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.SelectionJour;
{---------------------------------------------------------------------------------------}
var
  sDate : string;
  aDate : string;
  n     : Integer;
begin //OK@@
  if FBoConsultation then Exit;
  if not FPtgeManuel then Exit;

  sDate := FListe.Cells[COL_DATE, FListe.Row];

  if not IsValidDate(sDate) then Exit;

  for n := 1 to FListe.RowCount - 1 do begin
    aDate := FListe.Cells[COL_DATE, n];
    if not IsValidDate(sDate) then Continue;
    if aDate = sDate then CocheDecoche(n, '0');
  end;

  {Mise à jour des totaux}
  MajTotalisation;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.OnClickBSelectAll(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin //OK@@
  inherited;
  if FBoConsultation then Exit;
  if not FPtgeManuel then Exit;

  for n := 1 to FListe.RowCount - 1 do
    if FBoAllSelected then CocheDecoche(n, '')
                      else CocheDecoche(n, '0');

  FBoAllSelected := not FBoAllSelected;

  {Mise à jour des totaux}
  MajTotalisation;
end;

{11/05/07 : FQ 20284 : Ajout de l'export
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.BExportOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  ASaveDialog   : TSaveDialog;
  lStHint: string;
begin
  if not ExJaiLeDroitConcept(ccExportListe, True) then Exit;
  ASaveDialog := TSaveDialog.Create(Ecran);
  try
    ASaveDialog.Filter      := 'Fichier Texte (*.txt)|*.txt|Fichier Excel (*.xls)|*.xls|Fichier Ascii (*.asc)|*.asc|Fichier Lotus (*.wks)|*.wks|Fichier HTML (*.html)|*.html|Fichier XML (*.xml)|*.xml';
    ASaveDialog.DefaultExt  := 'XLS';
    ASaveDialog.FilterIndex := 1;
    ASaveDialog.Options     := ASaveDialog.Options + [ofOverwritePrompt, ofPathMustExist, ofNoReadonlyReturn, ofNoLongNames] - [ofEnableSizing];

    if ASaveDialog.Execute then begin
      {5 = .HTML}
      if ASaveDialog.FilterIndex = 5 then begin
        lStHint := FListe.Hint;
        FListe.Hint := Ecran.Caption;
        ExportGrid(FListe, nil, ASaveDialog.FileName, ASaveDialog.FilterIndex, True);
        FListe.Hint := lStHint;
      end
      else
        ExportGrid(FListe, nil, ASaveDialog.FileName, ASaveDialog.FilterIndex, True);
    end;
  finally
    FreeAndNil(ASaveDialog);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.SelectionRapide;
{---------------------------------------------------------------------------------------}
var
  i : Integer;
begin //OK@@
  if FBoConsultation then Exit;
  if not FPtgeManuel then Exit;

  for i := FListe.Row to FListe.RowCount - 1 do
    CocheDecoche(i, '0');

  {Mise à jour des totaux}
  MajTotalisation;
end;

{Verifie la présence d'écritures dans une devise differente de celle du compte de BANQUE(RIB)
 ou de la Devise PIVOT uniquement dans le cas où la devise utilisée est différente de la
 monnaie de tenue du DOSSIER.
 Renvoie True s'il y a plusieurs devises
{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.ControleDevise : Boolean;
{---------------------------------------------------------------------------------------}
begin //OK@@
  Result := False;
  {28/11/07 : On ne lance le traitement que si demandé ou ancienne version}
  if (not Assigned(CkLibExercice)) or CkLibExercice.Checked then
    try
      Result := ExisteSql('SELECT E_GENERAL FROM ECRITURE WHERE ' +
        'E_QUALIFPIECE = "N" AND ' + {28/11/07 : Pour optimiser}
        'E_GENERAL = "' + FGeneral + '" AND ' +
        'E_DEVISE <> "' + FDevise + '" AND ' +
        'E_DATECOMPTABLE <= "' + UsDateTime(FDtEE_DatePointage) + '"');
    except
      on E: Exception do PgiError(TraduireMemoire('Erreur SQL : ') + E.Message, Ecran.Caption);
    end;
end;

{Vérification effectue uniquement en Pointage sur journal
{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.ControleContrepartie: Boolean;
{---------------------------------------------------------------------------------------}
begin //OK@@
  Result := False;
  try
    Result := ExisteSql('SELECT E_GENERAL FROM ECRITURE WHERE ' +
      'E_GENERAL = "' + FGeneral + '" AND ' +
      'E_JOURNAL <> "' + FJournal + '" AND ' +
      'E_DATECOMPTABLE <="' + UsDateTime(FDtEE_DatePointage) + '"');
  except
    on E: Exception do PgiError(TraduireMemoire('Erreur SQL : ') + E.Message, Ecran.Caption);
  end;
end;

{16/08/07 : FQ 13017 : en pointage sur journal, affiche les pièces d'un autre journal
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.GetPieceAutreJournal;
{---------------------------------------------------------------------------------------}
var
  SQL : string;
  T   : TOB;
  O   : TObjEtats;
begin
  if VH^.PointageJal and (GetControlText('TCONTREPARTIE') <> '') then begin
    T := TOB.Create('MOMO', nil, -1);
    try
      SQL := 'SELECT J_LIBELLE, E_GENERAL, E_DATECOMPTABLE, E_NUMEROPIECE, E_DEBITDEV, E_CREDITDEV, E_DEVISE ' +
             'FROM ECRITURE LEFT JOIN JOURNAL ON J_JOURNAL = E_JOURNAL WHERE E_GENERAL = "' + FGeneral + '" AND ' +
             'E_JOURNAL <> "' + FJournal + '" AND E_DATECOMPTABLE <="' + UsDateTime(FDtEE_DatePointage) + '"';
      T.LoadDetailFromSQL(SQL);
      {Un relevé doit contenir au moins deux lignes : un 01 et un 07 }
      if T.Detail.Count > 0 then begin
        {Paramétrage de l'état}
        O := TObjEtats.Create(TraduireMemoire('Pièces sur un journal différent de ') + FJournal, T);
        O.MajTitre(0, TraduireMemoire('Journal'));
        O.MajTitre(1, TraduireMemoire('Compte'));
        O.MajTitre(2, TraduireMemoire('Date d''opération'));
        O.MajTitre(3, TraduireMemoire('Numéro pièce'));
        O.MajTitre(4, TraduireMemoire('Débit Dev.'));
        O.MajTitre(5, TraduireMemoire('Crédit Dev.'));
        O.MajTitre(6, TraduireMemoire('Devise'));

        O.MajAlign(0, ali_Gauche);
        O.MajAlign(1, ali_Gauche);
        O.MajAlign(2, ali_Centre);
        O.MajAlign(3, ali_Droite);
        O.MajAlign(4, ali_Droite);
        O.MajAlign(5, ali_Droite);
        O.MajAlign(6, ali_Centre);

        O.Imprimer;
      end
      else
        PGIInfo(TraduireMemoire('Ecritures introuvable'), Ecran.Caption);
    finally
      if Assigned(T) then FreeAndNil(T);
    end;
  end;
end;

{Controle si la référence de pointage est la dernière crée
 Renvoie True si on a une référence de pointage postérieure à celle en cours => En pointage
 sécurisé, on passe en lecture seule
{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.ControleRefPointage : Boolean;
{---------------------------------------------------------------------------------------}
var
  lQuery : TQuery;
  Code   : string;
  lMaxDatePointage : TDateTime;
begin //OK@@
  if VH^.PointageJal then Code := FJournal
                     else Code := FGeneral;
  Result := True;
  lMaxdatePointage := iDate1900;

  lQuery := OpenSql('SELECT MAX(EE_DATEPOINTAGE) DATEPOINTAGE FROM EEXBQ WHERE ' +
                    'EE_GENERAL = "' + Code + '"', True);
  try
    if not lQuery.EOF then begin
      if IsValidDate(lQuery.Findfield('DATEPOINTAGE').AsString) then
        lMaxDatePointage := lQuery.Findfield('DATEPOINTAGE').AsDateTime;
    end;
  finally
    Ferme(lQuery);
  end;

  {Il existe une référence de pointage postérieure à celle en cours => on renvoie True}
  if lMaxDatePointage > FDtEE_DatePointage then
    Exit;

  {La référence en cours est la dernière référence de pointage créée}
  Result := False;
end;

// RESTE LES SOLDES POINTES A GERER
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.InitAffichage(Arg : string);
{---------------------------------------------------------------------------------------}
var
  FBoCptEnDevise : Boolean;

    {----------------------------------------------------------------}
    procedure _RecupereCompte;
    {----------------------------------------------------------------}

       {------------------------------------------------------}
       function IsCompteDeBanque(Gene : string) : Boolean;
       {------------------------------------------------------}
       begin {03/05/07 : FQ TRESO 10449}
         Result := ExisteSQL('SELECT G_NATUREGENE FROM GENERAUX WHERE G_GENERAL = "' +
                             Gene + '" AND G_NATUREGENE IN ("BQE", "CAI")');
       end;
    var
      lQuery : TQuery;
    begin
      FNbDecimales := V_PGI.OkDecV;
      
      {Récupération du compte de compte de contrepartie en pointage sur journal}
      if VH^.PointageJal then begin
        lQuery := OpenSql('SELECT G_GENERAL, J_MODESAISIE FROM GENERAUX LEFT JOIN JOURNAL ' +
                          'ON J_CONTREPARTIE = G_GENERAL WHERE J_JOURNAL = "' + FGeneral + '"', True);
        if not lQuery.EOF then begin
          FBoBordereau := (lQuery.FindField('J_MODESAISIE').AsString = 'BOR') or
                          (lQuery.FindField('J_MODESAISIE').AsString = 'LIB');
          FJournal := FGeneral;
          FGeneral := lQuery.FindField('G_GENERAL').AsString;
        end
        else begin
          FGeneral := '';
          FJournal := '';
          PGIError(TraduireMemoire('Impossible de récupérer la contrepartie bancaire du journal'), Ecran.Caption);
        end;
        Ferme(lQuery);
      end
      {Récupération du compte de contrepartie}
      else begin
        lQuery := OpenSql('SELECT J_JOURNAL, J_MODESAISIE FROM JOURNAL WHERE J_CONTREPARTIE = "' + FGeneral + '"', True);
        if not lQuery.EOF then begin
          FJournal := lQuery.FindField('J_JOURNAL').AsString;
          FBoBordereau := (lQuery.FindField('J_MODESAISIE').AsString = 'BOR') or
                          (lQuery.FindField('J_MODESAISIE').AsString = 'LIB');
        end;
        Ferme(lQuery);
      end;

      {Récupération de la devise du compte}
      FDevise := V_PGI.DevisePivot;
      if EstPointageSurTreso then
        lQuery := OpenSql('SELECT BQ_DEVISE FROM BANQUECP WHERE BQ_CODE = "' + FGeneral + '" ORDER BY BQ_GENERAL', True)
      else
        lQuery := OpenSql('SELECT BQ_DEVISE FROM BANQUECP WHERE BQ_GENERAL = "' + FGeneral +
                          '" AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '" ORDER BY BQ_GENERAL', True);
      if not lQuery.EOF then begin
        FDevise := lQuery.FindField('BQ_DEVISE').AsString;
        FBoCptEnDevise := (FDevise <> V_PGI.DevisePivot) and (FDevise <> '');
        {Recherche du nombre de décimales de la devise}
        if FBoCptEnDevise then FNbDecimales := CalcDecimaleDevise(FDevise);
      end
      else begin
        {03/05/07 : FQ TRESO 10449 :Pour les personnes qui pointent les comptes divers (TVA notamment), on ne
                    peut taper sur BANQUECP, car la fiche BANQUECP n'est pas accessible depuis les Généraux}
        if IsCompteDeBanque(FGeneral) then
          PGIError(TraduireMemoire('Impossible de récupérer la devise du compte bancaire'), Ecran.Caption);

      end;
      Ferme(lQuery);
    end;

    {----------------------------------------------------------------}
    procedure _ChargeRefPointage(EE_GENERAL : string);
    {----------------------------------------------------------------}
    var
      lQuery: TQuery;
    begin
      lQuery := OpenSql('SELECT * FROM EEXBQ WHERE EE_GENERAL = "' + EE_GENERAL + '"' +
                        ' AND EE_REFPOINTAGE = "' + FStEE_RefPointage + '"' +
                        ' AND EE_DATEPOINTAGE = "' + UsDateTime(FDtEE_DatePointage) + '"' +
                        ' AND EE_NUMERO = ' + IntToStr(FItEE_Numero), True);

      try
        if not lQuery.Eof then begin
          if FBoCptEnDevise then begin
            FSoldeBqeDeb := Valeur(StrfMontant(lQuery.FindField('EE_NEWSOLDEDEB').AsFloat, 15, FNbDecimales, '', True));
            FSoldeBqeCre := Valeur(StrfMontant(lQuery.FindField('EE_NEWSOLDECRE').AsFloat, 15, FNbDecimales, '', True));
          end
          else begin
            FSoldeBqeDeb := Valeur(StrfMontant(lQuery.FindField('EE_NEWSOLDEDEBEURO').AsFloat, 15, FNbDecimales, '', True));
            FSoldeBqeCre := Valeur(StrfMontant(lQuery.FindField('EE_NEWSOLDECREEURO').AsFloat, 15, FNbDecimales, '', True));
          end;
        end;
      finally
        Ferme(lQuery);
      end;
    end;

var
  lSt: string;
begin
  {mode consultation ?}
  lSt := ReadTokenSt(Arg);
  FBoConsultation := (lSt = 'ACTION=CONSULTATION');
  {Compte GENERAL ou JOURNAL}
  lSt := ReadTokenSt(Arg);
  FGeneral := IIF(lSt <> '', lSt, '');
  {Date de Pointage}
  lSt := ReadTokenSt(Arg);
  FDtEE_DatePointage := StrToDate(IIF(lSt <> '', lSt, iDate1900));
  {Référence de POINTAGE}
  lSt := ReadTokenSt(Arg);
  FStEE_RefPointage := IIF(lSt <> '', lSt, '');
  {Numero interne du relevé}
  lSt := ReadTokenSt(Arg);
  FItEE_Numero := StrToInt(IIF(lSt <> '', lSt, '1'));
  {Y a-t-il un relevé rattaché à la référence de pointage ? CODEPOINTAGEMAN s'il n'y en a pas}
  lSt := ReadTokenSt(Arg);
  {14/11/07 : Gestion du pointage non bancaire}
  FPtgeManuel := (lSt = CODEPOINTAGEMAN) or (lSt = '') or (lSt = CODEPOINTAGENBQ);
  FPtgeNonBqe := (lSt = CODEPOINTAGENBQ) and (CtxCompta in V_PGI.PGIContexte) and not VH^.PointageJal;

  {Récupère la devise et, si on est en pointage sur journal, la contrepartie du journal}
  _RecupereCompte;
  {Charge la référence de pointage et affiche le solde bancaire}
  _ChargeRefPointage(iif(VH^.PointageJal, FJournal, FGeneral));

  {Mise à jour du Caption de la fiche
   12/07/07 : FQ 21057 : ajout du code à côté du libellé. En pointage sur Journal affichage du compte !}
  lSt := TraduireMemoire('Pointage du ') + FormatDateTime('dd/mm/yy', FDtEE_DatePointage) + ' (';
  if VH^.PointageJal then
    lSt := lSt + TraduireMemoire('Journal : ') + RechDom('TTJOURNAL', FJournal, False) + ' / ' + FGeneral + ')'
  else if EstPointageSurTreso then
    lSt := lSt + TraduireMemoire('Compte : ') + RechDom('TRBANQUECP', FGeneral, False) + ' / ' + FGeneral + ')'
  else
    lSt := lSt + TraduireMemoire('Compte : ') + RechDom('TTBANQUECP', FGeneral, False) + ' / ' + FGeneral + ')';
  lSt := lSt + '  /  ' + TraduireMemoire('Session n° ') + IntToStr(FItEE_Numero);
  Ecran.Caption := lSt;
  UpdateCaption(Ecran);

  if ctxTreso in V_PGI.PGIContexte then begin
    SetControlCaption('TTDATEVAL', TraduireMemoire('Date d''opération : '));
    SetControlProperty('TDATEVAL', 'LEFT', (GetControl('TTDATEVAL') as THLabel).Left +
                                           (GetControl('TTDATEVAL') as THLabel).Width + 3);
  end;

  FResteAPointe := 0;
  FEtatPointage := 0; {#ETAT}
  
  {En pointage sur journal, on n'affiche pas les dates de valeur mais les dates d'échéances}
  SetControlVisible('E_DATEECHEANCE' , VH^.PointageJal);
  SetControlVisible('E_DATEECHEANCE_', VH^.PointageJal);
  SetControlVisible('E_DATEVALEUR' , not VH^.PointageJal);
  SetControlVisible('E_DATEVALEUR_', not VH^.PointageJal);
  if VH^.PointageJal then SetControlCaption('TE_DATEECHEANCE', TraduireMemoire('Dates d''échéance du'))
                     else SetControlCaption('TE_DATEECHEANCE', TraduireMemoire('Dates de valeur du'));

  {On affiches tous les mouvements / écritures, pointés ou non}
  EtatAffichage := tv_Tout;
  {Mise à jour du libellé recapitulatif}
  AfficheEtatGrille;
  {Exclusion des dates de la gestion des filtres
   - Pour que la valeur des contrôles calculés (dépendants d'autres) ne soit pas enregistrée dans la sauvegarde
     des filtres, ni restaurée, il suffit de lui mettre un tag compris entre -9980 et  -9989.
   - Pour que la valeur des contrôles ne soit pas réinitialisée lors d'un nouveau filtre, il suffit de lui
     mettre un tag compris entre -9990 et -9999.
   - Pour les 2 à la fois, il suffit de mettre un tag compris entre -9970 et -9979.}
  (GetControl('E_DATECOMPTABLE' ) as THEdit).Tag := -9979;
  (GetControl('E_DATECOMPTABLE_') as THEdit).Tag := -9979;
  (GetControl('E_DATEVALEUR'    ) as THEdit).Tag := -9979;
  (GetControl('E_DATEVALEUR_'   ) as THEdit).Tag := -9979;
  (GetControl('E_DATEECHEANCE'  ) as THEdit).Tag := -9979;
  (GetControl('E_DATEECHEANCE_' ) as THEdit).Tag := -9979;

  {Mise à jour du format des totaux}
  (GetControl('TDLIGAFFICHE'    ) as THNumEdit).Decimals := FNbDecimales;
  (GetControl('TCLIGAFFICHE'    ) as THNumEdit).Decimals := FNbDecimales;
  (GetControl('TDLIGPOINTE'     ) as THNumEdit).Decimals := FNbDecimales;
  (GetControl('TCLIGPOINTE'     ) as THNumEdit).Decimals := FNbDecimales;
  (GetControl('TDMVTAFFICHE'    ) as THNumEdit).Decimals := FNbDecimales;
  (GetControl('TCMVTAFFICHE'    ) as THNumEdit).Decimals := FNbDecimales;
  (GetControl('TDMVTPOINTE'     ) as THNumEdit).Decimals := FNbDecimales;
  (GetControl('TCMVTPOINTE'     ) as THNumEdit).Decimals := FNbDecimales;
  (GetControl('TCAVANCEMENT'    ) as THNumEdit).Decimals := FNbDecimales;
  (GetControl('TDAVANCEMENT'    ) as THNumEdit).Decimals := FNbDecimales;
  (GetControl('TDSOLDEBANCAIRE' ) as THNumEdit).Decimals := FNbDecimales;
  (GetControl('TCSOLDEBANCAIRE' ) as THNumEdit).Decimals := FNbDecimales;
  (GetControl('TDSOLDECOMPTABLE') as THNumEdit).Decimals := FNbDecimales;
  (GetControl('TCSOLDECOMPTABLE') as THNumEdit).Decimals := FNbDecimales;

  if EstPointageSurTreso then
    SetControlCaption('TSOLDECOMPTABLE', TraduireMemoire('Solde de trésorerie'));

  {26/10/06 : FQ 10492 : Ajout d'une zone de gestion de la colonne référence pour opérations bancaires}
  SetControlVisible('TCBREFMVTBQE' , not FPtgeManuel);
  SetControlVisible('CBREFMVTBQE'  , not FPtgeManuel);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.InitControles;
{---------------------------------------------------------------------------------------}
begin//OK@@
  FListe := (GetControl('FLISTE') as THGrid);

  {Récupération du popup F11}
  POPF11 := TPopupMenu(GetControl('POPF11'));
  ActivateXPPopUp(POPF11);
  
  {Récupération des contrôles pour le rapprochement automatique}
  CCombinatoire := TCheckBox(GetControl('CCOMBINATOIRE'));
  CModeReg      := TCheckBox(GetControl('CMODEREG'));
  CRegAccro     := TCheckBox(GetControl('CREGACCRO'));
  CDate         := TCheckBox(GetControl('CDATE'));
  CNumCheque    := TCheckBox(GetControl('CNUMCHEQUE'));
  ckLibCib      := (GetControl('CKLIBCIB') as TCheckBox);
  CkLibExercice := (GetControl('CKLIBEXO') as TCheckBox); {28/11/07}
  TDate         := TSpinEdit(GetControl('TDATE'));
  TNiveau       := TSpinEdit(GetControl('TNIVEAU'));
  TNumCheque    := TSpinEdit(GetControl('TNUMCHEQUE'));
  TPosCheque    := TSpinEdit(GetControl('TPOSCHEQUE'));
  ComboRef      := THValCombobox(GetControl('COMBOREF'));
  ComboLibMvt   := THValCombobox(GetControl('CBREFMVTBQE')); {26/06/07 : FQ 10492}

  {28/12/07 : FQ 22119 : l'initialisation de la combo a été déplacé ici depuis GereTypeSession
              pour n'être exécuté qu'une seule fois et surtout avant le chargement des filtres}
  {Gestion du combo de traitement du rappro automatique}
  ComboRef.Items.Clear;
  ComboRef.Values.Clear;
  {En Tréso, on ne travaille que sur les référence internes et les libellés}
  if EstPointageSurTreso then begin
    ComboRef.Items.Add(TraduireMemoire('Libellé'));
    ComboRef.Items.Add(TraduireMemoire('Référence interne'));
    ComboRef.Values.Add('0');
    ComboRef.Values.Add('2');
  end
  else begin
    ComboRef.Items.Add(TraduireMemoire('Libellé'));
    ComboRef.Items.Add(TraduireMemoire('Référence externe'));
    ComboRef.Items.Add(TraduireMemoire('Référence interne'));
    ComboRef.Items.Add(TraduireMemoire('Référence libre'));
    ComboRef.Items.Add(TraduireMemoire('N° chèque stocké'));
    ComboRef.Values.Add('0');
    ComboRef.Values.Add('1');
    ComboRef.Values.Add('2');
    ComboRef.Values.Add('3');
    ComboRef.Values.Add('4');
  end;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.InitEvenement;
{---------------------------------------------------------------------------------------}
begin //OK@@
  CModeReg     .OnClick  := CModeRegOnClick;
  CRegAccro    .OnClick  := CRegAccroOnClick;
  ckLibCib     .OnClick  := CRegAccroOnClick;
  CCombinatoire.OnClick  := CCombinatoireOnClick;
  CDate        .OnClick  := CDateOnClick;
  CNumCheque   .OnClick  := CNumChequeOnClick;
  ComboRef     .OnChange := ComboRefOnChange;
  ComboLibMvt  .OnChange := ComboRefOnChange; {26/06/07 : FQ 10492}

  THLabel(GetControl('TCONTREPARTIE')).OnDblClick := ContrepartieDblClick; {18/07/06 : FQ 13017}

  TToolbarButton97(GetControl('BETATPREPOINTAGE')).OnClick := PrePointageOnClick;
  TToolbarButton97(GetControl('BPOINTAGEAUTO'   )).OnClick := BPointageAutoOnClick;
  TToolbarButton97(GetControl('BAGRANDIR'       )).OnClick := RedAgrandOnClick;
  TToolbarButton97(GetControl('BREDUIRE'        )).OnClick := RedAgrandOnClick;
  TToolbarButton97(GetControl('BVALIDER'        )).OnClick := OnClickBValider;
  {23/07/07 : FQ 21058 : Gestion des messages d'équilibre}
  TToolbarButton97(GetControl('BFERME'          )).OnClick := BFermeOnClick;

  Ecran.OnKeyDown := OnKeyDownEcran;
  Ecran.OnCloseQuery := FormOnCloseQuery;
  Ecran.OnDestroy := FormOnDestroy; {JP 27/12/07 : FQ 22118}
  AFindDialog.OnFind := OnFindAFindDialog;

  E_DateComptable_        := THEdit(GetControl('E_DATECOMPTABLE_', True));
  E_DateComptable_.Text   := DateToStr(FDtEE_DatePointage);
  E_DateComptable_.OnExit := OnExitE_DateComptable_;

  Ecran.OnResize          := OnResizeEcran;

  (GetControl('BRECHERCHER'    ) as TToolbarButton97).OnClick := BRechercheOnClick;
  (GetControl('BCHERCHE'       ) as TToolbarButton97).OnClick := BChercheOnClick;
  (GetControl('BVALIDEETRAPPRO') as TToolbarButton97).OnClick := OnClickBValideEtRappro;
  (GetControl('CKDEBITCREDIT'  ) as TCheckBox       ).OnClick := DebitCreditOnClick;

  if Assigned(GetControl('BSELECTALL')) then
    (GetControl('BSELECTALL') as TToolbarButton97).OnClick := OnClickBSelectAll;
  if Assigned(GetControl('BEXPORT')) then
    (GetControl('BEXPORT') as TToolbarButton97).OnClick := BExportOnClick;

  FListe.OnKeyDown     := OnKeyDownFListe;
  FListe.OnMouseDown   := OnMouseDownFListe;
  FListe.OnDblClick    := FListeOnDblClick;
  FListe.GetCellCanvas := GetCellCanvasFListe;
  FListe.OnRowEnter    := OnRowEnterFListe;
  FListe.PostDrawCell  := PostDrawCell;

  (GetControl('IVERT' ) as TImage).OnClick := CoherenceSession; {#ETAT#}
  (GetControl('IROUGE') as TImage).OnClick := CoherenceSession; {#ETAT#}

  {Popup Menu sur F11}
  {Acceder}
  TMenuItem(GetControl('ACCEDERECR'      )).OnClick := FListeOnDblClick;
  TMenuItem(GetControl('ACCEDERMVT'      )).OnClick := FListeOnDblClick;
  TMenuItem(GetControl('ACCEDERMULMVT'   )).OnClick := LancerMulEexBqLig;
  TMenuItem(GetControl('ACCEDERCIB'      )).OnClick := LancerCib;
  TMenuItem(GetControl('ACCEDERACCRO'    )).OnClick := LancerRegle;
  TMenuItem(GetControl('ACCEDERLIBCIB'   )).OnClick := LancerLibCib;
  TMenuItem(GetControl('AUTREJAL'        )).OnClick := ContrepartieDblClick; {18/07/06 : FQ 13017}
  TMenuItem(GetControl('RESTEAPOINTE'    )).OnClick := ResteAPointer;
  if Assigned(GetControl('COHERENCE')) then
    TMenuItem(GetControl('COHERENCE'     )).OnClick := CoherenceSession; {#ETAT#}
  {Acceder bis : 16/07/07 : FQ 21049 : Gestion des boutons sur le modèle du POPF11}
  TMenuItem(GetControl('ACCEDERECR1'     )).OnClick := FListeOnDblClick;
  TMenuItem(GetControl('ACCEDERMVT1'     )).OnClick := FListeOnDblClick;
  TMenuItem(GetControl('ACCEDERMULMVT1'  )).OnClick := LancerMulEexBqLig;
  TMenuItem(GetControl('ACCEDERCIB1'     )).OnClick := LancerCib;
  TMenuItem(GetControl('ACCEDERACCRO1'   )).OnClick := LancerRegle;
  TMenuItem(GetControl('ACCEDERLIBCIB1'  )).OnClick := LancerLibCib;
  TMenuItem(GetControl('AUTREJAL1'       )).OnClick := ContrepartieDblClick; {18/07/06 : FQ 13017}
  TMenuItem(GetControl('RESTEAPOINTE1'   )).OnClick := ResteAPointer;
  if Assigned(GetControl('COHERENCE1')) then
    TMenuItem(GetControl('COHERENCE1'    )).OnClick := CoherenceSession; {#ETAT#}
  {Traitement}
  TMenuItem(GetControl('INSMVTBQE'       )).OnClick := CreerEexBqLig;
  TMenuItem(GetControl('SAISIESIMPL'     )).OnClick := CreerEcriture;
  TMenuItem(GetControl('SAISIEPIECE'     )).OnClick := AccesSaisiePara;
  TMenuItem(GetControl('SUPPRPIECE'      )).OnClick := SupprEcriture;
  TMenuItem(GetControl('MODIFENTETE'     )).OnClick := ModifEntete;
  TMenuItem(GetControl('CTRLZ'           )).OnClick := AnnulerPointage;
  TMenuItem(GetControl('DEPOINTE'        )).OnClick := RAZPointage;
  if Assigned(GetControl('RAPPROCHER')) then {18/07/07 : FQ 21130}
    TMenuItem(GetControl('RAPPROCHER'    )).OnClick :=  BPointageAutoOnClick;
  TMenuItem(GetControl('VALIDER'         )).OnClick := OnClickBValider;
  TMenuItem(GetControl('VALIDERETRAPPRO' )).OnClick := OnClickBValideEtRappro;
  {Traitement bis : 16/07/07 : FQ 21049 : Gestion des boutons sur le modèle du POPF11}
  TMenuItem(GetControl('INSMVTBQE1'      )).OnClick := CreerEexBqLig;
  TMenuItem(GetControl('SAISIESIMPL1'    )).OnClick := CreerEcriture;
  TMenuItem(GetControl('SAISIEPIECE1'    )).OnClick := AccesSaisiePara;
  TMenuItem(GetControl('SUPPRPIECE1'     )).OnClick := SupprEcriture;
  TMenuItem(GetControl('MODIFENTETE1'    )).OnClick := ModifEntete;
  TMenuItem(GetControl('CTRLZ1'          )).OnClick := AnnulerPointage;
  TMenuItem(GetControl('DEPOINTE1'       )).OnClick := RAZPointage;
  TMenuItem(GetControl('VALIDER1'        )).OnClick := OnClickBValider;
  TMenuItem(GetControl('VALIDERETRAPPRO1')).OnClick := OnClickBValideEtRappro;
  {Sélection}
  TMenuItem(GetControl('TOUTSELECT'      )).OnClick := OnClickBSelectAll;
  TMenuItem(GetControl('SELECTRAPIDE'    )).OnClick := OnClickSelectionRapide;
  TMenuItem(GetControl('SELECTJOUR'      )).OnClick := OnClickSelectionJour;
  TMenuItem(GetControl('SELECTMOIS'      )).OnClick := OnClickSelectionMois;
  {Sélection bis : 16/07/07 : FQ 21049 : Gestion des boutons sur le modèle du POPF11}
  TMenuItem(GetControl('TOUTSELECT1'     )).OnClick := OnClickBSelectAll;
  TMenuItem(GetControl('SELECTRAPIDE1'   )).OnClick := OnClickSelectionRapide;
  TMenuItem(GetControl('SELECTJOUR1'     )).OnClick := OnClickSelectionJour;
  TMenuItem(GetControl('SELECTMOIS1'     )).OnClick := OnClickSelectionMois;
  {Déplacement}
  TMenuItem(GetControl('POINTESUIVANT'   )).OnClick := NextPriorByFieldOnClick;
  TMenuItem(GetControl('POINTEPRE'       )).OnClick := NextPriorByFieldOnClick;
  TMenuItem(GetControl('NONPOINTESUIVANT')).OnClick := NextPriorByFieldOnClick;
  TMenuItem(GetControl('NONPOINTEPRE'    )).OnClick := NextPriorByFieldOnClick;
  TMenuItem(GetControl('LIBELLESUIVANT'  )).OnClick := NextPriorByFieldOnClick;
  TMenuItem(GetControl('MOISSUIVANT'     )).OnClick := NextPriorByFieldOnClick;
  TMenuItem(GetControl('JOURSUIVANT'     )).OnClick := NextPriorByFieldOnClick;
  {Déplacement bis : 16/07/07 : FQ 21049 : Gestion des boutons sur le modèle du POPF11}
  TMenuItem(GetControl('POINTESUIVANT1'   )).OnClick := NextPriorByFieldOnClick;
  TMenuItem(GetControl('POINTEPRE1'       )).OnClick := NextPriorByFieldOnClick;
  TMenuItem(GetControl('NONPOINTESUIVANT1')).OnClick := NextPriorByFieldOnClick;
  TMenuItem(GetControl('NONPOINTEPRE1'    )).OnClick := NextPriorByFieldOnClick;
  TMenuItem(GetControl('LIBELLESUIVANT1'  )).OnClick := NextPriorByFieldOnClick;
  TMenuItem(GetControl('MOISSUIVANT1'     )).OnClick := NextPriorByFieldOnClick;
  TMenuItem(GetControl('JOURSUIVANT1'     )).OnClick := NextPriorByFieldOnClick;
  {Filtres}
  TMenuItem(GetControl('MVTBANCAIRE'     )).OnClick := PopupFiltreOnClick;
  TMenuItem(GetControl('MVTCOMPTABLE'    )).OnClick := PopupFiltreOnClick;
  TMenuItem(GetControl('MVTPOINTE'       )).OnClick := PopupFiltreOnClick;
  TMenuItem(GetControl('MVTNONPOINTE'    )).OnClick := PopupFiltreOnClick;
  TMenuItem(GetControl('MVTSESSION'      )).OnClick := PopupFiltreOnClick;
  TMenuItem(GetControl('MVTENCOURS'      )).OnClick := PopupFiltreOnClick;
  TMenuItem(GetControl('MVTTOUT'         )).OnClick := PopupFiltreOnClick;
  {Filtres bis : 16/07/07 : FQ 21049 : Gestion des boutons sur le modèle du POPF11}
  TMenuItem(GetControl('MVTBANCAIRE1'    )).OnClick := PopupFiltreOnClick;
  TMenuItem(GetControl('MVTCOMPTABLE1'   )).OnClick := PopupFiltreOnClick;
  TMenuItem(GetControl('MVTPOINTE1'      )).OnClick := PopupFiltreOnClick;
  TMenuItem(GetControl('MVTNONPOINTE1'   )).OnClick := PopupFiltreOnClick;
  TMenuItem(GetControl('MVTSESSION1'     )).OnClick := PopupFiltreOnClick;
  TMenuItem(GetControl('MVTENCOURS1'     )).OnClick := PopupFiltreOnClick;
  TMenuItem(GetControl('MVTTOUT1'        )).OnClick := PopupFiltreOnClick;

  ActivateXPPopUp(TPopupMenu(GetControl('POP27')));
  ADDMenuPop(TPopupMenu(GetControl('POP27')), '', '');

  {20/07/07 : FQ 21101 : Gestion dans le OnPopup de la visibilité de ceratins menus}
  TPopupMenu(GetControl('PPACCEDER')).OnPopup := PopupOnPopup;
  TPopupMenu(GetControl('POPF11'   )).OnPopup := PopupOnPopup;
end;

{Prépare la fiche en fonction du type de session : manuel ou non}
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.GereTypeSession;
{---------------------------------------------------------------------------------------}
var
  pc : TPageControl;
begin //OK@@
  FListe.MultiSelect := FPtgeManuel;
  pc := (GetControl('PAGECONTROL') as TPageControl);
  if FPtgeManuel then begin
    pc.Pages[1].TabVisible := False;
    pc.Pages[1].Visible    := False;
    {09/07/07 : FQ 10492 : Ajout de l'onglet affichage
    pc.Pages[0].TabVisible := False;
    pc.Height := pc.Height - 7;}
  end;

  pc.Pages[0].Visible := True;
  pc.ActivePageIndex  := 0;

  {PopupF11}
  {Accéder A}
  TMenuItem(GetControl('ACCEDERECR'      )).Visible := True;
  {20/07/07 : FQ 21101 : Une seule des 2 lignes n'est visible entre ACCEDERECR, ACCEDERMVT}
  TMenuItem(GetControl('ACCEDERMVT'      )).Visible := False;
  TMenuItem(GetControl('ACCEDERMULMVT'   )).Visible := not FPtgeNonBqe;
  TMenuItem(GetControl('ACCEDERCIB'      )).Visible := not FPtgeManuel;
  TMenuItem(GetControl('ACCEDERACCRO'    )).Visible := not FPtgeManuel;
  TMenuItem(GetControl('ACCEDERLIBCIB'   )).Visible := not FPtgeManuel;
  TMenuItem(GetControl('AUTREJAL'        )).Visible := not VH^.PointageJal; {18/07/06 : FQ 13017}
  TMenuItem(GetControl('RESTEAPOINTE'    )).Visible := True;
  if Assigned(GetControl('COHERENCE')) then
    TMenuItem(GetControl('COHERENCE'     )).Visible := not FPtgeManuel; {#ETAT#}
  {Accéder A bis : 16/07/07 : FQ 21049 : Gestion des boutons sur le modèle du POPF11}
  TMenuItem(GetControl('ACCEDERECR1'     )).Visible := True;
  {20/07/07 : FQ 21101 : Une seule des 2 lignes n'est visible entre ACCEDERECR1, ACCEDERMVT1}
  TMenuItem(GetControl('ACCEDERMVT1'     )).Visible := False;
  TMenuItem(GetControl('ACCEDERMULMVT1'  )).Visible := not FPtgeNonBqe;
  TMenuItem(GetControl('ACCEDERCIB1'     )).Visible := not FPtgeManuel;
  TMenuItem(GetControl('ACCEDERACCRO1'   )).Visible := not FPtgeManuel;
  TMenuItem(GetControl('ACCEDERLIBCIB1'  )).Visible := not FPtgeManuel;
  TMenuItem(GetControl('AUTREJAL1'       )).Visible := not VH^.PointageJal; {18/07/06 : FQ 13017}
  TMenuItem(GetControl('RESTEAPOINTE1'   )).Visible := True;
  if Assigned(GetControl('COHERENCE1')) then
    TMenuItem(GetControl('COHERENCE1'    )).Visible := not FPtgeManuel; {#ETAT#}

  {Traitements}
  TMenuItem(GetControl('TRAITEMENT'      )).Visible := not FBoConsultation ;
  TMenuItem(GetControl('INSMVTBQE'       )).Visible := FBoDroitMouvement and FPtgeManuel and not FBoConsultation and not FPtgeNonBqe;
  TMenuItem(GetControl('SAISIEPIECE'     )).Visible := FBoDroitEcritures and not FBoConsultation and not FPtgeNonBqe;// and
                                                       //(not FPtgeManuel or (CtxTreso in V_PGI.PGIContexte));
  TMenuItem(GetControl('SAISIESIMPL'     )).Visible := FBoDroitEcritures and not FBoConsultation and not FPtgeNonBqe;
  TMenuItem(GetControl('SUPPRPIECE'      )).Visible := FBoDroitEcritures and not FBoConsultation and not FBoBordereau;
  TMenuItem(GetControl('MODIFENTETE'     )).Visible := FBoDroitEcritures and not FBoConsultation and
                                                       not FBoBordereau  and not (CtxTreso in V_PGI.PGIContexte);
  TMenuItem(GetControl('CTRLZ'           )).Visible := not FBoConsultation;
  TMenuItem(GetControl('DEPOINTE'        )).Visible := not FBoConsultation;
  if Assigned(GetControl('RAPPROCHER')) then {18/07/07 : FQ 21130}
    TMenuItem(GetControl('RAPPROCHER'    )).Visible :=  not FPtgeManuel and not FBoConsultation;
  TMenuItem(GetControl('VALIDER'         )).Visible := not FBoConsultation;
  TMenuItem(GetControl('VALIDERETRAPPRO' )).Visible := not FBoConsultation;
  {Traitements bis : 16/07/07 : FQ 21049 : Gestion des boutons sur le modèle du POPF11}
  TMenuItem(GetControl('INSMVTBQE1'      )).Visible := FBoDroitMouvement and FPtgeManuel and not FBoConsultation and not FPtgeNonBqe;
  TMenuItem(GetControl('SAISIESIMPL1'     )).Visible := FBoDroitEcritures and not FBoConsultation and not FPtgeNonBqe;
  TMenuItem(GetControl('SAISIEPIECE1'    )).Visible := FBoDroitEcritures and not FBoConsultation and not FPtgeNonBqe;//and
                                                       //(not FPtgeManuel or (CtxTreso in V_PGI.PGIContexte));
  TMenuItem(GetControl('SUPPRPIECE1'     )).Visible := FBoDroitEcritures and not FBoConsultation and not FBoBordereau;
  TMenuItem(GetControl('MODIFENTETE1'    )).Visible := FBoDroitEcritures and not FBoConsultation and
                                                       not FBoBordereau  and not (CtxTreso in V_PGI.PGIContexte);
  TMenuItem(GetControl('CTRLZ1'          )).Visible := not FBoConsultation;
  TMenuItem(GetControl('DEPOINTE1'       )).Visible := not FBoConsultation;
  TMenuItem(GetControl('VALIDER1'        )).Visible := not FBoConsultation;
  TMenuItem(GetControl('VALIDERETRAPPRO1')).Visible := not FBoConsultation;
  SetControlVisible('BTRAITEMENT', not FBoConsultation);

  {Sélection}
  TMenuItem(GetControl('SELECTION'       )).Visible := FPtgeManuel;
  TMenuItem(GetControl('TOUTSELECT'      )).Visible := FPtgeManuel;
  TMenuItem(GetControl('SELECTRAPIDE'    )).Visible := FPtgeManuel;
  TMenuItem(GetControl('SELECTJOUR'      )).Visible := FPtgeManuel;
  TMenuItem(GetControl('SELECTMOIS'      )).Visible := FPtgeManuel;
  {Sélection bis : 16/07/07 : FQ 21049 : Gestion des boutons sur le modèle du POPF11}
  TMenuItem(GetControl('TOUTSELECT1'      )).Visible := FPtgeManuel;
  TMenuItem(GetControl('SELECTRAPIDE1'    )).Visible := FPtgeManuel;
  TMenuItem(GetControl('SELECTJOUR1'      )).Visible := FPtgeManuel;
  TMenuItem(GetControl('SELECTMOIS1'      )).Visible := FPtgeManuel;
  SetControlVisible('BSELECTION', FPtgeManuel);

  {Déplacement dans la grille}
  TMenuItem(GetControl('DEPLACEMENT'     )).Visible := True;
  TMenuItem(GetControl('POINTESUIVANT'   )).Visible := True;
  TMenuItem(GetControl('POINTEPRE'       )).Visible := True;
  TMenuItem(GetControl('NONPOINTESUIVANT')).Visible := True;
  TMenuItem(GetControl('NONPOINTEPRE'    )).Visible := True;
  TMenuItem(GetControl('LIBELLESUIVANT'  )).Visible := True;
  TMenuItem(GetControl('MOISSUIVANT'     )).Visible := True;
  TMenuItem(GetControl('JOURSUIVANT'     )).Visible := True;
  {Déplacement dans la grille bis : 16/07/07 : FQ 21049 : Gestion des boutons sur le modèle du POPF11}
  TMenuItem(GetControl('POINTESUIVANT1'   )).Visible := True;
  TMenuItem(GetControl('POINTEPRE1'       )).Visible := True;
  TMenuItem(GetControl('NONPOINTESUIVANT1')).Visible := True;
  TMenuItem(GetControl('NONPOINTEPRE1'    )).Visible := True;
  TMenuItem(GetControl('LIBELLESUIVANT1'  )).Visible := True;
  TMenuItem(GetControl('MOISSUIVANT1'     )).Visible := True;
  TMenuItem(GetControl('JOURSUIVANT1'     )).Visible := True;

  {Filtre sur la grille}
  TMenuItem(GetControl('MVTBANCAIRE'     )).Visible := not FPtgeNonBqe; {not FPtgeManuel; 18/07/07 : FQ 21132}
  TMenuItem(GetControl('MVTCOMPTABLE'    )).Visible := not FPtgeNonBqe; {not FPtgeManuel; 18/07/07 : FQ 21132}
  TMenuItem(GetControl('MVTPOINTE'       )).Visible := True;
  TMenuItem(GetControl('MVTNONPOINTE'    )).Visible := True;
  TMenuItem(GetControl('MVTSESSION'      )).Visible := True;
  TMenuItem(GetControl('MVTENCOURS'      )).Visible := True;
  TMenuItem(GetControl('MVTTOUT'         )).Visible := True;
  {Filtre sur la grille bis : 16/07/07 : FQ 21049 : Gestion des boutons sur le modèle du POPF11}
  TMenuItem(GetControl('MVTBANCAIRE1'    )).Visible := not FPtgeNonBqe; {not FPtgeManuel; 18/07/07 : FQ 21132}
  TMenuItem(GetControl('MVTCOMPTABLE1'   )).Visible := not FPtgeNonBqe; {not FPtgeManuel; 18/07/07 : FQ 21132}
  TMenuItem(GetControl('MVTPOINTE1'      )).Visible := True;
  TMenuItem(GetControl('MVTNONPOINTE1'   )).Visible := True;
  TMenuItem(GetControl('MVTSESSION1'     )).Visible := True;
  TMenuItem(GetControl('MVTENCOURS1'     )).Visible := True;
  TMenuItem(GetControl('MVTTOUT1'        )).Visible := True;

  {11/05/07 : FQ 20284 :Ajout du bouton allselect, à gérer comme les menus}
  if Assigned(GetControl('BSELECTALL')) then
    SetControlVisible('BSELECTALL', FPtgeManuel);

end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.InitVariables;
{---------------------------------------------------------------------------------------}
var
  RC : TControlFiltre;
begin //OK@@
  FPeutFermer := True;
  {Est-on en pointage sécurisé ?}
  FBoSecurise := GetParamSocSecur('SO_CPPOINTAGESECU', False);
  {Création des Tob pour le rapprochement automatique}
  TobReleve := TOB.Create('µRELEVE', nil, -1);
  TobCIB    := TOB.Create('µCIB', nil, -1);
  TobRegles := TOB.Create('µREGLE', nil, -1);
  TobLibCib := TOB.Create('LIBCIB', nil, -1);
  TobRegles.LoadDetailFromSQL('SELECT * FROM REGLEACCRO');
  TobLibCib.LoadDetailFromSQL('SELECT CC_CODE, CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE = "TCI"');
  {Nombre de décimales des montants}
  FNbDecimales := V_PGI.OkDecV;
  {Pour la recherche dans la grille}
  AFindDialog := TFindDialog.Create(Ecran);
  {Gestion des filtres}
  RC.Filtres  := (GetControl('FFILTRES'   ) as THValComboBox);
  RC.Filtre   := (GetControl('BFILTRE'    ) as TToolbarButton97);
  RC.PageCtrl := (GetControl('PAGECONTROL') as TPageControl);

  ObjFiltre := TObjFiltre.Create(RC, 'CPPOINTAGERAP');
  {28/12/07 : FQ 22119 : maintenant fait à la fin du OnArgument
  ObjFiltre.Charger;}
  
  FEtatRapproOk := False;
  {12/07/07 : FQ 21053 : à True : Pour ne pas demander de confirmation à la validation}
  FPasMessage := False;
  {28/11/07 : Date et exercice des derniers à-nouveaux issus d'une clôture définitive}
  FDateAno   := iDate1900;
  FExoAno    := '';
  FSaufVVert := False;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.BRechercheOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  FFindFirst := True;
  AFindDialog.Execute;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.OnFindAFindDialog(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  Rechercher(FListe, AFindDialog, FFindFirst);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.BChercheOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  RafraichitFiche(2);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.RedAgrandOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  if UpperCase((Sender as TToolbarButton97).Name) = 'BREDUIRE' then begin
    SetControlVisible('PNFILTRE'   , True);
    SetControlVisible('PAGECONTROL', True);
    SetControlVisible('BREDUIRE'   , False);
    SetControlVisible('BAGRANDIR'  , True);
  end
  else begin
    SetControlVisible('PAGECONTROL', False);
    SetControlVisible('PNFILTRE'   , False);
    SetControlVisible('BREDUIRE'   , True);
    SetControlVisible('BAGRANDIR'  , False);
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.GetModeAffichage : Boolean;
{---------------------------------------------------------------------------------------}
begin //OK@@
  Result := GetCheckBoxState('CKDEBITCREDIT') = cbChecked;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.GetDateAno : TDateTime; {28/11/07}
{---------------------------------------------------------------------------------------}
var
  Exo : TExoDate;
  MOK : Boolean;
  DtC : TDateTime;
  DtA : TDateTime;
begin
  if FDateAno = iDate1900 then begin
    DtC := StrToDate(E_DateComptable_.Text);
    {On commence par rechercher la Date des a-nouveaux}
    DtA := VH^.ExoV8.Deb;
    if Dtc >= GetEnCours.Deb then
      DtA := GetEnCours.Deb
    else begin
      QuelExoDate(DtC, DtC, Mok, Exo);
      if Exo.Deb > DtA then DtA := Exo.Deb;
    end;
    FDateAno := DtA;
  end;

  Result := FDateAno;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.GetExoAno : string;    {28/11/07}
{---------------------------------------------------------------------------------------}
var
  Exo : TExoDate;
  MOK : Boolean;
  DtC : TDateTime;
begin
  if FExoAno = '' then begin
    DtC := StrToDate(E_DateComptable_.Text);
    {On commence par rechercher la Date des a-nouveaux}
    if Dtc >= GetEnCours.Deb then
      FExoAno := GetEnCours.Code
    else begin
      QuelExoDate(DtC, DtC, Mok, Exo);
      FExoAno := Exo.Code;
    end;
  end;

  Result := FExoAno;
end;

{#ETAT# : Le montant de cohérence de la session est fonction du type de la session
{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.GetMntCoherence : Double;
{---------------------------------------------------------------------------------------}
begin //OK@@
  if FPtgeManuel then Result := FResteAPointe
                 else Result := FEtatPointage;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.DebitCreditOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  RafraichitFiche(2);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.FListeOnDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  if not Assigned(Fliste.Objects[0, Fliste.Row]) then Exit;

  if TOB(Fliste.Objects[0, Fliste.Row]).GetString('CLE_JOURNAL') = LIGNEBANCAIRE then
    LanceFicheEexBqLig(taModif)
  else
    LancerSaisie;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.AnnulerPointage(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  if PgiAsk(TraduireMemoire('Voulez-vous annuler le pointage en cours ?'), Ecran.Caption) = mrYes then
    ReinitialisePointage;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.RAZPointage(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin //OK@@
  if FBoConsultation then Exit;
  if PgiAsk(TraduireMemoire('Voulez-vous supprimer tout le pointage de la session courante ?'), Ecran.Caption) = mrYes then
    RemetRefPointageABlanc;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.FormOnCloseQuery(Sender : TObject; var CanClose : Boolean);
{---------------------------------------------------------------------------------------}
var
  Code : string;
  DtPg : string;
begin
  if FEtatRapproOk then begin
    if VH^.PointageJal then Code := FJournal
                       else Code := FGeneral;
    DtPg := DateToStr(FDtEE_DatePointage);
  end;

  {Si on n'est pas en consultation, on propose l'enregistrement du pointage
   23/07/07 : FQ 21053 et 21058 : devenu inutile, car appeler directement sur les évènements des boutons
   if not FBoConsultation then MajPointage;}

  {Si tout s'est bien passé ou qu'on est en consultation, on autorise la fermeture}
  CanClose := FPeutFermer or FBoConsultation;
  FPeutFermer := True;
  if CanClose then TFVierge(Ecran).FormCloseQuery(Sender, CanClose);
  if FEtatRapproOk then begin
    FEtatRapproOk := False;
    CC_LanceFicheEtatRapproDet(Code + ';' + DtPg + ';X');
  end;
end;

{JP 27/12/07 : FQ 22118 : Lors du Double VVErt, la destruction du TListeByUser de l'objet filtre
               entraine une violation d'accès s'il y a un filtre à l'écran => déplacement de la
               destruction de l'objet dans le FormDestroy et non plus dans le OnClose. Je ne
               serais pas surpris dans le destroy du TListeByUser on libère les composants filtre !
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.FormOnDestroy(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if Assigned(ObjFiltre) then FreeAndNil(ObjFiltre);
  TFVierge(Ecran).FormDestroy(Sender);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.AfficheEtatGrille;
{---------------------------------------------------------------------------------------}
var //OK@@
  Libelle : string;
begin
  case EtatAffichage of
    tv_Tout      : Libelle := TraduireMemoire('Affichage de tous les mouvements et écritures');
    tv_Bancaire  : Libelle := TraduireMemoire('Affichage de tous les mouvements bancaires');
    tv_Comptable : Libelle := TraduireMemoire('Affichage de toutes les écritures');
    tv_Pointe    : Libelle := TraduireMemoire('Affichage des mouvements et écritures pointés');
    tv_NonPointe : Libelle := TraduireMemoire('Affichage des mouvements et écritures non pointés');
    tv_Session   : Libelle := TraduireMemoire('Exclusion des écritures d''une autre session');
  end;

  SetControlCaption('TRECAPITULATIF', Libelle);
end;

{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.ConvertitTob(TobEcr : TOB; Code : string) : TOB;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  F : TOB;
  T : TOB;
  s : string;
begin //OK@@
  {02/10/07 : ajout d'une précaution}
  if not Assigned(TobEcr) then begin
    Result := nil;
    Exit;
  end;

  Result := TOB.Create('', nil, -1);
  if EstPointageSurTreso then s := 'TE_'
                         else s := 'E_';
  for n := 0 to TobEcr.Detail.Count - 1 do begin
    T := TobEcr.Detail[n];
    if (EstPointageSurTreso and IsTobAPointe(T)) or
       (not EstPointageSurTreso and IsTobAPointe(T) and {Vérifie si la ligne correspond bien au mouvement bancaire}
        ((TPieceCompta(TobEcr).EstBQE(n + 1) and not VH^.PointageJal) or
         (VH^.PointageJal and not TPieceCompta(TobEcr).EstBQE(n + 1)))) then begin
      F := TOB.Create('', Result, -1);
      AddChampPointage(F);
      F.SetString('CLE_JOURNAL'        , T.GetString(s + 'JOURNAL'));
      F.SetString('CLE_GENERAL'        , T.GetString(s + 'GENERAL'));
      F.SetString('CLE_LIBELLE'        , T.GetString(s + 'LIBELLE'));
      F.SetString('CLE_REFINTERNE'     , T.GetString(s + 'REFINTERNE'));
      F.SetString('CLE_DEVISE'         , T.GetString(s + 'DEVISE'));
      F.SetString('CLE_EXERCICE'       , T.GetString(s + 'EXERCICE'));
      F.SetInteger('CLE_NUMEROPIECE'   , T.GetInteger(s + 'NUMEROPIECE'));
      F.SetInteger('CLE_NUMLIGNE'      , T.GetInteger(s + 'NUMLIGNE'));
      F.SetInteger('CLE_NUMECHE'       , T.GetInteger(s + 'NUMECHE'));
      F.SetDateTime('CLE_DATECOMPTABLE', T.GetDateTime(s + 'DATECOMPTABLE'));
      F.SetDateTime('CLE_DATEVALEUR'   , T.GetDateTime(s + 'DATEVALEUR'));

      if EstPointageSurTreso then begin
        F.SetString(TRESO_NODOSSIER  , T.GetString('TE_NODOSSIER'));
        F.SetString(TRESO_USERCOMPTA , T.GetString('TE_USERCOMPTABLE'));
        F.SetString(TRESO_ORIGINE    , T.GetString('TE_QUALIFORIGINE'));
        F.SetString(TRESO_NUMLIGNE   , T.GetString('TE_NUMLIGNE'));
        F.SetString(TRESO_NUMTRANSAC , T.GetString('TE_NUMTRANSAC'));

        F.SetDouble('CLE_MONTANT'    , T.GetDouble('TE_MONTANT'));
        F.SetDouble('CLE_MONTANTDEV' , T.GetDouble('TE_MONTANTDEV'));
        if T.GetDouble('TE_MONTANT') < 0 then begin
          F.SetString('CLE_CIB'      , T.GetString('TE_CODECIB'));
          F.SetDouble('CLE_DEBIT'    , T.GetDouble('TE_MONTANT') * -1);
          F.SetDouble('CLE_DEBITDEV' , T.GetDouble('TE_MONTANTDEV') * -1);
        end else begin
          F.SetString('CLE_CIB'      , T.GetString('TE_CODECIB'));
          F.SetDouble('CLE_CREDIT'   , T.GetDouble('TE_MONTANT'));
          F.SetDouble('CLE_CREDITDEV', T.GetDouble('TE_MONTANTDEV'));
        end;
      end

      {Pointage sur ECRITURE}
      else begin
        F.SetString('CLE_QUALIFPIECE'  , T.GetString('E_QUALIFPIECE'));
        F.SetString('CLE_NUMTRAITECHQ' , T.GetString('E_NUMTRAITECHQ'));
        F.SetString('CLE_REFEXTERNE'   , T.GetString(s + 'REFEXTERNE'));
        F.SetString('CLE_REFLIBRE'     , T.GetString(s + 'REFLIBRE'));
        if VH^.PointageJal then begin
          F.SetDouble('CLE_MONTANT'    , T.GetDouble('E_CREDIT') -
                                         T.GetDouble('E_DEBIT'));
          F.SetDouble('CLE_MONTANTDEV' , T.GetDouble('E_CREDITDEV') -
                                         T.GetDouble('E_DEBITDEV'));
        end else begin
          F.SetDouble('CLE_MONTANT'    , T.GetDouble('E_DEBIT') -
                                         T.GetDouble('E_CREDIT'));
          F.SetDouble('CLE_MONTANTDEV' , T.GetDouble('E_DEBITDEV') -
                                         T.GetDouble('E_CREDITDEV'));
        end;
        F.SetDouble('CLE_CREDIT'     , T.GetDouble('E_CREDIT'));
        F.SetDouble('CLE_DEBIT'      , T.GetDouble('E_DEBIT'));
        F.SetDouble('CLE_DEBITDEV'   , T.GetDouble('E_DEBITDEV'));
        F.SetDouble('CLE_CREDITDEV'  , T.GetDouble('E_CREDITDEV'));

        F.SetString('CLE_CIB'        , T.GetString('E_MODEPAIE'));
      end;
      F.SetString('MODIFIE', 'X');
      F.SetString('CLE_QUALIFPIECE', 'N');

      if EstPointageSurTreso or not (FBoBordereau or VH^.PointageJal) then begin
        {27/07/07 : suite aux modifications de ce jour dans IsTobAPointe}
        if Assigned(TOB(FListe.Objects[0, FListe.Row])) and
           (TOB(FListe.Objects[0, FListe.Row]).GetString('CLE_JOURNAL') <> LIGNEBANCAIRE) then begin
          F.SetString('CLE_POINTE', '');
          F.SetDateTime('CLE_DATEPOINTAGE', iDate1900);
          F.SetString('CLE_REFPOINTAGE', '');
          Continue;
        end;
      end;

      F.SetString('CLE_POINTE', Code);
      F.SetDateTime('CLE_DATEPOINTAGE', FDtEE_DatePointage);
      F.SetString('CLE_REFPOINTAGE', FStEE_RefPointage);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.MajAvancement;
{---------------------------------------------------------------------------------------}
var
  stAvanc  : string;
  stGene   : string;
begin
  {Mise à jour de l'avancement du pointage.
   24/07/07 : #ETAT# : Ajout de la notion de cohérence du pointage dans les session de rapprochement}
  if MntCoherence = 0 then stAvanc := 'X'
                      else stAvanc := '-';

  if VH^.PointageJal then stGene := FJournal
                     else stGene := FGeneral;
  ExecuteSQL('UPDATE EEXBQ SET EE_AVANCEMENT = "' + stAvanc + '" WHERE EE_GENERAL = "' + stGene +
             '" AND EE_DATEPOINTAGE = "' + UsDateTime(FDtEE_DatePointage) +
             '" AND EE_NUMERO = ' + IntToStr(FItEE_Numero) +
             ' AND EE_REFPOINTAGE = "' + FStEE_RefPointage + '"');
end;

{Mise à jour des totaux pointés
 Je ne suis pas sûr de l'intérêt de s'embêter avec les totaux dans la mesure ou en n'étant
 plus utilisés dans le pointage, il ne doivent plus servir à grand chose.
 Par ailleurs, avec la saisie automatique, les lignes sont pointées sans mise à jour des
 totaux pointés, ce qui les rend faux !!!!
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.SetTotauxPointes;
{---------------------------------------------------------------------------------------}
var
  FTotDebPTP : Double;
  FTotDebPTD : Double;
  FTotCrePTP : Double;
  FTotCrePTD : Double;
  n : Integer;
  F : TOB;
begin
  if EstPointageSurTreso then Exit;
  
  FTotDebPTP := 0;
  FTotDebPTD := 0;
  FTotCrePTP := 0;
  FTotCrePTD := 0;
  for n := 0 to TobReleve.Detail.Count - 1 do begin
    F := TobReleve.Detail[n];
    {On est sur un mouvement bancaire, on passe à la ligne suivante}
    if F.GetString('CLE_JOURNAL') = LIGNEBANCAIRE then Continue;

    {La ligne a été modifiée et notamment la référence de pointage}
    if (F.GetString('CLE_REFPOINTAGE') <> F.GetString('CLE_OLDREF')) and
       (F.GetString('MODIFIE') = 'X') then begin
      {On vient de pointer l'écriture}
      if FStEE_refPointage = F.GetString('CLE_REFPOINTAGE') then begin
        if VH^.PointageJal then begin
          FTotDebPTP := FTotDebPTP + F.GetDouble('E_CREDIT');
          FTotDebPTD := FTotDebPTD + F.GetDouble('E_CREDITDEV');
          FTotCrePTP := FTotCrePTP + F.GetDouble('E_DEBIT');
          FTotCrePTD := FTotCrePTD + F.GetDouble('E_DEBITDEV');
        end else begin
          FTotDebPTP := FTotDebPTP + F.GetDouble('E_DEBIT');
          FTotDebPTD := FTotDebPTD + F.GetDouble('E_DEBITDEV');
          FTotCrePTP := FTotCrePTP + F.GetDouble('E_CREDIT');
          FTotCrePTD := FTotCrePTD + F.GetDouble('E_CREDITDEV');
        end;
      end

      {On vient de dépointer l'écriture}
      else begin
        if VH^.PointageJal then begin
          FTotDebPTP := FTotDebPTP - F.GetDouble('E_CREDIT');
          FTotDebPTD := FTotDebPTD - F.GetDouble('E_CREDITDEV');
          FTotCrePTP := FTotCrePTP - F.GetDouble('E_DEBIT');
          FTotCrePTD := FTotCrePTD - F.GetDouble('E_DEBITDEV');
        end else begin
          FTotDebPTP := FTotDebPTP - F.GetDouble('E_DEBIT');
          FTotDebPTD := FTotDebPTD - F.GetDouble('E_DEBITDEV');
          FTotCrePTP := FTotCrePTP - F.GetDouble('E_CREDIT');
          FTotCrePTD := FTotCrePTD - F.GetDouble('E_CREDITDEV');
        end;
      end;
    end;
  end;

  MajTotauxPointe(FGeneral, FTotDebPTP, FTotCrePTP, FTotDebPTD, FTotCrePTD, False, V_Pgi.SchemaName);
end;

{$IFDEF TRESO}
{Mets les reférence de pointage sur la pièce en train d'être générée
{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.MajRefPointage(var TOBPiece : TobPieceCompta; TobTreso : TOB = nil) : Boolean;
{---------------------------------------------------------------------------------------}
{$ELSE}
{---------------------------------------------------------------------------------------}
function TOF_CPPOINTAGERAP.MajRefPointage(var TOBPiece : TPieceCompta; TobTreso : TOB = nil) : Boolean;
{---------------------------------------------------------------------------------------}
{$ENDIF TRESO}
var //OK@@
  p : string;
  n : Integer;
  k : Integer;
  F : TOB;
  aTob : TOB;
  PtOk : Boolean; {27/07/07}
begin
  Result := False;
  if FPtgeManuel then P := '0'
                 else p := GetMaxNumPtge;

  {$IFDEF TRESO}
  {02/10/07 : On peut répondre non lors de la demande d'intégration en compta.
              Si on n'est pas en pointage sur TRECRITURE, on abandonne}
  if not Assigned(TOBPiece) and not EstPointageSurTreso then Exit;
  {..., sinon on travaille sur la tobTreso}
  if (Assigned(TOBPiece) and (TOBPiece.Detail.Count > 0)) or
     (TobTreso.Detail[0].GetString('TE_GENERAL') <> '') then begin
    {Ecriture des flux générés}
    if not TobTreso.InsertDB(nil) then Exit;
    {Mise à jour des soldes de Trésorerie}
    RecalculSolde(TobTreso.Detail[0].GetString('TE_GENERAL'), TobTreso.Detail[0].GetString('TE_DATECOMPTABLE'), 0, True);
    {Intégration en compta}
    {En Tréso, TOBPiece est un conteneur général et c'est le détail qui contient les pièces}
    if not Assigned(TOBPiece) or not TRIntegrationPieces(TOBPiece, False) or (TOBPiece.Detail.Count = 0)then
      {Si on n'est pas en pointage sur TRECRITURE, il ne sert à rien de poursuivre}
      if not EstPointageSurTreso then Exit;

    {25/07/07 : Les champ TE_NUMEROPIECE, TE_USERCOMPTABLE, TE_CPNUMLIGNE, ne sont plus à jour car ils ont été
                modifiés par requête lors de l'écriture de la pièce comptable en base : cf. UProcEcriture.TREnregistrePieces
     02/10/07 : On s'assure que l'on a demandé l'intégration en comptabilité}
    if EstPointageSurTreso and Assigned(TOBPiece) then begin
      aTob := TOBPiece.Detail[0].FindFirst(['E_GENERAL'], [GetGeneFromBqCode(TobTreso.Detail[0].GetString('TE_GENERAL'))], True);
      if Assigned(aTob) then begin
        TobTreso.Detail[0].SetInteger('TE_NUMEROPIECE', aTob.GetInteger('E_NUMEROPIECE'));
        TobTreso.Detail[0].SetInteger('TE_CPNUMLIGNE', aTob.GetInteger('E_NUMLIGNE'));
        TobTreso.Detail[0].SetString('TE_USERCOMPTABLE', V_PGI.User);
        {25/07/07 : Théoriquement TE_NUMECHE est à 1, mais par précaution car Stéphane le met à jour dans TREnregistrePieces}
        TobTreso.Detail[0].SetInteger('TE_NUMECHE', 1);
      end
      {25/07/07 : si l'on ne peut mettre à jour la tob, il va y avoir un problème lors de la mise à jour de la
                  table ECRITURE dans EcritPointage. Ce n'est pas bloquant, il suffit de recharger la tob par
                  une mouette bleue}
      else
        PGIError(TraduireMemoire('Le flux de Trésorerie n''a pu être mise à jour, ce qui va poser un') + #13 +
                 TraduireMemoire('lors de la mise à jour du pointage sur les écritures comptables.') + #13#13 +
                 TraduireMemoire('Pour contourner le problème, veuillez rafraîchir l''écran sans valider le pointage.'), Ecran.Caption);

    end;
  end
  else begin
    AfficheMessageErreur;
    Exit;
  end;

  {$ELSE}
  if not TPieceCompta(TOBPiece).Save then Exit;
  {$ENDIF TRESO}

                 if EstPointageSurTreso then aTob := ConvertitTob(TobTreso, p)
  else if ctxTreso in V_PGI.PGIContexte then aTob := ConvertitTob(TOBPiece.Detail[0], p)
                                        else aTob := ConvertitTob(TOBPiece, p);
  {02/10/07 : ajout d'une précaution}
  if not Assigned(aTob) then Exit;
  try
    k := 0;
    n := 0;
    if aTob.Detail.Count > 0 then begin
      {27/07/07 : suite aux modifications de ce jour dans IsTobAPointe et ConvertitTob, il est possible
                  d'ajouter des flux de Tréso non pointés}
      PtOk := not (EstPointageSurTreso or not (FBoBordereau or VH^.PointageJal)) or
                  (aTob.Detail[n].GetString('CLE_POINTE') <> '');

      {Mise à jour de la grille}
      if not FPtgeManuel and PtOk then CocheDecoche(FListe.Row, p);
      {Ajout de l'écriture qui vient d'être créé}
      for n := aTob.Detail.Count - 1 downto 0 do begin
        F := aTob.Detail[n];
        //if V_PGI.SAV then TobDebug(aTob);
        {Ajout de l'écriture qui vient d'être créé}
        FListe.InsertRow(FListe.Row + 1 + k);
        {Rattache la tob en cours à la tob du relevé}
        F.ChangeParent(TobReleve, -1);
        {Mise à jour de la grille}
        RemplitLigne(FListe.Row + 1 + k, F);
        {Mise à jour du pointage}
        if PtOk then CocheDecoche(FListe.Row + 1 + k, p);
        Inc(k);

        {Mise à jour des soldes sur écritures, pour éviter de relancer les requêtes}
        if VH^.PointageJal or EstPointageSurTreso then begin
          {En pointage sur journal, F pointe sur le compte lettrable alors que les soldes concernent le compte de
           banque => il faut donc inverser le sens}

          if FDevise = V_PGI.DevisePivot then begin
            FSoldeEcrDeb := FSoldeEcrDeb + Arrondi(F.GetDouble('CLE_CREDIT'), FNbDecimales);
            FSoldeEcrCre := FSoldeEcrCre + Arrondi(F.GetDouble('CLE_DEBIT'), FNbDecimales);
          end else begin
            FSoldeEcrDeb := FSoldeEcrDeb + Arrondi(F.GetDouble('CLE_CREDITDEV'), FNbDecimales);
            FSoldeEcrCre := FSoldeEcrCre + Arrondi(F.GetDouble('CLE_DEBITDEV'), FNbDecimales);
          end;
        end
        else begin
          if FDevise = V_PGI.DevisePivot then begin
            FSoldeEcrDeb := FSoldeEcrDeb + Arrondi(F.GetDouble('CLE_DEBIT'), FNbDecimales);
            FSoldeEcrCre := FSoldeEcrCre + Arrondi(F.GetDouble('CLE_CREDIT'), FNbDecimales);
          end else begin
            FSoldeEcrDeb := FSoldeEcrDeb + Arrondi(F.GetDouble('CLE_DEBITDEV'), FNbDecimales);
            FSoldeEcrCre := FSoldeEcrCre + Arrondi(F.GetDouble('CLE_CREDITDEV'), FNbDecimales);
          end;
        end;

      end;
    end;
  finally
    if Assigned(aTob) then FreeAndNil(aTob);
  end;

  {Mise à jour des totaux}
  MajTotalisation;

  Result := True;
end;

{23/07/07 : Mise à jour de l'image de l'état de pointage #ETAT#
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.MajImage;
{---------------------------------------------------------------------------------------}
begin
  if FPtgeManuel then begin
    SetControlVisible('IVERT' , False);
    SetControlVisible('IROUGE', False);
  end
  else begin
    SetControlVisible('IVERT' , FEtatPointage = 0);
    SetControlVisible('IROUGE', FEtatPointage <> 0);
  end;
end;

{$IFDEF TRESO}
{---------------------------------------------------------------------------------------}
procedure TOF_CPPOINTAGERAP.GenereEcriture;
{---------------------------------------------------------------------------------------}
var
  Cle : string;
  str : string;
  Ok  : Boolean;
  T   : TOB;
  n   : Integer;
  d   : Double;
  Q   : TQuery;
  TobCompta : TobPieceCompta;

    {-------------------------------------------------------------}
    function _GetCle : string;
    {-------------------------------------------------------------}
    begin
      Result := '';
      if EstPointageSurTreso then
        Result := FGeneral
      else begin
        Q := OpenSQL('SELECT BQ_CODE FROM BANQUECP WHERE BQ_GENERAL = "' + FGeneral +
                     '" AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '"', True);
        if not Q.Eof then Result := Q.FindField('BQ_CODE').AsString;
        Ferme(Q);
      end;
    end;

begin //OK@@
  n := FListe.Row;
  T := TOB(FListe.Objects[0, n]);
  if T.GetString('CLE_JOURNAL') <> LIGNEBANCAIRE then begin
    str :=  _GetCle + ';';
  end
  else begin
    {Constitution de la chaine de paramètres : Compte, "RAPPRO", Date Opération, Libellé, Montant, Date valeur}
    Cle := _GetCle;
    {JP 22/03/04 : C'est le sens du flux saisi qui détermine le signe de l'écriture de trésorerie}
    d := Abs(T.GetDouble('CLE_MONTANTDEV') + T.GetDouble('CLE_MONTANT'));
    str := Cle + ';RAPPRO;' + T.GetString('CLE_DATECOMPTABLE') + ';' + T.GetString('CLE_LIBELLE') +
                        ';' + FloatToStr(d) + ';' + T.GetString('CLE_DATEVALEUR') + ';' +
                        T.GetString('CLE_CIB') + ';';
  end;
  
  {Le retour est la clé en opération de l'écriture de trésorerie créée}
  T := TrSaisieFluxTob(str);
  {En cas d'abandon ... on sort}
  if not Assigned(T) then Exit;

  Ok := True;

  PutValueWholeTob(T, 'TE_NATURE', na_Realise);
  {Récupération des références de pointage}
  try

    {JP 23/04/04 : Si l'utilisateur n'a pas accès à l'écran d'intégration automatique, on ne l'autorise
                   pas à intégrer automatiquement les écritures en compta}
    if GetParamSocSecur('SO_TRINTEGAUTO', False) and AutoriseFonction(dac_Integration) and
       (PgiAsk(TraduireMemoire('Voulez-vous intégrer en comptabilité l''écriture créée ?'), Ecran.Caption) = mrYes) then begin
      {18/09/06 : Nouvelle gestion Multi-sociétés}
      TobCompta := TobPieceCompta.Create('***', Nil, -1);
      {S'il n'y a pas d'écritures pour la transaction !!?}
      if T.Detail.Count = 0 then Exit;
      {18/09/06 : Nouvelle intégration des écritures en comptabilité (gestion du multi-sociétés)}
      Ok := TRGenererPieceCompta(TobCompta, T);
      {15/10/07 : Affichage du IsValidePiece}
      if not Ok or (TobCompta.Detail.Count = 0) then begin
        AfficheMessageErreur;
        Ok := False;
      end;
    end;

    {Mise à jour des références de pointage et insertion en base}
    if Ok then MajRefPointage(TobCompta, T)

  finally
    if Assigned(T) then FreeAndNil(T);
    if Assigned(TobCompta) then FreeAndNil(TobCompta);
  end;
end;
{$ENDIF TRESO}

initialization
  RegisterClasses([TOF_CPPOINTAGERAP]);

end.

