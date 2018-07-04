{***********UNITE*************************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 09/01/2006
Modifié le ... : 09/01/2006
Description .. : Nouveau Source TOF de la FICHE : AMLISTE ()
Suite ........ : Ancien source TOF -> AMLISTE2_TOF
Suite ........ : FQ 17356 Sortie/Dépréciation pouvaient être non visibles avec CTL+A car
Suite ........ : on récupérait toute la liste d'immos au lieu des immos de la sélection en cours
Suite ........ : FQ 17357 Erreur en conversion de type variant si F11 sur liste vide
Suite ........ : FQ 17358 CTL+Suppr opérationnel mais bouton [Supprimer] non visible
Suite ........ : FQ 17358 Erreur en conversion de type variant si CTL+Suppr opérationnel mais sur liste vide
Suite ........ : FQ 17361 ALT+F3 création d'immos en série est une création incomplète
Suite ........ : FQ 17395 Changement de regroupement ne fonctionne pas sur une sélection par Ctl+A
Suite ........ : FQ 17427 Bouton d'informations de regroupement
Suite ........ : FQ 17516 Le changement de regroupement devient une opération ('REG')
Suite ........ : FQ 17518 Boucler sur l'éclatement en série des immos sélectionnées
Suite ........ : FQ 17826 Eclatement ne doit pas être ouvert sur une acquisition
Suite ........ : FQ 17451 Boucler sur la révision plan, la saisie d'exceptionnel, la modif bases en série des immos sélectionnées
Suite ........ : FQ 17881 En multi-sélection, on pouvait muter une immo acquise dans l'exercice
Suite......... : FQ 17923 Nouvelle opération : modif date de mise en service
Suite......... : FQ 17257 05/06 Filtre dérogatoire devient <=> présence plan fiscal
Suite......... : FQ 17881 05/06 Tests erronés pour ouvrir l'opération en série de Mutation
Suite......... : FQ 18119 05/06 Positionner les indicateurs de modif compta du PARAMSOC
Suite......... : FQ 17499 05/06 En suppression d'un CB, sur la levée de l'option, dégager le lien avec l'immo supprimée
Suite......... : FQ 18385 06/06 Fermer la saisie d'exceptionnel si date début amortissement ECO > date fin exo en cours
Suite......... : FQ 18386 06/06 Fermer la saisie de dépréciation si date début amortissement ECO > date fin exo en cours
Suite......... : FQ 18414 06/06 Plantage en création filtre ou modif exercice dans le multicritère
//               cause nom du champ i_dateamort modifié
Suite......... : FQ 18393 06/06 En série, reprendre la date de l'opération
Suite......... : FQ 18435 06/06 Eclatement unitaire et en série, ne pas ouvrir l'éclatement si présence d'exceptionnel
//               Se voit en série : on éclate une immo, on valide l'opération, on a un message l'immo ne peut être éclatée etc.
//               mais l'immo est bien éclatée, car le message s'applique à l'immo suivante
Suite......... : FQ 18394 06/06 En série, reprendre le no de serie pour qu'en annulation on puisse annuler les autres de la série
//               06/06 Suite FQ 18435 I_MONTANTEXC
Suite......... : FQ 14133 06/06 Ajout critère Exigibilité au dégrèvement dans l'onglet Recherche
Suite......... : 06/06 Gestion des DPI : nouveau critère  'Utilisation d'une DPI ds l'onglet Opérations
Suite......... : 06/06 Gestion des DPI : appel de la liste immos par le suivi des DPI
Suite......... : FQ 14324 07/06/06 - mbo - modif suppression immo : rÚactivation du test sur immo ayant subi une cloture
                                         + modif message pour harmonisation avec suppression fiche
Suite......... : FQ 18465 06/06 - BTY - Remplacer la liste d'exos de la combo FExercice2 par des exos relatifs  Encours, Précédent, N-2 etc.
Suite......... :          De plus . Restauration de filtre avec exo relatif et dates amortissement/achat, une correction
Suite......... :          est faite pour reprendre si possible les JJ MM des dates du filtre
Suite......... :                  . Restauration d'ancien filtre avec exo réel ne fonctionne pas => l'utilisateur devra le réenregistrer
// TGA - 04/07/2006 suppression des enregs de immomvtd
Suite......... : 07/06 Appel de la liste immos par le suivi des DPI, ôter la contrainte sur la date de mise en service datedebeco
Suite......... : FQ 18465 07/06 Pour SERIE1
Suite......... : MVG 12/07/2006 modif pour SERIE1
//XVI 10/01/2007 FQ (BL) 13735 suppression en cascade des ventilations analytiques (pour BL)
Suite......... : 07/06 Appel par le suivi des DPI, ôter la contrainte de réintégration + nouvx paramètres DPI
Suite......... : 07/06 Bloquer modif de dates si l'immo utilise une DPI
Suite......... : 09/06 MBO Nouvelle opération : saisie d'une prime d'équipement
Suite......... : 09/06 Appel par le suivi des DPI, bloquer les opérations
Suite......... : 09/06 FQ 16149 Appliquer les restrictions utilisateur au critère de sélection Etablissements
Suite......... : et en restauration de filtre
Suite......... : mbo - 22.09.06 - ajout ds popup plan : option suramortissement = plan pour prime d'équipement
Suite......... : 10/06 Nouveau critère Prime d'équipement
Suite......... : 10/06 MBO Nouvelle opération : saisie d'une subvention d'équipement
Suite......... : 10/06 BTY Nouvelle opération : réduction de la prime d'équipement
Suite......... : 10/06 FQ 18957 En retour d'une suppression d'immo, l'icône Opérations reste enable alors qu'aucune ligne n'est sélectionnée
Suite......... : 10/06 Nouveau critère Subvention d'investissement
Suite......... : 10/06 FQ 18994 Ne pas autoriser l'affectation DPI sur une immo sortie
Suite......... : 10/06 FQ 19002 Ne pas autoriser l'affectation DPI sur une immo de date de début amort. éco > exo en cours
Suite......... : et sur une immo de type simulation
Suite......... : 10/06 Ne pas saisir de prime ou subvention sur une immo de type simulation
Suite......... : 10/06 BTY FQ 19015 Pas de modif base sur une immo non amortissable
Suite......... : 11/06 BTY Nouvelle opération : réduction de la subvention
Suite......... : F 19070 - MBO - 21/11/2006 - pas menu visualisation écritures sur les biens SIMU et les immos financières
Suite......... : FQ 19260 -BTY- 12/06 Interdire la révision de plan à une immo NAM avec prime ou subvention
Suite......... : FQ 19846 - MBO - 16/03/07 - création du lieu géographique par défaut
Suite......... : FQ 17512 - MBO - 16/03/07 - ajout check gestion fiscale
                                           - interdire chgt méthode si gestion fiscale
                                           - interdire prime si gestion fiscale
Suite......... : MBO - 16/04/2007 - ajout option remplacement de composant
Suite......... : MBO - 26/04/2007 - pas de duplication d'un composant
Suite......... : FQ 20460 BTY 06/07 - Opération Réduction subvention visible à tort, à cause de la longueur de l'item dans AMLISTE (20 caractères)
Suite......... :                      'ReductionSubvention' -> 'ReductionSBV'
Suite......... : BTY - 06/07 - Top du composant remplaçant = I_STRING1 au lieu de I_REMPLACE
Suite......... : BTY - 06/07 - Message 'ERREUR: le champ I_LEVEEOPTION n'existe pas' vu en mode utilisateur SAV sur une sélection d'immos => I_OPELEVEEOPTION
Suite......... : BTY - 07/07 - Fonctions de gestion des exos relatifs rendues publiques dans Outils
Suite......... : MBO - 30/08/2007 - autoriser la suppression d'une immo remplaçante issue d'un import
Suite......... : FQ 14429 - MBO - 04/09/2007 - suppress° immo issue d'un éclatement ou autre opér° : préciser immo d'origine ds message
Suite......... : BTY - 15/10/07 Rendre CreationLieuParDefaut accessible par les pgms
Mots clefs ... : TOF;AMLISTE
*****************************************************************}

unit AMLISTE_TOF;

interface

uses StdCtrls,
  Controls,
  Classes,
  Menus,
  Graphics,
  Windows,
  uTob,
  HQry,
  dialogs,
 {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  mul,
  fe_main,
  HDB,
{$ELSE}
  eMul,
  MaineAgl,
{$ENDIF}
{$IFDEF VER150}
   Variants,
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  { AGL }
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  paramsoc,
  AGLInit,
  HStatus,
  HTB97,
  LookUp,
  { AMORTISSEMENT}
  ImSortie,
  ImEclate,
  ImMutati,
  ImLevOpt,
  ChanPlan,
  OpEnCour,
  OpeSerie,
  ImPlan,
  ImEnt,
  ImContra,
  ListeEch,
  PlanAmor,
{$IFNDEF CMPGIS35}
  PlanPrime,
  PlanSbv,
{$ENDIF}
  ImDupImo,
  ImModBas,
  IntegEcr,
  ImOutGen,
  AMImmo,
  ImCreSer,
  IMMO_TOM
{$IFDEF SERIE1}
  , S1Util, utModules
{$ELSE}
  ,Ent1
{$ENDIF}

  ,AMDEPRECIATION_TOF
  ,AMDEPRECGROUPE_TOF
  ,AMCHGTMETHODE_TOF
  ,AMMODIFSERVICE_TOF
  ,AMREGROUPEMENT_TOF
{$IFNDEF CMPGIS35}
  ,AMSYNTHESEDPI_TOF
  ,AMPRIME_TOF
  ,AMREDUCPRIME_TOF
  ,AMSUBVENTION_TOF
  ,AMREDUCSBV_TOF
  ,AMREMPLACEMENT_TOF
  ,AMDPI
{$ENDIF}
  ,Outils ;

type
  TOF_AMLISTE = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
  private
    { Contrôle }
    POPZoomAction: TPopupMenu;
    POPZoomCreation: TPopupMenu;
    POPZoomVisu: TPopupMenu;
    POPZoomRegroupement: TPopupMenu;

   {$IFDEF EAGLCLIENT}
    FListe: THGrid;
   {$ELSE}
    FListe: THDBGrid;
   {$ENDIF}

    { Variables }
    fAction: TActionFiche;
    fbHistorique: boolean;
    fRegroupement: string;
    fCodeImmo: string;
    fWhereDerog: string;
    fWhereGestionFiscale : string; // fq 17512 chantier fiscal
    fWhereOrig: string;
    fWhereAmortDiffere: string;
    fWhereBaseEcoTh: string;
    fWhereSansGroupe: string;
    fWhereSansOpe: string;
    fWhereOpe: string;
    fWhereBaseTPTh: string;
    fWhereReintegration : string;
    fWhereQuotePart : string;
    // 7.10 06/06 Appel Liste immos par le Suivi DPI
    fWhereCheckDPI : string;
    fDateDebExoDPI : TDateTime;
    fDateFinExoDPI : TDateTime;
    fMontantMaxiDPI : double;
{$IFNDEF CMPGIS35}
    fAmDPI : TAmDPI;
{$ENDIF}
    //
    // FQ 14133
    fWhereExigibleTP: string;
    // 10/06 Prime d'équipement
    fWhereCheckPRI : string;
    // 10/06 Subvention d'investissement
    fWhereCheckSBV : string;

    fOnSaveKeyDownEcran: procedure(Sender: TObject; var Key: Word; Shift:
      TShiftState) of object;

    { Méthodes }
    procedure OnFExercice2Change(Sender: TObject);
    procedure OnAmortDiffereClick(Sender: TObject);
    procedure OnAmortDerogClick(Sender: TObject);
    procedure OnGestionFiscaleClick(Sender: TObject);  // fq 17512
    procedure OnBaseEcoThClick(Sender: TObject);
    procedure OnSansGroupeClick(Sender: TObject);
    procedure OnSansOpeClick(Sender: TObject);
    procedure OnOperationChange (Sender: TObject);
    procedure OnReintegrationClick(Sender: TObject);
    procedure OnQuotePartClick(Sender: TObject);
    procedure OnBaseTPThClick(Sender: TObject);
    // 7.10 06/06 DPI Nouveau critère
    procedure OnDPIClick(Sender: TObject);
    // 10/06 Critère Prime d'équipement
    procedure OnPRIClick(Sender: TObject);
    // 10/06 Critère Subvention d'investissement
    procedure OnbSBVClick(Sender: TObject);
    // FQ 14133
    procedure OnExigibleTPClick(Sender: TObject);
    // FQ 18465 Critère exercice devient exercice relatif
{ 07/07 Bloc déplacé dans Outils
    function  AMPositionEncoursVHImmo : integer ;
    function  AMCodeRelatifExo (PositionExo : integer; PositionEncours : integer) : string ;
    function  AMExoRelatifToExercice (CodeExoRelatif : string) : string ;
    procedure AMInitComboExercice (ComBo : THValComboBox );
    procedure AMAppliquerExoRelatifToDates (CodeExoRelatif : string;
              TDateDebut, TDateFin : TControl ; bFiltre : boolean = False ) ; }

    // FQ 18465 En restauration de filtre, tenir compte des exercices relatifs
    procedure OnChangeFFiltres(Sender : TObject);
    //

    procedure RecupereCritereXX_WHERE(Sender: TObject);
//    procedure OnBChercheClick(Sender: TObject);
    procedure RafraichirListeImmo(stCode: string);
    procedure OnListeDblClick(Sender: TObject);
    procedure OnRowEnterListe(Sender: TObject); overload ;
    procedure OnRowEnterListe(stCode: string); overload ;
    procedure OnCreationPROClick(Sender: TObject);
    procedure OnCreationFIClick(Sender: TObject);
    procedure OnCreationCBClick(Sender: TObject);
    procedure OnCreationLOCClick(Sender: TObject);
    procedure OnCreationSerieClick(Sender: TObject);
    procedure OnDupliqueImmoClick(Sender: TObject);
    procedure OnCessionClick(Sender: TObject);
    procedure OnMutationClick(Sender: TObject);
    procedure OnEclatementClick(Sender: TObject);
    procedure OnLeveeOptionClick(Sender: TObject);
    procedure OnDepreciationActifClick(Sender: TObject);
    procedure ChangementMethodeClick(Sender: TObject);
    procedure ModifServiceClick(Sender: TObject);
    procedure OnPrimeClick(Sender: TObject);   // ajout mbo pour saisie prime d'équipement
    procedure OnSaisieSBVClick(Sender: TObject);   // ajout mbo pour saisie subvention d'équipement
    procedure OnReductionPrimeClick(Sender: TObject);  // 10/06 Réduction de la prime d'équipement
    procedure OnReductionSubventionClick(Sender: TObject);  // 11/06 Réduction de la subvention
    procedure OnSaisieRemplacementClick(Sender: TObject);   // ajout mbo pour remplacement composant
    procedure OnAnnulationOperationClick(Sender: TObject);
    procedure OnDureeMethodeClick(Sender: TObject);
    procedure OnElementExceptionnelClick(Sender: TObject);
    procedure OnChangeEtablissementClick(Sender: TObject);
    procedure OnChangeLieuClick(Sender: TObject);
    procedure OnModifBasesClick(Sender: TObject);
    procedure OnActionRegroupementClick(Sender: TObject);
    procedure OnPopZoomActionPopup(Sender: TObject);
    procedure OnPopZoomCreationPopup(Sender: TObject);
    procedure OnPopZoomVisuPopup(Sender: TObject);
    procedure OnPopZoomRegroupementPopup(Sender: TObject);
    procedure OnZoomPlanClick(Sender: TObject);
    procedure OnZoomPrimeClick(Sender: TObject);       // ajout mbo pour suramortissement prime
    procedure OnZoomSBVClick(Sender: TObject);       // ajout mbo pour suramortissement Subvention
    procedure OnZoomEcheanceClick(Sender: TObject);
    procedure OnZoomOperationsClick(Sender: TObject);
    procedure OnZoomEcrituresClick(Sender: TObject);
    procedure OnSupprimeClick(Sender: TObject);
    procedure OnReinitialiseClick(Sender: TObject);
    procedure OnPopupListeClick(Sender: TObject);
    procedure OnZoomInfoRegClick(Sender: TObject);
    procedure OnZoomDetailRegClick(Sender: TObject);
    // 04/06 FQ 17451
    procedure OperationSerieEclatRevExcMba (TypeOpe:TypeOperation);
    // 06/06 FQ 18393
    //function  ExecuteOperationEclatRevExcMba (TypeOpe:TypeOperation; CodeI:string):integer;
    function  ExecuteOperationEclatRevExcMba (TypeOpe:TypeOperation; CodeI:string; var Param:string):integer;
    // 04/06 FQ 17518
//    function  ImmosAEclater (stCode:string; TListeAll:Tob; var EnCoursEclatable:boolean): boolean;
    // 04/06 FQ 17451
    function  ImmosATraiter (stCode:string; TListeAll:Tob; TypeOpe:TypeOperation; var EnCoursATraiter:boolean): boolean;
    procedure EnableZoomAction;
    procedure EnableZoomVisu;
    procedure EnableZoomRegroupement;
    procedure EnableZoomActionSerie;
    procedure EnableZoomVisuSerie;
    procedure EnableZoomRegroupementSerie;
    // 09/06 Bloquer les opérations en perspective DPI
    procedure EnableZoomActionDPI;
    procedure EnableZoomAction1DPI;
    procedure EnableBoutonsAction(Action: TActionFiche);
    procedure RechargeListeChoix;
    procedure ExecuteOperationSerie(TypeOpe: TypeOperation);
    function  ImmosAMuterChanLieuEtab (stCode:string;stNature:string; TListeAll:Tob; TypeOpe:TypeOperation; var DateDerniereOpe:TDateTime): boolean;
    procedure SuppressionFicheImmo;
    procedure OnCompteelipsisClick(Sender: TObject);
    procedure OnOrganismeelipsisClick(Sender: TObject);
    procedure OnChangeChampLibre(Sender: TObject);
    procedure OnKeyDownFListe(Sender: TObject; var Key: Word; Shift:
      TShiftState);
    procedure OnKeyDownEcran(Sender: TObject; var Key: Word; Shift:
      TShiftState);
    function ListeVide: boolean;
    function SelectionUnique: boolean;
    function SelectionMultiple: boolean;
    procedure SelectionEnregistrement(stCode: string);
    procedure AfterShow;
    procedure OnFlipSelectionGrille (Sender : TObject);
    procedure OnSelectAllClick ( Sender : TObject );
    procedure RecupereAllSelect ( T : TOB );
    procedure AfficheOpeEnCours(Sender: TObject); overload;
    procedure AfficheOpeEnCours(stCode: string); overload;
    procedure ApresChangeFFiltres;  // FQ 16149

  end;

procedure AMLanceFiche_ListeDesImmobilisations(Compte: string; bHistorique:
  boolean = FALSE; Action: TActionFiche = taModif; Regroupement: string = '';
  ParamDPI : string = ''); // 7.10 06/06 Appel liste par le Suivi DPI
procedure SupprimeFicheImmo(Code: string; bFiche : boolean = True );
function JaiLeDroitSupprimerImmo ( CodeImmo, ChangeCode : string ) : boolean;
function ConfirmationSuppressionImmo ( CodeImmo, EtatImmo : string ) : boolean;
function ValideSuppressionImmo ( stCode : string; Appel : boolean ) : boolean;
function  RegroupementDejaDeprecie ( stGroupe : string) : boolean;
function  RegroupementADeprecier ( stCode,stGroupe : string) : boolean;
procedure CreationLieuParDefaut; // BTY Rendue publique 15/10/07


implementation

const
  HM: array[0..10] of string =
{0} {0}  ('Liste des immobilisations',
{1} {9}   'Aucune immobilisation sélectionnée',
{2} {13}  'Des opérations ont été effectuées sur cette immobilisation. Duplication de l''immobilisation impossible.',
{3} {15}  'Attention : cette immobilisation a été créée sur un exercice clos.#10#13La détruire va entraîner une modification de l''historique.#10#13Voulez-vous continuer ?',
{4} {18}  'Il n''existe pas de modèle pour ce regroupement.',
{5}       'Certaines immobilisations de la sélection ne peuvent pas être traitées.'+ #13 + #10 +
          ' L''opération portera uniquement sur l''immobilisation en cours si celle-ci l''autorise.',
{6}       'Il n''est pas possible de pratiquer plusieurs dépréciations/reprises dans l''exercice.',
{7}       'L''immobilisation est sortie.',
{8}       'Il n''est pas possible de pratiquer une dépréciation ' +
              'si l''immobilisation n''est pas amortie sur l''exercice en cours.',
{9}       'Les immobilisations de la sélection ne sont pas toutes de même nature.'+ #13 + #10 +
          ' L''opération portera uniquement sur l''immobilisation en cours.',
{10}      'Cette immobilisation est issue d''un remplacement de composant. Duplication de l''immobilisation impossible.'
  );
  HM2: array[0..11] of string =
  ('En cours :',
    'Mutation',
    'Eclatement',
    'Sortie',
    'Modification des bases',
    'Changement de lieu',
    'Changement d''établissement',
    'Changement de plan',
    'Levée de l''option',
    'Dépréciation de l''actif',
    'Changement de regroupement',
    'Remplacement de composant'
    );

procedure CreationLieuParDefaut;
var
  Q: TQuery;
  T: TOB;
begin
  //AvertirTable('TILIEUGEO');   // AJOUT MBO POUR FQ 19846

  Q := OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="GEO"', True);
  // FQ 19846 mbo 16.03.07 if (Q.Eof) and (VHImmo^.EtablisDefaut <> '') then
  if (Q.Eof) then
  begin
    T := TOB.Create('CHOIXCOD', nil, -1);
    T.PutValue('CC_TYPE', 'GEO');
    T.PutValue('CC_CODE', '001');
    T.PutValue('CC_LIBELLE', 'A définir');
    T.PutValue('CC_ABREGE', 'A définir');
    {mbo 16.03.07 6 FQ 19846
      T.PutValue('CC_CODE', VHImmo^.EtablisDefaut);
      T.PutValue('CC_LIBELLE', RechDom('TTETABLISSEMENT', VHImmo^.EtablisDefaut,False));
      T.PutValue('CC_ABREGE', RechDom('TTETABLISSEMENT', VHImmo^.EtablisDefaut, True));
    }
    T.InsertDB(nil);
    T.Free;
    AvertirTable('TILIEUGEO');
  end;
  Ferme(Q);
end;

procedure AMLanceFiche_ListeDesImmobilisations(Compte: string; bHistorique:
  boolean = FALSE; Action: TActionFiche = taModif; Regroupement: string = '';
  ParamDPI : string = ''); // 7.10 06/06 DPI
var
  stAction: string;
begin
  if ((GetParamSocSecur('SO_EXOCLOIMMO','') = VHImmo^.Encours.Code) or (bHistorique) or
    (Action = taConsult)) then
    stAction := 'ACTION=CONSULTATION'
  else
    stAction := 'ACTION=MODIFICATION';
  stAction := stAction + ';' + Compte;
  if bHistorique then
    stAction := stAction + ';HISTO'
  else
    stAction := stAction + ';';
  if Regroupement <> '' then
     stAction := stAction + ';' + Regroupement  // 7.10 06/06 DPI;
  // 7.10 06/06 DPI
  else
    stAction := stAction + ';';
  if ParamDPI <> '' then
     stAction := stAction + ';' + ParamDPI;
  //
  CreationLieuParDefaut;
  AGLLanceFiche('AM', 'AMLISTE', '', '', stAction);
end;


procedure TOF_AMLISTE.OnNew;
begin
  inherited;
end;

procedure TOF_AMLISTE.OnDelete;
begin
  inherited;
end;

procedure TOF_AMLISTE.OnUpdate;
begin
  inherited;
  // Requête a été faite sur le ONLOAD => Récup valeur initiale du critère complémentaire
  SetControlText('XX_WHERE', fWhereOrig);
  // Revoir boutons au cas où liste non vide
  EnableBoutonsAction(fAction);
  // Affichage en bas de l'écran des opérations sur l'immo courante
  AfficheOpeEnCours (nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : EVENEMENT ONLOAD
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnLoad;
begin
  inherited;

  TFMul(Ecran).Q.Manuel := True;
  if fAction = taConsult then
    begin
     TFMul(Ecran).SETDBLISTE ('AMLISTEHISTO');
     TFMul(Ecran).FNomFiltre := 'AMLISTEHISTO';
    end
  else
    begin
     TFMul(Ecran).SETDBLISTE ('AMLISTEIMMO');
     TFMul(Ecran).FNomFiltre := 'AMLISTEIMMO';
    end;
  TFMul(Ecran).Q.Manuel := False; // pour avoir l'exécution automatique de la requête

  // Appliquer les critères retenus dans chaque onglet du MUL
  RecupereCritereXX_WHERE (nil);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : EVENEMENT ONARGUMENT
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnArgument(S: string);
var
  St: string;
  okok: boolean;
  i: integer;
begin
  inherited;

////////////////////////////////////////////////////////////////////////////
//       Découpage paramètre en entrée
////////////////////////////////////////////////////////////////////////////

  // Action
  St := ReadTokenSt(S);
  if St = 'ACTION=MODIFICATION' then  fAction := taModif
  else   fAction := taConsult;
  // Compte
  St := ReadTokenSt(S);
  if St <> '' then  SetControlText('I_COMPTEIMMO', St);
  // Historique
  St := ReadTokenSt(S);
  fbHistorique := (St = 'HISTO');
  // Regroupement passé en entrée de la MUL
  // 7.10 06/06 DPI
  //fRegroupement := S;
  fRegroupement := ReadTokenSt(S);
  // Appel par le Suivi DPI => paramètres DPI passés en entrée
  fDateDebExoDPI := iDate1900;
  fDateFinExoDPI := iDate1900;
  fMontantMaxiDPI:= 0.00;
  St := ReadTokenSt(S);
  if St <> '' then fDateDebExoDPI := StrToDate (St);
  St := ReadTokenSt(S);
  if St <> '' then fDateFinExoDPI := StrToDate (St);
  St := ReadTokenSt(S);
  if St <> '' then fMontantMaxiDPI := StrToFloat (St);

  // 07/06 Paramètres DPI : dates 5 exos N-1 à N-5
{$IFNDEF CMPGIS35}
  fAmDPI := TAmDPI.Create;
  St := ReadTokenSt(S);
  for i:= 1 to 5 do
     begin
     if St <> '' then fAmDPI.Exercice[i].Deb := StrToDate (St);
     St := ReadTokenSt(S);
     if St <> '' then fAmDPI.Exercice[i].Fin := StrToDate (St);
     St := ReadTokenSt(S);
     end;
  // 07/06 Paramètres DPI : soldes 5 exos N-1 à N-5
  for i:= 1 to 5 do
     begin
     if St <> '' then fAmDPI.Solde[i] := StrToFloat(St);
     St := ReadTokenSt(S);
     end;
  //
{$ENDIF}

  // // Aide en ligne
{$IFDEF SERIE1}
  if fbHistorique then
    begin
    Ecran.HelpContext := 511500 ;
    Ecran.Caption:='Historique' ;
    UpdateCaption(ecran) ;
    end
  else
    Ecran.HelpContext := 511000;

  if (VS1.TypeProduit = S1BNC) and (not VS1.OKModImmo) then
  begin
    SetControlVisible('ImmoFinanciere', false);
    SetControlVisible('CreditBail', false);
    SetControlVisible('LocationFinanciere', false);
    SetControlVisible('CreationSerie', false);
    SetControlVisible('bZoomVisu', false);
  end;

{$ELSE}

  if fbHistorique then
    Ecran.HelpContext := 2115000
  else if fAction = taConsult then
    Ecran.HelpContext := 2110000
  else
    Ecran.HelpContext := 2111000;
{$ENDIF}


////////////////////////////////////////////////////////////////////////////
//         Initialisations MUL et Fliste
////////////////////////////////////////////////////////////////////////////

{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
  FListe := THGrid(GetControl('FLISTE'));
  FListe.SortEnabled := False;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
  FListe := THDBGrid(GetControl('FLISTE'));
  FListe.OnRowEnter := OnRowEnterListe;
  FListe.SortEnabled := True;
{$ENDIF}

  // FQ 18465 Compléter le changement de filtre de l'ancêtre MUL
  THValComboBox(GetControl('FFILTRES')).OnChange := OnChangeFFiltres;


////////////////////////////////////////////////////////////////////////////
//                 Initialisations des zones
////////////////////////////////////////////////////////////////////////////

  // // Boutons, PopUps
  SetControlVisible('bInsert', not fbHistorique);
  SetControlVisible('bSelectAll', not fbHistorique);
  SetControlVisible('bZoomAction', not fbHistorique);
  SetControlVisible('bZoomVisu', not fbHistorique);
  SetControlVisible('bRegroupement', not fbHistorique);
  SetControlVisible('bDelete', not fbHistorique);
  SetControlVisible('bReinit', False);
  SetControlVisible('POperation', not fbHistorique);
  SetControlVisible('PANELENCOURS', not fbHistorique);
  SetControlEnabled('BZoomAction', False);
  SetControlVisible('bRegroupement', not fbHistorique);
  SetControlEnabled('bRegroupement', not fbHistorique);

  POPZoomCreation := TPopUpMenu(GetControl('POPZOOMCREATION', True));
  POPZoomAction := TPopUpMenu(GetControl('POPZOOMACTION', True));
  POPZoomVisu := TPopUpMenu(GetControl('POPZOOMVISU', True));
  POPZoomRegroupement := TPopUpMenu(GetControl('POPREGROUPEMENT', True));
  EnableBoutonsAction (fAction);

  // // Init Onglet STANDARDS
  // 7.10 06/06 Appel par le Suivi DPI, griser zones
  if fDateDebExoDPI <> iDate1900 then
    begin
    SetControlEnabled('I_NATUREIMMO', False);
    SetControlText('I_NATUREIMMO', 'PRO;') ;
    end;

  // // Init Onglet COMPLEMENTS
  // 7.10 06/06 Appel par le Suivi DPI, griser zones
  {07/06 Pas de contrainte de date if fDateDebExoDPI <> iDate1900 then
    begin
    SetControlEnabled('FExercice2', False);
    SetControlEnabled('HLabel7', False);
    SetControlEnabled('I_DATEDEBECO', False);
    SetControlEnabled('I_DATEDEBECO_', False);
    SetControlEnabled('HLabel8', False);
    SetControlText('I_DATEDEBECO', DateToStr(fDateDebExoDPI));
    SetControlText('I_DATEDEBECO_', DateToStr(fDateFinExoDPI));
    end; }
  //

  // // Init Onglet AMORTISSEMENTS
  SetControlProperty('I_BASEECO', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('I_BASEECO_', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  // Exclure certains items de TIMETHODEIMMO
  SetControlProperty('I_METHODEFISC','Plus',' AND CO_CODE<>"VAR" AND CO_CODE<>"NAM"');
  // 7.10 06/06 Appel par le Suivi DPI, griser zones
  if fDateDebExoDPI <> iDate1900 then
    begin
     SetControlEnabled('I_METHODEECO', False);
     // 07/06 Rendre plus lisible SetControlText ('I_METHODEECO', 'DEG;LIN;VAR;');
     SetControlProperty('I_METHODEECO','Plus', ' AND CO_CODE<>"" AND CO_CODE<>"NAM"');
    end;

  // // Init Onglet OPERATIONS
{$IFDEF SERIE1}
  // Exclure les items de TIOPEAMOR autres que des opérations + changement d'établissement
  SetControlProperty('OPERATION','Plus', ' AND CO_LIBRE="OPE" AND CO_CODE<>"ETA"');
{$ELSE}
  // Exclure les items de TIOPEAMOR autres que des opérations
  SetControlProperty('OPERATION','Plus', ' AND CO_LIBRE="OPE"');
{$ENDIF}
  // 7.10 06/06 Appel par le Suivi DPI, griser certains champs
  if fDateDebExoDPI <> iDate1900 then
  begin
    // 07/06 SetControlEnabled('bReintegration', False);
    SetControlEnabled('bQuotePart', False);
    SetControlProperty('bDPI', 'State', cbUnChecked);
    SetControlEnabled('bDPI', False);
  end;
  //

  // // Init Onglet RECHERCHE
  SetControlProperty('I_MONTANTHT', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('I_MONTANTHT_', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('I_BASETAXEPRO', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('I_BASETAXEPRO_', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));

  // // Init Onglet CARACTERISTIQUES
{$IFDEF SERIE1}
  SetControlVisible('I_ETABLISSEMENT', False);
  SetControlVisible('t_ETABLISSEMENT', False);
{$ENDIF}
  // Code regroupement passé en entrée de la MUL
  if fRegroupement <> '' then
  begin
    SetControlText('I_GROUPEIMMO', fRegroupement);
    SetControlEnabled('I_GROUPEIMMO', False);
    SetControlEnabled('tI_GROUPEIMMO', False);
    SetControlEnabled('BSansGroupe', False);
  end;
  // Gestion restriction sur Etablissement
{$IFNDEF SERIE1}
  // BTY 09/06 FQ 16149
  ImGereEtablissement (THMultiValComboBox(GetControl('I_ETABLISSEMENT')) );
  //PositionneEtabUser( THMultiValComboBox(GetControl('I_ETABLISSEMENT')) ) ;
{$ENDIF}

  // // Init Onglet CHAMPS LIBRES avec S1, onglet TABLES LIBRES sinon
{$IFDEF SERIE1}
  ImLibellesTableLibre(TTabSheet(GetControl('PzLibreS1')), 'TT_TABLELIBREIMMO', '', 'I');
  Okok := false;
{$ELSE}
  ImLibellesTableLibre(TTabSheet(GetControl('PzLibre')), 'TI_TABLE', 'I_TABLE', 'I');
  Okok := true;
{$ENDIF}
  SetControlVisible('PzLibre', Okok);
  SetControlVisible('PzLibreS1', not Okok);



////////////////////////////////////////////////////////////////////////////
//                   Evènements
////////////////////////////////////////////////////////////////////////////

  // // Evènements généraux
  fListe.OnDblClick := OnListeDblClick;
  fListe.OnKeyDown := OnKeyDownFListe;
  fOnSaveKeyDownEcran := Ecran.OnKeyDown;
  Ecran.OnKeyDown := OnKeyDownEcran;
  TFMul(Ecran).OnAfterFormShow := AfterShow;
  fListe.OnFlipSelection := OnFlipSelectionGrille;
  TFMul(Ecran).bSelectAll.OnClick := OnSelectAllClick;

  // // Boutons
  TToolbarButton97(GetControl('bDelete')).OnClick := OnSupprimeClick;
  TToolbarButton97(GetControl('bReinit')).OnClick := OnReinitialiseClick;


  // // EVENTS Onglet STANDARDS
  THEdit(GetControl('I_COMPTEIMMO')).OnelipsisClick := OnCompteelipsisClick;
  THEdit(GetControl('I_COMPTELIE')).OnelipsisClick := OnCompteelipsisClick;

  // // EVENTS Onglet COMPLEMENTS
  TCheckBox(GetControl('bAmortDiffere')).OnClick := OnAmortDiffereClick;
  THValComboBox(GetControl('FExercice2')).OnChange := OnFExercice2Change;
  // FQ 18465 Liste des exercices à convertir en exos relatifs
  AMInitComboExercice(THValComboBox(GetControl('FExercice2')));

  // // EVENTS Onglet AMORTISSEMENTS
  TCheckBox(GetControl('bAmortDerog')).OnClick := OnAmortDerogClick;
  TCheckBox(GetControl('bBaseEcoTh')).OnClick := OnBaseEcoThClick;
  TCheckBox(GetControl('bGestionFiscale')).OnClick := OnGestionFiscaleClick;

  // // EVENTS Onglet OPERATIONS
  THMultiValComboBox(GetControl('Operation')).OnChange := OnOperationChange;
  TCheckBox(GetControl('bSansOpe')).OnClick := OnSansOpeClick;
  TCheckBox(GetControl('bReintegration')).OnClick := OnReintegrationClick;
  TCheckBox(GetControl('bQuotePart')).OnClick := OnQuotePartClick;
  // 7.10 06/06 DPI
  TCheckBox(GetControl('bDPI')).OnClick := OnDPIClick;
  // 10/06 Critère Prime d'équipement
  TCheckBox(GetControl('bPRI')).OnClick := OnPRIClick;
  // 10/06 Critère Subvention d'investissement
  TCheckBox(GetControl('bSBV')).OnClick := OnbSBVClick;

  // // EVENTS Onglet RECHERCHE
  TCheckBox(GetControl('bBaseTPTh')).OnClick := OnBaseTPThClick;
  // FQ 14133
  TCheckBox(GetControl('bExigibleTP')).OnClick := OnExigibleTPClick;

  // // EVENTS Onglet CARACTERISTIQUES
  TCheckBox(GetControl('bSansGroupe')).OnClick := OnSansGroupeClick;
  THEdit(GetControl('I_ORGANISMECB')).OnelipsisClick := OnOrganismeelipsisClick;

  // // EVENTS Onglet TABLES LIBRES
  THValComboBox(GetControl('TABLELIBRE1')).OnChange := OnChangeChampLibre;
  THValComboBox(GetControl('TABLELIBRE2')).OnChange := OnChangeChampLibre;
  THValComboBox(GetControl('TABLELIBRE3')).OnChange := OnChangeChampLibre;

  // // EVENTS Onglet CHAMPS LIBRES
  THValComboBox(GetControl('TABLELIBRE1')).OnChange := OnChangeChampLibre;
  THValComboBox(GetControl('TABLELIBRE2')).OnChange := OnChangeChampLibre;
  THValComboBox(GetControl('TABLELIBRE3')).OnChange := OnChangeChampLibre;

  // // EVENTS Popup Création
  PopZoomCreation.OnPopup := OnPopZoomCreationPopup;
  TMenuItem(GetControl('PleinePropriete')).OnClick := OnCreationPROClick;
  TMenuItem(GetControl('ImmoFinanciere')).OnClick := OnCreationFIClick;
  TMenuItem(GetControl('CreditBail')).OnClick := OnCreationCBClick;
  TMenuItem(GetControl('LocationFinanciere')).OnClick := OnCreationLOCClick;
  TMenuItem(GetControl('CreationSerie')).OnClick := OnCreationSerieClick;
  TMenuItem(GetControl('DupliqueImmo')).OnClick := OnDupliqueImmoClick;

  // // EVENTS Popup Action
  PopZoomAction.OnPopup := OnPopZoomActionPopup;
  TMenuItem(GetControl('Cession')).OnClick := OnCessionClick;
  TMenuItem(GetControl('Mutation')).OnClick := OnMutationClick;
  TMenuItem(GetControl('Eclatement')).OnClick := OnEclatementClick;
  TMenuItem(GetControl('LeveeOption')).OnClick := OnLeveeOptionClick;
  TMenuItem(GetControl('AnnulationOperation')).OnClick :=
    OnAnnulationOperationClick;
  TMenuItem(GetControl('DureeMethode')).OnClick := OnDureeMethodeClick;
  TMenuItem(GetControl('ElementExceptionnel')).OnClick :=
    OnElementExceptionnelClick;
  TMenuItem(GetControl('ChangeEtablissement')).OnClick :=
    OnChangeEtablissementClick;
  TMenuItem(GetControl('ChangeLieu')).OnClick := OnChangeLieuClick;
  TMenuItem(GetControl('ModifBases')).OnClick := OnModifBasesClick;
  TMenuItem(GetControl('ActionRegroupement')).OnClick :=
    OnActionRegroupementClick;
  TMenuItem(GetControl('DepActif')).OnClick := OnDepreciationActifClick;
  TMenuItem(GetControl('ChangementMethode')).OnClick := ChangementMethodeClick;
  TMenuItem(GetControl('ModifService')).OnClick := ModifServiceClick;   // fq 17923 - mbo
  TMenuItem(GetControl('Remplacement')).OnClick := OnSaisieRemplacementClick;   // mbo - 16.4.07
  TMenuItem(GetControl('Prime')).OnClick := OnPrimeClick;      // mbo - 29.08.06
  TMenuItem(GetControl('Subvention')).OnClick := OnSaisieSBVClick;   // mbo - 06.10.06
  TMenuItem(GetControl('ReductionPrime')).OnClick := OnReductionPrimeClick;  // 10/06 Réduction de la prime d'équipement
  TMenuItem(GetControl('ReductionSBV')).OnClick := OnReductionSubventionClick;  // 11/06 Réduction de la subvention


  // // EVENTS Popup Visualisation
  PopZoomVisu.OnPopup := OnPopZoomVisuPopup;
  TMenuItem(GetControl('ZoomPlan')).OnClick := OnZoomPlanClick;
  TMenuItem(GetControl('ZoomEcheance')).OnClick := OnZoomEcheanceClick;
  TMenuItem(GetControl('ZoomOperations')).OnClick := OnZoomOperationsClick;
  TMenuItem(GetControl('Ecritures')).OnClick := OnZoomEcrituresClick;
  // ajout mbo pour prime et SBV
  TMenuItem(GetControl('ZoomPrime')).OnClick := OnZoomPrimeClick;
  TMenuItem(GetControl('ZoomSubvention')).OnClick := OnZoomSBVClick;

  // // EVENTS Popup Regroupement
  PopZoomRegroupement.OnPopup := OnPopZoomRegroupementPopup;
  TMenuItem(GetControl('BINFOREG')).OnClick := OnZoomInfoRegClick;
  TMenuItem(GetControl('BDETAILREG')).OnClick := OnZoomDetailRegClick;

  // // EVENTS Popup Liste générale - F11
  TPopUpMenu(GetControl('PopupListe')).OnPopup := OnPopupListeClick;
  TMenuItem(GetControl('PleinePropriete1')).OnClick := OnCreationPROClick;
  TMenuItem(GetControl('ImmoFinanciere1')).OnClick := OnCreationFIClick;
  TMenuItem(GetControl('CreditBail1')).OnClick := OnCreationCBClick;
  TMenuItem(GetControl('LocationFinanciere1')).OnClick := OnCreationLOCClick;
  TMenuItem(GetControl('CreationSerie1')).OnClick := OnCreationSerieClick;
  TMenuItem(GetControl('DupliqueImmoListe')).OnClick := OnDupliqueImmoClick;
  TMenuItem(GetControl('ZoomPlan1')).OnClick := OnZoomPlanClick;
  // ajout mbo pour prime et SBV
  TMenuItem(GetControl('ZoomPrime1')).OnClick := OnZoomPrimeClick;
  TMenuItem(GetControl('ZoomSubvention1')).OnClick := OnZoomSBVClick;

  TMenuItem(GetControl('ZoomEcheance1')).OnClick := OnZoomEcheanceClick;
  TMenuItem(GetControl('ZoomOperations1')).OnClick := OnZoomOperationsClick;
  {$IFNDEF SERIE1}
  TMenuItem(GetControl('BInfoReg1')).OnClick := OnZoomInfoRegClick;
  {$ENDIF}
  TMenuItem(GetControl('BDetailReg1')).OnClick := OnZoomDetailRegClick;
  TMenuItem(GetControl('Ecritures1')).OnClick := OnZoomEcrituresClick;
  TMenuItem(GetControl('Cession1')).OnClick := OnCessionClick;
  TMenuItem(GetControl('Mutation1')).OnClick := OnMutationClick;
  TMenuItem(GetControl('Eclatement1')).OnClick := OnEclatementClick;
  TMenuItem(GetControl('LeveeOption1')).OnClick := OnLeveeOptionClick;
  TMenuItem(GetControl('AnnulationOperation1')).OnClick :=
    OnAnnulationOperationClick;
  TMenuItem(GetControl('DureeMethode1')).OnClick := OnDureeMethodeClick;
  TMenuItem(GetControl('ElementExceptionnel1')).OnClick :=
    OnElementExceptionnelClick;
  TMenuItem(GetControl('DepActif1')).OnClick := OnDepreciationActifClick;
  TMenuItem(GetControl('ChangementMethode1')).OnClick := ChangementMethodeClick;
  TMenuItem(GetControl('ModifService1')).OnClick := ModifServiceClick;

  TMenuItem(GetControl('Prime1')).OnClick := OnPrimeClick;
  TMenuItem(GetControl('Subvention1')).OnClick := OnSaisieSBVClick;
  TMenuItem(GetControl('Remplacement1')).OnClick := OnSaisieRemplacementClick;
  // 10/06 Réduction de la prime d'équipement
  TMenuItem(GetControl('ReductionPrime1')).OnClick := OnReductionPrimeClick;
  TMenuItem(GetControl('ReductionSBV1')).OnClick := OnReductionSubventionClick;

  TMenuItem(GetControl('ChangeEtablissement1')).OnClick :=
    OnChangeEtablissementClick;
  TMenuItem(GetControl('ChangeLieu1')).OnClick := OnChangeLieuClick;
  TMenuItem(GetControl('ModifBases1')).OnClick := OnModifBasesClick;
  TMenuItem(GetControl('ActionRegroupement1')).OnClick :=
    OnActionRegroupementClick;


  // Fiche appelée par un verbe OLE
  MakeZoomOle(Ecran.Handle);


////////////////////////////////////////////////////////////////////////////
//         Initialisation des critères de sélection
////////////////////////////////////////////////////////////////////////////

  // Critère complémentaire XX_WHERE par défaut
  if fbHistorique then
  begin
    if GetControlText('XX_WHERE') = '' then
      SetControlText('XX_WHERE', 'I_ETAT="FER"')
    else
      SetControlText('XX_WHERE', GetControlText('XX_WHERE') + ' AND I_ETAT="FER"');
  end
  else
  begin
    if GetControlText('XX_WHERE') = '' then
      SetControlText('XX_WHERE', 'I_ETAT<>"FER"')
    else
      SetControlText('XX_WHERE', GetControlText('XX_WHERE') + ' AND I_ETAT<>"FER"');
  end;
  // 7.10 06/06 DPI
  // En appel par le Suivi DPI, ajouter critères spécifiques par défaut
  // FQ 18994 + immo non sortie
  // FQ 19002 + immo de date début amortist éco <= exo en cours + immo non simulation
  // FQ 17512 MBO - pas de dpi si gestion fiscale
  if fDateDebExoDPI <> iDate1900 then
     begin
     SetControlText('XX_WHERE', GetControlText('XX_WHERE') +
                    ' AND I_NATUREIMMO="PRO" AND I_DPI<>"X"' +
                    ' AND I_METHODEECO<>"" AND I_METHODEECO<>"NAM"' +
                    ' AND I_QUOTEPART=0 AND I_QUANTITE <> 0' +
                    ' AND I_DATEDEBECO <= "' + UsDateTime (VHImmo^.Encours.Fin) + '"' +
                    ' AND I_QUALIFIMMO <> "S" AND I_NONDED <> "X"' );
                    // 07/06  AND I_REINTEGRATION=0' );
// 07/06            ' AND I_DATEDEBECO>="' + UsDateTime (fDateDebExoDPI) + '"' +
// 07/06            ' AND I_DATEDEBECO<="' + UsDateTime (fDateFinExoDPI) + '"');


     end;
  //

  // Critères initiaux à ajouter au XX_WHERE
  fWhereOrig := GetControlText('XX_WHERE');
  OnAmortDerogClick(nil);
  OnAmortDiffereClick(nil);
  OnBaseEcoThClick(nil);
  OnSansGroupeClick(nil);
  OnSansOpeClick(nil);
  OnOperationChange(nil);
  OnReintegrationClick(nil);
  OnQuotePartClick(nil);
  OnBaseTPThClick(nil);
  // 7.10 06/06 DPI
  OnDPIClick(nil);
  // FQ 14133
  OnExigibleTPClick(nil);
  // 10/06 Critère Prime d'équipement
  OnPRIClick(nil);
  // 10/06 Critère Subvention d'investissement
  OnbSBVClick(nil);
end;

procedure TOF_AMLISTE.OnClose;
begin
  TheMulQ := nil ; //BTY Erreur ancienne FQ 15508 de LGendreau
{$IFNDEF CMPGIS35}
  if fAmDPI <> nil then // 07/06
     fAmDPI.Free;
{$ENDIF}
  inherited;
end;

procedure TOF_AMLISTE.OnDisplay();
begin
  inherited;
  EnableBoutonsAction(fAction);
  if (not ListeVide) then
      OnRowEnterListe(nil);
  { Table libre 9 utilisée pour le stockage 'éligible TP' }
  SetControlVisible ('I_TABLE9',False);
  SetControlVisible ('TI_TABLE9',False);
end;

procedure TOF_AMLISTE.OnCancel();
begin
  inherited;
end;

procedure TOF_AMLISTE.AfterShow;
begin
  OnRowEnterListe(nil);
end;

procedure TOF_AMLISTE.EnableBoutonsAction(Action: TActionFiche);
begin
  SetControlVisible('BZoomAction', (Action <> taConsult) and (not ListeVide));
  SetControlVisible('BREINIT', False );
  SetControlVisible('BInsert', (Action <> taConsult));
  SetControlEnabled('BRegroupement', (GetControlVisible('BRegroupement') and (not ListeVide)));
  //{$IFDEF SERIE1}
//  SetControlVisible('BDelete', (Action <> taConsult) and (not ListeVide));
//{$ELSE}
//  SetControlVisible('BDelete', (Action <> taConsult) and (V_PGI.LaSerie = S5)
//     and (not ListeVide));
//{$ENDIF}
// BTY 01/06 Réafficher le bouton [Supprimer] comme en version 6
  SetControlVisible('BDelete', (Action <> taConsult) and (not ListeVide));
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 07/09/2006
Modifié le ... :   /  /
Description .. : Derrière le tableau de suivi des DPI (fDateDebExoDPI <> iDate1900)
Suite ........ : bloquer les opérations
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.EnableZoomActionDPI;
begin
  SetControlEnabled('Cession', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('Cession'))));
  SetControlEnabled('Eclatement', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('Eclatement'))));
  SetControlEnabled('Mutation', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('Mutation'))));
  SetControlEnabled('DureeMethode', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('DureeMethode'))));
  SetControlEnabled('ModifService', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('ModifService'))));
  SetControlEnabled('ElementExceptionnel', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('ElementExceptionnel'))));
  SetControlEnabled('LeveeOption', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('LeveeOption'))));
  //SetControlEnabled('AnnulationOperation', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('AnnulationOperation'))));
  SetControlEnabled('ChangeEtablissement', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('ChangeEtablissement'))));
  SetControlEnabled('ChangeLieu', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('ChangeLieu'))));
  SetControlEnabled('ModifBases', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('ModifBases'))));
  SetControlEnabled('ActionRegroupement', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('ActionRegroupement'))));
  SetControlEnabled('DepActif', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('DepActif'))));
  SetControlEnabled('ChangementMethode', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('ChangementMethode'))));
  SetControlEnabled('Prime', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('Prime'))));
  SetControlEnabled('Subvention', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('Subvention'))));
  SetControlEnabled('ReductionPrime', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('ReductionPrime'))));
  SetControlEnabled('ReductionSBV', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('ReductionSBV'))));
  SetControlEnabled('Remplacement', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('Remplacement'))));
end;

procedure TOF_AMLISTE.EnableZoomAction1DPI;
begin
  SetControlEnabled('Cession1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('Cession1'))));
  SetControlEnabled('Eclatement1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('Eclatement1'))));
  SetControlEnabled('Mutation1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('Mutation1'))));
  SetControlEnabled('DureeMethode1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('DureeMethode1'))));
  SetControlEnabled('ModifService1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('ModifService1'))));
  SetControlEnabled('ElementExceptionnel1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('ElementExceptionnel1'))));
  SetControlEnabled('LeveeOption1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('LeveeOption1'))));
  //SetControlEnabled('AnnulationOperation1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('AnnulationOperation1'))));
  SetControlEnabled('ChangeEtablissement1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('ChangeEtablissement1'))));
  SetControlEnabled('ChangeLieu1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('ChangeLieu1'))));
  SetControlEnabled('ModifBases1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('ModifBases1'))));
  SetControlEnabled('ActionRegroupement1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('ActionRegroupement1'))));
  SetControlEnabled('DepActif1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('DepActif1'))));
  SetControlEnabled('ChangementMethode1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('ChangementMethode1'))));
  SetControlEnabled('Prime1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('Prime1'))));
  SetControlEnabled('Subvention1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('Subvention1'))));
  SetControlEnabled('Remplacement1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('Remplacement1'))));
  SetControlEnabled('ReductionPrime1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('ReductionPrime1'))));
  SetControlEnabled('ReductionSBV1', ((fDateDebExoDPI = iDate1900) and (GetControlVisible('ReductionSBV1'))));
end;


procedure TOF_AMLISTE.EnableZoomAction;
var
  Nat: string;
  bDepotGar, bAcquisitionAnnee: boolean;
  bCedee, GestionFiscale, PlanFiscal: boolean;
  bDep, bDatePB, bExcep, bDPIUtilisee, ExistOpe : Boolean;
  bDejaPrime, bDejaSBV, bPrimeEncours, bPrime, ImmoReelle, bSBVEncours, bSBV : Boolean;
  stcode : String;
  bRevisionPlan, PasComposant : boolean;
begin
  bCedee := (GetField('I_QUANTITE') = 0);
  SetControlVisible('DupliqueImmo', not ListeVide);
  bAcquisitionAnnee := (GetField('I_DATEPIECEA') >= VHImmo^.Encours.Deb)
                        and (GetField('I_DATEPIECEA') <= VHImmo^.Encours.Fin);
  SetControlVisible('Cession', false);
  SetControlVisible('Eclatement', false);
  SetControlVisible('Mutation', false);
  SetControlVisible('DureeMethode', false);
  SetControlVisible('ModifService', False);
  SetControlVisible('ElementExceptionnel', false);
  SetControlVisible('LeveeOption', false);
  SetControlVisible('AnnulationOperation', false);
  SetControlVisible('ChangeEtablissement', false);
  SetControlVisible('ChangeLieu', false);
  SetControlVisible('ModifBases', false);
  SetControlVisible('ActionRegroupement', false);
  SetControlVisible('Ecritures1', SelectionUnique);
  SetControlVisible('DepActif', false);
  SetControlVisible('ChangementMethode', False);

  SetControlVisible('Prime', False);
  SetControlVisible('Subvention', False);
  SetControlVisible('ReductionPrime', False);
  SetControlVisible('ReductionSBV', False);
  SetControlVisible('Remplacement', False);

  // BTY FQ 17357 Sortir ici si liste vide
  // sinon Erreur en conversion de type variant sur le GetField('I_IMMO')
  if ListeVide or (not SelectionUnique) then
     exit;

  Stcode := GetField('I_IMMO');

  PlanFiscal := ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+stCode+'" AND I_METHODEFISC <>""');

  GestionFiscale := ExisteSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+stCode+'" AND I_NONDED="X"');

  ImmoReelle :=  ExisteSQL('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+stCode+'" AND  I_QUALIFIMMO="R"');

  if ExisteSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+stCode+'" AND (I_SBVPRI<>0 or I_SBVPRIC <>0)') then
    bDejaPrime:=True
  Else
    bDejaPrime:=False;

  if ExisteSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+stCode+'" AND (I_SBVMT<>0 or I_SBVMTC <>0)') then
    bDejaSBV:=True
  Else
    bDejaSBV:=False;

  // 10/06 Test opération saisie prime d'équipement OU réduction de prime dans l'exo en cours
  bPrimeEncours := ExisteSQL ('SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="' + Stcode +
                           '" AND (IL_TYPEOP = "PRI" OR IL_TYPEOP = "RPR") AND (IL_DATEOP>="' +
                           USDateTime(VHImmo^.Encours.Deb) + '" AND IL_DATEOP<="' +
                           USDateTime(VHImmo^.Encours.Fin) + '")');
  bPrime := ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+stCode+'" AND (I_SBVPRI<>0)');

  // 11/06 Test opération saisie subvention d'investissement OU réduction de subvention dans l'exo en cours
  bSBVEncours := ExisteSQL ('SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="' + Stcode +
                           '" AND (IL_TYPEOP = "SBV" OR IL_TYPEOP = "RSB") AND (IL_DATEOP>="' +
                           USDateTime(VHImmo^.Encours.Deb) + '" AND IL_DATEOP<="' +
                           USDateTime(VHImmo^.Encours.Fin) + '")');
  bSBV := ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+stCode+'" AND (I_SBVMT<>0)');

  // Test si immo déjà dépréciée
  // Pas d'éclatement, de cession, de modification de base, d'exceptionnel
  // de changement de méthode si bdep=True
  Stcode := GetField('I_IMMO');
  if ExisteSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+stCode+'" AND I_REVISIONECO<>0') then
    bdep:=True
  Else
    bdep:=False;

  // FQ 18385 Pas de saisie d'exceptionnel si date début amortissement ECO > date fin exo en cours
  bDatePB := ExisteSQL
   ('SELECT I_IMMO,I_DATEDEBECO FROM IMMO WHERE I_IMMO="'+ stCode + '" AND I_DATEDEBECO>"' +
   UsDateTime (VHImmo^.Encours.Fin) + '"');

  // FQ 17357
  //if ListeVide or (not SelectionUnique) then
  //   exit;
  Nat := GetField('I_NATUREIMMO');

  // FQ 18386 Pas de dépréciation si aucun regroupement et date début amortissement ECO > date fin exo en cours
  // FQ 18386 Pas de dépréciation si aucun regroupement et déjà dépréciée dans l'exercice
  // FQ 18386 Pas de dépréciation si aucun regroupement et immo sortie
  //SetControlVisible('DepActif', Nat = 'PRO');

  SetControlVisible ('DepActif', (( (Nat='PRO') and (GetField('I_GROUPEIMMO')<>'') ) or
                                  ( (Nat='PRO') and (GetField('I_GROUPEIMMO')='') and
                                    (not bDatePB) and (not bDep) and (not bCedee) ))  );

  // 07/06 Pas de modif date si DPI affectée à l'immo
  bDPIUtilisee := ExisteSQL ('SELECT I_DPI FROM IMMO WHERE I_IMMO="'+stCode+'" AND I_DPI = "X"');
  // ajout mbo 28.05.07  pas de modif date sur composant
  // 06/07 i_string1 au lieu de i_remplace
  PasComposant := ExisteSQL ('SELECT I_STRING1 FROM IMMO WHERE I_IMMO="'+stCode+'" AND I_STRING1 = ""');

  SetControlVisible('ModifService', (Nat = 'PRO')
    and (not bdep) and (not bCedee) and (not (GetField('I_OPELEVEEOPTION') = 'X'))
    and (PasComposant)  // mbo 28.05.07
    and (not bDPIUtilisee) );  // BTY 07/06    // fq 17923 mbo

  // 10/06 pas de prime sur une simulation
  SetControlVisible('Prime', (Nat = 'PRO')
    and (not bCedee) and (not (GetField('I_OPELEVEEOPTION') = 'X'))
    and (not bDejaPrime) and (ImmoReelle)
    and (not GestionFiscale));  // FQ 17512 pas de prime si gestion fiscale

  //10/06 pas de subvention sur une simulation
  SetControlVisible('Subvention', (Nat = 'PRO')
    and (not bCedee) and (not (GetField('I_OPELEVEEOPTION') = 'X'))
    and (not bDejaSBV) and (ImmoReelle )
    and (not GestionFiscale));   // fq 17512 pas de subvention si gestion fiscale

  // 10/06 Opération de réduction de prime d'équipement :
  // ouverte si une saisie de prime a été effectuée dans un exo antérieur à l'exo en cours
  SetControlVisible('ReductionPrime', (bPrime and (not bPrimeENcours)));

  // 11/06 Opération de réduction de subvention d'investissement :
  // ouverte si une saisie de subvention a été effectuée dans un exo antérieur à l'exo en cours
  SetControlVisible('ReductionSBV', (bSBV and (not bSBVEncours)));

  SetControlVisible('ChangementMethode', (Nat = 'PRO')
   and (not bAcquisitionAnnee)and not (getField('I_METHODEECO')='VAR') and (not bdep)and (not bCedee)
   and (not (Getfield('I_OPECHANGEPLAN') = 'X'))
   and (not (Getfield('I_OPEMODIFBASES') = 'X'))
   and (not GestionFiscale));  // ajout mbo fq 17512

  // Autorisation Cession
  SetControlVisible('Cession', ((Nat = 'PRO') or (Nat = 'FI')) and
    (not bCedee) and (not bdep) );

  // Autorisation Remplacement de composant
  SetControlVisible('Remplacement', (Nat = 'PRO')
   and (not bCedee) and (not bSBV)
   and (PlanFiscal));


  // Autorisation Eclatement
  //ATTENTION : le champ immolie est utilisé pour identifier l'immo d'origine pour un depot de gar.
  bDepotGar := (Nat = 'FI') and (GetField('I_IMMOLIE') <> '');

  // FA 18435 Pas d'éclatement si immo présente de l'exceptionnel
   bExcep := ExisteSQL
   ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+ stCode + '" AND I_MONTANTEXC<>0');

  // 04/06 FQ 17826
//  SetControlVisible('Eclatement', (Nat <> 'CB') and (Nat <> 'LOC') and
// (not bDepotGar and not((GetField('I_METHODEECO')='VAR') and not bAcquisitionAnnee) )
//  and (not bCedee) and (not (GetField('I_OPELEVEEOPTION') = 'X')) and (not bdep));
  SetControlVisible('Eclatement', (Nat <> 'CB') and (Nat <> 'LOC')  and
                    (not bDepotGar) and (not (GetField('I_METHODEECO')='VAR')) and
                    (not bAcquisitionAnnee) and (not bCedee) and
                    (not (GetField('I_OPELEVEEOPTION') = 'X')) and (not bdep) and (not bExcep));

  // Autorisation Mutation
  SetControlVisible('Mutation', (not bCedee) and (not
    (GetField('I_OPELEVEEOPTION') = 'X')) and (not bAcquisitionAnnee));

  // Autorisation Changement de plan
  // BTY 19260
  bRevisionPlan := True;
  if ((GetField('I_METHODEECO')='NAM') and (bDejaPrime or bDejaSBV)) then
      bRevisionPlan := False;

  SetControlVisible('DureeMethode', (Nat = 'PRO') and (not bAcquisitionAnnee)
    and (not bCedee) and (not (GetField('I_OPECHANGEPLAN') = 'X')) and ((GetField('I_METHODEECO')<>'VAR'))
    and (bRevisionPlan)  );

  // Autorisation exceptionnel - modif mbo le 14.11.2005
  SetControlVisible('ElementExceptionnel', (Nat = 'PRO') and (not bCedee) and
    (not (GetField('I_OPECHANGEPLAN') = 'X')) and (not (GetField('I_METHODEECO')='NAM'))
     and (not bdep) and (not bDatePB));     // FQ 18385  and (not bdep));

  // Autorisation Levée option
  SetControlVisible('LeveeOption', (Nat = 'CB') and (GetField('I_IMMOLIE') = '')
    and (not bCedee) and (not (GetField('I_OPELEVEEOPTION') = 'X')));

  // Opérations ?   mbo ajout du test sur remplacement de composant
  // on annule l'opération de remplacement sur l'immo de remplacement et non l'immo d'origine

  if ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+stCode+'" AND (I_OPEREMPL = "X")') then
     ExistOpe := SupRemplacement(stCode)
  else
     ExistOpe :=ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+stCode+'" AND (I_OPERATION="X") AND (I_OPEREMPL = "-")');

  SetControlVisible('AnnulationOperation', (ExistOpe));
  SetControlVisible('SeparAnnulAction', GetControlVisible('AnnulationOperation'));
  SetControlVisible('ChangeEtablissement', not bCedee and (not
                (GetField('I_OPELEVEEOPTION') = 'X')) and (VHImmo^.EtablisCpta = TRUE));
  SetControlVisible('ChangeLieu', not bCedee and (not
                (GetField('I_OPELEVEEOPTION') = 'X')));

  // Pas de modif base en location longue durée
  // FQ 19015 Pas de modif base sur une immo non amortissable
  SetControlVisible('ModifBases',(Nat <> 'LOC') and (Nat <> 'FI') and (not bCedee) and (not
    bAcquisitionAnnee) and (not (GetField('I_OPELEVEEOPTION') = 'X'))
     and (GetField('I_METHODEECO')<>'VAR') and (not bdep)
     and (GetField('I_METHODEECO')<>'NAM') );
  // Pas de regroupement si immo sortie
  SetControlVisible('ActionRegroupement', not bCedee);

  // BTY 09/06 En perspective DPI, bloquer les opérations
  EnableZoomActionDPI;
end;

procedure TOF_AMLISTE.EnableZoomActionSerie;
var
  i: integer;
  EnabledCession: boolean;
  PremNature: string;
  bDepotGar, bCedee, bAcquisitionAnnee: boolean;
  GroupeImmo : string;
  TListeAll  : TOB;
  NatureSelectee : string;
  DepSelectee, AcquisitionSelectee, bDatePBSelectee, bExcepSelectee : boolean;
  bDejaPrimeSBV, bDejaPrimeSBVSelectee, bRevisionPlan, bRevisionPlanSelectee : boolean;
begin
  // 04/06 FQ 17881 ATTENTION ! récupérer les infos immo APRES avoir fait GotoLeBookmark(0)
  // car à cet endroit l'immo courante n'est pas l'immo pointée par GotoLeBookmark(0)
  //bAcquisitionAnnee := (GetField('I_DATEPIECEA') >= VHImmo^.Encours.Deb)
  //                      and (GetField('I_DATEPIECEA') <= VHImmo^.Encours.Fin);
  //bCedee := (GetField('I_QUANTITE') = 0);
  SetControlVisible('DupliqueImmo', false);
  SetControlVisible('Cession', False);
  SetControlVisible('Mutation', False);
  SetControlVisible('Eclatement', false);
  SetControlVisible('ElementExceptionnel', false);
  SetControlVisible('LeveeOption', false);
  SetControlVisible('AnnulationOperation', false);
  SetControlVisible('SeparAnnulAction', false);
  SetControlVisible('ChangeEtablissement', False);
  SetControlVisible('ChangeLieu', false);
  SetControlVisible('ModifBases', false);
  SetControlVisible('ActionRegroupement', false);
  SetControlVisible('Ecritures', false);
  SetControlVisible('Ecritures1', false);
  SetControlVisible('DepActif', false);
  SetControlVisible('ChangementMethode', False);
  SetControlVisible('ModifService', False); // fq 17923 mbo
  SetControlVisible('Prime', False);
  SetControlVisible('Subvention', False);
  SetControlVisible('Remplacement', False);
  SetControlVisible('ReductionPrime', False);
  SetControlVisible('ReductionSBV', False);

  if not SelectionMultiple then
     exit;
  if FListe.AllSelected then
     TheMULQ.First
  else
     FListe.GotoLeBookmark(0);

  // 04/06 FQ 17881 Récup infos immos ICI APRES avoir positionné le curseur sur la 1e immo
  bAcquisitionAnnee := (GetField('I_DATEPIECEA') >= VHImmo^.Encours.Deb)
                        and (GetField('I_DATEPIECEA') <= VHImmo^.Encours.Fin);
  bCedee := (GetField('I_QUANTITE') = 0);

  // FQ 19260
  bDejaPrimeSBV := ExisteSQL (
  'SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+GetField('I_IMMO')+'" AND ((I_SBVPRI<>0) or (I_SBVPRIC <>0) or (I_SBVMT<>0) or (I_SBVMTC <>0))');

  PremNature := GetField('I_NATUREIMMO');
  if not VarIsNull ( TheMulQ.FindField('I_GROUPEIMMO').AsVariant) then
     GroupeImmo := GetField('I_GROUPEIMMO')
  else GroupeImmo := '';
  SetControlVisible('Cession', True);
  // FQ 17881 Mutation initialisée avec la 1e immo
  //SetControlVisible('Mutation', True);
  SetControlVisible('Mutation', (not bCedee) and
          (not (GetField('I_OPELEVEEOPTION') = 'X')) and (not bAcquisitionAnnee));

  // FQ 19099 et 19096
  //SetControlVisible('ChangeEtablissement', True);
  //SetControlVisible('ChangeLieu', True);
  //SetControlVisible('ActionRegroupement', True);
  SetControlVisible('ChangeEtablissement',(not bCedee) and (GetField('I_OPELEVEEOPTION') <> 'X')
                     and (VHImmo^.EtablisCpta = TRUE));
  SetControlVisible('ChangeLieu', (not bCedee) and (GetField('I_OPELEVEEOPTION') <> 'X'));
  SetControlVisible('ActionRegroupement', (not bCedee));

  // BTY 04/06 FQ 17518 Autoriser l'éclatement en série
  // Champ I_REVISIONECO pas accessible directement (absent de AMLISTEIMMO/AMLISTEHISTO)
  DepSelectee := ExisteSQL ('SELECT I_IMMO, I_REVISIONECO FROM IMMO WHERE I_IMMO="'+
                              GetField('I_IMMO')+ '" AND I_REVISIONECO<>0');
  bDepotGar := (PremNature = 'FI') and (GetField('I_IMMOLIE') <> '');

  // FQ 18385
  bDatePBSelectee := ExisteSQL
   ('SELECT I_IMMO,I_DATEDEBECO FROM IMMO WHERE I_IMMO="'+ GetField('I_IMMO') + '" AND I_DATEDEBECO>"' +
   UsDateTime (VHImmo^.Encours.Fin) + '"');

  // FA 18435 Pas d'éclatement si immo présente de l'exceptionnel
  bExcepSelectee := ExisteSQL
           ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+ GetField('I_IMMO') + '" AND I_MONTANTEXC<>0');
  SetControlVisible('Eclatement', (PremNature <> 'CB') and (PremNature <> 'LOC') and
                                  (not bDepotGar) and
                                  (not (GetField('I_METHODEECO')='VAR')) and
                                  (not bAcquisitionAnnee) and  (not bCedee) and
                                  (not (GetField('I_OPELEVEEOPTION') = 'X')) and
                                  (not DepSelectee) and (not bExcepSelectee));
  // BTY 04/06 FQ 17451 Autoriser la révision du plan, la saisie d'exceptionnel
  // et la modif de bases en série
  // BTY 19260
  bRevisionPlan := True;
  if ((GetField('I_METHODEECO')='NAM') and (bDejaPrimeSBV)) then
      bRevisionPlan := False;
  SetControlVisible('DureeMethode', (PremNature = 'PRO') and (not bAcquisitionAnnee)
                    and (not bCedee) and (not (GetField('I_OPECHANGEPLAN') = 'X'))
                    and (not (GetField('I_METHODEECO') = 'VAR')) and (bRevisionPlan) );
  SetControlVisible('ElementExceptionnel', (PremNature = 'PRO') and (not bCedee) and
                   (not (GetField('I_OPECHANGEPLAN') = 'X')) and
                   (not (GetField('I_METHODEECO')='NAM')) and (not DepSelectee) // );
                    and (not bDatePBSelectee));  // FQ 18385
  // FQ 19015 Pas de modif base sur une immo non amortissable
  SetControlVisible('ModifBases',(PremNature <> 'LOC') and (PremNature <> 'FI')
                  and (not bCedee) and (not bAcquisitionAnnee)
                    and (not (GetField('I_OPELEVEEOPTION') = 'X'))
                    and (not (GetField('I_METHODEECO') = 'VAR')) and (not DepSelectee)
                    and (not (GetField('I_METHODEECO') = 'NAM')) );

  // FQ 19099 et 19096

  if FListe.AllSelected then
  begin
    TListeAll := TOB.Create ('',nil,-1);
    try
      RecupereAllSelect ( TListeAll );

      // FQ 18386
      // SetControlVisible('DepActif', PremNature='PRO');
      SetControlVisible('DepActif', PremNature='PRO');

      for i := 0 to TListeAll.Detail.Count - 1 do
      begin
        EnabledCession := ((TListeAll.Detail[i].GetValue('I_NATUREIMMO') = 'PRO') or
          (TListeAll.Detail[i].GetValue('I_NATUREIMMO') = 'FI')) and (TListeAll.Detail[i].GetValue('I_QUANTITE') <> 0);
        SetControlVisible('Cession', GetControlVisible('Cession') and
             (EnabledCession) and (not bCedee));
        // FQ 17881 Cacher la mutation si l'une des immos n'est pas mutable
        // Voir plus bas
        //if (bCedee) or bAcquisitionAnnee or ((GetControlVisible('Mutation') and
        //   (TListeAll.Detail[i].GetValue('I_NATUREIMMO') <> PremNature))) then
        //   SetControlVisible('Mutation', false);
        SetControlVisible('ChangeEtablissement',
                           GetControlVisible('ChangeEtablissement')
                           and (TListeAll.Detail[i].GetValue('I_QUANTITE') <> 0)
                           and (TListeAll.Detail[i].GetValue('I_OPELEVEEOPTION') <> 'X') // FQ 19096
                           and (VHImmo^.EtablisCpta = TRUE));   // FQ 19096  (not bCedee) and (VHImmo^.EtablisCpta = TRUE));
        SetControlVisible('ChangeLieu',
                           GetControlVisible('ChangeLieu')
                           and (TListeAll.Detail[i].GetValue('I_OPELEVEEOPTION') <> 'X') // FQ 19096
                           and (TListeAll.Detail[i].GetValue('I_QUANTITE') <> 0));// FQ 19096 and (not bCedee));
        //Pas de regroupement si immo sortie
        //SetControlVisible('ActionRegroupement', True);
        SetControlVisible('ActionRegroupement',(TListeAll.Detail[i].GetValue('I_QUANTITE')<>0));// FQ 19096  not bCedee);

        // En multisélection, option Dépréciation ouverte
        // si toutes de nature = PRO et toutes de même regroupement
        if (TListeAll.Detail[i].GetValue('I_NATUREIMMO') <> 'PRO')
        or (TListeAll.Detail[i].GetValue('I_GROUPEIMMO') <> GroupeImmo) then
           SetControlVisible('DepActif', False);

        { FQ 14403 Pour la sortie de plusieurs immos, il faut qu'elles appartiennent au même groupe }
        if TListeAll.Detail[i].GetValue('I_GROUPEIMMO')<>GroupeImmo then
           SetControlVisible ('Cession',False);

        // BTY 04/06 FQ 17518 Tester si une immo au moins est éclatable
        NatureSelectee := TListeAll.Detail[i].GetValue('I_NATUREIMMO');
        AcquisitionSelectee := (TListeAll.Detail[i].GetValue('I_DATEPIECEA') >= VHImmo^.Encours.Deb)
                           and (TListeAll.Detail[i].GetValue('I_DATEPIECEA') <= VHImmo^.Encours.Fin);
        bDepotGar := (NatureSelectee='FI') and (TListeAll.Detail[i].GetValue('I_IMMOLIE')<>'');
        // Champ I_REVISIONECO dans la TOB TListeAll.Detail[i]
        DepSelectee := (TListeAll.Detail[i].GetValue('I_REVISIONECO') <> 0 );
        // FQ 18385 Champ I_DATEDEBECO dans la TOB
        bDatePBSelectee := (TListeAll.Detail[i].GetValue('I_DATEDEBECO') > VHImmo.Encours.Fin);
        // FQ 18435 Champ I_MONTANTEXC dans la TOB
        bExcepSelectee := (TListeAll.Detail[i].GetValue('I_MONTANTEXC') <> 0 );

        // FQ 17881
        // BTY-06/07- Message ERREUR:champ I_LEVEEOPTION n'existe pas => I_OPELEVEEOPTION
        if (TListeAll.Detail[i].GetValue('I_QUANTITE') = 0)
        or (TListeAll.Detail[i].GetValue('I_OPELEVEEOPTION') = 'X' )
        or AcquisitionSelectee then
           SetControlVisible('Mutation', False);

        if (NatureSelectee <> 'CB') and (NatureSelectee <> 'LOC') and
           (not bDepotGar) and (not (TListeAll.Detail[i].GetValue('I_METHODEECO') = 'VAR')) and
           (not AcquisitionSelectee) and (not (TListeAll.Detail[i].GetValue('I_QUANTITE')=0)) and
           (not (TListeAll.Detail[i].GetValue('I_OPELEVEEOPTION') = 'X')) and
           (not DepSelectee) and (not bExcepSelectee)  then   // FQ 18435
             SetControlVisible('Eclatement', True);
        // BTY 04/06 FQ 17451 Tester si la révision du plan, la saisie d'exceptionnel
        // et la modif de bases possible sur une immo au moins
        // FQ 19260
        bDejaPrimeSBVSelectee := ((TListeAll.Detail[i].GetValue('I_SBVPRI') <> 0) or
                                  (TListeAll.Detail[i].GetValue('I_SBVPRIC') <> 0) or
                                  (TListeAll.Detail[i].GetValue('I_SBVMT') <> 0) or
                                  (TListeAll.Detail[i].GetValue('I_SBVMTC') <> 0) );
        bRevisionPlanSelectee := True;
        if ((TListeAll.Detail[i].GetValue('I_METHODEECO')='NAM') and (bDejaPrimeSBVSelectee)) then
           bRevisionPlanSelectee := False;

        if (NatureSelectee = 'PRO') and (not AcquisitionSelectee) and
           (not (TListeAll.Detail[i].GetValue('I_QUANTITE') = 0)) and
           (not (TListeAll.Detail[i].GetValue('I_OPECHANGEPLAN') = 'X')) and
           (not (TListeAll.Detail[i].GetValue('I_METHODEECO')= 'VAR')) and
           (bRevisionPlanSelectee) then
                 SetControlVisible('DureeMethode', True);
        if (NatureSelectee = 'PRO') and
           (not (TListeAll.Detail[i].GetValue('I_QUANTITE') = 0)) and
           (not (TListeAll.Detail[i].GetValue('I_OPECHANGEPLAN') = 'X')) and
           (not (TListeAll.Detail[i].GetValue('I_METHODEECO')='NAM')) and (not DepSelectee) //
            and (not bDatePBSelectee) then // FQ 18385
                 SetControlVisible('ElementExceptionnel', True);
        if (NatureSelectee <> 'LOC') and (NatureSelectee <> 'FI') and
           (not AcquisitionSelectee) and
           (not (TListeAll.Detail[i].GetValue('I_QUANTITE') = 0))  and
           (not (TListeAll.Detail[i].GetValue('I_OPELEVEEOPTION') = 'X')) and
           (not (TListeAll.Detail[i].GetValue('I_METHODEECO')='VAR')) and (not DepSelectee) and
           (not (TListeAll.Detail[i].GetValue('I_METHODEECO')='NAM')) then  // FQ 19015
                SetControlVisible('ModifBases', True);
      end;
      // fListe.AllSelected := False;
    finally
      FreeAndNil (TListeAll);
    end;
  end
  else
  begin
    IF FListe.NbSelected - 1 >0 THEN
    begin

    // FQ 18386
    //SetControlVisible('DepActif', PremNature='PRO');
    SetControlVisible('DepActif', (PremNature='PRO') );

    for i := 0 to FListe.NbSelected - 1 do
    begin
      FListe.GotoLeBookmark(i);
      // Autorisation Cession
      EnabledCession := ((GetField('I_NATUREIMMO') = 'PRO') or
         (GetField('I_NATUREIMMO') = 'FI')) and (GetField('I_QUANTITE') <> 0);
      SetControlVisible('Cession', GetControlVisible('Cession') and
          (EnabledCession) and (not bCedee));

      // FQ 17881 Cacher la mutation si l'une des immos n'est pas mutable
      // Voir plus bas
      //if (bCedee) or bAcquisitionAnnee or ((GetControlVisible('Mutation') and
      //  (GetField('I_NATUREIMMO') <> PremNature))) then
      //  SetControlVisible('Mutation', false);

      SetControlVisible('ChangeEtablissement',
                         GetControlVisible('ChangeEtablissement')
                         and (GetField('I_QUANTITE') <> 0)
                         and (GetField('I_OPELEVEEOPTION') <> 'X') // FQ 19096
                         and (VHImmo^.EtablisCpta = TRUE)); // FQ 19096 and (not bCedee) and (VHImmo^.EtablisCpta = TRUE));
      SetControlVisible('ChangeLieu', GetControlVisible('ChangeLieu')
                         and (GetField('I_OPELEVEEOPTION') <> 'X') // FQ 19096
                         and (GetField('I_QUANTITE')<> 0));       // FQ 19096 and (not   bCedee));
      // Pas de regroupement si immo sortie
      //SetControlVisible('ActionRegroupement', True);
      SetControlVisible('ActionRegroupement', (GetField('I_QUANTITE')<> 0));// FQ 19096 not bCedee);

      // En multisélection, option Dépréciation ouverte
      // si toutes de nature = PRO et toutes de même regroupement
      if (GetField('I_NATUREIMMO') <> 'PRO') or (GetField('I_GROUPEIMMO') <> GroupeImmo) then
         SetControlVisible('DepActif', False);

      { FQ 14403 Pour la sortie de plusieurs immos, il faut qu'elles appartiennent au même groupe }
      if GetField('I_GROUPEIMMO')<>GroupeImmo then SetControlVisible ('Cession',False);

      // BTY 04/06 FQ 17518 Tester si une immo au moins est éclatable
      DepSelectee := ExisteSQL ('SELECT I_IMMO, I_REVISIONECO FROM IMMO WHERE I_IMMO="'+
                                 GetField('I_IMMO')+ '" AND I_REVISIONECO<>0');
      NatureSelectee := GetField('I_NATUREIMMO');
      AcquisitionSelectee := (GetField('I_DATEPIECEA') >= VHImmo^.Encours.Deb)
                         and (GetField('I_DATEPIECEA') <= VHImmo^.Encours.Fin);
      bDepotGar := (NatureSelectee = 'FI') and (GetField('I_IMMOLIE') <> '');
      // FQ 18385
      bDatePBSelectee := ExisteSQL
       ('SELECT I_IMMO,I_DATEDEBECO FROM IMMO WHERE I_IMMO="'+ GetField('I_IMMO') + '" AND I_DATEDEBECO>"' +
       UsDateTime (VHImmo^.Encours.Fin) + '"');

      // FQ 18435 Pas d'éclatement si immo présente de l'exceptionnel
      bExcepSelectee := ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+ GetField('I_IMMO')+
                                   '" AND I_MONTANTEXC<>0');
      // FQ 19260
      bDejaPrimeSBVSelectee := ExisteSQL (
      'SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+GetField('I_IMMO')+'" AND ((I_SBVPRI<>0) or (I_SBVPRIC <>0) or (I_SBVMT<>0) or (I_SBVMTC <>0))');
      bRevisionPlanSelectee := True;
      if ((GetField('I_METHODEECO')='NAM') and (bDejaPrimeSBVSelectee)) then
           bRevisionPlanSelectee := False;

      // FQ 17881
      // BTY-06/07- Message ERREUR:champ I_LEVEEOPTION n'existe pas => I_OPELEVEEOPTION
      if (GetField('I_QUANTITE') = 0)  or (GetField('I_OPELEVEEOPTION') = 'X' )
      or AcquisitionSelectee then
         SetControlVisible('Mutation', False);

      if (NatureSelectee <> 'CB') and (NatureSelectee <> 'LOC') and
         (not bDepotGar) and (not(GetField('I_METHODEECO')='VAR')) and
         (not AcquisitionSelectee) and (not (GetField('I_QUANTITE') = 0)) and
         (not (GetField('I_OPELEVEEOPTION') = 'X')) and (not DepSelectee) and (not bExcepSelectee) then // FQ 18435
          SetControlVisible('Eclatement', True);

      // BTY 04/06 FQ 17451 Tester si la révision du plan, la saisie d'exceptionnel
      // et la modif de bases possible sur une immo au moins
      if (NatureSelectee = 'PRO') and (not AcquisitionSelectee) and
         (not (GetField('I_QUANTITE') = 0)) and
         (not (GetField('I_OPECHANGEPLAN') = 'X')) and
         (not (GetField('I_METHODEECO')= 'VAR')) and
         (bRevisionPlanSelectee) then
               SetControlVisible('DureeMethode', True);
      if (NatureSelectee = 'PRO') and
         (not (GetField('I_QUANTITE') = 0)) and
         (not (GetField('I_OPECHANGEPLAN') = 'X')) and
         (not (GetField('I_METHODEECO')='NAM')) and (not DepSelectee) //
          and (not bDatePBSelectee) then // FQ 18385
               SetControlVisible('ElementExceptionnel', True);
      if (NatureSelectee <> 'LOC') and (NatureSelectee <> 'FI') and
         (not AcquisitionSelectee) and
         (not (GetField('I_QUANTITE') = 0))  and
         (not (GetField('I_OPELEVEEOPTION') = 'X')) and
         (not (GetField('I_METHODEECO')='VAR')) and (not DepSelectee) and
         (not (GetField('I_METHODEECO')='NAM')) then   // FQ 19015
              SetControlVisible('ModifBases', True);
    end;
    end;
  end;

  // BTY 09/06 En perspective DPI, bloquer les opérations
  EnableZoomActionDPI;
end;

procedure TOF_AMLISTE.EnableZoomVisu;
var
  QLog: TQuery;
  Nature: string;
  PasSimu : boolean;
begin
  inherited;
  SetControlVisible('ZoomPlan', false);

  SetControlVisible('ZoomPrime', false);    // ajout mbo pour prime
  SetControlVisible('ZoomSubvention', false);    // ajout mbo pour sbv

  SetControlVisible('ZoomEcheance', false);
  SetControlVisible('ZoomOperations', false);
  SetControlVisible('Ecritures',False);
  if ListeVide then
    exit;
  SetControlVisible('ZoomPlan', ImmoAmortie(GetField('I_METHODEECO'),
                varToStr(GetField('I_METHODEFISC')), GetField('I_NATUREIMMO')));

  // ajout mbo pour suramortissement prime d'équipement
  if ExisteSQL ('SELECT I_IMMO, I_SBVPRI, I_SBVPRIC FROM IMMO WHERE I_IMMO="'+
                                GetField('I_IMMO') + '" AND ((I_SBVPRI<>0) or (I_SBVPRIC<>0))') then
     SetControlVisible('ZoomPrime', true);
  // fin ajout prime

  // ajout mbo pour suramortissement subvention d'équipement
  if ExisteSQL ('SELECT I_IMMO, I_SBVMT, I_SBVMTC FROM IMMO WHERE I_IMMO="'+
                                GetField('I_IMMO') + '" AND ((I_SBVMT<>0) or (I_SBVMTC<>0))') then
     SetControlVisible('ZoomSubvention', true);
  // fin ajout prime

  Nature := GetField('I_NATUREIMMO');

  SetControlVisible('ZoomEcheance', ((Nature = 'CB') or (Nature = 'LOC')));

  QLog := OpenSQL('SELECT * from IMMOLOG WHERE IL_IMMO="' +
          GetField('I_IMMO') +
          '" AND IL_TYPEOP<>"ACQ" AND IL_TYPEOP<>"CLO" AND (IL_DATEOP>="' +
          USDateTime(VHImmo^.Encours.Deb) + '" AND IL_DATEOP<="' +
          USDateTime(VHImmo^.Encours.Fin) + '")', True);
  SetControlVisible('ZoomOperations', not (QLog.Eof));
  SetControlVisible('SeparOpeVisu', not (QLog.Eof));
  Ferme(QLog);
  SetControlVisible('SeparPlanVisu', GetControlVisible('ZoomEcheance') or
                     GetControlVisible('ZoomPlan'));

  // fq 19070 - ajout du test sur pas simulation et pas financière
    // mbo - fq 19070 - ajout du test sur qualif
  PasSimu := ExisteSQL ('SELECT I_IMMO, I_QUALIFIMMO FROM IMMO WHERE I_IMMO="'+
                                GetField('I_IMMO') + '" AND (I_QUALIFIMMO<>"S")');

  SetControlVisible('Ecritures', ((SelectionUnique) and (PasSimu) and (Nature <> 'FI')) );
end;


procedure TOF_AMLISTE.EnableZoomVisuSerie;
begin
  SetControlVisible('ZoomPlan', false);
  SetControlVisible('ZoomPrime', false);         // ajout mbo pour prime
  SetControlVisible('ZoomSubvention', false);    // ajout mbo pour subvention
  SetControlVisible('ZoomEcheance', false);
  SetControlVisible('SeparPlanVisu', false);
  SetControlVisible('ZoomOperations', false);
  SetControlVisible('Ecritures',False);
end;

procedure TOF_AMLISTE.ExecuteOperationSerie(TypeOpe: TypeOperation);
var
  FOpeSerie: TOperationSerie;
  i, OrdreSerie, ret: integer;
  QTmp, QSortie: TQuery;
  FCodeImmo: string;
  TListeAll : TOB;
  stCode : string;
  ToutATraiter : boolean;
begin
  stCode := GetField ('I_IMMO');
  ToutATraiter := True;

  FOpeSerie := TOperationSerie.Create(Application);

  try
    // Init entre autres de la date derniere opération
    FOpeSerie.InitOperationSerie(TypeOpe, FListe.NbSelected);
    if TypeOpe = toMutation then
       FOpeSerie.fNature := GetField('I_NATUREIMMO');

    // FQ 19208 Vérif toutes les immos sélectionnées sont traitables
    //          + Récup date d'opération la plus récente
    if fListe.AllSelected then
    begin
       TListeAll := TOB.Create ('', nil, -1);
       try
           RecupereAllSelect(TListeAll);
           if (TListeAll.Detail.Count > 0) then
           begin
             if (fListe.NbSelected >= 1) then fListe.GotoLeBookmark(0);
             // Vérif nature si mutation + récup plus récente date d'opération
             ToutATraiter := ImmosAMuterChanLieuEtab (stCode,FOpeSerie.fNature, TListeAll, TypeOpe, FOpeSerie.fDateDerniereOpe);
           end;
       finally
           FreeAndNil (TListeAll);
       end;
    end
    else
    begin
       if (fListe.NbSelected >= 1) then fListe.GotoLeBookmark(0);
       ToutATraiter := ImmosAMuterChanLieuEtab (stCode,FOpeSerie.fNature,nil,TypeOpe,FOpeSerie.fDateDerniereOpe);
    end;
    // Immos de nature différente
    if (not ToutATraiter) and (TypeOpe=toMutation) then
    begin
       FOpeSerie.fDateDerniereOpe := iDate1900;   // Ne pas préafficher date opération
       FOpeSerie.fCode := stCode;        // Ctl date ope en saisie opération sur immo courante
       PGIBox(HM[9], 'Mutation');
    end;

    ret := FOpeSerie.AfficheOperationSerie;

    if (ret = mrOK) or (ret = mrYes) then
    begin
{$IFDEF EAGLCLIENT}
{$ELSE}
      if TypeOpe = toCession then
        FOpeSerie.CalculCumulImmoCedee(THGrid(FListe), THQuery(TheMulQ));
{$ENDIF}
      QTmp := OpenSQL('SELECT MAX(IL_ORDRESERIE) FROM IMMOLOG', TRUE);
      if not QTmp.EOF then
         OrdreSerie := (QTmp.Fields[0].AsInteger + 1)
      else
         OrdreSerie := 1;
      Ferme(QTmp);
      if FListe.AllSelected then
      begin
        TListeAll := TOB.Create ('', nil, -1);
        try
          RecupereAllSelect(TListeAll);
          if (TListeAll.Detail.Count > 0) then
          begin
            // On se positionne sur un enregistrement de la sélection
            if (fListe.NbSelected >= 1) then fListe.GotoLeBookmark(0);

            // FQ 19208
            // Si Mutation, toutes les immos sont de même nature, sinon OK
            if ToutATraiter then
            // Enregistrement de l'opération pour toutes les immos
            begin
              for i := 0 to TListeAll.Detail.Count - 1 do
              begin
                FCodeImmo := TListeAll.Detail[i].GetValue('I_IMMO');
                case TypeOpe of
                  toCession:
                  begin
                    QSortie :=
                    OpenSQL('SELECT I_QUANTITE,I_MONTANTHT,I_BASEECO,I_BASEFISC FROM IMMO WHERE I_IMMO="'
                    + FCodeImmo + '"', True);
                    FOpeSerie.OperationSerieCession(FCodeImmo, OrdreSerie, QSortie);
                    Ferme(QSortie);
                  end;
                  toMutation:
                  begin
                    FOpeSerie.OperationSerieMutation(FCodeImmo, OrdreSerie);
                  end;
                  toChanEtabl:
                  begin
                    FOpeSerie.OperationSerieEtabl(FCodeImmo, OrdreSerie);
                  end;
                  toChanLieu:
                  begin
                    FOpeSerie.OperationSerieLieu(FCodeImmo, OrdreSerie);
                  end;
                  // FQ 17395 Changement de regroupement si Ctl+A
                  toRegroupement:
                  begin
                    //FQ 17516 FOpeSerie.OperationSerieRegroupement(FCodeImmo);
                    FOpeSerie.OperationSerieRegroupement(FCodeImmo, OrdreSerie);
                  end;
                end; // fin case
              end;   // for

            // FQ 19208
            end
            else
            begin
              // Enregistrement de l'opération uniquement pour l'immo en cours
              case TypeOpe of
                  toMutation:
                  begin
                   FOpeSerie.OperationSerieMutation(stCode, OrdreSerie);
                  end;
              end;
            end; // ToutAMuter
            //
          end;
          fListe.AllSelected := False;
        finally
          FreeAndNil (TListeAll);
        end;
      end
      else
      begin
        InitMove(FListe.NbSelected, '');
        if (fListe.NbSelected >= 1) then fListe.GotoLeBookmark(0);
        // FQ 19208
        // Si Mutation, immos sélectionnées sont de même nature, sinon OK
        if ToutATraiter then
        // Enregistrement de l'opération pour toute la sélection
        begin
          for i := 0 to FListe.NbSelected - 1 do
          begin
            MoveCur(False);
            FListe.GotoLeBookmark(i);
            FCodeImmo := GetField('I_IMMO');
            case TypeOpe of
              toCession:
                begin
                  QSortie :=
                  OpenSQL('SELECT I_QUANTITE,I_MONTANTHT,I_BASEECO,I_BASEFISC FROM IMMO WHERE I_IMMO="'
                  + FCodeImmo + '"', True);
                  FOpeSerie.OperationSerieCession(FCodeImmo, OrdreSerie, QSortie);
                  Ferme(QSortie);
                end;
              toMutation:
                begin
                  FOpeSerie.OperationSerieMutation(FCodeImmo, OrdreSerie);
                end;
              toChanEtabl:
                begin
                  FOpeSerie.OperationSerieEtabl(FCodeImmo, OrdreSerie);
                end;
              toChanLieu:
                begin
                  FOpeSerie.OperationSerieLieu(FCodeImmo, OrdreSerie);
                end;
              toRegroupement:
                begin
                  // FQ 17516
                  //FOpeSerie.OperationSerieRegroupement(FCodeImmo);
                  FOpeSerie.OperationSerieRegroupement(FCodeImmo, OrdreSerie);
                end;
            end; // fin case
          end;   // for

        // FQ 19208
        end
        else
        begin
          // Enregistrement de l'opération uniquement pour l'immo en cours
          case TypeOpe of
            toMutation:
            begin
              FOpeSerie.OperationSerieMutation(stCode, OrdreSerie);
            end;
          end;
        end; // ImmosAMuter
        FListe.ClearSelected;
      end;
      FiniMove;
    end;
  finally
    SetControlProperty('bSelectAll', 'Down', FListe.AllSelected);
    FOpeSerie.DetruitOperationSerie;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 23/11/2006
Modifié le ... :
Description .. : On vérifie que toutes les immos de la sélection
Suite ........ : peuvent subir la mutation <=> soient de même nature
Suite ........ : Autres opérations : ne fait rien
Mots clefs ... :
*****************************************************************}
function TOF_AMLISTE.ImmosAMuterChanLieuEtab (stCode:string; stNature:string; TListeAll:Tob; TypeOpe:TypeOperation; var DateDerniereOpe:TDateTime): boolean;
var i : integer;
    Q : TQuery;
    NatureSelectee : string;
begin
  result := True;

  // if TypeOpe <> toMutation then exit;

  if fListe.AllSelected then
  begin
    for i := 0 to TListeAll.Detail.Count - 1 do
    begin
      NatureSelectee := TListeAll.Detail[i].GetValue('I_NATUREIMMO');

      case TypeOpe of
      toMutation : begin
                   if NatureSelectee <> stNature then
                      begin
                      result := False;
                      Break;
                      end;
                   end;
      end; // case

      // Date de dernière opération dans l'exercice
      Q := OpenSQL ('SELECT IL_DATEOP FROM IMMOLOG WHERE IL_IMMO="' +
                     TListeAll.Detail[i].GetValue('I_IMMO') +
                     '" AND (IL_DATEOP >="' + USDateTime(VHImmo^.Encours.Deb) +
                     '" AND IL_DATEOP <="' + USDateTime(VHImmo^.Encours.Fin) + '")' +
                     ' ORDER BY IL_DATEOP DESC', True);
      if not Q.Eof then
         if Q.FindField ('IL_DATEOP').AsDateTime > DateDerniereOpe then
            DateDerniereOpe := Q.FindField ('IL_DATEOP').AsDateTime;
      Ferme (Q);
    end; // for

  end
  else
  begin

    for i := 0 to FListe.NbSelected - 1 do
    begin
      FListe.GotoLeBookmark(i);
      NatureSelectee := GetField('I_NATUREIMMO');

      case TypeOpe of
      toMutation : begin
                   if NatureSelectee <> stNature then
                      begin
                      result := False;
                      Break;
                      end;
                   end;
      end; // case

      // Date de dernière opération dans l'exercice
      Q := OpenSQL ('SELECT IL_DATEOP FROM IMMOLOG WHERE IL_IMMO="' +
                       GetField('I_IMMO') +
                       '" AND (IL_DATEOP >="' + USDateTime(VHImmo^.Encours.Deb) +
                       '" AND IL_DATEOP <="' + USDateTime(VHImmo^.Encours.Fin) + '")' +
                       ' ORDER BY IL_DATEOP DESC', True);
      if not Q.Eof then
         if Q.FindField ('IL_DATEOP').AsDateTime > DateDerniereOpe then
            DateDerniereOpe := Q.FindField ('IL_DATEOP').AsDateTime;
      Ferme (Q);
    end; // for
  end;

end;



function TOF_AMLISTE.ListeVide: boolean;
begin
{$IFDEF EAGLCLIENT}
   TheMulQ := TFMul(Ecran).Q.TQ;
   Result := (TheMulQ = nil) or ((TheMulQ <> nil) and (TheMulQ.Detail.Count = 0));
{$ELSE}
   TheMulQ := TFMul(Ecran).Q;
   Result := TheMulQ.isEmpty;
{$ENDIF}
end;

procedure TOF_AMLISTE.OnKeyDownEcran(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
 // BTY FQ 17361 Pas de duplication en série car ne recopie pas IMMOLOG
 // => passer par la duplication du menu [Nouveau]
{
    VK_F3:
      begin
        Key := 0;
        if Shift = [ssAlt] then
          LanceCreationEnSerie;
        TFMul(ECRAN).BChercheClick(nil);
      end;
}
    VK_DELETE :
      begin
        // BTY FQ 17358 Pas de CTL+Suppr si bouton bDelete non visible
        // if (Shift = [ssCtrl]) and
        //  (TToolbarButton97(GetControl('bDelete')).Enabled) then
        if (Shift = [ssCtrl]) and
           (GetControlVisible ('bDelete')) and
           (TToolbarButton97(GetControl('bDelete')).Enabled) then
        begin
          Key := 0;
          TToolbarButton97(GetControl('bDelete')).Click;
        end;
      end;
  else
    fOnSaveKeyDownEcran(Sender, Key, Shift);
  end;

end;

procedure TOF_AMLISTE.OnKeyDownFListe(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin

  case Key of
    VK_F11:
      begin
        Key := 0;
        if ((fAction <> taConsult) and (not fbHistorique)) then
              TPopupMenu(GetControl('PopupListe')).Popup(Mouse.CursorPos.x,
              Mouse.CursorPos.y);
      end;
    VK_F5:
      begin
        Key := 0;
        OnListeDblClick(nil);
      end;
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Recomposer le critère complémentaire XX_WHERE
Suite ........ : suivant les sélections constatées dans les onglets du MUL
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.RecupereCritereXX_WHERE(Sender: TObject);
begin
  // onglet Compléments
  SetControltext('XX_WHERE', GetControlText('XX_WHERE') + fWhereAmortDiffere);
  // onglet Amortissements
  //fq 17512 - SetControltext('XX_WHERE', GetControlText('XX_WHERE') + fWhereBaseEcoTh + fWhereDerog) ;
  SetControltext('XX_WHERE', GetControlText('XX_WHERE') + fWhereBaseEcoTh + fWhereDerog + fWhereGestionFiscale);
  // onglet Opérations
  SetControltext('XX_WHERE', GetControlText('XX_WHERE') + fWhereSansOpe + fWhereOpe
                             + fWhereReintegration + fWhereQuotePart + fWhereCheckDPI 
                             + fWhereCheckPRI + fWhereCheckSBV);    // 10/06 Critère prime d'équipement
                                                                    // 10/06 Subvention d'investissement
                             //   + fWhereReintegration + fWhereQuotePart); 7.10 06/06 DPI
  // onglet Recherche
  SetControltext('XX_WHERE', GetControlText('XX_WHERE') + fWhereBaseTPTh +
                             fWhereExigibleTP);  // FQ 14133

  // onglet Caractéristiques
  SetControltext('XX_WHERE', GetControlText('XX_WHERE') + fWhereSansGroupe);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : CREATION Crédit-bail
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnCreationCBClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := '';
  if not TheMulQ.Eof then
     stCode := GetField('I_IMMO');
  AMLanceFiche_FicheImmobilisation('', taCreatOne, 'CB');
  RechargeListeChoix;
  // Réactualiser, l'immo courante est perdue
  RafraichirListeImmo (stCode);
  AfficheOpeEnCours (stCode);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : CREATION Immo Financière
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnCreationFIClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := '';
  if not TheMulQ.Eof then
     stCode := GetField('I_IMMO');
  AMLanceFiche_FicheImmobilisation('', tacreatOne, 'FI');
  RechargeListeChoix;
  // Réactualiser, l'immo courante est perdue
  RafraichirListeImmo (stCode);
  AfficheOpeEnCours (stCode);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : CREATION Location longue durée
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnCreationLOCClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := '';
  if not TheMulQ.Eof then
     stCode := GetField('I_IMMO');
  AMLanceFiche_FicheImmobilisation('', taCreatOne, 'LOC');
  RechargeListeChoix;
  // Réactualiser, l'immo courante est perdue
  RafraichirListeImmo (stCode);
  AfficheOpeEnCours (stCode);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : CREATION Immo en pleine propriété
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnCreationPROClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := '';
  if TheMulQ.RecordCount <> 0 then
     stCode := GetField('I_IMMO');
  AMLanceFiche_FicheImmobilisation('', taCreatOne, 'PRO');
  RechargeListeChoix;
  // Réactualiser, l'immo courante est perdue
  RafraichirListeImmo (stCode);
  AfficheOpeEnCours (stCode);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : CREATION de plusieurs immos en série
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnCreationSerieClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := '';
  if not TheMulQ.Eof then
     stCode := GetField('I_IMMO');
  AMLanceFiche_FicheImmobilisation('', taCreat, '');
  RechargeListeChoix;
  // Réactualiser, l'immo courante est perdue
  RafraichirListeImmo (stCode);
  AfficheOpeEnCours (stCode);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : CREATION d'immo par duplication d'une immo
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnDupliqueImmoClick(Sender: TObject);
var
  NewCodeImmo: string;
  ret: integer;
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  if not IsOpeEnCours(nil, GetField('I_IMMO'), false) then
  begin
    // ajout mbo 26.04.2007 pas de duplic d'un composant
    // 06/07 i_string1 au lieu de i_remplace
    if (GetColonneSQL('IMMO','I_STRING1','I_IMMO="'+stCode+'"') <> '') then
    begin
       PGIBox(HM[10]);
    end else
    begin
       ret := ExecuteDuplication(GetField('I_IMMO'), NewCodeImmo);
       if ret = mrYes then
          begin
          AMLanceFiche_FicheImmobilisation(NewCodeImmo, taModif, '');
          // 05/06 FQ 18119 Positionner les indicateurs de modif
          VHImmo^.ChargeOBImmo := True;
          ImMarquerPublifi(True);
          end;
       // Réactualiser, l'immo courante est perdue
       RafraichirListeImmo (stCode);
       AfficheOpeEnCours (stCode);
    end;
  end
  else
    PGIBox(HM[2]); //PGIBox(HM[13]);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Touche ESPACE de sélection d'une immo
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnFlipSelectionGrille(Sender: TObject);
begin
  inherited;
  SetControlEnabled('BZOOMACTION', (FListe.NbSelected > 0) or (FListe.AllSelected));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Ouverture fiche immo
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnListeDblClick(Sender: TObject);
var
  stCode: string;
  stParamDPI, sSepar: string;
  i: integer;
begin
{$IFDEF EAGLCLIENT}
   TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
   TheMulQ := TFMul(Ecran).Q;
{$ENDIF}

  if TheMulQ.RecordCount <> 0 then
  begin
    stCode := GetField('I_IMMO');
    // 7.10 06/06 DPI
    // Appel par le Suivi DPI => affectation à une immo => paramètres à passer à la fiche immo
    stParamDPI := '';
    sSepar := ';';
{$IFNDEF CMPGIS35}
    if fDateDebExoDPI <> iDate1900 then
       begin
       stParamDPI := DateToStr (fDateDebExoDPI) + sSepar + DateToStr(fDateFinExoDPI)
                     + sSepar + FloatToStr(fMontantMaxiDPI);
       // 07/06 Paramètres DPI dates exos N-1 à N-5 + soldes N-1 à N-5
       for i:=1 to 5 do
         begin
         if  fAmDPI.Exercice[i].Deb <> iDate1900 then
             stParamDPI := stParamDPI + sSepar + DateToStr (fAmDPI.Exercice[i].Deb) +
                           sSepar + DateToStr (fAmDPI.Exercice[i].Fin);
         end;
       for i:=1 to 5 do
         begin
         stParamDPI := stParamDPI + sSepar + FloatToStr(fAmDPI.Solde[i]);
         end;
       end;
{$ENDIF}

    //AMLanceFiche_FicheImmobilisation(stCode, fAction, '');
    AMLanceFiche_FicheImmobilisation(stCode, fAction, '', stParamDPI);
    // FQ 15301 RAS
    RechargeListeChoix;
    // Réactualiser, l'immo courante est perdue
    RafraichirListeImmo (stCode);
    AfficheOpeEnCours (stCode);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Composition du menu F11 en fonction de l'immo courante
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnPopupListeClick(Sender: TObject);
begin
  OnPopZoomVisuPopup(Sender);
  OnPopZoomActionPopup(Sender);
  OnPopZoomRegroupementPopup(Sender);
  SetControlVisible('DupliqueImmoListe', (not ListeVide) and (FListe.NbSelected
    <= 1));
  SetControlVisible('Cession1', GetControlVisible('Cession'));
  SetControlVisible('Eclatement1', GetControlVisible('Eclatement'));
  SetControlVisible('Mutation1', GetControlVisible('Mutation'));
  SetControlVisible('DureeMethode1', GetControlVisible('DureeMethode'));
  SetControlVisible('ElementExceptionnel1',GetControlVisible('ElementExceptionnel'));
  SetControlVisible('DepActif1',GetControlVisible('DepActif'));
  SetControlVisible('ChangementMethode1',GetControlVisible('ChangementMethode'));
  SetControlVisible('ModifService1',GetControlVisible('ModifService'));  // fq 17923 mbo

  SetControlVisible('LeveeOption1', GetControlVisible('LeveeOption'));
  SetControlVisible('AnnulationOperation1', GetControlVisible('AnnulationOperation'));
  SetControlVisible('SeparAnnul', GetControlVisible('AnnulationOperation1'));
  SetControlVisible('ZoomPlan1', GetControlVisible('ZoomPlan'));

  SetControlVisible('ZoomPrime1', GetControlVisible('ZoomPrime'));           // ajout mbo pour prime
  SetControlVisible('ZoomSubvention1', GetControlVisible('ZoomSubvention')); // ajout mbo pour sbv

  SetControlVisible('ZoomEcheance1', GetControlVisible('ZoomEcheance'));
  SetControlVisible('SeparPlan', GetControlVisible('ZoomPlan1') or
       GetControlVisible('ZoomEcheance1'));
  SetControlVisible('ZoomOperations1', GetControlVisible('ZoomOperations'));
  SetControlVisible('SeparOpe', GetControlVisible('ZoomOperations1'));
  SetControlVisible('ChangeEtablissement1', GetControlVisible('ChangeEtablissement'));
  SetControlVisible('ChangeLieu1', GetControlVisible('ChangeLieu'));
  SetControlVisible('ModifBases1', GetControlVisible('ModifBases'));
  SetControlVisible('ActionRegroupement1', GetControlVisible('ActionRegroupement'));
  // Infos du regroupement
  SetControlVisible('SeparInfoReg1', GetControlVisible('BInfoReg'));
  SetControlVisible('BInfoReg1', GetControlVisible('BInfoReg'));
  SetControlVisible('BDetailReg1', GetControlVisible('BDetailReg'));

   SetControlVisible('Prime1', GetControlVisible('Prime'));
   SetControlVisible('Subvention1', GetControlVisible('Subvention'));
   SetControlVisible('Remplacement1', GetControlVisible('Remplacement'));
   SetControlVisible('ReductionPrime1', GetControlVisible('ReductionPrime'));
   SetControlVisible('ReductionSBV1', GetControlVisible('ReductionSBV'));

  // BTY 09/06 En perspective DPI, bloquer les opérations
  EnableZoomAction1DPI;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 06/02/2006
Modifié le ... :   /  /
Description .. : Ouverture/Fermeture des informations de regroupement en
Suite ........ : fonction de l'immo courante
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.EnableZoomRegroupement;
var  GroupeImmo : string;
begin
  if ListeVide then
     begin
     SetControlVisible('BInfoReg', False);
     SetControlVisible('BDetailReg', False);
     end
  else
     begin
     if not VarIsNull ( TheMulQ.FindField('I_GROUPEIMMO').AsVariant) then
        GroupeImmo := GetField('I_GROUPEIMMO')
     else GroupeImmo := '';
     SetControlVisible('BInfoReg', (GroupeImmo<>''));
     SetControlVisible('BDetailReg', (GroupeImmo<>''));
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 06/02/2006
Modifié le ... :   /  /
Description .. : Fermeture des informations de regroupement en multi-sélection
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.EnableZoomRegroupementSerie;
begin
  SetControlVisible('BInfoReg', False);
  SetControlVisible('BDetailReg', False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Composition du menu ACTION en fonction de la sélection
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnPopZoomActionPopup(Sender: TObject);
begin
  if (FListe.NbSelected <= 1) and (not FListe.AllSelected) then
     EnableZoomAction
  else
     EnableZoomActionSerie;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Recomposition du menu CREATION en fonction de la sélection
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnPopZoomCreationPopup(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
   TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
   TheMulQ := TFMul(Ecran).Q;
{$ENDIF}
 SetControlVisible('DupliqueImmo',(not TheMulQ.EOF) and (FListe.NbSelected <= 1));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Composition du menu VISU en fonction de l'immo courante
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnPopZoomVisuPopup(Sender: TObject);
begin
  if (FListe.NbSelected <= 1) and (not FListe.AllSelected) then
     EnableZoomVisu
  else
     EnableZoomVisuSerie;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Bouton Réinit (non visible)
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnReinitialiseClick(Sender: TObject);
begin
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Si liste immos non vide, afficher en bas de
Suite ........ : l'écran les opérations faites sur l'immo qui a le focus
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.AfficheOpeEnCours(Sender: TObject);
begin
  if ListeVide then
     SetControlCaption('sOperationEnCours', '')
  else
     OnRowEnterListe(nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : 2e définition, selon les paramètres d'appel
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.AfficheOpeEnCours(stCode: string);
begin
  if ListeVide then
     SetControlCaption('sOperationEnCours', '')
  else
     OnRowEnterListe(stCode);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Clic sur une immo ou défilement vertical dans la liste
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnRowEnterListe(Sender: TObject);
var
  sListeOpe: string;
begin
  if fbHistorique then
    exit;
  sListeOpe := '';
  if GetField('I_OPEMUTATION') = 'X' then
    sListeOpe := sListeOpe + HM2[1] + ', ';
  if GetField('I_OPEECLATEMENT') = 'X' then
    sListeOpe := sListeOpe + HM2[2] + ', ';
  if GetField('I_OPECESSION') = 'X' then
    sListeOpe := sListeOpe + HM2[3] + ', ';
  if GetField('I_OPEMODIFBASES') = 'X' then
    sListeOpe := sListeOpe + HM2[4] + ', ';
  if GetField('I_OPELIEUGEO') = 'X' then
    sListeOpe := sListeOpe + HM2[5] + ', ';
  if GetField('I_OPEETABLISSEMENT') = 'X' then
    sListeOpe := sListeOpe + HM2[6] + ', ';
  if GetField('I_OPECHANGEPLAN') = 'X' then
    sListeOpe := sListeOpe + HM2[7] + ', ';
  if GetField('I_OPELEVEEOPTION') = 'X' then
    sListeOpe := sListeOpe + HM2[8] + ', ';
  if GetField('I_OPEDEPREC') = 'X' then
    sListeOpe := sListeOpe + HM2[9] + ', ';
  // BTY 04/06 FQ 17516
  if GetField('I_OPEREG') = 'X' then
    sListeOpe := sListeOpe + HM2[10] + ', ';

  // mbo 04/2007     opération de remplacement de composant
  if GetField('I_OPEREMPL') = 'X' then
    sListeOpe := sListeOpe + HM2[11] + ', ';

  // Affichage des opérations de l'immo courante en bas de l'écran
  if Length(sListeOpe) > 0 then
  begin
    System.Delete(sListeOpe, Length(sListeOpe) - 1, 2);
    SetControlCaption('sOperationEnCours', HM2[0] + sListeOpe);
  end
  else
    SetControlCaption('sOperationEnCours', '');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : 2e implémentation de ROWENTERLISTE selon les paramètres d'appel
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnRowEnterListe(stCode: string);
var
  sListeOpe: string;
//  QTmp : TQuery;
begin
  if fbHistorique then
    exit;
  sListeOpe := '';

//  // On a perdu l'immo courante => la relire
//  QTmp := OpenSQL(
//  'SELECT I_IMMO,I_OPEMUTATION,I_OPEECLATEMENT,I_OPECESSION,I_OPEMODIFBASES,I_OPELIEUGEO,I_OPEETABLISSEMENT,I_OPECHANGEPLAN,I_OPELEVEEOPTION,I_OPEDEPREC FROM IMMO WHERE I_IMMO="'
//   + stCode + '"', true);
//  if not QTmp.Eof then
//  begin
    if GetField('I_OPEMUTATION') = 'X' then
    sListeOpe := sListeOpe + HM2[1] + ', ';
    if GetField('I_OPEECLATEMENT') = 'X' then
    sListeOpe := sListeOpe + HM2[2] + ', ';
    if GetField('I_OPECESSION') = 'X' then
    sListeOpe := sListeOpe + HM2[3] + ', ';
    if GetField('I_OPEMODIFBASES') = 'X' then
    sListeOpe := sListeOpe + HM2[4] + ', ';
    if GetField('I_OPELIEUGEO') = 'X' then
    sListeOpe := sListeOpe + HM2[5] + ', ';
    if GetField('I_OPEETABLISSEMENT') = 'X' then
    sListeOpe := sListeOpe + HM2[6] + ', ';
    if GetField('I_OPECHANGEPLAN') = 'X' then
    sListeOpe := sListeOpe + HM2[7] + ', ';
    if GetField('I_OPELEVEEOPTION') = 'X' then
    sListeOpe := sListeOpe + HM2[8] + ', ';
    if GetField('I_OPEDEPREC') = 'X' then
    sListeOpe := sListeOpe + HM2[9] + ', ';
    // BTY 04/06 FQ 17516
    if GetField('I_OPEREG') = 'X' then
    sListeOpe := sListeOpe + HM2[10] + ', ';
    // mbo 04/2007 remplacement de composant
    if GetField('I_OPEREMPL') = 'X' then
    sListeOpe := sListeOpe + HM2[11] + ', ';


    // Affichage des opérations de l'immo courante en bas de l'écran
    if Length(sListeOpe) > 0 then
    begin
      System.Delete(sListeOpe, Length(sListeOpe) - 1, 2);
      SetControlCaption('sOperationEnCours', HM2[0] + sListeOpe);
    end
    else
      SetControlCaption('sOperationEnCours', '');
//  end;
//  Ferme(QTmp);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : CTL+A Sélection de toute la liste
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnSelectAllClick(Sender: TObject);
begin
  TFMul(Ecran).bSelectAllClick ( Sender );
  OnFlipSelectionGrille(fListe);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /    
Description .. : Bouton Suppression
Mots clefs ... : 
*****************************************************************}
procedure TOF_AMLISTE.OnSupprimeClick(Sender: TObject);
var
  i: integer;
  stCode: string;
  TListeAll : TOB;
begin
  // BTY FQ 17358 Sortir tout de suite si liste vide
  // sinon Erreur en conversion de type variant sur le GetField('I_IMMO')
  //stCode := GetField('I_IMMO');
  if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
  begin
    PGIBox(HM[1]); // PGIBox(HM[9]);
    exit;
  end;

  stCode := GetField('I_IMMO'); // FQ 17358
  if FListe.AllSelected then
  begin
    InitMove(TheMulQ.RecordCount, '');
    TListeAll := TOB.Create ('',nil,-1);
    try
      RecupereAllSelect(TListeAll);
      for i := 0 to TListeAll.Detail.Count - 1 do
      begin
        MoveCur(False);
        FCodeImmo := TListeAll.Detail[i].GetValue('I_IMMO');
        // mbo fq 14324 ajout test si immo ayant fait l'objet d'une cloture (ConfirmationSuppressionImmo)
        if ValideSuppressionImmo ( FCodeImmo, true ) then
        begin
          if ConfirmationSuppressionImmo ( FCodeImmo, TListeAll.Detail[i].GetValue('I_ETAT') ) then
          begin
             if Transactions(SuppressionFicheImmo, 3) <> oeOk then
             begin
               MessageAlerte(HM[0]);
               Break;
             end else
             begin
             // 05/06 FQ 18119 Positionner les indicateurs de modif
             VHImmo^.ChargeOBImmo := True;
             ImMarquerPublifi(True);
             end;
          end;
        end;
      end;
      FListe.AllSelected := False;
      TToolbarButton97(GetControl('bSelectAll')).Down := False;
    finally
        FreeAndNil(TListeAll);
    end;
  end
  else
  begin
    InitMove(FListe.NbSelected, '');
    for i := 0 to FListe.NbSelected - 1 do
    begin
      MoveCur(False);
      FListe.GotoLeBookmark(i);
      FCodeImmo := GetField('I_IMMO');
      // ajout mbo fq 14324 - test si immo a fait l'objet d'une cloture (confirmationSuppressionImmo)
      if ValideSuppressionImmo ( FCodeImmo, true ) then
      begin
        if ConfirmationSuppressionImmo ( FCodeImmo, GetField('I_ETAT') ) then
        begin
           if Transactions(SuppressionFicheImmo, 3) <> oeOk then
           begin
             MessageAlerte(HM[0]);
             Break;
           end else
           begin
           // 05/06 FQ 18119 Positionner les indicateurs de modif
           VHImmo^.ChargeOBImmo := True;
           ImMarquerPublifi(True);
           end;
        end;
      end;
    end;
    FListe.ClearSelected;
  end;
  FiniMove;
  RafraichirListeImmo (stCode);
  AfficheOpeEnCours (stCode);

  // FQ 18957 
  SetControlEnabled('BZoomAction', False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 06/02/2006
Modifié le ... : 06/02/2006
Description .. : Composition du menu REGROUPEMENT en fonction de la sélection
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnPopZoomRegroupementPopup(Sender: TObject);
begin
  if (FListe.NbSelected <= 1) and (not FListe.AllSelected) then
     EnableZoomRegroupement
  else
     EnableZoomRegroupementSerie;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : VISU Plan de remboursement
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnZoomEcheanceClick(Sender: TObject);
var
  QTmp: TQuery;
  Contrat: TImContrat;
begin
  inherited;
  QTmp := OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="' + GetField('I_IMMO') +
    '"', true);
  if not QTmp.Eof then
  begin
    Contrat := TImContrat.Create;
    Contrat.Charge(QTmp);
    Contrat.ChargeTableEcheance;
    ListeDesEcheances(Contrat.ListeEcheances,
      QTmp.FindField('I_RESIDUEL').AsFloat);
    Contrat.Free;
  end;
  Ferme(QTmp);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : VISU Ecritures de l'immo
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnZoomEcrituresClick(Sender: TObject);
var
  Nature: string;
  ListeImmo: HTStrings;
begin
  ListeImmo := HTStringList.Create;
  Nature := GetField('I_NATUREIMMO');
  ListeImmo.Add(GetField('I_IMMO'));
  if (Nature = 'PRO') or (Nature = 'FI') then
     IntegrationEcritures(toDotation, ListeImmo, FALSE, TRUE)
  else if (Nature = 'CB') or (Nature = 'LOC') then
     IntegrationEcritures(toEcheance, ListeImmo, FALSE, TRUE);
  ListeImmo.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : REGROUPEMENT Caractéristiques du regroupement de l'immo courante
Suite  ....... : Message si le regroupement n'a pas de modèle
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnZoomInfoRegClick(Sender: TObject);
var GroupeImmo : string;
    T : TOB;
    ModeleExiste:boolean;
begin
  if not VarIsNull ( TheMulQ.FindField('I_GROUPEIMMO').AsVariant) then
     GroupeImmo := GetField('I_GROUPEIMMO')
  else GroupeImmo := '';

  if (GroupeImmo <> '') then
  begin
    // BTY 01/06 Afficher la fiche du regroupement modèle uniquement si les infos existent
    T := TOB.Create ('IMMO',nil,-1);
    try
     ModeleExiste := (T.SelectDB('"&#@MODELE_'+ GroupeImmo +'"', nil, True));
    finally
     T.Free;
    end;
    if not ModeleExiste then
       PGIInfo(HM[4], RechDom('AMREGROUPEMENT', GroupeImmo, False)) //PGIBox (HM[18])
    else
       AmLanceFiche_FicheRegroupement ('&#@MODELE_' + GroupeImmo);
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 06/02/2006
Modifié le ... :   /  /
Description .. : REGROUPEMENT Détail du regroupement de l'immo courante
Suite ........ : (liste des immos regroupées)
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnZoomDetailRegClick(Sender: TObject);
begin
  AMLanceFiche_ListeDesImmobilisations ( '', False, taConsult, GetField('I_GROUPEIMMO'));
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ACCES Liste opérations en cours (VISU et SUPPRESSION)
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnZoomOperationsClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');

  // ajout mbo pour la suppression de l'opération de remplacement sur l'immo d'origine
  if ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+stCode+'" AND (I_OPERATION="X") AND (I_OPEREMPL = "X")') then
  begin
     if (not SupRemplacement(stCode)) then
        OperationsEnCours(GetField('I_IMMO'), GetField('I_LIBELLE'), false, TaConsult)
     else
        OperationsEnCours(GetField('I_IMMO'), GetField('I_LIBELLE'), false, fAction);
  end else
     OperationsEnCours(GetField('I_IMMO'), GetField('I_LIBELLE'), false, fAction);

  // Rafraichir au cas où opération supprimée
  RafraichirListeImmo (stCode);
  AfficheOpeEnCours (stCode);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : VISU Plan d'amortissement
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnZoomPlanClick(Sender: TObject);
var
  QTmp: TQuery;
  PlanAmor: TPlanAmort;
begin
  inherited;
  QTmp := OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="' + GetField('I_IMMO') +
    '"', true);
  if not QTmp.Eof then
  begin
    PlanAmor := TPlanAmort.Create(true); // := CreePlan(true);
    try
      PlanAmor.Charge(QTmp);
      PlanAmor.Recupere(QTmp.FindField('I_IMMO').AsString,
        QTmp.FindField('I_PLANACTIF').AsString);
      AffichePlanAmortissement(PlanAmor);

    finally
      PlanAmor.free; //Detruit;
    end;
  end;
  Ferme(QTmp);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 22/09/2006
Modifié le ... :
Description .. : VISU suramortissement = plan d'amortissement prime
Suite ........ : d'équipement
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnZoomPrimeClick(Sender: TObject);
var
  QTmp: TQuery;
  PlanAmor: TPlanAmort;
begin
  inherited;
  QTmp := OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="' + GetField('I_IMMO') +
    '"', true);
  if not QTmp.Eof then
  begin
    PlanAmor := TPlanAmort.Create(true); // := CreePlan(true);
    try
      PlanAmor.Charge(QTmp);
      PlanAmor.Recupere(QTmp.FindField('I_IMMO').AsString,
                        QTmp.FindField('I_PLANACTIF').AsString);
{$IFNDEF CMPGIS35}
      AffichePlanPrime(PlanAmor);
{$ENDIF}
    finally
      PlanAmor.free; //Detruit;
    end;
  end;
  Ferme(QTmp);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 22/09/2006
Modifié le ... :
Description .. : VISU plan d'amortissement subvention
Suite ........ : d'équipement
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnZoomSBVClick(Sender: TObject);
var
  QTmp: TQuery;
  PlanAmor: TPlanAmort;
begin
  inherited;
  QTmp := OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="' + GetField('I_IMMO') +
    '"', true);
  if not QTmp.Eof then
  begin
    PlanAmor := TPlanAmort.Create(true); // := CreePlan(true);
    try
      PlanAmor.Charge(QTmp);
      PlanAmor.Recupere(QTmp.FindField('I_IMMO').AsString,
                        QTmp.FindField('I_PLANACTIF').AsString);
{$IFNDEF CMPGIS35}
      AffichePlanSBV(PlanAmor);
{$ENDIF}
    finally
      PlanAmor.free; //Detruit;
    end;
  end;
  Ferme(QTmp);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ACTION Mutation
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnMutationClick(Sender: TObject);
var
  mr: integer;
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  mr := mrYes;
  if (FListe.NbSelected <= 1) and (not FListe.AllSelected) then
     mr := ExecuteMutation(GetField('I_IMMO'))
  else
     ExecuteOperationSerie(toMutation);
  if mr = mrYes then
     begin
     RafraichirListeImmo (stCode);
     AfficheOpeEnCours (stCode);
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ACTION Changement de méthode
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.ChangementMethodeClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  if AmLanceFiche_FicheChgtMethode (GetField('I_IMMO')) = mrYes then
     begin
     RafraichirListeImmo (stCode);
     AfficheOpeEnCours (stCode);
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 24/04/2006
Modifié le ... :   /  /
Description .. : fq 17923
Suite ........ : ACTION appel modification de la date de mise en service
Mots clefs ... :
*****************************************************************}

procedure TOF_AMLISTE.ModifServiceClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  if AmLanceFiche_AMMODIFSERVICE (GetField('I_IMMO')) = mrYes then
     begin
     RafraichirListeImmo (stCode);
     AfficheOpeEnCours (stCode);
     end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 30/08/2006
Modifié le ... :   /  /
Description .. : ACTION Lancement de la fonction Saisie d'une prime d'équipement
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnPrimeClick(Sender: TObject);
var
  stCode: string;
begin
{$IFNDEF CMPGIS35}
  stCode := GetField('I_IMMO');
  if AmLanceFiche_AMPRIME (stCode) = mrYes then
     begin
     RafraichirListeImmo (stCode);
     AfficheOpeEnCours (stCode);
     end;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 06/10/2006
Modifié le ... :   /  /
Description .. : ACTION Lancement de la fonction Saisie d'une subvention d'équipement
Mots clefs ... :
*****************************************************************}

procedure TOF_AMLISTE.OnSaisieSBVClick(Sender: TObject);
var
  stCode: string;
begin
{$IFNDEF CMPGIS35}
  stCode := GetField('I_IMMO');
 if AmLanceFiche_AMSUBVENTION (stCode) = mrYes then
    begin
     RafraichirListeImmo (stCode);
     AfficheOpeEnCours (stCode);
    end;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 16/04/2007
Modifié le ... :   /  /
Description .. : Lancement de remplacement de composant
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnSaisieRemplacementClick(Sender: TObject);
var
  stCode: string;
begin
{$IFNDEF CMPGIS35}
  stCode := GetField('I_IMMO');
  if AmLanceFiche_ficheRemplacement (stCode) = mrYes then
  begin
    RafraichirListeImmo (stCode);
    AfficheOpeEnCours (stCode);
 end;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/10/2006
Modifié le ... :   /  /
Description .. : ACTION  Réduction de la prime d'équipement
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnReductionPrimeClick(Sender: TObject);
var
  stCode: string;
begin
{$IFNDEF CMPGIS35}
  stCode := GetField('I_IMMO');
  if AMLanceFiche_AMReducPrime (stCode) = mrYes then
     begin
     RafraichirListeImmo (stCode);
     AfficheOpeEnCours (stCode);
     end;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 03/11/2006
Modifié le ... :   /  /
Description .. : ACTION  Réduction de la subvention d'investissement
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnReductionSubventionClick(Sender: TObject);
var
  stCode: string;
begin
{$IFNDEF CMPGIS35}
  stCode := GetField('I_IMMO');
  if AMLanceFiche_AMReducSubvention (stCode) = mrYes then
     begin
     RafraichirListeImmo (stCode);
     AfficheOpeEnCours (stCode);
     end;
{$ENDIF}
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ACTION Change ETABLISSEMENT
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnChangeEtablissementClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  if (FListe.NbSelected <= 1) and (not FListe.AllSelected) then
     ExecuteEtablissement(stCode)
  else
     ExecuteOperationSerie(toChanEtabl);

  // Réactualiser, l'immo courante est perdue
  RafraichirListeImmo (stCode);
  AfficheOpeEnCours (stCode);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ACTION Annulation de la dernière opération
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnAnnulationOperationClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  OperationsEnCours(GetField('I_IMMO'), GetField('I_LIBELLE'), True, fAction);

  RafraichirListeImmo (stCode);
  AfficheOpeEnCours (stCode);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ACTION Sortie d'une immo ou d'un ensemble d'immos
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnCessionClick(Sender: TObject);
  function SortirToutLeRegroupement ( stCode, stGroupe : string) : boolean;
  var stMessage : string;
  begin
    // On se positionne sur un enregistrement de la sélection
    if (fListe.NbSelected >= 1) then fListe.GotoLeBookmark(0);
    if ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO<>"'+stCode+'" AND I_GROUPEIMMO<>"" AND I_GROUPEIMMO="'+stGroupe+'" AND I_OPECESSION<>"X"') then
    begin
      { Plusieurs immobilisations dans le regroupement }
      stMessage :=
        TraduireMemoire('Cette immobilisation est associée au regroupement ');
      stMessage := stMessage + RechDom('AMREGROUPEMENT', stGroupe, False);
      stMessage := stMessage + '.#10#13' +
        TraduireMemoire('La sortie concerne-t-elle l''ensemble du regroupement ?');
      Result := (PGIAsk(stMessage) = mrYes);
    end else Result := False;
  end;

var
  mr, i: integer;
  stCode: string;
  TListeAll : TOB;
  stGroupeImmo: string;
begin
  stCode := GetField('I_IMMO');
  if not VarIsNull ( TheMulQ.FindField('I_GROUPEIMMO').AsVariant) then
    stGroupeImmo := GetField('I_GROUPEIMMO')
  else stGroupeImmo := '';
  //mr := mrYes;
  begin
    if fListe.AllSelected then
    begin
      TListeAll := TOB.Create ('', nil, -1);
      try
        RecupereAllSelect(TListeAll);
        if (TListeAll.Detail.Count > 0) then
        begin
          if SortirToutLeRegroupement (stCode, stGroupeImmo) then ExecuteCessionRegroupement(TListeAll.Detail[0].GetValue('I_GROUPEIMMO'))
          else
          begin
            for i := 0 to TListeAll.Detail.Count - 1 do
            begin
              mr := ExecuteCession(TListeAll.Detail[i].GetValue('I_IMMO'));
              if mr <> mrYes then break;
            end;
          end;
        end;
        fListe.AllSelected := False;
      finally
        FreeAndNil (TListeAll);
      end;
    end
    else if (fListe.NbSelected >= 1) then
    begin
      if SortirToutLeRegroupement (stCode, stGroupeImmo) then ExecuteCessionRegroupement(GetField('I_GROUPEIMMO'))
      else
      begin
        for i := 0 to fListe.NbSelected - 1 do
        begin
          fListe.GotoLeBookmark(i);
          mr := ExecuteCession(GetField('I_IMMO'));
          if mr <> mrYes then
            break;
        end;
      end;
      fListe.ClearSelected;
    end;
  end;

  // Toujours rafraichir, GotoLeBookmark repositionne => immo courante perdue
  RafraichirListeImmo (stCode);
  AfficheOpeEnCours (stCode);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ACTION Dépréciation d'actif . En multisélection :
Suite ........ : - si aucune immo dans un regroupement =>dépr. individuelle
Suite ........ : - si toutes dans le même regroupement sans dépr. précédente
Suite ........ :   dans l'exo => dépr. d'ensemble après question
Suite ........ : - autres cas option fermée
Suite ........ : En monosélection :
Suite ........ : - si immo sans regroupement => dépr. individuelle
Suite ........ : - si immo ds un regroupement => dépr. d'ensemble après question
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnDepreciationActifClick(Sender: TObject);

var
  i : integer;
  stCode: string;
  TListeAll : TOB;
  stGroupeImmo: string;
  mr : string;
  EnCoursTraitable : boolean;
//  stRevisionEco : double;  message compile 7.10 06/06
begin
  stCode := GetField('I_IMMO');
  if not VarIsNull ( TheMulQ.FindField('I_GROUPEIMMO').AsVariant) then
    stGroupeImmo := GetField('I_GROUPEIMMO')
  else stGroupeImmo := '';
  //mr := IntToStr(mrYes);
//  stRevisionECo := 0;      message compile 7.10 06/06

  begin
  if fListe.AllSelected then
    begin
      TListeAll := TOB.Create ('', nil, -1);
      try
        // Option Dépréciation ouverte donc immos de même regroupement
        RecupereAllSelect(TListeAll);
        if (TListeAll.Detail.Count > 0) then
        begin
              // On se positionne sur un enregistrement de la sélection
              if (fListe.NbSelected >= 1) then fListe.GotoLeBookmark(0);

              if RegroupementADeprecier (stCode,stGroupeImmo) then
              begin
                // Dépréciation du regroupement si pas déjà déprécié dans l'exo
                If not RegroupementDejaDeprecie (stGroupeImmo) then
                begin
                AMLanceFiche_DeprecGroupe(TListeAll.Detail[0].GetValue('I_GROUPEIMMO'));
                end;
              end
              else
              begin
                // FQ 18386 Ne pas boucler si une immo a une date début amortissement ECO > date fin exo en cours
                if ImmosATraiter (stCode, TListeALL, toDeprec, EnCoursTraitable) then
                   // Toutes les immos sélectionnées sont traitables
                   // On boucle sur les dépréciations individuelles
                   for i := 0 to TListeAll.Detail.Count - 1 do
                   begin
                    // Dépréciations individuelles
                    mr := AMLanceFiche_DepreciationActif(TListeAll.Detail[i].GetValue('I_IMMO'));
                    if mr <> IntToStr (mrYes) then break;
                   end
                else
                begin
                  // Toutes les immos sélectionnées ne sont pas traitables
                  PGIBox(HM[5], HM2[9]);
                  // Si immo courante traitable, faire l'opération sur elle seule
                  if EnCoursTraitable then
                     AMLanceFiche_DepreciationActif(TListeAll.Detail[0].GetValue('I_IMMO'));
                end;
              end;
        end;
        fListe.AllSelected := False;
      finally
        FreeAndNil (TListeAll);
      end;

    end
    else if (fListe.NbSelected >= 1) then
    begin
       // Option Dépréciation ouverte donc immos de même regroupement
       // On se positionne sur un enregistrement de la sélection
       if (fListe.NbSelected >= 1) then fListe.GotoLeBookmark(0);

       if RegroupementADeprecier (stCode,stGroupeImmo) then
       begin
           // Dépréciation du regroupement si pas déjà déprécié dans l'exo
            If not RegroupementDejaDeprecie (stGroupeImmo) then
            begin
              // Dépréciation du regroupement
              AMLanceFiche_DeprecGroupe(stGroupeImmo);
            end;
       end
       else
       begin
           // FQ 18386 Ne pas boucler si une immo a une date début amortissement ECO > date fin exo en cours
           // FQ 18386  Ne pas boucler si une immo est déjà dépréciée dans l'exercice
           // FQ 18386  Ne pas boucler si une immo est sortie
           if ImmosATraiter (stCode, nil, toDeprec, EnCoursTraitable) then
              // Toutes les immos sélectionnées sont traitables
              // On boucle sur les opérations individuelles
              for i := 0 to fListe.NbSelected - 1 do
              begin
                fListe.GotoLeBookmark(i);
                mr := AMLanceFiche_DepreciationActif(GetField('I_IMMO'));
                if mr <> IntToStr(mrYes) then break;
              end
           else
           begin
              // Cause regroupement qui ouvre l'opération, on peut avoir ici
              // le cas d'une sélection unitaire => message circonstancié
              if fListe.NbSelected = 1 then
              begin
                 // Attention GetField ('I_REVISIONECO') répond Null (pas dans AMLISTEIMMO)
                 if ExisteSQL ('SELECT I_IMMO, I_REVISIONECO FROM IMMO WHERE I_IMMO="'+
                                GetField('I_IMMO') + '" AND I_REVISIONECO<>0') then
                     PGIBox(HM[6], HM2[9])                 // déjà dépréciée
                 // Attention GetField('I_QUANTITE') OK car appartient à la DBliste AMLISTEIMMO
                 else if GetField('I_QUANTITE') = 0  then  // sortie
                     PGIBox(HM[7], HM2[9])
                 // Attention GetField ('I_DATEDEBECO') répond Null (pas dans AMLISTEIMMO)
                 // 7.10 06/06
                 else if ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+ GetField('I_IMMO') +
                                    '" AND I_DATEDEBECO>"' + UsDateTime (VHImmo^.Encours.Fin) + '"' ) then
                     PGIBox(HM[8], HM2[9]);
                 end
              else
              begin
                // Toutes les immos sélectionnées ne sont pas traitables
                PGIBox(HM[5], HM2[9]);
                // Si immo courante traitable, faire l'opération sur elle seule
                if EnCoursTraitable then
                   AMLanceFiche_DepreciationActif(GetField('I_IMMO'));
              end;
           end;
        end;

      //end;
      fListe.ClearSelected;
   end;
  end;

  // Toujours rafraichir, GotoLeBookmark repositionne => immo courante perdue
  RafraichirListeImmo (stCode);
  AfficheOpeEnCours (stCode);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ACTION Révision du plan d'amortissement
Mots clefs ... :
*****************************************************************}
{procedure TOF_AMLISTE.OnDureeMethodeClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  if ExecuteChangePlan(GetField('I_IMMO'), 1) = mrYes then
     begin
     RafraichirListeImmo (stCode);
     AfficheOpeEnCours (stCode);
     end;
end; }
// 04/06 FQ 17451
procedure TOF_AMLISTE.OnDureeMethodeClick(Sender: TObject);
begin
  OperationSerieEclatRevExcMba (toRevision);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 13/04/2006
Modifié le ... :   /  /
Description .. : FQ 17451
Suite ........ : Exécution de l'opération individuelle passé en entrée
Suite ........ : FQ 18393 Boucler sur la même date d'opération
Suite ........ : => paramètre supplémentaire
Mots clefs ... :
*****************************************************************}
function TOF_AMLISTE.ExecuteOperationEclatRevExcMba (TypeOpe : TypeOperation; CodeI : string; var Param:string):integer;
var CodeParam :string;
begin
     result := mrYes;
     // FQ 18393
     CodeParam := CodeI + ';' + Param;

     case TypeOpe of
     // Eclatement
     toEclat : begin
               result := ExecuteEclatement (CodeParam);            //(CodeI);
               end;
     // Révision plan
     toRevision : begin
                result := ExecuteChangePlan (CodeParam, 1);        //(CodeI, 1);
                end;
     // Saisie d'exceptionnel
     toExceptionnel : begin
                      result := ExecuteChangePlan (CodeParam, 2);  //(CodeI, 2);
                      end;
     // Modif bases
     toModifBases : begin
                    result := ExecuteModificationBases (CodeParam); //(CodeI);
                    end;
     end;

     // FQ 18393
     if result = mrYes then Param := CodeParam;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 13/04/2006
Modifié le ... :   /  /
Description .. : FQ 17451 Opérations ouvertes au traitement en série :
Suite ........ : révision plan
Suite ........ : saisie exceptionnel
Suite ........ : modif bases
Suite ........ : éclatement
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OperationSerieEclatRevExcMba (TypeOpe : TypeOperation);

var
  i : integer;
  stCode, Titre : string;
  TListeAll : TOB;
  mr: integer;
  EnCoursTraitable : boolean;
  Param : string;
begin

  stCode := GetField('I_IMMO');
  // FQ 18394 2e paramètre : le n° d'ordre en série
  // => passer date opération = iDate1900 pour décryptage correct par ReadTokenSt
  //Param := '';
  Param := DateToStr(iDate1900) + ';' + IntToStr(TrouveNumeroOrdreSerieLogSuivant);


  case TypeOpe of
       toEclat : begin
                Titre := 'Eclatement';
                 end;
       toRevision : begin
                    Titre := 'Révision du plan d''amortissement';
                    end;
       toExceptionnel : begin
                        Titre := 'Eléments exceptionnels';
                        end;
       toModifBases : begin
                      Titre := 'Modification des bases';
                      end;
  end;

  if fListe.AllSelected then
  begin
      TListeAll := TOB.Create ('', nil, -1);
      try
        RecupereAllSelect(TListeAll);
        if (TListeAll.Detail.Count > 0) then
        begin
            // On se positionne sur un enregistrement de la sélection
            if (fListe.NbSelected >= 1) then fListe.GotoLeBookmark(0);

            if ImmosATraiter (stCode, TListeAll, TypeOpe, EnCoursTraitable) then
            // Toutes les immos sélectionnées peuvent être traitées
            // On boucle sur les opérations individuelles
            begin

              // // FQ 18394 Ajouter le n° d'ordre en série
              // if fListe.NbSelected > 1 then
              //    Param := Param + IntToStr(TrouveNumeroOrdreSerieLogSuivant);

               for i := 0 to TListeAll.Detail.Count - 1 do
               begin
                    mr := ExecuteOperationEclatRevExcMba (TypeOpe, TListeAll.Detail[i].GetValue('I_IMMO'), Param);
                    if   mr <> mrYes then break;
               end  ;
            end
            else
            begin
               // Toutes les immos sélectionnées ne sont pas traitables
               PGIBox(HM[5], Titre);
               // Si immo courante traitable, faire l'opération sur elle seule
               if EnCoursTraitable then
               begin
                  ExecuteOperationEclatRevExcMba(TypeOpe, TListeAll.Detail[0].GetValue('I_IMMO'), Param);
               end;
            end;
        end;
        fListe.AllSelected := False;
      finally
        FreeAndNil (TListeAll);
      end;

  end
  else if (fListe.NbSelected >= 1) then
  begin
      // On se positionne sur un enregistrement de la sélection
      if (fListe.NbSelected >= 1) then fListe.GotoLeBookmark(0);

      if ImmosATraiter (stCode, nil, TypeOpe, EnCoursTraitable) then
      // Toutes les immos sélectionnées sont traitables
      // On boucle sur les opérations individuelles
      begin
       // // FQ 18394 Ajouter le n° d'ordre en série
       // if fListe.NbSelected > 1 then
       //    Param := Param + IntToStr(TrouveNumeroOrdreSerieLogSuivant);

        for i := 0 to fListe.NbSelected - 1 do
        begin
              fListe.GotoLeBookmark(i);
              mr := ExecuteOperationEclatRevExcMba(TypeOpe, GetField('I_IMMO'), Param);
              if mr <> mrYes then break;
        end ;
      end
      else
      begin
        // Toutes les immos sélectionnées ne sont pas traitables
        PGIBox(HM[5], Titre);
        // Si immo courante traitable, faire l'opération sur elle seule
        if EnCoursTraitable then
           begin
           ExecuteOperationEclatRevExcMba(TypeOpe, GetField('I_IMMO'), Param);
           end;
      end;
      fListe.ClearSelected;
  end;

  // Toujours rafraichir, GotoLeBookmark repositionne => immo courante perdue
  RafraichirListeImmo (stCode);
  AfficheOpeEnCours (stCode);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... : 07/04/2006
Description .. : ACTION Eclatement
Suite ........ : Eclatement ouvert en série sur les immos sélectionnées
Mots clefs ... :
*****************************************************************}
{procedure TOF_AMLISTE.OnEclatementClick(Sender: TObject);
var
  mr: integer;
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  mr := ExecuteEclatement(GetField('I_IMMO'));
  if mr = mrYes then
     begin
     RafraichirListeImmo (stCode);
     AfficheOpeEnCours (stCode);
     end;
end;}

// 04/06 FQ 17518 Eclatement en série
{procedure TOF_AMLISTE.OnEclatementClick(Sender: TObject);

var
  i : integer;
  stCode: string;
  TListeAll : TOB;
  mr: integer;
  EnCoursEclatable : boolean;
begin
  stCode := GetField('I_IMMO');

  if fListe.AllSelected then
  begin
      TListeAll := TOB.Create ('', nil, -1);
      try
        RecupereAllSelect(TListeAll);
        if (TListeAll.Detail.Count > 0) then
        begin
            // On se positionne sur un enregistrement de la sélection
            if (fListe.NbSelected >= 1) then fListe.GotoLeBookmark(0);

            // 04/06 FQ 17451
            //if ImmosAEclater (stCode, TListeAll, EnCoursEclatable) then
            if ImmosATraiter (stCode, TListeAll, toEclat, EnCoursEclatable) then
            // Toutes les immos sélectionnées sont éclatables
               for i := 0 to TListeAll.Detail.Count - 1 do
               begin
                  // On boucle sur les éclatements individuels
                  mr := ExecuteEclatement(TListeAll.Detail[i].GetValue('I_IMMO'));
                  if mr <> mrYes then break;
               end
            else
            begin
               // Toutes les immos sélectionnées ne sont pas éclatables
               PGIBox(HM[5], 'Eclatement');
               // Si immo courante éclatable, l'éclater elle seule
               if EnCoursEclatable then
                  mr := ExecuteEclatement(TListeAll.Detail[0].GetValue('I_IMMO'));
            end;
        end;
        fListe.AllSelected := False;
      finally
        FreeAndNil (TListeAll);
      end;

  end
  else if (fListe.NbSelected >= 1) then
  begin
      // On se positionne sur un enregistrement de la sélection
      if (fListe.NbSelected >= 1) then fListe.GotoLeBookmark(0);

      // 04/06 FQ 17451
      //if ImmosAEclater (stCode, nil, EnCoursEclatable) then
      if ImmosATraiter (stCode, nil, toEclat, EnCoursEclatable) then
      // Toutes les immos sélectionnées sont éclatables
        for i := 0 to fListe.NbSelected - 1 do
        begin
            // On boucle sur les éclatements individuels
            fListe.GotoLeBookmark(i);
            mr := ExecuteEclatement(GetField('I_IMMO'));
            if mr <> mrYes then   break;
        end
      else
      begin
        // Toutes les immos sélectionnées ne sont pas éclatables
        PGIBox(HM[5],'Eclatement');
        // Si immo courante éclatable, l'éclater elle seule
        if EnCoursEclatable then
           mr := ExecuteEclatement(GetField('I_IMMO'));
      end;
      fListe.ClearSelected;
  end;

  // Toujours rafraichir, GotoLeBookmark repositionne => immo courante perdue
  RafraichirListeImmo (stCode);
  AfficheOpeEnCours (stCode);
end;  }
// 04/06 FQ 17451
procedure TOF_AMLISTE.OnEclatementClick(Sender: TObject);
begin
  OperationSerieEclatRevExcMba (toEclat);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 07/04/2006
Modifié le ... : 14/06/2006
Description .. : On vérifie que toutes les immos de la sélection
Suite ........ : peuvent subir l'opération passée en entrée
Suite ........ : FQ 18386 Ajout traitement dépréciation de plusieurs immos
Mots clefs ... :
*****************************************************************}
function TOF_AMLISTE.ImmosATraiter (stCode:string; TListeAll:Tob; TypeOpe : TypeOperation; var EnCoursATraiter:boolean): boolean;
var i : integer;
    AcquisitionSelectee, bDepotGar : boolean;
    DepSelectee : boolean;
    NatureSelectee : string;
    bDatePBSelectee, bExcepSelectee : boolean;
    bDejaPrimeSBVSelectee, bRevisionPlanSelectee : boolean;
begin
  result := True;
  EnCoursATraiter := result;

  if fListe.AllSelected then
  begin

    for i := 0 to TListeAll.Detail.Count - 1 do
    begin
      // Chercher dans la sélection si existe une immo ne pouvant pas subir TypeOpe

      NatureSelectee := TListeAll.Detail[i].GetValue('I_NATUREIMMO');
      AcquisitionSelectee := (TListeAll.Detail[i].GetValue('I_DATEPIECEA') >= VHImmo^.Encours.Deb)
                      and (TListeAll.Detail[i].GetValue('I_DATEPIECEA') <= VHImmo^.Encours.Fin);
      bDepotGar := (NatureSelectee='FI') and (TListeAll.Detail[i].GetValue('I_IMMOLIE')<>'');
      // Champ I_REVISIONECO dans la TOB TListeAll.Detail[i]
      DepSelectee := (TListeAll.Detail[i].GetValue('I_REVISIONECO') <> 0 );
      // FQ 18385 + 18386 Champ I_DATEDEBECO dans la TOB
      bDatePBSelectee := (TListeAll.Detail[i].GetValue('I_DATEDEBECO') > VHImmo^.Encours.Fin );
      // FQ 18435 Champ I_MONTANTEXC dans la TOB !!
      bExcepSelectee := (TListeAll.Detail[i].GetValue('I_MONTANTEXC') <> 0 );
      // FQ 19260
      bDejaPrimeSBVSelectee := ((TListeAll.Detail[i].GetValue('I_SBVPRI') <> 0) or
                                (TListeAll.Detail[i].GetValue('I_SBVPRIC') <> 0) or
                                (TListeAll.Detail[i].GetValue('I_SBVMT') <> 0) or
                                (TListeAll.Detail[i].GetValue('I_SBVMTC') <> 0) );
      bRevisionPlanSelectee := True;
      if ((TListeAll.Detail[i].GetValue('I_METHODEECO')='NAM') and (bDejaPrimeSBVSelectee)) then
           bRevisionPlanSelectee := False;

      case TypeOpe of
      // Eclatement
      toEclat : begin
               if (NatureSelectee = 'CB') or (NatureSelectee = 'LOC') or
                  (bDepotGar) or (TListeAll.Detail[i].GetValue('I_METHODEECO') = 'VAR') or
                  (AcquisitionSelectee) or (TListeAll.Detail[i].GetValue('I_QUANTITE')=0) or
                  (TListeAll.Detail[i].GetValue('I_OPELEVEEOPTION') = 'X') or
                  (DepSelectee) or (bExcepSelectee) then    // FQ 18435
                      result := false;
               // Récup situation de l'immo courante
               if (TListeAll.Detail[i].GetValue('I_IMMO') = stCode) then
                   EnCoursATraiter := not ( (NatureSelectee = 'CB') or
                                            (NatureSelectee = 'LOC') or
                                            (bDepotGar) or
                                            (TListeAll.Detail[i].GetValue('I_METHODEECO')='VAR') or
                                            (AcquisitionSelectee) or
                                            (TListeAll.Detail[i].GetValue('I_QUANTITE')=0) or
                                            (TListeAll.Detail[i].GetValue('I_OPELEVEEOPTION')='X') or
                                            (DepSelectee) or (bExcepSelectee)) ;  // FQ 18435
                end;
      // Révision plan d'amortissement
      toRevision : begin
                   if (NatureSelectee <> 'PRO') or (AcquisitionSelectee) or
                      (TListeAll.Detail[i].GetValue('I_QUANTITE') = 0) or
                      (TListeAll.Detail[i].GetValue('I_OPECHANGEPLAN') = 'X') or
                      (TListeAll.Detail[i].GetValue('I_METHODEECO')= 'VAR') or
                      (not bRevisionPlanSelectee) then
                         result := false;
                   // Récup situation de l'immo courante
                   if (TListeAll.Detail[i].GetValue('I_IMMO') = stCode) then
                       EnCoursATraiter := not ( (NatureSelectee <> 'PRO') or
                                                (AcquisitionSelectee) or
                                                (TListeAll.Detail[i].GetValue('I_QUANTITE')=0) or
                                                (TListeAll.Detail[i].GetValue('I_OPECHANGEPLAN')='X') or
                                                (TListeAll.Detail[i].GetValue('I_METHODEECO')='VAR') or
                                                (not bRevisionPlanSelectee) );
                   end;
      // Saisie d'exceptionnel
      toExceptionnel : begin
                     if (NatureSelectee <> 'PRO') or
                        (TListeAll.Detail[i].GetValue('I_QUANTITE') = 0) or
                        (TListeAll.Detail[i].GetValue('I_OPECHANGEPLAN') = 'X') or
                        (TListeAll.Detail[i].GetValue('I_METHODEECO')='NAM') or (DepSelectee) //then
                         or (bDatePBSelectee) then  // FQ 18385
                         result := false;
                     // Récup situation de l'immo courante
                     if (TListeAll.Detail[i].GetValue('I_IMMO') = stCode) then
                        EnCoursATraiter := not ( (NatureSelectee <> 'PRO') or
                                                 (TListeAll.Detail[i].GetValue('I_QUANTITE')=0) or
                                                 (TListeAll.Detail[i].GetValue('I_OPECHANGEPLAN')='X') or
                                                 (TListeAll.Detail[i].GetValue('I_METHODEECO')='NAM') or
                                                 (DepSelectee) or (bDatePBSelectee) ); // FQ 18385 );
                       end;
      // Modif bases
      toModifBases : begin
                     if (NatureSelectee = 'LOC') or (NatureSelectee = 'FI') or
                        (AcquisitionSelectee) or
                        (TListeAll.Detail[i].GetValue('I_QUANTITE') = 0)  or
                        (TListeAll.Detail[i].GetValue('I_OPELEVEEOPTION') = 'X') or
                        (TListeAll.Detail[i].GetValue('I_METHODEECO')='VAR') or
                        (DepSelectee) or
                        (TListeAll.Detail[i].GetValue('I_METHODEECO')='NAM') then // FQ 19015
                         result := false;
                     // Récup situation de l'immo courante
                     if (TListeAll.Detail[i].GetValue('I_IMMO') = stCode) then
                        EnCoursATraiter := not ( (NatureSelectee = 'LOC') or
                                                 (NatureSelectee = 'FI') or
                                                 (AcquisitionSelectee) or
                                                 (TListeAll.Detail[i].GetValue('I_QUANTITE')=0)  or
                                                 (TListeAll.Detail[i].GetValue('I_OPELEVEEOPTION')='X') or
                                                 (TListeAll.Detail[i].GetValue('I_METHODEECO')='VAR') or
                                                 (DepSelectee) or
                                                 (TListeAll.Detail[i].GetValue('I_METHODEECO')='NAM') );
                     end;

      // FQ 18386 Dépréciation d'actif
      toDeprec : begin
                 if (bDatePBSelectee) or (DepSelectee) or
                 (TListeAll.Detail[i].GetValue('I_QUANTITE')=0) then
                    result := false;
                 // Récup situation de l'immo courante
                 if (TListeAll.Detail[i].GetValue('I_IMMO') = stCode) then
                     EnCoursATraiter := not (bDatePBSelectee or DepSelectee or
                                            (TListeAll.Detail[i].GetValue('I_QUANTITE')=0) );
                 end;
      end;  //fin CASE
    end;

  end
  else
  begin

    for i := 0 to FListe.NbSelected - 1 do
    begin
      FListe.GotoLeBookmark(i);

      // Chercher dans la sélection si existe une immo non éclatable
      NatureSelectee := GetField('I_NATUREIMMO');
      AcquisitionSelectee := (GetField('I_DATEPIECEA') >= VHImmo^.Encours.Deb)
                         and (GetField('I_DATEPIECEA') <= VHImmo^.Encours.Fin);
      bDepotGar := (NatureSelectee = 'FI') and (GetField('I_IMMOLIE') <> '');
      // Champ I_REVISIONECO pas accessible directement (absent de AMLISTEIMMO/AMLISTEHISTO)
      DepSelectee := ExisteSQL ('SELECT I_IMMO, I_REVISIONECO FROM IMMO WHERE I_IMMO="'+
                                GetField('I_IMMO') + '" AND I_REVISIONECO<>0');
      // FQ 18385
      bDatePBSelectee := ExisteSQL
       ('SELECT I_IMMO,I_DATEDEBECO FROM IMMO WHERE I_IMMO="'+ GetField('I_IMMO') + '" AND I_DATEDEBECO>"' +
       UsDateTime (VHImmo^.Encours.Fin) + '"');
      // FQ 18435 Pas d'éclatement si immo présente de l'exceptionnel
      bExcepSelectee := ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+ GetField('I_IMMO')+
                                   '" AND I_MONTANTEXC<>0');
      // FQ 19260
      bDejaPrimeSBVSelectee := ExisteSQL (
      'SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+GetField('I_IMMO')+'" AND ((I_SBVPRI<>0) or (I_SBVPRIC <>0) or (I_SBVMT<>0) or (I_SBVMTC <>0))');
      bRevisionPlanSelectee := True;
      if ((GetField('I_METHODEECO')='NAM') and (bDejaPrimeSBVSelectee)) then
           bRevisionPlanSelectee := False;

      case TypeOpe of
      // Eclatement
      toEclat : begin
                if (NatureSelectee = 'CB') or (NatureSelectee = 'LOC') or
                   (bDepotGar) or (GetField('I_METHODEECO')='VAR') or
                   (AcquisitionSelectee) or
                   (GetField('I_QUANTITE') = 0) or
                   (GetField('I_OPELEVEEOPTION') = 'X') or
                   (DepSelectee) or (bExcepSelectee) then      // FQ 18435
                     result := false;
                // Récup situation de l'immo courante
                if (GetField('I_IMMO') = stCode) then
                    EnCoursATraiter := not ( (NatureSelectee = 'CB') or
                                             (NatureSelectee = 'LOC') or
                                             (bDepotGar) or (GetField('I_METHODEECO')='VAR') or
                                             (AcquisitionSelectee) or
                                             (GetField('I_QUANTITE')=0) or
                                             (GetField('I_OPELEVEEOPTION') = 'X') or
                                             (DepSelectee) or (bExcepSelectee) );   // FQ 18435
                end;
      // Révision plan d'amortissement
      toRevision : begin
                   if (NatureSelectee <> 'PRO') or (AcquisitionSelectee) or
                      (GetField('I_QUANTITE') = 0) or
                      (GetField('I_OPECHANGEPLAN') = 'X') or
                      (GetField('I_METHODEECO')= 'VAR') or
                      (not bRevisionPlanSelectee) then
                         result := false;
                   // Récup situation de l'immo courante
                   if (GetField('I_IMMO') = stCode) then
                       EnCoursATraiter := not ( (NatureSelectee <> 'PRO') or
                                                (AcquisitionSelectee) or
                                                (GetField('I_QUANTITE')=0) or
                                                (GetField('I_OPECHANGEPLAN')='X') or
                                                (GetField('I_METHODEECO')='VAR') or
                                                (not bRevisionPlanSelectee) );
                   end;
      // Saisie d'exceptionnel
      toExceptionnel : begin
                     if (NatureSelectee <> 'PRO') or
                        (GetField('I_QUANTITE') = 0) or
                        (GetField('I_OPECHANGEPLAN') = 'X') or
                        (GetField('I_METHODEECO')='NAM') or (DepSelectee) //
                         or (bDatePBSelectee) then  // FQ 18385
                         result := false;
                     // Récup situation de l'immo courante
                     if (GetField('I_IMMO') = stCode) then
                        EnCoursATraiter := not ( (NatureSelectee <> 'PRO') or
                                                 (GetField('I_QUANTITE')=0) or
                                                 (GetField('I_OPECHANGEPLAN')='X') or
                                                 (GetField('I_METHODEECO')='NAM') or
                                                 (DepSelectee) or (bDatePBSelectee) ); // FQ 18385
                       end;
      // Modif bases
      toModifBases : begin
                     if (NatureSelectee = 'LOC') or (NatureSelectee = 'FI') or
                        (AcquisitionSelectee) or
                        (GetField('I_QUANTITE') = 0)  or
                        (GetField('I_OPELEVEEOPTION') = 'X') or
                        (GetField('I_METHODEECO')='VAR') or
                        (DepSelectee) or
                        (GetField('I_METHODEECO')='NAM') then  // FQ 19015
                         result := false;
                     // Récup situation de l'immo courante
                     if (GetField('I_IMMO') = stCode) then
                        EnCoursATraiter := not ( (NatureSelectee = 'LOC') or
                                                 (NatureSelectee = 'FI') or
                                                 (AcquisitionSelectee) or
                                                 (GetField('I_QUANTITE')=0)  or
                                                 (GetField('I_OPELEVEEOPTION')='X') or
                                                 (GetField('I_METHODEECO')='VAR') or
                                                 (DepSelectee) or
                                                 (GetField('I_METHODEECO')='NAM') );// FQ 19015
                     end;

      // FQ 18386 Dépéciation d'actif
      toDeprec : begin
                 if (bDatePBSelectee or DepSelectee or
                    (GetField('I_QUANTITE') = 0)) then
                    result := false;
                 // Récup situation de l'immo courante
                 if (GetField('I_IMMO') = stCode) then
                     EnCoursATraiter := not (bDatePBSelectee or DepSelectee or
                                            (GetField('I_QUANTITE') = 0) );
                 end;
      end; // fin CASE
    end;
  end;

end;


{function TOF_AMLISTE.ImmosAEclater (stCode:string; TListeAll:Tob; var EnCoursEclatable:boolean): boolean;
var i : integer;
    AcquisitionSelectee, bDepotGar : boolean;
    bDepSelectee : boolean;
    NatureSelectee : string;
begin
  result := True;
  EnCoursEclatable := result;

  if fListe.AllSelected then
  begin

    for i := 0 to TListeAll.Detail.Count - 1 do
    begin
      // Chercher dans la sélection si existe une immo non éclatable

      NatureSelectee := TListeAll.Detail[i].GetValue('I_NATUREIMMO');
      AcquisitionSelectee := (TListeAll.Detail[i].GetValue('I_DATEPIECEA') >= VHImmo^.Encours.Deb)
                      and (TListeAll.Detail[i].GetValue('I_DATEPIECEA') <= VHImmo^.Encours.Fin);
      bDepotGar := (NatureSelectee='FI') and (TListeAll.Detail[i].GetValue('I_IMMOLIE')<>'');
      // Champ I_REVISIONECO dans la TOB TListeAll.Detail[i]
      bDepSelectee := (TListeAll.Detail[i].GetValue('I_REVISIONECO') <> 0 );

      if (NatureSelectee = 'CB') or (NatureSelectee = 'LOC') or
         (bDepotGar) or (TListeAll.Detail[i].GetValue('I_METHODEECO') = 'VAR') or
         (AcquisitionSelectee) or (TListeAll.Detail[i].GetValue('I_QUANTITE')=0) or
         (TListeAll.Detail[i].GetValue('I_OPELEVEEOPTION') = 'X') or (bDepSelectee) then
               result := false;

      // Récup situation de l'immo courante
      if (TListeAll.Detail[i].GetValue('I_IMMO') = stCode) then
          EnCoursEclatable := not ( (NatureSelectee = 'CB') or (NatureSelectee = 'LOC') or
                                    (bDepotGar) or
                                    (TListeAll.Detail[i].GetValue('I_METHODEECO') = 'VAR') or
                                    (AcquisitionSelectee) or
                                    (TListeAll.Detail[i].GetValue('I_QUANTITE')=0) or
                                    (TListeAll.Detail[i].GetValue('I_OPELEVEEOPTION')='X') or
                                    (bDepSelectee) ) ;
    end;

  end
  else
  begin

    for i := 0 to FListe.NbSelected - 1 do
    begin
      FListe.GotoLeBookmark(i);

      // Chercher dans la sélection si existe une immo non éclatable
      NatureSelectee := GetField('I_NATUREIMMO');
      AcquisitionSelectee := (GetField('I_DATEPIECEA') >= VHImmo^.Encours.Deb)
                         and (GetField('I_DATEPIECEA') <= VHImmo^.Encours.Fin);
      bDepotGar := (NatureSelectee = 'FI') and (GetField('I_IMMOLIE') <> '');
      // Champ I_REVISIONECO pas accessible directement (absent de AMLISTEIMMO/AMLISTEHISTO)
      bDepSelectee := ExisteSQL ('SELECT I_IMMO, I_REVISIONECO FROM IMMO WHERE I_IMMO="'+
                                GetField('I_IMMO') + '" AND I_REVISIONECO<>0');

      if (NatureSelectee = 'CB') or (NatureSelectee = 'LOC') or
         (bDepotGar or ((GetField('I_METHODEECO')='VAR') or
          AcquisitionSelectee)) or  ((GetField('I_QUANTITE') = 0)) or
         ((GetField('I_OPELEVEEOPTION') = 'X')) or (bDepSelectee) then
             result := false;

      // Récup situation de l'immo courante
      if (GetField('I_IMMO') = stCode) then
          EnCoursEclatable := not ((NatureSelectee = 'CB') or (NatureSelectee = 'LOC') or
                                   (bDepotGar or ((GetField('I_METHODEECO')='VAR') or
                                    AcquisitionSelectee)) or  ((GetField('I_QUANTITE')=0)) or
                                   ((GetField('I_OPELEVEEOPTION') = 'X')) or (bDepSelectee) );
    end;
  end;

end; }


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ACTION Saisie d'éléments exceptionnels
Mots clefs ... :
*****************************************************************}
{procedure TOF_AMLISTE.OnElementExceptionnelClick(Sender: TObject);
var
  stCode: string;
begin
  inherited;
  stCode := GetField('I_IMMO');
  if ExecuteChangePlan(GetField('I_IMMO'), 2) = mrYes then
     begin
     RafraichirListeImmo (stCode);
     AfficheOpeEnCours (stCode);
     end;
end; }
// 04/06 FQ 17451
procedure TOF_AMLISTE.OnElementExceptionnelClick(Sender: TObject);
begin
  inherited;
  OperationSerieEclatRevExcMba (toExceptionnel);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ACTION Levée d'option
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnLeveeOptionClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  if ExecuteLeveeOption(GetField('I_IMMO')) = mrYes then
     begin
     RafraichirListeImmo (stCode);
     AfficheOpeEnCours (stCode);
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ACTION Modification des bases
Mots clefs ... :
*****************************************************************}
{procedure TOF_AMLISTE.OnModifBasesClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  if (ExecuteModificationBases(GetField('I_IMMO')) = mrYes) then
     begin
     RafraichirListeImmo (stCode);
     AfficheOpeEnCours (stCode);
     end;
end; }
// 04/06 FQ 17451
procedure TOF_AMLISTE.OnModifBasesClick(Sender: TObject);
begin
  OperationSerieEclatRevExcMba (toModifBases)
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ACTION Change REGROUPEMENT
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnActionRegroupementClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  if (FListe.NbSelected <= 1) and (not FListe.AllSelected) then
     ExecuteRegroupement(stCode)
  else
     ExecuteOperationSerie(toRegroupement);

  RafraichirListeImmo (stCode);
  AfficheOpeEnCours (stCode);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ACTION Change LIEU Géographique
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnChangeLieuClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  if (FListe.NbSelected <= 1) and (not FListe.AllSelected) then
     ExecuteLieu(stCode)
  else
     ExecuteOperationSerie(toChanLieu);

  RechargeListeChoix;
  RafraichirListeImmo (stCode);
  AfficheOpeEnCours (stCode);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Recharger les listes des COMBOS susceptibles d'avoir évolué
Suite ........ : si l'utilisateur a créé de nouveaux items
Suite ........ : Peut se produire en retour de fiche immo ou de changement de lieu géo
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.RechargeListeChoix;
var
  OldVal: string;
begin
  { Lieu géographique }
  OldVal := GetControlText('I_LIEUGEO');
  THMultiValComboBox(GetControl('I_LIEUGEO')).Reload;
  SetControlText('I_LIEUGEO', OldVal);
  { Regroupement }
  OldVal := GetControlText('I_GROUPEIMMO');
  THMultiValComboBox(GetControl('I_GROUPEIMMO')).Reload;
  SetControlText('I_GROUPEIMMO', OldVal);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Charger en TOB toutes les immos de la multisélection
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.RecupereAllSelect(T: TOB);
var
 stWhere : string;
 Q : TQuery;
begin
  if T = nil then exit;

  // BTY FQ 17356 Récupérer la liste sélectionnée et pas toute la liste des immos
  // RecupWhereCritere reprend le XX_WHERE par défaut pas le critère complémentaire en cours
  //Q := OpenSql ('SELECT I_IMMO,I_GROUPEIMMO,I_NATUREIMMO,I_QUANTITE FROM IMMO ' + stWhere, True);

  stWhere := RecupWhereCritere(TFMul(Ecran).Pages);
  RecupereCritereXX_WHERE (nil);
  // 06/06 Attention I_MONTANTEXC au lieu de I_MONTANEXC pour la FQ 18435
  // 06/06 Champ pour tester l'éclatement
  // 06/06 FQ 18385 Champ pour test saisie exceptionnels
  // 04/06 FQ 17518 Récupérer certains champs pour test immo à éclater
  // 04/06 FQ 17451 Récupérer certains champs pour test révision plan/exceptionnel/modif bases
  //Q := OpenSql ('SELECT I_IMMO,I_GROUPEIMMO,I_NATUREIMMO,I_QUANTITE FROM IMMO ' +
  //Q := OpenSql (
  //     'SELECT I_IMMO,I_GROUPEIMMO,I_NATUREIMMO,I_QUANTITE, I_DATEPIECEA, I_IMMOLIE, I_OPELEVEEOPTION, I_REVISIONECO FROM IMMO ' +
  //              stWhere + ' AND ' + GetControlText('XX_WHERE'), True);
  //Q := OpenSql (
  //'SELECT I_IMMO,I_GROUPEIMMO,I_NATUREIMMO,I_QUANTITE, I_DATEPIECEA, I_IMMOLIE, I_OPELEVEEOPTION, I_REVISIONECO, I_METHODEECO, I_OPECHANGEPLAN FROM IMMO ' +
  // FQ 19260
  //'SELECT I_IMMO,I_GROUPEIMMO,I_NATUREIMMO,I_QUANTITE, I_DATEPIECEA, I_IMMOLIE, I_OPELEVEEOPTION, I_REVISIONECO, I_METHODEECO, I_OPECHANGEPLAN, I_DATEDEBECO, I_MONTANTEXC FROM IMMO ' +
  Q := OpenSql (
  'SELECT I_IMMO,I_GROUPEIMMO,I_NATUREIMMO,I_QUANTITE, I_DATEPIECEA, I_IMMOLIE, I_OPELEVEEOPTION, I_REVISIONECO, I_METHODEECO, I_OPECHANGEPLAN, I_DATEDEBECO, I_MONTANTEXC, ' +
                'I_SBVPRI, I_SBVPRIC, I_SBVMT, I_SBVMTC FROM IMMO ' +
                stWhere + ' AND ' + GetControlText('XX_WHERE'), True);
  // Remettre la valeur par défaut
  SetControlText('XX_WHERE', fWhereOrig);
  try
  if not Q.Eof then
      T.LoadDetailDB ('', '', '', Q, False);
  finally
    Ferme (Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Se repositionner sur l'immo courante
Suite ........ : Fait suite à une relance de la requête de multisélection
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.SelectionEnregistrement(stCode: string);
//var Param:array of variant;
begin
{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}
  if (stCode = '') then exit;
//  SetLength(Param,1);
//  Param[0] := stCode;
//TheMulQ.Locate('I_IMMO', [stCode], []);             ereur à la compile
//TheMulQ.Locate('I_IMMO', Param, []);                erreur à l'exécution
//TheMulQ.Locate('I_IMMO', vararrayof([stCode]), [])  erreur à l'exécution

TheMulQ.Locate('I_IMMO', stCode, []) ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Teste si plusieurs immos sélectionnées
Mots clefs ... :
*****************************************************************}
function TOF_AMLISTE.SelectionMultiple: boolean;
begin
  Result := (((fListe.AllSelected) and (TheMulQ.RecordCount > 1))
              or (fListe.NbSelected > 1));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Teste si une seule immo sélectionnée
Mots clefs ... :
*****************************************************************}
function TOF_AMLISTE.SelectionUnique: boolean;
begin
  Result := (((fListe.AllSelected) and (TheMulQ.RecordCount = 1))
             or (fListe.NbSelected = 1));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Fonction de suppression d'une immo
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.SuppressionFicheImmo;
begin
  SupprimeFicheImmo(FCodeImmo)
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : F9 Actionner l'exécution de la requête de multisélection
Suite ........ :
Mots clefs ... :
*****************************************************************}
{
procedure TOF_AMLISTE.OnBChercheClick(Sender: TObject);
begin
  // Recomposer le critère complémentaire
  RecupereCritereXX_WHERE(nil);

  // Relancer la requête
  TFMul(ECRAN).BChercheClick(nil);

  // Remettre le critère par défaut
  SetControlText('XX_WHERE', fWhereOrig);

  // Màj opérations sur l'immo en cours
  AfficheOpeEnCours (nil);
end;
}

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Rafraichir la liste des immos après une action,
Suite ........ : ceci conformément à la sélection des immos
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.RafraichirListeImmo(stCode: string);
begin
  // Relancer la requête de multisélection
  TFMul(ECRAN).BChercheClick(nil);
  // Se repositionner sur l'enreg courant
  SelectionEnregistrement(stCode);
  // Revoir les boutons au cas où liste devenue vide
  EnableBoutonsAction(fAction);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... : 11/01/2006
Description .. : ONGLET STANDARDS
Suite ........ : Liste des comptes en Lookup sur un champ Compte
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnCompteelipsisClick(Sender: TObject);
var
  stWhere: string;
begin
  if THEdit(Sender).Name = 'I_COMPTEIMMO' then
      stWhere := 'G_GENERAL LIKE "2%"'
  else if THEdit(Sender).Name = 'I_COMPTELIE' then
      stWhere := 'G_GENERAL LIKE "612%" OR G_GENERAL LIKE "613%"';
  LookupList(TControl(Sender), '', 'GENERAUX', 'G_GENERAL', 'G_LIBELLE', stWhere,
             'G_GENERAL', True, 0);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... : 11/01/2006
Description .. : ONGLET AMORTISSEMENTS
Suite ........ : Top d'amortissement dérogatoire
Suite ........ : mbo 29/05/2007 - cette coche conditionne l'accès à la coche gestion fiscale
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnAmortDerogClick(Sender: TObject);
begin
   case GetCheckBoxState('bAmortDerog') of
   cbChecked :
     begin
     // 05/06 FQ 17257 Dérogatoire <=> plan fiscal
     //fWhereDerog := ' AND I_METHODEFISC <>"" AND (I_METHODEFISC <> I_METHODEECO' +
     //               '  OR I_TAUXFISC <> I_TAUXECO OR I_DUREEFISC <> I_DUREEFISC)' ;
     fWhereDerog := ' AND I_METHODEFISC <>""';
     SetControlEnabled('BGestionFiscale', true);
     end;
   cbUnChecked :
     begin
     // 05/06 FQ 17257 Pas de dérogatoire <=> pas de fiscal
     //fWhereDerog := ' AND I_METHODEFISC <>"" AND I_METHODEFISC = I_METHODEECO' +
     //               ' AND I_TAUXFISC = I_TAUXECO AND I_DUREEFISC = I_DUREEFISC' ;
     fWhereDerog := ' AND I_METHODEFISC =""';
     SetControlEnabled('bGestionFiscale', False);
     end;
   else // cbGrayed <=> toutes les immos
     begin
     fWhereDerog := '';
     SetControlEnabled('bGestionFiscale', True);
     end;
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 16/03/2007
Modifié le ... :   /  /
Description .. : nouvelle coche : gestion fiscale
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnGestionFiscaleClick(Sender: TObject);
begin
   case GetCheckBoxState('bGestionFiscale') of
   cbChecked :
     begin
     fWhereGestionFiscale := ' AND I_NONDED = "X"';
     end;
   cbUnChecked :
     begin
     fWhereGestionFiscale := ' AND I_METHODEFISC <>"" AND I_NONDED ="-"';
     end;
   else // cbGrayed <=> toutes les immos
     begin
     fWhereGestionFiscale := '';
     end;
   end;
end;



{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 29/06/2006
Modifié le ... :
Description .. : FQ 18465 Action de changement du filtre (complète celui du MUL)
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnChangeFFiltres(Sender: TObject);
begin
  // Evènement de l'ancêtre MUL
  TFMUL(Ecran).FFiltresChange(Sender);

  // Si un exercice relatif + les dates amortissement ou achat sont dans le filtre :
  // correction pour appliquer JJ MM des dates du filtre si possible
  AMAppliquerExoRelatifToDates ( THValComboBox(GetControl('FExercice2')).Value,
                                 THEdit(GetControl('I_DATEDEBECO')),
                                 THEdit(GetControl('I_DATEDEBECO_')),
                                 True);
  AMAppliquerExoRelatifToDates ( THValComboBox(GetControl('FExercice2')).Value,
                                 THEdit(GetControl('I_DATEPIECEA')),
                                 THEdit(GetControl('I_DATEPIECEA_')),
                                 True);
  // 09/06 Revoir les établissements au cas où restauration incompatible avec les restrictions utilisateur
  ApresChangeFFiltres;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... : 12/01/2006
Description .. : ONGLET COMPLEMENTS
Suite ........ : Choix exclusif d'un exercice
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnFExercice2Change(Sender: TObject);
var
  FExercice2: THValComboBox;
begin
  FExercice2 := THValComboBox(GetControl('FExercice2'));
  if FExercice2.Value <> '' then
    begin

    // FQ 18465 Exercice Relatif à convertir en dates réelles
    // ImExoToDates(FExercice2.Value, THEdit(GetControl('I_DATEPIECEA')),
    //                 THEdit(GetControl('I_DATEPIECEA_')));
    // FQ 18414 i_dateamort changé au profit de i_datedebeco
    //ImExoToDates(FExercice2.Value, THEdit(GetControl('I_DATEAMORT')),
    //             THEdit(GetControl('I_DATEAMORT_')));
    //    ImExoToDates(FExercice2.Value, THEdit(GetControl('I_DATEDEBECO')),
    //                 THEdit(GetControl('I_DATEDEBECO_')));
   AMAppliquerExoRelatifToDates ( FExercice2.Value, THEdit(GetControl('I_DATEDEBECO')),
                                  THEdit(GetControl('I_DATEDEBECO_')));
   AMAppliquerExoRelatifToDates ( FExercice2.Value, THEdit(GetControl('I_DATEPIECEA')),
                                  THEdit(GetControl('I_DATEPIECEA_')));

    end
  else
    begin

    // FQ 18414
    // SetControlText('I_DATEAMORT', StDate1900);
    // SetControlText('I_DATEAMORT_', StDate2099);
    SetControlText('I_DATEDEBECO', StDate1900);
    SetControlText('I_DATEDEBECO_', StDate2099);

    SetControlText('I_DATEPIECEA', StDate1900);
    SetControlText('I_DATEPIECEA_', StDate2099);
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... : 11/01/2006
Description .. : ONGLET COMPLEMENTS
Suite ........ : Top de mise en service différée
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnAmortDiffereClick(Sender: TObject);
begin
  if GetCheckBoxState('bAmortDiffere') = cbChecked then
     fWhereAmortDiffere := ' AND I_DATEPIECEA <> I_DATEAMORT'
  else
     fWhereAmortDiffere := '';
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ONGLET AMORTISSEMENTS
Suite ........ : Top Base amortissements ECO différene de la valeur théorique
Suite ........ : (montant HT + TVA récupérable - TVA Récupérée)
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnBaseEcoThClick(Sender: TObject);
begin
   case GetCheckBoxState('bBaseEcoTh') of
   cbChecked :
     begin
     fWhereBaseEcoTh := ' AND I_METHODEECO <> "" AND I_BASEECO <> (I_MONTANTHT' +
                        ' + I_TVARECUPERABLE - I_TVARECUPEREE)';
     end;
   cbUnChecked :
     begin
     fWhereBaseEcoTh := ' AND I_METHODEECO <> "" AND I_BASEECO = (I_MONTANTHT' +
                        ' + I_TVARECUPERABLE - I_TVARECUPEREE)';
     end;
   else   // cbGrayed <=> toutes les immos
     begin
     fWhereBaseECoTh:= '';
     end;
   end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ONGLET OPERATIONS
Suite ........ : Multichoix de codes opération
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnOperationChange(Sender: TObject);
var St, TexteSelection : string;
    sOR, sSEP: string;

begin
  sSEP := ' AND (';
  sOR := ' OR';
  TexteSelection := THMultiValComboBox(GetControl('Operation')).Text;
  St := ReadTokenSt(TexteSelection);

  if St <> '' then
     begin
     fWhereOpe := '';
     while (St <> '') do
        begin
        if St ='MUT' then
           begin
           fWhereOpe := fWhereOpe + sSep + ' I_OPEMUTATION="X"';
           sSep := sOR;
           end;
        if St ='LEV' then
           begin
           fWhereOpe := fWhereOpe + sSep + ' I_OPELEVEEOPTION = "X"';
           sSep := sOR;
           end;
        if St ='CES' then
           begin
           fWhereOpe := fWhereOpe + sSep + ' I_OPECESSION = "X"';
           sSep := sOR;
           end;
        if St ='DPR' then
           begin
           fWhereOpe := fWhereOpe + sSep + ' I_OPEDEPREC = "X"';
           sSep := sOR;
           end;
        if St ='MBA' then
           begin
           fWhereOpe := fWhereOpe + sSep + ' I_OPEMODIFBASES = "X"';
           sSep := sOR;
           end;
        if St ='LIE' then
           begin
           fWhereOpe := fWhereOpe + sSep + ' I_OPELIEUGEO = "X"';
           sSep := sOR;
           end;
        if St ='ETA' then
           begin
           fWhereOpe := fWhereOpe + sSep + ' I_OPEETABLISSEMENT = "X"';
           sSep := sOR;
           end;
        if St ='ECL' then
           begin
           fWhereOpe := fWhereOpe + sSep + ' I_OPEECLATEMENT="X"';
           sSep := sOR;
           end;
        if St ='CHP' then
           begin
           fWhereOpe := fWhereOpe + sSep + ' I_OPECHANGEPLAN = "X"';
           sSep := sOR;
           end;
        // BTY 04/06 FQ 17516
        if St ='REG' then
           begin
           fWhereOpe := fWhereOpe + sSep + ' I_OPEREG = "X"';
           sSep := sOR;
           end;

        St := ReadTokenSt(TexteSelection);
        end;
        if fWhereOpe <> '' then  fWhereOpe := fWhereOpe + ')'
  end
  else
  begin
    fWhereOpe := '';
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ONGLET OPERATIONS
Suite ........ : Top d'immo sans aucune opération
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnSansOpeclick(Sender: TObject);
begin
     case GetCheckBoxState('bSansOpe') of
     cbChecked :
       begin
       fWhereSansOpe := ' AND I_OPERATION <> "X"';
       SetControlEnabled('Operation', False);
       SetControlEnabled('t_Operation', False);
       THMultiValComboBox(GetControl('OPERATION')).Value := '';
       end;
     cbUnChecked :
       begin
       fWhereSansOpe := ' AND I_OPERATION = "X"';
       SetControlEnabled('Operation', True);
       SetControlEnabled('t_Operation', True);
       end;
     else   // cbGrayed <=> toutes les immos
       begin
       fWhereSansOpe:= '';
       SetControlEnabled('Operation', True);
       SetControlEnabled('t_Operation', True);
       end;
     end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ONGLET OPERATIONS
Suite ........ : Top d'immo avec réintégration fiscale
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnReintegrationClick(Sender: TObject);
begin
     if GetCheckBoxState('bReintegration') = cbChecked then
       fWhereReintegration := ' AND I_REINTEGRATION<>0.0'
     else
       fWhereReintegration := '';  // ' AND I_REINTEGRATION=0.0'
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ONGLET OPERATIONS
Suite ........ : Top d'immo avec quote-part personnelle
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnQuotePartClick(Sender: TObject);
begin
     if GetCheckBoxState('bQuotePart') = cbChecked then
       fWhereQuotePart := ' AND I_QUOTEPART<>0'
     else
       fWhereQuotePart := '';  // ' AND I_QUOTEPART=0'
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 21/06/2006
Modifié le ... :   /  /
Description .. : ONGLET OPERATIONS
Suite ........ : Top d'immo sur utilisation d'une DPI 7.10 06/06 DPI
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnDPIclick(Sender: TObject);
begin
     case GetCheckBoxState('bDPI') of
     cbChecked :
       fWhereCheckDPI := ' AND I_DPI = "X"';
     cbUnChecked :
       fWhereCheckDPI := ' AND I_DPI <> "X"';
     else   // cbGrayed <=> toutes les immos
       fWhereCheckDPI := '';
     end;
end;
//

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2006
Modifié le ... :   /  /
Description .. : ONGLET OPERATIONS
Suite ........ : Top d'immo sur présence d'une prime d'équipement
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnPRIclick(Sender: TObject);
begin
     case GetCheckBoxState('bPRI') of
     cbChecked :
       fWhereCheckPRI := ' AND (I_SBVPRI <> 0.0 OR I_SBVPRIC <> 0.0)';
     cbUnChecked :
       fWhereCheckPRI := ' AND I_SBVPRI = 0 AND I_SBVPRIC = 0';
     else   // cbGrayed <=> toutes les immos
       fWhereCheckPRI := '';
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 18/10/2006
Modifié le ... :   /  /
Description .. : ONGLET OPERATIONS
Suite ........ : Top d'immo sur présence d'une subvention d'investissement
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnbSBVclick(Sender: TObject);
begin
     case GetCheckBoxState('bSBV') of
     cbChecked :
       fWhereCheckSBV := ' AND (I_SBVMT <> 0.0 OR I_SBVMTC <> 0.0)';
     cbUnChecked :
       fWhereCheckSBV := ' AND I_SBVMT = 0 AND I_SBVMTC = 0';
     else   // cbGrayed <=> toutes les immos
       fWhereCheckSBV := '';
     end;
end;





{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ONGLET RECHERCHE
Suite ........ : Top Base TP différene de la valeur théorique
Suite ........ : (montantHT + TVA récupérable - TVA Récupérée)
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnBaseTPThClick(Sender: TObject);
begin
   case GetCheckBoxState('bBaseTPTh') of
   cbChecked :
     begin
     fWhereBaseTPTh := ' AND I_BASETAXEPRO <> 0 AND I_BASETAXEPRO <> (I_MONTANTHT' +
                        ' + I_TVARECUPERABLE - I_TVARECUPEREE)';
     end;
   cbUnChecked :
     begin
     fWhereBaseTPTh := ' AND I_BASETAXEPRO <> 0 AND I_BASETAXEPRO = (I_MONTANTHT' +
                        ' + I_TVARECUPERABLE - I_TVARECUPEREE)';
     end;
   else   // cbGrayed <=> toutes les immos
     begin
     fWhereBaseTPTh:= '';
     end;
   end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 23/06/2006
Modifié le ... :   /  /
Description .. : ONGLET RECHERCHE
Suite ........ : FQ 14133 Top Exigibilité de la TP au dégrèvement
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnExigibleTPClick(Sender: TObject);
begin
   case GetCheckBoxState('bExigibleTP') of
   cbChecked :
     fWhereExigibleTP := ' AND I_TABLE9 = "X"';
   cbUnChecked :
     fWhereExigibleTP := ' AND I_TABLE9 <> "X"';
   else   // cbGrayed <=> toutes les immos
     fWhereExigibleTP := '';
   end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : ONGLET CARACTERISTIQUES
Suite ........ : Top d'immo sans aucun regroupement
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnSansGroupeClick(Sender: TObject);
begin

   if fRegroupement = '' then
     // Pas de regroupement passé en entrée de la MUL
     case GetCheckBoxState('bSansGroupe') of
     cbChecked :
       begin
       fWhereSansGroupe := ' AND I_GROUPEIMMO = ""';
       SetControlEnabled('I_GROUPEIMMO', False);
       SetControlEnabled('tI_GROUPEIMMO', False);
       THMultiValComboBox(GetControl('I_GROUPEIMMO')).Value := '';
       //THMultiValComboBox(GetControl('I_GROUPEIMMO')).SelectAll;
       end;
     cbUnChecked :
       begin
       fWhereSansGroupe := ' AND I_GROUPEIMMO <> ""';
       SetControlEnabled('I_GROUPEIMMO', True);
       SetControlEnabled('tI_GROUPEIMMO', True);
       end;
     else   // cbGrayed <=> toutes les immos
       begin
       fWhereSansGroupe:= '';
       SetControlEnabled('I_GROUPEIMMO', True);
       SetControlEnabled('tI_GROUPEIMMO', True);
       end;
     end
   else
   begin
     // Regroupement passé en entrée de la MUL
     fWhereSansGroupe:= '';
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... : 18/01/2006
Description .. : ONGLET CARACTERISTIQUES
Suite ........ : LookUp sur l'organisme bancaire d'une immo en Crédit-bail ou
Suite ........ : location longue durée
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnOrganismeelipsisClick(Sender: TObject);
begin
{$IFDEF SERIE1}
  LookUpList(TControl(Sender), TraduireMemoire('Auxiliaire'), 'TIERS',
             'T_AUXILIAIRE', 'T_LIBELLE', 'T_NATUREAUXI="FOU"', 'T_AUXILIAIRE', True, 1);
{$ELSE}
  LookUpList(TControl(Sender), TraduireMemoire('Auxiliaire'), 'TIERS',
             'T_AUXILIAIRE', 'T_LIBELLE', 'T_NATUREAUXI="FOU"', 'T_AUXILIAIRE', True, 2);
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 13/09/2006
Modifié le ... :
Description .. : ONGLET CARACERISTIQUES
Suite ........ : FQ 16149 Si un filtre restaure un établissement qui ne correspond pas
Suite ........ : aux restrictions utilisateur, alimenter la COMBO Etablissements
Suite ........ : selon les restrictions et pas selon le filtre
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.ApresChangeFFiltres;
begin
{$IFNDEF SERIE1}
  ImControlEtab (THMultiValComboBox(GetControl('I_ETABLISSEMENT')) );
{$ENDIF}
end;



{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... : 18/01/2006
Description .. : ONGLET CHAMPS LIBRES
Suite ........ : 3 champs libres
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTE.OnChangeChampLibre(Sender: TObject);
begin
  if THValComboBox(Sender).Name = 'TABLELIBRE1' then
     SetControlText('I_TABLE0', GetControlText('TABLELIBRE1'))
  else if THValComboBox(Sender).Name = 'TABLELIBRE2' then
     SetControlText('I_TABLE1', GetControlText('TABLELIBRE2'))
  else if THValComboBox(Sender).Name = 'TABLELIBRE3' then
     SetControlText('I_TABLE2', GetControlText('TABLELIBRE3'));
end;



{////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////

     P R O C E D U R E S / F O N C T I O N S   P U B L I Q U E S

/////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////}



{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 15/09/2004
Modifié le ... :   /  /
Description .. : Message d'avertissement avant supression d'une
Suite ........ : immobilisation.
Mots clefs ... :
*****************************************************************}
function ConfirmationSuppressionImmo ( CodeImmo, EtatImmo : string ) : boolean;
begin
  if EtatImmo <> 'OUV' then
  begin
    Result := not ( PGIAsk(HM[3] ) <> mrYes ); // ( PGIAsk(HM[15] ) <> mrYes );
  end else Result := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 15/09/2004
Modifié le ... :   /  /
Description .. : Contrôle de la suppression d'une immobilisation
Mots clefs ... :
*****************************************************************}
function JaiLeDroitSupprimerImmo(CodeImmo, ChangeCode: string): boolean;
begin
  Result := False;
  (*
    CA - 24/03/2005 - Modifications conditions de suppression d'une immobilisation
    FQ 15471 et 15044
    Conditions de suppression :
    - Message avertissement si opération en cours hors acquisition, clôture ou passage à l'euro
    - Pas de suppression possible si éclatement
  *)
  if ChangeCode<>'' then
  begin
    PGIBox('Cette immobilisation a été créée suite à une opération. La suppression est impossible.');
  end else
  if ExisteSQL ('SELECT I_IMMO FROM IMMOLOG WHERE IL_IMMO="'+CodeImmo+'"'+
                  ' AND IL_TYPEOP="ECL"') then
  begin
    PGIBox('Cette immobilisation a subi un éclatement. La suppression est impossible.');
  end else
  begin
    if ExisteSQL ('SELECT I_IMMO FROM IMMOLOG WHERE IL_IMMO="'+CodeImmo+'"'+
                  ' AND IL_TYPEOP<>"ACQ" AND IL_TYPEOP<>"EUR" OR IL_TYPEOP<>"CLO"') then
    begin
      Result :=  (PGIAsk('Cette immobilisation a subi des opérations. Voulez-vous continuer la suppression ?')=mrYes);
    end else Result := True;
  end;
end;

procedure SupprimeFicheImmo(Code: string; bFiche : boolean = True );
var
  Q: TQuery;
{$IFDEF SERIE1}
{$ELSE}
  i: integer;
{$ENDIF}
begin
  if bFiche then
    if ExecuteSQL('DELETE FROM IMMO WHERE I_IMMO = "' + Code + '"') <> 1 then
      V_PGI.IoError := oeSaisie
    ELSE
      Begin
        //Tga 28/06/2006 Maj Immomvtd suppression de l'immobilisation
{$IFNDEF CMPGIS35}
        AM_MAJ_IMMOMVTD('S',Code,'',0);
{$ENDIF}
        V_PGI.IoError := oeOk;
      End;  

  if V_PGI.IoError = oeOk then
    ExecuteSQL('DELETE FROM IMMOAMOR WHERE IA_IMMO = "' + Code + '"');
  if V_PGI.IoError = oeOk then
    ExecuteSQL('DELETE FROM IMMOLOG WHERE IL_IMMO = "' + Code + '"');
  if V_PGI.IoError = oeOk then
    ExecuteSQL('DELETE FROM IMMOECHE WHERE IH_IMMO = "' + Code + '"');
  if V_PGI.IoError = oeOk then
    ExecuteSQL('DELETE FROM IMMOUO WHERE IUO_IMMO = "' + Code + '"');
{$IFDEF SERIE1}
  //XVI 10/01/2007 FQ 13735 début
  if V_PGI.IoError = oeOk then
      ExecuteSQL('DELETE FROM VENTIL Where V_NATURE="IMM" AND V_COMPTE="' + Code + '"');
  //XVI 10/01/2007 FQ 13735 fin
{$ELSE}
  if V_PGI.IoError = oeOk then
  begin
    for i := 1 to ImMaxAxe do
    begin
      ExecuteSQL('DELETE FROM VENTIL Where V_NATURE="IM' + IntToStr(i) +
        '" AND V_COMPTE="' + Code + '"');
    end;
  end;
{$ENDIF}
  if V_PGI.IoError = oeOk then
  begin
{$IFDEF SERIE1}
{$ELSE}
    // Suppression du lien avec l'écriture éventuelle rattachée à l'immobilisation
    ExecuteSQL('UPDATE ECRITURE SET E_IMMO="", E_NUMEROIMMO=0, E_DATEMODIF="' +
      UsTime(NowH) + '" WHERE E_IMMO="' + Code + '"');
{$ENDIF}
    // Mise à jour des liens éventuels avec l'immobilisation supprimée
    Q :=
      //OpenSQL('SELECT I_IMMO,I_IMMOLIE,I_IMMOLIEGAR FROM IMMO WHERE I_IMMOLIEGAR="'
      //+ Code + '"' +
      //'OR I_IMMOLIE="' + Code + '"', TRUE);
      // 05/06 FQ 17499
      OpenSQL('SELECT I_IMMO,I_IMMOLIE,I_IMMOLIEGAR,I_CHANGECODE FROM IMMO WHERE I_IMMOLIEGAR="'
      + Code + '"' +
      'OR I_IMMOLIE="' + Code + '"' + ' OR I_CHANGECODE="' + Code + '"', TRUE);
    if not Q.EOF then
    begin
      Q.First;
      while not Q.Eof do
      begin
        if Q.FindField('I_IMMOLIE').AsString = Code then
          ExecuteSQL('UPDATE IMMO SET I_IMMOLIE="" WHERE I_IMMO="' +
            Q.FindField('I_IMMO').AsString + '"');
        if Q.FindField('I_IMMOLIEGAR').AsString = Code then
          ExecuteSQL('UPDATE IMMO SET I_IMMOLIEGAR="",I_DATEDEPOTGAR="' +
            USDateTime(iDate1900) + '",I_DEPOTGARANTIE=0 WHERE I_IMMO="' +
            Q.FindField('I_IMMO').AsString + '"');
        // 05/06 FQ 17499
        if Q.FindField('I_CHANGECODE').AsString = Code then
          ExecuteSQL('UPDATE IMMO SET I_CHANGECODE="" WHERE I_IMMO="' +
            Q.FindField('I_IMMO').AsString + '"');
        //
        Q.Next;
      end;
    end;
    Ferme(Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 24/03/2005
Modifié le ... : 24/03/2005
Description .. : Fonction de contrôle préalable à la supression d'une
Suite ........ : immobilisation.
Suite ........ : Cette fonction affiche tous les messages de confirmation.
Suite ........ : et renvoie True si la suppression phyique doit être lancée
Suite ........ : modif mbo - 27/06/2006 - fq 14324 - ajout d'un paramÞtre d'appel
Suite ........ : si true : affichage du message de confirmation de l'immo
Suite ........ : (appel par la liste)
Suite ........ : si false : pas affichage du message de confirmation (appel
Suite ........ : par la fiche)
Mots clefs ... :
*****************************************************************}
function ValideSuppressionImmo ( stCode : string; Appel : boolean) : boolean;
var stCaption, origine : string;
begin
  Result := False;
  stCaption := TraduireMemoire('Immobilisation n°')+stCode;
  (*
    CA - 24/03/2005 - Modifications conditions de suppression d'une immobilisation
    FQ 15471 et 15044
    Conditions de suppression :
    - Message avertissement si opération en cours hors acquisition, clôture ou passage à l'euro
    - Pas de suppression possible si éclatement
  *)
  if ExisteSQL ('SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="'+stCode+'"'+
                  ' AND IL_TYPEOP="ECL"') then
  begin
    PGIBox('Cette immobilisation a subi un éclatement. La suppression est impossible.',stCaption);
  end else

  if (GetColonneSQL('IMMO','I_OPEREMPL','I_IMMO="'+stCode+'"') <> '-') then
  begin
    PGIBox('Cette immobilisation a subi un remplacement de composant. La suppression est impossible.',stCaption);
  end else

  // 06/07 i_string1 au lieu de i_remplace
  // ajout mbo du test sur i_remplace pour les immos remplaçantes importées - 30.08.07
  if (GetColonneSQL('IMMO','I_STRING1','I_IMMO="'+stCode+'"') <> '') and
     (GetColonneSQL('IMMO','I_REMPLACE','I_IMMO="'+stCode+'"') <> '') then
  begin
     origine := GetColonneSQL('IMMO','I_REMPLACE','I_IMMO="'+stCode+'"');
     //FQ 14429 - PGIBox('Cette immobilisation est issue d''un remplacement de composant. La suppression est impossible.',stCaption);
     PGIBox('Cette immobilisation résulte du remplacement de l''immobilisation '+ origine+'. La suppression est impossible.',stCaption);
  end else

  if (GetColonneSQL('IMMO','I_CHANGECODE','I_IMMO="'+stCode+'"') <> '') then
  begin
    origine := GetColonneSQL('IMMO','I_CHANGECODE','I_IMMO="'+stCode+'"');
    //FQ 14429 - PGIBox('Cette immobilisation a été créée suite à une opération. La suppression est impossible.',stCaption);
    PGIBox('Immobilisation résultant d''une opération (origine : '+ origine +') : suppression impossible.',stCaption);
  end else

  begin
    if ExisteSQL ('SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="'+stCode+'"'+
                  ' AND IL_TYPEOP<>"ACQ" AND IL_TYPEOP<>"EUR" AND IL_TYPEOP<>"CLO"') then
    begin
      Result :=  (PGIAsk('Cette immobilisation a subi des opérations. Voulez-vous continuer la suppression ?',stCaption)=mrYes);
    end else Result := True;
  end;

  if Result then
    // ajout mbo fq 14324 - modif du message pour harmonisation avec suppression dans la fiche
    //Result := (PGIAsk('Confirmez-vous la suppression de l''immobilisation ?',stCaption)=mrYes);
    if Appel = true then
       Result := (PGIAsk('Confirmez-vous la suppression de l''enregistrement ?',stCaption)=mrYes);
    //Result := (PGIAsk('Confirmez-vous la suppression de l''immobilisation ?',stCaption)=mrYes);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/10/2005
Modifié le ... :   /  /
Description .. : En CRC2002-10 Dépréciation d'un regroupement
Suite ........ : On teste si la dépréciation a déjà été faite dans l'exo pour
Suite ........ : tout le regroupement
Mots clefs ... :
*****************************************************************}
function RegroupementDejaDeprecie ( stGroupe : string) : boolean;
  var stMessDeja : string;
      Q: TQuery;

  begin
      Result := True;
      stMessDeja :=TraduireMemoire('Il n''est pas possible de pratiquer plusieurs dépréciations/reprises dans l''exercice.');
      Q := OpenSQL(
      'SELECT * FROM IMMO WHERE I_NATUREIMMO="PRO" AND I_GROUPEIMMO="' +stGroupe+'"', true);
      While (not Q.EOF and Result)  do
      begin
        // Test chaque immo du regroupement déjà dépréciée dans l'exo
        Result :=ExisteSQL (
                 'SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="'+Q.FindField('I_IMMO').AsString+
                 '" AND IL_TYPEOP="DPR"' + ' AND IL_DATEOP>="' +USDateTime(VHImmo^.Encours.Deb)+'"' );
        Q.next;
      end;
      ferme(Q);
      if Result then
      begin
        // Pas trouvé d'immo non dépréciée dans l'exo => tout le regroupement l'est
        //PGIBox(stMessDeja, Ecran.Caption);
        PGIBox(stMessDeja);
      end;
  end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/09/0005
Modifié le ... :   /  /
Description .. : En CRC 2002-10, Top de dépréciation d'un regroupement
Suite ........ : possible ou non
Mots clefs ... :
*****************************************************************}
function RegroupementADeprecier ( stCode,stGroupe : string) : boolean;
  var stMessage : string;
  begin
    Result := False;
    if (stGroupe <> '') then
    begin
      // Même regroupement non nul
      stMessage :=
        TraduireMemoire('Cette immobilisation est associée au regroupement ');
      stMessage := stMessage + RechDom('AMREGROUPEMENT', stGroupe, False);
      stMessage := stMessage + '.#10#13' +
        TraduireMemoire('La dépréciation concerne-t-elle l''ensemble du regroupement ?');
      // Test regroupement concerne plusieurs immos
      if ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO<>"'+stCode+'" AND I_GROUPEIMMO<>"" AND I_GROUPEIMMO="'+stGroupe+'"') then
         Result := (PGIAsk(stMessage) = mrYes);
    end;
end;





initialization
  registerclasses([TOF_AMLISTE]);
end.




