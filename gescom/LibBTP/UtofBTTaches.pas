{***********UNITE*************************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 14/12/2000
ModIfié le ... : 28/02/2001
Description .. : Source TOF de la TABLE : AFTACHES ()
Suite ........ :
Suite ........ :
Suite ........ :
Mots clefs ... : TOF;AFTACHES
*****************************************************************}
unit UtofBTTaches;

interface

uses
{$IFDEF VER150}
  Variants,
{$ENDIF}
  buttons
{$IFDEF EAGLCLIENT}
  ,eFiche
  ,MaineAGL
{$ELSE}
  ,Fiche
  ,db
{$IFNDEF DBXPRESS}
  ,dbtables {BDE}
{$ELSE}
  ,uDbxDataSet
{$ENDIF}
  ,FE_Main
{$ENDIF}
  ,Lookup
  ,StdCtrls
  ,Controls
  ,Classes
  ,forms
  ,sysutils
  ,ComCtrls
  ,HCtrls
  ,HPanel
  ,UTOFAFTRADUCCHAMPLIBRE
  ,AFPLanningCst
  ,UTob
  ,HEnt1 // TActionFiche
  ,HTB97 // TToolBarButton97
  ,HRichOle
  ,UAFO_REGLES
	,uEntCommun
  ,UtilTOBPiece
  ,UdateUtils;

type

  TOF_BTTACHES = class(TOF_AFTRADUCCHAMPLIBRE)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(StArgument: string); override;
    procedure OnClose; override;
    Procedure DefiniLibelle(TobAppel : Tob);

  private

    fTobTaches    : TOB; // ensemble des tâches de l'affaire (Affichage de la grille
    fTobDet       : TOB; // tache courante de l'affaire (tache sélectionnée)
    fTobSup       : TOB; // tâches supprimées
    fTobSupRes    : TOB; // ressources supprimées
    fTobModele    : TOB; // Modèle de taches
    fTobTacheJour : TOB;
    fTobTypeAction: TOB; // Type d'action pour une tache

    fATATermineChange   : Boolean;
    fBoNoSuppression    : Boolean; // supression interdite quand on ouvre depuis le planning
    fBoAFPLANDECHARGE   : Boolean; // paramsoc indiquant si on est en plan de charge
    fBoModifRegles      : Boolean;
    fBoClose            : Boolean; // fermeture de la fenetre

    fCurAffaire         : AffaireCourante; // tache en cours
    fUpdateRes          : TTabRes; // ensemble des ressources a reaffecter

    fStLibelleParDefaut : string;
    fStLesCols          : string; // colonnes des taches
    fStOpenNumTache     : string; // numero de la tache a l'ouverture
    fStAfficheAffaire   : string;
    fStLignePiece       : string;
    fStBtEtat           : String; //Type d'action sur tache
    fStTiers            : String; //Tiers associé à l'affaire
    fStAuxilliaire      : String; //compte auxilliaire du tiers de l'affaire

    vAFReglesTache      : TAFRegles;

    fGSTaches : THGrid; // grille de la liste des tâches
    fGSRes    : THGrid; // grille des ressources

    fPaDet    : THPanel; // panel contenant les onglets de détail de la tâche

    fAction   : TActionFiche; // action sur la fiche
    fStatut   : TActionFiche; // statut interne de la fiche (pour enregistrement)

    fRB_MODEGENE1       : TRadioButton;
    fRB_MODEGENE2       : TRadioButton;
    fRB_QUOTIDIENNE     : TRadioButton;
    fRB_NBINTERVENTION  : TRadioButton;
    fRB_HEBDOMADAIRE    : TRadioButton;
    fRB_ANNUELLE        : TRadioButton;
    fRB_MENSUELLE       : TRadioButton;
    fRB_MOISMETHODE1    : TRadioButton;
    fRB_MOISMETHODE2    : TRadioButton;
    fRB_MOISMETHODE3    : TRadioButton;

    fPC_FREQUENCE       : TPageControl;

    fTS_QUOTIDIENNE     : TTAbSheet;
    fTS_NBINTERVENTION  : TTAbSheet;
    fTS_HEBDOMADAIRE    : TTAbSheet;
    fTS_ANNUELLE        : TTAbSheet;
    fTS_MENSUELLE       : TTAbSheet;

    FCB_RAPPLANNING     : TcheckBox;

    fED_AFFAIRE         : THEdit;

    fDtDebutAffaire     : TDateTime;
    fDtFinAffaire       : TDateTime;
    fDtDebPeriod        : TDateTime;
    fDtFinPeriod        : TDateTime;

    // Gestion des grids
    // Grid des tâches ...
    procedure InitGrid;
    procedure InitColGrid;
    procedure FormatColGridTaches;

    // ressources
    procedure InitGridRes;
    procedure AddUpdateRessList(aCol, aRow: Integer);
    procedure RefreshGridRes;
    procedure UpdateGridRes(pBoModifNbRes: Boolean; pInNum: Integer);

    // gestion de l'onglet règles
    procedure fRB_MODEGENE1OnClick(Sender: TObject);
    procedure fRB_MODEGENE2OnClick(Sender: TObject);
    procedure fRB_QUOTIDIENNEOnClick(Sender: TObject);
    procedure fRB_NBINTERVENTIONOnClick(Sender: TObject);
    procedure fRB_HEBDOMADAIREOnClick(Sender: TObject);
    procedure fRB_ANNUELLEOnClick(Sender: TObject);
    procedure fRB_MENSUELLEOnClick(Sender: TObject);
    procedure BRESSOURCEOnClick(Sender: TObject);

    procedure BCALCULRAPRESOnClick(Sender: TObject);
    procedure CALCULRAFRES;
    procedure PAGETACHEChange(Sender: TObject);

    // Gestion du grid
    procedure fGSTachesRowEnter(SEnder: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure fGSTachesRowExit(SEnder: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);

    procedure fGSResCellEnter(SEnder: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure fGSResCellExit(SEnder: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure fGSResKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure fGSResElipsisClick(SEnder: TObject);

    procedure FormKeyDown(SEnder: TObject; var Key: Word; ShIft: TShIftState);

    // Gestion des boutons
    procedure bInsertOnClick(SEnder: TObject);
    procedure bInsertResOnClick(SEnder: TObject);

    procedure bDeleteOnClick(SEnder: TObject);
    procedure bDeleteResOnClick(SEnder: TObject);

    procedure bCompetArticleOnClick(SEnder: TObject);
    procedure bCompetRessourceOnClick(SEnder: TObject);
    procedure bRessourceDispoOnClick(SEnder: TObject);
    procedure bFermerOnClick(SEnder: TObject);
    procedure bMemoOnClick(SEnder: TObject);
    procedure MnContact_OnClick(SEnder: TObject);
    procedure MnAdressesInt_OnClick(SEnder: TObject);
    procedure mnGeneAffaire_OnClick(SEnder: TObject);
    procedure mnGeneTache_OnClick(SEnder: TObject);
    procedure mnPlanChargeQuantite_OnClick(SEnder: TObject);
    procedure mnPlanChargePrixRevient_OnClick(SEnder: TObject);
    procedure mnPlanChargePrixVente_OnClick(SEnder: TObject);
    procedure MnPlanChargeDispo_OnClick(SEnder: TObject);
    procedure mnPlanningViewer_OnClick(SEnder: TObject);
    procedure mnPlanningRes_OnClick(SEnder: TObject);

    procedure MnPlanningTache_OnClick(SEnder: TObject);
    procedure MnPlanningRessource_OnClick(SEnder: TObject);
    procedure MnPlanningOccupRess_OnClick(SEnder: TObject);
    procedure MnPlanningOccupRessDetaille_OnClick(SEnder: TObject);
    procedure MnDetailPlanning_OnClick(SEnder: TObject);
    procedure MnPlanningActivite_OnClick(SEnder: TObject);

    // Gestion de la tâche en cours
    //procedure ATA_PERIODICITEOnClick(SEnder: TObject);
    // tob... sur tâche en cours
    // Tâche créée à partir de modèles tâches ou de ligne pièce affaire
    //procedure TOBCopieModeleTache(fCurAffaire :AffaireCourante; FromTOB, ToTOB : TOB);

    //Procedures....
    procedure AffichageEntete;
    procedure AlimTob(NoJour: integer);
    procedure AppelContact(sender: Tobject);
    procedure ARTICLEOnElipsisClick(SEnder: TObject);
    procedure ArticleOnExit(Sender: TObject);
    procedure ATA_PUPROnExit(Sender: TObject);
    procedure ATA_PUVENTEHTOnExit(Sender: TObject);
    procedure ATA_QTEINITIALEOnExit(SEnder: TObject);
    procedure ATA_QTEINITPLAOnExit(SEnder: TObject);
    procedure ATA_RAPPLANNINGOnClick(Sender: TObject);
    procedure ATA_TERMINEonClick(Sender: TObject);
    procedure ATA_UNITETEMPS_OnChange(Sender: TObject);
    procedure ATA_UNITETEMPSOnChange(Sender: TObject);
    procedure bDupliqueOnClick(SEnder: TObject);
    procedure bInsertModeleOnClick(SEnder: TObject);
    procedure BIntervenantsOnClick(Sender: TObject);
    procedure BloquerTaches;
    procedure BTLIGNEOnClick(SEnder: TObject);
    procedure BtnSelectionOnClick(Sender: TObject);
    procedure BtnEffacerOnClick(Sender: TObject);
    procedure CacheZoneSurTypeArt;
    procedure CalculQuantitePla(pBoEcran: Boolean);
    procedure ControleLastDateGene(var pListeJours: TListDate; pLastDateGene: TDateTime; pBoAjout: Boolean);
    procedure EcranOnCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FonctionOnExit(Sender: TObject);
    procedure fRB_MOISMETHODE1OnClick(SEnder: TObject);
    procedure fRB_MOISMETHODE2OnClick(SEnder: TObject);
    procedure fRB_MOISMETHODE3OnClick(SEnder: TObject);
    procedure GenereEvenement(TobA: Tob; NoAdresse: String);
    procedure GenererListesAppel(pListeJours: TListDate; NumLigne: Integer; Duree: Double);
    procedure GestionEcran;
    procedure GetRegles;
    procedure InitNewRegles;
    procedure InitNewTache(TobTache: Tob);
    procedure InsererModele(TobDepart, TobArrivee: Tob);
    procedure InsertTache;
    procedure LectureContact(contact: string);
    procedure LectureTypeAction(TypeAction: string);
    procedure LoadAffaire(pInIndiceTob: Integer);
    procedure ModifZoneExit(Sender: TObject);
    procedure PutRegles(pStModele: string);
    procedure RaffraichirArticle;
    procedure RefreshAdresse;
    procedure RefreshGrid;
    procedure RefreshLastDateGene;
    procedure RefreshLignePiece;
    procedure RefreshQuantitePla;
    procedure ScreenToTobRes;
    procedure TobResToScreen;
    procedure TraiteLignePiece;
    procedure TraitementArgument(pStArgument: string);
    procedure TypeActionOnExit(Sender: TObject);
    procedure TypeArticleOnExit(Sender: TObject);
    procedure UniteOnEnter(SEnder: TObject);
    procedure UniteOnEnter2(SEnder: TObject);
    procedure ValiderLesTaches;
    procedure ValiderLesTachesJour;
    procedure ZonesLectureSeule;

    //Fonctions....
    function ArticleUnique: Boolean;
    function CalendrierDefautOk : Boolean;
    function ClientFerme: Boolean;
    function ControleEcran: boolean;
    function ControleRessource(var pRdPrevuRes: Double): Integer;
    function ControleZonesOblig: boolean;
    function GenereAppel: boolean;
    function FamilleArticleUnique: Boolean;
    function GenereEvt: boolean;
    function IndiceTobTacheCourante: Integer;
    function MajItems(pTob: Tob): Boolean;
    //function RechRessPrinc(Affaire: String): String;
    function RessourceSaisie: Boolean;
    function ScreenToTob: Boolean;
    function TacheHeure(pBoGlobal : Boolean) : Boolean;
    function TestCoherence(pBoGeneration: Boolean; pStTache: string): Boolean;
    function TestCoherenceQte(pBoGeneration: Boolean): boolean;
    function TobToScreen(Row: Integer): Integer;
    function Valider(pBoForce, pBoSupControl: Boolean): Boolean;

  end;

function AFLanceFicheBTTaches(Argument: string): string;

implementation

uses AppelsUtil,
     FactAdresse,
     UtofBTModeleTache,
     PlanUtil,
    {,Spin,
     LookUp,
     messages,
     UtilArticle,
     TiersUtil,
     AFPlanningGene,
     UtomAffTiers,
     UTOF}

  UtofRessource_Mul, // PL le 13/01/04 pour changement de place de la fonction DecodeRefPiece
  BTPUtil,
  HMsgBox,
  Dicobtp,
  UtilTaches,
  SaisUtil,
  AGLInit,
  Vierge,
  windows,
  AGLInitGC,
  AFUtilArticle,
  CalcOleGenericAff,
  AffaireUtil,
  paramsoc,
  menus,
  UtilRessource,
  GereTobInterne,
  EntGC,
  FactUtil,
  FactTOB,
  UtilConfid,
  UtilPGI,
  // A réintégrer plus tard si besoin est !!!!
  //UtilPlanning,
  //AFPlanning,
  //UtofAFPlanningViewer,
  //UtofAFPlanningRes,
  //UTOFAFPLANCHARGE,
  //UTofAFPLANNINGGENERER,
  //UTofAFLIGPLANNING_Mul,
  //UtofAFRessourceDispo,
  //UtofAFRessourceOccup,
  //WCOMPETRESRCE_TOF,
  //UtofAFPlanningCompar,
  //UAFO_ValorisationPrix,
  {$ifndef OGC}
  //UTOFAFCOMPETARTICLE,
  {$endif}
  UAFO_Ressource,
  ActiviteUtil
  ,DateUtils
  ,CbpMCD
  ,CbpEnumerator
  ;

const
  TexteMessage: array[1..63] of string = (
    {1}'Saisie par affaire sur code affaire non valide.',
    {2}'La mise à jour des tâches ne s''est pas effectuée correctement.',
    {3}'Adresse d''intervention',
    {4}'Il est interdit d''avoir 2 fois le même code ressource pour une tâche.',
    {5}'Le libellé de tâche est obligatoire.',
    {6}'La famille de tâche est obligatoire.',
    {7}'La grille des tâches n''est pas trouvée dans la fiche tâche.',
    {8}'heure(s)',
    {9}'Saisir le code de la ressource !',
    {10}'jour(s)',
    {11}'Confirmez vous la suppression de la tâche ?',
    {12}'Confirmez vous la suppression de la ressource et du planning associé#13#10(si de l''activité a été générée pour ce planning, elle ne sera pas supprimée) ?',
    {13}'La date de fin de tâche ne peut pas être supérieure au %s, date de fin de l''affaire.',
    {14}'Aucune ressource n''est sélectionnée.',
    {15}'La fonction est obligatoire.',
    {16}'La date de début est supérieure à la date de fin.',
    {17}'La ressource n''existe pas.',
    {18}'Le type article est obligatoire.',
    {19}'L''article est obligatoire.',
    {20}'Un planning existe pour cette tâche.#13#10Confirmez-vous la suppression de cette tâche et du planning correspondant ?',
    {21}'Voulez-vous valider les modifications ?',
    {22}'La durée d''une intervention doit être entière et positive.',
    {23}'Attention, la valeur initiale de la tâche n''est pas cohérente avec les valeurs initiales des ressources !#13#10Voulez vous quand même calculer les quantités ?',
    {24}'Les quantités à planifier des ressources ne sont pas cohérentes avec la quantité à planifier de la tâche %s.#13#10Le planning ne peut donc pas être généré.',
    {25}'Aucune ressource n''est sélectionnée.',
    {26}'Le paramsoc du calendrier par défaut des ressources doit être renseigné pour pouvoir générer du planning.',
    {27}'Attention, cette tâche a un lien avec une ligne de la pièce affaire.#13#10Confirmez-vous la suppression de cette tâche ?',
    {28}'On ne peut pas avoir deux fois la même famille d''article dans une affaire.',
    {29}'On ne peut pas associer le même article à deux tâches dans une affaire.',
    {30}'Veuillez affecter des ressources à la tâche pour pouvoir générer le planning.',
    {31}'Attention, ce client est fermé. Vous ne devriez pas créer de tâches !',
    {32}'Attention, ce client est fermé. Vous ne devriez pas générer de planning !',
    {33}'La saisie du champ %s est obligatoire.',
    {34}'Il y a un planning existant sur cette tâche. On ne peut pas modifier l''unité.',
    {35}'L''article n''existe pas.',
    {36}'L''adresse d''intervention de cette tâche est alimentée par l''adresse du tiers',
    {37}'Voulez-vous marquer comme terminé toutes les interventions de cette tâche ?',
    {38}'Voulez-vous marquer rapport d''intervention reçu pour toutes les interventions de cette tâche ?',
    {39}'Confirmer-vous le remplacement de la ressource %s par la ressource %s .',
    {40}'(l''intégralité du planning sera transféré sur la nouvelle ressource même si de l''activité a été générée pour ce planning) ?',
    {41}'Action interdite lorsque la fiche est ouverte depuis le planning.',
    {42}'Ressource',
    {43}'Libellé',
    {44}'Qté à PLA',
    {45}'Qté RAP',
    {46}'Qté prévu PC',
    {47}'Qté RAF',
    {48}'statut',
    {49}'Pour une des tâches à générer, les quantités à planifier des ressources ne sont pas cohérentes avec la quantité à planifier de la tâche.#13#10 Le planning ne peut donc pas être généré.',
    {50}'L''option génération limitée à la quantité à planifier est cochée.#13#10 La quantité à planifier ne peut donc être nulle.',
    {51}'Un planning existe pour cette tâche, on ne peut pas changer l''article.',
    {52}'Le nombre de jours prévus pour les ressources est différent du nombre de jours de la tâche.#13#10 Voulez-vous alimenter le total de la tâche par celui des ressources ?',
    {53}'Attention, le planning n''est pas géré en heure actuellement, les heures seront arrondies à la journée supérieure, voulez-vous quand même générer les interventions ?',
    {54}'La durée d''une intervention doit être positive.',
    {55}'Le code appel n''a pu être généré. Veuillez vérifier votre paramétrage.',
    {56}'La date de début de tâche ne peut pas être inférieur au %s, date de début de l''affaire.',
    {57}'Affaire : %s.' + chr(10) + ' Tiers : %s. ' + Chr(10) + 'Aucune intervention n''a été générée pour la tâche %s. ',
    {58}'Si vous avez généré en utilisant ''Ajouter au planning existant'', vous planifiez à partir de la dernière date de génération précisée dans l''onglet ''Règles de la fiche Tâche''. ',
    {59}'Attention à la date de fin de génération paramètrée qui est actuellement au %s.',
    {60}'La génération de la tache %s c''est déroulée normalement. ' + Chr(10) + 'Nombre d''appel généré : %s ',
    {61}'La tâche à dupliquer a déjà été générée.' + chr(10) + 'Voulez-vous continuer en réinitialisant la tâche dupliquée ?',
    {62}'La date de fin de la tâche à générer est égale à 2099.' + chr(10) + 'Seule les 5 premières années seront générées !',
    {63}'La date de Début de la tâche à générer est égale à 1900.' + chr(10) + 'Cette date sera remplacé par la date de début de facturation !'

    );

function AFLanceFicheBTTaches(Argument: string): string;
begin
  result := AGLLanceFiche('BTP', 'BTTACHES', '', '', Argument);
end;

//************************** Evènements de la TOF ******************************

procedure TOF_BTTACHES.OnNew;
begin
  inherited;
end;

procedure TOF_BTTACHES.OnDelete;
begin
  inherited;
end;

procedure TOF_BTTACHES.OnUpdate;
var
  NomChamp: string;
begin
  inherited;

  if assigned(ecran.activecontrol) and (ecran.activecontrol.name = 'GSRES') then
    setFocusControl('ATA_FONCTION');

  // C.B 18/11/2004 on n'empeche pas d'enregistrer si qtés pas cohérentes
  // si on est en mode de génération planning jour, et que qté intiale, ilf aut que les qté par ressource
  // soit valide, sinon on interdit la validation
  //if (FCB_RAPPLANNING.checked) and not (TestCoherence (True)) then exit;
  if Valider(false, false) then
    RefreshGrid
  else
    exit;

  NomChamp := VerifierChampsObligatoires(Ecran, '');
  if NomChamp <> '' then
  begin
    NomChamp := ReadTokenSt(NomChamp);
    SetFocusControl(NomChamp);
    LastError := 33;
    LastErrorMsg := traduitGA(format(TexteMessage[LastError], [champToLibelle(NomChamp)]));
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : validation des modifications
               : pBoForce : forcer la maj
               : pBoSupControl : forcer la maj lors de suppressions
Mots clefs ... :
*****************************************************************}
function TOF_BTTACHES.Valider(pBoForce, pBoSupControl: Boolean): Boolean;
var
  io: TIoErr;

begin

  result := true;

  SetControlText('ATA_DATEDEBPERIOD', GetControlText('ATA_DATEDEBPERIOD'));
  SetControlText('ATA_DATEFINPERIOD', GetControlText('ATA_DATEDEBPERIOD_'));

  // validation des données saisies dans l'écran si modification
  if pBoSupControl or ControleEcran then
  begin
    fTobTacheJour := Tob.create('AFTACHEJOUR', nil, -1);
    try
      if pBoForce or ScreenToTob then
      begin
        io := Transactions(ValiderLesTaches, 1);
        if io <> oeOk then
        begin
          PGIBoxAF(TexteMessage[2], ecran.Caption); //La mise à jour des tâches ne s''est pas effectuée correctement.
          result := false;
        end
        else
        begin
          fStatut := taConsult;
          // les prix init ont été recalculer...  on réaffiche
          if not pBoSupControl then
          begin
            ValiderLesTachesJour;
            //SetControlText('ATA_INITPTPR', FtobDet.getvalue('ATA_INITPTPR'));
            //SetControlText('ATA_INITPTVENTEHT', FtobDet.getvalue('ATA_INITPTVENTEHT'));
          end;
        end;
      end;
    finally
      FreeAndNil(fTobTacheJour);
    end;
  end
  else
    result := false;

  fBoClose := result;

  //FV1 : 14/03/20018 - FS#3006 - ESPACS - En création de tâches contrat les lignes sont doublées
  SetControlVisible('BINSERT',        True);
  SetControlVisible('BDUPLIQUE',      True);
  SetControlVisible('BINSERTMODELE',  True);
  SetControlVisible('BDELETE',        True);
  SetControlEnabled('BDELETE',        True);
  SetControlVisible('BGENERER',       True);
end;

procedure TOF_BTTACHES.OnLoad;
var
  s: string;
begin
  inherited;

  AppliquerConfidentialite(Ecran, '');

  CacheZoneSurTypeArt;

  if getControlText('ATA_CODEARTICLE') <> '' then
     begin
     S := RechDom('AFTFAMARTICLE', GetControltext('ATA_CODEARTICLE'), false);
     if (s <> '') and (s <> 'Error') then
        SetControlText('FAMILLEARTICLE', RechDom('GCFAMILLENIV1', s, false));
     end;

  SetcontrolVisible('ATA_PAVANCEMENT', False);
  SetcontrolVisible('TATA_PAVANCEMENT', False);

  CalculQuantitePla(False);

  // C.B 07/11/2006
  // pour que le getecran ne change pas la valeur de modifie
  fTobDet.GetEcran(Ecran, fPaDet);
  fTobDet.modifie := false;

  // pour que le chargement des règles ne les mettent pas modifiées
  fBoModifRegles := false;

end;

procedure TOF_BTTACHES.CacheZoneSurTypeArt;
begin

  //C.B 21/03/2006
  //on cache RAP tant que le planning jour n'existe pas
  //if (getControlText('ATA_UNITETEMPS') = 'J') then
  //  setControlVisible('RAP', not fBoAFPLANDECHARGE)
  //else
    setControlVisible('RAP', false);

  {if IsTachePrestation(GetControlText('ATA_TYPEARTICLE')) then   //AB-200605 - FQ 12918
     begin
     if GetControlEnabled('ATA_QTEINITPLA') then SetControlEnabled('ATA_QTEINITPLA', not fBoAFPLANDECHARGE);
     SetControlVisible('GB_PRIXRAF', true);
     SetCOntrolVisible('ATA_QTEINITIALE', true);
     if GetControlEnabled('ATA_QTEINITIALE') then SetControlEnabled('ATA_QTEINITIALE', True);
     SetCOntrolVisible('TATA_QTEINITIALE', true);
     SetCOntrolVisible('QTEAPLANIFIERCALC', true);
     SetCOntrolVisible('T_QTEAPLANIFIERCALC', true);
     SetCOntrolEnabled('ATA_INITPTPR', False);
     SetCOntrolEnabled('ATA_INITPTVENTEHT', False);
     SetCOntrolEnabled('RAPPTPCALC', False);
     end
  else
     begin}
     SetControlVisible('GB_PRIXRAF', False);
     SetCOntrolVisible('ATA_QTEINITIALE', False);
     SetCOntrolVisible('TATA_QTEINITIALE', FAlse);
     SetCOntrolVisible('QTEAPLANIFIERCALC', False);
     SetCOntrolVisible('T_QTEAPLANIFIERCALC', False);
     SetCOntrolEnabled('ATA_INITPTPR', False);
     SetCOntrolEnabled('ATA_INITPTVENTEHT', False);
     SetCOntrolEnabled('RAPPTPCALC', False);
     //end;

  // C.B  25/07/2005
  // la gestion du RAP dans le  source de génération ne marche qu'en jour
  if (GetControlText('ATA_UNITETEMPS') = 'J') then
    SetControlenabled('ATA_RAPPLANNING', True)
  else
  begin
    SetControlChecked('ATA_RAPPLANNING', False);
    SetControlenabled('ATA_RAPPLANNING', False);
  end;

end;

procedure TOF_BTTACHES.OnArgument(StArgument: string);
var ed        : thedit;
    vArticle  : TControl;
begin
  inherited;

  fBoClose    := true;
  fUpdateRes  := nil;
  fStatut     := taConsult;

  fCurAffaire.StAffaire := '';
  fCurAffaire.StTiers   := '';

  fBoNoSuppression:= False;
  FCB_RAPPLANNING := TcheckBox(GetControl('ATA_RAPPLANNING'));

  // pas de lien taches / lignes  - maj dans loadaffaire
  SetControlVisible('BTLIGNE', GetParamsocSecur('SO_BTMODEPLANNING','') <> 'TAC');
  TraitementArgument(stArgument);

  // chargement d'une affaire
  if fCurAffaire.StAffaire <> '' then
     begin
     CodeAffaireDecoupe(fCurAffaire.StAffaire, fCurAffaire.StAff0, fCurAffaire.StAff1, fCurAffaire.StAff2,
                        fCurAffaire.StAff3, fCurAffaire.StAvenant, taModif, false);
     GestionEcran;
     LoadAffaire(0);
     end;

  vArticle := getcontrol('ATA_CODEARTICLE');
  if vArticle is THEdit then
     begin
     THEdit(vArticle).OnExit := ArticleOnexit;
     THEdit(vArticle).OnElipsisClick := ArticleOnElipsisClick;
     end;

  if (getcontrol('GBVALEUR') <> nil) then
     SetControlVisible('GBVALEUR', False);
  if (getcontrol('GBTEXTES') <> nil) then
     SetControlVisible('GBTEXTES', False);
  if (getcontrol('GBDECISION') <> nil) then
     SetControlVisible('GBDECISION', False);
  if (getcontrol('GBDATES__') <> nil) then
     SetControlVisible('GBDATES__', False);

  // Cacher le lien sur le planning pour algoe
  //fBoAFPLANDECHARGE := GetParamsocSecur('SO_BTPLANDECHARGE', false);
  fBoAFPLANDECHARGE := false;
  //
  //setControlVisible('BPLANNING', not fBoAFPLANDECHARGE); //planning jour, vu si pas PC uniquement
  //setControlVisible('BQTERESSOURCE', not fBoAFPLANDECHARGE); //planning jour, vu si pas PC uniquement
  //setControlVisible('BGENERER', not fBoAFPLANDECHARGE); //planning jour, vu si pas PC uniquement
  //setControlVisible('RAP', not fBoAFPLANDECHARGE); //planning jour, vu si pas PC uniquement
  // si gestion du RAF et gestion en plan de charge
  //setControlVisible('BPLANDECHARGE', fBoAFPLANDECHARGE);
  //setControlVisible('ATA_NUMEROTACHE', False);
  //setControlEnabled('ATA_NBRESSOURCE', False);
  //if fBoAFPLANDECHARGE then
     //begin
     //setControlEnabled('ATA_UNITETEMPS', False);
     //setControlEnabled('ATA_UNITETEMPS_', False);
     //end;
  //if not (GetParamsocSecur('SO_BTGESTIONRAF', false)) then
     //begin
     //SetControlVisible('GB_PRIXRAF', False);
     //SetControlVisible('GB_PC', False);
     //end;

  setControlVisible('BPLANNING', fBoAFPLANDECHARGE); //planning jour, vu si pas PC uniquement
  setControlVisible('BQTERESSOURCE', fBoAFPLANDECHARGE); //planning jour, vu si pas PC uniquement
  setControlVisible('BGENERER', not fBoAFPLANDECHARGE); //planning jour, vu si pas PC uniquement
  //setControlVisible('RAP', fBoAFPLANDECHARGE); //planning jour, vu si pas PC uniquement
  setControlVisible('RAP', false); //planning jour, vu si pas PC uniquement
  // si gestion du RAF et gestion en plan de charge                                      
  setControlVisible('BPLANDECHARGE', fBoAFPLANDECHARGE);
  setControlVisible('ATA_NUMEROTACHE', False);
  setControlEnabled('ATA_NBRESSOURCE', False);
  setControlEnabled('ATA_UNITETEMPS', False);
  setControlEnabled('ATA_UNITETEMPS_', False);

  setcontrolEnabled('RB_MODEGENE1', GetParamsocSecur('SO_AFGENELIMIT', false) and (valeur(GetControltext('ATA_QTEINITPLA')) <> 0));
  setcontrolEnabled('RB_MODEGENE2', GetParamsocSecur('SO_AFGENELIMIT', false) and (valeur(GetControltext('ATA_QTEINITPLA')) <> 0));

  // on cache tout ce qui concerne le RAF sur PC
  SetControlVisible('GB_PRIXRAF', False);
  SetControlVisible('GB_PC', False);
  SetControlVisible('TAB_COMPLEMENT', False);

  if VH_GC.GCIfDefCEGID then
  begin
    // C.B 20/06/2005
    // si le RI est reçu, on stocke la date et qui a saisi le RI
    SetControlEnabled('ATA_DATELIBRE1', False);
    SetControlEnabled('ATA_CHARLIBRE2', False);
  end;

  TCheckBox(GetControl('ATA_TERMINE')).OnClick := ATA_TERMINEonClick;
  TBitBtn(GetControl('BTN_SELECTION')).OnClick := BtnSelectionOnClick;
  TBitBtn(GetControl('BTN_EFFACER')).OnClick := BtnEffacerOnClick;

  // C.B 29/11/2005
  // la zone était saisiaable alors qu'elle est toujours calculée
  SetControlEnabled('QTERAP', False);

  //gestion de la modification de la Regle de calcul
  Ed := THEdit(GetControl('ATA_JOURINTERVAL'));
  Ed.OnExit := ModifZoneExit;

  //Semaine
  Ed := THEdit(GetControl('ATA_SEMAINEINTERV'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_JOUR1H'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_JOUR2H'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_JOUR3H'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_JOUR4H'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_JOUR5H'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_JOUR6H'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_JOUR7H'));
  Ed.OnExit := ModifZoneExit;

  //Mois
  Ed := THEdit(GetControl('ATA_MOISJOURFIXE'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_MOISFIXE0'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_MOISFIXE1'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_MOISFIXE2'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_MOISSEMAINE1'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_MOISSEMAINE2'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_MOISJOURLIB'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_JOUR1M'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_JOUR2M'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_JOUR3M'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_JOUR4M'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_JOUR5M'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_JOUR6M'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_JOUR7M'));
  Ed.OnExit := ModifZoneExit;

  //Année
  Ed := THEdit(GetControl('JOURAN'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('MOISAN'));
  Ed.OnExit := ModifZoneExit;
  Ed := THEdit(GetControl('ATA_ANNEENB'));
  Ed.OnExit := ModifZoneExit;

  //Nombre d'intervention
  Ed := THEdit(GetControl('ATA_NBINTERVENTION'));
  Ed.OnExit := ModifZoneExit;

  BloquerTaches;

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : 2 blocages possibles
                 1: En lien avec les commandes on ne peut pas manipuler les tâches
                    lancé qu'il seule fois, car identique pour toutes les taches
                 2: si lancer depuis le planning, on ne peut pas supprimer et créer
                    taches et ressources
Mots clefs ... :
*****************************************************************}
procedure TOF_BTTACHES.BloquerTaches;
var vBoEnabled: Boolean;

begin

  // gestion du planning récurrent
  // C.B 23/05/2006
  // l'onglet regles est visible en mode automatique
  //setcontrolvisible('TAB_REGLES', (not fBoAFPLANDECHARGE) and GetControlEnabled('TAB_TACHE'));

  // si mode automatique ou manuel
  vBoEnabled := (GetParamsocSecur('SO_BTMODEPLANNING', '') <> 'ART');

  SetControlEnabled('ATA_QTEINITPLA', vBoEnabled);
  SetControlEnabled('ATA_LIBELLETACHE1', vBoEnabled);
  SetControlEnabled('ATA_LIBELLETACHE2', vBoEnabled);
  SetControlEnabled('ATA_CODEARTICLE', vBoEnabled);

  //SetControlEnabled('ATA_TYPEARTICLE', vBoEnabled);
  SetControlText('ATA_TYPEARTICLE', 'PRE');
  SetControlEnabled('ATA_TYPEARTICLE', false);

  //SetControlEnabled('ATA_ACTIVITEREPRIS', vBoEnabled);
  SetControlEnabled('ATA_UNITETEMPS', vBoEnabled);
  SetControlEnabled('ATA_FAMILLETACHE', vBoEnabled);
  SetControlEnabled('ATA_DATEDEBPERIOD', vBoEnabled);
  SetControlEnabled('ATA_DATEDEBPERIOD_', vBoEnabled);
  SetControlEnabled('QTERAP', vBoEnabled);
  SetControlEnabled('ATA_FONCTION', vBoEnabled);
  SetControlEnabled('GSRES', vBoEnabled);

  //SetControlEnabled('ATA_STATUTPLA', vBoEnabled);
  SetControlVisible('ATA_ACTIVITEREPRIS', False);
  SetControlVisible('ATA_STATUTPLA', False);

  // C.B 07/11/2006
  // si mode automatique
  if not vBoEnabled then
     begin
     SetControlEnabled('BINSERT', vBoEnabled);
     SetControlEnabled('BINSERTMODELE', vBoEnabled);
     SetControlEnabled('BDUPLIQUE', vBoEnabled);
     SetControlEnabled('BDELETE', vBoEnabled);
     end
  else
     begin
     // si mode manuel, on vérifie si lancé depuis le planning
     // dans ce cas on ne peux pas créer et supprimer
     vBoEnabled := (not fBoNoSuppression) and (fAction <> taConsult);
     SetControlEnabled('BINSERT', vBoEnabled);
     SetControlEnabled('BDUPLIQUE', vBoEnabled);
     SetControlEnabled('BINSERTMODELE', vBoEnabled);
     SetControlEnabled('BDELETE', vBoEnabled);
     end;

  // on bloque également la création et la suppression de ressource
  // si lancer depuis le planning
  vBoEnabled := not fBoNoSuppression;
  SetControlEnabled('BINSERTRES', vBoEnabled);
  SetControlEnabled('BDELETERES', vBoEnabled);

end;

procedure TOF_BTTACHES.TypeArticleOnExit(Sender: TObject);
begin
  //CacheZoneSurTypeArt;
  //if not IsTachePrestation (GetControlText('ATA_TYPEARTICLE')) then //AB-200605 - FQ 12918
       //begin
   	   // si on a changer le code article, on initialise les qt& à 0 car non gérée
       //SetControltext('ATA_QTEINITIALE', '0');
       //SetControltext('ATA_QTEINITPLA', '0');
       //end;
end;

procedure TOF_BTTACHES.ArticleOnExit(Sender: TObject);
begin
  RaffraichirArticle;
end;

procedure TOF_BTTACHES.TypeActionOnExit(Sender: TObject);
var StReq : String;
    QQ    : TQuery;
Begin

  LectureTypeAction(GetControlText('ATA_BTETAT'));

  StReq := 'SELECT BTA_DUREEMINI FROM BTETAT WHERE BTA_BTETAT = "' + GetControlText('ATA_BTETAT') + '"';
  QQ := nil;

  try
    QQ := OpenSQL(StReq, True);
    if (not QQ.EOF) then
       begin
       setControlText('ATA_QTEINTERVENT', FloatToStr(QQ.findField('BTA_DUREEMINI').asfloat));
       end;
  finally
    ferme(QQ);
  end;

end;

procedure TOF_BTTACHES.FonctionOnExit(Sender: TObject);
var
  vst: string;
  vqr: tquery;
  Pr, Pv: double;
begin


  if (GetParamsocSecur('SO_AFVALOPR', '') <> 'FON') and (GetParamsocSecur('SO_AFVALOPV', '') <> 'FON') then
    exit;

  vSt := 'SELECT AFO_TAUXREVIENTUN, AFO_PVHT FROM FONCTION WHERE AFO_FONCTION = "' + GetControlText('ATA_FONCTION') + '"';
  vQr := nil;
  PR := 0;
  PV := 0;
  try
    vQR := OpenSQL(vSt, True);
    if (not vQR.EOF) then
    begin
      PR := vQr.findField('AFO_TAUXREVIENTUN').asfloat;
      PV := vQr.findField('AFO_PVHT').asfloat;
    end;
  finally
    ferme(vqr);
  end;

  if GetParamsocSecur('SO_AFVALOPR', '') = 'FON' then
  begin
    SetControlText('ATA_PUPR', FloattoStr(PR));
    ATA_PUPROnExit(self);
  end;

  if GetParamsocSecur('SO_AFVALOPV', '') = 'FON' then
  begin
    SetControlText('ATA_PUVENTEHT', FloattoStr(PV));
    ATA_PUVentehTOnExit(self);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 18/03/2002
Modifié le ... : 18/03/2002
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOF_BTTACHES.OnClose;
begin
  inherited;

  Valider(false, false);

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 18/03/2002
Modifié le ... :   /  /
Description .. : GESTION DE L'ENTETE
Mots clefs ... :
*****************************************************************}
procedure TOF_BTTACHES.AffichageEntete;
var
  vSt: string;
  vQr: TQuery;
  vStAffFormat: string;
  vStaffiche: string;

begin

  // Par Affaire
  if fCurAffaire.StAffaire = '' then
     begin
     PGIBoxAF(TexteMessage[1], ecran.Caption); //Saisie par affaire sur code affaire non valide.
     close;
     end
  else
     begin
     SetControlText('ATA_AFFAIRE', fCurAffaire.StAffaire);
     // C.B 30/03/2006
     // selection des dates de l'affaire pour l'initailisation des la tache
     vSt := 'SELECT AFF_LIBELLE, T_AUXILIAIRE, T_TIERS, T_LIBELLE, AFF_DEVISE, ';
     vSt := vSt + 'AFF_DATEDEBUT, AFF_DATEFIN, AFF_DATESIGNE, AFF_DATEDEBGENER, ';
     vSt := vSt + 'AFF_DATELIMITE, AFF_DATEFINGENER, AFF_DATERESIL, ';
     vSt := vSt + 'AFF_DATECLOTTECH, AFF_DATEGARANTIE, AFF_AFFAIREINIT, AFF_CHANTIER ';
     vSt := vSt + 'FROM TIERS, AFFAIRE WHERE AFF_AFFAIRE = "' + fCurAffaire.StAffaire + '" ';
     vSt := vSt + 'AND AFF_TIERS = T_TIERS';

     vQr := nil;
     try
        vQR := OpenSql(vSt, True);
        if not vQR.Eof then
           begin
           fCurAffaire.StLibAff := vQR.FindField('AFF_LIBELLE').AsString;
           fCurAffaire.StLibTiers := vQR.FindField('T_LIBELLE').AsString;
           SetControlText('ATA_TIERS', vQR.FindField('T_TIERS').AsString);
           fStTiers := GetControlText('ATA_TIERS');
           fStAuxilliaire := vQR.FindField('T_AUXILIAIRE').AsString;
           SetControlText('ATA_DEVISE', vQR.FindField('AFF_DEVISE').AsString);
           fCurAffaire.StTiers    := vQR.FindField('T_TIERS').AsString;
           fCurAffaire.StDevise   := vQR.FindField('AFF_DEVISE').AsString;
           fCurAffaire.StChantier := vQR.FindField('AFF_CHANTIER').AsString;
           fCurAffaire.StAffInit  := vQR.FindField('AFF_AFFAIREINIT').AsString;
           fCurAffaire.DtDebutAff := vQR.FindField('AFF_DATEDEBUT').AsDateTime;
           fCurAffaire.DtFinAff       := vQR.FindField('AFF_DATEFIN').AsDateTime;
           fCurAffaire.DtDateSigne    := vQR.FindField('AFF_DATESIGNE').AsDateTime;
           fCurAffaire.DtDateDebGener := vQR.FindField('AFF_DATEDEBGENER').AsDateTime;
           fCurAffaire.DtResil        := vQR.FindField('AFF_DATERESIL').AsDateTime;
           fCurAffaire.DtDateFinGener := vQR.FindField('AFF_DATEFINGENER').AsDateTime;
           fCurAffaire.DtDateLimite   := vQR.FindField('AFF_DATELIMITE').AsDateTime;
           fDtDebutAffaire            := vQR.FindField('AFF_DATEDEBUT').AsDateTime;
           fDtFinAffaire              := vQR.FindField('AFF_DATEFIN').AsDateTime;
           fCurAffaire.DtDateCloture  := vQR.FindField('AFF_DATECLOTTECH').AsDateTime;
           fCurAffaire.DtDateGarantie := vQR.FindField('AFF_DATEGARANTIE').AsDateTime;
           end;
        // affichage du tiers et de l'affaire
        vStAffFormat := CodeAffaireAffiche(fCurAffaire.StAffaire, ' ');
        vStaffiche := format('%s   %s / %s   %s', [fCurAffaire.StTiers, fCurAffaire.StLibTiers, VStAffFormat, fCurAffaire.StLibAff]);
        SetcontrolText('TLIBCLIAFF', vStaffiche);
        fStAfficheAffaire := vStAffFormat + '  ' + fCurAffaire.StLibAff;
     finally
        if vQR <> nil then ferme(vQR);
     end;
  end;
end;

//------------------------------------------------------------------------------
//                            Gestion des grilles
//------------------------------------------------------------------------------
{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 18/03/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... : fGSTaches : Liste des tâches
*****************************************************************}
procedure TOF_BTTACHES.InitGrid;
begin

  fGSTaches := THGRID(GetControl('GSTaches'));

  if assigned(fGSTaches) then
     begin
     fGSTaches.OnRowEnter := fGSTachesRowEnter;
     fGSTaches.OnRowExit := fGSTachesRowExit;
     fGSTaches.ListeParam := 'AFLISTETACHES';
     fGSTaches.ColWidths[0] := 18;
     FormatColGridTaches;
     TFVierge(Ecran).Hmtrad.ResizeGridColumns(fGSTaches);
     AffecteGrid(fGSTaches, taConsult);
     end
  else
     PGIBoxAF(TexteMessage[7], ecran.Caption); //La grille des tâches n''est pas trouvée dans la fiche tâche.

end;

procedure TOF_BTTACHES.FormatColGridTaches;
var
  vStNomCol: string;
  vStCols: string;
  i: Integer;
begin
  fStLesCols := fGSTaches.Titres[0];
  vStCols := fStLesCols;
  i := 0;
  repeat
    vStNomCol := AnsiUppercase(Trim(ReadTokenSt(vStCols)));
    if (vStNomCol <> '') then
    begin
      if (vStNomCol = 'ATA_DATEDEBPERIOD') or (vStNomCol = 'ATA_DATEFINPERIOD') then
      begin
        fGSTaches.ColTypes[i] := 'D';
        fGSTaches.ColFormats[i] := ShortDateFormat;
      end
    end;
    i := i + 1;
  until ((vStCols = '') or (vStNomCol = ''));
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 18/03/2002
Modifié le ... :   /  /
Description .. : grille des ressources de la tâches
Mots clefs ... : fGSTaches : Liste des tâches
*****************************************************************}
procedure TOF_BTTACHES.InitGridRes;
begin

  fGSRes := THGRID(GetControl('GSRes'));
  fGSRes.OnCellEnter := fGSResCellEnter;
  fGSRes.OnCellExit := fGSResCellExit;
  fGSRes.OnKeyDown := fGSResKeyDown;
  fGSRes.OnElipsisClick := fGSResElipsisClick;
  fGSRes.OnDblClick := fGSResElipsisClick;
  TFVierge(Ecran).Hmtrad.ResizeGridColumns(fGSRes);

  // Gestion des options de la grille
  AffecteGrid(fGSRes, fAction);

  //  fGSRes.colcount := 11;
  fGSRes.colcount := 8;
  fGSRes.rowcount := 2;
  fGSRes.row := 1;

  fGSRes.CellValues[0, 0] := '';
  fGSRes.CellValues[cInColRes, 0] := TraduitGA(TexteMessage[42]); //Ressource
  fGSRes.CellValues[cInColResLib, 0] := TraduitGA(TexteMessage[43]); //Libellé'

  if (not fBoAFPLANDECHARGE) then
  begin
    fGSRes.CellValues[cInColQte, 0] := TraduitGA(TexteMessage[44]); //Qté PLA
    fGSRes.CellValues[cInColQteRAP, 0] := TraduitGA(TexteMessage[45]); //Qté RAP
  end;

  if GetParamsocSecur('SO_BTGESTIONRAF', false) then // gestion du reste à faire sur PC
  begin
    fGSRes.CellValues[cInColQtePC, 0] := TraduitGA(TexteMessage[46]); //Qté Init PC
    fGSRes.CellValues[cInColQteRAF, 0] := TraduitGA(TexteMessage[47]); //Qté RAF
  end;

  fGSRes.CellValues[cInColStatut, 0] := TraduitGA(TexteMessage[48]); //Statut

  fGSRes.ColFormats[cInColQte] := '#,#00.00';
  fGSRes.ColTypes[cInColQte] := 'R';
  fGSRes.ColAligns[cInColQte] := taRightJustify;
  
  //fGSRes.ColFormats[cInColQtePC] := '#,#00.00';
  //fGSRes.ColTypes[cInColQtePC] := 'R';
  //fGSRes.ColAligns[cInColQtePC] := taRightJustify;

  fGSRes.ColAligns[cInColQteRAP] := taRightJustify;
  //fGSRes.ColAligns[cInColQteRAF] := taRightJustify;
  fGSRes.ColTypes[cInColQteRAP] := 'R';
  //fGSRes.ColTypes[cInColQteRAF] := 'R';
  fGSRes.ColFormats[cInColQteRAP] := '#,#00.00';
  //fGSRes.ColFormats[cInColQteRAF] := '#,#00.00';

  InitColGrid;

  // lignes
  fGSRes.RowHeights[0] := 22;
  fGSRes.ColEditables[cInColResLib] := False;
  fGSRes.ColEditables[cInColStatut] := False;
  fGSRes.ColEditables[cInColQteRAP] := False;
  //fGSRes.ColEditables[cInColQteRAF] := False;
  
end;
                                           
{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... :
Modifié le ... : 25/07/2005
Description .. : RAP  : generation en quantité limité planning jour
                 RAF  : estion des quantités restant à faire pour le plan de charge

                 liste de toutes les conbinaisons possibles :

                 Tout : RAP et RAF
                 Rien : ni RAF, ni RAP
                 RAP sans RAF
                 RAF sans RAP

                 18/10/2005 : On utilise la valeur du checkbox
                 "Generation limite a la qte planifiée" pour afficher les
                 colonnes RAP

Mots clefs ... :
*****************************************************************}
procedure TOF_BTTACHES.InitColGrid;
begin

  fGSRes.ColWidths[0] := 15;

  // Tout : RAF et RAP
  // gestion du RAF et pas que le plan de charge et gestion du RAP
  if GetParamsocSecur('SO_BTGESTIONRAF', false) and (not fBoAFPLANDECHARGE) and
    TCheckBox(GetControl('ATA_RAPPLANNING')).Checked then

  begin
    fGSRes.ColWidths[cInColQte] := 80;
    fGSRes.ColWidths[cInColQteRAP] := 80;
    //fGSRes.ColWidths[cInColQtePC] := 80;
    //fGSRes.ColWidths[cInColQteRAF] := 80;
    fGSRes.ColWidths[cInColRes] := 90;
    fGSRes.ColWidths[cInColResLib] := 140;
    fGSRes.ColWidths[cInColStatut] := 42;
  end

    // Rien : ni RAF, ni RAP
  else
    if (not TCheckBox(GetControl('ATA_RAPPLANNING')).Checked) and
    (not GetParamsocSecur('SO_BTGESTIONRAF', false)) then
  begin
    fGSRes.ColWidths[cInColQte] := -1;
    fGSRes.ColWidths[cInColQteRAP] := -1;
    //fGSRes.ColWidths[cInColQtePC] := -1;
    //fGSRes.ColWidths[cInColQteRAF] := -1;
    fGSRes.ColWidths[cInColRes] := 180;
    fGSRes.ColWidths[cInColResLib] := 292;
    fGSRes.ColWidths[cInColStatut] := 123;
  end

  // RAP sans RAF
  else
    if TCheckBox(GetControl('ATA_RAPPLANNING')).Checked and (not GetParamsocSecur('SO_BTGESTIONRAF', false)) then
  begin
    //fGSRes.ColWidths[cInColQtePC] := -1;
    //fGSRes.ColWidths[cInColQteRAF] := -1;
    fGSRes.ColWidths[cInColQte] := 100;
    fGSRes.ColWidths[cInColQteRAP] := 100;
    fGSRes.ColWidths[cInColRes] := 100;
    fGSRes.ColWidths[cInColResLib] := 192;
    fGSRes.ColWidths[cInColStatut] := 101;
  end

  // RAF sans RAP
  else
    if GetParamsocSecur('SO_BTGESTIONRAF', false) and (not TCheckBox(GetControl('ATA_RAPPLANNING')).Checked) then
  begin
    fGSRes.ColWidths[cInColQte] := -1;
    fGSRes.ColWidths[cInColQteRAP] := -1;
    //fGSRes.ColWidths[cInColQtePC] := 100;
    //fGSRes.ColWidths[cInColQteRAF] := 100;
    fGSRes.ColWidths[cInColRes] := 146;
    fGSRes.ColWidths[cInColResLib] := 146;
    fGSRes.ColWidths[cInColStatut] := 101;
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 18/03/2002
Modifié le ... :   /  /
Description .. : Gestion de la grille des Tâches
Mots clefs ... :
*****************************************************************}
procedure TOF_BTTACHES.fGSTachesRowEnter(SEnder: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var
  S: string;
begin

  TobToScreen(Ou - 1);
  PAGETACHEChange(self);

//  fTobDet.PutEcran(Ecran, fPaDet); modif BRL 20052009 car déjà fait dans TobToScreen FQ12510

  S := RechDom('AFTFAMARTICLE', GetControltext('ATA_CODEARTICLE'), false);
  if (s <> '') and (s <> 'Error') then
    SetControlText('FAMILLEARTICLE', RechDom('GCFAMILLENIV1', s, false))
  else
    SetControlText('FAMILLEARTICLE', '');

  // on ne peut pas changer l'article si du planning existe
  //SetControlEnabled('ATA_CODEARTICLE', not ExistePlanning(fCurAffaire.StAffaire, GetControlText('ATA_NUMEROTACHE')));
  //SetControlEnabled('ATA_TYPEARTICLE', GetControlEnabled('ATA_CODEARTICLE'));
  SetControlText('ATA_TYPEARTICLE', 'PRE');
  SetControlEnabled('ATA_TYPEARTICLE', False);

  if VH_GC.GCIfDefCEGID then
  begin
    if (fTobDet.GetString('ARTICLE_FORFAIT') = 'X') then
    begin
      SetControlVisible('ATA_BOOLLIBRE1', true);
      SetControlVisible('ATA_CHARLIBRE1', true);
      SetControlVisible('TATA_CHARLIBRE1', true);
    end
    else
    begin
      SetControlVisible('ATA_BOOLLIBRE1', false);
      SetControlVisible('ATA_CHARLIBRE1', false);
      SetControlVisible('TATA_CHARLIBRE1', false);
    end;
  end;

  // C.B 29/11/2005
  // les zones n'étaient pas réaffichées
  cacheZoneSurTypeArt;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 18/03/2002
Modifié le ... :   /  /
Description .. : controle des données obligatoires saisies
Mots clefs ... :
*****************************************************************}
procedure TOF_BTTACHES.fGSTachesRowExit(SEnder: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  fGSTaches.SynEnabled := false;
  cancel := not Valider(false, false);
  if not cancel then
    RefreshGrid;
  fGSTaches.SynEnabled := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 18/03/2002
Modifié le ... :   /  /
Description .. : Gestion de la grille Ressources / Tâches
Mots clefs ... :
*****************************************************************}
procedure TOF_BTTACHES.fGSResCellEnter(SEnder: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  // on affiche ... pour les colonnes ressources et fonctions
  if fAction <> taConsult then
  begin
    fGSRes.ElipsisButton := (fGSRes.Col = cInColRes);
    fTobDet.SetAllModifie(true);
  end;
end;
 
{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 03/06/2002
Modifié le ... :   /  /
Description .. : Reaffectation du planning d'une ressource
Mots clefs ... :
*****************************************************************}
procedure TOF_BTTACHES.fGSResCellExit(SEnder: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin

  if fStatut <> taModif then
    Exit;

  //forcer l'ecriture en majuscule
  fGSRes.Cells[ACol, ARow] := AnsiUppercase(fGSRes.Cells[ACol, ARow]);

  // controle de l'existence de la ressource
  if (ACol = cInColRes) and not ExisteRessource(fGSRes.Cells[ACol, ARow]) then
  begin
    PGIBoxAF(TexteMessage[17], ecran.Caption);  //La ressource n''existe pas.
    cancel := true;
  end
  else
    if (ACol = cInColRes) then
    fGSRes.Cells[cInColResLib, ARow] := RessourceLib(fGSRes.Cells[ACol, ARow]);

  //icone
  if fGSRes.Cells[cInColStatut, ARow] = '' then
    fGSRes.Cells[cInColStatut, ARow] := '#ICO#' + RechDom('AFSTATUTRESSOURCE', 'ACT', false);

  // initialisation de la quantite
  if fGSRes.Cells[cInColQte, ARow] = '' then
    fGSRes.Cells[cInColQte, ARow] := '0';
  //if fGSRes.Cells[cInColQtePC, ARow] = '' then
  //  fGSRes.Cells[cInColQtePC, ARow] := '0';

  // modification d'une ressource
  if (not cancel) and (aRow <> 0) and (ACol = cInColRes) and
    (fTobDet.Detail.Count >= aRow) and
    (fGSRes.Cells[ACol, ARow] <> fTobDet.Detail[aRow - 1].GetValue('ATR_RESSOURCE')) then
  begin
    if fGSRes.Cells[ACol, ARow] = '' then
    begin
      cancel := true;
      fGSRes.Cells[ACol, ARow] := fTobDet.Detail[aRow - 1].GetValue('ATR_RESSOURCE');
      fGSRes.Cells[cInColResLib, ARow] := RessourceLib(fGSRes.Cells[ACol, ARow]);
    end

    else if PGIAsk(format(traduitGA(TexteMessage[39]) + #13#10 + traduitGA(TexteMessage[40]), [fTobDet.Detail[aRow - 1].GetValue('ATR_RESSOURCE'), fGSRes.Cells[ACol, ARow]]), traduitGA(ecran.Caption)) = MrYes then //Confirmer-vous le remplacement de la ressource %s par la ressource %s .
      AddUpdateRessList(aCol, aRow)

    else
    begin
      cancel := true;
      fGSRes.Cells[ACol, ARow] := fTobDet.Detail[aRow - 1].GetValue('ATR_RESSOURCE');
      fGSRes.Cells[cInColResLib, ARow] := RessourceLib(fGSRes.Cells[ACol, ARow]);
    end;
  end;
  fGSRes.ElipsisButton := false;
end;

procedure TOF_BTTACHES.fGSResKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key <> VK_TAB) then
    fStatut := taModif;
  case key of
    VK_F5: fGSResElipsisClick(SEnder);
    VK_DOWN: if (fGSRes.rowcount < 2) then bInsertResOnClick(self);
    VK_RETURN: key := VK_TAB;
    VK_DELETE: if (Shift = [ssCtrl]) then bDeleteResOnClick(self);
    VK_UP:
      if (fGSRes.row = fGSRes.rowcount - 1) and (fGSRes.Cells[cInColRes, fGSRes.row] = '') then
      begin
        bDeleteResOnClick(self);
        fGSRes.row := fGSRes.rowcount - 2;
        key := vk_down;
      end;
  end;
end;

procedure TOF_BTTACHES.fGSResElipsisClick(SEnder: TObject);
var Title   : String;
    StChamp : string;
    StFonction  : String;
    StWhere : string;
    CodeRessource : THCritMaskEdit;
begin

  if fGSRes.Col <> cInColRes then exit;

  CodeRessource := ThCritMaskEdit.Create(TFvierge(ecran));

  title       := 'Recherche Ressource tâches';

  StChamp     := fGSRes.Cells[fGSRes.Col, fGSRes.Row];
  StFonction  := THValComboBox(GetControl('ATA_FONCTION')).Value;

  if (VH_GC.AFRechResAv) then
  begin
    stWhere     := 'TYPERESSOURCE:SAL,INT';
    if stFonction <> '' then
    begin
      if StWhere <> '' then stWhere := stWhere + ';ARS_FONCTION1:' + stFonction + '';
    end;
  end
  else
  begin
    stWhere     := 'ARS_TYPERESSOURCE="SAL" OR ARS_TYPERESSOURCE = "INT"';
    if stFonction <> '' then
    begin
      if StWhere <> '' then stWhere := stWhere + 'AND ARS_FONCTION1="' + stFonction + '"';
    end;
  end;

  DispatchRecherche(CodeRessource, 3, StWhere,'ARS_RESSOURCE=' + Trim(StChamp), '');

  if StChamp <> '' then
  begin
    fGSRes.Cells[fGSRes.Col, fGSRes.Row] := CodeRessource.text;
    fStatut := taModif;
  end;

  {*
  if fGSRes.Col = cInColRes then
  begin
    // fonction remplie
    if THValComboBox(GetControl('ATA_FONCTION')).Text <> '' then
    begin
      vStRes := DispatchRecherche(nil, 3,
        'ARS_RESSOURCE=' + fGSRes.Cells[fGSRes.Col, fGSRes.Row] +
        ';ARS_FONCTION1=' + THValComboBox(GetControl('ATA_FONCTION')).value, '', '');
      if vStRes <> '' then
      begin
        fGSRes.Cells[fGSRes.Col, fGSRes.Row] := vStRes;
        fStatut := taModif;
      end;
    end
    else
    begin
      vStRes := DispatchRecherche(nil, 3,
        'ARS_RESSOURCE=' + fGSRes.Cells[fGSRes.Col, fGSRes.Row], '', '');
      if vStRes <> '' then
      begin
        fGSRes.Cells[fGSRes.Col, fGSRes.Row] := vStRes;
        fStatut := taModif;
      end;
    end;
  end;
  *}

  CodeRessource.Free;

end;

procedure TOF_BTTACHES.FormKeyDown(SEnder: TObject; var Key: Word; ShIft: TShIftState);
var
  vBoClose: Boolean;
begin
  case Key of
    VK_ESCAPE:
      begin
        vBoClose := true;
        EcranOnCloseQuery(self, vBoClose)
      end;
    VK_F10:
      if Valider(false, false) then
        RefreshGrid;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 03/06/2002
Modifié le ... :   /  /
Description .. : Gestion des boutons
Mots clefs ... :
*****************************************************************}
procedure TOF_BTTACHES.bInsertOnClick(SEnder: TObject);
begin

  // C.B 07/11/2006
  if fBoNoSuppression then
    PGIBoxAF(TexteMessage[41], ecran.Caption) //Action interdite
  else
    begin
    if Valider(false, false) then
       begin
       RefreshGrid;
       InsertTache;
       SetActiveTabSheet('TAB_TACHE');
       SetFocusControl('ATA_LIBELLETACHE1');
       end;
    SetControlProperty('ATA_ANNEENB', 'Value', '1');
  end;

end;

function TOF_BTTACHES.ClientFerme: Boolean;
var
  vSt: string;
  vQr: TQuery;
begin
  Result := False;
  vSt := 'SELECT T_FERME FROM TIERS, AFFAIRE ';
  vSt := vSt + ' WHERE AFF_TIERS = T_TIERS ';
  vSt := vSt + ' AND T_TIERS = "' + fCurAffaire.StTiers + '"';
  vQr := nil;
  try
    vQR := OpenSql(vSt, True);
    if not vQR.Eof then
      Result := vQr.FindField('T_FERME').AsString = 'X';
  finally
    Ferme(vQR);
  end;
end;

{*********************************************************************
Auteur  ...... : FV
Créé le ...... : 01/03/2007
Description .. : Duplication d'une tache sélectionnée
**********************************************************************}
procedure TOF_BTTACHES.bDupliqueOnClick(SEnder: TObject);
var ARow      : integer;
    VstMsg    : String;
    RowSaved  : integer;
    OkInit    : Boolean;
    NoTacheDup: String;
Begin

    RowSaved := fGSTaches.Row-1;
    NoTacheDup := fTobTaches.detail[FtobTaches.detail.count-1].Getvalue('ATA_NUMEROTACHE');

    if not Valider(false, false) then exit;

    RefreshGrid;

    if ClientFerme then
       PGIBoxAF(TexteMessage[31], ecran.Caption); //Attention, ce client est fermé. Vous ne devriez pas créer de tâches !

    //controle si la tache à dupliquer n'est pas déjà générée
    if fTobTaches.detail[Rowsaved].GetValue('ATA_TERMINE') = 'X' then
       Begin
       VstMsg := traduitGA(TexteMessage[61]);
       if (PGIAsk(VstMsg) <> Mryes) then
          exit;
       end;

    fTobDet := TOB.Create('TACHE', fTobTaches, -1);
    fTobDet.AddChampSupValeur('QTEAPLANIFIERCALC', 0, false);
    fTobDet.AddChampSupValeur('RAPPTPRCAL', 0, false);
    fTobDet.AddChampSupValeur('RAPPTPVCAL', 0, false);
    fStatut := taCreat;

    if assigned(fGSTaches) then
       begin
       Arow := fTobTaches.detail.count;
       fGSTaches.InsertRow(ARow);
       fGSTaches.Row := ARow;
       fGSTaches.RowCount := fTobTaches.detail.count + 1;
       end;

    RefreshGridRes;

    InitNewTache(fTobDet);

    fTobDet.Dupliquer(fTobTaches.detail[RowSaved], false, true);

    SetControlText('ATA_NUMEROTACHE', IntToStr(StrToInt(NoTacheDup)+1));
    SetControlChecked('ATA_TERMINE', False);

    if not Valider(false, false) then exit;

    LoadAdresseTache(True, fTobDet);

    RefreshAdresse;

    RefreshGrid;

    fGSTaches.Row := TobToScreen(ARow-1);

    SetActiveTabSheet('TAB_TACHE');

    SetFocusControl('ATA_LIBELLETACHE1');

end;

{*********************************************************************
Auteur  ...... : AB
Créé le ...... : 25/03/2002
Description .. : Création de tâches à partir de modèles de tâches
**********************************************************************}
procedure TOF_BTTACHES.bInsertModeleOnClick(SEnder: TObject);
var
  vTOBModeles : TOB;
  i          : integer;
  NoTacheDup : string;
  ARow       : integer;
  VstMsg     : String;
  OkInit     : Boolean;
begin

  //if (fStatut <> taCreat) then exit;

  // C.B 07/11/2006
  if fBoNoSuppression then
     Begin
     PGIBoxAF(TexteMessage[41], ecran.Caption); //Action interdite
     exit;
     end;
   if FtobTaches.detail.count > 0 then
   begin
   	NoTacheDup := fTobTaches.detail[FtobTaches.detail.count-1].Getvalue('ATA_NUMEROTACHE');
    if StrToint(NoTacheDup) = 0 then
    begin
      fTobTaches.detail[FtobTaches.detail.count-1].free;
    end;
   	NoTacheDup := IntToStr(StrToInt(NoTacheDup) + 1);

   end else
   begin
   	notacheDup := '1';
   end;

   vTOBModeles := MaTobInterne('Liste ModeleTaches');

   AGLLanceFiche('BTP', 'BTTACHEMODELE_MUL', '', '', 'MULTI');

   //RowSaved := fGSTaches.Row-1;

//   fTobDet.free;
   if vTOBModeles.detail.count = 0 then begin fTOBDet := nil; exit; END;
   for i := 0 to vTOBModeles.detail.count-1 do
   Begin
     fTobDet := TOB.Create('TACHE', fTobTaches, -1);
     fTobDet.AddChampSupValeur('QTEAPLANIFIERCALC', 0, false);
     fTobDet.AddChampSupValeur('RAPPTPRCAL', 0, false);
     fTobDet.AddChampSupValeur('RAPPTPVCAL', 0, false);
     fStatut := taCreat;
     InitNewTache(fTobDet);
     InsererModele(vTOBModeles.detail[i], FtobDet);
     FtobDet.PutValue('ATA_NUMEROTACHE', IntToStr(StrToInt(NoTacheDup)+i));
     FtobDet.PutValue('ATA_FAITPARQUI', '');
     LoadAdresseTache(False, fTobDet);
     //FV1 - 20/04/2018 : FS#3065 - VIVIANE - En création tâche contrat perte information tâche après recherche adresse.
     //fTobDet.PutEcran(Ecran, fPaDet);
     RefreshAdresse;
     FtobDet.InsertDB(nil, true);
   end;

   RefreshGrid;

   fGSTaches.Row := TobToScreen(ARow-1);

   SetActiveTabSheet('TAB_TACHE');
   SetFocusControl('ATA_LIBELLETACHE1');

end;

Procedure TOF_BTTACHES.InsererModele(TobDepart, TobArrivee : Tob);
var IndZone,ITableLig,Ind : Integer;
    NomZone,NomChamp : String;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
begin
MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();

table := Mcd.getTable(mcd.PrefixetoTable('ATA'));
FieldList := Table.Fields;
FieldList.Reset();
	While FieldList.MoveNext do
    begin
  	NomZone := (FieldList.Current as IFieldCOM).name;
    if pos(NomZone,'ATA_DATEDEBPERIOD;ATA_DATEFINPERIOD;') > 0 then continue;
    ind := Pos ('_',NomZone);
    NomChamp := copy (NomZone,ind+1,255);
    if TOBDepart.FieldExists ('AFM_'+ Nomchamp) then
       TOBArrivee.PutValue(NomZone,TOBDepart.getvalue('AFM_'+NomChamp));
    end;

end;

procedure TOF_BTTACHES.bInsertResOnClick(SEnder: TObject);
var
  vTobRessources, vTobTacheRess, vTobRess: Tob;
  i_ind, i_row: integer;
  stRange: string;
  CodeRess : string;

  procedure AjoutRessource (CodeRessource : string);
  var fTOBRess : TOB;
  		QQ : TQuery;
  begin
  	fTOBRess := TOB.Create ('RESSOURCE',nil,-1);
    QQ := OpenSql ('SELECT ARS_RESSOURCE,ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE="'+CodeRessource+'"',true,1,'',true);
    fTOBREss.SelectDB('',QQ);
    //
    ferme (QQ);
    vTobTacheRess.InitValeurs;
    InitNewRessource(vTobTacheRess, fCurAffaire);
    vTobTacheRess.AddChampSupValeur('ARS_LIBELLE', fTobRess.getvalue('ARS_LIBELLE'), false);
    vTobTacheRess.putvalue('ATR_NUMEROTACHE', fTobDet.GetValue('ATA_NUMEROTACHE'));
    vTobTacheRess.putvalue('ATR_RESSOURCE', fTobRess.getvalue('ARS_RESSOURCE'));
    vTobTacheRess.putvalue('ATR_FONCTION', fTobDet.GetValue('ATA_FONCTION'));
    fGSRes.CacheEdit;
    fGSRes.SynEnabled := false;
    fGSRes.Cells[cInColRes, i_row] := vTobTacheRess.GetValue('ATR_RESSOURCE');
    fGSRes.Cells[cInColQte, i_row] := vTobTacheRess.GetValue('ATR_QTEINITPLA');
    //fGSRes.Cells[cInColQtePC, i_row] := vTobTacheRess.GetValue('ATR_QTEINITIALE');
    fGSRes.Cells[cInColStatut, i_row] := '#ICO#' + RechDom('AFSTATUTRESSOURCE', vTobTacheRess.GetValue('ATR_STATUTRES'), false);
    fGSRes.Cells[cInColResLib, i_row] := vTobTacheRess.GetValue('ARS_LIBELLE');
    fGSRes.MontreEdit;
    fGSRes.SynEnabled := true;
    freeAndNil(fTOBREss);
  end;
begin

  // C.B 07/11/2006
  if fBoNoSuppression then
    PGIBoxAF(TexteMessage[41], ecran.Caption) //Action interdite

  else
  begin
    if (fGSRes.row = 1) or ((fGSRes.row <> 1) and (fGSRes.cells[cInColRes, fGSRes.row] <> '')) then
    begin
      i_row := fGSRes.row;
      if THValComboBox(GetControl('ATA_FONCTION')).Text <> '' then
        stRange := 'ARS_FONCTION1=' + THValComboBox(GetControl('ATA_FONCTION')).value
      else
        stRange := '';
      CodeRess := AFLanceFiche_Rech_Ressource(stRange, 'MULTI');
      if CodeRess <> '' then
      begin
        fTobDet := TOB(fGSTaches.Objects[0, fGSTaches.Row]);
        fTobDet.SetAllModifie(true);

        if (fTobDet <> nil) and (fTobDet.detail.count > 0) then
        begin
          if fTobDet.FindFirst(['ATR_NUMEROTACHE', 'ATR_RESSOURCE'], [fTobDet.GetValue('ATA_NUMEROTACHE'), CodeRess], false) = nil then
          begin
            vTobTacheRess := TOB.Create('TACHERESSOURCE', fTobDet, i_row);
            inc(i_row);
            fGSRes.InsertRow(i_row);
        		AjoutRessource (CodeRess);
          end;
        end else
        begin
          vTobTacheRess := TOB.Create('TACHERESSOURCE', fTobDet, i_row - 1);
        	AjoutRessource (CodeRess);
        end;
      end;
    end;
    SetControlText('ATA_NBRESSOURCE', inttoStr(fGsRes.RowCount - 1)); //mcd 25/08/04 compteur non alimenté, donc alim palnning pas ok sans sortir (11309)
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : suppression des taches et des ressources
Mots clefs ... :
*****************************************************************}
procedure TOF_BTTACHES.bDeleteOnClick(SEnder: TObject);
var
  vStMsg: string;
  
begin

  // C.B 07/11/2006
  if fBoNoSuppression then
    PGIBoxAF(TexteMessage[41], ecran.Caption) //Suppression interdite

  // suppression d'une tache
  else if fGSTaches.Row <> 0 then
  begin
    fTobDet := TOB(fGSTaches.Objects[0, fGSTaches.Row]);
    // on verifie si des éléments planifiés existent
    //traduit dans pgiAskAF
    //if ExistePlanning(fCurAffaire.StAffaire, GetControlText('ATA_NUMEROTACHE')) then
    //  vStMsg := TexteMessage[20] //Un planning existe pour cette tâche.#13#10Confirmez-vous la suppression de cette tâche et du planning correspondant ?
    //else
      if fTobDet.getvalue('ATA_IDENTLIGNE') <> '' then
      vStMsg := TexteMessage[27] // Attention, cette tâche a un lien avec une ligne de la pièce affaire.#13#10Confirmez-vous la suppression de cette tâche ?
    else
      vStMsg := TexteMessage[11]; //Confirmez vous la suppression de la tâche ?

    if (PGIAskAF(vStMsg, ecran.Caption) = mrYes) then
    begin

      // mcd 28/01/04 il faut détruire les enrgt de la table tache jour
      ExecuteSql('DELETE FROM AFTACHEJOUR WHERE ATJ_TYPEJOUR="AFF" AND ATJ_AFFAIRE="'
        + fTobDet.GetValue('ATA_AFFAIRE') + '" AND ATJ_NUMEROTACHE="'
        + IntToStr(fTobDet.GetValue('ATA_NUMEROTACHE')) + '"');

      //mcd 14/11/2005 il faut aussi détruire le mémo correspondant !!!!
      ExecuteSql('DELETE FROM LIENSOLE WHERE LO_TABLEBLOB="ATA" AND LO_IDENTIFIANT="'
        + fTobDet.GetValue('ATA_AFFAIRE') + '/' + IntToStr(fTobDet.GetValue('ATA_NUMEROTACHE')) + '"');

      // il faut aussi détruire le planning correspondant !!!
      ExecuteSQL('DELETE FROM AFPLANNING where APL_NUMEROTACHE =' + GetControlText('ATA_NUMEROTACHE') + ' AND APL_AFFAIRE ="' + fCurAffaire.StAffaire + '"');

      // entete compte
      if fGSTaches.RowCount > 2 then
      begin
        if assigned(fTobDet) then
        begin
          fTobDet.ChangeParent(fTobSup, -1);
          fGSTaches.DeleteRow(fGSTaches.Row);
          if Valider(true, true) then
          begin
            RefreshGrid;
            // repositionnement sur la precedente
            TobToScreen(-1);
          end
        end
        else
        begin
          fGSTaches.DeleteRow(fGSTaches.Row);
          RefreshGrid;
          TobToScreen(-1);
        end;
      end
      else
      begin
        fTobDet := fTobTaches.Detail[0];
        if assigned(fTobDet) then
        begin
          fTobDet.ChangeParent(fTobSup, -1);
          if Valider(true, true) then
          begin
            fGSTaches.DeleteRow(1);
            InsertTache;
          end;
        end;
      end;
    end;
  end;
end;

procedure TOF_BTTACHES.bDeleteResOnClick(SEnder: TObject);
var
  vTob: Tob;
  vInRow: integer;

begin

  // C.B 07/11/2006
  if fBoNoSuppression then
    PGIBoxAF(TexteMessage[41], ecran.Caption) //Suppression interdite

  // suppression d'une ressource
  else
  begin
    if Screen.ActiveControl = fGSRes then
    begin

      if (fTobDet.detail.count = 0) and
        (GetControlText('ATA_NBRESSOURCE') > '1') then
      begin
        if fGSRes.RowCount > 2 then
        begin
          fGSRes.DeleteRow(fGSRes.Row);
          fGSRes.Row := 1;
          UpdateGridRes(True, -1);
          SetFocusControl('GSRES');
        end;
      end

      else
        if (fTobDet.detail.count = 0) and
        (GetControlText('ATA_NBRESSOURCE') = '1') then
        UpdateGridRes(False, 0)

      else
        if ((fGSRes.Cells[fGSRes.Col, fGSRes.Row] = '') and
        (fTobDet.Detail.count = (fGSRes.RowCount - 2))) or
        (PGIAskAF(TexteMessage[12], ecran.Caption) = mrYes) then //Confirmez vous la suppression de la ressource et du planning associé#13#10(si de l''activité a été générée pour ce planning, elle ne sera pas supprimée) ?
      begin

        if (fTobDet.Detail.count >= (fGSRes.RowCount - 1)) then
        begin
          //ajout a la liste de suppression
          SetLength(fUpdateRes, length(fUpdateRes) + 1);
          fUpdateRes[length(fUpdateRes) - 1].StOldRes := fTobDet.Detail[fGSRes.Row - 1].GetValue('ATR_RESSOURCE');
          fUpdateRes[length(fUpdateRes) - 1].StNewRes := fGSRes.Cells[fGSRes.Col, fGSRes.Row];
          fUpdateRes[length(fUpdateRes) - 1].rdAffaire := fCurAffaire;
          fUpdateRes[length(fUpdateRes) - 1].StNumTache := fTobDet.Detail[fGSRes.Row - 1].GetValue('ATR_NUMEROTACHE');
          fUpdateRes[length(fUpdateRes) - 1].InStatut := cInDelete;

          fTobDet := fTobTaches.Detail[fGSTaches.Row - 1];
          vTob := fTobDet.Detail[fGSRes.Row - 1];
          vTob.ChangeParent(fTobSupRes, -1);
        end;

        // pour gerer le pb de raffraichissement
        if fGSRes.RowCount > 2 then
        begin
          fGSRes.ElipsisButton := false;
          if fGSRes.Row = fGSRes.RowCount - 1 then
            vInRow := fGSRes.Row - 2
          else
            vInRow := fGSRes.Row - 1;
          if vInRow = 0 then
            vInRow := 1;

          fGSRes.DeleteRow(fGSRes.Row);
          UpdateGridRes(True, -1);
          // mettre dans la tob les ressources sans sauvegarder
          ScreenToTob;
          SetFocusControl('GSRES');
          fGSRes.ElipsisButton := true;
          fGSRes.Row := vInRow;
        end

        else
        begin
          fGSRes.DeleteRow(fGSRes.Row);
          UpdateGridRes(False, 0);
          fGSRes.Row := 1;
          fGSRes.FixedRows := 1;
          // mettre dans la tob les ressources sans sauvegarder
          ScreenToTob;
          SetFocusControl('GSRES');
          fGSRes.ElipsisButton := false;
        end;
      end;
      if fGSRes.RowCount < 2 then fGSRes.rowCount := 2;
    end
    else
      PGIBoxAF(TexteMessage[14], ecran.Caption);  //Aucune ressource n''est sélectionnée.
  end;
end;   

procedure TOF_BTTACHES.bCompetArticleOnClick(SEnder: TObject);
begin
  {$ifndef OGC}
  //AFLanceFicheAFCOMPETARTICLE(getControlText('ATA_ARTICLE'), '', ActionToString(fAction) + ';MONOENTETE');
  {$endif OGC}
end;

procedure TOF_BTTACHES.bCompetRessourceOnClick(SEnder: TObject);
begin
  // C.B 27/09/2005
  // tester si il y existe une ressource
  if fTobDet.Detail.count > 0 then
    //WLanceFicheWCOMPETRESRCE_FSL(fTobDet.Detail[fGSRes.Row - 1].GetValue('ATR_RESSOURCE'), '', ActionToString(fAction) + ';MONOENTETE')
  else
    PGIBoxAF(TexteMessage[14], ecran.Caption); //Aucune ressource n''est sélectionnée.
end;

procedure TOF_BTTACHES.bRessourceDispoOnClick(SEnder: TObject);
begin
  //AFLanceFicheAFRessourceDispo('', 'CODEARTICLE=' + getControlText('ATA_CODEARTICLE') + ';EXECUTIONDIRECTE');
end;

procedure TOF_BTTACHES.bFermerOnClick(SEnder: TObject);
begin
  inherited;
  fBoClose := True;
end;

procedure TOF_BTTACHES.bMemoOnClick(SEnder: TObject);
var
  vStAction: string;
begin
  if fAction = taConsult then
    vStAction := 'ACTION=CONSULTATION'
  else
    vStAction := 'ACTION=MODIFICATION';
  AglLanceFiche('YY', 'YYLIENSOLE', 'ATA;' + GetControlText('ATA_AFFAIRE') + '/' + GetControlText('ATA_NUMEROTACHE'), '', vStAction);
end;

procedure TOF_BTTACHES.EcranOnCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  canClose := fBoClose;
  if canclose and (fAction <> taConsult) then
  begin

    // pour prendre en compte les modifications à l'écran
    ScreenToTob;
    if (fTOBdet <> nil ) and (fTobDet.modifie) then
    begin
      if PGIAskAF(TexteMessage[21], ecran.Caption) = MrYes then //Voulez-vous valider les modifications ?
        Valider(false, false);
      if fTobTaches <> nil then
      begin
        fTobTaches.Free;
        fTobTaches := nil;
      end;
      if fTobSup <> nil then
      begin
        fTobSup.Free;
        fTobSup := nil;
      end;
      if fTobSupRes <> nil then
      begin
        fTobSupRes.Free;
        fTobSupRes := nil;
      end;
    end;
  end;
end;

procedure TOF_BTTACHES.ATA_QTEINITIALEOnExit(SEnder: TObject);
begin

  {if valeur(GetControlText('ATA_QTEINITIALE')) = 0 then
    SetControlText('ATA_QTEINITIALE', '0')
  else
    begin
    SetControlText('ATA_QTEINITUREF',
    floatToStr(AFConversionUnite(GetControlText('ATA_UNITETEMPS'),
    GetParamsocSecur('SO_AFMESUREACTIVITE', ''),
    valeur(GetControlText('ATA_QTEINITIALE')))));
    if (FCB_RAPPLANNING.checked) and (valeur(GetControlText('ATA_QTEINITPLA')) = 0) then
         SetControlText('ATA_QTEINITPLA', GetControlText('ATA_QTEINITIALE'));
    end;}

end;

procedure TOF_BTTACHES.ATA_QTEINITPLAOnExit(SEnder: TObject);
begin
  if valeur(GetControlText('ATA_QTEINITPLA')) = 0 then
  begin
    SetControlText('ATA_QTEINITPLA', '0');
    SetControlEnabled('RB_MODEGENE2', false);
    SetControlEnabled('RB_MODEGENE1', false);
  end
  else
  begin
    SetControlEnabled('RB_MODEGENE2', (GetControltext('ATA_QTEINITPLA')) <> '0');
    SetControlEnabled('RB_MODEGENE1', (GetControltext('ATA_QTEINITPLA')) <> '0');
    SetControlText('ATA_QTEINITREFPLA',
      floatToStr(AFConversionUnite(GetControlText('ATA_UNITETEMPS'),
      GetParamsocSecur('SO_AFMESUREACTIVITE', ''),
      valeur(GetControlText('ATA_QTEINITPLA')))));
    //if (GetParamsocSecur('SO_BTGESTIONRAF', false)) and (valeur(GetControlText('ATA_QTEINITIALE')) = 0) then
    //  SetControlText('ATA_QTEINITIALE', GetControlText('ATA_QTEINITPLA'));
  end;

  // on recalcul la quantité initiale
  CalculQuantitePla(True);
end;

procedure TOF_BTTACHES.ATA_PUPROnExit(SEnder: TObject);
var
  vDevise: RDevise;
begin
  vDevise.code := fTobDet.GetValue('ATA_DEVISE');
  if (vDevise.code = '') or (vDevise.code = #0) then
    vDevise.Code := V_PGI.DevisePivot;
  GetInfosDevise(vDevise);
end;

procedure TOF_BTTACHES.ATA_PUVENTEHTOnExit(SEnder: TObject);
var
  vDevise: RDevise;
begin
  vDevise.code := fTobDet.GetValue('ATA_DEVISE');
  if (vDevise.code = '') or (vDevise.code = #0) then
    vDevise.Code := V_PGI.DevisePivot;
  GetInfosDevise(vDevise);
end;

procedure TOF_BTTACHES.ArticleOnElipsisClick(SEnder: TObject);
begin

{MODIF fv on interdit d'aller dans le mul article, si le planning existe
  if (GetControlText('ATA_NUMEROTACHE') <> '0')
      and ExistePlanning(fCurAffaire.StAffaire, GetControlText('ATA_NUMEROTACHE')) then
  begin
    PGIBoxAF(TexteMessage[51], ecran.Caption);
    exit;
  end
  else
    begin}

  // C.B 25/10/2005
  // on stocke le libellé de création de l'article par défaut
  // pour le remplacer si on choisit un autre article et qu'on ne l'a pas modifié
  if GetControlText('ATA_LIBELLETACHE1') = fStLibelleParDefaut then SetControlText('ATA_LIBELLETACHE1', '');

  SetControlText('ATA_CODEARTICLE', DispatchRecherche(nil, 1, '', 'GA_CODEARTICLE='
+ GetControltext('ATA_CODEARTICLE') + ';RETOUR_CODEARTICLE=X;GA_TYPEARTICLE='
+ GetControlText('ATA_TYPEARTICLE'), ''));

  RaffraichirArticle;

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 25/10/2005
Modifié le ... :
Description .. : Raffraichir toutes les zones liées à l'article
Mots clefs ... :
*****************************************************************}
procedure TOF_BTTACHES.RaffraichirArticle;
var
  S               : string;
  vStCodeArticle  : string;
  vStTypeArticle  : string;
  vStArticle      : string;
  vStFacturable   : string;
  vStUnite        : string;
  vStLibelle      : string;
  vStStatutPla    : String;
  PR              : Double;
  PV              : Double;
  vStLien	        : String;
  vStEtat         : String;
  vBoPVOK :boolean;
  vTOBAffaires: Tob;
  vTobArticles: Tob;
  vTobvalo : Tob;
  vAFOAssistants: TAFO_Ressources;
  //AFO_Valorisation: TAFValorisationPrix;
begin

  vStCodeArticle := getControlText('ATA_CODEARTICLE');
  if (vStCodeArticle <> '') and
    (vStCodeArticle <> fTobDet.GetValue('ATA_CODEARTICLE')) then
  begin
    if controleCodeArticle(vStCodeArticle, vStTypeArticle, vStArticle, vStFacturable, vStUnite, vStLibelle, PR, PV, vStStatutPla) then
    begin
      SetControlText('ATA_TYPEARTICLE', vStTypeArticle);
      if IsTachePrestation (vStTypeArticle) then  //AB-200605 - FQ 12918
      begin
        SetControlText('ATA_UNITETEMPS', vStUnite);
      end
      else
        SetControlText('ATA_UNITETEMPS', 'J'); // pour frais et marchandise on met à J

      // C.B 27/04/2006
      // fq 11946
      //if GetControlText('ATA_LIBELLETACHE1') = '' then
      SetControlText('ATA_LIBELLETACHE1', vStLibelle);
      SetControlText('ATA_ARTICLE', vStArticle);

      // C.B 26/09/2006
{ModifFV if vStStatutPla <> '' then
        SetControlText('ATA_STATUTPLA', vStStatutPla)
      else
      begin
        vStEtat := getparamsocSecur('SO_AFPLANNINGETAT', '');
        if vStEtat <> '' then
          SetControlText('ATA_STATUTPLA', vStEtat)
        else
          SetControlText('ATA_STATUTPLA', '');
      end;

      if GetParamsocSecur('SO_AFVALOPR', '') = 'ART' then
      begin
        SetControlText('ATA_PUPR', FloattoStr(PR));
        ATA_PUPROnExit(self);
      end;

      if GetParamsocSecur('SO_AFVALOPV', '') = 'ART' then
      begin
        SetControlText('ATA_PUVENTEHT', FloattoStr(PV));
        ATA_PUVentehtOnExit(self);
      end;

      if GetParamsocSecur('SO_AFVALOPV', '') = 'TAR' then
      begin
        vAFOAssistants := TAFO_Ressources.Create;
        vTOBAffaires := TOB.Create('#AFFAIRE', nil, -1);
        vTOBArticles := TOB.Create('#ARTICLE', nil, -1);
        vTOBValo := nil;
        {AFO_Valorisation := TAFValorisationPrix.create(vAFOAssistants,vTobArticles,vTOBAffaires,nil);
        try
        with AFO_Valorisation do
        begin
          gStDevise := fCurAffaire.StDevise;
          gStValoActPR := GetparamSocSecur('SO_AFVALOPR','');
          gStValoActPV := GetparamSocSecur('SO_AFVALOPV','');
          gStAffaire := fCurAffaire.StAffaire;
          gStTiers := fCurAffaire.StTiers;
          gStCodeArticle := vStCodeArticle;
          gbPROK := false;
          gbPVOK := true;
          if MajTOBValo then
            vTOBValo := gTobArticleValorise;
          vBoPVOK := gbPVOK;
        end;
        finally
          AFO_Valorisation.Destroy;
        end;
        if (vTOBValo <> nil) and vBoPVOK then
          SetControlText('ATA_PUVENTEHT', StrF00 (vTobvalo.getdouble('GA_PVHT'), V_PGI.OkDecP));
        ATA_PUVentehtOnExit(self);
        vAFOAssistants.Free;
        vTobArticles.Free;
        vTOBAffaires.Free;
        vTobvalo.free;
      end;}

      SetControlText('ATA_ACTIVITEREPRIS', vStFacturable);

      S := RechDom('AFTFAMARTICLE', GetControltext('ATA_CODEARTICLE'), false);
      if (s <> '') and (s <> 'Error') then
        SetControlText('FAMILLEARTICLE', RechDom('GCFAMILLENIV1', s, false))
      else
        SetControlText('FAMILLEARTICLE', '');
    end
    else
    begin
      PGIBoxAF(TexteMessage[35], ecran.Caption); //L''article n''existe pas.
      Exit;
    end;

		vStLien := FindFamilleTacheLie(vStCodeArticle);
    if vStLien <> '' then
    	SetControlText('ATA_FAMILLETACHE', vStLien);

    cacheZoneSurTypeArt;
  end;
end;

procedure TOF_BTTACHES.UniteOnEnter2(SEnder: TObject);
begin

  // on n'est pas en création
  {if ExistePlanning(fCurAffaire.StAffaire, GetControlText('ATA_NUMEROTACHE')) then
  begin
    PGIBoxAF(TexteMessage[34], ecran.Caption); //Il y a un planning existant sur cette tâche. On ne peut pas modifier l''unité.
    SetFocusControl('ATA_QTEINTERVENT');
  end;}

end;

procedure TOF_BTTACHES.UniteOnEnter(SEnder: TObject);
begin

  // on n'est pas en création
  {if ExistePlanning(fCurAffaire.StAffaire, GetControlText('ATA_NUMEROTACHE')) then
  begin
    PGIBoxAF(TexteMessage[34], ecran.Caption);
    SetFocusControl('ATA_DATEDEBPERIOD');
  end;}

end;

procedure TOF_BTTACHES.ATA_UNITETEMPS_OnChange(SEnder: TObject);
begin
  if getControlText('ATA_UNITETEMPS') = getControlText('ATA_UNITETEMPS_') then
    exit;
  if getControlText('ATA_UNITETEMPS_') = 'J' then
    SetControlCaption('ATA_UNITETEMPS2', traduitGA(TexteMessage[10])) //jour(s)
  else
    SetControlCaption('ATA_UNITETEMPS2', traduitGA(TexteMessage[8])); //heure(s)
  SetControltext('ATA_UNITETEMPS', getControlText('ATA_UNITETEMPS_'));
end;

procedure TOF_BTTACHES.ATA_UNITETEMPSOnChange(SEnder: TObject);
begin
  //C.B 13/10/2005
  //on peut saisir des quantités a planifier même pour les tâches en heure
  {if (GetControlText('ATA_UNITETEMPS') = 'J') then
    SetControlenabled('ATA_RAPPLANNING', True)
  else
  begin
    SetControlChecked('ATA_RAPPLANNING', False);
    SetControlenabled('ATA_RAPPLANNING', False);
  end;
  }

  if getControlText('ATA_UNITETEMPS') <> getControlText('ATA_UNITETEMPS_') then
  begin
    if getControlText('ATA_UNITETEMPS') = 'J' then
      SetControlCaption('ATA_UNITETEMPS2', TexteMessage[10]) //jour(s)
    else
      SetControlCaption('ATA_UNITETEMPS2', TexteMessage[8]); //heure(s)
    SetControltext('ATA_UNITETEMPS_', getControlText('ATA_UNITETEMPS'));

    if (getControlText('ATA_UNITETEMPS') = 'J') then
      SetControlEnabled('ATA_QTEINITPLA', not fBoAFPLANDECHARGE);

    // C.B 06/04/2006
    // en fonction de l'unite (h ou j) on cache les quantites
		CacheZoneSurTypeArt;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 25/10/2005
Modifié le ... :
Description .. : Gestion de la tâche
Mots clefs ... :
*****************************************************************}
function TOF_BTTACHES.ControleZonesOblig: boolean;
var vStDate : string;
begin

  result := true;

  if getControlText('ATA_LIBELLETACHE1') = '' then
     begin
     PGIBoxAF(TexteMessage[5], ecran.Caption); //Le libellé de tâche est obligatoire.
     SetActiveTabSheet('TAB_TACHE');
     SetFocusControl('ATA_LIBELLETACHE1');
     result := false;
     end;

  {if getControlText('ATA_TYPEARTICLE') = '' then
      begin
      PGIBoxAF(TexteMessage[18], ecran.Caption); //Le type article est obligatoire.
      SetActiveTabSheet('TAB_TACHE');
      SetFocusControl('ATA_TYPEARTICLE');
      result := false;
      end;

   if getControlText('ATA_CODEARTICLE') = '' then
      begin
      PGIBoxAF(TexteMessage[19], ecran.Caption); //L''article est obligatoire.
      SetActiveTabSheet('TAB_TACHE');
      SetFocusControl('ATA_CODEARTICLE');
      result := false;
      end;

  //C.B 06/04/2006
  //hygiene office a les décimales
  {if (GetParamSocSecur('SO_AFCLIENT','') = cInClientHygieneOffice) and
      (valeur(GetControlText('ATA_QTEINTERVENT')) <= 0) and
      (not fBoAFPLANDECHARGE) then
      begin
      PGIBoxAF(TexteMessage[54], ecran.Caption); //La durée d''une intervention doit être positive.
      SetActiveTabSheet('TAB_REGLES');
      SetFocusControl('ATA_QTEINTERVENT');
      result := false;
      end;

    // C.B 12/01/2006
    // on ne peut pas remettre les décimales comme avant
    // principe : le generateur permettait de créer avec des décimales, mais la création de 1.5 jours
    // se fait sur 2 jour car les heures ne sont pas gérées dans le générateur.
    // sera remis quand le générateur prendra en compte les heures des demi-journées
    //  else if (StrToFloat(GetControlText('ATA_QTEINTERVENT')) <= 0) and
    //          (not fBoAFPLANDECHARGE) then
  if (GetParamSocSecur('SO_AFCLIENT','') <> cInClientHygieneOffice) and
     ((valeur(GetControlText('ATA_QTEINTERVENT')) <= 0) or
     (int(valeur(GetControlText('ATA_QTEINTERVENT'))) <> valeur(GetControlText('ATA_QTEINTERVENT')))) and
     (not fBoAFPLANDECHARGE) then
     begin
     PGIBoxAF(TexteMessage[22], ecran.Caption); //La durée d''une intervention doit être entière et positive.
     SetActiveTabSheet('TAB_REGLES');
     SetFocusControl('ATA_QTEINTERVENT');
     result := false;
     end;}

  // la fonction n'est obligatoire que pour algoe
  if fBoAFPLANDECHARGE and (getControlText('ATA_FONCTION') = '') and (GetParamsocSecur('SO_AFCLIENT',0) = cInClientAlgoe) then
     begin
     PGIBoxAF(TexteMessage[15], ecran.Caption); //La fonction est obligatoire.
     SetActiveTabSheet('TAB_RESSOURCE');
     SetFocusControl('ATA_FONCTION');
     result := false;
     end;

  if (GetParamsocSecur('SO_BTLIENARTTACHE','') = 'FAM') and (not FamilleArticleUnique) then
     begin
     PGIBoxAF(TexteMessage[28], ecran.Caption); //On ne peut pas avoir deux fois la même famille d''article dans une affaire.
     SetFocusControl('ATA_CODEARTICLE');
     result := false;
     end;

  if (GetParamsocSecur('SO_BTLIENARTTACHE','') = 'ART') and (not ArticleUnique) then
     begin
     PGIBoxAF(TexteMessage[29], ecran.Caption);  //On ne peut pas associer le même article à deux tâches dans une affaire.
     SetFocusControl('ATA_CODEARTICLE');
     result := false;
     end;

  // controle des dates saisie
  if strToDate(getControlText('ATA_DATEDEBPERIOD')) > strToDate(getControlText('ATA_DATEDEBPERIOD_')) then
     begin
     PGIBoxAF(TexteMessage[16], ecran.Caption); //La date de début est supérieure à la date de fin.
     SetActiveTabSheet('TAB_TACHE');
     SetFocusControl('ATA_DATEDEBPERIOD_');
     result := false;
     end;

  vStDate := GetParamsocSecur('SO_BTTACHEDATEDEB','');
  if vStDate = 'ACC' then
     Begin
     if StrToDate(GetControlText('ATA_DATEDEBPERIOD')) < fCurAffaire.DtDateSigne then
        begin
        PGIBoxAF(Format(TexteMessage[56], [dateToStr(fCurAffaire.DtDateSigne)]), ecran.Caption); //La date de fin de tâche ne peut pas être supérieure à la date de fin de l''affaire.
        result := false;
        end;
     end
  else if vStDate = 'AFF' then
     Begin
     if StrToDate(GetControlText('ATA_DATEDEBPERIOD')) < fCurAffaire.DtDebutAff then
        begin
        PGIBoxAF(Format(TexteMessage[56], [dateToStr(fCurAffaire.DtDebutAff)]), ecran.Caption); //La date de fin de tâche ne peut pas être supérieure à la date de fin de l''affaire.
        result := false;
        end;
     end
  else if vstDate = 'FAC' then
     Begin
//     if StrToDate(GetControlText('ATA_DATEDEBPERIOD')) < fCurAffaire.DtDateDebGener then
     if StrToDate(GetControlText('ATA_DATEDEBPERIOD')) < fCurAffaire.DtDebutAff then
        begin
        PGIBoxAF(Format(TexteMessage[56], [dateToStr(fCurAffaire.DtDateDebGener)]), ecran.Caption); //La date de fin de tâche ne peut pas être supérieure à la date de fin de l''affaire.
        result := false;
        end;
     end;

  vStDate := GetParamsocSecur('SO_BTTACHEDATEFIN','');
  if vStDate = 'AFF' then
     Begin
     if StrToDate(GetControlText('ATA_DATEFINPERIOD')) > fCurAffaire.DtFinAff then
        begin
        PGIBoxAF(Format(TexteMessage[13], [dateToStr(fCurAffaire.DtFinAff)]), ecran.Caption); //La date de fin de tâche ne peut pas être supérieure à la date de fin de l''affaire.
        result := false;
        end;
     End
  else if vStDate = 'FAC' then
     Begin
     if StrToDate(GetControlText('ATA_DATEFINPERIOD')) > fCurAffaire.DtDateFinGener then
        begin
        PGIBoxAF(Format(TexteMessage[13], [dateToStr(fCurAffaire.DtDateFinGener)]), ecran.Caption); //La date de fin de tâche ne peut pas être supérieure à la date de fin de l''affaire.
        result := false;
        end;
     End
  else if vStDate = 'LIM' then
     Begin
     if StrToDate(GetControlText('ATA_DATEFINPERIOD')) > fCurAffaire.DtDateLimite then
        begin
        PGIBoxAF(Format(TexteMessage[13], [dateToStr(fCurAffaire.DtDateLimite)]), ecran.Caption); //La date de fin de tâche ne peut pas être supérieure à la date de fin de l''affaire.
        result := false;
        end;
     end
  else if vStDate = 'CLO' then
     Begin
     if StrToDate(GetControlText('ATA_DATEFINPERIOD')) > fCurAffaire.DtDateCloture then
        begin
        PGIBoxAF(Format(TexteMessage[13], [dateToStr(fCurAffaire.DtDateCloture)]), ecran.Caption); //La date de fin de tâche ne peut pas être supérieure à la date de fin de l''affaire.
        result := false;
        end;
     End
  else if vStDate = 'GAR' then
     Begin
     if StrToDate(GetControlText('ATA_DATEFINPERIOD')) > fCurAffaire.DtDateGarantie then
        begin
        PGIBoxAF(Format(TexteMessage[13], [dateToStr(fCurAffaire.DtDateGarantie)]), ecran.Caption); //La date de fin de tâche ne peut pas être supérieure à la date de fin de l''affaire.
        result := false;
        end;
     End
  else if vStDate = 'RES' then
     Begin
     if StrToDate(GetControlText('ATA_DATEFINPERIOD')) > fCurAffaire.DtResil then
        begin
        PGIBoxAF(Format(TexteMessage[13], [dateToStr(fCurAffaire.DtResil)]), ecran.Caption); //La date de fin de tâche ne peut pas être supérieure à la date de fin de l''affaire.
        result := false;
        end;
     end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 07/05/2003
Modifié le ... :   /  /
Description .. : Test si ont n'a pas cette famille d'article dans une
                 autre tache de la meme affaire
Mots clefs ... :
*****************************************************************}
function TOF_BTTACHES.FamilleArticleUnique: Boolean;
var
  i: Integer;
  vStFam1: string;
  vStFam2: string;

begin
  result := True;
  vStFam2 := RechDom('AFTFAMARTICLE', GetControlText('ATA_CODEARTICLE'), false);
  for i := 0 to fTobTaches.detail.count - 1 do 
    if fTobTaches.detail[i] <> fTobDet then
    begin
      vStFam1 := RechDom('AFTFAMARTICLE', fTobTaches.detail[i].Getvalue('ATA_CODEARTICLE'), false);
      if vStFam1 = vStFam2 then
      begin
        result := False;
        break;
      end;
    end;
end;

function TOF_BTTACHES.ArticleUnique: Boolean;
var
  i: integer;

begin
  result := True;
  for i := 0 to fTobTaches.detail.count - 1 do
    if fTobTaches.detail[i] <> fTobDet then
    begin
      if fTobTaches.detail[i].Getvalue('ATA_CODEARTICLE') = GetControlText('ATA_CODEARTICLE') then
      begin
        result := False;
        break;
      end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 07/05/2003
Modifié le ... :   /  /
Description .. : planning jour
                 calcul de la qté restant à planifier : qté initiale -planifié
Mots clefs ... :
*****************************************************************}
procedure TOF_BTTACHES.CalculQuantitePla(pBoEcran: Boolean);
var
  vRdQte: Double;
begin

  if pBoEcran then
    vRdQte := valeur(GetControlText('ATA_QTEINITPLA')) - valeur(GetControlText('ATA_QTEPLANIFPLA'))
  else
    vRdQte := fTobDet.GetValue('ATA_QTEINITPLA') - fTobDet.GetValue('ATA_QTEPLANIFPLA');

  // C.B 22/08/2005
  // on ne passe pas en négatif que pour le client algoe !!!
  // C.B 09/12/2005
  // on passe en négatif pour personne
  if (vRdQte < 0) then vRdQte := 0;

  SetControlText('QTERAP', FloatToStr(vRdQte));

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
xCréé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : controle de l'unicite de la ressource pour la prestation
                 et l'affaire. controle de la saisie de la ressource
Mots clefs ... :
*****************************************************************}
function TOF_BTTACHES.ControleRessource(var pRdPrevuRes: Double): Integer;
var
  i, j: integer;
  vStRes: string;
  dPrevuRes, dPrevuTache: double;
begin
  result := 0;
  dPrevuRes := 0;
  dPrevuTache := 0;
  fGSRes.BeginUpdate;
  //if GetParamsocSecur('SO_BTGESTIONRAF', false) then
  //  dPrevuTache := valeur(GetControlText('ATA_QTEINITIALE'));
  // cas particulier
  // 0 ressource, on ne fait pas le controle
  if GetControlText('ATA_NBRESSOURCE') <> '0' then
    for i := 1 to fGSRes.RowCount - 1 do
    begin
      vStRes := fGSRes.CellValues[cInColRes, i];
      if vStRes = '' then
        result := -2
      else
        for j := i + 1 to fGSRes.RowCount - 1 do
        begin
          if vStRes = fGSRes.CellValues[cInColRes, j] then
          begin
            result := -1;
            break;
          end;
        end;
      if GetParamsocSecur('SO_BTGESTIONRAF', false) then
      begin
{Modif FV dPrevuRes := dPrevuRes + valeur(fGSRes.Cells[cInColQtePC, i]);
        if (dPrevuRes <> 0) and (dPrevuTache <> 0) and (dPrevuRes > dPrevuTache) then
        begin
          result := -3;
          pRdPrevuRes := dPrevuRes;
        end;}
      end;
      if (result < 0) then
        break;
    end;
  if GetParamsocSecur('SO_BTGESTIONRAF', false) then
    if (dPrevuRes <> 0) and (dPrevuTache <> 0) and (dPrevuRes <> dPrevuTache) then
    begin
      result := -3;
      pRdPrevuRes := dPrevuRes;
    end;
  fGSRes.EndUpdate;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_BTTACHES.ScreenToTob: Boolean;
begin

  if fTobDet <> nil then
  begin
    SetControlText('ATA_DATEFINPERIOD', GetControlText('ATA_DATEDEBPERIOD_'));
    FtobDet.PutValue('ATA_FAITPARQUI', GetControlText('CLIENTCONTACT'));
    FtobDet.PutValue('ATA_BTETAT', GetControlText('ATA_BTETAT'));
    fTobDet.GetEcran(Ecran, fPaDet);
    // on ne charge les règles que si modifications
    if fBoModifRegles then GetRegles;
    // si des ressources saisies ou modifiées
    if fTobDet.modifie then ScreenToTobRes;
  end;

  // si tâche modifiée ou si des ressources saisies ou modifiées
  //if fTobDet.modifie or (Length(fUpdateRes) <> 0) or (fStatut <> taConsult) or (fBoModifRegles) then
  if (fTOBdet <> nil ) and (fTobDet.modifie or (Length(fUpdateRes) <> 0)) then
  begin
    result := true;
    fStatut := taModif;
  end
  else
    result := false;
end;

procedure TOF_BTTACHES.GetRegles;
var Days : Word;
    Mnth : Word;
    Annee: Word;
    Mois : String;
    Jour : String;
begin

  if fRB_MODEGENE1.Checked then
     fTobDet.PutValue('ATA_MODEGENE', 1)
  else
     fTobDet.PutValue('ATA_MODEGENE', 2);

  ftobDet.putvalue('ATA_QTEINTERVENT', GetControlText('ATA_QTEINTERVENT'));

  if fRB_QUOTIDIENNE.Checked then
     fTobDet.PutValue('ATA_PERIODICITE', 'Q')
  else if fRB_NBINTERVENTION.Checked then
     fTobDet.PutValue('ATA_PERIODICITE', 'NBI')
  else if fRB_HEBDOMADAIRE.Checked then
     begin
     fTobDet.PutValue('ATA_PERIODICITE', 'S');
     if TCheckBox(GetControl('ATA_JOUR1H')).Checked then AlimTob(1);
     if TCheckBox(GetControl('ATA_JOUR2H')).Checked then AlimTob(2);
     if TCheckBox(GetControl('ATA_JOUR3H')).Checked then AlimTob(3);
     if TCheckBox(GetControl('ATA_JOUR4H')).Checked then AlimTob(4);
     if TCheckBox(GetControl('ATA_JOUR5H')).Checked then AlimTob(5);
     if TCheckBox(GetControl('ATA_JOUR6H')).Checked then AlimTob(6);
     if TCheckBox(GetControl('ATA_JOUR7H')).Checked then AlimTob(7);
     end
  else if fRB_ANNUELLE.Checked then
     begin
     DecodeDate(now, Annee, Mnth, Days);
     //Annee := Years(now);
     Mois  := GetControlText('MOISAN');
     Jour  := GetControlText('JOURAN');
     Mnth  := StrToInt(Mois);
     Days  := StrToInt(Jour);
     fTobDet.PutValue('ATA_PERIODICITE', 'A');
     //mis d'une année fictive dans la date .. seul jour et mois utilisé
     FtobDet.PutValue('ATA_DATEANNUELLE', Encodedate(Annee, Mnth, days));
     end
  else if fRB_MENSUELLE.Checked then
     begin
     fTobDet.PutValue('ATA_PERIODICITE', 'M');
     if fRB_MOISMETHODE1.Checked then
        begin
        fTobDet.PutValue('ATA_MOISMETHODE', '1');
        fTobDet.PutValue('ATA_MOISFIXE', GetControlText('ATA_MOISFIXE0'));
        end
     else if fRB_MOISMETHODE2.Checked then
        begin
        fTobDet.PutValue('ATA_MOISMETHODE', '2');
        fTobDet.PutValue('ATA_MOISSEMAINE', THValComboBox(GetControl('ATA_MOISSEMAINE1')).value);
        fTobDet.PutValue('ATA_MOISFIXE', GetControlText('ATA_MOISFIXE1'));
        fTobDet.PutValue('ATA_MOISJOURLIB', THValComboBox(GetControl('ATA_MOISJOURLIB')).Value);
        end
     else if fRB_MOISMETHODE3.Checked then
        begin
        fTobDet.PutValue('ATA_MOISMETHODE', '3');
        fTobDet.PutValue('ATA_MOISSEMAINE', THValComboBox(GetControl('ATA_MOISSEMAINE2')).value);
        fTobDet.PutValue('ATA_MOISFIXE', GetControlText('ATA_MOISFIXE2'));
        if TCheckBox(GetControl('ATA_JOUR1M')).Checked then AlimTob(1);
        if TCheckBox(GetControl('ATA_JOUR2M')).Checked then AlimTob(2);
        if TCheckBox(GetControl('ATA_JOUR3M')).Checked then AlimTob(3);
        if TCheckBox(GetControl('ATA_JOUR4M')).Checked then AlimTob(4);
        if TCheckBox(GetControl('ATA_JOUR5M')).Checked then AlimTob(5);
        if TCheckBox(GetControl('ATA_JOUR6M')).Checked then AlimTob(6);
        if TCheckBox(GetControl('ATA_JOUR7M')).Checked then AlimTob(7);
        end;
     end;

  fTobDet.PutValue('RAPPTPRCAL', Valeur(GetControlText('RAFPTPRCALC')));
  fTobDet.PutValue('RAPPTPVCAL', Valeur(GetControlText('RAFPTVENTEHTCALC')));

end;

procedure TOF_BTTACHES.PutRegles(pStModele: string);
var
  Tobjour, TobDet: TOB;
  Sql: string;
  ii: integer;
  annee, mm, jj: Word;
begin
  fBoModifRegles := false;
  TobJour := Tob.create('AFTACHEJOUR', nil, -1);
  if trim(pStModele) <> '' then
     begin
     Sql := 'SELECT * FROM AfTacheJour WHERE ATJ_TYPEJOUR="MOD" AND ATJ_MODELETACHE="' + pStModele + '"';
     end
  else
     begin
     Sql := 'SELECT * FROM AfTacheJour WHERE ATJ_TYPEJOUR="AFF" AND ATJ_AFFAIRE="'
          + fTobDet.GetValue('ATA_AFFAIRE') + '" AND ATJ_NUMEROTACHE="'
          + IntTostr(fTobDet.GetValue('ATA_NUMEROTACHE')) + '"';
     end;
  TobJour.LoadDetailDBFromSQL('AFTACHEJOUR', SQL, True);

  // raz de l'écran
  InitNewRegles;

  if fTobDet.GetValue('ATA_MODEGENE') = 1 then
    fRB_MODEGENE1.Checked := True
  else
    fRB_MODEGENE2.Checked := True;

  SetControlText('ATA_QTEINTERVENT', ftobDet.getvalue('ATA_QTEINTERVENT'));
  SetControlText('ATA_MOISJOURFIXE', fTobDet.GetValue('ATA_MOISJOURFIXE'));
  //SetControlText('QTEAPLANIFIERCALC', fTobDet.GetValue('QTEAPLANIFIERCALC'));
  //SetControlText('RAFPTPRCALC', fTobDet.GetValue('RAPPTPRCAL'));
  //SetControlText('RAFPTVENTEHTCALC', fTobDet.GetValue('RAPPTPVCAL'));

  if fTobDet.GetValue('ATA_PERIODICITE') = 'Q' then
    fRB_QUOTIDIENNE.Checked := true

  else if fTobDet.GetValue('ATA_PERIODICITE') = 'NBI' then
    fRB_NBINTERVENTION.Checked := true

  else if fTobDet.GetValue('ATA_PERIODICITE') = 'S' then
  begin
    fRB_HEBDOMADAIRE.Checked := true;
    for ii := 0 to Tobjour.detail.count - 1 do
    begin
      TobDet := TobJour.detail[ii];
      TCheckBox(GetControl('ATA_JOUR' + IntToStr(TobDEt.getvalue('ATj_JOURAPLANIF')) + 'H')).checked := True;
    end;
  end

  else if fTobDet.GetValue('ATA_PERIODICITE') = 'A' then
  begin
    fRB_ANNUELLE.Checked := true;
    Decodedate(FtobDet.GetValue('ATA_DATEANNUELLE'), annee, mm, jj);
    SetCOntrolText('JOURAN', IntToStr(jj));
    SetCOntrolText('MOISAN', IntTostr(mm));
  end

  else if fTobDet.GetValue('ATA_PERIODICITE') = 'M' then
  begin
    fRB_MENSUELLE.Checked := true;
    if fTobDet.GetValue('ATA_MOISMETHODE') = '1' then
    begin
      fRB_MOISMETHODE1.Checked := true;
      SetControlText('ATA_MOISFIXE0', fTobDet.GetValue('ATA_MOISFIXE'));
    end
    else if fTobDet.GetValue('ATA_MOISMETHODE') = '2' then
    begin
      fRB_MOISMETHODE2.Checked := true;
      THValComboBox(GetControl('ATA_MOISSEMAINE1')).value := fTobDet.GetValue('ATA_MOISSEMAINE');
      SetControlText('ATA_MOISFIXE1', fTobDet.GetValue('ATA_MOISFIXE'));
      THValComboBox(GetControl('ATA_MOISJOURLIB')).Value := fTobDet.GetValue('ATA_MOISJOURLIB');
    end
    else if fTobDet.GetValue('ATA_MOISMETHODE') = '3' then
    begin
      fRB_MOISMETHODE3.Checked := true;
      THValComboBox(GetControl('ATA_MOISSEMAINE2')).value := fTobDet.GetValue('ATA_MOISSEMAINE');
      SetControlText('ATA_MOISFIXE2', fTobDet.GetValue('ATA_MOISFIXE'));
      for ii := 0 to Tobjour.detail.count - 1 do
      begin
        TobDet := TobJour.detail[ii];
        TCheckBox(GetControl('ATA_JOUR' + IntToStr(TobDEt.getvalue('ATj_JOURAPLANIF')) + 'M')).checked := True;
      end;
    end;
  end;
  tobjour.free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 18/03/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_BTTACHES.InitNewTache(TobTache: TOB);
begin
  // ABU 16/04/04 init mis dans UtilTaches ds fct InitNewTobTAche pour utilisation PC
  InitNewTobTache(TobTache, GetParamsocSecur('SO_BTPRESTDEFAUT',''), fCurAffaire); //AB-20040416 déplacé dans Utiltaches
  fDtDebutAffaire := TobTache.GetValue('ATA_DATEDEBPERIOD');
  fDtFinAffaire := TobTache.GetValue('ATA_DATEFINPERIOD');
  TobTache.PutValue('RAPPTPRCAL', 0);
  TobTache.PutValue('RAPPTPVCAL', 0);
  TobTache.PutValue('QTEAPLANIFIERCALC', 0);
  SetControlText('QTEAPLANIFIERCALC', '0');
  SetControlText('RAFPTPRCALC', '0');
  SetControlText('RAFPTVENTEHTCALC', '0');
  SetControlText('FAMILLEARTICLE', '');

  // maj du nb de res
  UpdateGridRes(False, 0);
  CacheZoneSurTypeArt;

  // C.B 23/02/2006
  //on supprime l'accès à la création que pour le mode automatique
  if GetParamSocSecur('SO_BTMODEPLANNING', 'TAC') = 'ART' then
    SetControlEnabled('BInsert', False);

  // C.B 28/08/2005
  // on coche en création selon le paramsoc
  if (GetControlText('ATA_UNITETEMPS') = 'J') then
     begin
     SetControlChecked('ATA_RAPPLANNING', GetParamsocSecur('SO_AFGENELIMIT', false));
     SetControlenabled('ATA_RAPPLANNING', True);
     end
  else
     SetControlenabled('ATA_RAPPLANNING', False);

  SetControlEnabled('TAB_TACHE', True);
  SetControlEnabled('TAB_ADRESSE', True);
  SetControlEnabled('TAB_RESSOURCE', True);
  SetControlEnabled('TAB_REGLES', True);
  SetControlEnabled('TAB_COMPLEMENT', True);
  SetControlEnabled('PAGETACHE', True);
  SetControlVisible('BDelete', True);

  // C.B 27/09/2005
  // raffraichir la quantité rap
  CalculQuantitePla(False);

  // C.B 25/10/2005
  // on stocke le libellé de création de l'article par défaut
  // pour le remplacer si on choisit un autre article et qu'on ne l'a pas modifié
  fStLibelleParDefaut := GetControlText('ATA_LIBELLETACHE1');

  //FV1 : 14/03/20018 - FS#3006 - ESPACS - En création de tâches contrat les lignes sont doublées
  SetControlVisible('BINSERT',        False);
  SetControlVisible('BDUPLIQUE',      False);
  SetControlVisible('BINSERTMODELE',  False);
  SetControlVisible('BDELETE',        False);
  SetControlVisible('BGENERER',       False);


end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : affiche le contenu de la tob a l'ecran
Description .. : et retourne l'indice dans la grille
Mots clefs ... :
*****************************************************************}
function TOF_BTTACHES.TobToScreen(Row: Integer): integer;
begin

  if Row <> -1 then
     begin
     fTobDet := fTobTaches.Detail[Row];
     result := Row + 1;
     end
  else
     begin
     fTobDet := TOB(fGSTaches.Objects[0, fGSTaches.Row]);
     result := fGSTaches.Row;
     end;

  fATATermineChange := false;
  fTobDet.PutEcran(Ecran, fPaDet);
  ArticleOnexit(self);

  fDtDebPeriod := fTobDet.getvalue('ATA_DATEDEBPERIOD');
  fDtFinPeriod := fTobDet.getvalue('ATA_DATEFINPERIOD');

  SetControlText('ATA_DATEDEBPERIOD', DateTimeToStr(fDtDebPeriod));
  SetControlText('ATA_DATEDEBPERIOD_', DateTimeToStr(fDtFinPeriod));
  SetControlText('ATA_DATEFINPERIOD', DateTimeToStr(fDtFinPeriod));

  SetControlText('ATA_BTETAT', fTobDet.GetString('ATA_BTETAT'));

  SetControlVisible('BIntervenant', false);
  SetControlVisible('BCompetArticle', false);
  SetControlVisible('BPlanning', false);
  SetControlVisible('BPlanDeCharge', false);

  if fTobDet.GetString('ATA_TERMINE')='X' then
     Begin
     SetControlEnabled('TAB_TACHE', False);
     SetControlEnabled('TAB_ADRESSE', False);
     SetControlEnabled('TAB_RESSOURCE', False);
     SetControlEnabled('TAB_REGLES', False);
     SetControlEnabled('TAB_COMPLEMENT', False);
     SetControlVisible('BDelete', false);
     SetControlVisible('TATA_LASTDATEGENE', True);
     SetControlVisible('ATA_LASTDATEGENE', True);
     end
  else
     Begin
     SetControlEnabled('TAB_TACHE', True);
     SetControlEnabled('TAB_ADRESSE', True);
     SetControlEnabled('TAB_RESSOURCE', True);
     SetControlEnabled('TAB_REGLES', True);
     SetControlEnabled('TAB_COMPLEMENT', True);
     SetControlEnabled('PAGETACHE', True);
     SetControlVisible('BDelete', True);
     SetControlVisible('TATA_LASTDATEGENE', False);
     SetControlVisible('ATA_LASTDATEGENE', False);
     end;

  PutRegles('');

  CalculQuantitePla(False);

  // C.B 29/11/2005
  // on peut saisir la quantité initiale du plan de charge
  // tant qu'il n'y a pas eu d'écart de saisi
  //if (GetControlText('ATA_ECARTQTEINIT') = '0') and
  //  (GetParamsocSecur('SO_BTMODEPLANNING','') <> 'ART') and
  //  (Faction <> taConsult) then
  //  SetControlEnabled('ATA_QTEINITIALE', True)
  //else
  //  SetControlEnabled('ATA_QTEINITIALE', False);

  // C.B 08/11/2005
  // adresse
  RefreshAdresse;

  LectureTypeAction(GetControlText('ATA_BTETAT'));
  LectureContact(GetControlText('CLIENTCONTACT'));

  // les taches ressources
  TobResToScreen;

  if fTobDet.Detail.Count <> 0 then
    UpdateGridRes(False, fTobDet.Detail.Count)
  else
    UpdateGridRes(False, 0);

  // ajout de la gestion du type de tache
  RefreshLignePiece; //AB-LigneAffaire-Champs non modifiables

end;

procedure TOF_BTTACHES.TobResToScreen;
var
  i: Integer;
begin

  RefreshGridRes;
  for i := 0 to fTobDet.Detail.Count - 1 do
  begin
    fGSRes.Cells[cInColRes, i + 1] := fTobDet.Detail[i].GetValue('ATR_RESSOURCE');
    fGSRes.Cells[cInColQte, i + 1] := fTobDet.Detail[i].GetValue('ATR_QTEINITPLA');
    //fGSRes.Cells[cInColQtePC, i + 1] := fTobDet.Detail[i].GetValue('ATR_QTEINITIALE');
    //fGSRes.Cells[cInColStatut, i + 1] := RechDom('AFSTATUTRESSOURCE', fTobDet.Detail[i].GetValue('ATR_STATUTRES'), false);
    fGSRes.Cells[cInColStatut, i + 1] := '#ICO#' + RechDom('AFSTATUTRESSOURCE', fTobDet.Detail[i].GetValue('ATR_STATUTRES'), false);

    if (fTobDet.Detail[i].GetValue('ARS_LIBELLE') = '') or
      (fTobDet.Detail[i].GetValue('ARS_LIBELLE') = #0) then
      fGSRes.Cells[cInColResLib, i + 1] := RessourceLib(fGSRes.Cells[cInColRes, i + 1])
    else
      fGSRes.Cells[cInColResLib, i + 1] := fTobDet.Detail[i].GetValue('ARS_LIBELLE');

    // affiche calcul reste à faire PC
    if GetParamsocSecur('SO_BTGESTIONRAF', false) then
      //fGSRes.Cells[cInColQteRAF, i + 1] := fTobDet.Detail[i].GetValue('RAFQTERES');
  end;

  if fTobDet.Detail.Count <> 0 then
  begin
    fGSRes.Rowcount := fTobDet.Detail.Count + 1;
    fGSRes.enabled := True;
  end
  else
  begin
    fGSRes.Rowcount := 2;
    fGSRes.enabled := False;
  end;
end;

procedure TOF_BTTACHES.ScreenToTobRes;
var
  i: Integer;
  vTobRes: Tob;

begin

  for i := 1 to fGSRes.rowcount - 1 do
  begin

    vTobRes := fTobDet.FindFirst(['ATR_RESSOURCE'], [fGSRes.Cells[cInColRes, i]], true);

    if fGSRes.Cells[cInColRes, i] <> '' then
    begin
      if (vTobRes = nil) then
      begin
        vTobRes := TOB.Create('TACHERESSOURCE', fTobDet, -1);
        initNewRessource(vTobRes, fCurAffaire); //AB-20040416 déplacé dans Utiltaches
        vTobRes.AddChampSup('ARS_LIBELLE', True);
      end;

      vTobRes.putValue('ATR_RESSOURCE', fGSRes.Cells[cInColRes, i]);
      vTobRes.putValue('ARS_LIBELLE', fGSRes.Cells[cInColResLib, i]);
      //vTobRes.putValue('ATR_QTEINITIALE', valeur(fGSRes.Cells[cInColQtePC, i]));
      vTobRes.putValue('ATR_QTEINITPLA', valeur(fGSRes.Cells[cInColQte, i]));
      //vTobRes.putValue('ATR_QTEINITUREF', floatToStr(AFconversionUnite(GetControlText('ATA_UNITETEMPS'),
      //  GetParamsocSecur('SO_AFMESUREACTIVITE',''),
      //  valeur(fGSRes.Cells[cInColQtePC, i]))));
      //vTobRes.putValue('ATR_QTEINITREFPLA', floatToStr(AFconversionUnite(GetControlText('ATA_UNITETEMPS'),
      //  GetParamsocSecur('SO_AFMESUREACTIVITE',''),
      //  valeur(fGSRes.Cells[cInColQte, i]))));
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/04/2002
Modifié le ... :   /  /
Description .. : recherche l'indice dans la tob de la tache courante
                 dans la grille
Mots clefs ... :
*****************************************************************}
function TOF_BTTACHES.IndiceTobTacheCourante: Integer;
var
  i: Integer;

begin

  result := 1;
  if (fTobTaches <> nil) and (fTobTaches.Detail.count > 0) then
  begin
    for i := 0 to fTobTaches.Detail.count - 1 do
    begin
      if fTobDet = nil then
      begin
        fTobDet := fTobTaches.Detail[i];
        fTobDet.modifie := false;
        fATATermineChange := false;
      end
      else
        fTobDet := fTobTaches.Detail[i];

      // il n'y a pas de tache selectionné -> on se positionne sur la premiere
      if (fTobDet <> nil) and (fStOpenNumTache = '') then
      begin
        fStOpenNumTache := varAsType(fTobDet.GetValue('ATA_NUMEROTACHE'), varString);
        result := i;
        exit;
      end
      else
        if (fTobDet <> nil) and (varAsType(fTobDet.GetValue('ATA_NUMEROTACHE'), varString) = fStOpenNumTache) then
      begin
        result := i;
        exit;
      end;
    end;
  end;
end;

procedure TOF_BTTACHES.fRB_MODEGENE1OnClick(Sender: TObject);
begin
  fRB_MODEGENE2.checked := not fRB_MODEGENE1.Checked;
end;

procedure TOF_BTTACHES.fRB_MODEGENE2OnClick(Sender: TObject);
begin
  fRB_MODEGENE1.Checked := not fRB_MODEGENE2.Checked;
  fBoModifRegles := true;
end;

procedure TOF_BTTACHES.fRB_HEBDOMADAIREOnClick(Sender: TObject);
begin
  fRB_QUOTIDIENNE.checked := not fRB_HEBDOMADAIRE.checked;
  fRB_MENSUELLE.checked := not fRB_HEBDOMADAIRE.checked;
  fRB_ANNUELLE.checked := not fRB_HEBDOMADAIRE.checked;
  fRB_NBINTERVENTION.checked := not fRB_HEBDOMADAIRE.checked;
  TCheckBox(GetControl('ATA_JOUR1H')).checked := not fRB_HEBDOMADAIRE.checked;
  TCheckBox(GetControl('ATA_JOUR2H')).checked := not fRB_HEBDOMADAIRE.checked;
  TCheckBox(GetControl('ATA_JOUR3H')).checked := not fRB_HEBDOMADAIRE.checked;
  TCheckBox(GetControl('ATA_JOUR4H')).checked := not fRB_HEBDOMADAIRE.checked;
  TCheckBox(GetControl('ATA_JOUR5H')).checked := not fRB_HEBDOMADAIRE.checked;
  TCheckBox(GetControl('ATA_JOUR6H')).checked := not fRB_HEBDOMADAIRE.checked;
  TCheckBox(GetControl('ATA_JOUR7H')).checked := not fRB_HEBDOMADAIRE.checked;
  fPC_FREQUENCE.ActivePage := fTS_HEBDOMADAIRE;
  SetControlEnabled('RB_MODEGENE2', valeur(GetControltext('ATA_QTEINITPLA')) <> 0);
  SetControlEnabled('RB_MODEGENE1', valeur(GetControltext('ATA_QTEINITPLA')) <> 0);
  fBoModifRegles := true;
end;

procedure TOF_BTTACHES.fRB_QUOTIDIENNEOnClick(Sender: TObject);
begin
  fRB_HEBDOMADAIRE.checked := not fRB_QUOTIDIENNE.checked;
  fRB_MENSUELLE.checked := not fRB_QUOTIDIENNE.checked;
  fRB_ANNUELLE.checked := not fRB_QUOTIDIENNE.checked;
  fRB_NBINTERVENTION.checked := not fRB_QUOTIDIENNE.checked;
  fPC_FREQUENCE.ActivePage := fTS_QUOTIDIENNE;
  SetControlEnabled('RB_MODEGENE2', valeur(GetControltext('ATA_QTEINITPLA')) <> 0);
  SetControlEnabled('RB_MODEGENE1', valeur(GetControltext('ATA_QTEINITPLA')) <> 0);
  fBoModifRegles := true;
end;

procedure TOF_BTTACHES.fRB_MENSUELLEOnClick(Sender: TObject);
begin
  fRB_HEBDOMADAIRE.checked := not fRB_MENSUELLE.checked;
  fRB_QUOTIDIENNE.checked := not fRB_MENSUELLE.checked;
  fRB_ANNUELLE.checked := not fRB_MENSUELLE.checked;
  fRB_NBINTERVENTION.checked := not fRB_MENSUELLE.checked;
  fPC_FREQUENCE.ActivePage := fTS_MENSUELLE;
  TCheckBox(GetControl('ATA_JOUR1M')).checked := not fRB_MENSUELLE.checked;
  TCheckBox(GetControl('ATA_JOUR2M')).checked := not fRB_MENSUELLE.checked;
  TCheckBox(GetControl('ATA_JOUR3M')).checked := not fRB_MENSUELLE.checked;
  TCheckBox(GetControl('ATA_JOUR4M')).checked := not fRB_MENSUELLE.checked;
  TCheckBox(GetControl('ATA_JOUR5M')).checked := not fRB_MENSUELLE.checked;
  TCheckBox(GetControl('ATA_JOUR6M')).checked := not fRB_MENSUELLE.checked;
  TCheckBox(GetControl('ATA_JOUR7M')).checked := not fRB_MENSUELLE.checked;
  SetControlEnabled('RB_MODEGENE2', valeur(GetControltext('ATA_QTEINITPLA')) <> 0);
  SetControlEnabled('RB_MODEGENE1', valeur(GetControltext('ATA_QTEINITPLA')) <> 0);
  fBoModifRegles := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... :
Modifié le ... : 29/07/2005
Description .. : Mode de génération au plus tard pas disponible
                 avec le mode de génération annuelle
Mots clefs ... :
*****************************************************************}
procedure TOF_BTTACHES.fRB_ANNUELLEOnClick(Sender: TObject);
begin
  fRB_HEBDOMADAIRE.checked := not fRB_ANNUELLE.checked;
  fRB_QUOTIDIENNE.checked := not fRB_ANNUELLE.checked;
  fRB_MENSUELLE.checked := not fRB_ANNUELLE.checked;
  fRB_NBINTERVENTION.checked := not fRB_ANNUELLE.checked;
  fPC_FREQUENCE.ActivePage := fTS_ANNUELLE;
  SetControlEnabled('RB_MODEGENE2', False);
  SetControlEnabled('RB_MODEGENE1', False);
  fBoModifRegles := true;
end;

procedure TOF_BTTACHES.fRB_NBINTERVENTIONOnClick(Sender: TObject);
begin
  fRB_HEBDOMADAIRE.checked := not fRB_NBINTERVENTION.checked;
  fRB_QUOTIDIENNE.checked := not fRB_NBINTERVENTION.checked;
  fRB_MENSUELLE.checked := not fRB_NBINTERVENTION.checked;
  fRB_ANNUELLE.checked := not fRB_NBINTERVENTION.checked;
  fPC_FREQUENCE.ActivePage := fTS_NBINTERVENTION;
  SetControlEnabled('RB_MODEGENE2', false);
  SetControlEnabled('RB_MODEGENE1', false);
  fBoModifRegles := true;
end;

procedure TOF_BTTACHES.fRB_MOISMETHODE1OnClick(SEnder: TObject);
begin

  fRB_MOISMETHODE2.checked := not fRB_MOISMETHODE1.checked;
  fRB_MOISMETHODE3.checked := not fRB_MOISMETHODE1.checked;

  SetControlEnabled('ATA_MOISJOURFIXE', True);
  SetControlEnabled('ATA_MOISFIXE0', True);
  SetControlEnabled('ATA_MOISSEMAINE1', false);
  SetControlEnabled('ATA_MOISJOURLIB', false);
  SetControlEnabled('ATA_MOISFIXE1', false);
  SetControlEnabled('ATA_MOISSEMAINE2', false);
  SetControlEnabled('ATA_MOISFIXE2', false);

  SetControlEnabled('ATA_JOUR1M', false);
  SetControlEnabled('ATA_JOUR2M', false);
  SetControlEnabled('ATA_JOUR3M', false);
  SetControlEnabled('ATA_JOUR4M', false);
  SetControlEnabled('ATA_JOUR5M', false);
  SetControlEnabled('ATA_JOUR6M', false);
  SetControlEnabled('ATA_JOUR7M', false);

  fBoModifRegles := true;
end;

procedure TOF_BTTACHES.fRB_MOISMETHODE2OnClick(SEnder: TObject);
begin

  fRB_MOISMETHODE1.checked := not fRB_MOISMETHODE2.checked;
  fRB_MOISMETHODE3.checked := not fRB_MOISMETHODE2.checked;

  SetControlEnabled('ATA_MOISJOURFIXE', false);
  SetControlEnabled('ATA_MOISFIXE0', false);
  SetControlEnabled('ATA_MOISSEMAINE1', true);
  SetControlEnabled('ATA_MOISJOURLIB', true);
  SetControlEnabled('ATA_MOISFIXE1', true);
  SetControlEnabled('ATA_MOISSEMAINE2', false);
  SetControlEnabled('ATA_MOISFIXE2', false);

  SetControlEnabled('ATA_JOUR1M', false);
  SetControlEnabled('ATA_JOUR2M', false);
  SetControlEnabled('ATA_JOUR3M', false);
  SetControlEnabled('ATA_JOUR4M', false);
  SetControlEnabled('ATA_JOUR5M', false);
  SetControlEnabled('ATA_JOUR6M', false);
  SetControlEnabled('ATA_JOUR7M', false);

  fBoModifRegles := true;
end;

procedure TOF_BTTACHES.fRB_MOISMETHODE3OnClick(SEnder: TObject);
begin
  fRB_MOISMETHODE1.checked := not fRB_MOISMETHODE3.checked;
  fRB_MOISMETHODE2.checked := not fRB_MOISMETHODE3.checked;
  SetControlEnabled('ATA_MOISJOURFIXE', false);
  SetControlEnabled('ATA_MOISFIXE0', false);
  SetControlEnabled('ATA_MOISSEMAINE1', false);
  SetControlEnabled('ATA_MOISJOURLIB', false);
  SetControlEnabled('ATA_MOISFIXE1', false);
  SetControlEnabled('ATA_MOISSEMAINE2', true);
  SetControlEnabled('ATA_MOISFIXE2', true);
  SetControlEnabled('ATA_JOUR1M', true);
  SetControlEnabled('ATA_JOUR2M', true);
  SetControlEnabled('ATA_JOUR3M', true);
  SetControlEnabled('ATA_JOUR4M', true);
   SetControlEnabled('ATA_JOUR5M', true);
  SetControlEnabled('ATA_JOUR6M', true);
  SetControlEnabled('ATA_JOUR7M', true);
  fBoModifRegles := true;
end;

procedure TOF_BTTACHES.TraitementArgument(pStArgument: string);
var
  Tmp: string;
  champ: string;
  valeur: string;

begin
  // traitement des arguments
  Tmp := (Trim(ReadTokenSt(pStArgument)));

  while (Tmp <> '') do
     begin
     if Tmp <> '' then
        begin
        DecodeArgument(Tmp, Champ, valeur);
        if Champ = 'ATA_AFFAIRE' then fCurAffaire.StAffaire := valeur
        else if Champ = 'ATA_NUMEROTACHE' then fStOpenNumTache := valeur
     else if Champ = 'ATA_IDENTLIGNE' then fStLignePiece := findEtreplace(valeur, '|', ';', true)
     else if Champ = 'ZOOMTACHE' then
        begin
        SetControlEnabled('mnPlanningRessource', False);
        SetControlEnabled('mnPlanningTache', False);
        SetControlEnabled('BPLANDECHARGE', False);
        end
     else if Champ = 'ACTION' then
        begin
        if valeur = 'MODIFICATION' then  fAction := taModif
        else if valeur = 'CONSULTATION' then
           begin
           fAction := taConsult;
           ZonesLectureSeule;
           end;
        end
     else if Champ = 'SUPPRESSION_INTERDITE' then
        fBoNoSuppression := True;
     end;
     Tmp := (Trim(ReadTokenSt(pStArgument)));
  end;

end;

procedure TOF_BTTACHES.ZonesLectureSeule;
begin

  //Boutons généraux
  SetControlEnabled('BINSERT', False);
  SetControlEnabled('BINSERTMODELE', False);
  SetControlEnabled('BDELETE', False);
  SetControlEnabled('BDUPLIQUE', False);
  SetControlEnabled('BDEFAIRE', False);
  SetControlEnabled('BTN_SELECTION', False);
  SetControlEnabled('BTN_EFFACER', False);
  SetControlEnabled('BMENU', False);
  SetControlEnabled('BTLIGNE', False);
  SetControlEnabled('BGENERER', False);
  SetControlEnabled('BPLANNING', False);
  SetControlEnabled('BPLANDECHARGE', False);

//  SetControlEnabled('BINTERVENANT', False);
//  SetControlEnabled('BCOMPETARTICLE', False);

  // boutons ressource invisibles
  SetControlVisible('BINSERTRES', False);
  SetControlVisible('BDELETERES', False);
  SetControlVisible('BCALCULRAPRES', False);
  SetControlVisible('BCOMPETRESSOURCE', False);
  SetControlVisible('BRESSOURCE', False);

  //Onglets
  SetControlEnabled('TAB_TACHE', False);
  SetControlEnabled('TAB_REGLES', False);
  SetControlEnabled('TAB_RESSOURCE', False);

end;


procedure TOF_BTTACHES.LoadAffaire(pInIndiceTob: Integer);
var
  vBoGestionRAF: Boolean;
begin

  // Gestion des Grid
  InitGrid;
  initGridRes;

  // chargement du tiers et formatage des libelles affaire et tiers
  AffichageEntete;

  if assigned(fTobTaches) then fTobTaches.Free;
  if assigned(fTobSup) then fTobSup.Free;
  if assigned(fTobSupRes) then fTobSupRes.Free;

  fTobTaches := nil;
  fTobSup := nil;
  fTobSupRes := nil;

  // Chargement des taches
  fTobTaches := TOB.Create('Liste des taches', nil, -1);
  fTobSup := TOB.Create('Liste taches sup', nil, -1);
  fTobSupRes := TOB.Create('Liste des ressources modifiées', nil, -1);

  // des tâches existent ...
  // C.B Le calcul du RAF est conditionné au paramsoc PDC
  vBoGestionRAF := GetParamsocSecur('SO_BTGESTIONRAF', false);
  if ChargeLesTaches(fTobTaches, fCurAffaire.StAffaire, '', vBoGestionRAF) then
  begin
    fTobTaches.PutGridDetailOnListe(fGSTaches, 'AFLISTETACHES');
    fGSTaches.Fixedcols := 1;
    TFVierge(Ecran).Hmtrad.ResizeGridColumns(fGSTaches);
  end
  else
    fStatut := taCreat; // si aucune tache, on entre en creation

  TraiteLignePiece;

  if (fStatut = taCreat) then
     InsertTache
  // affiche la ligne sélectionnée positionnement dans la grille
  else if fTobTaches.Detail.Count > 0 then
       begin
       // calcul des quantités reste à faire des ressources pour le plan de charge
       //if GetParamsocSecur('SO_BTGESTIONRAF', false) then CalculRafRes;
       if pInIndiceTob = 0 then
          fGSTaches.Row := TobToScreen(IndiceTobTacheCourante)
       else
          fGSTaches.Row := TobToScreen(pInIndiceTob);
       // on ne peut pas changer l'article si du planning existe
       //SetControlEnabled('ATA_CODEARTICLE', not ExistePlanning(fCurAffaire.StAffaire, GetControlText('ATA_NUMEROTACHE')));
       //SetControlEnabled('ATA_TYPEARTICLE', GetControlEnabled('ATA_CODEARTICLE'));
       SetControlText('ATA_TYPEARTICLE', 'PRE');
       SetControlEnabled('ATA_TYPEARTICLE', False);
       end;

  SetActiveTabSheet('TAB_TACHE');
  SetFocusControl('ATA_LIBELLETACHE1');

  // on ne calcule pas le reste a planifier au chargement das affaires
  if not fBoAFPLANDECHARGE then CalculQuantitePla(False);

  if fStLignePiece <> '' then
     begin
     if assigned(fTobModele) then FreeAndNil(fTobModele);
     fStLignePiece := '';
     SetControlVisible('BTLIGNE', False);
     end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 24/04/2002
Modifié le ... :
Description .. :
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOF_BTTACHES.GestionEcran;
var
  Combo : THValComboBox;
  Bt    : TToolBarButton97;
  vMenu : TMenuItem;
  Ed    : THEdit;
  pcPage: TPageControl;
  cb    : TCheckBox;
begin

  // gestion des touches
  TFVierge(Ecran).OnKeyDown := FormKeyDown;

  // Gestion des Boutons
  Bt := TToolBarButton97(GetControl('BInsert'));
  if (Bt <> nil) then
    Bt.Onclick := BInsertOnClick;

  Bt := TToolBarButton97(GetControl('BInsertModele')); // AB-Modeletache-Création à partir d'un modele
  if (Bt <> nil) then
    Bt.Onclick := bInsertModeleOnClick;

  Bt := TToolBarButton97(GetControl('BDuplique')); // Duplication de la tache sélectionnée en une autre...
  if (Bt <> nil) then
    Bt.Onclick := bDupliqueOnClick;

  Bt := TToolBarButton97(GetControl('BInsertRes'));
  if (Bt <> nil) then
    Bt.Onclick := BInsertResOnClick;

  Bt := TToolBarButton97(GetControl('BDelete'));
  if (Bt <> nil) then
    Bt.Onclick := BDeleteOnClick;

  Bt := TToolBarButton97(GetControl('BDeleteRes'));
  if (Bt <> nil) then
    Bt.Onclick := BDeleteResOnClick;

  Bt := TToolbarButton97(GetControl('BFERME'));
  if (Bt <> nil) then
    Bt.Onclick := bFermerOnClick;

  Bt := TToolbarButton97(GetControl('BMEMO'));
  if (Bt <> nil) then
    Bt.Onclick := bMemoOnClick;

  Bt := TToolbarButton97(GetControl('BRESSOURCE'));
  if bt <> nil then
    Bt.Onclick := BRESSOURCEOnClick;

  Bt := TToolbarButton97(GetControl('BCALCULRAPRES'));
  if bt <> nil then
    Bt.Onclick := BCALCULRAPRESOnClick;
  //SetCOntrolVisible('BCALCULRAFRES', false);

  Bt := TToolbarButton97(GetControl('BTLIGNE'));
  if bt <> nil then
    Bt.Onclick := BTLIGNEOnClick;

  Bt := TToolbarButton97(GetControl('BINTERVENANT'));
  if bt <> nil then
    Bt.Onclick := BIntervenantsOnClick;

  Cb := TCheckBox(GetControl('ATA_RAPPLANNING'));
  if cb <> nil then
    Cb.Onclick := ATA_RAPPLANNINGOnClick;

  Bt := TToolbarButton97(GetControl('BCOMPETARTICLE'));
  if bt <> nil then
    Bt.Onclick := BCompetArticleOnClick;

  vMenu := TMenuItem(GetControl('mnGENEAFFAIRE'));
  if vMenu <> nil then
     vMenu.OnClick := mnGeneAffaire_OnClick;

  vMenu := TMenuItem(GetControl('MnContact'));
  if vMenu <> nil then
     vMenu.OnClick := MnContact_OnClick;

  vMenu := TMenuItem(GetControl('MnAdressesInt'));
  if vMenu <> nil then
     vMenu.OnClick := MnAdressesInt_OnClick;

  vMenu := TMenuItem(GetControl('mnGENETACHE'));
  if vMenu <> nil then
    vMenu.OnClick := mnGeneTache_OnClick;

  vMenu := TMenuItem(GetControl('mnPlanChargeQTE'));
  if vMenu <> nil then
    vMenu.OnClick := mnPlanChargeQuantite_OnClick;

  vMenu := TMenuItem(GetControl('mnPlanChargePRX'));
  if vMenu <> nil then
    vMenu.OnClick := mnPlanChargePrixRevient_OnClick;

  vMenu := TMenuItem(GetControl('mnPlanChargePRV'));
  if vMenu <> nil then
    vMenu.OnClick := mnPlanChargePrixVente_OnClick;

  vMenu := TMenuItem(GetControl('MnPlanChargeDispo'));
  if vMenu <> nil then
    vMenu.OnClick := MnPlanChargeDispo_OnClick;

  vMenu := TMenuItem(GetControl('mnPlanningViewer'));
  if vMenu <> nil then
    vMenu.OnClick := mnPlanningViewer_OnClick;

  vMenu := TMenuItem(GetControl('mnPlanningRes'));
  if vMenu <> nil then
    vMenu.OnClick := mnPlanningRes_OnClick;

  vMenu := TMenuItem(GetControl('mnPlanningTache'));
  if vMenu <> nil then
    vMenu.OnClick := mnPlanningTache_OnClick;

  vMenu := TMenuItem(GetControl('mnPlanningRessource'));
  if vMenu <> nil then
    vMenu.OnClick := mnPlanningRessource_OnClick;

  vMenu := TMenuItem(GetControl('MnPlanningOccup')); //AB-200509- Occupation par ressource sans détail affaire
  if vMenu <> nil then
    vMenu.OnClick := MnPlanningOccupRess_OnClick;

  vMenu := TMenuItem(GetControl('MnPlanningOccupDetail')); //AB-200509- Occupation par ressource avec détail affaire
  if vMenu <> nil then
    vMenu.OnClick := MnPlanningOccupRessDetaille_OnClick;

  vMenu := TMenuItem(GetControl('MnDetailPlanning'));
  if vMenu <> nil then
    vMenu.OnClick := MnDetailPlanning_OnClick;

  vMenu := TMenuItem(GetControl('MnPlanningActivite'));
  if vMenu <> nil then
    vMenu.OnClick := MnPlanningActivite_OnClick;

  vMenu := TMenuItem(GetControl('mnCompetRessource'));
  if vMenu <> nil then
    vMenu.Onclick := bCompetRessourceOnClick;

  vMenu := TMenuItem(GetControl('mnRecherche'));
  if vMenu <> nil then
    vMenu.Onclick := bRessourceDispoOnClick;

  pcPage := TPageControl(getControl('PAGETACHE'));
  if pcPage <> nil then
    pcpage.OnChange := PAGETACHEChange;

  // Evènements de la tâche en cours
  //Combo := THValComboBox(GetControl('ATA_PERIODICITE'));
  //If Combo <> Nil then Combo.Onclick:=ATA_PERIODICITEOnClick;

  fED_AFFAIRE := THEdit(GetControl('ATA_AFFAIRE'));

  fPaDet := THPanel(GetControl('PCLIENT'));

  ecran.OnCloseQuery := EcranOnCloseQuery;

  // gestion de l'onglet règles
  fRB_MODEGENE1 := TRadioButton(GetControl('RB_MODEGENE1'));
  fRB_MODEGENE2 := TRadioButton(GetControl('RB_MODEGENE2'));
  fRB_MODEGENE1.OnClick := fRB_MODEGENE1OnClick;
  fRB_MODEGENE2.OnClick := fRB_MODEGENE2OnClick;

  fRB_QUOTIDIENNE := TRadioButton(GetControl('RB_QUOTIDIENNE'));
  fRB_NBINTERVENTION := TRadioButton(GetControl('RB_NBINTERVENTION'));
  fRB_HEBDOMADAIRE := TRadioButton(GetControl('RB_HEBDOMADAIRE'));
  fRB_ANNUELLE := TRadioButton(GetControl('RB_ANNUELLE'));
  fRB_MENSUELLE := TRadioButton(GetControl('RB_MENSUELLE'));

  fRB_QUOTIDIENNE.OnClick := fRB_QUOTIDIENNEOnClick;
  fRB_NBINTERVENTION.OnClick := fRB_NBINTERVENTIONOnClick;
  fRB_HEBDOMADAIRE.OnClick := fRB_HEBDOMADAIREOnClick;
  fRB_ANNUELLE.OnClick := fRB_ANNUELLEOnClick;
  fRB_MENSUELLE.OnClick := fRB_MENSUELLEOnClick;

  fPC_FREQUENCE := TPageControl(GetControl('PC_FREQUENCE'));
  fTS_QUOTIDIENNE := TTabSheet(GetControl('TS_QUOTIDIENNE'));
  fTS_NBINTERVENTION := TTabSheet(GetControl('TS_NBINTERVENTION'));
  fTS_HEBDOMADAIRE := TTabSheet(GetControl('TS_HEBDOMADAIRE'));
  fTS_ANNUELLE := TTabSheet(GetControl('TS_ANNUELLE'));
  fTS_MENSUELLE := TTabSheet(GetControl('TS_MENSUELLE'));

  fTS_QUOTIDIENNE.TabVisible := false;
  fTS_HEBDOMADAIRE.TabVisible := false;
  fTS_NBINTERVENTION.TabVisible := false;
  fTS_ANNUELLE.TabVisible := false;
  fTS_MENSUELLE.TabVisible := false;

  fRB_MOISMETHODE1 := TRadioButton(GetControl('RB_MOISMETHODE1'));
  fRB_MOISMETHODE2 := TRadioButton(GetControl('RB_MOISMETHODE2'));
  fRB_MOISMETHODE3 := TRadioButton(GetControl('RB_MOISMETHODE3'));
  fRB_MOISMETHODE1.OnClick := fRB_MOISMETHODE1OnClick;
  fRB_MOISMETHODE2.OnClick := fRB_MOISMETHODE2OnClick;
  fRB_MOISMETHODE3.OnClick := fRB_MOISMETHODE3OnClick;

  combo := THValComboBox(GetControl('ATA_UNITETEMPS'));
  if combo <> nil then
  begin
    Combo.OnEnter := UniteOnEnter;
    combo.OnChange := ATA_UNITETEMPSOnChange;
  end;

  combo := THValComboBox(GetControl('ATA_UNITETEMPS_'));
  if combo <> nil then
  begin
    Combo.OnEnter := UniteOnEnter2;
    combo.OnChange := ATA_UNITETEMPS_OnChange;
  end;

  Ed := THEdit(GetControl('ATA_QTEINITIALE'));
  if Ed <> nil then
    ed.OnExit := ATA_QTEINITIALEOnExit;
  Ed := THEdit(GetControl('ATA_QTEINITPLA'));
  if Ed <> nil then
    ed.OnExit := ATA_QTEINITPLAOnExit;
  Ed := THEdit(GetControl('ATA_PUPR'));
  if Ed <> nil then
    ed.OnExit := ATA_PUPROnexit;
  Ed := THEdit(GetControl('ATA_PUVENTEHT'));
  if Ed <> nil then
    ed.OnExit := ATA_PUVENTEHTOnExit;

  Ed := THedit(GetControl('ATA_TYPEARTICLE'));
  Ed.OnExit := TypeArticleOnexit;

  Ed := THedit(GetControl('ATA_FONCTION'));
  Ed.OnExit := FonctionOnexit;

  Ed := THedit(GetControl('ATA_BTETAT'));
  Ed.OnExit := TypeActionOnexit;

  Ed :=THEdit(GetControl('CLIENTCONTACT'));
  Ed.OnElipsisClick := AppelContact;

  SetControlProperty('ATA_TYPEARTICLE', 'Plus', PlusTypeArticle(true));
  SetControlVisible('TATA_FONCTION', GetParamsocSecur('SO_AFFONCTION', false));
  SetControlVisible('ATA_FONCTION', GetParamsocSecur('SO_AFFONCTION', false));
  SetControlVisible('BCALCULRAPRES', false);

  // si on est en mode de valorisation ressource
  // les prix unitaires de valorisation n'ont pas de sens
  //SetControlCaption ('PR','Prix de revient par '+ RechDom('AFVALORISATION',GetParamsocSecur('SO_AFVALOPR'),FALSE)  );
  //SetControlCaption ('PV','Prix de vente par '+ RechDom('AFVALORISATION',GetParamsocSecur('SO_AFVALOPV'),FALSE)  );
  //if Not(V_PGI.SAV) then
  // begin

  SetControlVisible('PV', False); //affichage du mode de valorisation pour débug en sav
  SetControlVisible('PR', False);
  // end;
  if GetParamsocSecur('SO_AFVALOPR','') = 'RES' then
  begin
    SetControlVisible('ATA_PUPR', false);
    SetControlVisible('TATA_PUPR', false);
  end
  else
    if GetParamsocSecur('SO_AFVALOPR','') <> 'TAC' then
    SetControlEnabled('ATA_PUPR', false);
  if GetParamsocSecur('SO_AFVALOPV','') = 'RES' then
  begin
    SetControlVisible('ATA_PUVENTEHT', false);
    SetControlVisible('TATA_PUVENTEHT', false);
  end
  else
    if GetParamsocSecur('SO_AFVALOPV','') <> 'TAC' then
    SetControlEnabled('ATA_PUVENTEHT', false);

  //if (GetParamsocSecur('SO_AFVALOPR','') = 'RES') and
  //  (GetParamsocSecur('SO_AFVALOPV','') = 'RES') then
    SetControlVisible('GB_PRIXUNITAIRE', false);
    SetControlVisible('RB_NBINTERVENTION', false);

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 25/04/2002
Modifié le ... :
Description .. : controle la saisie de la tcahe courante
Suite ........ :
Mots clefs ... :
*****************************************************************}
function TOF_BTTACHES.ControleEcran: boolean;
var
  vInRes: Integer;
  vRdQte: Double;

begin

  result := False;
  if ControleZonesOblig then
  begin
    vInRes := ControleRessource(vRdQte);
    if vInRes = -1 then
    begin
      PGIBoxAF(TexteMessage[4], ecran.Caption); //Il est interdit d''avoir 2 fois le même code ressource pour une tâche.'
      // vider la zone
      fGSRes.Cells[fGSRes.Col, fGSRes.Row] := '';
    end

    else if vInRes = -2 then
      PGIBoxAF(TexteMessage[9], ecran.Caption) //Saisir le code de la ressource !

    else if vInRes = -3 then
    begin
      result := True;
      //if PGIAskAF(TexteMessage[52], ecran.Caption) = MrYes then  //La somme du prévu des ressources doit être égale à la quantité prévue de la tâche. Voulez-vous rafraichir la quantité prévue de la tâche ?
      //  SetControlText('ATA_QTEINITIALE', FloatToStr(vRdQte));
    end
    else
      result := True;
  end;
end;

procedure TOF_BTTACHES.AlimTob(NoJour: integer);
var
  tobdet: tob;
begin
  Tobdet := Tob.Create('AFTACHEJOUR', fTobTacheJour, -1);
  TOBDet.PutValue('Atj_Affaire', fTobDet.GetValue('ATA_AFFAIRE'));
  TOBDet.PutValue('atj_NumeroTache', fTobDet.GetValue('ATA_NumeroTAche'));
  TOBDet.PutValue('atj_JourHeureDeb', 0);
  TOBDet.PutValue('atj_JourNbHeures', 0);
  TOBdet.PutValue('atj_JourAplanif', noJour);
  TOBdet.PutValue('atj_ModeleTache', '');
  TOBdet.PutValue('atj_TypeJour', 'AFF');
end;

procedure TOF_BTTACHES.AddUpdateRessList(aCol, aRow: Integer);
begin
  SetLength(fUpdateRes, length(fUpdateRes) + 1);
  fUpdateRes[length(fUpdateRes) - 1].StOldRes := fTobDet.Detail[aRow - 1].GetValue('ATR_RESSOURCE');
  fUpdateRes[length(fUpdateRes) - 1].StNewRes := fGSRes.Cells[ACol, ARow];
  fUpdateRes[length(fUpdateRes) - 1].rdAffaire := fCurAffaire;
  fUpdateRes[length(fUpdateRes) - 1].StNumTache := fTobDet.Detail[aRow - 1].GetValue('ATR_NUMEROTACHE');
  fUpdateRes[length(fUpdateRes) - 1].InStatut := cInUpdate;
end;

procedure TOF_BTTACHES.RefreshGrid;
var
  vInRow: Integer;
begin

  if fGSTaches.row > 0 then
     begin
     vInRow := fGSTaches.row;
     fTobTaches.PutGridDetailOnListe(fGSTaches, 'AFLISTETACHES');
     fGSTaches.Fixedcols := 1;
     TFVierge(Ecran).Hmtrad.ResizeGridColumns(fGSTaches);
     fGSTaches.Row := vInRow;
     // lors de la creation, le numero de tache n'est pas rafraichit !
     if assigned(fTobDet) and (GetControlText('ATA_NUMEROTACHE') = '0') then
        begin
        fTobDet := TOB(fGSTaches.Objects[0, fGSTaches.Row]);
        SetControlText('ATA_NUMEROTACHE', fTobDet.GetValue('ATA_NUMEROTACHE'));
        end;
  end;
end;

procedure TOF_BTTACHES.RefreshAdresse;
begin

  if fTobDet.GetString('QUELLEADRESSE') = 'CLIENT' then
    SetControlText('GB_ADRESSE', traduitGA(TexteMessage[36])) //Adresse d''intervention du tiers

  else if fTobDet.GetString('QUELLEADRESSE') = 'TACHE' then
    SetControlText('GB_ADRESSE', traduitGA(TexteMessage[3])); //Adresse d''intervention

  SetControlText('CLIENTLIBELLE', fTobDet.GetString('ADR_LIBELLE'));
  SetControlText('CLIENTPRENOM', fTobDet.GetString('ADR_LIBELLE2'));
  SetControlText('CLIENTADRESSE1', fTobDet.GetString('ADR_ADRESSE1'));
  SetControlText('CLIENTADRESSE2', fTobDet.GetString('ADR_ADRESSE2'));
  SetControlText('CLIENTADRESSE3', fTobDet.GetString('ADR_ADRESSE3'));
  SetControlText('CLIENTCODEPOSTAL', fTobDet.GetString('ADR_CODEPOSTAL'));
  SetControlText('CLIENTVILLE', fTobDet.GetString('ADR_VILLE'));
  SetControlText('CLIENTTELEPHONE', fTobDet.GetString('ADR_TELEPHONE'));
  SetControlText('CLIENTCONTACT', fTobDet.GetString('ATA_FAITPARQUI'));

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_BTTACHES.InsertTache;
var ARow: integer;
begin

  if ClientFerme then
    PGIBoxAF(TexteMessage[31], ecran.Caption); //Attention, ce client est fermé. Vous ne devriez pas créer de tâches !

  fTobDet := TOB.Create('TACHE', fTobTaches, -1);
  fTobDet.AddChampSupValeur('QTEAPLANIFIERCALC', 0, false);
  fTobDet.AddChampSupValeur('RAPPTPRCAL', 0, false);
  fTobDet.AddChampSupValeur('RAPPTPVCAL', 0, false);
  fStatut := taCreat;

  if assigned(fGSTaches) then
  begin
    Arow := fTobTaches.detail.count;
    fGSTaches.InsertRow(ARow);
    fGSTaches.Row := ARow;
    fGSTaches.RowCount := fTobTaches.detail.count + 1;
  end;

  RefreshGridRes;
  InitNewTache(fTobDet);

  if fTobModele <> nil then
    TOBCopieModeleTache(fCurAffaire, fTobModele, fTobDet); // AB-Modeletache-copie la Tob

  // C.B 26/10/2006
  // affectation de la première adresse d'intervention de l'affaire par défaut
  // sinon, on prend l'adresse du tiers
  LoadAdresseTache(True, fTobDet);

  RefreshAdresse;

  //Recherche du Contact et du type d'action evenement en affichage des lignes de la grille
  LectureTypeAction(GetControlText('ATA_BTETAT'));
  LectureContact( GetControlText('CLIENTCONTACT'));

  fTobDet.PutEcran(Ecran, fPaDet);
  InitNewRegles;
  if (fTobModele <> nil) and (fTobModele.NomTable = 'AFMODELETACHE') then
    PutRegles(fTobModele.GetString('AFM_MODELETACHE')); // AB-20040924 - Modeletache-Reprend les règles du modèle

  SetControlText('ATA_DATEDEBPERIOD_', dateToStr(fDtFinAffaire));

  fGSTaches.Objects[0, fGSTaches.Row] := fTobDet;
  //type tache
  RefreshLignePiece; // AB-LigneAffaire-Champs non modifiables
  //SetControlEnabled('ATA_TYPEARTICLE', true);
  SetControlText('ATA_TYPEARTICLE', 'PRE');
  SetControlEnabled('ATA_TYPEARTICLE', False);

  // C.B 25/01/2005
  if assigned(fTobModele) then
    FreeAndNil(fTobModele);

end;

procedure TOF_BTTACHES.InitNewRegles;
var
	fStPeriod : String;

begin
  // gestion des règles
  fRB_MODEGENE1.checked := true;

  fStPeriod := GetParamsocSecur('SO_BTPERIODDEFAUT', 'Q');
  // C.B 23/05/2006
	// initialisation sur paramsoc
  if fStPeriod = 'A' then
		fRB_ANNUELLE.checked := true

  else if fStPeriod = 'M' then
	  fRB_MENSUELLE.checked := true

  else if fStPeriod = 'S' then
		fRB_HEBDOMADAIRE.checked := true

  else if fStPeriod = 'Q' then
		fRB_QUOTIDIENNE.checked := true

  else if fStPeriod = 'NBI' then
		fRB_NBINTERVENTION.checked := true
     
  else if fStPeriod = '' then
	  fRB_MENSUELLE.checked := true;

  fRB_MOISMETHODE1.checked := true;
  THValComboBox(GetControl('ATA_MOISSEMAINE1')).Value := '1';
  THValComboBox(GetControl('ATA_MOISJOURLIB')).Value := 'J1';
  THValComboBox(GetControl('ATA_MOISSEMAINE2')).Value := '1';
  SetCOntrolText('JOURAN', '1');
  SetCOntrolText('MOISAN', '1');
  SetControlText('ATA_MOISJOURFIXE', '1');
  SetControlText('ATA_MOISFIXE0', '1');
  SetControlText('ATA_MOISFIXE1', '1');
  SetControlText('ATA_MOISFIXE2', '1');
  TCheckBox(GetControl('ATA_JOUR1H')).checked := false;
  TCheckBox(GetControl('ATA_JOUR2H')).checked := false;
  TCheckBox(GetControl('ATA_JOUR3H')).checked := false;
  TCheckBox(GetControl('ATA_JOUR4H')).checked := false;
  TCheckBox(GetControl('ATA_JOUR5H')).checked := false;
  TCheckBox(GetControl('ATA_JOUR6H')).checked := false;
  TCheckBox(GetControl('ATA_JOUR7H')).checked := false;
  TCheckBox(GetControl('ATA_JOUR1M')).checked := false;
  TCheckBox(GetControl('ATA_JOUR2M')).checked := false;
  TCheckBox(GetControl('ATA_JOUR3M')).checked := false;
  TCheckBox(GetControl('ATA_JOUR4M')).checked := false;
  TCheckBox(GetControl('ATA_JOUR5M')).checked := false;
  TCheckBox(GetControl('ATA_JOUR6M')).checked := false;
  TCheckBox(GetControl('ATA_JOUR7M')).checked := false;
end;

function TOF_BTTACHES.MajItems(pTob: Tob): Boolean;
var
  vSt: string;

begin

  {result := True;

  //if (GetParamSocSecur('SO_AFCLIENT',0) = cInCLientRwd) then // gm
  begin
    vSt := 'UPDATE AFPLANNING SET APL_vallibre1 = '+ variantToSql(pTob.Getdouble('ATA_vallibre1')) ;
    vSt := vSt + ',APL_vallibre2 = '+ variantToSql(pTob.Getdouble('ATA_vallibre2'));
    vSt := vSt + ',APL_vallibre3 = '+ variantToSql(pTob.Getdouble('ATA_vallibre3'));
    vSt := vSt + ' WHERE APL_AFFAIRE = "' + pTob.GetString('ATA_AFFAIRE') + '"';
    vSt := vSt + ' AND APL_NUMEROTACHE = ' + pTob.GetString('ATA_NUMEROTACHE');
    executeSql(vSt);
  end;

  // spécifique cegid
  if VH_GC.GCIfDefCEGID then
  begin

    // C.B 20/06/2005
    // si le RI est reçu, on stocke la date et qui a saisi le RI
    if (pTob.GetString('ATA_BOOLLIBRE1') = 'X') then
    begin
      SetControlText('ATA_DATELIBRE1', DateToStr(Date));
      SetControlText('APL_CHARLIBRE2', V_PGI.UserName);
    end;

    if (pTob.GetString('ATA_BOOLLIBRE1') = 'X') and
      (PGIAskAF(TexteMessage[38], ecran.Caption) = MrYes) then //Voulez-vous marquer rapport d''intervention reçu pour toutes les interventions de cette tâche ?
    begin
      vSt := 'UPDATE AFPLANNING SET APL_BOOLLIBRE1 = "X"';
      vSt := vSt + ' WHERE APL_AFFAIRE = "' + pTob.GetString('ATA_AFFAIRE') + '"';
      vSt := vSt + ' AND APL_NUMEROTACHE = ' + pTob.GetString('ATA_NUMEROTACHE');
      executeSql(vSt);
    end;
  end;

  if (pTob.GetString('ATA_TERMINE') = 'X') and fATATermineChange then
  begin
    if PGIAskAF(TexteMessage[37], ecran.Caption) = MrYes then //Voulez-vous marquer comme terminé toutes les interventions de cette tâche ?
    begin
      vSt := 'UPDATE AFPLANNING SET APL_ETATLIGNE = "TER"';
      vSt := vSt + ' WHERE APL_AFFAIRE = "' + pTob.GetString('ATA_AFFAIRE') + '"';
      vSt := vSt + ' AND APL_NUMEROTACHE = ' + pTob.GetString('ATA_NUMEROTACHE');
      // C.B 29/12/2005
      // on ne change pas l'etat des interventions facturées
      vSt := vSt + ' AND APL_ETATLIGNE <> "FAC"';
      executeSql(vSt);
    end;
  end;}
  
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : On passe l'element a inserer ou a supprimer
Mots clefs ... :
*****************************************************************}
procedure TOF_BTTACHES.ValiderLesTaches;
var vSt :string;
begin

  // suppression des taches une par une
  if assigned(fTobSup) and (fTobSup.Detail.count = 1) then
  begin
    if ValideLesTaches(nil, fTobSup.detail[0], nil, fCurAffaire.StAffaire, fUpdateRes) then
    begin
      fTobSup.ClearDetail;
      fTobDet := nil;
    end;
  end

  // enregistrement de toutes les taches et de leurs ressources
  else
  begin
    if (fStatut <> taConsult) then
    begin
      if fTobDet.modifie then
      begin
        //MajItems(fTobDet);
        //AB-200602- Affaire cloturé si toutes les taches sont terminées pour affaire école ou affaire des pièces planifiables
        if fATATermineChange and (GetParamsoc('SO_BTMODEPLANNING') = 'ART') then
        begin
          if fTobTaches.FindFirst(['ATA_TERMINE'], ['-'], true) = nil then
          begin
            vSt := 'UPDATE AFFAIRE SET AFF_ETATAFFAIRE="CLO" WHERE AFF_AFFAIRE="'+fCurAffaire.StAffaire+'" ';
            vSt := vSt + ' AND (AFF_STATUTAFFAIRE="AFG" OR AFF_ISECOLE="X") ';
            ExecuteSQL(vSt);
          end
          else
          begin  //AB-200606- remettre l'affaire en cours si une tache non terminée
            vSt := 'UPDATE AFFAIRE SET AFF_ETATAFFAIRE="ENC" WHERE AFF_AFFAIRE="'+fCurAffaire.StAffaire+'" ';
            vSt := vSt + ' AND (AFF_STATUTAFFAIRE="AFG" OR AFF_ISECOLE="X") ';
            ExecuteSQL(vSt);
          end;
        end;
      end;
      if ValideLesTaches(fTobDet, nil, fTobSupRes, fCurAffaire.StAffaire, fUpdateRes) then
      begin
        fTobDet.modifie := false;
        fATATermineChange := false;
      end;
    end;
  end;

  fUpdateRes := nil;
  fTobSupRes.ClearDetail;

end;

procedure TOF_BTTACHES.ValiderLesTachesJour;
var
  indice, ii: integer;
  QQ: TQuery;
  Tobdet: Tob;
begin
  ExecuteSql('DELETE FROM AFTACHEJOUR WHERE ATJ_TYPEJOUR="AFF" AND ATJ_AFFAIRE="'
    + fTobDet.GetValue('ATA_AFFAIRE') + '" AND ATJ_NUMEROTACHE="'
    + IntToStr(fTobDet.GetValue('ATA_NUMEROTACHE')) + '"');
  indice := 0;
  QQ := OpenSQL('SELECT MAX(ATJ_JOURNUMERO) FROM AFTACHEJOUR', true);
  if not QQ.Eof then
    Indice := QQ.Fields[0].AsInteger + 1;
  Ferme(QQ);
  for ii := 0 to fTobTacheJour.detail.count - 1 do
  begin
    Tobdet := fTobTacheJour.detail[ii];
    // obligatoire de faire ceci pour Ok en création de tache
    if Tobdet.getvalue('ATJ_NUMEROTACHE') = 0 then
      TobDet.putvalue('ATJ_NUMEROTACHE', fTobDet.GetValue('ATA_NumeroTAche'));
    tobdet.putvalue('ATJ_JOURNUMERO', Indice);
    Inc(indice);
  end;
  fTobTacheJour.InsertOrUpdateDb(False);
end;

procedure TOF_BTTACHES.mnPlanChargeQuantite_OnClick(SEnder: TObject);
var
  vInCurIndiceTob: Integer;
  retour: string;
begin
  if (GetControlText('ATA_NUMEROTACHE') = '0') then
    fGSTaches.DeleteRow(fGSTaches.Row)
  else
    if not Valider(false, false) then
    Exit;
  vInCurIndiceTob := fGSTaches.row - 1;
  RefreshGrid;
  //retour := AFLanceFicheAFPlanCharge('', 'AFFAIRE:' + GetControlText('ATA_AFFAIRE') + ';QUANTITE;MONOFICHE');
  fStatut := taConsult;
  loadAffaire(vInCurIndiceTob);
  TobResToScreen;
  fGSTaches.Row := vInCurIndiceTob + 1;
  TobToScreen(fGSTaches.Row - 1);
end;

procedure TOF_BTTACHES.mnPlanChargePrixRevient_OnClick(SEnder: TObject);
var
  vInCurIndiceTob: Integer;
begin

  if (GetControlText('ATA_NUMEROTACHE') = '0') then
    fGSTaches.DeleteRow(fGSTaches.Row)
  else if not Valider(false, false) then
    Exit;

  vInCurIndiceTob := fGSTaches.row - 1;
  RefreshGrid;
  //AFLanceFicheAFPlanCharge('', 'AFFAIRE:' + GetControlText('ATA_AFFAIRE') + ';PRIXREVIENT');
  fStatut := taConsult;
  loadAffaire(vInCurIndiceTob);
  TobResToScreen;
  fGSTaches.Row := vInCurIndiceTob + 1;
  TobToScreen(fGSTaches.Row - 1);
end;

procedure TOF_BTTACHES.mnPlanChargePrixVente_OnClick(SEnder: TObject);
var
  vInCurIndiceTob: Integer;
begin

  if (GetControlText('ATA_NUMEROTACHE') = '0') then
    fGSTaches.DeleteRow(fGSTaches.Row)
  else if not Valider(false, false) then
    Exit;

  vInCurIndiceTob := fGSTaches.row - 1;
  RefreshGrid;
  //AFLanceFicheAFPlanCharge('', 'AFFAIRE:' + GetControlText('ATA_AFFAIRE') + ';PRIXVENTE');
  fStatut := taConsult;
  loadAffaire(vInCurIndiceTob);
  TobResToScreen;
  fGSTaches.Row := vInCurIndiceTob + 1;
  TobToScreen(fGSTaches.Row - 1);
end;

procedure TOF_BTTACHES.MnPlanChargeDispo_OnClick(SEnder: TObject);
begin
  //AFLanceFiche_OccupRessMoisDispo('EXECUTION=TRUE;TYPE=PC');
end;

procedure TOF_BTTACHES.mnPlanningViewer_OnClick(Sender: TObject);
begin
  if assigned(ecran.activecontrol) and (ecran.activecontrol.name = 'GSRES') then
    setFocusControl('ATA_FONCTION');
  if Valider(false, false) then
  begin
    RefreshGrid;
    //AFLanceFicheAFPlanningViewer(GetControlText('ATA_AFFAIRE'), 'AFFAIRE:' + GetControlText('ATA_AFFAIRE') + ';MONOFICHE');
  end;
end;

procedure TOF_BTTACHES.mnPlanningRes_OnClick(Sender: TObject);
begin
  if assigned(ecran.activecontrol) and (ecran.activecontrol.name = 'GSRES') then
    setFocusControl('ATA_FONCTION');
  if Valider(false, false) then
  begin
    RefreshGrid;
    //AFLanceFicheAFPlanningRes(GetControlText('ATA_AFFAIRE'), 'AFFAIRE:' + GetControlText('ATA_AFFAIRE') + ';MONOFICHE');
  end;
end;

procedure TOF_BTTACHES.MnPlanningTache_OnClick(Sender: TObject);
var
  vStAffaire: string;
  vStDemiJournee: string;
  vStDateDeb: string;
  vStDateFin: string;
                          
begin
  if Valider(false, false) then
  begin
    RefreshGrid;
    vStAffaire := ' APL_AFFAIRE = "' + fED_AFFAIRE.text + '"';

    if GetParamsocSecur('SO_AFDEMIJOURNEE', false) then
      vStDemiJournee := 'X'
    else
      vStDemiJournee := '-';

    // C.B 25/08/2005
    // on ouvre les plannings sur la première  interventions trouvée.
    // On ouvre un planning de la durée maximum définie dans les paramsoc.
    // Si il n'y a pas d'intervention, on ouvre sur la date de début de l'affaire
    //if not RecherchePremiereEtDerniereDatePlanning(fED_AFFAIRE.text, '', vStDateDeb, vStDateFin) then
    //  vStDateDeb := DateToStr(fDtDebutAffaire);
    //vStDateFin := DateToStr(PlusDate(StrToDate(vStDateDeb), GetParamsocSecur('SO_AFPLANAFFICHM1',0) + GetParamsocSecur('SO_AFPLANAFFICHM2',0), 'M'));
                                                
    // C.B 25/08/2005
    // on lance le planning 6
    //ExecPlanning(inttostr(cInMenuPlanning + 6), DateToStr(fDtDebutAffaire), vStDateFin, vStDateDeb, vStAffaire, '', '-', vStDemiJournee, 'ZOOMTACHE;ACTION=MODIFICATION');

    // C.B 02/09/2005
    // en retour du planning, on refresh la quantité planifiée de la tâche courante
    // relire la quantité planifiée en base
    RefreshQuantitePla;

    // on ne peut pas changer l'article si du planning existe
    //SetControlEnabled('ATA_CODEARTICLE', not ExistePlanning(fCurAffaire.StAffaire, GetControlText('ATA_NUMEROTACHE')));
    //SetControlEnabled('ATA_TYPEARTICLE', GetControlEnabled('ATA_CODEARTICLE'));
    SetControlText('ATA_TYPEARTICLE', 'PRE');
    SetControlEnabled('ATA_TYPEARTICLE', False);
  end;
end;

procedure TOF_BTTACHES.RefreshQuantitePla;
var
  i: Integer;
  vRd: Double;

begin

  // C.B 27/09/2005
  //En retour du planning on met a jour le rap de toutes les taches
  for i := 0 to fTobTaches.detail.count - 1 do
  begin
    //vRd := UpdatePlanifieTache(fTobTaches.detail[i].GetString('ATA_AFFAIRE'), fTobTaches.detail[i].GetString('ATA_NUMEROTACHE'), fTobTaches.detail[i].GetString('ATA_UNITETEMPS'));
    fTobTaches.detail[i].PutValue('ATA_QTEPLANIFPLA', vRd);

    // recalcul le reste à planifier de la tache courante
    if fTobDet.GetString('ATA_NUMEROTACHE') = fTobTaches.detail[i].GetString('ATA_NUMEROTACHE') then
    begin
      SetControlText('ATA_QTEPLANIFPLA', floatToStr(vRd));
      CalculQuantitePla(False);
    end;
  end;
end;

procedure TOF_BTTACHES.MnPlanningRessource_OnClick(Sender: TObject);
var
  vStAffaire: string;
  vStDemiJournee: string;
  vSt: string;
  vStDateDeb: string;
  vStDateFin: string;
  vQr: TQuery;

begin
  if Valider(false, false) then
  begin
    RefreshGrid;
    vStAffaire := ' ATA_AFFAIRE = "' + fED_AFFAIRE.text + '"';

    if GetParamsocSecur('SO_AFDEMIJOURNEE', false) then
      vStDemiJournee := 'X'
    else
      vStDemiJournee := '-';

    // C.B 25/08/2005
    // on ouvre les plannings sur la première  interventions trouvée.
    // On ouvre un planning de la durée maximum définie dans les paramsoc.
    // Si il n'y a pas d'intervention, on ouvre sur la date de début de l'affaire
    vSt := 'SELECT APL_DATEDEBPLA FROM AFPLANNING WHERE APL_AFFAIRE = "' + fED_AFFAIRE.text + '"';
    vQr := nil;
    try
      vQR := OpenSQL(vSt, True);
      if (not vQR.EOF) then
        vStDateDeb := vQr.findField('APL_DATEDEBPLA').asString
      else
        vStDateDeb := DateToStr(fDtDebutAffaire);
    finally
      ferme(vqr);
    end;
    vStDateFin := DateToStr(PlusDate(StrToDate(vStDateDeb), GetParamsocSecur('SO_AFPLANAFFICHM1',0) + GetParamsocSecur('SO_AFPLANAFFICHM2',0), 'M'));

    //ExecPlanning(inttostr(cInMenuPlanning + 3), DateToStr(fDtDebutAffaire), vStDateFin, vStDateDeb, vStAffaire, '', '-', vStDemiJournee, 'ZOOMTACHE;ACTION=MODIFICATION');

    // C.B 02/09/2005
    // en retour du planning, on refresh la quantité planifiée de la tâche courante
    // relire la quantité planifiée en base
    //RefreshQuantitePla;

    // on ne peut pas changer l'article si du planning existe
    //SetControlEnabled('ATA_CODEARTICLE', not ExistePlanning(fCurAffaire.StAffaire, GetControlText('ATA_NUMEROTACHE')));
    //SetControlEnabled('ATA_TYPEARTICLE', GetControlEnabled('ATA_CODEARTICLE'));
    SetControlText('ATA_TYPEARTICLE', 'PRE');
    SetControlEnabled('ATA_TYPEARTICLE', False);
  end;
  // rechargement des ressources a faire en V6+
end;

//AB-200509- Occupation par ressource sans détail affaire
procedure TOF_BTTACHES.MnPlanningOccupRess_OnClick(SEnder: TObject);
begin
  // en attendant les recherches par compétences
  //AFLanceFiche_PlanningOccupRess('TYPEPLA=PLA;');
end;

//AB-200509- Occupation par ressource avec détail affaire
procedure TOF_BTTACHES.MnPlanningOccupRessDetaille_OnClick(SEnder: TObject);
begin
  // en attendant les recherches par compétences
  //AFLanceFiche_PlanningOccupRessDetaille('TYPEPLA=PLA;');
end;

procedure TOF_BTTACHES.MnDetailPlanning_OnClick(SEnder: TObject);
begin
  //AFLanceFiche_DetailPlanning('AFFAIRE:' + fCurAffaire.StAffaire + ';NOFILTRE');
end;

procedure TOF_BTTACHES.MnPlanningActivite_OnClick(SEnder: TObject);
begin
  //AFLanceFiche_PlanningCompar('AFFAIRE:' + fCurAffaire.StAffaire);
end;

procedure TOF_BTTACHES.RefreshGridRes;
begin
  fGSRes.VidePile(False);
  initGridRes;
end;

procedure TOF_BTTACHES.UpdateGridRes(pBoModifNbRes: Boolean; pInNum: Integer);
var
  NbRes: integer;
begin

  // si modif du nb de ressource
  if pBoModifNbRes then
  begin
    NbRes := strToint(GetControlText('ATA_NBRESSOURCE')) + pInNum;
    SetControlText('ATA_NBRESSOURCE', inttoStr(NbRes));
    fGSRes.RowCount := NbRes + 1;
  end

  // fixe le nombre de ressource
  else
  begin
    if pInNum = 0 then
    begin
      RefreshGridRes;
      SetControlText('ATA_NBRESSOURCE', intToStr(pInNum));
      fGSRes.RowCount := 2;
    end
    else
    begin
      SetControlText('ATA_NBRESSOURCE', intToStr(pInNum));
      fGSRes.RowCount := pInNum + 1;
    end;
  end;

  // positionnement dans la grille
  fGSRes.row := fGSRes.rowcount - 1;
//  fGSRes.Enabled := (GetControlText('ATA_NBRESSOURCE') <> '0');
  fGSRes.Enabled := (not TCheckBox(GetControl('ATA_TERMINE')).Checked);
  fGSRes.ElipsisButton := fGSRes.Enabled;
end;             

procedure TOF_BTTACHES.BRESSOURCEOnClick(Sender: TObject);
begin
  if fGSRes.Row <= fTobDet.Detail.Count then
    AglLanceFiche('AFF', 'RESSOURCE', '', fTobDet.Detail[fGSRes.Row - 1].GetValue('ATR_RESSOURCE'), '')
  else
    PGIBoxAF(TexteMessage[25], ecran.Caption); //Aucune ressource n''est sélectionnée.
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Regarde la cohérence des qté initiales de la tache,
suite ........ : avec les qte initiale par ressource
suite ........ : pBoGeneration  : True si uniquement pour recalul qte raf ou rap,
suite ........ :           False si génération planning
*****************************************************************}
function TOF_BTTACHES.TestCoherence(pBoGeneration: Boolean; pStTache: string): boolean;
var
  i: Integer;
  vRdTotal: Double;

begin

  // C.B 27/07/2005
  // si on est en génération du planning
  // on ne teste la cohérence des quantités que si on est dans les options qui le nécessite
  // c'est à dire génération en quantité limitée
  if pBoGeneration then
  begin
    if TCheckbox(GetControl('ATA_RAPPLANNING')).checked then
    begin
      vRdTotal := 0;
      if (fGSRes.RowCount > 2) or (fGSRes.CellValues[cInColRes, 1] <> '') then
        // calcul de la qte par ressource
        for i := 1 to fGSRes.RowCount - 1 do
        begin
          if fGSRes.CellValues[cInColQte, i] = '' then
            fGSRes.CellValues[cInColQte, i] := '0';

          //if fGSRes.CellValues[cInColQtePC, i] = '' then
          //   fGSRes.CellValues[cInColQtePC, i] := '0';

          vRdTotal := vRdTotal + strtofloat(fGSRes.CellValues[cInColQte, i]);
        end;

      if ((GetControlText('ATA_QTEINITPLA') = '') and (vRdTotal <> 0)) or
        (vRdTotal <> valeur(GetControlText('ATA_QTEINITPLA'))) then
      begin
        PGIBox(format(TexteMessage[24], [GetControlText('ATA_LIBELLETACHE1')]), traduitGA(ecran.Caption)); //Les quantités à planifier des ressources ne sont pas cohérentes avec la quantité à planifier de la tâche %s.#13#10Le planning ne peut donc pas être généré.
        result := false;
      end
 
        // C.B 28/08/2005
        // si la quantité est nulle, on ne peut pas gérérer une quantité limité de 0
      else
        if valeur(THEdit(GetControl('ATA_QTEINITPLA')).Text) = 0 then
      begin
        PGIBoxAF(TexteMessage[50], ecran.Caption); //L''option génération limitée à la quantité à planifier est cochée.#13#10La quantité à planifier ne peut donc être nulle.
        result := False;
      end
      else
        result := true;
    end
    else
      result := True;
  end
  else
  begin
    vRdTotal := 0;
    if (fGSRes.RowCount > 2) or (fGSRes.CellValues[cInColRes, 1] <> '') then
      // calcul de la qte par ressource
      for i := 1 to fGSRes.RowCount - 1 do
      begin
        if fGSRes.CellValues[cInColQte, i] = '' then
          fGSRes.CellValues[cInColQte, i] := '0';

        //if fGSRes.CellValues[cInColQtePC, i] = '' then
        //  fGSRes.CellValues[cInColQtePC, i] := '0';

        vRdTotal := vRdTotal + strtofloat(fGSRes.CellValues[cInColQte, i]);
      end;                                

    if ((GetControlText('ATA_QTEINITPLA') = '') and (vRdTotal <> 0)) or
      (vRdTotal <> valeur(GetControlText('ATA_QTEINITPLA'))) then
    begin
      if PGIAskAF(TexteMessage[23], ecran.Caption) = MrYes then //Attention, la valeur initiale de la tâche n''est pas cohérente avec les valeurs initiales des ressources !#13#10Voulez vous quand même calculer les quantités ?
        result := true
      else
        result := false;
    end
    else
      result := True;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Regarde la cohérence des qté initiales de toutes les
suite ........ : taches avec les qte initiale par ressource
suite ........ : PBoRap  : True si uniquement pour recalul qte raf ou rap,
suite ........ :           False si génération planning
*****************************************************************}
function TOF_BTTACHES.TestCoherenceQte(pBoGeneration: Boolean): boolean;
var
  i: Integer;
  vBoTest: Boolean;

  function QuantiteOk: Boolean;
  var
    i, i_res: Integer;
    vTobTache: Tob;
    vRdTotal: Double;
    vRdRes: Double;
    vTobRes: Tob;
  begin
    result := false;
    for i := 0 to fTobTaches.detail.count - 1 do
    begin
      vRdRes := 0;
      vTobTache := fTobTaches.Detail[i];
      vRdTotal := vTobTache.GetDouble('ATA_QTEINITPLA');
      if IsTachePrestation (vTobTache.GetString('ATA_TYPEARTICLE')) then  //AB-200605 - FQ 12918
      begin
        for i_res := 0 to vTobTache.Detail.Count - 1 do
        begin
          vTobRes := vTobTache.Detail[i_res];
          vRdRes := vRdRes + vTobRes.GetDouble('ATR_QTEINITPLA');
        end;
      end;
      result := vRdRes = vRdTotal;
      if not result then
        break;
    end;
  end;

begin

  // C.B 27/07/2005
  // si on est en génération du planning
  // on ne teste la cohérence des quantités que si on est dans les options qui le nécessite
  // c'est à dire génération en quantité limitée
  result := false;
  if pBoGeneration then
  begin
    // on fait le test qui si une des taches à l'option génération limitée à la quantité planifiée
    vBoTest := False;
    for i := 0 to fTobTaches.detail.count - 1 do
    begin
      if fTobTaches.Detail[i].GetString('ATA_RAPPLANNING') = 'X' then
      begin
        vBoTest := True;
        break;
      end;
    end;

    if vBoTest and (not quantiteOk) then
    begin
      PGIBoxAF(TexteMessage[49], ecran.Caption); //Pour une des tâches à générer, les quantités à planifier des ressources ne sont pas cohérentes avec la quantité à planifier de la tâche.#13#10Le planning ne peut donc pas être généré.
      result := false;
    end
    else
      result := true;
  end
  else
  begin
    if (not quantiteOk) then
    begin
      if PGIAskAF(TexteMessage[23], ecran.Caption) = MrYes then //Attention, la valeur initiale de la tâche n''est pas cohérente avec les valeurs initiales des ressources !#13#10Voulez vous quand même calculer les quantités ?
        result := true
      else
        result := false;
    end
  end;
end;

procedure TOF_BTTACHES.BCALCULRAPRESOnClick(Sender: TObject);
var
  i: integer;
  vRdPlanifie: Double;
  vRdQte: Double;
begin // calcul du reste à planifier en qté pa rressource pour planning
  if assigned(ecran.activecontrol) and (ecran.activecontrol.name = 'GSRES') then
    setFocusControl('ATA_FONCTION');
  if valider(false, false) then
  begin
    RefreshGrid;
    if TestCoherence(False, fTobDet.GetValue('ATA_NUMEROTACHE')) then
    begin
      for i := 0 to fTobDet.Detail.Count - 1 do
      begin
        // C.B 04/11/2003
        vRdPlanifie := CalculPlanifie(fCurAffaire.StAffaire,
          fTobDet.Detail[i].GetValue('ATR_NUMEROTACHE'),
          'APL_QTEPLANIFIEE', fTobDet.Detail[i].GetValue('ATR_RESSOURCE'), 'PLA', False);
        vRdQte := valeur(fTobDet.Detail[i].GetValue('ATR_QTEINITPLA')) - vRdPlanifie;
        if vRdQte < 0 then
          vRdQte := 0;
        fGSRes.Cells[cInColQteRAP, i + 1] := floatToStr(vRdQte);
      end;
    end;
  end;
end;

// calcul du reste à planifier en qté pa rressource pour PC
procedure TOF_BTTACHES.CalculRafRes;
var
  vTobTache, vTobRes: Tob;
  i, i_res: integer;
  vRdPlanifie: Double;
  vRdQte: Double;
  vRdRealise: Double;

begin
  vRdQte := 0;
  vTobRes := nil;
  if not GetParamsocSecur('SO_BTGESTIONRAF', false) then
    Exit;
  for i := 0 to fTobTaches.Detail.Count - 1 do
  begin
    vTobTache := fTobTaches.Detail[i];
    if IsTachePrestation(vTobTache.GetString('ATA_TYPEARTICLE')) then //AB-200605 - FQ 12918
    begin
      for i_res := 0 to vTobTache.Detail.Count - 1 do
      begin
        vTobRes := vTobTache.Detail[i_res];
        vRdPlanifie := CalculPlanifie(fCurAffaire.StAffaire,
          vTobRes.GetValue('ATR_NUMEROTACHE'),
          'APL_QTEPLANIFIEE', vTobRes.GetValue('ATR_RESSOURCE'), 'PC', False);
        vRdQte := vTobRes.GetDouble('ATR_QTEINITIALE');
        vRdQte := vRdQte + vTobRes.GetDouble('ATR_ECARTQTEINIT');
        vRdQte := vRdQte - vRdPlanifie;
        vRdRealise := CalculRealise(fCurAffaire.StAffaire, vTobTache.GetString('ATA_FONCTION'), 'ACT_QTE', vTobRes.GetString('ATR_RESSOURCE'),
          vTobTache.GetString('ATA_ARTICLE'), vTobTache.GetString('ATA_UNITETEMPS'), True);
        vRdQte := vRdQte - vRdRealise;
        if (GetParamsocSecur('SO_AFCLIENT',0) = cInClientAlgoe) and (vRdQte < 0) then
          vRdQte := 0;
        vTobRes.AddChampSupValeur('RAFQTERES', floatToStr(vRdQte));
      end;
    end
    else
      for i_res := 0 to vTobTache.Detail.Count - 1 do
        if assigned(vTobRes) then
          vTobRes.AddChampSupValeur('RAFQTERES', floatToStr(vRdQte));
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... :
Modifié le ... : 25/10/2005
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOF_BTTACHES.PAGETACHEChange(Sender: TObject);
begin
  if (TPageControl(GetControl('PAGETACHE')).ActivePage = TTAbSheet(GetControl('TAB_RESSOURCE'))) then
    if (not fBoAFPLANDECHARGE) and
       (GetControlText('ATA_RAPPLANNING') = 'X') and
       (fAction <> taConsult) then
      SetControlVisible('BCALCULRAPRES', True)
    else
      SetControlVisible('BCALCULRAPRES', False);
end;

procedure TOF_BTTACHES.mnGeneAffaire_OnClick;
var i : integer;
begin

  //Génération de l'appel et/ou de l'évènement
  if Valider(false, false) then
     Begin
     //chargement de la tob avec les enregs de la grille
     For i := 0 to fTobTaches.Detail.Count -1 do
         begin
         FTobDet := fTobTaches.Detail[i];
         GenereAppel;
         end;
     end;

end;

function TOF_BTTACHES.RessourceSaisie: Boolean;
begin
  result := GetControlText('ATA_NBRESSOURCE') <> '0';
end;

function TOF_BTTACHES.TacheHeure(pBoGlobal : Boolean) : Boolean;
var
  vSt : String;

begin
  vSt := 'SELECT ATA_AFFAIRE FROM TACHE WHERE ATA_UNITETEMPS = "H" AND ATA_AFFAIRE = "' + fED_AFFAIRE.text + '"';
  if not pBoGlobal then
    vSt := vSt + ' AND ATA_NUMEROTACHE = ' + floatToStr(fTobDet.GetValue('ATA_NUMEROTACHE'));
  result := ExisteSql(vSt);
end;

procedure TOF_BTTACHES.mnGeneTache_OnClick;
begin

  //Génération de l'appel et/ou de l'évènement
  if Valider(false, false) then
     Begin
     GenereAppel
     end;

end;

Procedure Tof_BTTACHES.MnContact_OnClick;
Var Argument   : String;
		NumContact : String;
Begin

  Argument := '';
  NumContact := '';

  Argument := 'TYPE=T;';
  Argument := Argument + 'ACTION=CONSULTATION;';
  Argument := Argument + 'TYPE2=CLI;';
  Argument := Argument + 'PART=-;';
  Argument := Argument + 'TITRE='+ GetControlText('TLIBCLIAFF') +';';
  Argument := Argument + 'TIERS='+fStTiers+';';
  Argument := Argument + 'ALLCONTACT';

  AGLLanceFiche('YY','YYCONTACT','T;'+fStAuxilliaire,'', Argument);

end;
Procedure Tof_BTTACHES.MnAdressesInt_OnClick;
Var Range    : String;
		Argument : String;
    NumAdr   : string;
    Q        : TQuery;
Begin

   Range := 'ADR_TYPEADRESSE=TIE;ADR_NATUREAUXI=CLI;ADR_REFCODE=' + fStTiers;

   Argument := 'YTC_TIERSLIVRE=' + fStTiers;
   Argument := Argument + 'ACTION=CONSULTATION;';
   Argument := Argument + ';TYPEADRESSE=TIE';
   Argument := Argument + ';TITRE=' + GetControlText('TLIBCLIAFF');
   Argument := Argument + ';PART=-';
   Argument := Argument + ';CLI=' + fStAuxilliaire;
   Argument := Argument + ';TIERS=' + fStTiers;
   Argument := Argument + ';NATUREAUXI=CLI';

   AglLanceFiche('GC', 'GCADRESSES', Range, '', Argument);

end;


//Génération de l'appel en fonction du contrat et de la tache
Function TOF_BTTACHES.GenereAppel : boolean;
var AnneeDeb  : String;
    AnneeFin  : String;
    Affaire   : String;
    StTiers   : String;
    StLibelle : String;
    Frequence : String;
    NoTache   : String;
    TypeAction: String;
    vSt       : String;
    ListInt   : TListDate;
    NbJrDecal : Integer;
    NbAppel   : Integer;
    i         : Integer;
    Duree     : Double;
begin

  Result := false;
  if fTOBDet.GetValue('ATA_FAITPARQUI')= '' then
  begin
     PGIBoxAF('Vous devez renseigner un contact', ecran.caption);
     exit;
  end;

  if fTobDet.GetValue('ATA_TERMINE') = 'X' then
     begin
     PGIBoxAF('Cette tâche a déjà été générée', ecran.caption);
     exit;
     end;

  TypeAction := fTOBDet.GetValue('ATA_BTETAT');
  if TypeAction = '' then
     Begin
     PGIBoxAF('Le type d''action est obligatoire pour la génération d''une tâche', ecran.caption);
     exit;
     end;

  //Contrôle de la date de debut et de la date de fin
  AnneeDeb := FormatDateTime('yyyy',StrToDate(fTOBDet.GetValue('ATA_DATEDEBPERIOD')));
  AnneeFin := FormatDateTime('yyyy',StrToDate(fTOBDet.GetValue('ATA_DATEFINPERIOD')));

  if StrToDate(fTOBDet.GetValue('ATA_DATEDEBPERIOD')) = idate1900 then
     Begin
     PGIBoxAF(TexteMessage[63], ecran.Caption);
     SetControlText('ATA_DATEDEBPERIOD', DateToStr(now));
    AnneeDeb := FormatDateTime('yyyy',StrToDate(fTOBDet.GetValue('ATA_DATEDEBPERIOD')));
     end;

  if StrToDate(fTOBDet.GetValue('ATA_DATEFINPERIOD')) = idate2099 then
     Begin
     PGIBoxAF(TexteMessage[62], ecran.Caption);
     SetControlText('ATA_DATEFINPERIOD',dateToStr(PlusDate(now,5,'A')));
    ftobdet.putvalue('ATA_DATEFINPERIOD',fTOBDet.GetValue('ATA_DATEFINPERIOD'));
    AnneeFin := FormatDateTime('yyyy', StrToDate(fTOBDet.GetValue('ATA_DATEFINPERIOD')));
     end;

  if StrToDate(fTOBDet.GetValue('ATA_DATEFINPERIOD')) > PlusDate(now,5,'A') then
     Begin
     PGIBoxAF(TexteMessage[62], ecran.Caption);
     SetControlText('ATA_DATEFINPERIOD',dateToStr(PlusDate(now,5,'A')));
    ftobdet.putvalue('ATA_DATEFINPERIOD',fTOBDet.GetValue('ATA_DATEFINPERIOD'));
    AnneeFin := FormatDateTime('yyyy', StrToDate(fTOBDet.GetValue('ATA_DATEFINPERIOD')));
     end;

  //Définition de l'outil de calcul des règles d'échéance
  vAFReglesTache := TAFRegles.create;

  //génération en fonction des règles...
  NoTache   := FtobDet.GetValue('ATA_NUMEROTACHE');
  Affaire   := FtobDet.GetValue('ATA_AFFAIRE');
  StTiers   := FtobDet.GetValue('ATA_TIERS');
  Frequence := fTobDet.GetValue('ATA_PERIODICITE');
  NbJrDecal := fTobDet.GetValue('ATA_NBJOURSDECAL');
  Duree     := fTobDet.GetValue('ATA_QTEINTERVENT');
  StLibelle := fTobDet.GetValue('ATA_LIBELLETACHE1');

  ftobdet.UpdateDB;

  vAFReglesTache.LoadDBReglesTaches(Affaire, NoTache, 1);

  ListInt := vAFReglesTache.GenereListeJours(Duree, NbJrDecal, frequence, -1);
  if ListInt = nil then exit;

  ControleLastDateGene(ListInt, FtobDet.GetValue('ATA_LASTDATEGENE'), true);

  nbAppel := 0;

  if ListInt.count = 0 then
     begin
     //5 'Affaire : %s. Tiers : %s. Aucune intervention n''a été générée pour la tâche %s. Si vous avez généré en utilisant "Ajouter au planning existant", vous planifiez à partir de la dernière date de génération précisée dans l''onglet "Règles de la fiche Tâche".'
     vSt := format(TraduitGa(TexteMessage[57]), [affaire, StTiers, StLibelle]);
     vSt := vST + Chr(10) + format(TraduitGa(TexteMessage[59]), [dateToStr(FtobDet.GetValue('ATA_DATEANNUELLE'))]);
     if vAFReglesTache.DtDateFin < now then
        vSt := vSt + Chr(10) + format(TraduitGa(TexteMessage[58]), [dateToStr(vAFReglesTache.DtDateFin)]);
     PGIBoxAF(vst, ecran.caption);
     end
  else
     Begin
//     for i := ListInt.count - 1 downto 0 do
     for i := 0 to ListInt.count - 1 do
         begin
         //controle si la date de génération est supérieure à la date du jour...
         if ListInt.Items[i] > now then
            Begin
            GenererListesAppel(ListInt, i, Duree);
            NbAppel := NbAppel+1;
            end;
         end;
     fATATermineChange := false;
     SetControlProperty('ATA_TERMINE', 'Checked', True);
     SetControlText('ATA_LASTDATEGENE', DateToStr(now));
     fATATermineChange := True;
     ScreenToTob;
     valider(false, false);
     vst :=  'La génération de la tache ' + NoTache + ' s''est déroulée normalement. ' + Chr(10);
     vst := vst + 'Nombre d''appel(s) généré(s) : ' + inttostr(NbAppel);
     PGIBoxAF(vst, ecran.caption);
     End;

  vAFReglesTache.Free;

  ListInt.Free;

  RefreshLastDateGene;

  RefreshGrid;

  result := true;

end;


Procedure TOF_BTTACHES.GenererListesAppel(pListeJours : TListDate; NumLigne : Integer; Duree: Double);
Var CodeAppel   : String;
    P0          : String;
    P1          : String;
    P2          : String;
    P3          : String;
    Avn         : String;
    //
    C0          : THCritMaskEdit;
    C1          : THCritMaskEdit;
    C2          : THCritMaskEdit;
    C3          : THCritMaskEdit;
    CAvn        : THCritMaskEdit;
    //
    Req         : String;
    Ressource   : String;
    Contact     : String;
    BlocNote 	  : THRichEditOLE;
    NoAdresse   : Integer;
    DureeEnJour : Double;
    DateDeFin,DateDeb   : TDateTime;
    TobAppel    : Tob;
    TobAdrInt   : Tob;
    TobAdrIntB   : Tob;
    TobRes      : Tob;
    QQ          : TQuery;
    Ok_Saisie   : Boolean;
    iPartErreur,IDuree : integer;
begin

  //Vérification de la date de fin et de la date de début...
  Ok_Saisie := False;

  // Si en création calcul du N° d'appel
  Ok_Saisie := GetParamsocSecur('SO_SAISIECODEAPPEL', False);

  if Ok_Saisie then
     Begin
     C0 := ThCritMaskEdit.Create(TFvierge(ecran));
     C1 := ThCritMaskEdit.Create(TFvierge(ecran));
     C2 := ThCritMaskEdit.Create(TFvierge(ecran));
     C3 := ThCritMaskEdit.Create(TFvierge(ecran));
     Cavn := ThCritMaskEdit.Create(TFvierge(ecran));
     C0.parent := TFvierge(ecran);
     C1.parent := TFvierge(ecran);
     C2.parent := TFvierge(ecran);
     C3.parent := TFvierge(ecran);
     Cavn.parent := TFvierge(ecran);
     C0.Visible := false;
     C1.Visible := false;
     C2.Visible := false;
     C3.Visible := false;
     CAvn.Visible := false;
     C0.DataType := 'AFF_AFFAIRE0';
     C1.DataType := 'AFF_AFFAIRE1';
     C2.DataType := 'AFF_AFFAIRE2';
     C3.DataType := 'AFF_AFFAIRE3';
     CAvn.DataType:= 'AFF_AVENANT';
     end;

  //Détermination automatique du code appel dans le cas ou la saisie est requise
  //on prend le code contrat et on le transforme en code appel --->>>>>> naaaaaaannnnnnnn ... ce sont des oufzors
  if Ok_Saisie then
  Begin
     CodeAppel := GetControlText('ATA_AFFAIRE');
     CodeAffaireDecoupe(CodeAppel, P0, P1, P2, P3, Avn, TaCreat, False);
     C0.Text := 'W';
     C1.text := P1;
     C2.text := P2;
     C3.text := P3;
     Cavn.text := Avn;
     // CodeAppel := DechargeCleAffaire(C0,C1, C2, C3, CAvn, GetControlText('ATA_TIERS'), tacreat, True, True, false, iPartErreur);
     // correction --> omg
     CodeAppel := CalculCodeAppel('W', P1, P2, P3, Avn, GetControlText('ATA_TIERS'),true);
     // --
     P0 := C0.Text;
     P1 := C1.Text;
     P2 := C2.Text;
     P3 := C3.Text;
     Avn := CAvn.Text;
  end else
  begin
     CodeAppel := CalculCodeAppel('W', P1, P2, P3, Avn, GetControlText('ATA_TIERS'));
  end;

  if CodeAppel = '' then
     Begin
     PGIBoxAF(TexteMessage[55], ecran.caption); //Le code appel n'a pu être généré veuillez vérifier votre paramétrage
     exit;
     End;

  //Création de l'appel
  TobAppel := Tob.Create('AFFAIRE',Nil, -1);
  TobAdrInt := Tob.create('ADRESSES',nil, -1);
  TobAdrIntB := Tob.create('BADRESSES',nil, -1);

  //découpage du code Appel
  CodeAppelDecoupe(CodeAppel, P0, P1, P2, P3, Avn);

  //Mise à jour de la table Affaire avec les zones de l'écran
  TobAppel.PutValue('AFF_AFFAIRE', CodeAppel);
  TobAppel.PutValue('AFF_AFFAIREREF', CodeAppel);

  TobAppel.PutValue('AFF_AFFAIRE0', P0);
  TobAppel.PutValue('AFF_AFFAIRE1', P1);
  TobAppel.PutValue('AFF_AFFAIRE2', P2);
  TobAppel.PutValue('AFF_AFFAIRE3', P3);
  TobAppel.PutValue('AFF_AVENANT', Avn);
  TobAppel.PutValue('AFF_STATUTAFFAIRE', 'APP');
  TobAppel.PutValue('AFF_GENERAUTO', 'DIR');

  TobAppel.PutValue('AFF_DEVISE', fTOBDet.GetValue('ATA_DEVISE'));

  TobAppel.PutValue('AFF_TIERS', fTOBDet.GetValue('ATA_TIERS'));
  DateDeb := pListeJours.Items[NumLigne];
  TobAppel.PutValue('AFF_DATEDEBUT', DateDeb);

  // date fin = date deb + qte
  // regle utilisée : si on ne tombe pas sur un nombre de jour entier
  // on arrondi au nombre de jour supérieur
  (*
  DureeEnJour := AFConversionUnite(GetControlText('ATA_UNITETEMPS'), 'J', Duree);
  if int(DureeEnJour) < Duree then DureeEnJour := int(DureeEnJour) + 1;
  DateDeFin := plusDate(pListeJours.Items[NumLigne], round(DureeEnJour - 1), 'J');
  *)
  if Duree < 1 then Duree := 1;
  IDuree := StrToInt(FloatToStr(Duree));
  if fTOBDet.GetValue('ATA_UNITETEMPS') = 'H' Then
  begin
    DateDeFin := IncHour(DateDeb,IDuree);
  end else if fTOBDet.GetValue('ATA_UNITETEMPS') = 'J' Then
  begin
    DateDeFin := IncDay(DateDeb,IDuree);
  end else if fTOBDet.GetValue('ATA_UNITETEMPS') = 'MIN' Then
  begin
    DateDeFin := IncMinute(DateDeb,IDuree);
  end;

  TobAppel.PutValue('AFF_DATEFIN', DateDeFin);

  TobAppel.PutValue('AFF_DATELIMITE', iDate2099);
  TobAppel.PutValue('AFF_DATEGARANTIE', iDate2099);
  TobAppel.PutValue('AFF_DATEREPONSE', iDate2099);
  TobAppel.PutValue('AFF_ETATAFFAIRE', 'ECO');
  TobAppel.PutValue('AFF_AFFAIREINIT', fTOBDet.GetValue('ATA_AFFAIRE'));
  TobAppel.PutValue('AFF_CHANTIER', fCurAffaire.StChantier);
  //TobAppel.PutValue('AFF_CHANTIER', fCurAffaire.StAffaire);
  TobAppel.PutValue('AFF_NUMEROTACHE', fTOBDet.GetValue('ATA_NUMEROTACHE'));

  //Contrôle du contact
  Contact := fTOBDet.GetValue('ATA_FAITPARQUI');
  if Contact <> '' then
  begin
     TobAppel.PutValue('AFF_NUMEROCONTACT', Contact);
     TobAdrInt.PutValue('ADR_CONTACT',  fTOBDet.GetValue('ATA_FAITPARQUI')); //GetControlText('LCONTACT')
  end;

  TobAppel.PutValue('AFF_CREERPAR', 'TAC');

  //Descriptif de l'appel obligatoire
  if isBlobVide(Application.mainForm,fTOBDet,'ATA_DESCRIPTIF' ) then
  Begin
     TobAppel.PutValue('AFF_DESCRIPTIF', fTOBDet.GetValue('ATA_LIBELLETACHE1') + ' ' + fTOBDet.GetValue('ATA_LIBELLETACHE2'));
  end else
  Begin
     TobAppel.PutValue('AFF_DESCRIPTIF', fTOBDet.GetValue('ATA_DESCRIPTIF'));
  end;

  DefiniLibelle(TobAppel);
  
  if TobAppel.GetValue('AFF_LIBELLE') = '' then
  Begin
    TobAppel.PutValue('AFF_LIBELLE',    fTOBDet.GetValue('ATA_LIBELLETACHE1'));
    TobAppel.PutValue('AFF_DESCRIPTIF', fTOBDet.GetValue('ATA_LIBELLETACHE1') + ' ' + fTOBDet.GetValue('ATA_LIBELLETACHE2'));
  end;

  TobAppel.PutValue('AFF_PRIOCONTRAT', GetParamSocSecur('SO_BTIMPTACHE','1'));

  Ressource := RechRessPrinc(fTOBDet.GetValue('ATA_AFFAIRE'), fTOBDet.GetValue('ATA_NUMEROTACHE'));

  //recherche du responsable dans la tob des ressources associée à la grille
  if Ressource <> '' then
     Begin
     TobAppel.PutValue('AFF_RESPONSABLE', Ressource);
     TobAppel.PutValue('AFF_ETATAFFAIRE', 'AFF');
     end;

  //gestion de l'adresse
  //Vérification si adresse existe déjà pour l'appel
  NoAdresse := 0;
  Req := 'SELECT ADR_NUMEROADRESSE FROM ADRESSES LEFT JOIN BADRESSES ON BA0_TYPEADRESSE=ADR_TYPEADRESSE AND BA0_NUMEROADRESSE=ADR_NUMEROADRESSE WHERE ADR_REFCODE="' + CodeAppel + '" AND BA0_TYPEADRESSE="INT"';
  QQ := OpenSQL(Req, true);
  if QQ.EOF then
     Begin
     NoAdresse := GetSetNumAdresse;
     if NoAdresse <=0 then NoAdresse := 1;
     TobAdrInt.PutValue('ADR_NUMEROADRESSE', NoAdresse);
     end
  else
     Begin
     TobAdrInt.SelectDB('',QQ);
     TobAdrInt.PutValue('ADR_NUMEROADRESSE',  QQ.Fields[0].AsInteger);
     end;

  Ferme(QQ);

  TobAdrInt.PutValue('ADR_NATUREAUXI', '');
  TobAdrInt.PutValue('ADR_REFCODE', CodeAppel);
  TobAdrInt.PutValue('ADR_TYPEADRESSE', 'INT');
  //
  TobAdrInt.PutValue('ADR_JURIDIQUE',  '');
  TobAdrInt.PutValue('ADR_LIBELLE',    fTOBDet.GetValue('ADR_LIBELLE'));
  TobAdrInt.PutValue('ADR_LIBELLE2',   fTOBDet.GetValue('ADR_LIBELLE2'));
  TobAdrInt.PutValue('ADR_ADRESSE1',   fTOBDet.GetValue('ADR_ADRESSE1'));
  TobAdrInt.PutValue('ADR_ADRESSE2',   fTOBDet.GetValue('ADR_ADRESSE2'));
  TobAdrInt.PutValue('ADR_ADRESSE3',   fTOBDet.GetValue('ADR_ADRESSE3'));
  TobAdrInt.PutValue('ADR_CODEPOSTAL', fTOBDet.GetValue('ADR_CODEPOSTAL'));
  TobAdrInt.PutValue('ADR_VILLE',      fTOBDet.GetValue('ADR_VILLE'));
  TobAdrInt.PutValue('ADR_PAYS', '');
  TobAdrInt.PutValue('ADR_TELEPHONE',  fTOBDet.GetValue('ADR_TELEPHONE'));

  TobAdrInt.PutValue('ADR_EMAIL', '');
  TobAdrInt.PutValue('ADR_REGION', '');

  TobAdrInt.PutValue('ADR_LIVR', 'X');
  TobAdrInt.PutValue('ADR_FACT', 'X');
  TobAdrInt.PutValue('ADR_REGL', 'X');

  TobAdrIntB.PutValue('BA0_NUMEROADRESSE',  TobAdrInt.GetValue('ADR_NUMEROADRESSE'));
  TobAdrIntB.PutValue('BA0_TYPEADRESSE', 'INT');
  TobAdrIntB.PutValue('BA0_INT',  'X');

  if Ressource <> '' then
  begin
  	GenereEvenement(TobAppel, TobAdrInt.GetValue('ADR_NUMEROADRESSE'));
  end else
  begin
    TobAppel.PutValue('AFF_DATESIGNE', TobAppel.GetValue('AFF_DATEDEBUT'));
  end;    

  TobAppel.InsertOrUpdateDB(true);
  TobAdrInt.InsertorUpdateDB(true);
  TobAdrIntB.InsertorUpdateDB(true);

  TobAdrInt.free;
  TobAdrIntB.free;
  TobAppel.free;

end;

procedure Tof_BTTACHES.GenereEvenement(TobA : Tob; NoAdresse : String);
Var TobTachesRes  : Tob;
    TobEvenement  : Tob;
    TobTypeAction : Tob;
    NumEvent	    : Integer;
    NbEnreg		    : Integer;
    I             : Integer;
    DeltaDate	    : TDateTime;
    DateDeb		    : TDateTime;
    DateFin		    : TDateTime;
    Duree			    : Double;
    vQr           : TQuery;
    vStSQL        : String;
    NoTache       : String;
    TypeAction    : String;
    HeureDeb	    : TdateTime;
    HeureFin	    : TdateTime;
begin

    //chargement de la Tob Type d'action
    TypeAction := GetControlText('ATA_BTETAT');
    TobTypeAction := Tob.Create('BTETAT', nil, -1);
    vStSQL := 'SELECT BTA_BTETAT, BTA_MODIFIABLE, BTA_OBLIGATOIRE, BTA_ASSOSRES, BTA_DUREEMINI ';
    vStSQL := vStSQL + ' FROM BTETAT WHERE BTA_BTETAT = "' + TypeAction + '"';
    try
      vQR := OpenSQL(vStSQL, True);
      if not (vQR.EOF) then
         Begin
         TobTypeAction.LoadDetailDB('BTETAT', '', '', vQR, True);
         end;
    finally
      Ferme(vQR);
    end;

  // date fin = date deb + qte
  // par defaut l'heure de fion est équivalente à l'heure dans les PARAMSOC
    DateDeb   := TobA.GetValue('AFF_DATEDEBUT');
    DateFin   := TobA.GetValue('AFF_DATEFIN');

    Duree     := StrToFloat(GetControlText('ATA_QTEINTERVENT'));
    //Duree     := AFConversionUnite(GetControlText('ATA_UNITETEMPS'), 'H', Duree);
    if GetControlText('ATA_UNITETEMPS') = 'J' Then
    begin
    	Duree := Duree * DureeJour;
    end else
    begin
    	Duree := Duree * 60;
    end;
    HeureDeb 	:= StrToTime(TimeToStr(GetParamSocSecur('SO_BTAMDEBUT', '08:00')));
    DateDeb := StrToDate(DateToStr(DateDeb)) + heureDeb;

    DateFin := AjouteDuree (DateDeb,round(Duree));
    Duree := CalculDureeEvenement (DateDeb,DateFin);
    HeureFin := StrToTime(TimeToStr(DateFin));

    //lecture des ressources associées à la tâche sélectionnée
    NoTache := fTobDet.GetValue('ATA_NUMEROTACHE');
    TobTachesRes := Tob.Create('TACHERESSOURCE', nil, -1);
    vStSQL := 'SELECT ATR_RESSOURCE FROM TACHERESSOURCE WHERE ATR_AFFAIRE="' + GetControlText('ATA_AFFAIRE') + '"';
    vStSQL := vStSQL + ' AND ATR_NUMEROTACHE=' + NoTache;
    vQr := nil;
    try
      vQR := OpenSQL(vStSQL, True);
      if not (vQR.EOF) then
         Begin
         TobTachesRes.LoadDetailDB('TACHERESSOURCE', '', '', vQR, True);
         end;
    finally
      Ferme(vQR);
    end;

    //création de l'évènement associé à la ressource à la date de génération de la tache
    NbEnreg := TobTachesRes.detail.count -1;

    for I := 0 to NbEnreg do
        Begin
        //Chargement de la Tob des Evènements avec la tob de l'Appel
        if GetNumCompteur('BEP',iDate1900, NumEvent) then
           Begin
           TobEvenement := Tob.create('BTEVENPLAN',nil, -1);
           TobEvenement.PutValue('BEP_CODEEVENT', IntToStr(NumEvent));
           TobEvenement.PutValue('BEP_RESSOURCE', TobTachesRes.detail[I].GetValue('ATR_RESSOURCE'));
           TobEvenement.PutValue('BEP_BTETAT', GetControlText('ATA_BTETAT'));
           TobEvenement.PutValue('BEP_AFFAIRE', TobA.Getvalue('AFF_AFFAIRE'));
           TobEvenement.PutValue('BEP_TIERS', GetControlText('ATA_TIERS'));
           TobEvenement.PutValue('BEP_NUMEROADRESSE', NoAdresse);
           TobEvenement.PutValue('BEP_EQUIPERESS', '');
           TobEvenement.PutValue('BEP_MODIFIABLE', 'X');
           if i = 0 then
              TobEvenement.PutValue('BEP_RESPRINCIPALE', 'X')
           else
              TobEvenement.PutValue('BEP_RESPRINCIPALE', '-');
           TobEvenement.PutValue('BEP_EQUIPESEP', '-');
           TobEvenement.PutValue('BEP_OBLIGATOIRE', TobTypeAction.getvalue('BTA_OBLIGATOIRE'));
           TobEvenement.PutValue('BEP_DUREE', Duree);
           TobEvenement.PutValue('BEP_DATEDEB', DateDeb);
           TobEvenement.PutValue('BEP_HEUREDEB', HeureDeb);
           TobEvenement.PutValue('BEP_DATEFIN', DateFin);
           TobEvenement.PutValue('BEP_HEUREFIN', HeureFin);
           TobEvenement.PutValue('BEP_PERIODEDEBUT', '');
           TobEvenement.PutValue('BEP_PERIODEFIN', '');
           TobEvenement.PutValue('BEP_GEREPLAN', TobTypeAction.getvalue('BTA_ASSOSRES'));
           TobEvenement.PutValue('BEP_HEURETRAV', 'X');
           TobEvenement.PutValue('BEP_CREAPLANNING', 'X');
           TobEvenement.PutValue('BEP_BLOCNOTE', '');
           TobEvenement.InsertOrUpdateDB;
           end;
        end;

    TobA.PutValue('AFF_DATEREPONSE', TobEvenement.GetValue('BEP_DATEDEB'));

    TobEvenement.free;
    TobTachesRes.free;
    TobTypeAction.free;

end;

procedure Tof_BTTACHES.ControleLastDateGene(var pListeJours: TlistDate; pLastDateGene: TdateTime; pBoAjout: Boolean);
var
  i: Integer;
begin
  if pBoAjout then
    for i := pListeJours.count - 1 downto 0 do
    begin
      if pListeJours.Items[i] <= pLastDateGene then
        pListeJours.delete(i);
    end;
end;

{*
Function Tof_BTTACHES.RechRessPrinc(Affaire : String) : String;
Var TobTachesRes : Tob;
    vQr          : TQuery;
    vStSQL       : String;
    NoTache      : String;
Begin

    Result  := '';
    NoTache := fTobDet.GetValue('ATA_NUMEROTACHE');

    TobTachesRes := Tob.Create('TACHERESSOURCE', nil, -1);

    vStSQL := 'SELECT ATR_RESSOURCE FROM TACHERESSOURCE WHERE ATR_AFFAIRE="' + Affaire + '"';
    vStSQL := vStSQL + ' AND ATR_NUMEROTACHE=' + NoTache;

    vQr := nil;

    try
      vQR := OpenSQL(vStSQL, True);
      if not (vQR.EOF) then
         Begin
         TobTachesRes.LoadDetailDB('TACHERESSOURCE', '', '', vQR, True);
         Result := TobTachesRes.detail[0].GetValue('ATR_RESSOURCE');
         end;
    finally
      Ferme(vQR);
    end;

    TobTachesRes.free;

end;
*}

//Génération de l'évènement en fonction de l'appel généré et de la ressource associée
Function TOF_BTTACHES.GenereEvt : boolean;
begin
end;

procedure TOF_BTTACHES.RefreshLastDateGene;
var
  vSt: string;
  vQr: TQuery;

begin
  // raffraichir la date de derniere generation
  vSt := 'SELECT ATA_LASTDATEGENE FROM TACHE WHERE ATA_AFFAIRE = "' + fED_AFFAIRE.text + '" AND ATA_NUMEROTACHE = ' + floatToStr(fTobDet.GetValue('ATA_NUMEROTACHE'));
  vQr := OpenSQL(vSt, TRUE);
  try
    if not vQr.EOF then
      SetControlText('ATA_LASTDATEGENE', vQr.FindField('ATA_LASTDATEGENE').AsString);
  finally
    ferme(vQr);
  end;
end;

// champ non-modifiable pour tache créée à partir d'une ligne Affaire
procedure TOF_BTTACHES.RefreshLignePiece;
begin

  SetControlText('ATA_TYPEARTICLE', 'PRE');

  if fTobDet.GetValue('ATA_IDENTLIGNE') <> '' then
     begin
     SetControlEnabled('ATA_CODEARTICLE', false);
     SetControlEnabled('ATA_TYPEARTICLE', false);
     SetControlEnabled('ATA_QTEINITPLA', false);
     SetControlEnabled('ATA_UNITETEMPS', false);
     SetControlEnabled('ATA_UNITETEMPS_', false);
     SetControlEnabled('BTLIGNE', (fAction <> taConsult));
     end
  else
     begin
     SetControlEnabled('BTLIGNE', False);
     SetControlEnabled('ATA_CODEARTICLE', True);
     SetControlEnabled('ATA_TYPEARTICLE', False);
     SetControlEnabled('ATA_UNITETEMPS', True);
     SetControlEnabled('ATA_UNITETEMPS_', True);
     end;

end;

// Initialisation tache créée à partir d'une ligne Affaire
procedure TOF_BTTACHES.TraiteLignePiece;
var
  CleLigne: R_CleDoc;
  Q: TQuery;
  vTobDet: TOB;
begin
  if fStLignePiece = '' then
    exit; // on n'est pas sur une taceh en liaison avec ligne affaire
  vTobDet := fTobTaches.FindFirst(['ATA_IDENTLIGNE'], [fStLignePiece], true);
  if (vTobDet <> nil) then
    fStOpenNumTache := vTobDet.GetValue('ATA_NUMEROTACHE')

  else if fAction <> taConsult then
  begin
    fTobModele := Tob.create('ligne pour tache', nil, -1);
    DecodeRefPiece(fStLignePiece, CleLigne);
    Q := OpenSQL('SELECT GL_DATEPIECE,GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMORDRE,GL_LIBELLE,GL_CODEARTICLE,GL_QTEFACT,GL_QUALIFQTEVTE,GL_TYPEARTICLE,GL_PMRP,GL_PUHT,GL_PUHTDEV,GL_TOTALHT,GL_TOTALHTDEV,GL_FACTURABLE FROM LIGNE WHERE ' + WherePiece(CleLigne, ttdLigne, true, true) + ' ORDER BY GL_NUMLIGNE', True);
    fTobModele.SelectDB('', Q);
    Ferme(Q);
    fStatut := tacreat;
  end;
end;

// Consultation des lignes Affaire
procedure TOF_BTTACHES.BTLIGNEOnClick(SEnder: TObject);
var
  vStLigne: string;
  CleDoc: R_CleDoc;
  //Param: R_SaisiePieceParam;
begin
  vStLigne := fTobDet.Getvalue('ATA_IDENTLIGNE');
  if vStLigne <> '' then
  begin
    DecodeRefPiece(vStLigne, CleDoc);
    //SaisiePiece(CleDoc, taConsult, Param);
  end;
end;

procedure TOF_BTTACHES.BIntervenantsOnClick(Sender: TObject);
begin
  SaisieIntervenantsTache(fTobDet.GetString('ATA_AFFAIRE'), fTobDet.GetString('ATA_NUMEROTACHE'), fTobDet.GetString('ATA_IDENTLIGNE'), fTobDet.GetString('ATA_LIBELLETACHE1'), ActionToString(fAction));
end;

procedure TOF_BTTACHES.ATA_RAPPLANNINGOnClick(Sender: TObject);
begin
  initcolgrid;
end;

procedure TOF_BTTACHES.ATA_TERMINEonClick(Sender: TObject);
begin
  fATATermineChange := True;
end;

procedure TOF_BTTACHES.BtnSelectionOnClick(Sender: TObject);
var
  vStRange: string;
  vStParams: string;
  vStValeur: string;

begin

  vStRange := 'ADR_TYPEADRESSE=INT;ADR_REFCODE=' + fTobDet.GetString('ATA_AFFAIRE');
  vStParams := 'YTC_TIERSLIVRE=' + fTobDet.GetString('ATA_TIERS')
    + ';TYPEADRESSE=INT'
    + ';TITRE='+ traduitGA(TexteMessage[3]) //Adresse d''intervention
    + ';PART='
    + ';CLI='
    + ';TIERS=' + fTobDet.GetString('ATA_TIERS')
    + ';NATUREAUXI='
    + ';ACTION=CREATION'
    + ';AFFAIRE=' + fTobDet.GetString('ATA_AFFAIRE');

  vStValeur := AglLanceFiche('GC', 'GCADRESSES', vStRange, '', vStParams);
  if vStValeur <> '' then
  begin
    SetControlText('ATA_NUMEROADRESSE', vStValeur);
    SetControlText('ATA_TYPEADRESSE', 'INT');
    fTobDet.SetString('ATA_NUMEROADRESSE', vStValeur);
    fTobDet.SetString('ATA_TYPEADRESSE', 'INT');
    fTobDet.SetAllModifie(true);
  end;

  LoadAdresseTache(False, fTobDet);
  //FV1 - 20/04/2018 : FS#3065 - VIVIANE - En création tâche contrat perte information tâche après recherche adresse.
  //fTobDet.PutEcran(Ecran, fPaDet);

  RefreshAdresse;

  //Recherche du Contact et du type d'action evenement en affichage des lignes de la grille
  LectureTypeAction(GetControlText('ATA_BTETAT'));
  LectureContact(GetControlText('CLIENTCONTACT'));

end;

procedure TOF_BTTACHES.BtnEffacerOnClick(Sender: TObject);
begin
  SetControlText('ATA_NUMEROADRESSE', '0');
  SetControlText('ATA_TYPEADRESSE', '');
  fTobDet.PutValue('ATA_NUMEROADRESSE', 0);
  fTobDet.PutValue('ATA_TYPEADRESSE', '');
  fTobDet.SetAllModifie(true);
  LoadAdresseTache(false, fTobDet);
  //FV1 - 20/04/2018 : FS#3065 - VIVIANE - En création tâche contrat perte information tâche après recherche adresse.
  //fTobDet.PutEcran(Ecran, fPaDet);
  RefreshAdresse;
  //Recherche du Contact et du type d'action evenement en affichage des lignes de la grille
  LectureTypeAction(GetControlText('ATA_BTETAT'));
  LectureContact(GetControlText('CLIENTCONTACT'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Franck Vautrain
Créé le ...... : 11/05/2005
Modifié le ... :   /  /
Description .. : Appel de la recherche des contacts dans le cas ou le client
Suite ........ : dispose de + d'un contact
Mots clefs ... :
*****************************************************************}
procedure TOF_BTTACHES.AppelContact(sender : Tobject);
//Var Req         : String;
Var Argument   : String;
		NumContact : String;

begin
(*
  Req := '';
  Req := 'C_AUXILIAIRE="' + fStAuxilliaire + '"';

  SetControlProperty('CLIENTCONTACT', 'Datatype', 'BTCONTACT');
  SetControlProperty('CLIENTCONTACT', 'Plus', req);

  LookupCombo(GetControl('CLIENTCONTACT'));
*)
  Argument := '';
  NumContact := '';

  Argument := 'TYPE=T;';
  Argument := Argument + 'TYPE2=CLI;';
  Argument := Argument + 'PART=-;';
  Argument := Argument + 'TITRE=Contacts de '+ GetControlText('CLIENTLIBELLE') +';';
  Argument := Argument + 'TIERS='+fStTiers+';';
  Argument := Argument + 'ALLCONTACT';

  Numcontact := AGLLanceFiche('YY','YYCONTACT','T;'+fStAuxilliaire,'', Argument);

  if Numcontact = '' then
     NumContact := GetControlText('CLIENTCONTACT')
  Else
     SetControlText('CLIENTCONTACT', NumContact);

  LectureContact(GetControlText('CLIENTCONTACT'));

end;


procedure TOF_BTTACHES.LectureTypeAction (TypeAction : string);
var Req           : String;
    TobTypeAction : Tob;
begin

  SetControlText('LTYPEACTION', '');

  //Lecture du Type ACtion Evenement sélectionné et affichage des informations
  req := 'SELECT BTA_LIBELLE FROM BTETAT WHERE BTA_BTETAT="' + TypeAction + '"';

  TobTypeAction := Tob.Create('#BTETAT',Nil, -1);
  TobTypeAction.LoadDetailFromSQL(Req);
  if TobTypeAction.Detail.Count > 0 then
     Setcontroltext('LTYPEACTION', TobTypeAction.detail[0].GetString('BTA_LIBELLE'))
  else
     Setcontroltext('LTYPEACTION', '');

  TobTypeAction.free;

end;

procedure TOF_BTTACHES.LectureContact(contact : string);
Var Chaine      : String;
    param       : string;
    Req         : String;
    TobContact  : Tob;
    indice      : integer;
begin

  SetControlText('LCONTACT', '');

  Chaine := GetControlText('CLIENTCONTACT');

  if chaine = '' then exit;

  Indice := 0;

	repeat
    param := READTOKENST (Chaine);
    if Param <> '' then
       begin
       if indice = 0 then
          SetControlText('CLIENTCONTACT', Param)
       Else
          SetControlText('CLIENTTELEPHONE', Param);
       inc(indice);
       end;
  until Param = '';

  //Lecture du contact sélectionné et affichage des informations
  req := 'SELECT C_NOM FROM CONTACT WHERE C_TYPECONTACT="T" ';
  Req := Req + 'AND C_AUXILIAIRE="' + fStAuxilliaire + '" ';
//  Req := Req + 'AND C_TIERS="' + fStTiers + '" ';
  Req := Req + 'AND C_NUMEROCONTACT=' + GetControlText('CLIENTCONTACT');

  TobContact := Tob.Create('#CONTACT',Nil, -1);
  TobContact.LoadDetailFromSQL(Req);
  if TobContact.Detail.Count > 0 then
     Setcontroltext('LCONTACT', TobContact.detail[0].GetString('C_NOM'))
  else
     Setcontroltext('CLIENTCONTACT', '');

  TobContact.free;

end;

procedure TOF_BTTACHES.ModifZoneExit(Sender: TObject);
Begin

  // pour que la sortie de la zone permette la modification des règles
  fBoModifRegles := True;

end;

//
function TOF_BTTACHES.CalendrierDefautOk : Boolean;
begin
  result := GetParamSocSecur('SO_AFSTANDCALEN', '') <> '';
end;


procedure TOF_BTTACHES.DefiniLibelle;
var Memo: TMemo;
    lng : integer;
    Indice : integer;
begin
  Memo := TMemo (GetControl ('ATA_DESCRIPTIF') ) ;
  begin
    TobAppel.PutValue('AFF_LIBELLE','');
    // trouve la premiere ligne non vide
    Indice := 0;
    repeat
    	lng := Pos (#$D#$A, Memo.lines [indice] ) ;
      if lng > 1 then break else inc(indice);
    until indice > Memo.lines.count;

    if Indice > Memo.lines.count then exit;

    if (lng > 35) then
    begin
      TobAppel.PutValue ('AFF_LIBELLE', Copy (Memo.lines [Indice] , 0, 35) ) ;
    end
    else
    begin
      TobAppel.PutValue ('AFF_LIBELLE', Copy (Memo.lines [indice] , 0, lng - 1) ) ;
    end;
  end;
end;


initialization
  registerclasses([TOF_BTTACHES]);
end.
