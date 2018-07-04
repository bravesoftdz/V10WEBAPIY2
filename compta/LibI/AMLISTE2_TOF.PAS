{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 16/12/2003
Modifié le ... : 20/12/2004
Description .. : Source TOF de la FICHE : AMLISTE ()
Suite ........ : 20/12/2004 - CA - Suppression systématique du bouton de
Suite ........ : réinitialisation.
Suite ........ : 06/09/2005 - MBO - FQ 14277 - Pas de modif de base sur location longue durée

Mots clefs ... : TOF;AMLISTE
Suite ........ : 09/05 - BTY - En CRC2002-10, dépréciation d'un regoupement
                 TGA 8 Septembre 2005  Ajout "Dépréciation Actif"
                 TGA 09/11/2005 Test si immo déjà dépréciée
                 MBO 14/11/2005 autoriser la saisie de l'exceptionnel sur immo plan variable
                 TGA 17/11/2005 Gestion F11 sur dépréciation et changement de méthode
                 PGR 12/12/2005 Pas de regroupement si immo sortie
Suite..........: - FQ 17215 - TGA 21/12/2005 - GetParamSoc => GetParamSocSecur                 
Suite ........ : 01/06 - BTY - Ancien appel à la liste des immobilisations
*****************************************************************}

unit AMLISTE2_TOF;

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
  dbtables,
  mul,
  fe_main,
  HDB,
{$ELSE}
  eMul,
  MaineAgl,
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
  Outils,
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
  ImDupImo,
  ImModBas,
  IntegEcr,
  ImOutGen,
  AMImmo,
  ImCreSer,
  IMMO_TOM, AMLISTE_TOF   // BTY pointer les procédures publiques de AMLISTE
{$IFDEF SERIE1}
  , S1Util, utModules
{$ELSE}
  ,Ent1
{$ENDIF}

  // TGA - 09/2005 - Dépréciation d'actif
  ,AMDEPRECIATION_TOF
  ,AMDEPRECGROUPE_TOF
  //PGR - 09/2005 - Changement de méthode
  ,AMCHGTMETHODE_TOF;

type
  TOF_AMLISTE2 = class(TOF)
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
    I_METHODEFISC: THValComboBox;
    AMORTDEROG: TCheckBox;
    POPZoomAction: TPopupMenu;
    POPZoomCreation: TPopupMenu;
    POPZoomVisu: TPopupMenu;

   {$IFDEF EAGLCLIENT}
    fListe: THGrid;
   {$ELSE}
    fListe: THDBGrid;
   {$ENDIF}

    { Variables }
    fAction: TActionFiche;
    fbHistorique: boolean;
    fRegroupement: string;
    fCodeImmo: string;
    fWhereDerog: string;
    fWhereOpe: string;
    fWhereOrig: string;
    fWhereOpeAucune: string;
    fTypeOp: TypeOperation;
    fmnToutesOpe: TMenuItem;
    fmnSelectOpe: TMenuItem;
    fmnAucuneOpe: TMenuItem;
    fOnSaveKeyDownEcran: procedure(Sender: TObject; var Key: Word; Shift:
      TShiftState) of object;
    { Méthodes }
    procedure OnListeDblClick(Sender: TObject);
    procedure OnRowEnterListe(Sender: TObject);
    procedure OnAmortDerogClick(Sender: TObject);
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
    // TGA 08/09/2005 CRC200210
    procedure OnDepreciationActifClick(Sender: TObject);
    //PGR - 09/2005 - Changement de méthode
    procedure ChangementMethodeClick(Sender: TObject);
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
    procedure OnZoomPlanClick(Sender: TObject);
    procedure OnZoomEcheanceClick(Sender: TObject);
    procedure OnZoomOperationsClick(Sender: TObject);
    procedure OnZoomEcrituresClick(Sender: TObject);
    procedure OnSupprimeClick(Sender: TObject);
    procedure OnReinitialiseClick(Sender: TObject);
    procedure OnPopupListeClick(Sender: TObject);
    procedure OnFExercice2Change(Sender: TObject);
    procedure OnToutesOpeClick(Sender: TObject);
    procedure OnSelectOpeClick(Sender: TObject);
    procedure OnAucuneOpeClick(Sender: TObject);
    procedure OnOpeIndicateurClick(Sender: TObject);
    procedure EnableCheckBoxOperation;
    procedure EnableZoomAction;
    procedure EnableZoomVisu;
    procedure EnableZoomActionSerie;
    procedure EnableZoomVisuSerie;
    procedure EnableBoutonsAction(Action: TActionFiche);
    procedure RechargeListeChoix;
    procedure ExecuteOperationSerie(TypeOpe: TypeOperation);
    procedure SuppressionFicheImmo;
    procedure GereFiltreOpeEnCours(Etat: TCheckBoxState);
    procedure OnBChercheClick(Sender: TObject);
    procedure OnFiltresChange(Sender: TObject);
    procedure OnFiltreNouvelleRechercheClick(Sender: Tobject);
    procedure OnFiltreSauveClick(Sender: Tobject);
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
  end;

procedure AMLanceFiche2_ListeDesImmobilisations(Compte: string; bHistorique:
  boolean = FALSE; Action: TActionFiche = taModif; Regroupement: string = '');

{ BTY 01/06 Déclarées dans AMLISTE_TOF
*procedure SupprimeFicheImmo(Code: string; bFiche : boolean = True );
function JaiLeDroitSupprimerImmo ( CodeImmo, ChangeCode : string ) : boolean;
function ConfirmationSuppressionImmo ( CodeImmo, EtatImmo : string ) : boolean;
function ValideSuppressionImmo ( stCode : string ) : boolean;
// BTY 09/05 CRC2002-10 Dépréciation d'un regroupement
function  RegroupementDejaDeprecie ( stGroupe : string) : boolean;
function  RegroupementADeprecier ( stCode,stGroupe : string) : boolean;*}


implementation

const
  HM: array[0..17] of string =
  ('Liste des immobilisations',
    'Liste des immobilisations',
    'Cession d''une immobilisation',
    'Eclatement d''une immobilisation',
    'Mutation d''une immobilisation',
    'Changement de plan',
    'Levée de l''option',
    'Annulation de la dernière opération',
    'Confirmez-vous la suppression ?',
    'Aucune immobilisation sélectionnée',
    'Les immobilisations sélectionnées doivent être de nature identique.',
    'Des opérations ont été effectuées sur cette immobilisation. Veuillez les détruire avant de supprimer la fiche.',
    'Cette immobilisation a été créée suite à une opération. Pour la détruire, veuillez annuler l''opération.',
    'Des opérations ont été effectuées sur cette immobilisation. Duplication de l''immobilisation impossible.',
    'Pour effectuer l''intégration des ecritures, le compte fournisseur doit être impérativement renseigné.',
    'Attention : cette immobilisation a été créée sur un exercice clos.#10#13La détruire va entraîner une modification de l''historique.#10#13Voulez-vous continuer ?',
    'Des opérations ont été effectuées sur cette immobilisation. Duplication de l''immobilisation impossible.',
    'Aucune immobilisation sélectionnée'
    );
  HM2: array[0..8] of string =
  ('En cours :',
    'Mutation',
    'Eclatement',
    'Sortie',
    'Modification des bases',
    'Changement de lieu',
    'Changement d''établissement',
    'Changement de plan',
    'Levée de l''option'
    );

procedure CreationLieuParDefaut;
var
  Q: TQuery;
  T: TOB;
begin
  Q := OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="GEO"', True);
  if (Q.Eof) and (VHImmo^.EtablisDefaut <> '') then
  begin
    T := TOB.Create('CHOIXCOD', nil, -1);
    T.PutValue('CC_TYPE', 'GEO');
    T.PutValue('CC_CODE', VHImmo^.EtablisDefaut);
    T.PutValue('CC_LIBELLE', RechDom('TTETABLISSEMENT', VHImmo^.EtablisDefaut,
      False));
    T.PutValue('CC_ABREGE', RechDom('TTETABLISSEMENT', VHImmo^.EtablisDefaut,
      True));
    T.InsertDB(nil);
    T.Free;
    AvertirTable('TILIEUGEO');
  end;
  Ferme(Q);
end;

procedure AMLanceFiche2_ListeDesImmobilisations(Compte: string; bHistorique:
  boolean = FALSE; Action: TActionFiche = taModif; Regroupement: string = '');
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
    stAction := stAction + ';' + Regroupement;
  CreationLieuParDefaut;
  AGLLanceFiche('AM', 'AMLISTE2', '', '', stAction);
end;

function VerifNature(NatureIni, Nature: string): boolean;
begin
  if ((NatureIni = 'LOC') or (NatureIni = 'CB')) and ((Nature = 'LOC') or (Nature
    = 'CB')) then
    Result := True
  else if ((NatureIni = 'FI') or (NatureIni = 'PRO')) and ((Nature = 'PRO') or
    (Nature = 'FI')) then
    Result := True
  else
    Result := False;
end;

procedure TOF_AMLISTE2.OnNew;
begin
  inherited;
end;

procedure TOF_AMLISTE2.OnDelete;
begin
  inherited;
end;

procedure TOF_AMLISTE2.OnUpdate;
begin
  inherited;
end;

procedure TOF_AMLISTE2.OnLoad;
var
  F: TFMul;
begin
  inherited;
  F := TFMul(Ecran);
  F.Q.Manuel := True;
    // Evite l'exécution de la requête lors de la maj de Q.Liste
  // Affectation de la liste ad-hoc
  if fAction = taConsult then
    F.Q.Liste := 'MULVIMMOS'
  else
    F.Q.Liste := 'MULMIMMOS';
  F.Q.Manuel := False;
  if fRegroupement <> '' then
  begin
    SetControlText('I_GROUPEIMMO', fRegroupement);
    SetControlEnabled('I_GROUPEIMMO', False);
  end;
  //  OnRowEnterListe(nil);
end;

procedure TOF_AMLISTE2.OnArgument(S: string);
var
  St: string;
  okok: boolean;
begin
  inherited;
  { Action }
  St := ReadTokenSt(S);
  if St = 'ACTION=MODIFICATION' then
    fAction := taModif
  else
    fAction := taConsult;
  { Compte }
  St := ReadTokenSt(S);
  if St <> '' then
  begin
//    SetControlEnabled('I_COMPTEIMMO', False);
    SetControlText('I_COMPTEIMMO', St);
  end;
  { Historique }
  St := ReadTokenSt(S);
  fbHistorique := (St = 'HISTO');

  { Regroupement }
  fRegroupement := S;

  fTypeOp := toNone;

{$IFDEF SERIE1}
  if fbHistorique then
    Ecran.HelpContext := 511500
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

{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}
{$IFDEF EAGLCLIENT}
  fListe := THGrid(GetControl('FLISTE'));
{$ELSE}
  fListe := THDBGrid(GetControl('FLISTE'));
{$ENDIF}
  fListe.OnDblClick := OnListeDblClick;
{$IFDEF EAGLCLIENT}
  fListe.SortEnabled := False;
{$ELSE}
  fListe.OnRowEnter := OnRowEnterListe;
  fListe.SortEnabled := True;
{$ENDIF}
  fListe.OnKeyDown := OnKeyDownFListe;


  fOnSaveKeyDownEcran := Ecran.OnKeyDown;
  Ecran.OnKeyDown := OnKeyDownEcran;
  TFMul(Ecran).OnAfterFormShow := AfterShow;
  fListe.OnFlipSelection := OnFlipSelectionGrille;
  TFMul(Ecran).bSelectAll.OnClick := OnSelectAllClick;

  TToolbarButton97(GetControl('bDelete')).OnClick := OnSupprimeClick;
  TToolbarButton97(GetControl('bReinit')).OnClick := OnReinitialiseClick;
  TToolbarButton97(GetControl('bCherche')).OnClick := OnBChercheClick;
  THValComboBox(GetControl('FExercice2')).OnChange := OnFExercice2Change;
  THCritMaskEdit(GetControl('I_COMPTEIMMO')).OnelipsisClick :=
    OnCompteelipsisClick;
  THCritMaskEdit(GetControl('I_COMPTELIE')).OnelipsisClick :=
    OnCompteelipsisClick;
  THCritMaskEdit(GetControl('I_ORGANISMECB')).OnelipsisClick :=
    OnOrganismeelipsisClick;

  { Popup Création }
  POPZoomCreation := TPopUpMenu(GetControl('POPZOOMCREATION', True));
  PopZoomCreation.OnPopup := OnPopZoomCreationPopup;
  TMenuItem(GetControl('PleinePropriete')).OnClick := OnCreationPROClick;
  TMenuItem(GetControl('ImmoFinanciere')).OnClick := OnCreationFIClick;
  TMenuItem(GetControl('CreditBail')).OnClick := OnCreationCBClick;
  TMenuItem(GetControl('LocationFinanciere')).OnClick := OnCreationLOCClick;
  TMenuItem(GetControl('CreationSerie')).OnClick := OnCreationSerieClick;
  TMenuItem(GetControl('DupliqueImmo')).OnClick := OnDupliqueImmoClick;

  { Popup Action }
  POPZoomAction := TPopUpMenu(GetControl('POPZOOMACTION', True));
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
  // TGA 08/09/2005 CRC200210
  TMenuItem(GetControl('DepActif')).OnClick := OnDepreciationActifClick;
  //PGR - 09/2005 - Changement de méthode
  TMenuItem(GetControl('ChangementMethode')).OnClick := ChangementMethodeClick;
  { Popup Visualisation }
  POPZoomVisu := TPopUpMenu(GetControl('POPZOOMVISU', True));
  PopZoomVisu.OnPopup := OnPopZoomVisuPopup;
  TMenuItem(GetControl('ZoomPlan')).OnClick := OnZoomPlanClick;
  TMenuItem(GetControl('ZoomEcheance')).OnClick := OnZoomEcheanceClick;
  TMenuItem(GetControl('ZoomOperations')).OnClick := OnZoomOperationsClick;
  TMenuItem(GetControl('Ecritures')).OnClick := OnZoomEcrituresClick;

  { Popup Liste générale - F11 }
  TPopUpMenu(GetControl('PopupListe')).OnPopup := OnPopupListeClick;
  TMenuItem(GetControl('PleinePropriete1')).OnClick := OnCreationPROClick;
  TMenuItem(GetControl('ImmoFinanciere1')).OnClick := OnCreationFIClick;
  TMenuItem(GetControl('CreditBail1')).OnClick := OnCreationCBClick;
  TMenuItem(GetControl('LocationFinanciere1')).OnClick := OnCreationLOCClick;
  TMenuItem(GetControl('CreationSerie1')).OnClick := OnCreationSerieClick;
  TMenuItem(GetControl('DupliqueImmoListe')).OnClick := OnDupliqueImmoClick;
  TMenuItem(GetControl('ZoomPlan1')).OnClick := OnZoomPlanClick;
  TMenuItem(GetControl('ZoomEcheance1')).OnClick := OnZoomEcheanceClick;
  TMenuItem(GetControl('ZoomOperations1')).OnClick := OnZoomOperationsClick;
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

  // TGA 17/11/2005
  TMenuItem(GetControl('DepActif1')).OnClick := OnDepreciationActifClick;
  TMenuItem(GetControl('ChangementMethode1')).OnClick := ChangementMethodeClick;

  TMenuItem(GetControl('ChangeEtablissement1')).OnClick :=
    OnChangeEtablissementClick;
  TMenuItem(GetControl('ChangeLieu1')).OnClick := OnChangeLieuClick;
  TMenuItem(GetControl('ModifBases1')).OnClick := OnModifBasesClick;
  TMenuItem(GetControl('ActionRegroupement1')).OnClick :=
    OnActionRegroupementClick;

  { Opérations }
  fmnToutesOpe := TMenuItem(GetControl('ToutesOpe'));
  fmnSelectOpe := TMenuItem(GetControl('SelectOpe'));
  fmnAucuneOpe := TMenuItem(GetControl('AucuneOpe'));
  TMenuItem(GetControl('ToutesOpe')).OnClick := OnToutesOpeClick;
  TMenuItem(GetControl('SelectOpe')).OnClick := OnSelectOpeClick;
  TMenuItem(GetControl('AucuneOpe')).OnClick := OnAucuneOpeClick;
  I_METHODEFISC := THValComboBox(GetControl('I_METHODEFISC'));
  AMORTDEROG := TCheckBox(GetControl('AMORTDEROG'));
  AMORTDEROG.OnClick := OnAmortDerogClick;
  TCheckBox(GetControl('OPEMUTATION')).OnClick := OnOpeIndicateurClick;
  TCheckBox(GetControl('OPECHANGEPLAN')).OnClick := OnOpeIndicateurClick;
  TCheckBox(GetControl('OPELEVEEOPTION')).OnClick := OnOpeIndicateurClick;
  TCheckBox(GetControl('OPECESSION')).OnClick := OnOpeIndicateurClick;
  TCheckBox(GetControl('OPELIEUGEO')).OnClick := OnOpeIndicateurClick;
  TCheckBox(GetControl('OPEETABLISSEMENT')).OnClick := OnOpeIndicateurClick;
  TCheckBox(GetControl('OPEECLATEMENT')).OnClick := OnOpeIndicateurClick;
  TCheckBox(GetControl('REINTEGRATION')).OnClick := OnOpeIndicateurClick;
  TCheckBox(GetControl('QUOTEPART')).OnClick := OnOpeIndicateurClick;
  { Filtres }
  TMenuItem(GetControl('BNouvRech')).OnClick := OnFiltreNouvelleRechercheClick;
  TMenuItem(GetControl('BSaveFiltre')).OnClick := OnFiltreSauveClick;
  THValComboBox(GetControl('FFiltres')).OnChange := OnFiltresChange;
  { Tables Libres }
  THValComboBox(GetControl('TABLELIBRE1')).OnChange := OnChangeChampLibre;
  THValComboBox(GetControl('TABLELIBRE2')).OnChange := OnChangeChampLibre;
  THValComboBox(GetControl('TABLELIBRE3')).OnChange := OnChangeChampLibre;

  MakeZoomOle(Ecran.Handle);

  { Initialisation des zones }

  SetControlVisible('bInsert', not fbHistorique);
  SetControlVisible('bSelectAll', not fbHistorique);
  SetControlVisible('bZoomVisu', not fbHistorique);
  SetControlVisible('bZoomAction', not fbHistorique);

  SetControlVisible('bDelete', not fbHistorique);
//  SetControlVisible('bReinit', not fbHistorique);
  SetControlVisible('bReinit', False);
  SetControlVisible('POperations', not fbHistorique);
  SetControlVisible('PANELENCOURS', not fbHistorique);

{$IFDEF SERIE1}
  SetControlVisible('I_ETABLISSEMENT', False);
  SetControlVisible('tI_ETABLISSEMENT', False);
  SetControlVisible('OPEETABLISSEMENT', False);
{$ENDIF}

{$IFDEF SERIE1}
  ImLibellesTableLibre(TTabSheet(GetControl('PzLibreS1')), 'TT_TABLELIBREIMMO',
    '', 'I');
  Okok := false;
{$ELSE}
  ImLibellesTableLibre(TTabSheet(GetControl('PzLibre')), 'TI_TABLE', 'I_TABLE',
    'I');
  Okok := true;
{$ENDIF}
  SetControlVisible('PzLibre', Okok);
  SetControlVisible('PzLibreS1', not Okok);

  //  SetControlProperty('I_METHODEECO','Plus',' AND CO_CODE<>"VAR"');
  SetControlProperty('I_METHODEFISC','Plus',' AND CO_CODE<>"VAR" AND CO_CODE<>"NAM"');

  // FQ 16283 TGA 27/07/20005
  SetControlProperty('I_MONTANTHT', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('I_MONTANTHT_', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('I_BASETAXEPRO', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('I_BASETAXEPRO_', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  // Fin FQ 16283 TGA 27/07/20005

  if fbHistorique then
  begin
    if GetControlText('XX_WHERE') = '' then
      SetControlText('XX_WHERE', 'I_ETAT="FER"')
    else
      SetControlText('XX_WHERE', GetControlText('XX_WHERE') +
        ' AND I_ETAT="FER"');
  end
  else
  begin
    if GetControlText('XX_WHERE') = '' then
      SetControlText('XX_WHERE', 'I_ETAT<>"FER"')
    else
      SetControlText('XX_WHERE', GetControlText('XX_WHERE') +
        ' AND I_ETAT<>"FER"');
  end;

  fWhereOpeAucune := ' AND (I_OPEMUTATION="-" AND I_OPECHANGEPLAN="-" AND ' +
    'I_OPELIEUGEO="-" AND I_OPEETABLISSEMENT="-" AND ' +
    'I_OPELEVEEOPTION="-" AND I_OPECESSION="-" AND ' +
    'I_OPEECLATEMENT="-" AND I_REINTEGRATION=0.0 AND ' +
    'I_QUOTEPART=0)';
  fWhereOrig := GetControlText('XX_WHERE');
  fmnToutesOpe.Checked := GetCheckBoxState('FiltreOpe') = cbChecked;
  fmnAucuneOpe.Checked := GetCheckBoxState('FiltreOpe') = cbGrayed;
  fmnSelectOpe.Checked := GetCheckBoxState('FiltreOpe') = cbUnchecked;

  EnableCheckBoxOperation;
  OnAmortDerogClick(nil);

  SetControlEnabled('BZOOMACTION', False);

  EnableBoutonsAction(fAction);

  // Gestion restriction sur Etablissement
{$IFNDEF SERIE1}
  PositionneEtabUser( THMultiValComboBox(GetControl('I_ETABLISSEMENT')) ) ;
{$ENDIF}

end;

procedure TOF_AMLISTE2.OnClose;
begin
  inherited;
end;

procedure TOF_AMLISTE2.OnDisplay();
begin
  inherited;
  EnableBoutonsAction(fAction);
  OnRowEnterListe(nil);
  { Table libre 9 utilisée pour le stockage 'éligible TP' }
  SetControlVisible ('I_TABLE9',False);
  SetControlVisible ('TI_TABLE9',False);
end;

procedure TOF_AMLISTE2.OnCancel();
begin
  inherited;
end;

procedure TOF_AMLISTE2.OnListeDblClick(Sender: TObject);
var
  stCode: string;
begin
{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}

  if TheMulQ.RecordCount <> 0 then
  begin
    stCode := GetField('I_IMMO');
    //    if fTypeOp = toNone then FicheImmobilisation(TQuery (GetControl ('Q')), stCode, fAction,'') ;
    if fTypeOp = toNone then
      begin
        AMLanceFiche_FicheImmobilisation(stCode, fAction, '');

        // TGA 26/07/2005 FQ 15301
        if (fmnAucuneOpe.Checked) then
           SetControlText('XX_WHERE', GetControlText('XX_WHERE') + fWhereOpeAucune)
        else if (fmnToutesOpe.Checked) then
           SetControlText('XX_WHERE', GetControlText('XX_WHERE'))
        else
           SetControltext('XX_WHERE', GetControlText('XX_WHERE') + fWhereOpe);
        // Fin TGA 26/07/2005 FQ 15301

      end;

    TFMul(ECRAN).BChercheClick(nil);
    RechargeListeChoix;
    EnableBoutonsAction(fAction);
    SelectionEnregistrement(stCode);
  end;
end;

procedure TOF_AMLISTE2.EnableCheckBoxOperation;
begin
  SetControlEnabled('OPEECLATEMENT', fmnSelectOpe.Checked);
  SetControlEnabled('OPECHANGEPLAN', fmnSelectOpe.Checked);
  SetControlEnabled('OPELIEUGEO', fmnSelectOpe.Checked);
  SetControlEnabled('OPEETABLISSEMENT', fmnSelectOpe.Checked);
  SetControlEnabled('OPELEVEEOPTION', fmnSelectOpe.Checked);
  SetControlEnabled('OPECESSION', fmnSelectOpe.Checked);
  SetControlEnabled('OPEMUTATION', fmnSelectOpe.Checked);
  SetControlEnabled('REINTEGRATION', fmnSelectOpe.Checked);
  SetControlEnabled('QUOTEPART', fmnSelectOpe.Checked);
end;

procedure TOF_AMLISTE2.OnAmortDerogClick(Sender: TObject);
begin
  if GetCheckBoxState('AMORTDEROG') = cbGrayed then
    fWhereDerog := ''
  else if GetCheckBoxState('AMORTDEROG') = cbChecked then
    fWhereDerog := 'AND I_METHODEFISC <>""'
  else
    fWhereDerog := 'AND I_METHODEFISC =""';
end;

procedure TOF_AMLISTE2.EnableBoutonsAction(Action: TActionFiche);
begin

  SetControlVisible('BZoomAction', (Action <> taConsult) and (not ListeVide));
//  SetControlVisible('BREINIT', (Action <> taConsult) and (not ListeVide));
  SetControlVisible('BREINIT', False );
  SetControlVisible('BInsert', (Action <> taConsult));
{$IFDEF SERIE1}
  SetControlVisible('BDelete', (Action <> taConsult) and (not ListeVide));
{$ELSE}
  SetControlVisible('BDelete', (Action <> taConsult) and (V_PGI.LaSerie = S5)
    and (not ListeVide));
{$ENDIF}
end;

procedure TOF_AMLISTE2.RechargeListeChoix;
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

procedure TOF_AMLISTE2.OnCreationPROClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := '';
{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}
  if TheMulQ.RecordCount <> 0 then
    stCode := GetField('I_IMMO');
  AMLanceFiche_FicheImmobilisation('', taCreatOne, 'PRO');
  TFMul(ECRAN).BChercheClick(nil);
  RechargeListeChoix;
  SelectionEnregistrement(stCode);
  EnableBoutonsAction(fAction);
end;

procedure TOF_AMLISTE2.OnCessionClick(Sender: TObject);

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
  mr := mrYes;
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

  { On se repositionne sur l'enregistrement de départ }
  if mr = mrYes then
  begin
    TFMul(ECRAN).BChercheClick(nil);
    SelectionEnregistrement(stCode);
  end;
end;

procedure TOF_AMLISTE2.OnEclatementClick(Sender: TObject);
var
  mr: integer;
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  mr := ExecuteEclatement(GetField('I_IMMO'));
  if mr = mrYes then
  begin
    TFMul(ECRAN).BChercheClick(nil);
    SelectionEnregistrement(stCode);
  end;
end;

procedure TOF_AMLISTE2.OnMutationClick(Sender: TObject);
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
    TFMul(ECRAN).BChercheClick(nil);
    SelectionEnregistrement(stCode);
  end;
end;

procedure TOF_AMLISTE2.OnAnnulationOperationClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  OperationsEnCours(GetField('I_IMMO'), GetField('I_LIBELLE'), True, fAction);
  TFMul(ECRAN).BChercheClick(nil);
  SelectionEnregistrement(stCode);
  OnRowEnterListe(nil);
end;

procedure TOF_AMLISTE2.OnLeveeOptionClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  if ExecuteLeveeOption(GetField('I_IMMO')) = mrYes then
  begin
    TFMul(ECRAN).BChercheClick(nil);
    SelectionEnregistrement(stCode);
  end;
end;



// BTY 09/05 CRC200210 Ajout de la dépréciation d'un regroupement
// En multisélection
// option ouverte si aucune immo dans un regroupement => dépréciations individuelles
//             ou si toutes dans le même sans dépréciation précédente dans l'exo=> question pour déprécier tout le regroupement
//             sinon option fermée
// En monosélection
// si immo sans regroupement => dépréciation individuelle
// sinon sans dépréciation précédente dans l'exo => question pour déprécier tout le regroupement

procedure TOF_AMLISTE2.OnDepreciationActifClick (Sender: TObject);

var
  i : integer;
  stCode: string;
  TListeAll : TOB;
  stGroupeImmo: string;
  mr : string;
begin
  stCode := GetField('I_IMMO');
  if not VarIsNull ( TheMulQ.FindField('I_GROUPEIMMO').AsVariant) then
    stGroupeImmo := GetField('I_GROUPEIMMO')
  else stGroupeImmo := '';
  mr := IntToStr(mrYes);
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
                for i := 0 to TListeAll.Detail.Count - 1 do
                begin
                  // Dépréciations individuelles
                  mr := AMLanceFiche_DepreciationActif(TListeAll.Detail[i].GetValue('I_IMMO'));
                  if mr <> IntToStr (mrYes) then break;
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
          for i := 0 to fListe.NbSelected - 1 do
          begin
            // Dépréciations individuelles
            fListe.GotoLeBookmark(i);
            mr := AMLanceFiche_DepreciationActif(GetField('I_IMMO'));
            if mr <> IntToStr(mrYes) then   break;
          end;
        end;

      //end;
      fListe.ClearSelected;

    end;
  end;

  { On se repositionne sur l'enregistrement de départ }
  if mr = IntToStr(mrYes) then
  begin
    TFMul(ECRAN).BChercheClick(nil);
    SelectionEnregistrement(stCode);
  end;

end;


//PGR - 09/2005 - Changement de méthode
procedure TOF_AMLISTE2.ChangementMethodeClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');

  if AmLanceFiche_FicheChgtMethode (GetField('I_IMMO')) = mrYes then
  begin
    TFMul(ECRAN).BChercheClick(nil);
    SelectionEnregistrement(stCode);
  end;

end;

procedure TOF_AMLISTE2.OnDureeMethodeClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  if ExecuteChangePlan(GetField('I_IMMO'), 1) = mrYes then
  begin
    TFMul(ECRAN).BChercheClick(nil);
    SelectionEnregistrement(stCode);
  end;
end;

procedure TOF_AMLISTE2.OnElementExceptionnelClick(Sender: TObject);
var
  stCode: string;
begin
  inherited;
  stCode := GetField('I_IMMO');
  if ExecuteChangePlan(GetField('I_IMMO'), 2) = mrYes then
  begin
    TFMul(ECRAN).BChercheClick(nil);
    SelectionEnregistrement(stCode);
  end;
end;

procedure TOF_AMLISTE2.ExecuteOperationSerie(TypeOpe: TypeOperation);
var
  FOpeSerie: TOperationSerie;
  i, OrdreSerie, ret: integer;
  QTmp, QSortie: TQuery;
  FCodeImmo: string;
  TListeAll : TOB;
begin
  FOpeSerie := TOperationSerie.Create(Application);
  try
    if TypeOpe = toMutation then
      FOpeSerie.fNature := GetField('I_NATUREIMMO');
    FOpeSerie.InitOperationSerie(TypeOpe, FListe.NbSelected);
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
              end;
            end;
          end;
          fListe.AllSelected := False;
        finally
          FreeAndNil (TListeAll);
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
                FOpeSerie.OperationSerieRegroupement(FCodeImmo);
              end;
          end;
        end;
        FListe.ClearSelected;
      end;
      FiniMove;
    end;
  finally
    SetControlProperty('bSectAll', 'Down', FListe.AllSelected);
    FOpeSerie.DetruitOperationSerie;
  end;
end;

procedure TOF_AMLISTE2.EnableZoomAction;
var
  Nat: string;
  bDepotGar, bAcquisitionAnnee: boolean;
  bCedee: boolean;
  bDep : Boolean;
  stcode : String;
begin
  bCedee := (GetField('I_QUANTITE') = 0);
  SetControlVisible('DupliqueImmo', not ListeVide);
  bAcquisitionAnnee := (GetField('I_DATEPIECEA') >= VHImmo^.Encours.Deb)
    and (GetField('I_DATEPIECEA') <= VHImmo^.Encours.Fin);
  SetControlVisible('Cession', false);
  SetControlVisible('Eclatement', false);
  SetControlVisible('Mutation', false);
  SetControlVisible('DureeMethode', false);
  SetControlVisible('ElementExceptionnel', false);
  SetControlVisible('LeveeOption', false);
  SetControlVisible('AnnulationOperation', false);
  SetControlVisible('ChangeEtablissement', false);
  SetControlVisible('ChangeLieu', false);
  SetControlVisible('ModifBases', false);
  SetControlVisible('ActionRegroupement', false);
  SetControlVisible('Ecritures1', SelectionUnique);
  // TGA 08/09/2005 CRC200210
  SetControlVisible('DepActif', false);
  //PGR - 09/2005 - Changement de méthode
  SetControlVisible('ChangementMethode', False);

  // TGA 09/11/2005 Test si immo déjà dépréciée
  // Pas d'éclatement, de cession, de modification de base, d'exceptionnel
  // de changement de méthode si bdep=True
  Stcode := GetField('I_IMMO');
  if ExisteSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+stCode+'" AND I_REVISIONECO<>0') then
    bdep:=True
  Else
    bdep:=False;

  if ListeVide or (not SelectionUnique) then
    exit;

  Nat := GetField('I_NATUREIMMO');

  // TGA 08/09/2005 CRC200210 Autorisation Dépréciation Actif
  SetControlVisible('DepActif', Nat = 'PRO');

  // PGR 06/12/2005 Autorisation Changement de méthode
  SetControlVisible('ChangementMethode', (Nat = 'PRO')
   and (not bAcquisitionAnnee)and not (getField('I_METHODEECO')='VAR') and (not bdep)and (not bCedee)
   and (not (Getfield('I_OPECHANGEPLAN') = 'X')) and (not (Getfield('I_OPEMODIFBASES') = 'X')));

  // Autorisation Cession
  SetControlVisible('Cession', ((Nat = 'PRO') or (Nat = 'FI')) and
    (not bCedee) and (not bdep) );

  // Autorisation Eclatement
  //ATTENTION : le champ immolie est utilisé pour identifier l'immo d'origine pour un depot de gar.
  bDepotGar := (Nat = 'FI') and (GetField('I_IMMOLIE') <> '');
  SetControlVisible('Eclatement', (Nat <> 'CB') and (Nat <> 'LOC') and (not
    bDepotGar and not((GetField('I_METHODEECO')='VAR') and not bAcquisitionAnnee) )
    and (not bCedee) and (not (GetField('I_OPELEVEEOPTION') = 'X')) and (not bdep));

  // Autorisation Mutation
  SetControlVisible('Mutation', (not bCedee) and (not
    (GetField('I_OPELEVEEOPTION') = 'X')) and (not bAcquisitionAnnee));

  // Autorisation Changement de plan
  SetControlVisible('DureeMethode', (Nat = 'PRO') and (not bAcquisitionAnnee)
    and (not bCedee) and (not (GetField('I_OPECHANGEPLAN') = 'X')) and ((GetField('I_METHODEECO')<>'VAR')));

  // Autorisation exceptionnel - modif mbo le 14.11.2005
  SetControlVisible('ElementExceptionnel', (Nat = 'PRO') and (not bCedee) and
    (not (GetField('I_OPECHANGEPLAN') = 'X')) and (not (GetField('I_METHODEECO')='NAM'))
    and (not bdep));

  // Autorisation Levée option
  SetControlVisible('LeveeOption', (Nat = 'CB') and (GetField('I_IMMOLIE') = '')
    and (not bCedee) and (not (GetField('I_OPELEVEEOPTION') = 'X')));

  // Opérations ?
  SetControlVisible('AnnulationOperation', GetField('I_OPERATION') = 'X');
  SetControlVisible('SeparAnnulAction', GetControlVisible('AnnulationOperation'));
  SetControlVisible('ChangeEtablissement', not bCedee and (not
    (GetField('I_OPELEVEEOPTION') = 'X')) and (VHImmo^.EtablisCpta = TRUE));
  SetControlVisible('ChangeLieu', not bCedee and (not
    (GetField('I_OPELEVEEOPTION') = 'X')));

  // mbo FQ 14277 pas de modif base en location longue durée
  SetControlVisible('ModifBases',(Nat <> 'LOC') and (Nat <> 'FI') and (not bCedee) and (not
    bAcquisitionAnnee) and (not (GetField('I_OPELEVEEOPTION') = 'X'))
     and (GetField('I_METHODEECO')<>'VAR') and (not bdep) );
   //PGR - 12/12/2005 Pas de regroupement si immo sortie
  //SetControlVisible('ActionRegroupement', True);
  SetControlVisible('ActionRegroupement', not bCedee);
end;

procedure TOF_AMLISTE2.EnableZoomVisu;
var
  QLog: TQuery;
  Nature: string;
begin
  inherited;
  SetControlVisible('ZoomPlan', false);
  SetControlVisible('ZoomEcheance', false);
  SetControlVisible('ZoomOperations', false);
  SetControlVisible('Ecritures',False);
  if ListeVide then
    exit;
  SetControlVisible('ZoomPlan', ImmoAmortie(GetField('I_METHODEECO'),
    varToStr(GetField('I_METHODEFISC')), GetField('I_NATUREIMMO')));
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
  SetControlVisible('Ecritures', SelectionUnique);
end;

procedure TOF_AMLISTE2.OnPopZoomActionPopup(Sender: TObject);
begin
  if (FListe.NbSelected <= 1) and (not FListe.AllSelected) then
    EnableZoomAction
  else
    EnableZoomActionSerie;
end;

procedure TOF_AMLISTE2.OnZoomEcheanceClick(Sender: TObject);
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

procedure TOF_AMLISTE2.OnZoomPlanClick(Sender: TObject);
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

procedure TOF_AMLISTE2.OnZoomOperationsClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  OperationsEnCours(GetField('I_IMMO'), GetField('I_LIBELLE'), false, fAction);
  TFMul(ECRAN).BChercheClick(nil);
  SelectionEnregistrement(stCode);
end;

procedure TOF_AMLISTE2.OnPopZoomVisuPopup(Sender: TObject);
begin
  if (FListe.NbSelected <= 1) and (not FListe.AllSelected) then
    EnableZoomVisu
  else
    EnableZoomVisuSerie;
end;

procedure TOF_AMLISTE2.OnCreationCBClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := '';
  if not TheMulQ.Eof then
    stCode := GetField('I_IMMO');
  AMLanceFiche_FicheImmobilisation('', taCreatOne, 'CB');
  TFMul(ECRAN).BChercheClick(nil);
  RechargeListeChoix;
{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}
  SelectionEnregistrement(stCode);
  EnableBoutonsAction(fAction);
end;

procedure TOF_AMLISTE2.OnCreationFIClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := '';
  if not TheMulQ.Eof then
    stCode := GetField('I_IMMO');
  AMLanceFiche_FicheImmobilisation('', tacreatOne, 'FI');
  TFMul(ECRAN).BChercheClick(nil);
  RechargeListeChoix;
{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}
  SelectionEnregistrement(stCode);
  EnableBoutonsAction(fAction);
end;

procedure TOF_AMLISTE2.OnCreationLOCClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := '';
  if not TheMulQ.Eof then
    stCode := GetField('I_IMMO');
  AMLanceFiche_FicheImmobilisation('', taCreatOne, 'LOC');
  TFMul(ECRAN).BChercheClick(nil);
  RechargeListeChoix;
{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}
  SelectionEnregistrement(stCode);
  EnableBoutonsAction(fAction);
end;

procedure TOF_AMLISTE2.OnCreationSerieClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := '';
  if not TheMulQ.Eof then
    stCode := GetField('I_IMMO');
  AMLanceFiche_FicheImmobilisation('', taCreat, '');
  TFMul(ECRAN).BChercheClick(nil);
  RechargeListeChoix;
{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}
  SelectionEnregistrement(stCode);
  EnableBoutonsAction(fAction);
end;

procedure TOF_AMLISTE2.OnDupliqueImmoClick(Sender: TObject);
var
  NewCodeImmo: string;
  ret: integer;
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  if not IsOpeEnCours(nil, GetField('I_IMMO'), false) then
  begin
    ret := ExecuteDuplication(GetField('I_IMMO'), NewCodeImmo);
    if ret = mrYes then
      AMLanceFiche_FicheImmobilisation(NewCodeImmo, taModif, '');
    TFMul(ECRAN).BChercheClick(nil);
    SelectionEnregistrement(stCode);
  end
  else
    PGIBox(HM[13]);
end;

procedure TOF_AMLISTE2.OnSupprimeClick(Sender: TObject);
var
  i: integer;
  stCode: string;
  TListeAll : TOB;
begin
  stCode := GetField('I_IMMO');
  if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
  begin
    PGIBox(HM[9]);
    exit;
  end;
//  if ( PGIAsk(HM[8] ) <> mrYes) then exit;
//  if not ConfirmationSuppressionImmo ( GetField('I_IMMO'), GetField('I_ETAT') ) then exit;
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
//        if not JaiLeDroitSupprimerImmo(TListeAll.Detail[i].GetValue('I_IMMO'), TListeAll.Detail[i].GetValue('I_CHANGECODE')) then continue;
        if ValideSuppressionImmo ( FCodeImmo ) then
        begin
          if Transactions(SuppressionFicheImmo, 3) <> oeOk then
          begin
            MessageAlerte(HM[0]);
            Break;
          end else ImMarquerPublifi(True);
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
      if ValideSuppressionImmo ( FCodeImmo ) then
      begin
        if Transactions(SuppressionFicheImmo, 3) <> oeOk then
        begin
          MessageAlerte(HM[0]);
          Break;
        end;
      end;
    end;
    FListe.ClearSelected;
  end;
  FiniMove;
  TFMul(ECRAN).BChercheClick(nil);
  SelectionEnregistrement(stCode);
  EnableBoutonsAction(fAction);
end;

procedure TOF_AMLISTE2.SuppressionFicheImmo;
begin
  SupprimeFicheImmo(FCodeImmo)
end;

procedure TOF_AMLISTE2.OnPopupListeClick(Sender: TObject);
begin
  OnPopZoomVisuPopup(Sender);
  OnPopZoomActionPopup(Sender);
  SetControlVisible('DupliqueImmoListe', (not ListeVide) and (FListe.NbSelected
    <= 1));
  SetControlVisible('Cession1', GetControlVisible('Cession'));
  SetControlVisible('Eclatement1', GetControlVisible('Eclatement'));
  SetControlVisible('Mutation1', GetControlVisible('Mutation'));
  SetControlVisible('DureeMethode1', GetControlVisible('DureeMethode'));
  SetControlVisible('ElementExceptionnel1',GetControlVisible('ElementExceptionnel'));

  // TGA 17/11/2005
  SetControlVisible('DepActif1',GetControlVisible('DepActif'));
  SetControlVisible('ChangementMethode1',GetControlVisible('ChangementMethode'));

  SetControlVisible('LeveeOption1', GetControlVisible('LeveeOption'));
  SetControlVisible('AnnulationOperation1',
    GetControlVisible('AnnulationOperation'));
  SetControlVisible('SeparAnnul', GetControlVisible('AnnulationOperation1'));
  SetControlVisible('ZoomPlan1', GetControlVisible('ZoomPlan'));
  SetControlVisible('ZoomEcheance1', GetControlVisible('ZoomEcheance'));
  SetControlVisible('SeparPlan', GetControlVisible('ZoomPlan1') or
    GetControlVisible('ZoomEcheance1'));
  SetControlVisible('ZoomOperations1', GetControlVisible('ZoomOperations'));
  SetControlVisible('SeparOpe', GetControlVisible('ZoomOperations1'));
  SetControlVisible('ChangeEtablissement1',
    GetControlVisible('ChangeEtablissement'));
  SetControlVisible('ChangeLieu1', GetControlVisible('ChangeLieu'));
  SetControlVisible('ModifBases1', GetControlVisible('ModifBases'));
  SetControlVisible('ActionRegroupement1',
    GetControlVisible('ActionRegroupement'));
end;

procedure TOF_AMLISTE2.OnChangeEtablissementClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  if (FListe.NbSelected <= 1) and (not FListe.AllSelected) then
    ExecuteEtablissement(stCode)
  else
    ExecuteOperationSerie(toChanEtabl);
  TFMul(ECRAN).BChercheClick(nil);
  SelectionEnregistrement(stCode);
end;

procedure TOF_AMLISTE2.OnChangeLieuClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  if (FListe.NbSelected <= 1) and (not FListe.AllSelected) then
    ExecuteLieu(stCode)
  else
    ExecuteOperationSerie(toChanLieu);
  TFMul(ECRAN).BChercheClick(nil);
  SelectionEnregistrement(stCode);
end;

procedure TOF_AMLISTE2.EnableZoomVisuSerie;
begin
  SetControlVisible('ZoomPlan', false);
  SetControlVisible('ZoomEcheance', false);
  SetControlVisible('SeparPlanVisu', false);
  SetControlVisible('ZoomOperations', false);
  SetControlVisible('Ecritures',False);
end;

procedure TOF_AMLISTE2.EnableZoomActionSerie;
var
  i: integer;
  EnabledCession: boolean;
  PremNature: string;
  bCedee, bAcquisitionAnnee: boolean;
  GroupeImmo : string;
  TListeAll  : TOB;
begin
  bAcquisitionAnnee := (GetField('I_DATEPIECEA') >= VHImmo^.Encours.Deb)
    and (GetField('I_DATEPIECEA') <= VHImmo^.Encours.Fin);
  bCedee := (GetField('I_QUANTITE') = 0);
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
  //TGA - 09/2005 - Dépréciation d'actif
  SetControlVisible('DepActif', false);
  //PGR - 09/2005 - Changement de méthode
  SetControlVisible('ChangementMethode', False);

  if not SelectionMultiple then
    exit;
  if FListe.AllSelected then
    TheMULQ.First
  else
    FListe.GotoLeBookmark(0);
  PremNature := GetField('I_NATUREIMMO');
  if not VarIsNull ( TheMulQ.FindField('I_GROUPEIMMO').AsVariant) then
    GroupeImmo := GetField('I_GROUPEIMMO')
  else GroupeImmo := '';
  SetControlVisible('Cession', True);
  SetControlVisible('Mutation', True);
  SetControlVisible('ChangeEtablissement', True);
  SetControlVisible('ChangeLieu', True);
  SetControlVisible('ActionRegroupement', True);
  if FListe.AllSelected then
  begin
    TListeAll := TOB.Create ('',nil,-1);
    try
      RecupereAllSelect ( TListeAll );
      // BTY 09/05 CRC200210 Option Dépréciation
       SetControlVisible('DepActif', PremNature='PRO');

      for i := 0 to TListeAll.Detail.Count - 1 do
      begin
        EnabledCession := ((TListeAll.Detail[i].GetValue('I_NATUREIMMO') = 'PRO') or
          (TListeAll.Detail[i].GetValue('I_NATUREIMMO') = 'FI')) and (TListeAll.Detail[i].GetValue('I_QUANTITE') <> 0);
        SetControlVisible('Cession', GetControlVisible('Cession') and
          (EnabledCession) and (not bCedee));
        if (bCedee) or bAcquisitionAnnee or ((GetControlVisible('Mutation') and
          (TListeAll.Detail[i].GetValue('I_NATUREIMMO') <> PremNature))) then
          SetControlVisible('Mutation', false);
        SetControlVisible('ChangeEtablissement',
          GetControlVisible('ChangeEtablissement') and (not bCedee) and
          (VHImmo^.EtablisCpta = TRUE));
        SetControlVisible('ChangeLieu', GetControlVisible('ChangeLieu') and (not
          bCedee));
        //PGR - 12/12/2005 Pas de regroupement si immo sortie
        //SetControlVisible('ActionRegroupement', True);
        SetControlVisible('ActionRegroupement', not bCedee);

        // BTY 09/05 CRC200210 En multisélection, option Dépréciation ouverte
        // si toutes de nature = PRO et toutes de même regroupement
        if (TListeAll.Detail[i].GetValue('I_NATUREIMMO') <> 'PRO')
        or (TListeAll.Detail[i].GetValue('I_GROUPEIMMO') <> GroupeImmo) then
           SetControlVisible('DepActif', False);

        { FQ 14403 - CA 30/09/2004 - Pour la sortie de plusieurs immos, il faut qu'elles appartiennent au même groupe }
        if TListeAll.Detail[i].GetValue('I_GROUPEIMMO')<>GroupeImmo then SetControlVisible ('Cession',False);
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

     // BTY 09/05 CRC200210 Option Dépréciation
     SetControlVisible('DepActif', PremNature='PRO');

    for i := 0 to FListe.NbSelected - 1 do
    begin
      FListe.GotoLeBookmark(i);
      // Autorisation Cession
      EnabledCession := ((GetField('I_NATUREIMMO') = 'PRO') or
        (GetField('I_NATUREIMMO') = 'FI')) and (GetField('I_QUANTITE') <> 0);
      SetControlVisible('Cession', GetControlVisible('Cession') and
        (EnabledCession) and (not bCedee));
      if (bCedee) or bAcquisitionAnnee or ((GetControlVisible('Mutation') and
        (GetField('I_NATUREIMMO') <> PremNature))) then
        SetControlVisible('Mutation', false);
      SetControlVisible('ChangeEtablissement',
        GetControlVisible('ChangeEtablissement') and (not bCedee) and
        (VHImmo^.EtablisCpta = TRUE));
      SetControlVisible('ChangeLieu', GetControlVisible('ChangeLieu') and (not
        bCedee));
      //PGR - 12/12/2005 Pas de regroupement si immo sortie
      //SetControlVisible('ActionRegroupement', True);
      SetControlVisible('ActionRegroupement', not bCedee);

      // BTY 09/05 CRC200210 En multisélection, option Dépréciation ouverte
      // si toutes de nature = PRO et toutes de même regroupement

      if (GetField('I_NATUREIMMO') <> 'PRO') or (GetField('I_GROUPEIMMO') <> GroupeImmo) then
         SetControlVisible('DepActif', False);

      { FQ 14403 - CA 30/09/2004 - Pour la sortie de plusieurs immos, il faut qu'elles appartiennent au même groupe }
      if GetField('I_GROUPEIMMO')<>GroupeImmo then SetControlVisible ('Cession',False);
    end;
    end;
  end;
end;

procedure TOF_AMLISTE2.OnRowEnterListe(Sender: TObject);
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
  if Length(sListeOpe) > 0 then
  begin
    System.Delete(sListeOpe, Length(sListeOpe) - 1, 2);
    SetControlCaption('sOperationEnCours', HM2[0] + sListeOpe);
  end
  else
    SetControlCaption('sOperationEnCours', '');
end;

procedure TOF_AMLISTE2.OnFExercice2Change(Sender: TObject);
var
  FExercice2: THValComboBox;
begin
  FExercice2 := THValComboBox(GetControl('FExercice2'));
  if FExercice2.Value <> '' then
  begin
    ImExoToDates(FExercice2.Value, THCritMaskEdit(GetControl('I_DATEPIECEA')),
      THCritMaskEdit(GetControl('I_DATEPIECEA_')));
    ImExoToDates(FExercice2.Value, THCritMaskEdit(GetControl('I_DATEAMORT')),
      THCritMaskEdit(GetControl('I_DATEAMORT_')));
  end
  else
  begin
    SetControlText('I_DATEAMORT', StDate1900);
    SetControlText('I_DATEAMORT_', StDate2099);
    SetControlText('I_DATEPIECEA', StDate1900);
    SetControlText('I_DATEPIECEA_', StDate2099);
  end;
end;

procedure TOF_AMLISTE2.OnToutesOpeClick(Sender: TObject);
begin
  fmnToutesOpe.Checked := True;
  fmnAucuneOpe.Checked := False;
  fmnSelectOpe.Checked := False;
  SetControlChecked('FiltreOpe', True);
  GereFiltreOpeEnCours(cbChecked);
end;

procedure TOF_AMLISTE2.OnSelectOpeClick(Sender: TObject);
begin
  inherited;
  fmnToutesOpe.Checked := False;
  fmnAucuneOpe.Checked := False;
  fmnSelectOpe.Checked := True;
  SetControlProperty('FiltreOpe', 'State', cbUnChecked);
  fWhereOpe := fWhereOpeAucune;
  GereFiltreOpeEnCours(cbUnchecked);
end;

procedure TOF_AMLISTE2.OnAucuneOpeClick(Sender: TObject);
begin
  inherited;
  fmnToutesOpe.Checked := False;
  fmnAucuneOpe.Checked := True;
  fmnSelectOpe.Checked := False;
  SetControlProperty('FiltreOpe', 'State', cbGrayed);
  GereFiltreOpeEnCours(cbUnchecked);
end;

procedure TOF_AMLISTE2.GereFiltreOpeEnCours(Etat: TCheckBoxState);
begin
  fWhereOpe := '';
  EnableCheckBoxOperation;
  SetControlProperty('OPEECLATEMENT', 'State', Etat);
  SetControlProperty('OPECHANGEPLAN', 'State', Etat);
  SetControlProperty('OPELIEUGEO', 'State', Etat);
  SetControlProperty('OPEETABLISSEMENT', 'State', Etat);
  SetControlProperty('OPELEVEEOPTION', 'State', Etat);
  SetControlProperty('OPECESSION', 'State', Etat);
  SetControlProperty('OPEMUTATION', 'State', Etat);
  SetControlProperty('REINTEGRATION', 'State', Etat);
  SetControlProperty('QUOTEPART', 'State', Etat);
end;

procedure TOF_AMLISTE2.OnPopZoomCreationPopup(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}
  SetControlVisible('DupliqueImmo', (not TheMulQ.EOF) and (FListe.NbSelected <=
    1));
end;

procedure TOF_AMLISTE2.OnOpeIndicateurClick(Sender: TObject);
var
  sOR, sSEP: string;
begin
  inherited;
  sSEP := ' AND (';
  sOR := ' OR';
  fWhereOpe := '';
  if (GetCheckBoxState('OPEECLATEMENT') = cbChecked) then
  begin
    fWhereOpe := fWhereOpe + sSep + ' I_OPEECLATEMENT="X"';
    sSep := sOR;
  end;
  if (GetCheckBoxState('OPECHANGEPLAN') = cbChecked) then
  begin
    fWhereOpe := fWhereOpe + sSep + ' I_OPECHANGEPLAN="X"';
    sSep := sOR;
  end;
  if (GetCheckBoxState('OPELIEUGEO') = cbChecked) then
  begin
    fWhereOpe := fWhereOpe + sSep + ' I_OPELIEUGEO="X"';
    sSep := sOR;
  end;
  if (GetCheckBoxState('OPEETABLISSEMENT') = cbChecked) then
  begin
    fWhereOpe := fWhereOpe + sSep + ' I_OPEETABLISSEMENT="X"';
    sSep := sOR;
  end;
  if (GetCheckBoxState('OPELEVEEOPTION') = cbChecked) then
  begin
    fWhereOpe := fWhereOpe + sSep + ' I_OPELEVEEOPTION="X"';
    sSep := sOR;
  end;
  if (GetCheckBoxState('OPECESSION') = cbChecked) then
  begin
    fWhereOpe := fWhereOpe + sSep + ' I_OPECESSION="X"';
    sSep := sOR;
  end;
  if (GetCheckBoxState('OPEMUTATION') = cbChecked) then
  begin
    fWhereOpe := fWhereOpe + sSep + ' I_OPEMUTATION="X"';
    sSep := sOR;
  end;
  if (GetCheckBoxState('REINTEGRATION') = cbChecked) then
  begin
    fWhereOpe := fWhereOpe + sSep + ' I_REINTEGRATION<>0.0';
    sSep := sOR;
  end;
  if (GetCheckBoxState('QUOTEPART') = cbChecked) then
  begin
    fWhereOpe := fWhereOpe + sSep + ' I_QUOTEPART<>0';
    sSep := sOR;
  end;
  if fWhereOpe <> '' then
    fWhereOpe := fWhereOpe + ')'
  else
    fWhereOpe := fWhereOpeAucune;
end;

procedure TOF_AMLISTE2.OnModifBasesClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  if (ExecuteModificationBases(GetField('I_IMMO')) = mrYes) then
  begin
    TFMul(ECRAN).BChercheClick(nil);
    SelectionEnregistrement(stCode);
  end;
end;

procedure TOF_AMLISTE2.OnFiltreNouvelleRechercheClick(Sender: Tobject);
begin
  inherited;
  SetControlText('XX_WHERE', fWhereOrig);
  fmnToutesOpe.Checked := True;
  fmnAucuneOpe.Checked := False;
  fmnSelectOpe.Checked := False;
  GereFiltreOpeEnCours(cbChecked);
end;

procedure TOF_AMLISTE2.OnFiltreSauveClick(Sender: Tobject);
begin
  SetControlProperty('FiltreOpe', 'Color', clBtnFace);
  SetControlVisible('FiltreOpe', True);
  inherited;
  SetControlVisible('FiltreOpe', False);
end;

procedure TOF_AMLISTE2.OnCompteelipsisClick(Sender: TObject);
var
  stWhere: string;
begin
  if THcritMaskEdit(Sender).Name = 'I_COMPTEIMMO' then
    stWhere := 'G_GENERAL LIKE "2%"'
  else if THcritMaskEdit(Sender).Name = 'I_COMPTELIE' then
    stWhere := 'G_GENERAL LIKE "612%" OR G_GENERAL LIKE "613%"';
  LookupList(TControl(Sender), '', 'GENERAUX', 'G_GENERAL', 'G_LIBELLE', stWhere,
    'G_GENERAL', True, 0);
end;

procedure TOF_AMLISTE2.OnZoomEcrituresClick(Sender: TObject);
var
  Nature: string;
  ListeImmo: TStrings;
begin
  ListeImmo := TStringList.Create;
  Nature := GetField('I_NATUREIMMO');
  ListeImmo.Add(GetField('I_IMMO'));
  if (Nature = 'PRO') or (Nature = 'FI') then
    IntegrationEcritures(toDotation, ListeImmo, FALSE, TRUE)
  else if (Nature = 'CB') or (Nature = 'LOC') then
    IntegrationEcritures(toEcheance, ListeImmo, FALSE, TRUE);
  ListeImmo.Free;
end;

procedure TOF_AMLISTE2.OnKeyDownFListe(Sender: TObject; var Key: Word;
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

procedure TOF_AMLISTE2.OnChangeChampLibre(Sender: TObject);
begin
  if THValComboBox(Sender).Name = 'TABLELIBRE1' then
    SetControlText('I_TABLE0', GetControlText('TABLELIBRE1'))
  else if THValComboBox(Sender).Name = 'TABLELIBRE2' then
    SetControlText('I_TABLE1', GetControlText('TABLELIBRE2'))
  else if THValComboBox(Sender).Name = 'TABLELIBRE3' then
    SetControlText('I_TABLE2', GetControlText('TABLELIBRE3'));
end;

procedure TOF_AMLISTE2.OnOrganismeelipsisClick(Sender: TObject);
begin
  inherited;
{$IFDEF SERIE1}
  LookUpList(TControl(Sender), TraduireMemoire('Auxiliaire'), 'TIERS',
    'T_AUXILIAIRE', 'T_LIBELLE', 'T_NATUREAUXI="FOU"', 'T_AUXILIAIRE', True, 1);
{$ELSE}
  LookUpList(TControl(Sender), TraduireMemoire('Auxiliaire'), 'TIERS',
    'T_AUXILIAIRE', 'T_LIBELLE', 'T_NATUREAUXI="FOU"', 'T_AUXILIAIRE', True, 2);
{$ENDIF}
end;

procedure TOF_AMLISTE2.OnReinitialiseClick(Sender: TObject);
var
  stCode: string;
  i: integer;
begin
  stCode := GetField('I_IMMO');
  if (FListe.AllSelected) or (FListe.NbSelected >= 1) then
  begin
    if
      (PGIAsk('Attention : cette opération va réinitialiser les fiches sélectionnées.#10#13Voulez-vous continuer ?', '') = mrYes) then
    begin
      if FListe.AllSelected then
      begin
        TheMulQ.First;
        while not TheMulQ.Eof do
        begin
{$IFDEF AMORTISSEMENT}
          ReinitImmo(GetField('I_IMMO'));
          TheMulQ.Next;
{$ENDIF}
        end;
      end
      else
      begin
        for i := 0 to FListe.NbSelected - 1 do
        begin
          FListe.GotoLeBookmark(i);
{$IFDEF AMORTISSEMENT}
          ReinitImmo(GetField('I_IMMO'));
{$ENDIF}
        end;
      end;
      TFMul(ECRAN).BChercheClick(nil);
      SelectionEnregistrement(stCode);
      FListe.ClearSelected;
      PGIInfo('Traitement terminé.');
    end;
  end;
end;

procedure TOF_AMLISTE2.OnFiltresChange(Sender: TObject);
begin
  inherited;
  fmnToutesOpe.Checked := GetCheckBoxState('FiltreOpe') = cbChecked;
  fmnAucuneOpe.Checked := GetCheckBoxState('FiltreOpe') = cbGrayed;
  fmnSelectOpe.Checked := GetCheckBoxState('FiltreOpe') = cbUnchecked;
  EnableCheckBoxOperation;
end;

procedure TOF_AMLISTE2.OnBChercheClick(Sender: TObject);
begin
  SetControlText('XX_WHERE', GetControlText('XX_WHERE') + fWhereDerog);
  if (fmnAucuneOpe.Checked) then
    SetControlText('XX_WHERE', GetControlText('XX_WHERE') + fWhereOpeAucune)
  else if (fmnToutesOpe.Checked) then
    SetControlText('XX_WHERE', GetControlText('XX_WHERE'))
  else
    SetControltext('XX_WHERE', GetControlText('XX_WHERE') + fWhereOpe);
  TFMul(ECRAN).BChercheClick(nil);
  SetControlText('XX_WHERE', fWhereOrig);
end;

procedure TOF_AMLISTE2.OnKeyDownEcran(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    VK_F3:
      begin
        Key := 0;
        if Shift = [ssAlt] then
          LanceCreationEnSerie;
        TFMul(ECRAN).BChercheClick(nil);
      end;
    VK_DELETE :
      begin
        if (Shift = [ssCtrl]) and (TToolbarButton97(GetControl('bDelete')).Enabled) then
        begin
          Key := 0;
          TToolbarButton97(GetControl('bDelete')).Click;
        end;
      end;
  else
    fOnSaveKeyDownEcran(Sender, Key, Shift);
  end;
end;




function TOF_AMLISTE2.ListeVide: boolean;
begin
{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
  Result := (TheMulQ = nil) or ((TheMulQ <> nil) and (TheMulQ.Detail.Count =
    0));
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
  Result := TheMulQ.isEmpty;
{$ENDIF}
end;

procedure TOF_AMLISTE2.SelectionEnregistrement(stCode: string);
begin
  if (stCode = '') then
    exit;
{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}
  TheMulQ.Locate('I_IMMO', stCode, []);
end;

procedure TOF_AMLISTE2.OnActionRegroupementClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := GetField('I_IMMO');
  if (FListe.NbSelected <= 1) and (not FListe.AllSelected) then
    ExecuteRegroupement(stCode)
  else
    ExecuteOperationSerie(toRegroupement);
  TFMul(ECRAN).BChercheClick(nil);
  SelectionEnregistrement(stCode);
end;

function TOF_AMLISTE2.SelectionUnique: boolean;
begin
  Result := (((fListe.AllSelected) and (TheMulQ.RecordCount = 1))
    or (fListe.NbSelected = 1));
end;

function TOF_AMLISTE2.SelectionMultiple: boolean;
begin
  Result := (((fListe.AllSelected) and (TheMulQ.RecordCount > 1))
    or (fListe.NbSelected > 1));
end;

procedure TOF_AMLISTE2.AfterShow;
begin
  OnRowEnterListe(nil);
end;

procedure TOF_AMLISTE2.OnFlipSelectionGrille(Sender: TObject);
begin
  inherited;
  SetControlEnabled('BZOOMACTION', (fListe.NbSelected > 0) or (fListe.AllSelected));
end;

procedure TOF_AMLISTE2.OnSelectAllClick ( Sender : TObject );
begin
  TFMul(Ecran).bSelectAllClick ( Sender );
  OnFlipSelectionGrille(fListe);
end;
procedure TOF_AMLISTE2.RecupereAllSelect(T: TOB);
var
  stWhere : string;
  Q : TQuery;
begin
  if T = nil then exit;
  stWhere := RecupWhereCritere(TFMul(Ecran).Pages);
  Q := OpenSql ('SELECT I_IMMO,I_GROUPEIMMO,I_NATUREIMMO,I_QUANTITE FROM IMMO ' + stWhere, True);
  try
    if not Q.Eof then
      T.LoadDetailDB ('', '', '', Q, False);
  finally
    Ferme (Q);
  end;
end;



initialization
  registerclasses([TOF_AMLISTE2]);
end.




