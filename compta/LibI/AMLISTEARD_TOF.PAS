{***********UNITE*************************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 09/07/2007
Description .. : Source TOF de la FICHE : AMLISTEARD ()
Mots clefs ... : TOF;AMLISTEARD
*****************************************************************}

unit AMLISTEARD_TOF;

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
  Outils,
  ImEnt,
  ImPlan,
  PlanAmor,
  IntegEcr,
  ImOutGen,
  AmARDDotGlobale_TOF,
  IMMO_TOM
{$IFDEF SERIE1}
  , S1Util, utModules
{$ELSE}
  ,Ent1
{$ENDIF}
  ;

type
  TOF_AMLISTEARD = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
  private
    { Contrôles }
    POPZoomAction: TPopupMenu;
    POPZoomVisu: TPopupMenu;

   {$IFDEF EAGLCLIENT}
    FListe: THGrid;
   {$ELSE}
    FListe: THDBGrid;
   {$ENDIF}

    { Variables }
    fAction: TActionFiche;
    fWhereDerog: string;
    fWhereOrig: string;
    fWhereAmortDiffere: string;
    fWhereBaseEcoTh: string;
    fWhereSansGroupe: string;
    fWhereSansOpe: string;
    fWhereOpe: string;
    fWhereBaseTPTh: string;
    fWhereReintegration : string;
    fWhereQuotePart : string;
    fWhereCheckDPI : string;
    fWhereExigibleTP: string;
    fWhereCheckPRI : string;
    fWhereCheckSBV : string;
    fOnSaveKeyDownEcran: procedure(Sender: TObject; var Key: Word; Shift:
       TShiftState) of object;

    { Méthodes }
    procedure OnFExercice2Change(Sender: TObject);
    procedure OnAmortDiffereClick(Sender: TObject);
    procedure OnBaseEcoThClick(Sender: TObject);
    procedure OnSansGroupeClick(Sender: TObject);
    procedure OnSansOpeClick(Sender: TObject);
    procedure OnOperationChange (Sender: TObject);
    procedure OnReintegrationClick(Sender: TObject);
    procedure OnQuotePartClick(Sender: TObject);
    procedure OnBaseTPThClick(Sender: TObject);
    procedure OnDPIClick(Sender: TObject);
    procedure OnPRIClick(Sender: TObject);
    procedure OnbSBVClick(Sender: TObject);
    procedure OnExigibleTPClick(Sender: TObject);
    procedure OnChangeFFiltres(Sender : TObject);
    procedure RecupereCritereXX_WHERE(Sender: TObject);
    procedure RafraichirListeImmo(stCode: string);
    procedure OnListeDblClick(Sender: TObject);
    procedure OnPopZoomActionPopup(Sender: TObject);
    procedure OnPopZoomVisuPopup(Sender: TObject);
    procedure OnZoomPlanClick(Sender: TObject);
    procedure OnZoomEcrituresClick(Sender: TObject);
    procedure OnPopupListeClick(Sender: TObject);
    procedure EnableZoomAction;
    procedure EnableZoomActionSerie;
    procedure EnableZoomVisu;
    procedure EnableZoomVisuSerie;
    procedure EnableBoutonsAction;
    procedure OnCompteelipsisClick(Sender: TObject);
    procedure OnChangeChampLibre(Sender: TObject);
    procedure OnKeyDownFListe(Sender: TObject; var Key: Word; Shift:TShiftState);
    procedure OnKeyDownEcran(Sender: TObject; var Key: Word; Shift:TShiftState);
    function ListeVide: boolean;
    function SelectionUnique: boolean;
    function SelectionMultiple: boolean;
    procedure SelectionEnregistrement(stCode: string);
    procedure AfterShow;
    procedure OnFlipSelectionGrille (Sender : TObject);
    procedure OnSelectAllClick ( Sender : TObject );
    procedure RecupereAllSelect ( T : TOB );
   {$IFDEF EAGLCLIENT}
    procedure RecupereSelect (fListe : THGrid; T: Tob);
   {$ELSE}
    procedure RecupereSelect (fListe : THDBGrid; T: Tob);
   {$ENDIF}
    procedure ApresChangeFFiltres;
    procedure OnDotationMinimaleClick(Sender: TObject);
    procedure OnDotationMaximaleClick(Sender: TObject);
    procedure OnDotationTheoriqueClick(Sender: TObject);
    procedure OnDotationGlobaleClick(Sender: TObject);
    procedure OnDotationImmoClick(Sender: TObject);
    procedure OperationSerie (TypeOpe : TypeDotation);
    function ExecuteDotation (TypeOpe : TypeDotation; CodeI : string):integer;
    function ImmosATraiter (stCode:string; TListeAll:Tob; TypeOpe : TypeDotation; var EnCoursATraiter:boolean): boolean;
    function DotationMinimale (CodeI : string) : integer;
    function DotationMaximale (CodeI : string) : integer;
    function DotationTheorique (CodeI : string) : integer;
    function DotationImmo (CodeI : string) : integer;
    function DotationGlobale (TListe:Tob):integer;
    function EstDepreciee (CodeI : string; RepriseDep : Double) : boolean;
    procedure OnRowEnterListe(Sender: TObject); overload ;
  end;

procedure AMLanceFiche_ListeImmobilisationsARD(Action: TActionFiche = taModif; Param: string = '');
function AmGetResultatExercice : double;


implementation

const
  HM: array[0..9] of string =
{0}  ('Aucune immobilisation sélectionnée',
{1}  'Dotation minimale',
{2}  'Dotation maximale',
{3}  'Dotation théorique',
{4}  'Dotation globale',
{5}  'Dotation par immobilisation',
{6}  'Gestion des ARD',
{7}  'Certaines immobilisations de la sélection ne peuvent pas être traitées.'+ #13 + #10 +
     ' L''opération portera uniquement sur l''immobilisation en cours si celle-ci l''autorise.',
{8}  'Options non disponibles pour une immobilisation ayant fait l''objet d''une dépréciation d''actif.',
{9}  'Option non disponible, certaines immobilisations de la sélection ont fait l''objet d''une dépréciation d''actif.'
     );


function AmGetResultatExercice : double;
var Resultat: double;
    stSelect, stWhere, stWhereDate : string;
    Q: TQuery;
begin
  Resultat := 0;
  stSelect := 'SELECT SUM(E_CREDIT)-SUM(E_DEBIT)' +
              ' FROM ECRITURE LEFT JOIN GENERAUX ON (E_GENERAL=G_GENERAL)';
  stWhere := ' WHERE (G_NATUREGENE="CHA" OR G_NATUREGENE="PRO") AND (E_QUALIFPIECE="N")';
  stWhereDate := ' AND E_EXERCICE="' + VHImmo^.Encours.Code + '" AND E_DATECOMPTABLE >="'
                + USDateTime (VHImmo^.Encours.Deb) +
                '" AND E_DATECOMPTABLE <="' + USDateTime (VHImmo^.Encours.Fin) + '"';

  Q := OpenSQL (stSelect + stWhere + stWhereDate, True);
  if not Q.Eof then
     Resultat := Q.Fields[0].AsFloat;
  Ferme(Q);
  Result := Resultat;
end;



procedure CreationLieuParDefaut;
var
  Q: TQuery;
  T: TOB;
begin

  Q := OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="GEO"', True);
  if (Q.Eof) then
  begin
    T := TOB.Create('CHOIXCOD', nil, -1);
    T.PutValue('CC_TYPE', 'GEO');
    T.PutValue('CC_CODE', '001');
    T.PutValue('CC_LIBELLE', 'A définir');
    T.PutValue('CC_ABREGE', 'A définir');
    T.InsertDB(nil);
    T.Free;
    AvertirTable('TILIEUGEO');
  end;
  Ferme(Q);
end;


procedure AMLanceFiche_ListeImmobilisationsARD(Action: TActionFiche = taModif; Param: string = '');
var
  stAction: string;
begin
  if (GetParamSocSecur('SO_EXOCLOIMMO','') = VHImmo^.Encours.Code) then
    stAction := 'ACTION=CONSULTATION'
  else
    stAction := 'ACTION=MODIFICATION';
  stAction := stAction + ';' + Param;

  CreationLieuParDefaut;
  AGLLanceFiche('AM', 'AMLISTEARD', '', '', stAction);
end;


procedure TOF_AMLISTEARD.OnNew;
begin
  inherited;
end;

procedure TOF_AMLISTEARD.OnDelete;
begin
  inherited;
end;

procedure TOF_AMLISTEARD.OnUpdate;
begin
  inherited;
  // Requête a été faite sur le ONLOAD => Récup valeur initiale du critère complémentaire
  SetControlText('XX_WHERE', fWhereOrig);
  // Revoir boutons au cas où liste non vide
  EnableBoutonsAction;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : EVENEMENT ONLOAD
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.OnLoad;
begin
  inherited;

  TFMul(Ecran).Q.Manuel := True;
  TFMul(Ecran).SETDBLISTE ('AMLISTIMMOARD');
  TFMul(Ecran).FNomFiltre := 'AMLISTIMMOARD';
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
procedure TOF_AMLISTEARD.OnArgument(S: string);
var
  St : string;
  okok: boolean;
begin
  inherited;

////////////////////////////////////////////////////////////////////////////
//       Découpage paramètre en entrée
////////////////////////////////////////////////////////////////////////////

  // Action
  St := ReadTokenSt(S);
  if St = 'ACTION=MODIFICATION' then  fAction := taModif
  else   fAction := taConsult;

  // Param
  //St := ReadTokenSt(S);

  // // Aide en ligne
{$IFDEF SERIE1}
  //Ecran.HelpContext :=

  if (VS1.TypeProduit = S1BNC) and (not VS1.OKModImmo) then
  begin
    SetControlVisible('bZoomVisu', false);
  end;

{$ELSE}
   // Ecran.HelpContext :=
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

  // Compléter le changement de filtre de l'ancêtre MUL
  THValComboBox(GetControl('FFILTRES')).OnChange := OnChangeFFiltres;

  // Titre de la fenêtre
  if AmGetResultatExercice > 0 then
    TFMul(Ecran).Caption := 'Gestion des ARD de l''exercice du ' +
                          DateToStr(VHImmo^.Encours.Deb) + ' au ' +
                          DateToStr(VHImmo^.Encours.Fin) + ' (bénéficiaire)'
  else
    TFMul(Ecran).Caption := 'Gestion des ARD de l''exercice du ' +
                          DateToStr(VHImmo^.Encours.Deb) + ' au ' +
                          DateToStr(VHImmo^.Encours.Fin) + ' (déficitaire)';
  UpdateCaption(ECRAN);

////////////////////////////////////////////////////////////////////////////
//                 Initialisations des zones
////////////////////////////////////////////////////////////////////////////

  // // Boutons, PopUps
  SetControlVisible('bInsert', False);
  SetControlVisible('bDelete', False);
  SetControlVisible('bSelectAll', True);
  SetControlEnabled('BZoomAction', False);
  POPZoomAction := TPopUpMenu(GetControl('POPZOOMACTION', True));
  POPZoomVisu := TPopUpMenu(GetControl('POPZOOMVISU', True));
  EnableBoutonsAction;

  // // Init Onglet AMORTISSEMENTS
  SetControlProperty('I_BASEECO', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('I_BASEECO_', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('I_METHODEECO','Plus', ' AND CO_CODE<>"NAM"');

  // // Init Onglet OPERATIONS
{$IFDEF SERIE1}
  // Exclure les items de TIOPEAMOR autres que des opérations + changement d'établissement + levée option
  SetControlProperty('OPERATION','Plus', ' AND CO_LIBRE="OPE" AND CO_CODE<>"ETA" AND CO_CODE<>"LEV AND CO_CODE<>"DPR" AND CO_CODE<>"CES"');
{$ELSE}
  // Exclure les items de TIOPEAMOR autres que des opérations et levée option
  SetControlProperty('OPERATION','Plus', ' AND CO_LIBRE="OPE" AND CO_CODE<>"LEV" AND CO_CODE<>"DPR" AND CO_CODE<>"CES"');
{$ENDIF}
  // // Init Onglet RECHERCHE
  SetControlProperty('I_MONTANTHT', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('I_MONTANTHT_', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('I_BASETAXEPRO', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('I_BASETAXEPRO_', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));

  // // Init Onglet CARACTERISTIQUES
{$IFDEF SERIE1}
  SetControlVisible('I_ETABLISSEMENT', False);
  SetControlVisible('t_ETABLISSEMENT', False);
{$ELSE}
  // Gestion restriction sur Etablissement
  ImGereEtablissement (THMultiValComboBox(GetControl('I_ETABLISSEMENT')) );
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

  // // EVENTS Onglet STANDARDS
  THEdit(GetControl('I_COMPTEIMMO')).OnelipsisClick := OnCompteelipsisClick;

  // // EVENTS Onglet COMPLEMENTS
  TCheckBox(GetControl('bAmortDiffere')).OnClick := OnAmortDiffereClick;
  THValComboBox(GetControl('FExercice2')).OnChange := OnFExercice2Change;
  // Liste des exercices à convertir en exos relatifs
  AMInitComboExercice(THValComboBox(GetControl('FExercice2')));

  // // EVENTS Onglet AMORTISSEMENTS
  TCheckBox(GetControl('bBaseEcoTh')).OnClick := OnBaseEcoThClick;

  // // EVENTS Onglet OPERATIONS
  THMultiValComboBox(GetControl('Operation')).OnChange := OnOperationChange;
  TCheckBox(GetControl('bSansOpe')).OnClick := OnSansOpeClick;
  TCheckBox(GetControl('bReintegration')).OnClick := OnReintegrationClick;
  TCheckBox(GetControl('bQuotePart')).OnClick := OnQuotePartClick;
  TCheckBox(GetControl('bDPI')).OnClick := OnDPIClick;
  TCheckBox(GetControl('bPRI')).OnClick := OnPRIClick;
  TCheckBox(GetControl('bSBV')).OnClick := OnbSBVClick;

  // // EVENTS Onglet RECHERCHE
  TCheckBox(GetControl('bBaseTPTh')).OnClick := OnBaseTPThClick;
  TCheckBox(GetControl('bExigibleTP')).OnClick := OnExigibleTPClick;

  // // EVENTS Onglet CARACTERISTIQUES
  TCheckBox(GetControl('bSansGroupe')).OnClick := OnSansGroupeClick;

  // // EVENTS Onglet TABLES LIBRES
  THValComboBox(GetControl('TABLELIBRE1')).OnChange := OnChangeChampLibre;
  THValComboBox(GetControl('TABLELIBRE2')).OnChange := OnChangeChampLibre;
  THValComboBox(GetControl('TABLELIBRE3')).OnChange := OnChangeChampLibre;

  // // EVENTS Onglet CHAMPS LIBRES
  THValComboBox(GetControl('TABLELIBRE1')).OnChange := OnChangeChampLibre;
  THValComboBox(GetControl('TABLELIBRE2')).OnChange := OnChangeChampLibre;
  THValComboBox(GetControl('TABLELIBRE3')).OnChange := OnChangeChampLibre;

  // // EVENTS Popup Action
  PopZoomAction.OnPopup := OnPopZoomActionPopup;
  TMenuItem(GetControl('DotMin')).OnClick := OnDotationMinimaleClick;
  TMenuItem(GetControl('DotMax')).OnClick := OnDotationMaximaleClick;
  TMenuItem(GetControl('DotTheorique')).OnClick := OnDotationTheoriqueClick;
  TMenuItem(GetControl('DotGlobale')).OnClick := OnDotationGlobaleClick;
  TMenuItem(GetControl('DotImmo')).OnClick := OnDotationImmoClick;

  // // EVENTS Popup Visualisation
  PopZoomVisu.OnPopup := OnPopZoomVisuPopup;
  TMenuItem(GetControl('ZoomPlan')).OnClick := OnZoomPlanClick;
  TMenuItem(GetControl('Ecritures')).OnClick := OnZoomEcrituresClick;

  // // EVENTS Popup Liste générale - F11
  TPopUpMenu(GetControl('PopupListe')).OnPopup := OnPopupListeClick;
  TMenuItem(GetControl('ZoomPlan1')).OnClick := OnZoomPlanClick;
  TMenuItem(GetControl('Ecritures1')).OnClick := OnZoomEcrituresClick;
  TMenuItem(GetControl('DotMin1')).OnClick := OnDotationMinimaleClick;
  TMenuItem(GetControl('DotMax1')).OnClick := OnDotationMaximaleClick;
  TMenuItem(GetControl('DotTheorique1')).OnClick := OnDotationTheoriqueClick;
  TMenuItem(GetControl('DotGlobale1')).OnClick := OnDotationGlobaleClick;
  TMenuItem(GetControl('DotImmo1')).OnClick := OnDotationImmoClick;

  // Fiche appelée par un verbe OLE
  MakeZoomOle(Ecran.Handle);


////////////////////////////////////////////////////////////////////////////
//         Initialisation des critères de sélection
////////////////////////////////////////////////////////////////////////////

  // Critère complémentaire XX_WHERE par défaut
  if GetControlText('XX_WHERE') = '' then
     SetControlText('XX_WHERE', 'I_ETAT<>"FER"')
  else
     SetControlText('XX_WHERE', GetControlText('XX_WHERE') + ' AND I_ETAT<>"FER" AND I_OPEDEPREC<>"X"');

 // Critères initiaux à ajouter au XX_WHERE
  fWhereOrig := GetControlText('XX_WHERE');
  OnAmortDiffereClick(nil);
  OnBaseEcoThClick(nil);
  OnSansGroupeClick(nil);
  OnSansOpeClick(nil);
  OnOperationChange(nil);
  OnReintegrationClick(nil);
  OnQuotePartClick(nil);
  OnBaseTPThClick(nil);
  OnDPIClick(nil);
  OnExigibleTPClick(nil);
  OnPRIClick(nil);
  OnbSBVClick(nil);
end;

procedure TOF_AMLISTEARD.OnClose;
begin
  TheMulQ := nil;
  inherited;
end;

procedure TOF_AMLISTEARD.OnDisplay();
begin
  inherited;
  EnableBoutonsAction;
  if (not ListeVide) then
      OnRowEnterListe(nil);
  { Table libre 9 utilisée pour le stockage 'éligible TP' }
  SetControlVisible ('I_TABLE9',False);
  SetControlVisible ('TI_TABLE9',False);
end;

procedure TOF_AMLISTEARD.OnCancel();
begin
  inherited;
end;

procedure TOF_AMLISTEARD.AfterShow;
begin
  OnRowEnterListe(nil);
end;

procedure TOF_AMLISTEARD.EnableBoutonsAction;
begin
  SetControlVisible('BZoomAction', (fAction <> taConsult) and (not ListeVide));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Clic sur une immo ou défilement vertical dans la liste
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.OnRowEnterListe(Sender: TObject);
begin
  exit;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/07/2007
Modifié le ... :   /  /
Description .. : Test immo dépréciée :
Suite ........ : - dans un exerice clos
Suite ........ : - dans l'exercice en cours
Suite ........ : - en saisie antérieurs dépréciation
Mots clefs ... :
*****************************************************************}
function TOF_AMLISTEARD.EstDepreciee (CodeI : string; RepriseDep : Double) : boolean;
begin
// I_REVISIONECO pas accessible directement (absent de AMLISTIMMOARD)
result :=  (ExisteSQL ('SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="'+
            CodeI +'" AND IL_TYPEOP="CLO" AND IL_REVISIONECO<>0 AND IL_DATEOP = "'+
            UsDateTime((VHImmo^.EnCours.deb)-1)+'" '))
            or
           (ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+
            CodeI +'" AND (I_REVISIONECO<>0)'))
            or
           (RepriseDep <> 0);
end;



{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/07/2007
Modifié le ... :   /  /
Description .. : Ouverture/Fermeture des options du menu ACTION
Suite ........ : pour une seule immobilisation sélectionnée
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.EnableZoomAction;
var  stCode : string;
     bDeprec : boolean;

begin
  SetControlVisible('DotMin', false);
  SetControlVisible('DotMax', false);
  SetControlVisible('DotTheorique', false);
  SetControlVisible('DotGlobale', False);
  SetControlVisible('DotImmo', False);
  SetControlVisible('Ecritures1', SelectionUnique);

  if ListeVide or (not SelectionUnique) then
     exit;

  Stcode := GetField('I_IMMO');

  // Dotations fermées si immo dépréciée
  bDeprec := EstDepreciee (Stcode, GetField('I_REPRISEDEP'));
  SetControlVisible('Dotmin', (not bDeprec));
  SetControlVisible('Dotmax', (not bDeprec));
  SetControlVisible('Dottheorique', (not bDeprec));
  SetControlVisible('DotImmo', (not bDeprec));
  SetControlVisible('Dotglobale', False);
  // Non disponibles après une dépréciation
  if bDeprec then
     PGIBox(HM[8], HM[6]);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/07/2007
Modifié le ... :   /  /
Description .. : Ouverture/Fermeture des options du menu ACTION
Suite ........ : pour plusieurs immobilisations sélectionnées
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.EnableZoomActionSerie;
var  i: integer;
     bDeprec, bDeprecSelectee : boolean;
     TListeAll  : TOB;

begin
  SetControlVisible('Ecritures', false);
  SetControlVisible('Ecritures1', false);

  if not SelectionMultiple then
     exit;

  // Positionner le curseur sur la 1e immo
  if FListe.AllSelected then
     TheMULQ.First
  else
     FListe.GotoLeBookmark(0);

  // Test immo sous le curseur dépréciée
  bDeprec := EstDepreciee (GetField('I_IMMO'), GetField('I_REPRISEDEP'));

  // Init options pour l'immo sous le curseur
  SetControlVisible('DotGlobale', (not bDeprec));
  SetControlVisible('DotMin', (not bDeprec));
  SetControlVisible('DotMax', (not bDeprec));
  SetControlVisible('DotTheorique', (not bDeprec));
  SetControlVisible('DotImmo', (not bDeprec));

  // Toutes les immos sont sélectionnées
  if FListe.AllSelected then
  begin
    TListeAll := TOB.Create ('',nil,-1);
    try
      RecupereAllSelect ( TListeAll );

      for i := 0 to TListeAll.Detail.Count - 1 do
      begin
        bDeprecSelectee := EstDepreciee (TListeAll.Detail[i].GetValue('I_IMMO'),
                                         TListeAll.Detail[i].GetValue('I_REPRISEDEP'));

        // Dotations ouvertes si au moins une immo non dépréciée
        if (not bDeprecSelectee) then
        begin
           SetControlVisible('DotMin', True);
           SetControlVisible('DotMax', True);
           SetControlVisible('DotTheorique', True);
           SetControlVisible('DotGlobale', True);
           SetControlVisible('DotImmo', True);
        end;
      end;
    finally
      FreeAndNil (TListeAll);
    end;
  end
  else
  begin
    if FListe.NbSelected - 1 >0 then
    begin

      for i := 0 to FListe.NbSelected - 1 do
      begin
        FListe.GotoLeBookmark(i);

        bDeprecSelectee := EstDepreciee (GetField('I_IMMO'), GetField('I_REPRISEDEP'));

        // Dotations ouvertes si au moins une immo non dépréciée
        if (not bDeprecSelectee) then
        begin
           SetControlVisible('DotMin', True);
           SetControlVisible('DotMax', True);
           SetControlVisible('DotTheorique', True);
           SetControlVisible('DotGlobale', True);
           SetControlVisible('DotImmo', True);
        end;
      end;
    end; // nbselected
  end;
end;

procedure TOF_AMLISTEARD.EnableZoomVisu;
begin
  inherited;
  SetControlVisible('ZoomPlan', False);
  SetControlVisible('Ecritures',False);

  if ListeVide then
     exit;
  SetControlVisible('ZoomPlan', True);
  SetControlVisible('SeparPlanVisu', True);

  SetControlVisible('Ecritures', (SelectionUnique) );
end;


procedure TOF_AMLISTEARD.EnableZoomVisuSerie;
begin
  SetControlVisible('ZoomPlan', false);
  SetControlVisible('SeparPlanVisu', false);
  SetControlVisible('Ecritures',False);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/07/2007
Modifié le ... :   /  /
Description .. : ACTION Dotation minimale
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.OnDotationMinimaleClick(Sender: TObject);
begin
  OperationSerie (toDotMin);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/07/2007
Modifié le ... :   /  /
Description .. : ACTION Dotation maximale
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.OnDotationMaximaleClick(Sender: TObject);
begin
  OperationSerie (toDotMax);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/07/2007
Modifié le ... :   /  /
Description .. : ACTION Dotation théorique
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.OnDotationTheoriqueClick(Sender: TObject);
begin
  OperationSerie (toDotTheorique);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/07/2007
Modifié le ... :   /  /
Description .. : ACTION Dotation globale
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.OnDotationGlobaleClick(Sender: TObject);
begin
  OperationSerie (toDotGlobale);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/07/2007
Modifié le ... :   /  /
Description .. : ACTION Dotation par immobilisation
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.OnDotationImmoClick(Sender: TObject);
begin
  OperationSerie (toDotImmo);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/07/2007
Modifié le ... :   /  /
Description .. : Application des dotations ouverte en série si aucune immo dépréciée
Suite ........ : Sinon, message et traitement uniquement de l'immo courante (si pas dotation globale)
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.OperationSerie (TypeOpe : TypeDotation);

var
  i : integer;
  stCode, Titre : string;
  TListeAll, TListeSel : TOB;
  mr: integer;
  EnCoursTraitable : boolean;
begin

  stCode := GetField('I_IMMO');

  case TypeOpe of
       toDotMin : Titre := HM[1];
       toDotMax : Titre := HM[2];
       toDotTheorique : Titre := HM[3];
       toDotGlobale : Titre := HM[4];
       toDotImmo : Titre := HM[5];
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
            begin
               // Dotation globale appliquée en une fois
               if TypeOpe = toDotGlobale then
               begin
                   DotationGlobale (TListeAll);
               end
               else
               begin
                   // On boucle sur les opérations individuelles
                   for i := 0 to TListeAll.Detail.Count - 1 do
                   begin
                      mr := ExecuteDotation (TypeOpe, TListeAll.Detail[i].GetValue('I_IMMO'));
                      if   mr <> mrYes then break;
                   end;
               end;
            end
            else
            begin
               // Certaines immos sélectionnées ne sont pas traitables
               // Dotation globale à fermer
               if TypeOpe = toDotGlobale then
                 PGIBox(HM[9], Titre)
               else
               begin
                 PGIBox(HM[7], Titre);
                 // Si immo courante traitable, faire l'opération sur elle seule
                 if EnCoursTraitable then
                 begin
                   ExecuteDotation(TypeOpe, TListeAll.Detail[0].GetValue('I_IMMO'));
                 end;
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
      TListeSel := TOB.Create ('', nil, -1);
      try
        // Mettre la sélection en TOB
        RecupereSelect (fListe, TListeSel);

        // On se positionne sur un enregistrement de la sélection
        if (fListe.NbSelected >= 1) then fListe.GotoLeBookmark(0);

        if ImmosATraiter (stCode, nil, TypeOpe, EnCoursTraitable) then
        // Toutes les immos sélectionnées sont traitables
        begin
          // Dotation globale appliquée en une fois
          if TypeOpe = toDotGlobale then
          begin
            DotationGlobale (TListeSel);
          end
          else
          begin
            // On boucle sur les opérations individuelles
            for i := 0 to fListe.NbSelected - 1 do
            begin
              fListe.GotoLeBookmark(i);
              mr := ExecuteDotation(TypeOpe, GetField('I_IMMO'));
              if mr <> mrYes then break;
            end;
          end;
        end
        else
        begin
          // Certaines immos sélectionnées ne sont pas traitables
          // Dotation globale à fermer
          if TypeOpe = toDotGlobale then
            PGIBox(HM[9], Titre)
          else
          begin
            PGIBox(HM[7], Titre);
            // Si immo courante traitable, faire l'opération sur elle seule
            if EnCoursTraitable then
              begin
              ExecuteDotation(TypeOpe, GetField('I_IMMO'));
              end;
          end;
        end;
        fListe.ClearSelected;
      finally
        FreeAndNil (TListeSel);
      end;
  end;

  // Toujours rafraichir, GotoLeBookmark repositionne => immo courante perdue
  RafraichirListeImmo (stCode);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/07/2007
Modifié le ... :   /  /
Description .. : Branchement sur le calcul de dotation demandé
Mots clefs ... :
*****************************************************************}
function TOF_AMLISTEARD.ExecuteDotation (TypeOpe : TypeDotation; CodeI : string):integer;
begin
     result := mrYes;

     case TypeOpe of
     toDotMin : begin
               result := DotationMinimale (CodeI);
               end;
     toDotMax : begin
               result := DotationMaximale (CodeI);
               end;
     toDotTheorique : begin
               result := DotationTheorique (CodeI);
               end;
     toDotImmo : begin
               result := DotationImmo (CodeI);
               end;
     end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/07/2007
Modifié le ... : 17/07/2007
Description .. : Branchement sur la saisie de la dotation globale
Suite ........ : NOTA :
Suite ........ : TListe est soit issue d'un opensql sur la table immo
Suite ........ : soit alimentée à la main à partir des immos sélectionnées
Mots clefs ... :
*****************************************************************}
function TOF_AMLISTEARD.DotationGlobale (TListe:Tob):integer;
begin
   result := mrYes;

   // Stocker la Tob de la sélection dans la TOB générale TheTob accessible partout
   TheTob := TListe;
   if AMLanceFiche_ARDDotGlobale = mrYes then
   begin
   // ATTENTION : Ici TheTob remise à nil et n'a pas été màj dans AMLanceFiche_ARDDotGlobale
   // => laisser faire OperationSerie qui rafraichit la MUL par BChercheClick
   // à condition que les champs i_amtfor, etc. aient été ENREGISTRES EN BASE par AMLanceFiche_ARDDotGlobale
   end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/07/2007
Modifié le ... :
Description .. : Application de la dotation minimale à l'immo CodeI
Mots clefs ... :
*****************************************************************}
function TOF_AMLISTEARD.DotationMinimale (CodeI : string) : integer;
begin
  result := mrYes;
  PGIBox('Montant à appliquer à l''immo ' + CodeI, HM[1]);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/07/2007
Modifié le ... :
Description .. : Application de la dotation maximale à l'immo CodeI
Mots clefs ... :
*****************************************************************}
function TOF_AMLISTEARD.DotationMaximale (CodeI : string) : integer;
begin
  result := mrYes;
  PGIBox('Montant à appliquer à l''immo ' + CodeI, HM[2]);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/07/2007
Modifié le ... :
Description .. : Application de la dotation théorique à l'immo CodeI
Mots clefs ... :
*****************************************************************}
function TOF_AMLISTEARD.DotationTheorique (CodeI : string) : integer;
begin
  result := mrYes;
  PGIBox('Appliquer le calcul théorique à l''immo ' + CodeI, HM[3]);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/07/2007
Modifié le ... :
Description .. : Appel fenêtre de saisie d'une dotation applicable à une immo
Mots clefs ... :
*****************************************************************}
function TOF_AMLISTEARD.DotationImmo (CodeI : string) : integer;
begin
  result := mrYes;
  PGIBox('Appel fenêtre de saisie dotation par immo : immo ' + CodeI, HM[5]);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/07/2007
Modifié le ... :
Description .. : On vérifie que toutes les immos de la sélection
Suite ........ : peuvent subir la dotation demandée, sinon on vérifie l'immo sous le curseur
Suite ........ : Test fait sur la dépréciation
Mots clefs ... :
*****************************************************************}
function TOF_AMLISTEARD.ImmosATraiter (stCode:string; TListeAll:Tob; TypeOpe : TypeDotation;
         var EnCoursATraiter : boolean): boolean;
var i : integer;
    bDeprec : boolean;
begin
  result := True;
  EnCoursATraiter := result;

  if FListe.AllSelected then
  begin

    // Chercher dans la sélection si existe une immo ne pouvant pas subir TypeOpe
    for i := 0 to TListeAll.Detail.Count - 1 do
    begin
      bDeprec := EstDepreciee (TListeAll.Detail[i].GetValue('I_IMMO'), 
                               TListeAll.Detail[i].GetValue('I_REPRISEDEP') );
      if bDeprec then
         result := false;

      // Récup situation de l'immo courante
      if (TListeAll.Detail[i].GetValue('I_IMMO') = stCode) then
         EnCoursATraiter := (not bDeprec);
    end;

  end
  else
  begin

    // Chercher dans la sélection si existe une immo ne pouvant pas subir TypeOpe
    for i := 0 to FListe.NbSelected - 1 do
    begin
      FListe.GotoLeBookmark(i);

      bDeprec := EstDepreciee (GetField('I_IMMO'), GetField('I_REPRISEDEP'));
      if bDeprec then
         result := false;

      // Récup situation de l'immo courante
      if (GetField('I_IMMO') = stCode) then
         EnCoursATraiter := (not bDeprec);
    end;
  end;
end;



function TOF_AMLISTEARD.ListeVide: boolean;
begin
{$IFDEF EAGLCLIENT}
   TheMulQ := TFMul(Ecran).Q.TQ;
   Result := (TheMulQ = nil) or ((TheMulQ <> nil) and (TheMulQ.Detail.Count = 0));
{$ELSE}
   TheMulQ := TFMul(Ecran).Q;
   Result := TheMulQ.isEmpty;
{$ENDIF}
end;

procedure TOF_AMLISTEARD.OnKeyDownEcran(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  fOnSaveKeyDownEcran(Sender, Key, Shift);
end;

procedure TOF_AMLISTEARD.OnKeyDownFListe(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin

  case Key of
    VK_F11:
      begin
        Key := 0;
        if (fAction <> taConsult) then
           TPopupMenu(GetControl('PopupListe')).Popup(Mouse.CursorPos.x,Mouse.CursorPos.y);
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
procedure TOF_AMLISTEARD.RecupereCritereXX_WHERE(Sender: TObject);
begin
  // onglet Compléments
  SetControltext('XX_WHERE', GetControlText('XX_WHERE') + fWhereAmortDiffere);
  // onglet Amortissements
  SetControltext('XX_WHERE', GetControlText('XX_WHERE') + fWhereBaseEcoTh + fWhereDerog);
  // onglet Opérations
  SetControltext('XX_WHERE', GetControlText('XX_WHERE') + fWhereSansOpe + fWhereOpe
                             + fWhereReintegration + fWhereQuotePart + fWhereCheckDPI
                             + fWhereCheckPRI + fWhereCheckSBV);
  // onglet Recherche
  SetControltext('XX_WHERE', GetControlText('XX_WHERE') + fWhereBaseTPTh +
                             fWhereExigibleTP);
  // onglet Caractéristiques
  SetControltext('XX_WHERE', GetControlText('XX_WHERE') + fWhereSansGroupe);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Touche ESPACE de sélection d'une immo
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.OnFlipSelectionGrille(Sender: TObject);
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
procedure TOF_AMLISTEARD.OnListeDblClick(Sender: TObject);
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
    AMLanceFiche_FicheImmobilisation(stCode, taConsult, '', '');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Composition du menu F11 en fonction de l'immo courante
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.OnPopupListeClick(Sender: TObject);
begin
  OnPopZoomVisuPopup(Sender);
  OnPopZoomActionPopup(Sender);
  SetControlVisible('ZoomPlan1', GetControlVisible('ZoomPlan'));
  SetControlVisible('N31', GetControlVisible('ZoomPlan1'));
  SetControlVisible('DotMin1', GetControlVisible('DotMin'));
  SetControlVisible('DotMax1', GetControlVisible('DotMax'));
  SetControlVisible('DotTheorique1', GetControlVisible('DotTheorique'));
  SetControlVisible('DotGlobale1', GetControlVisible('DotGlobale'));
  SetControlVisible('DotImmo1', GetControlVisible('DotImmo'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Composition du menu ACTION en fonction de la sélection
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.OnPopZoomActionPopup(Sender: TObject);
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
Description .. : Composition du menu VISU en fonction de l'immo courante
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.OnPopZoomVisuPopup(Sender: TObject);
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
Description .. : CTL+A Sélection de toute la liste
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.OnSelectAllClick(Sender: TObject);
begin
  TFMul(Ecran).bSelectAllClick ( Sender );
  OnFlipSelectionGrille(fListe);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : VISU Ecritures de l'immo
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.OnZoomEcrituresClick(Sender: TObject);
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
Description .. : VISU Plan d'amortissement
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.OnZoomPlanClick(Sender: TObject);
var
  QTmp: TQuery;
  PlanAmor: TPlanAmort;
begin
  inherited;
  QTmp := OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="' + GetField('I_IMMO') +
    '"', true);
  if not QTmp.Eof then
  begin
    PlanAmor := TPlanAmort.Create(true);
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
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Charger en TOB toutes les immos de la multisélection
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.RecupereAllSelect(T: TOB);
var
 stWhere : string;
 Q : TQuery;
begin
  if T = nil then exit;

  // RecupWhereCritere reprend le XX_WHERE par défaut => compléter par le filtre en cours
  stWhere := RecupWhereCritere(TFMul(Ecran).Pages);
  RecupereCritereXX_WHERE (nil);
  Q := OpenSql (
  'SELECT I_IMMO,I_NATUREIMMO, I_DATEPIECEA, I_REVISIONECO, I_METHODEECO, I_MONTANTEXC, ' +
         'I_REPRISEDEP, I_DOTMIN, I_DOTMAX, I_DOTTHEORIQUE, I_AMTFOR, I_MONTANTHT FROM IMMO ' +
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
Description .. : Charger en TOB les immos sélectionnées par Barre Espace
Mots clefs ... :
*****************************************************************}
{$IFDEF EAGLCLIENT}
procedure TOF_AMLISTEARD.RecupereSelect (fListe : THGrid; T: Tob);
{$ELSE}
procedure TOF_AMLISTEARD.RecupereSelect (fListe : THDBGrid; T: Tob);
{$ENDIF}
var i : integer;
    TFille : Tob;
begin

   // On boucle sur les immos sélectionnées
   for i := 0 to fListe.NbSelected - 1 do
   begin
      fListe.GotoLeBookmark(i);
      // ajouter une TOB fille à T
      TFille := TOB.Create ('', T, -1);
{$IFDEF EAGLCLIENT}
      // ATTENTION Apparemment on ne peut pas filtrer les colonnes (I_LIBELLE inutile)
      // => Peut-être faire AddChampSupValeur comme en 2 tiers
      TFille.GetLigneGrid (fListe, i+1, 'I_IMMO;I_LIBELLE;I_COMPTEIMMO;I_DOTMIN;I_DOTMAX;I_DOTTHEORIQUE;I_AMTFOR;I_MONTANTHT');
{$ELSE}
      TFille.AddChampSupValeur('I_IMMO', GetField ('I_IMMO'));
      TFille.AddChampSupValeur('I_NATUREIMMO', GetField('I_NATUREIMMO'));
      TFille.AddChampSupValeur('I_DOTMIN', GetField ('I_DOTMIN'));
      TFille.AddChampSupValeur('I_DOTMAX', GetField ('I_DOTMAX'));
      TFille.AddChampSupValeur('I_DOTTHEORIQUE', GetField ('I_DOTTHEORIQUE'));
      TFille.AddChampSupValeur('I_AMTFOR', GetField ('I_AMTFOR'));
      TFille.AddChampSupValeur('I_MONTANTHT', GetField ('I_MONTANTHT'));
{$ENDIF}
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
procedure TOF_AMLISTEARD.SelectionEnregistrement(stCode: string);
begin
{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}
  if (stCode = '') then exit;

  TheMulQ.Locate('I_IMMO', stCode, []);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Teste si plusieurs immos sélectionnées
Mots clefs ... :
*****************************************************************}
function TOF_AMLISTEARD.SelectionMultiple: boolean;
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
function TOF_AMLISTEARD.SelectionUnique: boolean;
begin
  Result := (((fListe.AllSelected) and (TheMulQ.RecordCount = 1))
             or (fListe.NbSelected = 1));
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... :   /  /
Description .. : Rafraichir la liste des immos après une action,
Suite ........ : ceci conformément à la sélection des immos
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.RafraichirListeImmo(stCode: string);

begin
  // Relancer la requête de multisélection
  // ATTENTION fait perdre le FOCUS mais OBLIGATOIRE pour réactualiser les colonnes affichées
  // FOCUS conservé si pas de totaux affichés en cumul en bas de page
  TFMul(ECRAN).BChercheClick(nil);

  // Se repositionner sur l'enreg courant
  SelectionEnregistrement(stCode);

  // Remettre le FOCUS
  fListe.Setfocus;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 11/01/2006
Modifié le ... : 11/01/2006
Description .. : ONGLET STANDARDS
Suite ........ : Liste des comptes en Lookup sur un champ Compte
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.OnCompteelipsisClick(Sender: TObject);
var
  stWhere: string;
begin
  stWhere := 'G_GENERAL LIKE "2%"';
  LookupList(TControl(Sender), '', 'GENERAUX', 'G_GENERAL', 'G_LIBELLE', stWhere,
             'G_GENERAL', True, 0);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 29/06/2006
Modifié le ... :
Description .. : Action de changement du filtre (complète celui du MUL)
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.OnChangeFFiltres(Sender: TObject);
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
procedure TOF_AMLISTEARD.OnFExercice2Change(Sender: TObject);
var
  FExercice2: THValComboBox;
begin
  FExercice2 := THValComboBox(GetControl('FExercice2'));

  if FExercice2.Value <> '' then
    begin
    // Exercice Relatif à convertir en dates réelles
    AMAppliquerExoRelatifToDates ( FExercice2.Value, THEdit(GetControl('I_DATEDEBECO')),
                                  THEdit(GetControl('I_DATEDEBECO_')));
    AMAppliquerExoRelatifToDates ( FExercice2.Value, THEdit(GetControl('I_DATEPIECEA')),
                                  THEdit(GetControl('I_DATEPIECEA_')));
    end
  else
    begin
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
procedure TOF_AMLISTEARD.OnAmortDiffereClick(Sender: TObject);
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
procedure TOF_AMLISTEARD.OnBaseEcoThClick(Sender: TObject);
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
procedure TOF_AMLISTEARD.OnOperationChange(Sender: TObject);
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
procedure TOF_AMLISTEARD.OnSansOpeclick(Sender: TObject);
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
procedure TOF_AMLISTEARD.OnReintegrationClick(Sender: TObject);
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
procedure TOF_AMLISTEARD.OnQuotePartClick(Sender: TObject);
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
procedure TOF_AMLISTEARD.OnDPIclick(Sender: TObject);
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
procedure TOF_AMLISTEARD.OnPRIclick(Sender: TObject);
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
procedure TOF_AMLISTEARD.OnbSBVclick(Sender: TObject);
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
procedure TOF_AMLISTEARD.OnBaseTPThClick(Sender: TObject);
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
procedure TOF_AMLISTEARD.OnExigibleTPClick(Sender: TObject);
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
procedure TOF_AMLISTEARD.OnSansGroupeClick(Sender: TObject);
begin

     case GetCheckBoxState('bSansGroupe') of
     cbChecked :
       begin
       fWhereSansGroupe := ' AND I_GROUPEIMMO = ""';
       SetControlEnabled('I_GROUPEIMMO', False);
       SetControlEnabled('tI_GROUPEIMMO', False);
       THMultiValComboBox(GetControl('I_GROUPEIMMO')).Value := '';
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
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 13/09/2006
Modifié le ... :
Description .. : ONGLET CARACERISTIQUES
Suite ........ : Si un filtre restaure un établissement qui ne correspond pas
Suite ........ : aux restrictions utilisateur, alimenter la COMBO Etablissements
Suite ........ : selon les restrictions et pas selon le filtre
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEARD.ApresChangeFFiltres;
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
procedure TOF_AMLISTEARD.OnChangeChampLibre(Sender: TObject);
begin
  if THValComboBox(Sender).Name = 'TABLELIBRE1' then
     SetControlText('I_TABLE0', GetControlText('TABLELIBRE1'))
  else if THValComboBox(Sender).Name = 'TABLELIBRE2' then
     SetControlText('I_TABLE1', GetControlText('TABLELIBRE2'))
  else if THValComboBox(Sender).Name = 'TABLELIBRE3' then
     SetControlText('I_TABLE2', GetControlText('TABLELIBRE3'));
end;




initialization
  registerclasses([TOF_AMLISTEARD]);
end.




