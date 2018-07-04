{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 13/08/2002
Modifié le ... :   /  /
Description .. : Unit de définition de fonction génériques de la paie
Mots clefs ... : PAIE
*****************************************************************
PT1   | 16/12/2003 | JL | 50   | Ajout frais animateurs dans calcul investissement
PT2   | 25/01/2005 | JL | 60   | Modifs calcul total HT pour passerelle
      | 20/03/2006 | JL |      | Modification clé annuaire ----
      | 17/10/2006 | JL |      | Modification contrôle des exercices de formations -----
PT3   | 13/04/2007 | FL | 72   | FQ 13568 Prise en compte des paramètres prévisionnels
      |            |    |      |  pour le calcul du coût salarié
PT4   | 17/04/2007 | FL | 72   | FQ 14071 Pour le calcul du taux horaire salarié,
      |            |    |      |  vérifier si, au moment de la formation, il n'existe
      |            |    |      |  pas un taux historisé postérieur à la fin de la
      |            |    |      |  formation. Si oui, appliquer celui-ci.
PT5   | 15/05/2007 | FL | 72   | FQ 13567 Ajout du type de population pour les
      |            |    |      |  formations au prévisionnel
PT6   | 16/05/2007 | FL | 72   | FQ 11532 Mise à jour des plafonds
PT7   | 07/06/2007 | FL | 72   | Ajout de constantes pour la gestion des
      |            |    |      |  multi-sessions
PT8   | 13/07/2007 | VG | 72   | "Condition d''emploi" remplacé par "Caractéristique
      |            |    |      |  activité" - FQ N°14568
PT9   | 14/11/2007 | FL | 8.01 | Ajout de IFDEF PLUGIN pour l'utilisation allégée du source par le portail
PT10  | 28/09/2007 | FL | 7    | Emanager / Report / Adaptation du SQL pour assistant ou responsable
PT11  | 18/01/2008 | FL | 8.02 | Ajout de la fonction CalCoutPedagogique et mise à disposition des fonctions
                                 RecherchePopulation,CalcCtInvestSession,RendMillesimeRealise pour le portail
PT12  | 15/02/2008 | FL | 8.03 | Ajout de la fonction DossiersAInterroger pour le multidossier
PT13B  | 19/02/2008 | FL | 8.03 | Ajout de la fonction ElipsisSalarieMultidos
PT14  | 21/02/2008 | FL | 8.03 | Ajout de la fonction MAJResponsable pour le mise à jour des tables référençant
      |            |    |      |  les responsables (BUDGETPAIE, INSCFORMATION, FORMATIONS)
PT15  | 03/03/2008 | FL | 8.03 | Emanager / Généralisation de la fonction AdaptByResp à d'autres suffixes
PT16  | 14/03/2008 | FL | 8.03 | Ajout des fonctions AdaptMultiDosForm et GetNoDossierSalarie
PT17  | 18/03/2008 | FL | 8.03 | Modification de la fonction DossiersAInterroger pour ajouter le prédéfini au besoin (portail)
								+ Adaptations pour connexion sur la DB000000 + utilisation du GetBase
PT18  | 25/03/2008 | FL | 8.03 | Report de la version améliorée de DossiersAInterroger
PT19  | 02/04/2008 | FL | 8.03 | Gestion des groupes de travail
PT20  | 03/04/2008 | FL | 8.03 | Modification du mail envoyé
PT21  | 02/04/2008 | FL | 8.03 | Adaptation du partage formation à l'entreprise
PT22  | 07/04/2008 | FL | 8.03 | Correction de la récupération des infos des dossiers auxquels l'utilisateur a accès si pas de critère spécifié
PT23  | 16/04/2008 | FL | 8.04 | Création d'une variable pour l'activation uniquement du bundle Catalogue de formation + renommage PGAccesMultiForm en PGBundleInscFormation
PT24  | 24/04/2008 | FL | 8.04 | La liste des dossiers à intérroger ne doit pas prendre en compte le prédéfini s'il s'agit de la liste des intérimaires
PT25  | 30/04/2008 | FL | 8.04 | Adaptation pour le portail qui ne connait pas VH_PAIE
PT26  | 07/05/2008 | FL | 8.04 | Pour les mails envoyés dans le cadre du DIF, ne pas mettre les compteurs si la paie n'est pas gérée
PT27  | 15/05/2008 | FL | 8.04 | FQ 15371 Ajout du corps du mail en cas de sortie d'un salarié
PT28  | 28/05/2008 | FL | 8.04 | FQ 14052 Déplacement de fonctions depuis UTOMFormations pour la gestion du DIF
PT29  | 30/05/2008 | FL | 8.04 | Ajout du mail en cas d'annulation de DIF
PT30  | 02/06/2008 | FL | 8.04 | Passage en HTML des mails en provenance de PGI
PT31  | 04/06/2008 | FL | 8.04 | FQ 15458 Report V7 Ajout d'une fonction pour ne pas voir les salariés confidentiels
PT32  | 17/06/2008 | FL | 8.04 | Mise à jour des coûts au réalisé uniquement si formation effectuée
PT33  | 09/07/2008 | FL | 8.04 | Suivi plan d'intégration
}

unit PgoutilsFormation;

interface
uses StdCtrls, Controls, Classes, HSysMenu, HMsGBox,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} {$IFNDEF PLUGIN}Fe_Main,{$ENDIF}
{$ELSE}
  {$IFDEF PLUGIN}MaineAGL ,{$ENDIF}
{$ENDIF}
  sysutils, UTob, HCtrls, HEnt1, StrUtils
{$IFDEF PLUGIN}
  , ParamSoc,PGOutils ;
{$ELSE}
, EntPAie,  ParamSoc,PGOutils,PGCommun,
  SaisieList,UTableFiltre,HStatus, PGOutils2, PGPopulOutils, PgCalendrier,LookUp,UtilPGI; //PT6
{$ENDIF}

const ValPrevisionnel = True;
const ValReel = False;
Const TYP_POPUL_FORM_PREV = 'FOR'; //PT5
Const TYP_NO_MULTISESSION = 'AUC'; //PT7
Const TYP_ENTETE = 'ENT'; //PT7
Const TYP_SOUSSESSION = 'SOS'; //PT7

//PT21
Const BUNDLE_FORMATION  = 1; 
Const BUNDLE_HIERARCHIE = 2;
Const BUNDLE_CATALOGUE  = 3; //PT23

Const USER_MAIL_FORMATION = '~ÌÞ'; //PT21

{$IFDEF EMANAGER}
Const CYCLE_EN_COURS_EM = '12'; //PT10
{$ENDIF}

{$IFNDEF PLUGIN}

  Function  AfficheLibelle(Func, Params, WhereSQL: string; TT: TDataset; Total: Boolean): string;
  procedure VisibiliteChampFormation(Numero: string; Champ1, Champ2: TControl); // sur les champs libres de la formation pour la gestion de l'onglet PCOMPLEMENT des fiches multicritere
  procedure VisibiliteChampCompetence(Numero: string; Champ1, Champ2: TControl); // sur les champs libres des compétences pour la gestion de l'onglet PCOMPLEMENT des fiches multicritere
  procedure VisibiliteChampCompetenceRess(Numero: string; Champ1, Champ2: TControl); // sur les champs libres des compétences ressources pour la gestion de l'onglet PCOMPLEMENT des fiches multicritere
  procedure RecupInscForReporte(Millesime: string); // Copie les inscriptions au budget reportés sur le millesime suivant
  procedure RecupInscNonForm(Millesime: string); // copie les inscrptions au budgetnon suivi par une formation sur millesime suivant
  function  VerifPlafondFormation(LeFrais, LeMillesime: string; LeMontant: Double): Double;
  function  RecupSalaireFormation(LeSalarie: string): Double;
  function  RecupSalaireFormationNonNom(LeMillesime, LeLibEmploi: string): Double;
  procedure MAJCoutsFormation(LeMillesime, LeStage, LaSession: string;TF : TTableFiltre = Nil);
  function  NombreDeSessionFormationAPrevoir(NbMax, MbInsc: Integer): Double;
  procedure ChargeTobCreerFichierForm(NumEnvoiForm: Integer);
  procedure RecupDonneesEnvoiSta(NumEnvoiForm:Integer; OPCACP, OPCACS: String; Trace, TraceErr: TListBox);
  procedure RecupDonneesEnvoiAnim(NumEnvoiForm : Integer; OPCACP, OPCACS: String; Trace, TraceErr: TListBox);
  function  CalcCtPedagParSalarie(CtForm, CtPedag, CtUnit, NbHSal, NbHTot: Double): Double;
  function  RendPlafondFraisForm(LeFrais, LeLieu, OrgCollect, LeMillesime, Population: string): Double; //PT6
  procedure SuiviCarriereGrilles(GInit, GExt, GInt, GCompetEmp, GCompetSal: THGrid);
  procedure CarriereAfficheCompetences(Salarie, Emploi: string; GrilleEmploi, GrilleSalarie: THGrid);
  procedure CarriereAfficheFormations(LeSalarie: string; GrilleInit, GrilleInterne, GrilleExterne: THGrid);
  procedure CarriereAfficheHisto(LeSalarie: string; DateDeb, DateFin: TDateTime; Grille: THGrid; THCriteres: THMultiValComboBox);
  procedure CarriereAjoutLigneHisto(TobLigneHisto: Tob; Libelle, Valeur, Commentaire: string; DateModif: TDateTime);
  procedure CarriereAfficheContrats(LeSalarie: string; Grille: THGrid);
  procedure CarriereHistoInit( Combo : THMultiValComboBox);
  Function  PGCalcAllocationDIF(Salarie : String;NbHeuresNT : Double;DateSession : TDateTime) : Double;
  Function  PGDIFGenereAbs(HeuresT,HeuresNonT,Jours : Double;Salarie,DJ,FJ,TypeConge,Etablissement : String;DateD,DateF : TDateTime) : Integer;
  procedure PGDIFGenereRubEnTete(Heures : Double;Salarie,Etablissement,Rubrique,TypeAlim,Stage,Confidentiel : String;DateD,DateF,DatePaie : TDateTime);
  procedure PGDIFGenereRub(HeuresT,HeuresNonT : Double;Salarie,Etablissement,Rubrique,TypeAlim,Stage,Confidentiel : String;DateD,DateF,DatePaie : TDateTime;NumOrdre : Integer);
  procedure PGDIFSuppAbs(Salarie,TypeConge : String;DateDeb,DateFin : TDateTime);
  procedure PGDIFSuppRub(Salarie,Rubrique,TypeAlim : String;DatePaie : TDateTime);
  procedure PGRecupereAllocFormation(DD,DF : TDateTime);
  procedure RecupereLesFormationsDIF(Salarie,TypePlan,Rubrique,TypeAlim : String;DateVAL : TDateTime;Travail,HorsTravail : Boolean;LeStage : String = ''; Ordre : Integer = 0);
  procedure ReaffectPlafondForm;
  Procedure MajCoutPrev(Stage,Millesime : String;TF : TTableFiltre);
  Function  RendMillesimeEManager : String;
  Procedure MajSousSessionSal(Salarie,Stage,NMultiSession : String);
  Function  DonnePopulation(Salarie: String): String; //PT6
  Function  AdaptByRespEmanager (TypeUtilisateur,Prefixe,Resp : String; ServicesSup : Boolean) : String; overload; //PT9
  Function  AdaptByRespEmanager (TypeUtilisateur,Prefixe,Resp,Suffixe : String; ServicesSup : Boolean) : String; overload ; //PT15
  Procedure ElipsisSalarieMultidos (Sender : TObject; Where : String = ''); //PT13B
  Procedure ElipsisResponsableMultidos (Sender : TObject; Where : String = ''); //PT13B
  Procedure MAJResponsable (Service,TypeResp,NouveauResp : String); //PT14
  procedure MajCompetences (Salarie,Stage,Session,Millesime : String; Present : Boolean); //PT28
  procedure InsererDIFBulletin (TobInsc : TOB; InitDatePaie : TDateTime); //PT28
  procedure SupprimerDIF(TobInsc : TOB; InitDatePaie : TDateTime; Supp : Boolean; var SuppAbsence,SuppRubrique : Boolean); //PT28
  Function  RendNivAccesConfidentiel : String; //PT31
{$ENDIF}
  Procedure MultiDosFormation;
  Function  AdaptMultiDosForm (ChaineSQL : String; SelectStatic : Boolean = True) : String; //PT16
  Function  GetNoDossierSalarie (Salarie : String = '') : String; //PT16
  Function  RecherchePopulation (TInfos : TOB) : String; //PT6
  procedure CalcCtInvestSession(TypeMaj, LeStage, LeMillesime: string; LaSession: Integer);
  Function  RendMillesimeRealise(var DD,DF : TDateTime) : String;
  Function  RendMillesimePrevisionnel : String;
  procedure CreerStageAuPrevisionnel (CodeStage, Millesime: String);
  Procedure MajCoutsInscPrev (Stage, Millesime : String);
  //  Fonctions de calcul des salaires  :  renvoi le taux horaire
  function  ForTauxHoraireReel(LeSalarie: string; DebutPaie: TDateTime = 0; FinPaie: TDateTime = 0; MethodeCalc: string = ''; TypeCalc:Boolean=ValReel; DateFinFormation:TDateTime=0 ;LeDossier: string = ''): Double; //PT3 //PT4
  function  ForTauxHoraireCategoriel(LeLibEmploi, LeMillesime: string): Double;
  function  CalCoutPedagogique (Stage, Millesime, Etablissement, Rang: String; NbInsc : Integer): Double; //PT11
  Function  DossiersAInterroger (Predefini, Dossier, Prefixe : String; AvecPredefini : Boolean = False; AvecAnd : Boolean = False) : String; //PT12
  procedure PrepareMailFormation (Salarie, Responsable, Action : String; TobInfos : TOB; var Titre : String; var Texte : HTStrings; HTML : Boolean = False); //PT30
  Function  GetBaseSalarie (Salarie : String = '') : String;  //PT21
  Function  GetBasePartage(Bundle : Integer) : String; //PT21
var
{$IFNDEF PLUGIN}
  TobPrepAGEFOS, TobPasserelle, TobHistorique: Tob;
{$ELSE}
  GrpDossiers  : Boolean; //PT25
  BasesPartage : Array [0..9] of String; //PT25
{$ENDIF}
  PGBundleInscFormation,PGDroitMultiForm,PGBundleHierarchie,PGBundleCatalogue : Boolean; //PT23

implementation

uses
{$IFNDEF PLUGIN}
  Ent1;
{$ELSE}
  PGVIGUTIL;
{$ENDIF}

{$IFNDEF PLUGIN}

function AfficheLibelle(Func, Params, WhereSQL: string; TT: TDataset; Total: Boolean): string;
var
  Salarie: string;
  Nb: Integer;
begin
  if Func = 'LIBFORMATION' then
  begin
    Salarie := TT.FindField('PFI_SALARIE').AsString;
    if Salarie <> '' then Result := RechDom('PGSALARIE', Salarie, False);
  end
  else
  begin
    Nb := TT.FindField('PFI_NBINSC').AsInteger;
    Result := IntToStr(Nb) + ' Salariés';
  end;
end;

procedure VisibiliteChampFormation(Numero: string; Champ1, Champ2: TControl);
var
  TLieu: THLabel;
  ChampLieu: THValComboBox;
  Num: Integer;
begin
  Num := StrToInt(Numero);
  ChampLieu := THValComboBox(Champ1);
  TLieu := THLabel(Champ2);
  if Num > VH_Paie.NBFormationLibre then
  begin
    if ChampLieu <> nil then ChampLieu.Visible := FALSE;
    if TLieu <> nil then TLieu.Visible := FALSE;
    exit;
  end;
  if ChampLieu <> nil then
  begin
    ChampLieu.Enabled := TRUE;
    ChampLieu.Visible := TRUE;
  end;
  if TLieu <> nil then
  begin
    TLieu.Visible := TRUE;
    if Num = 1 then TLieu.Caption := VH_Paie.FormationLibre1
    else if Num = 2 then TLieu.Caption := VH_Paie.FormationLibre2
    else if Num = 3 then TLieu.Caption := VH_Paie.FormationLibre3
    else if Num = 4 then TLieu.Caption := VH_Paie.FormationLibre4
    else if Num = 5 then TLieu.Caption := VH_Paie.FormationLibre5
    else if Num = 6 then TLieu.Caption := VH_Paie.FormationLibre6
    else if Num = 7 then TLieu.Caption := VH_Paie.FormationLibre7
    else TLieu.Caption := VH_Paie.FormationLibre8;
  end;
end;

procedure VisibiliteChampCompetence(Numero: string; Champ1, Champ2: TControl);
var
  TLieu: THLabel;
  ChampLieu: THValComboBox;
  Num: Integer;
  Libelle: string;
begin
  Num := StrToInt(Numero);
  ChampLieu := THValComboBox(Champ1);
  TLieu := THLabel(Champ2);
  Libelle := RechDom('GCZONELIBRE', 'RH' + IntToStr(Num), False);
  if Libelle = 'Table libre ' + IntToStr(Num) then
  begin
    if ChampLieu <> nil then ChampLieu.Visible := FALSE;
    if TLieu <> nil then TLieu.Visible := FALSE;
    exit;
  end;
  if ChampLieu <> nil then
  begin
    ChampLieu.Enabled := TRUE;
    ChampLieu.Visible := TRUE;
  end;
  if TLieu <> nil then
  begin
    TLieu.Visible := TRUE;
    TLieu.Caption := Libelle;
  end;
end;

procedure VisibiliteChampCompetenceRess(Numero: string; Champ1, Champ2: TControl);
var
  TLieu: THLabel;
  ChampLieu: THValComboBox;
  Num: Integer;
  Libelle: string;
begin
  Num := StrToInt(Numero);
  ChampLieu := THValComboBox(Champ1);
  TLieu := THLabel(Champ2);
  Libelle := RechDom('GCZONELIBRE', 'CP' + IntToStr(Num), False);
  if Libelle = 'Table libre ' + IntToStr(Num) then
  begin
    if ChampLieu <> nil then ChampLieu.Visible := FALSE;
    if TLieu <> nil then
    begin
      TLieu.Visible := FALSE;
      TLieu.Caption := '';
    end;
    exit;
  end;
  if ChampLieu <> nil then
  begin
    ChampLieu.Enabled := TRUE;
    ChampLieu.Visible := TRUE;
  end;
  if TLieu <> nil then
  begin
    TLieu.Visible := TRUE;
    TLieu.Caption := Libelle;
  end;
end;

procedure RecupInscForReporte(Millesime: string);
var
  TobInsc, T: Tob;
  Q: TQuery;
  Rang: Integer;
  MillesimeSuiv: string;
begin
  MillesimeSuiv := IntToStr(StrToInt(Millesime) + 1);
  Q := OpenSQL('SELECT PFI_SALARIE,PFI_ETABLISSEMENT,PFI_RESPONSFOR,PFI_RANG,PFI_NBINSC,PFI_MILLESIME' +
    ' FROM INSCFORMATION WHERE ' +
    'PFI_REPORTELE <> "' + UsDateTime(IDate1900) + '" AND PFI_REFUSELE="' + UsDateTime(IDate1900) + '" ' +
    'AND PFI_DATEACCEPT="' + UsDateTime(IDate1900) + '"', True);
  TobInsc := Tob.Create('Les inscriptions', nil, -1);
  TobInsc.LoadDetailDB('INSCFORMATION', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT MAX(PFI_RANG) AS RANG FROM INSCFORMATION WHERE PFI_MILLESIME="' + MillesimeSuiv + '"', true);
  Rang := Q.FindField('RANG').AsInteger;
  Ferme(Q);
  T := TobInsc.FindFirst([''], [''], False);
  while T <> nil do
  begin
    executeSQL('INSERT INTO INSCFORMATION (PFI_SALARIE,PFI_ETABLISSEMENT,PFI_REPSONSFOR,PFI_RANG,PFI_NBINSC,PFI_MILLESIME)' +
      ' VALUES ' +
      '("' + T.GetValue('PFI_SALARIE') + '","' + T.GetValue('PFI_ETABLISSEMENT') + '","' + T.GetValue('RESPONSFOR') + '","' + IntToStr(Rang) + '",' +
      '"' + IntToStr(T.GetValue('NBINSC')) + '","' + MillesimeSuiv + '")');
    rang := Rang + 1;
    T := TobInsc.FindNext([''], [''], False);
  end;
end;

procedure recupInscNonForm(Millesime: string);
var
  Q: TQuery;
  TobInsc, T: Tob;
  rang: Integer;
  MillesimeSuiv: string;
begin
  MillesimeSuiv := IntToStr(StrToInt(Millesime) + 1);
  Q := OpenSQL('SELECT PFI_SALARIE,PFI_ETABLISSEMENT,PFI_RESPONSFOR,PFI_RANG,PFI_NBINSC,PFI_MILLESIME' +
    ' FROM INSCFORMATION WHERE ' +
    ' PFI_SALARIE <> "" AND PFI_SALARIE NOT IN (SELECT PFO_SALARIE FROM FORMATIONS WHERE PFO_MILLESIME="' + Millesime + '")' +
    ' AND PFI_MILLESIME="' + Millesime + '"', true);
  TobInsc := Tob.Create('Les inscriptions', nil, -1);
  TobInsc.LoadDetailDB('INSCFORMATION', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT MAX(PFI_RANG) AS RANG FROM INSCFORMATION WHERE PFI_MILLESIME="' + MillesimeSuiv + '"', true);
  Rang := Q.FindField('RANG').AsInteger;
  Ferme(Q);
  T := TobInsc.FindFirst([''], [''], False);
  while T <> nil do
  begin
    executeSQL('INSERT INTO INSCFORMATION (PFI_SALARIE,PFI_ETABLISSEMENT,PFI_REPSONSFOR,PFI_RANG,PFI_NBINSC,PFI_MILLESIME)' +
      ' VALUES ' +
      '("' + T.GetValue('PFI_SALARIE') + '","' + T.GetValue('PFI_ETABLISSEMENT') + '","' + T.GetValue('RESPONSFOR') + '","' + IntToStr(Rang) + '",' +
      '"' + IntToStr(T.GetValue('NBINSC')) + '","' + MillesimeSuiv + '")');
    rang := Rang + 1;
    T := TobInsc.FindNext([''], [''], False);
  end;
end;

function VerifPlafondFormation(LeFrais, LeMillesime: string; LeMontant: Double): Double;
var
  Q: TQuery;
  Plafond: Double;
begin
  plafond := 0;
  Q := OpenSQL('SELECT PFE_PLAFONDFL' + LeFrais + ' FROM EXERFORMATION WHERE PFE_MILLESIME="' + LeMillesime + '"', True);
  if not Q.eof then Plafond := Q.FindField('PFE_PLAFONDFL' + LeFrais).AsFloat;
  Ferme(Q);
  if (LeMontant > Plafond) and (plafond <> 0) then Result := Plafond
  else Result := LeMontant;
end;

function RecupSalaireFormation(LeSalarie: string): Double;
var
  Q: TQuery;
  Salaire: Double;
begin
  Salaire := 0;
  Q := OpenSQL('SELECT PSA_TAUXHORAIRE FROM SALARIES WHERE PSA_SALARIE="' + LeSalarie + '"', true);
  if not Q.eof then Salaire := Q.FindField('PSA_TAUXHORAIRE').AsFloat;
  Ferme(Q);
  Result := Arrondi(Salaire, 2);
end;

function RecupSalaireFormationNonNom(LeMillesime, LeLibEmploi: string): Double;
var
  Q: TQuery;
  Salaire: Double;
begin
  Salaire := 0;
  Q := OpenSQL('SELECT PSF_LIBEMPLOIFOR,PSF_MILLESIME,PSF_MONTANT ' +
    'FROM SALAIREFORM WHERE PSF_MILLESIME="' + LeMillesime + '" AND PSF_LIBEMPLOIFOR="' + LeLibEmploi + '"', True);
  if not Q.eof then Salaire := Q.FindField('PSF_MONTANT').AsFloat;
  Ferme(Q);
  Result := Arrondi(Salaire, 2);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jérôme LEFEVRE
Créé le ...... : 30/01/2003
Modifié le ... :   /  /
Description .. : Procédure qui met a jour le cout pédagogique d'une
Suite ........ : formation par salariés dans la table formation
Mots clefs ... :
*****************************************************************}

procedure MAJCoutsFormation(LeMillesime, LeStage, LaSession: string;TF : TTableFiltre = Nil);
var
  TobFormation, T: Tob;
  Q: TQuery;
  CoutSalarie, TotalHeures, HeuresSal, CoutAnim, CoutGlobal, CoutParStagiaire: Double;
  i: Integer;
begin
  TotalHeures := 0;
  CoutAnim := 0;
  CoutGlobal := 0;
  CoutParStagiaire := 0;
  Q := OpenSQL('SELECT PSS_COUTSALAIR,PSS_COUTPEDAG,PSS_COUTUNITAIRE FROM SESSIONSTAGE ' +
    'WHERE PSS_CODESTAGE="' + LeStage + '" AND PSS_MILLESIME="' + LeMillesime + '" AND PSS_ORDRE=' + LaSession + '', True); //DB2
  if not Q.Eof then
  begin
    CoutAnim := Q.FindField('PSS_COUTSALAIR').AsFloat;
    CoutGlobal := Q.FindField('PSS_COUTPEDAG').AsFloat;
    CoutParStagiaire := Q.FindField('PSS_COUTUNITAIRE').AsFloat;
  end;
  Ferme(Q);
  Q := OpenSQL('SELECT SUM(PFO_NBREHEURE) AS NBHEURE FROM FORMATIONS WHERE PFO_CODESTAGE="' + LeStage + '" AND PFO_ORDRE=' + LaSession + ' AND PFO_MILLESIME="' + LeMillesime + '" ' + //DB2
    'AND PFO_EFFECTUE="X"', True);
  if not Q.eof then TotalHeures := Q.FindField('NBHEURE').AsFloat;
  Ferme(Q);

  If TF = Nil then
  begin
         Q := OpenSQL('SELECT * FROM FORMATIONS WHERE PFO_CODESTAGE="' + LeStage + '" AND PFO_ORDRE=' + LaSession + ' AND PFO_MILLESIME="' + LeMillesime + '" AND PFO_EFFECTUE="X"', True); //DB2
         TobFormation := Tob.Create('FORMATIONS', nil, -1);
         TobFormation.LoadDetailDB('FORMATIONS', '', '', Q, False);
         Ferme(Q);
         for i := 0 to TobFormation.Detail.Count - 1 do
         begin
         CoutSalarie := 0;
         T := TobFormation.Detail[i];
         
         If T.GetValue('PFO_EFFECTUE') = 'X' Then //PT32
         Begin
	         if TotalHeures > 0 then
	         begin
	              HeuresSal := T.GetValue('PFO_NBREHEURE');
	              CoutSalarie := CoutAnim + CoutGlobal;
	              CoutSalarie := (CoutSalarie * HeuresSal) / TotalHeures;
	         end;
	         CoutSalarie := CoutSalarie + CoutParStagiaire;
	     End;
         T.PutValue('PFO_AUTRECOUT', CoutSalarie);
         T.UpdateDB(False);
    end;
    TobFormation.Free;
  end
  else
  begin
       TF.DisableTOM;
       for i:=1 to TF.LaGrid.RowCount - 1 do
       begin
            TF.SelectRecord(i);
            MoveCur(False);
            CoutSalarie := 0;
            
            If TF.GetValue('PFO_EFFECTUE') = 'X' Then //PT32
         	Begin
	            if TotalHeures > 0 then
	            begin
	              HeuresSal := TF.GetValue('PFO_NBREHEURE');
	              CoutSalarie := CoutAnim + CoutGlobal;
	              CoutSalarie := (CoutSalarie * HeuresSal) / TotalHeures;
	            end;
	            CoutSalarie := CoutSalarie + CoutParStagiaire;
	        End;
            TF.PutValue('PFO_AUTRECOUT',CoutSalarie);
            TF.Post;
       end;
       TF.EnableTOM;
  end;

end;

function NombreDeSessionFormationAPrevoir(NbMax, MbInsc: Integer): Double;
var
  NbCout, NbDivE, NbDivD: Double;
begin
  if NbMax > 0 then
  begin
    NbDivD := arrondi(MbInsc / NbMax, 2);
    NbDivE := arrondi(NbDivD, 0);
    if NbDivE - NbDivD < 0 then NbCout := NbDivE + 1
    else NbCout := NbDivE;
  end
  else NbCout := 1;
  Result := NbCout;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jérôme LEFEVRE
Créé le ...... : 20/02/2003
Modifié le ... :   /  /
Description .. : Chargement des données pour
Suite ........ : construire le fichier excel pour AGEFOS
Mots clefs ... :
*****************************************************************}

procedure ChargeTobCreerFichierForm(NumEnvoiForm: integer);
var
  Q: TQuery;
begin
  Q := OpenSQL('SELECT  PSS_ORDRE,PSS_CODESTAGE,PSS_MILLESIME,' +
    'PSS_COUTPEDAG,PSS_COUTUNITAIRE,PSS_LIEUFORM,PST_DUREESTAGE,PSS_FORMATION1,PSS_AVECCLIENT,PSS_NATUREFORM,PSS_CENTREFORMGU ' +
    'FROM SESSIONSTAGE ' +
    'LEFT JOIN STAGE ON PSS_CODESTAGE=PST_CODESTAGE AND PSS_MILLESIME=PST_MILLESIME ' +
    'WHERE PSS_NUMENVOI=' + IntToStr(NumEnvoiForm) + ' ', True); //DB2
  TobPrepAGEFOS := Tob.Create('Les sessions', nil, -1);
  TobPrepAGEFOS.LoadDetailDB('Les sessions', '', '', Q, False);
  ferme(Q);
  TobPasserelle := Tob.Create('Creation passerelle', nil, -1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jérôme LEFEVRE
Créé le ...... : 20/02/2003
Modifié le ... :   /  /
Description .. : Récupération des données concernant les animateurs pour
Suite ........ : construire le fichier excel pour AGEFOS
Mots clefs ... :
*****************************************************************}

procedure RecupDonneesEnvoiAnim(NumEnvoiForm : Integer; OPCACP, OPCACS: String; Trace, TraceErr: TListBox);
var
  TobAnim, TP, TobFrais, TF, TobMAjSessions, TMaj: Tob;
  Stage, Millesime, MillesimeEC, IntituleForm: string;
  Centreform,NomAnim, MatAnim, PrenomAnim, Etablissement: string;
  Session, i, j, Age, NbSal: Integer;
  Q: TQuery;
  DateDebut, DateFin: TDateTime;
  OPCAPedag, OPCASal: String;
  Duree, DureeJours, DureeForm, TotalFrais, TotalHT, TotalMaj, CtSalarial, CtPedag: Double;
  CoutSalarial, TauxCharge, SalHoraire: Double;
  Lieu, AvecClient, Axe, Nature, Statut, LibOPCASal, LibOPCAPedag: string;
  NumAdherent, LibCentreForm, CP, Adresse, Ville, Contact, Tel, Fax, Agrement: string;
  FraisKM, FraisRepas, FraisHotel, FraisTrain, FraisMetro, FraisPeage, FraisParking, FraisSeminaire, FraisEssence, FraisDivers, FraisPedag, FraisLoc: Double;
  aaaa, mm, jj, dd, oo, yyyy: Word;
  NbHotel, Nbrepas, NbKM, CoutUnitaire, NbHTot, CtPedagAnim: Double;
begin
  MillesimeEC := '0000';
  Q := OpenSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE PFE_ACTIF="X" AND PFE_CLOTURE="-"', True);
  if not Q.Eof then MillesimeEC := Q.FindField('PFE_MILLESIME').AsString;
  Ferme(Q);
  for i := 0 to TobPrepAGEFOS.Detail.Count - 1 do
  begin
    TotalMaj := 0;
    Stage := TobPrepAGEFOS.Detail[i].GetValue('PSS_CODESTAGE');
    Millesime := TobPrepAGEFOS.Detail[i].GetValue('PSS_MILLESIME');
    Session := TobPrepAGEFOS.Detail[i].GetValue('PSS_ORDRE');
    Duree := 0;
    DateDebut := IDate1900;
    DateFin := IDate1900;
    DureeJours := 0;
    CoutUnitaire := 0;
    CtPedag := 0;
    {Récupération des infos concernant la session traitée}
    Q := OpenSQL('SELECT SUM(PAN_NBREHEURE) NBH,PSS_COUTUNITAIRE,PSS_DATEDEBUT,' +
      'PSS_DATEFIN,PST_LIBELLE,PSS_ORGCOLLECTPGU,PSS_ORGCOLLECTSGU,PSS_COUTPEDAG,PSS_LIEUFORM,PSS_DUREESTAGE,PSS_JOURSTAGE,PSS_FORMATION1,PSS_AVECCLIENT,PSS_NATUREFORM,PSS_CENTREFORMGU' +
      ' FROM SESSIONSTAGE LEFT JOIN STAGE ON PSS_MILLESIME=PST_MILLESIME AND PSS_CODESTAGE=PST_CODESTAGE ' +
      'LEFT JOIN SESSIONANIMAT ON PSS_MILLESIME=PAN_MILLESIME AND PAN_CODESTAGE=PST_CODESTAGE AND PSS_ORDRE=PAN_ORDRE ' +
      'WHERE PSS_MILLESIME="' + Millesime + '" AND PSS_ORDRE=' + IntToStr(Session) + ' AND PSS_CODESTAGE="' + Stage + '" ' + //DB2
      'GROUP BY PSS_COUTUNITAIRE,PSS_DATEDEBUT,PSS_DATEFIN,PST_LIBELLE,PSS_ORGCOLLECTPGU,PSS_ORGCOLLECTSGU,PSS_COUTPEDAG,PSS_LIEUFORM,PSS_DUREESTAGE,PSS_JOURSTAGE,PSS_FORMATION1,PSS_AVECCLIENT,PSS_NATUREFORM,PSS_CENTREFORMGU', True);
    if not Q.eof then
    begin
      Duree := Q.FindField('PSS_DUREESTAGE').AsFloat;
      DureeJours := Q.FindField('PSS_JOURSTAGE').AsFloat;
      Lieu := Q.FindField('PSS_LIEUFORM').AsString;
      CentreForm := Q.FindField('PSS_CENTREFORMGU').AsString;
      AvecClient := Q.FindField('PSS_AVECCLIENT').AsString;
      if AvecClient = 'X' then AvecClient := 'Oui'
      else AvecClient := 'Non';
      Axe := Q.FindField('PSS_FORMATION1').AsString;
      Nature := Q.FindField('PSS_NATUREFORM').AsString;
      OPCAPedag := Q.FindField('PSS_ORGCOLLECTPGU').AsString;
      OPCASal := Q.FindField('PSS_ORGCOLLECTSGU').AsString;
      IntituleForm := Q.FindField('PST_LIBELLE').AsString;
      DateDebut := Q.FindField('PSS_DATEDEBUT').AsdateTime;
      DateFin := Q.FindField('PSS_DATEFIN').AsDateTime;
      CtPedag := Q.FindField('PSS_COUTPEDAG').Asfloat;
      NbHTot := Q.FindField('NBH').Asfloat;
      CoutUnitaire := Q.FindField('PSS_COUTUNITAIRE').Asfloat;
    end;
    Ferme(Q);
    NbSal := 0;
    Q := OpenSQL(' SELECT COUNT (PFO_SALARIE) NBSAL FROM FORMATIONS ' +
      'WHERE PFO_MILLESIME="' + Millesime + '" AND PFO_ORDRE=' + IntToStr(Session) + ' AND PFO_CODESTAGE="' + Stage + '"', True); //DB2
    if not Q.Eof then NbSal := Q.FindField('NBSAL').AsInteger;
    Ferme(Q);
    CoutUnitaire := NbSal * CoutUnitaire;
    //If (OPCAPedag <> OPCACP) and (OPCACP >= 0) then continue; {Non traité si OPCA session différent de OPCA destinataire, message d'erreur lors traitement salariés}
   // If (OPCASal <> OPCACS) and (OPCACS >= 0) then continue;
    LibCentreForm := '';
    CP := '';
    Adresse := '';
    Ville := '';
    Contact := '';
    Tel := '';
    Fax := '';
    LibOPCASal := '';
    LibOPCAPedag := '';
    Agrement := '';
    NumAdherent := '';
    if Nature = '002' then {Si nature = 002 : formation externe dc récupération des données du centre de formation}
    begin
      Q := OpenSQL('SELECT ANN_NOADMIN,ANN_FAX,ANN_APNOM,ANN_APRUE1,ANN_APRUE2,ANN_APRUE3,ANN_ALCP,ANN_ALVILLE,ANN_TEL1 FROM ' +
        'ANNUAIRE WHERE ANN_GUIDPER="' + CentreForm + '"', True); // DB2
      if not Q.eof then
      begin
        LibCentreForm := Q.FindField('ANN_APNOM').AsString;
        CP := Q.FindField('ANN_ALCP').AsString;
        Adresse := Q.FindField('ANN_APRUE1').AsString + ' ' + Q.FindField('ANN_APRUE2').AsString + ' ' + Q.FindField('ANN_APRUE3').AsString;
        Ville := Q.FindField('ANN_ALVILLE').AsString;
        Contact := Q.FindField('ANN_ALCP').AsString;
        Tel := Q.FindField('ANN_TEL1').AsString;
        Fax := Q.FindField('ANN_FAX').AsString;
        Agrement := Q.FindField('ANN_NOADMIN').AsString;
      end;
      Ferme(Q);
    end;
    {Récupération des n° adhérent et libellé pour OPCA Des cou^ts pédagogiques et salariés}
    Q := OpenSQL('SELECT ANN_NOM1 FROM ANNUAIRE WHERE ANN_GUIDPER="' + OPCAPedag + '"', True); // DB2
    if not Q.Eof then LibOPCAPedag := Q.FindField('ANN_NOM1').AsString;
    Ferme(Q);
    Q := OpenSQL('SELECT ANN_NOM1,ANN_NOADMIN FROM ANNUAIRE WHERE ANN_GUIDPER="' + OPCASal + '"', True); // DB2
    if not Q.Eof then
    begin
      NumAdherent := Q.FindField('ANN_NOADMIN').AsString;
      LibOPCASal := Q.FindField('ANN_NOM1').AsString;
    end;
    Ferme(Q);
    {Récupération des infos concernant les salariés animateurs}
    Q := OpenSQL('SELECT PSA_ETABLISSEMENT,PSA_SEXE,PSA_DATENAISSANCE,PSA_DADSCAT,PAN_SALAIREANIM,PAN_NBREHEURE,PAN_CODESTAGE,PAN_SALARIE,PSA_LIBELLE,PSA_PRENOM FROM SESSIONANIMAT ' +
      'LEFT JOIN SALARIES ON PSA_SALARIE=PAN_SALARIE ' +
      ' WHERE PAN_SALARIE <> "" AND PAN_MILLESIME="' + Millesime + '" AND PAN_ORDRE=' + IntToStr(Session) + ' AND PAN_CODESTAGE="' + Stage + '"', True); //DB2
    TobAnim := Tob.Create('Les animateurs', nil, -1);
    TobAnim.LoadDetailDB('Les animateurs', '', '', Q, False);
    Ferme(Q);
    for j := 0 to TobAnim.Detail.Count - 1 do
    begin
      DureeForm := TobAnim.Detail[j].GetValue('PAN_NBREHEURE');
      MatAnim := TobAnim.Detail[j].GetValue('PAN_SALARIE');
      NomAnim := TobAnim.Detail[j].GetValue('PSA_LIBELLE');
      PrenomAnim := TobAnim.Detail[j].GetValue('PSA_PRENOM');
      Etablissement := TobAnim.Detail[j].GetValue('PSA_ETABLISSEMENT');
      CtSalarial := TobAnim.Detail[j].GetValue('PAN_SALAIREANIM');
      if DureeForm <= 0 then
      begin
        if TraceErr <> nil then TraceErr.Items.Add('Le nombre d''heures n''est pas renseigné pour le formateur ' +
            MatAnim + ' ' + NomAnim + '' + PrenomAnim + ' lors de la formation ' + IntituleForm);
        Continue;
      end;
      TP := Tob.Create('Une fille anim', TobPasserelle, -1);
      TP.AddChampSupValeur('MILLESIME', MillesimeEC, False); ///*********
      TP.AddChampSupValeur('ADHERENT', NumAdherent, false);
      TP.AddChampSupValeur('ENVOI', NumEnvoiForm, false);
      TP.AddChampSupValeur('ACTION', IntToStr(Session) + Stage, false);
      TP.AddChampSupValeur('OPCA', '', false);
      TP.AddChampSupValeur('AXE', RechDom('PGFORMATION1', Axe, False), false);
      TP.AddChampSupValeur('INTITULE', IntituleForm, false);
      TP.AddChampSupValeur('CPLIEU', '', False); ///*********
      TP.AddChampSupValeur('LIEU', RechDom('PGLIEUFORMATION', Lieu, False), false);
      TP.AddChampSupValeur('FORMATIONCLIENT', AvecClient, false);
      TP.AddChampSupValeur('DUREEFORMATION', Duree, false);
      TP.AddChampSupValeur('NBJOURSFORMATION', DureeJours, false); ///*********
      TP.AddChampSupValeur('DUREEFORMATEUR', DureeForm, false);
      TP.AddChampSupValeur('DUREESTAGIAIRE', 0, false);
      TP.AddChampSupValeur('DATEDEBUT', DateToStr(DateDebut), false); ///*********
      TP.AddChampSupValeur('DATEFIN', dateToStr(DateFin), false); ///*********
      TP.AddChampSupValeur('MATRICULE', MatAnim, false);
      TP.AddChampSupValeur('TYPE', 'Formateur', false);
      TP.AddChampSupValeur('NOM', NomAnim, false);
      TP.AddChampSupValeur('PRENOM', PrenomAnim, false);
      TP.AddChampSupValeur('ETABLISSEMENT', Etablissement, false); ///*********
      Statut := TobAnim.Detail[j].GetValue('PSA_DADSCAT');
      if (Statut = '01') or (Statut = '02') then TP.AddChampSupValeur('STATUT', 'Cadre')
      else TP.AddChampSupValeur('STATUT', 'Non cadre');
      DecodeDate(TobAnim.Detail[j].GetValue('PSA_DATENAISSANCE'), aaaa, mm, jj);
      DecodeDate(Date, yyyy, oo, dd);
      Age := yyyy - aaaa;
      TP.AddChampSupValeur('AGE', Age, false);
      TP.AddChampSupValeur('SEXE', TobAnim.Detail[j].GetValue('PSA_SEXE'), false);
      Q := OpenSQL('SELECT PFS_FRAISSALFOR,PFS_MONTANT,PFS_SALARIE,PFS_QUANTITE FROM FRAISSALFORM ' +
        'WHERE PFS_SALARIE="' + TobAnim.Detail[j].GetValue('PAN_SALARIE') + '"' +
        ' AND PFS_MILLESIME="' + Millesime + '" AND PFS_ORDRE=' + IntToStr(Session) + ' ' + //DB2
        'AND PFS_CODESTAGE="' + Stage + '"', True);
      TobFrais := Tob.Create('Les frais', nil, -1);
      TobFrais.LoadDetailDB('FRAISSALFORM', '', '', Q, False);
      Ferme(Q);
      NbKM := 0;
      FraisKM := 0;
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['003'], False);
      if TF <> nil then
      begin
        FraisKM := TF.GetValue('PFS_MONTANT');
        NbKM := TF.GetValue('PFS_QUANTITE');
      end;
      TP.AddChampSupValeur('NBKM', NbKM, False);
      TP.AddChampSupValeur('FRAISKM', FraisKM, False);
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['001'], False);
      NbRepas := 0;
      FraisRepas := 0;
      if TF <> nil then
      begin
        FraisRepas := TF.GetValue('PFS_MONTANT');
        NbRepas := TF.GetValue('PFS_QUANTITE');
      end;
      TP.AddChampSupValeur('NBREPAS', NbRepas, False);
      TP.AddChampSupValeur('FRAISREPAS', FraisRepas, False);
      NbHotel := 0;
      FraisHotel := 0;
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['002'], False);
      if TF <> nil then
      begin
        FraisHotel := TF.GetValue('PFS_MONTANT');
        NbHotel := TF.GetValue('PFS_QUANTITE');
      end;
      TP.AddChampSupValeur('NBNUITS', NbHotel, False);
      TP.AddChampSupValeur('FRAISHOTEL', FraisHotel, False);
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['005'], False);
      if TF <> nil then FraisTrain := TF.GetValue('PFS_MONTANT')
      else FraisTrain := 0;
      TP.AddChampSupValeur('FRAISTRAIN_AVION', FraisTrain, False);
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['010'], False);
      if TF <> nil then FraisMetro := TF.GetValue('PFS_MONTANT')
      else FraisMetro := 0;
      TP.AddChampSupValeur('FRAISMETRO', FraisMetro, False);
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['007'], False);
      if TF <> nil then FraisParking := TF.GetValue('PFS_MONTANT')
      else FraisParking := 0;
      TP.AddChampSupValeur('FRAISPARKING', FraisParking, False);
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['008'], False);
      if TF <> nil then FraisSeminaire := TF.GetValue('PFS_MONTANT')
      else FraisSeminaire := 0;
      TP.AddChampSupValeur('FRAISSEMINAIRE', FraisSeminaire, False);
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['011'], False);
      if TF <> nil then FraisEssence := TF.GetValue('PFS_MONTANT')
      else FraisEssence := 0;
      TP.AddChampSupValeur('FRAISESSENCE', FraisEssence, False);
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['004'], False);
      if TF <> nil then FraisLoc := TF.GetValue('PFS_MONTANT')
      else FraisLoc := 0;
      TP.AddChampSupValeur('FRAISLOCATIONVOITURE', FraisLoc, False);
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['012'], False);
      if TF <> nil then FraisPedag := TF.GetValue('PFS_MONTANT')
      else FraisPedag := 0;
      TP.AddChampSupValeur('FRAISMATERIELPEDAGOGIQUE', FraisPedag, False);
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['009'], False);
      if TF <> nil then FraisDivers := TF.GetValue('PFS_MONTANT')
      else FraisDivers := 0;
      TP.AddChampSupValeur('FRAISDIVERS', FraisDivers, False);
      TobFrais.free;
      TotalFrais := FraisKM + FraisRepas + FraisHotel + FraisTrain + FraisMetro + FraisPeage + FraisParking + FraisSeminaire + FraisEssence + FraisDivers;
      if AvecClient = 'OUI' then TotalFrais := 0;
      if AvecClient = 'OUI' then CtSalarial := 0;
      TP.AddChampSupValeur('TOTALFRAIS', TotalFrais, False);
      TP.AddChampSupValeur('CTSALARIAL', CtSalarial, false);
      CtPedagAnim := CtPedag + CoutUnitaire;
      CtPedagAnim := CalcCtPedagParSalarie(0, CtPedagAnim, 0, DureeForm, NbHTot);
      if AvecClient = 'OUI' then CtPedag := 0;
      if Nature = '001' then TP.AddChampSupValeur('CTPEDAGSTAFI', CtPedagAnim, false)
      else TP.AddChampSupValeur('CTPEDAGSTAFI', 0, false);
      TP.AddChampSupValeur('ORGANISME', LibCentreForm, false);
      TP.AddChampSupValeur('AGREMENT', Agrement, false);
      TP.AddChampSupValeur('ADRESSE', Adresse, false);
      TP.AddChampSupValeur('CODEPOST', CP, false);
      TP.AddChampSupValeur('VILLE', Ville, false);
      TP.AddChampSupValeur('CONTACT', Contact, false);
      TP.AddChampSupValeur('TELEPHONE', Tel, false);
      TP.AddChampSupValeur('TELECOPIE', Fax, false);
      TP.AddChampSupValeur('CTPEDAGSTAFE', 0, false);
      TP.AddChampSupValeur('OPCACTSAL', LibOPCASal, false);
      TP.AddChampSupValeur('OPCACOUTPEDAG', LibOPCAPedag, false);
      TP.AddChampSupValeur('TOTALHT', CtPedagAnim + TotalFrais + CtSalarial, false);
    end;
    TobAnim.free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jérôme LEFEVRE
Créé le ...... : 20/02/2003
Modifié le ... :   /  /
Description .. : Récupération des données concernant les stagiaires pour
Suite ........ : construire le fichier excel pour AGEFOS
Mots clefs ... :
*****************************************************************}

procedure RecupDonneesEnvoiSta(NumEnvoiForm : Integer; OPCACP, OPCACS: String; Trace, TraceErr: TListBox);
var
  TobStagiaire, TP, TobFrais, TF, TobMAjSessions, TMaj: Tob;
  Stage, Millesime, MillesimeEC, IntituleForm, NumAdherent: string;
  MatSal, NomSal, PrenomSal, Etablissement: string;
  Session, i, j: Integer;
  Q: TQuery;
  TotalMaj: Double;
  DateDebut, DateFin: TDateTime;
  Age : Integer;
  Centreform,OPCAPedag, OPCASal : String;
  Duree, DureeJours, TotalFrais, CoutPedagogique, CoutPedagogiqueSta, SalaireAnim, FraisAnim, SalHoraire, CoutSalarial, DureeStagiaire, TauxCharge, CoutUnitaire, NbHTot, TotalHT: Double;
  Lieu, AvecClient, Axe, Nature, Statut, LibOPCAPedag, LibOPCASal: string;
  LibCentreForm, CP, Adresse, Ville, Contact, Tel, Fax, Agrement: string;
  FraisKM, FraisRepas, FraisHotel, FraisTrain, FraisMetro, FraisPeage, FraisParking, FraisSeminaire, FraisEssence, FraisDivers, FraisPedag, FraisLoc: Double;
  aaaa, mm, jj, dd, oo, yyyy: Word;
  NbKM, NbRepas, NbHotel: Integer;
  CoutSuppAnim: Double;
begin
  MillesimeEC := '0000';
  Q := OpenSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE PFE_ACTIF="X" AND PFE_CLOTURE="-"', True);
  if not Q.Eof then MillesimeEC := Q.FindField('PFE_MILLESIME').AsString;
  Ferme(Q);
  for i := 0 to TobPrepAGEFOS.Detail.Count - 1 do
  begin
    TotalMaj := 0;
    Stage := TobPrepAGEFOS.Detail[i].GetValue('PSS_CODESTAGE');
    Millesime := TobPrepAGEFOS.Detail[i].GetValue('PSS_MILLESIME');
    Session := TobPrepAGEFOS.Detail[i].GetValue('PSS_ORDRE');
    Duree := 0;
    DureeJours := 0;
    CoutUnitaire := 0;
    CoutPedagogique := 0;
    NbHTot := 0;
    DateDebut := IDate1900;
    DateFin := IDate1900;
    Q := OpenSQL('SELECT PSS_DATEDEBUT,PSS_DATEFIN,SUM (PFO_NBREHEURE) NBHTOT,PST_LIBELLE,PSS_ORGCOLLECTPGU,PSS_ORGCOLLECTSGU,PSS_COUTUNITAIRE,PSS_COUTPEDAG,PSS_LIEUFORM,PSS_DUREESTAGE,PSS_JOURSTAGE,PSS_FORMATION1,PSS_AVECCLIENT,PSS_NATUREFORM,PSS_CENTREFORMGU' +
      ' FROM SESSIONSTAGE LEFT JOIN STAGE ON PSS_MILLESIME=PST_MILLESIME AND PSS_CODESTAGE=PST_CODESTAGE ' +
      'LEFT JOIN FORMATIONS ON PSS_ORDRE=PFO_ORDRE AND PSS_CODESTAGE=PFO_CODESTAGE AND PSS_MILLESIME=PFO_MILLESIME ' +
      'WHERE PSS_MILLESIME="' + Millesime + '" AND PSS_ORDRE=' + IntToStr(Session) + ' AND PSS_CODESTAGE="' + Stage + '" ' + //DB2
      'GROUP BY PSS_DATEDEBUT,PSS_DATEFIN,PST_LIBELLE,PSS_ORGCOLLECTPGU,PSS_ORGCOLLECTSGU,PSS_COUTUNITAIRE,PSS_COUTPEDAG,PSS_LIEUFORM,PSS_DUREESTAGE,PSS_JOURSTAGE,PSS_FORMATION1,PSS_AVECCLIENT,PSS_NATUREFORM,PSS_CENTREFORMGU', True);
    if not Q.eof then
    begin
      CoutPedagogique := Q.FindField('PSS_COUTPEDAG').AsFloat;
      Duree := Q.FindField('PSS_DUREESTAGE').AsFloat;
      DureeJours := Q.FindField('PSS_JOURSTAGE').AsFloat;
      Lieu := Q.FindField('PSS_LIEUFORM').AsString;
      CentreForm := Q.FindField('PSS_CENTREFORMGU').AsString;
      AvecClient := Q.FindField('PSS_AVECCLIENT').AsString;
      if AvecClient = 'X' then AvecClient := 'Oui'
      else AvecClient := 'Non';
      Axe := Q.FindField('PSS_FORMATION1').AsString;
      Nature := Q.FindField('PSS_NATUREFORM').AsString;
      CoutUnitaire := Q.FindField('PSS_COUTUNITAIRE').AsFloat;
      NbHTot := Q.FindField('NBHTOT').AsFloat;
      OPCAPedag := Q.FindField('PSS_ORGCOLLECTPGU').AsString;
      OPCASal := Q.FindField('PSS_ORGCOLLECTSGU').AsString;
      IntituleForm := Q.FindField('PST_LIBELLE').AsString;
      DateDebut := Q.FindField('PSS_DATEDEBUT').AsdateTime;
      DateFin := Q.FindField('PSS_DATEFIN').AsDateTime;
    end;
    Ferme(Q);
    SalaireAnim := 0;
    FraisAnim := 0;
    Q := OpenSQL('SELECT SUM(PAN_SALAIREANIM) SALAIRE FROM SESSIONANIMAT' +
      ' WHERE PAN_MILLESIME="' + Millesime + '" AND PAN_ORDRE=' + IntToStr(Session) + ' AND PAN_CODESTAGE="' + Stage + '" ', True); //DB2
    if not Q.Eof then SalaireAnim := Q.FindField('SALAIRE').AsFloat;
    Ferme(Q);
    Q := OpenSQL('SELECT SUM(PFS_MONTANT) MONTANT FROM SESSIONANIMAT' +
      ' LEFT JOIN FRAISSALFORM ON PAN_SALARIE=PFS_SALARIE AND PAN_CODESTAGE=PFS_CODESTAGE AND PAN_MILLESIME=PFS_MILLESIME AND PAN_ORDRE=PFS_ORDRE' +
      ' WHERE PAN_MILLESIME="' + Millesime + '" AND PAN_ORDRE=' + IntToStr(Session) + ' AND PAN_CODESTAGE="' + Stage + '" ', True); //DB2
    if not Q.Eof then FraisAnim := Q.FindField('MONTANT').AsFloat;
    Ferme(Q);
    if AvecClient = 'OUI' then FraisAnim := 0;
    SalaireAnim := SalaireAnim + FraisAnim;
    if NbHTot <= 0 then
    begin
      if TraceErr <> nil then TraceErr.Items.Add('La formation ' + IntituleForm + ' N° ' + IntToStr(Session) + Stage + ' n''a aucune heure de formation');
      Continue;
    end;
    if (OPCAPedag <> OPCACP) and (OPCACP <>'') then
    begin
      TraceErr.Items.Add('La formation ' + IntituleForm + ' N° ' + IntToStr(Session) + Stage +
        ' ne correspond pas à l''organisme collecteur choisi pour le coût pédagogique');
      continue;
    end;
    if (OPCASal <> OPCACS) and (OPCACS <> '') then
    begin
      TraceErr.Items.Add('La formation ' + IntituleForm + ' N° ' + IntToStr(Session) + Stage +
        ' ne correspond pas à l''organisme collecteur choisi pour les coûts salariés');
      continue;
    end;
    LibCentreForm := '';
    CP := '';
    Adresse := '';
    Ville := '';
    Contact := '';
    Tel := '';
    Fax := '';
    LibOPCAPedag := '';
    LibOPCASal := '';
    Agrement := '';
    NumAdherent := '';
    Ferme(Q);
    if Nature = '002' then
    begin
      Q := OpenSQL('SELECT ANN_NOADMIN,ANN_FAX,ANN_APNOM,ANN_APRUE1,ANN_APRUE2,ANN_APRUE3,ANN_ALCP,ANN_ALVILLE,ANN_TEL1 FROM ' +
        'ANNUAIRE WHERE ANN_GUIDPER="' + CentreForm + '"', True); // DB2
      if not Q.eof then
      begin
        LibCentreForm := Q.FindField('ANN_APNOM').AsString;
        CP := Q.FindField('ANN_ALCP').AsString;
        Adresse := Q.FindField('ANN_APRUE1').AsString + ' ' + Q.FindField('ANN_APRUE2').AsString + ' ' + Q.FindField('ANN_APRUE3').AsString;
        Ville := Q.FindField('ANN_ALVILLE').AsString;
        Contact := Q.FindField('ANN_ALCP').AsString;
        Tel := Q.FindField('ANN_TEL1').AsString;
        Fax := Q.FindField('ANN_FAX').AsString;
        Agrement := Q.FindField('ANN_NOADMIN').AsString;
      end;
      Ferme(Q);
    end;
    Q := OpenSQL('SELECT ANN_NOM1 FROM ANNUAIRE WHERE ANN_GUIDPER="' + OPCAPedag + '"', True); //DB2
    if not Q.Eof then LibOPCAPedag := Q.FindField('ANN_NOM1').AsString;
    Ferme(Q);
    Q := OpenSQL('SELECT ANN_NOADMIN,ANN_NOM1 FROM ANNUAIRE WHERE ANN_GUIDPER="' + OPCASal + '"', True); // DB2
    if not Q.Eof then
    begin
      LibOPCASal := Q.FindField('ANN_NOM1').AsString;
      NumAdherent := Q.FindField('ANN_NOADMIN').AsString;
    end;
    Ferme(Q);
    Q := OpenSQL('SELECT PSA_ETABLISSEMENT,PSA_SEXE,PSA_DATENAISSANCE,PSA_DADSCAT,PFO_COUTREELSAL,PFO_NBREHEURE,PFO_CODESTAGE,PFO_SALARIE,PSA_LIBELLE,PSA_PRENOM FROM FORMATIONS ' +
      'LEFT JOIN SALARIES ON PSA_SALARIE=PFO_SALARIE ' +
      ' WHERE PFO_EFFECTUE="X" AND PFO_MILLESIME="' + Millesime + '" AND PFO_ORDRE=' + IntToStr(Session) + ' AND PFO_CODESTAGE="' + Stage + '"', True); //DB2
    TobStagiaire := Tob.Create('Les stagiaires', nil, -1);
    TobStagiaire.LoadDetailDB('Les stagiaires', '', '', Q, False);
    Ferme(Q);
    for j := 0 to TobStagiaire.Detail.Count - 1 do
    begin

      DureeStagiaire := TobStagiaire.Detail[j].GetValue('PFO_NBREHEURE');
      MatSal := TobStagiaire.Detail[j].GetValue('PFO_SALARIE');
      NomSal := TobStagiaire.Detail[j].GetValue('PSA_LIBELLE');
      PrenomSal := TobStagiaire.Detail[j].GetValue('PSA_PRENOM');
      Etablissement := TobStagiaire.Detail[j].GetValue('PSA_ETABLISSEMENT');
      if NbHTot <= 0 then
      begin
        if TraceErr <> nil then TraceErr.Items.Add('Le nombre d''heures n''est pas renseigné pour le salarié ' +
            MatSal + ' ' + NomSal + ' ' + PrenomSal + ' lors de la formation ' + IntituleForm + ' N° ' + IntToStr(Session) + Stage);
        Continue;
      end;
      TP := Tob.Create('Une fille anim', TobPasserelle, -1);
      TP.AddChampSupValeur('MILLESIME', MillesimeEC, False); ///*********
      TP.AddChampSupValeur('ADHERENT', NumAdherent, False);
      TP.AddChampSupValeur('ENVOI', NumEnvoiForm, False);
      TP.AddChampSupValeur('ACTION', IntToStr(Session) + Stage, False);
      TP.AddChampSupValeur('OPCA', '', False);
      TP.AddChampSupValeur('AXE', RechDom('PGFORMATION1', Axe, False), False);
      TP.AddChampSupValeur('INTITULE', IntituleForm, False);
      TP.AddChampSupValeur('CPLIEU', '', False); ///*********
      TP.AddChampSupValeur('LIEU', RechDom('PGLIEUFORMATION', Lieu, False), False);
      TP.AddChampSupValeur('FORMATIONCLIENT', AvecClient, False);
      TP.AddChampSupValeur('DUREEFORMATION', Duree, False);
      TP.AddChampSupValeur('NBJOURSFORMATION', DureeJours, false); ///*********
      TP.AddChampSupValeur('DUREEFORMATEUR', 0, False);
      TP.AddChampSupValeur('DUREESTAGIAIRE', DureeStagiaire, False);
      TP.AddChampSupValeur('DATEDEBUT', DateToStr(DateDebut), false); ///*********
      TP.AddChampSupValeur('DATEFIN', DateToStr(DateFin), false); ///*********
      TP.AddChampSupValeur('MATRICULE', MatSal, False);
      TP.AddChampSupValeur('TYPE', 'Stagiaire', False);
      TP.AddChampSupValeur('NOM', NomSal, False);
      TP.AddChampSupValeur('PRENOM', PrenomSal, False);
      TP.AddChampSupValeur('ETABLISSEMENT', Etablissement, false); ///*********
      Statut := TobStagiaire.Detail[j].GetValue('PSA_DADSCAT');
      if (Statut = '01') or (Statut = '02') then TP.AddChampSupValeur('STATUT', 'Cadre')
      else TP.AddChampSupValeur('STATUT', 'Non cadre', False);
      DecodeDate(TobStagiaire.Detail[j].GetValue('PSA_DATENAISSANCE'), aaaa, mm, jj);
      DecodeDate(Date, yyyy, oo, dd);
      Age := yyyy - aaaa;
      TP.AddChampSupValeur('AGE', Age, False);
      TP.AddChampSupValeur('SEXE', TobStagiaire.Detail[j].GetValue('PSA_SEXE'), False);
      CoutSalarial := TobStagiaire.Detail[j].GetValue('PFO_COUTREELSAL');
      // If (Statut = '01') or (Statut = '02') then TauxCharge := 1.17
      // Else TauxCharge := 1.12;
      TauxCharge := 1; // Calcul du taux de charge déplacé
      CoutSalarial := Arrondi(CoutSalarial * TauxCharge, 2);
      Q := OpenSQL('SELECT PFS_FRAISSALFOR,PFS_MONTANT,PFS_SALARIE,PFS_MONTANTPLAF,PFS_QUANTITE FROM FRAISSALFORM ' +
        'WHERE PFS_SALARIE="' + TobStagiaire.Detail[j].GetValue('PFO_SALARIE') + '"' +
        ' AND PFS_MILLESIME="' + Millesime + '" AND PFS_ORDRE=' + IntToStr(Session) + ' ' + //DB2
        'AND PFS_CODESTAGE="' + Stage + '"', True);
      TobFrais := Tob.Create('Les frais', nil, -1);
      TobFrais.LoadDetailDB('FRAISSALFORM', '', '', Q, False);
      Ferme(Q);
      NbKM := 0;
      FraisKM := 0;
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['003'], False);
      if TF <> nil then
      begin
        FraisKM := TF.GetValue('PFS_MONTANTPLAF');
        NbKM := TF.GetValue('PFS_QUANTITE');
      end;
      TP.AddChampSupValeur('NBKM', NbKM, False);
      TP.AddChampSupValeur('FRAISKM', FraisKM, False);
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['001'], False);
      NbRepas := 0;
      FraisRepas := 0;
      if TF <> nil then
      begin
        FraisRepas := TF.GetValue('PFS_MONTANTPLAF');
        NbRepas := TF.GetValue('PFS_QUANTITE');
      end;
      TP.AddChampSupValeur('NBREPAS', NbRepas, False);
      TP.AddChampSupValeur('FRAISREPAS', FraisRepas, False);
      NbHotel := 0;
      FraisHotel := 0;
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['002'], False);
      if TF <> nil then
      begin
        FraisHotel := TF.GetValue('PFS_MONTANTPLAF');
        NbHotel := TF.GetValue('PFS_QUANTITE');
      end;
      TP.AddChampSupValeur('NBNUITS', NbHotel, False);
      TP.AddChampSupValeur('FRAISHOTEL', FraisHotel, False);
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['005'], False);
      if TF <> nil then FraisTrain := TF.GetValue('PFS_MONTANT')
      else FraisTrain := 0;
      TP.AddChampSupValeur('FRAISTRAIN_AVION', FraisTrain, False);
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['010'], False);
      if TF <> nil then FraisMetro := TF.GetValue('PFS_MONTANT')
      else FraisMetro := 0;
      TP.AddChampSupValeur('FRAISMETRO', FraisMetro, False);
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['007'], False);
      if TF <> nil then FraisParking := TF.GetValue('PFS_MONTANT')
      else FraisParking := 0;
      TP.AddChampSupValeur('FRAISPARKING', FraisParking, False);
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['008'], False);
      if TF <> nil then FraisSeminaire := TF.GetValue('PFS_MONTANTPLAF')
      else FraisSeminaire := 0;
      TP.AddChampSupValeur('FRAISSEMINAIRE', FraisSeminaire, False);
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['011'], False);
      if TF <> nil then FraisEssence := TF.GetValue('PFS_MONTANT')
      else FraisEssence := 0;
      TP.AddChampSupValeur('FRAISTAXI', FraisEssence, False);
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['004'], False);
      if TF <> nil then FraisLoc := TF.GetValue('PFS_MONTANT')
      else FraisLoc := 0;
      TP.AddChampSupValeur('FRAISLOCATIONVOITURE', FraisLoc, False);
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['012'], False);
      if TF <> nil then FraisPedag := TF.GetValue('PFS_MONTANT')
      else FraisPedag := 0;
      TP.AddChampSupValeur('FRAISMATERIELPEDAGOGIQUE', FraisPedag, False);
      TF := TobFrais.FindFirst(['PFS_FRAISSALFOR'], ['009'], False);
      if TF <> nil then FraisDivers := TF.GetValue('PFS_MONTANT')
      else FraisDivers := 0;
      TP.AddChampSupValeur('FRAISDIVERS', FraisDivers, False);
      TobFrais.free;
      TotalFrais := FraisKM + FraisRepas + FraisHotel + FraisTrain + FraisMetro + FraisPeage + FraisParking + FraisSeminaire + FraisEssence + FraisDivers;
      TP.AddChampSupValeur('TOTALFRAIS', TotalFrais, False);
      TP.AddChampSupValeur('CTSALARIAL', CoutSalarial, False);
      if NbHTot > 0 then CoutPedagogiqueSta := CalcCtPedagParSalarie(SalaireAnim, CoutPedagogique, CoutUnitaire, DureeStagiaire, NbHTot)
      else CoutPedagogiqueSta := 0;
      if (CoutPedagogique = 0) and (CoutUnitaire = 0) then CoutSuppAnim := CoutPedagogiqueSta //PT2
      else if NbHTot = 0 then CoutSuppAnim := 0
      else CoutSuppAnim := (SalaireAnim * DureeStagiaire) / NbHTot;

      if Nature = '001' then TP.AddChampSupValeur('CTPEDAGSTAFI', CoutPedagogiqueSta, False)
      else TP.AddChampSupValeur('CTPEDAGSTAFI', 0, False);
      TP.AddChampSupValeur('ORGANISME', LibCentreForm, False);
      TP.AddChampSupValeur('AGREMENT', Agrement, False);
      TP.AddChampSupValeur('ADRESSE', Adresse, False);
      TP.AddChampSupValeur('CODEPOST', CP, False);
      TP.AddChampSupValeur('VILLE', Ville, False);
      TP.AddChampSupValeur('CONTACT', Contact, False);
      TP.AddChampSupValeur('TELEPHONE', Tel, False);
      TP.AddChampSupValeur('TELECOPIE', Fax, False);
      if Nature = '002' then TP.AddChampSupValeur('CTPEDAGSTAFE', CoutPedagogiqueSta, False)
      else TP.AddChampSupValeur('CTPEDAGSTAFE', 0, False);
      TP.AddChampSupValeur('OPCACTSAL', LibOPCASal, False);
      TP.AddChampSupValeur('OPCACOUTPEDAG', LibOPCAPedag, False);
      if nature = '001' then TotalHT := CoutPedagogiqueSta + TotalFrais + CoutSalarial
      else TotalHT := TotalFrais + CoutSalarial;
      TotalMaj := TotalMaj + TotalHT;
      TP.AddChampSupValeur('TOTALHT', TotalHT - CoutSuppAnim, False); //PT2

    end;
    TobStagiaire.free;
    if Trace <> nil then Trace.Items.Add('La formation ' + IntituleForm + ' N° ' + IntToStr(Session) + Stage + ' a été traité avec succès');
    Q := OpenSQL('SELECT * FROM SESSIONSTAGE WHERE ' +
      'PSS_CODESTAGE="' + Stage + '" AND PSS_ORDRE=' + IntToStr(Session) + ' AND PSS_MILLESIME="' + Millesime + '"', True); //DB2
    TobMAjSessions := Tob.Create('SESSIONSTAGE', nil, -1);
    TobMAjSessions.LoadDetailDB('SESSIONSTAGE', '', '', Q, False);
    Ferme(Q);
    TMaj := TobMAjSessions.FindFirst(['PSS_CODESTAGE', 'PSS_ORDRE', 'PSS_MILLESIME'], [Stage, Session, Millesime], False);
    if TMaj <> nil then
    begin
      TMaj.PutValue('PSS_TOTALHT', TotalMaj + TMaj.GetValue('PSS_TOTALHT'));
      TMaj.UpdateDB(False);
    end;
    TobMAjSessions.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jérôme LEFEVRE
Créé le ...... : 20/02/2003
Modifié le ... :   /  /
Description .. : Calcul du cout pédagogique pour un salarié en fct des
Suite ........ : différents coût et par rapport au nombre d'heures total pour
Suite ........ : la session
Mots clefs ... :
*****************************************************************}

function CalcCtPedagParSalarie(CtForm, CtPedag, CtUnit, NbHSal, NbHTot: Double): Double;
var
  CoutTotal: Double;
begin
  if NbHTot = 0 then
  begin
    CoutTotal := 0;
  end
  else
  begin
    CoutTotal := CtForm + CtPedag;
    CoutTotal := (CoutTotal * NbHSal) / NbHTot;
    CoutTotal := CoutTotal + CtUnit;
  end;
  Result := arrondi(CoutTotal, 2);
end;

function RendPlafondFraisForm(LeFrais, LeLieu, OrgCollect, LeMillesime, Population: string): Double; //PT6
var
  Q: TQuery;
  St: string;
  Plafond: Double;
begin
  Q := OpenSQL('SELECT PLF_PLAFPARIS FROM LIEUFORMATION WHERE PLF_LIEUFORM="' + LeLieu + '"', True);
  if not Q.Eof then
  begin
    if Q.FindField('PLF_PLAFPARIS').AsString = 'X' then St := 'AND PFP_PARIS="X"'
    else St := 'AND PFP_PROVINCE="X"';
  end;
  Ferme(Q);
  //PT6 - Début
  If (VH_Paie.PGForGestPlafByPop) Then
     St := St + ' AND PFP_POPULATION="'+Population+'"'
  Else
     St := St + ' AND PFP_POPULATION=""';
  //PT6 - Fin
  Q := OpenSQL('SELECT PFP_MONTANT FROM FRAISSALPLAF WHERE' +
    ' PFP_MILLESIME="' + LeMillesime + '" AND PFP_FRAISSALFOR="' + LeFrais + '" AND PFP_ORGCOLLECTGU="' + OrgCollect + '" ' + St, True); //DB2
  Plafond := 0;
  if not Q.eof then Plafond := Q.FindField('PFP_MONTANT').AsFloat;
  Ferme(Q);
  Result := Plafond;
end;

procedure SuiviCarriereGrilles(GInit, GExt, GInt, GCompetEmp, GCompetSal: THGrid);
var
  i: Integer;
  Libelle: string;
begin
  // Formations
  if GInit <> nil then
  begin
    GInit.ColFormats[0] := 'CB=PGSTAGEFORMCOMPLET';
    GInit.ColTypes[1] := 'D';
    GInit.colformats[1] := ShortDateFormat;
    GInit.ColFormats[2] := 'CB=PGCENTREFORMATION';
  end;
  if GExt <> nil then
  begin
    GExt.ColFormats[0] := 'CB=PGSTAGEFORMCOMPLET';
    GExt.ColTypes[1] := 'D';
    GExt.colformats[1] := ShortDateFormat;
    GExt.ColTypes[2] := 'D';
    GExt.colformats[2] := ShortDateFormat;
    GExt.ColFormats[3] := 'CB=PGCENTREFORMATION';
  end;
  if GInt <> nil then
  begin
    GInt.ColFormats[0] := 'CB=PGSTAGEFORMCOMPLET';
    GInt.ColTypes[1] := 'D';
    GInt.colformats[1] := ShortDateFormat;
    GInt.ColTypes[2] := 'D';
    GInt.colformats[2] := ShortDateFormat;
    GInt.ColFormats[3] := 'CB=PGCENTREFORMATION';
  end;
  //Compétences
  if GCompetEmp <> nil then
  begin
    GCompetEmp.ColFormats[0] := 'CB=PGRHCOMPETENCES';
    GCompetEmp.ColFormats[1] := '# ##0.00';
    GCompetEmp.ColAligns[1] := taRightJustify;
    GCompetEmp.ColFormats[2] := '# ##0.00';
    GCompetEmp.ColAligns[2] := taRightJustify;
    GCompetEmp.ColFormats[3] := 'CB=PGOUINONGRAPHIQUE';
    for i := 1 to 5 do
    begin
      Libelle := RechDom('GCZONELIBRE', 'RH' + IntToStr(i), False);
      if Libelle <> 'Table libre ' + IntToStr(i) then
      begin
        GCompetEmp.ColCount := GCompetEmp.ColCount + 1;
        GCompetEmp.ColFormats[GCompetEmp.ColCount - 1] := 'CB=PGTABLELIBRERH' + IntToStr(i);
        GCompetEmp.Titres.Add(Libelle);
        GCompetEmp.UpdateTitres;
      end;
    end;
  end;
  if GCompetSal <> nil then
  begin
    GCompetSal.ColFormats[0] := 'CB=PGRHCOMPETENCES';
    GCompetSal.ColFormats[1] := '# ##0.00';
    GCompetSal.ColAligns[1] := taRightJustify;
    for i := 1 to 5 do
    begin
      Libelle := RechDom('GCZONELIBRE', 'RH' + IntToStr(i), False);
      if Libelle <> 'Table libre ' + IntToStr(i) then
      begin
        GCompetSal.ColCount := GCompetSal.ColCount + 1;
        GCompetSal.ColFormats[GCompetSal.ColCount - 1] := 'CB=PGTABLELIBRERH' + IntToStr(i);
        GCompetSal.Titres.Add(Libelle);
        GCompetSal.UpdateTitres;
      end;
    end;
  end;
end;

procedure CarriereAfficheCompetences(Salarie, Emploi: string; GrilleEmploi, GrilleSalarie: THGrid);
var
  Competence, NiveauAtteint, Libelle: string;
  Q: TQuery;
  TobCompetencesSal, TS, TobCompetencesEmploi, TE, TobComparaison, TC: Tob;
  DegreMaitriseE, DegreMaitriseS: Double;
  a, i: Integer;
begin
  //        GrilleEmploi := THGrid(GetControl('GCOMPETEMPLOI'));
  //        GrilleSalarie := THGrid(GetControl('GCOMPETSAL'));
  //        Salarie := GetControlText('PSA_SALARIE');
  //        Emploi := GetControlText('PSA_LIBELLEEMPLOI');
  GrilleEmploi.ColFormats[0] := 'CB=PGRHCOMPETENCES';
  GrilleEmploi.ColFormats[1] := '# ##0.00';
  GrilleEmploi.ColAligns[1] := taRightJustify;
  GrilleEmploi.ColFormats[2] := '# ##0.00';
  GrilleEmploi.ColAligns[2] := taRightJustify;
  GrilleEmploi.ColFormats[3] := 'CB=PGOUINONGRAPHIQUE';
  for i := 1 to 5 do
  begin
    Libelle := RechDom('GCZONELIBRE', 'RH' + IntToStr(i), False);
    if Libelle <> 'Table libre ' + IntToStr(i) then
    begin
      GrilleEmploi.ColCount := GrilleEmploi.ColCount + 1;
      GrilleEmploi.ColFormats[GrilleEmploi.ColCount - 1] := 'CB=PGTABLELIBRERH' + IntToStr(i);
      GrilleEmploi.Titres.Add(Libelle);
      GrilleEmploi.UpdateTitres;
    end;
  end;
  GrilleSalarie.ColFormats[0] := 'CB=PGRHCOMPETENCES';
  GrilleSalarie.ColFormats[1] := '# ##0.00';
  GrilleSalarie.ColAligns[1] := taRightJustify;
  for i := 1 to 5 do
  begin
    Libelle := RechDom('GCZONELIBRE', 'RH' + IntToStr(i), False);
    if Libelle <> 'Table libre ' + IntToStr(i) then
    begin
      GrilleSalarie.ColCount := GrilleSalarie.ColCount + 1;
      GrilleSalarie.ColFormats[GrilleSalarie.ColCount - 1] := 'CB=PGTABLELIBRERH' + IntToStr(i);
      GrilleSalarie.Titres.Add(Libelle);
      GrilleSalarie.UpdateTitres;
    end;
  end;
  TobComparaison := Tob.Create('Comparatif', nil, -1);
  Q := OpenSQL('SELECT * FROM STAGEOBJECTIF LEFT JOIN RHCOMPETENCES ON POS_COMPETENCE=PCO_COMPETENCE WHERE POS_NATOBJSTAGE="EMP" ' +
    'AND POS_CODESTAGE="' + Emploi + '"', True);
  TobCompetencesEmploi := Tob.Create('emploi', nil, -1);
  TobCompetencesEmploi.LoadDetailDB('emploi', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT PCH_COMPETENCE,PCH_DEGREMAITRISE,PCO_TABLELIBRERH1,PCO_TABLELIBRERH2,PCO_TABLELIBRERH3,PCO_TABLELIBRERH4' +
    ',PCO_TABLELIBRERH5 FROM RHCOMPETRESSOURCE LEFT JOIN RHCOMPETENCES ON PCH_COMPETENCE=PCO_COMPETENCE WHERE PCH_TYPERESSOURCE="SAL" ' +
    'AND PCH_SALARIE="' + Salarie + '"', True);
  TobCompetencesSal := Tob.Create('Salarie', nil, -1);
  TobCompetencesSal.LoadDetailDB('Salarie', '', '', Q, False);
  Ferme(Q);
  for i := 0 to TobCompetencesEmploi.Detail.Count - 1 do
  begin
    Competence := TobCompetencesEmploi.Detail[i].GetValue('POS_COMPETENCE');
    DegreMaitriseE := TobCompetencesEmploi.Detail[i].GetValue('POS_DEGREMAITRISE');
    DegreMaitriseS := 0;
    TS := TobCompetencesSal.FindFirst(['PCH_COMPETENCE'], [Competence], False);
    while TS <> nil do
    begin
      if TS.GetValue('PCH_DEGREMAITRISE') > DegreMaitriseS then DegreMaitriseS := TS.GetValue('PCH_DEGREMAITRISE');
      TS := TobCompetencesSal.FindNext(['PCH_COMPETENCE'], [Competence], False);
    end;
    if DegreMaitriseS >= DegreMaitriseE then NiveauAtteint := 'OUI'
    else NiveauAtteint := 'NON';
    TC := Tob.Create('FilleCompare', TobComparaison, -1);
    TC.AddChampSupValeur('COMPETENCE', Competence, False);
    TC.AddChampSupValeur('MAITRISEEMP', DegreMaitriseE, False);
    TC.AddChampSupValeur('MAITRISESAL', DegreMaitriseS, False);
    TC.AddChampSupValeur('ATTEINT', NiveauAtteint, False);
    for a := 1 to 5 do
    begin
      Libelle := RechDom('GCZONELIBRE', 'RH' + IntToStr(a), False);
      if Libelle <> 'Table libre ' + IntToStr(a) then
      begin
        TC.AddChampSupValeur('LIBRE' + IntToStr(a), TobCompetencesEmploi.Detail[i].GetValue('PCO_TABLELIBRERH' + intToStr(a)), False);
      end;
    end;
  end;
  TobComparaison.PutGridDetail(GrilleEmploi, False, True, '', False);
  if TobComparaison.Detail.Count > 0 then GrilleEmploi.RowCount := TobComparaison.Detail.Count + 1
  else GrilleEmploi.RowCount := 2;
  TobComparaison.Free;
  TobComparaison := Tob.Create('Comparatif', nil, -1);
  for i := 0 to TobCompetencesSal.Detail.Count - 1 do
  begin
    Competence := TobCompetencesSal.Detail[i].GetValue('PCH_COMPETENCE');
    DegreMaitriseS := TobCompetencesSal.Detail[i].GetValue('PCH_DEGREMAITRISE');
    TE := TobCompetencesEmploi.FindFirst(['POS_COMPETENCE'], [Competence], False);
    if TE = nil then
    begin
      TC := Tob.Create('FilleCompare', TobComparaison, -1);
      TC.AddChampSupValeur('COMPETENCE', Competence, False);
      TC.AddChampSupValeur('MAITRISE', DegreMaitriseS, False);
      for a := 1 to 5 do
      begin
        Libelle := RechDom('GCZONELIBRE', 'RH' + IntToStr(a), False);
        if Libelle <> 'Table libre ' + IntToStr(a) then
        begin
          TC.AddChampSupValeur('LIBRE' + IntToStr(a), TobCompetencesSal.Detail[i].GetValue('PCO_TABLELIBRERH' + intToStr(a)), False);
        end;
      end;
    end;
  end;
  TobComparaison.PutGridDetail(GrilleSalarie, False, True, '', False);
  if TobComparaison.Detail.Count > 0 then GrilleSalarie.RowCount := TobComparaison.Detail.Count + 1
  else GrilleSalarie.RowCount := 2;
  TobCompetencesSal.Free;
  TobCompetencesEmploi.Free;
  TobComparaison.Free;
  //HMTrad.ResizeGridColumns(GrilleSalarie);
  //HMTrad.ResizeGridColumns(GrilleEmploi);
end;

procedure CarriereAfficheFormations(LeSalarie: string; GrilleInit, GrilleInterne, GrilleExterne: THGrid);
var
  Q: TQuery;
  TobFormationsInit, TobFormationsInterne, TobFormationsExt: Tob;
//  HMTrad: THSystemMenu;
  i: Integer;
begin
  if grilleInit <> nil then
  begin
    grilleInit.ColFormats[0] := 'CB=PGSTAGEFORMCOMPLET';
    grilleInit.ColTypes[1] := 'D';
    grilleInit.colformats[1] := ShortDateFormat;
    grilleInit.ColFormats[2] := 'CB=PGCENTREFORMATION';
    for i := 1 to GrilleInit.RowCount - 1 do
    begin
      GrilleInit.Rows[i].Clear;
    end;
    Q := OpenSQL('SELECT PFO_CODESTAGE,PFO_DATEFIN,PFO_CENTREFORMGU FROM FORMATIONS WHERE PFO_SALARIE="' + LeSalarie + '"' +
      'AND PFO_NATUREFORM="004"', True);
    TobFormationsInit := Tob.Create('LesFormationsInit', nil, -1);
    TobFormationsInit.LoadDetailDB('LesFormationsInit', '', '', Q, False);
    Ferme(Q);
    TobFormationsInit.PutGridDetail(GrilleInit, False, True, '', False);
    if TobFormationsInit.Detail.count > 0 then GrilleInit.RowCount := TobFormationsInit.Detail.count + 1
    else GrilleInit.RowCount := 2;
    TobFormationsInit.Free;
    //HMTrad.ResizeGridColumns(GrilleInit);
  end;
  if GrilleInterne <> nil then
  begin
    GrilleInterne.ColFormats[0] := 'CB=PGSTAGEFORMCOMPLET';
    GrilleInterne.ColTypes[1] := 'D';
    GrilleInterne.colformats[1] := ShortDateFormat;
    GrilleInterne.ColTypes[2] := 'D';
    GrilleInterne.colformats[2] := ShortDateFormat;
    GrilleInterne.ColFormats[3] := 'CB=PGCENTREFORMATION';
    for i := 1 to GrilleInterne.RowCount - 1 do
    begin
      GrilleInterne.Rows[i].Clear;
    end;
    Q := OpenSQL('SELECT PFO_CODESTAGE,PFO_DATEDEBUT,PFO_DATEFIN FROM FORMATIONS WHERE PFO_SALARIE="' + LeSalarie + '"' +
      'AND PFO_NATUREFORM="001"', True);
    TobFormationsInterne := Tob.Create('LesFormationsInit', nil, -1);
    TobFormationsInterne.LoadDetailDB('LesFormationsInit', '', '', Q, False);
    Ferme(Q);
    TobFormationsInterne.PutGridDetail(GrilleInterne, False, True, '', False);
    if TobFormationsInterne.Detail.count > 0 then GrilleInterne.RowCount := TobFormationsInterne.Detail.count + 1
    else GrilleInterne.RowCount := 2;
    TobFormationsInterne.Free;
    //HMTrad.ResizeGridColumns(GrilleInterne);
  end;
  if GrilleExterne <> nil then
  begin
    GrilleExterne.ColFormats[0] := 'CB=PGSTAGEFORMCOMPLET';
    GrilleExterne.ColTypes[1] := 'D';
    GrilleExterne.colformats[1] := ShortDateFormat;
    GrilleExterne.ColTypes[2] := 'D';
    GrilleExterne.colformats[2] := ShortDateFormat;
    GrilleExterne.ColFormats[3] := 'CB=PGCENTREFORMATION';
    for i := 1 to GrilleExterne.RowCount - 1 do
    begin
      GrilleExterne.Rows[i].Clear;
    end;
    Q := OpenSQL('SELECT PFO_CODESTAGE,PFO_DATEDEBUT,PFO_DATEFIN,PFO_CENTREFORMGU FROM FORMATIONS WHERE PFO_SALARIE="' + LeSalarie + '"' +
      'AND PFO_NATUREFORM="002"', True);
    TobFormationsExt := Tob.Create('LesFormationsInit', nil, -1);
    TobFormationsExt.LoadDetailDB('LesFormationsInit', '', '', Q, False);
    Ferme(Q);
    TobFormationsExt.PutGridDetail(GrilleExterne, False, True, '', False);
    if TobFormationsExt.Detail.count > 0 then GrilleExterne.RowCount := TobFormationsExt.Detail.count + 1
    else GrilleExterne.RowCount := 2;
    TobFormationsExt.Free;
    //HMTrad.ResizeGridColumns(GrilleExterne);
  end;
end;

procedure CarriereAfficheHisto(LeSalarie: string; DateDeb, DateFin: TDateTime; Grille: THGrid; THCriteres: THMultiValComboBox);
var
  Q: TQuery;
  TobHisto, TMC, TobMesChamps: Tob;
//  HMTrad: THSystemMenu;
  i, x, Num: Integer;
  DateModif: TDateTime;
  Valeur, Commentaire, LaTablette, LeChamp, LeChampBool, LeType, LeLibelle: string;
  LesCriteres, ChampCritere: string;
begin
  LesCriteres := THCriteres.Value;
  TobMesChamps := Tob.Create('LesCahmps', nil, -1);
  //CodeEmploi
  TMC := Tob.Create('TobMesCahmps', TobMesChamps, -1);
  TMC.AddchampSupValeur('LECHAMP', 'PHS_CODEEMPLOI', False);
  TMC.AddchampSupValeur('LECHAMPBOOL', 'PHS_BCODEEMPLOI', False);
  TMC.AddchampSupValeur('LATABLETTE', 'PGCODEPCSESE', False);
  TMC.AddchampSupValeur('LETYPE', 'S', False);
  TMC.AddchampSupValeur('LELIBELLE', 'Code emploi', False);
  TMC.AddchampSupValeur('PRISENCOMPTE', '-', False);
  //CodeEmploi
  TMC := Tob.Create('TobMesCahmps', TobMesChamps, -1);
  TMC.AddchampSupValeur('LECHAMP', 'PHS_LIBELLEEMPLOI', False);
  TMC.AddchampSupValeur('LECHAMPBOOL', 'PHS_BLIBELLEEMPLOI', False);
  TMC.AddchampSupValeur('LATABLETTE', 'PGLIBEMPLOI', False);
  TMC.AddchampSupValeur('LETYPE', 'S', False);
  TMC.AddchampSupValeur('LELIBELLE', 'Libelle emploi', False);
  TMC.AddchampSupValeur('PRISENCOMPTE', '-', False);
  //Qualification
  TMC := Tob.Create('TobMesCahmps', TobMesChamps, -1);
  TMC.AddchampSupValeur('LECHAMP', 'PHS_QUALIFICATION', False);
  TMC.AddchampSupValeur('LECHAMPBOOL', 'PHS_BQUALIFICATION', False);
  TMC.AddchampSupValeur('LATABLETTE', 'PGLIBQUALIFICATION', False);
  TMC.AddchampSupValeur('LETYPE', 'S', False);
  TMC.AddchampSupValeur('LELIBELLE', 'Qualification', False);
  TMC.AddchampSupValeur('PRISENCOMPTE', '-', False);
  //Coeff
  TMC := Tob.Create('TobMesCahmps', TobMesChamps, -1);
  TMC.AddchampSupValeur('LECHAMP', 'PHS_COEFFICIENT', False);
  TMC.AddchampSupValeur('LECHAMPBOOL', 'PHS_BCOEFFICIENT', False);
  TMC.AddchampSupValeur('LATABLETTE', 'PGLIBCOEFFICIENT', False);
  TMC.AddchampSupValeur('LETYPE', 'S', False);
  TMC.AddchampSupValeur('LELIBELLE', 'Coefficient', False);
  TMC.AddchampSupValeur('PRISENCOMPTE', '-', False);
  //Indice
  TMC := Tob.Create('TobMesCahmps', TobMesChamps, -1);
  TMC.AddchampSupValeur('LECHAMP', 'PHS_INDICE', False);
  TMC.AddchampSupValeur('LECHAMPBOOL', 'PHS_BINDICE', False);
  TMC.AddchampSupValeur('LATABLETTE', 'PGLIBINDICE', False);
  TMC.AddchampSupValeur('LETYPE', 'S', False);
  TMC.AddchampSupValeur('LELIBELLE', 'Indice', False);
  TMC.AddchampSupValeur('PRISENCOMPTE', '-', False);
  //Niveau
  TMC := Tob.Create('TobMesCahmps', TobMesChamps, -1);
  TMC.AddchampSupValeur('LECHAMP', 'PHS_NIVEAU', False);
  TMC.AddchampSupValeur('LECHAMPBOOL', 'PHS_BNIVEAU', False);
  TMC.AddchampSupValeur('LATABLETTE', 'PGLIBNIVEAU', False);
  TMC.AddchampSupValeur('LETYPE', 'S', False);
  TMC.AddchampSupValeur('LELIBELLE', 'Niveau', False);
  TMC.AddchampSupValeur('PRISENCOMPTE', '-', False);
  //Code Stat
  if VH_Paie.PGLibCodeStat <> '' then
  begin
    TMC := Tob.Create('TobMesCahmps', TobMesChamps, -1);
    TMC.AddchampSupValeur('LECHAMP', 'PHS_CODESTAT', False);
    TMC.AddchampSupValeur('LECHAMPBOOL', 'PHS_BCODESTAT', False);
    TMC.AddchampSupValeur('LATABLETTE', 'PGCODESTAT', False);
    TMC.AddchampSupValeur('LETYPE', 'S', False);
    TMC.AddchampSupValeur('LELIBELLE', VH_Paie.PGLibCodeStat, False);
    TMC.AddchampSupValeur('PRISENCOMPTE', '-', False);
  end;
  //Champ libres
  for Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    if Num > 4 then Break;
    TMC := Tob.Create('TobMesCahmps', TobMesChamps, -1);
    TMC.AddchampSupValeur('LECHAMP', 'PHS_TRAVAILN' + IntToStr(Num), False);
    TMC.AddchampSupValeur('LECHAMPBOOL', 'PHS_BTRAVAILN' + IntToStr(Num), False);
    TMC.AddchampSupValeur('LATABLETTE', 'PGTRAVAILN' + IntToStr(Num), False);
    TMC.AddchampSupValeur('LETYPE', 'S', False);
    TMC.AddchampSupValeur('LELIBELLE', GetParamSoc('SO_PGLIBORGSTAT' + IntToStr(Num)), False);
    TMC.AddchampSupValeur('PRISENCOMPTE', '-', False);
  end;
  //Salaires
  for Num := 1 to VH_Paie.PgNbSalLib do
  begin
    TMC := Tob.Create('TobMesCahmps', TobMesChamps, -1);
    TMC.AddchampSupValeur('LECHAMP', 'PHS_SALAIREMOIS' + IntToStr(Num), False);
    TMC.AddchampSupValeur('LECHAMPBOOL', 'PHS_BSALAIREMOIS' + IntToStr(Num), False);
    TMC.AddchampSupValeur('LATABLETTE', '', False);
    TMC.AddchampSupValeur('LETYPE', 'I', False);
    TMC.AddchampSupValeur('LELIBELLE', GetParamSoc('SO_PGSALLIB' + IntToStr(Num)) + ' mensuel', False);
    TMC.AddchampSupValeur('PRISENCOMPTE', '-', False);
  end;
  for Num := 1 to VH_Paie.PgNbSalLib do
  begin
    TMC := Tob.Create('TobMesCahmps', TobMesChamps, -1);
    TMC.AddchampSupValeur('LECHAMP', 'PHS_SALAIREANN' + IntToStr(Num), False);
    TMC.AddchampSupValeur('LECHAMPBOOL', 'PHS_BSALAIREANN' + IntToStr(Num), False);
    TMC.AddchampSupValeur('LATABLETTE', '', False);
    TMC.AddchampSupValeur('LETYPE', 'I', False);
    TMC.AddchampSupValeur('LELIBELLE', GetParamSoc('SO_PGSALLIB' + IntToStr(Num)) + ' annuel', False);
    TMC.AddchampSupValeur('PRISENCOMPTE', '-', False);
  end;
  //Horaires
  TMC := Tob.Create('TobMesCahmps', TobMesChamps, -1);
  TMC.AddchampSupValeur('LECHAMP', 'PHS_PGHHORHEBDO', False);
  TMC.AddchampSupValeur('LECHAMPBOOL', 'PHS_BPGHHORHEBDO', False);
  TMC.AddchampSupValeur('LATABLETTE', '', False);
  TMC.AddchampSupValeur('LETYPE', 'I', False);
  TMC.AddchampSupValeur('LELIBELLE', 'Horaire hebdomadaire', False);
  TMC.AddchampSupValeur('PRISENCOMPTE', '-', False);
  //Horaires
  TMC := Tob.Create('TobMesCahmps', TobMesChamps, -1);
  TMC.AddchampSupValeur('LECHAMP', 'PHS_PGHHORANNUEL', False);
  TMC.AddchampSupValeur('LECHAMPBOOL', 'PHS_BPGHHORANNUEL', False);
  TMC.AddchampSupValeur('LATABLETTE', '', False);
  TMC.AddchampSupValeur('LETYPE', 'I', False);
  TMC.AddchampSupValeur('LELIBELLE', 'Horaire annuel', False);
  TMC.AddchampSupValeur('PRISENCOMPTE', '-', False);
  //Statut professionnel DADS
  TMC := Tob.Create('TobMesCahmps', TobMesChamps, -1);
  TMC.AddchampSupValeur('LECHAMP', 'PHS_DADSPROF', False);
  TMC.AddchampSupValeur('LECHAMPBOOL', 'PHS_BDADSPROF', False);
  TMC.AddchampSupValeur('LATABLETTE', 'PGSPROFESSIONNEL', False);
  TMC.AddchampSupValeur('LETYPE', 'S', False);
  TMC.AddchampSupValeur('LELIBELLE', 'Statut professionnel DADS', False);
  TMC.AddchampSupValeur('PRISENCOMPTE', '-', False);
  //Statut catégoriel DADS
  TMC := Tob.Create('TobMesCahmps', TobMesChamps, -1);
  TMC.AddchampSupValeur('LECHAMP', 'PHS_DADSCAT', False);
  TMC.AddchampSupValeur('LECHAMPBOOL', 'PHS_BDADSCAT', False);
  TMC.AddchampSupValeur('LATABLETTE', 'PGSCATEGORIEL', False);
  TMC.AddchampSupValeur('LETYPE', 'S', False);
  TMC.AddchampSupValeur('LELIBELLE', 'Statut catégoriel DADS', False);
  TMC.AddchampSupValeur('PRISENCOMPTE', '-', False);
  //Condition d''emploi
  TMC := Tob.Create('TobMesCahmps', TobMesChamps, -1);
  TMC.AddchampSupValeur('LECHAMP', 'PHS_CONDEMPLOI', False);
  TMC.AddchampSupValeur('LECHAMPBOOL', 'PHS_BCONDEMPLOI', False);
  TMC.AddchampSupValeur('LATABLETTE', 'PGCONDEMPLOI', False);
  TMC.AddchampSupValeur('LETYPE', 'S', False);
{PT8
  TMC.AddchampSupValeur('LELIBELLE', 'Condition d''emploi', False);
}
  TMC.AddchampSupValeur('LELIBELLE', 'Caractéristique activité', False);
//FIN PT8
  TMC.AddchampSupValeur('PRISENCOMPTE', '-', False);
  if LesCriteres <> '' then
  begin
    while LesCriteres <> '' do
    begin
      ChampCritere := ReadTokenPipe(LesCriteres, ';');
      TMC := TobMesChamps.FindFirst(['LECHAMP'], ['PHS_' + ChampCritere], False);
      if TMC <> nil then TMC.PutValue('PRISENCOMPTE', 'X');
    end;
    TMC := TobMesChamps.FindFirst([''], [''], False);
    while TMC <> nil do
    begin
      if TMC.GetValue('PRISENCOMPTE') <> 'X' then TMC.Free;
      TMC := TobMesChamps.FindNext([''], [''], False);
    end;
  end;
  for i := 1 to Grille.RowCount - 1 do
  begin
    Grille.Rows[i].Clear;
  end;
  Q := OpenSQL('SELECT * FROM HISTOSALARIE WHERE PHS_SALARIE="' + LeSalarie + '" ' +
    'AND PHS_DATEEVENEMENT>="' + UsDateTime(DateDeb) + '" AND PHS_DATEEVENEMENT<="' + UsDateTime(DateFin) + '"', True);
  TobHisto := Tob.Create('Histo', nil, -1);
  TobHisto.LoadDetailDB('Histo', '', '', Q, False);
  Ferme(Q);
  TobHistorique := Tob.Create('Edition', nil, -1);
  for i := 0 to TobHisto.Detail.Count - 1 do
  begin
    DateModif := TobHisto.Detail[i].GetValue('PHS_DATEEVENEMENT');
    Commentaire := TobHisto.Detail[i].GetValue('PHS_COMMENTAIRE');
    if Copy(Commentaire, 1, 4) <> 'Init' then
    begin
      for x := 0 to TobMesChamps.Detail.Count - 1 do
      begin
        LeChamp := TobMesChamps.Detail[X].GetValue('LECHAMP');
        LeChampBool := TobMesChamps.Detail[X].GetValue('LECHAMPBOOL');
        LaTablette := TobMesChamps.Detail[X].GetValue('LATABLETTE');
        LeType := TobMesChamps.Detail[X].GetValue('LETYPE');
        LeLibelle := TobMesChamps.Detail[X].GetValue('LELIBELLE');
        if TobHisto.Detail[i].GetValue(LeChampBool) = 'X' then
        begin
          if LeType = 'S' then Valeur := TobHisto.Detail[i].Getvalue(LeChamp)
          else Valeur := FloatToStr(TobHisto.Detail[i].GetValue(LeChamp));
          if LaTablette <> '' then Valeur := RechDom(LaTablette, Valeur, False);
          CarriereAjoutLigneHisto(TobHisto.Detail[i], LeLibelle, Valeur, Commentaire, DateModif);
        end;
      end;
    end;
  end;
  TobMesChamps.Free;
  TobHisto.Free;
  if TobHistorique.Detail.count > 0 then Grille.RowCount := TobHistorique.Detail.count + 1
  else Grille.RowCount := 2;
  TobHistorique.PutGridDetail(Grille, False, True, '', False);
  TobHistorique.Free;
  //HMTrad.ResizeGridColumns(Grille);
end;

procedure CarriereAjoutLigneHisto(TobLigneHisto: Tob; Libelle, Valeur, Commentaire: string; DateModif: TDateTime);
var
  TE: Tob;
begin
  TE := Tob.Create('FilleEtat', TobHistorique, -1);
  TE.AddchampSupValeur('DATEMODIF', DateModif, False);
  TE.AddchampSupValeur('COMMENTAIRE', Commentaire, False);
  TE.AddchampSupValeur('LIBELLEMODIF', Libelle, False);
  TE.AddchampSupValeur('VALEURMODIF', Valeur, False);
end;

procedure CarriereAfficheContrats(LeSalarie: string; Grille: THGrid);
var
  Q: TQuery;
  TobContrats: Tob;
//  HMTrad: THSystemMenu;
  i: Integer;
begin
  Grille.ColFormats[0] := 'CB=PGTYPECONTRAT';
  Grille.ColTypes[1] := 'D';
  Grille.colformats[1] := ShortDateFormat;
  Grille.ColTypes[2] := 'D';
  Grille.colformats[2] := ShortDateFormat;
  for i := 1 to Grille.RowCount - 1 do
  begin
    Grille.Rows[i].Clear;
  end;
  Q := OpenSQL('SELECT PCI_TYPECONTRAT,PCI_DEBUTCONTRAT,PCI_FINCONTRAT FROM CONTRATTRAVAIL ' +
    'WHERE PCI_SALARIE="' + LeSalarie + '"', True);
  TobContrats := Tob.Create('LesContrats', nil, -1);
  TobContrats.LoadDetailDB('LesContrats', '', '', Q, False);
  Ferme(Q);
  TobContrats.PutGridDetail(Grille, False, True, '', False);
  if TobContrats.Detail.count > 0 then Grille.RowCount := TobContrats.Detail.count + 1
  else Grille.RowCount := 2;
  TobContrats.Free;
  //HMTrad.ResizeGridColumns(Grille);
end;


procedure CarriereHistoInit( Combo : THMultiValComboBox);
var Num : Integer;
begin
  if Not Assigned(Combo) then exit;
  Combo.Items.Add('Code emploi');
  Combo.Values.Add('CODEEMPLOI');
  Combo.Items.Add('Libelle emploi');
  Combo.Values.Add('LIBELLEEMPLOI');
  Combo.Items.Add('Qualification');
  Combo.Values.Add('QUALIFICATION');
  Combo.Items.Add('Coefficient');
  Combo.Values.Add('COEFFICIENT');
  Combo.Items.Add('Indice');
  Combo.Values.Add('INDICE');
  if VH_Paie.PGLibCodeStat <> '' then
  begin
    Combo.Items.Add(VH_Paie.PGLibCodeStat);
    Combo.Values.Add('CODESTAGE');
  end;
  for Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    if Num = 1 then
    begin
      Combo.Items.Add(VH_Paie.PGLibelleOrgStat1);
      Combo.Values.Add('TRAVAILN1');
    end;
    if Num = 2 then
    begin
      Combo.Items.Add(VH_Paie.PGLibelleOrgStat2);
      Combo.Values.Add('TRAVAILN2');
    end;
    if Num = 3 then
    begin
      Combo.Items.Add(VH_Paie.PGLibelleOrgStat3);
      Combo.Values.Add('TRAVAILN3');
    end;
    if Num = 4 then
    begin
      Combo.Items.Add(VH_Paie.PGLibelleOrgStat4);
      Combo.Values.Add('TRAVAILN4');
    end;
  end;
  for Num := 1 to VH_Paie.PgNbSalLib do
  begin
    Combo.Items.Add(GetParamSoc('SO_PGSALLIB' + IntToStr(Num)) + ' mensuel');
    Combo.Values.Add('SALAIREMOIS' + IntToStr(Num));
  end;
  for Num := 1 to VH_Paie.PgNbSalLib do
  begin
    Combo.Items.Add(GetParamSoc('SO_PGSALLIB' + IntToStr(Num)) + ' annuel');
    Combo.Values.Add('SALAIREANN' + IntToStr(Num));
  end;
  Combo.Items.Add('Horaire hebdomadaire');
  Combo.Values.Add('PGHHORHEBDO');
  Combo.Items.Add('Horaire annuel');
  Combo.Values.Add('PGHHORANNUEL');
  Combo.Items.Add('Statut professionnel DADS');
  Combo.Values.Add('DADSPROF');
  Combo.Items.Add('Statut catégoriel DADS');
  Combo.Values.Add('DADSCAT');
{PT8
  Combo.Items.Add('Condition d''emploi');
}
  Combo.Items.Add('Caractéristique activité');
//FIN PT8
  Combo.Values.Add('CONDEMPLOI');
end;

{***********A.G.L.***********************************************
Auteur  ...... : JL
Créé le ...... : 21/04/2005
Modifié le ... :   /  /    
Description .. : Fonction de calcul de l'allocation de formation pour les 
Suite ........ : heures de DFI hors temps de travail
Suite ........ : Alloc = 50% des salaires versés sur 12 derniers mois
Mots clefs ... : 
*****************************************************************}
Function PGCalcAllocationDIF(Salarie : String;NbHeuresNT : Double;DateSession : TDateTime) : Double;
var Q : TQuery;
    DateDebPaie : TDateTime;
    Salaires, NbHeuresPaie : Double;
begin
     DateDebPaie := PlusDate(DateSession , -1 , 'A');
     Q := OpenSQL('SELECT SUM(PPU_CNETAPAYER) SALAIRE,SUM(PPU_HEURESREELLES) HEURES FROM PAIEENCOURS '+
     'WHERE PPU_SALARIE="'+Salarie+'" AND PPU_DATEFIN>="'+UsDateTime(DateDebpaie)+'"',True);
     If Not Q.Eof then
     begin
         Salaires := Q.FindField('SALAIRE').AsFloat;
         NbHeuresPaie := Q.FindField('HEURES').AsFloat;
     end
     else
     begin
         Salaires := 0;
         NbHeuresPaie := 0;
     end;
     Ferme(Q);
     If NbHeuresPaie > 0 then Salaires := Salaires / NbHeuresPaie
     else Salaires := 0;
     Result := Salaires * NbHeuresNT;
end;

Function PGDIFGenereAbs(HeuresT,HeuresNonT,Jours : Double;Salarie,DJ,FJ,TypeConge,Etablissement : String;DateD,DateF : TDateTime) : Integer;
var TobAbs,TA : Tob;
    No_Ordre : Integer;
    Heures : Double;
    St : String;
begin
     Result := 0;
     Heures := HeuresT + HeuresNonT;
     If Heures = 0 then Exit;
     st := 'SELECT PCN_DATEDEBUTABS,PCN_DATEFINABS,PCN_TYPECONGE,PCN_TYPEMVT FROM ABSENCESALARIE ' +
    ' WHERE PCN_SALARIE = "' + Salarie + '" AND PCN_DEBUTDJ="' + DJ + '" AND PCN_FINDJ="' + FJ + '"' +
    ' AND (PCN_TYPECONGE = "PRI" OR ' +
    ' (PCN_TYPEMVT = "ABS" AND PCN_SENSABS="-"))' +
    ' AND ((PCN_DATEDEBUTABS >="' + usdatetime(DateD) + '" AND PCN_DATEDEBUTABS <= "' +
    usdatetime(DateF) + '")' +
    'OR (PCN_DATEFINABS >="' + usdatetime(DateD) + '" AND PCN_DATEDEBUTABS <= "' +
    usdatetime(DateF) + '")' +
    'OR (PCN_DATEFINABS >="' + usdatetime(DateF) + '" AND PCN_DATEDEBUTABS <= "' +
    usdatetime(DateD) + '")' +
    'OR(PCN_DATEFINABS <="' + usdatetime(DateF) + '" AND PCN_DATEDEBUTABS >= "' +
    usdatetime(DateD) + '"))';
    If ExisteSQL(St) then
    begin
         result := -1;
         Exit;
    end;
     TobAbs := TOB.Create('Lesabsences', Nil, -1);
     TA := TOB.Create('ABSENCESALARIE', TobAbs, -1);
     No_Ordre := IncrementeSeqNoOrdre('ABS',Salarie);
     InitialiseTobAbsenceSalarie(TA);
     TA.PutValue('PCN_SALARIE', Salarie);
     TA.PutValue('PCN_ORDRE', No_Ordre);
     TA.PutValue('PCN_TYPEMVT', 'ABS');
     TA.PutValue('PCN_TYPECONGE', TypeConge);
     TA.PutValue('PCN_SENSABS', '-');
     TA.PutValue('PCN_LIBELLE', 'Formation du '+DateToStr(DateD)+' au '+DateToStr(DateF));
     TA.PutValue('PCN_DATEVALIDITE', DateF);
     TA.PutValue('PCN_DATEDEBUTABS', DateD);
     TA.PutValue('PCN_DATEFINABS', DateF);
     TA.PutValue('PCN_JOURS', Jours);
     TA.PutValue('PCN_HEURES', Heures);
     TA.PutValue('PCN_ETABLISSEMENT', Etablissement);
     TA.PutValue('PCN_VALIDSALARIE','SAL');
     TA.PutValue('PCN_VALIDRESP','VAL');
     TA.PutValue('PCN_EXPORTOK','X');
     TA.PutValue('PCN_DEBUTDJ',DJ);
     TA.PutValue('PCN_FINDJ',FJ);
     TA.PutValue('PCN_MVTORIGINE','SAL');
     TA.InsertOrUpdateDB;
     TobAbs.Free;
end;

procedure PGDIFGenereRubEnTete(Heures : Double;Salarie,Etablissement,Rubrique,TypeAlim,Stage,Confidentiel : String;DateD,DateF,DatePaie : TDateTime);
var TobRub,TR : Tob;
    DateDebRub,DateFinRub : TDateTime;
begin
     DateDebRub := DebutDeMois(DatePaie);
     DateFinRub := FinDeMois(DatePaie);
     If Heures = 0 then Exit;
     TobRub := TOB.Create('LesRubriques', Nil, -1);
     TR := TOB.Create('HISTOSAISRUB', TobRub, -1);
     TR.PutValue('PSD_SALARIE', Salarie);
     TR.PutValue('PSD_ORDRE', 200);
     TR.PutValue('PSD_ORIGINEMVT', 'MHE');
     TR.PutValue('PSD_DATEDEBUT', DateDebRub);
     TR.PutValue('PSD_DATEFIN', DateFinRub);
     TR.PutValue('PSD_RUBRIQUE', Rubrique);
     TR.PutValue('PSD_LIBELLE', 'Droit Individuel à la Formation');
     TR.PutValue('PSD_RIBSALAIRE','' );
     TR.PutValue('PSD_BANQUEEMIS', '');
     TR.PutValue('PSD_TOPREGLE', '-');
     TR.PutValue('PSD_DATEPAIEMENT', Idate1900);
     TR.PutValue('PSD_ETABLISSEMENT', Etablissement);
     TR.PutValue('PSD_TYPALIMPAIE', TypeAlim);
     If typeAlim = 'BAS' then TR.PutValue('PSD_BASE', Heures)
     else TR.PutValue('PSD_BASE', 0);
     TR.PutValue('PSD_TAUX', 0);
     TR.PutValue('PSD_COEFF', 0);
     If typeAlim = 'MON' then TR.PutValue('PSD_MONTANT',Heures )
     else TR.PutValue('PSD_MONTANT',0 );
     TR.PutValue('PSD_DATEINTEGRAT', IDate1900);
     TR.PutValue('PSD_DATECOMPT', IDate1900);
     TR.PutValue('PSD_AREPORTER', 'NON');
     TR.PutValue('PSD_CONFIDENTIEL', Confidentiel);
     TR.PutValue('PSD_TOPCONVERT', '-');
     TR.InsertOrUpdateDB;
     TobRub.Free;
end;

procedure PGDIFGenereRub(HeuresT,HeuresNonT : Double;Salarie,Etablissement,Rubrique,TypeAlim,Stage,Confidentiel : String;DateD,DateF,DatePaie : TDateTime;NumOrdre : Integer);
var TobRub,TR : Tob;
    Heures : Double;
    DateDebRub,DateFinRub : TDateTime;
    LibCom : String;
begin
     DateDebRub := DebutDeMois(DatePaie);
     DateFinRub := FinDeMois(DatePaie);
     Heures := HeuresT + HeuresNonT;
     If Heures = 0 then Exit;
     TobRub := TOB.Create('LesRubriques', Nil, -1);
     LibCom := 'Formation du '+formatdateTime('dd/mm/yy',DateD) + ' au '+FormatDateTime('dd/mm/yy',DateF);
     TR := TOB.Create('HISTOSAISRUB', TobRub, -1);
     TR.PutValue('PSD_SALARIE', Salarie);
     TR.PutValue('PSD_ORDRE', NumOrdre);
     TR.PutValue('PSD_ORIGINEMVT', 'MLB');
     TR.PutValue('PSD_DATEDEBUT', DateDebRub);
     TR.PutValue('PSD_DATEFIN', DateFinRub);
     TR.PutValue('PSD_RUBRIQUE', Rubrique);
     TR.PutValue('PSD_LIBELLE', LibCom );
     TR.PutValue('PSD_RIBSALAIRE','' );
     TR.PutValue('PSD_BANQUEEMIS', '');
     TR.PutValue('PSD_TOPREGLE', '-');
     TR.PutValue('PSD_DATEPAIEMENT', Idate1900);
     TR.PutValue('PSD_ETABLISSEMENT', Etablissement);
     TR.PutValue('PSD_TYPALIMPAIE', TypeAlim);
     If typeAlim = 'BAS' then TR.PutValue('PSD_BASE', Heures)
     else TR.PutValue('PSD_BASE', 0);
     TR.PutValue('PSD_TAUX', 0);
     TR.PutValue('PSD_COEFF', 0);
     If typeAlim = 'MON' then TR.PutValue('PSD_MONTANT',Heures )
     else TR.PutValue('PSD_MONTANT',0 );
     TR.PutValue('PSD_DATEINTEGRAT', IDate1900);
     TR.PutValue('PSD_DATECOMPT', IDate1900);
     TR.PutValue('PSD_AREPORTER', 'NON');
     TR.PutValue('PSD_CONFIDENTIEL', Confidentiel);
     TR.PutValue('PSD_TOPCONVERT', '-');
     TR.InsertOrUpdateDB;
     TobRub.Free;
end;

procedure PGDIFSuppAbs(Salarie,TypeConge : String;DateDeb,DateFin : TDateTime);
begin
     ExecuteSQL('DELETE FROM ABSENCESALARIE WHERE '+
     'PCN_TYPECONGE="'+TypeConge+'" AND '+
     'PCN_SALARIE="'+Salarie+'" AND '+
     'PCN_DATEDEBUTABS="'+UsDateTime(DateDeb)+'" AND '+
     'PCN_DATEFINABS="'+UsDateTime(DateFin)+'"');
end;

procedure PGDIFSuppRub(Salarie,Rubrique,TypeAlim : String;DatePaie : TDateTime);
var DateDebRub,DateFinRub : TDateTime;
begin
     DateDebRub := DebutDeMois(DatePaie);
     DateFinRub := FinDeMois(DatePaie);
     ExecuteSQL('DELETE FROM HISTOSAISRUB WHERE '+
     '(PSD_ORIGINEMVT="MHE" OR PSD_ORIGINEMVT="MLB") AND '+
     'PSD_SALARIE="'+Salarie+'" AND '+
     'PSD_DATEDEBUT="'+UsDateTime(DateDebRub)+'" AND '+
     'PSD_DATEFIN="'+UsDateTime(DateFinRub)+'" AND '+
     'PSD_RUBRIQUE="'+Rubrique+'" AND '+
     'PSD_ORDRE>=200');
end;

procedure PGRecupereAllocFormation(DD,DF : TDateTime);
var Q : TQuery;
    TF,TobAllocations : Tob;
    SommeHeures,Heures,Taux,Alloc : Double;
    i,a : Integer;
    DateDeb,DateFin : TDateTime;
    Montant : Double;
    Salarie,Rubrique : String;
begin
//     DD := EncodeDate(annee,mois,1);
//     DF := FinDeMois(DD);
     Rubrique := GetParamSoc('SO_PGDIFRUBALLOC');
     Q := OpenSQL('SELECT SUM(PHB_MTREM) MONTANT,PHB_SALARIE,PHB_DATEDEBUT,PHB_DATEFIN FROM HISTOBULLETIN WHERE PHB_DATEFIN>="'+UsDateTime(DD)+'" AND PHB_DATEFIN<="'+UsDateTime(DF)+'" '+
     'AND PHB_RUBRIQUE="'+Rubrique+'" GROUP BY PHB_SALARIE,PHB_DATEDEBUT,PHB_DATEFIN',true);
     TobAllocations := Tob.Create('LesAllocs',Nil,-1);
     TobAllocations.LoadDetailDB('LesAllocs','','',Q,False);
     Ferme(Q);
     SommeHeures := 0;
     For a := 0 to TobAllocations.Detail.Count - 1 do
     begin
           Salarie := TobAllocations.Detail[a].GetString('PHB_SALARIE');
           Montant := TobAllocations.Detail[a].GetDouble('MONTANT');
           DateDeb := TobAllocations.Detail[a].GetDouble('PHB_DATEDEBUT');
           DateFin := TobAllocations.Detail[a].GetDouble('PHB_DATEFIN');
           Q := OpenSQL('SELECT * FROM FORMATIONS WHERE PFO_SALARIE="'+Salarie+'" AND PFO_DATEPAIE>="'+UsDateTime(DateDeb)+'" AND PFO_DATEPAIE<="'+UsDateTime(DateFin)+'" '+
           'AND PFO_TYPEPLANPREV="DIF" AND PFO_HTPSNONTRAV>0 AND PFO_EFFECTUE="X"',True);
           TF := Tob.Create('FORMATIONS',Nil,-1);
           TF.LoadDetailDB('FORMATIONS','','',Q,False);
           Ferme(Q);
           For i := 0 to TF.Detail.Count - 1 do
           begin
                SommeHeures := SommeHeures + TF.Detail[i].GetValue('PFO_HTPSNONTRAV');
           end;
           For i := 0 to TF.Detail.Count - 1 do
           begin
                Heures := TF.Detail[i].GetValue('PFO_HTPSNONTRAV');
                If SommeHeures <> 0 then Taux := Heures / SommeHeures
                else Taux := 0;
                Alloc := Montant * taux;
                TF.Detail[i].PutValue('PFO_ALLOCFORM',Alloc);
                TF.Detail[i].PutValue('PFO_PAYER','X');
                TF.Detail[i].UpdateDB;
           end;
           TF.Free;
     end;
     TobAllocations.Free;
end;

procedure RecupereLesFormationsDIF(Salarie,TypePlan,Rubrique,TypeAlim : String;DateVAL : TDateTime;Travail,HorsTravail : Boolean;LeStage : String = ''; Ordre : Integer = 0);
var Q : TQuery;
    DD,DF : TDateTime;
    TobFormations : Tob;
    i,NumOrdre : Integer;
    HeuresTAlim,HeuresNonTAlim,NbhT,NbhNonT : double;
    Etab,DJ,FJ,Confidentiel,Stage,WhereH,WhereFormation : String;
    dateD,dateF : TDateTime;
    DatePaie : TdateTime;
    TotalHeures : Double;

begin
     DD := DebutDeMois(DateVal);
     DF := FinDeMois(DateVal);
     DateD := IDate1900;
     DateF := IDate1900;
     DatePaie := IDate1900;
     WhereH := '';
     NumOrdre := 200;
     TotalHeures := 0;
     If (Travail) and not (HorsTravail) then WhereH := ' AND PFO_HTPSTRAV>0';
     If (HorsTravail) and not (Travail) then WhereH := ' AND PFO_HTPSNONTRAV>0';
     If (LeStage <> '') and (Ordre > 0) then WhereFormation := ' AND (PFO_CODESTAGE<>"'+Stage+'" AND PFO_ORDRE<>'+IntToStr(Ordre)+')'
     else WhereFormation := '';
     Q := OpenSQL('SELECT * FROM FORMATIONS WHERE PFO_SALARIE="'+Salarie+'" AND  PFO_DATEPAIE>="'+UsDateTime(DD)+'" AND PFO_DATEPAIE<="'+UsDateTime(DF)+'" '+
     ' AND PFO_TYPEPLANPREV="'+TypePlan+'" AND PFO_EFFECTUE="X"'+WhereH+WhereFormation,True);
     TobFormations := Tob.Create('LesFormations',Nil,-1);
     TobFormations.LoadDetailDB('LesFormations','','',Q,False);
     Ferme(Q);
     For i := 0 to TobFormations.Detail.Count - 1 do
     begin
          NbhT := TobFormations.Detail[i].GetValue('PFO_HTPSTRAV');
          NbhNonT := TobFormations.Detail[i].GetValue('PFO_HTPSNONTRAV');
          Salarie := TobFormations.Detail[i].GetValue('PFO_SALARIE');
          Etab := TobFormations.Detail[i].GetValue('PFO_ETABLISSEMENT');
          DateD := TobFormations.Detail[i].GetValue('PFO_DATEDEBUT');
          DateF := TobFormations.Detail[i].GetValue('PFO_DATEFIN');
          DJ := TobFormations.Detail[i].GetValue('PFO_DEBUTDJ');
          FJ := TobFormations.Detail[i].GetValue('PFO_FINDJ');
          DatePaie := TobFormations.Detail[i].GetValue('PFO_DATEPAIE');
          Confidentiel := TobFormations.Detail[i].GetValue('PFO_CONFIDENTIEL');
          Stage := TobFormations.Detail[i].GetValue('PFO_CODESTAGE');
          HeuresTAlim := 0;
          HeuresNonTAlim := 0;
          If Travail then HeuresTAlim := NbhT;
          If HorsTravail then HeuresNonTAlim := NbhNonT;
          TotalHeures := TotalHeures + HeuresNonTAlim + HeuresTAlim;
          PGDIFGenereRub(HeuresTAlim,HeuresNonTAlim,Salarie,Etab,Rubrique,TypeAlim,Stage,Confidentiel,DateD,DateF,DatePaie,NumOrdre);
          NumOrdre := NumOrdre + 1;
     end;
     PGDIFGenereRubEnTete(TotalHeures,Salarie,Etab,Rubrique,TypeAlim,Stage,Confidentiel,DateD,DateF,DatePaie);
     TobFormations.Free;
end;

procedure ReaffectPlafondForm;
var TPlafond,TNew,T : Tob;
    i : Integer;
    Q : TQuery;
    Guid : String;
begin
     If ExisteSQL('SELECT PFP_ORGCOLLECTGU FROM FRAISSALPLAF WHERE PFP_ORGCOLLECT >= 0 AND (PFP_ORGCOLLECTGU IS NULL OR PFP_ORGCOLLECTGU="")') then
     begin
          TPlafond := Tob.Create('FRAISSALPLAF',Nil,-1);
          TPlafond.LoadDetailDB('FRAISSALPLAF', '', '', Nil, False);
          For i := TPlafond.Detail.Count - 1 downto 0 do
          begin
               Guid := '';
               TNew := Tob.Create('FRAISSALPLAF', Nil, -1);
               T := TPlafond.Detail[i];
               TNew.Dupliquer(T, True, True, True);
               Q := OpenSQL('SELECT ANN_GUIDPER FROM ANNUAIRE WHERE ANN_CODEPER='+IntToStr(T.GetValue('PFP_ORGCOLLECT')),True);
               If not Q.Eof then Guid := Q.FindField('ANN_GUIDPER').AsString;
               Ferme(Q);
               TNew.PutValue('PFP_ORGCOLLECTGU',Guid);
               TNew.insertOrUpdateDB(False);
               T.DeleteDB;
               TNew.free;
          end;
          TPlafond.Free;
     end;
end;

Procedure MajCoutPrev(Stage,Millesime : String;TF : TTableFiltre);
var Q : TQuery;
    i : Integer;
    CoutFonc,CoutUnit,CoutAnim,NbMax,NbDivD,NbDivE : Double;
    NbInsc,NbCout,CoutFoncSta,CoutAnimSta : Double;
    NbInscStage : Integer;
    TobInsc     : TOB;
begin
     CoutFonc := 0;
     CoutUnit := 0;
     CoutAnim := 0;
     NbMax := 0;

     Q := OpenSQL('SELECT PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR,PST_NBSTAMAX,PST_DUREESTAGE,PST_JOURSTAGE FROM STAGE ' +
     'WHERE PST_CODESTAGE="' + Stage + '" AND PST_MILLESIME="' + Millesime + '"', True);
     If Not Q.eof then
     begin
          CoutFonc := Q.FindField('PST_COUTBUDGETE').AsFloat;
          CoutUnit := Q.FindField('PST_COUTUNITAIRE').AsFloat;
          CoutAnim := Q.FindField('PST_COUTSALAIR').AsFloat;
          NbMax := Q.FindField('PST_NBSTAMAX').AsInteger;
     end;
     Ferme(Q);
     Q := OpenSQL('SELECT SUM(PFI_NBINSC) NBINSC FROM INSCFORMATION ' +
                          'WHERE PFI_CODESTAGE="' + Stage + '" AND PFI_MILLESIME="' + Millesime + '" ' +
                          'AND (PFI_ETATINSCFOR="ATT" OR PFI_ETATINSCFOR="VAL")', True);
     if not Q.Eof then NbInscStage := Q.FindField('NBINSC').AsInteger
     else NbInscStage := 1;
     Ferme(Q);
     
     
     //PT31 - Début
     // Plus de traitement à partir du TF pour que tous les salariés soient traités, même ceux qui ne sont pas visibles (confidentiels)
     (*TF.DisableTom;
     TF.StartUpdate;
     for i:=1 to TF.LaGrid.RowCount - 1 do
     begin
          TF.SelectRecord(i);
          MoveCur(False);
          NbInsc  := TF.GetValue('PFI_NBINSC');
          if NbMax > 0 then
          begin
               NbDivD := NbInscStage / NbMax;
               NbDivE := arrondi(NbDivD, 0);
               if NbDivE - NbDivD < 0 then NbCout := NbDivE + 1
               else NbCout := NbDivE;
          end
          else NbCout := 1;
          CoutFoncSta := (NbCout * CoutFonc) / NbInscStage; //PT5
          CoutAnimSta := (NbCout * CoutAnim) / NbInscStage;
          If (TF.GetValue('PFI_ETATINSCFOR')='ATT') or (TF.GetValue('PFI_ETATINSCFOR')='VAL') then TF.PutValue('PFI_AUTRECOUT',NbInsc * (CoutFoncSta + CoutAnimSta + Coutunit))
          else TF.PutValue('PFI_AUTRECOUT',0);
          TF.Post;
     end;
     TF.EndUpdate;
     TF.EnableTom;*)
     TobInsc := TOB.Create('~LesInscrits', Nil, -1);
     TobInsc.LoadDetailDBFromSQL('INSCFORMATION','SELECT * FROM INSCFORMATION WHERE PFI_CODESTAGE="' + Stage + '" AND PFI_MILLESIME="' + Millesime + '"');
     
     for i:=0 to TobInsc.Detail.Count - 1 do
     begin
          NbInsc  := TobInsc.Detail[i].GetValue('PFI_NBINSC');
          if NbMax > 0 then
          begin
               NbDivD := NbInscStage / NbMax;
               NbDivE := arrondi(NbDivD, 0);
               if NbDivE - NbDivD < 0 then NbCout := NbDivE + 1
               else NbCout := NbDivE;
          end
          else NbCout := 1;
          CoutFoncSta := (NbCout * CoutFonc) / NbInscStage; //PT5
          CoutAnimSta := (NbCout * CoutAnim) / NbInscStage;
          If (TobInsc.Detail[i].GetValue('PFI_ETATINSCFOR')='ATT') or (TobInsc.Detail[i].GetValue('PFI_ETATINSCFOR')='VAL') then 
          	TobInsc.Detail[i].PutValue('PFI_AUTRECOUT',NbInsc * (CoutFoncSta + CoutAnimSta + Coutunit))
          else 
          	TobInsc.Detail[i].PutValue('PFI_AUTRECOUT',0);
     end;
     
     TobInsc.UpdateDB;
     FreeAndNil(TobInsc);
     
     TF.RefreshLignes;
     //PT31 - Fin
end;

Function RendMillesimeEManager : String;
var Q : TQuery;
begin
     Result := '0000';
     Q := OpenSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE PFE_ENCOURS="X" ORDER BY PFE_MILLESIME DESC',True);
     If Not Q.Eof then Result := Q.FindField('PFE_MILLESIME').AsString;
     Ferme(Q);
end;

Procedure MajSousSessionSal(Salarie,Stage,NMultiSession : String);
var Q : TQuery;
    TobSousSession,TobEnTete : Tob;
    i : Integer;
begin
  ExecuteSQL('DELETE FROM FORMATIONS WHERE PFO_SALARIE="'+Salarie+'" AND PFO_ORDRE='+NMultiSession+' AND PFO_CODESTAGE="'+Stage+'"');
  Q := OpenSQL('SELECT PFO_TYPEPLANPREV,SUM(PFO_NBREHEURE) NBHEURES,SUM(PFO_HTPSTRAV) SURTT,SUM(PFO_HTPSNONTRAV) HORSTT FROM FORMATIONS '+
   'WHERE PFO_SALARIE="'+Salarie+'" AND PFO_NOSESSIONMULTI='+NMultiSession+' AND PFO_CODESTAGE="'+Stage+'" GROUP BY PFO_TYPEPLANPREV',True);
  TobSousSession := Tob.Create('SommeSousSession',Nil,-1);
  TobSousSession.LoadDetailDB('SommeSousSession','','',Q,False);
  Ferme(Q);
  For i := 0 to TobSousSession.Detail.Count - 1 do
  begin
    TobEnTete := Tob.Create('FORMATIONS',Nil,-1);
    TobEnTete.PutValue('PFO_PGTYPESESSION','ENT');
    TobEnTete.PutValue('PFO_MILLESIME','0000');
    TobEnTete.PutValue('PFO_CODESTAGE',Stage);
    TobEnTete.PutValue('PFO_ORDRE',NMultiSession);
    TobEnTete.PutValue('PFO_SALARIE',Salarie);
    TobEnTete.PutValue('PFO_EFFECTUE','X');
    TobEnTete.PutValue('PFO_TYPEPLANPREV',TobSousSession.Detail[i].GetValue('PFO_TYPEPLANPREV'));
    TobEnTete.PutValue('PFO_NBREHEURE',TobSousSession.Detail[i].GetValue('NBHEURES'));
    TobEnTete.PutValue('PFO_HTPSTRAV',TobSousSession.Detail[i].GetValue('SURTT'));
    TobEnTete.PutValue('PFO_HTPSNONTRAV',TobSousSession.Detail[i].GetValue('HORSTT'));
    TobEnTete.InserTDB(Nil);
    FreeAndNil(TobEnTete);
  end;
  TobSousSession.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 16/05/2007 / PT6
Modifié le ... :   /  /    
Description .. : Retourne la population d'un salarié
Mots clefs ... : 
*****************************************************************}
function DonnePopulation(Salarie: String): String;
Var Q : TQuery;
begin
     Result := '';

     Q := OpenSQL('SELECT PNA_POPULATION FROM SALARIEPOPUL WHERE PNA_TYPEPOP="'+TYP_POPUL_FORM_PREV+'" AND PNA_SALARIE="'+Salarie+'"', True);
     If Not Q.EOF Then Result := Q.FindField('PNA_POPULATION').AsString;
     Ferme(Q);
end;

//PT10 - Début
{***********************************************
Fonction permettant d'adpater une requête SQL Emanager selon
qu'on soit responsable ou assistant.
***********************************************}
Function AdaptByRespEmanager (TypeUtilisateur,Prefixe,Resp : String; ServicesSup : Boolean) : String;
Begin
	Result := AdaptByRespEmanager (TypeUtilisateur,Prefixe,Resp,'FOR',ServicesSup); //PT10
End;
//PT9 - Fin


{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 03/03/2008 / PT10
Modifié le ... :   /  /
Description .. : Fonction permettant d'adpater une requête SQL Emanager selon
Suite ........ : qu'on soit responsable ou assistant de manière générique (avec suffixe).
Mots clefs ... :
*****************************************************************}
Function AdaptByRespEmanager (TypeUtilisateur,Prefixe,Resp,Suffixe : String; ServicesSup : Boolean) : String;
Begin
  If TypeUtilisateur = 'R' then
  Begin
    If Not ServicesSup Then
      Result := ' ('+Prefixe+'_RESPONS'+Suffixe+'="'+Resp+'")'
    Else
      Result := ' ('+Prefixe+'_RESPONS'+Suffixe+'="'+Resp+'" OR '+Prefixe+'_RESPONS'+Suffixe+' IN '+
      '(SELECT PGS_RESPONS'+Suffixe+' FROM SERVICES WHERE PGS_CODESERVICE IN '+
      '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND'+
      ' PGS_RESPONS'+Suffixe+'="'+Resp+'")))';
  End
  Else
  Begin
    If Not ServicesSup Then
      Result := ' ('+Prefixe+'_RESPONS'+Suffixe+' IN (SELECT PGS_RESPONS'+Suffixe+' FROM SERVICES WHERE PGS_SECRETAIRE'+Suffixe+'="'+Resp+'"))'
    Else
      Result := ' ('+Prefixe+'_RESPONS'+Suffixe+' IN '+
      '(SELECT PGS_RESPONS'+Suffixe+' FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
      'WHERE PGS_SECRETAIRE'+Suffixe+'="'+Resp+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIRE'+Suffixe+'="'+Resp+'")))';
  End;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 21/02/2008 / PT14
Modifié le ... :   /  /
Description .. : Met à jour les tables BUDGETPAIE, INSCFORMATION et
Suite ........ : FORMATIONS avec le nouveau responsable d'un salarié
Suite ........ : suite à une mise à jour de la hierarchie.
Mots clefs ... : REPONSABLE;VARIABLE;FORMATION
*****************************************************************}
Procedure MAJResponsable (Service,TypeResp,NouveauResp : String);
Begin
    If TypeResp = 'VAR' Then
    Begin
		ExecuteSQL( 'UPDATE BUDGETPAIE SET PBG_RESPONSVAR="'+NouveauResp+'" '+
                    'WHERE PBG_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_CODESERVICE="'+Service+'")'); // AND PBG_ANNEE="'+GetParamSoc('SO_PGAUGANNEE')+'"');
        AvertirTable('PGRESPONSVARIABLE');
	End
    Else If TypeResp = 'FOR' Then
	Begin
		ExecuteSQL( 'UPDATE FORMATIONS SET PFO_RESPONSFOR="'+NouveauResp+'" '+
                    'WHERE PFO_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_CODESERVICE="'+Service+'")');

		ExecuteSQL( 'UPDATE INSCFORMATION SET PFI_RESPONSFOR="'+NouveauResp+'" '+
                    'WHERE PFI_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_CODESERVICE="'+Service+'")');
    End;
End;


{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 19/02/2008 / PT13
Modifié le ... :   /  /
Description .. : Gestion de l'elipsis salarié en cas de partage des
Suite ........ : dossiers formation.
Mots clefs ... :
*****************************************************************}
Procedure ElipsisSalarieMultidos (Sender : TObject; Where : String = ''); 
var StWhere,StOrder,StListe : String;
Begin
	StListe := 'PSI_LIBELLE,PSI_PRENOM';
	StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT")';
	If Not PGDroitMultiForm Then
		StWhere := StWhere + ' AND PSI_NODOSSIER="'+V_PGI.NoDossier+'"'
	Else
	Begin
		StListe := StListe + ',DOS_LIBELLE';
		If V_PGI.ModePCL='1' Then StWhere := StWhere + ' AND PSI_NODOSSIER IN (SELECT DOS_NODOSSIER FROM DOSSIER WHERE '+GererCritereGroupeConfTous+')'; //PT19
	End;
	StWhere := StWhere + ' AND ' + Where;
    StOrder := 'PSI_INTERIMAIRE';
    StWhere := RecupClauseHabilitationLookupList(StWhere);
    StWhere := StWhere + ' AND PSI_CONFIDENTIEL<="'+RendNivAccesConfidentiel()+'"'; //PT31

    LookupList(THEdit(Sender),'Liste des salariés','INTERIMAIRES LEFT JOIN DOSSIER ON PSI_NODOSSIER=DOS_NODOSSIER','PSI_INTERIMAIRE',StListe,StWhere,StOrder, True,-1);
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 19/02/2008 / PT13
Modifié le ... :   /  /
Description .. : Gestion de l'elipsis du Responsable en cas de partage des
Suite ........ : dossiers formation.
Mots clefs ... :
*****************************************************************}
Procedure ElipsisResponsableMultidos (Sender : TObject; Where : String = ''); 
var StWhere,StOrder,StListe : String;
Begin
	StListe := 'PSI_LIBELLE,PSI_PRENOM';
	If Where <> '' Then StWhere := Where
	Else
	Begin
		StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES)';
		If Not PGDroitMultiForm Then
			StWhere := StWhere + ' AND PSI_NODOSSIER="'+V_PGI.NoDossier+'"'
		Else
		Begin
			StListe := StListe + ',DOS_LIBELLE';
			If V_PGI.ModePCL='1' Then StWhere := StWhere + ' AND PSI_NODOSSIER IN (SELECT DOS_NODOSSIER FROM DOSSIER WHERE '+GererCritereGroupeConfTous+')'; //PT19
		End;
	End;
	
    StOrder := 'PSI_INTERIMAIRE';
    StWhere := RecupClauseHabilitationLookupList(StWhere);
    StWhere := StWhere + ' AND PSI_CONFIDENTIEL<="'+RendNivAccesConfidentiel()+'"'; //PT31

    LookupList(THEdit(Sender),'Liste des responsables de formation','INTERIMAIRES  LEFT JOIN DOSSIER ON PSI_NODOSSIER=DOS_NODOSSIER','PSI_INTERIMAIRE',StListe,StWhere,StOrder, True,-1);
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 26/05/2008 / PT28
Modifié le ... :   /  /
Description .. : Met à jour les compétences du salarié suite à la réalisation d'une formation
Mots clefs ... :
*****************************************************************}
procedure MajCompetences (Salarie,Stage,Session,Millesime : String; Present : Boolean);
var Competence : String;
    Q : TQuery;
    IMax,i : Integer;
    TobCompetStage,T : Tob;
    DateValid : TDateTime;
begin
        Q := OpenSQL('SELECT PSS_DATEFIN FROM SESSIONSTAGE WHERE PSS_ORDRE='+Session+' AND PSS_CODESTAGE="'+Stage+'" AND PSS_MILLESIME="'+Millesime+'"',True);
        If Not Q.Eof then DateValid := Q.FindField('PSS_DATEFIN').AsDateTime
        Else DateValid := IDate1900;
        Ferme(Q);

        If Present = True then
        begin
                Q := OpenSQL('SELECT * FROM STAGEOBJECTIF WHERE POS_NATOBJSTAGE="FOR" AND POS_CODESTAGE="'+Stage+'"',True);
                TobCompetStage := Tob.Create('CompetencesStages',Nil,-1);
                TobCompetStage.LoadDetailDB('CompetencesStages','','',Q,False);
                Ferme(Q);

                For i := 0 to TobCompetStage.Detail.Count - 1 do
                begin
                        Competence := TobCompetStage.Detail[i].Getvalue('POS_COMPETENCE');
                        T := Tob.Create('RHCOMPETRESSOURCE',Nil,-1);
                        Q:=OpenSQL('SELECT MAX(PCH_RANG) FROM RHCOMPETRESSOURCE WHERE PCH_TYPERESSOURCE="SAL" AND PCH_SALARIE="'+Salarie+'"',TRUE) ;
                        if Not Q.EOF then imax:=Q.Fields[0].AsInteger+1 else iMax:=1 ;
                        Ferme(Q) ;

                        T.PutValue('PCH_TYPERESSOURCE','SAL');
                        T.PutValue('PCH_SALARIE',Salarie);
                        T.PutValue('PCH_COMPETENCE',Competence);
                        T.PutValue('PCH_RANG',IMax);
                        T.PutValue('PCH_CODESTAGE',Stage);
                        T.PutValue('PCH_SESSIONSTAGE',StrToInt(Session));
                        T.PutValue('PCH_DATEVALIDATION',DateValid);
                        T.PutValue('PCH_DEGREMAITRISE',TobCompetStage.Detail[i].GetValue('POS_DEGREMAITRISE'));
                        T.PutValue('PCH_TABLELIBRECR1',TobCompetStage.Detail[i].GetValue('POS_TABLELIBRECR1'));
                        T.PutValue('PCH_TABLELIBRECR2',TobCompetStage.Detail[i].GetValue('POS_TABLELIBRECR2'));
                        T.PutValue('PCH_TABLELIBRECR3',TobCompetStage.Detail[i].GetValue('POS_TABLELIBRECR3'));
                        T.InsertOrUpdateDB(False);
                        T.Free;
                end;
                TobCompetStage.Free;
        end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 26/05/2008 / PT28
Modifié le ... :   /  /
Description .. : Génère une absence pour le DIF
Mots clefs ... :
*****************************************************************}
procedure InsererDIFBulletin (TobInsc : TOB; InitDatePaie : TDateTime);
var Q : TQuery;
    NbhT,NbhNonT : Double;
    TypeFor : String;
    TobMvtDIF : Tob;
    i : Integer;
    Travail,HorsTravail : String;
    Rubrique,TypeConge,TypeAlim : String;
    Salarie,Etab : String;
    DateD,DateF : TDateTime;
    DJ,FJ : String;
    HeuresTAlim,HeuresNonTAlim,Jours : Double;
    AbsOk : Integer;
    TypeForm,Predef : String;
    SuppAbsence,SuppRubrique : Boolean;
begin
//     If (InitTypePlan <> GetField('PFO_TYPEPLANPREV')) or (GetField('PFO_DATEPAIE') <> InitDatePaie) then SupprimerDIF(True,SuppAbsence,SuppRubrique)
//     else SupprimerDIF(False,SuppAbsence,SuppRubrique);
     SupprimerDIF(TobInsc,InitDatePaie,((TobInsc.IsFieldModified('PFO_TYPEPLANPREV')) or (TobInsc.GetValue('PFO_DATEPAIE') <> InitDatePaie)),SuppAbsence,SuppRubrique);
     If TobInsc.GetValue('PFO_EFFECTUE') = '-' then Exit;
     NbhT := TobInsc.GetValue('PFO_HTPSTRAV');
     NbhNonT := TobInsc.GetValue('PFO_HTPSNONTRAV');
     TypeFor := TobInsc.GetValue('PFO_TYPEPLANPREV');
     
     Q := OpenSQL('SELECT PSS_JOURSTAGE FROM SESSIONSTAGE WHERE PSS_ORDRE='+IntToStr(TobInsc.GetValue('PFO_ORDRE'))+' AND PSS_CODESTAGE="'+TobInsc.GetValue('PFO_CODESTAGE')+'" AND PSS_MILLESIME="'+TobInsc.GetValue('PFO_MILLESIME')+'"',True);
     If Not Q.Eof then Jours := Q.FindField('PSS_JOURSTAGE').AsFloat
     Else Jours := IDate1900;
     Ferme(Q);
     
     Q := OpenSQL('SELECT * FROM PARAMFORMABS WHERE ##PPF_PREDEFINI## PPF_TYPEPLANPREV="'+TypeFor+'"',True);
     TobMvtDIF := Tob.Create('Lesmvts',Nil,-1);
     TobMvtDIF.LoadDetailDB('Lesmvts','','',Q,False);
     Ferme(Q);
     
     Salarie := TobInsc.GetValue('PFO_SALARIE');
     Etab := TobInsc.GetValue('PFO_ETABLISSEMENT');
     DateD := TobInsc.GetValue('PFO_DATEDEBUT');
     DateF := TobInsc.GetValue('PFO_DATEFIN');
     DJ := TobInsc.GetValue('PFO_DEBUTDJ');
     FJ := TobInsc.GetValue('PFO_FINDJ');
     
     For i := 0 TO TobMvtDIF.Detail.Count - 1 do
     begin
          Travail := TobMvtDIF.Detail[i].GetValue('PPF_CHEURETRAV');
          HorsTravail := TobMvtDIF.Detail[i].GetValue('PPF_CHEURENONTRAV');
          Predef := TobMvtDIF.Detail[i].GetValue('PPF_PREDEFINI');
          If Predef <> 'DOS' then
          begin
               TypeForm := TobMvtDIF.Detail[i].GetValue('PPF_TYPEPLANPREV');
               If TobMvtDIF.FindFirst(['PPF_PREDEFINI','PPF_CHEURETRAV','PPF_CHEURENONTRAV','PPF_TYPEPLANPREV'],['DOS',Travail,HorsTravail,TypeForm],False) <> Nil then Continue;
               If Predef = 'CEG' then
               begin
                    If TobMvtDIF.FindFirst(['PPF_PREDEFINI','PPF_CHEURETRAV','PPF_CHEURENONTRAV','PPF_TYPEPLANPREV'],['STD',Travail,HorsTravail,TypeForm],False) <> Nil then Continue;
               end;
          end;
          HeuresTAlim := 0;
          HeuresNonTAlim := 0;
          If Travail = 'X' then HeuresTAlim := NbhT;
          If HorsTravail = 'X' then HeuresNonTAlim := NbhNonT;
          If (TobMvtDIF.Detail[i].GetValue('PPF_ALIMABS') = 'X') and (SuppAbsence) and (SuppRubrique) then
          begin
               TypeConge := TobMvtDIF.Detail[i].GetValue('PPF_TYPECONGE');
               AbsOk := PGDIFGenereAbs(HeuresTAlim,HeuresNonTAlim,Jours,Salarie,DJ,FJ,TypeConge,Etab,DateD,DateF);
               If AbsOk = -1 then
               begin
                    PGIBox('Impossible de créer l''absence correspondante car le salarié possède déjà une absence sur cette période');
                    Exit;
               end;

          end;
          If (TobMvtDIF.Detail[i].GetValue('PPF_ALIMRUB') = 'X')  and (SuppAbsence) and (SuppRubrique) then
          begin
               Rubrique := TobMvtDIF.Detail[i].GetValue('PPF_RUBRIQUE');
               Travail := TobMvtDIF.Detail[i].GetValue('PPF_CHEURETRAV');
               HorsTravail := TobMvtDIF.Detail[i].GetValue('PPF_CHEURENONTRAV');
               TypeAlim := TobMvtDIF.Detail[i].GetValue('PPF_ALIMENT');
               RecupereLesFormationsDIF(Salarie,TypeFor,Rubrique,TypeAlim,TobInsc.GetValue('PFO_DATEPAIE'),Travail = 'X',HorsTravail = 'X');
          end;
     end;
     TobMvtDIF.free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 26/05/2008 / PT28
Modifié le ... :   /  /
Description .. : Supprime une absence pour le DIF
Mots clefs ... :
*****************************************************************}
procedure SupprimerDIF(TobInsc : TOB; InitDatePaie : TDateTime; Supp : Boolean; var SuppAbsence,SuppRubrique : Boolean);
var Q : TQuery;
    TypeFor : String;
    TobMvtDIF : Tob;
    i : Integer;
    Travail,HorsTravail : String;
    Rubrique,TypeConge,TypeAlim,Predef : String;
    DateDebRub,DateFinRub : TDateTime;
begin
     SuppRubrique := True;
     SuppAbsence := True;
     //TypeFor := InitTypePlan;
     TypeFor := TobInsc.GetValue('PFO_TYPEPLANPREV');
     
     Q := OpenSQL('SELECT * FROM PARAMFORMABS WHERE ##PPF_PREDEFINI## PPF_TYPEPLANPREV="'+TypeFor+'"',True);
     TobMvtDIF := Tob.Create('Lesmvts',Nil,-1);
     TobMvtDIF.LoadDetailDB('Lesmvts','','',Q,False);
     Ferme(Q);
     
     For i := 0 TO TobMvtDIF.Detail.Count - 1 do
     begin
          Travail := TobMvtDIF.Detail[i].GetValue('PPF_CHEURETRAV');
          HorsTravail := TobMvtDIF.Detail[i].GetValue('PPF_CHEURENONTRAV');
          Predef := TobMvtDIF.Detail[i].GetValue('PPF_PREDEFINI');
          
          If Predef <> 'DOS' then
          begin
               If TobMvtDIF.FindFirst(['PPF_PREDEFINI','PPF_CHEURETRAV','PPF_CHEURENONTRAV','PPF_TYPEPLANPREV'],['DOS',Travail,HorsTravail,TypeFor],False) <> Nil then Continue;
               If Predef = 'CEG' then
               begin
                    If TobMvtDIF.FindFirst(['PPF_PREDEFINI','PPF_CHEURETRAV','PPF_CHEURENONTRAV','PPF_TYPEPLANPREV'],['STD',Travail,HorsTravail,TypeFor],False) <> Nil then Continue;
               end;
          end;
          If TobMvtDIF.Detail[i].GetValue('PPF_ALIMABS') = 'X' then
          begin
               TypeConge := TobMvtDIF.Detail[i].GetValue('PPF_TYPECONGE');
               If ExisteSQL('SELECT PCN_DATEDEBUTABS,PCN_DATEFINABS FROM ABSENCESALARIE WHERE '+
               'PCN_TYPECONGE="'+TypeConge+'" AND '+
               'PCN_SALARIE="'+TobInsc.GetValue('PFO_SALARIE')+'" AND '+
               '(PCN_CODETAPE="P" OR PCN_CODETAPE="S") AND '+
               'PCN_DATEDEBUTABS="'+UsDateTime(TobInsc.GetValue('PFO_DATEDEBUT'))+'" AND '+
               'PCN_DATEFINABS="'+UsDateTime(TobInsc.GetValue('PFO_DATEFIN'))+'"') then
               begin
                    SuppAbsence := True;
                    PGIBox('L''absence pour le salarié '+RechDom('PGSALARIE',TobInsc.GetValue('PFO_SALARIE'),False)+
                    ' du '+DateToStr(TobInsc.GetValue('PFO_DATEDEBUT'))+' au '+DateToStr(TobInsc.GetValue('PFO_DATEFIN'))+
                    '#13#10 a déjà été intégrée dans le bulletin. Aucune modification ne sera effectuée','Génération absence');
               end
               else PGDIFSuppAbs(TobInsc.GetValue('PFO_SALARIE'),TypeConge,TobInsc.GetValue('PFO_DATEDEBUT'),TobInsc.GetValue('PFO_DATEFIN'));
          end;
          If TobMvtDIF.Detail[i].GetValue('PPF_ALIMRUB') = 'X' then
          begin
               Rubrique := TobMvtDIF.Detail[i].GetValue('PPF_RUBRIQUE');
               TypeAlim := TobMvtDIF.Detail[i].GetValue('PPF_ALIMENT');
               DateDebRub := DebutDeMois(InitDatePaie);
               DateFinRub := FinDeMois(InitDatePaie);
               If ExisteSQL('SELECT PPU_DATEDEBUT,PPU_DATEFIN FROM PAIEENCOURS WHERE '+
               'PPU_DATEDEBUT>="'+UsDateTime(dateDebRub)+'" AND '+
               'PPU_DATEFIN<="'+UsDateTime(dateFinRub)+'" AND '+
               'PPU_SALARIE="'+TobInsc.GetValue('PFO_SALARIE')+'"') then
               begin
                    SuppRubrique := False;
                    If InitDatePaie <> TobInsc.GetValue('PFO_DATEPAIE') then
                    begin
                         PGIBox('Le bulletin pour le salarié '+RechDom('PGSALARIE',TobInsc.GetValue('PFO_SALARIE'),False)+
                         ' du '+DateToStr(dateDebRub)+' au '+DateToStr(dateFinRub)+
                         '#13#10 a déjà été créé. La date de validité ne peut être modifiée et aucune modification ne sera effectuée pour les rubriques','Génération absence');
                         TobInsc.PutValue('PFO_DATEPAIE',InitDatePaie);
                    end
                    else
                    begin
                         PGIBox('Le bulletin pour le salarié '+RechDom('PGSALARIE',TobInsc.GetValue('PFO_SALARIE'),False)+
                         ' du '+DateToStr(dateDebRub)+' au '+DateToStr(dateFinRub)+
                         '#13#10 a déjà été créé. Aucune modification ne sera effectuée pour les rubriques','Génération absence');
                    end;
               end
               else
               begin
                    PGDIFSuppRub(TobInsc.GetValue('PFO_SALARIE'),Rubrique,TypeAlim,InitDatePaie);
                    If Supp then RecupereLesFormationsDIF(TobInsc.GetValue('PFO_SALARIE'),TypeFor,Rubrique,TypeAlim,InitDatePaie,Travail = 'X',HorsTravail = 'X',TobInsc.GetValue('PFO_CODESTAGE'),TobInsc.GetValue('PFO_ORDRE'));
               end;
          end;
     end;
     TobMvtDIF.free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 04/06/2008 / PT31
Modifié le ... :   /  /
Description .. : Renvoie le niveau de confidentialité de l'utilisateur
Mots clefs ... :
*****************************************************************}
Function RendNivAccesConfidentiel : String;
var NivAcces : Integer;
    Q        : TQuery;
Begin
	NivAcces := 0;
	Q := OpenSQL('SELECT UG_CONFIDENTIEL,UG_NIVEAUACCES FROM USERGRP WHERE UG_GROUPE="'+V_PGI.FGroupe+'"', True);
	If Not Q.EOF Then
	Begin
	   If Q.FindField('UG_CONFIDENTIEL').AsString = '1' Then
	      NivAcces := Q.FindField('UG_NIVEAUACCES').AsInteger;
	End;
	Ferme(Q);
	Result := IntToStr(NivAcces);
End;

{$ENDIF} //IFNDEF PLUGIN

Procedure MultiDosFormation;
Var
{$IFDEF PLUGIN}
 ModePCL,BaseCommune : Boolean;
{$ENDIF}
    Q                   : TQuery ;
begin
	PGBundleInscFormation := False;
	PGDroitMultiForm      := False;
	PGBundleHierarchie    := False;
	PGBundleCatalogue     := False; //PT23

	// Bundle Inscriptions aux formations
	Q := OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE="FORMATIONS"', True); //PT23
	If Not Q.EOF Then
	Begin
		PGBundleInscFormation := True;
        {$IFNDEF PLUGIN}VH_PAIE.{$ENDIF}BasesPartage[BUNDLE_FORMATION] := Q.FindField('DS_NOMBASE').AsString; //PT25
	End;
	Ferme(Q);
	
	//PT23
	// Bundle Catalogue de stage
	Q := OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE="STAGE"', True);
	If Not Q.EOF Then
	Begin
		PGBundleCatalogue := True;
		{$IFNDEF PLUGIN}VH_PAIE.{$ENDIF}BasesPartage[BUNDLE_CATALOGUE] := Q.FindField('DS_NOMBASE').AsString; //PT25
	End;
	Ferme(Q);
	
	// Bundle Hierarchie
	Q := OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE="SERVICES"', True);
	If Not Q.EOF Then
	Begin
		PGBundleHierarchie := True;
		{$IFNDEF PLUGIN}VH_PAIE.{$ENDIF}BasesPartage[BUNDLE_HIERARCHIE] := Q.FindField('DS_NOMBASE').AsString; //PT25
	End;
	Ferme(Q);
	
    If JaiLeDroitTag(200171) then PGDroitMultiForm := True;

  	//PT17 - Début
  	{$IFDEF PLUGIN}
  	{ Cas particulier du portail en PCL. On se connecte uniquement sur la base partagée.
      Dans ce cas, on récupère le dossier du salarié connecté.
      Si aucun salarié rattaché, le dossier reste le même. }
    ModePCL := False;
    Q :=OpenSQL('SELECT DISTINCT(SO_MODEFONC) FROM SOCIETE', True);
    If Not Q.EOF Then ModePCL := Q.FindField('SO_MODEFONC').AsString = 'DB0';
    Ferme(Q);

    If ModePCL Then GrpDossiers := True; //PT25

    BaseCommune := False;
    Q := OpenSQL('SELECT DISTINCT(DS_NOMBASE) FROM DESHARE', True);
    If Not Q.EOF Then BaseCommune := Q.FindField('DS_NOMBASE').AsString = 'DB'+V_PGI.NoDossier;
    Ferme(Q);

	If (PGBundleInscFormation) And (ModePCL) And (BaseCommune) And (V_PGI.UserSalarie <> '') Then
   	Begin
        V_PGI.NoDossier := GetNoDossierSalarie;
	End
  	{$ENDIF}
  	//PT17 - Fin
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 10/05/2007 / PT6
Modifié le ... :   /  /    
Description .. : Recherche la population correspondant aux données
Suite ........ : présentes sans spécification d'un salarié particulier.
Suite ........ : La TOB Tinfos doit contenir les valeurs nécessaires
Suite ........ : à la recherche (provenant d'un écran ou autre).
Mots clefs ... : POPULATION
*****************************************************************}
Function RecherchePopulation (TInfos : TOB) : String;
Var
  sql1,sql2,sql3, population, codepop, populsalarie, sti, numident,nomchamp, nodossier, pred: String;
  Q1, QQ,  Q3, Q0 : TQuery;
  tob_Identsal, Tids , tob_valident , tob_paieparim , TD: TOB;
  i,  nbident: integer;
  trouve : boolean;
  Tob_TempValIdent,Tob_TempIdentSal : TOB;
  Saisie, Valeurs, ValeurTemp : String;

  {$IFDEF PLUGIN}
  // Redéfinition de la fonction ici pour éviter d'inclure des tas de fichiers dans les vignettes
  Function GetPredefiniPopulation(TypePop : String) : String;
  var
    sql1 : String;
    TobPop : Tob;
  begin
  //Optimisation de la recherche du prédéfini -> 1 seul requete
  sql1 := 'select PPC_POPULATION, PPO_CODEPOP, PPO_PREDEFINI,PPO_NODOSSIER from ORDREPOPULATION, CODEPOPULATION where ##PPC_PREDEFINI## AND ##PPO_PREDEFINI## AND PPC_CODEPOP = PPO_CODEPOP'+  //PT13
  ' and PPO_TYPEPOP LIKE "%'+ typepop +'%"';
  TobPop := Tob.Create('recherche du prédéfini',nil,-1);
  TobPop.LoadDetailFromSQL(sql1);
  result := 'DOS';
  if not Assigned(TobPop.FindFirst(['PPO_PREDEFINI','PPO_NODOSSIER'],['DOS',V_PGI.NoDossier],False)) then
  begin
    result := 'STD';
    if not Assigned(TobPop.FindFirst(['PPO_PREDEFINI'],['STD'],False)) then
    begin
      result := 'CEG';
      if not Assigned(TobPop.FindFirst(['PPO_PREDEFINI'],['CEG'],False)) then
      begin
        result := '';
      end;
    end;
  end;
  FreeAndNil(TobPop);
  end;
  {$ENDIF}
Begin
  Trouve := False;
  populsalarie := '';

  // Recherche des champs de PAIEPARIM utilisables comme critères de population
  Q0 := Opensql('select PAI_IDENT,PAI_PREFIX,PAI_SUFFIX from PAIEPARIM where ##PAI_PREDEFINI## AND PAI_UTILISABLEPOP <> ""', true); //PT13
  Tob_paieparim := Tob.Create('Critères population', nil, -1);
  Tob_paieparim.LoadDetailDB('Critères population', '', '', Q0, False);
  ferme(Q0);

  // Lecture des codes critères concernés par le type de population
  {$IFNDEF PLUGIN}
  nodossier := PgrendNodossier();
  {$ELSE}
  nodossier := V_PGI.NoDossier;
  {$ENDIF}

  // Détermination du prédéfini (niveau de population : CEG, STD ou DOS)
  pred := GetPredefiniPopulation(TYP_POPUL_FORM_PREV);

  // Constitution de la requête listant les codes population
  if pred = 'DOS' then
    sql1 := 'select PPO_CODEPOP from CODEPOPULATION where ##PPO_PREDEFINI## AND PPO_PREDEFINI= "DOS" and PPO_NODOSSIER = "'+nodossier+'" and'+ //PT13
            ' PPO_TYPEPOP LIKE "%'+ TYP_POPUL_FORM_PREV +'%"'
  else
    sql1 := 'select PPO_CODEPOP from CODEPOPULATION where ##PPO_PREDEFINI## AND PPO_PREDEFINI= "'+pred+'" and PPO_TYPEPOP LIKE "%'+ TYP_POPUL_FORM_PREV +'%"' ; //PT13

  QQ:= Opensql(sql1, true);
  While Not QQ.Eof Do
  begin
    codepop := QQ.findfield('PPO_CODEPOP').asstring;

    // Recherche des uplets correspondant aux critères possibles pour le code population
    sql2 := 'SELECT PPO_NBIDENT, PPO_IDENT1, PPO_IDENT2, PPO_IDENT3, PPO_IDENT4,PPO_LIBELLE,PPO_TYPEPOP FROM CODEPOPULATION' +
            ' Where ##PPO_PREDEFINI## AND PPO_CODEPOP = "'+codepop+'" ';  //PT13

    Q1 := OPENSQL(sql2, True);
    if not Q1.Eof then
    begin
      // Création d'une TOB contenant le nom de chaque champ de la table avec son numéro d'index
      nbident := Q1.findfield('PPO_NBIDENT').asinteger;
      if assigned(Tob_Identsal) then FreeAndNil(Tob_identsal);
      Tob_Identsal := Tob.Create('Nom identifiants salarie', nil, -1);
      for i := 1 to nbident do
      begin
        Sti := IntToStr(i);
        numident := Q1.findfield('PPO_IDENT' + sti).asstring;
        TD := Tob_paieparim.findfirst(['PAI_IDENT'], [numident], true);
        nomchamp := 'PFI' + '_' + TD.getvalue('PAI_SUFFIX');  // forçage du préfixe à PFI
        Tids := Tob.Create('Identifiant', Tob_Identsal, -1);
        Tids.AddChampSup('NOMCHAMP', False);
        Tids.PutValue('NOMCHAMP', Nomchamp);
        Tids.AddChampSup('ORDREIDENT',False);
        Tids.PutValue('ORDREIDENT',sti);
      end;
      
      // Lecture des populations pour ce code population
      sql3 := 'SELECT PPC_POPULATION, PPC_LIBELLE,PPC_VALIDENT1,PPC_VALIDENT2, PPC_VALIDENT3, PPC_VALIDENT4'+
              ' FROM ORDREPOPULATION WHERE ##PPC_PREDEFINI## PPC_CODEPOP = "' + codepop + '" ';
      Q3 := OPENSQL(sql3, True);
      if assigned(Tob_valident) then FreeAndNil(Tob_valident);
      Tob_ValIdent := Tob.Create('Valeur identifiants', nil, -1);
      Tob_ValIdent.LoadDetailDB('Valeur identifiants', '', '', Q3, False);
      ferme(Q3);

      Trouve := False;

      // Recherche des critères et valeurs qui correspondent avec les éléments saisis
      // Parcours de chaque uplet de valeurs
      Tob_TempValIdent := Tob_ValIdent.FindFirst([''],[''],False);
      While Tob_TempValIdent <> nil do
      begin
        // Récupération de la population en cours
        population := Tob_TempValIdent.GetString('PPC_POPULATION');

        Trouve := True;

        // Parcours de chaque champ à contrôler
        Tob_TempIdentSal := Tob_IdentSal.FindFirst([''],[''],False);
        While Tob_TempIdentSal <> Nil Do
        Begin
          // Récupération du champ de saisie
          If (TInfos.GetNumChamp(Tob_TempIdentSal.GetValue('NOMCHAMP')) > -1) Then
          Begin
            // Le champ est trouvé dans le formulaire. On vérifie maintenant qu'une valeur a été saisie
            Saisie := TInfos.GetValue(Tob_TempIdentSal.GetValue('NOMCHAMP'));
            If (Saisie <> '') Then
            Begin
              Valeurs := Tob_TempValIdent.GetValue('PPC_VALIDENT'+Tob_TempIdentSal.GetValue(('ORDREIDENT')));
              // On recherche si la valeur existe dans la liste des valeurs possibles de la population
              ValeurTemp := ReadTokenSt(Valeurs);
              While (ValeurTemp <> '') Do
              Begin
                If (Saisie = ValeurTemp) Then Break;
                ValeurTemp := ReadTokenSt(Valeurs);
              End;

              If (ValeurTemp = '') Then
              Begin
                // La valeur saisie ne fait pas partie de la population en cours. On passe à la suivante
                Trouve := False; Break;
              End;
            End
            Else
            Begin
              // Le champ n'a pas encore été saisi. On ne peut donc pas vérifier l'appartenance à une population
              Trouve := False; Break;
            End;
          End
          Else
          Begin
            // Le champ de saisie n'est pas trouvé, ce n'est pas la peine d'aller plus loin car
            // la population ne sera jamais vérifiée
            Trouve := False; Break;
          End;

          // Passage au champ suivant
          Tob_TempIdentSal := Tob_IdentSal.FindNext([''],[''],False);
        End;

        // Si en sortant de la boucle de parcours de chaque champ on a Trouve=True, c'est bon
        If Trouve Then Break;

        // Passage à la liste de valeurs suivante pour le même code population
        Tob_TempValIdent := Tob_ValIdent.FindNext([''],[''],False);
      end;
    end;
    ferme(Q1);

    If Trouve Then
    Begin
      // On a trouvé une population qui satisfait la saisie.
      populsalarie := population;
      break;
    End;

    // Passage au code population suivant
    QQ.next;
  end;
  
  ferme(QQ);
  if assigned(Tob_valident) then FreeAndNil(Tob_valident);
  if assigned(Tob_Identsal) then FreeAndNil(Tob_identsal);
  if assigned(Tob_Paieparim) then FreeAndNil(Tob_Paieparim);
  
  Result:= populsalarie;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jérôme LEFEVRE
Créé le ...... : 02/06/2003
Modifié le ... :   /  /
Description .. : Fonction de maj en temps réel des différents coûts pour
Suite ........ : suivi de l'investissement.
Suite ........ : Type MAJ :
Suite ........ : - "TOUS" = Tous les coûts
Suite ........ : - "PEDAGOGIQUE" = Coût pédagogique (animateur +
Suite ........ : session)
Suite ........ : - "SALAIRE" = Coût salaire
Suite ........ : - "FRAIS" = Frais des salariés
Mots clefs ... :
*****************************************************************}

procedure CalcCtInvestSession(TypeMaj, LeStage, LeMillesime: string; LaSession: Integer);
var
  Q: TQuery;
  TInvest: Tob;
  i: Integer;
  Maj: Boolean;
  CoutSalaire, Frais,  CoutPedag, CoutTotal, TauxP, TauxS, TauxF: Double;
begin
  Q := OpenSQL('SELECT * FROM INVESTSESSION WHERE PIS_CODESTAGE="' + LeStage + '" AND ' +
    'PIS_ORDRE=' + IntToStr(LaSession) + ' AND PIS_MILLESIME="' + LeMillesime + '"', True);
  TInvest := Tob.Create('INVESTSESSION', nil, -1);
  TInvest.LoadDetailDB('INVESTSESSION', '', '', Q, False);
  Ferme(Q);
  for i := 0 to TInvest.Detail.Count - 1 do
  begin
    Maj := False;
    if TypeMaj = 'TOUS' then Maj := True;
    if TypeMAJ = 'PEDAGOGIQUE' then
    begin
      if TInvest.Detail[i].GetValue('PIS_COUTPEDAG') = 'X' then Maj := True;
    end;
    if TypeMAJ = 'SALAIRE' then
    begin
      if TInvest.Detail[i].GetValue('PIS_COUTSALAIRE') = 'X' then Maj := True;
    end;
    if TypeMAJ = 'FRAIS' then
    begin
      if TInvest.Detail[i].GetValue('PIS_FRAISFORM') = 'X' then Maj := True;
    end;
    if Maj = True then
    begin
      CoutSalaire := 0;
      Frais := 0;
      CoutPedag := 0;
      if (TInvest.Detail[i].GetValue('PIS_COUTSALAIRE') = 'X') or (TInvest.Detail[i].GetValue('PIS_FRAISFORM') = 'X') then
      begin
        Q := OpenSQL('SELECT SUM(PFO_COUTREELSAL) SALAIRE, SUM(PFO_FRAISREEL) FRAIS FROM FORMATIONS WHERE PFO_CODESTAGE="' + LeStage + '" AND ' +
          'PFO_ORDRE=' + IntToStr(LaSession) + ' AND PFO_MILLESIME="' + LeMillesime + '"', True); //DB2
        if not Q.Eof then
        begin
          if TInvest.Detail[i].GetValue('PIS_COUTSALAIRE') = 'X' then CoutSalaire := Q.FindField('SALAIRE').AsFloat;
          if TInvest.Detail[i].GetValue('PIS_FRAISFORM') = 'X' then Frais := Q.FindField('FRAIS').AsFloat;
        end;
        Ferme(Q);
      end;
      if TInvest.Detail[i].GetValue('PIS_COUTPEDAG') = 'X' then
      begin
        Q := OpenSQL('SELECT PSS_COUTPEDAG FROM SESSIONSTAGE WHERE PSS_CODESTAGE="' + LeStage + '" AND ' +
          'PSS_ORDRE=' + IntToStr(LaSession) + ' AND PSS_MILLESIME="' + LeMillesime + '"', True); //DB2
        if not Q.Eof then CoutPedag := Q.FindField('PSS_COUTPEDAG').AsFloat;
        Ferme(Q);
        Q := OpenSQL('SELECT SUM (PAN_SALAIREANIM) SALAIREANIM FROM SESSIONANIMAT WHERE PAN_CODESTAGE="' + LeStage + '" AND ' +
          'PAN_ORDRE=' + IntToStr(LaSession) + ' AND PAN_MILLESIME="' + LeMillesime + '"', True); //DB2
        if not Q.Eof then CoutPedag := CoutPedag + Q.FindField('SALAIREANIM').AsFloat;
        Ferme(Q);
        //DEBUT PT1
        Q := OpenSQL('SELECT SUM (PFS_MONTANT) FRAISANIM FROM FRAISSALFORM LEFT JOIN SESSIONANIMAT ON PFS_SALARIE=PAN_SALARIE AND PFS_CODESTAGE=PAN_CODESTAGE AND PAN_ORDRE=PFS_ORDRE ' +
          'WHERE PAN_CODESTAGE="' + LeStage + '" AND PAN_FRASPEDAG="X" AND ' +
          'PAN_ORDRE=' + IntToStr(LaSession) + ' AND PAN_MILLESIME="' + LeMillesime + '"', True); //DB2
        if not Q.Eof then CoutPedag := CoutPedag + Q.FindField('FRAISANIM').AsFloat;
        Ferme(Q);
        //FIN PT1
      end;
      TauxP := TInvest.Detail[i].GetValue('PIS_TAUXPEDAG');
      TauxS := TInvest.Detail[i].GetValue('PIS_TAUXSALAIRE');
      TauxF := TInvest.Detail[i].GetValue('PIS_TAUXSALFRAIS');
      CoutSalaire := CoutSalaire * TauxS;
      CoutPedag := CoutPedag * TauxP;
      Frais := Frais * TauxF;
      CoutTotal := CoutSalaire + Frais + CoutPedag;
      TInvest.Detail[i].PutValue('PIS_MONTANT', CoutTotal);
      TInvest.Detail[i].UpdateDB(False);
    end;
  end;
  TInvest.Free;
end;


Function RendMillesimeRealise(var DD,DF : TDateTime) : String;
var Q : TQuery;
begin
     Result := '0000';
     DF := IDate1900;
     DD := IDate1900;
     Q := OpenSQL('SELECT PFE_MILLESIME,PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_ACTIF="X" AND PFE_CLOTURE="-" ORDER BY PFE_MILLESIME DESC',True);
     If Not Q.Eof then
     begin
          DD := Q.FindField('PFE_DATEDEBUT').AsDateTime;
          DF := Q.FindField('PFE_DATEFIN').AsDateTime;
          Result := Q.FindField('PFE_MILLESIME').AsString;
     end;
     Ferme(Q);
end;

Function RendMillesimePrevisionnel : String;
var Q : TQuery;
begin
     Result := '0000';
     Q := OpenSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE PFE_OUVREPREV="X" AND PFE_CLOTUREPREV="-" ORDER BY PFE_MILLESIME DESC',True);
     If Not Q.Eof then Result := Q.FindField('PFE_MILLESIME').AsString;
     Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jérôme LEFEVRE
Créé le ...... : 30/01/2003
Modifié le ... : 13/04/2007 (FL / PT3)
Modifié le ... : 17/04/2007 (FL / PT4)
Description .. : Fonction qui renvoie le taux horaire du salarié selon le mode
Suite ........ : de calcul dans paramsoc et selon les méthodes de calcul prévisionnel ou réel.
Suite ........ : Prend en compte également l'historique du T.H. dans le cas où la formation aurait
Suite ........ : été saisie à postériori.
Mots clefs ... : TAUX HORAIRE;PREVISIONNEL;REEL;
*****************************************************************}

function ForTauxHoraireReel(LeSalarie: string; DebutPaie: TDateTime = 0; FinPaie: TDateTime = 0; MethodeCalc: string = ''; TypeCalc:Boolean=ValReel; DateFinFormation:TDateTime=0; LeDossier: string = ''): Double; //PT3 //PT4
var
  Q: TQuery;
  Brut, NbHeures, TauxHoraire: Double;
  Methode : string;
  DateApplication : TDateTime;
  DernierHisto : Boolean;
begin
  TauxHoraire := 0;
  DateApplication := iDate1900;
  DernierHisto := False;
  //PT3 - Début
  If MethodeCalc = '' Then
  Begin
     If (TypeCalc = ValReel) Then
        {$IFNDEF PLUGIN}
          Methode := VH_Paie.PGForMethodeCalc
        {$ELSE}
          Methode := GetParamSoc('SO_PGFORMETHODECALC')
        {$ENDIF}
     Else
        {$IFNDEF PLUGIN}
          Methode := VH_Paie.PGForMethodeCalcPrev;
        {$ELSE}
          Methode := GetParamSoc('SO_PGFORMETHODECALCPREV');
        {$ENDIF}
  End
  //PT3 - Fin
  else Methode := MethodeCalc;

  if Methode = 'THS' then
  begin
    //PT4 - Début
    If (DateFinFormation <> 0) Then
    Begin
      DernierHisto := False;

      { Recherche d'un taux horaire historisé actif à la date de fin de formation, donc ayant une
        date d'application antérieure à la date de fin de formation }
      If LeDossier <> '' then
      begin
        Q := OpenSQL('SELECT PHS_DATEAPPLIC, PHS_TAUXHORAIRE FROM '+GetBase(LeDossier, 'HISTOSALARIE')+' WHERE PHS_SALARIE="' + LeSalarie + '" AND '  //PT17
                 + 'PHS_BTAUXHORAIRE="X" ORDER BY PHS_DATEAPPLIC ASC', True);
      end
      else Q := OpenSQL('SELECT PHS_DATEAPPLIC, PHS_TAUXHORAIRE FROM HISTOSALARIE WHERE PHS_SALARIE="' + LeSalarie + '" AND '
                 + 'PHS_BTAUXHORAIRE="X" ORDER BY PHS_DATEAPPLIC ASC', True);
      If Not Q.Eof Then
      Begin
        Q.First;
        Repeat
        Begin
          DateApplication := Q.FindField('PHS_DATEAPPLIC').AsDateTime;

          If (DateApplication < DateFinFormation) Then
            TauxHoraire := Q.FindField('PHS_TAUXHORAIRE').AsFloat
          Else
            { Dans le cas où le taux horaire n'était pas encore renseigné à une date antérieure, il apparaît à 0 en table.
             Dès lors, on récupère la valeur suivante la plus proche }
            If (TauxHoraire = 0) Then
              TauxHoraire := Q.FindField('PHS_TAUXHORAIRE').AsFloat;

          Q.Next;

          If (Q.Eof) Then DernierHisto := True;
        End
        Until ((DateApplication >= DateFinFormation) Or (Q.Eof));
      End;
      Ferme(Q);
    End;
    If ((TauxHoraire = 0) Or ((DateApplication < DateFinFormation) And (DernierHisto = True))) Then
    //PT4 - Fin
    Begin
      { 2 possibilités :
        - Pas de taux historisé : on prend l'actuel
        - On a trouvé un taux historisé mais il est antérieur à la formation et c'est le dernier de la liste : on prend l'actuel }
      If LeDossier <> '' then Q := OpenSQL('SELECT PSA_TAUXHORAIRE FROM '+GetBase(LeDossier, 'SALARIES')+' WHERE PSA_SALARIE="' + LeSalarie + '"', True)  //PT17
      else Q := OpenSQL('SELECT PSA_TAUXHORAIRE FROM SALARIES WHERE PSA_SALARIE="' + LeSalarie + '"', True);
      if not Q.Eof then TauxHoraire := Q.FindField('PSA_TAUXHORAIRE').AsFloat;
      Ferme(Q);
    End;
  end;

  if Methode = 'DPE' then
  begin
   If LeDossier <> '' then
   begin
    if (DebutPaie > IDate1900) and (FinPaie > IDate1900) then
      Q := OpenSQL('SELECT PPU_CBRUT,PPU_CHEURESTRAV FROM '+GetBase(LeDossier, 'PAIEENCOURS')+' WHERE PPU_SALARIE="' + LeSalarie + '" AND PPU_DATEDEBUT>="' + UsDateTime(DebutPaie) + '" AND PPU_DATEFIN<="' + UsDateTime(FinPaie) + '"', True)  //PT17
    else Q := OpenSQL('SELECT PPU_CBRUT,PPU_CHEURESTRAV FROM '+GetBase(LeDossier, 'PAIEENCOURS')+' WHERE PPU_SALARIE="' + LeSalarie + '" ORDER BY PPU_DATEDEBUT DESC', True)
   end
   else
   begin
    if (DebutPaie > IDate1900) and (FinPaie > IDate1900) then
      Q := OpenSQL('SELECT PPU_CBRUT,PPU_CHEURESTRAV FROM PAIEENCOURS WHERE PPU_SALARIE="' + LeSalarie + '" AND PPU_DATEDEBUT>="' + UsDateTime(DebutPaie) + '" AND PPU_DATEFIN<="' + UsDateTime(FinPaie) + '"', True)
    else Q := OpenSQL('SELECT PPU_CBRUT,PPU_CHEURESTRAV FROM PAIEENCOURS WHERE PPU_SALARIE="' + LeSalarie + '" ORDER BY PPU_DATEDEBUT DESC', True)
   end;
    if not Q.Eof then
    begin
      Q.First;
      Brut := Q.FindField('PPU_CBRUT').AsFloat;
      NbHeures := Q.FindField('PPU_CHEURESTRAV').AsFloat;
      if NbHeures > 0 then TauxHoraire := Brut / NbHeures;
    end;
    Ferme(Q);
  end;
  Result := arrondi(TauxHoraire, 2);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jérôme LEFEVRE
Créé le ...... : 30/01/2003
Modifié le ... :
Description .. : Fonction qui renvoie le taux horaire catégoriel qui se trouve
Suite ........ : dans la table SALAIREFORM
Mots clefs ... :
*****************************************************************}

function ForTauxHoraireCategoriel(LeLibEmploi, LeMillesime: string): Double;
var
  Q: TQuery;
  TauxHoraire: Double;
begin
  TauxHoraire := 0;
  Q := OpenSQL('SELECT PSF_MONTANT FROM SALAIREFORM WHERE ' +
    'PSF_LIBEMPLOIFOR="' + LeLibEmploi + '" AND PSF_MILLESIME="' + LeMillesime + '"', True);
  if not Q.eof then TauxHoraire := Q.FindField('PSF_MONTANT').AsFloat;
  Ferme(Q);
  //       TauxCharge :=
  Result := Arrondi(TauxHoraire, 2);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 08/11/2007
Modifié le ... :   /  /    
Description .. : Mise à jour des coûts des inscrits suite à une
Suite ........ : nouvelle inscription au prévisionnel
Mots clefs ... : 
*****************************************************************}
Procedure MajCoutsInscPrev (Stage, Millesime : String);
var Q : TQuery;
    i : Integer;
    CoutFonc,CoutUnit,CoutAnim,NbMax,NbDivD,NbDivE : Double;
    NbInsc,NbCout,CoutFoncSta,CoutAnimSta : Double;
    NbInscStage : Integer;
    TobInsc : TOB;
begin
     CoutFonc   := 0;
     CoutUnit   := 0;
     CoutAnim   := 0;
     NbMax      := 0;


     // Récupération des coûts théoriques de la formation
     Q := OpenSQL('SELECT PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR,PST_NBSTAMAX,PST_DUREESTAGE,PST_JOURSTAGE FROM STAGE ' +
     'WHERE PST_CODESTAGE="' + Stage + '" AND PST_MILLESIME="' + Millesime + '"', True);
     If Not Q.eof then
     begin
          CoutFonc  := Q.FindField('PST_COUTBUDGETE').AsFloat;
          CoutUnit  := Q.FindField('PST_COUTUNITAIRE').AsFloat;
          CoutAnim  := Q.FindField('PST_COUTSALAIR').AsFloat;
          NbMax     := Q.FindField('PST_NBSTAMAX').AsInteger;
     end;
     Ferme(Q);

     // Récupération du nombre total d'inscrits
     Q := OpenSQL('SELECT SUM(PFI_NBINSC) NBINSC FROM INSCFORMATION ' +
                          'WHERE PFI_CODESTAGE="' + Stage + '" AND PFI_MILLESIME="' + Millesime + '" ' +
                          'AND (PFI_ETATINSCFOR="ATT" OR PFI_ETATINSCFOR="VAL")', True);
     if not Q.Eof then NbInscStage := Q.FindField('NBINSC').AsInteger
     else NbInscStage := 1;
     Ferme(Q);

     TobInsc := TOB.Create('LesInscrits', Nil, -1);
     TobInsc.LoadDetailDBFromSQL('INSCFORMATION', 'SELECT * FROM INSCFORMATION WHERE PFI_CODESTAGE="'+Stage+'" AND PFI_MILLESIME="'+Millesime+'"');

     for i:=0 to TobInsc.Detail.Count-1 do
     begin
          NbInsc := TobInsc.Detail[i].GetValue('PFI_NBINSC');
          if NbMax > 0 then
          begin
               NbDivD := NbInscStage / NbMax;
               NbDivE := Arrondi(NbDivD, 0);
               if ((NbDivE - NbDivD) < 0) then NbCout := NbDivE + 1
               else NbCout := NbDivE;
          end
          else NbCout := 1;
          CoutFoncSta := (NbCout * CoutFonc) / NbInscStage;
          CoutAnimSta := (NbCout * CoutAnim) / NbInscStage;
          If (TobInsc.Detail[i].GetValue('PFI_ETATINSCFOR')='ATT') or (TobInsc.Detail[i].GetValue('PFI_ETATINSCFOR')='VAL') then TobInsc.Detail[i].PutValue('PFI_AUTRECOUT',NbInsc * (CoutFoncSta + CoutAnimSta + Coutunit))
          else TobInsc.Detail[i].PutValue('PFI_AUTRECOUT',0);
     end;
     TobInsc.InsertOrUpdateDB;
     FreeAndNil(TobInsc);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 08/11/2007
Modifié le ... :   /  /
Description .. : Création d'un stage au prévisionnel
Mots clefs ... :
*****************************************************************}
procedure CreerStageAuPrevisionnel (CodeStage, Millesime: String);
var
  Q, QAnim: TQuery;
  TobStage, T, TobInvest, TI: Tob;
  NbHeure: Integer;
  SalaireAnim, NbAnim: Double;
begin
    if not ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_CODESTAGE="' + CodeStage + '" AND PST_MILLESIME="' + Millesime + '"') then
    begin
        QAnim := OpenSQl('SELECT PFE_SALAIREANIM FROM EXERFORMATION WHERE PFE_MILLESIME="' + Millesime + '"', True);
        salaireanim := 0;
        if not QAnim.eof then SalaireAnim := QAnim.FindField('PFE_SALAIREANIM').AsFloat;
        Ferme(QAnim);

        Q := OpenSQL('SELECT * FROM STAGE WHERE PST_CODESTAGE="' + CodeStage + '" AND PST_MILLESIME="0000"', True);
        TobStage := Tob.Create('STAGE', nil, -1);
        TobStage.loadDetailDB('STAGE', '', '', Q, False);
        Ferme(Q);

        T := TobStage.FindFirst(['PST_CODESTAGE'], [CodeStage], False);
        if T <> nil then
        begin
            NbHeure := T.GetValue('PST_DUREESTAGE');
            NbAnim := T.GetValue('PST_NBANIM');
            T.ChangeParent(TobStage, -1);
            T.PutValue('PST_MILLESIME', Millesime);
            T.PutValue('PST_COUTSALAIR', SalaireAnim * NbHeure * NbAnim);
            T.SetModifieField('PST_BLOCNOTE', True); // Pour forcer la récupération du mémo
            T.InsertOrUpdateDB;
        end;
        TobStage.Free;

        // Duplication investissement pour la formation
        Q := OpenSQL('SELECT * FROM INVESTSESSION ' +
                     'WHERE PIS_CODESTAGE="' + CodeStage + '" AND PIS_ORDRE=-1 AND PIS_MILLESIME="0000"', True); 
        TobInvest := Tob.Create('INVESTSESSION', nil, -1);
        TobInvest.LoadDetailDB('INVESTSESSION', '', '', Q, False);
        Ferme(Q);

        TI := TobInvest.FindFirst(['PIS_CODESTAGE'], [CodeStage], False);
        while TI <> nil do
        begin
            TI.ChangeParent(TobInvest, -1);
            TI.PutValue('PIS_MILLESIME', Millesime);
            TI.InsertOrUpdateDB(False);
            TI := TobInvest.FindNext(['PIS_CODESTAGE'], [CodeStage], False);
        end;
        TobInvest.Free;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 15/01/2008  / PT11
Modifié le ... :   /  /
Description .. : Calcul le coût pédagogique d'un salarié lors de l'inscription
Mots clefs ... :
*****************************************************************}
function CalCoutPedagogique (Stage, Millesime, Etablissement, Rang: String; NbInsc : Integer): Double;
var
  NbAnim,NbInscStage, NbMax: Integer;
  NbHeure ,SalaireAnim,CoutFonc, CoutUnit, CoutAnim, NbDivD, NbDivE, NbCout: Double;
  Q,QAnim : TQuery;
begin
  // Récupération des coûts globaux pour le stage millésimé
  Q := OpenSQL('SELECT PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR,PST_NBSTAMAX FROM STAGE ' +
    'WHERE PST_CODESTAGE="' + Stage + '" AND PST_MILLESIME="' + Millesime + '"', True);
  if not Q.Eof then
  begin
    CoutFonc := Q.FindField('PST_COUTBUDGETE').AsFloat;
    CoutUnit := Q.FindField('PST_COUTUNITAIRE').AsFloat;
    CoutAnim := Q.FindField('PST_COUTSALAIR').AsFloat;
    NbMax := Q.FindField('PST_NBSTAMAX').AsInteger;
    Ferme(Q);
  end
  else
  // Récupération des coûts globaux pour le stage au catalogue
  begin
    Ferme(Q);
    Q := OpenSQL('SELECT PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR,PST_NBSTAMAX,PST_DUREESTAGE,PST_NBANIM FROM STAGE ' +
      'WHERE PST_CODESTAGE="' + Stage + '" AND PST_MILLESIME="0000"', True);
    if not Q.Eof then
    begin
      CoutFonc := Q.FindField('PST_COUTBUDGETE').AsFloat;
      CoutUnit := Q.FindField('PST_COUTUNITAIRE').AsFloat;
      NbMax := Q.FindField('PST_NBSTAMAX').AsInteger;
      NbHeure := Q.FindField('PST_DUREESTAGE').AsFloat;
      NbAnim := Q.FindField('PST_NBANIM').AsInteger;
      Ferme(Q);
      QAnim := OpenSQl('SELECT PFE_SALAIREANIM FROM EXERFORMATION WHERE PFE_MILLESIME="' + Millesime + '"', True);
      salaireanim := 0;
      if not QAnim.eof then SalaireAnim := QAnim.FindField('PFE_SALAIREANIM').AsFloat;
      Ferme(QAnim);
      CoutAnim := SalaireAnim * NbHeure * NbAnim;
    end
    else
    begin
      CoutUnit := 0;
      CoutAnim := 0;
      Coutfonc := 0;
      NbMax := 0;
      Ferme(Q);
    end;
  end;

  // Récupération du nombre d'inscriptions à la formation
  Q := OpenSQL('SELECT SUM(PFI_NBINSC) NBINSC FROM INSCFORMATION ' +
    'WHERE PFI_CODESTAGE="' + Stage + '" AND PFI_MILLESIME="' + Millesime + '" ' +
    'AND (PFI_ETATINSCFOR="ATT" OR PFI_ETATINSCFOR="VAL") ' +
    'AND (PFI_ETABLISSEMENT<>"' + Etablissement + '" OR PFI_RANG<>' + Rang + ')', True);
  if not Q.Eof then NbInscStage := Q.FindField('NBINSC').AsInteger
  else NbInscStage := 0;
  Ferme(Q);

  NbInscStage := NbInscStage + NbInsc;
  if NbMax > 0 then
  begin
    NbDivD := NbInscStage / NbMax;
    NbDivE := arrondi(NbDivD, 0);
    if NbDivE - NbDivD < 0 then NbCout := NbDivE + 1
    else NbCout := NbDivE
  end
  else NbCout := 1;

  CoutFonc := (NbCout * CoutFonc) / NbInscStage;
  CoutAnim := (NbCout * CoutAnim) / NbInscStage;

  Result := (NbInsc) * (coutFonc + CoutAnim + Coutunit);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 15/02/2008 / PT12
Modifié le ... : 24/04/2008 / PT24
Description .. : Renvoie la partie Dossier du SqlWhere adaptée au choix Predefini / N°Dossier
Suite ........ : effectué par l'utilisateur.
Mots clefs ... :
*****************************************************************}
Function DossiersAInterroger (Predefini, Dossier, Prefixe : String; AvecPredefini : Boolean = False; AvecAnd : Boolean = False) : String;
Var DosCur : String;
    ForceSansPredef : Boolean;
Begin
	Result := '';

    If Prefixe = 'PSI' Then ForceSansPredef := True Else ForceSansPredef := False; //PT24

    If Dossier = '' Then
    Begin
    	If (AvecPredefini) And (Predefini <> '') Then
    		Result := Prefixe+'_PREDEFINI="'+Predefini+'"';

    	//PT19
		If {$IFNDEF PLUGIN} (V_PGI.ModePCL='1') {$ELSE} GrpDossiers {$ENDIF} And (Predefini <> 'STD') Then //PT25
		Begin
        	If (Predefini = '') And (Not ForceSansPredef) Then Result := '('+Prefixe+'_PREDEFINI="STD" OR ('+Prefixe+'_PREDEFINI="DOS" AND ' //PT22 //PT24
			Else If Result <> '' Then Result := Result + ' AND ';
			Result := Result + Prefixe + '_NODOSSIER IN (SELECT DOS_NODOSSIER FROM DOSSIER WHERE '+GererCritereGroupeConfTous+')';
            If (Predefini = '') And (Not ForceSansPredef) Then Result := Result + '))'; //PT22 //PT24
		End;
    End
    Else
    Begin
    	If (Predefini = 'STD') Then
        Begin
            If (AvecPredefini) Then Result := Prefixe+'_PREDEFINI="STD"';
        End
    	Else
    	Begin
        	If (Predefini = '') And (Not ForceSansPredef) Then Result := '('+Prefixe+'_PREDEFINI="STD" OR ('+Prefixe+'_PREDEFINI="DOS" AND ' //PT24
        	Else If (Predefini = 'DOS') And AvecPredefini Then Result := '('+Prefixe+'_PREDEFINI="'+Predefini+'" AND ';

        	// Prise en compte de plusieurs dossiers
        	If Pos(';', Dossier) = 0 Then
        		Result := Result + Prefixe+'_NODOSSIER="'+Dossier+'"'
        	Else
        	Begin
        		Result := Result + Prefixe+'_NODOSSIER IN (';
        		While Dossier <> '' Do
        		Begin
        			DosCur := ReadTokenSt(Dossier);
        			Result := Result + '"' + DosCur + '",';
        		End;
        		Result := Copy (Result,1,Length(Result)-1);
        		Result := Result + ')';
        	End;

        	If (Predefini = '') And (Not ForceSansPredef) Then Result := Result + '))' //PT24
            Else If (Predefini = 'DOS') And AvecPredefini Then Result := Result + ')';
        End;
    End;
    
    If (Result <> '') And AvecAnd Then Result := ' AND ' + Result;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 04/03/2008
Modifié le ... : 02/06/2008 / PT30
Description .. : Envoie un mail au destinataire (Responsable dans le cas
Suite ........ : d'une demande salarié, ou salarié dans le cas d'une validation
Suite ........ : par un responsable).
Suite ........ : La TobInfos contient les informations à insérer dans le mail
Suite ........ : provenant idéalement de la table PFI.
Mots clefs ... :
*****************************************************************}
procedure PrepareMailFormation (Salarie, Responsable, Action : String; TobInfos : TOB; var Titre : String; var Texte : HTStrings; HTML : Boolean = False);
var Q        : TQuery;
    AvantLe  : TDateTime;
    Acquis,Pris,Solde,SoldePrev,Demande : Double;
    Cumul    : String;
    Tablette : String;
    Chaine   : String;
    T,T2     : TOB;
    i        : Integer;
    Cursus   : Boolean;
    RC       : String;
    
    procedure AfficheRappelFormation (T : TOB);
    
    Begin
    	If HTML Then Texte.Add('<table border=0 cellspacing=1 cellpadding=1>');
    	
    	If HTML Then Texte.Add('<tr><td width="300" style="border-bottom:solid">Intitulé</td><td>'+RechDom('PGSTAGEFORM',T.GetValue('PFI_CODESTAGE'),False)+'</td></tr>')
		Else Texte.Add(' - Intitulé                     : '+RechDom('PGSTAGEFORM',T.GetValue('PFI_CODESTAGE'),False));
        //PT20 On n'affiche plus le centre de formation . Par contre on affiche le lieu.
        // On ajoute le champ suivant pour plus de facilité
(*        Q := OpenSQL('SELECT PST_CENTREFORMGU FROM STAGE WHERE PST_CODESTAGE="'+T.GetValue('PFI_CODESTAGE')+'" AND PST_MILLESIME="0000"', True);
        If Not Q.EOF Then
            T.AddChampSupValeur('PFI_CENTREFORMGU', Q.FindField('PST_CENTREFORMGU').AsString);
        Ferme(Q);
        // Récupération de l'adresse du centre de formation
        If T.GetValue('PFI_CENTREFORMGU') <> '' Then
        Begin
		    Texte.Add(' - Centre de formation          : '+RechDom('PGCENTREFORMATION',T.GetValue('PFI_CENTREFORMGU'),False));
            Q := OpenSQL('SELECT ANN_ALRUE1,ANN_ALRUE2,ANN_ALCP,ANN_ALVILLE FROM ANNUAIRE WHERE ANN_GUIDPER="'+T.GetValue('PFI_CENTREFORMGU')+'"', True);
            If Not Q.EOF Then
            Begin
                Texte.Add('                                  '+Q.FindField('ANN_ALRUE1').AsString);
                If Q.FindField('ANN_ALRUE2').AsString <> '' Then
                Texte.Add('                                  '+Q.FindField('ANN_ALRUE2').AsString);
                Texte.Add('                                  '+Q.FindField('ANN_ALCP').AsString+' '+Q.FindField('ANN_ALVILLE').AsString);
                Texte.Add('');
            End;
            Ferme(Q);
        End;*)
        If T.GetValue('PFI_LIEUFORM') <> '' Then
        Begin
		    Q := OpenSQL('SELECT * FROM LIEUFORMATION WHERE PLF_LIEUFORM="'+T.GetValue('PFI_LIEUFORM')+'"', True);
		    If Not Q.EOF Then
		    Begin
		    	If HTML Then Texte.Add('<tr><td valign="top" style="border-bottom:solid">Lieu de formation</td><td>'+Q.FindField('PLF_LIBELLE').AsString)
	    		Else Texte.Add(' - Lieu de formation            : '+Q.FindField('PLF_LIBELLE').AsString);
		    	If Q.FindField('PLF_ADRESSE1').AsString <> '' Then Texte.Add(RC+'                                  '+Q.FindField('PLF_ADRESSE1').AsString);
		    	If Q.FindField('PLF_ADRESSE2').AsString <> '' Then Texte.Add(RC+'                                  '+Q.FindField('PLF_ADRESSE2').AsString);
		    	If Q.FindField('PLF_ADRESSE3').AsString <> '' Then Texte.Add(RC+'                                  '+Q.FindField('PLF_ADRESSE3').AsString);
	            Texte.Add(RC+'                                  '+Q.FindField('PLF_CODEPOSTAL').AsString+' '+Q.FindField('PLF_VILLE').AsString);
	            If HTML Then Texte.Add('</td></tr>');
		    End;
		    Ferme(Q);
        End;
        
        If HTML Then Texte.Add('<tr><td style="border-bottom:solid">Nature</td><td>'+RechDom('PGNATUREFORM',T.GetValue('PFI_NATUREFORM'),False)+'</td></tr>')
        Else Texte.Add(' - Nature                       : '+RechDom('PGNATUREFORM',T.GetValue('PFI_NATUREFORM'),False));
        
        If HTML Then Texte.Add('<tr><td>Durée prévue</td><td>'+IntToStr(T.GetValue('PFI_JOURSTAGE'))+' jour(s) / '+IntToStr(T.GetValue('PFI_DUREESTAGE')) + ' heure(s)</td></tr>')
		Else Texte.Add(' - Durée prévue                 : '+IntToStr(T.GetValue('PFI_JOURSTAGE'))+' jour(s) / '+IntToStr(T.GetValue('PFI_DUREESTAGE')) + ' heure(s)');
    End;
begin
	If PGBundleHierarchie Then Tablette := 'PGSALARIEINT'
	Else Tablette := 'PGSALARIE';
	
	If HTML Then RC := '<br>' Else RC := '';
	
	// Saisie d'un salarié dans le cadre du plan
	If Action = 'SAISIEPLF' Then
	Begin	
		Titre := 'Demande de formation du salarié '+RechDom(Tablette, Salarie, False);

		Texte := HTStringList.Create;
		If HTML Then Texte.Add('<html>');
		
		Texte.Add(RechDom(Tablette, Salarie, False)+' a posé une demande pour la formation suivante ');
		Texte.Add('dans le cadre du '+RechDom('PGTYPEPLANPREV',RightStr(Action,3),False)+'.');
		Texte.Add(RC);
		
		Chaine := 'Projet de formation :';
		If HTML Then Texte.Add('<h3 style="background=''#CEDDF0''">'+Chaine+'</h3>') 
		Else Begin Texte.Add(Chaine); Texte.Add('---------------------'); End;
		
		AfficheRappelFormation(TobInfos);
		If HTML Then Texte.Add('</table>');
		Texte.Add(RC);
		If HTML Then Texte.Add('</html>');
	End
	// Décision du responsable dans le cadre du plan
	Else If (Action = 'DECISIONPLF') Or (Action = 'DECISIONDIF') then
	Begin
		Texte := HTStringList.Create;
		If HTML Then Texte.Add('<html>');

		// Cas des cursus : plusieurs lignes. On prend 
		If TobInfos.Detail.Count > 1 Then Cursus := True
		Else Cursus := False;

        If TobInfos.Detail.Count = 0 Then T := TobInfos
        Else T := TobInfos.Detail[0];
		
		Chaine := 'Le responsable '+RechDom(Tablette, Responsable, False)+' a ';
		If T.GetValue('PFI_ETATINSCFOR') = 'VAL' Then
		Begin
			Titre := 'Validation de votre demande de formation';
			Chaine := Chaine + 'validé';
		End
		Else If T.GetValue('PFI_ETATINSCFOR') = 'REF' Then
		Begin
			Titre := 'Refus de votre demande de formation';
			Chaine := Chaine + 'invalidé';
		End
		Else If T.GetValue('PFI_ETATINSCFOR') = 'REP' Then
		Begin
			Titre := 'Report de votre demande de formation';
			Chaine := Chaine + 'reporté';
		End
		Else If T.GetValue('PFI_ETATINSCFOR') = 'NAN' Then
		Begin
			Titre := 'Annulation de votre demande de formation';
			Chaine := Chaine + 'annulé';
		End;
		Chaine := Chaine + ' votre demande pour la formation suivante ';
		
		Texte.Add(Chaine);
		Texte.Add('dans le cadre du '+RechDom('PGTYPEPLANPREV',RightStr(Action,3),False)+'.');
		If HTML Then Texte.Add(RC);
		
		If T.GetValue('PFI_ETATINSCFOR') <> 'VAL' Then
        If T.GetValue('PFI_MOTIFETATINSC') <> '' Then
        Begin
        	Texte.Add(RC);
        	Chaine := 'Le motif est le suivant : ';
        	If HTML Then Chaine := Chaine + '<b>';
        	Chaine := Chaine + RechDom ('PGMOTIFETATINSC',T.GetValue('PFI_MOTIFETATINSC'),False);
        	If HTML Then Chaine := Chaine + '</b>'; 
	        Texte.Add(Chaine);
        End;
		Texte.Add(RC);
		
		Chaine := 'Rappel de votre projet de formation :';
		If HTML Then Texte.Add ('<h3 style="background=''#CEDDF0''">'+Chaine+'</h3>')
		Else Begin Texte.Add(Chaine); Texte.Add('-------------------------------------'); End;
		
		If Cursus Then
		Begin
			Chaine := 'Intitulé du cursus : ';
			If HTML Then Chaine := Chaine + '<b>' + RechDom('PGCURSUS',T.GetValue('PFI_CURSUS'),False) + '</b>'
			Else Chaine := Chaine + RechDom('PGCURSUS',T.GetValue('PFI_CURSUS'),False);
			Texte.Add(Chaine);
			
			T := TobInfos.FindFirst(['PFI_CODESTAGE'],['--CURSUS--'],False);
			If T <> Nil Then
			Texte.Add('      Durée totale : '+IntToStr(T.GetValue('PFI_JOURSTAGE'))+' jour(s) / '+IntToStr(T.GetValue('PFI_DUREESTAGE')) + ' heure(s)');
			Texte.Add(RC);
			Texte.Add(RC);
			
			For i := 0 To TobInfos.Detail.Count-1 Do
			Begin
				T := TobInfos.Detail[i];
				If T.GetValue('PFI_CODESTAGE') <> '--CURSUS--' Then
				Begin
					AfficheRappelFormation(T);
					If HTML Then Texte.Add('</table>');
					Texte.Add(RC);
				End;
			End;
		End
		Else
		Begin
			AfficheRappelFormation(T);
			If HTML Then Texte.Add('</table>');
		End;
		Texte.Add(RC);
		
		If Action = 'DECISIONDIF' Then
		Begin
			//PT26 - Compteurs uniquement si gestion de la paie
			If GetParamSocSecur('SO_PGCUMULDIFACQUIS', '') <> '' Then
			Begin
				Demande := 0.0;
				Acquis  := 0;
				Pris    := 0;
			
				Q := OpenSQL('SELECT SUM(PFO_NBREHEURE) HEURES FROM FORMATIONS WHERE PFO_EFFECTUE="X" AND PFO_TYPEPLANPREV="DIF" AND PFO_SALARIE="'+Salarie+'"',True);
				If Not Q.Eof then Pris := Q.FindField('HEURES').AsFloat;
				Ferme(Q);
				
				Cumul := GetParamSoc('SO_PGCUMULDIFACQUIS');
				Q := OpenSQL('SELECT SUM(PHC_MONTANT) ACQUIS '+
							 'FROM HISTOCUMSAL WHERE PHC_CUMULPAIE="'+Cumul+'" AND PHC_SALARIE="'+Salarie+'"',True);
				If Not Q.Eof then Acquis := Q.FindField('ACQUIS').AsFloat;
				Ferme(Q);
				
				Q := OpenSQL('SELECT SUM (PFI_DUREESTAGE) NBHEURES FROM INSCFORMATION '+
							 'WHERE PFI_ETATINSCFOR="VAL" AND PFI_CODESTAGE<>"--CURSUS--" AND PFI_REALISE="-" AND PFI_TYPEPLANPREV="DIF" AND PFI_SALARIE="'+Salarie+'"',True);
				If Not Q.Eof then Demande := Q.FindField('NBHEURES').AsFloat;
				Ferme(Q);
				
				Solde        := Acquis - Pris;
				SoldePrev    := Acquis - Pris - Demande;
			
				Texte.Add(RC);
				
				Chaine := 'Etat de mon compteur DIF :';
				If HTML Then Texte.Add('<h3 style="background=''#CEDDF0''">'+Chaine+'</h3>')
				Else Begin Texte.Add(Chaine);Texte.Add('-------------------------------');End;
				
				If HTML Then Texte.Add('<table border=0 cellspacing=1 cellpadding=1>');
				Chaine := 'Heures acquises';
				If HTML Then Texte.Add('<tr><td style="border-bottom:solid">'+Chaine+'</td><td>'+FloatToStr(Acquis)+'</td></tr>')
				Else Texte.Add(' - '+Chaine+'    : '+FloatToStr(Acquis));
				
				Chaine := 'Heures prises';
				If HTML Then Texte.Add('<tr><td style="border-bottom:solid">'+Chaine+'</td><td>'+FloatToStr(Pris)+'</td></tr>')
				Else Texte.Add(' - '+Chaine+'      : '+FloatToStr(Pris));

				Chaine := 'Solde';
				If HTML Then Texte.Add('<tr><td style="border-bottom:solid">'+Chaine+'</td><td>'+FloatToStr(Solde)+' heure(s)</td></tr>')
				Else Texte.Add(' - '+Chaine+'              : '+FloatToStr(Solde) + ' heure(s)');

				Chaine := 'Solde prévisionnel';
				If HTML Then Texte.Add('<tr><td style="border-bottom:solid">'+Chaine+'</td><td>'+FloatToStr(SoldePrev)+' heure(s)</td></tr>')
				Else Texte.Add(' - '+Chaine+' : '+FloatToStr(SoldePrev)+ ' heure(s)');
				
				If HTML Then Texte.Add('</table>');
			End;
			Texte.Add(RC);
		End;

		If HTML Then Texte.Add('</html>');
	End
	// Saisie d'un salarié dans le cadre du DIF
	Else If Action = 'SAISIEDIF' then
	begin
		//If GetParamSocSecur('SO_PGDIFMAILSAISIE', False) = False Then Exit;
		AvantLe      := PlusDate(TobInfos.GetValue('PFI_DATEDIF'),30,'J');
	
		Titre := 'Accusé de réception - Demande de DIF de '+ TobInfos.GetValue('PFI_LIBELLE');
		
		Texte := HTStringList.Create;
		If HTML Then Texte.Add('<html>');
		
		Texte.Add('Objet : Demande de formation du '+DateToStr(TobInfos.GetValue('PFI_DATEDIF'))+ ' au titre du DIF.'+RC);
		Texte.Add(RC);
		
		Chaine := 'Je dois répondre à '+RechDom(Tablette, Salarie, False)+' au maximum le ';
		If HTML Then Chaine := Chaine + '<b><u><font color="#993300">';
		Chaine := Chaine + DateToStr(AvantLe);
		If HTML Then Chaine := Chaine + '</font></u></b>';
		Texte.Add(Chaine+'.'+RC);
		Texte.Add(RC);
		
		Chaine := 'Projet de formation :';
		If HTML Then Texte.Add('<h3 style="background=''#CEDDF0''">'+Chaine+'</h3>')
		Else Begin Texte.Add(Chaine);Texte.Add('---------------------');End;
		
		AfficheRappelFormation(TobInfos);
		
		Chaine := 'dont '+FloatToStr(TobInfos.GetValue('PFI_HTPSTRAV'))+' sur le temps de travail';
		If HTML Then Texte.Add('<tr><td style="border-bottom:solid"></td><td>'+Chaine)
		Else Texte.Add('                                   '+Chaine);

		Chaine := 'et '+FloatToStr(TobInfos.GetValue('PFI_HTPSNONTRAV'))+' hors temps de travail.';
		If HTML Then Texte.Add(RC+Chaine+'</td></tr>')
		Else Texte.Add('                                   '+Chaine);
		
		If TobInfos.GetValue('PFI_MOTIFINSCFOR') <> '' Then
		Begin
			Chaine := 'Motif d''inscription';
			If HTML Then Texte.Add('<tr><td style="border-bottom:solid">'+Chaine+'</td><td>'+RechDom('PGMOTIFINSCFORMATION',TobInfos.GetValue('PFI_MOTIFINSCFOR'),False)+'</td></tr>')
			Else Texte.Add(' - '+Chaine+'          : '+RechDom('PGMOTIFINSCFORMATION',TobInfos.GetValue('PFI_MOTIFINSCFOR'),False));
		End;
		
		If HTML Then Texte.Add('</table>');
		Texte.Add(RC);
		
		Chaine := 'Planning :';
		If HTML Then Texte.Add('<h3 style="background=''#CEDDF0''">'+Chaine+'</h3>')
		Else Begin Texte.Add(Chaine); Texte.Add('----------'); End;
		
		If HTML Then Texte.Add('<table border=0 cellspacing=1 cellpadding=1>');
		
		Chaine := 'Date de la demande';
		If HTML Then Texte.Add('<tr><td width="300" style="border-bottom:solid">'+Chaine+'</td><td>'+DateToStr(TobInfos.GetValue('PFI_DATEDIF'))+'</td></tr>')
		Else Texte.Add(' - '+Chaine+' : '+DateToStr(TobInfos.GetValue('PFI_DATEDIF')));
		
		{$IFNDEF EAGLSERVER} // Pas pour le portail
		Chaine := 'Date de saisie sous PGI Formation';
		If HTML Then Texte.Add('<tr><td style="border-bottom:solid">'+Chaine+'</td><td>'+DateToStr(TobInfos.GetValue('PFI_DATECREATION'))+'</td></tr>')
		Else Texte.Add(' - '+Chaine+' : '+DateToStr(TobInfos.GetValue('PFI_DATECREATION')));
		{$ENDIF}
		
		Chaine := 'Date maxi de la réponse par le responsable';
		If HTML Then Texte.Add('<tr><td style="border-bottom:solid">'+Chaine+'</td><td>'+DateToStr(AvantLe)+'</td></tr>')
		Else  Texte.Add(' - '+Chaine+' : '+DateToStr(AvantLe));

		If HTML Then Texte.Add('</table>');
		Texte.Add(RC);
		
		//PT26 - Compteurs uniquement si Paie gérée par PGI
		If GetParamSocSecur('SO_PGCUMULDIFACQUIS', '') <> '' Then
		Begin
			Demande := 0.0;
			Acquis  := 0;
			Pris    := 0;
		
			Q := OpenSQL('SELECT SUM(PFO_NBREHEURE) HEURES FROM FORMATIONS WHERE PFO_EFFECTUE="X" AND PFO_TYPEPLANPREV="DIF" AND PFO_SALARIE="'+Salarie+'"',True);
			If Not Q.Eof then Pris := Q.FindField('HEURES').AsFloat;
			Ferme(Q);
			
			Cumul := GetParamSoc('SO_PGCUMULDIFACQUIS');
			Q := OpenSQL('SELECT SUM(PHC_MONTANT) ACQUIS '+
						 'FROM HISTOCUMSAL WHERE PHC_CUMULPAIE="'+Cumul+'" AND PHC_SALARIE="'+Salarie+'"',True);
			If Not Q.Eof then Acquis := Q.FindField('ACQUIS').AsFloat;
			Ferme(Q);
			
			Q := OpenSQL('SELECT SUM (PFI_DUREESTAGE) NBHEURES FROM INSCFORMATION '+
						 'WHERE PFI_ETATINSCFOR="VAL" AND PFI_CODESTAGE<>"--CURSUS--" AND PFI_REALISE="-" AND PFI_TYPEPLANPREV="DIF" AND PFI_SALARIE="'+Salarie+'"',True);
			If Not Q.Eof then Demande := Q.FindField('NBHEURES').AsFloat;
			Ferme(Q);
			
			Solde        := Acquis - Pris;
			SoldePrev    := Acquis - Pris - Demande;
		
			Chaine := 'Compteur DIF du collaborateur :';
			If HTML Then Texte.Add('<h3 style="background=''#CEDDF0''">'+Chaine+'</h3>')
			Else Begin Texte.Add(Chaine);Texte.Add('-------------------------------');End;
			
			If HTML Then Texte.Add('<table border=0 cellspacing=1 cellpadding=1>');
			Chaine := 'Heures acquises';
			If HTML Then Texte.Add('<tr><td width="300" style="border-bottom:solid">'+Chaine+'</td><td>'+FloatToStr(Acquis)+'</td></tr>')
			Else Texte.Add(' - '+Chaine+'    : '+FloatToStr(Acquis));
			
			Chaine := 'Heures prises';
			If HTML Then Texte.Add('<tr><td style="border-bottom:solid">'+Chaine+'</td><td>'+FloatToStr(Pris)+'</td></tr>')
			Else Texte.Add(' - '+Chaine+'      : '+FloatToStr(Pris));

			Chaine := 'Solde';
			If HTML Then Texte.Add('<tr><td style="border-bottom:solid">'+Chaine+'</td><td>'+FloatToStr(Solde)+' heure(s)</td></tr>')
			Else Texte.Add(' - '+Chaine+'              : '+FloatToStr(Solde) + ' heure(s)');

			Chaine := 'Solde prévisionnel';
			If HTML Then Texte.Add('<tr><td style="border-bottom:solid">'+Chaine+'</td><td>'+FloatToStr(SoldePrev)+' heure(s)</td></tr>')
			Else Texte.Add(' - '+Chaine+' : '+FloatToStr(SoldePrev)+ ' heure(s)');
			
			If HTML Then Texte.Add('</table>');
			Texte.Add(RC);
		End
		Else
			Texte.Add('Pensez à consulter le compteur DIF du collaborateur avant de valider la demande.');
		Texte.Add(RC);
		Texte.Add(RC);
		//PT26 - Suppression du parpaing ci-dessous : propre à Cegid
		(*Texte.Add('Le cadre de validation des demandes de DIF est le plan de formation et la gestion des compétences du groupe.');
		Texte.Add('');
		Texte.Add('Par principe, les responsables peuvent valider les formations internes.');
		Texte.Add('Les formations externes seront validées dans le cadre du budget et avec l''accord de la DRH.');
		Texte.Add('');
		Texte.Add('Attention, l''absence de réponse équivaut à l''accord pour la formation.');
		Texte.Add('');
		Texte.Add('La réponse "oui", implique que la formation devra se mettre en oeuvre en '+TobInfos.GetValue('PFI_MILLESIME')+'.');
		Texte.Add('C''est au responsable de planifier l''organisation de la session.');
		Texte.Add('');
		Texte.Add('Le service formation de la DRH est à votre disposition pour tout support.');
		Texte.Add('');
		Texte.Add('Dès que la réponse sera saisie dans le module Formation/DIF, vous recevrez un e-mail de réponse à transmettre et à faire signer au collaborateur.');*)
		Texte.Add('Attention, l''absence de réponse équivaut à l''accord pour la formation.');
		If HTML Then Texte.Add('</html>');
    End
    //PT29
	// Annulation d'une demande de DIF
	Else If Action = 'ANNULATIONDIF' then
	begin
		Titre := 'Annulation de la demande de DIF de '+ TobInfos.GetValue('PFI_LIBELLE');

		Texte := HTStringList.Create;
		If HTML Then Texte.Add('<html>');
		
		Texte.Add(TobInfos.GetValue('PFI_LIBELLE') +' a annulé sa demande de formation datant du '+DateToStr(TobInfos.GetValue('PFI_DATEDIF'))+ ' au titre du DIF.');
		Texte.Add(RC);
		If HTML Then Texte.Add('</html>');
	//PT27
	End
	// Annulation des formations suite à sortie du salarié
	Else If Action = 'SORTIESALARIE' then
	begin
		Titre := 'Sortie du salarié '+RechDom(Tablette, Salarie, False);

		Texte := HTStringList.Create;
		Texte.Add('<html>');
		Texte.Add('Suite à la sortie ');
        If TobInfos.GetValue('PSA_DATESORTIE') > Date Then Texte.Add('prochaine ');
        Texte.Add('du salarié '+RechDom(Tablette, Salarie, False));
		Texte.Add(' en date du '+DateToStr(TobInfos.GetValue('PSA_DATESORTIE')));
		Texte.Add(', les demandes de formation suivantes ont été annulées automatiquement :<br>');
        Texte.Add('<br>');
        Texte.Add('<ul type="circle">');
        T2 := TobInfos.FindFirst(['TYPEFOR'],['REAL'], True);
        If T2 <> Nil Then
        Begin
        	Texte.Add('<li><h3>Plan de formation :</h3></li><ul type="square">');
	        While T2 <> Nil Do
	        begin
	            Texte.Add('<li><b>'+T2.GetValue('LIBELLE')+'</b> du '+DateToStr(T2.GetValue('PFO_DATEDEBUT'))+' au '+DateToStr(T2.GetValue('PFO_DATEFIN'))+' dans le cadre du <i>'+RechDom('PGTYPEPLANPREV',T2.GetValue('PFO_TYPEPLANPREV'),False)+'</i> anciennement dans l''état <i>"'+RechDom('PGETATVALIDATION',T2.GetValue('PFO_ETATINSCFOR'),False)+'"</i></li>');
	            T2 := TobInfos.FindNext(['TYPEFOR'],['REAL'], True);
	        end;
	        Texte.Add('</ul><br>');
	    End;
        T := TobInfos.FindFirst(['TYPEFOR'],['PREV'], True);
        If T <> Nil Then
        Begin
        	Texte.Add('<li><h3>Budget prévisionnel :</h3></li><ul type="square">');
	        While T <> Nil Do
	        begin
	            Texte.Add('<li><b>'+T.GetValue('LIBELLE')+'</b> pour le millésime '+T.GetValue('PFI_MILLESIME')+' dans le cadre du <i>'+RechDom('PGTYPEPLANPREV',T.GetValue('PFI_TYPEPLANPREV'),False)+'</i></li>');
	            T := TobInfos.FindNext(['TYPEFOR'],['PREV'], True);
	        end;
	        Texte.Add('</ul><br>');
	    End;
        Texte.Add('</ul>');
        Texte.Add('Il se peut que d''autres inscriptions soient également à annuler. Merci de vérifier les demandes validées et reportées.<br>');
        Texte.Add('</html>');
	End
	// Aucune annulation des formations suite à sortie du salarié
	Else If Action = 'SORTIESALARIENOP' then
	begin
		Titre := 'Sortie du salarié '+RechDom(Tablette, Salarie, False);

		Texte := HTStringList.Create;
		Texte.Add('<html>');
		Texte.Add('Suite à la sortie ');
        If TobInfos.GetValue('PSA_DATESORTIE') > Date Then Texte.Add('prochaine ');
        Texte.Add('du salarié '+RechDom(Tablette, Salarie, False));
		Texte.Add(' en date du '+DateToStr(TobInfos.GetValue('PSA_DATESORTIE')));
		Texte.Add(', les demandes de formation suivantes sont peut-être à annuler :<br>');
        Texte.Add('<br>');
        Texte.Add('<ul type="circle">');
        T2 := TobInfos.FindFirst(['TYPEFOR'],['REAL'], True);
        If T2 <> Nil Then
        Begin
        	Texte.Add('<li><h3>Plan de formation :</h3></li><ul type="square">');
	        While T2 <> Nil Do
	        begin
	            Texte.Add('<li><b>'+T2.GetValue('LIBELLE')+'</b> du '+DateToStr(T2.GetValue('PFO_DATEDEBUT'))+' au '+DateToStr(T2.GetValue('PFO_DATEFIN'))+' dans le cadre du <i>'+RechDom('PGTYPEPLANPREV',T2.GetValue('PFO_TYPEPLANPREV'),False)+'</i> dans l''état <i>"'+RechDom('PGETATVALIDATION',T2.GetValue('PFO_ETATINSCFOR'),False)+'"</i></li>');
	            T2 := TobInfos.FindNext(['TYPEFOR'],['REAL'], True);
	        end;
	        Texte.Add('</ul><br>');
	    End;
        T := TobInfos.FindFirst(['TYPEFOR'],['PREV'], True);
        If T <> Nil Then
        Begin
        	Texte.Add('<li><h3>Budget prévisionnel :</h3></li><ul type="square">');
	        While T <> Nil Do
	        begin
	            Texte.Add('<li><b>'+T.GetValue('LIBELLE')+'</b> pour le millésime '+T.GetValue('PFI_MILLESIME')+' dans le cadre du <i>'+RechDom('PGTYPEPLANPREV',T.GetValue('PFI_TYPEPLANPREV'),False)+'</i></li>');
	            T := TobInfos.FindNext(['TYPEFOR'],['PREV'], True);
	        end;
	        Texte.Add('</ul><br>');
	    End;
        Texte.Add('</ul>');
        Texte.Add('Il se peut que d''autres inscriptions soient également à annuler. Merci de vérifier les demandes validées et reportées.<br>');
        Texte.Add('</html>');
	End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 14/03/2008 / PT16
Modifié le ... :   /  /
Description .. : Adapte une requête faisant référence à la table "Salaries"
Suite ........ : en la faisant pointer sur la table "Interimaires".
Mots clefs ... :
*****************************************************************}
Function AdaptMultiDosForm (ChaineSQL : String; SelectStatic : Boolean = True) : String; 
Var Select, Reste : String;
    PosFrom       : Integer;
Begin
	If PGBundleHierarchie Then
	Begin
		Select := '';
		Reste  := ChaineSQL;
		
		If SelectStatic Then
		Begin
			// Découpage de la chaîne
			PosFrom := Pos(' FROM ', ChaineSQL);
			Select := Copy(ChaineSQL,1,PosFrom);
			Reste  := Copy(ChaineSQL,PosFrom+1,Length(ChaineSQL));
			
			// Traitement de la clause SELECT
			Select := StringReplace (Select, 'PSA_SALARIE',       'PSI_INTERIMAIRE AS PSA_SALARIE',         [rfignorecase, rfreplaceall]);
			Select := StringReplace (Select, 'PSA_LIBELLE ',      'PSI_LIBELLE AS PSA_LIBELLE ',            [rfignorecase, rfreplaceall]);
			Select := StringReplace (Select, 'PSA_LIBELLE,',      'PSI_LIBELLE AS PSA_LIBELLE,',            [rfignorecase, rfreplaceall]);
			Select := StringReplace (Select, 'PSA_PRENOM',        'PSI_PRENOM AS PSA_PRENOM',               [rfignorecase, rfreplaceall]);
			Select := StringReplace (Select, 'PSA_LIBELLEEMPLOI', 'PSI_LIBELLEEMPLOI AS PSA_LIBELLEEMPLOI', [rfignorecase, rfreplaceall]);
			Select := StringReplace (Select, 'PSA_ETABLISSEMENT', 'PSI_ETABLISSEMENT AS PSA_ETABLISSEMENT', [rfignorecase, rfreplaceall]);
//PT33
			Select := StringReplace (Select, 'PSA_DATEENTREE',    'PSI_DATEENTREE AS PSA_DATEENTREE', 		[rfignorecase, rfreplaceall]);
			Select := StringReplace (Select, 'PSA_TRAVAILN1',    'PSI_TRAVAILN1 AS PSA_TRAVAILN1', 			[rfignorecase, rfreplaceall]);
			Select := StringReplace (Select, 'PSA_TRAVAILN2',    'PSI_TRAVAILN1 AS PSA_TRAVAILN2', 			[rfignorecase, rfreplaceall]);
			Select := StringReplace (Select, 'PSA_TRAVAILN3',    'PSI_TRAVAILN1 AS PSA_TRAVAILN3', 			[rfignorecase, rfreplaceall]);
			Select := StringReplace (Select, 'PSA_TRAVAILN4',    'PSI_TRAVAILN1 AS PSA_TRAVAILN4', 			[rfignorecase, rfreplaceall]);
//FIN PT33
		End;

		// Traitement du reste de la requête
		Reste  := StringReplace (Reste, 'SALARIES', 'INTERIMAIRES', [rfignorecase, rfreplaceall]);
		Reste  := StringReplace (Reste, 'PSA_SALARIE', 'PSI_INTERIMAIRE', [rfignorecase, rfreplaceall]);
		Reste  := StringReplace (Reste, 'PSA_', 'PSI_', [rfignorecase, rfreplaceall]);
		Result := Select + Reste;
	End
	Else
		Result := ChaineSQL;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 17/03/2008 / PT16
Modifié le ... :   /  /
Description .. : Retourne le numéro de dossier du salarié
Mots clefs ... :
*****************************************************************}
Function  GetNoDossierSalarie (Salarie : String = '') : String; 
Var Q : TQuery;
Begin
	Result := '';
	If Salarie = '' Then Salarie := V_PGI.UserSalarie;
	
	Q := OpenSQL('SELECT PSI_NODOSSIER FROM INTERIMAIRES WHERE PSI_INTERIMAIRE="'+Salarie+'"', True);
	If Not Q.EOF Then Result := Q.FindField('PSI_NODOSSIER').AsString;
	Ferme(Q);
End;


{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 02/04/2008 / PT21
Modifié le ... :   /  /
Description .. : Retourne la base centrale pour le partage des données
Mots clefs ... :
*****************************************************************}
Function GetBasePartage(Bundle : Integer) : String;
Begin
	If (Bundle > -1) And (Bundle < 10) Then
		Result := {$IFNDEF PLUGIN}VH_PAIE.{$ENDIF}BasesPartage[Bundle] //PT25
	Else
		Result := '';
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 02/04/2008 / PT21
Modifié le ... :   /  /
Description .. : Retourne le n° de dossier d'un salarié
Mots clefs ... :
*****************************************************************}
Function  GetBaseSalarie (Salarie : String = '') : String; 
Var Q       : TQuery;
    Dossier : String;
Begin
	Result := '';
	If Salarie = '' Then Salarie := V_PGI.UserSalarie;
	
	Dossier := GetNoDossierSalarie(Salarie);
	If Dossier <> '' Then
	Begin
		Q := OpenSQL('SELECT DOS_NOMBASE FROM DOSSIER WHERE DOS_NODOSSIER="'+Dossier+'"', True);
		If Not Q.EOF Then Result := Q.FindField('DOS_NOMBASE').AsString;
		Ferme(Q);
	End;
End;

end.

