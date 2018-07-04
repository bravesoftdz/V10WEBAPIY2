{***********UNITE*************************************************
Auteur  ...... : N. ACHINO et D. CARRET
Créé le ...... : 20/06/2001
Modifié le ... : 19/06/2002
Description .. : Assistant pour constituer le fichier de démarrage d'une
Suite ........ : caisse.
Mots clefs ... : FO
*****************************************************************}
unit FOExportCais;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls, Grids,
  HPanel, Mask, HRichEdt, HRichOLE, ParamSoc, LookUp, HEnt1, dbtables, UTOB,
  FileCtrl, FE_Main, SAISUTIL, FOToxUtil, Ent1, UtilPGI,
{$IFDEF EAGLCLIENT}
{$ELSE}
  MajTable,
{$ENDIF}
{$IFDEF NOMADE}
  uToxaPlat, EntGc,UtilNomade,UtilUtilitaires, LicUtil, UtilDispGC,
{$ELSE NOMADE}
  {$IFDEF NOMADESERVER}
  uToxaPlat, EntGc,UtilNomade,UtilUtilitaires, LicUtil,
  {$ENDIF NOMADESERVER}
{$ENDIF NOMADE}
  UTox, uToxConf, uToxClasses, uToxFiche, UToxNet, uToxConst, gcToxWork, FOImportCais,
  MenuOLG, Spin, HDB;

procedure FOExportCaisse(Caisse: string);
procedure PCPAssistInit(Repr: string);

type
  TFExportCais = class(TFAssist)
    PAGE1: TTabSheet;
    PTITRE1: THPanel;
    TITRE1: THLabel;
    LGPK_CAISSE: THLabel;
    GPK_CAISSE: THCritMaskEdit;
    PAGE3: TTabSheet;
    LSITECAISSE: THLabel;
    SITECAISSE: THCritMaskEdit;
    PTITRE3: THPanel;
    TITRE3: THLabel;
    PTITRE31: THPanel;
    TITRE31: THLabel;
    LSITESIEGE: THLabel;
    SITESIEGE: THCritMaskEdit;
    SO_LIBELLE: TEdit;
    LSO_LIBELLE: THLabel;
    PAGE4: TTabSheet;
    PAGE2: TTabSheet;
    PTITRE2: THPanel;
    TITRE2: THLabel;
    PTITRE4: THPanel;
    TITRE4: THLabel;
    TITRE41: THLabel;
    LNOMDOSSIER: THLabel;
    NOMDOSSIER: THCritMaskEdit;
    SO_REGIMEDEFAUT: THValComboBox;
    LSO_REGIMEDEFAUT: THLabel;
    LSO_GCMODEREGLEDEFAUT: THLabel;
    SO_GCMODEREGLEDEFAUT: THValComboBox;
    LSO_GCTIERSDEFAUT: THLabel;
    SO_GCTIERSDEFAUT: THCritMaskEdit;
    LSO_GCPERBASETARIF: THLabel;
    SO_GCPERBASETARIF: THValComboBox;
    LSO_GCPHOTOFICHE: THLabel;
    SO_GCPHOTOFICHE: THValComboBox;
    PAGE5: TTabSheet;
    PTITRE5: THPanel;
    TITRE5: THLabel;
    CPTRENDU: THRichEditOLE;
    TITRE51: THLabel;
    LIB_CAISSE: THLabel;
    PAGE7: TTabSheet;
    PTITRE7: THPanel;
    TITRE7: THLabel;
    LANCETOX: TCheckBox;
    LSEV_CODEEVENT: THLabel;
    SEV_CODEEVENT: THCritMaskEdit;
    BSITECAISSE: TToolbarButton97;
    BSITESIEGE: TToolbarButton97;
    BCODEEVENT: TToolbarButton97;
    US_UTILISATEUR: THMultiValComboBox;
    LUS_UTILISATEUR: THLabel;
    PAGE8: TTabSheet;
    PTITRE8: THPanel;
    TITRE8: THLabel;
    GEVTS: THGrid;
    PAGE9: TTabSheet;
    PTITRE9: THPanel;
    TITRE9: THLabel;
    TITRE91: THLabel;
    LSO_DEVISEPRINC: THLabel;
    SO_DEVISEPRINC: THValComboBox;
    SO_TENUEEURO: TCheckBox;
    PAGE21: TTabSheet;
    LREPRESENTANT: THLabel;
    LNOMREP: THLabel;
    HPanel1: THPanel;
    TITREPOP: THLabel;
    US_UTILISATEUR_: THCritMaskEdit;
    US_LIBELLE: TEdit;
    lEtapeManu: THLabel;
    DBCENTRAL: TComboBox;
    TDBCENTRAL: TLabel;
    PB: TProgressBar;
    PAGE22: TTabSheet;
    PTITRE22: THPanel;
    TITR22: THLabel;
    GEVTINIT: THGrid;
    LstRecap: TListBox;
    PAGE211: TTabSheet;
    PAGE212: TTabSheet;
    GP_CODEIFART: TGroupBox;
    LSO_GCLGNUMART_: THLabel;
    LSO_GCPREFIXEART_: THLabel;
    LSO_GCCOMPTEURART_: THLabel;
    SO_GCLGNUMART_: THDBSpinEdit;
    SO_GCPREFIXEART_: TEdit;
    SO_GCCOMPTEURART_: THNumEdit;
    HPanel2: THPanel;
    HLabel1: THLabel;
    SO_LIBELLE_: TEdit;
    LSO_MAIL: THLabel;
    SO_MAIL_: TEdit;
    LSO_ADRESSE: THLabel;
    SO_ADRESSE1_: TEdit;
    LSO_CODEPOSTAL: THLabel;
    SO_CODEPOSTAL_: TEdit;
    LSO_DIVTERRIT: THLabel;
    SO_DIVTERRIT_: TEdit;
    SO_VILLE_: TEdit;
    LSO_VILLE: THLabel;
    SO_ADRESSE2_: TEdit;
    SO_ADRESSE3_: TEdit;
    LSO_PAYS: THLabel;
    SO_PAYS_: THValComboBox;
    SO_FAX_: TEdit;
    LSO_FAX: THLabel;
    SO_TELEPHONE_: TEdit;
    LSO_TELEPHONE: THLabel;
    HLabel2: THLabel;
    HPanel3: THPanel;
    HLabel3: THLabel;
    GP_CODIFTIERS: TGroupBox;
    LSO_GCLGNUMTIERS_: THLabel;
    LSO_GCPREFIXETIERS_: THLabel;
    LSO_GCCOMPTEURTIERS_: THLabel;
    SO_GCLGNUMTIERS_: THDBSpinEdit;
    SO_GCPREFIXETIERS_: TEdit;
    SO_GCCOMPTEURTIERS_: THNumEdit;
    LSO_GCPREFIXEAUXI_: THLabel;
    SO_GCPREFIXEAUXI_: TEdit;
    SO_GCPREFIXEAUXIFOU_: TEdit;
    LSO_GCPREFIXEAUXIFOU_: THLabel;
    LCONFIGPCP: THLabel;
    CONFIGPCP: TEdit;
    bImprimer: TToolbarButton97;
    GP_PARAMPCES: TGroupBox;
    GP_VERROUILLEPCES_: TCheckBox;
    GP_IMPIMMEDIATE_: TCheckBox;
    GP_APERCUAVIMP_: TCheckBox;
    GP_PARAMPCES_: TCheckBox;
    bParamPce: TToolbarButton97;
    procedure bPrecedentClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SO_LIBELLEEnter(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure LANCETOXClick(Sender: TObject);
    procedure BSITECAISSEClick(Sender: TObject);
    procedure SITECAISSEDblClick(Sender: TObject);
    procedure BSITESIEGEClick(Sender: TObject);
    procedure SITESIEGEDblClick(Sender: TObject);
    procedure SEV_CODEEVENTDblClick(Sender: TObject);
    procedure BCODEEVENTClick(Sender: TObject);
    procedure SO_DEVISEPRINCChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure US_UTILISATEUR_Change(Sender: TObject);
    procedure DBCENTRAL_OnChange(Sender: TObject);
{$IFDEF NOMADE}
    procedure bImprimerClick(Sender: TObject);
    procedure GP_IMPIMMEDIATE_Click(Sender: TObject);
    procedure bParamPceClick(Sender: TObject);
    procedure bAnnulerClick(Sender: TObject);
{$ENDIF}

  private
    { Déclarations privées }
		Superviseur : boolean;
    CodeCaisse: string;
    TOBPCaisse: TOB;
    TOBEtab: TOB;
    CodeUsr: string;
    TOBRepr: TOB;
    CurrentPage: integer; // Numéro de la page courante
    PagesCount: integer; // Nb de pages utilisées
    SocForm: THValComboBox; // Sélection de la base central
    BasePCP: string; // Nom de la base PCP
    bFinOk: Boolean;
    TobRepresentant : TOB; // Gestion PCP par user
    TobNatPces : TOB; // Pièces à gérer dans PCP
    LibSite : string;
    EvtOblInit, EvtOblRemontee : string; // Contient les évènements obligatoire par rapport au paramètrage user
    procedure ActiveBtn(Active: Boolean);
    procedure AfficheMemo(Texte: string; Style: TFontStyles; Saut, Traduire: Boolean);
    procedure ChargeSiteCentral;
    procedure ChargeSiteCaisse;
    procedure SauveCaisse;
    procedure SauvePCP;
    function NomFichierConf(AvecChemin: Boolean): string;
    procedure ChangeVariable(TOBL: TOB; NomChamp, NomVar, Valeur: string);
    procedure ChangeInfoComp(TOBL: TOB; NomChamp, NomInfo, TypeInfo, Valeur: string);
    procedure ExportTablette(NomTable: string);
    procedure ExportTable(NomTable: string);
    procedure ExportSite(TOBData: TOB);
    procedure ExportParamSoc;
    procedure AffecteSouche(TOBData: TOB);
    procedure FabriqueDemarrage;
    procedure AffEtape;
    procedure LoadData;
    procedure PCPDeconnecteDB;
    {$IFDEF NOMADE}
    procedure LogToxAPlat(PhaseId: Cardinal; MessageStr: array of variant);
    function ConnectBase(QuelleBase : string; VerifBase, BaseUsr : boolean) : boolean;
    function TestDataBase : boolean;
    function TestDataUser : boolean;
    procedure CreerRapport;
    {$ENDIF}
    procedure GereErreur(NumErreur : integer; Chps : THEdit; Txt : string);
  public
    { Déclarations publiques }
  end;

const
  ouFO = 1;
  ouPCP = 2;
  // Nom des pages disponibles pour chaque environnement (FO/PCP) -> P.Pages[i].Name
  OngletUtile: array[1..2] of string = (
    {1 FO}';PAGE1;PAGE2;PAGE3;PAGE4;PAGE5;PAGE6;PAGE7;PAGE8;PAGE9;'
    {2 PCP}, ';PAGE21;PAGE211;PAGE212;PAGE22;PAGE8;PAGE5;'
    );

var
  BC_LesSites: TCollectionSites; // Sites de la Base Centrale.

implementation

{$R *.DFM}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/06/2001
Modifié le ... : 20/06/2001
Description .. : Constitution du fichier de démarrage d'une caisse
Mots clefs ... : FO
*****************************************************************}

procedure FOExportCaisse(Caisse: string);
var X: TFExportCais;
begin
  X := TFExportCais.Create(Application);
  try
    X.CodeCaisse := Caisse;
    X.ShowModal;
  finally
    X.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : D. Carret
Créé le ...... : 04/07/2002
Modifié le ... : 04/07/2002
Description .. : Constitution du fichier de démarrage pour un utilisateur
Mots clefs ... : PCP
*****************************************************************}
procedure PCPAssistInit(Repr: string);
var X: TFExportCais;
begin
  X := TFExportCais.Create(Application);
  try
    X.CodeUsr := Repr;
    // Mise en place des libellés spécifiques de l'assistant PCP
    X.Caption := TraduireMemoire('Assistant d''initialisation d''un poste Nomade');
    X.TITRE3.Caption := TraduireMemoire('Indiquez le site qui correspond à l''utilisateur');
    X.LSITECAISSE.Caption := TraduireMemoire('Choix du site utilisateur');
    X.TITRE8.Caption := TraduireMemoire('Indiquez les évènements qui seront exécutés depuis le portable');
    X.TITRE4.Caption := TraduireMemoire('Indiquez le dossier de stockage du fichier de démarrage');
    X.TITRE5.Caption := TraduireMemoire('Le bouton ''Fin'' lance l''initialisation de la base utilisateur');
    X.ShowModal;
  finally
    X.Free;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  EstDansLaListe : Vérifie si une chaîne de caractères appartient à une liste au format multi-combo
///////////////////////////////////////////////////////////////////////////////////////

function EstDansLaListe(Chaine, Liste: string): Boolean;
begin
  Result := False;
  if Chaine = '' then Exit;
  if Liste = '' then Exit;
  if Liste[1] <> ';' then Liste := ';' + Liste;
  if Liste[Length(Liste)] <> ';' then Liste := Liste + ';';
  if Chaine[1] <> ';' then Chaine := ';' + Chaine;
  if Chaine[Length(Chaine)] <> ';' then Chaine := Chaine + ';';
  Result := (Pos(Chaine, Liste) <> 0);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FabriqueSQLIN : Fabrique une clause IN pour un SELECT à partir des codes sélectionnés d'un multi-combo
///////////////////////////////////////////////////////////////////////////////////////

function FabriqueSQLIN(Liste, NomChamp: string; WithAnd, NotIn, ChampNumeric: Boolean): string;
var Stg, Operateur, Separateur: string;
begin
  Result := '';
  if (Liste = '') or (NomChamp = '') then Exit;
  Stg := Liste;
  Stg := Trim(Stg);
  // Cas du <<Tous>>
  if (Pos('<<', Stg) = 1) and (Pos('>>', Stg) = Length(Stg) - 1) then Stg := '';
  if Stg <> '' then
  begin
    if Stg[Length(Stg)] = ';' then Delete(Stg, Length(Stg), 1);
    if WithAnd then Result := ' and ';
    if ChampNumeric then Separateur := '' else Separateur := '"';
    Stg := Separateur + FindEtReplace(Stg, ';', Separateur + ',' + Separateur, True) + Separateur;
    if NotIn then Operateur := 'not in' else Operateur := 'in';
    Result := Result + NomChamp + ' ' + Operateur + ' (' + Stg + ') '
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ActiveBtn : active les boutons de l'assistant
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.ActiveBtn(Active: Boolean);
begin
  bAide.Enabled := Active;
  bAnnuler.Enabled := Active;
  bPrecedent.Enabled := Active;
  bSuivant.Enabled := Active;
  bImprimer.Enabled := Active;
  if (CurrentPage = PagesCount) then bFin.Enabled := Active;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AfficheMemo : affiche un texte dans le mémo.
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.AfficheMemo(Texte: string; Style: TFontStyles; Saut, Traduire: Boolean);
var Stg: string;
begin
  if (P.ActivePage.Name = 'PAGE5') and (CPTRENDU <> nil) then
  begin
    if Saut then CPTRENDU.lines.Append('');
    CPTRENDU.SelAttributes.Style := Style;
    if Traduire then Stg := TraduireMemoire(Texte) else Stg := Texte;
    CPTRENDU.lines.Append(Stg);
    LstRecap.Items.Add(Stg);
    if CPTRENDU.CanFocus then CPTRENDU.SetFocus;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeSiteCentral : Recherche du site central
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.ChargeSiteCentral;
begin
  with TCollectionSites.Create(TCollectionSite, True) do
  begin
    if LeSiteLocal <> nil then
    begin
      SITESIEGE.Text := LeSiteLocal.SSI_CODESITE;
      NOMDOSSIER.Text := LeSiteLocal.SSI_CDEPART;
    end;
    Free;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeSiteCaisse : Recherche du site de la caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.ChargeSiteCaisse;
var LeSite: TCollectionSite;
begin
  with TCollectionSites.Create(TCollectionSite, True) do
  begin
    LeSite := Find(TobPCaisse.GetValue('GPK_ETABLISSEMENT'));
    if LeSite <> nil then
      SITECAISSE.Text := LeSite.SSI_CODESITE
      else
      SITECAISSE.Text := GPK_CAISSE.Text;
    Free;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SauveCaisse : construction du fichier de démarrage de la caisse.
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.SauveCaisse;
var Fichier: string;
begin
  // Suppression du fichier de démarrage
  Fichier := NomFichierConf(True);
  if Fichier = '' then Exit;
  DeleteFile(Fichier);
  // Enregistrement des Régimes de facturation
  AfficheMemo('Régimes de facturation', [fsItalic], False, True);
  ExportTablette('TTREGIMETVA');
  ExportTablette('TTTVA');
  // Enregistrement des Photos pour les fiches
  AfficheMemo('Photos pour les fiches', [fsItalic], False, True);
  ExportTablette('GCEMPLOIBLOB');
  // Enregistrement des Périodes de base pour les tarifs
  AfficheMemo('Périodes de base pour les tarifs', [fsItalic], False, True);
  ExportTable('TARIFPER');
  // Enregistrement des Devises
  AfficheMemo('Devises', [fsItalic], False, True);
  ExportTable('DEVISE');
  // Enregistrement des Modes de règlement
  AfficheMemo('Modes de règlement', [fsItalic], False, True);
  ExportTable('MODEREGL');
  // Enregistrement des Modes de paiement
  AfficheMemo('Modes de paiement', [fsItalic], False, True);
  ExportTable('MODEPAIE');
  // Enregistrement des Tiers
  AfficheMemo('Tiers', [fsItalic], False, True);
  ExportTable('TIERS');
  // Enregistrement des Etablissements
  AfficheMemo('Etablissements', [fsItalic], False, True);
  ExportTable('ETABLISS');
  // Enregistrement des Dépôts
  AfficheMemo('Dépôts', [fsItalic], False, True);
  ExportTable('DEPOTS');
  // Enregistrement des Caisses
  AfficheMemo('Caisses', [fsItalic], False, True);
  ExportTable('PARCAISSE');
  // Enregistrement des Souches
  AfficheMemo('Souches', [fsItalic], False, True);
  ExportTable('SOUCHE');
  // Enregistrement des Sites
  AfficheMemo('Variables de la TOX', [fsItalic], False, True);
  ExportTable('STOXVARS');
  AfficheMemo('Sites', [fsItalic], False, True);
  ExportTable('STOXSITES');
  AfficheMemo('Informations complémentaires des sites', [fsItalic], False, True);
  ExportTable('STOXINFOCOMP');
  AfficheMemo('Requêtes de la TOX', [fsItalic], False, True);
  ExportTable('STOXQUERYS');
  AfficheMemo('Evenements de la TOX', [fsItalic], False, True);
  ExportTable('STOXEVENTS');
  // Enregistrement des Utilisateurs
  AfficheMemo('Groupes d''utilisateurs', [fsItalic], False, True);
  ExportTable('USERGRP');
  AfficheMemo('Confidentialité', [fsItalic], False, True);
  ExportTable('MENU');
  AfficheMemo('Utilisateurs', [fsItalic], False, True);
  ExportTable('UTILISAT');
  // Enregistrement des Paramètres société
  AfficheMemo('Paramètres société', [fsItalic], False, True);
  ExportParamSoc;
  AfficheMemo('Construction du fichier de démarrage terminée', [fsBold, fsUnderline], False, True);
  AfficheMemo('  => ' + Fichier, [], False, False);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SauvePCP : construction du fichier de démarrage pour un poste PCP.
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.SauvePCP;
{$IFDEF NOMADE}
var Fichier: string;
  YY: TFImportCais;
  Toxcr: boolean;
  FormDBName: string;
  NbOcc, Ind: integer;
  {$ENDIF}
begin
  {$IFDEF NOMADE}
  // Suppression du fichier de démarrage
  Fichier := NomFichierConf(True);
  if Fichier = '' then Exit;
  DeleteFile(Fichier);
  AfficheMemo('Execices', [fsItalic], False, True);                    ExportTable('EXERCICE');
  AfficheMemo('Comptes comptable', [fsItalic], False, True);           ExportTable('GENERAUX');
  AfficheMemo('Etablissements', [fsItalic], False, True);              ExportTable('ETABLISS');
  AfficheMemo('Dépôts', [fsItalic], False, True);                      ExportTable('DEPOTS');
  AfficheMemo('Exceptions sur pièces', [fsItalic], False, True);       ExportTable('PARPIECE'); ExportTable('SITEPIECE');
  AfficheMemo('Souches', [fsItalic], False, True);                     ExportTable('SOUCHE');
  AfficheMemo('Paramètres de la TOX', [fsItalic], False, True);        ExportTable('STOXPARMS');
  AfficheMemo('Variables de la TOX', [fsItalic], False, True);         ExportTable('STOXVARS');
  AfficheMemo('Sites', [fsItalic], False, True);                       ExportTable('STOXSITES');
  AfficheMemo('Infos complémentaires sites', [fsItalic], False, True); ExportTable('STOXINFOCOMP');
  AfficheMemo('Requêtes de la TOX', [fsItalic], False, True);          ExportTable('STOXQUERYS');
  AfficheMemo('Evenements de la TOX', [fsItalic], False, True);        ExportTable('STOXEVENTS');
  AfficheMemo('Chronos de la TOX', [fsItalic], False, True);           ExportTable('STOXCHRONO');
  AfficheMemo('Groupes d''utilisateurs', [fsItalic], False, True);     ExportTable('USERGRP');
  AfficheMemo('Confidentialité', [fsItalic], False, True);             ExportTable('MENU');
  AfficheMemo('Utilisateurs', [fsItalic], False, True);                ExportTable('UTILISAT');
  AfficheMemo('Paramètres utilisateurs', [fsItalic], False, True);     ExportTable('CPPROFILUSERC');
  AfficheMemo('Paramètres société', [fsItalic], False, True);          ExportParamSoc;
  AfficheMemo('Construction du fichier de démarrage terminée', [fsBold, fsUnderline], False, True);
  AfficheMemo('  => ' + Fichier, [], False, False);

  // Import dans la base PCP
  AfficheMemo('Déconnexion de la base du site central', [], False, False);
  PCPDeconnecteDB; // Déconnexion de la base courante = Base Centrale

  // Connexion à la base PCP
  if not ConnectBase(BasePCP, False, True) then exit;

  // Import du fichier d'initialisation
  YY := TFImportCais.Create(Application);
  try
    // Chargement des données
    AfficheMemo('Chargement des données en cours ...', [fsBold, fsUnderline], False, True);
    YY.NOMDOSSIER.Text := NOMDOSSIER.Text;
    YY.GPK_CAISSE.Text := US_UTILISATEUR_.Text;
    YY.ChargeFichierConf;
    YY.AfficheTOB; // Mise en place des paramsoc dans les zones de la fiche FOImportCais
    YY.SauvePCP;
  finally
    //YY.SortieHalley ;
    YY.Free;
  end;

  // Descente des articles sur la base représentant
  // ==============================================
  AfficheMemo('Déconnexion de la base utilisateur', [], False, False);
  PCPDeconnecteDB; // Déconnexion de la base courante

  // Connexion sur la base du serveur
  FormDBName := SocForm.Text;
  if not ConnectBase(FormDBName, False, False) then exit;

  // Transfert des données via la TOX A PLAT ( Si évènement sélectionné ) !
  Toxcr := True;
  if GEVTINIT.nbSelected > 0 then
  begin
    AfficheMemo('Transfert des données de base en cours ...', [fsBold, fsUnderline], False, False);
    NbOcc := 1000;
    for Ind := GEVTINIT.FixedRows to GEVTINIT.RowCount - 1 do
      if GEVTINIT.IsSelected(Ind) then
      begin
        if not TX_ExecuteEvenement(SITECAISSE.Text, BasePCP, GEVTINIT.CellValues[0, Ind], LogToxAPlat, NbOcc, FALSE) then
        begin
          AfficheMemo('ATTENTION, Pb lors du transfert des données !', [fsBold, fsUnderline], False, False);
          Toxcr := False;
        end;
      end;
  end else
    AfficheMemo('Aucun évènement d''initilisation n''est sélectionné !!', [fsBold, fsUnderline], False, False);

  AfficheMemo('Déconnexion de la base du site central', [], False, False);
  PCPDeconnecteDB; // Déconnexion de la base courante = Base Centrale

  // Connexion à la base PCP
  if not ConnectBase(BasePCP, False, True) then exit;

  if Toxcr then
  begin
    AfficheMemo('INITIALISATION TERMINEE AVEC SUCCES !', [fsBold, fsUnderline], False, True);
    AfficheMemo('Vous pouvez maintenant vous connecter avec votre compte utilisateur.', [], False, True);
  end
  else AfficheMemo('ANOMALIE ! FIN ANORMALE DU TRAITEMENT !', [fsBold, fsUnderline], False, True);
  // bFinOk:=True ; NON, on force la déconnexion de l'utilisateur CEGID/CEGID.
  {$ENDIF}
end;

///////////////////////////////////////////////////////////////////////////////////////
//  NomFichierConf : Retourne le nom du fichier de configuration
///////////////////////////////////////////////////////////////////////////////////////

function TFExportCais.NomFichierConf(AvecChemin: Boolean): string;
var racine, suffixe: string;
begin
  if ctxFO in V_PGI.PGIContexte then
  begin
    if AvecChemin then
      Result := IncludeTrailingBackslash(NOMDOSSIER.Text)
      else
      Result := '';
    racine := 'FO';
    suffixe := Trim(GPK_CAISSE.Text)
  end else
  begin
    if AvecChemin then
      Result := ExtractFileDrive(Application.ExeName) + '\PGI00\Std\'
      else
      Result := '';
    racine := 'PCP';
    suffixe := Trim(US_UTILISATEUR_.Text)
  end;
  Result := Result + racine + 'Cnf' + suffixe + '.txt';
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChangeVariable : change la valeur d'une variable TOX
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.ChangeVariable(TOBL: TOB; NomChamp, NomVar, Valeur: string);
var Ind: Integer;
  StgLst: TStringList;
begin
  StgLst := TStringList.Create;
  StgLst.Text := TOBL.GetValue(NomChamp);
  Ind := StgLst.IndexOfName(NomVar);
  if Ind < 0 then StgLst.Add(NomVar + '=' + Valeur) else StgLst.Strings[Ind] := NomVar + '=' + Valeur;
  TOBL.PutValue(NomChamp, StgLst.Text);
  StgLst.Free;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChangeInfoComp : change la valeur d'une information complémentaire de la TOX
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.ChangeInfoComp(TOBL: TOB; NomChamp, NomInfo, TypeInfo, Valeur: string);
var Ind, Taille: Integer;
  Ok: Boolean;
  StgLst: TStringList;
begin
  Ok := False;
  StgLst := TStringList.Create;
  StgLst.Text := TOBL.GetValue(NomChamp);
  Taille := Length(NomInfo);
  for Ind := 0 to StgLst.Count - 1 do if Copy(StgLst.Strings[Ind], 1, Taille) = NomInfo then
    begin
      StgLst.Strings[Ind] := NomInfo + TypeInfo + Valeur;
      Ok := True;
      Break;
    end;
  if not Ok then StgLst.Add(NomInfo + TypeInfo + Valeur);
  TOBL.PutValue(NomChamp, StgLst.Text);
  StgLst.Free;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ExportTablette : exporte les données d'une tablette dans un fichier
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.ExportTablette(NomTable: string);
var Prefixe, Table, Cle, Fichier, Stg: string;
  TOBData: TOB;
  QQ: TQuery;
begin
  if NomTable = '' then Exit;
  Fichier := NomFichierConf(True);
  if Fichier = '' then Exit;
  Table := '';
  Cle := '';
  Prefixe := '';
  Stg := 'select DO_PREFIXE,DO_TYPE from DECOMBOS where DO_COMBO="' + NomTable + '"';
  QQ := OpenSQL(Stg, True);
  if not QQ.EOF then
  begin
    Prefixe := QQ.FindField('DO_PREFIXE').AsString;
    Cle := QQ.FindField('DO_TYPE').AsString;
  end;
  Ferme(QQ);
  TOBData := TOB.Create('Les Données de ' + NomTable, nil, -1);
  if Prefixe = 'CO' then Table := 'COMMUN' else
    if Prefixe = 'CC' then Table := 'CHOIXCOD' else
    if Prefixe = 'YX' then Table := 'CHOIXEXT';
  if (Table <> '') and (Cle <> '') then TOBData.LoadDetailDB(Table, '"' + Cle + '"', '', nil, False, True);
  TOBData.SaveToFile(Fichier, True, True, False);
  AfficheMemo(TraduireMemoire('Export de') + ' ' + IntToStr(TOBData.Detail.Count) + ' ' + TraduireMemoire('lignes'), [], False, False);
  TOBData.Free;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ExportTable : exporte les données d'une table dans un fichier
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.ExportTable(NomTable: string);
var TOBData, TOBL, TobTmp : TOB;
  QQ: TQuery;                      
  Ind, Cpt: Integer;
  Ok: Boolean;
  Fichier, Stg, sSql, sCle, sBlodName, ListeUtilisat: string;
  NatCourante, NatSuivante : string;

  procedure RequeteEvts;
  var Cpt : integer;
  begin
    if GEVTS.nbSelected > 0 then
    begin
      Ok := False;
      sSql := 'SELECT * FROM STOXEVENTS WHERE SEV_TYPETRT="001" AND SEV_CODEEVENT in (';
      for Cpt := GEVTS.FixedRows to GEVTS.RowCount - 1 do
      begin
        if GEVTS.IsSelected(Cpt) then
        begin
          if Ok then
            sSql := sSql + ','
            else
            Ok := True;
          sSql := sSql + '"' + GEVTS.CellValues[0, Cpt] + '"';
        end;
      end;
      sSql := sSql + ')';
    end else
      sSql := 'SELECT * FROM STOXEVENTS WHERE SEV_TYPETRT="" AND SEV_CODEEVENT=""';
  end;

begin
  if NomTable = '' then Exit;
  Fichier := NomFichierConf(True);
  if Fichier = '' then Exit;
  TOBData := TOB.Create('Les Données de ' + NomTable, nil, -1);
  sCle := '';
  sSql := '';
  sBlodName := '';
  QQ := nil;
  // Constitution des critères de sélection
  if NomTable = 'TIERS' then
  begin
    sSql := 'select * from TIERS where T_TIERS in ("' + SO_GCTIERSDEFAUT.Text + '",'
      + '"' + TOBPCaisse.GetValue('GPK_TIERS') + '","' + GetParamSoc('SO_GCTIERSINV') + '",'
      + '"' + GetParamSoc('SO_GCTIERSMVSTK') + '","' + GetParamSoc('SO_GCTIERSREASSORT') + '")';
  end else
  if NomTable = 'DEPOTS' then
  begin
    if ctxFO in V_PGI.PGIContexte then sCle := '"' + TOBPCaisse.GetValue('GPK_DEPOT') + '"';
  end else
  if NomTable = 'PARCAISSE' then
  begin
    sSql := 'select * from PARCAISSE where GPK_CAISSE="' + TOBPCaisse.GetValue('GPK_CAISSE') + '"'
      + ' or GPK_ETABLISSEMENT="' + TOBPCaisse.GetValue('GPK_ETABLISSEMENT') + '"';
  end else
  if NomTable = 'ETABLISS' then
  begin
    if ctxFO in V_PGI.PGIContexte then sCle := '"' + TOBPCaisse.GetValue('GPK_ETABLISSEMENT') + '"';
  end else
  if NomTable = 'SOUCHE' then
  begin
    sCle := '"GES"';
  end else
  if NomTable = 'STOXSITES' then
  begin
    sSql := 'select * from STOXSITES where SSI_CODESITE in ("' + SITECAISSE.Text + '","' + SITESIEGE.Text + '")';
    sBlodName := 'SSI_VARIABLES';
  end else
  if NomTable = 'STOXINFOCOMP' then
  begin
    sSql := 'select * from STOXINFOCOMP where SIC_CODESITE in ("' + SITECAISSE.Text + '","' + SITESIEGE.Text + '")';
    sBlodName := 'SIC_INFO';
  end else
  if NomTable = 'STOXQUERYS' then
  begin
    if ctxFO in V_PGI.PGIContexte then
      sSql := 'select * from STOXQUERYS where SQE_TYPETRT="001"' // SQE_PREDEFINI<>"CEG"'
      else
      sSql := 'select * from STOXQUERYS where SQE_TYPETRT="001" and SQE_CODEREQUETE like "PCP%"'; // and SQE_PREDEFINI="CEG"' ;
    sBlodName := 'SQE_INFOS';
  end else
  if NomTable = 'STOXEVENTS' then
  begin
    RequeteEvts;
    sBlodName := 'SEV_INFOS';
  end else
  if NomTable = 'MENU' then
  begin
//    sSql := 'select * from MENU where MN_2 in (60,26,167,168,169,170,171,279,280,281)';
    sSql := 'SELECT * FROM MENU WHERE MN_2 IN(169,279,280,281)';
  end else
  if NomTable = 'UTILISAT' then
  begin
    if ctxFO in V_PGI.PGIContexte then
      ListeUtilisat := US_UTILISATEUR.Text
      else
      ListeUtilisat := US_UTILISATEUR_.Text + ';CEG';
    sSql := 'select * from UTILISAT where ' + FabriqueSQLIN(ListeUtilisat, 'US_UTILISATEUR', False, False, False);
  end else
  if NomTable = 'CPPROFILUSERC' then
  begin
    if ctxFO in V_PGI.PGIContexte then
      ListeUtilisat := US_UTILISATEUR.Text
      else
      ListeUtilisat := US_UTILISATEUR_.Text + ';CEG';
    sSql := 'select * from CPPROFILUSERC where ' + FabriqueSQLIN(ListeUtilisat, 'CPU_USER', False, False, False);
  end else
  if NomTable = 'SITEPIECE' then
  begin
    if ctxFO in V_PGI.PGIContexte then sCle := '"' + SITECAISSE.Text + '"'
    else sSql := 'select * from SITEPIECE where GSP_CODESITE="' + SITECAISSE.Text + '"';
  end else
  if NomTable = 'USERGRP' then
  begin
    sSql := 'SELECT * FROM USERGRP WHERE UG_GROUPE IN(SELECT US_GROUPE FROM UTILISAT WHERE US_UTILISATEUR="'+US_UTILISATEUR_.Text+'")';
  end;

  // Lecture des données de la table
  if sSql <> '' then QQ := OpenSQL(sSql, True);
  TOBData.LoadDetailDB(NomTable, sCle, '', QQ, False, True);
  if QQ <> nil then Ferme(QQ);
  // Modification des lignes sélectionnées
  if NomTable = 'DEVISE' then
  begin
    for Ind := 0 to TOBData.Detail.Count - 1 do
    begin
      TOBL := TOBData.Detail[Ind];
      if (TOBL.GetValue('D_DEVISE') = SO_DEVISEPRINC.Value) and
        (TOBL.GetValue('D_MONNAIEIN') = 'X') then Stg := 'X' else Stg := '-';
      TOBL.PutValue('D_FONGIBLE', Stg);
    end;
  end else
  if NomTable = 'DEPOTS' then
  begin
    for Ind := 0 to TOBData.Detail.Count - 1 do
    begin
      TOBL := TOBData.Detail[Ind];
      if ctxFO in V_PGI.PGIContexte then TOBL.PutValue('GDE_SURSITE', 'X')
        // Pour PCP, pas de modification, tous les dépots du SC seront disponibles (en tablette).
        // Les contrôles sur stock sur le poste PCP doivent être désactivés.
    end;
  end else
  if NomTable = 'ETABLISS' then
  begin
    for Ind := 0 to TOBData.Detail.Count - 1 do
    begin
      TOBL := TOBData.Detail[Ind];
      if ctxFO in V_PGI.PGIContexte then TOBL.PutValue('ET_SURSITE', 'X')
        // Pour PCP, pas de modification, tous les établissements du SC seront disponibles (en tablette).
        // Les contrôles sur stock sur le poste PCP doivent être désactivés.
    end;
  end else
  if NomTable = 'SOUCHE' then
  begin
  if ctxFO in V_PGI.PGIContexte then AffecteSouche(TOBData);
  end else
    if NomTable = 'STOXPARMS' then
  begin
    for Ind := 0 to TOBData.Detail.Count - 1 do
    begin
      TOBL := TOBData.Detail[Ind];
      TOBL.PutValue('STP_LASTCHRONOREC', 0);
      TOBL.PutValue('STP_LASTCHRONOSEND', 0);
      TOBL.PutValue('STP_PRIORITYTHREAD', '002'); 
    end;
  end else
  if NomTable = 'STOXEVENTS' then
  begin
    for Ind := 0 to TOBData.Detail.Count - 1 do
    begin
      TOBL := TOBData.Detail[Ind];
      TOBL.PutValue('SEV_ACTIF', 'X');
    end;
  end else
  if NomTable = 'STOXSITES' then
  begin
    ExportSite(TOBData);
  end else
  if NomTable = 'STOXINFOCOMP' then
  begin
    // site de la caisse
    TOBL := TOBData.FindFirst(['SIC_CODESITE'], [SITECAISSE.Text], False);
    if TOBL = nil then
    begin
      TOBL := TOB.Create('STOXINFOCOMP', TOBData, -1);
      TOBL.InitValeurs;
      TOBL.PutValue('SIC_CODESITE', SITECAISSE.Text);
    end;
    ChangeInfoComp(TOBL, 'SIC_INFO', 'DEVISE', ';C;TTDEVISETOUTES;', SO_DEVISEPRINC.Value);
    if SO_TENUEEURO.Checked then ChangeInfoComp(TOBL, 'SIC_INFO', 'BASCULEEURO', ';B;', 'X');
    // site central
    TOBL := TOBData.FindFirst(['SIC_CODESITE'], [SITESIEGE.Text], False);
    if TOBL = nil then
    begin
      TOBL := TOB.Create('STOXINFOCOMP', TOBData, -1);
      TOBL.InitValeurs;
      TOBL.PutValue('SIC_CODESITE', SITESIEGE.Text);
    end;
    ChangeInfoComp(TOBL, 'SIC_INFO', 'DEVISE', ';C;TTDEVISETOUTES;', GetParamSoc('SO_DEVISEPRINC'));
    if GetParamSoc('SO_TENUEEURO') then ChangeInfoComp(TOBL, 'SIC_INFO', 'BASCULEEURO', ';B;', 'X');
  end else
  if NomTable = 'STOXCHRONO' then //Création des chronos pour chaque évènements du site de la base
  begin
    //Charge les évènements sélectionnés
    RequeteEvts;
    TobL := TOB.Create('STOXEVENTS',nil,-1);
    if sSql <> '' then
      QQ := OpenSQL(sSql, True);
    TobL.LoadDetailDB('STOXEVENTS','','',QQ,false);
    Ferme(QQ);
    // Pour chaque évènements, ajout d'un enr dans TOBData avec site courant pour la date courante
    for Ind := 0 to TobL.detail.count-1 do
    begin
      TobTmp := TOB.create('STOXCHRONO',TobData,-1);
      TobTmp.InitValeurs;
      TobTmp.SetString('SXC_CODEEVENT',TobL.detail[Ind].GetString('SEV_CODEEVENT'));
//      TobTmp.SetDateTime('SXC_DATEEXECUTION',V_PGI.DateEntree-1);
      TobTmp.SetDateTime('SXC_DATEEXECUTION',NowH);
      TobTmp.SetString('SXC_NUMCHRONO',IntToStr(Ind));
      TobTmp.SetString('SXC_ETATCHRONO','050');
      TobTmp.SetString('SXC_CODESITE',SITECAISSE.Text);
      TobTmp := TOB.create('STOXCHRONO',TobData,-1);
      TobTmp.InitValeurs;
      TobTmp.SetString('SXC_CODEEVENT',TobL.detail[Ind].GetString('SEV_CODEEVENT'));
      TobTmp.SetDateTime('SXC_DATEEXECUTION',NowH);
      TobTmp.SetString('SXC_NUMCHRONO',IntToStr(Ind));
      TobTmp.SetString('SXC_ETATCHRONO','100');
      TobTmp.SetString('SXC_CODESITE',SITECAISSE.Text);
    end;
    if TobL <> nil then FreeAndNil(TobL);
  end else
  if NomTable = 'UTILISAT' then
  begin
    for Ind := 0 to TOBData.Detail.Count - 1 do
    begin
      TOBL := TOBData.Detail[Ind];
      TOBL.PutValue('US_PRESENT', '-');
    end;
  end;
{$IFDEF NOMADE} //JT 10/12/2003 - Modif paramètrage PARPIECE de CC que pour PGE
  if NomTable = 'PARPIECE' then
  begin
    for Ind := TOBData.detail.count - 1 downto 0 do
    begin
      TOBL := TOBData.detail[ind];
      if (TOBL.GetValue('GPP_TRSFVENTE') = '-') and (TOBL.GetValue('GPP_TRSFACHAT') = '-' ) then
        TOBL.free
      else if ((not VH_GC.PCPUsVte) and (TOBL.GetValue('GPP_TRSFVENTE') = 'X')) then
        TOBL.Free
      else if ((not VH_GC.PCPUsVte) and (TOBL.GetValue('GPP_TRSFACHAT') = 'X')) then
        TOBL.Free
      else
      begin
        if (not (ctxMode in V_PGI.PGIContexte)) and (TOBL.GetValue('GPP_NATUREPIECEG') = 'CC') then // Si Cde client
        begin                                                                                       // test si PRO+CLI
          if pos('PRO',TOBL.GetValue('GPP_NATURETIERS')) = 0 then
            TOBL.PutValue('GPP_NATURETIERS',TOBL.GetValue('GPP_NATURETIERS')+'PRO;');
          if TOBL.GetValue('GPP_PROCLI') = '-' then
            TOBL.PutValue('GPP_PROCLI','X');
        end;
        if GP_VERROUILLEPCES_.Checked then // Vérouiller les pièces après trsf Tox
          TOBL.PutValue('GPP_ACTIONFINI','TOX');
        if GP_IMPIMMEDIATE_.Checked then // Impression à la validation des pièces
        begin
          TOBL.PutValue('GPP_IMPIMMEDIATE','X');
          if GP_APERCUAVIMP_.Checked then
            TOBL.PutValue('GPP_APERCUAVIMP','X')
            else
            TOBL.PutValue('GPP_APERCUAVIMP','-');
        end else
        begin
          TOBL.PutValue('GPP_IMPIMMEDIATE','-');
          TOBL.PutValue('GPP_APERCUAVIMP','-');
        end;
        TOBL.PutValue('GPP_TYPEECRCPTA','RIE'); // Pas de compta
      end;
    end;
    // Maj des pièces suivantes
    TobTmp := TOB.Create('Les natures',nil,-1);
    TobTmp.Dupliquer(TobData,true,true);
    for Ind := 0 to TobData.detail.count - 1  do
    begin
      Stg := TobData.Detail[Ind].GetString('GPP_NATURESUIVANTE');
      TobData.Detail[Ind].SetString('GPP_NATURESUIVANTE','');
      // Recherche si les natures suivantes existent dans le parpièce
      while Stg <> '' do
      begin
        NatSuivante := ReadTokenSt(Stg);
        for Cpt := 0 to TobTmp.detail.count-1 do
        begin
          if pos(NatSuivante,TobTmp.Detail[Cpt].GetString('GPP_NATUREPIECEG')) > 0 then
          begin
            NatCourante := TobData.Detail[Ind].GetString('GPP_NATURESUIVANTE')+NatSuivante+';';
            TobData.Detail[Ind].SetString('GPP_NATURESUIVANTE',NatCourante);
            break;
          end;
        end;
      end;
    end;
    if TobTmp <> nil then FreeAndNil(TobTmp);
  end;
{$ENDIF NOMADE}

  // Export des données
  TOBData.SaveToFile(Fichier, True, True, False, sBlodName);
  AfficheMemo(TraduireMemoire('Export de') + ' ' + IntToStr(TOBData.Detail.Count) + ' ' + TraduireMemoire('lignes'), [], False, False);
  TOBData.Free;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ExportSite : exporte les données de la table des sites de la TOX
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.ExportSite(TOBData: TOB);
{$IFDEF TOX_V1_PLUS_UTILISE}
var TOBC, TOBL: TOB;
  {$ELSE}
var TOBL: TOB;
  {$ENDIF}
  Ind: Integer;
  bSiteLocal: boolean;
  sVariable, sValue: string;
begin
  if ctxFO in V_PGI.PGIContexte then
  begin
    sVariable := '$ETAB';
    sValue := TOBPCaisse.GetValue('GPK_ETABLISSEMENT');
  end else
  begin
    sVariable := '$UTIL';
    sValue := TOBRepr.GetValue('US_UTILISATEUR');
  end;
  // recherche du site de la caisse
  {$IFDEF TOX_V1}
  TOBC := TOBData.FindFirst(['SSI_CODESITE'], [SITECAISSE.Text], False);
  {$ENDIF}
  for Ind := 0 to TOBData.Detail.Count - 1 do
  begin
    TOBL := TOBData.Detail[Ind];
    TOBL.PutValue('SSI_SITEENABLED', 'X');
    bSiteLocal := TOBL.GetValue('SSI_CODESITE') = SITECAISSE.Text;
    if (ctxFO in V_PGI.PGIContexte) or (bSiteLocal)
      then ChangeVariable(TOBL, 'SSI_VARIABLES', sVariable, sValue);
    {$IFDEF TOX_V1_PLUS_UTILISE}
    if TOBL.GetValue('SSI_CODESITE') = SITECAISSE.Text then
    begin
      // le site de la caisse devient local
      TOBC.PutValue('SSI_STYPESITE', '002');
    end else
    begin
      // le site central devient distant
      TOBL.PutValue('SSI_STYPESITE', '001');
    end;
    {$ELSE}
    TOBL.PutValue('SSI_CURRENTSITE', CheckToString(bSiteLocal));
    {$ENDIF}
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ExportParamSoc : exporte les paramètres société dans un fichier
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.ExportParamSoc;

  procedure MajParam(TOBData: TOB; NomParam: string; ValeurParam: Variant);
  var TOBL: TOB;
  begin
    TOBL := TOBData.FindFirst(['SOC_NOM'], [NomParam], False);
    if TOBL = nil then
    begin
      TOBL := TOB.Create('PARAMSOC', TOBData, -1);
      TOBL.InitValeurs;
      TOBL.PutValue('SOC_NOM', NomParam);
    end;
    TOBL.PutValue('SOC_DATA', ValeurParam);
  end;

var TOBData: TOB;
  QQ: TQuery;
  Cpt, Taille: Integer;
  Fichier, sSql, sEtab, Stg, sMin, sMax: string;
begin
  Fichier := NomFichierConf(True);
  if Fichier = '' then Exit;
  sEtab := TOBEtab.GetValue('ET_ETABLISSEMENT');
  TOBData := TOB.Create('Les Données de PARAMSOC', nil, -1);
  // Sélection des paramètres de la comptabilité et de la gestion commerciale
  sSql := 'select SOC_NOM,SOC_DATA from PARAMSOC where (SOC_TREE like "001;001;%" or SOC_TREE like "001;003;%")'
    + ' and SOC_TREE not like "%;000;"';
  QQ := OpenSQL(sSql, True);
  if not QQ.Eof then TOBData.LoadDetailDB('PARAMSOC', '', '', QQ, False, True);
  Ferme(QQ);

  // Modification des lignes sélectionnées
  if ctxFO in V_PGI.PGIContexte then
  begin
    MajParam(TOBData, 'SO_SOCIETE', GPK_CAISSE.Text);
    MajParam(TOBData, 'SO_LIBELLE', SO_LIBELLE.Text);
    MajParam(TOBData, 'SO_NATUREJURIDIQUE', '...');
    MajParam(TOBData, 'SO_ADRESSE1', TOBEtab.GetValue('ET_ADRESSE1'));
    MajParam(TOBData, 'SO_ADRESSE2', TOBEtab.GetValue('ET_ADRESSE2'));
    MajParam(TOBData, 'SO_ADRESSE3', TOBEtab.GetValue('ET_ADRESSE3'));
    MajParam(TOBData, 'SO_CODEPOSTAL', TOBEtab.GetValue('ET_CODEPOSTAL'));
    MajParam(TOBData, 'SO_VILLE', TOBEtab.GetValue('ET_VILLE'));
    MajParam(TOBData, 'SO_PAYS', TOBEtab.GetValue('ET_PAYS'));
    MajParam(TOBData, 'SO_TELEPHONE', TOBEtab.GetValue('ET_TELEPHONE'));
    MajParam(TOBData, 'SO_FAX', TOBEtab.GetValue('ET_FAX'));
    MajParam(TOBData, 'SO_MAIL', TOBEtab.GetValue('ET_EMAIL'));
    MajParam(TOBData, 'SO_TELEX', TOBEtab.GetValue('ET_TELEX'));
    MajParam(TOBData, 'SO_RVA', '');
    MajParam(TOBData, 'SO_CONTACT', '');
    MajParam(TOBData, 'SO_NIF', '');
    MajParam(TOBData, 'SO_APE', '');
    MajParam(TOBData, 'SO_SIRET', '');
    MajParam(TOBData, 'SO_RC', '');
    MajParam(TOBData, 'SO_TXTJURIDIQUE', '');
    MajParam(TOBData, 'SO_CAPITAL', '0');
    MajParam(TOBData, 'SO_ETABLISDEFAUT', sEtab);
    MajParam(TOBData, 'SO_REGIMEDEFAUT', SO_REGIMEDEFAUT.Value);
    MajParam(TOBData, 'SO_GCMODEREGLEDEFAUT', SO_GCMODEREGLEDEFAUT.Value);
    MajParam(TOBData, 'SO_GCTIERSDEFAUT', SO_GCTIERSDEFAUT.Text);
    MajParam(TOBData, 'SO_GCPHOTOFICHE', SO_GCPHOTOFICHE.Value);
    MajParam(TOBData, 'SO_GCTIERSPAYS', TOBEtab.GetValue('ET_PAYS'));
    MajParam(TOBData, 'SO_GCDEPOTDEFAUT', TOBPCaisse.GetValue('GPK_DEPOT'));
    MajParam(TOBData, 'SO_GCPERBASETARIF', SO_GCPERBASETARIF.Value);
    if Trim(TOBPCaisse.GetValue('GPK_APPELPRIXTIC')) = 'HT ' then MajParam(TOBData, 'SO_GCDEFFACTUREHT', 'X')
    else MajParam(TOBData, 'SO_GCDEFFACTUREHT', '-');
    MajParam(TOBData, 'SO_GCMONTRENUMERO', '-');
    MajParam(TOBData, 'SO_GCDESACTIVECOMPTA', 'X');
    MajParam(TOBData, 'SO_GCFOCAISREFTOX', GPK_CAISSE.Text);
    if not Superviseur then
    begin
    // Attribution du chrono client
    Taille := GetParamsoc('SO_GCLGNUMTIERS') - Length(sEtab);
    if Taille > 0 then
    begin
      sMin := sEtab + StringOfChar('0', Taille);
      sMax := sEtab + StringOfChar('9', Taille);
      Cpt := StrToInt(sMin);
      sSql := 'select max(T_TIERS) from TIERS where T_TIERS >= "' + sMin + '" and T_TIERS <= "' + sMax + '"';
      QQ := OpenSQL(sSql, True);
      if not QQ.Eof then
      begin
        Stg := QQ.Fields[0].AsString;
        if IsNumeric(Stg) then Cpt := StrToInt(Stg) + 1;
      end;
      Ferme(QQ);
      MajParam(TOBData, 'SO_GCCOMPTEURTIERS', Cpt);
      end;
    end;
    MajParam(TOBData, 'SO_GCNUMTIERSAUTO', 'X');
    // Traitement de la devise de tenue
    MajParam(TOBData, 'SO_DECQTE', 0);
    MajParam(TOBData, 'SO_DEVISEPRINC', SO_DEVISEPRINC.Value);
    sSql := 'select D_DECIMALE,D_PARITEEURO from DEVISE where D_DEVISE="' + SO_DEVISEPRINC.Value + '"';
    QQ := OpenSQL(sSql, True);
    if not QQ.Eof then
    begin
      MajParam(TOBData, 'SO_DECVALEUR', QQ.FindField('D_DECIMALE').AsInteger);
      MajParam(TOBData, 'SO_TAUXEURO', QQ.FindField('D_PARITEEURO').AsFloat);
      MajParam(TOBData, 'SO_DECPRIX', QQ.FindField('D_DECIMALE').AsInteger);
    end;
    Ferme(QQ);
    /////if SO_TENUEEURO.Checked then MajParam(TOBData, 'SO_TENUEEURO', 'X') ;

  end else // PCP
  begin
    //Coordonnées
    MajParam(TOBData, 'SO_LIBELLE', SO_LIBELLE_.Text);
    MajParam(TOBData, 'SO_ADRESSE1', SO_ADRESSE1_.Text);
    MajParam(TOBData, 'SO_ADRESSE2', SO_ADRESSE2_.Text);
    MajParam(TOBData, 'SO_ADRESSE3', SO_ADRESSE3_.Text);
    MajParam(TOBData, 'SO_CODEPOSTAL', SO_CODEPOSTAL_.Text);
    MajParam(TOBData, 'SO_DIVTERRIT', SO_DIVTERRIT_.Text);
    MajParam(TOBData, 'SO_VILLE', SO_VILLE_.Text);
    MajParam(TOBData, 'SO_PAYS', SO_PAYS_.Text);
    MajParam(TOBData, 'SO_TELEPHONE', SO_TELEPHONE_.Text);
    MajParam(TOBData, 'SO_FAX', SO_FAX_.Text);
    MajParam(TOBData, 'SO_MAIL', SO_MAIL_.Text);
    //Codification article
    MajParam(TOBData, 'SO_GCNUMARTAUTO', 'X');
    MajParam(TOBData, 'SO_GCLGNUMART', SO_GCLGNUMART_.Value);
    MajParam(TOBData, 'SO_GCPREFIXEART', SO_GCPREFIXEART_.Text);
    MajParam(TOBData, 'SO_GCCOMPTEURART', SO_GCCOMPTEURART_.Text);
    //Codification clients et fournisseurs
    MajParam(TOBData, 'SO_GCNUMTIERSAUTO', 'X');
    MajParam(TOBData, 'SO_GCNUMFOURAUTO', 'X');
    MajParam(TOBData, 'SO_GCPREFIXETIERS', SO_GCPREFIXEART_.Text);
    MajParam(TOBData, 'SO_GCCOMPTEURSTIERS', '0');
    MajParam(TOBData, 'SO_GCLGNUMTIERS', SO_GCLGNUMTIERS_.Value);
    //Auxiliaire clients et fournisseurs
    MajParam(TOBData, 'SO_GCPREFIXEAUXI', SO_GCPREFIXEAUXI_.Text);
    MajParam(TOBData, 'SO_GCPREFIXEAUXIFOU', SO_GCPREFIXEAUXIFOU_.Text);
//    if VH_GC.PCPVenteSeria then
    if VH_GC.PCPUsVte then
      MajParam(TOBData, 'SO_GCCREATPRO', 'X')
      else
      MajParam(TOBData, 'SO_GCCREATPRO', '-');
    //Paramètres comptable
    MajParam(TOBData, 'SO_GCDESACTIVECOMPTA ', 'X');
  end;

  // Export des données
  TOBData.SaveToFile(Fichier, True, True, False);
  AfficheMemo(TraduireMemoire('Export de') + ' ' + IntToStr(TOBData.Detail.Count) + ' ' + TraduireMemoire('lignes'), [], False, False);
  TOBData.Free;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AffecteSouche : Attribue les numéros des souches
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.AffecteSouche(TOBData: TOB);
var TOBNum, TOBL, TOBN: TOB;
  sSql, sEtab: string;
  Ind, Cpt: Integer;
  QQ: TQuery;
begin
  TOBNum := TOB.Create('Numeros', nil, -1);
  // Sélection des derniers numéros attribués dans la plage par type de souche
  sEtab := TOBEtab.GetValue('ET_ETABLISSEMENT');
  sSql := 'select GP_SOUCHE, max(GP_NUMERO) as GP_NUMERO from PIECE'
    + ' where GP_NUMERO >= ' + sEtab + '000000 and GP_NUMERO <= ' + sEtab + '999999'
    + ' group by GP_SOUCHE';
  QQ := OpenSQL(sSql, True);
  if not QQ.Eof then TOBNum.LoadDetailDB('PIECE', '', '', QQ, False, True);
  Ferme(QQ);
  // Mise à jour des souches
  for Ind := 0 to TOBData.Detail.Count - 1 do
  begin
    TOBL := TOBData.Detail[Ind];
    TOBN := TOBNum.FindFirst(['GP_SOUCHE'], [TOBL.GetValue('SH_SOUCHE')], False);
    if TOBN = nil then Cpt := StrToInt(sEtab + '000000')
    else Cpt := TOBN.GetValue('GP_NUMERO') + 1;
    TOBL.PutValue('SH_NUMDEPART', Cpt);
  end;
  TOBNum.Free;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FabriqueDemarrage : Constitution de l'événement de démarrage à partir de l'événement choisi
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.FabriqueDemarrage;
var NewEvt, Stg: string;
  TOBE: TOB;
  StgLst: TStringList;
begin
  // lecture de l'événement de démarrage
  AfficheMemo('Chargement de l''événement de démarrage', [], False, True);
  TOBE := TOB.Create('STOXEVENTS', nil, -1);
  TOBE.SelectDB('"001";"' + SEV_CODEEVENT.Text + '"', nil);
  // Recopie de l'événement choisi avec le nom de l'événement choisi auquel on ajoute un '-' puis le code du site
  NewEvt := SEV_CODEEVENT.Text;
  if Length(NewEvt) > 13 then Delete(NewEvt, 14, 4);
  NewEvt := NewEvt + '-' + GPK_CAISSE.Text;
  TOBE.PutValue('SEV_CODEEVENT', NewEvt);
  TOBE.PutValue('SEV_ACTIF', 'X');
  TOBE.PutValue('SEV_TYPE', 'EM');
  TOBE.PutValue('SEV_TYPECONDITION', '003');
  // Remplacement du site (QUERYS=StgLst[0] - SITES=StgLst[1] - GROUPES=StgLst[2])
  StgLst := TStringList.Create;
  StgLst.Text := TOBE.GetValue('SEV_INFOS');
  StgLst[1] := SITECAISSE.Text + ';';
  TOBE.PutValue('SEV_INFOS', StgLst.Text);
  StgLst.Free;
  // Création du nouvel événement
  TOBE.SetAllModifie(True);
  TOBE.InsertDB(nil);
  TOBE.Free;
  Stg := TraduireMemoire('Copie de') + ' ' + SEV_CODEEVENT.Text + ' ' + TraduireMemoire('en') + ' ' + NewEvt;
  AfficheMemo(Stg, [], False, False);
end;

procedure TFExportCais.AffEtape;
begin
  lEtapeManu.Caption := Msg.Mess[0] + ' ' + IntToStr(CurrentPage) + '/' + IntToStr(PagesCount);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  bPrecedentClick : bouton PRECEDENT
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.bPrecedentClick(Sender: TObject);
begin
  // Gestion des onglets non visibles !
  bImprimer.Visible := False;
  bParamPce.Visible := False;
  while (P.ActivePage.PageIndex > 1) and not P.Pages[P.ActivePage.PageIndex - 1].TabVisible
    do P.ActivePage.PageIndex := P.ActivePage.PageIndex - 1;
  inherited;
  // Maj label étape + Inactivation des boutons PRECEDENT et FIN
  dec(CurrentPage);
  AffEtape;
  bPrecedent.Enabled := (P.ActivePage.Name <> 'PAGE1') and (P.ActivePage.Name <> 'PAGE21');
  bFin.Enabled := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  bSuivantClick : bouton SUIVANT
///////////////////////////////////////////////////////////////////////////////////////
Procedure TFExportCais.GereErreur(NumErreur : integer; Chps : THEdit; Txt : string);
begin
  if NumErreur <> 0 then
  begin
    if NumErreur > 0 then
      Msg.Execute(NumErreur, Caption, '');
    LstRecap.Clear;
    if Chps.CanFocus then Chps.SetFocus;
  end else
    LStRecap.Items.Add(Txt);
end;

procedure TFExportCais.bSuivantClick(Sender: TObject);
var QQ: TQuery;
  TOBEvt, TOBE: TOB;
  Ind(*, Cpt*) : Integer;
  StgLst: TStringList;
  sSql, EvtInitDefaut : string;
begin
  // Verification des paramètres saisis
  if P.ActivePage.Name = 'PAGE1' then
  begin
    if (GPK_CAISSE.Text = '') or (not LookUpValueExist(GPK_CAISSE)) then
    begin
      Msg.Execute(2, Caption, '');
      if GPK_CAISSE.CanFocus then GPK_CAISSE.SetFocus;
      Exit;
    end;
    if SO_LIBELLE.Text = '' then
    begin
      Msg.Execute(5, Caption, '');
      if SO_LIBELLE.CanFocus then SO_LIBELLE.SetFocus;
      Exit;
    end;
    if TOBPCaisse.GetValue('GPK_CAISSE') <> GPK_CAISSE.Text then
    begin
      if TOBPCaisse <> nil then TOBPCaisse.Free;
      TOBPCaisse := TOB.Create('PARCAISSE', nil, -1);
      TOBPCaisse.SelectDB('"' + GPK_CAISSE.Text + '"', nil);
    end;
    if TOBEtab.GetValue('ET_ETABLISSEMENT') <> TOBPCaisse.GetValue('GPK_ETABLISSEMENT') then
    begin
      if TOBEtab <> nil then TOBEtab.Free;
      TOBEtab := TOB.Create('ETABLISS', nil, -1);
      TOBEtab.SelectDB('"' + TOBPCaisse.GetValue('GPK_ETABLISSEMENT') + '"', nil);
    end;
    // Recherche du site de la caisse
    ChargeSiteCaisse;
  end;

  //Choix de évènements d'initialisation Mode
  if P.ActivePage.Name = 'PAGE3' then
  begin
    if (SITECAISSE.Text = '') or (not LookUpValueExist(SITECAISSE)) then
    begin
      Msg.Execute(3, Caption, '');
      if SITECAISSE.CanFocus then SITECAISSE.SetFocus;
      Exit;
    end;
    if (SITESIEGE.Text = '') or (not LookUpValueExist(SITESIEGE)) or (SITESIEGE.Text = SITECAISSE.Text) then
    begin
      Msg.Execute(3, Caption, '');
      if SITESIEGE.CanFocus then SITESIEGE.SetFocus;
      Exit;
    end else
    begin
      // Chargement de la grille de sélection des événements TOX
      sSql := 'SELECT * FROM STOXEVENTS WHERE SEV_TYPETRT="001"';
      QQ := OpenSQL(sSql, True);
      TOBEvt := TOB.Create('Les STOXEVENTS', nil, -1);
      TOBEvt.LoadDetailDB('STOXEVENTS', '', '', QQ, False);
      Ferme(QQ);
      TOBEvt.PutGridDetail(GEVTS, False, False, 'SEV_CODEEVENT;SEV_LIBELLE', True);
      TOBEvt.PutGridDetail(GEVTINIT, False, False, 'SEV_CODEEVENT;SEV_LIBELLE', True);
      EvtInitDefaut := 'BOUTIQUE';
      // Chargement de la grille de sélection de l'évènement d'initialisation
      StgLst := TStringList.Create;
      for Ind := 0 to TOBEvt.Detail.Count - 1 do
      begin
        TOBE := TOBEvt.Detail[Ind];
        StgLst.Text := TOBE.GetValue('SEV_INFOS');
        if EstDansLaListe(SITESIEGE.Text, StgLst[1]) then
          GEVTS.FlipSelection(GEVTS.FixedRows + Ind);
        // Sélection évt d'initialisation
          if TOBE.GetValue('SEV_CODEEVENT') = EvtInitDefaut then GEVTINIT.FlipSelection(GEVTINIT.FixedRows + Ind);
      end;
      StgLst.Free;
      TOBEvt.Free;
    end;
  end;

  //Choix des évènements de remontée
  if P.ActivePage.Name = 'PAGE8' then
  begin
    if GEVTS.nbSelected = 0 then
    begin
      Msg.Execute(21, Caption, '');
      Exit;
    end;
    if not (ctxFO in V_PGI.PGIContexte) then
      CreerRapport;
  end;

  // Où mettre le fichier de démarrage
  if P.ActivePage.Name = 'PAGE4' then
  begin
    NOMDOSSIER.Text := Trim(NOMDOSSIER.Text);
    if not DirectoryExists(NOMDOSSIER.Text) then
    begin
      Msg.Execute(4, Caption, '');
      if NOMDOSSIER.CanFocus then NOMDOSSIER.SetFocus;
      Exit;
    end;
  end;

  if P.ActivePage.Name = 'PAGE7' then
  begin
    if (LANCETOX.Checked) and ((SEV_CODEEVENT.Text = '') or (not LookUpValueExist(SEV_CODEEVENT))) then
    begin
      Msg.Execute(7, Caption, '');
      if SEV_CODEEVENT.CanFocus then SEV_CODEEVENT.SetFocus;
      Exit;
    end;
  end;

  // Choix de la base Site Central et Utilisateur
  if P.ActivePage.Name = 'PAGE21' then
  begin
    if (US_UTILISATEUR_.text = '') or (not LookUpValueExist(US_UTILISATEUR_)) then
    begin
      GereErreur(8, THEdit(US_UTILISATEUR_), '');
      exit;
    end;
  end;

  //Coordonées
  if P.ActivePage.Name = 'PAGE211' then
  begin
    SO_GCPREFIXEAUXI_.Visible     := VH_GC.PCPVenteSeria;
    LSO_GCPREFIXEAUXI_.Visible    := VH_GC.PCPVenteSeria;
    SO_GCPREFIXEAUXIFOU_.Visible  := VH_GC.PCPAchatSeria;
    LSO_GCPREFIXEAUXIFOU_.Visible := VH_GC.PCPAchatSeria;
    if ctxMode in V_PGI.PGIContexte then
    begin
      GP_VERROUILLEPCES_.Checked := False;
      SO_GCLGNUMTIERS_.Visible := True;
      LSO_GCLGNUMTIERS_.Visible := True;
    end else
    begin
      GP_VERROUILLEPCES_.Checked := True;
      SO_GCLGNUMTIERS_.Visible := False;
      LSO_GCLGNUMTIERS_.Visible := False;
    end;
  end;

  //Codification article et tiers
  if P.ActivePage.Name = 'PAGE212' then
  begin
    if (ctxMode in V_PGI.PGIContexte) and (SO_GCLGNUMTIERS_.Value = 0) then
    begin
      GereErreur(25, THEdit(SO_GCLGNUMTIERS_), '');
      Exit;
    end;
    //Test si auxiliaires saisis
{
    if (VH_GC.PCPVente) and (SO_GCPREFIXEAUXI_.Text = '') then
    begin
      GereErreur(22, THEdit(SO_GCPREFIXEAUXI_), '');
      Exit;
    end;
    if (VH_GC.PCPAchat) and (SO_GCPREFIXEAUXIFOU_.Text = '') then
    begin
      GereErreur(23, THEdit(SO_GCPREFIXEAUXIFOU_), '');
      Exit;
    end;
}
    //Si vente et achat et auxiliaire saisi, test s'ils sont différents
    if (not Superviseur) then
    begin
    if (VH_GC.PCPVenteSeria) and (VH_GC.PCPAchatSeria) and
       (SO_GCPREFIXEAUXI_.Text <> '') and (SO_GCPREFIXEAUXIFOU_.Text <> '') and
       (SO_GCPREFIXEAUXI_.Text = SO_GCPREFIXEAUXIFOU_.Text) then
    begin
      GereErreur(24, THEdit(SO_GCPREFIXEAUXI_), '');
      Exit;
      end;
    end;
  end;

  // Sélection de l'évènement d'initialisation du poste PCP
  if P.ActivePage.Name = 'PAGE22' then
  begin
    if GEVTINIT.nbSelected = 0 then
    begin
      Msg.Execute(20, Caption, '');
      Exit;
    end else
    // Test si les evts obligatoire (comparés au param user) sont sélectionnés
    begin
(*      if EvtOblInit <> '' then
      begin
        for Cpt := GEVTINIT.RowCount -1 downto 1 do
          if (pos(GEVTINIT.Cells[0,Cpt],EvtOblInit) > 0) and (not GEVTINIT.IsSelected(Cpt))then
          begin
            Msg := TraduireMemoire('La nature de pièce "')+Libelle+
                   TraduireMemoire('" ne dispose pas de souche pour ')+US_LIBELLE.Text+'.';
            HShowMessage('-1;'+Caption+';'+Msg+';E;O;O;O','','');
          end;
      end;
*)
    end;
  end;

  // Gestion des onglets non visibles !
  while (P.ActivePage.PageIndex < P.PageCount - 1) and not P.Pages[P.ActivePage.PageIndex + 1].TabVisible do
    P.ActivePage.PageIndex := P.ActivePage.PageIndex + 1;
  inherited;
  if P.ActivePage.Name = 'PAGE5' then
  begin
    if ctxFO in V_PGI.PGIContexte then
    begin
      TITRE51.Caption := GPK_CAISSE.Text + ' - ' + LIB_CAISSE.Caption
    end else
    begin
      bImprimer.Visible := True;
      TITRE51.Caption := US_UTILISATEUR_.Text + ' - ' + US_LIBELLE.Text;
    end;
  end else
    bImprimer.Visible := False;
    bParamPce.Visible := False;
  // Maj label étape + Activation du bouton FIN
  inc(CurrentPage);
  AffEtape;
  if (CurrentPage = PagesCount) then bFin.Enabled := True;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  bFinClick : bouton FIN
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.bFinClick(Sender: TObject);
{$IFDEF PB_COMPILE}
var  IdApplication: string;
  {$ENDIF}
begin
  inherited;
  if not (ctxFO in V_PGI.PGIContexte) then
  begin
    LstRecap.Visible := False;
    LstRecap.Clear;
    CPTRENDU.Visible := True;
  end;
  ActiveBtn(False);
  AfficheMemo('Construction du fichier de démarrage', [fsBold, fsUnderline], True, True);
  if ctxFO in V_PGI.PGIContexte then
    SauveCaisse
    else
    SauvePCP;
  if LANCETOX.Checked then
  begin
    // Constitution de l'événement de démarrage
    AfficheMemo('Constitution de l''événement de démarrage', [fsBold], True, True);
    FabriqueDemarrage;
    // Démarrage des échanges FO - BO
    {$IFDEF PB_COMPILE}
    if AglToxServeurActifNet then //if not AglStatusTox(WhenStart) then
    begin
      AfficheMemo('Démarrage des échanges FO-BO', [fsBold], False, True);
      IdApplication := FOGetToxIdApp;
      AglToxConf(aceStart, IdApplication, GescomModeNotConfirmeTox, GescomModeTraitementTox, AglToxFormError);
    end;
    {$ENDIF}
  end;
{$IFDEF NOMADE}
  bParamPce.Visible := (GP_PARAMPCES_.Checked);
  if GP_PARAMPCES_.Checked then //Paramètrage des pièces demandé
  begin
    if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
      AglLanceFiche ('GC','GCPARPIECE','','','') // LanceparPiece
      else
      AglLanceFiche ('GC','GCPARPIECEUSR','','','') ; // LanceparPiece  utilisateur
  end;
{$ENDIF NOMADE}
  bAnnuler.Caption := TraduireMemoire('&Fermer');
  bAnnuler.Enabled := True;
  bImprimer.Enabled := True;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormCreate :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.FormCreate(Sender: TObject);
var IndiceOnglet: Integer;
  bVisible: boolean;
  StOnglet, StTmp: string;
begin
	Superviseur := false;
  // Affichage ou non des onglets suivant le type d'assistant : FO / PCP
  PagesCount := 0;
  CurrentPage := 1;
  if not (ctxFO in V_PGI.PGIContexte) then
  begin
    StOnglet := OngletUtile[ouPCP];
    P.Pages[1].PageIndex := 0;
    bFinOk := False;
    BC_LesSites := nil; // Sites de la base centrale non chargés.
  end else
    StOnglet := OngletUtile[ouFO];
  for IndiceOnglet := 0 to P.PageCount - 1 do
  begin
    StTmp := ';' + P.Pages[IndiceOnglet].Name + ';';
    bVisible := (pos(StTmp, StOnglet) > 0);
    P.Pages[IndiceOnglet].TabVisible := bVisible;
    P.Pages[IndiceOnglet].TabStop := bVisible;
    if bVisible then inc(PagesCount);
  end;
  AffEtape; // Initialisation
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormShow :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.FormShow(Sender: TObject);
var Ind: Integer;
begin
  inherited;
  bFin.Enabled := False;
  CPTRENDU.Clear;
  LIB_CAISSE.Caption := '';

  if ctxFO in V_PGI.PGIContexte then
  begin
    TOBPCaisse := TOB.Create('PARCAISSE', nil, -1);
    TOBEtab := TOB.Create('ETABLISS', nil, -1);
    if CodeCaisse <> '' then
    begin
      TOBPCaisse.SelectDB('"' + CodeCaisse + '"', nil);
      GPK_CAISSE.Text := TOBPCaisse.GetValue('GPK_CAISSE');
      LIB_CAISSE.Caption := TOBPCaisse.GetValue('GPK_LIBELLE');
      SO_LIBELLE.Text := LIB_CAISSE.Caption;
      TOBEtab.SelectDB('"' + TOBPCaisse.GetValue('GPK_ETABLISSEMENT') + '"', nil);
    end;
    // Valeurs par défaut
    SO_DEVISEPRINC.Value := GetParamSoc('SO_DEVISEPRINC');
    SO_REGIMEDEFAUT.Value := GetParamSoc('SO_REGIMEDEFAUT');
    SO_GCMODEREGLEDEFAUT.Value := GetParamSoc('SO_GCMODEREGLEDEFAUT');
    SO_GCTIERSDEFAUT.Text := GetParamSoc('SO_GCTIERSDEFAUT');
    SO_GCPERBASETARIF.Value := GetParamSoc('SO_GCPERBASETARIF');
    SO_GCPHOTOFICHE.Value := GetParamSoc('SO_GCPHOTOFICHE');
    SEV_CODEEVENT.Text := 'DEMARRAGE';
    ChargeSiteCentral;
    LstRecap.Visible := False;
    CPTRENDU.Visible := True;
  end else
  begin // Contexte PCP -> assistant
    TobRepresentant := TOB.Create('COMMERCIAL', nil, -1); // JTR Gestion PCP par user
    TobNatPces := TOB.Create('LES SOUCHES', nil, -1);
    BasePCP := V_PGI.CurrentAlias; // Sauvegarde du nom de la base PCP
    SocForm := THValComboBox(DBCENTRAL);
    ChargeDossier(SocForm.Items, True);
    // Suppression de la base courante !!
    Ind := SocForm.Items.IndexOf(V_PGI.CurrentAlias);
    if Ind > -1 then SocForm.Items.Delete(Ind);
    US_LIBELLE.Text := '';
    // Chargement de l'utilisateur
    if CodeUsr <> '' then
    begin
      TOBRepr := TOB.Create('UTILISAT', nil, -1);
      TOBRepr.SelectDB('"' + CodeUsr + '"', nil, True);
      US_UTILISATEUR_.Text := TOBRepr.GetValue('US_UTILISATEUR');
      US_lIBELLE.Text := TOBRepr.GetValue('US_LIBELLE');
    end;
    LstRecap.Visible := True;
    CPTRENDU.Visible := False;
{$IFDEF BTP}
		SO_GCPREFIXEART_.enabled := true;
		SO_GCPREFIXETIERS_.enabled := true;
		SO_GCCOMPTEURART_.enabled := true;
		SO_GCCOMPTEURTIERS_.enabled := true;
{$ENDIF}
  end;
end;

procedure TFExportCais.LoadData;
begin
  // Chargement de la caisse
  TOBPCaisse := TOB.Create('PARCAISSE', nil, -1);
  TOBEtab := TOB.Create('ETABLISS', nil, -1);
  if CodeCaisse <> '' then
  begin
    TOBPCaisse.SelectDB('"' + CodeCaisse + '"', nil);
    GPK_CAISSE.Text := TOBPCaisse.GetValue('GPK_CAISSE');
    LIB_CAISSE.Caption := TOBPCaisse.GetValue('GPK_LIBELLE');
    SO_LIBELLE.Text := LIB_CAISSE.Caption;
    TOBEtab.SelectDB('"' + TOBPCaisse.GetValue('GPK_ETABLISSEMENT') + '"', nil);
  end;
  // Valeurs par défaut
  SO_DEVISEPRINC.Value := GetParamSoc('SO_DEVISEPRINC');
  SO_REGIMEDEFAUT.Value := GetParamSoc('SO_REGIMEDEFAUT');
  SO_GCMODEREGLEDEFAUT.Value := GetParamSoc('SO_GCMODEREGLEDEFAUT');
  SO_GCTIERSDEFAUT.Text := GetParamSoc('SO_GCTIERSDEFAUT');
  SO_GCPERBASETARIF.Value := GetParamSoc('SO_GCPERBASETARIF');
  SO_GCPHOTOFICHE.Value := GetParamSoc('SO_GCPHOTOFICHE');
  SEV_CODEEVENT.Text := 'DEMARRAGE';
  // Chargement des sites de la base "centrale" sélectionnée
  if BC_LesSites <> nil then BC_LesSites.Free;
  BC_LesSites := TCollectionSites.Create(TCollectionSite, True);
  // Recherche du site central
  if (BC_LesSites <> nil) and (BC_LesSites.LeSiteLocal <> nil) then
  begin
    SITESIEGE.Text := BC_LesSites.LeSiteLocal.SSI_CODESITE;
    NOMDOSSIER.Text := BC_LesSites.LeSiteLocal.SSI_CDEPART;
  end;
end;

procedure TFExportCais.PCPDeconnecteDB;
begin
  Logout;
  DeconnecteHalley;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormClose :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
  if TOBEtab <> nil then TOBEtab.Free;
  if TOBPCaisse <> nil then TOBPCaisse.Free;
  if LstRecap <> nil then FreeAndNil(LstRecap);
  if BC_LesSites <> nil then BC_LesSites.Free;
  if not (ctxFO in V_PGI.PGIContexte) then
  begin
    if TobRepresentant<> nil then FreeAndNil(TobRepresentant);
    if TobNatPces <> nil then FreeAndNil(TobNatPces);
    if TOBRepr <> nil then FreeAndNil(TOBRepr);
    if not bFinOk then FMenuG.FermeSoc; // Initialisation annulée => on ferme l'appli !
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SO_LIBELLEEnter :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.SO_LIBELLEEnter(Sender: TObject);
begin
  inherited;
  if SO_LIBELLE.Text = '' then SO_LIBELLE.Text := LIB_CAISSE.Caption;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  LANCETOXClick :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.LANCETOXClick(Sender: TObject);
begin
  inherited;
  SEV_CODEEVENT.Enabled := LANCETOX.Checked;
  BCODEEVENT.Enabled := LANCETOX.Checked;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BSITECAISSEClick : création du site de la caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.BSITECAISSEClick(Sender: TObject);
begin
  inherited;
  AglLanceFiche('TOX', 'TOXSAISIESITES', '', '', 'ACTION=CREATION');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BSITESIEGEClick : création du site central
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.BSITESIEGEClick(Sender: TObject);
begin
  inherited;
  AglLanceFiche('TOX', 'TOXSAISIESITES', '', '', 'ACTION=CREATION');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SITECAISSEDblClick : consultation du site de la caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.SITECAISSEDblClick(Sender: TObject);
begin
  inherited;
  if SITECAISSE.Text = '' then LookupCombo(SITECAISSE)
  else AglLanceFiche('TOX', 'TOXSAISIESITES', '', SITECAISSE.Text, '');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SITESIEGEDblClick : consultation du site central
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.SITESIEGEDblClick(Sender: TObject);
begin
  inherited;
  if SITESIEGE.Text = '' then LookupCombo(SITESIEGE)
  else AglLanceFiche('TOX', 'TOXSAISIESITES', '', SITESIEGE.Text, '');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BCODEEVENTClick : création d'un événement
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.BCODEEVENTClick(Sender: TObject);
begin
  inherited;
  AglLanceFiche('TOX', 'TOXSAISIEEVENTS', '', '', 'ACTION=CREATION');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SEV_CODEEVENTDblClick : consultation d'un événement
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.SEV_CODEEVENTDblClick(Sender: TObject);
begin
  inherited;
  if SEV_CODEEVENT.Text = '' then LookupCombo(SEV_CODEEVENT)
  else AglLanceFiche('TOX', 'TOXSAISIEEVENTS', '', '001;' + SEV_CODEEVENT.Text, '');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SO_DEVISEPRINCChange : changement de devise de tenue de la caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TFExportCais.SO_DEVISEPRINCChange(Sender: TObject);
begin
  inherited;
  if EstMonnaieIN(SO_DEVISEPRINC.Value) then
  begin
    SO_TENUEEURO.Enabled := True;
  end else
  begin
    SO_TENUEEURO.Checked := False;
    SO_TENUEEURO.Enabled := False;
  end;
end;

procedure TFExportCais.US_UTILISATEUR_Change(Sender: TObject);

  procedure ErreurUser;
  begin
    US_UTILISATEUR_.text := '';
    US_LIBELLE.Text := '';
    CONFIGPCP.Text := '';
    SITECAISSE.Text := '';
    LibSite := '';
    SO_GCPREFIXEAUXI_.Text := '';
    SO_GCPREFIXEAUXIFOU_.Text := '';
    SO_LIBELLE_.Text := '';
    SO_MAIL_.Text    := '';
  end;

begin
  inherited;
  if (TOBRepr = nil) or (TOBRepr.GetValue('US_UTILISATEUR') <> US_UTILISATEUR_.Text) then
  begin
    if TOBRepr <> nil then TOBRepr.Free;
    TOBRepr := TOB.Create('UTILISAT', nil, -1);
    TOBRepr.SelectDB('"' + US_UTILISATEUR_.Text + '"', nil);
  end;
  if TOBRepr <> nil then
  begin
    if US_UTILISATEUR_.text = '' then
    begin
      ErreurUser;
      exit;
    end;
    if not TestDataUser then
    begin
      ErreurUser;
    end else
    begin
      US_LIBELLE.Text  := TOBRepr.GetValue('US_LIBELLE');        
      SO_LIBELLE_.Text := US_LIBELLE.text;
      SO_MAIL_.Text    := TobRepr.GetValue('US_EMAIL');
      if SO_MAIL_.Text = '' then
        SO_MAIL_.Text := GetParamSoc('SO_MAIL');
    end;
  end
end;

procedure TFExportCais.DBCENTRAL_OnChange(Sender: TObject);
var FormDBName: string;
//    iNumDispatch : SetInteger;
begin
  inherited;
  if DBCENTRAL.Text <> '' then
  begin
    PCPDeconnecteDB; // Déconnexion de la base courante
    FormDBName := SocForm.Text; // Connexion à la base sélectionnée
    if not ConnectBase(FormDBName, False, False) then exit;
{
    SetLength(iNumDispatch, 3);
    iNumDispatch[0] := 10;
    iNumDispatch[1] := 13;
    iNumDispatch[3] := 16;
    V_PGI.ZoomOLE := FALSE;
    V_PGI.ZoomOLE := TRUE;
    AGLLanceFiche('F','GENS','','','') ;
    FMenuG.LanceDispatchReconnect(FormDBName,iNumDispatch,true);
    }
    LoadData; // Chargement des infos de la base courante
    if TestDataBase then
    begin
      US_UTILISATEUR_.Enabled := True;
      US_UTILISATEUR_.SetFocus;
    end else
    begin
      PCPDeconnecteDB;
      DBCENTRAL.text := '';
      DBCENTRAL.SetFocus;
    end;
  end;
end;

{$IFDEF NOMADE}

procedure TFExportCais.LogToxAPlat(PhaseId: Cardinal; MessageStr: array of variant);
  function AffPhaseId(PhaseId: Cardinal): boolean;
  begin
    // A faire : gérer des niveaux de traces ...
    if V_PGI.SAV then result := True
    else Result := (PhaseId in [1..4, 7, 8, 14, 24]);
  end;
begin
  if PhaseId = TOXP_POURCENTAGE then PB.Position := Cardinal(MessageStr[0])
  else if AffPhaseId(PhaseId)
    then AfficheMemo(Format('%s : %3.3d : %s', [FormatDateTime('hh:nn:ss', Now), PhaseId, MessageStr[0]]),
      [], False, False);
end;

function TFExportCais.ConnectBase(QuelleBase : string; VerifBase, BaseUsr : boolean) : boolean;
begin
  if BaseUsr then
    AfficheMemo('Connexion sur la base utilisateur', [], False, False)
    else
    AfficheMemo('Connexion sur la base centrale', [], False, False);
  if not ConnecteHalley(QuelleBase, VerifBase, nil, nil, nil, nil) then
  begin
    PgiBox('Impossible de se connecter à la société ' + QuelleBase, 'Problème de connexion');
    Result := False;
  end else
  begin
    ChargeDescriGC;
    Result := True;
  end
end;

function TFExportCais.TestDataBase : boolean;
begin
  //Métier non renseigné dans ParamSoc
  if GetParamSoc('SO_METIER') = '' then
  begin
    GereErreur(26, THEdit(DBCENTRAL), '');
    Result := False;
    LstRecap.Clear;
    exit;
  end else
  begin
    InitMetier;
  end;
  //Recherche coordonnées
  SO_ADRESSE1_.Text   := GetParamSoc('SO_ADRESSE1');
  SO_ADRESSE2_.Text   := GetParamSoc('SO_ADRESSE2');
  SO_ADRESSE3_.Text   := GetParamSoc('SO_ADRESSE3');
  SO_CODEPOSTAL_.Text := GetParamSoc('SO_CODEPOSTAL');
  SO_DIVTERRIT_.Text  := GetParamSoc('SO_DIVTERRIT');
  SO_VILLE_.Text      := GetParamSoc('SO_VILLE');
  SO_PAYS_.Text       := GetParamSoc('SO_PAYS');
  SO_TELEPHONE_.Text  := GetParamSoc('SO_TELEPHONE');
  SO_FAX_.Text        := GetParamSoc('SO_FAX');
  Result := True;
end;

function TFExportCais.TestDataUser : boolean;
var LeSite: TCollectionSite;
    Qry : TQuery;
    Cpt : integer;
    OkSouche : boolean;
    Req, Msg, Nature, Libelle, EvtInitDefaut, EvtRemontee : string;
    Souche : variant;
    TobEvt : TOB;

  function RechercheChrono(Requete : string) : integer;
  begin
    Qry := OpenSQL('SELECT COUNT(*) AS CPT FROM '+Requete, True);
    if not Qry.Eof then
      Result := Qry.FindField('CPT').AsInteger
      else
      Result := 0;
    Ferme(Qry);
  end;

  procedure Erreur(NumErr : integer);
  begin
    GereErreur(NumErr, THEdit(US_UTILISATEUR_), '');
    LstRecap.Clear;
    Result := False;
  end;

begin
  GetVarPCP(US_UTILISATEUR_.text);
  if (not VH_GC.PCPUsVte) and (not VH_GC.PCPUsAch) then
  begin
    Erreur(19);
    exit;
  end;

  if (VH_GC.PCPUsVte) and (VH_GC.PCPUsAch) then
    CONFIGPCP.Text := TraduireMemoire('VENTE ET ACHAT')
  else if VH_GC.PCPUsVte then
    CONFIGPCP.Text := TraduireMemoire('VENTE')
  else if VH_GC.PCPUsAch then
    CONFIGPCP.Text := TraduireMemoire('ACHAT');

  if (VH_GC.PCPUsVte) and (not VH_GC.PCPVenteSeria) then
    V_PGI.VersionDemo := True;
  if (VH_GC.PCPUsAch) and (not VH_GC.PCPAchatSeria) then
    V_PGI.VersionDemo := True;
  if V_PGI.VersionDemo then
  begin
    if PgiAsk('Produit non sérialisé. Voulez-vous sérialiser maintenant ?') = mrYes then
    begin
      TestSeriaNomade;
      FMenuG.LanceDispatch(29);
      bAnnulerClick(nil);
      exit;
    end;
  end;

{$IFDEF BTP}
	Superviseur := TOBRepr.GetValue ('US_SUPERVISEUR')='X';
{$ENDIF}

  // Recherche du site associé à l'utilisateur
  if BC_LesSites <> nil then
  begin
    LeSite := BC_LesSites.FindVariableValue('$UTIL', US_UTILISATEUR_.Text, '');
    if LeSite = nil then
    begin
      Erreur(10);
      Exit;
    end else
    begin
      SITECAISSE.Text := LeSite.SSI_CODESITE;
      LibSite := LeSite.SSI_LIBELLE;
    end;
  end;

  // Contrôle de l'existence d'une souche pour chaque nature de pièce paramétrée.
  Req := 'SELECT GPP_VENTEACHAT, GPP_LIBELLE, GPP_TRSFVENTE, GPP_TRSFACHAT, GPP_NATUREPIECEG, GPP_SOUCHE, GSP_NATUREPIECEG FROM PARPIECE '+
         'LEFT JOIN SITEPIECE ON GSP_NATUREPIECEG=GPP_NATUREPIECEG AND GSP_CODESITE="'+SITECAISSE.Text+'" ';
  if (VH_GC.PCPUsVte) and (VH_GC.PCPUsAch) then
    Req := Req + 'WHERE (GPP_TRSFVENTE = "X" OR GPP_TRSFACHAT = "X") '
  else if VH_GC.PCPUsVte then
    Req := Req + 'WHERE GPP_TRSFVENTE = "X" '
  else if VH_GC.PCPUsAch then
    Req := Req + 'WHERE GPP_TRSFACHAT = "X" ';
  Req := Req + 'ORDER BY GPP_VENTEACHAT DESC, GPP_NATUREPIECEG';
  Qry := OpenSQL(Req,True);
  TobNatPces.ClearDetail;
  TobNatPces.LoadDetailDB('PARPIECE', '', '', Qry, False, True);
  Ferme(Qry);
  if TobNatPces.detail.count > 0 then
  begin
    OkSouche := True;
    for Cpt := 0 to TobNatPces.detail.count-1 do
    begin
    // DBR Agl580
      Nature := TobNatPces.Detail[Cpt].GetString ('GPP_NATUREPIECEG');
      Souche := TobNatPces.Detail[Cpt].GetString ('GSP_NATUREPIECEG');
      Libelle := TobNatPces.Detail[Cpt].GetString ('GPP_LIBELLE');
      // MODIF BTP
      if (Souche = '') and (superviseur) then
      begin
      	Souche := TobNatPces.Detail[Cpt].GetString ('GPP_SOUCHE');
      end;

      if Souche = '' then
    {
      Nature := TobNatPces.Detail[Cpt].GetValue('GPP_NATUREPIECEG');
      Souche := TobNatPces.Detail[Cpt].GetValue('GSP_NATUREPIECEG');
      Libelle := TobNatPces.Detail[Cpt].GetValue('GPP_LIBELLE');
      if VarIsNull(Souche) then
    }
      begin
        OkSouche := False;
        Msg := TraduireMemoire('La nature de pièce "')+Libelle+
               TraduireMemoire('" ne dispose pas de souche pour ')+US_LIBELLE.Text+'.';
        HShowMessage('-1;'+Caption+';'+Msg+';E;O;O;O','','');
        Break;
      end;
    end;
  end else
  begin
    OkSouche := False;
    if (VH_GC.PCPUsVte) and (VH_GC.PCPUsAch) then
      GereErreur(11, THEdit(US_UTILISATEUR_), '')
    else if VH_GC.PCPUsVte then
      GereErreur(12, THEdit(US_UTILISATEUR_), '')
    else if VH_GC.PCPUsAch then
      GereErreur(13, THEdit(US_UTILISATEUR_), '');
  end;
  if not OkSouche then
  begin
    Erreur(-1);
    exit;
  end;

  if Superviseur then
  begin
    Qry := OpenSQL('select GCL_COMMERCIAL, GCL_LIBELLE from COMMERCIAL', True);
  end else
  begin
  // Recherche du/des représentants associé(s) à l'utilisateur
  Qry := OpenSQL('select GCL_COMMERCIAL, GCL_LIBELLE from COMMERCIAL where GCL_UTILASSOCIE ="'+ US_UTILISATEUR_.Text + '"', True);
  end;
  TobRepresentant.LoadDetailDB('COMMERCIAL','','',Qry,False);
  Ferme(Qry);
  if TobRepresentant.Detail.count = 0 then
  begin
    Erreur(15);
    Exit;
  end;

  //Auxiliaire des tiers et préparation des évènements
  EvtInitDefaut := 'PCPU_GENERIQUE';
  EvtRemontee   := '';
  SO_GCPREFIXEAUXI_.Text := '';
  SO_GCPREFIXEAUXIFOU_.Text := '';
  if VH_GC.PCPUsVte then
  begin
{$IFDEF BTP}
		if Superviseur then
    begin
    	// Mise en place des portables de type administrateur
      EvtInitDefaut := EvtInitDefaut+';PCPS_INIT_VTE';
      EvtRemontee   := ';PCPS_REMONTEE_VTE';
    end else
    begin
      EvtInitDefaut := EvtInitDefaut+';PCPU_INIT_VTE';
      EvtRemontee   := ';PCPU_REMONTEE_VTE';
    end;
{$ELSE}
    EvtInitDefaut := EvtInitDefaut+';PCPU_INIT_VTE';
    EvtRemontee   := ';PCPU_REMONTEE_VTE';
{$ENDIF}
    SO_GCPREFIXEAUXI_.Text := GetParamSoc('SO_GCPREFIXEAUXI');
  end;
  if VH_GC.PCPUsAch then
  begin
    EvtInitDefaut := EvtInitDefaut+';PCPU_INIT_ACH';
    EvtRemontee   := EvtRemontee+';PCPU_REMONTEE_ACH';
    SO_GCPREFIXEAUXIFOU_.Text := GetParamSoc('SO_GCPREFIXEAUXIFOU');
  end;
//  EvtRemontee := Copy(EvtRemontee,2,length(EvtREmontee)); PCPU_DEMANDE
  EvtRemontee := Copy(EvtRemontee,2,length(EvtREmontee))+';PCPU_DEMANDE';
  //Chargement des grilles
{$IFDEF BTP}
	if Superviseur then
  begin
  	Qry := OpenSQL('SELECT * FROM STOXEVENTS WHERE (SEV_CODEEVENT LIKE "PCPS%") OR (SEV_CODEEVENT="PCPU_GENERIQUE") OR (SEV_CODEEVENT="PCPU_DEMANDE")', True);
  end else
  begin
  	Qry := OpenSQL('SELECT * FROM STOXEVENTS WHERE SEV_CODEEVENT LIKE "PCPU%"', True);
  end;
{$ELSE}
  Qry := OpenSQL('SELECT * FROM STOXEVENTS WHERE SEV_CODEEVENT LIKE "PCPU%"', True);
{$ENDIF}
  TOBEvt := TOB.Create('Les STOXEVENTS', nil, -1);
  TOBEvt.LoadDetailDB('STOXEVENTS', '', '', Qry, False);
  Ferme(Qry);
  TOBEvt.PutGridDetail(GEVTS, False, False, 'SEV_CODEEVENT;SEV_LIBELLE', True);
  TOBEvt.PutGridDetail(GEVTINIT, False, False, 'SEV_CODEEVENT;SEV_LIBELLE', True);
  // Chargement de la grille de sélection de l'évènement d'initialisation
  for Cpt := GEVTINIT.RowCount -1 downto 1 do
//    if (GEVTINIT.Cells[0,Cpt] <> 'PCPU_GENERIQUE') and (Pos('PCPU_INIT_',GEVTINIT.Cells[0,Cpt]) = 0) then
    if Pos(GEVTINIT.Cells[0,Cpt],EvtInitDefaut) = 0 then
      GEVTINIT.DeleteRow(Cpt);
  for Cpt := GEVTS.RowCount -1 downto 1 do
    if Pos(GEVTS.Cells[0,Cpt],EvtRemontee) = 0 then
      GEVTS.DeleteRow(Cpt);
  //Selection des evts
  GEVTINIT.ClearSelected;
  GEVTS.ClearSelected;
  for Cpt := 1 to GEVTINIT.RowCount -1 do
    if pos(GEVTINIT.Cells[0,Cpt],EvtInitDefaut) > 0 then GEVTINIT.FlipSelection(Cpt);
  for Cpt := 1 to GEVTS.RowCount -1 do
    if pos(GEVTS.Cells[0,Cpt],EvtRemontee) > 0 then GEVTS.FlipSelection(Cpt);
  TOBEvt.Free;
  EvtOblInit := EvtInitDefaut;
  EvtOblRemontee := EvtRemontee;


  //Paramètres articles
  if GetParamSoc('SO_GCLGNUMART') = 0 then
    SO_GCLGNUMART_.Value := 8
    else
    SO_GCLGNUMART_.Value  := GetParamSoc('SO_GCLGNUMART') + length(VH_GC.PCPPrefixe);
  SO_GCPREFIXEART_.Text   := VH_GC.PCPPrefixe;
  SO_GCCOMPTEURART_.Value := RechercheChrono('ARTICLE WHERE GA_ARTICLE LIKE "'+US_UTILISATEUR_.Text+'%" ');
  //Code des tiers
  if ctxMode in V_PGI.PGIContexte then                              
  begin
    SO_GCLGNUMTIERS_.Visible := True;
    LSO_GCLGNUMTIERS_.Visible := True;
    SO_GCLGNUMTIERS_.Value := GetParamSoc('SO_GCLGNUMTIERS');
  end else
  begin
    LSO_GCLGNUMTIERS_.Visible := False;
    SO_GCLGNUMTIERS_.Visible := False;
    SO_GCLGNUMTIERS_.Value := 0;
  end;
  SO_GCPREFIXETIERS_.Text   := VH_GC.PCPPrefixe;
  SO_GCCOMPTEURTIERS_.Value := RechercheChrono('TIERS WHERE T_NATUREAUXI="CLI" AND T_TIERS LIKE "'+US_UTILISATEUR_.Text+'%" '+
                                               'AND T_REPRESENTANT IN (SELECT GCL_COMMERCIAL FROM '+
                                               'COMMERCIAL WHERE GCL_UTILASSOCIE="'+US_UTILISATEUR_.Text+'")');
  Result := True;
end;

procedure TFExportCais.bImprimerClick(Sender: TObject);
var TobToPrint : TOB;
    Cpt : integer;
begin
  inherited;
  TobToPrint := TOB.Create('',nil,-1);
  for Cpt := 0 to LstRecap.Items.Count -1 do
    UtilTobCreat(TobToPrint,'','',LstRecap.Items[Cpt],'');
  UtilTobPrint(TobToPrint,Caption,1);
  if TobToPrint <> nil then FreeAndNil(TobToPrint);
end;

procedure TFExportCais.CreerRapport;
var Cpt : integer;
begin
  LstRecap.Clear;
  //Utilisateur
  LstRecap.Items.Add(TraduireMemoire('INFORMATIONS UTILISATEUR'));
  LstRecap.Items.Add(TraduireMemoire(' Code et nom   : '+US_UTILISATEUR_.text+' - '+US_LIBELLE.text));
  LstRecap.Items.Add(TraduireMemoire(' Site distant  : '+SITECAISSE.Text+' - '+LibSite));
  for Cpt := 0 to TobRepresentant.detail.count -1 do
  begin
    if Cpt = 0 then
      LstRecap.Items.Add(TraduireMemoire(' Représentant  : ')+TobRepresentant.detail[Cpt].GetValue('GCL_LIBELLE'))
      else
      LstRecap.Items.Add(TraduireMemoire('                 ')+TobRepresentant.detail[Cpt].GetValue('GCL_LIBELLE'));
  end;
  LstRecap.Items.Add(TraduireMemoire(' Configuration : '+ CONFIGPCP.Text));
  //Paramètres pièces
  LstRecap.Items.Add('');
  LstRecap.Items.Add(TraduireMemoire('PARAMETRE PIECES'));
  for Cpt := 0 to TobNatPces.detail.count-1 do
  begin
    if Cpt = 0 then
      LstRecap.Items.Add(TraduireMemoire(' Natures à transférer     : ')+TobNatPces.Detail[Cpt].GetValue('GPP_LIBELLE'))
      else
      LstRecap.Items.Add('                          : '+TobNatPces.Detail[Cpt].GetValue('GPP_LIBELLE'));
  end;
  if GP_VERROUILLEPCES_.Checked then
    LstRecap.Items.Add(TraduireMemoire(' Vérouiller après envoi   : Oui'))
    else
    LstRecap.Items.Add(TraduireMemoire(' Vérouiller après envoi   : Non'));
  if GP_IMPIMMEDIATE_.Checked then
    begin
    if GP_APERCUAVIMP_.Checked then
      LstRecap.Items.Add(TraduireMemoire(' Imprimer à la validation : Oui (avec aperçu)'))
      else
      LstRecap.Items.Add(TraduireMemoire(' Imprimer à la validation : Oui (sans aperçu)'));
    end else
    LstRecap.Items.Add(TraduireMemoire(' Imprimer à la validation : Non'));
  //Codification
  LstRecap.Items.Add('');
  LstRecap.Items.Add(TraduireMemoire('CODIFICATION ARTICLE'));
  LstRecap.Items.Add(TraduireMemoire(' Préfixe            : ')+SO_GCPREFIXEART_.text);
  LstRecap.Items.Add(TraduireMemoire(' Chrono départ      : ')+FloatToStr(SO_GCCOMPTEURART_.value));
  LstRecap.Items.Add(TraduireMemoire(' Longueur           : ')+IntToStr(SO_GCLGNUMART_.value));
  LstRecap.Items.Add('');
  LstRecap.Items.Add(TraduireMemoire('CODIFICATION TIERS '));
  LstRecap.Items.Add(TraduireMemoire(' Préfixe            : ')+SO_GCPREFIXETIERS_.text);
  LstRecap.Items.Add(TraduireMemoire(' Chrono départ      : ')+FloatToStr(SO_GCCOMPTEURTIERS_.value));
  if ctxMode in V_PGI.PGIContexte then
    LstRecap.Items.Add(TraduireMemoire(' Longueur           : ')+IntToStr(SO_GCLGNUMTIERS_.value));
  if VH_GC.PCPVenteSeria then
    LstRecap.Items.Add(TraduireMemoire(' Préfixe auxiliaire client : ')+SO_GCPREFIXEAUXI_.text);
  if VH_GC.PCPAchatSeria then
    LstRecap.Items.Add(TraduireMemoire(' Préfixe auxiliaire fourn. : ')+SO_GCPREFIXEAUXIFOU_.text);
  //Coordonnées
  LstRecap.Items.Add('');
  LstRecap.Items.Add(TraduireMemoire('COORDONNEES'));
  LstRecap.Items.Add(TraduireMemoire(' Nom                  : ')+SO_LIBELLE_.text);
  LstRecap.Items.Add(TraduireMemoire(' Adresse              : ')+SO_ADRESSE1_.text);
  LstRecap.Items.Add('                      : '+SO_ADRESSE2_.text);
  LstRecap.Items.Add('                      : '+SO_ADRESSE3_.text);
  LstRecap.Items.Add(TraduireMemoire(' Code postal          : ')+SO_CODEPOSTAL_.text);
  LstRecap.Items.Add(TraduireMemoire(' Div. Territoriale    : ')+SO_DIVTERRIT_.text);
  LstRecap.Items.Add(TraduireMemoire(' Ville                : ')+SO_VILLE_.text);
  LstRecap.Items.Add(TraduireMemoire(' Pays                 : ')+SO_PAYS_.text);
  LstRecap.Items.Add(TraduireMemoire(' Téléphone            : ')+SO_TELEPHONE_.text);
  LstRecap.Items.Add(TraduireMemoire(' Télécopie            : ')+SO_FAX_.text);
  LstRecap.Items.Add(TraduireMemoire(' E-mail               : ')+SO_MAIL_.text);
end;
{$ENDIF}

procedure TFExportCais.GP_IMPIMMEDIATE_Click(Sender: TObject);
begin
  inherited;
  GP_APERCUAVIMP_.Checked := False;
  GP_APERCUAVIMP_.Enabled := (GP_IMPIMMEDIATE_.Checked);
end;

procedure TFExportCais.bParamPceClick(Sender: TObject);
begin
  inherited;
{$IFDEF BTP}
	AglLanceFiche ('GC','GCPARPIECE','','','') // LanceparPiece
{$ELSE}
  if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
    AglLanceFiche ('GC','GCPARPIECE','','','') // LanceparPiece
    else
    AglLanceFiche ('GC','GCPARPIECEUSR','','','') ; // LanceparPiece  utilisateur
{$ENDIF}
end;

procedure TFExportCais.bAnnulerClick(Sender: TObject);
begin
  inherited;
  ;
end;

end.
