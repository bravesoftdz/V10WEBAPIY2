{***********UNITE*************************************************
Auteur  ...... : N. ACHINO et D. CARRET
Créé le ...... : 20/07/2001
Modifié le ... : 19/06/2002
Description .. : Assistant d'importation du fichier de démarrage d'une caisse
Mots clefs ... : FO
*****************************************************************}
unit FOImportCais;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
  Mask, Spin, Psock, NMFtp, HRichEdt, HRichOLE, HEnt1, UTOB, dbtables,
  Buttons, HStatus, HPanel, inifiles, ParamSoc, FileCtrl, FOToxUtil,
  UTox, uToxConf, uToxFiche, uToxConst, uToxNet, UToxFuncs, uToxClient, gcToxWork;

function FOImportCaisse: Boolean;
function PCPImportRepr: Boolean;

type
  TFImportCais = class(TFAssist)
    FTPSC: TNMFTP;
    PAGE2: TTabSheet;
    PAGE6: TTabSheet;
    PAGE4: TTabSheet;
    GDEVISE: TGroupBox;
    GCOORD: TGroupBox;
    LSO_ASRESSE1: TLabel;
    LSO_MAIL: TLabel;
    LSO_FAX: TLabel;
    LSO_TELEPHONE: TLabel;
    SO_SOCIETE: TEdit;
    SO_LIBELLE: TEdit;
    SO_ADRESSE1: TEdit;
    SO_ADRESSE2: TEdit;
    SO_ADRESSE3: TEdit;
    SO_CODEPOSTAL: TEdit;
    SO_VILLE: TEdit;
    SO_TELEPHONE: TEdit;
    SO_FAX: TEdit;
    SO_MAIL: TEdit;
    GETAB: TGroupBox;
    LSO_ETABLISDEFAUT: THLabel;
    SO_ETABLISDEFAUT: THValComboBox;
    PAGE1: TTabSheet;
    GORIGINE: TRadioGroup;
    PFICHIER: TPanel;
    LNOMDOSSIER: THLabel;
    NOMDOSSIER: THCritMaskEdit;
    PTELECOM: TPanel;
    LADRESSEIP1: THLabel;
    LADRESSEIP2: THLabel;
    LADRESSEIP3: THLabel;
    LUSER: THLabel;
    LMOTPASSE: THLabel;
    LADRESSEIP4: THLabel;
    ADRESSEIP1: TSpinEdit;
    ADRESSEIP2: TSpinEdit;
    ADRESSEIP3: TSpinEdit;
    USER: THCritMaskEdit;
    MOTPASSE: THCritMaskEdit;
    ADRESSEIP4: TSpinEdit;
    PAGE3: TTabSheet;
    CPTRDURECUP: THRichEditOLE;
    bCharge: TSpeedButton;
    PTITRE1: THPanel;
    TITRE1: THLabel;
    PTITRE2: THPanel;
    TITRE2: THLabel;
    PTITRE3: THPanel;
    TITRE3: THLabel;
    PAGE5: TTabSheet;
    GDIVERS: TGroupBox;
    LSO_REGIMEDEFAUT: THLabel;
    LSO_GCPERBASETARIF: THLabel;
    SO_REGIMEDEFAUT: THValComboBox;
    SO_GCPERBASETARIF: THValComboBox;
    PTITRE5: THPanel;
    TITRE5: THLabel;
    PTITRE4: THPanel;
    TITRE4: THLabel;
    LSO_GCMODEREGLEDEFAUT: THLabel;
    SO_GCMODEREGLEDEFAUT: THValComboBox;
    LSO_GCTIERSDEFAUT: THLabel;
    LSO_GCPHOTOFICHE: THLabel;
    SO_GCPHOTOFICHE: THValComboBox;
    PAGE7: TTabSheet;
    PTITRE7: THPanel;
    CPTRENDU: THRichEditOLE;
    TITRE7: THLabel;
    TITRE71: THLabel;
    LGPK_CAISSE: TLabel;
    GPK_CAISSE: THCritMaskEdit;
    GACTIONTABLE: TRadioGroup;
    PAGE8: TTabSheet;
    PTITRE8: THPanel;
    TITRE8: THLabel;
    LANCETOX: TCheckBox;
    SO_PAYS: THValComboBox;
    SO_GCTIERSDEFAUT: THValComboBox;
    lEtapeManu: THLabel;
    LSO_DEVISEPRINC: THLabel;
    LSO_DECVALEUR: THLabel;
    LSO_TAUXEURO: THLabel;
    HLabel1: THLabel;
    SO_DEVISEPRINC: THValComboBox;
    SO_DECVALEUR: TSpinEdit;
    SO_TAUXEURO: THNumEdit;
    SO_FONGIBLE: THValComboBox;
    SO_TENUEEURO: TCheckBox;
    procedure GORIGINEClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure bChargeClick(Sender: TObject);
    procedure ORIGINEChange(Sender: TObject);
    procedure AffEtape;
  private
    { Déclarations privées }
    TOBReg: TOB;
    TOBTva: TOB;
    TOBPhoto: TOB;
    TOBDevise: TOB;
    TOBModeRegl: TOB;
    TOBDModePaie: TOB;
    TOBTiers: TOB;
    TOBExercice: TOB;
    TOBEtb: TOB;
    TOBDepot: TOB;
    TOBCaisse: TOB;
    TOBSouche: TOB;
    TOBPer: TOB;
    TOBToxParms: TOB;
    TOBToxvars: TOB;
    TOBToxChronos: TOB;
    TOBSite: TOB;
    TOBInfo: TOB;
    TOBToxQry: TOB;
    TOBToxEvt: TOB;
    TOBUserGrp: TOB;
    TOBMenu: TOB;
    TOBUtil: TOB;
    TOBCPProfilUserC : TOB;
    TOBParam: TOB;
    TOBGeneraux : TOB;
    // Ajout pour PCP
    TOBSitePiece: TOB;
    TOBParPiece : TOB;
    // Fin ajout pour PCP
    DonneesChargees: Boolean;
    //    SortieHalley    : Boolean ;
    EtabOrg: string;
    DeviseOrg: string;
    CurrentPage: integer; // Numéro de la page courante
    PagesCount: integer; // Nb de pages utilisées
    procedure AfficheMemo(Texte: string; Style: TFontStyles; Saut, Traduire: Boolean);
    procedure ActiveBtn(Active: Boolean);
    function NomFichierConf(AvecChemin: Boolean): string;
    procedure RenommeFichierConf;
    procedure TOBToCombo(DataType: Integer; TOBData: TOB; Combo: THValComboBox);
    function RecupFichierConf: Boolean;
    procedure OnSuccess(Trans_Type: TCmdType);
    procedure OnFailure(var Handled: Boolean; Trans_Type: TCmdType);
    procedure OnPacketRecvd(Sender: TObject);
    procedure OnTransactionStart(Sender: TObject);
    procedure OnTransactionStop(Sender: TObject);
    //    function  ChargeFichierConf : Boolean ;
    procedure FusionneTOB(LaTOB, TOBData: TOB; NomTable, Sql: string; Cle: array of string; Valeur: string);
    function ImportTOB(LaTOB: TOB): Integer;
    procedure SauveCaisse;
    //    procedure SauvePCP ;
    procedure AvertirEuro;
    procedure MAJTable(TOBData: TOB; EraseTable: Boolean);
    procedure MAJParamSoc;
    procedure MAJSouche;
    procedure MAJDevise;
    procedure MAJCaisse;
    procedure MAJTOXSite;
  public
    { Déclarations publiques }
    SortieHalley: Boolean;
    function ChargeFichierConf: Boolean;
    procedure SauvePCP;
    procedure AfficheTOB;
  end;

const
  ouFO = 1;
  ouPCP = 2;
  // Nom des pages disponibles pour chaque environnement (FO/PCP) -> P.Pages[i].Name
  OngletUtile: array[1..2] of string = (
    {1 FO}';PAGE1;PAGE2;PAGE3;PAGE4;PAGE5;PAGE6;PAGE7;PAGE8;'
    {2 PCP}, ';PAGE1;PAGE2;PAGE3;PAGE7;PAGE8;'
    );
implementation

uses StockUtil;

{$R *.DFM}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 20/07/2001
Modifié le ... : 23/07/2001
Description .. : Installation d'une caisse à partir du fichier de démarrage.
Mots clefs ... : FO
*****************************************************************}

function FOImportCaisse: Boolean;
var X: TFImportCais;
begin
  X := TFImportCais.Create(Application);
  try
    X.ShowModal;
  finally
    Result := X.SortieHalley;
    X.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : D. Carret
Créé le ...... : 04/07/2002
Modifié le ... : 04/07/2002
Description .. : Installation d'une caisse à partir du fichier de démarrage.
Mots clefs ... : PCP
*****************************************************************}

function PCPImportRepr: Boolean;
var X: TFImportCais;
begin
  X := TFImportCais.Create(Application);
  try
    // Mise en place des libellés spécifiques de l'assistant PCP
    X.Caption := TraduireMemoire('Import du fichier de démarrage d''un poste PCP');
    X.TITRE1.Caption := TraduireMemoire(#10 + #13 + 'Cet assistant va vous aider à initialiser votre poste');
    X.LGPK_CAISSE.Caption := TraduireMemoire('Entrez votre code représentant');
    X.TITRE2.Caption := TraduireMemoire('Indiquez le dossier de stockage de démarrage de votre poste');
    X.TITRE3.Caption := TraduireMemoire('Lancez le chargement du fichier de données à partir du fichier de démarrage de votre poste');
    X.TITRE8.Caption := TraduireMemoire('Vous pouvez enchaîner sur la récupération des données disponibles sur le site central');
    X.LANCETOX.Caption := TraduireMemoire('Démarrage des échanges : portable - site central');
    X.TITRE7.Caption := TraduireMemoire('Le bouton ''Fin'' lance l''intégration du fichier de démarrage de votre poste');
    X.ShowModal;
  finally
    Result := X.SortieHalley;
    X.Free;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormShow :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.FormShow(Sender: TObject);
begin
  inherited;
  // On inactive le bouton FIN
  bFin.Enabled := False;
  CPTRDURECUP.Clear;
  CPTRENDU.Clear;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormCreate :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.FormCreate(Sender: TObject);
var IndiceOnglet: Integer;
  bVisible: boolean;
  StOnglet, StTmp: string;
begin
  inherited;
  DonneesChargees := False;
  SortieHalley := False;
  EtabOrg := '';
  DeviseOrg := '';
  TOBReg := nil;
  TOBTva := nil;
  TOBPhoto := nil;
  TOBDevise := nil;
  TOBModeRegl := nil;
  TOBDModePaie := nil;
  TOBTiers := nil;
  TOBExercice := nil;
  TOBEtb := nil;
  TOBDepot := nil;
  TOBCaisse := nil;
  TOBSouche := nil;
  TOBPer := nil;
  TOBToxParms := nil;
  TOBToxvars := nil;
  TOBToxChronos := nil;
  TOBSite := nil;
  TOBInfo := nil;
  TOBToxQry := nil;
  TOBToxEvt := nil;
  TOBUserGrp := nil;
  TOBMenu := nil;
  TOBUtil := nil;
  TOBCPProfilUserC := nil;
  TOBParam := nil;
  TOBGeneraux := nil;
  TOBSitePiece := nil;
  TOBParPiece := nil;

  // Affichage ou non des onglets suivant le type d'assistant : FO / PCP
  PagesCount := 0;
  CurrentPage := 1;
  if ctxFO in V_PGI.PGIContexte then StOnglet := OngletUtile[ouFO] else StOnglet := OngletUtile[ouPCP];
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
//  FormClose :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
  if TOBReg <> nil then FreeAndNil(TOBReg);
  if TOBTva <> nil then FreeAndNil(TOBTva);
  if TOBPhoto <> nil then FreeAndNil(TOBPhoto);
  if TOBDevise <> nil then FreeAndNil(TOBDevise);
  if TOBModeRegl <> nil then FreeAndNil(TOBModeRegl);
  if TOBDModePaie <> nil then FreeAndNil(TOBDModePaie);
  if TOBTiers <> nil then FreeAndNil(TOBTiers);
  if TOBExercice <> nil then FreeAndNil(TOBExercice);
  if TOBEtb <> nil then FreeAndNil(TOBEtb);
  if TOBDepot <> nil then FreeAndNil(TOBDepot);
  if TOBCaisse <> nil then FreeAndNil(TOBCaisse);
  if TOBSouche <> nil then FreeAndNil(TOBSouche);
  if TOBPer <> nil then FreeAndNil(TOBPer);
  if TOBToxParms <> nil then FreeAndNil(TOBToxParms);
  if TOBToxvars <> nil then FreeAndNil(TOBToxvars);
  if TOBToxChronos <> nil then FreeAndNil(TOBToxChronos);
  if TOBSite <> nil then FreeAndNil(TOBSite);
  if TOBInfo <> nil then FreeAndNil(TOBInfo);
  if TOBToxQry <> nil then FreeAndNil(TOBToxQry);
  if TOBToxEvt <> nil then FreeAndNil(TOBToxEvt);
  if TOBUserGrp <> nil then FreeAndNil(TOBUserGrp);
  if TOBMenu <> nil then FreeAndNil(TOBMenu);
  if TOBUtil <> nil then FreeAndNil(TOBUtil);
  if TOBParam <> nil then FreeAndNil(TOBParam);
  if TOBGeneraux <> nil then FreeAndNil(TOBGeneraux);
  if TOBSitePiece <> nil then FreeAndNil(TOBSitePiece);
  if TOBParPiece <> nil then FreeAndNil(TOBParPiece);
  if TOBCPProfilUserC <> nil then FreeAndNil(TOBCPProfilUserC);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GORIGINEClick :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.GORIGINEClick(Sender: TObject);
begin
  inherited;
  PTELECOM.Visible := (GORIGINE.ItemIndex = 1);
  DonneesChargees := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ORIGINEChange
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.ORIGINEChange(Sender: TObject);
begin
  inherited;
  DonneesChargees := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  bSuivantClick : bouton SUIVANT
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.bSuivantClick(Sender: TObject);
begin
  // Verification des paramètres FTP
  if P.ActivePage.Name = 'PAGE1' then
  begin
    if Trim(GPK_CAISSE.Text) = '' then
    begin
      if CtxFO in V_PGI.PGIContexte then Msg.Execute(6, Caption, '')
      else Msg.Execute(10, Caption, '');
      if GPK_CAISSE.CanFocus then GPK_CAISSE.SetFocus;
      Exit;
    end;
  end
  else if P.ActivePage.Name = 'PAGE2' then
  begin
    if GORIGINE.ItemIndex = 1 then
    begin
      if (ADRESSEIP1.Value = 0) and (ADRESSEIP2.Value = 0) and
        (ADRESSEIP3.Value = 0) and (ADRESSEIP4.Value = 0) then
      begin
        Msg.Execute(5, Caption, '');
        if ADRESSEIP1.CanFocus then ADRESSEIP1.SetFocus;
        Exit;
      end;
      if Trim(USER.Text) = '' then
      begin
        Msg.Execute(2, Caption, '');
        if USER.CanFocus then USER.SetFocus;
        Exit;
      end;
      if Trim(MOTPASSE.Text) = '' then
      begin
        Msg.Execute(3, Caption, '');
        if MOTPASSE.CanFocus then MOTPASSE.SetFocus;
        Exit;
      end;
    end;
    NOMDOSSIER.Text := Trim(NOMDOSSIER.Text);
    if not DirectoryExists(NOMDOSSIER.Text) then
    begin
      Msg.Execute(4, Caption, '');
      if NOMDOSSIER.CanFocus then NOMDOSSIER.SetFocus;
      Exit;
    end;
  end
  else if P.ActivePage.Name = 'PAGE3' then
  begin
    if not DonneesChargees then
    begin
      Msg.Execute(7, Caption, '');
      Exit;
    end;
  end;

  // Gestion des onglets non visibles !
  while (P.ActivePage.PageIndex < P.PageCount - 1) and not P.Pages[P.ActivePage.PageIndex + 1].TabVisible do
    P.ActivePage.PageIndex := P.ActivePage.PageIndex + 1;
  inherited;
  if P.ActivePage.Name = 'PAGE7' then
    if ctxFO in V_PGI.PGIContexte then TITRE71.Caption := GPK_CAISSE.Text //zzz+ ' - ' + LIB_CAISSE.Caption
    else TITRE71.Caption := GPK_CAISSE.Text; //+ ' - ' + GCL_LIBELLE.Text ;
  // Maj label étape + Activation du bouton FIN
  inc(CurrentPage);
  AffEtape;
  if (CurrentPage = PagesCount) then bFin.Enabled := True;
end;

procedure TFImportCais.AffEtape;
begin
  lEtapeManu.Caption := Msg.Mess[0] + ' ' + IntToStr(CurrentPage) + '/' + IntToStr(PagesCount);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  bPrecedentClick : bouton PRECEDENT
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.bPrecedentClick(Sender: TObject);
begin
  // Gestion des onglets non visibles !
  while (P.ActivePage.PageIndex > 1) and not P.Pages[P.ActivePage.PageIndex - 1].TabVisible
    do P.ActivePage.PageIndex := P.ActivePage.PageIndex - 1;
  inherited;
  // Maj label étape + Inactivation des boutons PRECEDENT et FIN
  dec(CurrentPage);
  AffEtape;
  bPrecedent.Enabled := (P.ActivePage.Name <> 'PAGE1');
  bFin.Enabled := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  bFinClick : bouton FIN
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.bFinClick(Sender: TObject);
var Nb: Integer;
  QQ: TQuery;
  IdApplication: string;
  WhenStart: TDateTime;
begin
  inherited;
  ActiveBtn(False);
  SourisSablier;
  AfficheMemo('Mise à jour de la base', [fsBold, fsUnderline], True, True);
  QQ := OpenSQL('select count(*) from PIECE', True);
  if QQ.Eof then Nb := 0 else Nb := QQ.Fields[0].AsInteger;
  Ferme(QQ);
  if (Nb = 0) or (PGIAsk('Attention, ce dossier n''est pas vide.#10Voulez-vous continuer ?', Caption) = mrYes) then
  begin
    if ctxFO in V_PGI.PGIContexte then SauveCaisse else SauvePCP;
    RenommeFichierConf;
  end;
  AfficheMemo('Mise à jour de la base terminée.', [fsBold, fsUnderline], True, True);
  //{$IFDEF PB_COMPILE}
  if LANCETOX.Checked then
  begin
    // Lancement du Tox Serveur
    //if not AglToxServeurActifNet then
    if not AglToxCommuniCam(UST_RUNNING, nil) then
    begin
      AfficheMemo('Lancement du serveur de communication ...', [fsBold], False, True);
      //AglToxServeurRunNet ;
      AglToxCommuniCam(UST_START, nil);
      Delay(3000); // attente de 3 secondes
    end;

    // Lancement d'une connexion immédiate
    //if AglToxServeurRunConnexion then
    if AglToxCommuniCam(UST_RUNNING, nil) then
    begin
      AfficheMemo('Lancement d''une connexion immédiate ...', [fsBold], False, True);
      AglToxCommuniCam(UST_ForceExecute, nil);
      Delay(5000); // attente de 5 secondes
      // Démarrage des échanges FO - BO
      if not AglStatusTox(WhenStart) then
      begin
        AfficheMemo('Démarrage des échanges FO-BO', [fsBold], False, True);
        IdApplication := FOGetToxIdApp;
        AglToxConf(aceStart, IdApplication, GescomModeNotConfirmeTox, GescomModeTraitementTox, AglToxFormError);
      end;
    end else
    begin
      AfficheMemo('Problème pour transférer les fichiers !', [fsBold], False, True);
    end;
  end else SortieHalley := True;
  //{$ENDIF}
  SourisNormale;
  bAnnuler.Caption := TraduireMemoire('&Fermer');
  bAnnuler.Enabled := True;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  bChargeClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.bChargeClick(Sender: TObject);
var Ok: Boolean;
begin
  inherited;
  bCharge.Enabled := False;
  ActiveBtn(False);
  SourisSablier;
  // Chargement des données
  AfficheMemo('Chargement des données', [fsBold, fsUnderline], False, True);
  if GORIGINE.ItemIndex = 1 then Ok := RecupFichierConf
  else Ok := True;
  if Ok then Ok := ChargeFichierConf;
  if Ok then AfficheTOB;
  AfficheMemo('Chargement des données terminé.', [fsBold, fsUnderline], False, True);
  DonneesChargees := Ok;
  SourisNormale;
  ActiveBtn(True);
  bCharge.Enabled := True;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  NomFichierConf : Retourne le nom du fichier de configuration
///////////////////////////////////////////////////////////////////////////////////////

function TFImportCais.NomFichierConf(AvecChemin: Boolean): string;
var Racine : string;
begin
  if ctxFO in V_PGI.PGIContexte then
  begin
    if AvecChemin then
      Result := IncludeTrailingBackslash(NOMDOSSIER.Text)
      else
      Result := '';
    racine := 'FO';
  end else
  begin
    if AvecChemin then
      Result := ExtractFileDrive(Application.ExeName) + '\PGI00\Std\'
      else
      Result := '';
    racine := 'PCP';
  end;
  Result := Result + racine + 'Cnf' + Trim(GPK_CAISSE.Text) + '.txt';
end;

///////////////////////////////////////////////////////////////////////////////////////
//  RenameFichierConf : Renomme le nom du fichier de configuration
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.RenommeFichierConf;
var OldName, NewName: string;
begin
  OldName := NomFichierConf(True);
  if FileExists(OldName) then
  begin
    NewName := ChangeFileExt(OldName, '.old');
    if FileExists(NewName) then DeleteFile(NewName);
    RenameFile(OldName, NewName);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  RecupFichierConf : Récupération par FTP du fichier de configuration
///////////////////////////////////////////////////////////////////////////////////////

function TFImportCais.RecupFichierConf: Boolean;
begin
  Result := False;
  FTPSC.UserID := USER.Text;
  FTPSC.Host := IntToStr(ADRESSEIP1.Value) + '.' + IntToStr(ADRESSEIP2.Value) + '.'
    + IntToStr(ADRESSEIP3.Value) + '.' + IntToStr(ADRESSEIP4.Value);
  FTPSC.Password := MOTPASSE.Text;
  FTPSC.Vendor := NMOS_AUTO;
  AfficheMemo('Début de la communication avec le site central', [fsBold], False, True);
  try
    FTPSC.Connect;
  except
    AfficheMemo('La connexion au site central a échouée', [fsBold], False, True);
  end;
  if FTPSC.Connected then
  begin
    Result := True;
    FTPSC.Mode(MODE_ASCII);
    FTPSC.OnSuccess := OnSuccess;
    FTPSC.OnFailure := OnFailure;
    FTPSC.OnPacketRecvd := OnPacketRecvd;
    FTPSC.OnTransactionStart := OnTransactionStart;
    FTPSC.OnTransactionStop := OnTransactionStop;
    AfficheMemo('Chargement du fichier de configuration', [fsBold], False, True);
    try
      FTPSC.Download(NomFichierConf(False), NomFichierConf(True));
    except
      Result := False;
    end;
    AfficheMemo('Fin de la communication avec le site central', [fsBold], False, True);
    FTPSC.Disconnect;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnSuccess :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.OnSuccess(Trans_Type: TCmdType);
begin
  AfficheMemo('Le fichier de configuration a été récupéré avec succès', [fsBold], False, True);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnFailure :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.OnFailure(var Handled: Boolean; Trans_Type: TCmdType);
begin
  Handled := False;
  AfficheMemo('Le fichier de configuration n''a pu être récupéré', [fsBold], False, True);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnPacketRecvd :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.OnPacketRecvd(Sender: TObject);
begin
  MoveCur(False);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnTransactionStart :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.OnTransactionStart(Sender: TObject);
begin
  InitMove(20, '');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnTransactionStop :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.OnTransactionStop(Sender: TObject);
begin
  FiniMove;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeFichierConf : Chargement du fichier de configuration
///////////////////////////////////////////////////////////////////////////////////////

function TFImportCais.ChargeFichierConf: Boolean;
var Fichier: string;
begin
  Fichier := NomFichierConf(True);
  if not FileExists(Fichier) then
  begin
    if ctxFO in V_PGI.PGIContexte then Msg.Execute(8, Caption, '')
    else Msg.Execute(11, Caption, '');
    Result := False;
    Exit;
  end;
  AfficheMemo('Chargement des données depuis le fichier de configuration', [fsBold], False, True);
  AfficheMemo('  => ' + Fichier, [], False, False);
  Result := TOBLoadFromFile(Fichier, ImportTOB);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TOBToCombo : constitue la liste de valeur d'un combo à partir d'une TOB
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.TOBToCombo(DataType: Integer; TOBData: TOB; Combo: THValComboBox);
var TOBL: TOB;
  Ind: Integer;
  Ok: Boolean;
  Code, Libelle, Filtre, Valeur: string;
begin
  case DataType of
    1: // Devises
      begin
        Code := 'D_DEVISE';
        Libelle := 'D_LIBELLE';
        Filtre := '';
        Valeur := '';
      end;
    2: // Etablissements
      begin
        Code := 'ET_ETABLISSEMENT';
        Libelle := 'ET_LIBELLE';
        Filtre := '';
        Valeur := '';
      end;
    3: // Régimes de facturation
      begin
        Code := 'CC_CODE';
        Libelle := 'CC_LIBELLE';
        Filtre := 'CC_TYPE';
        Valeur := 'RTV';
      end;
    4: // Mode de règlement tiers
      begin
        Code := 'MR_MODEREGLE';
        Libelle := 'MR_LIBELLE';
        Filtre := '';
        Valeur := '';
      end;
    5: // Tiers des transferts
      begin
        Code := 'T_TIERS';
        Libelle := 'T_TIERS';
        Filtre := '';
        Valeur := '';
      end;
    6: // Période de base pour les tarifs
      begin
        Code := 'GFP_CODEPERIODE';
        Libelle := 'GFP_LIBELLE';
        Filtre := '';
        Valeur := '';
      end;
    7: // Photo pour les fiches
      begin
        Code := 'CC_CODE';
        Libelle := 'CC_LIBELLE';
        Filtre := 'CC_TYPE';
        Valeur := 'GEB';
      end;
  else Exit;
  end;
  Combo.Items.Clear;
  Combo.Values.Clear;
  Combo.DataType := '';
  for Ind := 0 to TOBData.Detail.Count - 1 do
  begin
    TOBL := TOBData.Detail[Ind];
    if Filtre <> '' then Ok := (TOBL.GetValue(Filtre) = Valeur)
    else Ok := True;
    if Ok then
    begin
      Combo.Items.Add(TOBL.GetValue(Libelle));
      Combo.Values.Add(TOBL.GetValue(Code));
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FusionneTOB : Fusionne deux TOB
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.FusionneTOB(LaTOB, TOBData: TOB; NomTable, Sql: string; Cle: array of string; Valeur: string);
var QQ: TQuery;
  TOBO, TOBD: TOB;
  Ind, ii: Integer;
  ValeurCle: array of Variant;
begin
  SetLength(ValeurCle, High(Cle) + 1);
  QQ := OpenSQL(Sql, True);
  if not QQ.Eof then
  begin
    TOBData.LoadDetailDB(NomTable, '', '', QQ, False, True);
    for Ind := 0 to LaTOB.Detail.Count - 1 do
    begin
      TOBO := LaTOB.Detail[Ind];
      for ii := Low(Cle) to High(Cle) do ValeurCle[ii] := TOBO.GetValue(Cle[ii]);
      TOBD := TOBData.FindFirst(Cle, ValeurCle, False);
      if TOBD <> nil then
      begin
        TOBD.PutValue(Valeur, TOBO.GetValue(Valeur));
        TOBD.SetAllModifie(True);
      end;
    end;
  end else
  begin
    TOBData.Dupliquer(LaTob, True, True);
  end;
  Ferme(QQ);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ImportTOB : Importation d'une TOB
///////////////////////////////////////////////////////////////////////////////////////

function TFImportCais.ImportTOB(LaTOB: TOB): Integer;
var sSql: string;
begin
  if Pos('TTREGIMETVA', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Régimes de facturation', [fsItalic], False, True);
    if TOBReg <> nil then TOBReg.Free;
    TOBReg := TOB.Create('Régimes', nil, -1);
    TOBReg.Dupliquer(LaTob, True, True);
    TOBToCombo(3, TOBReg, SO_REGIMEDEFAUT);
  end else
  if Pos('TTTVA', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Codes TVA', [fsItalic], False, True);
    if TOBTva <> nil then TOBTva.Free;
    TOBTva := TOB.Create('Codes TVA', nil, -1);
    TOBTva.Dupliquer(LaTob, True, True);
  end else
  if Pos('GCEMPLOIBLOB', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Photo pour les fiches', [fsItalic], False, True);
    if TOBPhoto <> nil then TOBPhoto.Free;
    TOBPhoto := TOB.Create('Photos', nil, -1);
    TOBPhoto.Dupliquer(LaTob, True, True);
    TOBToCombo(7, TOBPhoto, SO_GCPHOTOFICHE);
  end else
  if Pos('DEVISE', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Devises', [fsItalic], False, True);
    if TOBDevise <> nil then TOBDevise.Free;
    TOBDevise := TOB.Create('Devises', nil, -1);
    TOBDevise.Dupliquer(LaTob, True, True);
    TOBToCombo(1, TOBDevise, SO_DEVISEPRINC);
  end else
  if Pos('MODEREGL', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Modes de règlement', [fsItalic], False, True);
    if TOBModeRegl <> nil then TOBModeRegl.Free;
    TOBModeRegl := TOB.Create('Règlements', nil, -1);
    TOBModeRegl.Dupliquer(LaTob, True, True);
    TOBToCombo(4, TOBModeRegl, SO_GCMODEREGLEDEFAUT);
  end else
  if Pos('MODEPAIE', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Modes de paiement', [fsItalic], False, True);
    if TOBDModePaie <> nil then TOBDModePaie.Free;
    TOBDModePaie := TOB.Create('Paiements', nil, -1);
    TOBDModePaie.Dupliquer(LaTob, True, True);
  end else
  if Pos('TIERS', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Tiers', [fsItalic], False, True);
    if TOBTiers <> nil then TOBTiers.Free;
    TOBTiers := TOB.Create('Tiers', nil, -1);
    TOBTiers.Dupliquer(LaTob, True, True);
    TOBToCombo(5, TOBTiers, SO_GCTIERSDEFAUT);
  end else
  if Pos('EXERCICE', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Exercice', [fsItalic], False, True);
    if TOBExercice <> nil then TOBExercice.Free;
    TOBExercice := TOB.Create('Exercice', nil, -1);
    TOBExercice.Dupliquer(LaTob, True, True);
  end else
  if Pos('ETABLISS', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Etablissements', [fsItalic], False, True);
    if TOBEtb <> nil then TOBEtb.Free;
    TOBEtb := TOB.Create('Etablissements', nil, -1);
    TOBEtb.Dupliquer(LaTob, True, True);
    TOBToCombo(2, TOBEtb, SO_ETABLISDEFAUT);
  end else
  if Pos('DEPOTS', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Dépôts', [fsItalic], False, True);
    if TOBDepot <> nil then TOBDepot.Free;
    TOBDepot := TOB.Create('Dépôts', nil, -1);
    TOBDepot.Dupliquer(LaTob, True, True);
  end else
  if Pos('PARCAISSE', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Caisses', [fsItalic], False, True);
    if TOBCaisse <> nil then TOBCaisse.Free;
    TOBCaisse := TOB.Create('Caisses', nil, -1);
    TOBCaisse.Dupliquer(LaTob, True, True);
  end else
  if Pos('SOUCHE', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Souches', [fsItalic], False, True);
    if TOBSouche <> nil then TOBSouche.Free;
    TOBSouche := TOB.Create('Souches', nil, -1);
    TOBSouche.Dupliquer(LaTob, True, True);
  end else
  if Pos('TARIFPER', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Périodes d''application des tarifs', [fsItalic], False, True);
    if TOBPer <> nil then TOBPer.Free;
    TOBPer := TOB.Create('Périodes', nil, -1);
    TOBPer.Dupliquer(LaTob, True, True);
    TOBToCombo(6, TOBPer, SO_GCPERBASETARIF);
  end else
  if Pos('STOXPARMS', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Paramètres TOX', [fsItalic], False, True);
    if TOBToxParms <> nil then TOBToxParms.Free;
    TOBToxParms := TOB.Create('Paramètres', nil, -1);
    TOBToxParms.Dupliquer(LaTob, True, True);
  end else
  if Pos('STOXVARS', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Variables TOX', [fsItalic], False, True);
    if TOBToxvars <> nil then TOBToxvars.Free;
    TOBToxvars := TOB.Create('Variables', nil, -1);
    TOBToxvars.Dupliquer(LaTob, True, True);
  end else
  if Pos('STOXCHRONO',LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Chonos', [fsItalic], False, True);
    if TOBToxChronos <> nil then TOBToxChronos.Free;
    TOBToxChronos := TOB.Create('Chronos', nil, -1);
    TOBToxChronos.Dupliquer(LaTob, True, True);
  end else
  if Pos('STOXSITES', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Sites', [fsItalic], False, True);
    if TOBSite <> nil then TOBSite.Free;
    TOBSite := TOB.Create('Sites', nil, -1);
    TOBSite.Dupliquer(LaTob, True, True);
  end else
  if Pos('STOXINFOCOMP', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Infos complémentaires des sites', [fsItalic], False, True);
    if TOBInfo <> nil then TOBInfo.Free;
    TOBInfo := TOB.Create('Infos', nil, -1);
    TOBInfo.Dupliquer(LaTob, True, True);
  end else
  if Pos('STOXQUERYS', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Requêtes', [fsItalic], False, True);
    if TOBToxQry <> nil then TOBToxQry.Free;
    TOBToxQry := TOB.Create('Requêtes', nil, -1);
    TOBToxQry.Dupliquer(LaTob, True, True);
  end else
  if Pos('STOXEVENTS', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Evénements', [fsItalic], False, True);
    if TOBToxEvt <> nil then TOBToxEvt.Free;
    TOBToxEvt := TOB.Create('Evénements', nil, -1);
    TOBToxEvt.Dupliquer(LaTob, True, True);
  end else
  if Pos('USERGRP', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Groupes d''utilisateurs', [fsItalic], False, True);
    if TOBUserGrp <> nil then TOBUserGrp.Free;
    TOBUserGrp := TOB.Create('Groupes utilisateurs', nil, -1);
    TOBUserGrp.Dupliquer(LaTob, True, True);
  end else
  if Pos('MENU', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Confidentialité', [fsItalic], False, True);
    if TOBMenu <> nil then TOBMenu.Free;
    TOBMenu := TOB.Create('Confidentialité', nil, -1);
    sSql := 'select * from MENU where MN_1 in (26, 107, 108, 109, 111, 112)';
    FusionneTOB(LaTOB, TOBMenu, 'MENU', sSql, ['MN_1', 'MN_2', 'MN_3', 'MN_4'], 'MN_ACCESGRP');
  end else
  if Pos('UTILISAT', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Utilisateurs', [fsItalic], False, True);
    if TOBUtil <> nil then TOBUtil.Free;
    TOBUtil := TOB.Create('Utilisateurs', nil, -1);
    TOBUtil.Dupliquer(LaTob, True, True);
  end else
  if Pos('PARAMSOC', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Paramètres société', [fsItalic], False, True);
    if TOBParam <> nil then TOBParam.Free;
    TOBParam := TOB.Create('Paramètres', nil, -1);
    sSql := 'select * from PARAMSOC where (SOC_TREE like "001;001;%" or SOC_TREE like "001;003;%")'
      + ' and SOC_TREE not like "%;000;"';
    FusionneTOB(LaTOB, TOBParam, 'PARAMSOC', sSql, ['SOC_NOM'], 'SOC_DATA');
  end else
    // Ajout pour PCP
  if Pos('SITEPIECE', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Exceptions sites/pièces', [fsItalic], False, True);
    if TOBSitePiece <> nil then TOBSitePiece.Free;
    TOBSitePiece := TOB.Create('Sites_Pièces', nil, -1);
    TOBSitePiece.Dupliquer(LaTob, True, True);
  end else
  if Pos('PARPIECE', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Paramètres pièces', [fsItalic], False, True);
    if TOBParPiece <> nil then TOBParPiece.Free;
    TOBParPiece := TOB.Create('Param Pièces', nil, -1);
    TOBParPiece.Dupliquer(LaTob, True, True);
  end else
  if Pos('CPPROFILUSERC', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Paramètres utilisateurs', [fsItalic], False, True);
    if TOBCPProfilUserC <> nil then TOBCPProfilUserC.Free;
    TOBCPProfilUserC := TOB.Create('Param User', nil, -1);
    TOBCPProfilUserC.Dupliquer(LaTob, True, True);
  end else
  if Pos('GENERAUX', LaTOB.NomTable) > 0 then
  begin
    AfficheMemo('Comptes généraux', [fsItalic], False, True);
    if TOBGeneraux <> nil then TOBGeneraux.Free;
    TOBGeneraux := TOB.Create('Comptes généraux', nil, -1);
    TOBGeneraux.Dupliquer(LaTob, True, True);
  end;
  LaTob.Free;
  Result := 0;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AfficheTOB : Affichage des TOB
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.AfficheTOB;
var TOBL: TOB;
begin
  // Comptabilité - Coordonnées
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_SOCIETE'], False);
  if TOBL <> nil then SO_SOCIETE.Text := TOBL.GetValue('SOC_DATA');
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_LIBELLE'], False);
  if TOBL <> nil then SO_LIBELLE.Text := TOBL.GetValue('SOC_DATA');
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_ADRESSE1'], False);
  if TOBL <> nil then SO_ADRESSE1.Text := TOBL.GetValue('SOC_DATA');
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_ADRESSE2'], False);
  if TOBL <> nil then SO_ADRESSE2.Text := TOBL.GetValue('SOC_DATA');
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_ADRESSE3'], False);
  if TOBL <> nil then SO_ADRESSE3.Text := TOBL.GetValue('SOC_DATA');
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_CODEPOSTAL'], False);
  if TOBL <> nil then SO_CODEPOSTAL.Text := TOBL.GetValue('SOC_DATA');
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_VILLE'], False);
  if TOBL <> nil then SO_VILLE.Text := TOBL.GetValue('SOC_DATA');
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_PAYS'], False);
  if TOBL <> nil then SO_PAYS.Value := TOBL.GetValue('SOC_DATA');
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_TELEPHONE'], False);
  if TOBL <> nil then SO_TELEPHONE.Text := TOBL.GetValue('SOC_DATA');
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_FAX'], False);
  if TOBL <> nil then SO_FAX.Text := TOBL.GetValue('SOC_DATA');
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_MAIL'], False);
  if TOBL <> nil then SO_MAIL.Text := TOBL.GetValue('SOC_DATA');
  // Comptabilité - Comptables
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_ETABLISDEFAUT'], False);
  if TOBL <> nil then
  begin
    SO_ETABLISDEFAUT.Value := TOBL.GetValue('SOC_DATA');
    EtabOrg := SO_ETABLISDEFAUT.Value;
  end;
  // Comptabilité - Euro
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_TENUEEURO'], False);
  if TOBL <> nil then
  begin
    SO_TENUEEURO.Checked := TOBL.GetValue('SOC_DATA') = 'X';
  end;

  // MODIF LM 28/08/02 pour gestion EURO PGI
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_DEVISEPRINC'], False);
  if TOBL <> nil then
  begin
    if SO_TENUEEURO.Checked then
    begin
      SO_DEVISEPRINC.Value := 'EUR';
      SO_FONGIBLE.Value := TOBL.GetValue('SOC_DATA');
      SO_TAUXEURO.Value := 1.00000;
    end else
    begin
      SO_DEVISEPRINC.Value := TOBL.GetValue('SOC_DATA');
      SO_FONGIBLE.Value := 'EUR';
      TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_TAUXEURO'], False);
      if TOBL <> nil then SO_TAUXEURO.Value := TOBL.GetValue('SOC_DATA');
    end;
    //Deviseorg := SO_FONGIBLE.Value ;
  end;
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_DECVALEUR'], False);
  if TOBL <> nil then SO_DECVALEUR.Value := TOBL.GetValue('SOC_DATA');

  // Gestion commerciale - Paramètres par défaut
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_REGIMEDEFAUT'], False);
  if TOBL <> nil then SO_REGIMEDEFAUT.Value := TOBL.GetValue('SOC_DATA');
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_GCPERBASETARIF'], False);
  if TOBL <> nil then SO_GCPERBASETARIF.Value := TOBL.GetValue('SOC_DATA');
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_GCMODEREGLEDEFAUT'], False);
  if TOBL <> nil then SO_GCMODEREGLEDEFAUT.Value := TOBL.GetValue('SOC_DATA');
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_GCTIERSDEFAUT'], False);
  if TOBL <> nil then SO_GCTIERSDEFAUT.Value := TOBL.GetValue('SOC_DATA');
  TOBL := TOBParam.FindFirst(['SOC_NOM'], ['SO_GCPHOTOFICHE'], False);
  if TOBL <> nil then SO_GCPHOTOFICHE.Value := TOBL.GetValue('SOC_DATA');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AfficheMemo : affiche un texte dans le mémo.
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.AfficheMemo(Texte: string; Style: TFontStyles; Saut, Traduire: Boolean);
var Stg: string;
begin
  if (P.ActivePage.Name = 'PAGE3') and (CPTRDURECUP <> nil) then
  begin
    if Saut then CPTRDURECUP.lines.Append('');
    CPTRDURECUP.SelAttributes.Style := Style;
    if Traduire then Stg := TraduireMemoire(Texte) else Stg := Texte;
    CPTRDURECUP.lines.Append(Stg);
    if CPTRDURECUP.CanFocus then CPTRDURECUP.SetFocus;
  end else
    if (P.ActivePage.Name = 'PAGE7') and (CPTRENDU <> nil) then
  begin
    if Saut then CPTRENDU.lines.Append('');
    CPTRENDU.SelAttributes.Style := Style;
    if Traduire then Stg := TraduireMemoire(Texte) else Stg := Texte;
    CPTRENDU.lines.Append(Stg);
    if CPTRENDU.CanFocus then CPTRENDU.SetFocus;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ActiveBtn : active les boutons de l'assistant
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.ActiveBtn(Active: Boolean);
begin
  bAide.Enabled := Active;
  bAnnuler.Enabled := Active;
  bPrecedent.Enabled := Active;
  bSuivant.Enabled := Active;
  if (CurrentPage = PagesCount) then bFin.Enabled := Active;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SauveCaisse : mise à jour de la base
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.SauveCaisse;
begin
  // Mise à jour des tables
  MAJTable(TOBReg, False);
  MAJTable(TOBTva, False);
  MAJTable(TOBPhoto, False);
  MAJTable(TOBModeRegl, True);
  MAJTable(TOBDModePaie, True);
  MAJTable(TOBTiers, True);
  MAJTable(TOBEtb, True);
  MAJTable(TOBDepot, True);
  MAJTable(TOBPer, True);
  MAJTable(TOBToxvars, True);
  MAJTable(TOBToxChronos, True);
  MAJTable(TOBToxQry, True);
  MAJTable(TOBToxEvt, True);
  MAJTable(TOBInfo, True);
  // Mise à jour des sites
  MAJTable(TOBSite, True);
  MAJTOXSite;
  MAJTable(TOBUserGrp, True);
  MAJTable(TOBMenu, False);
  MAJTable(TOBUtil, True);
  // Mise à jour des paramètres sociétés
  MAJParamSoc;
  // Mise à jour des souches
  if EtabOrg = SO_ETABLISDEFAUT.Value then MAJTable(TOBSouche, False)
  else MAJSouche;
  // Mise à jour des devises
  MAJTable(TOBDevise, True);
  MAJDevise;

  // Mise à jour des caisses
  MAJTable(TOBCaisse, True);
  MAJCaisse;
  // Rechargement des tablettes en mémoire
  VideTablettes;
  AvertirTable('TTDEVISE');
  AvertirTable('TTDEVISETOUTES');
  // Demande la bascule Euro si besoin
  if not GetParamSoc('SO_TENUEEURO') then AvertirEuro;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SauvePCP : mise à jour de la base
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.SauvePCP;
begin
  // Mise à jour des tables
  MAJTable(TOBExercice, True);
  MAJTable(TOBEtb, True);
  MAJTable(TOBDepot, True);
  MAJTable(TOBSitePiece, True); // Mise à jour des exceptions sites/pièces
  MAJTable(TOBParPiece, True);
  MAJTable(TOBToxParms, True);
  MAJTable(TOBToxvars, True);
  MAJTable(TOBToxChronos, True);
  MAJTable(TOBToxQry, True);
  MAJTable(TOBToxEvt, True);
  MAJTable(TOBInfo, True);
  // Mise à jour des sites
  MAJTable(TOBSite, True);
  MAJTOXSite;
  MAJTable(TOBUserGrp, True);
  MAJTable(TOBMenu, False);
  MAJTable(TOBUtil, True);
  MAJTable(TOBCPProfilUserC, True);
  MAJTable(TOBGeneraux, True);
  // Mise à jour des paramètres sociétés
  MAJParamSoc;
  // Mise à jour des souches
  if EtabOrg = SO_ETABLISDEFAUT.Value then MAJTable(TOBSouche, False)
  else MAJSouche;
  // Mise à jour des devises
  if Deviseorg = SO_DEVISEPRINC.Value then MAJTable(TOBDevise, True)
  else MAJDevise;
  // Rechargement des tablettes en mémoire
  VideTablettes;
  AvertirTable('TTDEVISE');
  AvertirTable('TTDEVISETOUTES');
  // Demande la bascule Euro si besoin
  if not GetParamSoc('SO_TENUEEURO') then AvertirEuro;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AvertirEuro : averti l'utilisateur de lancer la bascule Euro
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.AvertirEuro;
var TOBL: TOB;
  Stg: string;
  Ind: Integer;
  StgLst: TStringList;
begin
  TOBL := TOBSite.FindFirst(['SSI_STYPESITE'], ['002'], False);
  if TOBL = nil then Exit;
  Stg := TOBL.GetValue('SSI_CODESITE');
  TOBL := TOBInfo.FindFirst(['SIC_CODESITE'], [Stg], False);
  if TOBL = nil then Exit;
  StgLst := TStringList.Create;
  StgLst.Text := TOBL.GetValue('SIC_INFO');
  for Ind := 0 to StgLst.Count - 1 do if StgLst.Strings[Ind] = 'BASCULEEURO;B;X' then
    begin
      // Vous devez lancer la bascule Euro.
      Stg := Msg.Mess[9];
      ReadTokenST(Stg);
      ReadTokenST(Stg);
      AfficheMemo(ReadTokenST(Stg), [fsBold, fsItalic], True, False);
      Msg.Execute(9, Caption, '');
      // On ne lance pas la TOX
      if LANCETOX.Checked then LANCETOX.Checked := False;
      Break;
    end;
  StgLst.Free;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MAJTable : mise à jour d'une table
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.MAJTable(TOBData: TOB; EraseTable: Boolean);
var sSql: string;
  TOBL: TOB;
  iTob: integer;
begin
  if TOBData = nil then Exit;
  if TOBData.NomTable = '' then Exit;
  if TOBData.Detail.Count < 1 then Exit;
  AfficheMemo(' * ' + TOBData.NomTable, [], False, False);
  TOBData.SetAllModifie(True);
  if (EraseTable) and (GACTIONTABLE.ItemIndex = 0) and (TOBData.Detail.Count > 0) then
  begin
    TOBL := TOBData.Detail[0];
    sSql := 'delete ' + TOBL.NomTable;
    if TOBData.NomTable = 'STOXQUERYS' then
    begin
      if ctxFO in V_PGI.PGIContexte then sSql := sSql + ' where SQE_PREDEFINI<>"CEG"'
      else sSql := sSql + ' where SQE_PREDEFINI="CEG"';
    end
    else if TOBData.NomTable = 'STOXEVENTS' then sSql := sSql + ' where SEV_TYPETRT="001"';
    ExecuteSQL(sSql);

    if pos(' UTILISAT', TOBData.NomTable) > 0 then
    begin
      // Depuis AGL550b, la table UTILISAT ne peut pas être maj par InsertDB : écrasement US_UTILISATEUR !
      for iTob := 0 to TOBData.Detail.Count - 1 do
        ExecuteSql(MakeInsertSQL(TOBData.Detail[iTob]));
    end
    else TOBData.InsertDB(nil);
  end else TOBData.InsertOrUpdateDB;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MAJParamSoc : mise à jour des paramètres société
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.MAJParamSoc;
{$IFNDEF NOMADE}
var Cpt: Integer;
{$ENDIF}

  procedure MajParam(NomParam: string; ValeurParam: Variant);
  var TOBL: TOB;
  begin
    TOBL := TOBParam.FindFirst(['SOC_NOM'], [NomParam], False);
    if TOBL <> nil then
    begin
      TOBL.PutValue('SOC_DATA', ValeurParam);
      TOBL.SetAllModifie(True);
    end;
  end;

begin
  // Chargement des paramètres société
  AfficheMemo('Mise à jour des paramètres société', [], False, True);
  // Comptabilité - Coordonnées
  AfficheMemo('  * Comptabilité - Coordonnées', [], False, True);
  MajParam('SO_SOCIETE', SO_SOCIETE.Text);
  MajParam('SO_LIBELLE', SO_LIBELLE.Text);
  MajParam('SO_ADRESSE1', SO_ADRESSE1.Text);
  MajParam('SO_ADRESSE2', SO_ADRESSE2.Text);
  MajParam('SO_ADRESSE3', SO_ADRESSE3.Text);
  MajParam('SO_CODEPOSTAL', SO_CODEPOSTAL.Text);
  MajParam('SO_VILLE', SO_VILLE.Text);
  MajParam('SO_PAYS', SO_PAYS.Value);
  MajParam('SO_TELEPHONE', SO_TELEPHONE.Text);
  MajParam('SO_FAX', SO_FAX.Text);
  MajParam('SO_MAIL', SO_MAIL.Text);
  // Comptabilité - Comptables
  AfficheMemo('  * Comptabilité - Comptables', [], False, True);
  MajParam('SO_ETABLISDEFAUT', SO_ETABLISDEFAUT.Value);
  // Comptabilité - Euro
  AfficheMemo('  * Comptabilité - Euro', [], False, True);
  MajParam('SO_DEVISEPRINC', SO_DEVISEPRINC.Value);
  MajParam('SO_DECVALEUR', SO_DECVALEUR.Text);
  MajParam('SO_TAUXEURO', SO_TAUXEURO.Text);
  if SO_TENUEEURO.Checked then MajParam('SO_TENUEEURO', 'X')
  else MajParam('SO_TENUEEURO', '-');
  // Comptabilité - Divers
  AfficheMemo('  * Comptabilité - Divers', [], False, True);
  MajParam('SO_DECPRIX', SO_DECVALEUR.Text);
  //MajParam('SO_DECQTE', 0) ;
  MajParam('SO_GCMODEREGLEDEFAUT', SO_GCMODEREGLEDEFAUT.Value);
  // Gestion commerciale - Paramètres par défaut
  AfficheMemo('  * Gestion commerciale - Paramètres par défaut', [], False, True);
  MajParam('SO_REGIMEDEFAUT', SO_REGIMEDEFAUT.Value);
  MajParam('SO_GCTIERSDEFAUT', SO_GCTIERSDEFAUT.Text);
  MajParam('SO_GCTIERSPAYS', SO_PAYS.Value);
  MajParam('SO_GCDEPOTDEFAUT', SO_ETABLISDEFAUT.Value);
  MajParam('SO_GCPERBASETARIF', SO_GCPERBASETARIF.Value);
  // Gestion commerciale - Préférences
  MajParam('SO_GCPHOTOFICHE', SO_GCPHOTOFICHE.Value);
  // Gestion commerciale - Clients Fournisseurs
  {$IFDEF NOMADE}
  MajParam('SO_GCCOMPTEURTIERS', 0);
  MajParam('SO_GCNUMTIERSAUTO', 'X');
  {$ELSE}
  if EtabOrg <> SO_ETABLISDEFAUT.Value then
  begin
    AfficheMemo('  * Gestion commerciale - Clients Fournisseurs', [], False, True);
    Cpt := StrToint(SO_ETABLISDEFAUT.Value + '00000');
    MajParam('SO_GCCOMPTEURTIERS', Cpt);
    MajParam('SO_GCNUMTIERSAUTO', 'X');
  end;
  {$ENDIF}
  if ctxFO in V_PGI.PGIContexte then
  begin
    // Gestion commerciale - Front Office
    AfficheMemo('  * Gestion commerciale - Front Office', [], False, True);
    MajParam('SO_GCFOCAISREFTOX', GPK_CAISSE.Text);
  end;
  // Mise à jour de la table
  TOBParam.UpdateDB(True);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MAJSouche : mise à jour des souches
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.MAJSouche;
var Stg: string;
begin
  AfficheMemo('Mise à jour des souches', [], False, True);
  Stg := 'update SOUCHE set SH_NUMDEPART=' + SO_ETABLISDEFAUT.Value + '000000';
  ExecuteSQL(Stg);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MAJDevise : mise à jour des devises
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.MAJDevise;
var Stg: string;
  QQ: TQuery;
begin
  AfficheMemo('Mise à jour des devises', [], false, True);
  Stg := 'update DEVISE set D_FONGIBLE="-"';
  ExecuteSQL(Stg);
  Stg := 'select D_MONNAIEIN from DEVISE where D_DEVISE="' + SO_FONGIBLE.Value + '"';
  QQ := OpenSQL(Stg, True);
  if not (QQ.EOF) and (QQ.Fields[0].AsString = 'X') then
  begin
    Stg := 'update DEVISE set D_FONGIBLE="X" where D_DEVISE="' + SO_FONGIBLE.Value + '"';
    ExecuteSQL(Stg);
  end;
  Ferme(QQ);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MAJCaisse : mise à jour des caisses
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.MAJCaisse;
var Stg, Heure1, Heure2: string;
  TOBL: TOB;
begin
  //AfficheMemo('Mise à jour des caisses', [], false, True) ;
  Stg := 'update PARCAISSE set GPK_FERME="X" where GPK_CAISSE<>"' + GPK_CAISSE.Text + '"';
  ExecuteSQL(Stg);
  Stg := 'update PARCAISSE set GPK_FERME="-" where GPK_CAISSE="' + GPK_CAISSE.Text + '"';
  ExecuteSQL(Stg);
  // Enregistrement de la configuration du ToxServeur de la caisse de référence
  TOBL := TOBCaisse.FindFirst(['GPK_CAISSE'], [GPK_CAISSE.Text], False);
  if TOBL <> nil then
  begin
    Heure1 := TOBL.GetValue('GPK_TOXAPPEL1');
    Heure2 := TOBL.GetValue('GPK_TOXAPPEL2');
    if Heure1 <> '' then FOToxServeurMajHeure(Heure1, Heure2);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MAJTOXSite : mise à jour du fichier des sites du ToxServeur
///////////////////////////////////////////////////////////////////////////////////////

procedure TFImportCais.MAJTOXSite;
begin
  AfficheMemo('Génération du fichier pour le ToxServeur', [], false, True);
  AglMakeToxIni;
end;

end.
