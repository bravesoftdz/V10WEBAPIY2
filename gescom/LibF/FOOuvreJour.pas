{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 30/10/2003
Description .. : Assistant d'ouverture d'une journée de vente du Front Office
Mots clefs ... : FO
*****************************************************************}
unit FOOuvreJour;

interface

uses
  Forms, assist, ImgList, Controls, StdCtrls, ComCtrls, Hctrls, Mask,
  Grids, HPanel, HSysMenu, hmsgbox, ExtCtrls, HTB97, Classes,
  Windows, SysUtils, Graphics, Hent1, HDebug,
  {$IFNDEF EAGLCLIENT}
  dbtables,
  {$ENDIF}
  UTOB, kb_Ecran;

procedure FOOuvreJournee(PRien: THPanel);

type
  TFOuvertJour = class(TFAssist)
    PChoixCaisse: TTabSheet;
    GJC_CAISSE: THValComboBox;
    TGPK_ETABLISSEMENT: THLabel;
    LGPK_ETABLISSEMENT: TEdit;
    TGPK_DEPOT: THLabel;
    LGPK_DEPOT: TEdit;
    GPK_DEPOT: THValComboBox;
    GPK_ETABLISSEMENT: THValComboBox;
    TGJC_JOURNEE: THLabel;
    GJC_JOURNEE: THCritMaskEdit;
    PfondCaisse: TTabSheet;
    TitreChoix: THLabel;
    TitreFond: THLabel;
    TGJC_CAISSE: THLabel;
    PResume: TTabSheet;
    CompteRendu: TRichEdit;
    ListeImages: TImageList;
    LGJC_CAISSE: TEdit;
    TGJC_VENDOUV: THLabel;
    GJC_VENDOUV: THCritMaskEdit;
    GS: THGrid;
    PPave: TPanel;
    BLEFT: TToolbarButton97;
    BRIGHT: TToolbarButton97;
    BUP: TToolbarButton97;
    BDOWN: TToolbarButton97;
    PPIECBIL: THPanel;
    PCUMUL: THPanel;
    PBILLET: THPanel;
    PPIECE: THPanel;
    PCUMULBILLET: THPanel;
    PCUMULPIECE: THPanel;
    CUMMTBILLET: THNumEdit;
    CUMMTPIECE: THNumEdit;
    CUMMONTANT: THNumEdit;
    GBILLET: THGrid;
    GPIECE: THGrid;
    BENTETE: TToolbarButton97;
    BBILLET: TToolbarButton97;
    procedure GJC_CAISSEChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure bFinClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure PChange(Sender: TObject); override;
    procedure GJC_VENDOUVEnter(Sender: TObject);
    procedure GJC_VENDOUVDblClick(Sender: TObject);
    procedure GJC_VENDOUVElipsisClick(Sender: TObject);
    procedure GJC_JOURNEEExit(Sender: TObject);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSEnter(Sender: TObject);
    procedure GSExit(Sender: TObject);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GBilletEnter(Sender: TObject);
    procedure GBilletExit(Sender: TObject);
    procedure GBilletCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GBilletCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GBilletColumnWidthsChanged(Sender: TObject);
    procedure GPieceEnter(Sender: TObject);
    procedure GPieceExit(Sender: TObject);
    procedure GPieceCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GPieceCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GPieceColumnWidthsChanged(Sender: TObject);
    procedure BEnteteClick(Sender: TObject);
    procedure BBilletClick(Sender: TObject);
    procedure BLEFTClick(Sender: TObject);
    procedure BRIGHTClick(Sender: TObject);
    procedure BUPClick(Sender: TObject);
    procedure BDOWNClick(Sender: TObject);
    procedure bAnnulerClick(Sender: TObject);
  private
    { Déclarations privées }
    TOBCaisse: TOB; // TOB de la caisse choisie
    TOBJCaisse: TOB; // TOB de la journée de la caisse
    TOBFDCais: TOB; // TOB du fond de caisse
    TOBBillet: TOB; // TOB des billets (montants par type de billets)
    TOBPiece: TOB; // TOB des pièces (montants par type de pièces)
    TOBBilMdp: TOB; // TOB des billets du mode de paiement courant
    TOBPieMdp: TOB; // TOB des pièces du mode de paiement courant
    NbPage: Integer; // Nombre de pages actives
    NoPage: Integer; // Numéro de la page courante
    PnlBoutons: TClavierEcran; // pavé tactile
    BoutonEncours: Boolean; // traitement d'un bouton en cours
    PiecBilEnCours: Boolean; // saisie des pièces et billets en cours
    ApplicEncours: Boolean; // recopie en cours du total (détail ou pièces/billet) sur la grille du fond de caisse
    StPiecBilCur: string; // valeur de la cellule courante de la grille des pièces ou billets
    procedure ChargeCaisse;
    function SauveCaisse: boolean;
    procedure EnregistreSaisie;
    procedure GenereFondCaisse;
    procedure ImprimeTicket;
    procedure ChargeCompteRendu;
    function VerifieVendOuv: Boolean;
    function VerifieDateOuv: Boolean;
    function VerifiePieceBillet(ModePaie: string; TotSaisi: Double): Boolean;
    function VerifieFdCaisse: Boolean;
    function VerifiePageActive(TT: TTabSheet): Boolean;
    procedure ChoixVendeurPave;
    procedure ActivebFin;
    procedure GPiecBilColAligns(Grille: THGrid);
    procedure GPiecBilPutGrid(Grille: THGrid; TOBPB: TOB);
    procedure ChargeFondCaisse;
    function GetTOBFDCais(ARow: Integer): TOB;
    procedure GSColAligns;
    procedure FormateZoneSaisie(Grille: THGrid; ACol, ARow: Longint);
    procedure CalculMontantPieceBillet(Grille: THGrid; TOBParent: TOB; ACol, ARow: Longint; var Cancel: Boolean);
    function CalculTotalPieceBillet(TOBBil, TOBPie: TOB; Affiche: Boolean): Double;
    procedure SaisieClavierEcran(Concept, Code, Extra: string; Qte, Prix: Double);
    procedure BoutonCalculetteClick(Val: string);
    procedure SaisieCalculette(Val: Double);
    procedure AppliqueTotal;
    procedure GotoPied;
    procedure GotoEntete;
    procedure ReprisePieceBillet(ModePaie, Texte: string);
  public
    { Déclarations publiques }
    function PreviousPage: TTabSheet; override;
    function NextPage: TTabSheet; override;
  end;

implementation
uses
  {$IFDEF TOXCLIENT}
  uToxConf, uToxConst,
  {$ENDIF}
  EntGC, SaisUtil, TickUtilFO, MFOCTRLECAISSE_TOF, FOUtil, FODefi;

// définition des colonnes de la grille de saisie du fond de caisse
const
  GRD_MODEPAIE = 0;
  GRD_DEVISE = 1;
  GRD_MONTANT = 2;

  // définition des colonnes des grilles des billets et pièces
const
  COLB_LIBELLE = 0;
  COLB_QUANTITE = 1;
  COLB_MONTANT = 2;

  {$R *.DFM}

  {***********A.G.L.Privé.*****************************************
  Auteur  ...... : N. ACHINO
  Créé le ...... : 23/07/2001
  Modifié le ... : 23/07/2001
  Description .. : Ouverture d'une journée de vente du Front Office
  Mots clefs ... : FO
  *****************************************************************}

procedure FOOuvreJournee(PRien: THPanel);
var X: TFOuvertJour;
begin
  FOAlerteAnnoncesNonValidees(FOGetParamCaisse('GPK_ETABLISSEMENT'), FOGetParamCaisse('GPK_DEPOT'));
  PRien.InsideTitle.Caption := TraduireMemoire('Ouverture de la journée de vente');
  X := TFOuvertJour.Create(Application);
  try
    X.ShowModal;
  finally
    X.Free;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeCaisse : charge la TOB de la caisse choisie
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.ChargeCaisse;
begin
  // Lecture du paramétrage de la caisse
  TOBCaisse.SelectDB('"' + GJC_CAISSE.Value + '"', nil);
  // Affichage des données complémentaires
  LGJC_CAISSE.Text := GJC_CAISSE.Text;
  GPK_ETABLISSEMENT.Value := TOBCaisse.GetValue('GPK_ETABLISSEMENT');
  LGPK_ETABLISSEMENT.Text := GPK_ETABLISSEMENT.Text;
  GPK_DEPOT.Value := TOBCaisse.GetValue('GPK_DEPOT');
  LGPK_DEPOT.Text := GPK_DEPOT.Text;
  // Initialisation de la TOB de la journée de caisse
  TOBJCaisse.InitValeurs;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SauveCaisse : enregistre la TOB de la caisse choisie
///////////////////////////////////////////////////////////////////////////////////////

function TFOuvertJour.SauveCaisse: boolean;
var nbOuv, numZ, Ind: Integer;
  TOBL: TOB;
  io: TIOErr;
  Caisse: string;
begin
  Result := False;
  Caisse := TOBCaisse.GetValue('GPK_CAISSE');
  // Met à jour la journée de la caisse
  NumZ := FOGetNumZCaisse(Caisse) + 1;
  if TOBJCaisse.GetValue('GJC_NUMZCAISSE') = 0 then
  begin
    TOBJCaisse.PutValue('GJC_NUMZCAISSE', NumZ);
  end;
  TOBJcaisse.PutValue('GJC_CAISSE', Caisse);
  TOBJcaisse.PutValue('GJC_ETAT', ETATJOURCAISSE[2]);
  //TOBJcaisse.PutValue('GJC_DATEOUV', NowH) ;
  TOBJcaisse.PutValue('GJC_DATEOUV', StrToDate(GJC_JOURNEE.Text));
  TOBJcaisse.PutValue('GJC_HEUREOUV', TimeToStr(Time));
  TOBJcaisse.PutValue('GJC_USER', V_PGI.FUser);
  if TOBJcaisse.FieldExists('GJC_VENDOUV') then TOBJcaisse.PutValue('GJC_VENDOUV', GJC_VENDOUV.Text);
  nbOuv := TOBJcaisse.GetValue('GJC_NBOUV');
  if (nbOuv > 0) then Inc(nbOuv) else nbOuv := 1;
  TOBJcaisse.PutValue('GJC_NBOUV', nbOuv);
  // Met à jour le fond de caisse
  if TOBFDCais.Detail.Count > 0 then
  begin
    for Ind := 0 to TOBFDCais.Detail.Count - 1 do
    begin
      TOBL := TOBFDCais.Detail[Ind];
      TOBL.PutValue('GJM_CAISSE', Caisse);
      TOBL.PutValue('GJM_NUMZCAISSE', NumZ);
      TOBL.PutValue('GJM_PIECBILOUV', FOMAJTotalPieceBillet(TOBBillet, TOBPiece, TOBL.GetValue('GJM_MODEPAIE'), 'GJM'));
      TOBL.UpdateDateModif;
    end;
  end;
  // Mise à jour de la base
  io := Transactions(EnregistreSaisie, 2);
  case io of
    oeOk:
      begin
        FOReChargeVHGCCaisse;
        GenereFondCaisse;
        Result := True;
      end;
    oeUnknown: MessageAlerte(Msg.Mess[3]);
    oeSaisie: MessageAlerte(Msg.Mess[4]);
  else MessageAlerte(Msg.Mess[3]);
  end;
  // Conserve le code du dernier vendeur
  FOMajChampSupValeur(VH_GC.TOBPCaisse, 'LASTVENDEUR', GJC_VENDOUV.Text);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  EnregistreSaisie : enregistre les données saisies
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.EnregistreSaisie;
var Caisse: string;
begin
  // la caisse ne doit pas être ouverte
  Caisse := TOBCaisse.GetValue('GPK_CAISSE');
  if FOEtatCaisse(Caisse, False) = ETATJOURCAISSE[2] then
  begin
    V_PGI.IoError := oeSaisie;
    Exit;
  end;
  // enregistre la TOB de la jounée de la caisse
  TOBJCaisse.UpdateDateModif;
  if not TOBJCaisse.InsertOrUpdateDB(False) then
  begin
    V_PGI.IoError := oeUnknown;
    Exit;
  end;
  // enregistre la TOB du fond de caisse
  if not TOBFDCais.InsertOrUpdateDB(False) then
  begin
    V_PGI.IoError := oeUnknown;
    Exit;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GenereFondCaisse : génération de la pièce pour le fond de caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GenereFondCaisse;
var
  CodeArt, CodeVen, ModePaie, CodeDev, CodeTiers: string;
  MntEche: Double;
  Ind: Integer;
  TOBL: TOB;
begin
  CodeArt := VH_GC.TOBPCaisse.GetValue('GPK_FDCAISSE');
  CodeVen := GJC_VENDOUV.Text;
  for Ind := 0 to TOBFDCais.Detail.Count - 1 do
  begin
    TOBL := TOBFDCais.Detail[Ind];
    if (TOBL.GetValue('GJM_MODEPAIE') <> '') and (TOBL.GetValue('GJM_FDCAISOUV') <> 0) then
    begin
      TOBL := TOBFDCais.Detail[Ind];
      ModePaie := TOBL.GetValue('GJM_MODEPAIE');
      CodeDev := TOBL.GetValue('GJM_DEVISE');
      MntEche := TOBL.GetValue('GJM_FDCAISOUV');
//      {$IFDEF GESCOM}
//      CodeTiers := '&#@';
//      {$ELSE}
//      CodeTiers := '';
//      {$ENDIF}
      if FOGenereTicket(CodeArt, CodeVen, ModePaie, CodeDev, CodeTiers, Caption, MntEche, True, True, True) = 0 then
      begin
        V_PGI.IoError := oeUnknown;
        Exit;
      end;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeCompteRendu : Initialisation des zones du compte rendu
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.ChargeCompteRendu;
var Ind: Integer;
begin
  CompteRendu.Text := '';
  CompteRendu.SelAttributes.Style := [fsBold];
  CompteRendu.lines.Append(TraduireMemoire('Vous allez ouvrir la journée du'));
  CompteRendu.SelAttributes.Style := [fsBold, fsItalic];
  CompteRendu.lines.Append(GJC_JOURNEE.Text);
  CompteRendu.SelAttributes.Style := [fsBold];
  CompteRendu.lines.Append('');
  CompteRendu.lines.Append(TraduireMemoire('pour la caisse'));
  CompteRendu.SelAttributes.Style := [fsBold, fsItalic];
  CompteRendu.lines.Append(GJC_CAISSE.Value + ' - ' + GJC_CAISSE.Text);
  if FOEstVrai(TOBCaisse, 'GPK_GEREFONDCAISSE') then
  begin
    CompteRendu.lines.Append('');
    CompteRendu.lines.Append(TraduireMemoire('Avec en fonds de caisse'));
    for Ind := GS.FixedRows to GS.RowCount do
      CompteRendu.lines.Append(GS.Cells[GRD_MONTANT, Ind] + ' ' + GS.Cells[GRD_DEVISE, Ind]);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifieVendOuv : vérifie la cohérence du code vendeur
///////////////////////////////////////////////////////////////////////////////////////

function TFOuvertJour.VerifieVendOuv: Boolean;
begin
  Result := True;
  if GJC_VENDOUV.Visible then
  begin
    if GJC_VENDOUV.Text = '' then
    begin
      if FOEstVrai(TOBCaisse, 'GPK_VENDOBLIG') then
      begin
        //Msg.Execute(6, Caption, '') ;
        Result := False
      end;
    end else
    begin
      if not FORepresentantExiste(GJC_VENDOUV.Text, 'VEN', nil) then
      begin
        Msg.Execute(7, Caption, '');
        Result := False;
      end;
    end;
  end;
  if not Result then
  begin
    if GJC_VENDOUV.CanFocus then GJC_VENDOUV.SetFocus;
    GJC_VENDOUVElipsisClick(Self);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifieDateOuv : vérifie la cohérence de la date d'ouverture
///////////////////////////////////////////////////////////////////////////////////////

function TFOuvertJour.VerifieDateOuv: Boolean;
var
  DateOuv: TDateTime;
  sSql: string;
  QQ: TQuery;
  Nb: Integer;
  MaskOk: boolean;
begin
  Result := True;
  MaskOk := True;
  try
    GJC_JOURNEE.ValidateEdit;
  except
    MaskOk := False;
  end;
  if MaskOk then
  begin
    DateOuv := StrToDate(GJC_JOURNEE.Text);
    if DateOuv > Date then
    begin
      // Il est interdit d'ouvrir la journée pour une date dans le futur
      Msg.Execute(8, Caption, '');
      Result := False;
    end else if DateOuv < Date then
    begin
      // On vérifie si une journée n'a pas déjà été ouverte à cette date
      sSql := 'SELECT COUNT(*) FROM JOURSCAISSE WHERE GJC_CAISSE="' + GJC_CAISSE.Value + '"'
        + ' AND GJC_DATEOUV="' + USDate(GJC_JOURNEE) + '"';
      QQ := OpenSQL(sSql, True);
      if QQ.Eof then Nb := 0 else Nb := QQ.Fields[0].AsInteger;
      Ferme(QQ);
      if Nb > 0 then Result := (Msg.Execute(9, Caption, '') = mrYes);
    end;
  end else
  begin
    Msg.Execute(10, Caption, '');
    Result := False;
  end;
  if not (Result) and (GJC_JOURNEE.CanFocus) then GJC_JOURNEE.SetFocus;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifiePieceBillet : vérifie la cohérence du détail piéces et billets
///////////////////////////////////////////////////////////////////////////////////////

function TFOuvertJour.VerifiePieceBillet(ModePaie: string; TotSaisi: Double): Boolean;
var TOBB, TOBP: TOB;
  dTot: Double;
begin
  Result := True;
  if TOBBillet <> nil then
    TOBB := TOBBillet.FindFirst(['GJM_MODEPAIE'], [ModePaie], False)
  else TOBB := nil;

  if TOBPiece <> nil then
    TOBP := TOBPiece.FindFirst(['GJM_MODEPAIE'], [ModePaie], False)
  else TOBP := nil;

  if (TOBB <> nil) or (TOBP <> nil) then
  begin
    dTot := CalculTotalPieceBillet(TOBB, TOBP, False);
    Debug('Saisi=' + FloatToStr(TotSaisi) + ' Calculé=' + FloatToStr(dTot) + ' Ecart=' + FloatToStr((TotSaisi - dTot)));
    if (dTot <> 0) or (FOGetParamCaisse('GPK_CTRLPIECBIL') = 'X') then
    begin
      Result := (Arrondi(dTot, V_PGI.okDecV) = Arrondi(TotSaisi, V_PGI.okDecV));
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifieFdCaisse : vérifie la cohérence du fond de caisse
///////////////////////////////////////////////////////////////////////////////////////

function TFOuvertJour.VerifieFdCaisse: Boolean;
var Ind: Integer;
  TOBL: TOB;
begin
  Result := False;
  if TOBFDCais = nil then Exit;
  if BoutonEnCours then Exit;
  BoutonEncours := True;
  GSExit(nil);
  Result := True;
  for Ind := 0 to TOBFDCais.Detail.Count - 1 do
  begin
    TOBL := TOBFDCais.Detail[Ind];
    if not VerifiePieceBillet(TOBL.GetValue('GJM_MODEPAIE'), TOBL.GetValue('GJM_FDCAISOUV')) then
    begin
      PGIBox('Le détail des pièces et de billets est incorrect !', Caption);
      Result := False;
      Break;
    end;
  end;
  BoutonEncours := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormCreate
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.FormCreate(Sender: TObject);
begin
  inherited;
  // Creation des TOB de la caisse
  TOBCaisse := TOB.Create('PARCAISSE', nil, -1);
  TOBJCaisse := TOB.Create('JOURSCAISSE', nil, -1);
  TOBFDCais := TOB.Create('FondCaisse', nil, -1);
  TOBBillet := TOB.Create('Les billets', nil, -1);
  TOBPiece := TOB.Create('Les pieces', nil, -1);
  BoutonEncours := False;
  PiecBilEnCours := False;
  PnlBoutons := nil;
  if FOGetParamCaisse('GPK_CLAVIERECRAN') = 'X' then WindowState := wsMaximized;
  // Calcul du nombres d'étapes
  NoPage := 1;
  NbPage := P.PageCount;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormClose
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
  // Libération des TOB de la caisse
  TOBCaisse.Free;
  TOBJCaisse.Free;
  TOBFDCais.Free;
  TOBBillet.Free;
  TOBPiece.Free;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormShow
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.FormShow(Sender: TObject);
var Ind: integer;
  CashDef: string;
begin
  inherited;
  // Lecture de la caisse choisie à la connexion
  CashDef := VH_GC.TOBPCaisse.GetValue('GPK_LIBELLE');
  Ind := GJC_CAISSE.Items.IndexOf(CashDef);
  if Ind >= 0 then GJC_CAISSE.ItemIndex := Ind else GJC_CAISSE.ItemIndex := 0;
  GJC_Caisse.OnChange(nil);
  // En mode SAV, on peut ouvrir la journée à une autre date que la date système
  if V_PGI.SAV then
  begin
    GJC_JOURNEE.Enabled := True;
    GJC_JOURNEE.Color := clWindow;
  end;
  // Activation de la saisie du vendeur
  if (TOBJcaisse.FieldExists('GJC_VENDOUV')) and not (FOEstVrai(TOBCaisse, 'GPK_VENDSAISIE')) and
    not (FOEstVrai(TOBCaisse, 'GPK_VENDSAISLIG')) then
  begin
    GJC_VENDOUV.Visible := False;
    GJC_VENDOUV.Visible := False;
  end;
  // On inactive le bouton FIN
  ActivebFin;
  // Grille des billets
     // branchement des événements
  GBillet.OnColumnWidthsChanged := GBilletColumnWidthsChanged;
  // colonnes inaccessibles
  GBillet.ColLengths[COLB_LIBELLE] := -1;
  GBillet.ColLengths[COLB_MONTANT] := -1;
  // format de saisie
  GBillet.ColFormats[COLB_QUANTITE] := '###0';
  // alignement des données
  GPiecBilColAligns(GBillet);
  // Grille des pièces
     // branchement des événements
  GPiece.OnColumnWidthsChanged := GPieceColumnWidthsChanged;
  // colonnes inaccessibles
  GPiece.ColLengths[COLB_LIBELLE] := -1;
  GPiece.ColLengths[COLB_MONTANT] := -1;
  // format de saisie
  GPiece.ColFormats[COLB_QUANTITE] := '###0';
  // alignement des données
  GPiecBilColAligns(GPiece);
  PPIECE.Enabled := True;
  if FOGetParamCaisse('GPK_CLAVIERECRAN') <> 'X' then
  begin
    PPIECBIL.Align := alTop;
    PPIECBIL.Height := GroupBox1.Top - 1;
    GroupBox1.Anchors := GroupBox1.Anchors - [akRight, akBottom] + [akTop];
    if GroupBox1.Top > BBILLET.Top then
    begin
      Ind := (GroupBox1.Top + GroupBox1.Height) - BBILLET.Top + 2;
      Height := Height + Ind;
      BENTETE.Left := BBILLET.Left;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormResize :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.FormResize(Sender: TObject);
///////////////////////////////////////////////////////////////////////////////////////
  procedure PlaceBoutons(Btn: TToolbarButton97; ALeft: Integer);
  begin
    Btn.Height := 50;
    Btn.Width := 120;
    Btn.Top := Height - 78;
    Btn.Left := ALeft;
    Btn.Font.Size := 11;
  end;
  ///////////////////////////////////////////////////////////////////////////////////////
  procedure PlaceBtnFleche(Btn: TToolbarButton97; Rang: Integer);
  begin
    Btn.Height := 50;
    Btn.Width := 60;
    Btn.Top := Height - 78;
    Btn.Left := (Width - (Rang * (Btn.Width + 1)) - 14);
    Btn.Font.Size := 11;
    Btn.Visible := True;
  end;
  ///////////////////////////////////////////////////////////////////////////////////////
var
  AHeight: Integer;
  Ctrl: TComponent;
begin
  inherited;
  if FOGetParamCaisse('GPK_CLAVIERECRAN') = 'X' then
  begin
    lEtape.Visible := False;
    GroupBox1.Top := Height - 85;
    PlaceBoutons(bAide, 20);
    PlaceBoutons(bAnnuler, 145);
    PlaceBoutons(bPrecedent, 271);
    PlaceBoutons(bSuivant, 392);
    PlaceBoutons(bFin, 518);
    PlaceBtnFleche(BLEFT, 4);
    PlaceBtnFleche(BRIGHT, 3);
    PlaceBtnFleche(BUP, 2);
    PlaceBtnFleche(BDOWN, 1);
    BENTETE.Left := 1;
    BENTETE.Top := Height - 78;
    BBILLET.Left := 52;
    BBILLET.Top := Height - 78;
    PPave.Left := 0;
    PPave.Width := Width;
    PPave.Top := P.Top + P.Height;
    AHeight := GroupBox1.Top - PPave.Top;
    if AHeight > 0 then PPave.Height := AHeight;
    PPave.Visible := True;
    if PnlBoutons = nil then
    begin
      FOCreateClavierEcran(PnlBoutons, Self, PPave, False, True);
      PnlBoutons.LanceBouton := SaisieClavierEcran;
      PnlBoutons.LanceCalculette := SaisieCalculette;
      PnlBoutons.BoutonCalculette := BoutonCalculetteClick;
      PnlBoutons.Caisse := FOCaisseCourante;

      if (P.ActivePage.Name = 'PChoixCaisse') and (GJC_VENDOUV.Focused) then ChoixVendeurPave;
      // détail pièces billets
      PPIECBIL.Parent := PPave;
      PPIECBIL.Align := alLeft;
      Ctrl := PnlBoutons.FindComponent('CALCULETTE');
      if (Ctrl <> nil) and (Ctrl is THPanel) then
      begin
        PPIECBIL.Height := THPanel(Ctrl).Height;
        PPIECBIL.Width := THPanel(Ctrl).Left;
      end;
      PBILLET.Width := PPIECBIL.Width div 2;
      PPIECE.Width := PPIECBIL.Width - PBILLET.Width;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormKeyDown :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  inherited;
  case Key of
    VK_RETURN: Key := VK_TAB;
    VK_F10: if Shift = [] then
      begin
        Key := 0;
        BoutonCalculetteClick('ENTER');
      end;
    VK_HOME: if Shift = [ssCtrl] then
      begin
        Key := 0;
        GotoEntete;
      end;
    VK_END: if Shift = [ssCtrl] then
      begin
        Key := 0;
        GotoPied;
      end;
    VK_ESCAPE: if Shift = [] then
      begin
        Key := 0;
        BoutonCalculetteClick('CLEAR');
      end;
  end;
  if (Key = VK_DOWN) or ((Key = VK_TAB) and (Shift = [])) then
  begin
    if (ActiveControl = GBILLET) and (GBILLET.Row = GBILLET.RowCount - 1) then
    begin
      if GPIECE.CanFocus then
      begin
        GPIECE.Row := GPIECE.FixedRows;
        GPIECE.Col := COLB_QUANTITE;
        GPIECE.SetFocus;
      end;
    end;
    if (ActiveControl = GPiece) and (GPIECE.Row = GPIECE.RowCount - 1) then
    begin
      if GBILLET.CanFocus then
      begin
        GBILLET.Row := GBILLET.FixedRows;
        GBILLET.Col := COLB_QUANTITE;
        GBILLET.SetFocus;
      end;
    end;
  end;
  if (Key = VK_UP) or ((Key = VK_TAB) and (Shift = [ssShift])) then
  begin
    if (ActiveControl = GBILLET) and (GBILLET.Row = GBILLET.FixedRows) then
    begin
      if GPIECE.CanFocus then
      begin
        GPIECE.Row := GPIECE.RowCount - 1;
        GPIECE.Col := COLB_QUANTITE;
        GPIECE.SetFocus;
      end;
    end;
    if (ActiveControl = GPiece) and (GPIECE.Row = GPIECE.FixedRows) then
    begin
      if GBILLET.CanFocus then
      begin
        GBILLET.Row := GBILLET.RowCount - 1;
        GBILLET.Col := COLB_QUANTITE;
        GBILLET.SetFocus;
      end;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ActivebFin : Activation du bouton FIN
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.ActivebFin;
var Ok: Boolean;
begin
  Ok := (P.ActivePage.Name = 'PResume');
  bFin.Enabled := Ok;
  if FOGetParamCaisse('GPK_CLAVIERECRAN') = 'X' then
  begin
    bFin.Visible := Ok;
    BLEFT.Visible := (not Ok);
    BRIGHT.Visible := (not Ok);
    BUP.Visible := (not Ok);
    BDOWN.Visible := (not Ok);
  end;
  Ok := (P.ActivePage.Name = 'PfondCaisse');
  if FOGetParamCaisse('GPK_CLAVIERECRAN') = 'X' then
  begin
    BENTETE.Visible := Ok;
    BBILLET.Visible := Ok;
  end else
  begin
    BBILLET.Visible := Ok;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GPiecBilColAligns : définition de l'alignement des données des pièces et billets
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GPiecBilColAligns(Grille: THGrid);
begin
  if COLB_QUANTITE >= 0 then Grille.ColAligns[COLB_QUANTITE] := taRightJustify;
  if COLB_MONTANT >= 0 then Grille.ColAligns[COLB_MONTANT] := taRightJustify;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GPiecBilPutGrid : affichage des lignes de la TOB dans la grille des pièces ou dans celle de billets
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GPiecBilPutGrid(Grille: THGrid; TOBPB: TOB);
var
  sColName: string;
  NewTob: boolean;
begin
  if TOBPB = nil then
  begin
    TOBPB := TOB.Create('', nil, -1);
    NewTob := True;
  end else
    NewTob := False;
  if Grille = nil then Exit;
  sColName := 'GPI_LIBELLE;QUANTITE;MONTANT';
  TOBPB.PutGridDetail(Grille, False, False, sColName, True);
  HMTrad.ResizeGridColumns(Grille);
  if NewTob then TOBPB.Free;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  bFinClick : bouton FIN
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.bFinClick(Sender: TObject);
begin
  if not bFin.Enabled then Exit;
  inherited;
  // Interdit les boutons
  bAide.Enabled := False;
  bAnnuler.Enabled := False;
  bPrecedent.Enabled := False;
  bSuivant.Enabled := False;
  bFin.Enabled := False;
  // Enregistre les données saisies
  if SauveCaisse then
  begin
    ImprimeTicket;
    // Arrêt des échanges FO - BO
    {$IFDEF TOXCLIENT}
    AglToxConf(aceStop, '', nil, nil, nil);
    {$ENDIF}
  end;
  Close;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  bSuivantClick : bouton SUIVANT
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.bSuivantClick(Sender: TObject);
begin
  if not bSuivant.Enabled then Exit;
  // Verification du code vendeur et du fond de caisse
  if (P.ActivePage.Name = 'PChoixCaisse') and ((not VerifieVendOuv) or (not VerifieDateOuv)) then Exit;
  if (P.ActivePage.Name = 'PfondCaisse') and (not VerifieFdCaisse) then Exit;
  // Incrementation du n° de page avant traitement du changement de page
  Inc(Nopage);
  inherited;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  bPrecedentClick : bouton PRECEDENT
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.bPrecedentClick(Sender: TObject);
begin
  if not bPrecedent.Enabled then Exit;
  // Décrementation du n° de page avant traitement du changement de page
  Dec(Nopage);
  inherited;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  PChange
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.PChange(Sender: TObject);
begin
  inherited;
  // Affiche le n° réelle de la page, en fonction du nombre de pages à traiter.
  lEtape.Caption := Msg.Mess[0] + ' ' + IntToStr(Nopage) + '/' + IntToStr(NbPage);
  // Affichage de l'image de l'onglet
  if (P.ActivePage.ImageIndex >= 0) and (P.ActivePage.ImageIndex < Image.Images.Count) then
  begin
    Image.Visible := True;
    Image.ImageIndex := P.ActivePage.ImageIndex;
  end else
  begin
    Image.Visible := False;
  end;
  lEtape.Visible := not (P.ActivePage.Name = 'PfondCaisse');
  if (P.ActivePage.Name <> 'PfondCaisse') then
  begin
    BBILLET.Visible := False;
    BENTETE.Visible := False;
  end;
  // ouverture du tiroir caisse si nécessaire
  if (P.ActivePage.Name = 'PfondCaisse') then FOOuvreTiroir(False, True);
  // Initialisation des zones du compte rendu
  if (P.ActivePage.Name = 'PResume') then ChargeCompteRendu;
  // Active le détail pièce billet
  PPIECBIL.Visible := ((P.ActivePage.Name = 'PfondCaisse') and (FOGetParamCaisse('GPK_CLAVIERECRAN') = 'X'));
  // Activation du bouton FIN
  ActivebFin;
  // Efface le pavé tactile
  if PnlBoutons <> nil then PnlBoutons.PageCourante := 1;
  // Donne le focus au 1er champ de la page
  if (P.ActivePage.Name = 'PChoixCaisse') and (GJC_VENDOUV.CanFocus) then GJC_VENDOUV.SetFocus;
  if (P.ActivePage.Name = 'PfondCaisse') and (GS.CanFocus) then
  begin
    GS.SetFocus;
    if GS.InplaceEditor <> nil then GS.InplaceEditor.SelectAll;
  end;
  if (P.ActivePage.Name = 'PResume') and (CompteRendu.CanFocus) then CompteRendu.SetFocus;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  PreviousPage
///////////////////////////////////////////////////////////////////////////////////////

function TFOuvertJour.PreviousPage: TTabSheet;
var TT: TTabsheet;
  Boucle: Integer;
begin
  TT := inherited PreviousPage;
  // Active ou non les onglets en fonction du paramétrage de la caisse
  Boucle := -1;
  while (not VerifiePageActive(TT)) do
  begin
    TT := P.Pages[P.ActivePage.PageIndex - 1 + Boucle];
    Dec(Boucle);
  end;
  Result := TT;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  NextPage
///////////////////////////////////////////////////////////////////////////////////////

function TFOuvertJour.NextPage: TTabSheet;
var TT: TTabsheet;
  Boucle: Integer;
begin
  TT := inherited NextPage;
  // Active ou non les onglets en fonction du paramétrage de la caisse
  Boucle := 1;
  while (not VerifiePageActive(TT)) do
  begin
    TT := P.Pages[P.ActivePage.PageIndex + 1 + Boucle];
    Inc(Boucle);
  end;
  Result := TT;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifiePageActive : vérifie si une page doit être affichée
///////////////////////////////////////////////////////////////////////////////////////

function TFOuvertJour.VerifiePageActive(TT: TTabSheet): Boolean;
begin
  Result := True;
  // Active ou non les onglets en fonction du paramétrage de la caisse
  if (TT <> nil) and (TOBCaisse <> nil) then
  begin
    if TT.Name = 'PfondCaisse' then
    begin
      Result := FOEstVrai(TOBCaisse, 'GPK_GEREFONDCAISSE');
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChoixVendeurPave : choix du vendeur avec le pavé tactile
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.ChoixVendeurPave;
var sPlus: string;
begin
  if PnlBoutons <> nil then
  begin
    sPlus := 'GCL_TYPECOMMERCIAL="VEN" and GCL_ETABLISSEMENT="' + GPK_ETABLISSEMENT.Value + '"'
      + ' and GCL_DATESUPP>"' + USDateTime(Date) + '"';
    PnlBoutons.LanceChargePageMenu('VEN', '', sPlus, clBtnFace, 40);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GJC_VENDOUVEnter
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GJC_VENDOUVEnter(Sender: TObject);
begin
  inherited;
  ChoixVendeurPave;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GJC_VENDOUVDblClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GJC_VENDOUVDblClick(Sender: TObject);
begin
  inherited;
  FOGetVendeur(GJC_VENDOUV, Msg.Mess[5], '', GPK_ETABLISSEMENT.Value);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GJC_VENDOUVlipsisClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GJC_VENDOUVElipsisClick(Sender: TObject);
begin
  inherited;
  FOGetVendeur(GJC_VENDOUV, Msg.Mess[5], '', GPK_ETABLISSEMENT.Value);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GJC_CAISSEChange
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GJC_CAISSEChange(Sender: TObject);
var Ind: Integer;
  TOBL: TOB;
  sMdp, sDev: string;
begin
  inherited;
  // Recherche de la caisse choisie
  ChargeCaisse;
  // Initialisation du fond de caisse et de la liste des pièces et billets
  ChargeFondCaisse;
  for Ind := 0 to TOBFDCais.Detail.Count - 1 do
  begin
    TOBL := TOBFDCais.Detail[Ind];
    sMdp := TOBL.GetValue('GJM_MODEPAIE');
    sDev := TOBL.GetValue('GJM_DEVISE');
    FOChargeTOBPiecBil(TOBL, TOBBillet, TOBPiece, sMdp, sDev, 'GJM');
    if FOGetParamCaisse('GPK_REPRISEFDCAIS') = 'X' then ReprisePieceBillet(sMdp, TOBL.GetValue('GJM_PIECBILOUV'));
  end;
  // Calcul du nombres d'étapes
  NbPage := 0;
  for Ind := 0 to P.PageCount - 1 do if VerifiePageActive(P.Pages[Ind]) then Inc(NbPage);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ReprisePieceBillet : reprise des quantité de pièces et billets
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.ReprisePieceBillet(ModePaie, Texte: string);
var TOBPrinc, TOBL: TOB;
  Stg: string;
  iCode: Integer;
  dQte: Double;
begin
  if Texte = '' then Exit;
  TOBBilMdp := TOBBillet.FindFirst(['GJM_MODEPAIE'], [ModePaie], False);
  TOBPieMdp := TOBPiece.FindFirst(['GJM_MODEPAIE'], [ModePaie], False);
  Stg := ReadTokenST(Texte);
  while Stg <> '' do
  begin
    if Copy(Stg, 1, 1) = 'B' then TOBPrinc := TOBBilMdp else TOBPrinc := TOBPieMdp;
    Delete(Stg, 1, 1);
    iCode := StrToInt(ReadTokenPipe(Stg, '='));
    dQte := StrToFloat(ReadTokenPipe(Stg, '='));
    if TOBPrinc <> nil then TOBL := TOBPrinc.FindFirst(['GPI_PIECEBILLET'], [iCode], False)
    else TOBL := nil;
    if TOBL <> nil then TOBL.PutValue('QUANTITE', dQte);
    Stg := ReadTokenST(Texte);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GJC_JOURNEEExit
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GJC_JOURNEEExit(Sender: TObject);
begin
  inherited;
  VerifieDateOuv;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ImprimeTicket : impression des tickets de fond de caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.ImprimeTicket;
var sWhere: string;
  NumZ: Integer;
  Ok: Boolean;
begin
  if (FOEstVrai(TOBCaisse, 'GPK_GEREFONDCAISSE')) and (FOExisteCodeEtat(efoFdCaisse)) then
  begin
    Ok := FOGetFromRegistry(REGJOURNEE, REGFERMEFDCAISSE, True);
    if Ok then
    begin
      CompteRendu.lines.Append(TraduireMemoire('Impression du fonds de caisse'));
      NumZ := TOBJcaisse.GetValue('GJC_NUMZCAISSE');
      sWhere := 'GJM_CAISSE="' + GJC_CAISSE.Value + '" and GJM_NUMZCAISSE="' + IntToStr(NumZ) + '"';
      FOLanceImprimeLP(efoFdCaisse, sWhere, False, nil);
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeFondCaisse : Initialisation du tableau de saisie du fond de caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.ChargeFondCaisse;
var TOBMDP, TOBL, TOBF: TOB;
  Select, sDev, sSym: string;
  Ind, NumZ: Integer;
  dMt: Double;
begin
  NumZ := FOGetNumZCaisse(GJC_CAISSE.Value);
  // colonnes inaccessibles
  GS.ColLengths[GRD_MODEPAIE] := -1;
  GS.ColLengths[GRD_DEVISE] := -1;
  // alignement des données
  GSColAligns;
  // positionnement sur la colonne de saisie et ajustement automatique de la taille des colonnes
  GS.Col := GRD_MONTANT;
  // recherche des modes de paiement concernés par le fond de caisse
  Select := 'select MP_MODEPAIE, MP_LIBELLE, MP_DEVISEFO, D_SYMBOLE, D_DECIMALE from MODEPAIE, DEVISE '
    + 'where MP_UTILFO="X" and MP_TYPEMODEPAIE="' + TYPEPAIEESPECE + '" '
    + 'and MP_DEVISEFO=D_DEVISE' + FOFabriqueListeMDP('GPK_MDPFDCAIS', 'MP_MODEPAIE');
  TOBMDP := TOB.Create('ModePaiement', nil, -1);
  TOBMDP.LoadDetailFromSQL(Select);
  // constitution de la TOB du fond de caisse
  for Ind := 0 to TOBMDP.Detail.Count - 1 do
  begin
    TOBL := TOBMDP.Detail[Ind];
    if TOBL <> nil then
    begin
      sDev := TOBL.GetValue('MP_DEVISEFO');
      if FOTestCodeEuro(sDev) then sSym := SIGLEEURO else sSym := TOBL.GetValue('D_SYMBOLE');
      TOBF := TOB.Create('CTRLCAISMT', TOBFDCais, -1);
      TOBF.InitValeurs;
      TOBF.PutValue('GJM_CAISSE', GJC_CAISSE.Value);
      TOBF.PutValue('GJM_NUMZCAISSE', 0);
      TOBF.PutValue('GJM_MODEPAIE', TOBL.GetValue('MP_MODEPAIE'));
      TOBF.PutValue('GJM_DEVISE', sDev);
      TOBF.PutValue('GJM_FDCAISSEDEV', 0);
      TOBF.PutValue('GJM_FDCAISOUV', 0);
      TOBF.AddChampSupValeur('LIBELLE', TOBL.GetValue('MP_LIBELLE'));
      TOBF.AddChampSupValeur('SYMBOLE', sSym);
      TOBF.AddChampSupValeur('NBDEC', TOBL.GetValue('D_DECIMALE'));
    end;
  end;
  TOBMDP.Free;
  // Reprise des montants saisis lors de la fermeture précédente
  if FOGetParamCaisse('GPK_REPRISEFDCAIS') = 'X' then
  begin
    Select := 'select GJM_MODEPAIE,GJM_FDCAISSEDEV,GJM_PIECBILTOT from CTRLCAISMT '
      + 'where GJM_CAISSE="' + GJC_CAISSE.Value + '"'
      + ' and GJM_NUMZCAISSE=' + IntToStr(NumZ)
      + ' and ' + FOFabriqueListeMDP('GPK_MDPFDCAIS', 'GJM_MODEPAIE');
    TOBMDP := TOB.Create('Montants précédents', nil, -1);
    TOBMDP.LoadDetailFromSQL(Select);
    for Ind := 0 to TOBMDP.Detail.Count - 1 do
    begin
      TOBL := TOBMDP.Detail[Ind];
      TOBF := TOBFDCais.FindFirst(['GJM_MODEPAIE'], [TOBL.GetValue('GJM_MODEPAIE')], False);
      if TOBF <> nil then
      begin
        dMt := TOBL.GetValue('GJM_FDCAISSEDEV');
        TOBF.PutValue('GJM_FDCAISOUV', dMt);
        TOBF.PutValue('GJM_FDCAISSEDEV', dMt);
        ////TOBF.PutValue('GJM_PIECBILOUV', TOBL.GetValue('GJM_PIECBILTOT')) ;
      end;
    end;
    TOBMDP.Free;
  end;
  TOBFDCais.PutGridDetail(GS, False, False, 'LIBELLE;SYMBOLE;GJM_FDCAISOUV', True);
  HMTrad.ResizeGridColumns(GS);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GetTOBFDCais : retourne TOB correspondante à une ligne de la grille
///////////////////////////////////////////////////////////////////////////////////////

function TFOuvertJour.GetTOBFDCais(ARow: Integer): TOB;
begin
  Result := nil;
  if (Arow < 1) or (ARow > TOBFDCais.Detail.Count) then Exit;
  Result := TOBFDCais.Detail[ARow - 1];
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSColAligns : définition de l'alignement des données de la grille du fond de caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GSColAligns;
begin
  GS.ColAligns[GRD_DEVISE] := taCenter;
  GS.ColAligns[GRD_MONTANT] := taRightJustify;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSCellEnter :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  inherited;
  if not PiecBilEnCours then
  begin
    if (FOGetParamCaisse('GPK_CTRLPIECBIL') = 'X') and
       (FOGetParamCaisse('GPK_CLAVIERECRAN') = 'X') then
      GotoPied;
  end else
    PiecBilEnCours := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSCellExit :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var Montant: Double;
  TOBL: TOB;
begin
  inherited;
  TOBL := GetTOBFDCais(ARow);
  if TOBL = nil then
  begin
    Cancel := True;
    Exit;
  end;
  Montant := Valeur(GS.Cells[ACol, ARow]);
  if Montant < 0 then
  begin
    PGIBox('Montant incorrect !', 'Fond de caisse');
    Cancel := True;
  end else
  begin
    FormateZoneSaisie(GS, ACol, ARow);
    TOBL.PutValue('GJM_FDCAISSEDEV', Montant); // à conserver pour cmpatibilté avec liste du fond de caisse
    TOBL.PutValue('GJM_FDCAISOUV', Montant);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSEnter :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GSEnter(Sender: TObject);
var Bc, Cancel: Boolean;
  ACol, ARow: Integer;
begin
  inherited;
  GS.TitleBold := True;
  GSColAligns;
  ARow := GS.Row;
  ACol := GS.Col;
  Cancel := False;
  Bc := False;
  GSRowEnter(Self, ARow, Bc, False);
  GSCellEnter(Self, ACol, ARow, Cancel);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSExit :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GSExit(Sender: TObject);
var ARow, ACol: Integer;
  Cancel: Boolean;
begin
  GS.TitleBold := False;
  GSColAligns;
  ARow := GS.Row;
  ACol := GS.Col;
  Cancel := False;
  GSCellExit(Self, ACol, ARow, Cancel);
  if Cancel then
  begin
    if GS.CanFocus then GS.SetFocus;
    Exit;
  end;
  inherited;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSRowEnter :
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var TOBC: TOB;
  sTitre, sMdp: string;
begin
  inherited;
  TOBBilMdp := nil;
  TOBPieMdp := nil;
  TOBC := GetTOBFDCais(GS.Row);
  if TOBC <> nil then
  begin
    sTitre := TOBC.GetValue('LIBELLE');
    sMdp := TOBC.GetValue('GJM_MODEPAIE');
    TOBBilMdp := TOBBillet.FindFirst(['GJM_MODEPAIE'], [sMdp], False);
    TOBPieMdp := TOBPiece.FindFirst(['GJM_MODEPAIE'], [sMdp], False);
    BBILLET.Enabled := ((TOBBilMdp <> nil) or (TOBPieMdp <> nil));
    if FOGetParamCaisse('GPK_CLAVIERECRAN') = 'X' then
      PPIECBIL.Visible := ((TOBBilMdp <> nil) or (TOBPieMdp <> nil));
    GPiecBilPutGrid(GBillet, TOBBilMdp);
    GPiecBilPutGrid(GPiece, TOBPieMdp);
    CalculTotalPieceBillet(TOBBilMdp, TOBPieMdp, True);
    PCUMUL.Caption := sTitre;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GBilletEnter : OnEnter sur la grille des billets
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GBilletEnter(Sender: TObject);
var Cancel: Boolean;
  ACol, ARow: Integer;
begin
  PiecBilEnCours := True;
  GPiece.TitleBold := True;
  GPiecBilColAligns(GPiece);
  GBillet.TitleBold := True;
  GPiecBilColAligns(GBillet);
  Cancel := False;
  ACol := GBillet.Col;
  ARow := GBillet.Row;
  GBilletCellEnter(nil, ACol, ARow, Cancel);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GBilletExit : OnExit sur la grille des billets
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GBilletExit(Sender: TObject);
var ACol, ARow: Integer;
  Cancel: Boolean;
begin
  Cancel := False;
  ACol := GBillet.Col;
  ARow := GBillet.Row;
  GBilletCellExit(nil, ACol, ARow, Cancel);
  GPiece.TitleBold := False;
  GPiecBilColAligns(GPiece);
  GBillet.TitleBold := False;
  GPiecBilColAligns(GBillet);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GBilletCellEnter : OnCellEnter sur la grille des billets
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GBilletCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var OldEna: Boolean;
begin
  if GBillet.Col <> COLB_QUANTITE then
  begin
    OldEna := GBillet.SynEnabled;
    GBillet.SynEnabled := False;
    GBillet.Col := COLB_QUANTITE;
    GBillet.SynEnabled := OldEna;
  end;
  StPiecBilCur := GBillet.Cells[GBillet.Col, GBillet.Row];
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GBilletCellExit : OnCellExit sur la grille des billets
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GBilletCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  //if csDestroying in ComponentState then Exit ;
  if GBillet.Cells[ACol, ARow] = StPiecBilCur then Exit;
  if ACol = COLB_QUANTITE then
  begin
    FormateZoneSaisie(GBillet, ACol, ARow);
    CalculMontantPieceBillet(GBillet, TOBBilMdp, ACol, ARow, Cancel);
    if not Cancel then CalculTotalPieceBillet(TOBBilMdp, TOBPieMdp, True);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GBilletColumnWidthsChanged : OnColumnWidthsChanged sur la grille des billets
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GBilletColumnWidthsChanged(Sender: TObject);
var Coord: TRect;
  ALeft, AWidth: Integer;
begin
  Coord := GBillet.CellRect(COLB_MONTANT, 0);
  ALeft := Coord.Left + 1;
  AWidth := GBillet.ColWidths[COLB_MONTANT] + 1;
  CUMMTBILLET.Left := ALeft;
  CUMMTBILLET.Width := AWidth;
  CUMMTBILLET.Ctl3D := False;
  CUMMTBILLET.ParentCtl3D := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GPieceEnter : OnEnter sur la grille des pièces
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GPieceEnter(Sender: TObject);
var Cancel: Boolean;
  ACol, ARow: Integer;
begin
  PiecBilEnCours := True;
  GPiece.TitleBold := True;
  GPiecBilColAligns(GPiece);
  GBillet.TitleBold := True;
  GPiecBilColAligns(GBillet);
  Cancel := False;
  ACol := GPiece.Col;
  ARow := GPiece.Row;
  GPieceCellEnter(nil, ACol, ARow, Cancel);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GPieceExit : OnExit sur la grille des pièces
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GPieceExit(Sender: TObject);
var ACol, ARow: Integer;
  Cancel: Boolean;
begin
  Cancel := False;
  ACol := GPiece.Col;
  ARow := GPiece.Row;
  GPieceCellExit(nil, ACol, ARow, Cancel);
  GPiece.TitleBold := False;
  GPiecBilColAligns(GPiece);
  GBillet.TitleBold := False;
  GPiecBilColAligns(GBillet);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GPieceCellEnter : OnCellEnter sur la grille des pièces
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GPieceCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var OldEna: Boolean;
begin
  if GPiece.Col <> COLB_QUANTITE then
  begin
    OldEna := GPiece.SynEnabled;
    GPiece.SynEnabled := False;
    GPiece.Col := COLB_QUANTITE;
    GPiece.SynEnabled := OldEna;
  end;
  StPiecBilCur := GPiece.Cells[GPiece.Col, GPiece.Row];
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GPieceCellExit : OnCellExit sur la grille des pièces
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GPieceCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  //if csDestroying in ComponentState then Exit ;
  if GPiece.Cells[ACol, ARow] = StPiecBilCur then Exit;
  if ACol = COLB_QUANTITE then
  begin
    FormateZoneSaisie(GPiece, ACol, ARow);
    CalculMontantPieceBillet(GPiece, TOBPieMdp, ACol, ARow, Cancel);
    if not Cancel then CalculTotalPieceBillet(TOBBilMdp, TOBPieMdp, True);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GPieceColumnWidthsChanged : OnColumnWidthsChanged sur la grille des pièces
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GPieceColumnWidthsChanged(Sender: TObject);
var Coord: TRect;
  ALeft, AWidth: Integer;
begin
  Coord := GPiece.CellRect(COLB_MONTANT, 0);
  ALeft := Coord.Left + 1;
  AWidth := GPiece.ColWidths[COLB_MONTANT] + 1;
  CUMMTPIECE.Left := ALeft;
  CUMMTPIECE.Width := AWidth;
  CUMMTPIECE.Ctl3D := False;
  CUMMTPIECE.ParentCtl3D := False;
  ALeft := PPIECE.Left + 1;
  AWidth := Coord.Right;
  CUMMONTANT.Left := ALeft;
  CUMMONTANT.Width := AWidth;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormateZoneSaisie : formatage de la cellule
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.FormateZoneSaisie(Grille: THGrid; ACol, ARow: Longint);
var Stg: string;
  iNbDec: Integer;
  TOBL: TOB;
begin
  iNbDec := 0;
  if Grille = GS then
  begin
    TOBL := GetTOBFDCais(GS.Row);
    if TOBL <> nil then iNbDec := TOBL.GetValue('NBDEC') else iNbDec := V_PGI.OkDecV;
  end;
  Stg := StrS(Valeur(Grille.Cells[ACol, ARow]), iNbDec);
  Grille.Cells[ACol, ARow] := Stg;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  CalculMontantPieceBillet : calcul du montant d'une ligne de pièces et de billets
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.CalculMontantPieceBillet(Grille: THGrid; TOBParent: TOB; ACol, ARow: Longint; var Cancel: Boolean);
var TOBL, TOBE: TOB;
  dQte, dMt: Double;
  iLig, iNbdec: Integer;
begin
  TOBL := nil;
  iLig := ARow - Grille.FixedRows;
  TOBE := GetTOBFDCais(GS.Row);
  if TOBE <> nil then iNbDec := TOBE.GetValue('NBDEC') else iNbDec := V_PGI.OkDecV;
  if (iLig >= 0) and (iLig < TOBParent.Detail.Count) then TOBL := TOBParent.Detail[iLig];
  if TOBL <> nil then dQte := Valeur(Grille.Cells[ACol, ARow]) else dQte := 0;
  if dQte < 0 then
  begin
    PGIBox('Vous devez saisir une quantité positive !', Caption);
    Cancel := True;
    Exit;
  end;
  TOBL.PutValue('QUANTITE', dQte);
  dMt := dQte * TOBL.GetValue('GPI_VALEUR');
  TOBL.PutValue('MONTANT', dMt);
  Grille.Cells[COLB_MONTANT, ARow] := StrS(dMt, iNbdec);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  CalculTotalPieceBillet : calcul le total des pièces et billets
///////////////////////////////////////////////////////////////////////////////////////

function TFOuvertJour.CalculTotalPieceBillet(TOBBil, TOBPie: TOB; Affiche: Boolean): Double;
var TOBL: TOB;
  Ind: Integer;
  dBil, dPiec, dQte, dVal: Double;
begin
  dBil := 0;
  if TOBBil <> nil then
  begin
    for Ind := 0 to TOBBil.Detail.Count - 1 do
    begin
      TOBL := TOBBil.Detail[Ind];
      dQte := TOBL.GetValue('QUANTITE');
      dVal := TOBL.GetValue('GPI_VALEUR');
      dBil := dBil + Arrondi((dQte * dVal), V_PGI.okDecV);
      Debug('Quantité=' + FloatToStr(dQte) + ' Valeur=' + FloatToStr(dVal) + ' Montant=' + FloatToStr((dQte * dVal)) + ' Total billets=' + FloatToStr(dBil));
    end;
  end;
  if Affiche then CUMMTBILLET.Value := dBil;
  dPiec := 0;
  if TOBPie <> nil then
  begin
    for Ind := 0 to TOBPie.Detail.Count - 1 do
    begin
      TOBL := TOBPie.Detail[Ind];
      dQte := TOBL.GetValue('QUANTITE');
      dVal := TOBL.GetValue('GPI_VALEUR');
      dPiec := dPiec + Arrondi((dQte * dVal), V_PGI.okDecV);
      Debug('Quantité=' + FloatToStr(dQte) + ' Valeur=' + FloatToStr(dVal) + ' Montant=' + FloatToStr((dQte * dVal)) + ' Total pièces=' + FloatToStr(dPiec));
    end;
  end;
  if Affiche then CUMMTPIECE.Value := dPiec;
  Result := dBil + dPiec;
  Debug('Total billets et pièces=' + FloatToStr((dBil + dPiec)));
  if Affiche then CUMMONTANT.Value := Result;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SaisieClavierEcran : interprète les boutons du pavé
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.SaisieClavierEcran(Concept, Code, Extra: string; Qte, Prix: Double);
begin
  if BoutonEncours then Exit;
  BoutonEncours := True;
  if Concept = 'VEN' then // changement de vendeur
  begin
    if P.ActivePage.Name = 'PChoixCaisse' then
    begin
      if GJC_VENDOUV.CanFocus then GJC_VENDOUV.SetFocus;
      GJC_VENDOUV.Text := Code;
      FOSimuleClavier(VK_TAB);
    end;
  end else
    ;
  BoutonEncours := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BoutonCalculetteClick : interprète les boutons de la calculette
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.BoutonCalculetteClick(Val: string);
begin
  if BoutonEncours then Exit;
  BoutonEncours := True;
  if Val = 'ENTER' then
  begin
    if (P.ActivePage.Name = 'PfondCaisse') and (ActiveControl <> GS) then
    begin
      AppliqueTotal;
    end else
    begin
      BoutonEncours := False;
      BSuivantClick(nil);
    end;
  end else
    if Val = 'CLEAR' then
  begin
    Close;
  end;
  BoutonEncours := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SaisieCalculette : prend en compte un chiffre saisi sur la calculette
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.SaisieCalculette(Val: Double);
var Texte, sDate: string;
  CC: TWinControl;
  ACol, ARow: Integer;
begin
  if BoutonEncours then Exit;
  BoutonEncours := True;
  Texte := FloatToStr(Val);
  CC := Screen.ActiveControl;
  if CC is THGrid then
  begin
    ACol := THGrid(CC).Col;
    ARow := THGrid(CC).Row;
    THGrid(CC).Cells[ACol, ARow] := Texte;
  end else
    if CC is THCritMaskEdit then
  begin
    if THCritMaskEdit(CC).OpeType = otDate then
    begin
      // on attend une date avec des séparateurs
      sDate := '';
      if (Length(Texte) = 3) or (Length(Texte) = 5) then Texte := '0' + Texte;
      if Length(Texte) = 4 then Texte := Texte + FODonneAn(0);
      if Length(Texte) >= 6 then
      begin
        sDate := Copy(Texte, 1, 2) + DateSeparator + Copy(Texte, 3, 2) + DateSeparator;
        if Length(Texte) >= 8 then sDate := sDate + Copy(Texte, 5, 4)
        else sDate := sDate + Copy(Texte, 5, 2);
      end;
      if IsValidDate(sDate) then THCritMaskEdit(CC).Text := sDate;
    end else
    begin
      THCritMaskEdit(CC).Text := Texte;
    end;
  end else
    if CC is TCustomEdit then
  begin
    TCustomEdit(CC).Text := Texte;
  end else
    if CC is THValComboBox then
  begin
    THValComboBox(CC).Value := Texte;
  end;
  FOSimuleClavier(VK_TAB);
  BoutonEncours := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AppliqueTotal : recopie du total (détail ou pièces/billet) sur la grille du fond de caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.AppliqueTotal;
begin
  if ApplicEncours then Exit;
  if not PPIECBIL.Visible then Exit;
  ApplicEncours := True;
  if (P.ActivePage.Name = 'PfondCaisse') then
  begin
    GS.Cells[GRD_MONTANT, GS.Row] := StrS(CUMMONTANT.Value, V_PGI.OkDecV);
    if GS.CanFocus then GS.SetFocus;
    FOSimuleClavier(VK_TAB);
  end;
  ApplicEncours := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GotoEntete : se positionne sur la grille (fond de caisse) de l'en-tête
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GotoEntete;
begin
  if (P.ActivePage.Name <> 'PfondCaisse') then Exit;
  if FOGetParamCaisse('GPK_CLAVIERECRAN') <> 'X' then
  begin
    PPIECBIL.SendToBack;
    PPIECBIL.Visible := False;
    BBILLET.Visible := True;
    BENTETE.Visible := False;
    bSuivant.Enabled := True;
  end;
  if GS.CanFocus then GS.SetFocus;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GotoPied : se positionne sur la grille (détail ou billet) de pied
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.GotoPied;
begin
  if (P.ActivePage.Name <> 'PfondCaisse') then Exit;
  if (TOBBilMdp = nil) and (TOBPieMdp = nil) then Exit;
  if FOGetParamCaisse('GPK_CLAVIERECRAN') <> 'X' then
  begin
    PPIECBIL.Visible := True;
    PPIECBIL.BringToFront;
    BBILLET.Visible := False;
    BENTETE.Visible := True;
    bSuivant.Enabled := False;
  end;
  if not PPIECBIL.Visible then Exit;
  if GBILLET.CanFocus then
  begin
    GBILLET.Row := GBILLET.FixedRows;
    GBILLET.Col := COLB_QUANTITE;
    GBILLET.SetFocus;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BEnteteClick : OnClick du bouton BENTETE
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.BEnteteClick(Sender: TObject);
begin
  AppliqueTotal;
  GotoEntete;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BBilletClick : OnClick du bouton BBILLET
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.BBilletClick(Sender: TObject);
begin
  GotoPied;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BLEFTClick : OnClick sur le bouton BLEFT
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.BLEFTClick(Sender: TObject);
begin
  inherited;
  FOSimuleClavier(VK_TAB, True);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BRIGHTClick : OnClick sur le bouton BRIGHT
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.BRIGHTClick(Sender: TObject);
begin
  inherited;
  FOSimuleClavier(VK_TAB, False);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BUPClick : OnClick sur le bouton BUP
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.BUPClick(Sender: TObject);
begin
  inherited;
  FOSimuleClavier(VK_UP, False);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BDOWNClick : OnClick sur le bouton BDOWN
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.BDOWNClick(Sender: TObject);
begin
  inherited;
  FOSimuleClavier(VK_DOWN, False);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  bAnnulerClick : OnClick sur le bouton bAnnuler
///////////////////////////////////////////////////////////////////////////////////////

procedure TFOuvertJour.bAnnulerClick(Sender: TObject);
begin
  inherited;
  NormalClose := False;
end;

end.
