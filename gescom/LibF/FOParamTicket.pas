{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2002
Modifié le ... : 23/07/2002
Description .. : Paramétrage de la fiche de saisie des tickets
Mots clefs ... :
*****************************************************************}
unit FOParamTicket;

interface
uses
  Forms, Dialogs, hmsgbox, HSysMenu, ExtCtrls, StdCtrls, Spin, Hctrls,
  HTB97, HFLabel, Grids, Mask, LCD_Lab, HPanel, Controls, Classes,
  SysUtils, AGLInit, M3FP,
  {$IFDEF EAGLCLIENT}
  UTob, MaineAGL,
  {$ELSE}
  dbtables, FE_Main, Prefs,
  {$ENDIF}
  Graphics, HEnt1, KB_Ecran, TickUtilFO, Menus;

type R_IdentCol = record
    ColName, ColTyp: string;
  end;

  // Gestion des afficheurs
type TQuoiAffiche = (qaLibreTexte, qaLibreMontant, qaTotal, qaLigne, qaTimer);
  TOuAffiche = (ofInterne, ofClient, ofTimer);
  TSetOuAffiche = set of TOuAffiche;

type
  TFParamTicket = class(TForm)
    DockBottom: TDock97;
    Outils97: TToolbar97;
    BpropPanel: TToolbarButton97;
    Valide97: TToolbar97;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    PEntete: THPanel;
    HGP_DATEPIECE: THLabel;
    HGP_TIERS: THLabel;
    fAffTXT: TLCDLabel;
    GP_NUMEROPIECE: THLabel;
    HGP_REPRESENTANT: THLabel;
    GP_DATEPIECE: THCritMaskEdit;
    GP_TIERS: THCritMaskEdit;
    GP_REPRESENTANT: THCritMaskEdit;
    PMain: TPanel;
    PImage: THPanel;
    TimerLCD: TTimer;
    TimerGen: TTimer;
    HMTrad: THSystemMenu;
    HPiece: THMsgBox;
    HTitres: THMsgBox;
    HErr: THMsgBox;
    TWPropPanel: TToolWindow97;
    BValidePropPanel: TToolbarButton97;
    BCancelPropPanel: TToolbarButton97;
    TFCoulorStart: THLabel;
    TFLigne: THLabel;
    TFBackGround: THLabel;
    CCBackGround: THValComboBox;
    ChoixCouleur: TColorDialog;
    TWPropLCD: TToolWindow97;
    BValidePropLCD: TToolbarButton97;
    BCancelPropLCD: TToolbarButton97;
    TFColorFond: THLabel;
    TFPixelON: THLabel;
    TFPixelOFF: THLabel;
    BPropLCD: TToolbarButton97;
    TFGSColor: THLabel;
    BColorFond: TToolbarButton97;
    BPixelON: TToolbarButton97;
    BPixelOFF: TToolbarButton97;
    BColorStart: TToolbarButton97;
    BColorEnd: TToolbarButton97;
    BGSColor: TToolbarButton97;
    CBTwoColors: TCheckBox;
    TWChoixPave: TToolWindow97;
    BValideChoixPave: TToolbarButton97;
    BCancelChoixPave: TToolbarButton97;
    BClavier: TToolbarButton97;
    BCaisse: TToolbarButton97;
    BChoixPave: TToolbarButton97;
    BPreference: TToolbarButton97;
    CBAffInverse: TCheckBox;
    MsgAccueil: TEdit;
    TFMsgAccueil: THLabel;
    BEche: TToolbarButton97;
    ToolbarSep971: TToolbarSep97;
    ToolbarSep972: TToolbarSep97;
    TWPropCell: TToolWindow97;
    BValidePropCell: TToolbarButton97;
    BCancelPropCell: TToolbarButton97;
    CtrAlign: TRadioGroup;
    RBRien: TRadioButton;
    RBArtFi: TRadioButton;
    RBTypeRemise: TRadioButton;
    RBModePaie: TRadioButton;
    RBVendeur: TRadioButton;
    RBPage: TRadioButton;
    SPNoPage: TSpinEdit;
    GBDIM: TGroupBox;
    TFLargeur: THLabel;
    SPLargeur: TSpinEdit;
    BPropCell: TToolbarButton97;
    BDefaultConfig: TToolbarButton97;
    CBAffClient: TCheckBox;
    TFNbRows: THLabel;
    SPNbRows: TSpinEdit;
    TWOrdreTab: TToolWindow97;
    BValidelOrdreTab: TToolbarButton97;
    BCancelOrdreTab: TToolbarButton97;
    GSOrdTab: THGrid;
    BBas: TToolbarButton97;
    BHaut: TToolbarButton97;
    BOrdTab: TToolbarButton97;
    PGrille: TPanel;
    PClavier2: TPanel;
    PnlCorps: THPanel;
    GS: THGrid;
    GSReg: THGrid;
    PPied: THPanel;
    ImageArticle: TImage;
    LblPrixU: THLabel;
    LblA_LIBELLE: THLabel;
    T_PrixU: TFlashingLabel;
    A_QTESTOCK: TFlashingLabel;
    PI_NETAPAYERDEV: TLCDLabel;
    TPI_NETAPAYERDEV: THLabel;
    LIBELLETIERS: THLabel;
    LIBELLEARTICLE: THLabel;
    LGP_TOTALQTEFACT: TLabel;
    TLGP_TOTALQTEFACT: TLabel;
    POPZ: TPopupMenu;
    mPavePrincipal: TMenuItem;
    mPaveSecond: TMenuItem;
    TWPropPave2: TToolWindow97;
    BValidePropPave2: TToolbarButton97;
    BCancelPropPave2: TToolbarButton97;
    CtrAlignPave2: TRadioGroup;
    GroupBox1: TGroupBox;
    TSpLargPave2: THLabel;
    SpLargPave2: TSpinEdit;
    CBPave2: TCheckBox;
    BPave2: TToolbarButton97;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure TimerGenTimer(Sender: TObject);
    procedure TimerLCDTimer(Sender: TObject);
    procedure PEnteteResize(Sender: TObject);
    procedure PEnteteDblClick(Sender: TObject);
    procedure PPiedDblClick(Sender: TObject);
    procedure BValidePropPanelClick(Sender: TObject);
    procedure BCancelPropPanelClick(Sender: TObject);
    procedure BColorStartClick(Sender: TObject);
    procedure BColorEndClick(Sender: TObject);
    procedure BGSColorClick(Sender: TObject);
    procedure BValidePropLCDClick(Sender: TObject);
    procedure BCancelPropLCDClick(Sender: TObject);
    procedure fAffTXTDblClick(Sender: TObject);
    procedure PI_NETAPAYERDEVDblClick(Sender: TObject);
    procedure BColorFondClick(Sender: TObject);
    procedure BPixelONClick(Sender: TObject);
    procedure BPixelOFFClick(Sender: TObject);
    procedure RadioButtonClick(Sender: TObject);
    procedure BValideChoixPaveClick(Sender: TObject);
    procedure BCancelChoixPaveClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BPropPanelClick(Sender: TObject);
    procedure BPropLCDClick(Sender: TObject);
    procedure BClavierClick(Sender: TObject);
    procedure GP_TIERSEnter(Sender: TObject);
    procedure GP_TIERSDblClick(Sender: TObject);
    procedure GP_REPRESENTANTDblClick(Sender: TObject);
    procedure GP_REPRESENTANTEnter(Sender: TObject);
    procedure BPreferenceClick(Sender: TObject);
    procedure BCaisseClick(Sender: TObject);
    procedure BChoixPaveClick(Sender: TObject);
    procedure BEcheClick(Sender: TObject);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure GSDblClick(Sender: TObject);
    procedure GSEnter(Sender: TObject);
    procedure GSMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure GSRegEnter(Sender: TObject);
    procedure GSRegDblClick(Sender: TObject);
    procedure GSRegCellEnter(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure GSRegMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure BValidePropCellClick(Sender: TObject);
    procedure BCancelPropCellClick(Sender: TObject);
    procedure BPropCellClick(Sender: TObject);
    procedure BDefaultConfigClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BValidelOrdreTabClick(Sender: TObject);
    procedure BCancelOrdreTabClick(Sender: TObject);
    procedure BOrdTabClick(Sender: TObject);
    procedure BHautClick(Sender: TObject);
    procedure BBasClick(Sender: TObject);
    procedure BPave2Click(Sender: TObject);
    procedure BValidePropPave2Click(Sender: TObject);
    procedure BCancelPropPave2Click(Sender: TObject);
  private
    { Déclarations privées }
    IdentCols: array[0..19] of R_IdentCol;
    function ValideSaisie: Boolean;
    procedure ProprieteParDefaut;
    function CaculLargeur(Indice: Integer; ValeurDefaut: Boolean): Integer;
    function IsPropParDefaut(TypeObj, Ind: Integer): Boolean;
    procedure MiseEnFormeFiche(CreationPave: Boolean);
    procedure MiseEnFormePave;
    procedure MiseEnFormePave2;
    procedure ChargeChoixDesPaves(var PropPv2: RPropPave2);
    procedure ChargeFromNature;
    procedure EtudieColsListe;
    procedure FormateColsReg;
    procedure EtudieColsGSReg;
    procedure RetailleGrille(GrilleReglement: Boolean);
    procedure ChargeProprietes;
    procedure InitEnteteDefaut;
    procedure InvertirAfficheur;
    procedure Resizerafficheur;
    function TraiteTexteAff(ST: string): string;
    procedure AffichageLCD(Texte: string; Qte, Mnt: Double; NbDec: Integer; Sigle: string; Ou: TSetOuAffiche; Quoi: TQuoiAffiche; Attente: Boolean);
    procedure LCDVisible(Visible: Boolean);
    procedure BoutonCalculetteClick(Val: string);
    procedure OuvreOrdreTabPanel;
    procedure OuvrePropPanel;
    procedure OuvrePropLCD;
    procedure OuvreChoixPave(ACol: Integer);
    procedure OuvrePropCell(ACol: Integer);
    procedure ActiveOptionsChoixPave;
    procedure AffichePaveChoisi(ACol: Integer);
    procedure TitreFenetre(ACol: Integer; var Fen: TToolWindow97);
    procedure OuvrePropPave2;
  public
    { Déclarations publiques }
    Action: TActionFiche; // action de la fiche
    NaturePiece: string; // nature de la pièce
    InMaximise: Boolean; // maximisation en cours
    MajParam: Boolean; // les paramètres doivent être mise à jour
    PnlBoutons: TClavierEcran; // pavé tactile
    PnlBtn2: TClavierEcran;
    PropPv2: RPropPave2;
    DefPropPv2: RPropPave2;
    LesColonnes: string; // liste des colonnes de la grille
    MesgAfficheur: string; // message de l'afficheur
    QuoiSurLCD: TQuoiAffiche; // Type du message affiché actuellement sur le LCD
    ColCarac: array[0..41] of RColCarac; // Caractéristiques des colonnes
    //    de  0 à 19 => colonnes de la grille des lignes (GS)
    //    de 20 à 39 => colonnes de la grille des échéances (GSReg)
    //    de 40 à 41 => champs d'en-tête (40=Client, 41=Vendeur)
    ColInd: Integer; // Rang dans ColCarac à mettre à jour
    DefColorStart: TColor; // Couleur de début par défaut
    DefColorEnd: TColor; // Couleur de fin par défaut
    DefBackGround: TDirection; // Effet par défaut du fond
    DefColorFond: TColor; // Couleur par défaut des panels
    DefPixelON: TColor; // Couleur par défaut pour le Pixel On
    DefPixelOFF: TColor; // Couleur par défaut pour le Pixel Off
    DefTwoColors: Boolean; // Couleur alternée de la grille
    DefNbRows: integer; // Nombre par défaut de lignes visibles dans la grille
    DefTabOrder: array[0..2] of integer; // Ordre de tabulation par défaut
    DefGSColor: TColor; // Couleur par défaut de l'entête de grille
    DefCarac: array[0..41] of RColCarac; // Caractéristiques par défaut des colonnes
    //    de  0 à 19 => colonnes de la grille des lignes (GS)
    //    de 20 à 39 => colonnes de la grille des échéances (GSReg)
    //    de 40 à 41 => champs d'en-tête (40=Client, 41=Vendeur)
    DefTextWidth: Integer; // Taille d'un caractère
  end;

procedure SaisieParamTicketFO(NaturePiece: string; Action: TActionFiche; Caisse: string = '');

implementation
uses
  EntGC, SAISUTIL, FactUtil, FactComm, FactTOB,
  KB_Param, FOUtil, MC_Lib;

// Colonnes du Grid
const SG_Demarque: integer = -1;
  SG_PxNet: integer = -1;
  SG_MontNet: integer = -1;
  SGR_NL: integer = -1;
  SGR_Mode: integer = -1;
  SGR_Lib: integer = -1;
  SGR_Mont: integer = -1;
  SGR_Dev: integer = -1;
  SGR_Ech: integer = -1;

  // Gestion des Timers
const MaxInterval = 5000;
  MinInterval = 500;
  ToMaximise = 100;

  {$R *.DFM}

  {***********A.G.L.Privé.*****************************************
  Auteur  ...... : N. ACHINO
  Créé le ...... : 31/12/2002
  Modifié le ... : 31/12/2002
  Description .. : Paramétrage de la fiche de saisie des tickets
  Mots clefs ... :
  *****************************************************************}

procedure SaisieParamTicketFO(NaturePiece: string; Action: TActionFiche; Caisse: string = '');
var XX: TFParamTicket;
  OldCais: string;
  {$IFNDEF EAGL}
    DrawXP: boolean;
  {$ENDIF}
begin
  {$IFNDEF EAGL}
    DrawXP:=V_PGI.DrawXP;
    V_PGI.DrawXP:=False;
  {$ENDIF}
  SourisSablier;
  OldCais := FOCaisseCourante;
  if Caisse = '' then Caisse := FOCaisseCourante;
  FOChargeVHGCCaisse(Caisse);
  XX := TFParamTicket.Create(Application);
  XX.Action := Action;
  XX.NaturePiece := NaturePiece;
  try
    XX.ShowModal;
  finally
    XX.Free;
    if OldCais <> FOCaisseCourante then FOChargeVHGCCaisse(OldCais);
  end;
  SourisNormale;
  {$IFNDEF EAGL}
    V_PGI.DrawXP:=DrawXP;
  {$ENDIF}
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormCreate
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.FormCreate(Sender: TObject);
begin
  InitLesCols;
  SG_Demarque := -1;
  SG_PxNet := -1;
  SG_MontNet := -1;
  SGR_NL := -1;
  SGR_Mode := -1;
  SGR_Lib := -1;
  SGR_Mont := -1;
  SGR_Dev := -1;
  SGR_Ech := -1;
  FillChar(IdentCols, Sizeof(IdentCols), #0);
  InMaximise := FALSE;
  MajParam := FALSE;
  FOInitColCarac(ColCarac);
  FOInitColCarac(DefCarac);
  if not (IsInside(Self)) and (FOGetParamCaisse('GPK_TOOLBAR') = 'X') then FoShowEtiquette(Self, True);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormShow
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.FormShow(Sender: TObject);
begin
  if InMaximise then Exit;
  // Mise en forme de la fiche
  MiseEnFormeFiche(True);
  ProprieteParDefaut;
  ChargeChoixDesPaves(PropPv2);
  MiseEnFormePave2;
  RetailleGrille(False);
  FOAppliqueColCarac(GSReg, ColCarac, 20);
  RetailleGrille(True);
  ChargeProprietes;
  // Appel de la fonction d'empilage dans la liste des fiches
  AglEmpileFiche(Self);
  // Maximisation de la forme
  if not TimerGen.Enabled then
  begin
    TimerGen.Interval := ToMaximise;
    TimerGen.Enabled := True;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormDestroy
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.FormDestroy(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormResize
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.FormResize(Sender: TObject);
var Position: Integer;
begin
  PMain.Tag := PMain.Height - PGrille.Height;
  // Redimensionnent du LCD client
  ResizerAfficheur;
  // Redimensionnent du LCD total général
  PI_NETAPAYERDEV.left := pPied.width - PI_NETAPAYERDEV.Width - 4;
  TPI_NETAPAYERDEV.Left := PI_NETAPAYERDEV.left - TPI_NETAPAYERDEV.width - 8;
  // Affichage de la photo de l'article
  if (FOAfficheImageArt) or (FOLogoCaisse <> '') then
  begin
    Position := (ImageArticle.Left * 2) + ImageArticle.Width;
    LIBELLETIERS.Left := Position;
    LblPrixU.Left := Position;
    LblA_LIBELLE.Left := Position;
    LIBELLEARTICLE.Left := Position;
  end;
  // Redimensionnent des libellés du tiers et de l'article
  LIBELLETIERS.Width := PI_NETAPAYERDEV.Left - LIBELLETIERS.Left - 6;
  LIBELLEARTICLE.Width := PI_NETAPAYERDEV.Left - LIBELLEARTICLE.Left - 6;
  // Repositionnement de la quantité totale et du total en contre-valeur
  TLGP_TOTALQTEFACT.Left := PI_NETAPAYERDEV.Left;
  LGP_TOTALQTEFACT.Left := TLGP_TOTALQTEFACT.Left + TLGP_TOTALQTEFACT.Width + 6;
  // Redimensionnent des boutons tactiles
  if assigned(PnlBoutons) then PnlBoutons.ResizeClavier(nil);
  // Redimensionnent de l'image de base d'écran
  {**
  if not FOExisteClavierEcran then
     BEGIN
     // Visualisation des titres et de 12 lignes dans la grille
     Taille:=PMain.Height-PPied.Height-((GS.DefaultRowHeight+1)*13)-3 ;
     if Taille < 1 then Taille:=1 ;
     PImage.Height:=Taille ;
     Top:=(PImage.Height div 3)-SigleCegid.Height ;
     if Top < 10 then Top:=10 ;
     SigleCegid.Top:=Top ;
     Top:=2*(PImage.Height div 3) ;
     if Top < (SigleCegid.Top+SigleCegid.Height) then Top:=PImage.Height-NomAppli.Height ;
     NomAppli.Top:=Top ;
     END ;
  **}
  RetailleGrille(False);
  RetailleGrille(True);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormCloseQuery
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var Rep: Integer;
begin
  if MajParam then
  begin
    Rep := PGIAskCancel('Voulez-vous enregistrer les modifications ?', Caption);
    case Rep of
      mrYes: CanClose := (ValideSaisie);
      mrNo: ;
    else CanClose := False;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ValideSaisie
///////////////////////////////////////////////////////////////////////////////////////

function TFParamTicket.ValideSaisie: Boolean;
var
  Stg: string;
  StgLst: TStrings;
  Ind: Integer;
begin
  Result := True;
  if CBAffClient.Checked <> (FOGetParamCaisse('GPK_AFFICHEUR') = 'X') then
  begin
    if CBAffClient.Checked then
    begin
      Stg := 'X';
    end else
    begin
      Stg := '-';
      MsgAccueil.Text := '';
      CBAffInverse.Checked := False;
    end;
    VH_GC.TOBPCaisse.PutValue('GPK_AFFICHEUR', Stg);
  end;
  if MsgAccueil.Text <> FOGetParamCaisse('GPK_AFFMESG') then
  begin
    VH_GC.TOBPCaisse.PutValue('GPK_AFFMESG', Trim(MsgAccueil.Text));
  end;
  if CBAffInverse.Checked <> (FOGetParamCaisse('GPK_AFFINVERSE') = 'X') then
  begin
    if CBAffInverse.Checked then Stg := 'X' else Stg := '-';
    VH_GC.TOBPCaisse.PutValue('GPK_AFFINVERSE', Stg);
  end;
  // Enregistrement de la présentation choisie
  StgLst := TStringList.Create;
  with StgLst do
  begin
    Clear;
    Add('[' + NaturePiece + ']');
    // propriétés des panels
    if not IsPropParDefaut(1, 0) then
    begin
      Add('PANEL=' + IntToStr(Ord(BColorStart.Color)) + ';'
        + IntToStr(Ord(BColorEnd.Color)) + ';'
        + IntToStr(CCBackGround.ItemIndex) + ';');
    end;
    // propriétés des LCD
    if not IsPropParDefaut(2, 0) then
    begin
      Add('LCD=' + IntToStr(Ord(BColorFond.Color)) + ';'
        + IntToStr(Ord(BPixelON.Color)) + ';'
        + IntToStr(Ord(BPixelOFF.Color)) + ';');
    end;
    // propriétés des grilles
    if not IsPropParDefaut(3, 0) then
    begin
      if CBTwoColors.Checked then Stg := 'X' else Stg := '-';
      Add('GRID=' + IntToStr(Ord(BGSColor.Color)) + ';'
        + Stg + ';'
        + IntToStr(SPNbRows.Value) + ';');
    end;
    // ordre de tabulation en-tête
    if not IsPropParDefaut(6, 0) then
    begin
      Add('ETABORD=' + IntToStr(GP_TIERS.TabOrder) + ';'
        + IntToStr(GP_REPRESENTANT.TabOrder) + ';'
        + IntToStr(GP_DATEPIECE.TabOrder) + ';');
    end;
    // choix des pavés
    for Ind := Low(ColCarac) to High(ColCarac) do
    begin
      if not IsPropParDefaut(4, Ind) then
      begin
        Add('P' + IntToStr(Ind) + '=' + IntToStr(ColCarac[Ind].NoPage));
      end;
    end;
    // choix des propriétés des colonnes
    for Ind := Low(ColCarac) to High(ColCarac) do
    begin
      if not IsPropParDefaut(5, Ind) then
      begin
        Add('C' + IntToStr(Ind) + '=' + IntToStr(CaculLargeur(Ind, False)) + ';'
          + IntToStr(Ord(ColCarac[Ind].Align)) + ';');
      end;
    end;
    // propriétés du pavé secondaire
    if not IsPropParDefaut(7, 0) then
    begin
      Add('PAVE2=' + TrueFalseSt(PropPv2.Visible) + ';'
        + IntToStr(PropPv2.Largeur) + ';'
        + IntToStr(Ord(PropPv2.Align)) + ';'
        + IntToStr(PropPv2.NbrBtnWidth) + ';'
        + IntToStr(PropPv2.NbrBtnHeight) + ';'
        + IntToStr(Ord(PropPv2.ClcPosition)) + ';'
        + TrueFalseSt(PropPv2.ClcVisible) + ';');
    end;
  end;
  if VH_GC.TOBPCaisse.FieldExists('GPK_INFOS') then VH_GC.TOBPCaisse.PutValue('GPK_INFOS', StgLst.Text);
  StgLst.Free;
  // mise à jour de la table des caisses
  if VH_GC.TOBPCaisse.UpdateDB then MajParam := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ProprieteParDefaut
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.ProprieteParDefaut;
var Ind: Integer;
begin
  // Propriétés des panels
  DefColorStart := PPied.ColorStart;
  DefColorEnd := PPied.ColorEnd;
  DefBackGround := PPied.BackGroundEffect;
  // Propriétés des grilles
  DefGSColor := GS.FixedColor;
  DefTwoColors := GS.TwoColors;
  DefNbRows := (PGrille.Height - 114) div GS.RowHeights[1];
  if not FOExisteClavierEcran then DefNbRows := 12;
  // Ordre de tabulation
  DefTabOrder[0] := GP_TIERS.TabOrder;
  DefTabOrder[1] := GP_REPRESENTANT.TabOrder;
  DefTabOrder[2] := GP_DATEPIECE.TabOrder;
  // Propriétés des LCD
  DefColorFond := fAffTXT.BackGround;
  DefPixelON := fAffTXT.PixelOn;
  DefPixelOFF := fAffTXT.PixelOff;
  // Caractéristiques par défaut des colonnes
  FOInitColCarac(DefCarac);
  for Ind := 0 to GS.ColCount - 1 do
  begin
    DefCarac[Ind].Largeur := GS.ColWidths[Ind];
    DefCarac[Ind].Align := GS.ColAligns[Ind];
  end;
  for Ind := 0 to GSReg.ColCount - 1 do
  begin
    DefCarac[20 + Ind].Largeur := GSReg.ColWidths[Ind];
    DefCarac[20 + Ind].Align := GSReg.ColAligns[Ind];
  end;
  DefTextWidth := GS.Canvas.TextWidth('W');
  FillChar(DefPropPv2, Sizeof(DefPropPv2), #0);
  DefPropPv2.Visible := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  CaculLargeur : convertit la largeur d'une colonne de grille en largeur de liste
///////////////////////////////////////////////////////////////////////////////////////

function TFParamTicket.CaculLargeur(Indice: Integer; ValeurDefaut: Boolean): Integer;
var MaxWidth, Largeur, Deb, Fin: Integer;
begin
  Result := 0;
  if Indice > 39 then Exit;
  if ValeurDefaut then
  begin
    Largeur := DefCarac[Indice].Largeur;
    Result := Round(Largeur / DefTextWidth);
  end else
  begin
    Largeur := ColCarac[Indice].Largeur;
    MaxWidth := 0;
    if Indice < 20 then Deb := 0 else Deb := 20;
    Fin := Deb + 19;
    for Indice := Deb to Fin do
      if MaxWidth < ColCarac[Indice].Largeur then MaxWidth := ColCarac[Indice].Largeur;
    Result := Round(((Largeur * 100) div MaxWidth) / 6.3);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  IsPropParDefaut
///////////////////////////////////////////////////////////////////////////////////////

function TFParamTicket.IsPropParDefaut(TypeObj, Ind: Integer): Boolean;
begin
  Result := True;
  case TypeObj of
    1:
      begin // Propriétés des panels
        if (DefColorStart <> BColorStart.Color) or
          (DefColorEnd <> BColorEnd.Color) or
          (Ord(DefBackGround) <> CCBackGround.ItemIndex) then Result := False;
      end;
    2:
      begin // propriétés des LCD
        if (DefColorFond <> BColorFond.Color) or
          (DefPixelON <> BPixelON.Color) or
          (DefPixelOFF <> BPixelOFF.Color) then Result := False;
      end;
    3:
      begin // propriétés des grilles
        if (DefTwoColors <> CBTwoColors.Checked) or
          (DefGSColor <> BGSColor.Color) or
          (DefNbRows <> SPNbRows.Value) then Result := False;
      end;
    4:
      begin // choix des pavés
        if DefCarac[Ind].NoPage <> ColCarac[Ind].NoPage then Result := False;
      end;
    5:
      begin // choix des pavés et propriétés des colonnes
        if (CaculLargeur(Ind, True) <> CaculLargeur(Ind, False)) or
          (DefCarac[Ind].Align <> ColCarac[Ind].Align) then Result := False;
      end;
    6:
      begin // ordre de tabulation
        if (DefTabOrder[0] <> GP_TIERS.TabOrder) or
          (DefTabOrder[1] <> GP_REPRESENTANT.TabOrder) or
          (DefTabOrder[2] <> GP_DATEPIECE.TabOrder) then Result := False;
      end;
    7:
      begin // 2ème pavé
        if (DefPropPv2.Visible <> PropPv2.Visible) or
          (DefPropPv2.Largeur <> PropPv2.Largeur) or
          (DefPropPv2.Align <> PropPv2.Align) or
          (DefPropPv2.NbrBtnWidth <> PropPv2.NbrBtnWidth) or
          (DefPropPv2.NbrBtnHeight <> PropPv2.NbrBtnHeight) or
          (DefPropPv2.ClcPosition <> PropPv2.ClcPosition) or
          (DefPropPv2.ClcVisible <> PropPv2.ClcVisible) then Result := False;
      end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MiseEnFormeFiche
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.MiseEnFormeFiche(CreationPave: Boolean);
var St: string;
  Ind: Integer;
begin
  {$IFDEF EAGLCLIENT}
  BPreference.Visible := False;
  {$ENDIF}
  ChargeFromNature;
  EtudieColsListe;
  InitEnteteDefaut;
  AffecteGrid(GS, Action);
  if Action = taConsult then GS.MultiSelect := False;
  // Suppression de la colonne N° de ligne
  if SG_NL <> -1 then
  begin
    GS.ColWidths[SG_NL] := 0;
  end;
  // Suppression du champ Client sur l'en-tête
  GP_TIERS.Visible := (VH_GC.TOBPCaisse.GetValue('GPK_CLISAISIE') = 'X');
  HGP_TIERS.Visible := GP_TIERS.Visible;
  LIBELLETIERS.Visible := GP_TIERS.Visible;
  // Suppression du champ Vendeur sur l'en-tête
  GP_REPRESENTANT.Visible := (VH_GC.TOBPCaisse.GetValue('GPK_VENDSAISIE') = 'X');
  HGP_REPRESENTANT.Visible := GP_REPRESENTANT.Visible;
  // Suppression de la colonne Vendeur
  if (VH_GC.TOBPCaisse.GetValue('GPK_VENDSAISLIG') = '-') and (SG_Rep <> -1) then
  begin
    GS.ColWidths[SG_Rep] := -1;
  end;
  // Suppression des Remises
  if (VH_GC.TOBPCaisse.GetValue('GPK_REMSAISIE') = '-') or (VH_GC.TOBPCaisse.GetValue('GPK_REMAFFICH') = '-') then
  begin
    if SG_Rem <> -1 then
    begin
      GS.ColWidths[SG_Rem] := -1;
      if VH_GC.TOBPCaisse.GetValue('GPK_REMSAISIE') = '-' then SG_Rem := -1;
    end;
    if SG_RV <> -1 then
    begin
      GS.ColWidths[SG_RV] := -1;
      if VH_GC.TOBPCaisse.GetValue('GPK_REMSAISIE') = '-' then SG_RV := -1;
    end;
    if SG_RL <> -1 then
    begin
      GS.ColWidths[SG_RL] := -1;
      if VH_GC.TOBPCaisse.GetValue('GPK_REMSAISIE') = '-' then SG_RL := -1;
    end;
  end;
  // Suppression de la colonne Démarque
  if (not FOGereDemarque(False)) and (SG_Demarque <> -1) then
  begin
    GS.ColWidths[SG_Demarque] := -1;
    SG_Demarque := -1;
  end;
  // Ecran tactile
  if CreationPave then MiseEnFormePave;
  // Afficheur client
  if FOGetParamCaisse('GPK_AFFICHEUR') <> 'X' then LCDVisible(False);
  // Photo de l'article
  if (FOAfficheImageArt) or (FOLogoCaisse <> '') then
  begin
    ImageArticle.Visible := True;
    FOAffichePhotoArticle('', ImageArticle);
  end else ImageArticle.Visible := False;
  // Grille des échéances
  GSReg.Ctl3D := GS.Ctl3D;
  GSReg.FixedColor := GS.FixedColor;
  GSReg.TwoColors := GS.TwoColors;
  GSReg.Font := GS.Font;
  FormateColsReg;
  AffecteGrid(GSReg, Action);
  // Pour mieux visualiser le cadrage des colonnes
  for Ind := 0 to GS.ColCount - 1 do
  begin
    if Ind = SG_RefArt then St := 'xxxxx' else
      if Ind = SG_Lib then St := 'xxxxxxxxxx' else
      if (Ind = SG_Px) or (Ind = SG_RV) or (Ind = SG_RL) or (Ind = SG_Montant) then St := 'xxx,xx' else
      St := 'xxx';
    GS.Cells[Ind, GS.FixedRows] := St;
  end;
  for Ind := 0 to GSReg.ColCount - 1 do
  begin
    if Ind = SGR_Lib then St := 'xxxxxxxxxx' else
      if Ind = SGR_Mont then St := 'xxx,xx' else
      if Ind = SGR_Ech then St := 'xx/xx/xx' else
      St := 'xxx';
    GSReg.Cells[Ind, GSReg.FixedRows] := St;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MiseEnFormePave
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.MiseEnFormePave;
begin
    if FOExisteClavierEcran then
    begin
      PImage.Visible := False;
      PMain.Tag := PMain.Height - PGrille.Height;
      FOCreateClavierEcran(PnlBoutons, Self, Pmain);
      PnlBoutons.BoutonCalculette := BoutonCalculetteClick;
    end else
    begin
      if Assigned(PnlBoutons) then FreeAndNil(PnlBoutons);
      PImage.Visible := True;
      PImage.Align := alClient;
    end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MiseEnFormePave2
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.MiseEnFormePave2;
begin
  if PropPv2.Visible then
  begin
    PClavier2.Visible := True;
    PClavier2.Width := PropPv2.Largeur;
    if PropPv2.Align = taLeftJustify then
      PClavier2.Align := alLeft
    else
      PClavier2.Align := alRight;
    PClavier2.Tag := 0;
    if Assigned(PnlBtn2) then FreeAndNil(PnlBtn2);
    PnlBtn2 := TClavierEcran.Create(Self, 1);
    PnlBtn2.Parent := PClavier2;
    PnlBtn2.Align := alBottom;
    PnlBtn2.Height := PClavier2.Tag;
    PnlBtn2.colornb := 6;
    PnlBtn2.LanceBouton := nil;
    PnlBtn2.LanceCalculette := nil;
    PnlBtn2.BoutonCalculette := BoutonCalculetteClick;
    PnlBtn2.Caisse := FOAlphaCodeNumeric(FOCaisseCourante);
    PnlBtn2.NbrBtnWidth := PropPv2.NbrBtnWidth;
    PnlBtn2.NbrBtnHeight := PropPv2.NbrBtnHeight;
    PnlBtn2.ClcPosition := PropPv2.ClcPosition;
    PnlBtn2.ClcVisible := PropPv2.ClcVisible;

    BClavier.DropdownArrow := True;
    BClavier.DropdownMenu := POPZ;
    BClavier.OnClick := nil;
  end else
  begin
    PClavier2.Visible := False;
    PClavier2.Align := alNone;
    if Assigned(PnlBtn2) then FreeAndNil(PnlBtn2);
    BClavier.DropdownArrow := False;
    BClavier.DropdownMenu := nil;
    BClavier.OnClick := BClavierClick;
  end;
  BPave2.Visible := True;
  RetailleGrille(False);
  RetailleGrille(True);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeChoixDesPaves
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.ChargeChoixDesPaves(var PropPv2: RPropPave2);
begin
  ChargePaveContextuel([PEntete, PPied], [GS], [fAffTXT, PI_NETAPAYERDEV], ColCarac, [GP_TIERS, GP_REPRESENTANT, GP_DATEPIECE], PropPv2);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeFromNature
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.ChargeFromNature;
begin
  GS.ListeParam := GetInfoParPiece(NaturePiece, 'GPP_LISTESAISIE');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  EtudieColsListe
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.EtudieColsListe;
var i: integer;
  Nam, St: string;
begin
  LesColonnes := GS.Titres[0];
  if VH_GC.TOBPCaisse.GetValue('GPK_APPELPRIXTIC') = 'TTC' then
  begin
    LesColonnes := FindEtReplace(LesColonnes, 'GL_PUHTDEV', 'GL_PUTTCDEV', False);
    LesColonnes := FindEtReplace(LesColonnes, 'GL_PUHTNETDEV', 'GL_PUTTCNETDEV', False);
    LesColonnes := FindEtReplace(LesColonnes, 'GL_TOTALHTDEV', 'GL_TOTALTTCDEV', False);
    LesColonnes := FindEtReplace(LesColonnes, 'GL_MONTANTHTDEV', 'GL_MONTANTTTCDEV', False);
  end;
  EtudieColsGrid(GS, True);
  St := LesColonnes;
  for i := 0 to GS.ColCount - 1 do
  begin
    Nam := ReadTokenSt(St);
    IdentCols[i].ColName := Nam;
    IdentCols[i].ColTyp := ChampToType(Nam);
    if IdentCols[i].ColTyp = 'DATE' then
    begin
      GS.ColTypes[i] := 'D';
      GS.ColFormats[i] := ShortdateFormat;
    end;
    if Nam = 'GL_TYPEREMISE' then
    begin
      ///GS.ColFormats[i]:='CB=GCTYPEREMISE' ;
      SG_Demarque := i;
    end;
    if (Nam = 'GL_MONTANTHTDEV') or (Nam = 'GL_MONTANTTTCDEV') then SG_MontNet := i;
    if (Nam = 'GL_PUHTNETDEV') or (Nam = 'GL_PUTTCNETDEV') then SG_PxNet := i;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FormateColsReg : définit les caractéristiques de chaque colonne de la grille GSReg
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.FormateColsReg;
var Ind: Integer;
begin
  GSReg.ListeParam := 'GCREGLEFO';
  LesColonnes := GSReg.Titres[0];
  LesColonnes := FindEtReplace(LesColonnes, '(GPE_LIBELLE)', 'GPE_LIBELLE', False);
  EtudieColsGSReg;
  // Code devise
  if SGR_Dev <> -1 then GSReg.ColFormats[SGR_Dev] := 'CB=TTDEVISETOUTES';
  // Date d'échéance
  if SGR_Ech <> -1 then
  begin
    GSReg.ColTypes[SGR_Ech] := 'D';
    GSReg.ColFormats[SGR_Ech] := ShortDateFormat;
  end;
  // Uniquement le mode de paiement, le montant et la date d'échéance sont saisissables
  for Ind := 0 to GSReg.ColCount - 1 do
  begin
    if (Ind <> SGR_Mode) and (Ind <> SGR_Mont) and (Ind <> SGR_Ech) then GSReg.ColLengths[Ind] := -1;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  EtudieColsGSReg : recherche le champ associé à chaque colonne
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.EtudieColsGSReg;
var NomCol, LesCols: string;
  icol, ichamp, iTableLigne: integer;
begin
  LesCols := GSReg.Titres[0];
  icol := 0;
  iTableLigne := PrefixeToNum('GPE');
  repeat
    NomCol := Uppercase(Trim(ReadTokenSt(LesCols)));
    if NomCol <> '' then
    begin
      ichamp := ChampToNum(NomCol);
      if ichamp >= 0 then
      begin
        if Pos('X', V_PGI.Dechamps[iTableLigne, ichamp].Control) > 0 then
        begin
          // "X" interdit la saisie sauf si aussi "Y" pour FrontOffice
          if ((Pos('Y', V_PGI.Dechamps[iTableLigne, ichamp].Control) > 0) and (ctxFO in V_PGI.PGIContexte))
            then else GSReg.ColLengths[icol] := -1;
        end;
        if NomCol = 'GPE_NUMECHE' then SGR_NL := icol else
          if NomCol = 'GPE_MODEPAIE' then SGR_Mode := icol else
          if NomCol = 'GPE_MONTANTENCAIS' then SGR_Mont := icol else
          if NomCol = 'GPE_DEVISEESP' then SGR_Dev := icol else
          if NomCol = 'GPE_DATEECHE' then SGR_Ech := icol else
          ;
      end else
      begin
        if NomCol = '(GPE_LIBELLE)' then SGR_Lib := icol else
      end;
    end;
    Inc(icol);
  until ((LesCols = '') or (NomCol = ''));
end;

///////////////////////////////////////////////////////////////////////////////////////
//  RetailleGrille
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.RetailleGrille(GrilleReglement: Boolean);
var Ind, Deb: Integer;
  Grille: THGrid;
begin
  if GrilleReglement then
  begin
    Grille := GSReg;
    Deb := 20;
  end else
  begin
    Grille := GS;
    Deb := 0;
  end;
  HMTrad.ResizeGridColumns(Grille);
  for Ind := 0 to Grille.ColCount - 1 do
  begin
    ColCarac[Deb + Ind].Largeur := Grille.ColWidths[Ind];
    ColCarac[Deb + Ind].Align := Grille.ColAligns[Ind];
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChargeProprietes
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.ChargeProprietes;
begin
  // Propriétés des panels
  BColorStart.Color := PPied.ColorStart;
  BColorEnd.Color := PPied.ColorEnd;
  CCBackGround.ItemIndex := Ord(PPied.BackGroundEffect);
  // Propriétés des grilles
  BGSColor.Color := GS.FixedColor;
  CBTwoColors.Checked := GS.TwoColors;
  SPNbRows.Value := (PGrille.Height - 114) div GS.RowHeights[1];
  // Propriétés des LCD
  BColorFond.Color := fAffTXT.BackGround;
  BPixelON.Color := fAffTXT.PixelOn;
  BPixelOFF.Color := fAffTXT.PixelOff;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  InitEnteteDefaut
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.InitEnteteDefaut;
begin
  GP_DATEPIECE.Text := DateToStr(V_PGI.DateEntree);
  // Propriétés des panels, des grilles et des LCD
  ChargeProprietes;
  // Message par défaut sur l'afficheur
  InvertirAfficheur;
  MesgAfficheur := Trim(VH_GC.TOBPCaisse.GetValue('GPK_AFFMESG'));
  if MesgAfficheur = '' then MesgAfficheur := 'Bienvenue';
  MesgAfficheur := MesgAfficheur + ' ';
  AffichageLCD(MesgAfficheur, 0, 0, 0, '', [ofInterne, ofClient], qaLibreTexte, False);
  MsgAccueil.Text := MesgAfficheur;
  CBAffInverse.Checked := fAffTXT.UpSideDown;
  CBAffClient.Checked := fAffTXT.Visible;
  if (not TimerLCD.Enabled) and (TimerLCD.Interval <> MaxInterval) then TimerLCD.Interval := MaxInterval;
  TimerLCD.Enabled := True;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TimerGenTimer : Maximisation de la forme
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.TimerGenTimer(Sender: TObject);
begin
  if not TimerGen.Enabled then exit;
  if TimerGen.interval = ToMaximise then
  begin
    InMaximise := True;
    TimerGen.Interval := MaxInterval;
    BorderStyle := bsNone;
    Caption := '';
    BorderIcons := [];
    Visible := False;
    WindowState := wsMaximized;
    Visible := True;
    InMaximise := False;
    TimerGen.Enabled := False;
    // positionnement des barres de boutons
    Valide97.DragHandleStyle := dhNone;
    Valide97.BorderStyle := bsNone;
    Valide97.DockPos := DockBottom.Width - Valide97.Width;
    Outils97.DragHandleStyle := dhNone;
    Outils97.BorderStyle := bsNone;
    Outils97.DockPos := 0;
    // pré-positionnement dans la grille
    if VH_GC.TOBPCaisse.GetValue('GPK_PREPOSGRID') = 'X' then if GS.CanFocus then GS.SetFocus;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TimerLCDTimer
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.TimerLCDTimer(Sender: TObject);
var Stg: string;
begin
  if not TimerLCD.Enabled then exit;
  Stg := FODecaleGauche(fAffTXT.Caption, 1, fAffTXT.NoOfChars);
  AffichageLCD(Stg, 0, 0, 0, '', [ofInterne], qaLibreTexte, False);
  if TimerLCD.Interval <> MinInterval then TimerLCD.Interval := MinInterval;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  PEnteteResize
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.PEnteteResize(Sender: TObject);
begin
  ResizerAfficheur;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  InvertirAfficheur : place l'afficheur en mode inversé
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.InvertirAfficheur;
begin
  fAffTXT.Caption := '';
  FAffTXT.UpSideDown := (VH_GC.TOBPCaisse.GetValue('GPK_AFFINVERSE') = 'X');
  //FAffTXT.Visible:=(VH_GC.TOBPCaisse.GetValue('GPK_AFFICHEUR') = 'X') ;
  LCDVisible((VH_GC.TOBPCaisse.GetValue('GPK_AFFICHEUR') = 'X'));
  fAffTXT.Invalidate;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  Resizerafficheur : ajuste la taille de l'afficheur
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.Resizerafficheur;
begin
  fAffTXT.Width := PEntete.Width - 2 * fAffTXT.left;
  fAffTXT.Left := (PEntete.Width - fAffTXT.Width) div 2;
  fAffTXT.top := 4;
  fAffTXT.CalcCharSize;
  fAffTXT.Invalidate;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TraiteTexteAff : convertit un texte pour l'afficheur client
///////////////////////////////////////////////////////////////////////////////////////

function TFParamTicket.TraiteTexteAff(ST: string): string;
var i: integer;
begin
  result := AnsiUppercase(traduirememoire(St));
  for i := 1 to length(Result) do
    case Result[i] of
      'á', 'à', 'â', 'ä', 'Á', 'À', 'Â', 'Ä': Result[i] := 'A';
      'é', 'è', 'ê', 'ë', 'É', 'È', 'Ê', 'Ë': Result[i] := 'E';
      'í', 'ì', 'î', 'ï', 'Í', 'Ì', 'Î', 'Ï': Result[i] := 'I';
      'ó', 'ò', 'ô', 'ö', 'Ó', 'Ò', 'Ô', 'Ö': Result[i] := 'O';
      'ú', 'ù', 'û', 'ü', 'Ú', 'Ù', 'Û', 'Ü': Result[i] := 'U';
      '£': Result[i] := 'L';
      '¥': Result[i] := 'Y';
      /////    '€'                              : Result[i]:='E' ;
    end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AffichageLCD : affiche un texte, une quantité ou un montant sur l'afficheur client
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.AffichageLCD(Texte: string; Qte, Mnt: Double; NbDec: Integer; Sigle: string; Ou: TSetOuAffiche; Quoi: TQuoiAffiche; Attente: Boolean);
var Libelle, Valeur: string;
  Taille, Pos, Nb: Integer;
begin
  if not fAffTXT.Visible then exit;
  if fAffTXT.Tag <> 0 then Delay(2500); //If tag<>0, le dernier message doit être affiché plus temps
  fAffTXT.Tag := ord(Attente);
  QuoiSurLCD := Quoi;
  Valeur := '';
  Libelle := '';
  if not (ofTimer in ou) then Libelle := traitetexteAff(Texte);
  if Qte <> 0 then
  begin
    Taille := fAffTXT.NoOfChars;
    if Length(Libelle) > Taille then
    begin
      // on tronque le texte si on ne dispose pas d'assez de place pour le montant
      Pos := Taille;
      Nb := Length(Libelle) - Taille;
      Delete(Libelle, Pos, Nb);
    end else
    begin
      // on complete par des espaces pour se placer sur la 2ème ligne
      while Length(Libelle) < Taille do Libelle := Libelle + ' ';
    end;
    Libelle := Libelle + StrfMontant(Qte, fAffTXT.noOfChars, V_PGI.OkDecQ, '', True) + ' X '
  end;
  if Mnt <> 0 then Valeur := StrfMontant(Mnt, fAffTXT.noOfChars, NbDec, Sigle, True);
  if Valeur <> '' then
  begin
    Taille := fAffTXT.NoOfChars * fAffTXT.TextLines;
    // on tronque le texte si on ne dispose pas d'assez de place pour le montant
    if Length(Valeur) > (Taille - Length(Libelle) + 1) then
    begin
      Pos := Taille - Length(Valeur);
      Nb := Length(Libelle) - Length(Valeur) + 1;
      Delete(Libelle, Pos, Nb);
    end;
    // cadrage du montant à gauche
    while Length(Valeur) < (Taille - length(Libelle)) do Valeur := ' ' + Valeur;
  end;
  if fAffTXT.Caption <> Libelle + Valeur then fAffTXT.Caption := Libelle + Valeur;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  LCDVisible : rend visible l'afficheur client
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.LCDVisible(Visible: Boolean);
var Ind: Integer;
begin
  if Visible = fAffTXT.Visible then Exit;
  Ind := fAffTXT.Top + fAffTXT.Height;
  if Visible then
  begin
    fAffTXT.Visible := True;
    TimerLCD.Enabled := True;
  end else
  begin
    fAffTXT.Visible := False;
    TimerLCD.Enabled := False;
    Ind := Ind * (-1);
  end;
  PEntete.Height := PEntete.Height + Ind;
  GP_NUMEROPIECE.Top := GP_NUMEROPIECE.Top + Ind;
  HGP_DATEPIECE.Top := HGP_DATEPIECE.Top + Ind;
  GP_DATEPIECE.Top := GP_DATEPIECE.Top + Ind;
  HGP_TIERS.Top := HGP_TIERS.Top + Ind;
  GP_TIERS.Top := GP_TIERS.Top + Ind;
  HGP_REPRESENTANT.Top := HGP_REPRESENTANT.Top + Ind;
  GP_REPRESENTANT.Top := GP_REPRESENTANT.Top + Ind;
  // Redimensionnent des boutons tactiles
  FormResize(nil);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BoutonCalculetteClick : interprète les boutons de la calculette
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BoutonCalculetteClick(Val: string);
begin
  if Val = 'ENTER' then
  begin
    BValiderClick(nil);
  end else
    if Val = 'CLEAR' then
  begin
    BAbandonClick(nil);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OuvreOrdreTabPanel
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BOrdTabClick(Sender: TObject);
begin
  OuvreOrdreTabPanel;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OuvreOrdreTabPanel
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.OuvreOrdreTabPanel;
var ARow, Ind: integer;
  TabCode, TabLib: array[0..2] of string;
begin
  for Ind := Low(TabCode) to High(TabCode) do TabCode[Ind] := '';
  if (GP_TIERS.Visible) and (GP_TIERS.Enabled) then
  begin
    ARow := GP_TIERS.TabOrder;
    if (ARow >= 0) and (ARow <= 2) then
    begin
      TabCode[ARow] := 'GP_TIERS';
      TabLib[ARow] := HGP_TIERS.Caption;
    end;
  end;
  if (GP_REPRESENTANT.Visible) and (GP_REPRESENTANT.Enabled) then
  begin
    ARow := GP_REPRESENTANT.TabOrder;
    if (ARow >= 0) and (ARow <= 2) then
    begin
      TabCode[ARow] := 'GP_REPRESENTANT';
      TabLib[ARow] := HGP_REPRESENTANT.Caption;
    end;
  end;
  if (GP_DATEPIECE.Visible) and (GP_DATEPIECE.Enabled) then
  begin
    ARow := GP_DATEPIECE.TabOrder;
    if (ARow >= 0) and (ARow <= 2) then
    begin
      TabCode[ARow] := 'GP_DATEPIECE';
      TabLib[ARow] := ChampToLibelle('GP_DATEPIECE');
    end;
  end;
  ARow := 0;
  GSOrdTab.RowCount := 3;
  for Ind := Low(TabCode) to High(TabCode) do
    if TabCode[Ind] <> '' then
    begin
      GSOrdTab.Cells[0, ARow] := TabCode[Ind];
      GSOrdTab.Cells[1, ARow] := TabLib[Ind];
      Inc(ARow);
    end;
  GSOrdTab.RowCount := ARow;
  GSOrdTab.ColWidths[0] := -1;
  GSOrdTab.ColLengths[0] := -1;
  HMTrad.ResizeGridColumns(GSOrdTab);
  CentreTW(Self, TWOrdreTab);
  TWOrdreTab.Visible := True;
  BValidelOrdreTab.Enabled := (Action <> taConsult);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BHautClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BHautClick(Sender: TObject);
var ARow: integer;
begin
  ARow := GSOrdTab.Row;
  if ARow > GSOrdTab.FixedRows then
  begin
    GSOrdTab.ExchangeRow(ARow, ARow - 1);
    GSOrdTab.Row := ARow - 1;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BBasClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BBasClick(Sender: TObject);
var ARow: integer;
begin
  ARow := GSOrdTab.Row;
  if ARow < GSOrdTab.RowCount - 1 then
  begin
    GSOrdTab.ExchangeRow(ARow, ARow + 1);
    GSOrdTab.Row := ARow + 1;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BValidelOrdreTabClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BValidelOrdreTabClick(Sender: TObject);
var ARow: integer;
begin
  if Action = taConsult then Exit;
  TWOrdreTab.Visible := False;
  for ARow := GSOrdTab.FixedRows to GSOrdTab.RowCount - 1 do
  begin
    if GSOrdTab.Cells[0, ARow] = 'GP_TIERS' then
      GP_TIERS.TabOrder := ARow
    else if GSOrdTab.Cells[0, ARow] = 'GP_REPRESENTANT' then
      GP_REPRESENTANT.TabOrder := ARow
    else if GSOrdTab.Cells[0, ARow] = 'GP_DATEPIECE' then
      GP_DATEPIECE.TabOrder := ARow;
  end;
  MajParam := True;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BCancelOrdreTabClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BCancelOrdreTabClick(Sender: TObject);
begin
  TWOrdreTab.Visible := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  PEnteteDblClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.PEnteteDblClick(Sender: TObject);
begin
  OuvrePropPanel;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  PPiedDblClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.PPiedDblClick(Sender: TObject);
begin
  OuvrePropPanel;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OuvrePropPanel
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.OuvrePropPanel;
begin
  if FOExisteClavierEcran then
    SPNbRows.MaxValue := 20
  else
    SPNbRows.MaxValue := 22;
  CentreTW(Self, TWPropPanel);
  TWPropPanel.Visible := True;
  BValidePropPanel.Enabled := (Action <> taConsult);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BValidePropPanelClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BValidePropPanelClick(Sender: TObject);
var aHeight: integer;
begin
  if Action = taConsult then Exit;
  TWPropPanel.Visible := False;
  PEntete.ColorStart := BColorStart.Color;
  PEntete.ColorEnd := BColorEnd.Color;
  PEntete.BackGroundEffect := TDirection(CCBackGround.ItemIndex);
  if PEntete.BackGroundEffect = bdFond then PEntete.ColorNB := 6 else PEntete.ColorNB := 100;
  PPied.ColorStart := BColorStart.Color;
  PPied.ColorEnd := BColorEnd.Color;
  PPied.BackGroundEffect := TDirection(CCBackGround.ItemIndex);
  if PPied.BackGroundEffect = bdFond then PPied.ColorNB := 6 else PEntete.ColorNB := 100;
  GS.FixedColor := BGSColor.Color;
  GSReg.FixedColor := BGSColor.Color;
  if GS.TwoColors <> CBTwoColors.Checked then
  begin
    GS.TwoColors := CBTwoColors.Checked;
    GSReg.TwoColors := GS.TwoColors;
    if GS.Visible then GS.Repaint else GSReg.Repaint;
  end;
  if SPNbRows.Value <> GS.VisibleRowCount then
  begin
    aHeight := PGrille.Height;
    PGrille.Height := 84 + GS.RowHeights[0] + 2 + ((GS.RowHeights[1] + 1) * SPNbRows.Value);
    Dec(aHeight, PGrille.Height);
    if PnlBoutons <> nil then
    begin
      Pmain.Tag := PnlBoutons.Height + aHeight;
      PnlBoutons.ResizeClavier(nil);
    end;
    if PnlBtn2 <> nil then PnlBtn2.ResizeClavier(nil);
  end;
  MajParam := True;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BCancelPropPanelClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BCancelPropPanelClick(Sender: TObject);
begin
  TWPropPanel.Visible := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BColorStartClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BColorStartClick(Sender: TObject);
begin
  ChoixCouleur.Color := BColorStart.Color;
  ChoixCouleur.Execute;
  BColorStart.Color := ChoixCouleur.Color;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BColorEndClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BColorEndClick(Sender: TObject);
begin
  ChoixCouleur.Color := BColorEnd.Color;
  ChoixCouleur.Execute;
  BColorEnd.Color := ChoixCouleur.Color;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BGSColorClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BGSColorClick(Sender: TObject);
begin
  ChoixCouleur.Color := BGSColor.Color;
  ChoixCouleur.Execute;
  BGSColor.Color := ChoixCouleur.Color;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  fAffTXTDblClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.fAffTXTDblClick(Sender: TObject);
begin
  OuvrePropLCD;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  PI_NETAPAYERDEVDblClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.PI_NETAPAYERDEVDblClick(Sender: TObject);
begin
  OuvrePropLCD;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OuvrePropLCD
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.OuvrePropLCD;
begin
  CentreTW(Self, TWPropLCD);
  TWPropLCD.Visible := True;
  BValidePropLCD.Enabled := (Action <> taConsult);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BValidePropLCDClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BValidePropLCDClick(Sender: TObject);
var Texte: string;
begin
  if Action = taConsult then Exit;
  TWPropLCD.Visible := False;
  fAffTXT.BackGround := BColorFond.Color;
  fAffTXT.PixelOn := BPixelON.Color;
  fAffTXT.PixelOff := BPixelOFF.Color;
  PI_NETAPAYERDEV.BackGround := BColorFond.Color;
  PI_NETAPAYERDEV.PixelOn := BPixelON.Color;
  PI_NETAPAYERDEV.PixelOff := BPixelOFF.Color;
  // modification de l'afficheur
  fAffTXT.Caption := '';
  fAffTXT.UpSideDown := CBAffInverse.Checked;
  LCDVisible(CBAffClient.Checked);
  fAffTXT.Invalidate;
  Texte := Trim(MsgAccueil.Text);
  if Texte = '' then Texte := 'Bienvenue';
  Texte := Texte + ' ';
  AffichageLCD(Texte, 0, 0, 0, '', [ofInterne, ofClient], qaLibreTexte, False);
  MajParam := True;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BCancelPropLCDClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BCancelPropLCDClick(Sender: TObject);
begin
  TWPropLCD.Visible := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BColorFondClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BColorFondClick(Sender: TObject);
begin
  ChoixCouleur.Color := BColorFond.Color;
  ChoixCouleur.Execute;
  BColorFond.Color := ChoixCouleur.Color;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BPixelONClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BPixelONClick(Sender: TObject);
begin
  ChoixCouleur.Color := BPixelON.Color;
  ChoixCouleur.Execute;
  BPixelON.Color := ChoixCouleur.Color;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BPixelOFFClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BPixelOFFClick(Sender: TObject);
begin
  ChoixCouleur.Color := BPixelOFF.Color;
  ChoixCouleur.Execute;
  BPixelOFF.Color := ChoixCouleur.Color;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  TitreFenetre : fabrique le titre de la fenêtre
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.TitreFenetre(ACol: Integer; var Fen: TToolWindow97);
var Stg: string;
  Ind: Integer;
begin
  Ind := Pos(':', Fen.Caption);
  if Ind > 0 then Dec(Ind) else Ind := Length(Fen.Caption);
  Stg := Copy(Fen.Caption, 1, Ind);
  if ACol >= 0 then
  begin
    if ACol < 20 then Stg := Stg + ': ' + GS.Cells[ColInd, 0] else
      if ACol < 40 then Stg := Stg + ': ' + GSReg.Cells[ColInd - 20, 0] else
      if ACol = 40 then Stg := Stg + ': ' + HGP_TIERS.Caption else
      if ACol = 41 then Stg := Stg + ': ' + HGP_REPRESENTANT.Caption;
  end;
  Fen.Caption := Stg;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OuvreChoixPave
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.OuvreChoixPave(ACol: Integer);
var Stg: string;
  QQ: TQuery;
  Ind: Integer;
begin
  CentreTW(Self, TWChoixPave);
  TWChoixPave.Visible := True;
  ColInd := ACol;
  if ColCarac[ColInd].NoPage > 0 then
  begin
    RBPage.Checked := True;
    SPNoPage.Value := ColCarac[ColInd].NoPage;
  end else
  begin
    case ColCarac[ColInd].NoPage of
      -1: RBArtFi.Checked := True;
      -2: RBTypeRemise.Checked := True;
      -3: RBVendeur.Checked := True;
      -4: RBModePaie.Checked := True;
    else RBRien.Checked := True;
    end;
    SPNoPage.Value := 0;
  end;
  ActiveOptionsChoixPave;
  // recherche des n° page du clavier
  Stg := 'SELECT MIN(CE_PAGE) AS MINI,MAX(CE_PAGE) AS MAXI FROM CLAVIERECRAN'
    + ' WHERE CE_CAISSE="' + FOGetParamCaisse('GPK_CAISSE') + '"';
  QQ := OpenSQL(Stg, True);
  if not QQ.EOF then
  begin
    Ind := QQ.FindField('MINI').AsInteger + 1;
    if Ind < 0 then Ind := 1;
    SPNoPage.MinValue := Ind;
    Ind := QQ.FindField('MAXI').AsInteger + 1;
    if (Ind < SPNoPage.MinValue) then Ind := SPNoPage.MinValue;
    SPNoPage.MaxValue := Ind;
  end;
  Ferme(QQ);
  BValideChoixPave.Enabled := (Action <> taConsult);
  // titre de la fenêtre
  TitreFenetre(ColInd, TWChoixPave);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ActiveOptionsChoixPave
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.ActiveOptionsChoixPave;
begin
  SPNoPage.Enabled := (RBPage.Checked);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  RadioButtonClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.RadioButtonClick(Sender: TObject);
begin
  ActiveOptionsChoixPave;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BValideChoixPaveClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BValideChoixPaveClick(Sender: TObject);
begin
  if Action = taConsult then Exit;
  TWChoixPave.Visible := False;
  ColCarac[ColInd].NoPage := 0;
  if RBArtFi.Checked then ColCarac[ColInd].NoPage := -1 else
    if RBTypeRemise.Checked then ColCarac[ColInd].NoPage := -2 else
    if RBVendeur.Checked then ColCarac[ColInd].NoPage := -3 else
    if RBModePaie.Checked then ColCarac[ColInd].NoPage := -4 else
    if RBPage.Checked then ColCarac[ColInd].NoPage := SPNoPage.Value;
  AffichePaveChoisi(ColInd);
  MajParam := True;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BCancelChoixPaveClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BCancelChoixPaveClick(Sender: TObject);
begin
  TWChoixPave.Visible := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  AffichePaveChoisi
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.AffichePaveChoisi(ACol: Integer);
begin
  if PnlBoutons = nil then Exit;
  if (ACol < Low(ColCarac)) or (ACol > High(ColCarac)) then Exit;
  ColInd := ACol;
  AffichePaveContextuel(PnlBoutons, ColCarac[ACol].NoPage);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OuvrePropCell
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.OuvrePropCell(ACol: Integer);
var Grille: THGrid;
begin
  ColInd := ACol;
  if ColInd < 20 then
  begin
    Grille := GS;
    ACol := ColInd;
  end else
  begin
    Grille := GSReg;
    ACol := ColInd - 20;
  end;
  if (ACol < Grille.FixedCols) and (ACol >= Grille.ColCount) then Exit;
  CentreTW(Self, TWPropCell);
  TWPropCell.Visible := True;
  // alignement de la cellule
  if Grille.ColAligns[ACol] = taLeftJustify then CtrAlign.ItemIndex := 0 else
    if Grille.ColAligns[ACol] = taCenter then CtrAlign.ItemIndex := 1 else
    if Grille.ColAligns[ACol] = taRightJustify then CtrAlign.ItemIndex := 2;
  // largeur de la cellule
  SPLargeur.Value := Grille.ColWidths[ACol];
  // titre de la fenêtre
  TitreFenetre(ColInd, TWPropCell);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BValidePropCellClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BValidePropCellClick(Sender: TObject);
var Grille: THGrid;
  ACol: Integer;
  GrilReg: Boolean;
begin
  if Action = taConsult then Exit;
  TWPropCell.Visible := False;
  if ColInd < 20 then
  begin
    Grille := GS;
    ACol := ColInd;
    GrilReg := False;
  end else
  begin
    Grille := GSReg;
    ACol := ColInd - 20;
    GrilReg := True;
  end;
  if (ACol < Grille.FixedCols) and (ACol >= Grille.ColCount) then Exit;
  // alignement de la cellule
  case CtrAlign.ItemIndex of
    0: Grille.ColAligns[ACol] := taLeftJustify;
    1: Grille.ColAligns[ACol] := taCenter;
    2: Grille.ColAligns[ACol] := taRightJustify;
  end;
  ColCarac[ColInd].Align := Grille.ColAligns[ACol];
  // largeur de la cellule
  if Grille.ColWidths[ACol] <> SPLargeur.Value then
  begin
    Grille.ColWidths[ACol] := SPLargeur.Value;
    RetailleGrille(GrilReg);
  end;
  MajParam := True;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BCancelPropCellClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BCancelPropCellClick(Sender: TObject);
begin
  TWPropCell.Visible := False;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GP_TIERSDblClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.GP_TIERSDblClick(Sender: TObject);
begin
  OuvreChoixPave(40);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GP_TIERSEnter
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.GP_TIERSEnter(Sender: TObject);
begin
  AffichePaveChoisi(40);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GP_REPRESENTANTDblClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.GP_REPRESENTANTDblClick(Sender: TObject);
begin
  OuvreChoixPave(41);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GP_REPRESENTANTEnter
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.GP_REPRESENTANTEnter(Sender: TObject);
begin
  AffichePaveChoisi(41);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSEnter
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.GSEnter(Sender: TObject);
var Cancel: Boolean;
  ACol, ARow: Integer;
begin
  Cancel := False;
  ACol := GS.Col;
  ARow := GS.Row;
  GSCellEnter(GS, ACol, ARow, Cancel);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSDblClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.GSDblClick(Sender: TObject);
begin
  OuvreChoixPave(GS.Col);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSCellEnter
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.GSCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
  AffichePaveChoisi(GS.Col);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSMouseUp
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.GSMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var Ind: Integer;
begin
  for Ind := 0 to GS.ColCount - 1 do
  begin
    if GS.ColWidths[Ind] <> ColCarac[Ind].Largeur then
    begin
      RetailleGrille(False);
      Break;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSRegEnter
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.GSRegEnter(Sender: TObject);
var Cancel: Boolean;
  ACol, ARow: Integer;
begin
  Cancel := False;
  ACol := GSReg.Col;
  ARow := GSReg.Row;
  GSRegCellEnter(GSReg, ACol, ARow, Cancel);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSRegDblClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.GSRegDblClick(Sender: TObject);
begin
  OuvreChoixPave(20 + GSReg.Col);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSRegCellEnter
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.GSRegCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
  AffichePaveChoisi(20 + GSReg.Col);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GSRegMouseUp
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.GSRegMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var Ind: Integer;
begin
  for Ind := 0 to GSReg.ColCount - 1 do
  begin
    if GSReg.ColWidths[Ind] <> ColCarac[20 + Ind].Largeur then
    begin
      RetailleGrille(True);
      Break;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BValiderClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BValiderClick(Sender: TObject);
begin
  ValideSaisie;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BAbandonClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BAbandonClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BAideClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BpropPanelClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BpropPanelClick(Sender: TObject);
begin
  OuvrePropPanel;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BPropLCDClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BPropLCDClick(Sender: TObject);
begin
  OuvrePropLCD;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BClavierClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BClavierClick(Sender: TObject);
var sParam: string;
  sOldParam: string;
begin
  if Sender = mPaveSecond then
  begin
    // pavé secondaire
    sParam := IntToStr(PropPv2.NbrBtnWidth) + ';' + IntToStr(PropPv2.NbrBtnHeight) + ';'
      + IntToStr(Ord(PropPv2.ClcPosition)) + ';' + TrueFalseSt(PropPv2.ClcVisible) + ';';
    sOldParam := sParam;
    Param_ClavierEcran(FOAlphaCodeNumeric(FOCaisseCourante), sParam, Action);
    if sParam <> sOldParam then
    begin
      PropPv2.NbrBtnWidth := Valeuri(ReadTokenst(sParam));
      PropPv2.NbrBtnHeight := Valeuri(ReadTokenst(sParam));
      PropPv2.ClcPosition := tPosClc(Valeuri(ReadTokenst(sParam)));
      PropPv2.ClcVisible := (ReadTokenst(sParam) <> '-');
      MajParam := True;
    end;
    MiseEnFormePave2;
  end else
  begin
    // pavé principal
    sParam := FOGetParamCaisse('GPK_PARAMSCE');
    sOldParam := sParam;
    Param_ClavierEcran(FOGetParamCaisse('GPK_CAISSE'), sParam, Action);
    MiseEnFormePave;
    if sParam <> sOldParam then
    begin
      VH_GC.TOBPCaisse.PutValue('GPK_PARAMSCE', sParam);
      MajParam := True;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BPreferenceClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BPreferenceClick(Sender: TObject);
begin
  {$IFNDEF EAGLCLIENT}
  if Preferences(['Pièces', '', '', '', 'Tickets'], False) then PnlBoutons.ColorNb := 6;
  {$ENDIF}
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BCaisseClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BCaisseClick(Sender: TObject);
begin
  AGLLanceFiche('MFO', 'PCAISSE', FOGetParamCaisse('GPK_CAISSE'), '', ActionToString(Action));
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BChoixPaveClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BChoixPaveClick(Sender: TObject);
begin
  if Screen.ActiveControl = GS then OuvreChoixPave(GS.Col) else
    if Screen.ActiveControl = GSReg then OuvreChoixPave(20 + GSReg.Col) else
    if Screen.ActiveControl = GP_TIERS then OuvreChoixPave(40) else
    if Screen.ActiveControl = GP_REPRESENTANT then OuvreChoixPave(41);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BPropCellClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BPropCellClick(Sender: TObject);
begin
  if Screen.ActiveControl = GS then OuvrePropCell(GS.Col) else
    if Screen.ActiveControl = GSReg then OuvrePropCell(20 + GSReg.Col);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BEcheClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BEcheClick(Sender: TObject);
begin
  if BEche.Down then
  begin
    GS.Visible := False;
    GSReg.Visible := True;
    if GSReg.CanFocus then GSReg.SetFocus;
  end else
  begin
    GSReg.Visible := False;
    GS.Visible := True;
    if GS.CanFocus then GS.SetFocus;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BDefaultConfigClick
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BDefaultConfigClick(Sender: TObject);
begin
  // Propriétés des panels
  PPied.ColorStart := DefColorStart;
  PPied.ColorEnd := DefColorEnd;
  PPied.BackGroundEffect := DefBackGround;
  // Propriétés des LCD
  fAffTXT.BackGround := DefColorFond;
  fAffTXT.PixelOn := DefPixelON;
  fAffTXT.PixelOff := DefPixelOFF;
  // Propriétés des grilles
  GS.FixedColor := DefGSColor;
  GS.TwoColors := DefTwoColors;
  GS.ListeParam := '';
  GSReg.ListeParam := '';
  PGrille.Height := 114 + (GS.RowHeights[1] * DefNbRows);
  // Ordre de tabulation
  GP_TIERS.TabOrder := DefTabOrder[0];
  GP_REPRESENTANT.TabOrder := DefTabOrder[1];
  GP_DATEPIECE.TabOrder := DefTabOrder[2];
  MiseEnFormeFiche(False);
  BValidePropLCDClick(nil);
  BValidePropPanelClick(nil);
  {***********************************************
  for Ind := Low(ColCarac) to High(ColCarac) do
     BEGIN
     ColCarac[Ind].NoPage := DefCarac[Ind].NoPage ;
     ColCarac[Ind].Largeur := DefCarac[Ind].Largeur ;
     ColCarac[Ind].Align := DefCarac[Ind].Align ;
     END ;
  for Ind := 0 to GS.ColCount -1 do
     BEGIN
     GS.ColWidths[Ind] := DefCarac[Ind].Largeur ;
     GS.ColAligns[Ind] := DefCarac[Ind].Align ;
     END ;
  for Ind := 0 to GSReg.ColCount -1 do
     BEGIN
     GSReg.ColWidths[Ind] := DefCarac[20+Ind].Largeur ;
     GSReg.ColAligns[Ind] := DefCarac[20+Ind].Align ;
     END ;
  **************************************************}
  // 2ème pavé
  PropPv2 := DefPropPv2;
  MiseEnFormePave2;
  MiseEnFormePave;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BPave2Click
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BPave2Click(Sender: TObject);
begin
  OuvrePropPave2;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OuvrePropPave2
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.OuvrePropPave2;
begin
  CentreTW(Self, TWPropPave2);
  TWPropPave2.Visible := True;
  BValidePropPave2.Enabled := (Action <> taConsult);
  CBPave2.Checked := PropPv2.Visible;
  SpLargPave2.Value := PropPv2.Largeur;
  if PropPv2.Align = taLeftJustify then
    CtrAlignPave2.ItemIndex := 0
  else
    CtrAlignPave2.ItemIndex := 1;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BValidePropPave2Click
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BValidePropPave2Click(Sender: TObject);
var
  wAlign: TAlignment;
begin
  if Action = taConsult then Exit;
  TWPropPave2.Visible := False;
  if PropPv2.Visible <> CBPave2.Checked then
  begin
    PropPv2.Visible := CBPave2.Checked;
    MajParam := True;
  end;
  if PropPv2.Largeur <> SpLargPave2.Value then
  begin
    PropPv2.Largeur := SpLargPave2.Value;
    MajParam := True;
  end;
  if CtrAlignPave2.ItemIndex = 0 then
    wAlign := taLeftJustify
  else
   wAlign := taRightJustify;
  if PropPv2.Align <> wAlign then
  begin
    PropPv2.Align := wAlign;
    MajParam := True;
  end;
  if MajParam then MiseEnFormePave2;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  BCancelPropPave2Click
///////////////////////////////////////////////////////////////////////////////////////

procedure TFParamTicket.BCancelPropPave2Click(Sender: TObject);
begin
  TWPropPave2.Visible := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : FOCopieCaisse = recopie le paramétrage d'une caisse
Suite ........ : depuis le script d'une fiche
Suite ........ :  - Parms[0] = Nature de pièce
Suite ........ :  - Parms[1] = Code action
Suite ........ :  - Parms[2] = Code caisse
Mots clefs ... : FO
*****************************************************************}

procedure FOAGLSaisieParamTicket(parms: array of variant; nb: integer);
var Action: TActionFiche;
begin
  Action := TActionFiche(vInteger(Parms[1]));
  SaisieParamTicketFO(string(Parms[0]), Action, string(Parms[2]));
end;

///////////////////////////////////////////////////////////////////////////////////////
//  Initialization
///////////////////////////////////////////////////////////////////////////////////////

initialization
  RegisterAglProc('FOSaisieParamTicket', False, 3, FOAGLSaisieParamTicket);

end.
