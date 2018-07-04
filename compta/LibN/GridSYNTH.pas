{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 16/03/2004
Modifié le ... : 28/09/2006
Description .. : - FQ 13347 - 16/03/2004 - CA
Suite ........ : - 21/02/2005 - CA - Gestion plus "propre" des Tags des
Suite ........ : grilles :
Suite ........ :           _ Modèle : Tag = 0
Suite ........ :           _ Etat : Tag = 1
Suite ........ :           _ Détail : Tag = 2
Suite ........ :           _ Non affectés : Tag = 3
Suite ........ :           _ Bilan passif : Tag = 4
Suite ........ :           _ Détail bilan passif : Tag = 5
Suite ........ :
Suite ........ : - FQ 17421 - 29/06/2006 : Correction stOleErr
Suite ........ : - FQ 18884 - 28/09/2006 : Résolution non renseignée
Mots clefs ... :
*****************************************************************}
unit GridSYNTH;

interface

uses
  Windows,
  // Messages,
  SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls,
  // HPanel,
  Grids, Hctrls, UTob, HEnt1, Formule, HTB97,
  ComCtrls, StdCtrls, HStatus, Buttons, HColor,
  QRGridPCL,
  // HXLSPas,
  uTilXLS, LookUp,
{$IFDEF EAGLCLIENT}
{$ELSE}
{$IFNDEF DBXPRESS}dbtables, {$ELSE}uDbxDataSet, {$ENDIF}
  EdtREtat,
  PrintDBG, // PrintDBGrid
{$ENDIF}
  CritEdt,
  // Hcompte,
  Ent1,
  HMsgBox,
  HSysMenu, ulibExercice, utilpgi, calcole,
  ed_tools,
  // uLibWindows,
  Mask,
  uPrintF1Book,
  HPdfPrev,
  HPdfViewer, Printers,
  ZCumul,
{$IFDEF GPL_PLAQPGI} //cv4 le 19/06/2007 -> impression dans GPL -> NE PAS ENLEVER CE CODE SVP !!!!!
  LIASSELIB,
  GdP_Global,
{$ENDIF GPL_PLAQPGI}

  uYFILESTD, HPanel, TntStdCtrls, TntExtCtrls, TntButtons, TntGrids;

function LanceLiasse(ModeleName: string; VoirDetail, VoirNom: Boolean;
  FormuleCol, FormatCol, FormatColDetail: array of string;
  var Crit: tCritEdtPCL; Demo: Boolean;
  vCritEdtChaine: TCritEdtChaine): boolean;

const
  MAX_COL = 20;

type
  TFEtatLiasse = class(TForm)
    HPB: TToolWindow97;
    Dock971: TDock97;
    Pages: TPageControl;
    PResult: TTabSheet;
    PDetails: TTabSheet;
    FListe: THGrid;
    PParams: TTabSheet;
    GParams: THGrid;
    FD: TFontDialog;
    CD: TColorDialog;
    FTestFont: TLabel;
    Panel1: TPanel;
    Label1: TLabel;
    FTitre: TEdit;
    BDelLigne: THBitBtn;
    BInsLigne: THBitBtn;
    bSauve: TToolbarButton97;
    bColor: TToolbarButton97;
    bPolice: TToolbarButton97;
    bDroite: TToolbarButton97;
    bCentre: TToolbarButton97;
    bGauche: TToolbarButton97;
    bFrame: TPaletteButton97;
    FDetail: THGrid;
    PNonAffectes: TTabSheet;
    FListeNonAffectes: THGrid;
    pnl_Criteres: TPanel;
    lbl_lblAxe: TLabel;
    lbl_lblSectionDe: TLabel;
    lbl_lblSectionA: TLabel;
    lbl_Axe: TLabel;
    lbl_SectionA: TLabel;
    lbl_SectionDe: TLabel;
    HMtrad: THSystemMenu;
    Timer1: TTimer;
    Panel2: TPanel;
    BImprimer: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    BValider: TToolbarButton97;
    FType: THValComboBox;
    BExport: TToolbarButton97;
    SD: TSaveDialog;
    Cache: THCritMaskEdit;
    BOUVRIR: TToolbarButton97;
    PNOMFICHIERPGE: THPanel;
    HLabel1: THLabel;
    FFICHIERPGE: THCritMaskEdit;
    BENREGISTREPGE: TToolbarButton97;
    FCHEMINFICHIERPGE: THCritMaskEdit;
    BANNULER: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FListeDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect;
      State: TGridDrawState);
    procedure BImprimerClick(Sender: TObject);
    procedure PagesChange(Sender: TObject);
    procedure bSauveClick(Sender: TObject);
    procedure GParamsRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean;
      Chg: Boolean);
    procedure GParamsRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean;
      Chg: Boolean);
    procedure GParamsCellExit(Sender: TObject; var ACol, ARow: Integer; var
      Cancel: Boolean);
    procedure BDelLigneClick(Sender: TObject);
    procedure BInsLigneClick(Sender: TObject);
    procedure GParamsRowMoved(Sender: TObject; FromIndex, ToIndex: Integer);
    procedure bPoliceClick(Sender: TObject);
    procedure bColorClick(Sender: TObject);
    procedure bGaucheClick(Sender: TObject);
    procedure bFrameChange(Sender: TObject);
    procedure GParamsDblClick(Sender: TObject);
    procedure GParamsKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FListeDblClick(Sender: TObject);
    procedure FListeMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BCValideClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BExportClick(Sender: TObject);
    procedure BOUVRIRClick(Sender: TObject);
    procedure BENREGISTREPGEClick(Sender: TObject);
    procedure BANNULERClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
  private { Déclarations privées }
    Flafonte: tfont;
    FModeleName: string;
    GX, GY: Integer;
    FCurrentTOB: TOB;
    FTL: HTStringList;
    FVoirDetail, FVoirNom, FDetailCalcule, FNonAffectes: Boolean;
    FFormuleCol, FFormatCol, FFormatColDetail: array[1..MAX_COL] of string;
    FCritEdtChaine: TCritEdtChaine;
    { Ajout CA le 13/02/2001 }
    FTitreNA: array[0..7] of string;
    FTitrePassif: array[0..7] of string;
    FEcartResultat: array[1..MAX_COL] of double;
    FBilanActif: array[1..MAX_COL] of double;
    FBilanPassif: array[1..MAX_COL] of double;
    FBEcartEtat: boolean;
    FLigneFinActif, FLigneFinDetailActif: integer;
    { Fin Ajout CA le 13/02/2001 }
    FCurrentCol, FNbCol: Integer;
    FLaTOB, FDetailTob: TOB;
    Crit: tCritEdtPCL;
    Demo: Boolean;
//    FListeDesComptes: TOB;
    FListeDesSections: TOB;
    fbStopEdition: boolean;
    fLesRubriques: TZCumul;
//    fLesRubriquesMem: TZCumulRubrique;
    fTLesGrilles: TOB;
    fTitreMaquette: array[0..4] of string; // Titre de l'état indiqué dans la maquette
    procedure Charge;
    procedure ChargeDetail;
    procedure CalculRub;
    procedure CalculRubDetail;
    function LeCumul(Nom: string; LaCol: Integer; var LesCptes: tStringList):
      Variant;
    function LeCumulDetail(Cpt: string; LaCol, LaLig: Integer; bComparatifBilan:
      boolean): Double;
    function GetCellCumul(St: hstring): Variant;
    //    procedure GetRubDetail (Nom : String ; Noms,Libelles : TStringList ; R : Integer);
    procedure ChargeDetailCompteRubrique(Noms, Libelles: TStringList; R:
      integer);
    procedure DecodeFont(lafonte: TFOnt; St: string);
    function EncodeFont: string;
    function AddGras(St: string): string;
    function IsLigneVide(ARow: Integer): Boolean;
    procedure GDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState:
      TGridDrawState);
    procedure ChargeNonAffectes;
    procedure UpdateRowLengths(G: THGrid; bMontant0: boolean);
    function QuelleUnite(XX: Double; bDetail: boolean = False): Double;
    procedure AfficheMentionBilan;
    procedure AfficheMentionResultat;
    procedure AjouteMentionCellule(Col: integer; Valeur: double);
    procedure CreeLigneMention;
    procedure UpdateTitreBilan;
    procedure AjusteTailleColonne(G: THGrid; bEtat: boolean);
    procedure ValideCtxPCL;
    //    procedure ValideCtxNotPCL;
    function GetResultatSurPeriode(stDate: string; var stErr: string;
      EnMonnaieOpposee, Etab, BalSit: string): double;
    procedure AfficheCell(FromPrinter: boolean; lecanvas: tcanvas; letag:
      integer; letext, lestyle: string; f: TAlignment; lafonte: tfont; ACol, ARow:
      Integer; Rect: TRect; State:
      TGridDrawState; fx, fy: integer);
    procedure PrepareCell(legrid: tHgrid; var letag: integer; var lestyle,
      letext: string; var laligment: TAlignment; lafonte: tfont; State:
      TGridDrawState; acol, arow:
      integer);
    procedure ChargeMaquette;
//    procedure ChargeListeDesComptes;
    procedure ChargeListeDesSections;
//    function LibelleDuCompte(stCompte: string): string;
    function LibelleDuCompte(pStDetail: HString): HString;
    function LibelleDeLaSection(stSection: string): string;
    procedure DoWriteCell(const LeGrid: THGrid; const ARow, ACol: integer;
      TMere: TOB);
    procedure GrilleToTob(TMere: TOB);
    procedure ExportVersExcel(T: TOB; FileName: string; bModele: boolean);
    function ImprimeLesGrilles: string;
    function ConstruireParametreCritere: string;
    function ImprimeUneGrille(TUneGrille: TOB): boolean;
    function FaitLibelleEtatSynthese(TE: TTypeEtatSynthese; bEtat, bDetail,
      bPassif, bPublifi: boolean): string;
    function TitreEtatSynthese(Crit: tCritEdtPCL; iTag: integer): string;
    procedure MajTitreGrilles;
    procedure CalculLesGrilles;
    procedure RenommePourPublifi;
    procedure MajCaptionAnalytique;
    procedure FaitTitreMaquette(St: string);
    function ControleStandardPCL(var pStCrit3: string; var pstPredefini: string): boolean;
    function GrilleVersMaquette(TL: HTStringList): integer;
    function JePeuxEnregistrerMaquette(pEtatOk: integer): boolean;
    procedure EnregistreMaquette(TL: HTStringList; pstFichier, pstCrit3, pstPredefini: string);
  public { Déclarations publiques }
  end;

function DecaleLibelle(St: string): string;
function GetNomCabinet: string;
function FaitMasqueDecimales: string;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  CPVersion,
  {$ENDIF MODENT1}
  ParamSoc;

const
  POS_COL: INTEGER = 2;
  POS_LIBELLE: Integer = 1;
  POS_NOM: Integer = 0;
  Default_Style = 'Arial;8;;clBlack;clWhite;----;';
  CHPCTRL_ACTIF = 'ACTIF';
  CHPCTRL_PASSIF = 'PASSIF';
  CHPCTRL_RESULTAT = 'RESULTAT';

procedure VidePileGrid(G: THGrid; FreeObj: Boolean);
var
  C, R: integer;
begin
  for R := G.FixedRows to G.RowCount - 1 do
    for C := 0 to G.ColCount - 1 do
    begin
      if ((FreeObj) and (G.Objects[C, R] <> nil)) then G.Objects[C, R].Free;
      G.Cells[C, R] := '';
      G.Objects[C, R] := nil;
    end;
  G.RowCount := G.FixedRows + 1;
end;

function TFEtatLiasse.QuelleUnite(XX: Double; bDetail: boolean = False): Double;
var
  szResol: string;
begin
  Result := XX;
  if Demo then
    Exit;
  if crit.Resolution = '' then
    exit;
  if bDetail then
    szResol := crit.ResolutionDetail
  else
    szResol := crit.Resolution;
  if Length(szResol) > 0 then // CA - 28/09/2006 - FQ 18884 : Résolution on renseignée
  begin
    case szResol[1] of
      'C': ;
      'F': ;
      'K': Result := XX / 1000;
      'M': Result := XX / 1000000;
    end;
  end;
end;

function LanceLiasse(ModeleName: string; VoirDetail, VoirNom: Boolean;
  FormuleCol, FormatCol, FormatColDetail: array of string;
  var Crit: tCritEdtPCL; Demo: Boolean;
  vCritEdtChaine: TCritEdtChaine): boolean;
var
  FEtatLiasse: TFEtatLiasse;
  i: Integer;
begin
  FEtatLiasse := TFEtatLiasse.Create(Application);
  FEtatLiasse.FModeleName := ModeleName;
  FEtatLiasse.FVoirDetail := VoirDetail;
  FEtatLiasse.FVoirNom := VoirNom;
  FEtatLiasse.Crit := Crit;
  FEtatLiasse.Demo := Demo;
  FEtatLiasse.FCritEdtChaine := vCritEdtChaine;
  for i := 1 to MAX_COL do
  begin
    FEtatLiasse.FFormuleCol[i] := '';
    FEtatLiasse.FFormatCol[i] := '';
    FEtatLiasse.FFormatColDetail[i] := '';
  end;
  for i := Low(FormuleCol) to High(FormuleCol) do
    FEtatLiasse.FFormuleCol[i + 1] := FormuleCol[i];
  for i := Low(FormatCol) to High(FormatCol) do
    FEtatLiasse.FFormatCol[i + 1] := FormatCol[i];
  for i := Low(FormatColDetail) to High(FormatColDetail) do
    FEtatLiasse.FFormatColDetail[i + 1] := FormatColDetail[i];
  FEtatLiasse.FNbCol := High(FormuleCol) + 1;
  FEtatLiasse.ShowModal;
  Result := not FEtatLiasse.fbStopEdition;
  FEtatLiasse.Free;
end;

procedure TFEtatLiasse.FormShow(Sender: TObject);
var
  stFileName: string;
  bOldNoPrintDialog, bOldSilentMode, bOldPDF: boolean;
  iCopie: integer;
begin
  { Initialisation du ZCumul }
  if Crit.Corresp = '1' then fLesRubriques.Corresp := coCorresp1
  else if Crit.Corresp = '2' then fLesRubriques.Corresp := coCorresp2
  else fLesRubriques.Corresp := coGeneral;

  { Interface utilisateur }
  BValider.Visible := (not (ctxStandard in V_PGI.PGIContexte));
  BOuvrir.Visible := not (ctxPCL in V_PGI.PGIContexte);

  { Initialisation des variables }
  FLigneFinActif := 0;
  FLigneFinDetailActif := 0;
  if not Crit.AvecDetail then
    PDetails.TabVisible := FALSE;
  Pages.ActivePageIndex := 0;
  if FVoirNom then
  begin
    POS_NOM := 0;
    POS_LIBELLE := 1;
    POS_COL := 2;
  end
  else
  begin
    POS_NOM := -1;
    POS_LIBELLE := 0;
    POS_COL := 1;
  end;

  { Initialisation de la grille }
  if POS_LIBELLE >= 0 then
  begin
    FListe.Cells[POS_LIBELLE, 0] := '';
    FListe.ColWidths[POS_LIBELLE] := 270; // 250
  end;
  if POS_NOM >= 0 then
  begin
    FListe.Cells[POS_NOM, 0] := '';
    FListe.ColWidths[POS_NOM] := 80;
  end;
  FListe.ColCount := FNbCol + 1 + Ord(FVoirNom);
  FDetail.ColCount := FListe.ColCount;
  if POS_LIBELLE >= 0 then
    FDetail.ColWidths[POS_LIBELLE] := FListe.ColWidths[POS_LIBELLE];
  if POS_NOM >= 0 then
    FDetail.ColWidths[POS_NOM] := FListe.ColWidths[POS_NOM];

  if Demo then
  begin
    Pages.ActivePage := PParams;
    if ctxPCL in V_PGI.PGIContexte then
    begin
      PResult.TabVisible := FALSE;
      PNonAffectes.TabVisible := False;
    end;
    PDetails.TabVisible := FALSE;
    GParams.SetFocus;
    PagesChange(nil);
    exit;
  end
  else
    PParams.TabVisible := False;
      // CA - 02/12/2002 - Onglet modèle invisible en édition

  if (((ctxPCL in V_PGI.PGIContexte) and (not Demo)) or (not (ctxPCL in
    V_PGI.PGIContexte))) then
    if not Crit.AvecDetail then
    begin
      Pages.ActivePage := PResult;
      PDetails.TabVisible := FALSE;
      FListe.SetFocus;
      PagesChange(nil);
    end;

  if (Crit.TE = esBIL) or (Crit.TE = esBILA) then
    UpdateTitreBilan;

  if ctxStandard in V_PGI.PGIContexte then
    BValider.visible := FALSE;

  fTLesGrilles := TOB.Create('Maman', nil, -1);
  try
    { Constitution des grilles destinées à l'édition }
    CalculLesGrilles;
    if fLesRubriques.LastError <> 0 then
    begin
      PGIBox(fLesRubriques.LastErrorMsg);
      FreeAndNil(fTLesGrilles);
      Timer1.Enabled := True;
      exit;
    end;

    if Crit.bExport then
    begin
      if SD.Execute then
      begin
        ExportVersExcel(fTLesGrilles, SD.FileName, False);
      end;
    end else
    begin
      { Mémorisation du paramétrage de l'utilisateur }
      bOldPDF := V_PGI.QRPDF;
      bOldNoPrintDialog := V_PGI.NoPrintDialog;
      bOldSilentMode := V_PGI.SilentMode;

      { Activation du mode Pdf et sauvegarde des valeurs initiales }
      V_PGI.QRPDF := True;
      V_PGI.NoPrintDialog := True;
      V_PGI.SilentMode := True;

      { Constitution de l'aperçu sur le disque }
      if (FCritEdtchaine.Utiliser) and (FCritEdtChaine.NomPDF <> '') then
      begin
        for iCopie := 0 to fCritEdtChaine.NombreExemplaire - 1 do
          stFileName := ImprimeLesGrilles;
        sleep(2000); // Temporisation pour éviter de 'perdre' certaines pages de l'état en édition PDF
      end else stFileName := ImprimeLesGrilles;

      { Affichage de l'aperçu avant impression }

{$IFDEF GPL_PLAQPGI} //cv4 le 19/06/2007 -> impression dans GPL -> NE PAS ENLEVER CE CODE SVP !!!!!
      if V_PGILIA.BOOK then
      begin
        V_PGILIA.BOOKPATH := '';
        V_PGILIA.BOOKPATH := LeGdP.RecupPrintPDF(stFileName);
      end
      else
      begin
{$ENDIF GPL_PLAQPGI}
        if not FCritEdtchaine.Utiliser then PreviewPDFFile('Etat de synthèse', stFileName, True)
        else if FCritEdtChaine.NomPDF = '' then
        begin
          if FCritEdtchaine.Utiliser then
          begin
            for iCopie := 0 to fCritEdtChaine.NombreExemplaire - 1 do
              PrintPDF(stFileName, Printer.Printers[Printer.PrinterIndex], '');
          end
          else PrintPDF(stFileName, Printer.Printers[Printer.PrinterIndex], '');
        end;
      { On renomme le fichier pour Publifi Edition }
        RenommePourPublifi;
{$IFDEF GPL_PLAQPGI} //cv4 le 19/06/2007 -> impression dans GPL -> NE PAS ENLEVER CE CODE SVP !!!!!
      end;
{$ENDIF GPL_PLAQPGI}


      { Retour aux valeurs initiales }
      V_PGI.QrPdf := bOldPDF;
      V_PGI.NoPrintDialog := bOldNoPrintDialog;
      V_PGI.SilentMode := bOldSilentMode;
      V_PGI.QRPDFMerge := '';
      V_PGI.QRPDFQueue := '';
    end;
  finally
    FreeAndNil(fTLesGrilles);
  end;
  // CA - 26/02/2007 - Pour éviter le problème de focus en CascadeForms
  HMTrad.ActiveResize := False;
  Timer1.Enabled := True;
end;

procedure TFEtatLiasse.FormDestroy(Sender: TObject);
begin
//  FListeDesComptes.Free;
  FListeDesSections.Free;
  freeandnil(flafonte);
  FTL.Free;
  FLaTOB.Free;
  FDetailTOB.Free;
  fLesRubriques.Free;
//  fLesRubriquesMem.Free;
end;

procedure TFEtatLiasse.FormCreate(Sender: TObject);
begin
  flafonte := tfont.Create;
  FDetailCalcule := FALSE;
  FNonAffectes := False;
  FTL := HTStringList.Create;
  FLaTOB := TOB.Create('Les rubriques', nil, -1);
  FDetailTOB := TOB.Create('Les rubriques detail', nil, -1);
//  FListeDesComptes := TOB.Create('', nil, -1);
  FListeDesSections := TOB.Create('', nil, -1);
  fbStopEdition := False;
  fLesRubriques := TZCumul.Create;
//  fLesRubriquesMem := TZCumulRubrique.Create;
  fLesRubriques.ModeFonc := mfMemory;
  fLesRubriques.AvecLibelle := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 03/01/2003
Modifié le ... :   /  /
Description .. : Chargement de la maquette d'édition
Mots clefs ... : ETAT;SYNTHESE;MAQUETTE
*****************************************************************}

procedure TFEtatLiasse.ChargeMaquette;
var
  First: boolean;
  i, nb, ih, kk: integer;
  St, F1, FA: string;
  T: TOB;
begin
  FTL.LoadFromFile(FModeleName);
//  FLaTOB.ClearDetail;
//  FDetailTOB.ClearDetail;
  { Mémorisation du titre de l'état - FQ 17608 }
  FaitTitreMaquette(FTL[0]);

  Caption := '';

  First := TRUE;
  FTL.Delete(0);
  FTL.Delete(0);
//  FListe.RowCount := 2;
  VidePileGrid(FListe, True);
  VidePileGrid(FDetail, True);
  for i := 0 to FTL.Count - 1 do
  begin
    St := FTL[i];
    T := TOB.Create('Une TOB', FLaTob, -1);
    T.AddChampSupValeur('TYPE', Trim(UpperCase(ReadTokenPipe(St, '|'))));
    T.AddChampSupValeur('ALIGN', Trim(UpperCase(ReadTokenPipe(St, '|'))));
    T.AddChampSupValeur('STYLE', Trim(UpperCase(ReadTokenPipe(St, '|'))));
    T.AddChampSupValeur('LIBELLE', Trim(ReadTokenPipe(St, '|')));
    T.AddChampSupValeur('NOM', Trim(UpperCase(ReadTokenPipe(St, '|'))));
    T.AddChampSupValeur('BASE', Trim(UpperCase(ReadTokenPipe(St, '|'))));
    for kk := 1 to MAX_COL do
    begin
      FA := '';
      if kk = 1 then
        F1 := Trim(UpperCase(ReadTokenPipe(St, '|')))
      else
        FA := Trim(UpperCase(ReadTokenPipe(St, '|')));
      if FA = '' then
        FA := F1;
      T.AddChampSupValeur('FORMULE' + IntToStr(kk), FA);
      T.AddChampSupValeur('VALEUR' + IntToStr(kk), 0.00);
    end;
    ih := ValeurI(ReadTokenPipe(St, '|'));
    if ih <= 0 then
      ih := 18; //RowHeight
    if (T.GetValue('TYPE') <> 'I') and (T.GetValue('TYPE') <> 'L') then
    begin
      if not First then
        FListe.RowCount := FListe.RowCount + 1
      else
        First := FALSE;
      nb := FListe.RowCount - 1;
    end
    else
      Nb := -1;
    T.AddChampSupValeur('LIGNE', nb);
    if nb > 0 then
    begin
      if POS_LIBELLE >= 0 then
        FListe.Cells[POS_LIBELLE, nb] := DecaleLibelle(T.GetValue('LIBELLE'));
      if POS_NOM >= 0 then
        FListe.Cells[POS_NOM, nb] := T.GetValue('NOM');
      FListe.Objects[0, nb] := T;
      FListe.RowHeights[nb] := ih;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 15/03/2006
Modifié le ... :   /  /
Description .. : Décompose la première ligne de la maquette d'édition pour
Suite ........ : en extraire les titres de l'état. Cette ligne doit être découpée
Suite ........ : de la manière suivante :
Suite ........ : - Pour CR et SIG : TITRE_ETAT;TITRE_ETAT_DETAIL;
Suite ........ : - Pour BIL :
Suite ........ : TITRE_ETAT_ACTIF;TITRE_ETAT_ACTIF_DETAIL;TITR
Suite ........ : E_ETAT_PASSIF;TITRE_ETAT_PASSIF_DETAIL
Mots clefs ... :
*****************************************************************}

procedure TFEtatLiasse.FaitTitreMaquette(St: string);
var
  stTitre: string;
  i, pT: integer;
begin
  if St = '' then St := TraduireMemoire('Titre non défini');
  { Titre de la maquette }
  pT := Pos('(', St);
  if pT > 0 then
  begin
    fTitreMaquette[0] := Copy(St, 1, pT - 1);
    St := Copy(St, pT + 1, Length(St) - pT - 1);
  end;
  { Titre des états }
  for i := 1 to 4 do
  begin
    stTitre := ReadTokenSt(St);
    if (stTitre <> '') then
      fTitreMaquette[i] := stTitre
    else if (i > 1) then fTitreMaquette[i] := fTitreMaquette[i - 1];
  end;
end;

procedure TFEtatLiasse.Charge;
begin

  ChargeMaquette;

  CalculRub;

  if fLesRubriques.LastError <> 0 then exit;

  case Crit.TE of
    esBIL, esBILA: AfficheMentionBilan;
    esCR, esSIG, esCRA, esSIGA: AfficheMentionResultat;
  end;

  if PDetails.TabVisible then
  begin
    ChargeDetail;
  end;

  if FBEcartEtat then
  begin
    ChargeDetail; // ChargeDetail détecte seul si le chargement a déjà été fait.
    ChargeNonAffectes;
  end;

  PNonAffectes.TabVisible := FBEcartEtat;
  AjusteTailleColonne(FListe, False);
  if PDetails.TabVisible then
    AjusteTailleColonne(FDetail, False);
end;

procedure TFEtatLiasse.CalculRub;
var
  kk, i, j, nbl: Integer;
  T: TOB;
  okok: boolean;
  bColonnePourcent: boolean;
  tipe, St, StF, StFF, stMask: string;
  X: Double;
  VV: Variant;
  StErr, stErr2: string;
  LesCptes: tStringList;
begin
  //InitMove(FLaTob.Detail.Count*4*2,'') ;
  InitMove(1000, '');
  for kk := 1 to 4 do
    for i := 0 to FLaTob.Detail.Count - 1 do
    begin
      MoveCur(FALSE);
      T := FLaTOB.Detail[i];
      FCurrentTob := T;
      nbl := T.GetValue('LIGNE');
      tipe := T.GetValue('TYPE');
      // Ajout CA le 20/02/2001
      if T.GetValue('NOM') = CHPCTRL_ACTIF then
        FLigneFinActif := T.GetValue('LIGNE');
      // Fin Ajout CA le 20/02/2001
      if (tipe <> 'C') then
      begin
        for j := 1 to FNbCol do
        begin
          FCurrentCol := j;
          okok := FALSE;
          St := T.GetValue('FORMULE' + IntToStr(j));
          StFF := FFormuleCol[j];
          X := 0;
          stMask := '#,##0.00'; // Obligatoire pour les calculs
          if (kk = 1) and (St <> '') and (Copy(St, 1, 1) <> '=') and (Copy(StFF,
            1, 1) <> '=') then
          begin
            StErr := '';
            LesCptes := nil;
            VV := LeCumul(St, j, LesCptes);
            if fLesRubriques.LastError <> 0 then exit;
            OkOk := TRUE;
            VV := QuelleUnite(VV);
            if LesCptes <> nil then
            begin
              if NBL > 0 then
                FListe.Objects[POS_COL + j - 1, nbl] := LesCptes
              else
                LesCptes.Free;
            end;
            (* FQ 17421 - CA - 29/06/2006 : avec ZCumul, StOLEErr n'est plus mis à jour }
            if V_PGI.StOLEErr = '' then
              X := VV
            else
            begin
              StErr := VV;
              X := 0;
            end;
            *)
            X := VV
          end;
          if (kk = 2) and (Copy(St, 1, 1) = '=') and (Copy(StFF, 1, 1) <> '=')
            then
          begin
            StFF := St;
            Delete(StFF, 1, 1);
            StFF := '{"' + stMask + '"' + StFF + '}';
            StF := GFormule(StFF, GetCellCumul, nil, 1);
            X := Valeur(StF);
            okok := TRUE;
          end;
          if (kk = 3) and (St <> '') and (Copy(St, 1, 1) <> '=') and (Copy(StFF,
            1, 1) = '=') then
          begin
            Delete(StFF, 1, 1);
            StFF := '{"' + stMask + '"' + StFF + '}';
            StF := GFormule(StFF, GetCellCumul, nil, 1);
            X := Valeur(StF);
            okok := TRUE;
          end;
          if (kk = 4) and (Copy(St, 1, 1) = '=') and (Copy(StFF, 1, 1) = '=')
            then
          begin
            Delete(StFF, 1, 1);
            StFF := '{"' + stMask + '"' + StFF + '}';
            StF := GFormule(StFF, GetCellCumul, nil, 1);
            X := Valeur(StF);
            okok := TRUE;
          end;

          if okok then
          begin
            T.PutValue('VALEUR' + IntToStr(j), X);
            if nbl > 0 then
            begin
              bColonnePourcent := (Crit.AvecPourcent and ((j = 2) or (j = 4) or
                (j = 6) or (j = 8)));
              if StErr <> '' then
                FListe.Cells[POS_COL + j - 1, nbl] := StErr
              else
                FListe.Cells[POS_COL + j - 1, nbl] := FormatFloat(FFormatcol[j],
                  X);
              // Pour éviter les blancs soulignés !!!
              if FListe.Cells[POS_COL + j - 1, nbl] = ' ' then
                FListe.Cells[POS_COL + j - 1, nbl] := ''
                  // Mise à 0 des colonnes inutiles pour le passif
              else if ((Crit.TE = esBIL) and ((FLigneFinActif > 0) and (nbl >
                FLigneFinActif))
                and ((j = 2) or (j = 3))) then
                FListe.Cells[POS_COL + j - 1, nbl] := ''
                  // Si pourcent > 999 % non significatif
              else if (Crit.AvecPourcent and ((j = 2) or (j = 4) or (j = 6) or (j
                = 8)) and (abs(X) > 999.99)) then
                FListe.Cells[POS_COL + j - 1, nbl] := 'NS'
              { FQ 17582 - si 4 colonnes = on arrondi les pourcentages }
              else if (Crit.AvecPourcent and ((j = 2) or (j = 4) or (j = 6) or (j
                = 8)) and (FNbCol = 8)) then
                FListe.Cells[POS_COL + j - 1, nbl] := IntToStr(Round(X));
              if T.GetValue('NOM') = CHPCTRL_RESULTAT then
              begin
                if not bColonnePourcent then
                begin
                  FEcartResultat[j] := Arrondi(X - QuelleUnite(GetResultatSurPeriode(FFormuleCol[j],
                    stErr2, Crit.EnMonnaieOpposee, Crit.Etab, Crit.Col[j -
                    1].BalSit)), V_PGI.OkDecV);
                  if stErr2 <> '' then
                    FEcartResultat[j] := 0;
                end;
              end
              else if T.GetValue('NOM') = CHPCTRL_ACTIF then
              begin
                FBilanActif[j] := X;
              end
              else if T.GetValue('NOM') = CHPCTRL_PASSIF then
                FBilanPassif[j] := X;
            end;
          end;
        end;
      end;
    end;
  UpdateRowLengths(FListe, Crit.AvecMontant0);
  FiniMove;
end;

function GetNextRub(var StO, StE, StNext, StSigne: string): Boolean;
var
  i1, i2, j: Integer;
begin
  Result := FALSE;
  i1 := Pos('+', StO);
  i2 := Pos('-', StO);
  if (i1 > 0) and (i2 > 0) then
  begin
    if i1 > i2 then
      j := i2
    else
      j := i1;
  end
  else if i1 > 0 then
    j := i1
  else if i2 > 0 then
    j := i2
  else
    j := 0;
  StE := '';
  StNext := '';
  if StO = '' then
    Exit;
  if (i2 > 0) or (i1 > 0) then
  begin
    StSigne := Copy(StO, j, 1);
    StE := Copy(StO, 1, j - 1);
    StO := Copy(StO, j + 1, Length(StO) - j);
    StNext := StO;
  end
  else
    StE := StO;
  Result := TRUE;
end;

function TFEtatLiasse.LeCumul(Nom: string; LaCol: Integer; var LesCptes:
  tStringList): Variant;
var
  Appel: string;
  St1, StNext, StNom, StSigne, StSigneNext: string;
  OkOk: Boolean;
  vCumul: variant;
  StComplementTL, StComplementTLEx: string;
begin
  Result := 0;
  if Demo then
    Result := Valeur(Nom) * 100 * LaCol
  else
  begin
    if VH^.bAnalytique then
    begin
      if Crit.bTLI then
      begin
        Appel := '(' + #22 + Crit.stAxe + ';' + Crit.StSectionDe + ';' +
          Crit.StSectionA + ')RUBRIQUE';
        StComplementTL := AnalyseCompte(Crit.StSectionDe, AxeToFb(Crit.stAxe), False, True);
        if (Crit.StSectionA <> '') then
          StComplementTLEx := AnalyseCompte(Crit.StSectionA, AxeToFb(Crit.stAxe), False, True)
        else StComplementTLEx := '';
      end
      else
        Appel := '(' + Crit.stAxe + ';' + Crit.StSectionDe + ';' +
          Crit.StSectionA + ')RUBRIQUE';
    end
    else
      Appel := 'RUBRIQUE';
    stSigne := '+';
    if Crit.AvecDetail then
      LesCptes := TStringList.Create;
    St1 := Nom;
    repeat
      OkOk := GetNextRUB(St1, StNom, StNext, StSigneNext);
      if OkOk and (StNom <> '') then
      begin
        if fLesRubriques.InitCriteres(Crit.StTypEcr, Crit.Etab, Crit.Devise,
          FFormuleCol[LaCol], Crit.Col[LaCol - 1].BalSit) then
        begin
          (*
          fLesRubriquesMem.InitCriteres(Crit.StTypEcr, Crit.Etab, Crit.Devise,
          FFormuleCol[LaCol], Crit.Col[LaCol - 1].BalSit);
          *)
          if Crit.stAxe <> '' then
          begin
            if Crit.bTLI then
              fLesRubriques.InitAnalytique(Crit.stAxe, '', '') // Sinon chaîne TL trop longue pour HashTable
            else
              fLesRubriques.InitAnalytique(Crit.stAxe, Crit.StSectionDe, Crit.StSectionA);
            fLesRubriques.ComplementTL := stComplementTL;
            if stComplementTLEx <> '' then
              fLesRubriques.ComplementTL := fLesRubriques.ComplementTL + 'AND NOT(' + stComplementTLEx + ')';
            (*
            fLesRubriquesMem.InitAnalytique(Crit.stAxe, Crit.StSectionDe,
              Crit.StSectionA);
            *)
          end;
          (*
          if V_PGI.Sav then
            vCumul := fLesRubriquesMem.GetValeur( stNom, LesCptes)
          else
          *)
          vCumul := fLesRubriques.GetValeur('RUBRIQUE', stNom, LesCptes);
        end;

        (*  --- Ancien mode de calcul via GetCumul ---

        vCumul := Get_CumulPCL(Appel, StNom, Crit.StTypEcr, Crit.Etab, Crit.Devise, FFormuleCol[LaCol], '', Crit.DeviseEnPivot, Crit.EnMonnaieOpposee, LesCptes,
            Crit.Col[LaCol - 1].BalSit);

        *)
        if fLesRubriques.LastError <> 0 then exit;

        if stSigne = '+' then
          Result := Result + vCumul
        else if stSigne = '-' then
          Result := Result - vCumul;
        stSigne := stSigneNext;
      end;
    until StNext = '';
  end;
end;

function TrouveIndice(LC: tStringList; Cpt: string): Integer;
var
  i, l: Integer;
  St, Cpt1: string;
begin
  Result := -1;
  if LC = nil then
    Exit;
  for i := 0 to LC.count - 1 do
  begin
    St := LC[i];
    l := Pos(':', St);
    if l > 0 then
    begin
      Cpt1 := Copy(St, 1, l - 1);
      if Cpt = Cpt1 then
      begin
        Result := i;
        Break;
      end;
    end;
  end;
end;

function IsEntier(St: string): Boolean;
var
  i: Integer;
begin
  Result := TRUE;
  if St = '' then
  begin
    Result := FALSE;
    Exit;
  end;
  for i := 1 to length(St) do
  begin
    if (St[i] in ['0'..'9']) = FALSE then
    begin
      Result := FALSE;
      Exit;
    end;
  end;
end;

function TFEtatLiasse.LeCumulDetail(Cpt: string; LaCol, LaLig: Integer;
  bComparatifBilan: boolean): Double;
var
  LesCptes: tStringList;
  i, j, k: Integer;
  St: string;
  Tst: array[0..7] of string;
begin
  //FFormuleCol[LaCol]
  Result := 0;
  if Demo then
    Result := 2 * 100 * LaCol
  else
  begin
    LesCptes := tStringList(Fliste.Objects[LaCol, LaLig]);
    if LesCptes = nil then
      Exit;
    i := TrouveIndice(LesCptes, Cpt);
    if i = -1 then
      Exit;
    FillChar(Tst, SizeOf(Tst), #0);
    St := LesCptes[i];
    j := Pos(':', St);
    k := 0;
    if St[length(ST)] <> ':' then
      st := St + ':';
    while j > 0 do
    begin
      Tst[k] := Copy(St, 1, Pos(':', St) - 1);
      Delete(St, 1, Pos(':', St));
      Inc(k);
      if k > 7 then
        Break;
      j := Pos(':', St);
    end;
    if IsEntier(Tst[6]) then
    begin
      i := StrToInt(Tst[6]);
      case i of
        7: Result := QuelleUnite(Valeur(Tst[1]), True);
        6: Result := QuelleUnite(Valeur(Tst[2]), True);
        5: Result := QuelleUnite(Valeur(Tst[3]), True);
        3: Result := QuelleUnite(Valeur(Tst[4]), True);
        2: Result := QuelleUnite(Valeur(Tst[5]), True);
      end;
    end;

    { CA - 27/09/2004 - FQ 12201 }
    { CA - 10/11/2006 - FQ 12201 - à priori pas bon, on déplace après calcul du signe}
    { if bComparatifBilan then
      exit; // on ne retraite pas le signe dans le cas du comparatif du bilan. }

    { si on est en solde créditeur, on place le signe - }
    { pour se replacer dans le contexte débiteur et appliquer NEG ou POS }
    if (i = 3) or (i = 6) then
      Result := Result * (-1);

    { CA - 10/11/2006 - FQ 12201 - Pour avoir le bon signe}
    if bComparatifBilan then
      exit; // on ne retraite pas le signe dans le cas du comparatif du bilan.

    if (Result < 0) then // au crédit
    begin
      Result := Abs(Result);
      if (Copy(Tst[7], 1, 3) = 'POS') then // CA - 15/09/2006 - FQ 18629
        Result := Result * (-1);
    end
    else
    begin // au débit
      Result := Abs(Result);
      if (Copy(Tst[7], 1, 3) = 'NEG') then // CA - 15/09/2006 - FQ 18629
        Result := Result * (-1);
    end;
  end;
end;

function TFEtatLiasse.GetCellCumul(St: hstring): Variant;
var
  T: TOB;
  i: Integer;
  StB: string;
begin
  if Copy(St, 1, 3) = 'COL' then
  begin
    i := ValeurI(St);
    if FCurrentTOB <> nil then
      Result := FCurrentTOB.GetValue('VALEUR' + IntToStr(i))
    else
      result := 0;
  end
  else if Copy(St, 1, 4) = 'BASE' then
  begin
    if FCurrentTOB = nil then
    begin
      result := 0;
      exit;
    end;
    StB := FCurrentTOB.GetValue('BASE');
    T := FLaTob.FindFirst(['NOM'], [Stb], FALSE);
    if T <> nil then
      Result := T.GetValue('VALEUR' + IntToStr(ValeurI(St)))
    else
      result := 0;
  end
  else
  begin
    T := FLaTob.FindFirst(['NOM'], [St], FALSE);
    if T <> nil then
      Result := T.GetValue('VALEUR' + IntToStr(FCurrentCol))
    else
      result := 0;
  end;
end;

function GetMontantDetail(St: string; i: Integer): Double;
begin
  Result := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 27/01/2003
Modifié le ... :   /  /
Description .. : Charge les infos des comptes d'une rubrique  ( libellé )
Mots clefs ... :
*****************************************************************}

procedure TFEtatLiasse.ChargeDetailCompteRubrique(Noms, Libelles: TStringList;
  R: integer);
var
  i, k, l: Integer;
  LesCptes: tStringList;
  Cpt, Lib, Section: string;
  iPos, nbNoms: integer;
  lStLeDetail: string;
begin
  if Noms <> nil then
    Noms.Clear;
  if Libelles <> nil then
    Libelles.Clear;
  if (R < 0) then
    Exit;
  for i := 1 to FNbCol do
  begin
    if Fliste.Objects[POS_COL + i - 1, R] = nil then
      continue;
    LesCptes := TStringList(FListe.Objects[POS_COL + i - 1, R]);
    if LesCptes.Count = 0 then
      Continue;
    for k := 0 to LesCptes.Count - 1 do
    begin
      lStLeDetail := LesCptes[k];
      l := Pos(':', LesCptes[k]);
      if l > 0 then
      begin
        if (VH^.bAnalytique) then
        begin
          Section := Copy(LesCptes[k], 1, l - 1);
          nbNoms := Noms.Count;
            // CA - 22/03/2005 - on mémorise le nombre de noms pour tester le cas ou le noms existe déjà dans la liste
          iPos := Noms.Add(Section);
          if ((iPos >= 0) and (Noms.Count > nbNoms)) then
            { Insertion du code en position iPos }
          begin
            Cpt := ReadTokenSt(Section);
            if ((Crit.bLibelleSection) and (not Crit.bUnEtatParSection)) then
            begin
              if Crit.bLibelleCompte then
                Lib := ' - ' + Copy(LibelleDuCompte(LesCptes[k]), 1, 17) + ' - ' +
                  Copy(LibelleDeLaSection(Section), 1, 17)
              else
                Lib := ' - ' + Cpt + ' - ' + LibelleDeLaSection(Section);
            end
            else
            begin
              if Crit.bUnEtatParSection then
              begin
                if Crit.bLibelleCompte then
                  Lib := ' - ' + Cpt + ' - ' + LibelleDuCompte(LesCptes[k])
                else
                  Lib := ' - ' + Cpt;
              end
              else
              begin
                if Crit.bLibelleCompte then
                  Cpt := LibelleDuCompte(Cpt);
                if Crit.bLibelleSection then
                  Lib := ' - ' + Cpt + ' - ' + LibelleDeLaSection(Section)
                else
                  Lib := ' - ' + Cpt + ' - ' + Section;
              end;
            end;
            // Libelles.Add(Lib);
            Libelles.Insert(iPos, Lib);
              // Insertion du libellé à la même position que le code
          end;
        end
        else
        begin
          Cpt := Copy(LesCptes[k], 1, l - 1);
          Noms.Add(Cpt);
          Lib := ' - ' + Cpt + ' ' + LibelleDuCompte(lStLeDetail);
          Libelles.Add(Lib);
        end;
      end;
    end;
  end;
end;

{procedure TFEtatLiasse.GetRubDetail(Nom: String; Noms,Libelles: TStringList ; R : Integer);
Var i,k,l : Integer ;
    LesCptes : tStringList ;
    Cpt,Lib : String ;
    Q : tQuery ;
begin
  Noms.Clear ;
  Libelles.Clear ;
  if (R < 0) then Exit ;
  For i:=1 to FNbCol do
  BEGIN
    If Fliste.Objects[POS_COL+i-1,R]=Nil Then Continue ;
    LesCptes:=TStringList(FListe.Objects[POS_COL+i-1,R]) ;
    If LesCptes.Count=0 Then Continue ;
    For k:=0 To LesCptes.Count-1 Do
    BEGIN
      l:=Pos(':',LesCptes[k]) ;
      If l>0 Then
      BEGIN
        Cpt:=Copy(LesCptes[k],1,l-1) ;
        Noms.Add(Cpt) ;
        if not Crit.SansCompte then Lib:=' - '+Cpt
        else Lib := ' - ';
//        Q:=OpenSQL('SELECT G_LIBELLE FROM GENERAUX WHERE G_GENERAL="'+Cpt+'" ',TRUE) ;
//        If Not Q.Eof Then Lib:=Lib+' '+Q.Fields[0].AsString ;
//        Ferme(Q) ;
        Lib := Lib + ' ' + LibelleDuCompte( Cpt );
        Libelles.Add(Lib) ;
      END ;
    END ;
  END ;
end;}

procedure TFEtatLiasse.CalculRubDetail;
var
  kk, i, j, nbl, NblPrinc: Integer;
  T: TOB;
  okok: boolean;
  tipe, St, StF, StFF, stMask: string;
  X: Double;
  StCpt: string;
begin
  //InitMove(FLaTob.Detail.Count*4,'') ;
  for kk := 1 to 4 do
    for i := 0 to FDetailTob.Detail.Count - 1 do
    begin
      MoveCur(FALSE);
      T := FDetailTob.Detail[i];
      FCurrentTob := T;
      if T.GetValue('NOM') = CHPCTRL_ACTIF then
        FLigneFinDetailActif := T.GetValue('LIGNE');
      nbl := T.GetValue('LIGNE');
      tipe := T.GetValue('TYPE');
      for j := 1 to FNbCol do
      begin
        St := T.GetValue('FORMULE' + IntToStr(j));
        StCpt := T.GetValue('NOM');
        StFF := FFormuleCol[j];
        X := 0;
        stMask := '#,##0.00'; // Obligatoire pour les calculs
        if (tipe = 'D') then
        begin
          FCurrentCol := j;
          okok := FALSE;
          if (kk = 1) and (St <> '') and (Copy(St, 1, 1) <> '=') and (Copy(StFF,
            1, 1) <> '=') then
          begin
            nblPrinc := T.GetValue('LIGNEPRINC');
            if (((Crit.TE = esBIL) and (j = 4)) and ((nbl < FLigneFinDetailActif)
              or (FLigneFinDetailActif = 0))) then
              // Cas du bilan ACTIF en comparatif
              X := LeCumulDetail(StCpt, POS_COL + j - 1, NbLPrinc, True)
            else
              X := LeCumulDetail(StCpt, POS_COL + j - 1, NbLPrinc, False);
            okok := TRUE;
          end;
          if (kk = 2) and (Copy(St, 1, 1) = '=') and (Copy(StFF, 1, 1) <> '=')
            then
          begin
            StFF := St;
            Delete(StFF, 1, 1);
            StFF := '{"' + stMask + '"' + StFF + '}';
            StF := GFormule(StFF, GetCellCumul, nil, 1);
            X := Valeur(StF);
            okok := TRUE;
          end;
          if (kk = 3) and (St <> '') and (Copy(St, 1, 1) <> '=') and (Copy(StFF,
            1, 1) = '=') then
          begin
            Delete(StFF, 1, 1);
            StFF := '{"' + stMask + '"' + StFF + '}';
            StF := GFormule(StFF, GetCellCumul, nil, 1);
            X := Valeur(StF);
            okok := TRUE;
          end;
          if (kk = 4) and (Copy(St, 1, 1) = '=') and (Copy(StFF, 1, 1) = '=')
            then
          begin
            Delete(StFF, 1, 1);
            StFF := '{"' + stMask + '"' + StFF + '}';
            StF := GFormule(StFF, GetCellCumul, nil, 1);
            X := Valeur(StF);
            okok := TRUE;
          end;

          if okok then
          begin
            T.PutValue('VALEUR' + IntToStr(j), X);
            if nbl > 0 then
              FDetail.Cells[POS_COL + j - 1, nbl] :=
                FormatFloat(FFormatcolDetail[j], X);
            // Pour éviter las blancs soulignés !!!
            if FDetail.Cells[POS_COL + j - 1, nbl] = ' ' then
              FDetail.Cells[POS_COL + j - 1, nbl] := ''
                // Mise à 0 des colonnes inutiles pour le passif (détail)
            else if ((Crit.TE = esBIL) and ((FLigneFinDetailActif > 0) and (nbl
              > FLigneFinDetailActif))
              and ((j = 2) or (j = 3))) then
              FDetail.Cells[POS_COL + j - 1, nbl] := ''
                // Si pourcent > 999 % non significatif
            else if (Crit.AvecPourcent and ((j = 2) or (j = 4) or (j = 6) or (j
              = 8)) and (abs(X) > 999.99)) then
              FDetail.Cells[POS_COL + j - 1, nbl] := 'NS'
            { FQ 17582 - si 4 colonnes = on arrondi les pourcentages }
            else if (Crit.AvecPourcent and ((j = 2) or (j = 4) or (j = 6) or (j
              = 8)) and (FNbCol = 8) and (X <> 0)) then
              FDetail.Cells[POS_COL + j - 1, nbl] := IntToStr(Round(X));
          end;
        end
        else
        begin
          if (tipe = 'F') or (tipe = 'R') or (tipe = 'L') then
            if (St <> '') or (Copy(StFF, 1, 1) = '=') then
            begin
              if ((Crit.TE = esBIL) and ((FLigneFinDetailActif > 0) and (nbl >
                FLigneFinDetailActif))
                and ((j = 2) or (j = 3))) then
                FDetail.Cells[POS_COL + j - 1, nbl] := ''
                  // Si pourcent > 999 % non significatif
              else if (Crit.AvecPourcent and ((j = 2) or (j = 4) or (j = 6) or (j
                = 8)) and (Abs(T.GetValue('VALEUR' + IntToStr(j))) > 999.99)) then
                FDetail.Cells[POS_COL + j - 1, nbl] := 'NS'
            { FQ 17582 - si 4 colonnes = on arrondi les pourcentages }
              else if (Crit.AvecPourcent and ((j = 2) or (j = 4) or (j = 6) or (j
                = 8)) and (FNbCol = 8) and (T.GetValue('VALEUR' + IntToStr(j)) <> 0)) then
                FDetail.Cells[POS_COL + j - 1, nbl] := IntToStr(Round(T.GetValue('VALEUR' + IntToStr(j))))
              else if nbl > 0 then
                FDetail.Cells[POS_COL + j - 1, nbl] :=
                  FormatFloat(FFormatcolDetail[j], T.GetValue('VALEUR' +
                  IntToStr(j)));
            end;
        end;
      end;
    end;
  // UpdateRowLengths(FDetail, Crit.AvecMontant0);
  // CA - 18/11/2002 - On n'affiche jamais les lignes à 0 du détail
  UpdateRowLengths(FDetail, False);
  FiniMove;
end;

procedure TFEtatLiasse.ChargeDetail;
var
  i, nb, j, ih, ih1, kk: Integer;
  First: Boolean;
  T1, T2, T3: TOB;
  TN, TL: TStringList;
  StF, stLibelleSansCompte: string;
begin
  if FDetailCalcule then
    exit;
  FDetailTOB.ClearDetail;
  First := TRUE;
  FDetailCalcule := TRUE;
  FDetail.RowCount := 2;
  TN := TStringList.Create;
  TL := TStringList.Create;
  TN.Sorted := TRUE;
  TN.Duplicates := DupIgnore;
  if (VH^.bAnalytique) then
  begin
    TL.Sorted := False;
    TL.Duplicates := DupAccept;
  end
  else
  begin
    TL.Sorted := True;
    TL.Duplicates := DupIgnore;
  end;

  (*if (VH^.bAnalytique) then
    begin // - CA 16/03/2004 - FQ 13347
      TL.Duplicates := DupAccept;
      TN.Duplicates  := DupAccept;
    end
    else *)
//  ChargeListeDesComptes;
  if VH^.bAnalytique then
    ChargeListeDesSections;
  for i := 0 to FLaTob.Detail.Count - 1 do
  begin
    T1 := FLaTob.Detail[i];
    ih := 18;
    ih1 := T1.GetValue('LIGNE');
    if ih1 > 0 then
      ih := FListe.RowHeights[ih1];
    if (T1.GetValue('TYPE') = 'R') or (T1.GetValue('TYPE') = 'L') then
    begin
      // GetRubDetail(T1.GetValue('FORMULE1'),TN,TL,ih1) ;
      ChargeDetailCompteRubrique(TN, TL, ih1);
      for j := 0 to TN.Count - 1 do
      begin
        T3 := TOB.Create('Une TOB detail', FDetailTob, -1);
        T3.AddChampSupValeur('TYPE', 'D');
        T3.AddChampSupValeur('ALIGN', T1.GetValue('ALIGN'));
        T3.AddChampSupValeur('STYLE', T1.GetValue('STYLE'));
        if (Crit.bLibelleCompte and (not VH^.bAnalytique)) then
        begin
          stLibelleSansCompte := TL[j];
          Delete(stLibelleSansCompte, 1, 3 + VH^.Cpta[fbGene].Lg + 1);
          T3.AddChampSupValeur('LIBELLE', ' - ' + stLibelleSansCompte);
        end
        else
          T3.AddChampSupValeur('LIBELLE', TL[j]);
        T3.AddChampSupValeur('NOM', TN[j]);
        T3.AddChampSupValeur('BASE', T1.GetValue('BASE'));
        for kk := 1 to MAX_COL do
        begin
          StF := T1.GetValue('FORMULE' + IntToStr(kk));
          if StF = '' then
            T3.AddChampSupValeur('FORMULE' + IntToStr(kk), '')
          else if Copy(StF, 1, 1) <> '=' then
            T3.AddChampSupValeur('FORMULE' + IntToStr(kk), TN[j])
          else
            T3.AddChampSupValeur('FORMULE' + IntToStr(kk), T1.GetValue('FORMULE'
              + IntToStr(kk)));
          T3.AddChampSupValeur('VALEUR' + IntToStr(kk), 0.00);
        end;
        if not First then
          FDetail.RowCount := FDetail.RowCount + 1
        else
          First := FALSE;
        nb := FDetail.RowCount - 1;
        T3.AddChampSupValeur('LIGNE', nb);
        T3.AddChampSupValeur('LIGNEPRINC', ih1);
        //        T3.SaveToFile ('c:\test.txt',True,True,True );
        if POS_LIBELLE >= 0 then
          FDetail.Cells[POS_LIBELLE, nb] :=
            DecaleLibelle(T3.GetValue('LIBELLE'));
        if POS_NOM >= 0 then
          FDetail.Cells[POS_NOM, nb] := T3.GetValue('NOM');
        FDetail.Objects[0, nb] := T3;
        FDetail.RowHeights[nb] := ih;
      end;
    end;
    T2 := TOB.Create('Une TOB', FDetailTob, -1);
    T2.Dupliquer(T1, FALSE, TRUE);
    T2.PutValue('STYLE', AddGras(T2.GetValue('STYLE')));
    if (T2.GetValue('TYPE') <> 'I') and (T2.GetValue('TYPE') <> 'L') then
    begin
      if not First then
        FDetail.RowCount := FDetail.RowCount + 1
      else
        First := FALSE;
      nb := FDetail.RowCount - 1;
    end
    else
      Nb := -1;
    T2.AddChampSupValeur('LIGNE', nb);
    if nb > 0 then
    begin
      if POS_LIBELLE >= 0 then
        FDetail.Cells[POS_LIBELLE, nb] := DecaleLibelle(T2.GetValue('LIBELLE'));
      if POS_NOM >= 0 then
        FDetail.Cells[POS_NOM, nb] := T2.GetValue('NOM');
      FDetail.Objects[0, nb] := T2;
      FDetail.RowHeights[nb] := ih;
    end;
  end;
  TN.Free;
  TL.Free;
  CalculRubDetail;
  {  for i:=1 To FNbCol Do
    begin
        FDetail.Cells[POS_COL+i-1,0]:=Crit.Col[i-1].StTitre ;
    END ; }
end;

type
  TXLabel = class(TLabel);

procedure TFEtatLiasse.AfficheCell(FromPrinter: boolean; lecanvas: tcanvas;
  letag: integer; letext: string; lestyle: string; f: TAlignment; lafonte: tfont;
  ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState; fx, fy: integer);

  procedure AfficheText(C: TCanvas; LocAlign: TAlignment; Text: string; R:
    TRect);
  var
    Ht, Wt, X, Y, H: Integer;
    S, St: string;
    N, I: Integer;
  begin
    Windows.SetTextalign(C.Handle, TA_NOUPDATECP);

    N := THPrintGridFacility.AglCountChar(#13, Text) + 1;

    { Pas de text }
    if N = 0 then
      Exit;

    { Hauteur d'une ligne }
    H := (R.Bottom - R.Top) div n;

    S := StringReplace(Text, #13#10, '£', [rfReplaceAll]);
    St := ReadTokenPipe(S, '£');
    X := 0;
    I := 0;
    while St <> '' do
    begin
      { Calcul largeur et hauteur du texte en fonction de la fonte en cours dans le canvas }
      Wt := C.TextWidth(St);
      Ht := C.TextHeight(St);
      { Centrage vertical }
      with R do
        Y := (H - Ht) div 2;

      case LocAlign of
        taLeftJustify: ;
        taRightJustify: if (R.Right - R.Left) > Wt then
            X := ((R.Right - R.Left) - Wt)
          else
            X := 2;
        taCenter: X := ((R.Right - R.Left) - Wt) div 2;
      end;

      if X < 0 then
        X := 0;

      ExtTextOut(C.Handle, R.Left + X, (R.Top + (I * H)) + Y, ETO_CLIPPED, @R,
        pchar(St), StrLen(pchar(St)), nil);

      St := ReadTokenPipe(S, '£');
      Inc(I);
    end;
  end;

var
  //    St     : String ;
  OnTitle: Boolean;
  OkCrit: Boolean;
  //    oldfont: thandle ;
  afont: tfont;
  CouleurDuFond: TColor;
begin
  OnTitle := arow = 0;

  { décodage du style }
  afont := tfont.create;
  afont.assign(lecanvas.font);
  lecanvas.font.assign(lafonte);

  { la couleur }
  if (gdfixed in State) then
    CouleurDuFond := fliste.FixedColor
  else
  begin
    if ((gdSelected in State) or (gdFocused in State)) then
    begin
      CouleurDuFond := clHighlight;
      leCanvas.Font.Color := clHighlightText;
    end
    else
    begin
      CouleurDuFond := FTestFont.Color;
    end;
  end;

  if OnTitle then
  begin
    //    OkCrit := (ACol >= Pos_Col) and (letag = 1);
    // CA - 21/02/2005 - Nouvelle gestion des tags
    OkCrit := (ACol >= Pos_Col) and ((letag = 1) or (letag = 2));
    if OkCrit then
      LeText := Crit.Col[ACol - Pos_Col].StTitre
        //    else if (leTag = 2) then LeText := FTitreNA[ACol]
      // CA - 21/02/2005 - Nouvelle gestion des tags
    else if (leTag = 3) then
      LeText := FTitreNA[ACol]
        //    else if (leTag = 3) then LeText := FTitrePassif[ACol]
      // CA - 21/02/2005 - Nouvelle gestion des tags
    else if ((leTag = 4) or (leTag = 5)) then
      LeText := FTitrePassif[ACol]
        { Fin Ajout CA le 20/02/20001 }
    else
      LeText := Text;

    { CA - 13/04/2005 - Ne plus afficher le titre dans la première cellule
                        sauf pour les non affectés }
    if (ACol = 0) and (leTag <> 3) then
      LeText := '';

    LeCanvas.Brush.Color := $00E6E6E6;
    LeCanvas.FillRect(Rect);
    // Appel de la méthode d'affichage du TXLabel
    LeCanvas.Font.Style := LeCanvas.Font.Style + [fsbold];
    InflateRect(Rect, -2, -2);
    AfficheText(lecanvas, tacenter, letext, Rect);
  end
  else
  begin
    lecanvas.brush.color := CouleurDuFond;
    case f of
      taRightJustify: ExtTextOut(lecanvas.handle, Rect.Right -
          lecanvas.TextWidth(letext) - (8 * fx), Rect.Top + (2 * fy), ETO_OPAQUE or
          ETO_CLIPPED, @Rect, PChar(leText),
          Length(leText), nil);
      taCenter: ExtTextOut(lecanvas.handle, Rect.Left + ((Rect.Right - Rect.Left
          - lecanvas.TextWidth(leText)) div 2), Rect.Top + 2, ETO_OPAQUE or
          ETO_CLIPPED, @Rect,
          PChar(leText), Length(leText), nil);
    else
      ExtTextOut(lecanvas.handle, Rect.Left + (2 * fx), Rect.Top + (2 * fy),
        ETO_OPAQUE or ETO_CLIPPED, @Rect, PChar(leText), Length(leText), nil);
    end;
  end;

  { traitement particulier des lignes avec un style }
  if (leStyle <> '') and not ontitle then
  begin
    leCanvas.Pen.Style := psSolid;
    leCanvas.Pen.Color := clBlack;
    if FromPrinter then
      lecanvas.pen.width := 5;
    if akTop in FTestFont.Anchors then
    begin
      leCanvas.MoveTo(Rect.Left, Rect.Top);
      leCanvas.LineTo(Rect.Right + 1, Rect.Top);
    end;
    if akBottom in FTestFont.Anchors then
    begin
      leCanvas.MoveTo(Rect.Left, Rect.Bottom);
      leCanvas.LineTo(Rect.Right + 1, Rect.Bottom);
    end;
  end;

  lecanvas.font.assign(afont);
  freeandnil(afont);
end;

procedure TFEtatLiasse.PrepareCell(legrid: THgrid; var letag: integer; var
  lestyle: string; var letext: string; var laligment: TAlignment; lafonte: tfont;
  State:
  TGridDrawState; acol, arow: integer);
var
  s: string;
  t: tob;
begin
  letag := legrid.tag;
  if gdfixed in state then
  begin
    lestyle := 'Arial;8;;clBlack;clWhite;----;';
    laligment := tacenter;
  end
  else
  begin
    T := TOB(legrid.Objects[0, ARow]);
    if t = nil then
      exit;
    lestyle := T.GetValue('STYLE');
    //    if ((ACol = POS_NOM) or ((leTag = 2) and (ACol = 1))) then laligment := taLeftJustify
    // CA - 21/02/2005 - Nouvelle gestion des tags
    if ((ACol = POS_NOM) or ((leTag = 3) and (ACol = 1))) then
      laligment := taLeftJustify
    else if ACol <> POS_LIBELLE then
      laligment := taRightJustify
    else
    begin
      s := T.GetValue('ALIGN');
      if S = 'D' then
        laligment := taRightJustify
      else if S = 'C' then
        laligment := taCenter
      else
        laligment := taLeftJustify;
    end;
  end;
  if leStyle = '' then
    leStyle := Default_Style;

  { le text }
  letext := legrid.Cells[ACol, ARow];
  if (letext = #0) or ((ACol = POS_LIBELLE) and (letext = '.')) then
    letext := '';

  { Décodage de la fonte }
  DecodeFont(lafonte, lestyle);
end;

procedure TFEtatLiasse.FListeDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  lestyle: string;
  letag: integer;
  letext: string;
  laligment: talignment;
begin
  { préparation de la cellule }
  preparecell(thgrid(sender), letag, lestyle, letext, laligment,
    ftestfont.font, state, acol, arow);

  AfficheCell(False, THGrid(sender).canvas, letag, letext, lestyle,
    laligment, ftestfont.font, ACol, ARow, Rect, State, 1, 1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 16/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFEtatLiasse.DoWriteCell(const LeGrid: THGrid; const ARow, ACol:
  integer; TMere: TOB);
var
  CouleurFond: TColor;
  LeText: string;
  Align: TAlignment;
  T: TOB;
  LeStyle: string;
  Ancrage: TAnchors;
  s: string;
begin

  // Init
  Ancrage := [];
  Align := taLeftJustify;
  LeStyle := Default_Style;

  // Le Text par défaut
  LeText := Legrid.Cells[ACol, ARow];
  if (LeText = #0) or ((Acol = POS_LIBELLE) and (LeText = '.')) then
    LeText := '';

  // Décodage de la fonte
  DecodeFont(FLaFonte, LeStyle);

  // Colonne ou ligne fixe
  if (Arow < LeGrid.FixedRows) or (ACol < LeGrid.FixedCols) then
  begin
    CouleurFond := LeGrid.FixedColor;
    LeStyle := 'Arial;8;;clBlack;clWhite;----;';
    Align := taCenter;

    // Décodage de la fonte
    DecodeFont(FLaFonte, LeStyle);
  end
  else
  begin
    t := tob(LeGrid.Objects[0, ARow]);
    if assigned(t) then
    begin
      if t.FieldExists('STYLE') then
        lestyle := t.GetValue('STYLE')
      else
        LeStyle := '';

      // Décodage de la fonte
      if LeStyle <> '' then
        DecodeFont(FLaFonte, LeStyle);

      //      if ((ACol = POS_NOM) or ((LeGrid.Tag = 2) and (ACol = 1))) then Align := taLeftJustify
      // CA - 21/02/2005 - Nouvelle gestion des tags
      if ((ACol = POS_NOM) or ((LeGrid.Tag = 3) and (ACol = 1))) then
        Align := taLeftJustify
      else if ACol <> POS_LIBELLE then
        Align := taRightJustify
      else
      begin
        if T.FieldExists('ALIGN') then
          s := T.GetValue('ALIGN')
        else
          S := '';
        if S = 'D' then
          Align := taRightJustify
        else if S = 'C' then
          Align := taCenter
        else
          Align := taLeftJustify;
      end;
    end;
    CouleurFond := FTestFont.Color;
  end;

  // Je suis sur la première ligne
  if ARow = 0 then
  begin
    // if (ACol >= Pos_Col) and (LeGrid.tag = 1) then LeText := Crit.Col[ACol - Pos_Col].StTitre
    // CA - 21/02/2005 - Nouvelle gestion des tags
    if (ACol >= Pos_Col) and ((LeGrid.tag = 1) or (LeGrid.tag = 2)) then
      LeText := Crit.Col[ACol - Pos_Col].StTitre
        //    else if (LeGrid.Tag = 2) then LeText := FTitreNA[ACol]
    else if (LeGrid.Tag = 3) then
      LeText := FTitreNA[ACol]
        //    else if (LeGrid.Tag = 3) then LeText := FTitrePassif[ACol]
    else if ((LeGrid.Tag = 4) or (LeGrid.Tag = 5)) then
      LeText := FTitrePassif[ACol]
    else if (LeGrid.Tag = 0) then // Cas de l'export du modèle
      LeText := GParams.Cells[ACol, 0]
    else Letext := Text;
    CouleurFond := $00E6E6E6;
  end;

  // Cas des lignes et des colonnes fixent
  if (ACol < LeGrid.FixedCols) or (ARow < LeGrid.FixedRows) then
  begin
    Ancrage := [akLeft, akTop, akBottom, akRight];
  end
  else
  begin
    Ancrage := [akLeft, akRight];
  end;

  // traitement particulier des lignes avec un style
  if (leStyle <> '') and (ARow > 0) then
  begin
    if akTop in FTestFont.Anchors then
      Ancrage := Ancrage + [akTop];
    if akBottom in FTestFont.Anchors then
      Ancrage := Ancrage + [akBottom];
  end;

  THPrintGridFacility.AglAddCellToTob(TMere, ACol, ARow, LeText, FLaFonte,
    CouleurFond, Align, Ancrage);
end;

procedure TFEtatLiasse.BImprimerClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  if (not (ctxStandard in V_PGI.PGIContexte)) then
    BCValideClick(nil);
{$ELSE}
  if (not (ctxStandard in V_PGI.PGIContexte)) then
    BCValideClick(nil)
  else
    PrintDBGrid(GParams, nil, 'Edition des maquettes', '');
{$ENDIF}
end;

procedure TFEtatLiasse.PagesChange(Sender: TObject);
var
  i, kk, ih: Integer;
  TL: HTStringList;
  St: string;
begin
  bSauve.Visible := Pages.ActivePage = PParams;
  if Pages.ActivePage = PDetails then
  begin
    ChargeDetail;
  end
  else if Pages.ActivePage = PParams then
  begin
    GParams.GetCellCanvas := GDrawCell;
    GParams.RowCount := 2;
    GParams.ColCount := 7 + MAX_COL;
    GParams.ColLengths[3] := -1;
    GParams.ColWidths[3] := 0;
    GParams.ColLengths[2] := -1;
    GParams.ColWidths[2] := 0;
    TL := HTStringList.Create;
    TL.LoadFromFile(FModeleName);
    Ftitre.text := TL[0];
    TL.Delete(0);
    Caption := TraduireMemoire('Paramétrage');
    for i := 1 to TL.Count - 1 do
    begin
      St := TL[i];
      GParams.Cells[0, i] := IntTostr(i);
      for kk := 1 to 6 + MAX_COL do
        GParams.CellValues[kk, i] := Trim(ReadTokenPipe(St, '|'));
      ih := ReadTokenI(St);
      if ih > 0 then
        GParams.RowHeights[i] := ih;

      GParams.RowCount := GParams.RowCount + 1;
    end;
    TL.Free;
  end;
  if GParams.RowCount > 2 then
    GParams.RowCount := GParams.RowCount - 1;
  GParams.FixedRows := 1;
  if Pages.ActivePage = PParams then
    GParams.AutoResizeColumns(10);
end;

procedure TFEtatLiasse.bSauveClick(Sender: TObject);
var
  TL: HTStringList;
  EtatOK: integer;
  lStCrit3, lStPredefini: string;
begin
  if ctxPCL in V_PGI.PGIContexte then
  begin
    if ControleStandardPCL(lStCrit3, lstPredefini) then
    begin
      TL := HTStringList.Create;
      try
        EtatOk := GrilleVersMaquette(TL);
        if JePeuxEnregistrerMaquette(EtatOk) then
          EnregistreMaquette(TL, FModeleName, lstCrit3, lstPredefini);
      finally
        TL.Free;
      end;
    end;
  end else
  begin
    FFICHIERPGE.Text := ExtractFileName(FModeleName);
    FCHEMINFICHIERPGE.Text := FModeleName;
    PNOMFICHIERPGE.Visible := True;
  end;
end;

function TFEtatLiasse.EncodeFont: string;
begin
  Result := FTestFont.Font.Name + ';' + IntToStr(FTestFont.Font.Size) + ';';
  if fsBold in FTestFont.Font.Style then
    Result := Result + 'B';
  if fsItalic in FTestFont.Font.Style then
    Result := Result + 'I';
  if fsUnderline in FTestFont.Font.Style then
    Result := Result + 'U';
  if fsStrikeOut in FTestFont.Font.Style then
    Result := Result + 'S';
  Result := Result + ';' + ColorToString(FTestFont.Font.Color) + ';' +
    ColorToString(FTestFont.Color) + ';';
  if akLeft in FTestFont.Anchors then
    Result := Result + 'X'
  else
    Result := Result + '-';
  if akTop in FTestFont.Anchors then
    Result := Result + 'X'
  else
    Result := Result + '-';
  if akRight in FTestFont.Anchors then
    Result := Result + 'X'
  else
    Result := Result + '-';
  if akBottom in FTestFont.Anchors then
    Result := Result + 'X'
  else
    Result := Result + '-';
  Result := Result + ';';
end;

procedure TFEtatLiasse.DecodeFont(lafonte: TFont; St: string);
var
  Frame, StS: string;
begin
  if St = '' then
    St := Default_Style;
  try
    lafonte.Name := ReadTokenSt(St);
    lafonte.Size := ReadTokenI(St);
    StS := ReadTokenSt(St);
    lafonte.Style := [];
    if Pos('B', StS) > 0 then
      lafonte.Style := lafonte.Style + [fsBold];
    if Pos('I', StS) > 0 then
      lafonte.Style := lafonte.Style + [fsItalic];
    if Pos('U', StS) > 0 then
      lafonte.Style := lafonte.Style + [fsUnderline];
    if Pos('S', StS) > 0 then
      lafonte.Style := lafonte.Style + [fsStrikeOut];

    if St <> '' then
      lafonte.Color := StringToColor(ReadTokenSt(St));
    if St <> '' then
      FTestFont.Color := StringToColor(ReadTokenSt(St));
    Frame := ReadTokenSt(St);
    if Length(Frame) < 4 then
      Frame := Frame + '    ';
    FTestFont.Anchors := [];
    FTestFont.Font.Assign(laFonte);

    if (Frame[1] = 'X') then
      FTestFont.Anchors := FTestFont.Anchors + [akLeft];
    if (Frame[2] = 'X') then
      FTestFont.Anchors := FTestFont.Anchors + [akTop];
    if (Frame[3] = 'X') then
      FTestFont.Anchors := FTestFont.Anchors + [akRight];
    if (Frame[4] = 'X') then
      FTestFont.Anchors := FTestFont.Anchors + [akBottom];

  except
  end;
end;

function TFEtatLiasse.IsLigneVide(ARow: Integer): Boolean;
var
  i: Integer;
begin
  Result := TRUE;
  for i := 4 to GParams.ColCount - 1 do
    if Trim(GParams.Cells[i, ARow]) <> '' then
    begin
      Result := FALSE;
      Exit;
    end;
end;

procedure TFEtatLiasse.GParamsRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  if (GParams.Row = GParams.RowCount - 1) and (not IsLigneVide(GParams.Row))
    then
  begin
    GParams.RowCount := GParams.RowCount + 1;
    GParams.Cells[0, GParams.RowCount - 1] := IntToStr(GParams.RowCount - 1);
  end;
  if GParams.Cells[2, GParams.Row] = 'G' then
    bGauche.Down := TRUE
  else if GParams.Cells[2, GParams.Row] = 'D' then
    bDroite.Down := TRUE
  else
    bCentre.Down := TRUE;
end;

procedure TFEtatLiasse.GParamsRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  if (GParams.Row = GParams.RowCount - 3) and (ou = GParams.RowCount - 2)
    and (IsLigneVide(Ou + 1)) then
    GParams.RowCount := GParams.RowCount - 1;
end;

procedure TFEtatLiasse.GDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState:
  TGridDrawState);
var
  Style: string;
  Rect: TRect;
begin
  if (ARow > 0) and (ACol = 4) then
  begin
    Style := Gparams.Cells[3, ARow];
    if Style = '' then
      Style := Default_Style;
    DecodeFont(FTestFont.Font, Style);
    Rect := GParams.CellRect(ACol, ARow);
    GParams.Canvas.Font := FTestFont.Font;
    GParams.Canvas.Brush.Color := FTestFont.Color;
    GParams.Canvas.Pen.Style := psSolid;
    GParams.Canvas.Pen.Color := clBlack;
    if akTop in FTestFont.Anchors then
    begin
      GParams.Canvas.MoveTo(Rect.Left, Rect.Top);
      GParams.Canvas.LineTo(Rect.Right + 1, Rect.Top);
    end;
    //      if akLeft in FTestFont.Anchors then BEGIN GParams.Canvas.MoveTo(Rect.Left,Rect.Top) ; GParams.Canvas.LineTo(Rect.Left,Rect.Bottom+1) ;END ;
    //      if akRight in FTestFont.Anchors then BEGIN GParams.Canvas.MoveTo(Rect.Right,Rect.Top) ; GParams.Canvas.LineTo(Rect.Right,Rect.Bottom+1) ; END ;
    if akBottom in FTestFont.Anchors then
    begin
      GParams.Canvas.MoveTo(Rect.Left, Rect.Bottom);
      GParams.Canvas.LineTo(Rect.Right + 1, Rect.Bottom);
    end;
  end;
end;

procedure TFEtatLiasse.GParamsCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
  if (GParams.Row = GParams.RowCount - 1) and (not IsLigneVide(GParams.Row))
    then
  begin
    GParams.RowCount := GParams.RowCount + 1;
    GParams.Cells[0, GParams.RowCount - 1] := IntToStr(GParams.RowCount - 1);
  end;
end;

procedure TFEtatLiasse.BDelLigneClick(Sender: TObject);
var
  i: Integer;
begin
  if GParams.RowCount < 3 then
    exit;
  GParams.DeleteRow(GParams.Row);
  for i := 1 to GParams.RowCount - 1 do
    GParams.Cells[0, i] := IntToStr(i);
end;

procedure TFEtatLiasse.BInsLigneClick(Sender: TObject);
var
  i: Integer;
begin
  GParams.InsertRow(GParams.Row);
  for i := 1 to GParams.RowCount - 1 do
    GParams.Cells[0, i] := IntToStr(i);
  GParams.Row := GParams.Row - 1;
  GParams.Cells[3, GParams.Row] := Default_Style;
end;

procedure TFEtatLiasse.GParamsRowMoved(Sender: TObject; FromIndex, ToIndex:
  Integer);
var
  i: Integer;
begin
  for i := 1 to GParams.RowCount - 1 do
    GParams.Cells[0, i] := IntToStr(i);
end;

procedure TFEtatLiasse.bPoliceClick(Sender: TObject);
begin
  with FD do
  begin
    DecodeFont(FTestFont.Font, GParams.Cells[3, GParams.Row]);
    Font.Assign(FTestFont.Font);
    if Execute then
    begin
      FTestFont.Font.Assign(Font);
      GParams.Cells[3, GParams.Row] := EncodeFont;
      GParams.InvalidateRow(GParams.Row);
    end;
  end;
end;

procedure TFEtatLiasse.bColorClick(Sender: TObject);
begin
  with CD do
  begin
    DecodeFont(FTestFont.Font, GParams.Cells[3, GParams.Row]);
    Color := FTestFont.Color;
    if Execute then
    begin
      FTestFont.Color := Color;
      GParams.Cells[3, GParams.Row] := EncodeFont;
      GParams.InvalidateRow(GParams.Row);
    end;
  end;
end;

procedure TFEtatLiasse.bGaucheClick(Sender: TObject);
begin
  if bGauche.Down then
    GParams.Cells[2, GParams.Row] := 'G'
  else if bDroite.Down then
    GParams.Cells[2, GParams.Row] := 'D'
  else
    GParams.Cells[2, GParams.Row] := 'C';
  GParams.InvalidateRow(GParams.Row);
end;

procedure TFEtatLiasse.bFrameChange(Sender: TObject);
begin
  DecodeFont(FTestFont.Font, GParams.Cells[3, GParams.Row]);
  FTestFont.Anchors := [];
  if bFrame.CurrentChoix in [1, 5, 7] then
    FTestFont.Anchors := FTestFont.Anchors + [aktop];
  if bFrame.CurrentChoix in [2, 6, 7] then
    FTestFont.Anchors := FTestFont.Anchors + [akLeft];
  if bFrame.CurrentChoix in [3, 5, 7] then
    FTestFont.Anchors := FTestFont.Anchors + [akBottom];
  if bFrame.CurrentChoix in [4, 6, 7] then
    FTestFont.Anchors := FTestFont.Anchors + [akRight];
  GParams.Cells[3, GParams.Row] := EncodeFont;
  GParams.InvalidateRow(GParams.Row);
end;

function TFEtatLiasse.AddGras(St: string): string;
begin
  DecodeFont(FTestFont.Font, St);
  FTestFont.Font.Style := FTestFont.Font.Style + [fsBold];
  result := EncodeFont;
end;

procedure Dechiffre(i: Integer; var St, St1, St2, St3: string);
var
  i1, i2, j: Integer;
  k1, k2: Integer;
  Ok2: Boolean;
begin
  Inc(i);
  i1 := Pos('+', St);
  i2 := Pos('-', St);
  St2 := '';
  St1 := '';
  St3 := '';
  if St = '' then
    Exit;
  Ok2 := FALSE;
  if (i2 > 0) or (i1 > 0) then
  begin
    k1 := 0;
    for j := i - 1 downto 1 do
      if St[j] in ['+', '-'] then
      begin
        k1 := j;
        break;
      end;
    k2 := Length(St);
    for j := i to k2 do
      if St[j] in ['+', '-'] then
      begin
        Ok2 := TRUE;
        k2 := j;
        break;
      end;
    if not Ok2 then
      k2 := 0;
    for j := 1 to Length(St) do
    begin
      if j <= k1 then
        St1 := St1 + St[j];
      if (J > k1) and ((j < k2) or (k2 = 0)) then
        St2 := St2 + St[j];
      if (j >= k2) and (k2 > 0) then
        St3 := St3 + St[j];
    end;
  end
  else
    St2 := St;
end;

procedure TFEtatLiasse.GParamsDblClick(Sender: TObject);
var
  St, St1, St2, St3: string;
  i: Integer;
begin
  if GParams.Col >= 7 then
  begin
    St := GParams.Cells[GParams.Col, GParams.Row];
    if (St = '') or (Copy(St, 1, 1) <> '=') then
    begin
      if GParams.inplaceEditor <> nil then
      begin
        i := GParams.inplaceEditor.SelStart;
        Dechiffre(i, St, St1, St2, St3);
      end
      else
      begin
        i := Length(St);
        Dechiffre(i, St, St1, St2, St3);
      end;
      Cache.Text := St2;
      if LookupList(Cache, TraduireMemoire('Recherche d''une rubrique'),
        'RUBRIQUE', 'RB_RUBRIQUE', 'RB_LIBELLE',
        'RB_NATRUB="CPT"', 'RB_RUBRIQUE', True, 7701) then
        //      if GChercheCompte(Cache, nil) then
      begin
        St := St1 + Cache.Text + St3;
        GParams.Cells[GParams.Col, GParams.Row] := St;
        if i > Length(St) then
          i := Length(St);

        // GCO - 23/12/2004 - FQ 15149
        if GParams.InplaceEditor <> nil then
          GParams.InplaceEditor.SelStart := i;

        if GParams.Cells[POS_LIBELLE, GParams.Row] = '' then
          // GParams.Cells[POS_LIBELLE, GParams.Row] := LibRub.Caption;
          GParams.Cells[POS_LIBELLE, GParams.Row] := RechDom('TZRUBRIQUE',
            Cache.Text, False);
      end;
    end;
  end;
end;

procedure TFEtatLiasse.GParamsKeyDown(Sender: TObject; var Key: Word; Shift:
  TShiftState);
var
  OkG, Vide: boolean;
begin
  OkG := (Screen.ActiveControl = GParams);
  Vide := (Shift = []);
  case Key of
    VK_F5: if ((OkG) and (Vide)) then
      begin
        Key := 0;
        GParamsDblClick(nil);
      end;
  end;
end;

procedure TFEtatLiasse.FListeDblClick(Sender: TObject);
var
  C, R, C1, R1: longint;
  St: string;
begin
  Exit;
  FListe.MouseToCell(GX, GY, C, R);
  if R <= 0 then
    Exit;
  R1 := R;
  C1 := POS_COL + C - 1;
  (*
  Appel:='RUBRIQUE' ;
    Result:=Get_Cumul(Appel,Nom,Crit.StTypEcr,Crit.Etab,Crit.Devise,FFormuleCol[LaCol],'') ;
    FListe.Cells[POS_COL+i-1,0]:=Crit.Col[i-1].StTitre ;
  i:=POS_COL+i-1,0
  *)
  St := GParams.Cells[6 + C1, R1];
  (*
  If (St='') Or (Copy(St,1,1)<>'=') Then
    GetBalanceSimple('RUBRIQUE',Fliste.Cells[0,Fliste.Row],Crit.StTypEcr,Crit.Etab,Crit.Devise,
                     Crit.Col[i].Exo.Code,Crit.Col[i].Date1,Crit.Col[i].Date2,False) ;
  *)
end;

procedure TFEtatLiasse.FListeMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  GX := X;
  GY := Y;
end;

procedure TFEtatLiasse.FormClose(Sender: TObject; var Action: TCloseAction);
var
  C, R: integer;
begin
  for R := Fliste.FixedRows to FListe.RowCount - 1 do
    for C := 0 to FListe.ColCount - 1 do
    begin
      if (FListe.Objects[C, R] <> nil) then
        Fliste.Objects[C, R].Free;
      FListe.Cells[C, R] := '';
      FListe.Objects[C, R] := nil;
    end;
  for R := FlisteNonAffectes.FixedRows to FlisteNonAffectes.RowCount - 1 do
    for C := 0 to FlisteNonAffectes.ColCount - 1 do
    begin
      if (FlisteNonAffectes.Objects[C, R] <> nil) then
        FlisteNonAffectes.Objects[C, R].Free;
      FlisteNonAffectes.Cells[C, R] := '';
      FlisteNonAffectes.Objects[C, R] := nil;
    end;
end;

procedure TFEtatLiasse.ChargeNonAffectes;
var
  T, TGen, TGenDetail: TOB;
  i, j, k: integer;
  stRub, stLastRub, stWhere, stCompte: string;
  TypeEtat: TTypeEtatSynthese;
  bFirst: boolean;
  Q, QRub: TQuery;
begin
  if FNonAffectes then
    exit;
  TypeEtat := esBil;
  FNonAffectes := True;
  bFirst := True;
  TGen := TOB.Create('', nil, -1);
  //TGen.LoadDetailDB('GENERAUX', '', '', nil, False, False);
  if ((Crit.TE = esBIL) or (Crit.TE = esBILA)) then stWhere := '(G_NATUREGENE<>"CHA" AND G_NATUREGENE<>"PRO")'
  else stWhere := '(G_NATUREGENE="CHA" OR G_NATUREGENE="PRO")';
  TGen.LoadDetailFromSQL('SELECT G_GENERAL,G_LIBELLE,G_TOTDEBP,G_TOTCREP,G_TOTDEBE,G_TOTCREE,G_TOTDEBS,G_TOTCRES FROM GENERAUX WHERE ' + stWhere);
  InitMove(FLaTOB.Detail.Count, '');
  for i := 0 to FLaTOB.Detail.Count - 1 do
  begin
    MoveCur(False);
    T := FLaTOB.Detail[i];
    if ((T.GetValue('TYPE') = 'R') or (T.GetValue('TYPE') = 'L')) then
    begin
      for k := 1 to MAX_COL do
      begin
        stRub := T.GetValue('FORMULE' + IntToStr(k));
        if stRub <> stLastRub then
        begin
          QRub :=
            OpenSQL('SELECT RB_RUBRIQUE,RB_COMPTE1 FROM RUBRIQUE WHERE RB_RUBRIQUE="'
            + stRub + '"', True);
          while not QRub.Eof do
          begin
            stWhere := AnalyseCompte(QRub.FindField('RB_COMPTE1').AsString,
              fbGene, False, True);
            Q := OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE ' + stWhere,
              True);
            while not Q.Eof do
            begin
              stCompte := Q.FindField('G_GENERAL').AsString;
              if bFirst then
              begin
                bFirst := False;
                if (stCompte[1] in ['1', '2', '3', '4', '5']) then
                  TypeEtat := esBIL
                else if (stCompte[1] in ['6', '7']) then
                  TypeEtat := esCR
                else
                  bFirst := True;
                if not bFirst then
                begin
                  j := 0;
                  while j < (TGen.Detail.Count - 1) do
                  begin
                    TGenDetail := TGen.Detail[j];
                    stCompte := TGenDetail.GetValue('G_GENERAL');
                    if ((not (stCompte[1] in ['1', '2', '3', '4', '5']) and
                      (TypeEtat = esBIL)) or
                      (not (stCompte[1] in ['6', '7']) and (TypeEtat = esCR)))
                      then
                      TGenDetail.Free
                    else
                      Inc(j, 1);
                  end;
                end;
              end;
              TGenDetail := TGen.FindFirst(['G_GENERAL'],
                [Q.FindField('G_GENERAL').AsString], False);
              if TGenDetail <> nil then
                TGenDetail.Free;
              Q.Next;
            end;
            Ferme(Q);
            QRub.Next;
          end;
          Ferme(QRub);
        end;
        stLastRub := stRub;
      end;
    end;
  end;
  if TGen.Detail.Count = 0 then
    FListeNonAffectes.RowCount := 2
  else
    FListeNonAffectes.RowCount := TGen.Detail.Count + 1;
  FListeNonAffectes.ColCount := 8;
  FListeNonAffectes.DefaultRowHeight := 18;
  FTitreNA[0] := TraduireMemoire('Compte');
  FTitreNA[1] := TraduireMemoire('Libellé');
  FTitreNA[2] := TraduireMemoire('Débit N-1');
  FTitreNA[3] := TraduireMemoire('Crédit N-1');
  FTitreNA[4] := TraduireMemoire('Débit N');
  FTitreNA[5] := TraduireMemoire('Crédit N');
  FTitreNA[6] := TraduireMemoire('Débit N+1');
  FTitreNA[7] := TraduireMemoire('Crédit N+1');
  for i := 0 to TGen.Detail.Count - 1 do
  begin
    FListeNonAffectes.Cells[0, i + 1] := TGen.Detail[i].GetValue('G_GENERAL');
    T := TOB.Create('Une TOB', nil, -1);
    T.AddChampSupValeur('TYPE', 'C');
    T.AddChampSupValeur('ALIGN', 'L');
    T.AddChampSupValeur('STYLE', 'ARIAL;8;;CLBLACK;CLWHITE;----;');
    T.AddChampSupValeur('LIBELLE', '');
    FListe.RowHeights[FListe.RowCount - 1] := 18;
    FListeNonAffectes.Objects[0, i + 1] := T;
    FListeNonAffectes.Cells[1, i + 1] := TGen.Detail[i].GetValue('G_LIBELLE');
    FListeNonAffectes.Cells[2, i + 1] := TGen.Detail[i].GetValue('G_TOTDEBP');
    FListeNonAffectes.Cells[3, i + 1] := TGen.Detail[i].GetValue('G_TOTCREP');
    FListeNonAffectes.Cells[4, i + 1] := TGen.Detail[i].GetValue('G_TOTDEBE');
    FListeNonAffectes.Cells[5, i + 1] := TGen.Detail[i].GetValue('G_TOTCREE');
    FListeNonAffectes.Cells[6, i + 1] := TGen.Detail[i].GetValue('G_TOTDEBS');
    FListeNonAffectes.Cells[7, i + 1] := TGen.Detail[i].GetValue('G_TOTCRES');
  end;
  FiniMove;
  TGen.Free;
end;

procedure TFEtatLiasse.UpdateRowLengths(G: THGrid; bMontant0: boolean);
var
  i, k: integer;
  bLigneAZero: boolean;
  T: TOB;
begin
  if bMontant0 then
    exit;
  for i := 1 to G.RowCount - 1 do
  begin
    bLigneAZero := True;
    T := TOB(G.Objects[0, i]);
    if (T.GetValue('TYPE') <> 'C') and (T.GetValue('TYPE') <> 'F') then
    begin
      for k := POS_COL to G.ColCount - 1 do
      begin
        if Valeur(G.Cells[k, i]) <> 0 then
        begin
          bLigneAZero := False;
          break;
        end;
      end;
      //      if bLigneAZero then G.RowHeights[i]:=0;
              { Lignes à 0 invisibles }
      if bLigneAZero then
        G.RowHeights[i] := -1;
    end;
  end;
end;

procedure TFEtatLiasse.AfficheMentionResultat;
var
  i: integer;
begin
  if VH^.bAnalytique then
    exit; // Pas d'affichage de mention si analytique
  FBEcartEtat := False;
  for i := 1 to MAX_COL do
  begin
    if FEcartResultat[i] <> 0 then
    begin
      if not FBEcartEtat then
      begin
        CreeLigneMention;
        FBEcartEtat := True;
      end;
      AjouteMentionCellule((POS_COL - 1) + i, FEcartResultat[i]);
    end;
  end;
end;

procedure TFEtatLiasse.AfficheMentionBilan;
var
  txN, txN1: integer;
begin
  if VH^.bAnalytique then
    exit; // Pas d'affichage de mention si analytique
  FBEcartEtat := False;
  if Crit.AvecPourcent then
  begin
    txN := 2;
    txN1 := 3;
  end
  else
  begin
    txN := 0;
    txN1 := 0;
  end;
  if ((FBilanActif[3 + txN] <> FBilanPassif[1]) or (FBilanActif[4 + txN1] <>
    FBilanPassif[4 + txN1])) then
  begin
    CreeLigneMention;
    FBEcartEtat := True;
    AjouteMentionCellule(POS_COL, FBilanActif[3 + txN] - FBilanPassif[1]);
    if Crit.AvecPourcent then
      AjouteMentionCellule(POS_COL + 6, FBilanActif[4 + txN1] - FBilanPassif[4 +
        txN1])
    else
      AjouteMentionCellule(POS_COL + 3, FBilanActif[4 + txN1] - FBilanPassif[4 +
        txN1]);
  end;
end;

procedure TFEtatLiasse.CreeLigneMention;
var
  T: TOB;
  stMention: string;
begin
  FListe.RowCount := FListe.RowCount + 2;
  FListe.RowHeights[FListe.RowCount - 2] := 18;
  T := TOB.Create('Une TOB', nil, -1);
  T.AddChampSupValeur('TYPE', 'C');
  T.AddChampSupValeur('ALIGN', 'D');
  T.AddChampSupValeur('STYLE', 'ARIAL;8;BU;clRed;clWhite;----;');
  T.AddChampSupValeur('LIBELLE', '');
  FListe.RowHeights[FListe.RowCount - 1] := 18;
  FListe.Objects[0, FListe.RowCount - 1] := T;
  case Crit.TE of
    esBIL, esBILA: stMention := TraduireMemoire('Ecart entre actif et passif');
    esCR, esSIG, esCRA, esSIGA: stMention :=
      TraduireMemoire('Ecart sur le résultat');
  end;
  FListe.Cells[0, FListe.RowCount - 1] := stMention;
end;

procedure TFEtatLiasse.AjouteMentionCellule(Col: integer; Valeur: double);
begin
  if Valeur <> 0 then
    FListe.Cells[Col, FListe.RowCount - 1] := FormatFloat(FFormatcol[Col],
      Valeur);
end;

function GetNomCabinet: string;
var
  QCom: TQuery;
begin
  if ctxPCL in V_PGI.PGIContexte then
  begin
    QCom :=
      OpenSQL('SELECT SOC_DATA FROM ##DP##.PARAMSOC WHERE SOC_NOM="SO_LIBELLE"',
      True);
    if not QCom.Eof then
      Result := QCom.Fields[0].AsString
    else
      Result := '###Erreur###';
    Ferme(QCom);
  end;
end;

function DecaleLibelle(St: string): string;
var
  i: integer;
begin
  if (Length(St) > 0) then
  begin
    i := 1;
    while ((i <= length(St)) and (St[i] = '.')) do // Correction SBO 06/05/2004
    begin
      St[i] := ' ';
      Inc(i, 1);
    end;
  end;
  Result := St;
end;

procedure TFEtatLiasse.BCValideClick(Sender: TObject);
begin
  ValideCtxPCL;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 07/11/2003
Modifié le ... :   /  /
Description .. : Met à jour une TOB avec les éléments à imprimer : une tob
Suite ........ : mère passée en paramètre contiendrait tout les grilles
Suite ........ : (chaque grille est vue comme une fille)
Mots clefs ... :
*****************************************************************}

procedure TFEtatLiasse.GrilleToTob(TMere: TOB);

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 12/11/2003
Modifié le ... :   /  /
Description .. : Suppression des lignes inutiles dans la TOB image de la
Suite ........ : grille
Mots clefs ... :
*****************************************************************}
  procedure NettoieTob(T: TOB);
  var
    i: integer;
    nLigneAvant: integer;
  begin
    nLigneAvant := T.Detail.Count;
    { Suppression des lignes inutiles en début }
    if (T.GetValue('FirstLine') <> 0) then
    begin
      i := 0;
      while (i < T.GetValue('FirstLine')) do
      begin
        T.Detail[0].Free;
        Inc(i, 1);
      end;
    end;
    { Suppression des lignes inutiles en fin }
    if ((T.GetValue('LastLine') <> 0) and (T.GetValue('LastLine') <>
      nLigneAvant)) then
    begin
      i := nLigneAvant - T.Detail.Count - 1;
      while (i <> 0) do
      begin
        T.Detail[T.Detail.Count - 1].Free;
        Dec(i, 1);
      end;
    end;
  end;

var
  Grille, GPassif, GActif: THGrid;
  T: TOB;
  i, j, x: integer;
  col, tcol: integer;
  nc, nl: integer;
  bDetailBilan: boolean;
  dLigneFinActif: integer;
begin
  bDetailBilan := False;
  for i := 0 to Pages.PageCount - 1 do
  begin
    if ((Pages.Pages[i].TabVisible) and (Pages.Pages[i].Name <> 'PParams')) then
    begin
      for j := 0 to Pages.Pages[i].ControlCount - 1 do
      begin
        if Pages.Pages[i].Controls[j] is THGrid then
        begin
          Grille := THGrid(Pages.Pages[i].Controls[j]);
          //          if (Grille.Tag = 1) and ((Crit.TE = esBil) or (Crit.TE = esBILA)) then
                    // CA - 21/02/2005 - Nouvelle gestion des tags
          if ((Grille.Tag = 1) or (Grille.Tag = 2)) and ((Crit.TE = esBil) or
            (Crit.TE = esBILA)) then
          begin // édition d'un bilan
            if bDetailBilan then
              dLigneFinActif := FLigneFinDetailActif
            else
              dLigneFinActif := FligneFinActif;
            { 1 - Constitution de la grille du bilan actif }
            GActif := THGrid.Create(nil);
            GActif.ColCount := Grille.ColCount;
            GActif.RowCount := dLigneFinActif + 1;
            GActif.OnDrawCell := FListeDrawCell;
            GActif.FixedRows := Grille.FixedRows;
            GActif.FixedCols := Grille.FixedCols;
            GActif.FixedColor := $00E6E6E6;
            GActif.Tag := Grille.Tag;
            GActif.Options := [goFixedVertLine, goFixedHorzLine, goVertLine,
              goColSizing, goRowSelect];
            for x := 0 to GActif.ColCount - 1 do
              GActif.ColWidths[x] := Grille.ColWidths[x];
            for x := 0 to GActif.RowCount - 1 do
            begin
              GActif.RowHeights[x] := Grille.RowHeights[x];
              GActif.Rows[x].Assign(Grille.Rows[x]);
            end;
            {            for x:=0 to GActif.RowCount - 1 do
                          GActif.RowHeights[x]:=Grille.RowHeights[x];
                        for x := 0 to GActif.ColCount - 1 do
                        begin
                          GActif.ColWidths[x]:= Grille.ColWidths[x];
                          GActif.Cols[x].Assign(Grille.Cols[x]);
                        end;}
            T := THPrintGridFacility.AglInitGridToTob(GActif);
            T.ChangeParent(TMere, -1);
            for nl := 0 to T.Detail.Count - 1 do
              for nc := 0 to GActif.colCount - 1 do
                DoWriteCell(GActif, nl, nc, t.detail[nl]);

            GActif.Free;

            { 2 - Constitution de la grille du bilan passif }

            GPassif := THGrid.Create(nil);

            if Crit.AvecPourcent then
            begin
              case (Grille.ColCount - POS_COL) of
                8: GPassif.ColCount := 4 + POS_COL;
                6: GPassif.ColCount := 2 + POS_COL;
              end;
            end
            else
              GPassif.ColCount := Grille.ColCount - 2;

            GPassif.RowCount := Grille.RowCount - dLigneFinActif;
            GPassif.OnDrawCell := FListeDrawCell;
            GPassif.FixedRows := Grille.FixedRows;
            GPassif.FixedCols := Grille.FixedCols;
            GPassif.FixedColor := $00E6E6E6;
            if bDetailBilan then
              GPassif.Tag := 5
            else
              GPassif.Tag := 4;
            GPassif.Options := [goFixedVertLine, goFixedHorzLine, goVertLine,
              goColSizing, goRowSelect];
            GPassif.RowHeights[0] := Grille.RowHeights[0];
            GPassif.Rows[0].Assign(Grille.Rows[0]);
              // Ligne des titres de colonnes

            { Mise à jour des données }

            for x := dLigneFinActif + 1 to Grille.RowCount - 1 do
            begin
              GPassif.RowHeights[x - dLigneFinActif] := Grille.RowHeights[x];
              for col := 0 to GPassif.ColCount - 1 do
              begin
                if Crit.AvecPourcent then
                begin
                  if (col > POS_COL) then
                    tcol := 4
                  else
                    tcol := 0;
                end
                else
                begin
                  if (col > POS_COL) then
                    tcol := 2
                  else
                    tcol := 0;
                end;
                GPassif.Cells[col, x - dLigneFinActif] := Grille.Cells[col +
                  tcol, x];
                GPassif.Objects[col, x - dLigneFinActif] := Grille.Objects[col +
                  tcol, x];
                { Mise à jour des largeurs de colonne - FQ 15050 - CA - 01/12/2004 }
                if (x = (dLigneFinActif + 1)) then
                  GPassif.ColWidths[col] := Grille.ColWidths[col + tcol];
              end;
              // GPassif.Rows[x-dLigneFinActif].Assign(Grille.Rows[x]);
            end;

            (*
                        { Mise à jour des hauteurs de ligne }
                        for x := dLigneFinActif+1 to Grille.RowCount - 1 do
                        begin
                          GPassif.RowHeights[x-dLigneFinActif]:=Grille.RowHeights[x];
                        end;
                        { Mise à jour des données dans la grille du passif }
                        for x := 0 to GPassif.ColCount - 1 do
                        begin
                          { Décalage des colonnes }
                          case Crit.AvecPourcent of
                            True: if (x < (GPassif.ColCount - 2)) then tx := 0 else tx := 4;
                            False: if x <= 1 then tx := 0 else tx := 2;
                            else tx := 0;
                          end;
                          { Mise à jour de la taille des colonnes }
                          GPassif.ColWidths[x] := Grille.ColWidths[x + tx];
                          { Copie des données }
                          for y:=1 to GPassif.RowCount do
                            GPassif.Objects[x,y] := Grille.Objects[y+dLigneFinActif+1,x+tx];
                            GPassif.Cells[x,y] := Grille.Cells[y+dLigneFinActif+1,x+tx];
                        end;
            *)

            T := THPrintGridFacility.AglInitGridToTob(GPassif);
            T.ChangeParent(TMere, -1);
            for nl := 0 to T.Detail.Count - 1 do
              for nc := 0 to GPassif.colCount - 1 do
                DoWriteCell(GPassif, nl, nc, t.detail[nl]);
            GPassif.Free;
            bDetailBilan := True;
              // Le bilan est édité - Si un autre bilan , c'est le détail.
          end
          else
          begin // édition d'un compte de résultat ou d'un SIG
            Grille := THGrid(Pages.Pages[i].Controls[j]);
            T := THPrintGridFacility.AglInitGridToTob(Grille);
            T.ChangeParent(TMere, -1);
            for nl := 0 to T.Detail.Count - 1 do
              for nc := 0 to Grille.colCount - 1 do
                DoWriteCell(Grille, nl, nc, t.detail[nl]);
            //            T.AddChampSupValeur('FirstLine', 0);
            //            T.AddChampSupValeur('LastLine', Grille.RowCount);
            //            NettoieTob ( T );
          end;
        end;
      end;
    end;
  end;
end;

procedure TFEtatLiasse.ValideCtxPCL;
var
  stTitre, stMonnaie, stCriteres: string;
  Exo: TExoDate;
  Err: integer;
  DD1, DD2: TDateTime;
  GPassif: THGrid;
  LesCriteres: string;
  TMere: TOB;
  SaveValeurQrPdf: Boolean;
  stTl1: string;
  stBl1, stBl2, stBr1, stBr2: string;
  stFichierPS: string;
begin
  { Redimensionnement des colonnes des grilles à imprimer }
  AjusteTailleColonne(FListe, True);
  if PDetails.TabVisible then
    AjusteTailleColonne(FDetail, True);

  GPassif := nil;
  if WhatDate(FFormuleCol[1], DD1, DD2, Err, Exo) then
    stTitre := TraduireMemoire('Etats de synthèse au ') + DateToStr(DD2);

  // Création de la TOB mére des impressions. Elle sera supprimée en fin de procédure.
  TMere := TOB.Create('Maman', nil, -1);
  try
    GrilleToTob(TMere);

    if Crit.EnMonnaieOpposee = 'X' then
    begin
      if VH^.TenueEuro then
        stMonnaie := TraduireMemoire('En francs')
      else
        stMonnaie := TraduireMemoire('En euros');
    end;

    if VH^.bAnalytique then
    begin
      { CA - 26/08/2003 - Libellé = Axe n°x - SECTION si un état par section }
      if Crit.bUnEtatParSection then
        stCriteres := lbl_Axe.Caption + ' - ' + lbl_SectionDe.Caption
      else
      begin
        stCriteres := lbl_Axe.Caption;
        if lbl_SectionDe.Caption <> '' then
          stCriteres := stCriteres + ' ' + TraduireMemoire('de') + ' ' +
            lbl_SectionDe.Caption;
        if lbl_SectionA.Caption <> '' then
          stCriteres := stCriteres + ' ' + TraduireMemoire('à') + ' ' +
            lbl_SectionA.Caption;
      end;
    end
    else
      stCriteres := '';
    if Crit.bTLI then
      stCriteres := #22 + stCriteres;
    if Crit.bJoker then
      stCriteres := #23 + stCriteres;

    { On renseigne les informations de bas de page }
    if Crit.bImpressionDate then
    begin
      stBl1 := 'Imprimé le ' + DateToStr(date) + ' à ' + FormatDateTime('hh:mm',
        Time);
      stBl2 := 'Par ' + V_PGI.UserName;
    end;
    stBr1 := stCriteres;
    if ctxPCL in V_PGI.PGIContexte then
    begin
      if Crit.InfoLibre <> '' then
        stBr2 := Crit.InfoLibre + ' / ' + TraduireMemoire('Dossier n°') +
          V_PGI.NoDossier
      else
        stBr2 := V_PGI.DefaultSectionName + ' / ' + TraduireMemoire('Dossier n°')
          + V_PGI.NoDossier;
    end
    else
      stBr2 := '';
    stTl1 := GetParamSocSecur('SO_LIBELLE', '');
    if stMonnaie <> '' then
      stTl1 := stTl1 + ' - ' + stMonnaie;
    // Les critères
    LesCriteres := 'TITRE=' + Caption + '`';
    LesCriteres := LesCriteres + 'TOPLEFT=' + stTl1 + '`';
    LesCriteres := LesCriteres + 'TOPRIGHT=' + StTitre + '`';
    LesCriteres := LesCriteres + 'BOTTOMLEFT1=' + stBl1 + '`';
    LesCriteres := LesCriteres + 'BOTTOMRIGHT1=' + stBr1 + '`';
    LesCriteres := LesCriteres + 'BOTTOMLEFT2=' + stBl2 + '`';
    LesCriteres := LesCriteres + 'BOTTOMRIGHT2=' + stBr2 + '`';

    // Alimentation du pseudo champ critère qui sera passé telque à la fonction lanceetat()
    TMere.AddChampSupValeur('CRITERES', LesCriteres);

    // Gestion du mode silence

    // Si mode silence
    SaveValeurQrPdf := v_pgi.QRPdf;

    if FCritEdtChaine.Utiliser then
    begin
      if FCritEdtChaine.NomPDF <> '' then
      begin
        V_PGI.QRPdf := True;
        V_PGI.NoPrintDialog := True;
        if FCritEdtChaine.MultiPdf then
          V_PGI.QRPDFQueue := ExtractFilePath(FCritEdtChaine.NomPDF) + '\' +
            V_PGI.NoDossier + '-' + ExtractFileName(FCritEdtChaine.NomPDF)
        else
          V_PGI.QRPDFQueue := FCritEdtChaine.NomPDF;
        V_PGI.QRPDFMerge := '';
        if (V_PGI.QRPDFQueue <> '') and FileExists(V_PGI.QRPDFQueue) then
          V_PGI.QRPDFMerge := V_PGI.QRPDFQueue;
      end;
      // Nombre de copies
      V_PGI.NbDocCopies := FCritEdtCHaine.NombreExemplaire;
    end;

    fbStopEdition := not AglPrintGridPcl(TMere, '', '', '', FCritEdtChaine.NomPDF
      <> '', False, True, True, False, FCritEdtChaine.Utiliser);
    if (FileExists(ChangeStdDatPath('$DOS\' +
      'PGIExpert_CCS5_Etats de synthèse.PS')) or
      FileExists(ChangeStdDatPath('$DOS\' +
      'PGIExpert_CCS5_Etat de synthèse.PS'))) then
    begin
      if Crit.TE = esBIL then
        stFichierPS := TraduireMemoire('Bilan')
      else if Crit.TE = esBILA then
        stFichierPS := TraduireMemoire('Bilan(analytique)')
      else if Crit.TE = esCR then
        stFichierPS := TraduireMemoire('Compte de résultat')
      else if Crit.TE = esCRA then
        stFichierPS := TraduireMemoire('Compte de résultat(analytique)')
      else if Crit.TE = esSIG then
        stFichierPS := TraduireMemoire('SIG')
      else if Crit.TE = esSIGA then
        stFichierPS := TraduireMemoire('SIG(analytique)');
      stFichierPS := stFichierPS + '.PS';
      if FileExists(ChangeStdDatPath('$DOS\' + stFichierPS)) then
        DeleteFile(ChangeStdDatPath('$DOS\' + stFichierPS));
      if not RenameFile(ChangeStdDatPath('$DOS\' +
        'PGIExpert_CCS5_Etat de synthèse.PS'), ChangeStdDatPath('$DOS\' +
        stFichierPS)) then
        RenameFile(ChangeStdDatPath('$DOS\' +
          'PGIExpert_CCS5_Etats de synthèse.PS'), ChangeStdDatPath('$DOS\' +
          stFichierPS));
    end;

    // Sympa, on remet les choses dans l'ordre
    V_PGI.QRPDFMerge := '';
    V_PGI.QRPDFQueue := '';
    v_pgi.QRPdf := SaveValeurQrPdf;

    if ((GPassif <> nil) and (Crit.TE = esBil)) then
      GPassif.Free;

  finally
    { Redimensionnement des colonnes des grilles à imprimer }
    AjusteTailleColonne(FListe, False);
    if PDetails.TabVisible then
      AjusteTailleColonne(FDetail, False);
    FreeAndNil(TMere);
  end;
end;

procedure TFEtatLiasse.UpdateTitreBilan;
begin
  // Toujours plus crado , mais bon ... Pour l'instant, on suppose que 'il y aura toujours le même nombre de colonnes !!!!
  if Crit.AvecPourcent then
  begin
    Crit.Col[0].StTitre := TraduireMemoire('Brut');
    Crit.Col[1].StTitre := '%';
    Crit.Col[2].StTitre := TraduireMemoire('Amortissements') + #13 + #10 +
      TraduireMemoire('Provisions');
    Crit.Col[3].StTitre := '%';
    Crit.Col[4].StTitre := TraduireMemoire('Net au ') + #13 + #10 +
      FormatDateTime('dd/mm/yy', Crit.Col[0].Date2);
    Crit.Col[5].StTitre := '%';
    Crit.Col[6].StTitre := TraduireMemoire('Net au ') + #13 + #10 +
      FormatDateTime('dd/mm/yy', Crit.Col[3].Date2);
    Crit.Col[7].StTitre := '%';
    FTitrePassif[POS_COL] := TraduireMemoire('Net au ') + #13 + #10 +
      FormatDateTime('dd/mm/yy', Crit.Col[0].Date2);
    FTitrePassif[POS_COL + 1] := '%';
    FTitrePassif[POS_COL + 2] := TraduireMemoire('Net au ') + #13 + #10 +
      FormatDateTime('dd/mm/yy', Crit.Col[3].Date2);
    FTitrePassif[POS_COL + 3] := '%';
  end
  else
  begin
    Crit.Col[0].StTitre := TraduireMemoire('Brut');
    Crit.Col[1].StTitre := TraduireMemoire('Amortissements') + #13 + #10 +
      TraduireMemoire('Provisions');
    Crit.Col[2].StTitre := TraduireMemoire('Net au ') + #13 + #10 +
      FormatDateTime('dd/mm/yy', Crit.Col[0].Date2);
    Crit.Col[3].StTitre := TraduireMemoire('Net au ') + #13 + #10 +
      FormatDateTime('dd/mm/yy', Crit.Col[3].Date2);
    FTitrePassif[POS_COL] := TraduireMemoire('Net au ') + #13 + #10 +
      FormatDateTime('dd/mm/yy', Crit.Col[0].Date2);
    FTitrePassif[POS_COL + 1] := TraduireMemoire('Net au ') + #13 + #10 +
      FormatDateTime('dd/mm/yy', Crit.Col[3].Date2);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 07/12/2004
Modifié le ... : 07/12/2004
Description .. : Mise à jour des tailles des colonnes
Suite ........ : - si bEtat = False : dans la grille affichée suite aux calculs
Suite ........ : - si bEtat = True : dans l'édition
Suite ........ : On utilise :
Suite ........ :  - LCOLA pour les dimensions de l'aperçu
Suite ........ :  - LCOLE pour les dimensions de l'état
Mots clefs ... :
*****************************************************************}

procedure TFEtatLiasse.AjusteTailleColonne(G: THGrid; bEtat: boolean);
const
  { Largeur de colonne pour les montants si moins de 6 colonnes }
  LCOLA_MONTANTMAX = 190;
  LCOLE_MONTANTMAX = 165;
  { Largeur de colonne pour les montants si plus de 6 colonnes }
  LCOLA_MONTANTMIN = 130;
  LCOLE_MONTANTMIN = 105;
  { Largeur de colonne pour les pourcentages }
  LCOLA_POURCENT = 75;
  LCOLE_POURCENT = 60;
var
  i, CW: integer;

begin
  if (not Crit.AvecPourcent) then
  begin
    for i := POS_COL + 1 to G.ColCount - 1 do
      G.ColWidths[i] := 80;
  end
  else
  begin
    case G.ColCount of
      1..6: if bEtat then
          CW := LCOLE_MONTANTMAX
        else
          CW := LCOLA_MONTANTMAX;
      7..9: if bEtat then
          CW := LCOLE_MONTANTMIN
        else
          CW := LCOLA_MONTANTMIN;
    else if bEtat then
      CW := LCOLE_POURCENT
    else
      CW := LCOLA_POURCENT;
    end;
    for i := POS_COL to G.ColCount - 1 do
    begin
      if (((i - POS_COL) mod 2) = 0) then
        G.ColWidths[i] := CW
      else
      begin
        if bEtat then
          G.ColWidths[i] := LCOLE_POURCENT
        else
          G.ColWidths[i] := LCOLA_POURCENT;
      end;
    end;
  end;
end;

function TFEtatLiasse.GetResultatSurPeriode(stDate: string;
  var stErr: string; EnMonnaieOpposee, Etab, BalSit: string): double;
var
  Q: TQuery;
  stSelect, stChamps, stWhere, stWhereEtab, stWhereDate, stWhereQualifPiece:
  string;
  DD1, DD2: TDateTime;
  Err: integer;
  Exo: TExoDate;
  Resultat: double;
begin
  stErr := '';
  Resultat := 0;
  if WhatDate(stDate, DD1, DD2, Err, Exo) then
  begin
    if BalSit <> '' then
    begin
      stChamps := 'SUM(BSE_CREDIT)-SUM(BSE_DEBIT)';
      stSelect := 'SELECT ' + stChamps +
        ' FROM CBALSITECR LEFT JOIN GENERAUX ON (BSE_COMPTE1=G_GENERAL)';
      stWhere := ' WHERE BSE_CODEBAL="' + BalSit +
        '" AND (G_NATUREGENE="CHA" OR G_NATUREGENE="PRO") ';
      Q := OpenSQL(stSelect + stWhere, True);
      if not Q.Eof then
      begin
        if EnMonnaieOpposee = 'X' then
        begin
          if not VH^.TenueEuro then
            Resultat := PivotToEuro(Q.Fields[0].AsFloat)
          else
            Resultat := EuroToPivot(Q.Fields[0].AsFloat);
        end
        else
          Resultat := Q.Fields[0].AsFloat;
      end;
      Ferme(Q);
    end
    else
    begin
      stChamps := 'SUM(E_CREDIT)-SUM(E_DEBIT)';
      stSelect := 'SELECT ' + stChamps +
        ' FROM ECRITURE LEFT JOIN GENERAUX ON (E_GENERAL=G_GENERAL)';
      stWhere := ' WHERE (G_NATUREGENE="CHA" OR G_NATUREGENE="PRO") ';
      // CA - 29/11/2001 - Prise en compte des 'autres' écritures.
      stWhereQualifPiece := ' AND ( E_QUALIFPIECE="N"';
      if Crit.TypEcr[Simu] then
        stWhereQualifPiece := stWhereQualifPiece + ' OR E_QUALIFPIECE="S"';
      if Crit.TypEcr[Situ] then
        stWhereQualifPiece := stWhereQualifPiece + ' OR E_QUALIFPIECE="U"';
      if Crit.TypEcr[Revi] then
        stWhereQualifPiece := stWhereQualifPiece + ' OR E_QUALIFPIECE="R"';
      if Crit.TypEcr[Previ] then
        stWhereQualifPiece := stWhereQualifPiece + ' OR E_QUALIFPIECE="P"';
      if Crit.TypEcr[Ifrs] then
        stWhereQualifPiece := stWhereQualifPiece + ' OR E_QUALIFPIECE="I"';
          // Modif IFRS 05/05/2004
      stWhereQualifPiece := stWhereQualifPiece + ')';
      stWhere := stWhere + stWhereQualifPiece;
      if Etab <> '' then
        stWhereEtab := ' AND E_ETABLISSEMENT="' + Etab + '"'
      else
        stWhereEtab := '';
      stWhereDate := ' AND E_EXERCICE="' + Exo.Code + '" AND E_DATECOMPTABLE>="'
        + USDateTime(DD1) +
        '" AND E_DATECOMPTABLE<="' + USDateTime(DD2) + '"';
      Q := OpenSQL(stSelect + stWhere + stWhereEtab + stWhereDate, True);
      if not Q.Eof then
        Resultat := Q.Fields[0].AsFloat;
      Ferme(Q);
    end;
  end
  else
    stErr := 'Erreur Date';
  Result := Resultat;
end;

procedure TFEtatLiasse.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled := False;
  BFerme.click;
  Close;
end;

(*
procedure TFEtatLiasse.ChargeListeDesComptes;
var
  Q: TQuery;
  StQuery: string;
begin{>>GpProfile} ProfilerEnterProc(83); try {GpProfile>>}
  stQuery := '';
  case Crit.TE of
    esBil, esBILA: stQuery :=
      'SELECT G_GENERAL,G_LIBELLE FROM GENERAUX WHERE G_NATUREGENE<>"CHA" AND G_NATUREGENE<>"PRO"';
    esSIG, esSIGA, esCR, esCRA: stQuery :=
      'SELECT G_GENERAL,G_LIBELLE FROM GENERAUX WHERE G_NATUREGENE="CHA" OR G_NATUREGENE="PRO"';
  end;
  if stQuery <> '' then
  begin
    Q := OpenSQL(stQuery, True);
    if not q.eof then
      FListeDesComptes.LoadDetailDB('GENERAUX', '', '', Q, False);
    Ferme(q); // Et alors, on oublie de fermer la requête !!
  end;
{>>GpProfile} finally ProfilerExitProc(83); end; {GpProfile>>}end;

*)

function TFEtatLiasse.LibelleDuCompte(pStDetail: HString): HString;
var lSt: string;
begin
  lSt := ReadTokenTab(pStDetail);
  Result := pStDetail;
end;

procedure TFEtatLiasse.ChargeListeDesSections;
begin
  FListeDesSections.LoadDetailDB('SECTION', '', '', nil, False);
end;

function TFEtatLiasse.LibelleDeLaSection(stSection: string): string;
var
  T: TOB;
begin
  T := FListeDesSections.FindFirst(['S_SECTION'], [stSection], False);
  if T <> nil then
    Result := T.GetValue('S_LIBELLE')
  else
    Result := '';
end;

procedure TFEtatLiasse.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Key = vk_Imprime) and (ssCtrl in Shift) then
  begin
    key := 0;
    BImprimerClick(nil);
  end;
end;

procedure TFEtatLiasse.BExportClick(Sender: TObject);
var
  T, TMere: TOB;
  i, j: integer;
begin
  if not ExJaiLeDroitConcept(ccExportListe, True) then
    Exit;

  if SD.Execute then
  begin
    TMere := TOB.Create('Maman', nil, -1);
    try
      if Pages.ActivePage = PParams then
      begin

        { Export de la grille de paramétrage }
        T := THPrintGridFacility.AglInitGridToTob(GParams);
        T.ChangeParent(TMere, -1);
        for i := 0 to T.Detail.Count - 1 do
          for j := 0 to GParams.colCount - 1 do
            DoWriteCell(GParams, i, j, t.detail[i]);
      end else
        { Export des grilles de résultats (états) }
        GrilleToTob(TMere);

      { Export vers Excel }
      ExportVersExcel(TMere, SD.FileName, (Pages.ActivePage = PParams));
    finally
      TMere.Free;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 03/04/2007
Modifié le ... :   /  /
Description .. : Fonction d'exportation des grilles vers Excel (fichier .xls)
Mots clefs ... :
*****************************************************************}

procedure TFEtatLiasse.ExportVersExcel(T: TOB; FileName: string; bModele: boolean);

  function _LigneVide (pT : TOB ; bMontant0 : boolean ) : boolean;
  var ii : integer;
      bTitre : boolean;
      bZero : boolean;
  begin
    // Est-ce qu'on a un titre sur la ligne ?
    bTitre := True;
    for ii := 0 to pT.Detail.Count - 1 do
    begin
      if (pT.Detail[ii].GetValue('COL') <= POS_LIBELLE) then
        bTitre := bTitre and (Trim(pT.Detail[ii].GetValue('TEXT'))<>'')
      else break;
    end;
    // Est-ce que la ligne est composée uniquement de montant à 0
    bZero := True;
    for ii := 0 to pT.Detail.Count - 1 do
    begin
      if (pT.Detail[ii].GetValue('COL') > POS_LIBELLE) then
        bZero := bZero and (Trim(pT.Detail[ii].GetValue('TEXT'))='');
    end;
    // La ligne est vide si on n'a pas de titre ou si toutes les colonnes de montants
    // sont à zéro et qu'on ne veut pas éditer les montants à 0
    Result := ((not bTitre) and bZero) or ((not bMontant0) and bZero);
  end;

var
  i, j, k, iRowCount: integer;
  nRowCount: integer;
  XLS: TXLSWriter;
  TGrille, TLigne, TCellule: TOB;
  St: string;
begin
  XLS := TXLSWriter.Create(FileName);
  try
    nRowCount := 0;
    for i := 0 to T.Detail.Count - 1 do
      nRowCount := nRowCount + T.Detail[i].GetValue('ROWCOUNT');
    XLS.SetDimensions(T.Detail[0].GetValue('COLCOUNT'), nRowCount);
    iRowCount := 0;
    // Pour toutes les grilles
    for i := 0 to T.Detail.Count - 1 do
    begin
      TGrille := T.Detail[i];
      // Pour chaque ligne de la grille
      for j := 0 to TGrille.Detail.Count - 1 do
      begin
        TLigne := TGrille.Detail[j];
        if _LigneVide(TLigne, Crit.AvecMontant0) then continue;
        Inc(iRowCount);
        // Pour chaque cellule de la ligne
        for k := 0 to TLigne.Detail.Count - 1 do
        begin
          TCellule := TLigne.Detail[k];
          St := TCellule.GetValue('TEXT');
          { Si c'est un montant à 0, on traite comme montant dans Excel }
          if (TCellule.GetValue('COL') > POS_LIBELLE) and (not bModele) and (St = '') then
            XLS.WriteNumber(TCellule.GetValue('COL') + 1, iRowCount, 0, StrfMask(V_PGI.OkDecV, '', True))
          { Si c'est une zone de type texte }
          else if (TCellule.GetValue('COL') <= POS_LIBELLE) or (j = 0) or (St = '') or (bModele)
            then
            XLS.WriteText(TCellule.GetValue('COL') + 1, iRowCount, St)
          { Si c'est un montant }
          else
            XLS.WriteNumber(TCellule.GetValue('COL') + 1, iRowCount, Valeur(St),
              StrfMask(V_PGI.OkDecV, '', True))
        end;
      end;
    end;
  finally
    XLS.CloseFile;
    FreeAndNil(XLS);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 24/08/2005
Modifié le ... :   /  /
Description .. : Procédure de calcul des grilles à imprimer
Mots clefs ... :
*****************************************************************}

procedure TFEtatLiasse.CalculLesGrilles;
var
  TSection: TOB;
  QSection: TQuery;
  i: integer;
  stWhere: string;
begin
  if (Crit.stSectionDe <> Crit.stSectionA) and (Crit.bUnEtatParSection) then
  begin
    TSection := TOB.Create('', nil, -1);
    try
      if Crit.bJoker then
        stWhere := ' S_SECTION LIKE "' + TraduitJoker(Crit.stSectionDe) + '" '
      else
        stWhere := ' S_SECTION>="' + Crit.stSectionDe + '" AND S_SECTION<="' +
          Crit.stSectionA + '" ';
      QSection := OpenSQL('SELECT S_SECTION FROM SECTION WHERE S_AXE="' +
        Crit.stAxe + '" AND ' +
        stWhere + ' ORDER BY S_SECTION', True);
      TSection.LoadDetailDb('', '', '', QSection, False);
      Ferme(QSection);

      InitMoveProgressForm(nil, 'Etat de synthèse',
        'Préparation de l''édition en cours, veuillez patienter ...', (TSection.Detail.Count * 2) + 1, True,
        True);

      for i := 0 to TSection.Detail.Count - 1 do
      begin
        Crit.stSectionDe := TSection.Detail[i].GetValue('S_SECTION');
        Crit.stSectionA := TSection.Detail[i].GetValue('S_SECTION');

        if not MoveCurProgressForm('Préparation de l''édition (phase 1) : section ' + Crit.stSectionA) then break;

        Charge;

        if fLesRubriques.LastError <> 0 then break;

        FDetailCalcule := False;
        MajCaptionAnalytique;

        { Passage des grilles en TOB }
        if not MoveCurProgressForm('Préparation de l''édition (phase 2) : section ' + Crit.stSectionA) then break;
        GrilleToTob(fTLesGrilles);

        { Mise à jour des Titres }
        MajTitreGrilles;
      end;
    finally
      FiniMoveProgressForm;
      TSection.Free;
    end;
  end else
  begin
    InitMoveProgressForm(nil, 'Etat de synthèse',
      'Préparation de l''édition en cours, veuillez patienter ...', 3, False,
      True);
    MoveCurProgressForm('Préparation de l''édition (phase 1)...');
    Charge;

    if fLesRubriques.LastError <> 0 then
    begin
      FiniMoveProgressForm;
      exit;
    end;

    { Passage des grilles en TOB }
    MoveCurProgressForm('Préparation de l''édition (phase 2)...');
    GrilleToTob(fTLesGrilles);

    { Mise à jour des Titres }
    MajTitreGrilles;
    FiniMoveProgressForm;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 14/06/2005
Modifié le ... :   /  /
Description .. : Procédure d'impression des grilles renseignées
Mots clefs ... :
*****************************************************************}

function TFEtatLiasse.ImprimeLesGrilles: string;
var
  {TLesGrilles, }TUneGrille: TOB;
  stFileName: string;
  bStop: boolean;
  i: integer;
  OffsetPage: integer;
  bChainePDF: boolean;
begin
  OffsetPage := 0;
  bChainePDF := True;

  { Redimensionnement des colonnes des grilles à imprimer }
  AjusteTailleColonne(FListe, True);
  if PDetails.TabVisible then
    AjusteTailleColonne(FDetail, True);

  TUneGrille := TOB.Create('', nil, -1);

  try
    { Cas des états chainés }
    if ((FCritEdtChaine.Utiliser) and (FCritEdtChaine.NomPDF <> '')) then
    begin
      if FCritEdtChaine.MultiPdf then
        V_PGI.QRPDFQueue := ExtractFilePath(FCritEdtChaine.NomPDF) + '\' +
          V_PGI.NoDossier + '-' + ExtractFileName(FCritEdtChaine.NomPDF)
      else V_PGI.QRPDFQueue := FCritEdtChaine.NomPDF;
      V_PGI.QRPDFMerge := '';
      if (V_PGI.QRPDFQueue <> '') and FileExists(V_PGI.QRPDFQueue) then
        V_PGI.QRPDFMerge := V_PGI.QRPDFQueue;
      stFileName := V_PGI.QRPDFQueue;
      bChainePDF := true;
    end else
    begin
      stFileName := TempFileName();
      V_PGI.QrPdfQueue := stFileName; // Nom du fichier imprimé
(*
      {$IFDEF EAGLCLIENT}
      V_PGI.QRPDFMerge := TempFileName();
      {$ELSE}*)
      V_PGI.QRPDFMerge := stFileName;
(*      {$ENDIF}*)
      if (FCritEdtChaine.Utiliser) then
        V_PGI.NbDocCopies := FCritEdtCHaine.NombreExemplaire;
    end;

    try
{$IFDEF EAGLCLIENT}
      THPrintBatch.StartPdfBatch(V_PGI.QRPDFMerge);
{$ELSE}
      if not bChainePDF then
      begin
        THPrintBatch.StartPdfBatch(stFileName);
      end;
{$ENDIF}
      { Page de garde }
      // LanceEtat('ESY','ESY','',True,False,False,nil,'','',False);
      // THPrintBatch.AjoutPdf (V_PGI.QrPdfQueue, True) ;

      i := 0;
      bStop := False;
      InitMoveProgressForm(nil, 'Etat de synthèse',
        'Constitution de l''aperçu en cours, veuillez patienter ...', fTLesGrilles.Detail.Count + 1, True,
        True);
      try

        while ((not bStop) and (i < fTLesGrilles.Detail.Count)) do
        begin
          { FQ 17576 - 30/05/2006 - CA - Pour afficher le premier état en états chaînés PDF }
{$IFDEF EAGLCLIENT}
{$ELSE}
          if bChainePDF then
            if (V_PGI.QRPDFQueue <> '') and FileExists(V_PGI.QRPDFQueue) then
              V_PGI.QRPDFMerge := V_PGI.QRPDFQueue;
{$ENDIF}
          if not MoveCurProgressForm(fTLesGrilles.Detail[i].GetValue('TITRE_ETAT') + ' ...') then break;
          { Si l'état n'est pas demandé, on ne l'imprime pas }
          if (not Crit.bEtat) and ((fTLesGrilles.Detail[i].GetValue('TAG') = 1)
            or (fTLesGrilles.Detail[i].GetValue('TAG') = 4)) then
          begin
            Inc(i);
            continue;
          end;
          { On déplace la grille à imprimer dans la grille courante }
          fTLesGrilles.Detail[i].ChangeParent(TUneGrille, -1);

          { Mise à jour des titres de la grille }
          TUneGrille.AddChampSupValeur('TITRE_ETAT',
            TUneGrille.Detail[0].GetValue('TITRE_ETAT'));
          TUneGrille.AddChampSupValeur('CRITERES',
            TUneGrille.Detail[0].GetValue('CRITERES'));
          TUneGrille.AddChampSupValeur('OFFSETPAGE', OffsetPage);
          TUneGrille.AddChampSupValeur('DUPLICATA',
            TUneGrille.Detail[0].GetValue('DUPLICATA'));

          { Impression de la grille }
          bStop := ImprimeUneGrille(TUneGrille);
          OffsetPage := OffsetPage + TUneGrille.GetValue('OFFSETPAGE');

          { On réintègre la grille à la TOB initiale }
          TUneGrille.Detail[0].ChangeParent(fTLesGrilles, i);

          { Passage à la grille suivante }
          Inc(i);
        end;
      finally
{$IFDEF EAGLCLIENT}
        THPrintBatch.StopPdfBatch();
{$ELSE}
        if not bChainePDF then
        begin
          THPrintBatch.StopPdfBatch();
        end;
{$ENDIF}

        FiniMoveProgressForm;
      end;
    finally
    end;

  finally
    FreeAndNil(TUneGrille);
  end;

{$IFDEF EAGLCLIENT}
  stFileName := V_PGI.QRPDFMerge;
{$ENDIF}
  Result := stFileName;
end;

function TFEtatLiasse.ImprimeUneGrille(TUneGrille: TOB): boolean;
var
  bStop: boolean;
begin
  { Lancement de l'impression }
  bStop := not AglPrintGridPcl(TUneGrille, '', '', '', True, False, True, True,
    True, True);
  THPrintBatch.AjoutPdf(V_PGI.QrPdfQueue, True);
  result := bStop;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 14/06/2005
Modifié le ... :   /  /
Description .. : Construction de la chaîne de critères à passer à la routine
Suite ........ : d'édition pour les faire apparaître sur l'état
Mots clefs ... :
*****************************************************************}

function TFEtatLiasse.ConstruireParametreCritere: string;
var
  stLesCriteres: string;
  stBl1, stBl2, stBr1, stBr2: string;
  DD1, DD2: TDateTime;
  Exo: TExoDate;
  Err: integer;
begin
  { Informations de bas de page }
  if Crit.bImpressionDate then
  begin
    stBl1 := 'Imprimé le ' + DateToStr(date) + ' à ' + FormatDateTime('hh:mm',
      Time);
    stBl2 := 'Par ' + V_PGI.UserName;
  end;
  if ctxPCL in V_PGI.PGIContexte then
  begin
    if Crit.InfoLibre <> '' then
      stBr2 := Crit.InfoLibre + ' / ' + TraduireMemoire('Dossier n°') +
        V_PGI.NoDossier
    else
      stBr2 := V_PGI.DefaultSectionName + ' / ' + TraduireMemoire('Dossier n°')
        + V_PGI.NoDossier;
  end
  else
    stBr2 := GetParamSocSecur('SO_LIBELLE', '');
  if VH^.bAnalytique then
  begin
    { CA - 26/08/2003 - Libellé = Axe n°x - SECTION si un état par section }
    if Crit.bUnEtatParSection then
      stBr1 := lbl_Axe.Caption + ' - ' + lbl_SectionDe.Caption
    else
    begin
      stBr1 := lbl_Axe.Caption;
      if lbl_SectionDe.Caption <> '' then
        stBr1 := stBr1 + ' ' + TraduireMemoire('de') + ' ' +
          lbl_SectionDe.Caption;
      if lbl_SectionA.Caption <> '' then
        stBr1 := stBr1 + ' ' + TraduireMemoire('à') + ' ' +
          lbl_SectionA.Caption;
    end;
  end
  else
    stBr1 := '';
  if Crit.bTLI then
    stBr1 := #22 + stBr1;
  if Crit.bJoker then
    stBr1 := #23 + stBr1;

  { Alimentation du pseudo champ critère qui sera passé telque à la fonction lanceetat() }
  stLesCriteres := stLesCriteres + 'TOPLEFT=' + GetParamSocSecur('SO_LIBELLE', '') + '`';
  if WhatDate(FFormuleCol[1], DD1, DD2, Err, Exo) then
    stLesCriteres := stLesCriteres + 'TOPRIGHT=' +
      TraduireMemoire('Etats de synthèse au ') + DateToStr(DD2) + '`'
  else
    stLesCriteres := stLesCriteres + 'TOPRIGHT=`';
  stLesCriteres := stLesCriteres + 'BOTTOMLEFT1=' + stBl1 + '`';
  stLesCriteres := stLesCriteres + 'BOTTOMRIGHT1=' + stBr1 + '`';
  stLesCriteres := stLesCriteres + 'BOTTOMLEFT2=' + stBl2 + '`';
  stLesCriteres := stLesCriteres + 'BOTTOMRIGHT2=' + stBr2 + '`';
//  if (Crit.MentionFiligrane <> '') then
//    stLesCriteres := stLesCriteres + 'DUPLICATA=' + Crit.MentionFiligrane + '`';

  Result := stLesCriteres;
end;

function TFEtatLiasse.FaitLibelleEtatSynthese(TE: TTypeEtatSynthese; bEtat,
  bDetail, bPassif, bPublifi: boolean): string;
begin
  Result := '';
  case Crit.TE of
    esSIG, esSIGA, esCR, esCRA:
      begin
        if bEtat then
          Result := fTitreMaquette[1]
        else if bDetail then
          Result := fTitreMaquette[2];
      end;
    esBIL, esBILA:
      begin
        if bPublifi then
        begin
          if bDetail and not bEtat then
            Result := TraduireMemoire('Détail du bilan')
          else
            Result := TraduireMemoire('Bilan');
        end
        else
        begin
          if (bEtat and bPassif) then
            Result := fTitreMaquette[3]
          else if (bDetail and bPassif) then
            Result := fTitreMaquette[4]
          else if (bDetail and not bPassif) then
            Result := fTitreMaquette[2]
          else if (not bDetail and not bPassif) then
            Result := fTitreMaquette[1];
        end;
      end;
  end;
  Result := TraduireMemoire(Result);
end;

procedure TFEtatLiasse.MajTitreGrilles;
var
  stTitre, stLesCriteres: string;
  i: integer;
begin
  stLesCriteres := ConstruireParametreCritere;
  for i := 0 to fTLesGrilles.Detail.Count - 1 do
  begin
    stTitre := TitreEtatSynthese(Crit, fTLesGrilles.Detail[i].GetValue('TAG'));
    { On mémorise les infos du titre au niveau de la grille pour les remonter au moment de l'édition au niveau de la mère }
    if not fTLesGrilles.Detail[i].FieldExists('CRITERES') then
    begin
      fTLesGrilles.Detail[i].AddChampSupValeur('TITRE_ETAT', stTitre);
      fTLesGrilles.Detail[i].AddChampSupValeur('CRITERES', 'TITRE=' + stTitre + '`'
        + stLesCriteres);
      fTLesGrilles.Detail[i].AddChampSupValeur('DUPLICATA', Crit.MentionFiligrane);
    end;
  end;
end;

procedure TFEtatLiasse.RenommePourPublifi;
var
  stFichierPS: string;
begin
  if (FileExists(ChangeStdDatPath('$DOS\' +
    'PGIExpert_CCS5_Etats de synthèse.PS')) or
    FileExists(ChangeStdDatPath('$DOS\' + 'PGIExpert_CCS5_Etat de synthèse.PS')))
    then
  begin
    stFichierPS := FaitLibelleEtatSynthese(Crit.TE, Crit.bEtat, Crit.AvecDetail,
      False, True);
    if VH^.bAnalytique then
      stFichierPS := stFichierPS + TraduireMemoire('(analytique)');
    stFichierPS := stFichierPS + '.PS';
    if FileExists(ChangeStdDatPath('$DOS\' + stFichierPS)) then
      DeleteFile(ChangeStdDatPath('$DOS\' + stFichierPS));
    if not RenameFile(ChangeStdDatPath('$DOS\' +
      'PGIExpert_CCS5_Etat de synthèse.PS'), ChangeStdDatPath('$DOS\' +
      stFichierPS)) then
      RenameFile(ChangeStdDatPath('$DOS\' +
        'PGIExpert_CCS5_Etats de synthèse.PS'), ChangeStdDatPath('$DOS\' +
        stFichierPS));
  end;
end;

function TFEtatLiasse.TitreEtatSynthese(Crit: tCritEdtPCL; iTag: integer):
  string;
var
  stComplAna: string;
begin
  stComplAna := '';
  if VH^.bAnalytique then
  begin
    if Crit.bTLI then
      stComplAna := TraduireMemoire(' sur tables libres analytiques')
    else if Crit.bUnEtatParSection then
    begin
      if Crit.bLibelleDansLesTitres then
        stComplAna := ' sur ' + GetColonneSQL('SECTION', 'S_ABREGE',
          'S_SECTION="' + Crit.stSectionDe + '"')
      else
        stComplAna := ' sur ' + Crit.stSectionDe;
    end
    else
      stComplAna := TraduireMemoire(' sur sections analytiques');
  end;
  Result := FaitLibelleEtatSynthese(Crit.TE, (iTag = 1) or (iTag = 4), (iTag = 2)
    or (iTag = 5), (iTag = 4) or (iTag = 5), False);
  if (iTag = 3) then
    Result := TraduireMemoire('Non Affectés');
  Result := TraduireMemoire(Result);
  Result := Result + stComplAna;
end;

procedure TFEtatLiasse.MajCaptionAnalytique;
begin
  lbl_Axe.Caption := RechDom('AFAXE', Crit.stAxe, False);
  if Crit.bTLI then
  begin
    Caption := Caption + ' SUR TABLES LIBRES ANALYTIQUES';
    lbl_lblSectionDe.Caption := 'Tables ';
    lbl_SectionDe.SetBounds(16 + lbl_lblSectionDe.Width, 25, 490 - 16 -
      lbl_lblSectionDe.Width, 15);
    lbl_SectionDe.Caption := Crit.stSectionDe;
    lbl_lblSectionA.Caption := 'Tables exceptées';
    lbl_SectionA.SetBounds(16 + lbl_lblSectionA.Width, 46, 550 - 16 -
      lbl_lblSectionA.Width, 15);
    lbl_SectionA.Caption := Crit.stSectionA;
  end
  else
  begin
    if Crit.bUnEtatParSection then
    begin
      if Crit.bLibelleDansLesTitres then
        Caption := Caption + ' sur ' + GetColonneSQL('SECTION', 'S_ABREGE',
          'S_SECTION="' + Crit.stSectionDe + '"')
      else Caption := Caption + ' sur ' + Crit.stSectionDe;
    end else
      Caption := Caption + ' SUR SECTIONS ANALYTIQUES';
    if (Crit.bJoker) then
    begin
      lbl_lblSectionDe.Caption := 'Sections ';
      lbl_SectionDe.SetBounds(16 + lbl_lblSectionDe.Width, 25, 490 - 16 -
        lbl_lblSectionDe.Width, 15);
      lbl_SectionDe.Caption := Crit.stSectionDe;
      lbl_lblSectionA.Visible := False;
      lbl_SectionA.Visible := False;
    end else
    begin
      lbl_SectionDe.Caption := Crit.stSectionDe;
      lbl_SectionA.Caption := Crit.stSectionA;
    end;
  end;
end;


procedure TFEtatLiasse.BOUVRIRClick(Sender: TObject);
var
  OD: TOpenDialog;
begin
  OD := TOpenDialog.Create(Application);
  OD.DefaultExt := 'txt';
  OD.Filter := '*.*|*.txt';
  if OD.Execute then
  begin
    FModeleName := OD.FileName;
    PagesChange(nil);
  end;
  OD.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 02/04/2007
Modifié le ... :   /  /
Description .. : Enregistrement de la maquette en base
Mots clefs ... :
*****************************************************************}

procedure TFEtatLiasse.EnregistreMaquette(TL: HTStringList; pstFichier, pstCrit3, pstPredefini: string);
var
  pT: integer;
  lRet: integer;
  stTitre: string;
  lstCrit2: string;
  iRet: integer;
begin
  { Enregistrement du fichier sur disque }
  TL.SaveToFile(pstFichier);
  {Enregistrement de la maquette en base }
  case Crit.TE of
    esBIL, esBILA: lStCrit2 := 'BIL';
    esCR, esCRA: lStCrit2 := 'CR';
    esSIG, esSIGA: lStCrit2 := 'SIG';
  end;
  pT := Pos('(', FTitre.Text);
  if pT > 0 then stTitre := Copy(FTitre.Text, 1, pT - 1)
  else stTitre := FTitre.Text;
  iRet := mrYes;
  if ExisteSQL('SELECT YFS_NOM FROM YFILESTD WHERE YFS_CODEPRODUIT="COMPTA" AND ' +
    'YFS_LANGUE="' + V_PGI.LanguePrinc + '" AND YFS_PREDEFINI="' + pstPredefini + '" AND ' +
    'YFS_CRIT1="ETATSYNTH" AND YFS_CRIT2="' + lstCrit2 + '" AND YFS_CRIT3="' + pstCrit3 +
    '" AND YFS_NOM="' + ExtractFileName(pstFichier) + '"') then
    iRet := PGIAsk('Le fichier existe déjà. Voulez-vous l''écraser ?');
  if (iRet = mrYes) then
  begin
    lRet := AGL_YFILESTD_IMPORT(pStFichier, 'COMPTA', ExtractFileName(pstFichier), ExtractFileExt(pstFichier),
      'ETATSYNTH', lstCrit2, pStCrit3, '', '', '-', '-', '-', '-', '-', V_PGI.LanguePrinc, pStPredefini, stTitre, '000000');
    if lRet = -1 then
      PGIInfo('Enregistrement de la maquette terminé.', stTitre)
    else
      PGIInfo(AGL_YFILESTD_GET_ERR(lRet) + #13#10 + pstFichier);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 02/04/2007
Modifié le ... :   /  /
Description .. : Contrôle le droit d'enregistrer la maquette;
Mots clefs ... :
*****************************************************************}

function TFEtatLiasse.JePeuxEnregistrerMaquette(pEtatOk: integer): boolean;
begin
  Result := ((((Crit.TE = esBIL) or (Crit.TE = esBILA)) and (pEtatOk = 2)) or
    (((Crit.TE = esCR) or (Crit.TE = esSIG) or (Crit.TE = esCRA) or (Crit.TE =
    esSIGA)) and (pEtatOk = 1)));
  if not Result then
  begin
    if (Crit.TE = esBIL) or (Crit.TE = esBILA) then
      MessageAlerte('Champs de contrôle ' +
        CHPCTRL_ACTIF + ' et ' + CHPCTRL_PASSIF +
        ' absents.#10#13Vous ne pouvez pas enregistrer cette maquette.')
    else if ((Crit.TE = esCR) or (Crit.TE = esSIG) or (Crit.TE = esCRA) or (Crit.TE
      = esSIGA)) then
      MessageAlerte('Champs de contrôle ' + CHPCTRL_RESULTAT +
        ' absent.#10#13Vous ne pouvez pas enregistrer cette maquette.');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 02/04/2007
Modifié le ... :   /  /
Description .. : Contrôle du niveau d'enregistrement en PCL
Mots clefs ... :
*****************************************************************}

function TFEtatLiasse.ControleStandardPCL(var pStCrit3: string; var pstPredefini: string): boolean;
var
  iStd: integer;
  NumStd: integer;
  bStdCegid: boolean;
  stFileName: string;
begin
  Result := False;
  pStCrit3 := '';
  pStPredefini := 'STD';
  if not (ctxPCL in V_PGI.PGIContexte) then exit;
  case Crit.TE of
    esCR, esCRA: iStd := 3;
    esBIL, esSIG, esSIGA, esBILA: iStd := 4;
  else iStd := 0;
  end;
  stFileName := ExtractFileName(FModeleName);
  if IsNumeric(Copy(stFileName, iStd, 2)) then
  begin
    NumStd := StrToInt(Copy(stFileName, iStd, 3));
    bStdCegid := (EstSpecif('51502') and (ctxStandard in V_PGI.PGIContexte));
    if ((not bStdCegid) and (NumStd <= 20)) then
    begin
      MessageAlerte('Enregistrement impossible : Standard CEGID !');
      exit;
    end
    else if ((not (ctxStandard in V_PGI.PGIContexte)) and (NumStd < 100)) then
    begin
      MessageAlerte('Enregistrement impossible : Maquette standard !');
      exit;
    end;
    pStCrit3 := IntToStr(NumStd);
    if NumStd <= 20 then pStPredefini := 'CEG';
    Result := True;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 02/04/2007
Modifié le ... :   /  /
Description .. : Fonction de copie de la maquette de la grille telle que saisie
Suite ........ : par l'utilisateur vers la liste en mémoire
Mots clefs ... :
*****************************************************************}

function TFEtatLiasse.GrilleVersMaquette(TL: HTStringList): integer;
var
  i, kk: integer;
  St: string;
  stCell: string;
  EtatOk: integer;
begin
  Result := -1;
  EtatOk := 0;
  if TL = nil then exit
  else TL.Clear;
  { Copie du titre }
  TL.Add(Ftitre.text);
  { On parcourt la grille }
  for i := 0 to GParams.RowCount - 1 do
  begin
    St := '';
    for kk := 1 to 6 + MAX_COL do
    begin
      stCell := Trim(GParams.CellValues[kk, i]);
      St := St + stCell + '|';
      if ((kk = 5) and ((Crit.TE = esBIL) or (Crit.TE = esBILA)) and ((stCell =
        CHPCTRL_PASSIF) or (stCell = CHPCTRL_ACTIF))) then
        Inc(EtatOK, 1)
      else if ((kk = 5) and ((Crit.TE = esSIG) or (Crit.TE = esCR) or (Crit.TE =
        esCRA) or (Crit.TE = esSIGA)) and (stCell = CHPCTRL_RESULTAT)) then
        Inc(EtatOK, 1);
    end;
    St := St + IntToStr(GParams.RowHeights[i]) + '|';
    if ctxPCL in V_PGI.PGIContexte then
    begin
      case Crit.TE of
        esBIL, esBILA: St := St + 'BIL' + '|';
        esCR, esCRA: St := St + 'CR' + '|';
        esSIG, esSIGA: St := St + 'SIG' + '|';
      end;
    end
    else
    begin
      case Crit.TE of
        esBIL: St := St + 'BIL' + '|';
        esCR: St := St + 'CR' + '|';
        esSIG: St := St + 'SIG' + '|';
        esBILA: St := St + 'BILA' + '|';
        esCRA: St := St + 'CRA' + '|';
        esSIGA: St := St + 'SIGA' + '|';
      end;
    end;
    if not IsLigneVide(i) then
      TL.Add(st);
  end;
  Result := EtatOk;
end;

procedure TFEtatLiasse.BENREGISTREPGEClick(Sender: TObject);
var
  TL: HTStringList;
  EtatOk: integer;
  stFichier: string;
begin
  TL := HTStringList.Create;
  try
    EtatOk := GrilleVersMaquette(TL);
    if JePeuxEnregistrerMaquette(EtatOk) then
    begin
      stFichier := ExtractFilePath(FCHEMINFICHIERPGE.Text) + '\' + FFICHIERPGE.Text;
      EnregistreMaquette(TL, stFichier, '', 'STD');
    end;
  finally
    PNOMFICHIERPGE.Visible := False;
    TL.Free;
  end;
end;

procedure TFEtatLiasse.BANNULERClick(Sender: TObject);
begin
  PNOMFICHIERPGE.Visible := False;
end;

procedure TFEtatLiasse.FormResize(Sender: TObject);
begin
  PNOMFICHIERPGE.Left := bSauve.Left - 180;
end;

function FaitMasqueDecimales: string;
var St: string;
  i: integer;
begin
  St := '0.';
  for i := 0 to V_PGI.OkDecV - 1 do
    St := St + '0';
  Result := St;
end;

end.

