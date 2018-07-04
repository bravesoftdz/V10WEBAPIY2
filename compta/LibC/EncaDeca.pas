unit EncaDeca;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, FileCtrl,
  StdCtrls, Buttons, ExtCtrls, Grids, Hcompte, Hctrls, ComCtrls,
{$IFDEF EAGLCLIENT}
  uTob, 
{$ELSE}
  DBGrids,
  DBCtrls,
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  HDB,
  ParamDBG,
  PrintDBG,
{$ENDIF}
{$IFDEF VER150}
   Variants,
{$ENDIF}
  hmsgbox, HTB97, Mask, Ent1, HEnt1, Menus, HSysMenu, Hqry,
  SaisUtil, SaisComm, Saisie, LettUtil, EncUtil,
  CPTiers_TOM,
  FichComm,
  EcheMono, HStatus, EtapeReg, EcheUnit,
  UObjFiltres, {JP 23/08/04}
  ParamDat, EcheMPA,
  Hspliter, EncTiers, DecaBqe, Choix, TofVerifRib,
  BANQUECP_TOM,
{$IFDEF V530}
  EdtEtat,
{$ELSE}
  EdtREtat,
{$ENDIF}
  ed_tools, HPanel, UiUtil, UtilPGI,
  // Lancement Saisie Piece ert Bor :
  uTofCPMulMvt,
  SaisBor,
  ULibEcriture,
  ADODB;

procedure EncaisseDecaisse(Enc: boolean; Circuit1, Circuit2: string; Generation, AffectBanque: boolean; SorteLettre: TSorteLettre);
procedure DecaisseCircuit(n: Integer);
procedure EncDecChequeTraiteBOR(Chq: TSorteLettre);

type
  TFEncaDeca = class(TForm)
    DockBottom: TDock97;
    HPB: TToolWindow97;
    BImprimer: TToolbarButton97;
    BValide: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BAide: TToolbarButton97;
    bSelectAll: TToolbarButton97;
    BLanceSaisieBor: TToolbarButton97;
    BLanceSaisie: TToolbarButton97;
    BAgrandir: TToolbarButton97;
    BRecherche: TToolbarButton97;
    BZoomPiece: TToolbarButton97;
    BModifEche: TToolbarButton97;
    bAffectBqe: TToolbarButton97;
    BCtrlRib: TToolbarButton97;
    BModifRIB: TToolbarButton97;
    BZoomModele: TToolbarButton97;
    Pages: TPageControl;
    PComptes: TTabSheet;
    Bevel1: TBevel;
    HLabel4: THLabel;
    HEtape: THLabel;
    HCpteGen: THLabel;
    H_GENERAL: TLabel;
    H_GENEGENERATION: TLabel;
    BZOOMETAPE: TToolbarButton97;
    H_TOTDEVISE: TLabel;
    E_GENERAL: THCpteEdit;
    ETAPE: THValComboBox;
    GENEGENERATION: THCpteEdit;
    FSelectAll: TCheckBox;
    rDetail: TRadioButton;
    rMasse: TRadioButton;
    cBAP: TCheckBox;
    cModeleBAP: THValComboBox;
    ERSOLDED: THNumEdit;
    PParam: TTabSheet;
    Bevel2: TBevel;
    HLabel7: THLabel;
    Label14: TLabel;
    HLabel1: THLabel;
    HLabel17: THLabel;
    HLabel9: THLabel;
    Label1: TLabel;
    Label2: TLabel;
    MP_CATEGORIE: THValComboBox;
    E_MODEPAIE: THValComboBox;
    GUIDE: THValComboBox;
    E_DEVISE: THValComboBox;
    TYPEETAPE: THValComboBox;
    CodeAFB: THMultiValComboBox;
    PGenere: TTabSheet;
    Bevel5: TBevel;
    H_MODEGENERE: TLabel;
    HDateReg: THLabel;
    MODEGENERE: THValComboBox;
    DATEGENERATION: THCritMaskEdit;
    CAutoEnreg: TCheckBox;
    GLOBAL: TCheckBox;
    DATEECHEANCE: THCritMaskEdit;
    CChoixDate: TCheckBox;
    GereAccept: TCheckBox;
    PGestionPDF: TTabSheet;
    Bevel7: TBevel;
    LNBLIGNES: THLabel;
    TREPSPOOLER: THLabel;
    CAutoEnreg1: TCheckBox;
    NBLIGNES: THCritMaskEdit;
    Spooler: TCheckBox;
    RepSpooler: THCritMaskEdit;
    XFichierSpooler: TCheckBox;
    PCircuit: TTabSheet;
    Label3: TLabel;
    Panel1: TPanel;
    g1: TGroupBox;
    HLabel6: THLabel;
    HLabel12: THLabel;
    bListeLots: TToolbarButton97;
    E_SUIVDEC: THValComboBox;
    E_NOMLOT: THCritMaskEdit;
    g2: TGroupBox;
    HLabel13: THLabel;
    HLotDest: THLabel;
    CircuitDest: THValComboBox;
    LotDest: TEdit;
    XX_WHERE1: TEdit;
    PEcritures: TTabSheet;
    Bevel3: TBevel;
    TE_JOURNAL: THLabel;
    TE_NATUREPIECE: THLabel;
    TE_DATECOMPTABLE: THLabel;
    TE_DATECOMPTABLE2: THLabel;
    TE_EXERCICE: THLabel;
    TE_DATEECHEANCE: THLabel;
    TE_DATEECHEANCE2: THLabel;
    E_EXERCICE: THValComboBox;
    E_JOURNAL: THValComboBox;
    E_NATUREPIECE: THValComboBox;
    E_DATECOMPTABLE: THCritMaskEdit;
    E_DATECOMPTABLE_: THCritMaskEdit;
    E_NUMECHE: THCritMaskEdit;
    XX_WHERE: TEdit;
    MP_ENCAISSEMENT: THCritMaskEdit;
    E_QUALIFPIECE: THCritMaskEdit;
    E_ECHE: THCritMaskEdit;
    E_ETATLETTRAGE: THCritMaskEdit;
    E_TRESOLETTRE: THCritMaskEdit;
    XX_WHEREENC: TEdit;
    cVoirCircuit: TCheckBox;
    XX_WHERE3: TEdit;
    E_DATEECHEANCE: THCritMaskEdit;
    E_DATEECHEANCE_: THCritMaskEdit;
    cFactCredit: TCheckBox;
    PComplements: TTabSheet;
    Bevel4: TBevel;
    TE_NUMEROPIECE: THLabel;
    HLabel2: THLabel;
    HLabel3: THLabel;
    TE_DEBIT: TLabel;
    TE_DEBIT_: TLabel;
    TE_CREDIT: TLabel;
    TE_CREDIT_: TLabel;
    iTiers: TToolbarButton97;
    iPiece: TToolbarButton97;
    H_OrdreTri: THLabel;
    TE_ETABLISSEMENT: THLabel;
    E_NUMEROPIECE: THCritMaskEdit;
    E_NUMEROPIECE_: THCritMaskEdit;
    E_AUXILIAIRE: THCpteEdit;
    E_DEBIT: THCritMaskEdit;
    E_DEBIT_: THCritMaskEdit;
    E_CREDIT: THCritMaskEdit;
    E_CREDIT_: THCritMaskEdit;
    OrdreTri: THValComboBox;
    E_ETABLISSEMENT: THValComboBox;
    PLibres: TTabSheet;
    Bevel6: TBevel;
    TT_TABLE0: THLabel;
    TT_TABLE1: THLabel;
    TT_TABLE2: THLabel;
    TT_TABLE3: THLabel;
    TT_TABLE4: THLabel;
    TT_TABLE5: THLabel;
    TT_TABLE6: THLabel;
    TT_TABLE7: THLabel;
    TT_TABLE8: THLabel;
    TT_TABLE9: THLabel;
    T_TABLE4: THCpteEdit;
    T_TABLE3: THCpteEdit;
    T_TABLE2: THCpteEdit;
    T_TABLE1: THCpteEdit;
    T_TABLE0: THCpteEdit;
    T_TABLE5: THCpteEdit;
    T_TABLE6: THCpteEdit;
    T_TABLE7: THCpteEdit;
    T_TABLE8: THCpteEdit;
    T_TABLE9: THCpteEdit;
    Dock971: TDock97;
    PFiltres: TToolWindow97;
    BFiltre: TToolbarButton97;
    BChercher: TToolbarButton97;
    FFiltres: THValComboBox;
    G: THGrid;
    QEnc: THQuery;
    FindMvt: TFindDialog;
    POPS: TPopupMenu;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    HMTrad: THSystemMenu;
    HDiv: THMsgBox;
    HMEnc: THMsgBox;
    BReduire: TToolbarButton97; // FQ 21450

    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GDblClick(Sender: TObject);
    procedure GMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BZoomPieceClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure BAgrandirClick(Sender: TObject);
    procedure BRechercheClick(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure DATEGENERATIONExit(Sender: TObject);
    procedure FindMvtFind(Sender: TObject);
    procedure BModePaieClick(Sender: TObject);
    procedure POPSPopup(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure E_DEVISEChange(Sender: TObject);
    procedure MP_CATEGORIEChange(Sender: TObject);
    procedure ETAPEChange(Sender: TObject);
    procedure E_GENERALExit(Sender: TObject);
    procedure GENEGENERATIONExit(Sender: TObject);
    procedure E_GENERALChange(Sender: TObject);
    procedure GENEGENERATIONChange(Sender: TObject);
    procedure BZOOMETAPEClick(Sender: TObject);
    procedure BModifEcheClick(Sender: TObject);
    procedure E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
    procedure BAideClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure cTypeVisuChange(Sender: TObject);
    procedure GRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure bAffectBqeClick(Sender: TObject);
    procedure bListeLotsClick(Sender: TObject);
    procedure cBAPClick(Sender: TObject);
    procedure cVoirCircuitClick(Sender: TObject);
    procedure CChoixDateClick(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure BModifRIBClick(Sender: TObject);
    procedure BZoomModeleClick(Sender: TObject);
    procedure BCtrlRibClick(Sender: TObject);
    procedure CAutoEnreg1Click(Sender: TObject);
    procedure SpoolerClick(Sender: TObject);
    procedure BLanceSaisieClick(Sender: TObject);
    procedure BLanceSaisieBorClick(Sender: TObject);
    procedure bSelectAllClick(Sender: TObject); // FQ 21450

  private
    ObjFiltre: TObjFiltre; {JP 23/08/04}
    DateGene, NowFutur: TDateTime;
    DEV: RDEVISE;
    CollecContreP, CollecTreso, NoBanque, BanqueVide, JalFerme: Boolean;
    OkMultiSession: Boolean;
    TGuide, LesM: TList;
    LaSelection: HTStringList;
    EntGuide: RMVT;
    TauxRef: Double;
    GX, GY: integer;
    PreVirMP, FormatCFONB, Document, EnvoiTrans: String3;
    PreVirEche: TDateTime;
    PreVirEgal, OkCFONB, OkBordereau: boolean;
    XX: array[1..100] of string;
    WMinX, WMinY: Integer;
    ParTiers, EnDev: Boolean;
    TitreF: string;
    MinNum, MaxNum: Longint;
    StWRib: string;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    procedure GetCellCanvas(Acol, ARow: LongInt; Canvas: TCanvas; AState: TGridDrawState);
    // Entete, Chargements
    procedure MajCaption;
    procedure EtudierGuide;
    procedure PositionneFerme(Jal: string);
    procedure GUIDEChange(Sender: TObject);
    procedure SwapCat(Vide: Boolean);
    procedure AffecteColTypes;
    procedure InitEcran;
    procedure InitCircuit;
    procedure InitLettreCheque;
    procedure InitLettreTraite;
    procedure InitGeneration;
    procedure InitLeReste;
    procedure InitEncOuDec;
    procedure PositionneBB;
    // Encaissement
    procedure CreerGuideTemp;
    procedure CreerLignesTemp(k: integer);
    procedure LGContreCommun(Q: TQuery; OGuiContre: TOBM; Auxi, GeneMP, OldRef, OldLib, OldPrevi, OldRib, OldReftire: string; k: integer);
    procedure LGContreLigne(Q: TQuery; DCur, CCur: double; OldEche: TDateTime; OldPaie: String3);
    procedure LGPremLigne(Q: TQuery; O: TOBM);
    function MontantGuideSt(O: TOBM; Debit: boolean): string;
    function MontantARegler(O: TOBM; Debit, AvecCouv: boolean; Quoi: TSorteMontant): Double;
    function JalBanque(k: integer): string;
    procedure FabricLesM;
    procedure RetoucheLesM(var CarePDF: Boolean);
    // Contrôles
    function JeValidePas: boolean;
    function ModeGenereAlimente: Boolean;
    function CoherGroupe: byte;
    function AcqModeDate(var MP: String3; var DD: TDateTime): boolean;
    function CoherVirPre(var MP: String3; var DD: TDateTime): boolean;
    function LesMeme: boolean;
    function VerifLesEche: boolean;
    function VerifLesDevise: boolean;
    function VerifLesRIB: boolean;
    function VerifLesPrevi: boolean;
    function EstCptTicTid : Boolean ; // FQ 15590 SBO 04/04/20005
    // Recherches, SQL
    procedure RempliGrid;
    procedure RempliTiers(TEche: TList; Lig: integer);
    procedure ChangeTypeGrid;
    // Sélections
    procedure CocheDecoche(Lig: integer; Next: boolean);
    procedure CalculDebitCredit;
    procedure AjouteOBM(O: TOBM; Lig1, Lig2: Integer; Plus: Boolean);
    procedure ConstitueLaSelection;
    function GetLeDebitCur(O: TOBM; AvecCouv: boolean): double;
    function GetLeCreditCur(O: TOBM; AvecCouv: boolean): double;
    function GetLaCouvCur(O: TOBM): double;
    // Circuit
    procedure UpdateCircuit;
    procedure UpdateOBMCircuit(O: TOBM);
    procedure GereOngletCircuit(Montrer: boolean);
    // Affectation prévisionnelle des banques
    procedure AddGridAffBqe(G: THGrid; O: TOBM);
    procedure DoAffectBqe;
    function CountBqe: integer;
    // Saisie pièce et borderau
    procedure ClickSaisiePiece;
    procedure ClickSaisieBor;
    procedure InitCaption;
  public
    Enc, FindFirst, Generation: Boolean;
    Circuit1, Circuit2: string;
    AffectBanque: Boolean;
    SorteLettre: TSorteLettre;
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  ULibExercice,
  CPProcGen,
  CPProcMetier,
  {$ENDIF MODENT1}
  ParamSoc;

const
  Decoupe: Boolean = FALSE;
const
  OkDecoupe: Boolean = TRUE;
const
  IndDecoupe: Integer = 0;

procedure EncaisseDecaisse(Enc: boolean; Circuit1, Circuit2: string; Generation, AffectBanque: boolean; SorteLettre: TSorteLettre);
var
  X: TFEncaDeca;
  PP: THPanel;
begin
  if PasCreerDate(V_PGI.DateEntree) then Exit;
  if Enc then
  begin
    if _Blocage(['nrBatch', 'nrCloture', 'nrLettrage', 'nrEnca', 'nrRelance'], True, 'nrEnca') then Exit;
  end else
  begin
    if _Blocage(['nrBatch', 'nrCloture', 'nrLettrage', 'nrDeca'], True, 'nrDeca') then Exit;
  end;
  X := TFEncaDeca.Create(Application);
  X.ENC := Enc;
  X.Circuit1 := Circuit1;
  X.Circuit2 := Circuit2;
  X.Generation := Generation;
  X.AffectBanque := AffectBanque;
  X.SorteLettre := SorteLettre;
  PP := FindInsidePanel;
  if PP = nil then
  begin
    try
      X.ShowModal;
    finally
      X.Free;
      if Enc then _Bloqueur('nrEnca', False) else _Bloqueur('nrDeca', False);
    end;
    Screen.Cursor := SyncrDefault;
  end else
  begin
    InitInside(X, PP);
    X.Show;
  end;
end;

procedure EncDecChequeTraiteBOR(Chq: TSorteLettre);
var
  Q: TQuery;
  s1, s2, stSQL: string;
begin
  s1 := '';
  s2 := '';
  if Chq <> tslTraite then
  begin
    stSQL := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="CID" AND CC_LIBELLE<>"" ORDER BY CC_CODE DESC';
    Q := OpenSQL(stSQL, True);
    if not Q.EOF then
    begin
      s2 := Q.FindField('CC_CODE').AsString;
      Q.Next;
      if not Q.EOF then s1 := Q.FindField('CC_CODE').AsString;
    end;
    Ferme(Q);
    EncaisseDecaisse(False, s1, s2, True, True, Chq);
  end else
  begin
    EncaisseDecaisse(True, '', '', True, True, Chq);
  end;
end;

procedure DecaisseCircuit(n: Integer);
var
  Q: TQuery;
  s1, s2, s3, stSQL: string;
  Generation: Boolean;
  AffectBanque: Boolean;
begin
  if n = 1 then s1 := '' else s1 := 'CD' + IntToStr(n - 1);
  s2 := 'CD' + IntToStr(n);
  //Charger l'étape du circuit
  stSQL := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="CID" AND CC_LIBELLE<>"" AND CC_CODE="' + s2 + '"';
  Q := OpenSQL(stSQL, True);
  AffectBanque := (Q.FindField('CC_ABREGE').AsString = 'X');
  Ferme(Q);
  //Il faut vérifier si c'est est la dernière étape du circuit
  s3 := 'CD' + Inttostr(n + 1);
  stSQL := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="CID" AND CC_LIBELLE<>"" AND CC_CODE="' + s3 + '"';
  Q := OpenSQL(stSQL, True);
  Generation := Q.EOF;
  Ferme(Q);
  EncaisseDecaisse(False, s1, s2, Generation, AffectBanque, tslAucun);
end;

{====================================== FORM, GRID ==========================================}
procedure TFEncaDeca.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  DetruitGuideEnc(True);
  VideListe(TGuide);
  TGuide.Free;
  VideListe(LesM);
  LesM.Free;
  LaSelection.Clear;
  LaSelection.Free;
  G.ClearSelected;
  PurgePopup(POPS);
  G.VidePile(True);
  RegSaveToolbarPos(Self, 'EncaDeca');
  if Parent is THPanel then
  begin
    if ENC then _Bloqueur('nrEnca', False) else _Bloqueur('nrDeca', False);
  end;

  {JP 23/08/2004 : Nouvelle gestion des filtres}
  if Assigned(ObjFiltre) then FreeAndNil(ObjFiltre);
end;

procedure TFEncaDeca.FormCreate(Sender: TObject);
var
  Composants: TControlFiltre; {JP 23/08/2004}
begin
  PopUpMenu := ADDMenuPop(PopUpMenu, '', '');
  TGuide := TList.Create;
  LesM := TList.Create;
  LaSelection := HTStringList.Create;
  LaSelection.Sorted := True;
  LaSelection.Clear;
  G.ClearSelected;
  OkCFONB := False;
  FormatCFONB := '';
  OkBordereau := False;
  Document := '';
  EnvoiTrans := '';
  WMinX := Width;
  WMinY := 258;
  MinNum := 999999999;
  MaxNum := 0;
  JalFerme := False;
  RegLoadToolbarPos(Self, 'EncaDeca');
  {JP 23/08/2004 : Nouvelle gestion des filtres}
  Composants.PopupF := POPF;
  Composants.Filtres := FFILTRES;
  Composants.Filtre := BFILTRE;
  Composants.PageCtrl := PAGES;
  ObjFiltre := TObjFiltre.Create(Composants, '');
end;

procedure TFEncaDeca.AffecteColTypes;
begin
  if G.ListeParam = 'ENCEURODET' then
  begin
    G.ColTypes[6] := 'R';
    G.ColTypes[7] := 'R';
    G.ColTypes[8] := 'R';
    G.ColTypes[9] := 'D';
  end else
  begin
    G.ColTypes[3] := 'R';
    G.ColTypes[4] := 'R';
    G.ColTypes[5] := 'R';
  end;
end;

procedure TFEncaDeca.InitEcran;
begin
  {Généralités, Grid}
  G.ListeParam := 'ENCEURODET';
  AffecteColTypes;
  HMTrad.ResizeGridColumns(G);
  PGestionPDF.TabVisible := SorteLettre in [tslCheque, tslTraite, tslBOR];
  OkDecoupe := TRUE;
  if not PGestionPDF.TabVisible then
  begin
    CAutoEnreg1.Checked := FALSE;
    NBLIGNES.Text := '0';
    OkDecoupe := FALSE;
    Spooler.Checked := FALSE;
    RepSpooler.Text := '';
  end;
  if SorteLettre = tslCheque then TitreF := 'CHEQUE' else
    if SorteLettre = tslTraite then TitreF := 'ENCTRAITE' else
    if SorteLettre = tslBOR then TitreF := 'DECBOR' else
    TitreF := 'ENCDEC' + Chr(48 + Ord(ENC));
  G.GetCellCanvas := GetCellCanvas;
  G.RowCount := 2;
  OrdreTri.ItemIndex := 0;
  NoBanque := True;
  BanqueVide := False;
  {Type de liste}
  rDetail.Checked := True;
  cTypeVisuChange(nil);
  {Initialisations zones}
  DateGene := V_PGI.DateEntree;
  E_EXERCICE.Value := EXRF(VH^.Entree.Code);
  {JP 21/08/07 : FQ 20025 : Si pas de filtre, mettre à jour les dates par rapport à l'exercice
  E_DATECOMPTABLE.Text := StDate1900;
  E_DATECOMPTABLE_.Text := StDate2099;}
  ExoToEdDates(E_EXERCICE.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_);
  DATEGENERATION.Text := DateToStr(DateGene);
  DATEECHEANCE.Text := StDate1900;
  E_DATEECHEANCE.Text := StDate1900;
  E_DATEECHEANCE_.Text := StDate2099;
  E_DEVISE.Value := V_PGI.DevisePivot;
  {JP 21/08/07 : FQ 21158 : on supprime l'underscore de ER_SOLDED, pour que le champ ne soit pas repris dans la requête}
  ChangeMask(ERSOLDED, V_PGI.OkDecV, V_PGI.SymbolePivot);
  {Affectations diverses}
  if AffectBanque then bAffectBqe.Hint := HMEnc.Mess[39]
  else bAffectBqe.Hint := HMEnc.Mess[40];
  CircuitDest.Value := Circuit2;
  CircuitDest.Enabled := False;
  {Positionnement}
  Pages.ActivePage := PComptes;
end;

procedure TFEncaDeca.InitCircuit;
var
  Q: TQuery;
begin
  PCircuit.TabVisible := (Circuit2 <> '');
  if PCircuit.TabVisible then
  begin
    if Circuit1 <> '' then E_SUIVDEC.Value := Circuit1;
    g1.Visible := (Circuit1 <> '');
    E_SUIVDEC.Value := Circuit1;
    E_SUIVDEC.Enabled := False;
    cVoirCircuit.Visible := SorteLettre in [tslCheque, tslBOR];
    cBAP.Enabled := True;
    cBAP.Checked := True;
    cModeleBAP.Enabled := True;
    BAffectBqe.Enabled := AffectBanque;
  end else
  begin
    cBAP.Enabled := False;
    cModeleBAP.Enabled := False;
    if not ENC then
    begin
      Q := OpenSQL('Select count(*) from CHOIXCOD Where CC_TYPE="CID" and CC_LIBELLE<>""', True);
      if Q.Fields[0].AsInteger <= 0 then
      begin
        cVoirCircuit.Visible := False;
        cVoirCircuit.Checked := False;
        GereOngletCircuit(False);
      end;
      Ferme(Q);
    end;
  end;
end;

procedure TFEncaDeca.InitLettreCheque;
begin
  if SorteLettre <> tslCheque then Exit;
  CircuitDest.DataType := '';
  CircuitDest.Items.Add(HMEnc.Mess[45]);
  CircuitDest.ItemIndex := 0;
  XX_WHERE3.Text := 'MP_LETTRECHEQUE="X"';
  GLOBAL.Checked := False;
  GLOBAL.Enabled := False;
  E_MODEPAIE.Vide := False;

  {JP  23/08/04 : FQ 13855
  BModifEche.Enabled:=False ;}
  BModifEche.Visible := False;

  cVoirCircuit.Checked := False;
  GereOngletCircuit(False);
end;

procedure TFEncaDeca.InitLettreTraite;
begin
  if ((SorteLettre <> tslTraite) and (SorteLettre <> tslBOR)) then Exit;
  CircuitDest.DataType := '';
  CircuitDest.Items.Add(HMEnc.Mess[47]);
  CircuitDest.ItemIndex := 0;
  XX_WHERE3.Text := 'MP_LETTRETRAITE="X"';
  E_MODEPAIE.Vide := False;

  {JP  23/08/04 : FQ 13855
  BModifEche.Enabled:=False ;}
  BModifEche.Visible := False;

  cVoirCircuit.Checked := False;
  GereOngletCircuit(False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 30/04/2004
Modifié le ... : 30/04/2004
Description .. : modification du caption de la fenetre selon l'etape du circuit
Suite ........ : !
Mots clefs ... :
*****************************************************************}
procedure TFEncaDeca.InitCaption;
var
  Q: Tquery;
begin
  if (SorteLettre = tslBOR) then Caption := HMEnc.Mess[48]
  else if (SorteLettre = tslTraite) then Caption := HMEnc.Mess[51]
  else if (SorteLettre = tslCheque) then Caption := HMEnc.Mess[44]
  else if (Circuit2 <> '') then
  begin
    Q := OpenSQL('SELECT CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="CID" AND CC_LIBELLE<>"" AND CC_CODE="' + Circuit2 + '"', true);
    if (not Q.EOF) then Caption := Caption + ' - ' + Q.Fields[0].AsString;
    Ferme(Q);
  end;
end;

procedure TFEncaDeca.InitGeneration;
begin
  if Generation then
  begin
    HEtape.Caption := HMEnc.Mess[42];
    LotDest.Text := '';
    LotDest.Visible := False;
    HLotDest.Visible := False;
    cBAP.Visible := False;
    CBAP.Checked := False;
    cModeleBAP.Visible := False;
  end else
  begin
    HEtape.Caption := HMEnc.Mess[43];
    H_GENEGENERATION.Visible := False;
    HCpteGen.Visible := False;
    GeneGeneration.Visible := False;
    Global.Visible := False;
    HDateReg.Visible := False;
    DateGeneration.Visible := False;
  end;
  PGenere.TabVisible := Generation;
  BModifRIB.Enabled := Generation;
end;

procedure TFEncaDeca.InitEncOuDec;
begin
  if ENC then
  begin
    HelpContext := 7496000;
    if SorteLettre = tslTraite then ETAPE.DataType := 'ttEtapeEncTraite'
    else ETAPE.DataType := 'ttEtapeEncais';
    BAffectBqe.Enabled := False;
    E_GENERAL.ZoomTable := tzGEncais;
    Caption := HMEnc.Mess[1];
    (*
 // JLD 3564   XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC"))' ;
    XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND E_NATUREPIECE="AC")' ;
    *)
 (*
 {$IFDEF JACADI}
    XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC"))' ;
 {$ELSE}
    XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND E_NATUREPIECE="AC")' ;
 {$ENDIF}
 *)
    XX_WHEREENC.Text := 'E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC"))';
    MP_ENCAISSEMENT.Text := 'DEC';
    cVoirCircuit.Visible := False;
  end else
  begin
    HelpContext := 7500000;
    if SorteLettre = tslCheque then ETAPE.DataType := 'ttEtapeDecCheque' else
      if SorteLettre = tslBOR then ETAPE.DataType := 'ttEtapeDecTraite' else
      if ((Generation) and (Circuit2 <> '')) then ETAPE.DataType := 'ttEtapeDecRegle' else
      ETAPE.DataType := 'ttEtapeDecais';
    E_GENERAL.ZoomTable := tzGDecais;
    Caption := HMEnc.Mess[2];
    (*
 // JLD 3564   XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF"))' ;
    XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND E_NATUREPIECE="AF")' ;
    *)
 (*
 {$IFDEF JACADI}
    XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF"))' ;
 {$ELSE}
    XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND E_NATUREPIECE="AF")' ;
 {$ENDIF}
 *)
    XX_WHEREENC.Text := 'E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF"))';
    MP_ENCAISSEMENT.Text := 'ENC';
    PLibres.TabVisible := False;
    cFactCredit.Caption := HDiv.Mess[9];
  end;
  InitTablesLibresTiers(PLibres);
end;

procedure TFEncaDeca.InitLeReste;
begin
  StWRib := '';
  if ETAPE.Values.Count = 1 then ETAPE.Value := ETAPE.Values[0];
  if SorteLettre <> tslAucun then
  begin
    if E_MODEPAIE.Values.Count > 0 then E_MODEPAIE.Value := E_MODEPAIE.Values[0];
  end;
  OkMultiSession := OkDecoupe and (SorteLettre in [tslCheque, tslBOR, tslTraite]);
  CAutoEnreg1.Visible := FALSE;
  NBLIGNES.Visible := FALSE;
  LNBLIGNES.Visible := FALSE;
  if ((SorteLettre = tslAucun) and (not ENC) and (Generation)) then
  begin
    CAutoEnreg.Visible := True;
    CAutoEnreg1.Visible := FALSE;
    NBLIGNES.Visible := FALSE;
    LNBLIGNES.Visible := FALSE;
  end else
  begin
    CAutoEnreg.Visible := FALSE;
    if OkMultiSession then
    begin
      CAutoEnreg1.Visible := TRUE;
      NBLIGNES.Visible := TRUE;
      LNBLIGNES.Visible := TRUE;
    end;
  end;
  // Helpcontext
  BZoomModele.Visible := False;
  if SorteLettre = tslCheque then
  begin
    HelpContext := 7505000;
    PComptes.HelpContext := 7505100;
    PParam.HelpContext := 7505200;
    PGenere.HelpContext := 7505300;
    PEcritures.HelpContext := 7505400;
    PComplements.HelpContext := 7505500;
    BZoomModele.Enabled := True;
    BZoomModele.Visible := True;
  end else if ((Circuit1 = '') and (Circuit2 <> '')) then
  begin
    HelpContext := 7496200;
    PComptes.HelpContext := 7502100;
    PParam.HelpContext := 7502200;
    PGenere.HelpContext := 0;
    PCircuit.HelpContext := 7502300;
    PEcritures.HelpContext := 7502400;
    PComplements.HelpContext := 7502500;
  end else if ((Circuit1 <> '') and (Circuit2 <> '')) then
  begin
    HelpContext := 7496200;
    PComptes.HelpContext := 7503100;
    PParam.HelpContext := 7503200;
    PGenere.HelpContext := 7503300;
    PCircuit.HelpContext := 7503400;
    PEcritures.HelpContext := 7503500;
    PComplements.HelpContext := 7503600;
  end else if not ENC then
  begin
    HelpContext := 7500000;
    PComptes.HelpContext := 7501100;
    PParam.HelpContext := 7501200;
    PGenere.HelpContext := 7501300;
    PCircuit.HelpContext := 0;
    PEcritures.HelpContext := 7501400;
    PComplements.HelpContext := 7501500;
  end;
  if SorteLettre = tslTraite then HelpContext := 7496100;
  if SorteLettre = tslBOR then HelpContext := 7505760;
end;

procedure TFEncaDeca.FormShow(Sender: TObject);
begin
  InitEcran;
  InitCircuit;
  InitGeneration;
  InitEncOuDec;
  InitLettreCheque;
  InitLettreTraite;
  InitLeReste;
  InitCaption;
  PositionneEtabUser(E_ETABLISSEMENT);
  {JP 23/08/2004 : Nouvelle gestion des filtres
  ChargeFiltre(TitreF,FFiltres,Pages) ;}
  ObjFiltre.FFI_TABLE := TitreF;
  ObjFiltre.Charger;
  UpdateCaption(Self);
end;

procedure TFEncaDeca.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
  OkG, Vide: boolean;
begin
  OkG := G.Focused;
  Vide := (Shift = []);
  case Key of
    VK_F9: if Vide then BChercherClick(nil);
    VK_F10: if Vide then BValideClick(nil);
    VK_SPACE: if ((OkG) and (Vide)) then CocheDecoche(G.Row, False);
    VK_F5:
      begin
        Key := 0;
        BZoomPieceClick(nil);
      end;
    {Alt-B}66:
      begin
        Key := 0;
        if (Shift = [ssAlt]) then ClickSaisieBor;
      end;
    80: if Shift = [ssCtrl] then
      begin
        Key := 0;
        BImprimerClick(nil);
      end;
    {Alt-S}83:
      begin
        Key := 0;
        if (Shift = [ssAlt]) then ClickSaisiePiece;
      end;
    KEY_MENUPOP:
      begin
        Key := 0;
        PopS.PopUp(Mouse.CursorPos.X, Mouse.CursorPos.Y);
      end; {F11 : Menu popup}
  end;
end;

procedure TFEncaDeca.GMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  C, R: Longint;
begin
  GX := X;
  GY := Y;
  if ((ssCtrl in Shift) and (Button = mbLeft)) then
  begin
    G.MouseToCell(X, Y, C, R);
    if (R > 0) then CocheDecoche(G.Row, False); //and R<G.RowCount-1
  end;
end;

{======================================= SELEC =============================================}
function TFEncaDeca.GetLaCouvCur(O: TOBM): double;
begin
  if E_DEVISE.Value <> V_PGI.DevisePivot then Result := O.GetMvt('E_COUVERTUREDEV') else
    Result := O.GetMvt('E_COUVERTURE');
end;

function TFEncaDeca.GetLeDebitCur(O: TOBM; AvecCouv: boolean): double;
begin
  if E_DEVISE.Value <> V_PGI.DevisePivot then Result := MontantARegler(O, True, AvecCouv, tsmDevise) else
    Result := MontantARegler(O, True, AvecCouv, tsmPivot);
end;

function TFEncaDeca.GetLeCreditCur(O: TOBM; AvecCouv: boolean): double;
begin
  if E_DEVISE.Value <> V_PGI.DevisePivot then Result := MontantARegler(O, False, AvecCouv, tsmDevise) else
    Result := MontantARegler(O, False, AvecCouv, tsmPivot);
end;

procedure TFEncaDeca.CalculDebitCredit;
var
  TotDeb, TotCred: double;
  i, j: integer;
  O: TOBM;
  TEche: TList;
begin
  TotDeb := 0;
  TotCred := 0;
  if ParTiers then
  begin
    {Calcul des totaux par tiers}
    for i := 1 to G.RowCount - 1 do
    begin
      if EstSelect(G, i) then
      begin
        TEche := TList(G.Objects[0, i]);
        if TEche = nil then Continue;
        for j := 0 to TEche.Count - 1 do
        begin
          O := TOBM(TEche[j]);
          if O = nil then Break;
          TotDeb := TotDeb + GetLeDebitCur(O, True);
          TotCred := TotCred + GetLeCreditCur(O, True);
        end;
      end;
    end;
  end else
  begin
    {Calcul des totaux en détail}
    for i := 1 to G.RowCount - 1 do if EstSelect(G, i) then
      begin
        O := TOBM(G.Objects[0, i]);
        if O = nil then Break;
        TotDeb := TotDeb + GetLeDebitCur(O, True);
        TotCred := TotCred + GetLeCreditCur(O, True);
      end;
  end;
  {JP 21/08/07 : FQ 21158 : on supprime l'underscore de ER_SOLDED, pour que le champ ne soit pas repris dans la requête}
  AfficheLeSolde(ERSOLDED, TotDeb, TotCred);
end;

procedure TFEncaDeca.AjouteOBM(O: TOBM; Lig1, Lig2: Integer; Plus: Boolean);
var
  St, MP: string;
  ii: integer;
begin
  MP := O.GetMvt('E_MODEPAIE');
  St := Format_String(O.GetMvt('E_AUXILIAIRE'), 17) + ';'
    + Format_String(O.GetMvt('E_CONTREPARTIEGEN'), 17) + ';'
    + InttoStr(Lig1) + ';' + Format_String(MP, 3) + ';'
    + DateToStr(O.GetMvt('E_DATEECHEANCE')) + ';'
    + InttoStr(Lig2) + ';';
  if Plus then LaSelection.Add(St) else
  begin
    ii := LaSelection.IndexOf(St);
    if ii >= 0 then LaSelection.Delete(ii);
  end;
end;

procedure TFEncaDeca.CocheDecoche(Lig: integer; Next: boolean);
var
  Cancel: boolean;
begin
  //if Lig=G.RowCount-1 then Exit ;
  if EstSelect(G, Lig) then
  begin
    G.Cells[G.ColCount - 1, Lig] := ' ';
  end
  else
  begin
    G.Cells[G.ColCount - 1, Lig] := '+';
  end;
  Cancel := False;
  GRowEnter(nil, Lig, Cancel, False);
  if ((Lig = G.RowCount - 1) and (Next)) then Next := False;
  if Next then G.Row := Lig + 1;
  G.Invalidate;
  CalculDebitCredit;
end;

procedure TFEncaDeca.GDblClick(Sender: TObject);
var
  C, R: longint;
begin
  G.MouseToCell(GX, GY, C, R);
  if (R > 0) then CocheDecoche(G.Row, True);
end;

{======================================= SQL ===============================================}
procedure TFEncaDeca.ChangeTypeGrid;
begin
  ParTiers := (rMasse.Checked);
  if (G.RowCount <= 2) then //liste vide
  begin
    if (ParTiers) then G.ListeParam := 'ENCEUROT' else G.ListeParam := 'ENCEURODET';
    HMTrad.ResizeGridColumns(G);
    AffecteColTypes;
  end;
  G.SortedCol := -1;
  if ParTiers then
  begin
    bModifEche.Hint := iTiers.Hint;
    bModifEche.Glyph := iTiers.Glyph;
    bZoomPiece.Hint := HMEnc.Mess[33];
    {JP  23/08/04 : FQ 13855
    BModifEche.Enabled:=True ;}
    BModifEche.Visible := True;
  end else
  begin
    bModifEche.Hint := iPiece.Hint;
    bModifEche.Glyph := iPiece.Glyph;
    bZoomPiece.Hint := HMEnc.Mess[34];
    {JP  23/08/04 : FQ 13855
    BModifEche.Enabled:=(SorteLettre=tslAucun);}
    BModifEche.Visible := (SorteLettre = tslAucun);
  end;
end;

procedure TFEncaDeca.RempliGrid;
var
  O, O1: TOBM;
  i: integer;
  OldGeneAuxi, NewGeneAuxi, StS: string;
  Premier, Okok: Boolean;
  TEche, TPiece: TList;
  XD, XC, XV : Double;
begin
  LaSelection.Clear;
  G.ClearSelected;
  TEche := TList.Create;
  TPiece := TList.Create;
  Premier := True;
  Okok := False;
  while not QEnc.EOF do
  begin
    //LG*
    O := TOBM.Create(EcrGen, '', True);
    Okok := True;
    O.ChargeMvtP(QEnc, 'E');
    if FSelectAll.Checked then O.PutMvt('E_ETAT', '*') else O.PutMvt('E_ETAT', ' ');
    {Sauvegarde mémoire des infos}
    O.PutMvt('E_MULTIPAIEMENT', O.GetMvt('E_LIBRETEXTE0'));
    StS := O.GetMvt('E_CONTREPARTIEGEN') + ';' + O.GetMvt('E_CONTREPARTIEAUX');
    O.PutMvt('E_TRACE', StS);
    {Maj mémoire des infos}
    O.PutMvt('E_LIBRETEXTE0', QEnc.FindField('T_LIBELLE').AsString); {ruse fichier TIERS}
    O.PutMvt('E_CONTREPARTIEGEN', QEnc.FindField('MP_GENERAL').AsString); {ruse fichier MP}
    O.PutMvt('E_CONTREPARTIEAUX', QEnc.FindField('MP_ENCAISSEMENT').AsString); {ruse fichier MP}
    NewGeneAuxi := O.GetMvt('E_GENERAL') + O.GetMvt('E_AUXILIAIRE');
    if Premier then OldGeneAuxi := NewGeneAuxi;
    Premier := False;
    if (ParTiers) and (NewGeneAuxi <> OldGeneAuxi) then
    begin
      RempliTiers(TEche, G.RowCount - 1);
      TEche := TList.Create;
      OldGeneAuxi := NewGeneAuxi;
      G.RowCount := G.RowCount + 1;
    end;
    if ParTiers then
    begin
      O.PutMvt('E_ETAT', '*');
      TEche.Add(O);
    end else
    begin
      TPiece.Add(O);
    end;
    QEnc.Next;
  end;
  {Cas du dernier tiers}
  if ParTiers then
  begin
    if Okok then //Au moins 1 tiers
    begin
      RempliTiers(TEche, G.RowCount - 1);
      G.RowCount := G.RowCount + 1;
    end;
  end;
  {Cas du traitement en détail : on traite TPieces}
  if not ParTiers then
    {Lister tous les mouvements}
  begin
    for i := 0 to TPiece.Count - 1 do
    begin
      O1 := TOBM(TPiece[i]);
      O := nil;
      EgaliseOBM(O1, O);
      {Sauvegarde des zones OBM}
      XD := O.GetMvt('E_DEBIT');
      XC := O.GetMvt('E_CREDIT');
      XV := O.GetMvt('E_COUVERTURE');
      {Manipulation des zones pour affichage contextuel}
      O.PutMvt('E_DEBIT', GetLeDebitCur(O, False));
      O.PutMvt('E_CREDIT', GetLeCreditCur(O, False));
      O.PutMvt('E_COUVERTURE', GetLaCouvCur(O));
      {Affichage}
      ComCom1(G, O);
      G.Objects[0, G.RowCount - 1] := O;
      {Restauration des zones OBM}
      O.PutMvt('E_DEBIT', XD);
      O.PutMvt('E_CREDIT', XC);
      O.PutMvt('E_COUVERTURE', XV);
      if FSelectAll.Checked then G.Cells[G.ColCount - 1, G.RowCount - 1] := '*';
      G.RowCount := G.RowCount + 1;
    end;
    VideListe(TPiece);
  end;
  if G.RowCount > 2 then G.RowCount := G.RowCount - 1;
  TPiece.Free;
end;

procedure TFEncaDeca.RempliTiers(TEche: TList; Lig: integer);
var
  i: integer;
  O: TOBM;
  Gene, Aux, Libelle: string;
  DebitCur, CreditCur, TotDebCur, TotCredCur: Double;
  CouvCur, TotCouvCur: Double;
begin
  TotDebCur := 0;
  TotCredCur := 0;
  TotCouvCur := 0;
  for i := 0 to TEche.Count - 1 do
  begin
    O := TOBM(TEche[i]);
    Gene := O.GetMvt('E_GENERAL');
    Aux := O.GetMvt('E_AUXILIAIRE');
    Libelle := O.GetMvt('E_LIBRETEXTE0');
    if O.GetMvt('E_ETAT') = '*' then
    begin
      DebitCur := GetLeDebitCur(O, False);
      CreditCur := GetLeCreditCur(O, False);
      CouvCur := GetLaCouvCur(O);
      TotDebCur := TotDebCur + DebitCur;
      TotCredCur := TotCredCur + CreditCur;
      if DebitCur <> 0 then TotCouvCur := TotCouvCur + CouvCur else TotCouvCur := TotCouvCur - CouvCur;
    end;
  end;
  if not Enc then TotCouvCur := -TotCouvCur;
  G.Cells[0, Lig] := Gene;
  G.Cells[1, Lig] := Aux;
  G.Cells[2, Lig] := Libelle;
  G.Cells[3, Lig] := FormatFloat(G.ColFormats[3], TotDebCur);
  G.Cells[4, Lig] := FormatFloat(G.ColFormats[4], TotCredCur);
  G.Cells[5, Lig] := FormatFloat(G.ColFormats[5], TotCouvCur);
  G.Objects[0, Lig] := TEche;
  if FSelectAll.Checked then G.Cells[G.ColCount - 1, Lig] := '*' else G.Cells[G.ColCount - 1, Lig] := '';
end;

{--------------------------------------------------------------------------------------------}
procedure TFEncaDeca.GetCellCanvas(Acol, ARow: LongInt; Canvas: TCanvas; AState: TGridDrawState);
begin
  if (ARow <> 0) and (EstSelect(G, ARow)) then
    G.Canvas.Font.Style := G.Canvas.Font.Style + [fsItalic] else
    G.Canvas.Font.Style := G.Canvas.Font.Style - [fsItalic];
end;

procedure TFEncaDeca.BChercherClick(Sender: TObject);
var
  StV8, StXP, StXN, StXP2, StXN2, St: string;
begin
  (*
  {$IFDEF JACADI}
  if ENC then
     BEGIN
     if cFactCredit.Checked
          then XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC" OR E_NATUREPIECE="FC"))'
          else XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC"))' ;
  //      then XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="FC"))'
  //      else XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND E_NATUREPIECE="AC")' ;
     END else
     BEGIN
     if cFactCredit.Checked
        then XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF" OR E_NATUREPIECE="FF"))'
        else XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF"))' ;
  //      then XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="FF"))'
  //      else XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND E_NATUREPIECE="AF")' ;
     END ;
  {$ELSE}
  if ENC then
     BEGIN
     if cFactCredit.Checked
        then XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="FC"))'
        else XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND E_NATUREPIECE="AC")' ;
     END else
     BEGIN
     if cFactCredit.Checked
        then XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="FF"))'
        else XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND E_NATUREPIECE="AF")' ;
     END ;
  {$ENDIF}
  *)
  if ENC then
  begin
    if cFactCredit.Checked
      then XX_WHEREENC.Text := 'E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC" OR E_NATUREPIECE="FC"))'
    else XX_WHEREENC.Text := 'E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC"))';
  end else
  begin
    if cFactCredit.Checked
      then XX_WHEREENC.Text := 'E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF" OR E_NATUREPIECE="FF"))'
    else XX_WHEREENC.Text := 'E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF"))';
  end;
  if PCircuit.TabVisible then
  begin
    if ((G1.Visible) and (E_SUIVDEC.Value = '')) then
    begin
      HMEnc.Execute(26, caption, '');
      Exit;
    end;
    if ((E_SUIVDEC.Value <> '') and (Trim(E_NOMLOT.Text) = '')) then
    begin
      HMEnc.Execute(37, caption, '');
      Exit;
    end;
  end;
  if SorteLettre in [tslCheque, tslBOR] then
  begin
    if cVoirCircuit.Checked then XX_WHERE1.Text := 'E_SUIVDEC<>""' else
    begin
      XX_WHERE1.Text := 'E_SUIVDEC=""';
      E_SUIVDEC.Value := '';
      E_NOMLOT.Text := '';
    end;
  end else
  begin
    if cVoirCircuit.Visible then
    begin
      if cVoirCircuit.Checked then XX_WHERE1.Text := '' else XX_WHERE1.Text := 'E_SUIVDEC=""';
    end else
    begin
      if ((E_SUIVDEC.Value = '') and (not Generation) and (not ENC)) then XX_WHERE1.Text := 'E_SUIVDEC=""';
    end;
  end;
  if ETAPE.Value = '' then
  begin
    HMEnc.Execute(27, caption, '');
    if ETAPE.CanFocus then ETAPE.SetFocus;
    Exit;
  end;
  if E_GENERAL.Text = '' then
  begin
    HMEnc.Execute(0, caption, '');
    if E_GENERAL.CanFocus then E_GENERAL.SetFocus;
    Exit;
  end;
  if not BonCompteEnc(ENC, E_GENERAL.Text) then
  begin
    HMEnc.Execute(8, caption, '');
    if E_GENERAL.CanFocus then E_GENERAL.SetFocus;
    Exit;
  end;
  if GUIDE.Value = '' then
  begin
    HMEnc.Execute(7, caption, '');
    if GUIDE.CanFocus then GUIDE.SetFocus;
    Exit;
  end;
  if ((MP_CATEGORIE.Value = '') and (not MP_CATEGORIE.Vide)) then
  begin
    HMEnc.Execute(16, caption, '');
    if MP_CATEGORIE.CanFocus then MP_CATEGORIE.SetFocus;
    Exit;
  end;
  EnDev := (E_DEVISE.Value <> V_PGI.DevisePivot);
  G.VidePile(True);
  ChangeTypeGrid;
  StWRib := '';
  St := '';
  QEnc.Close;
  QEnc.SQL.Clear;
  QEnc.SQL.Add('Select Ecriture.*,Tiers.T_LIBELLE,ModePaie.* from Ecriture ');
  St := St + ' Left outer join Tiers on E_AUXILIAIRE=T_AUXILIAIRE';
  St := St + ' Left outer join Modepaie on E_MODEPAIE=MP_MODEPAIE';
  St := St + ' Where ' + WhereGeneCritEcr(Pages, QEnc, False);
  StV8 := LWhereV8;
  if StV8 <> '' then St := St + ' AND ' + StV8; //QEnc.SQL.Add(' AND '+StV8) ;
  if E_DEVISE.Value <> V_PGI.DevisePivot then
  begin
    St := St + ' AND (E_ETATLETTRAGE<>"PL" OR E_LETTRAGEDEV="X")';
    StXP := StrFPoint(9 * Resolution(DEV.Decimale + 1));
    StXN := StrFPoint(-9 * Resolution(DEV.Decimale + 1));
    St := St + ' and (E_DEBITDEV+E_CREDITDEV-E_COUVERTUREDEV not between ' + StXN + ' AND ' + StXP + ')';
  end else
  begin
    St := St + ' AND ((E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL")) ';
    StXP := StrFPoint(9 * Resolution(V_PGI.OkDecV + 1));
    StXN := StrFPoint(-9 * Resolution(V_PGI.OkDecV + 1));
    StXP2 := StrFPoint(9 * Resolution(V_PGI.OkDecE + 1));
    StXN2 := StrFPoint(-9 * Resolution(V_PGI.OkDecE + 1));
    St := St + ' And (E_DEBIT+E_CREDIT-E_COUVERTURE not between ' + StXN + ' AND ' + StXP + ')';
  end;
  if TypeEtape.Value = 'POR' then
  begin
    St := St + ' AND ((E_NATUREPIECE<>"FC" AND E_NATUREPIECE<>"AC" AND E_NATUREPIECE<>"FF" AND E_NATUREPIECE<>"AF") OR (MP_GENERAL<>""))';
  end;
  if ((GereAccept.Visible) and (TypeEtape.Value = 'BQE') and (ENC) and (GereAccept.Checked)) then
  begin
    { FQ 19123 BVE 25.09.07
    St := St + ' AND (MP_CATEGORIE<>"LCR" OR MP_CODEACCEPT="ACC" OR MP_CODEACCEPT="NON" OR MP_CODEACCEPT="BOR")';}
    St := St + ' AND (MP_CATEGORIE<>"LCR" OR MP_CODEACCEPT="ACC" OR MP_CODEACCEPT="NON" OR MP_CODEACCEPT="BOR" OR E_CODEACCEPT="ACC")';
    { END FQ 19123 }
  end;
  if ((not ENC) and (TypeEtape.Value = 'BQE') and (SorteLettre = tslAucun) and (Circuit1 = '') and (Circuit2 = '')) then
  begin
    St := St + ' AND (E_BANQUEPREVI="" OR E_BANQUEPREVI="' + GeneGeneration.Text + '")';
  end;
  if St <> '' then QEnc.SQL.Add(St);
  StWRib := St;
  if ParTiers then
  begin
    QEnc.SQL.Add('order by E_GENERAL, E_AUXILIAIRE, E_EXERCICE, E_DATECOMPTABLE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ');
  end else
  begin
    case OrdreTri.ItemIndex of
      0: QEnc.SQL.Add('order by E_GENERAL, E_EXERCICE, E_DATECOMPTABLE, E_AUXILIAIRE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ');
      1: QEnc.SQL.Add('order by E_AUXILIAIRE, E_GENERAL, E_EXERCICE, E_DATECOMPTABLE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ');
      2: QEnc.SQL.Add('order by E_EXERCICE, E_DATECOMPTABLE, E_GENERAL, E_AUXILIAIRE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ');
      3: QEnc.SQL.Add('order by E_DATEECHEANCE, E_GENERAL, E_AUXILIAIRE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ');
    end;
  end;
  ChangeSQL(QEnc);
  EnableControls(Self, False);
  BZOOMETAPE.Enabled := False;
  QEnc.Open;
  EnableControls(Self, True);
  BZOOMETAPE.Enabled := True;
  RempliGrid;
  QEnc.Close;
  G.Enabled := True;
  CalculDebitCredit;
  PreVirEgal := False;
  PreVirMP := '';
  PreVirEche := 0;
end;

{======================================== ENTETE ===========================================}
procedure TFEncaDeca.E_DEVISEChange(Sender: TObject);
begin
  if E_DEVISE.Value = '' then
  begin
    FillChar(DEV, Sizeof(DEV), #0);
    Exit;
  end;
  DEV.Code := E_DEVISE.Value;
  GetInfosDevise(DEV);
  {JP 21/08/07 : FQ 21158 : on supprime l'underscore de ER_SOLDED, pour que le champ ne soit pas repris dans la requête}
  ChangeMask(ERSOLDED, DEV.Decimale, DEV.Symbole);
end;

procedure TFEncaDeca.MP_CATEGORIEChange(Sender: TObject);
begin
  CatToMP(MP_CATEGORIE.Value, E_MODEPAIE.Items, E_MODEPAIE.Values, SorteLettre, False);
  if ((E_MODEPAIE.Items.Count > 0) and (E_MODEPAIE.Vide) and (SorteLettre <> tslAucun)) then E_MODEPAIE.ItemIndex := 0;
  G.Enabled := False;
end;

procedure TFEncaDeca.PositionneFerme(Jal: string);
var
  QQ: TQuery;
begin
  if Jal = '' then Exit;
  QQ := OpenSQL('Select J_FERME From JOURNAL Where J_JOURNAL="' + Jal + '"', True);
  if not QQ.EOF then
  begin
    if QQ.Fields[0].AsString = 'X' then JalFerme := True;
  end;
  Ferme(QQ);
end;

procedure TFEncaDeca.EtudierGuide;
var
  Code: String3;
  Q: TQuery;
  O: TOBM;
  Trouv: boolean;
  Auxi1, Auxi2, Coll1, Coll2: String17;
begin
  VideListe(TGuide);
  FillChar(EntGuide, Sizeof(EntGuide), #0);
  Code := GUIDE.Value;
  JalFerme := False;
  if Code = '' then Exit;
  {Entête du guide}
  if ENC then Q := OpenSQL('Select * from GUIDE where GU_TYPE="ENC" AND GU_GUIDE="' + Code + '"', True)
  else Q := OpenSQL('Select * from GUIDE where GU_TYPE="DEC" AND GU_GUIDE="' + Code + '"', True);
  Trouv := not Q.EOF;
  if Trouv then
  begin
    EntGuide.Etabl := Q.FindField('GU_ETABLISSEMENT').AsString;
    EntGuide.Jal := Q.FindField('GU_JOURNAL').AsString;
    EntGuide.Nature := Q.FindField('GU_NATUREPIECE').AsString;
    EntGuide.CodeD := Q.FindField('GU_DEVISE').AsString;
  end;
  Ferme(Q);
  if not Trouv then Exit;
  PositionneFerme(EntGuide.Jal);
  EntGuide.DateC := DateGene;
  EntGuide.Exo := QuelExo(DATEGENERATION.Text);
  EntGuide.Simul := 'N';
  EntGuide.General := E_GENERAL.Text;
  EntGuide.LeGuide := Code;
  EntGuide.Valide := False;
  if ENC then EntGuide.TypeGuide := 'ENC' else EntGuide.TypeGuide := 'DEC';
  if ((EntGuide.CodeD <> '') and (EntGuide.CodeD <> DEV.Code)) then
  begin
    HMEnc.Execute(20, caption, '');
    Exit;
  end;
  {Lignes du guide}
  if ENC then Q := OpenSQL('Select * from ECRGUI where EG_TYPE="ENC" AND EG_GUIDE="' + Code + '" Order By EG_TYPE, EG_GUIDE, EG_NUMLIGNE', True)
  else Q := OpenSQL('Select * from ECRGUI where EG_TYPE="DEC" AND EG_GUIDE="' + Code + '" Order By EG_TYPE, EG_GUIDE, EG_NUMLIGNE', True);
  Trouv := not Q.EOF;
  if Trouv then
  begin
    while not Q.EOF do
    begin
      O := TOBM.Create(EcrGui, '', True);
      O.ChargeMvt(Q);
      TGuide.Add(O);
      Q.Next;
    end;
  end;
  Ferme(Q);
  if (VH^.PaysLocalisation<>CodeISOES) and (Not Trouv) then Exit ; //XVI 24/02/2005
  {cohérence du guide}
  if TGuide.Count <> 2 then
  begin
    HMEnc.Execute(17, caption, '');
    Exit;
  end;
  Auxi1 := TOBM(TGuide[0]).GetMvt('EG_AUXILIAIRE');
  Coll1 := TOBM(TGuide[0]).GetMvt('EG_GENERAL');
  Auxi2 := TOBM(TGuide[1]).GetMvt('EG_AUXILIAIRE');
  Coll2 := TOBM(TGuide[1]).GetMvt('EG_GENERAL');
  CollecTreso := EstCollectif(Coll1);
  CollecContreP := EstCollectif(Coll2);
  if ((TypeEtape.Value = 'BQE') and (CollecContreP)) then
  begin
    HMEnc.Execute(21, caption, '');
    Exit;
  end;
end;

procedure TFEncaDeca.MajCaption;
var
  St: string;
begin
  if Generation then
  begin
    if ENC then St := HMEnc.Mess[1] else St := HMEnc.Mess[2];
    St := St + ' - ' + Lowercase(TypeEtape.Text);
    Caption := St;
  end else
  begin
    Caption := CircuitDest.Items[CircuitDest.ItemIndex];
  end;
end;

procedure TFEncaDeca.GUIDEChange(Sender: TObject);
begin
  EtudierGuide;
end;

procedure TFEncaDeca.DATEGENERATIONExit(Sender: TObject);
var
  DD: TDateTime;
  Err: integer;
begin
  if csDestroying in ComponentState then Exit;
  if not IsValidDate(DATEGENERATION.Text) then
  begin
    Pages.ActivePage := PComptes;
    HMEnc.Execute(9, caption, '');
    DATEGENERATION.Text := DateToStr(V_PGI.DateEntree);
    DateGene := V_PGI.DateEntree;
  end else
  begin
    DD := StrToDate(DATEGENERATION.Text);
    Err := DateCorrecte(DD);
    if Err > 0 then
    begin
      Pages.ActivePage := PComptes;
      HMEnc.Execute(9 + Err, caption, '');
      DATEGENERATION.Text := DateToStr(V_PGI.DateEntree);
      DateGene := V_PGI.DateEntree;
    end else
    begin
      if RevisionActive(DD) then
      begin
        DATEGENERATION.Text := DateToStr(V_PGI.DateEntree);
        DateGene := V_PGI.DateEntree;
      end else
      begin
        DateGene := DD;
      end;
    end;
  end;
end;

{======================================= OUTILS ============================================}
procedure TFEncaDeca.BAgrandirClick(Sender: TObject);
begin
  Pages.Visible := not Pages.Visible;                    
  { BVE 20.09.07 : Changement d'ergonomie de la Fiche
  if Pages.Visible then
  begin
    BAgrandir.Glyph := BSwapA.Glyph;
    BAgrandir.Hint := BSwapA.Hint;
    BAgrandir.Caption := BSwapA.Caption;
  end else
  begin
    BAgrandir.Glyph := BSwapR.Glyph;
    BAgrandir.Hint := BSwapR.Hint;
    BAgrandir.Caption := BSwapR.Caption;
  end;                                                } 
  BAgrandir.Visible := Pages.Visible;
  BReduire.Visible  := not(Pages.Visible);
end;

procedure TFEncaDeca.FindMvtFind(Sender: TObject);
begin
  Rechercher(G, FindMvt, FindFirst);
end;

procedure TFEncaDeca.BRechercheClick(Sender: TObject);
begin
  FindFirst := True;
  FindMvt.Execute;
end;

procedure TFEncaDeca.BModePaieClick(Sender: TObject);
begin
  FicheModePaie_AGL('');
end;

procedure TFEncaDeca.POPSPopup(Sender: TObject);
begin
  InitPopUp(Self);
end;

procedure TFEncaDeca.BImprimerClick(Sender: TObject);
begin
  PrintDBGrid(G, Pages, Caption, '');
end;

procedure TFEncaDeca.E_GENERALChange(Sender: TObject);
begin
  G.Enabled := False;
end;

procedure TFEncaDeca.GENEGENERATIONChange(Sender: TObject);
begin
  G.Enabled := False;
end;

procedure TFEncaDeca.BZoomPieceClick(Sender: TObject);
var
  Lig: integer;
  O: TOBM;
  M: RMVT;
begin
  Lig := G.Row;
  if ((Lig <= 0) or (Lig > G.RowCount - 1)) then Exit;
  if ParTiers then
  begin
    {Zoom tiers}
    if G.Cells[1, Lig] <> '' then
      FicheTiers(nil, '', G.Cells[1, Lig], taConsult, 1);
  end else
  begin
    {Zoom pièce}
    if G.Cells[1, Lig] <> '' then
    begin
      O := TOBM(G.Objects[0, Lig]);
      if O = nil then Exit;
      M := OBMToIdent(O, False);
      if ((M.ModeSaisieJal <> '') and (M.ModeSaisieJal <> '-')) then LanceSaisieFolioOBM(O, taConsult)
      else LanceSaisie(nil, taConsult, M)
    end;
  end;
end;

function TFEncaDeca.ModeGenereAlimente: Boolean;
begin
  Result := TRUE;
  if ModeGenere.Visible = FALSE then Result := FALSE else
  begin
    if ModeGenere.vide and (ModeGenere.ItemIndex = 0) then Result := FALSE else
    begin
      if ModeGenere.value = '' then Result := FALSE;
    end;
  end;
end;

{================================= CONTROLES ==========================================}
function TFEncaDeca.CoherGroupe: byte;
var
  k: integer;
  OldMP, NewMP: String3;
  St, Junk, Auxi, OldAux: string;
begin
  Result := 0;
  if LaSelection.Count <= 0 then Exit;
  {Cohérence des modes de paiement}
  if ((TypeEtape.Value = 'BQE') and (GLOBAL.Checked)) then
  begin
    for k := 0 to LaSelection.Count - 1 do
    begin
      St := LaSelection[k];
      Auxi := Trim(ReadtokenSt(St));
      Junk := Trim(ReadTokenSt(St));
      Junk := ReadTokenSt(St);
      NewMP := ReadTokenSt(St);
      //       if (MODEGENERE.Value<>'')  then NewMP:=MODEGENERE.Value ;
      if ModeGenereAlimente then NewMP := MODEGENERE.Value;
      if k = 0 then
      begin
        OldMP := NewMP;
        OldAux := Auxi;
      end;
      if ((OldAux = Auxi) and (OldMP <> NewMP)) then
      begin
        Result := 2;
        Exit;
      end;
      OldMP := NewMP;
      OldAux := Auxi;
    end;
  end;
end;

function TFEncaDeca.AcqModeDate(var MP: String3; var DD: TDateTime): boolean;
var
  X: T_MONOECH;
begin
  Result := False;
  if HMEnc.Execute(24, caption, '') <> mrYes then Exit;
  X.DateEche := DD;
  X.ModePaie := MP;
  X.Cat := MP_CATEGORIE.Value;
  X.Treso := False;
  X.OkInit := False;
  X.OkVal := False;
  X.Action := taModif;
  X.DateValeur := DD;
  X.DateMvt := 0;
  if SaisirMonoEcheance(X) then
  begin
    DD := X.DateEche;
    MP := X.ModePaie;
    PreVirEche := DD;
    PreVirMP := MP;
    Result := True;
  end;
end;

function TFEncaDeca.CoherVirPre(var MP: String3; var DD: TDateTime): boolean;
var
  St, NewMP, OldMP, Junk: string;
  NewDate, OldDate: TDateTime;
  k, i: integer;
begin
  Result := False;
  OldDate := 0;
  for k := 0 to LaSelection.Count - 1 do
  begin
    St := LaSelection[k];
    for i := 1 to 3 do Junk := Trim(ReadtokenSt(St));
    NewMP := ReadTokenSt(St);
    //    if MODEGENERE.Value<>'' then NewMP:=MODEGENERE.Value ;
    if ModeGenereAlimente then NewMP := MODEGENERE.Value;
    NewDate := StrToDate(ReadTokenSt(St));
    if DateEcheance.Enabled then NewDate := StrToDate(DateEcheance.Text);
    if k = 0 then
    begin
      OldMP := NewMP;
      OldDate := NewDate;
    end;
    if ((OldMP <> NewMP) or (OldDate <> NewDate)) then
    begin
      MP := NewMP;
      DD := NewDate;
      Exit;
    end;
    OldMP := NewMP;
    OldDate := NewDate;
  end;
  Result := True;
end;

function TFEncaDeca.LesMeme: boolean;
var
  OldDate, NewDate: TDateTime;
  St, Junk: string;
  OldMP, NewMP: String3;
  Idem: boolean;
  k, i: integer;
  Auxi, OldAux: string;
begin
  Idem := True;
  Result := True;
  OldDate := 0;
  if not GLOBAL.Checked then Exit;
  for k := 0 to LaSelection.Count - 1 do
  begin
    St := LaSelection[k];
    Auxi := Trim(ReadtokenSt(St));
    for i := 1 to 2 do Junk := Trim(ReadtokenSt(St));
    NewMP := ReadTokenSt(St);
    //    if MODEGENERE.Value<>'' then NewMP:=MODEGENERE.Value ;
    if ModeGenereAlimente then NewMP := MODEGENERE.Value;
    NewDate := StrToDate(ReadTokenSt(St));
    if DateEcheance.Enabled then NewDate := StrToDate(DateEcheance.Text);
    if k = 0 then
    begin
      OldMP := NewMP;
      OldDate := NewDate;
    end;
    if (OldAux = Auxi) and ((OldMP <> NewMP) or (OldDate <> NewDate)) then
    begin
      Idem := False;
      Break;
    end;
    OldMP := NewMP;
    OldDate := NewDate;
    OldAux := Auxi;
  end;
  if not Idem then
  begin
    if HMEnc.Execute(30, caption, '') = mrYes then Idem := True;
  end;
  Result := Idem;
end;

function TFEncaDeca.VerifLesPrevi: boolean;
var
  Okok: boolean;
  Lig, j: integer;
  O: TOBM;
  TEche: TList;
  Previ: string;
begin
  Result := True;
  Okok := True;
  if not ENC then Exit;
  if SorteLettre <> tslAucun then Exit;
  for Lig := 1 to G.RowCount - 1 do if EstSelect(G, Lig) then
    begin
      if ParTiers then
      begin
        TEche := TList(G.Objects[0, Lig]);
        for j := 0 to TEche.Count - 1 do
        begin
          O := TOBM(TEche[j]);
          if O.GetMvt('E_ETAT') <> '*' then Continue;
          Previ := O.GetMvt('E_BANQUEPREVI');
          if ((Previ <> '') and (Previ <> GeneGeneration.Text)) then
          begin
            Okok := False;
            Break;
          end;
        end;
        if not Okok then Break;
      end else
      begin
        O := TOBM(G.Objects[0, Lig]);
        if O = nil then Continue;
        Previ := O.GetMvt('E_BANQUEPREVI');
        if ((Previ <> '') and (Previ <> GeneGeneration.Text)) then
        begin
          Okok := False;
          Break;
        end;
      end;
    end;
  if not Okok then
  begin
    if HMEnc.Execute(53, caption, '') = mrYes then Okok := True;
  end;
  Result := Okok;
end;

function TFEncaDeca.VerifLesEche: boolean;
var
  NewEche: TDateTime;
  Junk, St: string;
  k, i: integer;
  okok: boolean;
begin
  Okok := True;
  for k := 0 to LaSelection.Count - 1 do
  begin
    St := LaSelection[k];
    for i := 1 to 4 do Junk := Trim(ReadtokenSt(St));
    NewEche := StrToDate(ReadTokenSt(St));
    if DateEcheance.Enabled then NewEche := StrToDate(DateEcheance.Text);
    if not NbJoursOk(DATEGENE, NewEche) then
    begin
      Okok := False;
      Break;
    end;
  end;
  Result := Okok;
  if not Okok then HMEnc.Execute(31, caption, '');
end;

function TFEncaDeca.VerifLesRIB: boolean;
var
  Okok: boolean;
  Lig, j: integer;
  O: TOBM;
  TEche: TList;
begin
  Okok := True;
  Result := True;
  if SorteLettre <> tslTraite then Exit;
  for Lig := 1 to G.RowCount - 1 do if EstSelect(G, Lig) then
    begin
      if ParTiers then
      begin
        TEche := TList(G.Objects[0, Lig]);
        for j := 0 to TEche.Count - 1 do
        begin
          O := TOBM(TEche[j]);
          if O.GetMvt('E_ETAT') <> '*' then Continue;
          if Trim(O.GetMvt('E_RIB')) = '' then
          begin
            Okok := False;
            Break;
          end;
        end;
        if not Okok then Break;
      end else
      begin
        O := TOBM(G.Objects[0, Lig]);
        if O = nil then Continue;
        if Trim(O.GetMvt('E_RIB')) = '' then
        begin
          Okok := False;
          Break;
        end;
      end;
    end;
  Result := Okok;
  if not Okok then HMEnc.Execute(52, caption, '');
end;

function TFEncaDeca.VerifLesDevise: boolean;
var
  Okok, Premier: boolean;
  Taux: Double;
  Lig, j: integer;
  O: TOBM;
  TEche: TList;
begin
  Premier := True;
  Okok := True;
  Result := True;
  TauxRef := 1;
  Taux := 1;
  if DEV.Code = V_PGI.DevisePivot then Exit;
  for Lig := 1 to G.RowCount - 1 do if EstSelect(G, Lig) then
    begin
      if ParTiers then
      begin
        TEche := TList(G.Objects[0, Lig]);
        for j := 0 to TEche.Count - 1 do
        begin
          O := TOBM(TEche[j]);
          if O.GetMvt('E_ETAT') <> '*' then Continue;
          if Premier then
          begin
            Taux := O.GetMvt('E_TAUXDEV');
            Premier := False;
          end else
          begin
            if Arrondi(O.GetMvt('E_TAUXDEV') - Taux, 8) <> 0 then
            begin
              Okok := False;
              Break;
            end;
          end;
        end;
        if not Okok then Break;
      end else
      begin
        O := TOBM(G.Objects[0, Lig]);
        if O = nil then Continue;
        if Premier then
        begin
          Taux := O.GetMvt('E_TAUXDEV');
          Premier := False;
        end else
        begin
          if Arrondi(O.GetMvt('E_TAUXDEV') - Taux, 8) <> 0 then
          begin
            Okok := False;
            Break;
          end;
        end;
      end;
    end;
  Result := Okok;
  TauxRef := Taux;
  if not Okok then HMEnc.Execute(32, caption, '');
end;

function TFEncaDeca.JeValidePas: boolean;
var
  ii: integer;
  MP: String3;
  DD: TDateTime;
begin
  Result := True;
  PreVirEgal := False;
  if ETAPE.Value = '' then
  begin
    HMEnc.Execute(27, caption, '');
    Exit;
  end;
  if not G.Enabled then
  begin
    HMEnc.Execute(14, caption, '');
    Exit;
  end;
  if E_General.Text = '' then
  begin
    HMEnc.Execute(8, caption, '');
    Exit;
  end;
  if ((TypeEtape.Value <> 'POR') and (GENEGENERATION.Text = '')) then
  begin
    HMEnc.Execute(25, caption, '');
    Exit;
  end;
  if GUIDE.Value = '' then
  begin
    HMEnc.Execute(7, caption, '');
    Exit;
  end;
  if LaSelection.Count <= 0 then
  begin
    HMEnc.Execute(6, caption, '');
    Exit;
  end;
  if TypeEtape.Value = '' then
  begin
    HMEnc.Execute(19, caption, '');
    Exit;
  end;
  if ((Generation) and (JalFerme)) then
  begin
    HMEnc.Execute(55, caption, '');
    Exit;
  end;
  if RevisionActive(DateGene) then Exit;
  if not LesMeme then Exit;
  if not VerifLesEche then Exit;
  if not VerifLesDevise then Exit;
  if not VerifLesRIB then Exit;
  if not VerifLesPrevi then Exit;
  if TypeEtape.Value = 'BQE' then
  begin
    ii := CoherGroupe;
    if ii > 0 then
    begin
      HMEnc.Execute(21 + ii, caption, '');
      Exit;
    end;
    if (GLOBAL.Checked) and ((MP_CATEGORIE.Value = 'PRE') or (MP_CATEGORIE.Value = 'VIR')) then
    begin
      if not CoherVirPre(MP, DD) then
      begin
        if not AcqModeDate(MP, DD) then Exit else PreVirEgal := True;
      end;
    end;
  end;
  case SorteLettre of
    tslCheque: if HMEnc.Execute(46, caption, '') <> mrYes then Exit;
    tslBOR: if HMEnc.Execute(49, caption, '') <> mrYes then Exit;
    tslTraite: if HMEnc.Execute(50, caption, '') <> mrYes then Exit;
  else if HMEnc.Execute(5 - Ord(ENC), caption, '') <> mrYes then Exit;
  end;
  OkMultiSession := OkMultiSession and CAutoEnreg1.Checked;
  Result := False;
end;

{================================ GENERATION ==========================================}
function TFEncaDeca.MontantARegler(O: TOBM; Debit, AvecCouv: boolean; Quoi: TSorteMontant): Double;
var
  X: double;
begin
  X := 0;
  case Quoi of
    tsmPivot:
      begin
        if Debit then X := O.GetMvt('E_DEBIT') else X := O.GetMvt('E_CREDIT');
        if ((X <> 0) and (AvecCouv)) then X := X - O.GetMvt('E_COUVERTURE');
      end;
    tsmDevise:
      begin
        if Debit then X := O.GetMvt('E_DEBITDEV') else X := O.GetMvt('E_CREDITDEV');
        if ((X <> 0) and (AvecCouv)) then X := X - O.GetMvt('E_COUVERTUREDEV');
      end;
  end;
  Result := X;
end;

function TFEncaDeca.MontantGuideSt(O: TOBM; Debit: boolean): string;
begin
  if Debit then Result := FloatToStr(GetLeDebitCur(O, True))
  else Result := FloatToStr(GetLeCreditCur(O, True));
end;

procedure TFEncaDeca.LGPremLigne(Q: TQuery; O: TOBM);
var
  R: RMVT;
  T: HTStrings;
  St: string;
  VV: Variant;
begin
  Q.FindField('EG_GENERAL').AsString := O.GetMvt('E_GENERAL');
  Q.FindField('EG_AUXILIAIRE').AsString := O.GetMvt('E_AUXILIAIRE');
  Q.FindField('EG_RIB').AsString := O.GetMvt('E_RIB');
  Q.FindField('EG_REFEXTERNE').AsString := O.GetMvt('E_REFEXTERNE');
  Q.FindField('EG_BANQUEPREVI').AsString := O.GetMvt('E_BANQUEPREVI');
  Q.FindField('EG_DEBITDEV').AsString := MontantGuideSt(O, False);
  Q.FindField('EG_CREDITDEV').AsString := MontantGuideSt(O, True);
  (*
  if MODEGENERE.Value<>'' then Q.FindField('EG_MODEPAIE').AsString:=MODEGENERE.Value
                          else Q.FindField('EG_MODEPAIE').AsString:=O.GetMvt('E_MODEPAIE') ;
  *)
  if ModeGenereAlimente then Q.FindField('EG_MODEPAIE').AsString := MODEGENERE.Value
  else Q.FindField('EG_MODEPAIE').AsString := O.GetMvt('E_MODEPAIE');
  if DateEcheance.Enabled then VV := StrToDate(DateEcheance.Text) else VV := O.GetMvt('E_DATEECHEANCE');
  Q.FindField('EG_DATEECHEANCE').AsString := VarAsType(VV, VarString);
  if Q.FindField('EG_REFINTERNE').AsString = '' then Q.FindField('EG_REFINTERNE').AsString := O.GetMvt('E_REFINTERNE');
  if Q.FindField('EG_LIBELLE').AsString = '' then Q.FindField('EG_LIBELLE').AsString := O.GetMvt('E_LIBELLE');
  R := OBMToIdent(O, True);
  St := EncodeLC(R);
  T := HTStringList.Create;
  T.Add(St);
  TMemoField(Q.FindField('EG_ECHEANCES')).Assign(T);
  T.Free;
end;

procedure TFEncaDeca.LGContreCommun(Q: TQuery; OGuiContre: TOBM; Auxi, GeneMP, OldRef, OldLib, OldPrevi, OldRib, OldRefTire: string; k: integer);
begin
  if TypeEtape.Value = 'POR' then Q.FindField('EG_GENERAL').AsString := GeneMP else
  begin
    if ((NoBanque) or (not Generation) or (TypeEtape.Value <> 'BQE'))
      then Q.FindField('EG_GENERAL').AsString := OGuiContre.GetMvt('EG_GENERAL')
    else Q.FindField('EG_GENERAL').AsString := XX[k];
  end;
  if ((CollecContreP) and (NoBanque) and (TypeEtape.Value <> 'BQE'))
    then Q.FindField('EG_AUXILIAIRE').AsString := Auxi
  else Q.FindField('EG_AUXILIAIRE').AsString := '';
  if ((OldRef <> '') and (OGuiContre.GetMvt('EG_REFINTERNE') = '')) then Q.FindField('EG_REFINTERNE').AsString := OldRef;
  if ((OldLib <> '') and (OGuiContre.GetMvt('EG_LIBELLE') = '')) then Q.FindField('EG_LIBELLE').AsString := OldLib;
  if ((OldRib <> '') and (OGuiContre.GetMvt('EG_RIB') = '')) then Q.FindField('EG_RIB').AsString := OldRib;
  if ((OldrefTire <> '') and (OGuiContre.GetMvt('EG_REFEXTERNE') = '')) then Q.FindField('EG_REFEXTERNE').AsString := OldRefTire;
  if OldPrevi <> '' then Q.FindField('EG_BANQUEPREVI').AsString := OldPrevi;
end;

procedure TFEncaDeca.LGContreLigne(Q: TQuery; DCur, CCur: double; OldEche: TDateTime; OldPaie: String3);
var
  VV: Variant;
  Deci: integer;
begin
  if E_DEVISE.Value <> V_PGI.DevisePivot then Deci := DEV.Decimale else Deci := V_PGI.OkDecV;
  if Abs(DCur) > Abs(CCur) then
  begin
    DCur := DCur - CCur;
    CCur := 0;
    if ((TypeEtape.Value = 'BQE') and (DCur < 0)) then
    begin
      CCur := -DCur;
      DCur := 0;
    end;
  end else
  begin
    CCur := CCur - DCur;
    DCur := 0;
    if ((TypeEtape.Value = 'BQE') and (CCur < 0)) then
    begin
      DCur := -CCur;
      CCur := 0;
    end;
  end;
  if DCur <> 0 then Q.FindField('EG_DEBITDEV').AsString := StrS(DCur, Deci);
  if CCur <> 0 then Q.FindField('EG_CREDITDEV').AsString := StrS(CCur, Deci);
  Q.FindField('EG_MODEPAIE').AsString := OldPaie;
  VV := OldEche;
  Q.FindField('EG_DATEECHEANCE').AsString := VarAsType(VV, VarString);
end;

procedure TFEncaDeca.CreerLignesTemp(k: integer);
var
  i, Ind, Lig, Lig2, NumL: integer;
  OEcr, OGuiAux, OGuiContre: TOBM;
  Q: TQuery;
  St, Auxi, GeneMP, OldAux, OldGeneMP, OldRef, OldLib, OldPrevi, OldRefTire, OldRib: string;
  DCur, CCur: double;
  OldPaie: String3;
  OldEche: TDateTime;
  Rupt, Premier: Boolean;
  Junk, StB: string;
  TEche: TList;
  NbLigne, MaxLignes: Integer;
begin
  if (not OkDecoupe) or (not Decoupe) then IndDecoupe := 0;
  Ind := IndDecoupe;
  IndDecoupe := 0;
  MaxLignes := StrToInt(NBLIGNES.Text);
  if VH^.PaysLocalisation=CodeISOES then
  begin
     OGuiAux:=nil ;
     OGuiContre:=nil ;
     if TGuide.Count>=1 then
        OGuiAux:=TOBM(TGuide[0]) ;
     if OGuiAux=Nil then
        Exit ;
     if TGuide.Count>=2 then
        OGuiContre:=TOBM(TGuide[1]) ;
     if OGuiContre=Nil then
        Exit ;
  End else
  Begin
     OGuiAux := TOBM(TGuide[0]);
     if OGuiAux = nil then Exit;
     OGuiContre := TOBM(TGuide[1]);
     if OGuiContre = nil then Exit;
  End ;
  DCur := 0;
  CCur := 0;
  OldEche := 0;
  NbLigne := 0;
  OldRef := '';
  OldLib := '';
  OldPrevi := '';
  OldRefTire := '';
  OldRib := '';
  Q := OpenSQL('Select * from ECRGUI WHERE EG_TYPE="RrR"', False);
  NumL := 0;
  InitMove(LaSelection.Count, '');
  OldAux := '';
  OldGeneMP := '';
  Premier := True;
  //for i:=0 to LaSelection.Count-1 do
  for i := Ind to LaSelection.Count - 1 do
  begin
    MoveCur(False);
    St := LaSelection[i];
    Auxi := Trim(ReadtokenSt(St));
    GeneMP := Trim(ReadTokenSt(St));
    St := Trim(St);
    Lig := ReadTokenI(St);
    Junk := ReadTokenSt(St);
    Junk := ReadTokenSt(St);
    Lig2 := ReadTokenI(St);
    //    if ((Lig<=0) or (Lig>=G.RowCount)) then Continue ;
    if TypeEtape.Value = 'POR' then CollecContreP := EstCollectif(GeneMP);
    if ParTiers then
    begin
      TEche := TList(G.Objects[0, Lig]);
      OEcr := TOBM(TEche[Lig2 - 1]);
    end else OEcr := TOBM(G.Objects[0, Lig]);
    StB := OEcr.GetMvt('E_BANQUEPREVI');
    if StB = '' then StB := GENEGENERATION.Text;
    if ((not NoBanque) and (StB <> XX[k])) then Continue;
    { Rupture sur Tiers ou contrepartie systématique}
    if not Premier then
    begin
      Rupt := False;
      if not GLOBAL.Checked then Rupt := True else
        if ((TypeEtape.Value = 'POR') and (OldGeneMP <> GeneMP)) then Rupt := True else
        if ((TypeEtape.Value = 'BQE') or (TypeEtape.Value = 'EFF')) and (not CollecContreP) then Rupt := False else
        if Auxi <> OldAux then Rupt := True;
      if ((SorteLettre in [tslBOR, tslTraite]) and (TypeEtape.Value <> 'BQE') and (not CollecContreP)) then Rupt := True;
      if Rupt then
      begin
        Q.Insert;
        OGuiContre.EgalChamps(Q);
        Inc(NumL);
        LGCommun(Q, NumL);
        LGContreCommun(Q, OGuiContre, OldAux, OldGeneMP, OldRef, OldLib, OldPrevi, OldRib, OldRefTire, k);
        if PreVirEgal then LGContreLigne(Q, DCur, CCur, PreVirEche, PreVirMP)
        else LGContreLigne(Q, DCur, CCur, OldEche, OldPaie);
        if NoBanque then Q.FindField('EG_GUIDE').AsString := V_PGI.User
        else Q.FindField('EG_GUIDE').AsString := FormatFloat('000', k);
        Q.Post;
        Inc(NbLigne);
        if OkDecoupe and Decoupe and (NbLigne > MaxLignes) and (MaxLignes > 10) then
        begin
          IndDecoupe := i;
          Break;
        end;
        DCur := 0;
        CCur := 0;
      end;
    end;
    { Report 1 pour 1 de la première ligne }
    Q.Insert;
    OGuiAux.EgalChamps(Q);
    Inc(NumL);
    LGCommun(Q, NumL);
    LGPremLigne(Q, OEcr);
    if NoBanque then Q.FindField('EG_GUIDE').AsString := V_PGI.User
    else Q.FindField('EG_GUIDE').AsString := FormatFloat('000', k);
    Q.Post;
    Inc(NbLigne);
    DCur := DCur + GetLeDebitCur(OEcr, True);
    CCur := CCur + GetLeCreditCur(OEcr, True);
    OldAux := Auxi;
    if DateEcheance.Enabled then OldEche := StrToDate(DateEcheance.Text) else OldEche := OEcr.GetMvt('E_DATEECHEANCE');
    //    if MODEGENERE.Value<>'' then OldPaie:=MODEGENERE.Value else OldPaie:=OEcr.GetMvt('E_MODEPAIE') ;
    if ModeGenereAlimente then OldPaie := MODEGENERE.Value else OldPaie := OEcr.GetMvt('E_MODEPAIE');
    OldGeneMP := GeneMP;
    OldRefTire := OEcr.GetMvt('E_REFEXTERNE');
    OldRib := OEcr.GetMvt('E_RIB');
    if not GLOBAL.Checked then
    begin
      OldRef := OEcr.GetMvt('E_REFINTERNE');
      OldLib := OEcr.GetMvt('E_LIBELLE');
      OldPrevi := OEcr.GetMvt('E_BANQUEPREVI');
    end;
    Premier := False;
  end;
  FiniMove;
  if V_PGI.IOError <> oeOK then
  begin
    Ferme(Q);
    Exit;
  end;
  if OkDecoupe and Decoupe and (IndDecoupe > 0) then Exit;
  { Dernière rupture }
  Q.Insert;
  OGuiContre.EgalChamps(Q);
  Inc(NumL);
  LGCommun(Q, NumL);
  LGContreCommun(Q, OGuiContre, OldAux, GeneMP, OldRef, OldLib, OldPrevi, OldRib, OldRefTire, k);
  if PreVirEgal then LGContreLigne(Q, DCur, CCur, PreVirEche, PreVirMP)
  else LGContreLigne(Q, DCur, CCur, OldEche, OldPaie);
  if NoBanque then Q.FindField('EG_GUIDE').AsString := V_PGI.User
  else Q.FindField('EG_GUIDE').AsString := FormatFloat('000', k);
  Q.Post;
  Ferme(Q);
end;

function TFEncaDeca.CountBqe: integer;
var
  i, j, k: integer;
  O: TOBM;
  Nb: integer;
  Trouv: boolean;
  StB: string;
  TEche: TList;
begin
  Nb := 0;
  FillChar(XX, Sizeof(XX), #0);
  NoBanque := False;
  BanqueVide := False;
  if ((ENC) or (SorteLettre in [tslBOR, tslTraite])) then
  begin
    Result := 1;
    NoBanque := True;
    Exit;
  end;
  if ((not ENC) and (Generation) and (TypeEtape.Value <> 'BQE')) then
  begin
    Result := 1;
    NoBanque := True;
    Exit;
  end;
  if ParTiers then
  begin
    for i := 1 to G.RowCount - 1 do if EstSelect(G, i) then
      begin
        TEche := TList(G.Objects[0, i]);
        for j := 0 to TEche.Count - 1 do
        begin
          O := TOBM(TEche[j]);
          if O.GetMvt('E_ETAT') = '*' then
          begin
            StB := O.GetMvt('E_BANQUEPREVI');
            if StB = '' then
            begin
              StB := GENEGENERATION.Text;
              BanqueVide := True;
            end;
            Trouv := False;
            for k := 1 to Nb do if StB = XX[k] then
              begin
                Trouv := True;
                Break;
              end;
            if not Trouv then
            begin
              Inc(Nb);
              XX[Nb] := StB;
            end;
          end;
        end;
      end;
  end else
  begin
    for i := 1 to G.RowCount - 1 do
    begin
      if EstSelect(G, i) then
      begin
        O := TOBM(G.Objects[0, i]);
        StB := O.GetMvt('E_BANQUEPREVI');
        if StB = '' then
        begin
          StB := GENEGENERATION.Text;
          BanqueVide := True;
        end;
        Trouv := False;
        for k := 1 to Nb do if StB = XX[k] then
          begin
            Trouv := True;
            Break;
          end;
        if not Trouv then
        begin
          Inc(Nb);
          XX[Nb] := StB;
        end;
      end;
    end;
  end;
  if ((Nb = 1) and ((XX[1] = '') or (XX[1] = GENEGENERATION.Text))) then NoBanque := True;
  Result := Nb;
end;

function TFEncaDeca.JalBanque(k: integer): string;
var
  QJ: TQuery;
begin
  Result := '';
  if XX[k] = '' then Exit;
  QJ := OpenSQL('Select J_JOURNAL from JOURNAL where J_CONTREPARTIE="' + XX[k] + '" AND J_NATUREJAL="BQE"', True);
  if not QJ.EOF then Result := QJ.Fields[0].AsString;
  Ferme(QJ);
end;

procedure TFEncaDeca.CreerGuideTemp;
var
  QS, QD: TQuery;
  i, Nb, k: integer;
  JalB: string;
begin
  BanqueVide := False;
  DetruitGuideEnc(True);
  Nb := CountBqe;
  Decoupe := FALSE;
  if OkDecoupe then Decoupe := (Nb = 1) and OkMultiSession;
  if ENC then QS := OpenSQL('Select * from GUIDE where GU_TYPE="ENC" AND GU_GUIDE="' + GUIDE.Value + '"', True)
  else QS := OpenSQL('Select * from GUIDE where GU_TYPE="DEC" AND GU_GUIDE="' + GUIDE.Value + '"', True);
  if not QS.EOF then
  begin
    for k := 1 to Nb do
    begin
      QD := OpenSQL('Select * from GUIDE WHERE GU_TYPE="Rer"', False);
      QD.Insert;
      InitNew(QD);
      for i := 0 to QS.FieldCount - 1 do QD.Fields[i].Assign(QS.Fields[i]);
      QD.FindField('GU_TYPE').AsString := V_PGI.User;
      if NoBanque then QD.FindField('GU_GUIDE').AsString := V_PGI.User else
      begin
        QD.FindField('GU_GUIDE').AsString := FormatFloat('000', k);
        JalB := JalBanque(k);
        if JalB <> '' then QD.FindField('GU_JOURNAL').AsString := JalB;
      end;
      if EntGuide.CodeD = '' then EntGuide.CodeD := DEV.Code;
      QD.FindField('GU_DEVISE').AsString := EntGuide.CodeD;
      QD.Post;
      Ferme(QD);
      if V_PGI.IOError = oeOK then CreerLignesTemp(k);
    end;
  end else
  begin
    V_PGI.IOError := oeUnknown;
  end;
  Ferme(QS);
end;

procedure TFEncaDeca.ConstitueLaSelection;
{Constitution de la liste de sélection avant validation}
var
  i, j: integer;
  TEche: TList;
  O: TOBM;
begin
  LaSelection.Clear;
  if ParTiers then
  begin
    for i := 1 to G.RowCount - 1 do if EstSelect(G, i) then
      begin
        TEche := TList(G.Objects[0, i]);
        for j := 0 to TEche.Count - 1 do
        begin
          O := TOBM(TEche[j]);
          if O.GetMvt('E_ETAT') = '*' then AjouteOBM(O, i, j + 1, True);
        end;
      end;
  end else
  begin
    for i := 1 to G.RowCount - 1 do
    begin
      if EstSelect(G, i) then
      begin
        O := TOBM(G.Objects[0, i]);
        AjouteOBM(O, i, 0, True);
      end;
    end;
  end;
end;

procedure TFEncaDeca.UpdateCircuit;
var
  O: TOBM;
  StS, CPG, CPA: string;
  i, j: Integer;
  TEche: TList;
begin
  if not PCircuit.TabVisible then exit;
  if ParTiers then {masse}
  begin
    for i := 1 to G.RowCount - 1 do if EstSelect(G, i) then
      begin
        TEche := TList(G.Objects[0, i]);
        for j := 0 to TEche.Count - 1 do
        begin
          O := TOBM(TEche[j]);
          if O.GetMvt('E_ETAT') = '*' then
          begin
            {Restore infos "bricolées" en mémoire}
            O.PutMvt('E_LIBRETEXTE0', O.GetMvt('E_MULTIPAIEMENT'));
            StS := O.GetMvt('E_TRACE');
            CPG := ReadTokenSt(StS);
            CPA := ReadTokenSt(StS);
            O.PutMvt('E_CONTREPARTIEGEN', CPG);
            O.PutMvt('E_CONTREPARTIEAUX', CPA);
            O.PutMvt('E_MULTIPAIEMENT', '');
            O.PutMvt('E_TRACE', '');
            {Maj circuit}
            O.PutMvt('E_SUIVDEC', CircuitDest.Value);
            O.PutMvt('E_NOMLOT', LotDest.Text);
            O.PutMvt('E_FLAGECR', V_PGI.User);
            UpdateOBMCircuit(O);
            if V_PGI.IOError <> oeOK then break;
          end;
        end;
        if V_PGI.IOError <> oeOK then break;
      end;
  end else {détail}
  begin
    for i := 1 to G.RowCount - 1 do
    begin
      if EstSelect(G, i) then
      begin
        O := TOBM(G.Objects[0, i]);
        {Restore infos "bricolées" en mémoire}
        O.PutMvt('E_LIBRETEXTE0', O.GetMvt('E_MULTIPAIEMENT'));
        StS := O.GetMvt('E_TRACE');
        CPG := ReadTokenSt(StS);
        CPA := ReadTokenSt(StS);
        O.PutMvt('E_CONTREPARTIEGEN', CPG);
        O.PutMvt('E_CONTREPARTIEAUX', CPA);
        O.PutMvt('E_MULTIPAIEMENT', '');
        O.PutMvt('E_TRACE', '');
        {Maj circuit}
        O.PutMvt('E_SUIVDEC', CircuitDest.Value);
        O.PutMvt('E_NOMLOT', LotDest.Text);
        O.PutMvt('E_FLAGECR', V_PGI.User);
        UpdateOBMCircuit(O);
        if V_PGI.IOError <> oeOK then break;
      end;
    end;
  end;
end;

procedure TFEncaDeca.UpdateOBMCircuit(O: TOBM);
var
  St, SQL: string;
  M: RMVT;
  NumP: Longint;
begin
  St := O.StPourUpdate;
  if St = '' then Exit;
  M := OBMToIdent(O, True);
  NumP := O.GetMvt('E_NUMEROPIECE');
  if NumP > MaxNum then MaxNum := NumP;
  if NumP < MinNum then MinNum := NumP;
  SQL := 'UPDATE ECRITURE SET ' + St + ', E_DATEMODIF="' + USTime(NowFutur) + '"';
  SQL := SQL + ' Where  ' + WhereEcriture(tsGene, M, True) + ' AND E_DATEMODIF="' + USTime(O.GetMvt('E_DATEMODIF')) + '"';
  if ExecuteSQL(SQL) <> 1 then V_PGI.IOError := oeUnknown;
end;

procedure TFEncaDeca.FabricLesM;
var
  Q: TQuery;
  Code: string;
  P: P_MV;
  M: RMVT;
  k: integer;
begin
  {Init communs}
  FillChar(M, Sizeof(M), #0);
  VideListe(LesM);
  { FQ 15737 BVE 25.09.07}
  DateGene := StrToDate(DATEGENERATION.Text);
  M.DateC  := DateGene;
  { END FQ 15737 }
  M.Exo := QuelExo(DATEGENERATION.Text);
  M.Simul := 'N';
  M.Valide := False;
  M.MajDirecte := ((CAutoEnreg.Checked) and (Generation) and (not ENC) and (SorteLettre = tslAucun));
  M.MSED.MultiSessionEncaDeca := (CAutoEnreg1.Checked);
  if ENC then M.TypeGuide := 'ENC' else M.TypeGuide := 'DEC';
  M.ExportCFONB := OKCFONB;
  M.FormatCFONB := FormatCFONB;
  M.EnvoiTrans := EnvoiTrans;
  M.Bordereau := OKBordereau;
  M.Document := Document;
  M.SorteLettre := SorteLettre;
  M.Globalise := ((GLOBAL.Checked) and (SorteLettre in [tslAucun, tslCheque]));
  M.CodeD := DEV.Code;
  M.DateTaux := DateGene;
  M.TauxD := TauxRef;
  M.TIDTIC := EstCptTicTid ; // FQ 15590 SBO 04/04/20005
  {Recherche du ou des guides créés}
  Q := OpenSQL('SELECT GU_GUIDE, GU_TYPE from GUIDE Where GU_TYPE="' + V_PGI.User + '" Order By GU_TYPE, GU_GUIDE', True);
  while not Q.EOF do
  begin
    if NoBanque then
    begin
      M.LeGuide := V_PGI.User; {ruse pour contreparties}
      if TypeEtape.Value = 'BQE' then M.General := GENEGENERATION.Text else
        if TypeEtape.Value <> 'POR' then M.Section := GENEGENERATION.Text;
    end else
    begin
      Code := Q.FindField('GU_GUIDE').AsString;
      k := StrToInt(Copy(Code, 3, 1));
      M.LeGuide := Code;
      M.General := XX[k];
    end;
    P := P_MV.Create;
    P.R := M;
    LesM.Add(P);
    Q.Next;
  end;
  Ferme(Q);
end;

procedure VerifRep(LaSousDir: string);
begin
  if LaSousDir <> '' then
  begin
    if LaSousDir[Length(LaSousDir)] = '\' then LaSousDir := Copy(LaSousDir, 1, Length(LaSousDir) - 1);
    if not DirectoryExists(LaSousDir) then CreateDir(LaSousDir);
  end;
end;

procedure TFEncaDeca.RetoucheLesM(var CarePDF: Boolean);
var
  i: Integer;
begin
  if LesM = nil then Exit;
  for i := 0 to LesM.Count - 1 do
  begin
    P_MV(LesM[i]).R.MSED.Spooler := Spooler.Checked;
    if P_MV(LesM[i]).R.MSED.Spooler and (not V_PGI.QRPDF) then
    begin
      V_PGI.QRPDF := TRUE;
      CarePdf := TRUE;
    end;
    P_MV(LesM[i]).R.MSED.RepSpooler := RepSpooler.Text;
    VerifRep(P_MV(LesM[i]).R.MSED.RepSpooler);
    //  M.MSED.SoucheSpooler:=SoucheSpooler ;
    P_MV(LesM[i]).R.MSED.SoucheSpooler := 0;
    P_MV(LesM[i]).R.MSED.XFichierSpooler := XFichierSpooler.Checked;
    case SorteLettre of
      tslCheque: P_MV(LesM[i]).R.MSED.RacineSpooler := GetparamSoc('SO_CPRACINECHEQUE');
      tslTraite: P_MV(LesM[i]).R.MSED.RacineSpooler := GetparamSoc('SO_CPRACINETRAITE');
      tslBOR: P_MV(LesM[i]).R.MSED.RacineSpooler := GetparamSoc('SO_CPRACINEBOR');
    end;
  end;
end;

procedure TFEncaDeca.BValideClick(Sender: TObject);
var
  M: RMVT;
  io: TIOErr;
  s, SQL: string;
  DecoupeFaite: Boolean;
  SessionFaite: Boolean;
  ModeleMultiSession: string;
  SoucheSpooler: Integer;
  CarePDF: Boolean;
label
  0;
begin
  NowFutur := NowH;
  DecoupeFaite := FALSE;
  if (CircuitDest.Value <> '') and (not Generation) then
  begin
    if Trim(LotDest.Text) = '' then
    begin
      HMEnc.Execute(35, caption, '');
      Exit;
    end;
    if HMEnc.Execute(54, caption, ' ' + Trim(LotDest.Text) + ' ?') <> mrYes then Exit;
    if Transactions(UpdateCircuit, 3) = oeOk then
    begin
      if (cBAP.Checked) and (cModeleBAP.Value <> '') then
      begin
        s := 'E_FLAGECR="' + V_PGI.User + '" AND E_GENERAL="' + E_GENERAL.Text + '" '
          + 'AND E_NUMEROPIECE>=' + IntToStr(MinNum) + ' AND E_NUMEROPIECE<=' + IntToStr(MaxNum);
        LanceEtat('E', 'BAP', cModeleBAP.Value, True, False, False, nil, s, '', False);
      end;
      // RAZ de E_FLAGECR
      SQL := 'UPDATE ECRITURE SET E_FLAGECR="" WHERE E_GENERAL="' + E_GENERAL.Text + '" '
        + 'AND E_QUALIFPIECE="N" AND E_ECHE="X" AND E_NUMECHE>0 AND E_ETATLETTRAGE<>"TL" '
        + 'AND E_NUMEROPIECE>=' + IntToStr(MinNum) + ' AND E_NUMEROPIECE<=' + IntToStr(MaxNum) + ' '
        + 'AND E_FLAGECR="' + V_PGI.User + '"';
      ExecuteSQL(SQL);
      if not Generation then HMEnc.Execute(38, caption, '');
    end else
    begin
      MessageAlerte(HDiv.Mess[2]);
      Exit;
    end;
  end;
  CarePdf := FALSE;
  if Generation then
  begin
    { Validation normale }
    ConstitueLaSelection;
    if JeValidePas then Exit;
    Application.ProcessMessages;
    IndDecoupe := 0;
    SessionFaite := FALSE;
    ModeleMultiSession := '';
    SoucheSpooler := 0;
    CarePdf := FALSE;
    0: io := Transactions(CreerGuideTemp, 5);
    case io of
      oeOK: ;
      oeUnknown:
        begin
          MessageAlerte(HDiv.Mess[1]);
          Exit;
        end;
      oeSaisie:
        begin
          MessageAlerte(HDiv.Mess[2]);
          Exit;
        end;
    end;
    FabricLesM;
    if (not OkDecoupe) or (not Decoupe) or (IndDecoupe <= 0) then G.VidePile(True);
    if Decoupe and OkDecoupe and (IndDecoupe > 0) and OkMultiSession then DecoupeFaite := TRUE;
    if ((Generation) and (not ENC)) then
    begin
      if (SorteLettre in [tslCheque, tslTraite, tslBOR]) and Spooler.Checked then RetoucheLesM(CarePdf);
      LanceMultiSaisie(LesM);
    end else
    begin
      M := P_MV(LesM[0]).R;
      M.MajDirecte := False;
      M.MSED.MultiSessionEncaDeca := (CAutoEnreg1.Checked) and OkMultiSession and DecoupeFaite;
      M.MSED.SessionFaite := SessionFaite;
      if M.MSED.MultiSessionEncaDeca then M.MSED.ModeleMultiSession := ModeleMultiSession;
      M.MSED.Spooler := Spooler.Checked;
      if M.MSED.Spooler and (not V_PGI.QRPDF) then
      begin
        V_PGI.QRPDF := TRUE;
        CarePdf := TRUE;
      end;
      M.MSED.RepSpooler := RepSpooler.Text;
      VerifRep(M.MSED.RepSpooler);
      M.MSED.SoucheSpooler := SoucheSpooler;
      M.MSED.XFichierSpooler := XFichierSpooler.Checked;
      {$IFDEF SPEC350}
      case SorteLettre of
        tslCheque: M.MSED.RacineSpooler := 'CHQ';
        tslTraite: M.MSED.RacineSpooler := 'TRA';
        tslBOR: M.MSED.RacineSpooler := 'BOR';
      end;
      {$ELSE}
      case SorteLettre of
        tslCheque: M.MSED.RacineSpooler := GetparamSoc('SO_CPRACINECHEQUE');
        tslTraite: M.MSED.RacineSpooler := GetparamSoc('SO_CPRACINETRAITE');
        tslBOR: M.MSED.RacineSpooler := GetparamSoc('SO_CPRACINEBOR');
      end;
      {$ENDIF}
      LanceSaisie(nil, taCreat, M);
      ModeleMultiSession := M.MSED.ModeleMultiSession;
      SoucheSpooler := M.MSED.SoucheSpooler;
      SessionFaite := TRUE;
      {$IFNDEF SPEC350}
      if M.MSED.Spooler then
        case SorteLettre of
          tslCheque: SetparamSoc('SO_CPCHEMINCHEQUE', M.MSED.RepSpooler);
          tslTraite: SetparamSoc('SO_CPCHEMINTRAITE', M.MSED.RepSpooler);
          tslBOR: SetparamSoc('SO_CPCHEMINBOR', M.MSED.RepSpooler);
        end;
      {$ENDIF}
    end;
    if OkDecoupe and Decoupe and (IndDecoupe > 0) and OkMultiSession then goto 0;
  end;
  Application.ProcessMessages;
  if CarePDF then V_PGI.QRPDF := FALSE;
  BChercherClick(nil);
  G.ClearSelected;
end;

procedure TFEncaDeca.SwapCat(Vide: Boolean);
var
  ValCat, ValMode: string;
begin
  ValCat := MP_CATEGORIE.Value;
  ValMode := E_MODEPAIE.Value;
  MP_CATEGORIE.Vide := Vide;
  MP_Categorie.Enabled := True;
  MP_CATEGORIE.Reload;
  MP_CATEGORIE.Value := ValCat;
  E_MODEPAIE.Value := ValMode;
end;

procedure TFEncaDeca.PositionneBB;
begin
  if ((Circuit2 = '') and (AffectBanque)) then
  begin
    if ((not ENC) and (SorteLettre = tslBOR)) then BAffectBqe.Enabled := True else
      if ((ENC) and (SorteLettre = tslTraite)) then BAffectBqe.Enabled := True else
      BAffectBqe.Enabled := (TypeEtape.Value = 'BQE');
  end;
end;

procedure TFEncaDeca.ETAPEChange(Sender: TObject);
var
  Q: TQuery;
  StW, TypeET: string;
  Okok, OkCpte: boolean;
begin
  G.Enabled := False;
  Okok := True;
  OkCFONB := False;
  FormatCFONB := '';
  EnvoiTrans := '';
  OkBordereau := False;
  Document := '';
  JalFerme := False;
  if ETAPE.Value = '' then Exit;
  if ENC then StW := 'ER_ENCAISSEMENT="X"' else StW := 'ER_ENCAISSEMENT="-"';
  Q := OpenSQL('Select * from ETAPEREG Where ' + StW + ' AND ER_ETAPE="' + ETAPE.Value + '"', True);
  if not Q.EOF then
  begin
    E_GENERAL.Text := Q.FindField('ER_CPTEDEPART').AsString;
    TypeET := Q.FindField('ER_TYPEETAPE').AsString;
    if ((TYPEET = 'POR') or (TYPEET = 'EFF')) then
    begin
      if ENC then GENEGENERATION.ZoomTable := tzGBQCaissCli else GENEGENERATION.ZoomTable := tzGBQCaissFou;
    end else if TYPEET = 'BQE' then
    begin
      GENEGENERATION.ZoomTable := tzGBanque;
    end;
    Okok := ((Okok) and (E_GENERAL.ExisteH > 0));
    if Okok then
    begin
      GENEGENERATION.Text := Q.FindField('ER_CPTEARRIVEE').AsString;
      if (GENEGENERATION.Text <> '') then OkCpte := (GENEGENERATION.ExisteH > 0) else OkCpte := (TypeET = 'POR');
      Okok := ((Okok) and (OkCpte));
      if Okok then
      begin
        MP_CATEGORIE.Value := '';
        if ((ENC) and (SorteLettre = tslAucun)) then MP_CATEGORIE.Value := Q.FindField('ER_CATEGORIEMP').AsString else
          if ((not ENC) and (Generation) and (SorteLettre = tslAucun)) then MP_CATEGORIE.Value := Q.FindField('ER_CATEGORIEMP').AsString;
        if Generation then E_MODEPAIE.Value := Q.FindField('ER_MODEPAIE').AsString;
        E_DEVISE.Value := Q.FindField('ER_DEVISE').AsString;
        GUIDE.Reload;
        GUIDE.Value := Q.FindField('ER_GUIDE').AsString;
        GLOBAL.Checked := (Q.FindField('ER_GLOBALISE').AsString = 'X');
        OkCFONB := (Q.FindField('ER_EXPORTCFONB').AsString = 'X');
        FormatCFONB := Q.FindField('ER_FORMATCFONB').AsString;
        {$IFNDEF SPEC302}
        if VH^.OldTeleTrans then EnvoiTrans := Q.FindField('ER_ENVOITRANS').AsString;
        {$ENDIF}
        OkBordereau := (Q.FindField('ER_BORDEREAU').AsString = 'X');
        Document := Q.FindField('ER_DOCUMENT').AsString;
        CodeAFB.Text := Q.FindField('ER_CODEAFB').AsString;
        TYPEETAPE.Value := TypeET;
        GereAccept.Visible := ((ENC) and (TypeET = 'BQE'));
        PositionneBB;
      end;
    end;
  end;
  Ferme(Q);
  if not Okok then
  begin
    E_GENERAL.Text := '';
    GENEGENERATION.Text := '';
    HMEnc.Execute(29, caption, '');
    Etape.Value := '';
    Exit;
  end;
  E_GENERAL.Enabled := True;
  DateGeneration.Enabled := True;
  Global.Enabled := (SorteLettre in [tslAucun, tslCheque]);
  if TYPEETAPE.Value = 'POR' then
  begin
    GENEGENERATION.Enabled := False;
    GENEGENERATION.Text := '**********';
    H_GENEGENERATION.Caption := HDiv.Mess[0];
  end else if ETAPE.Value = 'EFF' then
  begin
    GENEGENERATION.Enabled := True;
  end else if ETAPE.Value = 'BQE' then
  begin
    GENEGENERATION.Enabled := False;
  end;
  MODEGENERE.Visible := True;
  H_MODEGENERE.Visible := True;
  MODEGENERE.ItemIndex := 0;
  if SorteLettre <> tslAucun then
  begin
    MP_CATEGORIE.Vide := True;
    MP_CATEGORIE.Reload;
    MP_CATEGORIE.Value := ''; (*MP_CATEGORIE.ItemIndex:=0 ;*)
    MP_Categorie.Enabled := False;
    //   MODEGENERE.Visible:=False ; MODEGENERE.Value:='' ; H_MODEGENERE.Visible:=False ;
  end else if OkCFONB then
  begin
    SwapCat(False);
    //   MODEGENERE.Visible:=False ; MODEGENERE.Value:='' ; H_MODEGENERE.Visible:=False ;
  end else
  begin
    SwapCat(True);
    //   MODEGENERE.Visible:=True ; H_MODEGENERE.Visible:=True ;
  end;
  if SorteLettre = tslAucun then MajCaption;
  GuideChange(nil);
end;

procedure TFEncaDeca.E_GENERALExit(Sender: TObject);
var
  St: string;
begin
  if csDestroying in ComponentState then Exit;
  St := E_GENERAL.Text;
  if St = '' then
  begin
    H_GENERAL.Caption := '';
    Exit;
  end;
  if St = GENEGENERATION.Text then
  begin
    HMEnc.Execute(28, caption, '');
    Exit;
  end;
  E_GENERAL.ExisteH;
  if GUIDE.Value = '' then Exit;
  if ENC then EXECUTESQL('UPDATE ECRGUI SET EG_GENERAL="' + St + '" WHERE EG_TYPE="ENC" AND EG_GUIDE="' + GUIDE.Value + '" AND EG_NUMLIGNE=1')
  else EXECUTESQL('UPDATE ECRGUI SET EG_GENERAL="' + St + '" WHERE EG_TYPE="DEC" AND EG_GUIDE="' + GUIDE.Value + '" AND EG_NUMLIGNE=1');
  GuideChange(nil);
end;

procedure TFEncaDeca.GENEGENERATIONExit(Sender: TObject);
var
  St: string;
begin
  if csDestroying in ComponentState then Exit;
  St := GENEGENERATION.Text;
  if St = '' then Exit;
  if St = E_GENERAL.Text then
  begin
    HMEnc.Execute(28, caption, '');
    Exit;
  end;
  if GUIDE.Value = '' then Exit;
  if TypeEtape.Value = 'POR' then Exit;
  if ENC then EXECUTESQL('UPDATE ECRGUI SET EG_GENERAL="' + St + '" WHERE EG_TYPE="ENC" AND EG_GUIDE="' + GUIDE.Value + '" AND EG_NUMLIGNE=2')
  else EXECUTESQL('UPDATE ECRGUI SET EG_GENERAL="' + St + '" WHERE EG_TYPE="DEC" AND EG_GUIDE="' + GUIDE.Value + '" AND EG_NUMLIGNE=2');
  GuideChange(nil);
end;

procedure TFEncaDeca.BZOOMETAPEClick(Sender: TObject);
var
  VV: String3;
  OldG: Boolean;
begin
  VV := Etape.Value;
  ParamEtapeReg(ENC, VV, False);
  OldG := V_PGI.GestionMessage;
  V_PGI.GestionMessage := True;
  Etape.Reload;
  Etape.Value := VV;
  V_PGI.GestionMessage := OldG;
end;

procedure TFEncaDeca.BModifEcheClick(Sender: TObject);
var
  M: RMVT;
  EU: T_ECHEUNIT;
  TAN: String3;
  Lig, k: integer;
  O: TOBM;
  Aux, Coll: string;
  TotDebCur, TotCredCur, TotCouvCur: Double;
  TEche: TList;
begin
  Lig := G.Row;
  if ((Lig <= 0) or (Lig > G.RowCount - 1)) then Exit;
  if ParTiers then
  begin {Liste des échéances du tiers}
    Aux := G.Cells[1, Lig];
    TotDebCur := 0;
    TotCredCur := 0;
    TotCouvCur := 0;
    TEche := TList(G.Objects[0, Lig]);
    if TEche = nil then exit;
    if EncaDecaTiers(Aux, TEche, TotDebCur, TotCredCur, TotCouvCur, Enc, DEV, SorteLettre <> tslAucun, false) then
    begin
      G.Cells[3, Lig] := FormatFloat(G.ColFormats[3], TotDebCur);
      G.Cells[4, Lig] := FormatFloat(G.ColFormats[4], TotCredCur);
      G.Cells[5, Lig] := FormatFloat(G.ColFormats[5], TotCouvCur);
      // Sélection dynamique du tiers
      if TotDebCur + TotCredCur <> 0 then G.Cells[G.ColCount - 1, Lig] := '+' else G.Cells[G.ColCount - 1, Lig] := ' ';
      CalculDebitCredit;
    end;
  end else
  begin
    O := TOBM(G.Objects[0, Lig]);
    if O = nil then Exit;
    if O.GetMvt('E_ANA') = 'X' then Exit;
    M := OBMToIdent(O, True);
    FillChar(EU, Sizeof(EU), #0);
    EU.DateEche := O.GetMvt('E_DATEECHEANCE');
    EU.ModePaie := O.GetMvt('E_MODEPAIE');
    EU.DebitDEV := O.GetMvt('E_DEBITDEV');
    EU.CreditDEV := O.GetMvt('E_CREDITDEV');
    EU.Debit := O.GetMvt('E_DEBIT');
    EU.Credit := O.GetMvt('E_CREDIT');
    EU.DEVISE := O.GetMvt('E_DEVISE');
    EU.TauxDEV := O.GetMvt('E_TAUXDEV');
    EU.DateComptable := O.GetMvt('E_DATECOMPTABLE');
    EU.DateModif := O.GetMvt('E_DATEMODIF');
    EU.ModeSaisie := O.GetMvt('E_MODESAISIE');
    {#TVAENC}
    if VH^.OuiTvaEnc then
    begin
      Coll := O.GetMvt('E_GENERAL');
      if EstCollFact(Coll) then
      begin
        for k := 1 to 4 do EU.TabTva[k] := O.GetMvt('E_ECHEENC' + IntToStr(k));
        EU.TabTva[5] := O.GetMvt('E_ECHEDEBIT');
      end;
    end;
    TAN := O.GetMvt('E_ECRANOUVEAU');
    if TAN = 'OAN' then
    begin
      if M.CodeD <> V_PGI.DevisePivot then Exit;
      if ((VH^.EXOV8.Code <> '') and (M.DateC < VH^.EXOV8.Deb)) then Exit;
    end;
    if ModifUneEcheance(M, EU) then BChercherClick(nil);
  end;
end;

procedure TFEncaDeca.WMGetMinMaxInfo(var MSG: Tmessage);
begin
  with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do
  begin
    X := WMinX;
    Y := WMinY;
  end;
end;

procedure TFEncaDeca.E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
begin
  ParamDate(Self, Sender, Key);
end;

procedure TFEncaDeca.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

procedure TFEncaDeca.BAnnulerClick(Sender: TObject);
begin
  Close;
  if IsInside(Self) then
  begin
    CloseInsidePanel(Self) ;
  end;
end;

procedure TFEncaDeca.cTypeVisuChange(Sender: TObject);
begin
  OrdreTri.Visible := (rDetail.Checked);
  H_OrdreTri.Visible := OrdreTri.Visible;
  BModifRIB.Enabled := ((Generation) and (OrdreTri.Visible));
end;

procedure TFEncaDeca.GRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin
end;

{------------------ Affectation prévisionnelle des banques --------------------}

procedure TFEncaDeca.AddGridAffBqe(G: THGrid; O: TOBM);
begin
  // Ajoute l'échéance dans le grid détail d'affectation
  if AffectBanque then O.PutMvt('E_NUMPIECEINTERNE', O.GetMvt('E_BANQUEPREVI'));
  if G.Cells[0, G.RowCount - 1] <> '' then G.RowCount := G.RowCount + 1;
  ComCom1(G, O);
  G.Objects[0, G.RowCount - 1] := O;
end;

procedure TFEncaDeca.DoAffectBqe;
var
  X: TFAffBqe;
  i, j, n: Integer;
  TEche: TList;
  O: TOBM;
begin
  X := TFAffBqe.Create(Application);
  if not AffectBanque then X.cVisu.Checked := True;
  X.Devise.Text := E_DEVISE.Value;
  n := 0;
  if ParTiers then
  begin
    for i := 1 to G.RowCount - 1 do if EstSelect(G, i) then
      begin
        TEche := TList(G.Objects[0, i]);
        for j := 0 to TEche.Count - 1 do
        begin
          O := TOBM(TEche[j]);
          if O.GetMvt('E_ETAT') = '*' then
          begin
            AddGridAffBqe(X.G, O);
            Inc(n);
          end;
        end;
      end;
  end else
  begin
    for i := 1 to G.RowCount - 1 do
    begin
      if EstSelect(G, i) then
      begin
        O := TOBM(G.Objects[0, i]);
        AddGridAffBqe(X.G, O);
        Inc(n);
      end;
    end;
  end;
  try
    if n > 0 then X.ShowModal;
  finally
    X.Free;
  end;
end;

procedure TFEncaDeca.bAffectBqeClick(Sender: TObject);
begin
  if not BAffectBqe.Enabled then Exit;
  DoAffectBqe;
end;

procedure TFEncaDeca.bListeLotsClick(Sender: TObject);
var
  sWhere: string;
begin
  if Trim(E_GENERAL.Text) = '' then
  begin
    HMEnc.Execute(27, caption, '');
    Exit;
  end;
  if E_SUIVDEC.Value = '' then
  begin
    HMEnc.Execute(26, caption, '');
    Exit;
  end;
  sWhere := 'E_GENERAL="' + E_GENERAL.Text + '" AND E_SUIVDEC="' + E_SUIVDEC.Value + '"';
  E_NOMLOT.Text := Choisir(HMEnc.Mess[41], 'ECRITURE', 'DISTINCT E_NOMLOT', '', sWhere, 'E_NOMLOT');
end;

procedure TFEncaDeca.cBAPClick(Sender: TObject);
begin
  cModeleBAP.Enabled := cBAP.Checked;
end;

procedure TFEncaDeca.GereOngletCircuit(Montrer: boolean);
begin
  if SorteLettre in [tslCheque, tslBOR, tslAucun] then
  begin
    PCircuit.TabVisible := cVoirCircuit.Checked;
    if CVoirCircuit.Checked then
    begin
      E_SUIVDEC.Value := Circuit1;
    end else
    begin
      E_SUIVDEC.Value := '';
      E_NOMLOT.Text := '';
    end;
  end;
  if Montrer then Pages.ActivePage := PEcritures else Pages.ActivePage := PComptes;
end;

procedure TFEncaDeca.cVoirCircuitClick(Sender: TObject);
begin
  GereOngletCircuit(True);
end;

procedure TFEncaDeca.CChoixDateClick(Sender: TObject);
begin
  if CChoixDate.Checked then
  begin
    DateEcheance.Enabled := True;
    if DateEcheance.Text = StDate1900 then DateEcheance.Text := DateToStr(V_PGI.DateEntree);
  end else
  begin
    DateEcheance.Text := StDate1900;
    DateEcheance.Enabled := False;
  end;
end;

procedure TFEncaDeca.E_EXERCICEChange(Sender: TObject);
begin
  {JP 21/08/07 : FQ 20025 : gestion du cas ou E_EXERCICE.Value = ''}
  ExoToEdDates(E_EXERCICE.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_);
end;

procedure TFEncaDeca.BModifRIBClick(Sender: TObject);
var
  Lig: integer;
  O: TOBM;
  RIB, Aux: string;
  MM: RMVT;
begin
  if not Generation then Exit;
  if ParTiers then Exit;
  Lig := G.Row;
  if ((Lig <= 0) or (Lig > G.RowCount - 1)) then Exit;
  O := TOBM(G.Objects[0, Lig]);
  if O = nil then Exit;
  Aux := O.GetMvt('E_AUXILIAIRE');
  if Aux = '' then Exit;
  RIB := O.GetMvt('E_RIB');
  if ModifLeRIB(RIB, Aux) then
  begin
    O.PutMvt('E_RIB', RIB);
    MM := OBMToIdent(O, True);
    {JP 16/11/07 : FQ 21847 : Gestion de E_UTILISATEUR}
    ExecuteSQL('UPDATE ECRITURE SET E_RIB = "' + RIB + '", E_UTILISATEUR = "' + V_PGI.User +
               '", E_DATEMODIF = "' + UsTime(NowH) + '" WHERE ' + WhereEcriture(tsGene, MM, True));
  end;
end;

procedure TFEncaDeca.BZoomModeleClick(Sender: TObject);
begin
  if SorteLettre <> tslCheque then Exit;
  if ETAPE.Value = '' then Exit;
  if GeneGeneration.Text = '' then Exit;
  FicheBanqueCP(GeneGeneration.text, taModif, 0);
end;

(*
QEnc.SQL.Add('Select Ecriture.*,Tiers.T_LIBELLE,ModePaie.* from Ecriture ') ;
QEnc.SQL.Add(' Left outer join Tiers on E_AUXILIAIRE=T_AUXILIAIRE') ;
QEnc.SQL.Add(' Left outer join Modepaie on E_MODEPAIE=MP_MODEPAIE') ;
QEnc.SQL.Add(' Where '+WhereGeneCritEcr(Pages,QEnc,False)) ;
StV8:=LWhereV8 ; if StV8<>'' then QEnc.SQL.Add(' AND '+StV8) ;
if E_DEVISE.Value<>V_PGI.DevisePivot then
   BEGIN
   QEnc.SQL.Add(' AND (E_ETATLETTRAGE<>"PL" OR E_LETTRAGEDEV="X")') ;
   StXP:=StrFPoint(9*Resolution(DEV.Decimale+1)) ; StXN:=StrFPoint(-9*Resolution(DEV.Decimale+1)) ;
   QEnc.SQL.Add(' and (E_DEBITDEV+E_CREDITDEV-E_COUVERTUREDEV not between '+StXN+' AND '+StXP+')') ;
   END else
   begin
   if CModeSaisie.Value='ECU' then
      BEGIN
      QEnc.SQL.Add(' AND ((E_ETATLETTRAGE="AL" AND E_SAISIEEURO="X") OR (E_ETATLETTRAGE="PL" AND E_LETTRAGEEURO="X"))') ;
      END else if CModeSaisie.Value=V_PGI.DevisePivot then
      BEGIN
      QEnc.SQL.Add(' AND ((E_ETATLETTRAGE="AL" AND E_SAISIEEURO="-") OR (E_ETATLETTRAGE="PL" AND E_LETTRAGEEURO="-"))') ;
      END else
      BEGIN
      QEnc.SQL.Add(' AND ((E_ETATLETTRAGE="AL") OR (E_ETATLETTRAGE="PL" AND E_LETTRAGEEURO="-"))') ;
      END ;
   StXP:=StrFPoint(9*Resolution(V_PGI.OkDecV+1))  ; StXN:=StrFPoint(-9*Resolution(V_PGI.OkDecV+1)) ;
   StXP2:=StrFPoint(9*Resolution(V_PGI.OkDecE+1)) ; StXN2:=StrFPoint(-9*Resolution(V_PGI.OkDecE+1)) ;
   QEnc.SQL.Add(' And (E_DEBIT+E_CREDIT-E_COUVERTURE not between '+StXN+' AND '+StXP+')') ;
   QEnc.SQL.Add(' And (E_DEBITEURO+E_CREDITEURO-E_COUVERTUREEURO not between '+StXN2+' AND '+StXP2+')') ;
   end ;
if TypeEtape.Value='POR' then
   begin
   QEnc.SQL.Add(' AND ((E_NATUREPIECE<>"FC" AND E_NATUREPIECE<>"AC" AND E_NATUREPIECE<>"FF" AND E_NATUREPIECE<>"AF") OR (MP_GENERAL<>""))') ;
   end ;
if ((GereAccept.Visible) and (TypeEtape.Value='BQE') and (ENC) and (GereAccept.Checked)) then
   BEGIN
   QEnc.SQL.Add(' AND (MP_CATEGORIE<>"LCR" OR MP_CODEACCEPT="ACC" OR MP_CODEACCEPT="NON" OR MP_CODEACCEPT="BOR")') ;
   END ;
if ((Not ENC) and (TypeEtape.Value='BQE') and (SorteLettre=tslAucun) and (Circuit1='') and (Circuit2='')) then
   BEGIN
   QEnc.SQL.Add(' AND (E_BANQUEPREVI="" OR E_BANQUEPREVI="'+GeneGeneration.Text+'")') ;
   END ;
if ParTiers then
   begin
   QEnc.SQL.Add('order by E_GENERAL, E_AUXILIAIRE, E_EXERCICE, E_DATECOMPTABLE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ') ;
   end else
   begin
   case OrdreTri.ItemIndex of
      0 : QEnc.SQL.Add('order by E_GENERAL, E_EXERCICE, E_DATECOMPTABLE, E_AUXILIAIRE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ') ;
      1 : QEnc.SQL.Add('order by E_AUXILIAIRE, E_GENERAL, E_EXERCICE, E_DATECOMPTABLE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ') ;
      2 : QEnc.SQL.Add('order by E_EXERCICE, E_DATECOMPTABLE, E_GENERAL, E_AUXILIAIRE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ') ;
      end ;
   end ;

*)
procedure TFEncaDeca.BCtrlRibClick(Sender: TObject);
var
  StWRib: string;
  i: Integer;
begin
  // Code non testé car BCtrlRib.visible = FALSE !
  inherited;
  StWRib := RecupWhereCritere(Pages);
  if (StWRib = '') then Exit;
  if ((not G.AllSelected) and (G.NbSelected > 0) and (G.NbSelected < 100)) then
  begin // Si on n'a pas tous sélectionné ET qu'il y a au moins 1 et 100 au plus lignes sélectionnées
    // Rajoute une clause au WHERE
    StWRib := StWRib + ' AND (';
    for i := 0 to G.NbSelected - 1 do
    begin
      G.GotoLeBookmark(i);
      StWRib := StWRib + ' (E_NUMEROPIECE=' + QEnc.FindField('E_NUMEROPIECE').AsString + ' AND E_NUMLIGNE=' + QEnc.FindField('E_NUMLIGNE').AsString + ' AND E_JOURNAL="' +
        QEnc.FindField('E_JOURNAL').AsString + '") OR';
    end;
    // Efface le dernier OR et rajoute ')'
    delete(StWRib, length(StWRib) - 2, 3);
    StWRib := StWRib + ')';
  end;
  if StWRib <> '' then CPLanceFiche_VerifRib('WHERE=' + StWRib);
end;

procedure TFEncaDeca.CAutoEnreg1Click(Sender: TObject);
begin
  if not CAutoEnreg1.Checked then NbLignes.Text := '0' else if NbLignes.Text = '0' then NbLignes.Text := '100';
  NbLignes.Enabled := CAutoEnreg1.Checked;
  LNbLignes.Enabled := CAutoEnreg1.Checked;
end;

procedure TFEncaDeca.SpoolerClick(Sender: TObject);
begin
  RepSpooler.Enabled := Spooler.Checked;
  TRepSpooler.Enabled := Spooler.Checked;
  XFichierSpooler.Enabled := Spooler.Checked;
  if Spooler.Checked and (RepSpooler.text = '') then
  begin
    {$IFDEF SPEC350}
    RepSpooler.text := ExtractFilePath(Application.EXEName)
      {$ELSE}
    case SorteLettre of
      tslCheque: RepSpooler.text := GetparamSoc('SO_CPCHEMINCHEQUE');
      tslTraite: RepSpooler.text := GetparamSoc('SO_CPCHEMINTRAITE');
      tslBOR: RepSpooler.text := GetparamSoc('SO_CPCHEMINBOR');
    end;
    {$ENDIF}
  end;
end;

procedure TFEncaDeca.BLanceSaisieClick(Sender: TObject);
begin
  // A faire
  ClickSaisiePiece;
end;

procedure TFEncaDeca.BLanceSaisieBorClick(Sender: TObject);
begin
  // A faire
  ClickSaisieBor;
end;

procedure TFEncaDeca.ClickSaisieBor;
begin
  SaisieFolio(taModif);
end;

procedure TFEncaDeca.ClickSaisiePiece;
begin
  MultiCritereMvt(taCreat, 'N', False);
end;

function TFEncaDeca.EstCptTicTid: Boolean; // FQ 15590 SBO 04/04/20005
begin
  result := False ;
  if E_GENERAL.Text = '' then Exit ;
  Result := ExisteSQL( 'SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = "' + E_GENERAL.Text + '"'
                       + ' AND G_NATUREGENE IN ("TIC", "TID")' ) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 18/09/2007
Modifié le ... :   /  /    
Description .. : FQ 21450
Mots clefs ... : 
*****************************************************************}
procedure TFEncaDeca.bSelectAllClick(Sender: TObject);
var
  i : integer ;
begin
  for i:=1 to G.RowCount-1 do
  begin
     if G.AllSelected and EstSelect(G,i) then
        CocheDecoche(i,False)
     else if not(G.AllSelected) and not(EstSelect(G,i)) then
        CocheDecoche(i,False) ;
  end ;
  G.AllSelected := not G.AllSelected;
  G.Invalidate ;
end;

end.

