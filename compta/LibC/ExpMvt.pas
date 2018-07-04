unit ExpMvt;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, Hctrls, Mask, hmsgbox, DB, Ent1, SaisComm,
  HStatus, HEnt1, Hcompte, HDB, Menus, ComCtrls,ParamDat, EDI, (*ImpUtil,*)RapSuppr,Paramsoc,
  HSysMenu,CritEdt,CPTEUTIL,UTILEDT, HTB97, ed_tools, HPanel, UiUtil,RappType,ImpFicU,TImpFic,UObjFiltres,
  ADODB, udbxDataset {SG6 10/11/04 Gestion des Filtres FQ 18826} ;

type tcritexpmvt = Record
                   Format,Exo,Etab,Jal1,Jal2,TypeEcr : String ;
                   Date1,Date2 : tdatetime ;
                   NumP1,NumP2 : Integer ;
                   NomFic : String ;
                   OkExport : Boolean ;
                   AvecODA : Boolean ;
                   End ;

Procedure ExportMvt(Lequel : String) ;
procedure LanceExportMvtExt(Lequel : String) ;
Procedure ExternalExportMvt(Lequel : String ; Var CritExpMvt : TCritExpMvt) ;

type TFNumPiece=Class
     Min,Max : Integer ;
     END ;

Type TMonoLigne = Record
                  mMontantD,mMontantE,mMontantF,mPourcent,mPourcentQte1,mPourcentQte2 : Double ;
                  mAxe : String ;
                  PasDeGeneSurCptVentilable : Boolean ;
                  ODAnal : Boolean ;
                  Tva,TPF,TvaEnc,CtrGen,CtrAux : String ;
                  End ;

type
  TFExpMvt = class(TForm)
    MsgBox: THMsgBox;
    Sauve: TSaveDialog;
    PListe: TPanel;
    EJOURNALAUTORISE: THLabel;
    PFenGuide: TPanel;
    Panel2: TPanel;
    BCValide: TToolbarButton97;
    BCAide: TToolbarButton97;
    BCAbandon: TToolbarButton97;
    Bdetag: TToolbarButton97;
    BTag: TToolbarButton97;
    JALNONEXPORTE: TListBox;
    Pages: TPageControl;
    PCritere: TTabSheet;
    PGene: TPanel;
    HLabel6: THLabel;
    Label7: THLabel;
    HLabel7: THLabel;
    HLabel5: THLabel;
    FFormat: THValComboBox;
    EXERCICE: THValComboBox;
    PFichier: TPanel;
    HLabel1: THLabel;
    HLabel2: THLabel;
    FTypeEcr: THValComboBox;
    FileName: TEdit;
    FExport: TCheckBox;
    RechFile: TToolbarButton97;
    Query: TQuery;
    FDateEcr1: THCritMaskEdit;
    FDateEcr2: THCritMaskEdit;
    PBalance: TTabSheet;
    PJournal: TGroupBox;
    HLabel10: THLabel;
    HLabel8: THLabel;
    HMTrad: THSystemMenu;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    PAvance: TTabSheet;
    FEtab: THValComboBox;
    HLabel11: THLabel;
    Status: THStatusBar;
    Dock971: TDock97;
    PFiltres: TToolWindow97;
    BFiltre: TToolbarButton97;
    FFiltres: THValComboBox;
    PTools: TToolWindow97;
    BEDI: TToolbarButton97;
    BFormat: TToolbarButton97;
    BZoom: TToolbarButton97;
    BStop: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    PCB: TPanel;
    CTri: TCheckBox;
    CBInfosComptes: TCheckBox;
    CBDisting: TCheckBox;
    CBOdaCegid: TCheckBox;
    AvecRevision: TCheckBox;
    Lettre: TCheckBox;
    PNature: TPanel;
    ENature: THLabel;
    EModePaie: THLabel;
    EEche1: THLabel;
    EEche2: THLabel;
    FNatureAux: THValComboBox;
    FModePaie: THValComboBox;
    FEche1: THCritMaskEdit;
    FEche2: THCritMaskEdit;
    PCompte: TPanel;
    HLabel4: THLabel;
    EDEPART: THLabel;
    EARRIVEE: THLabel;
    HLabel9: THLabel;
    GNature: THLabel;
    CPTEDEBUT: THCpteEdit;
    CPTEFIN: THCpteEdit;
    FNatureGene: THValComboBox;
    ENATUREGENE: THLabel;
    FNATJAL1: THValComboBox;
    Label1: THLabel;
    FNATJAL2: THValComboBox;
    Label3: THLabel;
    SaufJal: TEdit;
    ZoomJAL: TToolbarButton97;
    HLabel3: THLabel;
    Numpiece1: THNumEdit;
    Label2: THLabel;
    Numpiece2: THNumEdit;
    JALAN: THCpteEdit;
    JALN: THCpteEdit;
    LJALAN: THLabel;
    LJALN: THLabel;
    PDivers: TGroupBox;
    Label4: THLabel;
    RefInterne: TEdit;
    Label5: THLabel;
    Libelle: TEdit;
    TSelectCpte: THLabel;
    FSelectCpte: THValComboBox;
    FCollectif: TCheckBox;
    PSISCO: TTabSheet;
    GSISCO: TGroupBox;
    LCompteResultatSISCO: THLabel;
    ECompteResultatSISCO: THLabel;
    FCompteResultatSISCO: THCpteEdit;
    HLabel12: THLabel;
    FBen: THCpteEdit;
    HLabel13: THLabel;
    HLabel14: THLabel;
    FPerte: THCpteEdit;
    HLabel15: THLabel;
    LNumPlanSISCO: THLabel;
    FNumPlanSISCO: TEdit;
    FNumDossierSISCO: TEdit;
    LNumDossierSISCO: THLabel;
    TSCENARIOEXPORT: THLabel;
    SCENARIOEXPORT: THCritMaskEdit;
    procedure RechFileClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FFormatChange(Sender: TObject);
    procedure ZoomJALClick(Sender: TObject);
    procedure BCValideClick(Sender: TObject);
    procedure BCAbandonClick(Sender: TObject);
    procedure BTagClick(Sender: TObject);
    procedure BdetagClick(Sender: TObject);
    procedure BFormatClick(Sender: TObject);
    procedure BEDIClick(Sender: TObject);
    procedure PFenGuideMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure PFenGuideMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure BStopClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure EXERCICEClick(Sender: TObject);
    procedure FDateEcr1KeyPress(Sender: TObject; var Key: Char);
    procedure FDateEcr2KeyPress(Sender: TObject; var Key: Char);
    procedure BZoomClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FNatureGeneChange(Sender: TObject);
    procedure SCENARIOEXPORTKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SCENARIOEXPORTChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    ObjFiltre:TObjFiltre; //SG6 10/11/04 Gestion des filtres FQ 14826
    Composants : TControlFiltre; //SG6   Gestion des Filtes 10/11/04   FQ 14826
    Prefixe : String3 ;
    Arreter: boolean ;
    NbMvt : LongInt ;
    TJAl : TStringList ;
    NbPiece,OldNumP: Integer ;
    OldJal,FirstJal,OldNat,OldQualP : String ;
    OldDateC :TDateTime ;
    TotCred, TotDeb : Double ;
    StTotCred, StTotDeb : String ;
    Fichier : TextFile ;
    CodeJournal   : String3 ;
    DateDeb       : TDateTime;
    DateFin       : TDateTime;
    QtiteJournal  : integer ;
    DPTotJournal,DNTotJournal,
    CPTotJournal,CNTotJournal  : Double ;
    QtiteJournaux : integer ;
    DPTotJournaux,DNTotJournaux,
    CPTotJournaux,CNTotJournaux : Double ;
    QtiteExtract  : integer ;
    DPTotExtract,DNTotExtract,
    CPTotExtract,CNTotExtract  : Double ;
    FiltreExp : String ;
    QtiteComptes : integer ;
    CompteGene : String17 ;
    NbEcr         : LongInt ;
    QBalC : TQuery ;
    DPTotCompte,DNTotCompte,CPTotCompte,CNTotCompte     : Double ;
    SFDCompte,SFCCompte,SNDCompte,SNCCompte             : Double ;
    DPTotComptes,DNTotComptes,CPTotComptes,CNTotComptes : Double ;
    SFDExtract,SFCExtract,SNDExtract,SNCExtract : Double ;
    SourisX,SourisY : Integer ;
    SectAttente : Array[1..5] of String ;
    TDev :  TList ;
    GCalc : TGCalculCum ;
    PourVisu : boolean ;
    ListeNumJal : TStringList ;
    ListeCptCegid : TStringList ;
    SemiAuto : Boolean ;
    Scenario : TScenario ;
    LoadingFiltre : Boolean ;
    NatExport : tNatImpExp ;
    ExternalExp : Boolean ;
    CritExpMvt : TCritExpMvt ;
    Procedure RajouteCompteCegid ;
    procedure LanceExport ;
    procedure ExporteMvt ;
    procedure ExporteMvtOdaCegid(AppendFichier : Boolean) ;
    procedure ExporteLesAnals(Var BlocML : TMonoLigne ; LJalB : tStringlist ; Q1 : TQuery) ;
    procedure ExporteLesEcritures(ListeMvtL,ListeJalB : TStringList ; Q1 : TQuery) ;
    procedure ExporteLesEcrituresODAna(Q1 : TQuery) ;
    function  EnRupture(QEcr : TQuery) : boolean ;
    function  TestBreak(QEcr : TQuery) : boolean ;
    function  Fmt_Date_Saari(Dat : String) : String ;
    function  ChercheCategPaie(MPaie : String3) : String3;
    function  CHERCHEUNMODE ( CategP : String3 ; Mode : String3 ; OkTreso,OkV7 : boolean ) : Char ;
    procedure InitComptes ;
    Function  StWhereBase(Pref,NomChampJal,NomChampQualif : String)  : String ;
    Function  FabriqueStWhere(What : Integer ; Pref : String ; OkPourLettrage,ODAnalCegid : Boolean) : String ;
    Function  FabriqueStWhereAna(Q1 : TQuery) : String ;
    procedure ConstruitRequete(Pref : String3 ; OkUpdate : boolean ; ODAnalCegid : Boolean) ;
    Function ReqSQL(Pref : String3 ; OkUpdate : boolean ; ODAnalCegid : Boolean ; Q1 : TQuery) : String ;
    procedure EcrireLigne(Q : TQuery ; P : Char ; var St : String ; Var BlocML : TMonoLigne ; LJalB : tStringList) ;
    procedure TagDetag(ok : boolean) ;
    procedure MajSelect ;
    procedure RempliJal ;
    function  VerifCodeJal : boolean ;
    procedure AjouteZonesHal(Q : TQuery ; var St : String ; P : Char) ;
    procedure LanceExportEcrEDI ;
    procedure EnteteMvt ;
    procedure Journal(Q1 : tQuery) ;
    procedure ExtractionGene ;
    procedure DetailMvt ;
    procedure ResumeMvt ;
    procedure ExporteBal ;
    procedure ExporteBalSISCO ;
    procedure LanceExportBalEDI ;
    procedure EnteteBal ;
    procedure DetailBal ;
    procedure CalcTotCpte ;
    procedure ResumeBal ;
    procedure CalculSoldes ;
    procedure RecupExo(var Exo : TExoDate ; var Date1,Date2,Date11,Date21,DateDeb,DateFin : TDateTime) ;
    function  WhereSQL : String;
    procedure VisiblePourBalance ;
    procedure ExporteLettre(ListeMvtL : TStringList) ;
    procedure ChargeDevEtSect ;
    function  FQuelAN(Inclure : Boolean ; Date1,Date11 : TDateTime) : TQuelAN ;
    function  AjouteInfosCompte(Q : TQuery ; P : String) : String ;
    procedure MajListeNumJal ;
    Function  OkEcrireLigne (QEcr : TQuery ; P : Char) : Boolean ;
    procedure ZoneVisible(BalSISCO : Boolean);
    procedure RecupExternalCrit ;
    Procedure SiscoDefaut ;
    Function  CritSISCOOk(Var CritEdt : TCritEdt) : Integer ;
  public
  Lequel : String ;
  end;

implementation

uses Filtre,FmtChoix,SaisUtil,LettUtil,VisuExp,ImporFmt, QRBalGen, SISCO, tCalcCum, ScenacoE, UtilTrans ;

{$R *.DFM}

Function FirstLigneFic(Format,Lequel : String) : String ;
Var St : String ;
{$IFDEF SPEC302}
    QSoc : TQuery ;
{$ENDIF}
BEGIN
St:=Format_String(' ',70) ;
{$IFDEF SPEC302}
QSoc:=OpenSQL('Select SO_LIBELLE From SOCIETE',True) ;
if not QSOC.Eof then BEGIN QSOC.First ; St:=Format_String(QSOC.Fields[0].AsString,70) ; END ;
Ferme(QSOC) ;
{$ELSE}
St:=GetParamSoc('SO_LIBELLE') ; St:=Format_String(St,30) ;
{$ENDIF}
//St:=St+Format+Lequel ;
Result:=St ;
END ;

Procedure ExportMvt(Lequel : String) ;
var FExpMvt:TFExpMvt ;
    PP : THPanel ;
BEGIN
FExpMvt:=TFExpMvt.Create(Application) ;
FExpMvt.Lequel:=Lequel ;

Fillchar(FExpMVt.CritExpMvt,SizeOf(FExpMVt.CritExpMvt),#0) ;
FExpMVt.ExternalExp:=FALSE ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   try
     FExpMVt.SemiAuto:=FALSE ;
     FExpMVt.ShowModal ;
     finally
     FExpMvt.Free ;
     END ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(FExpMvt,PP) ;
   FExpMvt.Show ;
   END ;
END ;

procedure LanceExportMvtExt(Lequel : String) ;
var FExpMvt:TFExpMvt ;
BEGIN
FExpMvt:=TFExpMvt.Create(Application) ;
FExpMvt.Lequel:=Lequel ;
Fillchar(FExpMVt.CritExpMvt,SizeOf(FExpMVt.CritExpMvt),#0) ;
FExpMVt.ExternalExp:=FALSE ;
try
  FExpMVt.SemiAuto:=TRUE ;
  FExpMVt.ShowModal ;
  finally
  FExpMvt.Free ;
  END ;
SourisNormale ;
END ;

Procedure ExternalExportMvt(Lequel : String ; Var CritExpMvt : TCritExpMvt) ;
var FExpMvt:TFExpMvt ;
    PP : THPanel ;
BEGIN
FExpMvt:=TFExpMvt.Create(Application) ;
FExpMvt.Lequel:=Lequel ;
FExpMVt.ExternalExp:=TRUE ;
FExpMVt.CritExpMvt:=CritExpMvt ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   try
     FExpMVt.SemiAuto:=FALSE ;
     FExpMVt.ShowModal ;
     finally
     FExpMvt.Free ;
     END ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(FExpMvt,PP) ;
   FExpMvt.Show ;
   END ;
END ;

procedure TFExpMvt.RecupExternalCrit ;
BEGIN
With Critexpmvt Do
  BEGIN
  FFormat.Value:=Format ;
  Exercice.Value:=Exo ;
  FDateEcr1.Text:=DateToStr(Date1) ;
  FDateEcr2.Text:=DateToStr(Date2) ;
  If Jal1<>'' Then FNatJal1.Value:=Jal1 ;
  If Jal2<>'' Then FNatJal2.Value:=Jal2 ;
  NumPiece1.Value:=NumP1 ;
  NumPiece2.Value:=NumP2 ;
  FTypeEcr.Value:=TypeEcr ;
  FileName.Text:=NomFic ;
  If AvecODA Then CBOdaCEGID.State:=cbGrayed ;
  END ;
END ;

Procedure TFExpMvt.SiscoDefaut ;
{$IFDEF SPEC302}
Var QSoc : TQuery ;
{$ENDIF}
BEGIN
If (Lequel<>'FBA') Or (FFormat.Value<>'SIS') then Exit ;
{$IFDEF SPEC302}
QSoc:=OpenSQL('Select SO_OUVREPERTE, SO_OUVREBEN, FROM SOCIETE WHERE SO_SOCIETE="'+V_PGI.CodeSociete+'"',TRUE) ;
If Not QSoc.Eof Then
  BEGIN
  If Fperte.Text='' Then FPerte.text:=QSoc.Fields[0].AsString ;
  If FBen.Text=''   Then FBen.text:=QSoc.Fields[1].AsString ; END ;
Ferme(QSoc) ;
{$ELSE}
If Fperte.Text='' Then FPerte.text:=GetParamSoc('SO_OUVREPERTE')   ;
If FBen.Text='' Then FBen.text:=GetParamSoc('SO_OUVREBEN')   ;
{$ENDIF}
END ;

procedure TFExpMvt.FormShow(Sender: TObject);
var Index,Titre,Reste : string ;
begin
LoadingFiltre:=FALSE ; LJalN.Caption:='' ; LJalAN.Caption:='' ;
NatExport:=nPiece ;
Fillchar(Scenario,SizeOf(Scenario),#0) ;
If SemiAuto Then
  BEGIN
  VStatus:=Status ;
  Status.Caption := Copyright ;
  END Else Status.Visible:=FALSE ;
FEche1.text:=StDate1900 ; FEche2.text:=StDate2099 ;
FiltreExp:='EXPMVT' ;
Prefixe:='E' ;
If Lequel='FEC' then
  BEGIN
  FiltreExp:='EXPMVT' ;
  Caption:=MsgBox.Mess[15] ; //exportation des écritures
  HelpContext:=6341000 ;
  END else
If Lequel='FOD' then
  BEGIN
  Prefixe:='Y' ;
  FiltreExp:='EXPMVTOD' ;
  FNatJal1.DataType:='ttJalAnalytique' ;
  FNatJal2.DataType:='ttJalAnalytique' ;
  Caption:=MsgBox.Mess[34] ;   //Exportation des écritures analytiques
  HelpContext:=6341000 ;
  END else
If Lequel='FBE' then
  BEGIN
  Prefixe:='BE' ;
  FiltreExp:='EXPMVTBU' ;
  EXERCICE.DataType:='ttExerciceBudget' ;
  FNatJal1.DataType:='ttBudJal' ;
  FNatJal2.DataType:='ttBudJal' ;
  FTypeEcr.DataType:='ttNatEcrBud' ;
  Caption:=MsgBox.Mess[35] ; //Exportation des écritures budgétaires
  HelpContext:=6341000 ;
  END else
If Lequel='FBA' then
  BEGIN
  FiltreExp:='EXPBAL';
  CPTEDEBUT.ZoomTable:=tzGeneral ;
  CPTEFIN.ZoomTable:=tzGeneral ;
  InitComptes ;
  Caption:=MsgBox.Mess[16] ;  //Exportation de la balance générale
  HelpContext:=6350000 ;
  PCritere.HelpContext:=6351000 ;
  PBalance.HelpContext:=6352000 ;
//  FDateEcr1.Enabled:=False ; FDateEcr2.Enabled:=False ;
  FSelectCpte.ItemIndex:=0 ; FSelectCpte.Enabled:=False ;
  END ;
//astuce pour avoir le bon titre de message en fonction du type d'export
Reste:=MsgBox.Mess[1] ; Index:=ReadTokenSt(Reste) ;Titre:=ReadTokenSt(Reste) ;
MsgBox.Mess[1]:=Index+';'+caption+';'+Reste;

FNatJal1.ItemIndex:=0 ;
FNatJal2.ItemIndex:=FNatJal2.Items.Count-1 ;
FTypeEcr.ItemIndex:=1 ;
Exercice.ItemIndex:=0 ;ExerciceClick(Nil) ;

ChangeDataTypeFmt('-',Lequel,FFormat) ;
VisiblePourBalance ;
GereFMTChoix('-',Lequel,FFormat,FileName,RechFile) ;
if FFormat.ItemIndex=-1 then FFormatChange(nil) ;
RempliJal ;

//SG6 10/11/04 Gestion des filtres FQ 14826
ObjFiltre.FFI_TABLE:=FiltreExp;
ObjFiltre.Charger;

ChargeDevEtSect ;
Lettre.Visible:=(Lequel='FEC') ;
(*
CBInfosComptes.Visible:=V_PGI.SAV and ((FFormat.Value='HLI') or (FFormat.Value='HAL') or (FFormat.Value='SAA') or (FFormat.Value='SN2')) ;
CBInfosComptes.Visible:=(FFormat.Value='CGN') Or FFormat.Value='CGE') ;
*)
CBInfosComptes.Enabled:=(FFormat.Value='CGN') Or (FFormat.Value='CGE') ;
If Not CBInfosComptes.Enabled Then CBInfosComptes.Checked:=FALSE ;
CBODACegid.Visible:=(FFormat.Value='CGE') Or (FFormat.Value='CGN') ;
If Not CBODACegid.Visible Then CBDisting.Checked:=FALSE ;
ListeNumJal:=TStringList.Create ;
Pages.ActivePage:=Pages.Pages[0] ;
UpdateCaption(Self) ;
If ExternalExp Then
  BEGIN
  RecupExternalCrit ; BValider.ModalResult:=mrOk ; BValiderClick(Nil) ; PostMessage(Handle,WM_CLOSE,0,0) ; Exit ;
  END ;
SiscoDefaut ;
ScenarioExport.Visible:=FALSE ; TScenarioExport.Visible:=FALSE ;
//ScenarioExport.Visible:=(Lequel='FEC') ; TScenarioExport.Visible:=(Lequel='FEC') ;
If Not ScenarioExport.Visible Then ScenarioExport.Text:='' ;
end;


procedure TFExpMvt.BFormatClick(Sender: TObject);
begin
ChoixFormatImpExp('-' ,Lequel) ;
If not (ChoixFmt.OkSauve) then Exit ;
With ChoixFmt do
  BEGIN
  if Format<>'' then FFormat.Value:=Format ;
  FFormat.Enabled:=not (FixeFmt) ;
  if Fichier<>'' then FileName.Text:=Fichier ;
  FileName.Enabled:=not (FixeFichier) ;
  RechFile.Enabled:=not (FixeFichier) ;
  END ;
END ;

procedure TFExpMvt.RechFileClick(Sender: TObject);
begin
DirDefault(Sauve,FileName.Text) ;
if Sauve.Execute then FileName.Text:=Sauve.FileName ;
end;

procedure TFExpMvt.BFermeClick(Sender: TObject);
begin
Close;
//SG6 13/01/05 FQ 15242
if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

procedure TFExpMvt.EXERCICEClick(Sender: TObject);
begin
If LoadingFiltre Then Exit ;
ExoToDates(EXERCICE.Value,FDATEECR1,FDATEECR2) ;
end;

procedure TFExpMvt.FDateEcr1KeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

procedure TFExpMvt.FDateEcr2KeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

function Fmt_Date_HAL(Dat : String) : String ;
var D : TDateTime ;
BEGIN
D:=StrToDate(Dat) ; Result:=FormatDateTime('ddmmyyyy',D) ;
END ;

function TFExpMvt.Fmt_Date_Saari(Dat : String) : String ;
var D : TDateTime ;
BEGIN
D:=StrToDate(Dat) ; Result:=FormatDateTime('ddmmyy',D) ;
END ;

{================ Mode de paiement SAARI ====================}

function TFExpMvt.ChercheCategPaie(MPaie : String3) : String3 ;
var Q : TQuery ;
BEGIN
Result:='' ;
Q:=OpenSQL('Select MP_CATEGORIE FROM MODEPAIE WHERE MP_MODEPAIE="'+MPaie+'"',True) ;
if Not Q.Eof then Result:=Q.Fields[0].AsString ;
Ferme(Q) ;
END ;

FUNCTION TFExpMvt.CHERCHEUNMODE ( CategP : String3; Mode : String3 ; OkTreso,OkV7 : boolean ) : Char ;
BEGIN
// Catégories/Modes de paiements connus;
if CategP='ESP' then ChercheUnMode:='O' else
if CategP='CHQ' then ChercheUnMode:='C' else
if CategP='CB' then BEGIN if ((not OkTreso) and (Not OkV7)) then ChercheUnMode:='C' else ChercheUnMode:='U' END else
if CategP='VIR' then ChercheUnMode:='V' else
if CategP='PRE' then ChercheUnMode:='P' else
// Mode paramétré suivant type LCR ...mais quoi d'autre ???
if Mode='TRD' then ChercheUnMode:='T' else
if Mode='TRA' then BEGIN if OkV7 then ChercheUnMode:='A' else ChercheUnMode:='T' END else
if Mode='LCR' then ChercheUnMode:='L' else
if Mode='BOR' then ChercheUnMode:='B' else
// Autres type de LCR...
if CategP='LCR' then ChercheUnMode:='T' else ChercheUnMode:='S' ;
END ;

procedure EnabledPanel(GB1 : TPanel ; OkOk : Boolean) ;
Var l   : Integer ;
    CC1 : TControl ;
BEGIN
For l:=0 To GB1.ControlCount-1 Do
  BEGIN
  CC1:=GB1.Controls[l] ; CC1.Enabled:=OkOk ;
  If Not OkOk Then
    BEGIN
    If CC1 is THvalComboBox Then THvalComboBox(CC1).ItemIndex:=0 ;
    If CC1 is THCpteEdit Then THCpteEdit(CC1).Text:='' ;
    If CC1 is THCritMaskEdit Then
      BEGIN
      If Pos('1',CC1.Name)>0 Then THCritMaskEdit(CC1).Text:=StDate1900 Else THCritMaskEdit(CC1).Text:=StDate2099;
      END ;
    END ;
  END ;
GB1.enabled:=OkOk ;
END ;

procedure TFExpMvt.FFormatChange(Sender: TObject);
Var {H,}i : Integer ;
    CodeScenario : String ;
begin
//BZoom.Enabled:=(FFormat.Value<>'RAP') and (FFormat.Value<>'MP') ;

CBInfosComptes.Enabled:=(FFormat.Value='CGN') Or (FFormat.Value='CGE') ;
If Not CBInfosComptes.Enabled Then CBInfosComptes.Checked:=FALSE ;
if Pages.ActivePage=PBalance then Pages.ActivePage:=PCritere ;
Pages.Pages[3].TabVisible:=(Lequel='FBA') And (FFormat.Value='SIS');
Pages.Pages[2].TabVisible:=(((FFormat.Value='SAA') or (FFormat.Value='HAL')) and (Lequel='FBA')) ;
Pages.Pages[1].TabVisible:=(Lequel<>'FBA') ;
(*
H:=28+PGene.Height+PFichier.Height+PFiltres.Height+PTools.Height ;
If SemiAuto Then H:=H+Status.Height ;
*)
CodeScenario:='' ; If ScenarioExport.Text<>'' Then CodeScenario:=ScenarioExport.Text ;
Fillchar(Scenario,SizeOf(Scenario),#0) ;
If Lequel='FEC' Then ChargeScenario('-',FFormat.Value,Lequel,CodeScenario,Scenario,FALSE) ;
//PCompte.Visible:=FALSE ; PNature.Visible:=FALSE ;
EnabledPanel(PCompte,FALSE) ; EnabledPanel(PNature,FALSE) ;
If Scenario.EstCharge Then
  BEGIN
  If Scenario.Nature='PIE' Then NatExport:=nPiece Else If Scenario.Nature='VIR' Then NatExport:=nVirement Else
  If Scenario.Nature='PRE' Then NatExport:=nPrelevement Else If Scenario.Nature='LCR' Then NatExport:=nLcr Else
  If Scenario.Nature='RAP' Then NatExport:=nRappro Else If Scenario.Nature='TRE' Then NatExport:=nTreso ;
  If Scenario.FiltreGen Then
    BEGIN
    If Not LoadingFiltre Then
      BEGIN
      If Scenario.NatGen<>'' Then FNatureGene.Value:=Scenario.NatGen Else FNatureGene.ItemIndex:=0 ;
      END ;
    InitComptes ;
    EnabledPanel(PCompte,TRUE) ;
    (* PCompte.Top:=PFichier.top ; PCompte.Visible:=True ; H:=H+PCompte.Height ; *)
    END ;
  If Scenario.FiltreAux Then
    BEGIN
    If Not LoadingFiltre Then
      BEGIN
      If Scenario.NatAux<>'' Then FNatureAux.Value:=Scenario.NatAux Else FNatureAux.ItemIndex:=0 ;
      END ;
    EnabledPanel(PNature,TRUE) ;
    (*
    PNature.Top:=PFichier.top ; PNature.Visible:=True ; H:=H+PNature.Height ;
    FNatureAux.ItemIndex:=0 ; FModePaie.ItemIndex:=0 ;
    *)
    END ;
(*
  if (FFormat.Value='RAP') or (FFormat.Value='CRA') then //or (Lequel='FBA') then
    BEGIN
    FNatJal1.DataType:='ttJalBanque' ;
    FNatJal2.DataType:='ttJalBanque' ;
    PCompte.Top:=PFichier.top ;
    PCompte.Visible:=True ;
    InitComptes ; H:=H+PCompte.Height ;
    END else
  if (FFormat.value='MP') then
    BEGIN
    FNatJal1.DataType:='ttJournal' ;
    FNatJal2.DataType:='ttJournal' ;
    PNature.Top:=PFichier.top ; PNature.Visible:=True ; H:=H+PNature.Height ;
    FNature.ItemIndex:=0 ; FModePaie.ItemIndex:=0 ;
    END ;
*)
  END Else
  BEGIN
  NatExport:=nPiece ;
  END ;
if (Lequel='FBA') then
  BEGIN
  EXERCICE.Enabled:=not (FFormat.Value='CGF') ; FTypeEcr.Enabled:=not (FFormat.Value='CGF') ;
  for i:=0 to PGene.ControlCount-1 do
    if PGene.Controls[i].Tag=1 then TControl(PGene.Controls[i]).Enabled:=not (FFormat.Value='CGF') ;
  END ;
(*
Pages.Height:=H-(PFiltres.Height+PTools.Height) ;
ClientHeight:=H ;
*)
BEDI.Enabled:=(FFormat.Value='EDI') ;
FNatJal1.ItemIndex:=0 ;
FNatJal2.ItemIndex:=FNatJal2.Items.Count-1 ;
if FFormat.Value='EDI' then Sauve.FilterIndex:=4 else
  if FFormat.Value='HAL' then Sauve.FilterIndex:=1 else
    if (FFormat.Value='CGN') Or (FFormat.Value='CGE') then Sauve.FilterIndex:=1 else
      if (FFormat.Value='SAA') or (FFormat.Value='SN2') or (FFormat.Value='RAP')
         or (FFormat.Value='MP') then Sauve.FilterIndex:=3 ;
CBDisting.enabled:=(FFormat.Value='CGE') Or (FFormat.Value='CGN') ;
CBDisting.Checked:=CBDisting.enabled ;
CBODACegid.Visible:=(FFormat.Value='CGE') Or (FFormat.Value='CGN') ;
If Not CBODACegid.Visible Then CBDisting.Checked:=FALSE ;
If Scenario.EstCharge And (NatExport=nRappro) Then
  BEGIN
  Lettre.Enabled:=TRUE ;
  Lettre.Caption:=MsgBox.Mess[39] ;
  Lettre.AllowGrayed:=True ;
  Lettre.State:=cbUnchecked ;
  END Else
  BEGIN
  Lettre.State:=cbUnchecked ;
  END ;
BZoom.Enabled:=NatExport=nPiece ;
ZoneVisible((FFormat.Value='SIS') And (Lequel='FBA')) ;
end;

Procedure TFExpMvt.InitComptes ;
var Q : TQuery ;
    St,StWhere : String ;

BEGIN
If LoadingFiltre Then Exit ;
St:='Select G_GENERAL FROM GENERAUX' ;
StWhere:=RecupWhere(CpteDebut.ZoomTable) ;
If StWhere<>'' Then StWhere:='WHERE '+Copy(StWhere,5,Length(StWhere)-4) ;
//If Lequel<>'FBA' then St:=St+' WHERE G_POINTABLE="X"' ;
St:=St+' '+StWhere+' ORDER BY G_GENERAL' ;
Q:=OpenSql(St,True) ;
If Not Q.EOF then
  BEGIN
  Q.First ; CPTEDEBUT.Text:=Q.Fields[0].AsString ;
  Q.Last  ; CPTEFIN.Text:=Q.Fields[0].AsString ;
  END ;
Ferme(Q) ;
CPTEDEBUT.ExisteH ;CPTEFIN.ExisteH  ;
END ;

{================ Arrêt de l'exportation ====================}

procedure TFExpMvt.BStopClick(Sender: TObject);
begin Arreter:=True ; end;

procedure TFExpMvt.MajListeNumJal ;
var FNumPiece : TFNumPiece ;
BEGIN
if OldJal='' then Exit ;
if ListeNumJal.IndexOf(OldJal)=-1 then
  BEGIN
  FNumPiece:=TFNumPiece.Create ;
  FNumPiece.Min:=OldNumP ;
  FNumPiece.Max:=OldNumP ;
  ListeNumJal.AddObject(OldJal,FNumPiece) ;
  END else
  BEGIN
  FNumPiece:=TFNumPiece(ListeNumJal.Objects[ListeNumJal.IndexOf(OldJal)]) ;
  if FNumPiece.Min>OldNumP then FNumPiece.Min:=OldNumP ;
  if FNumPiece.Max<OldNumP then FNumPiece.Max:=OldNumP ;
  ListeNumJal.Objects[ListeNumJal.IndexOf(OldJal)]:=FNumPiece ;
  END ;
END ;

Function TFExpMvt.EnRupture(QEcr : TQuery) : Boolean ;
var Ok,First : Boolean ;
BEGIN
Result:=False ;
Ok:=False ;
First:=(OldJal='') ;
if Prefixe='BE' then
  BEGIN
  If (QEcr.FindField(Prefixe+'_BUDJAL').AsString<>OldJal) then
    BEGIN
    OldJal:=QEcr.Findfield(Prefixe+'_BUDJAL').AsString ;
    if FirstJal='' then FirstJal:=OldJal ;
    END ;
  if (QEcr.FindField(Prefixe+'_NUMEROPIECE').AsInteger<>OldNumP) then
    BEGIN
    Ok:=True ; OldNumP:=QEcr.Findfield(Prefixe+'_NUMEROPIECE').AsInteger ;
    END ;
  END else
  BEGIN
  If (QEcr.FindField(Prefixe+'_JOURNAL').AsString<>OldJal) then
    BEGIN
    Ok:=True ; OldJal:=QEcr.Findfield(Prefixe+'_JOURNAL').AsString ;
    if FirstJal='' then FirstJal:=OldJal ;
    END ;
  if (QEcr.FindField(Prefixe+'_NATUREPIECE').AsString<>OldNat) then
    BEGIN
    Ok:=True ; OldNat:=QEcr.Findfield(Prefixe+'_NATUREPIECE').AsString ;
    END ;
  if (QEcr.FindField(Prefixe+'_DATECOMPTABLE').AsDateTime<>OldDateC) then
    BEGIN
    Ok:=True ; OldDateC:=QEcr.Findfield(Prefixe+'_DATECOMPTABLE').AsDateTime ;
    END ;
  if (QEcr.FindField(Prefixe+'_NUMEROPIECE').AsInteger<>OldNumP) then
    BEGIN
    Ok:=True ; OldNumP:=QEcr.Findfield(Prefixe+'_NUMEROPIECE').AsInteger ;
    END ;
  if (QEcr.FindField(Prefixe+'_QUALIFPIECE').AsString<>OldQualP) then BEGIN Ok:=True ; OldQualP:=QEcr.Findfield(Prefixe+'_QUALIFPIECE').AsString ; END ;
  END ;
if Ok then
  BEGIN
  if not First then Inc(NbPiece) ;
  Result:=True ;
  END ;
if not Arreter then MajListeNumJal ;
END;

function TFExpMvt.TestBreak(QEcr : TQuery) : Boolean ;
BEGIN
Application.ProcessMessages ;
Result:=False ;
if EnRupture(QEcr) then
  BEGIN
  if Arreter then if MsgBox.Execute(21,caption,'')<>mryes then Arreter:=False ;
  END else Exit ;
Result:=Arreter ;
END ;

{================ Journaux à exclure ====================}

procedure TFExpMvt.ZoomJALClick(Sender: TObject);
//var H : Integer ;
begin
if not VerifCodeJal then Exit ;
MajSelect ;
ActivePanels(Self,False,True) ;
PListe.Left:=(PGene.Width-PListe.Width) div 2 ;
//H:=PGene.Height+PFichier.Height ;
//if PCompte.Visible then H:=H+PCompte.Height ;
//if PNature.Visible then H:=H+PNature.Height ;
//PListe.Top:=(H-PListe.Height) div 2 ;
PListe.Top:=FEtab.Top ;
BTag.Visible:=True ; BDeTag.Visible:=False ;
end;

procedure TFExpMvt.BCValideClick(Sender: TObject);
var i : integer ;
    St : String ;
    Tous : Boolean ;
begin
  inherited;
BCAbandonClick(Nil) ;
St:='' ;
Tous:=True ;
for i:=0 to JALNONEXPORTE.Items.Count-1 do
  BEGIN
  if JALNONEXPORTE.Selected[i] then St:=St+TJAL.Strings[i]+';' else Tous:=False ;
  END ;
if not Tous then SAUFJAL.Text:=St else SAUFJAL.Text:='' ;
end;

procedure TFExpMvt.BCAbandonClick(Sender: TObject);
begin
ActivePanels(Self,True,False) ;
SAUFJAL.SetFocus ;
end;

procedure TFExpMvt.BTagClick(Sender: TObject);
begin TagDetag(TRUE) ; end;

procedure TFExpMvt.BdetagClick(Sender: TObject);
begin TagDetag(FALSE) ; end;

procedure TFExpMvt.TagDetag (Ok : Boolean) ;
var i : integer ;
BEGIN
BTag.Visible:=not Ok ;
BDeTag.Visible:=Ok ;
for i:=0 to JALNONEXPORTE.Items.Count - 1 do
  if JALNONEXPORTE.Selected[i]=not Ok then JALNONEXPORTE.Selected[i]:=Ok ;
END ;

procedure  TFExpMvt.MajSelect ;
var ListeJal : String;
    Code : String ;
BEGIN
ListeJal:=SAUFJAL.Text ;
While Length(ListeJal)<>0 do
  BEGIN
  Code:=ReadTokenSt(ListeJal) ;
  if TJAL.IndexOf(Code) <>-1 then JALNONEXPORTE.Selected[TJAL.IndexOf(Code)]:=True ;
  END
END ;

procedure TFExpMvt.RempliJAL ;
var i : integer ;
BEGIN
TJAL:=TStringList.Create ;
For i:=0 to FNatJal1.Values.Count-1 do
  BEGIN
  TJAL.Add(FNatJal1.Values[i]) ;
  JALNONEXPORTE.Items.Add(FNatJal1.Items[i]+' ('+FNatJal1.Values[i]+')') ;
  END ;
MajSelect ;
END ;

Function TFExpMvt.VerifCodeJal : boolean ;
Var Code,St : string ;
BEGIN
Result:=False ;
St:=SAUFJAL.Text ;
While length(St)<>0 do
  BEGIN
  Code:=ReadTokenSt(St) ;
  if TJAL.IndexOf(Code)=-1 then BEGIN MsgBox.execute(10,caption,'') ; SAUFJAL.SetFocus ; Exit ; END ;
  END;
Result:=True ;
END ;

procedure TFExpMvt.ChargeDevEtSect ;
var Q       : TQuery ;
    TDevise : TFDevise ;
    i : integer ;
BEGIN
i:=0 ;
Q:=OpenSQL('SELECT X_SECTIONATTENTE FROM AXE ORDER BY X_AXE',True) ;
While not Q.Eof do BEGIN Inc(i) ; SectAttente[i]:=Q.Fields[0].AsString ; Q.Next ; END ;
Ferme(Q) ;
Q:=OpenSQL('SELECT D_DEVISE,D_DECIMALE,D_QUOTITE FROM DEVISE ORDER BY D_DEVISE',True) ;
TDev:=TList.Create ;
While not Q.Eof do
  BEGIN
  TDevise:=TFDevise.Create ;
  TDevise.Code:=Q.Fields[0].AsString ;
  TDevise.Decimale:=Q.Fields[1].AsInteger ;
  TDevise.Quotite:=Q.Fields[2].AsFloat ;
  TDev.Add(TDevise) ;
  Q.Next ;
  END ;
Ferme(Q) ;
END ;

{================ Lancement de l'exportation ====================}

procedure TFExpMvt.BValiderClick(Sender: TObject);
begin
try
  //Transactions(LanceExport,1) ;
  BeginTrans ;
  LanceExport ;
  CommitTrans ;
  finally
  ActivePanels(Self,True,False) ;
  END ;
end;


Procedure InitRequete(Var Q : TQuery ; What : Integer) ;
Var St : String ;
BEGIN
Q:=Nil ;
Case What Of
  0 : St:='SELECT * FROM GENERAUX WHERE G_GENERAL=:CPTE' ;
  1 : St:='SELECT * FROM TIERS    WHERE T_AUXILIAIRE=:CPTE' ;
  2 : St:='SELECT * FROM SECTION  WHERE S_SECTION=:CPTE AND S_AXE=:AXE' ;
  END ;
Q:=PrepareSQL(St,TRUE) ;
END ;

Procedure TFExpMvt.RajouteCompteCegid ;
Var St1,Cpt,Axe : String ;
    i,j : Integer ;
    Qui : String ;
    Q : Array[0..2] Of TQuery ;
    NewNom : String ;
    FichierCpt,FichierMvt : TextFile ;
BEGIN
If CBInfosComptes.State=cbUnchecked Then Exit ;
If (FFormat.Value<>'CGN') And (FFormat.Value<>'CGE') Then Exit ;
If ListeCptCegid.Count<=0 Then Exit ;
NewNom:=NewNomFic(FileName.Text,'Cpt') ;
AssignFile(FichierCpt,NewNom) ;
{$I-} Rewrite(FichierCpt) ;{$I+}
St1:=FirstLigneFic(FFormat.Value,Lequel) ;
Writeln(FichierCpt,St1) ;
For i:=0 To 2 Do InitRequete(Q[i],i) ;
InitMove(ListeCptCegid.Count,'') ;
For i:=0 To ListeCptCegid.Count-1 Do
  BEGIN
  St1:='' ;
  MoveCur(FALSE) ;
  Qui:=Copy(ListeCptCegid[i],1,1) ; Cpt:=Copy(ListeCptCegid[i],2,Length(ListeCptCegid[i])-1) ;
  Axe:=Cpt ; Cpt:=ReadTokenSt(Axe) ;
  j:=0 ; Case Qui[1] Of 'G' : j:=0 ; 'X' : j:=1 ; 'S' : j:=2 ; END ;
  Q[j].Close ;   Q[j].Params[0].AsString:=Cpt ;
  If j=2 Then BEGIN If Axe='' Then Axe:='A1' ; Q[j].Params[1].AsString:=Axe ; END ;
  Q[j].Open ;
  If Not Q[j].Eof Then BEGIN St1:=FaitStCPTCEGID(Qui[1],Q[j]) ; If St1<>'' Then Writeln(FichierCpt,St1) ; END ;
  END ;
FiniMove ;
For i:=0 To 2 Do Ferme(Q[i]) ;
If CBInfosComptes.State=cbGrayed Then
  BEGIN
  AssignFile(FichierMvt,FileName.text) ;
  InitMove(1000,'') ;
  Reset(FichierMvt) ;
  If Not Eof(FichierMvt) Then Readln(FichierMvt,St1) ;
  While Not Eof(FichierMvt) Do
    BEGIN
    MoveCur(FALSE) ; ReadLn(FichierMvt,St1) ; Writeln(FichierCpt,St1) ;
    END ;
  CloseFile(FichierMvt) ;
  FiniMove ;
  END ;
CloseFile(FichierCpt) ;
{$i-} AssignFile(FichierMvt,FileName.Text) ; Erase(FichierMvt) ; {$i+}
renamefile(NewNom,FileName.Text) ;
END ;

procedure TFExpMvt.LanceExport ;
var RapportBal : String ;
    Mss,Num : integer ;
    OldStOLEErr : String ;
    QueLEsOdaCegid : Boolean ;
    StSQL : String ;
BEGIN
Arreter:=False ; nbEcr:=0 ;
if not (IsValidDate(FDateEcr1.Text) and IsValidDate(FDateEcr2.Text) and IsValidDate(FEche1.Text) and IsValidDate(FEche2.Text)) then BEGIN Msgbox.Execute(9,caption,'') ; exit ; END ;
if FileName.Text='' then BEGIN Pages.ActivePage:=PCritere ; MsgBox.Execute(2,caption,'') ; Exit ; END ;
if FFormat.ItemIndex=-1 then BEGIN Pages.ActivePage:=PCritere ; MsgBox.Execute(14,caption,'') ; Exit ; END ;
if (Lequel='FBA') and (FFormat.Value='CGF') then
  BEGIN
  OldStOLEErr:=V_PGI.StOLEErr ;
  V_PGI.StOLEErr:=FileName.Text ;
  BalanceGeneCERG ;
  V_PGI.StOLEErr:=OldStOLEErr ;
  Exit ;
  END ;
Application.ProcessMessages ;
AssignFile(Fichier,FileName.Text) ;
{$I-} ReWrite (Fichier) ; {$I+}
if IoResult<>0 then BEGIN Msgbox.Execute(3,caption,'') ; Exit ; END ;
CloseFile(Fichier) ;
ActivePanels(Self,False,False) ;
SourisSablier ; QueLesODACegid:=FALSE ;
ListeCptCegid.Clear ;
ListeCptCegid.Sorted:=TRUE ;
ListeCptCegid.Duplicates:=DupIgnore ;
If (Lequel<>'FBA') then
  BEGIN
  if FFormat.Value='EDI' then LanceExportEcrEDI else
    BEGIN
    Case CBOdaCegid.State Of
      cbUnchecked : ExporteMvt ;
      CbChecked   : BEGIN Prefixe:='Y' ; ExporteMvtODACegid(FALSE) ; QueLesODACegid:=TRUE ; END ;
      cbGrayed    : BEGIN
                    ExporteMvt ; Prefixe:='Y' ; ExporteMvtODACegid(TRUE) ; Prefixe:='E' ;
                    END ;
      End ;
    RajouteCompteCegid ;
    END ;
  if (NbPiece>0) And (Not ExternalExp) then //(not Arreter) and
    BEGIN
    Mss:=36 ;
    if (NbPiece=1) then Mss:=37 ;
    if ChoixFmt.CompteRendu then
      BEGIN
      MsgBox.Execute(1,IntToStr(NbPiece)+' ',MsgBox.Mess[Mss]+Chr(13)+Chr(13)+MsgBox.Mess[7]+StTotDeb+'.'+Chr(13)+MsgBox.Mess[8]+StTotCred+'.') ;
//      if (FFormat.Value<>'RAP') and (FFormat.Value<>'MP') and (FFormat.Value<>'CRA') then
      If NatExport=nPiece Then
        BEGIN
        PourVisu:=True ; StSQL:=ReqSQL(Prefixe,False,QueLesODACegid,NIL) ; PourVisu:=False ;
//        QEcr.Close ;
        If ((FFormat.Value='CGN') Or (FFormat.Value='CGE')) And (CbInfosComptes.State=cbChecked) Then Else ListeExportes(StSQL,Prefixe) ;
        END ;
      END else MsgBox.Execute(6,caption,'') ;
    END ;
  END  else
  BEGIN
  if FFormat.Value='EDI' then LanceExportBalEDI else If FFormat.Value='SIS' Then ExporteBalSISCO Else ExporteBal ;
  if (NbEcr>0) then //(not Arreter)
    BEGIN
    if (ChoixFmt.CompteRendu) then
      BEGIN
      Num:=0 ;
      if FFormat.Value<>'EDI' then Num:=6 ;
      RapportBal:=Chr(13)+Chr(13)+Msgbox.Mess[22+Num]+StrfMontant(Abs(SNDExtract),20,V_PGI.OkDecV,V_PGI.SymbolePivot,False)+Chr(13)+
                  MsgBox.Mess[23+Num]+StrfMontant(Abs(SNCExtract),20,V_PGI.OkDecV,V_PGI.SymbolePivot,False)+Chr(13)+Chr(13)+
                  MsgBox.Mess[24+Num]+StrfMontant(Abs(DPTotComptes-CNtotComptes),20,V_PGI.OkDecV,V_PGI.SymbolePivot,False)+Chr(13)+
                  MsgBox.Mess[25+Num]+StrfMontant(Abs(CPTotComptes-DNtotComptes),20,V_PGI.OkDecV,V_PGI.SymbolePivot,False)+Chr(13)+Chr(13)+
                  MsgBox.Mess[26]+StrfMontant(Abs(SFDExtract),20,V_PGI.OkDecV,V_PGI.SymbolePivot,False)+Chr(13)+
                  MsgBox.Mess[27]+StrfMontant(Abs(SFCExtract),20,V_PGI.OkDecV,V_PGI.SymbolePivot,False) ;
      MessageDlg(IntToStr(NbEcr)+' '+MsgBox.Mess[32]+RapportBal, mtInformation, [mbOk], 0) ;
      END else MsgBox.Execute(17,caption,'') ;
    END ;
  END ;
ActivePanels(Self,True,False) ;
SourisNormale ;
END ;

{================ Requête pour l'exportation ====================}
(*
Procedure TFExpMvt.EcrireCompteCEGID ;
BEGIN

END ;
*)

Function TFExpMvt.StWhereBase(Pref,NomChampJal,NomChampQualif : String)  : String ;
Var StWhere : String ;
    NumP1,NumP2,StSaufJal : String ;
BEGIN
NumP1:=StrfPoint(Valeur(NUMPIECE1.Text)) ;
NumP2:=StrfPoint(Valeur(NUMPIECE2.Text)) ;
StWhere:=' WHERE '+Pref+'_DATECOMPTABLE>="'+USDate(FDateEcr1)+'" AND '+Pref+'_DATECOMPTABLE<="'+USDate(FDateEcr2) +'"' ;
if Exercice.ItemIndex>=0 then StWhere:=StWhere+' AND '+Pref+'_EXERCICE="'+EXERCICE.Value+'"' ;
if FEtab.ItemIndex>0 then StWhere:=StWhere+' AND '+Pref+'_ETABLISSEMENT="'+FEtab.Value+'"' ;
StWhere:=StWhere+' AND '+Fourchette(Pref+'_NUMEROPIECE',NumP1,NumP2,ftInteger) ;
StWhere:=StWhere+' AND '+Fourchette(Pref+'_'+NomChampJal,FNatJal1.Value,FNatJal2.Value,ftString) ;
if SAUFJAL.Text<>'' then
  BEGIN
  StSaufJal:=SaufJal.Text ;
  while StSaufJal<>'' do StWhere:=StWhere+' AND ' +Pref+'_'+NomChampJal+'<>"'+ReadTokenSt(StSaufJal)+'"' ;
  END ;
if FTypeEcr.ItemIndex>0 then StWhere:=StWhere+' AND '+Pref+'_'+NomChampQualif+'="'+FTypeEcr.Value+'"' ;
Result:=StWhere ;
END ;

Function TFExpMvt.FabriqueStWhere(What : Integer ; Pref : String ; OkPourLettrage,ODAnalCegid : Boolean) : String ;
Var Jal,StWhere : String ;
BEGIN
StWhere:='' ;
Jal:='JOURNAL' ;
if (Pref='BE') then Jal:='BUDJAL' ;
Case What Of
  0 : BEGIN // Recherche Ecriture Gene ou OD analytiues CEGID
      StWhere:=StWhereBase(Pref,'JOURNAL','QUALIFPIECE') ;
      if OkPourLettrage then StWhere:=StWhere+' AND '+Pref+'_LETTRAGE<>""' ;
      If ODAnalCegid Then StWhere:=StWhere+' AND '+Pref+'_TYPEANALYTIQUE="X" ' ;
      END ;
  1 : BEGIN // Recherche Ecriture Budgétaire
      StWhere:=StWhereBase(Pref,'BUDJAL','NATUREBUD') ;
      END ;
  2 : BEGIN // Recherche Ecriture Analytique couplées à une écriture géné.
      StWhere:=' WHERE '+Pref+'_JOURNAL=:JOURNAL AND '+Pref+'_EXERCICE=:EXERCICE' ;
      StWhere:=StWhere+' AND '+Pref+'_DATECOMPTABLE=:DATECOMPTABLE AND '+Pref+'_NATUREPIECE=:NATUREPIECE'
                      +' AND '+Pref+'_NUMEROPIECE=:NUMEROPIECE AND '+Pref+'_QUALIFPIECE=:QUALIFPIECE AND '+Pref+'_NUMLIGNE=:NUMLIGNE' ;
      END ;
  3 : BEGIN // Recherche Ecriture Gene ou OD analytiques CEGID
      StWhere:=StWhereBase(Pref,'JOURNAL','QUALIFPIECE') ;
      StWhere:=StWhere+' AND '+Pref+'_TYPEANALYTIQUE="X" ' ;
      END ;
  END ;
Result:=StWhere ;
END ;

Function TFExpMvt.FabriqueStWhereAna(Q1 : TQuery) : String ;
Var StWhere : String ;
BEGIN
StWhere:=' WHERE Y_JOURNAL="'+Q1.FindField('E_JOURNAL').AsString+'" AND Y_EXERCICE="'+Q1.FindField('E_EXERCICE').AsString+'" ' ;
StWhere:=StWhere+' AND Y_DATECOMPTABLE="'+USDateTime(Q1.FindField('E_DATECOMPTABLE').AsDateTime)+'" AND Y_NATUREPIECE="'+Q1.FindField('E_NATUREPIECE').AsString+'" '
                +' AND Y_NUMEROPIECE='+IntToStr(Q1.FindField('E_NUMEROPIECE').AsInteger)+' AND Y_QUALIFPIECE="'+Q1.FindField('E_QUALIFPIECE').AsString+'" '
                +' AND Y_NUMLIGNE='+IntToStr(Q1.FindField('E_NUMLIGNE').AsInteger)+' ' ;
Result:=StWhere ;
END ;

procedure TFExpMvt.ConstruitRequete(Pref : String3 ; OkUpdate : boolean ; ODAnalCegid : Boolean) ;
var StTable,Jal,St,StP,StWhere,StOrder : String ;
    OkPourLettrage : boolean ;
    FNumPiece : TFNumPiece ;
    i : integer ;
BEGIN
OkPourLettrage:=False ;
Jal:='JOURNAL' ; if (Pref='BE') then Jal:='BUDJAL' ;
if Pref='L' then BEGIN OkPourLettrage:=True ; Pref:='E' ; END ;
StTable:=PrefixeToTable(Pref) ;
If OkUpdate Then St:='UPDATE '+StTable+' SET '+Pref+'_EXPORTE="X" ' Else
  BEGIN
  if ((FFormat.Value='MP') and (FNatureAux.Value<>'')) then St:='SELECT *, '+GetChampsTiers(ctExpMvt)+' FROM '+StTable
                                                    else St:='SELECT * FROM '+StTable ;
  END ;
if (Pref<>'Y') Or ODAnalCegid then
  BEGIN
  If Pref='BE' Then StWhere:=FabriqueStWhere(1,Pref,OkPourLettrage,ODAnalCegid)
               Else StWhere:=FabriqueStWhere(0,Pref,OkPourLettrage,ODAnalCegid) ;
  END ;

if (Lequel<>'FOD') and (Not ODAnalCegid) And
   (
   (FFormat.Value ='SN2') or (FFormat.Value ='SAA') or (FFormat.Value ='HAL') or
   (FFormat.Value ='HLI') or (FFormat.Value ='CGE') or (FFormat.Value ='CGN')
   ) then
  BEGIN
  if Pref='Y' then
    BEGIN
    // Mise à jour des analytiques.
    if OkUpdate then StWhere:=FabriqueStWhere(0,Pref,OkPourLettrage,ODAnalCegid)
    // Recherche des analytiques sur l'écriture en cours.
                Else StWhere:=FabriqueStWhere(2,Pref,OkPourLettrage,ODAnalCegid) ;
    END ;
  If Scenario.EstCharge And (Scenario.FiltreGen Or Scenario.FiltreAux) And (Lequel='FEC') And ((FFormat.Value ='CGE') or (FFormat.Value ='CGN')) Then
    BEGIN
    If Scenario.FiltreGen Then StWhere:=StWhere+' AND '+Fourchette(Pref+'_GENERAL',CpteDebut.Text,CpteFin.Text,ftString) ;
    If Scenario.FiltreAux Then
       BEGIN
// GP le 20/05/2001       StWhere:=StWhere+' AND '+Pref+'_ECHE="X" AND '+Pref+'_ETATLETTRAGE<>"RI"' ;
       StWhere:=StWhere+' AND '+Pref+'_ECHE="X" ' ;
       if FNatureAux.Value<>'' then BEGIN
                                    If OkUpdate Then
                                      BEGIN
                                      StWhere:=StWhere+' AND E_AUXILIAIRE IN(SELECT T_AUXILIAIRE FROM TIERS WHERE T_NATUREAUXI="'+FNatureAux.Value+'") ' ;
                                      END Else
                                      BEGIN
                                      St:=St+' LEFT JOIN TIERS ON '+Pref+'_AUXILIAIRE=T_AUXILIAIRE' ;
                                      StWhere:=StWhere+' AND T_NATUREAUXI="'+FNatureAux.Value+'"' ;
                                      END ;
                                    END ;
       if FModePaie.Value<>'' then StWhere:=StWhere+' AND '+Pref+'_MODEPAIE="'+FModePaie.Value+'" AND '+Pref+'_DATEECHEANCE>="'
                                           +USDate(FEche1)+'" AND '+Pref+'_DATEECHEANCE<="'+USDate(FEche2) +'"' ;
       END ;
    END ;
  END else
  BEGIN
  If Scenario.FiltreGen Or Scenario.FiltreAux Then
    BEGIN
    If Scenario.FiltreGen Then StWhere:=StWhere+' AND '+Fourchette(Pref+'_GENERAL',CpteDebut.Text,CpteFin.Text,ftString) ;
    If Scenario.FiltreAux Then
       BEGIN
// GP le 20/05/2001       StWhere:=StWhere+' AND '+Pref+'_ECHE="X" AND '+Pref+'_ETATLETTRAGE<>"RI"' ;
       StWhere:=StWhere+' AND '+Pref+'_ECHE="X" ' ;
       if FNatureAux.Value<>'' then BEGIN
                                    (*
                                    St:=St+' LEFT JOIN TIERS ON '+Pref+'_AUXILIAIRE=T_AUXILIAIRE' ;
                                    StWhere:=StWhere+' AND T_NATUREAUXI="'+FNatureAux.Value+'"' ;
                                    *)
                                    If OkUpdate Then
                                      BEGIN
                                      StWhere:=StWhere+' AND E_AUXILIAIRE IN(SELECT T_AUXILIAIRE FROM TIERS WHERE T_NATUREAUXI="'+FNatureAux.Value+'") ' ;
                                      END Else
                                      BEGIN
                                      St:=St+' LEFT JOIN TIERS ON '+Pref+'_AUXILIAIRE=T_AUXILIAIRE' ;
                                      StWhere:=StWhere+' AND T_NATUREAUXI="'+FNatureAux.Value+'"' ;
                                      END ;
                                    END ;
       if FModePaie.Value<>'' then StWhere:=StWhere+' AND '+Pref+'_MODEPAIE="'+FModePaie.Value+'" AND '+Pref+'_DATEECHEANCE>="'
                                           +USDate(FEche1)+'" AND '+Pref+'_DATEECHEANCE<="'+USDate(FEche2) +'"' ;
       END ;
    END Else
    BEGIN
    if (Lequel='FOD') then StWhere:=FabriqueStWhere(3,Pref,OkPourLettrage,ODAnalCegid) ;
    END ;
  END ;

if not PourVisu and not (FExport.State=cbGrayed) then
  if not FExport.Checked then StWhere:=StWhere+' AND '+Pref+'_EXPORTE<>"X"'
                         else StWhere:=StWhere+' AND '+Pref+'_EXPORTE="X"' ;

// maj de mouvements pour visu

if (PourVisu) or OkUpdate then
  BEGIN
  StP:='' ;
  For i:=0 to ListeNumJal.Count-1 do
    BEGIN
    if StP='' then StP:=' AND ('
              else StP:=StP+' OR ' ;
    FNumPiece:=TFNumPiece(ListeNumJal.Objects[i]) ;
    StP:=StP+'('+Pref+'_'+Jal+'="'+ListeNumJal[i]+'" AND '+Pref+'_NUMEROPIECE>='+IntToStr(FNumPiece.Min)+' AND '+Pref+'_NUMEROPIECE<='+IntToStr(FNumPiece.Max)+')' ;
    END ;
    if StP<>'' then StP:=StP+')' ;
    StWhere:=StWhere+StP ;
  if PourVisu And (Pref<>'BE') then
    BEGIN
    If Not OkPourLettrage Then
      BEGIN
      If (Lequel<>'FOD') And (Not ODAnalCEGID)
        then StWhere:=StWhere+' AND '+Pref+'_NUMLIGNE=1 AND '+Pref+'_NUMECHE<=1'
        Else StWhere:=StWhere+' AND '+Pref+'_NUMLIGNE=0 AND '+Pref+'_NUMVENTIL=1' ;
      END ;
    END ;
  END ;

if not OkUpdate then
  BEGIN
  If OkPourLettrage Then StOrder:=' ORDER BY E_AUXILIAIRE, E_GENERAL, E_ETATLETTRAGE, E_LETTRAGE, E_DATECOMPTABLE, E_DATEECHEANCE, E_NUMECHE ' Else
  BEGIN
    if CTri.Checked then
      BEGIN
      StOrder:=' ORDER BY '+Pref+'_NUMEROPIECE,'+Pref+'_QUALIFPIECE' ;
      If (Pref<>'Y') And (Not ODAnalCegid) then
        BEGIN
        if Pref='BE' then StOrder:=StOrder+','+Pref+'_DATECOMPTABLE'
                     else StOrder:=StOrder+','+Pref+'_NUMLIGNE,'+Pref+'_NUMECHE,'+Pref+'_DATECOMPTABLE' ;
        END else StOrder:=StOrder+','+Pref+'_AXE, '+Pref+'_NUMVENTIL, '+Pref+'_DATECOMPTABLE'  ;
      END else
      BEGIN
      StOrder:=' ORDER BY '+Pref+'_'+Jal+','+Pref+'_EXERCICE, '+Pref+'_DATECOMPTABLE, '+Pref+'_NUMEROPIECE, ' ;
      StOrder:=StOrder+Pref+'_QUALIFPIECE' ;
      If (Pref<>'Y') And (Not ODAnalCegid) then
        BEGIN
        if Pref<>'BE' then StOrder:=StOrder+','+Pref+'_NUMLIGNE,'+Pref+'_NUMECHE' ;
        END else StOrder:=StOrder+','+Pref+'_AXE, '+Pref+'_NUMVENTIL' ;
      END ;
    END ;
  END ;
// Ventilations analytiques des écritures générales
if (Lequel<>'FOD') and (Pref='Y') and (Not ODAnalCegid) then
  BEGIN
  ExecuteSQL(St+' '+StWhere+' '+StOrder) ;
  END else
  BEGIN
  ExecuteSQL(St+' '+StWhere+' '+StOrder) ;
  SourisNormale ;
  END ;
END ;

Function TFExpMvt.ReqSQL(Pref : String3 ; OkUpdate : boolean ; ODAnalCegid : Boolean ; Q1 : TQuery) : String ;
var StTable,Jal,St,StP,StWhere,StOrder : String ;
    OkPourLettrage : boolean ;
    FNumPiece : TFNumPiece ;
    i : integer ;
BEGIN
OkPourLettrage:=False ;
Jal:='JOURNAL' ; if (Pref='BE') then Jal:='BUDJAL' ;
if Pref='L' then BEGIN OkPourLettrage:=True ; Pref:='E' ; END ;
StTable:=PrefixeToTable(Pref) ;
If OkUpdate Then St:='UPDATE '+StTable+' SET '+Pref+'_EXPORTE="X" ' Else
  BEGIN
  if ((FFormat.Value='MP') and (FNatureAux.Value<>'')) then St:='SELECT *, '+GetChampsTiers(ctExpMvt)+' FROM '+StTable
                                                    else St:='SELECT * FROM '+StTable ;
  END ;
if (Pref<>'Y') Or ODAnalCegid then
  BEGIN
  If Pref='BE' Then StWhere:=FabriqueStWhere(1,Pref,OkPourLettrage,ODAnalCegid)
               Else StWhere:=FabriqueStWhere(0,Pref,OkPourLettrage,ODAnalCegid) ;
  END ;

if (Lequel<>'FOD') and (Not ODAnalCegid) And
   (
   (FFormat.Value ='SN2') or (FFormat.Value ='SAA') or (FFormat.Value ='HAL') or
   (FFormat.Value ='HLI') or (FFormat.Value ='CGE') or (FFormat.Value ='CGN')
   ) then
  BEGIN
  if Pref='Y' then
    BEGIN
    // Mise à jour des analytiques.
    if OkUpdate then StWhere:=FabriqueStWhere(0,Pref,OkPourLettrage,ODAnalCegid)
    // Recherche des analytiques sur l'écriture en cours.
//                Else StWhere:=FabriqueStWhere(2,Pref,OkPourLettrage,ODAnalCegid) ;
                Else StWhere:=FabriqueStWhereAna(Q1) ;
    END ;
  If Scenario.EstCharge And (Scenario.FiltreGen Or Scenario.FiltreAux) And (Lequel='FEC') And ((FFormat.Value ='CGE') or (FFormat.Value ='CGN')) Then
    BEGIN
    If Scenario.FiltreGen Then StWhere:=StWhere+' AND '+Fourchette(Pref+'_GENERAL',CpteDebut.Text,CpteFin.Text,ftString) ;
    If Scenario.FiltreAux Then
       BEGIN
// GP le 20/05/2001       StWhere:=StWhere+' AND '+Pref+'_ECHE="X" AND '+Pref+'_ETATLETTRAGE<>"RI"' ;
       StWhere:=StWhere+' AND '+Pref+'_ECHE="X" ' ;
       if FNatureAux.Value<>'' then BEGIN
                                    St:=St+' LEFT JOIN TIERS ON '+Pref+'_AUXILIAIRE=T_AUXILIAIRE' ;
                                    StWhere:=StWhere+' AND T_NATUREAUXI="'+FNatureAux.Value+'"' ;
                                    END ;
       if FModePaie.Value<>'' then StWhere:=StWhere+' AND '+Pref+'_MODEPAIE="'+FModePaie.Value+'" AND '+Pref+'_DATEECHEANCE>="'
                                           +USDate(FEche1)+'" AND '+Pref+'_DATEECHEANCE<="'+USDate(FEche2) +'"' ;
       END ;
    END ;
  END else
  BEGIN
  If Scenario.FiltreGen Or Scenario.FiltreAux Then
    BEGIN
    If Scenario.FiltreGen Then StWhere:=StWhere+' AND '+Fourchette(Pref+'_GENERAL',CpteDebut.Text,CpteFin.Text,ftString) ;
    If Scenario.FiltreAux Then
       BEGIN
// GP le 20/05/2001       StWhere:=StWhere+' AND '+Pref+'_ECHE="X" AND '+Pref+'_ETATLETTRAGE<>"RI"' ;
       StWhere:=StWhere+' AND '+Pref+'_ECHE="X" ' ;
       if FNatureAux.Value<>'' then BEGIN
                                    St:=St+' LEFT JOIN TIERS ON '+Pref+'_AUXILIAIRE=T_AUXILIAIRE' ;
                                    StWhere:=StWhere+' AND T_NATUREAUXI="'+FNatureAux.Value+'"' ;
                                    END ;
       if FModePaie.Value<>'' then StWhere:=StWhere+' AND '+Pref+'_MODEPAIE="'+FModePaie.Value+'" AND '+Pref+'_DATEECHEANCE>="'
                                           +USDate(FEche1)+'" AND '+Pref+'_DATEECHEANCE<="'+USDate(FEche2) +'"' ;
       END ;
    END Else
    BEGIN
    if (Lequel='FOD') then StWhere:=FabriqueStWhere(3,Pref,OkPourLettrage,ODAnalCegid) ;
    END ;
  END ;

if not PourVisu and not (FExport.State=cbGrayed) then
  if not FExport.Checked then StWhere:=StWhere+' AND '+Pref+'_EXPORTE<>"X"'
                         else StWhere:=StWhere+' AND '+Pref+'_EXPORTE="X"' ;

If Scenario.EstCharge And (NatExport=nRappro) And (Pref='E') Then
  BEGIN
  Case Lettre.State Of
    CBUnchecked : StWhere:=StWhere+' AND '+Pref+'_REFPOINTAGE="" ' ;
    CBchecked : StWhere:=StWhere+' AND '+Pref+'_REFPOINTAGE<>"" ' ;
    END ;
  END ;
// maj de mouvements pour visu

if (PourVisu) or OkUpdate then
  BEGIN
  StP:='' ;
  For i:=0 to ListeNumJal.Count-1 do
    BEGIN
    if StP='' then StP:=' AND ('
              else StP:=StP+' OR ' ;
    FNumPiece:=TFNumPiece(ListeNumJal.Objects[i]) ;
    StP:=StP+'('+Pref+'_'+Jal+'="'+ListeNumJal[i]+'" AND '+Pref+'_NUMEROPIECE>='+IntToStr(FNumPiece.Min)+' AND '+Pref+'_NUMEROPIECE<='+IntToStr(FNumPiece.Max)+')' ;
    END ;
    if StP<>'' then StP:=StP+')' ;
    StWhere:=StWhere+StP ;
  if PourVisu And (Pref<>'BE') then
    BEGIN
    If Not OkPourLettrage Then
      BEGIN
      If (Lequel<>'FOD') And (Not ODAnalCEGID)
        then StWhere:=StWhere+' AND '+Pref+'_NUMLIGNE=1 AND '+Pref+'_NUMECHE<=1'
        Else StWhere:=StWhere+' AND '+Pref+'_NUMLIGNE=0 AND '+Pref+'_NUMVENTIL=1' ;
      END ;
    END ;
  END ;

if not OkUpdate then
  BEGIN
  If OkPourLettrage Then StOrder:=' ORDER BY E_AUXILIAIRE, E_GENERAL, E_ETATLETTRAGE, E_LETTRAGE, E_DATECOMPTABLE, E_DATEECHEANCE, E_NUMECHE ' Else
  BEGIN
    if CTri.Checked then
      BEGIN
      StOrder:=' ORDER BY '+Pref+'_NUMEROPIECE,'+Pref+'_QUALIFPIECE' ;
      If (Pref<>'Y') And (Not ODAnalCegid) then
        BEGIN
        if Pref='BE' then StOrder:=StOrder+','+Pref+'_DATECOMPTABLE'
                     else StOrder:=StOrder+','+Pref+'_NUMLIGNE,'+Pref+'_NUMECHE,'+Pref+'_DATECOMPTABLE' ;
        END else StOrder:=StOrder+','+Pref+'_AXE, '+Pref+'_NUMVENTIL, '+Pref+'_DATECOMPTABLE'  ;
      END else
      BEGIN
      StOrder:=' ORDER BY '+Pref+'_'+Jal+','+Pref+'_EXERCICE, '+Pref+'_DATECOMPTABLE, '+Pref+'_NUMEROPIECE, ' ;
      StOrder:=StOrder+Pref+'_QUALIFPIECE' ;
      If (Pref<>'Y') And (Not ODAnalCegid) then
        BEGIN
        if Pref<>'BE' then StOrder:=StOrder+','+Pref+'_NUMLIGNE,'+Pref+'_NUMECHE' ;
        END else StOrder:=StOrder+','+Pref+'_AXE, '+Pref+'_NUMVENTIL' ;
      END ;
    END ;
  END ;
Result:=St+' '+StWhere+' '+StOrder ;
END ;

{================ Format Halley / SAARI ====================}

procedure TFExpMvt.ExporteLesAnals(Var BlocML : TMonoLigne ; LJalB : tStringlist ; Q1 : TQuery) ;
Var StInfosComptes : String ;
    St,StSQL2 : String ;
    Q2 : TQuery ;
BEGIN
If ((FFormat.Value='MP') Or (FFormat.Value='RAP')) then Exit ;
if (Q1.FindField('E_ANA').AsString='X') then
  BEGIN
  if (FFormat.Value ='SN2') or (FFormat.Value ='SAA') or
     (FFormat.Value ='HAL') or (FFormat.Value ='HLI') or
     (FFormat.Value ='CGE') or (FFormat.Value ='CGN') then
    BEGIN
    StSQL2:=ReqSQL('Y',False,FALSE,Q1) ;
    Q2:=OpenSQL(StSQL2,TRUE) ;
    (*
    QAna.ParamByName('JOURNAL').AsString:=Q1.FindField('E_JOURNAL').AsString ;
    QAna.ParamByName('EXERCICE').AsString:=Q1.FindField('E_EXERCICE').AsString ;
    QAna.ParamByName('DATECOMPTABLE').AsDateTime:=Q1.FindField('E_DATECOMPTABLE').AsDateTime ;
    QAna.ParamByName('NATUREPIECE').AsString:=Q1.FindField('E_NATUREPIECE').AsString ;
    QAna.ParamByName('NUMEROPIECE').AsInteger:=Q1.FindField('E_NUMEROPIECE').AsInteger ;
    QAna.ParamByName('QUALIFPIECE').AsString:=Q1.FindField('E_QUALIFPIECE').AsString ;
    QAna.ParamByName('NUMLIGNE').AsInteger:=Q1.FindField('E_NUMLIGNE').AsInteger ;
    QAna.Open ;
    *)
    StInfosComptes:=AjouteInfosCompte(Q2,'Y') ;
    While not Q2.EOF do
      BEGIN
      EcrireLigne(Q2,'Y',St,BlocML,LJalB) ;
      if (Lequel='FEC') and ((FFormat.Value='SAA') or (FFormat.Value='SN2')) then St:=St+StInfosComptes ;
      Writeln(Fichier,St) ;
      if (Lequel='FEC') and (FFormat.Value='HAL') and (StInfosComptes<>'') then WriteLn(Fichier,StInfosComptes) ;
      Q2.Next ;
      END ;
    Ferme(Q2) ;
    END ;
  END ;
END ;

Function TFExpMvt.OkEcrireLigne (QEcr : TQuery ; P : Char) : Boolean ;
BEGIN
Result:=TRUE ;
If ((FFormat.Value='CGE') Or (FFormat.Value='CGN')) And (P='E') And (Not CBDisting.Checked) And
   (QEcr.FindField('E_ANA').AsString='X') And (QEcr.FindField('E_ECHE').AsString<>'X') Then Result:=FALSE ;
END ;

Function CleMvtL(Aux,Gen,CodeLettre : String) : String ;
BEGIN
If Trim(CodeLettre)='' Then Exit ;
Result:=Format_String(Aux,17)+Format_String(Gen,17)+Format_String(CodeLettre,4) ;
END ;


procedure TFExpMvt.ExporteLesEcritures(ListeMvtL,ListeJalB : TStringList ; Q1 : TQuery) ;
Var TotLigneDeb,TotLigneCred, TotLigne : Double ;
    StGene,St : String ;
    StEche : Array[1..12] of String ;
    Sens : Char ;
    i, NumEche,NumPiece, OldNumEche : integer ;
    Jal,Exo : String ;
    DateC : TDateTime ;
    StInfosComptes : String ;
    BlocML : TmonoLigne ;
    CleL : String ;
BEGIN
Inc(NbMvt) ;
StInfosComptes:=AjouteInfosCompte(Q1,'E') ;
MoveCur(False) ; Fillchar(BlocML,SizeOf(BlocML),#0) ; BlocML.PasDeGeneSurCptVentilable:=FALSE ;
if (not Q1.Eof) and (Q1.FindField('E_ECHE').AsString='X') and (Q1.FindField('E_ETATLETTRAGE').AsString<>'RI') then
  BEGIN
// Format Négoce V2
  if FFormat.Value='SN2' then
      BEGIN
      TotLigneDeb:=0 ; TotLigneCred:=0 ; NumEche:=0 ;
      For i:=1 to 12 do StEche[i]:='' ;
      EcrireLigne(Q1,'E',StGene,BlocML,ListeJalB) ;
      Jal:=Q1.FindField('E_JOURNAL').AsString ; Exo:=Q1.FindField('E_EXERCICE').AsString ;
      DateC:=Q1.FindField('E_DATECOMPTABLE').AsDateTime ; NumPiece:=Q1.FindField('E_NUMEROPIECE').AsInteger ;
      OldNumEche:=0 ;
      While not (Q1.EOF) and (Q1.FindField('E_JOURNAL').AsString=Jal)
        and (Q1.FindField('E_EXERCICE').AsString=Exo)
        and (Q1.FindField('E_DATECOMPTABLE').AsDateTime=DateC)
        and (Q1.FindField('E_NUMEROPIECE').AsInteger=NumPiece)
        and (Q1.FindField('E_ECHE').AsString='X')
        and (Q1.FindField('E_NUMECHE').AsInteger>OldNumEche)
        and (Q1.FindField('E_ETATLETTRAGE').AsString<>'RI') do
        BEGIN
        OldNumEche:=Q1.FindField('E_NUMECHE').AsInteger ;
        NumEche:=NumEche+1 ;
        TotLigneDeb:=TotLigneDeb+Q1.FindField('E_DEBIT').AsFloat ;
        TotLigneCred:=TotLigneCred+Q1.FindField('E_CREDIT').AsFloat ;
        EcrireLigne(Q1,'E',StEche[NumEche],BlocML,ListeJalB) ;
        Q1.Next ;
        Jal:=Q1.FindField('E_JOURNAL').AsString ; Exo:=Q1.FindField('E_EXERCICE').AsString ;
        DateC:=Q1.FindField('E_DATECOMPTABLE').AsDateTime ; NumPiece:=Q1.FindField('E_NUMEROPIECE').AsInteger ;
        END ;
      TotLigne:=TotLigneDeb-TotLigneCred ;
      if TotLigne<>0 then
        BEGIN
        TotDeb:=TotDeb+TotLigneDeb ;
        TotCred:=TotCred+TotLigneCred ;
        if TotLigne<0 then Sens:='C' else Sens:='D' ;
        StGene:=Copy(StGene,1,83)+Sens+AlignDroite(StrfMontant(Abs(TotLigne),20,V_PGI.OkDecV,'',False),20)+Copy(StGene,105,length(StGene)-104) ;
        if (Lequel='FEC') and (FFormat.Value='SN2') then StGene:=StGene+StInfosComptes ;
        Writeln(Fichier,StGene) ;
        if not Q1.Eof then Q1.Prior ;
        END else
        BEGIN
        TotDeb:=TotDeb+Q1.Findfield('E_DEBIT').AsFloat ;
        TotCred:=TotCred+Q1.FindField('E_CREDIT').AsFloat ;
        END ;
      For i:=1 to NumEche do
        BEGIN
        StEche[i]:=Copy(StEche[i],1,24)+'E'+Copy(StEche[i],26,length(StEche[i])-25) ;
        if (Lequel='FEC') and (FFormat.Value='SN2') then StEche[i]:=StEche[i]+StInfosComptes ;
        Writeln(Fichier,StEche[i]) ;
        END ;
      END else
      //Autres... (Coeur de gamme...)
      BEGIN
      TotDeb:=TotDeb+Q1.Findfield('E_DEBIT').AsFloat ;
      TotCred:=TotCred+Q1.findField('E_CREDIT').AsFloat ;
      EcrireLigne(Q1,'E',St,BlocML,ListeJalB) ;
      if (Lequel='FEC') and ((FFormat.Value='SAA') or (FFormat.Value='HLI')) then St:=St+StInfosComptes ;
      if FFormat.Value='RAP' then
        BEGIN
        case Lettre.State of
          cbGrayed    : Writeln(Fichier,St) ;
          cbChecked   : if Copy(St,Length(St),1)='1' then Writeln(Fichier,St) ;
          cbUnchecked : if Copy(St,Length(St),1)='0' then Writeln(Fichier,St) ;
          END ;
        END else Writeln(Fichier,St) ;
      if (Lequel='FEC') and (FFormat.Value='HAL') and (StInfosComptes<>'') then WriteLn(Fichier,StInfosComptes) ;
      END ;
  If (ListeMvtL<>NIL) And Lettre.Checked And (Lequel='FEC')Then
    BEGIN
    CleL:=CleMvtL(Q1.FindField('E_AUXILIAIRE').AsString,Q1.FindField('E_GENERAL').AsString,Q1.FindField('E_LETTRAGE').AsString) ;
    ListeMvtL.Add(CleL) ;
    END ;
  END else
  BEGIN
  TotDeb:=TotDeb+Q1.Findfield('E_DEBIT').AsFloat ;
  TotCred:=TotCred+Q1.findField('E_CREDIT').AsFloat ;
  BlocML.PasDeGeneSurCptVentilable:=Not OkEcrireLigne(Q1,'E') ;
  EcrireLigne(Q1,'E',St,BlocML,ListeJalB) ;
  if (Lequel='FEC') and ((FFormat.Value='SAA') or (FFormat.Value='SN2') or (FFormat.Value='HLI')) then St:=St+StInfosComptes ;
  If Not BlocML.PasDeGeneSurCptVentilable Then Writeln(Fichier,St) ;
  if (Lequel='FEC') and (FFormat.Value='HAL') and (StInfosComptes<>'') then WriteLn(Fichier,StInfosComptes) ;
  END ;
// Lignes d'analytique...

ExporteLesAnals(BlocML,ListeJalB,Q1) ;
END ;

procedure TFExpMvt.ExporteLesEcrituresODAna(Q1 : tQuery) ;
Var St : String ;
    BlocML : TmonoLigne ;
BEGIN
Inc(NbMvt) ;
MoveCur(False) ; Fillchar(BlocML,SizeOf(BlocML),#0) ; BlocML.OdAnal:=TRUE ;
TotDeb:=TotDeb+Q1.Findfield('Y_DEBIT').AsFloat ;
TotCred:=TotCred+Q1.findField('Y_CREDIT').AsFloat ;
EcrireLigne(Q1,'Y',St,BlocML,NIL) ; Writeln(Fichier,St) ;
END ;

Procedure AlimJalB(LJB : tStringList) ;
Var QJ : TQuery ;
BEGIN
QJ:=OpenSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_MODESAISIE="BOR"',TRUE) ;
While Not QJ.Eof Do BEGIN LJB.Add(QJ.Fields[0].AsString) ; QJ.NExt ; END ;
Ferme(QJ) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 03/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TFExpMvt.ExporteMvt;
var
    ExpParam : boolean;
    Entete : TFmtEntete;
    Detail : TTabFmtDetail;
    Debut : Boolean;
    St : String;
    ListeMvtL : TStringList;
    ListeJalB : tStringList;
    StSQL1 : String;
    Q1 : TQuery;
    Q1Eof : Boolean;
begin
    NbMvt := 0;
    NbPiece := 0;
    TotDeb := 0;
    TotCred := 0;
    PourVisu := false;
    ListeMvtL := nil;

    if ((Lettre.Checked) And (NatExport = nPiece)) then
    begin
        ListeMvtL := TStringList.Create;
        ListeMvtL.Sorted := true;
        ListeMvtL.Duplicates := DupIgnore;
    end ;

    ListeJalB := TStringList.Create;
    ListeJalB.Sorted := true;
    ListeJalB.Duplicates := DupIgnore;
    AlimJalB(ListeJalB);

    // GG ConstruitRequete(Prefixe,False,FALSE) ;
    StSQL1 := ReqSQL(Prefixe,false,false,nil) ;

    // BPY le 03/10/2003 : test a faire HORS du try ..... parce ke si le test foire on OUVRE pas le fichié !
    ExpParam := (Lequel = 'FEC') and ((FFormat.Value <> 'SAA') and (FFormat.Value <> 'EDI') and (FFormat.Value <> 'SN2') and (FFormat.Value <> 'HAL') and (FFormat.Value <> 'HLI') and (FFormat.Value <> 'MP') and (FFormat.Value <> 'RAP') and (FFormat.Value <> 'CGN') and (FFormat.Value <> 'CGE')) or (Lequel = 'FBE') or (Lequel = 'FOD');
    if (ExpParam) then if (not ChargeFormat(Fichier,Filename.Text,Lequel,'-',FFormat.Value,Entete,Detail,Debut)) then
    begin
        MsgBox.Execute(43,caption,'');
        exit;
    end;
    // fin BPY

    // GG if QEcr.Eof then BEGIN MsgBox.Execute(4,caption,'') ; Exit ; END ;
    try
        // BPY : le 03/10/2003
//        ExpParam := (Lequel = 'FEC') and ((FFormat.Value <> 'SAA') and (FFormat.Value <> 'EDI') and (FFormat.Value <> 'SN2') and (FFormat.Value <> 'HAL') and (FFormat.Value <> 'HLI') and (FFormat.Value <> 'MP') and (FFormat.Value <> 'RAP') and (FFormat.Value <> 'CGN') and (FFormat.Value <> 'CGE')) or (Lequel = 'FBE') or (Lequel = 'FOD');
//
//        if (ExpParam) then
//        begin
//            if (not ChargeFormat(Fichier,Filename.Text,Lequel,'-',FFormat.Value,Entete,Detail,Debut)) then exit;
//        end
//        else
//        begin
//            AssignFile(Fichier,FileName.Text) ;
//            {$I-} ReWrite (Fichier) ; {$I+}
//            St:=FirstLigneFic(FFormat.Value,Lequel) ;
//            Writeln(Fichier,St) ;
//        end;
        if (not ExpParam) then
        begin
            AssignFile(Fichier,FileName.Text) ;
            {$I-} ReWrite (Fichier) ; {$I+}
            St:=FirstLigneFic(FFormat.Value,Lequel) ;
            Writeln(Fichier,St) ;
        end;
        // fin BPY

        OldJal := '';
        OldNat := '';
        OldDateC := iDate1900;
        OldNumP := 0;
        FirstJal := '';
        OldQualP := '';
        //  InitMove(RecordsCount(QEcr)+2,'') ;
        MoveCur(false);
        (*
        if not ExpParam and ((FFormat.Value<>'MP') and (FFormat.Value<>'RAP')) then
        BEGIN
        StSQL2:=ReqSQL('Y',False,FALSE) ;
        END ;
        *)
        //QEcr.First ;
        Q1 := OpenSQL(StSQL1,TRUE);
        if (Q1.Eof) then
        begin
            Q1Eof := TRUE;
            MsgBox.Execute(4,caption,'');
        end
        else
        begin
            Q1Eof := FALSE;
            InitMove(RecordsCount(Q1)+2,'');
        end;

        while (not Q1.Eof) do
        begin
            if (TestBreak(Q1)) then break;
            if (ExpParam) then
            begin
                EcrireFormat(Fichier,Entete,Detail,Debut,Q1);
                TotDeb := TotDeb + Q1.Findfield(Prefixe + '_DEBIT').AsFloat;
                TotCred := TotCred + Q1.FindField(Prefixe + '_CREDIT').AsFloat;
            end
            else
            begin
                ExporteLesEcritures(ListeMvtL,ListeJalB,Q1);
            end;
            Q1.Next;
            MoveCur(False);
        end;
        Ferme(Q1);

        // BPY le 03/10/2003 : de l'utilité de 2* le meme test .....
//        if (Not Q1Eof) then Inc(NbPiece);
//        If (Not Q1Eof) then
//        begin
//            // GG Update à faire
//            ConstruitRequete(Prefixe,true,false);
//            MoveCur(false);
//            if (not ExpParam and ((FFormat.Value<>'MP') and (FFormat.Value<>'RAP'))) then ConstruitRequete('Y',true,false);
//            FiniMove;
//        end;
        If (Not Q1Eof) then
        begin
            Inc(NbPiece);
            // GG Update à faire
            ConstruitRequete(Prefixe,true,false);
            MoveCur(false);
            if (not ExpParam and ((FFormat.Value<>'MP') and (FFormat.Value<>'RAP'))) then ConstruitRequete('Y',true,false);
            FiniMove;
        end;
        // fin BPY
    finally
        CloseFile(Fichier);
    end;

    StTotDeb := AlignDroite(StrfMontant(Abs(TotDeb),20,V_PGI.OkDecV,V_PGI.SymbolePivot,true),20);
    StTotCred := AlignDroite(StrfMontant(Abs(TotCred),20,V_PGI.OkDecV,V_PGI.SymbolePivot,true),20);

    if ((Lettre.Checked) And (NatExport = nPiece)) then
    begin
        ExporteLettre(ListeMvtL);
        ListeMvtL.Clear;
        ListeMvtL.Free;
    end;
    ListeJalB.Clear;
    ListeJalB.Free;
end;

Procedure TFExpMvt.ExporteMvtODACegid(AppendFichier : Boolean) ;
var St       : String ;
    Q1       : TQuery ;
    Q1Eof : Boolean ;
    StSQL : String ;
BEGIN
If Not AppendFichier Then BEGIN NbMvt:=0 ; NbPiece:=0 ; TotDeb:=0 ; TotCred:=0 ; END ;
PourVisu:=False ;
Prefixe:='Y' ;
StSQL:=ReqSQL(Prefixe,FALSE,TRUE,NIL) ;
Q1:=OpenSQL(StSQL,TRUE) ; Q1Eof:=Q1.Eof ;
//ConstruitRequete(Prefixe,False,TRUE) ;
if (Q1Eof) And (Not AppendFichier) then BEGIN MsgBox.Execute(4,caption,'') ; Ferme(Q1) ; Exit ; END ;
try
  AssignFile(Fichier,FileName.Text) ;
  {$I-}
  If AppendFichier Then Append(Fichier) Else ReWrite (Fichier) ;
  {$I+}
  If Not AppendFichier Then
    BEGIN
    St:=FirstLigneFic(FFormat.Value,Lequel) ;
    Writeln(Fichier,St) ;
    END ;
  OldJal:='' ; OldNat:='' ; OldDateC:=iDate1900 ; OldNumP:=0 ;
  FirstJal:='' ; OldQualP:='' ;
  InitMove(RecordsCount(Q1)+2,'') ;
  MoveCur(False) ;
  While not Q1.Eof do
    BEGIN
    if TestBreak(Q1) then Break ;
    ExporteLesEcrituresODAna(Q1) ;
    Q1.Next ; MoveCur(False) ;
    END ;
  if Q1.Eof then Inc(NbPiece) ;
  ConstruitRequete(Prefixe,True,TRUE) ;
  MoveCur(False);
  FiniMove ;
  finally
  Ferme(Q1) ;
  CloseFile(Fichier) ;
  end ;
StTotDeb:=AlignDroite(StrfMontant(Abs(TotDeb),20,V_PGI.OkDecV,V_PGI.SymbolePivot,True),20) ;
StTotCred:=AlignDroite(StrfMontant(Abs(TotCred),20,V_PGI.OkDecV,V_PGI.SymbolePivot,True),20) ;
END ;


(*
Function QuelMontant(Q : TQuery ; P : String ; b : Byte ; Decim : Integer ; Var Sens : String) : String ;
Var MD,MC,Montant : Double ;
    Sup : String ;
    LeSens : String ;
BEGIN
Result:='' ; Sup:='' ; Montant:=0 ;
Case b Of 1 : Sup:='DEV' ; 2 : Sup:='EURO' ; END ;
MD:=Q.Findfield(P+'_DEBIT'+Sup).AsFloat ; MC:=Q.Findfield(P+'_CREDIT'+Sup).AsFloat ;
MD:=Arrondi(MD,Decim) ; MC:=Arrondi(MC,Decim) ;
if ((MD<0) or (MD>0)) and (MC=0) then BEGIN Montant:=MD ; LeSens:='D' ; END else
 if ((MC<0) or (MC>0)) and (MD=0) then BEGIN Montant:=MC ; LeSens:='C' ; END ;
If Sens='' Then Sens:=LeSens ;
Result:=StrfMontant(Montant,20,Decim,'',False) ;
END ;
*)
Function QuelCouverture(Q : TQuery ; P : String ; b : Byte ; Decim : Integer) : String ;
Var Montant : Double ;
    Sup : String ;
BEGIN
Result:='' ; Sup:='';
//Case b Of 1 : Sup:='DEV' ; 2 : Sup:='EURO' ; END ;
Case b Of 1 : Sup:='DEV' ; 2 : Sup:='' ; END ;  // CA - 30/09/2003 - Suppression contre-valeur
Montant:=Q.Findfield(P+'_COUVERTURE'+Sup).AsFloat ;
Montant:=Arrondi(Montant,Decim) ;
Result:=StrfMontant(Montant,20,Decim,'',False) ;
END ;

Procedure AjouteZonesCEGID(Q : TQuery ; Var St : String ; P : Char ; Decim : Integer ; Var BlocML : TMonoLigne) ;
Var StMontantCegid : Array[1..3] Of String ;
    Tva,Tpf,TvaEnc,CtrGen,CtrAux : String ;
BEGIN
Fillchar(StMontantCegid,SizeOf(StMontantCegid),#0) ;
If P='E' Then
  BEGIN
  If VH^.TenueEuro Then
    BEGIN
    StMontantCegid[3]:=QuelCouverture(Q,P,0,V_PGI.OkDecV) ;
    StMontantCegid[2]:=QuelCouverture(Q,P,1,Decim) ;
    StMontantCegid[1]:=QuelCouverture(Q,P,2,V_PGI.OkDecE) ;
    END Else
    BEGIN
    StMontantCegid[1]:=QuelCouverture(Q,P,0,V_PGI.OkDecV) ;
    StMontantCegid[2]:=QuelCouverture(Q,P,1,Decim) ;
    StMontantCegid[3]:=QuelCouverture(Q,P,2,V_PGI.OkDecE) ;
    END ;
  If BlocML.PasDeGeneSurCptVentilable Then
    BEGIN
    BlocML.TVA:=Q.FindField(P+'_TVA').AsString ;
    BlocML.TPF:=Q.FindField(P+'_TPF').AsString ;
    BlocML.TvaEnc:=Q.FindField(P+'_TVAENCAISSEMENT').AsString ;
    BlocML.CtrGen:=Q.FindField(P+'_CONTREPARTIEGEN').AsString ;
    BlocML.CtrAux:=Q.FindField(P+'_CONTREPARTIEAUX').AsString ;
    END ;
  St:=St+Format_String(Q.FindField(P+'_REFEXTERNE').AsString,35)
         +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATEREFEXTERNE').AsDateTime))
         +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATECREATION').AsDateTime))
         +Format_String(V_PGI.CodeSociete,3)
         +Format_String(Q.FindField(P+'_AFFAIRE').AsString,17)
         +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATETAUXDEV').AsDateTime))
         +Format_String(Q.FindField(P+'_ECRANOUVEAU').AsString,3)
         +AlignDroite(StrfMontant(Q.FindField(P+'_QTE1').AsFloat,20,4,'',False),20)
         +AlignDroite(StrfMontant(Q.FindField(P+'_QTE2').AsFloat,20,4,'',False),20)
         +Format_String(Q.FindField(P+'_QUALIFQTE1').AsString,3)
         +Format_String(Q.FindField(P+'_QUALIFQTE2').AsString,3)
         +Format_String(Q.FindField(P+'_REFLIBRE').AsString,35)
         +Format_String(Q.FindField(P+'_TVAENCAISSEMENT').AsString,1)
         +Format_String(Q.FindField(P+'_REGIMETVA').AsString,3)
         +Format_String(Q.FindField(P+'_TVA').AsString,3)
         +Format_String(Q.FindField(P+'_TPF').AsString,3)
         +Format_String(Q.FindField(P+'_CONTREPARTIEGEN').AsString,17)
         +Format_String(Q.FindField(P+'_CONTREPARTIEAUX').AsString,17)
         +Format_String(Q.FindField(P+'_REFPOINTAGE').AsString,17)
         +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATEPOINTAGE').AsDateTime))
         +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATERELANCE').AsDateTime))
         +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATEVALEUR').AsDateTime))
         +Format_String(Q.FindField(P+'_RIB').AsString,35)
         +Format_String(Q.FindField(P+'_REFRELEVE').AsString,10)
         +Format_String(' ',17)
         +Format_String(Q.FindField(P+'_LIBRETEXTE0').AsString,30)
         +Format_String(Q.FindField(P+'_LIBRETEXTE1').AsString,30)
         +Format_String(Q.FindField(P+'_LIBRETEXTE2').AsString,30)
         +Format_String(Q.FindField(P+'_LIBRETEXTE3').AsString,30)
         +Format_String(Q.FindField(P+'_LIBRETEXTE4').AsString,30)
         +Format_String(Q.FindField(P+'_LIBRETEXTE5').AsString,30)
         +Format_String(Q.FindField(P+'_LIBRETEXTE6').AsString,30)
         +Format_String(Q.FindField(P+'_LIBRETEXTE7').AsString,30)
         +Format_String(Q.FindField(P+'_LIBRETEXTE8').AsString,30)
         +Format_String(Q.FindField(P+'_LIBRETEXTE9').AsString,30)
         +Format_String(Q.FindField(P+'_TABLE0').AsString,3)
         +Format_String(Q.FindField(P+'_TABLE1').AsString,3)
         +Format_String(Q.FindField(P+'_TABLE2').AsString,3)
         +Format_String(Q.FindField(P+'_TABLE3').AsString,3)
         +AlignDroite(StrfMontant(Q.FindField(P+'_LIBREMONTANT0').AsFloat,20,4,'',False),20)
         +AlignDroite(StrfMontant(Q.FindField(P+'_LIBREMONTANT1').AsFloat,20,4,'',False),20)
         +AlignDroite(StrfMontant(Q.FindField(P+'_LIBREMONTANT2').AsFloat,20,4,'',False),20)
         +AlignDroite(StrfMontant(Q.FindField(P+'_LIBREMONTANT3').AsFloat,20,4,'',False),20)
         +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_LIBREDATE').AsDateTime))
         +Format_String(Q.FindField(P+'_LIBREBOOL0').AsString,1)
         +Format_String(Q.FindField(P+'_LIBREBOOL1').AsString,1)
         +Format_String(Q.FindField(P+'_CONSO').AsString,3)
         +AlignDroite(StMontantCegid[1],20)
         +AlignDroite(StMontantCegid[2],20)
         +AlignDroite(StMontantCegid[3],20)
         +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATEPAQUETMIN').AsDateTime))
         +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATEPAQUETMAX').AsDateTime))
         +Format_String(Q.FindField(P+'_LETTRAGE').AsString,5)
         +Format_String(Q.FindField(P+'_LETTRAGEDEV').AsString,1)
         +Format_String(' ',1)
         +Format_String(Q.FindField(P+'_ETATLETTRAGE').AsString,3) ;
  END Else
  BEGIN
  If BlocML.PasDeGeneSurCptVentilable Then
    BEGIN
    CtrGen:=BlocML.CtrGen ;
    CtrAux:=BlocML.CtrAux ;
    Tva:=BlocML.TVA ;
    Tpf:=BlocML.TPF ;
    TvaEnc:=BlocML.TvaEnc ;
    END Else
    BEGIN
    CtrGen:=Q.FindField(P+'_CONTREPARTIEGEN').AsString ;
    CtrAux:=Q.FindField(P+'_CONTREPARTIEAUX').AsString ;
    Tva:='' ; Tpf:='' ; TvaEnc:='' ;
    END ;
  St:=St+Format_String(Q.FindField(P+'_REFEXTERNE').AsString,35)
         +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATEREFEXTERNE').AsDateTime))
         +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATECREATION').AsDateTime))
         +Format_String(V_PGI.CodeSociete,3)
         +Format_String(Q.FindField(P+'_AFFAIRE').AsString,17)
         +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATETAUXDEV').AsDateTime))
         +Format_String(Q.FindField(P+'_ECRANOUVEAU').AsString,3)
         +AlignDroite(StrfMontant(Q.FindField(P+'_QTE1').AsFloat,20,4,'',False),20)
         +AlignDroite(StrfMontant(Q.FindField(P+'_QTE2').AsFloat,20,4,'',False),20)
         +Format_String(Q.FindField(P+'_QUALIFQTE1').AsString,3)
         +Format_String(Q.FindField(P+'_QUALIFQTE2').AsString,3)
         +Format_String(Q.FindField(P+'_REFLIBRE').AsString,35)
         +Format_String(TvaEnc,1)
         +Format_String(' ',3)
         +Format_String(Tva,3)
         +Format_String(Tpf,3)
         +Format_String(CtrGen,17)
         +Format_String(CtrAux,17)
         +Format_String(' ',17)
         +Format_String(' ',8)
         +Format_String(' ',8)
         +Format_String(' ',8)
         +Format_String(' ',35)
         +Format_String(' ',10)
         +Format_String(' ',17)
         +Format_String(Q.FindField(P+'_LIBRETEXTE0').AsString,30)
         +Format_String(Q.FindField(P+'_LIBRETEXTE1').AsString,30)
         +Format_String(Q.FindField(P+'_LIBRETEXTE2').AsString,30)
         +Format_String(Q.FindField(P+'_LIBRETEXTE3').AsString,30)
         +Format_String(Q.FindField(P+'_LIBRETEXTE4').AsString,30)
         +Format_String(Q.FindField(P+'_LIBRETEXTE5').AsString,30)
         +Format_String(Q.FindField(P+'_LIBRETEXTE6').AsString,30)
         +Format_String(Q.FindField(P+'_LIBRETEXTE7').AsString,30)
         +Format_String(Q.FindField(P+'_LIBRETEXTE8').AsString,30)
         +Format_String(Q.FindField(P+'_LIBRETEXTE9').AsString,30)
         +Format_String(Q.FindField(P+'_TABLE0').AsString,3)
         +Format_String(Q.FindField(P+'_TABLE1').AsString,3)
         +Format_String(Q.FindField(P+'_TABLE2').AsString,3)
         +Format_String(Q.FindField(P+'_TABLE3').AsString,3)
         +AlignDroite(StrfMontant(Q.FindField(P+'_LIBREMONTANT0').AsFloat,20,4,'',False),20)
         +AlignDroite(StrfMontant(Q.FindField(P+'_LIBREMONTANT1').AsFloat,20,4,'',False),20)
         +AlignDroite(StrfMontant(Q.FindField(P+'_LIBREMONTANT2').AsFloat,20,4,'',False),20)
         +AlignDroite(StrfMontant(Q.FindField(P+'_LIBREMONTANT3').AsFloat,20,4,'',False),20)
         +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_LIBREDATE').AsDateTime))
         +Format_String(Q.FindField(P+'_LIBREBOOL0').AsString,1)
         +Format_String(Q.FindField(P+'_LIBREBOOL1').AsString,1)
         +Format_String(' ',3)
         +Format_String(' ',20)
         +Format_String(' ',20)
         +Format_String(' ',20)
         +Format_String(' ',8)
         +Format_String(' ',8)
         +Format_String(' ',5)
         +Format_String(' ',1)
         +Format_String(' ',1)
         +Format_String(' ',1) ;
  END ;
END ;

procedure TFExpMvt.EcrireLigne(Q : TQuery ; P : Char ; var St : String ; Var BlocML : TMonoLigne ; LJalB : tStringList) ;
var Jal,DateCpta,Ref,Lib,MPaie,Ech,NumPiece,StSoldeLettr,StMontant,CodeMontant,CodeDev,Axe,NumEche,Etab :string ;
    Montant,MD,MC : Double ;
    TypeEcr,RefPointage : String1 ;
    CatPaie,TypeCpte : String ;
    CpteGene,AuxiAna,TypePiece,Sens : String ;
    LgCpte : Byte ;
    OkEuro,SaisieOpposee,ComptaEnEuro : Boolean ;
    StMontantCegid : Array[1..3] Of String ;
    DecDev : integer ;
    Quotite,TauxDev: Double ;
    EnDevise,EnMontantOppose : Boolean ;
BEGIN
LgCpte:=13 ; Fillchar(StMontantCegid,SizeOf(StMontantCegid),#0) ;
DecDev:=V_PGI.OkDecV ; Etab:=VH^.EtablisDefaut ;
if (FFormat.Value='HLI') or (FFormat.Value='HAL') Or
   (FFormat.Value='CGN') or (FFormat.Value='CGE')then LgCpte:=17 ;
Jal:=' ';DateCpta:=StDate1900 ; TypePiece:=' ';CpteGene:=' ';AuxiAna:=' ';Ref:=' ';
Lib:=' ';MPaie:=' ';Ech:=StDate1900 ;TypeEcr:=' ';NumPiece:=' '; StSoldeLettr:=' ' ;
Axe:='  ' ; NumEche:=' 0' ;
(*
Sens:='D' ;
Montant:=Q.Findfield(P+'_DEBIT').AsFloat-Q.FindField(P+'_CREDIT').AsFloat ;
if Montant<0 then Sens:='C' ;
*)
// si on garde le signe
Jal:=Copy(Q.FindField(P+'_JOURNAL').AsString,1,3) ; DateCpta:=DateToStr(Q.FindField(P+'_DATECOMPTABLE').AsDateTime) ;
OkEuro:=Q.FindField(P+'_DATECOMPTABLE').AsDateTime>=V_PGI.DateDebutEuro ;
ComptaEnEuro:=VH^.TenueEuro ;
If OkEuro Then
  BEGIN
  CodeMontant:='FDE' ;
  END Else
  BEGIN
  CodeMontant:='FD-' ;
  END ;
if (FFormat.Value='CGN') or (FFormat.Value='CGE') then
   BEGIN
   Etab:=Q.FindField(P+'_ETABLISSEMENT').AsString ;
   TypePiece:=Copy(Q.FindField(P+'_NATUREPIECE').AsString,1,2) ;
   Sens:='' ;
   CodeDev:=Q.FindField(P+'_DEVISE').AsString ;
   EnDevise:=RecupDevise(CodeDev,DecDev,Quotite,TDev) ;
   If EnDevise Then
     BEGIN
     // CodeMontant:='DFE' ;
      CodeMontant:='DE-' ;
      StMontantCegid[1]:=QuelMontant(Q,P,1,DecDev,Sens) ;
//     StMontantCegid[2]:=QuelMontant(Q,P,0,V_PGI.OkDecV,Sens) ;
      StMontantCegid[2]:=QuelMontant(Q,P,0,V_PGI.OkDecV,Sens) ;
//     StMontantCegid[3]:=QuelMontant(Q,P,2,V_PGI.OkDecE,Sens) ;
      StMontantCegid[3]:=Format_String(' ',20) ;
     If ComptaEnEuro Then CodeMontant:='DE-' ;
     END Else
     BEGIN
     If OkEuro Then
       BEGIN
       EnMontantOppose:=False;
       If EnMontantOppose Then
         BEGIN
         If ComptaEnEuro Then
           BEGIN
           CodeMontant:='FE-' ;
           StMontantCegid[2]:=QuelMontant(Q,P,0,V_PGI.OkDecV,Sens) ;
           StMontantCegid[1]:=QuelMontant(Q,P,2,V_PGI.OkDecE,Sens) ;
           StMontantCegid[3]:=Format_String(' ',20) ;
           END Else
           BEGIN
           CodeMontant:='EF-' ;
           StMontantCegid[2]:=QuelMontant(Q,P,0,V_PGI.OkDecV,Sens) ;
           StMontantCegid[1]:=QuelMontant(Q,P,2,V_PGI.OkDecE,Sens) ;
           StMontantCegid[3]:=Format_String(' ',20) ;
           END ;
         END Else
         BEGIN
         If ComptaEnEuro Then
           BEGIN
           CodeMontant:='E--' ;
           StMontantCegid[1]:=QuelMontant(Q,P,0,V_PGI.OkDecV,Sens) ;
           //StMontantCegid[2]:=QuelMontant(Q,P,2,V_PGI.OkDecE,Sens) ;
           StMontantCegid[2]:=Format_String(' ',20) ;
           StMontantCegid[3]:=Format_String(' ',20) ;
           END Else
           BEGIN
           CodeMontant:='FE-' ;
           StMontantCegid[1]:=QuelMontant(Q,P,0,V_PGI.OkDecV,Sens) ;
           StMontantCegid[2]:=QuelMontant(Q,P,2,V_PGI.OkDecE,Sens) ;
           StMontantCegid[3]:=Format_String(' ',20) ;
           END ;
         END ;
       END Else
       BEGIN
       CodeMontant:='FE-' ;
       StMontantCegid[1]:=QuelMontant(Q,P,0,V_PGI.OkDecV,Sens) ;
       StMontantCegid[2]:=QuelMontant(Q,P,2,V_PGI.OkDecE,Sens) ;
       StMontantCegid[3]:=Format_String(' ',20) ;
       END ;
     END ;
   CpteGene:=Copy(Q.FindField(P+'_GENERAL').AsString,1,17) ;
   Ref:=Copy(Q.FindField(P+'_REFINTERNE').AsString,1,35) ; Lib:=Copy(Q.FindField(P+'_LIBELLE').AsString,1,35) ;
   NumPiece:=Copy(Q.FindField(P+'_NUMEROPIECE').AsString,1,8) ;
//   If LJalB.IndexOf(Jal)<0 Then NumPiece:=Copy(Q.FindField(P+'_NUMEROPIECE').AsString,1,8) ;
   TypeEcr:=Copy(Q.Findfield(P+'_QUALIFPIECE').AsString,1,1) ;
   END Else
   BEGIN
   Montant:=0 ;
   MD:=Q.Findfield(P+'_DEBIT').AsFloat ; MC:=Q.Findfield(P+'_CREDIT').AsFloat ;
   if ((MD<0) or (MD>0)) and (MC=0) then BEGIN Sens:='D' ; Montant:=Q.Findfield(P+'_DEBIT').AsFloat ; END else
    if ((MC<0) or (MC>0)) and (MD=0) then BEGIN Sens:='C' ; Montant:=Q.Findfield(P+'_CREDIT').AsFloat END ;
   StMontant:=StrfMontant(Montant,20,V_PGI.OkDecV,'',False) ;
   TypePiece:=Copy(Q.FindField(P+'_NATUREPIECE').AsString,1,2) ;
// CP : 09/02/98
// Pour format SAARI -> pas d' acomptes cli. ou Fou.
// Format HALLEY étendu : NATUREPIECE est en fin d'enregistrement
   if TypePiece='OC' then TypePiece:='RC' else if TypePiece='OF' then TypePiece:='RF' ;
   CpteGene:=Copy(Q.FindField(P+'_GENERAL').AsString,1,LgCpte) ;
   Ref:=Copy(Q.FindField(P+'_REFINTERNE').AsString,1,13) ; Lib:=Copy(Q.FindField(P+'_LIBELLE').AsString,1,25) ;
//   If LJalB.IndexOf(Jal)<0 Then NumPiece:=Copy(Q.FindField(P+'_NUMEROPIECE').AsString,1,7) ;
   NumPiece:=Copy(Q.FindField(P+'_NUMEROPIECE').AsString,1,7) ;
   TypeEcr:=Copy(Q.Findfield(P+'_QUALIFPIECE').AsString,1,1) ;
   END ;
ListeCptCegid.Add('G'+Trim(CpteGene)) ;
If P='E' Then NumEche:=Copy(Q.FindField(P+'_NUMECHE').AsString,1,2)
         Else Axe:=Q.FindField(P+'_AXE').AsString ;
if P='E'then
      BEGIN
      AuxiAna:=Copy(Q.FindField(P+'_AUXILIAIRE').AsString,1,LgCpte) ;
      ListeCptCegid.Add('X'+Trim(AuxiAna)) ;
      if (FFormat.Value='CGN') or (FFormat.Value='CGE') then MPaie:=Q.FindField(P+'_MODEPAIE').AsString Else
         BEGIN
         CatPaie:=ChercheCategPaie(Q.FindField(P+'_MODEPAIE').AsString) ;
         MPaie:=CHERCHEUNMODE(CatPaie,Q.FindField(P+'_MODEPAIE').AsString,False,False) ;
         END ;
      if Q.FindField(P+'_ECHE').AsString='X' then Ech:=DateToStr(Q.FindField(P+'_DATEECHEANCE').AsDateTime) ;
      if AuxiAna <>'' then TypeCpte:= 'X' else TypeCpte:=' ' ;
      if (FFormat.Value='RAP') then
        BEGIN
        if (Q.FindField(P+'_REFPOINTAGE').AsString<>'') then RefPointage:='1' else RefPointage:='0' ;
        END ;
      (* .. Pas de solde lettrage .....
      Case FFormat.ItemIndex of
        0: BEGIN
           SoldeLettr:=' ' ;
           //SoldeLettr:=Abs(Montant)-Q.FindField(P+'_COUVERTURE').asFloat ;
           //StSoldeLettr:=StrfMontant(Abs(SoldeLettr),20,2,'',False) ;
           END ;
        1: if RefPointage<>'' then StSoldeLettr:='1' else StSoldeLettr:='0' ;
        2: StSoldeLettr:=' ' ;
        END ;
 *)
      END Else
      BEGIN
      TypeCpte:='A' ;
      AuxiAna:=Copy(Q.FindField(P+'_SECTION').AsString,1,LgCpte) ;
      If BlocML.PasDeGeneSurCptVentilable Then
        BEGIN
        TypeCpte:='H' ;
        END ;
      If BlocML.ODAnal Then TypeCpte:='O' ;
      ListeCptCegid.Add('S'+Trim(AuxiAna)+';'+Trim(Axe)) ;
      END ;
 if (FFormat.Value='HAL') or (FFormat.Value='HLI') then
  BEGIN
  St:=Format_String(Jal,3)+Format_String(Fmt_Date_HAL(DateCpta),8)+Format_String(TypePiece,2)
     +Format_String(CpteGene,LgCpte)+Format_String(TypeCpte,1)+Format_String(AuxiAna,LgCpte)
     +Format_String(Ref,13)+Format_String(Lib,25)+Format_String(MPaie,1)
     +Format_String(Fmt_Date_HAL(Ech),8)+Format_String(Sens,1)+AlignDroite(StMontant,20)
     +Format_String(TypeEcr,1)+AlignDroite(NumPiece,7); //+AlignDroite(StSoldeLettr,20) ;
  if (FFormat.Value='HAL') then AjouteZonesHal(Q,St,P) ;
  END else
 if (FFormat.Value='CGN') or (FFormat.Value='CGE') then
  BEGIN
  If OkEuro Then BEGIN
    SaisieOpposee:=False ;
    If CodeDev=V_PGI.DevisePivot Then BEGIN
      If ((VH^.TenueEuro) And (Not SaisieOpposee)) Or
         ((Not VH^.TenueEuro) And (SaisieOpposee)) Then CodeDev:='EUR' ;
    END ;
    TauxDev:=Q.FindField(P+'_TAUXDEV').AsFloat/V_PGI.TauxEuro ;
    END
  Else BEGIN
    TauxDev:=Q.FindField(P+'_TAUXDEV').AsFloat/V_PGI.TauxEuro ;
  END ;
  St:=Format_String(Jal,3)+Format_String(Fmt_Date_HAL(DateCpta),8)+Format_String(TypePiece,2)
     +Format_String(CpteGene,LgCpte)+Format_String(TypeCpte,1)+Format_String(AuxiAna,LgCpte)
     +Format_String(Ref,35)+Format_String(Lib,35)+Format_String(MPaie,3)
     +Format_String(Fmt_Date_HAL(Ech),8)+Format_String(Sens,1)+AlignDroite(StMontantCegid[1],20)
     +Format_String(TypeEcr,1)+AlignDroite(NumPiece,8)+Format_String(CodeDev,3)
     +AlignDroite(StrfMontant(TauxDev,20,7,'',False),10)+CodeMontant
     +AlignDroite(StMontantCegid[2],20)+AlignDroite(StMontantCegid[3],20)+Format_String(Etab,3)
     +Format_String(Axe,2)+AlignDroite(NumEche,2) ;
  if (FFormat.Value='CGE') then AjouteZonesCEGID(Q,St,P,DecDev,BlocML) ;
  END else
  St:=Format_String(Jal,3)+Fmt_Date_Saari(DateCpta)+Format_String(TypePiece,2)
     +Format_String(CpteGene,13)+Format_String(TypeCpte,1)+Format_String(AuxiAna,13)
     +Format_String(Ref,13)+Format_String(Lib,25)+Format_String(MPaie,1)
     +Format_String(Fmt_Date_Saari(Ech),6)+Format_String(Sens,1)+AlignDroite(StMontant,20)
     +Format_String(TypeEcr,1)+AlignDroite(NumPiece,7); //+AlignDroite(StSoldeLettr,20) ;

if FFormat.Value='RAP' then St:=St+AlignDroite(StSoldeLettr,20)+RefPointage ;

if ChoixFmt.Ascii then St:=Ansi2Ascii(St) ;
END ;


procedure TFExpMvt.AjouteZonesHal(Q : TQuery ; var St : String ; P : Char) ;
var ZonesHalley  : String ;
    DecDev : integer ;
    Quotite: Double ;
    Code : String3 ;

BEGIN
ZonesHalley:='  '; //DecDev:=2 ; Quotite:=1 ;
if P='E' then
  BEGIN
  Code:=Q.FindField(P+'_DEVISE').Value ;
  RecupDevise(Code,DecDev,Quotite,TDev) ;
  if (Q.FindField(P+'_NUMECHE').AsInteger>=1) then ZonesHalley:=FormatFloat('00',Q.FindField(P+'_NUMECHE').AsInteger) ;
  ZonesHalley:=ZonesHalley+Format_String(Q.FindField(P+'_VALIDE').AsString,1)+Format_String(Q.FindField(P+'_REFEXTERNE').AsString,35)+Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATEREFEXTERNE').AsDateTime))
  +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATECREATION').AsDateTime))+Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATEMODIF').AsDateTime))+Format_String(Q.FindField(P+'_SOCIETE').AsString,3)
  +Format_String(Q.FindField(P+'_ETABLISSEMENT').AsString,3)+Format_String(Q.FindField(P+'_AFFAIRE').AsString,17)+AlignDroite(StrfMontant(0,20,2,'',False),20)
  +AlignDroite(StrfMontant(0,20,2,'',False),20)
  +AlignDroite(StrfMontant(V_PGI.TauxEuro,20,V_PGI.OkDecE,'',False),20)+Format_String(Q.FindField(P+'_DEVISE').AsString,3)+AlignDroite(StrfMontant(Q.FindField(P+'_DEBITDEV').AsFloat,20,DecDev,'',False),20)
  +AlignDroite(StrfMontant(Q.FindField(P+'_CREDITDEV').AsFloat,20,DecDev ,'',False),20)+AlignDroite(StrfMontant(Q.FindField(P+'_TAUXDEV').AsFloat,20,DecDev,'',False),20)+Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATETAUXDEV').AsDateTime))
  +AlignDroite(StrfMontant(Quotite,20,DecDev ,'',False),20)+Format_String(Q.FindField(P+'_ECRANOUVEAU').AsString,3)+AlignDroite(StrfMontant(Q.FindField(P+'_QTE1').AsFloat,20,4,'',False),20)
  +AlignDroite(StrfMontant(Q.FindField(P+'_QTE2').AsFloat,20,4,'',False),20)
  +Format_String(Q.FindField(P+'_QUALIFQTE1').AsString,3)+Format_String(Q.FindField(P+'_QUALIFQTE2').AsString,3)
  +Format_String(Q.FindField(P+'_REFLIBRE').AsString,35)+Format_String(Q.FindField(P+'_TVAENCAISSEMENT').AsString,1)+Format_String(Q.FindField(P+'_REGIMETVA').AsString,3)
  +Format_String(Q.FindField(P+'_TVA').AsString,3)+Format_String(Q.FindField(P+'_TPF').AsString,3)
  +Format_String(Q.FindField(P+'_CONTREPARTIEGEN').AsString,17)+Format_String(Q.FindField(P+'_CONTREPARTIEAUX').AsString,17)+AlignDroite(StrfMontant(Q.FindField(P+'_COUVERTURE').AsFloat,20,V_PGI.OkDecV,'',False),20)
  +Format_String(Q.FindField(P+'_LETTRAGE').AsString,5)+Format_String(Q.FindField(P+'_LETTRAGEDEV').AsString,1)+Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATEPAQUETMAX').AsDateTime))
  +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATEPAQUETMIN').AsDateTime))+Format_String(Q.FindField(P+'_REFPOINTAGE').AsString,17)+Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATEPOINTAGE').AsDateTime))
  +Format_String(' ',4)
  //+Format_String(Q.FindField(P+'_LETTREPOINTLCR').AsString,4)
  +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATERELANCE').AsDateTime))
  +Format_String(Q.FindField(P+'_CONTROLE').AsString,1)
  +AlignDroite(StrfMontant(0,20,V_PGI.OkDecV,'',False),20)
  +AlignDroite(StrfMontant(0,20,V_PGI.OkDecV,'',False),20)
  //+AlignDroite(StrfMontant(Q.FindField(P+'_TOTALTVAENC').AsFloat,20,V_PGI.OkDecV,'',False),20)
  //+AlignDroite(StrfMontant(Q.FindField(P+'_RELIQUATTVAENC').AsFloat,20,V_PGI.OkDecV,'',False),20)
  +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATEVALEUR').AsDateTime))+Format_String(Q.FindField(P+'_RIB').AsString,35)+Format_String(Q.FindField(P+'_REFRELEVE').AsString,35)
  +AlignDroite(StrfMontant(Q.FindField(P+'_COUVERTUREDEV').AsFloat,20,DecDev,'',False),20)+Format_String(Q.FindField(P+'_ETATLETTRAGE').AsString,3)+Format_String(Q.FindField(P+'_NUMPIECEINTERNE').AsString,35)
  +Format_String(Q.FindField(P+'_ENCAISSEMENT').AsString,3)+Format_String(Q.FindField(P+'_TYPEANOUVEAU').AsString,3)+Format_String(Q.FindField(P+'_ECHE').AsString,1)+Format_String(Q.FindField(P+'_ANA').AsString,1)
  +Format_String(Q.FindField(P+'_MODEPAIE').AsString,3)+Format_String(Q.FindField(P+'_NATUREPIECE').AsString,3) ;
  END else
if P='Y' then
  BEGIN
  Code:=Q.FindField(P+'_DEVISE').Value ;
  RecupDevise(Code,DecDev,Quotite,TDev) ;
  ZonesHalley:=Format_String(Q.FindField(P+'_AXE').AsString,2)
  +Format_String(Q.FindField(P+'_VALIDE').AsString,1)+Format_String(Q.FindField(P+'_REFEXTERNE').AsString,35)+Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATEREFEXTERNE').AsDateTime))
  +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATECREATION').AsDateTime))+Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATEMODIF').AsDateTime))+Format_String(Q.FindField(P+'_SOCIETE').AsString,3)
  +Format_String(Q.FindField(P+'_ETABLISSEMENT').AsString,3)+Format_String(Q.FindField(P+'_AFFAIRE').AsString,17)
  +AlignDroite(StrfMontant(0,20,2,'',False),20)+AlignDroite(StrfMontant(0,20,2,'',False),20)+AlignDroite(StrfMontant(V_PGI.TauxEuro,20,V_PGI.OkDecE,'',False),20)
  +Format_String(Q.FindField(P+'_DEVISE').AsString,3)+AlignDroite(StrfMontant(Q.FindField(P+'_DEBITDEV').AsFloat,20,DecDev,'',False),20)
  +AlignDroite(StrfMontant(Q.FindField(P+'_CREDITDEV').AsFloat,20,DecDev,'',False),20)+AlignDroite(StrfMontant(Q.FindField(P+'_TAUXDEV').AsFloat,20,DecDev,'',False),20)
  +Fmt_Date_HAL(DateToStr(Q.FindField(P+'_DATETAUXDEV').AsDateTime))
  +AlignDroite(StrfMontant(Quotite,20,DecDev ,'',False),20)+Format_String(Q.FindField(P+'_ECRANOUVEAU').AsString,3)
  +AlignDroite(StrfMontant(Q.FindField(P+'_QTE1').AsFloat,20,4,'',False),20)+AlignDroite(StrfMontant(Q.FindField(P+'_QTE2').AsFloat,20,4,'',False),20)
  +Format_String(Q.FindField(P+'_QUALIFQTE1').AsString,3)+Format_String(Q.FindField(P+'_QUALIFQTE2').AsString,3)
  +AlignDroite(StrfMontant(Q.FindField(P+'_POURCENTAGE').AsFloat,20,4,'',False),20)
  //+AlignDroite(StrfMontant(Q.FindField(P+'_TOTALECRITURE').AsFloat,20,2,'',False),20)
  +Format_String(Q.FindField(P+'_CONTROLE').AsString,1)+AlignDroite(IntToStr(Q.FindField(P+'_NUMVENTIL').AsInteger),6)
  +AlignDroite(StrfMontant(Q.FindField(P+'_POURCENTQTE1').AsFloat,20,4,'',False),20)
  +AlignDroite(StrfMontant(Q.FindField(P+'_POURCENTQTE2').AsFloat,20,4,'',False),20)
  {+AlignDroite(StrfMontant(Q.FindField(P+'_TOTALDEVISE').AsFloat,20,DecDev,'',False),20)
  +AlignDroite(StrfMontant(Q.FindField(P+'_TOTALEURO').AsFloat,20,2,'',False),20)
  +AlignDroite(StrfMontant(Q.FindField(P+'_TOTALQTE1').AsFloat,20,4,'',False),20)
  +AlignDroite(StrfMontant(Q.FindField(P+'_TOTALQTE2').AsFloat,20,4,'',False),20)
  }
  +Format_String(Q.FindField(P+'_TYPEANOUVEAU').AsString,3)+Format_String(Q.FindField(P+'_TYPEMVT').AsString,3)
  //+Format_String(Q.FindField(P+'_QUALIFECRQTE1').AsString,3)+Format_String(Q.FindField(P+'_QUALIFECRQTE2').AsString,3)
  +Format_String(Q.FindField(P+'_TYPEANALYTIQUE').AsString,1) ;
  END ;
St:=St+ZonesHalley ;
END ;

{================ écriture de l'entête de fichier EDI ====================}

procedure TFExpMvt.LanceExportEcrEDI ;
var FTMP : TextFile ;
    St : String ;
    Fic : String ;
begin
If not (Dossier.OkSauve) then InitDossier('-',Lequel) ;
Fic:=FileName.Text ;
AssignFile(Fichier,Fic) ;
{$I-} ReWrite(Fichier) ; {$I+}
EnteteMvt ;
DetailMvt ;
ResumeMvt ;
CloseFile(Fichier) ;
// Conversion Ascii...
If ChoixFmt.Ascii then
  BEGIN
  AssignFile(Fichier,Fic) ;
  {$I-} Reset(Fichier) ; {$I+}
  AssignFile(FTMP,ExpandFileName(ExtractFileName(Fic)+'.$$$')) ;
  {$I-} ReWrite(FTMP) ; {$I+}
  While Not EOF(Fichier) do
    BEGIN
    Readln(Fichier,St) ;
    St:=ANSI2ASCII(St) ;
    WriteLn(FTMP,St) ;
    END ;
  CloseFile(Fichier) ;
  CloseFile(FTMP) ;
  Erase(Fichier) ;
  RenameFile(ExpandFileName(ExtractFileName(Fic)+'.$$$'),Fic) ;
  END ;
StTotCred:=MTF(CPTotExtract+CNTotExtract,20,2) ;
StTotDeb:=MTF(DPTotExtract+DNTotExtract,20,2) ;
END ;

procedure TFExpMvt.EnteteMvt ;
BEGIN
With Dossier do
  BEGIN
  Writeln(Fichier, '00010'+'UNOC'+'1'+Format_String(EditEmet,35)+'OEC '+
                   Format_String(EditAdrEmet,14)+Format_String(EditRec,35)+
                   'OEC '+Format_String(EditAdrRec,14)+FormatDateTime('yymmddhhnn',NowH)+
                   Format_String(EditRefIntChg,14)+CBoxAccRec,CBoxEssai) ;
  Writeln(Fichier, '00020'+'ENTREC'+Format_String(EditEmet,35)+'OEC '+
                   Format_String(EditRec,35)+'OEC '+FormatDateTime('yymmddhhnn',NowH)+
                   Format_String(EditRefDem,17)+'RT'+'  2'+'  0') ;

  Writeln(Fichier, '00030'+Format_String(EditRefIntChg,14)+'ENTREC'+
                 '  2'+'  1'+'RT'+'E03'+'CPT'+'PEE'+Format_String(EditIdentEnv,17)+Blanc(3)+'DFP'+
                   FormatDateTime('yyyymmdd',StrToDate(FDateEcr1.Text))+
                   FormatDateTime('yyyymmdd',StrToDate(FDateEcr2.Text))+'242'+
                   Format_String(FormatDateTime('yyyymmddhhnn',NowH),16)+'203') ;//+FILLER{5719}) ;
  Writeln(Fichier, '00110'+'ACF'+Format_String(DEditSiret,17)+'100'+'107'+
                   Format_String(DEditPerChrg,35)+Format_String(DEditRue1,35)+
                   Format_String(DEditRue2,35)+Format_String(DEditVille,35)+
                   Format_String(DEditDiv,9)+Format_String(DEditCodPost,9)+
                   Format_String(DEditPays,2)) ;
  Writeln(Fichier, '00115'+'ACF'+Format_String(DEditFctCont,3)+Format_String(DEditPerChrg,35)+
                   Format_String(DEditNumCom,25)+Format_String(DEditModCom,3)+'AAA'+
                   Format_String(EditRefDem,35));
  Writeln(Fichier, '00130'+'ACS'+Format_String(EEditSiret,17)+'100'+'107'+
                   Format_String(EEditPerChrg,35)+Format_String(EEditRue1,35)+
                   Format_String(EEditRue2,35)+Format_String(EEditVille,35)+
                   Format_String(EEditDiv,9)+Format_String(EEditCodPost,9)+
                   Format_String(EEditPays,2)) ;
  Writeln(Fichier, '00135'+'ACS'+Format_String(EEditFctCont,3)+Format_String(EEditPerChrg,35)+
                   Format_String(EEditNumCom,25)+Format_String(EEditModCom,3)+'ASV'+
                   Format_String(EditNumDosEmt,35)) ;
  Writeln(Fichier, '00150'+'ACR'+Format_String(REditSiret,17)+'100'+'107'+
                   Format_String(RComboPerChrg,35)+Format_String(REditRue1,35)+
                   Format_String(REditRue2,35)+Format_String(REditVille,35)+
                   Format_String(REditDiv,9)+Format_String(REditCodPost,9)+
                   Format_String(REditPays,2)) ;
  Writeln(Fichier, '00155'+'ACR'+Format_String(REditFctCont,3)+Format_String(RComboPerChrg,35)+
                   Format_String(REditNumCom,25)+Format_String(REditModCom,3)+'ARV'+
                   Format_String(EditNumDosRec,35)) ;
  //Writeln(Fichier, '00190'+);
  END ;
END ;

{================== écriture du corps du fichier EDI =====================}

procedure TFExpMvt.DetailMvt ;
BEGIN
QtiteExtract:=0 ; DPTotExtract:=0 ; DNTotExtract:=0 ; CPTotExtract:=0 ; CNTotExtract:=0 ;
ExtractionGene ;

Writeln(Fichier, '00450'+
{DTM} FormatDateTime('yyyymmdd',StrToDate(FDateEcr1.Text))+
      FormatDateTime('yyyymmdd',StrToDate(FDateEcr2.Text))+
{QTY}   ' 01'+NBS(QtiteJournaux,15,' ')+'RJL'+
{MOA}   'TDP'+MTF(DPTotJournaux,19,2)+'FRF'+'DT'+
{MOA}   'TDN'+MTF(Abs(DNTotJournaux),19,2)+'FRF'+'CT'+
{MOA}   'TCP'+MTF(CPTotJournaux,19,2)+'FRF'+'DT'+
{MOA}   'TCN'+MTF(Abs(CNTotJournaux),19,2)+'FRF'+'CT') ;
(* fin 00450 *)

DPTotExtract:=DPTotExtract+DPTotJournaux ;
DNTotExtract:=DNTotExtract+DNTotJournaux ;
CPTotExtract:=CPTotExtract+CPTotJournaux ;
CNTotExtract:=CNTotExtract+CNTotJournaux ;
Inc(QtiteExtract) ;

Writeln(Fichier, '00480'+
{QTY}   ' 01'+NBS(QtiteExtract,15,' ')+'DTM'+
{MOA}   'TDP'+MTF(DPTotExtract,20,2)+'FRF'+'DT'+
{MOA}   'TDN'+MTF(Abs(DNTotExtract),20,2)+'FRF'+'CT'+
{MOA}   'TCP'+MTF(CPTotExtract,20,2)+'FRF'+'DT'+
{MOA}   'TCN'+MTF(Abs(CNTotExtract),20,2)+'FRF'+'CT') ;

END;

procedure TFExpMvt.ExtractionGene ;
var NatureJal     : String3 ;
    Q1,QQ : TQuery ;
    StSQL1 : String ;
BEGIN
NbEcr:=0 ; NbPiece:=0 ; NbMvt:=0 ;
QtiteJournaux:=0 ; DPTotJournaux:=0 ; DNTotJournaux:=0 ; CPTotJournaux:=0 ; CNTotJournaux:=0 ;


Writeln(Fichier, '00200'+
{DTM}   'DFP'+FormatDateTime('yyyymmdd',StrToDate(FDateEcr1.Text))+
        FormatDateTime('yyyymmdd',StrToDate(FDateEcr2.Text))+'718') ;

StSQL1:=ReqSQL('E',False,FALSE,NIL) ;
//ConstruitRequete('E',False,FALSE) ;
Q1:=OpenSQL(StSQL1,TRUE) ;
if Q1.Eof then MsgBox.Execute(4,caption,'')
          Else InitMove(RecordsCount(Q1)+1,'');
(*
if (QEcr.EOF) then BEGIN MsgBox.Execute(4,caption,'') ; Exit ; END ;
InitMove(RecordsCount(QEcr)+1,'') ;MoveCur(False) ;
*)
While Not Q1.EOF do
  BEGIN
  CodeJournal:=Q1.FindField('E_JOURNAL').AsString ;

  if Arreter then Break ;
  Inc(NbEcr) ;
  CodeJournal:=Q1.FindField('E_JOURNAL').AsString ;
  Inc(QtiteJournaux) ;
  QQ:=OpenSQL('select J_LIBELLE,J_NATUREJAL from JOURNAL where J_JOURNAL = "'+CodeJournal+'"',TRUE) ;
  (*
  Query.Close ;
  Query.sql.Clear ;
  Query.sql.add('select J_LIBELLE,J_NATUREJAL from JOURNAL '+
                'where J_JOURNAL = "'+CodeJournal+'"');
  ChangeSQL(Query) ;
  Query.Open ;
  *)
  NatureJal:=QS(QQ.FindField('J_NATUREJAL').asString,3);
  NatureJal:='' ;
  if NatureJal ='VTE' then NatureJal:='VEN' else
    if (NatureJal ='BQE') or (NatureJal ='CAI')  then NatureJal:='TRE' else
      if (NatureJal ='ANO') then NatureJal:='SAN' else
        if NatureJal <>'ACH' then NatureJal:='OPD' ;
  Writeln(Fichier, '00250'+CodeJournal+Blanc(3)+'ZZZ'+Blanc(3)+QS(QQ.FindField('J_LIBELLE').asString,35)+Format_string(NatureJal,6)+'TED'+'PEE') ;

  Journal(Q1) ;

  Writeln(Fichier, '00400'+
{RJL}  CodeJournal+Blanc(3)+
{QTY}     ' 01'+NBS(QtiteJournal,15,' ')+'ENT'+
{MOA}     'TDP'+MTF(DPTotJournal,18,2)+'FRF'+'DT'+
{MOA}     'TDN'+MTF(Abs(DNTotJournal),18,2)+'FRF'+'CT'+
{MOA}     'TCP'+MTF(CPTotJournal,18,2)+'FRF'+'DT'+
{MOA}     'TCN'+MTF(Abs(CNTotJournal),18,2)+'FRF'+'CT') ;
(* fin 00400 *)

  DPTotJournaux:=DPTotJournaux+DPTotJournal ;
  DNTotJournaux:=DNTotJournaux+DNTotJournal ;
  CPTotJournaux:=CPTotJournaux+CPTotJournal ;
  CNTotJournaux:=CNTotJournaux+CNTotJournal ;
  END ;
Ferme(Q1) ;
//ConstruitRequete('E',True,FALSE) ;
FiniMove ;
END ;

{================Recencement des écritures par journal ===================}

procedure TFExpMvt.Journal(Q1 : tQuery) ;
var MonnaieRef, QualifPiece,Sens,Etat,TvaEnc  : String3 ;
    Montant : String ;
BEGIN
MonnaieRef:='FRF' ;
if (Q1.findField('E_QUALIFPIECE').asString='S') or (Q1.findField('E_QUALIFPIECE').asString='U') or (Q1.findField('E_QUALIFPIECE').asString='P') then QualifPiece:='SIM' else
  if (Q1.findField('E_QUALIFPIECE').asString='N') then
    if (Q1.findField('E_VALIDE').asString='X') then QualifPiece:='VAL' else QualifPiece:='NVL' ;
QtiteJournal:=0 ; DPTotJournal:=0 ; DNTotJournal:=0 ; CPTotJournal:=0 ; CNTotJournal:=0 ;
While (CodeJournal=Q1.FindField('E_JOURNAL').AsString) and (not Q1.EOF) do
  BEGIN
  if TestBreak(Q1) then Break ;
  Inc(QtiteJournal) ; Inc(NbMvt) ;
  if (Q1.FindField('E_DEBIT').asFloat<>0) then
           BEGIN
           if (Q1.FindField('E_DEBIT').asFloat>=0) then
             BEGIN
             DPTotJournal:=DPTotJournal+Q1.FindField('E_DEBIT').asFloat ;
             Montant:=MTF(Q1.FindField('E_DEBIT').AsFloat,17,2);
             Sens:='DP' ;
             END else
             BEGIN
             DNTotJournal:=DNTotJournal+Q1.FindField('E_DEBIT').asFloat ;
             Montant:=MTF(Abs(Q1.FindField('E_DEBIT').AsFloat),17,2);
             Sens:='DN' ;
             END ;
           END else
           BEGIN
           if (Q1.FindField('E_CREDIT').asFloat>=0) then
             BEGIN
             CPTotJournal:=CPTotJournal+Q1.FindField('E_CREDIT').asFloat ;
             Montant:=MTF(Q1.FindField('E_CREDIT').AsFloat,17,2);
             Sens:='CP' ;
             END else
             BEGIN
             CNTotJournal:=CNTotJournal+Q1.FindField('E_CREDIT').asFloat ;
             Montant:=MTF(Abs(Q1.Findfield('E_CREDIT').AsFloat),17,2);
             Sens:='CN' ;
             END ;
           END ;
  Writeln(Fichier, '00285'+NBS(Q1.findField('E_NUMEROPIECE').AsInteger,12,' ')
         +QI(Q1.FindField('E_NUMLIGNE').asInteger,'000000')+'000001'+'ECG'+QS(Q1.FindField('E_GENERAL').AsString,17)+' PC'+'PEE'+'ECD' (*ou ECA si analytique*)

         +QS(Q1.FindField('E_AUXILIAIRE').AsString,17)+' PC'+'PEE'
         +'SEC'+QualifPiece+' PC'+'PEE'+'DVC'+QD(Q1.FindField('E_DATECOMPTABLE').AsDateTime,'yyyymmdd')+Blanc(8)+'102'
         +'DPJ'+QD(Q1.FindField('E_DATEREFEXTERNE').AsDateTime,'yyyymmdd')+Blanc(8)+'102'
         +'265'+QD(Q1.FindField('E_DATEECHEANCE').AsDateTime,'yyyymmdd')+Blanc(8)+'102'
         +'MEC'+Montant+MonnaieRef+'003'+Sens

{RFF}    +'RPC'+QS(Q1.FindField('E_REFEXTERNE').AsString,35)+'RPL'+QS(Q1.FindField('E_LETTRAGE').asString,5)
{RFF}    +'RDP'+QS(Q1.FindField('E_REFPOINTAGE').AsString,17)) ;

(* Fin 00285 *)

  if (Q1.findField('E_ETAT').asString='IM') then Etat:='IMP' else Etat:='MAN' ;
  if (Q1.findField('E_TVAENCAISSEMENT').asString='X') then TvaEnc:='TSE' else TvaEnc:='TSD' ;

{FLE}  Writeln(Fichier, '00287'+NBS(Q1.findField('E_NUMEROPIECE').asInteger,12,' ')+Blanc(6)+Blanc(3)+Blanc(3)+
{QTE}  MTF(Q1.FindField('E_QTE1').asFloat,15,3)+QS(Q1.FindField('E_QUALIFQTE1').AsString,3)+
{TAX}  Blanc(3)+{TVA à faire}Blanc(7)+
{AFC}  'ORE'+Etat+' PC'+'PEE'+'TVA'+TvaEnc+' PC'+'PEE'+
{RFF}  'RIP'+QS(Q1.FindField('E_REFINTERNE').asString,17)+' FC'+Blanc(17)) ;

(* Fin 00287 *)

  Writeln(Fichier, '00370'+
  {FLE}   NBS(Q1.findField('E_NUMEROPIECE').asInteger,12,' ')+QI(Q1.FindField('E_NUMLIGNE').asInteger,'000000')+Blanc(6)+
  {RFF}   'ECR'+QS(Q1.FindField('E_REFINTERNE').asString,17)+
  {FTX}   'ACC'+'ACC'+QS(Q1.FindField('E_NATUREPIECE').AsString,3)+'LAC'+'PEE'+
          QS(Q1.FindField('E_LIBELLE').AsString,70)) ;
(* Fin 00370 *)
  Q1.Next ;
  MoveCur(False) ;
  END ;
if Q1.Eof then Inc(NbPiece) ;
END ;

{================= écriture de la fin de fichier EDI =====================}

procedure TFExpMvt.ResumeMvt ;
BEGIN
With Dossier do
  Writeln(Fichier, '00800'+'ACC'+Format_String(CombFctTxt,3)+blanc(3)+
        Format_String(EditRefTxt,35)+Format_String(EditTxt,35)+Blanc(143)) ;
END;

procedure TFExpMvt.BEDIClick(Sender: TObject);
begin ParametresEDI('-',Lequel); End ;

procedure TFExpMvt.PFenGuideMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
if (ssleft in shift) then SourisX:=X ;SourisY:=Y ;
end;

procedure TFExpMvt.PFenGuideMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
if not (ssleft in shift) then Exit ;
if (PListe.Left+(X-SourisX)>=ClientWidth-5) then X:=SourisX ;
if ((PListe.Left+(X-SourisX)+PListe.Width)=10) then X:=SourisX ;
if (PListe.Top+(Y-SourisY)>=ClientHeight-5) then Y:=SourisY ;
if ((PListe.Top+(Y-SourisY)+PListe.Height)=10) then Y:=SourisY ;
PListe.Left:=PListe.Left+(X-SourisX) ;
PListe.Top:=PListe.Top+(Y-SourisY) ;
end;

{============================= Exportation de la BALANCE SISCO=====================================}
Function StrToInt1(St : String) : Integer ;
Var i : Integer ;
BEGIN
i:=Pos(V_PGI.SepMillier,St) ;
While i>0 Do BEGIN Delete(St,i,1) ; i:=Pos(V_PGI.SepMillier,St) ; END ;
Result:=StrToInt(St) ;
END ;

Function TFExpMvt.CritSISCOOk(Var CritEdt : TCritEdt) : Integer ;
BEGIN
Result:=0 ;
If CritEdt.CptLibre11='' Then
  BEGIN
  Result:=41 ; Pages.ActivePage:=PSisco ; //FCompteResultatSISCO.SetFocus ;
  END ;
END ;

procedure TFExpMvt.ExporteBalSISCO ;
var CritEdt : TCritEdt ;
    Err : Integer ;
BEGIN
Fillchar(CritEdt,SizeOf(CritEdt),#0) ;
With CritEdt Do
  BEGIN
  DateDeb:=StrToDate(FDateEcr1.Text) ; DateFin:=StrToDate(FDateEcr2.Text) ;
  QuelExoDate(DateDeb,DateDeb,MonoExo,CritEdt.Exo) ;
  If FEtab.ItemIndex>0 Then Etab:=FEtab.Value ;
  GL.NumPiece1:=StrToInt1(NUMPIECE1.Text) ;
  GL.NumPiece2:=StrToInt1(NUMPIECE2.Text) ;
  Cpt1:=FNatJal1.Value ; Cpt2:=FNatJal2.Value ;
  GL.Sauf:=SaufJal.Text ;
  CritEdt.Decimale:=V_PGI.OkDecV ;
  CritEdt.CptLibre1:=Trim(FBen.Text) ; CritEdt.CptLibre2:=Trim(FPerte.Text) ;
  CritEdt.CptLibre11:=Trim(FCOmpteResultatSISCO.Text) ;
  If VH^.TenueEuro Then CritEdt.Monnaie:=2 Else CritEdt.Monnaie:=0 ;
  END ;
Err:=CritSISCOOk(CritEdt) ;
If Err=0 Then
  BEGIN
  ExportBALSISCO(FileName.Text,CritEdt,FTypeEcr.Value,AvecRevision.State) ; MsgBox.Execute(42,Caption,'') ;
  END Else MsgBox.Execute(Err,Caption,'') ;
END ;

{============================= Exportation de la BALANCE =====================================}

procedure TFExpMvt.ExporteBal ;
var Jal,St :String ;
    i : integer ;
    FTMP : TextFile ;
    First,Ok : boolean ;
BEGIN
if (JalAN.Text='') or (JALN.Text='') then
  BEGIN
  Pages.ActivePage:=PBalance ;
  MsgBox.Execute(33,caption,'') ;
  Exit ;
  END ;
try
  AssignFile(Fichier,FileName.Text) ;
  {$I-} Rewrite(Fichier) ; {$I+}
  St:=FirstLigneFic(FFormat.Value,Lequel) ;
  Writeln(Fichier,St) ;
  DetailBal ;
  finally
  CloseFile(Fichier) ;
  end ;

AssignFile(FTMP,ExpandFileName(ExtractFileName(FileName.Text)+'.$$$')) ;
{$I-} ReWrite(FTMP) ; {$I+}
for i:=1 to 2 do
  BEGIN
  AssignFile(Fichier,FileName.Text) ;
  {$I-} Reset(Fichier) ; {$I+}
  Readln(Fichier,St) ;
  if i=1 then WriteLn(FTMP,St) ;
  Jal:='' ; First:=True ;
  While Not EOF(Fichier) do
    BEGIN
    Readln(Fichier,St) ;
    if First then Jal:=Copy(St,1,3) ;
    Ok:=((Jal=Copy(St,1,3)) and (i=1)) or ((Jal<>Copy(St,1,3)) and (i=2)) ;
    if Ok then WriteLn(FTMP,St) ;
    First:=False ;
    END ;
  CloseFile(Fichier) ;
  END ;
CloseFile(FTMP) ;
Erase(Fichier) ;
RenameFile(ExpandFileName(ExtractFileName(FileName.Text)+'.$$$'),FileName.Text) ;
END ;

{========================= Exportation de la BALANCE EDI ==============================}

procedure TFExpMvt.LanceExportBalEDI ;
var FTMP : TextFile ;
    St : String ;
begin
If not (Dossier.OkSauve) then InitDossier('-',Lequel) ;
AssignFile(Fichier,FileName.Text) ;
{$I-} Rewrite(Fichier) ; {$I+}
EnteteBal ; DetailBal ; ResumeBal ;
CloseFile(Fichier) ;
// Conversion Ascii...
If ChoixFmt.Ascii then
  BEGIN
  AssignFile(Fichier,FileName.Text) ; {$I-} Reset(Fichier) ; {$I+}
  AssignFile(FTMP,ExpandFileName(ExtractFileName(FileName.Text)+'.$$$')) ;
  {$I-} ReWrite(FTMP) ; {$I+}
  While Not EOF(Fichier) do BEGIN Readln(Fichier,St) ; St:=ANSI2ASCII(St) ; WriteLn(FTMP,St) ; END ;
  CloseFile(Fichier) ; CloseFile(FTMP) ; Erase(Fichier) ;
  RenameFile(ExpandFileName(ExtractFileName(FileName.Text)+'.$$$'),FileName.Text) ;
  END ;
END ;

{================ écriture de l'entête de fichier EDI ====================}

procedure TFExpMvt.EnteteBal ;
BEGIN
With Dossier do
  BEGIN
  Writeln(Fichier, '01010'+'UNOC'+'1'+Format_String(EditEmet,35)+'OEC '+
                   Format_String(EditAdrEmet,14)+Format_String(EditRec,35)+
                   'OEC '+Format_String(EditAdrRec,14)+FormatDateTime('yymmddhhnn',NowH)+
                   Format_String(EditRefIntChg,14)+CBoxAccRec,CBoxEssai) ;
  {UNH} Writeln(Fichier, '01030'+Format_String(EditRefIntChg,14)+'BALANC'+
                 '  2'+'  1'+'RT'+'BCG'+'BAL'+'PEE'+Format_String(EditIdentEnv,17)+Blanc(3)+'DPB'+
                   FormatDateTime('yyyymmdd',StrToDate(FDateEcr1.Text))+
                   FormatDateTime('yyyymmdd',StrToDate(FDateEcr2.Text))+'718'+'242'+
                   Format_String(FormatDateTime('yyyymmddhhnn',NowH),16)+'203') ;//+FILLER{5719}) ;
        Writeln(Fichier, '01110'+'ACF'+Format_String(DEditSiret,17)+'100'+'107'+
                   Format_String(DEditPerChrg,35)+Format_String(DEditRue1,35)+
                   Format_String(DEditRue2,35)+Format_String(DEditVille,35)+
                   Format_String(DEditDiv,9)+Format_String(DEditCodPost,9)+
                   Format_String(DEditPays,2)) ;
        DEditModCom:='TE' ;
        Writeln(Fichier, '01115'+'ACF'+Format_String(DEditFctCont,3)+Format_String(DEditPerChrg,35)+
                   Format_String(DEditNumCom,25)+Format_String(DEditModCom,3)+'AAA'+
                   Format_String(EditRefDem,35));
        Writeln(Fichier, '01130'+'ACS'+Format_String(EEditSiret,17)+'100'+'107'+
                   Format_String(EEditPerChrg,35)+Format_String(EEditRue1,35)+
                   Format_String(EEditRue2,35)+Format_String(EEditVille,35)+
                   Format_String(EEditDiv,9)+Format_String(EEditCodPost,9)+
                   Format_String(EEditPays,2)) ;
        EEditModCom:='TE' ;
        Writeln(Fichier, '01135'+'ACS'+Format_String(EEditFctCont,3)+Format_String(EEditPerChrg,35)+
                   Format_String(EEditNumCom,25)+Format_String(EEditModCom,3)+'ASV'+
                   Format_String(EditNumDosEmt,35)) ;
        Writeln(Fichier, '01150'+'ACR'+Format_String(REditSiret,17)+'100'+'107'+
                   Format_String(RComboPerChrg,35)+Format_String(REditRue1,35)+
                   Format_String(REditRue2,35)+Format_String(REditVille,35)+
                   Format_String(REditDiv,9)+Format_String(REditCodPost,9)+
                   Format_String(REditPays,2)) ;
        REditModCom:='TE' ;
        Writeln(Fichier, '01155'+'ACR'+Format_String(REditFctCont,3)+Format_String(RComboPerChrg,35)+
                   Format_String(REditNumCom,25)+Format_String(REditModCom,3)+'ARV'+
                   Format_String(EditNumDosRec,35)) ;
  END ;
END ;

{============== écriture du corps du fichier compte par compte =====================}

procedure TFExpMvt.DetailBal ;
var St : String ;
    OkFirst : boolean ;
BEGIN
QtiteExtract:=0 ; SFDExtract:=0 ; SFCExtract:=0 ; SNDExtract:=0 ; SNCExtract:=0 ;
QtiteComptes:=0 ; DPTotComptes:=0 ; DNTotComptes:=0 ; CPTotComptes:=0 ; CNTotComptes:=0 ;
NbEcr:=0 ;
OkFirst:=False ;
Query.Close ;
Query.RequestLive:=False ;
Query.sql.Clear ;
St:='SELECT DISTINCT G_GENERAL,G_COLLECTIF, G_CENTRALISABLE, G_LIBELLE, G_VENTILABLE, G_VENTILABLE1, '
   +'G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5 FROM GENERAUX Q WHERE G_GENERAL<>"'+W_W+'" ' ;
if (CPTEDEBUT.Text<>'') then BEGIN St:=St+' AND G_GENERAL>="'+CPTEDEBUT.Text+'"' ; OkFirst:=True ; END ;
if (CPTEFIN.Text<>'') then BEGIN if OkFirst then St:=St+' AND ' ; St:=St+'G_GENERAL<="'+CPTEFIN.Text+'"' ; END ;
Query.SQL.Add(St) ;
St:=WhereSQL ;
if (St<>'') then Query.sql.Add(' AND '+St) ;
Query.sql.add('ORDER BY G_GENERAL') ;
ChangeSQL(Query) ;
Query.Open ;
NbEcr:=0 ;
if (Query.EOF) then BEGIN (*CloseFile(Fichier) ;*) MsgBox.Execute(4,caption,'') ; Exit ; END ;
InitMove(RecordsCount(Query)+1,'') ;MoveCur(False) ;
CalcTotCpte ;
FiniMove ;
If QBalC<>NIL Then QBalC.Free ;
GCalc.Free ;
END;


{================ Totaux/Soldes des Comptes ===================}

procedure TFExpMvt.CalcTotCpte ;
var Coll : String3 ;
    Montant : Double ;
    TypeCpte,Sens: String1 ;
    Q : TQuery ;
    DateEche,Ana,AuxiOuAna : String ;
    i: integer ;
    CatPaie,ModeRegle : String ;
    MPaie : String3 ;
BEGIN
NbEcr:=0 ; QBalC:=nil ; GCalc:=nil ;
While not Query.Eof do
  BEGIN
  Application.ProcessMessages ; Inc(NbEcr) ;
  if Arreter then if MsgBox.Execute(21,Caption,'')<>mryes then Arreter:=False else Break ;
  Inc(QtiteComptes) ;
  CompteGene:=QS(Query.FindField('G_GENERAL').asString,17) ;
  if (Query.FindField('G_COLLECTIF').AsString='X') then Coll:='CCO' else Coll:='CNC' ;
  DPTotCompte:=0 ; DNTotCompte:=0 ; CPTotCompte:=0 ; CNTotCompte:=0 ;
  SNDCompte:=0 ; SNCCompte:=0 ; SFDCompte:=0 ; SFCCompte:=0 ;
  if FFormat.Value='EDI' then
     Writeln(Fichier, '01200'+
   {SEQ} NBS(QtiteComptes,6,' ')+
   {CPT} 'ECG'+CompteGene+'PC '+'PEE'+QS(Query.FindField('G_LIBELLE').AsString,17)+Blanc(3)+
   {AFC} Coll+'PC '+'PEE') ;

  CalculSoldes ;

  if FFormat.Value='EDI' then
     BEGIN
     Writeln(Fichier, '01220'+CompteGene+Blanc(17)+Blanc(17)+Blanc(17)+
   {MOA} 'MDP'+MTF(DPTotCompte,19,2)+'FRF'+'DT'+
   {MOA} 'MCP'+MTF(CPTotCompte,19,2)+'FRF'+'CT'+
   {MOA} 'MDN'+MTF(Abs(DNTotCompte),19,2)+'FRF'+'DT'+
   {MOA} 'MCN'+MTF(Abs(CNTotCompte),19,2)+'FRF'+'CT') ;
     Writeln(Fichier, '01230'+CompteGene+Blanc(17)+Blanc(17)+Blanc(17)+
   {MOA} 'SFD'+MTF(SFDCompte,19,2)+'FRF'+'DT'+
   {MOA} 'SFC'+MTF(SFCCompte,19,2)+'FRF'+'CT'+
   {MOA} 'SND'+MTF(SNDCompte,19,2)+'FRF'+'DT'+
   {MOA} 'SNC'+MTF(SNCCompte,19,2)+'FRF'+'CT') ;
     END else
     BEGIN
     TypeCpte:=' ' ; AuxiOuAna:='' ; MPaie:=' ' ; ModeRegle:='' ;
     Ana:='  ' ;
     if (Coll='CCO') then //and (Query.FindField('G_CENTRALISABLE').AsString='X') then
       BEGIN
{$IFDEF SPEC302}
       Q:=OpenSQL('Select SO_CLIATTEND,SO_FOUATTEND FROM SOCIETE',True) ;
       if Not Q.EOF then
         if Copy(CompteGene,1,2)='41' then AuxiOuAna:=Q.Fields[0].AsString else AuxiOuAna:=Q.Fields[1].AsString ;
       TypeCpte:='X' ;
       Ferme(Q) ;
{$ELSE}
       if Copy(CompteGene,1,2)='41' then AuxiOuAna:=GetParamSoc('SO_CLIATTEND') else AuxiOuAna:=GetParamSoc('SO_FOUATTEND') ;
       TypeCpte:='X' ;
{$ENDIF}
       if AuxiOuAna<>'' then
         BEGIN
         Q:=OpenSQL('SELECT T_MODEREGLE FROM TIERS WHERE T_AUXILIAIRE="'+AuxiOuAna+'"',True) ;
         if not Q.Eof then ModeRegle:=Q.Fields[0].AsString ;
         Ferme(Q) ;
         MPaie:=FirstModePaie(ModeRegle) ;
         CatPaie:=ChercheCategPaie(MPaie) ;
         MPaie:=CHERCHEUNMODE(CatPaie,MPaie,False,False) ;
         END ;
       END ;
     //DateEche:=DateToStr(iDate1900) ;
     //if Query.FindField('G_LETTRABLE').AsString='X' then ;
     DateEche:=FDateEcr1.Text ;
     // Ecritures générales
     // A nouveau
     Montant:=(SNDCompte-SNCCompte) ;
     if Montant<>0 then
       BEGIN
       if (Montant>=0) then Sens:='D' else Sens:='C' ;
       Writeln(Fichier,Format_String(JALAN.Text,3)+Fmt_Date_HAL(FDateEcr1.text)+'OD'+Format_String(CompteGene,17)+TypeCpte         +Format_String(AuxiOuAna,17)+Format_String(RefInterne.Text,13)+Format_String(Libelle.Text,25)
         +' '+Fmt_Date_HAL(DateEche)+Sens+MTF(Abs(Montant),19,2)+Copy(FTypeEcr.Value,1,1)+'      1'+Ana) ;
       END ;
     // courantes
     //if Query.FindField('G_LETTRABLE').AsString='X' then ;
     DateEche:=FDateEcr2.Text ;
     Montant:=(DPTotCompte-CNTotCompte)-(CPTotCompte-DNTotCompte) ;
     if Montant<>0 then
       BEGIN
       if (Montant>=0) then Sens:='D' else Sens:='C' ;
       Writeln(Fichier,Format_String(JALN.Text,3)+Fmt_Date_HAL(FDateEcr2.text)+'OD'+Format_String(CompteGene,17)+TypeCpte
         +Format_String(AuxiOuAna,17)+Format_String(RefInterne.Text,13)+Format_String(Libelle.Text,25)
         +' '+Fmt_Date_HAL(DateEche)+Sens+MTF(Abs(Montant),19,2)+Copy(FTypeEcr.Value,1,1)+'      1'+Ana) ;
       END ;

     // Ecritures analytiques
     if (Query.FindField('G_VENTILABLE').AsString='X') then
       BEGIN
       TypeCpte:='A' ; Ana:='A1' ; AuxiOuAna:=SectAttente[1] ; MPaie:=' ' ;
       for i:=2 to 5 do
         if (Query.FindField('G_VENTILABLE'+IntToStr(i)).AsString='X') then
           BEGIN
           AuxiOuAna:=SectAttente[i] ;
           Ana:='A'+IntToStr(i) ;
           Break ;
           END ;
       // A nouveau
       Montant:=(SNDCompte-SNCCompte) ;
       if Montant<>0 then
         BEGIN
         if (Montant>=0) then Sens:='D' else Sens:='C' ;
         Writeln(Fichier,Format_String(JALAN.Text,3)+Fmt_Date_HAL(FDateEcr1.text)+'OD'+Format_String(CompteGene,17)+TypeCpte
           +Format_String(AuxiOuAna,17)+Format_String(RefInterne.Text,13)+Format_String(Libelle.Text,25)
           +' '+Fmt_Date_HAL(DateToStr(iDate1900))+Sens+MTF(Abs(Montant),19,2)+Copy(FTypeEcr.Value,1,1)+'      1'+Ana) ;
         END ;
       // courantes
       Montant:=(DPTotCompte-CNTotCompte)-(CPTotCompte-DNTotCompte) ;
       if Montant<>0 then
         BEGIN
         if (Montant>=0) then Sens:='D' else Sens:='C' ;
         Writeln(Fichier,Format_String(JALN.Text,3)+Fmt_Date_HAL(FDateEcr2.text)+'OD'+Format_String(CompteGene,17)+TypeCpte
           +Format_String(AuxiOuAna,17)+Format_String(RefInterne.Text,13)+Format_String(Libelle.Text,25)
           +' '+Fmt_Date_HAL(FDateEcr2.Text)+Sens+MTF(Abs(Montant),19,2)+Copy(FTypeEcr.Value,1,1)+'      1'+Ana) ;
         END ;
       END ;
     END ;
  DPTotComptes:=DPTotComptes+DPTotCompte ;
  DNTotComptes:=DNTotComptes+DNTotCompte ;
  CPTotComptes:=CPTotComptes+CPTotCompte ;
  CNTotComptes:=CNTotComptes+CNTotCompte ;

  SNDExtract:=SNDExtract+SNDCompte ;
  SNCExtract:=SNCExtract+SNCCompte ;
  SFDExtract:=SFDExtract+SFDCompte ;
  SFCExtract:=SFCExtract+SFCCompte ;

  Query.Next ;
  MoveCur(False) ;
  END ;
END ;

{================= écriture de la fin de fichier EDI =====================}

procedure TFExpMvt.ResumeBal ;
BEGIN
Writeln(Fichier, '01475'+'S') ;
Writeln(Fichier, '01480'+
{QTY}   ' 01'+NBS(QtiteComptes,15,' ')+Blanc(3)+
{MOA}   'TDP'+MTF(DPTotComptes,20,2)+'FRF'+'DT'+
{MOA}   'TCP'+MTF(CPTotComptes,20,2)+'FRF'+'CT'+
{MOA}   'TDN'+MTF(Abs(DNTotComptes),20,2)+'FRF'+'DT'+
{MOA}   'TCN'+MTF(Abs(CNTotComptes),20,2)+'FRF'+'CT'+

{MOA}   'TFD'+MTF(SFDExtract,20,2)+'FRF'+'DT'+
{MOA}   'TFC'+MTF(SFCExtract,20,2)+'FRF'+'CT'+
{MOA}   'TND'+MTF(SNDExtract,20,2)+'FRF'+'DT'+
{MOA}   'TNC'+MTF(SNCExtract,20,2)+'FRF'+'CT') ;
With Dossier do
  Writeln(Fichier, '01800'+'ACC'+Format_String(CombFctTxt,3)+blanc(3)+
        Format_String(EditRefTxt,35)+Format_String(EditTxt,35)+Blanc(143)) ;
END;

{================= Calcul des soldes du compte en cours =====================}

Function QuelTypeCpt(Valeur : String) : Integer ;
BEGIN
Result:=0 ;
If Valeur='EXO' Then Result:=0 Else If Valeur='FOU' Then Result:=1 Else
If Valeur='NSL' Then Result:=2 Else If Valeur='PER' Then Result:=3 ;
END ;

function TFExpMvt.FQuelAN(Inclure : Boolean ; Date1,Date11 : TDateTime) : TQuelAN ;
Var LTypCpt : Integer ;
     MonoExo,OkExo : Boolean ;
BEGIN
Result:=AvecAN ;
MonoExo:=(Exercice.ItemIndex>=0) ;
LTypCpt:=QuelTypeCpt(FSelectCpte.Value) ;
if not MonoExo then LTypCpt:=3 ;
OkExo:=MonoExo ;
OkExo:=OkExo And (Date1=Date11) ;
//OkExo:=OkExo And (Date1=Date11) ;
If (LTypCpt=3) And (Not OkExo) And (Not Inclure) Then Result:=SansAN ;
END ;

procedure TFExpMvt.CalculSoldes ;
var TotCpt1,TotCpt2,
    TotCpt1M,TotCpt1S,
    TotCpt2M,TotCpt2S,Tot   : TabTot ;
    //TotCpt1Euro,TotCpt2Euro,TotEuro,
    TotCpt1MEuro,TotCpt1SEuro,
    TotCpt2MEuro,TotCpt2SEuro   : TabTot ;
    Date1,Date2,Date11,Date21,DateDeb,DateFin : TDateTime ;
    Exo : TExoDate ;
    Sommation,PrColl : Boolean ;
    QuelAN : TQuelAN ;
    QuelTypeEcr : SetttTypePiece ;
BEGIN
Sommation:=True ;
RecupExo(Exo,Date1,Date2,Date11,Date21,DateDeb,DateFin) ;
QuelAN:=FQuelAN(True,Date1,Date11) ;
PrColl:=FCollectif.Checked ;
QuelTypeEcr:=WhatTypeEcr(FTypeEcr.Value,V_PGI.Controleur,AvecRevision.State) ;
if QBalC=nil then QBalC:=PrepareTotCptSolde(QuelTypeEcr,False,False,True,FALSE) ;
if GCalc=nil then
  BEGIN
  GCalc:=TGCalculCum.Create(Un,fbGene,fbGene,QuelTypeEcr,False,False,True,False,False,V_PGI.OkDecV,V_PGI.OkDecE) ;
  GCalc.InitCalcul('','','','','','',Exo.Code,
                        Date1,Date11,TRUE) ;
  END ;
if (Query.FindField('G_COLLECTIF').AsString='X') and PrColl Then
      BEGIN
      If QuelAN<>SansAN Then
         BEGIN
         ExecuteTotCptSolde(QBalC,Trim(CompteGene), Date1,Date1,
                   '','',Exo.Code,TotCpt1M,TotCpt1S,TotCpt1MEURO,TotCpt1SEURO,FALSE,Sommation,V_PGI.OkDecV,V_PGI.OkDecE,FALSE,FALSE,QuelTypeEcr) ;
         GCAlc.ReInitCalcul(CompteGene,'',Date1,Date11) ;
         GCalc.Calcul ; TotCpt1:=GCalc.ExecCalc.TotCpt ;
         END ;

      GCAlc.ReInitCalcul(Trim(CompteGene),'',Date21,Date2) ;
      GCalc.Calcul ; TotCpt2:=GCalc.ExecCalc.TotCpt ;

      ExecuteTotCptSolde(QBalC,Trim(CompteGene), Date1,Date2,
                         '','',Exo.Code,TotCpt2M,TotCpt2S,TotCpt2MEURO,TotCpt2SEURO,TRUE,Sommation,V_PGI.OkDecV,V_PGI.OkDecE,FALSE,FALSE,QuelTypeEcr) ;
      TotCpt1[0]:=TotCpt1S[0] ;
      END Else
      BEGIN
      If QuelAN<>SansAN Then
         BEGIN
         GCAlc.ReInitCalcul(Trim(CompteGene),'',Date1,Date11) ;
         GCalc.Calcul ; TotCpt1:=GCalc.ExecCalc.TotCpt ;
         END ;
      GCAlc.ReInitCalcul(Trim(CompteGene),'',Date21,Date2) ;
      GCalc.Calcul ; TotCpt2:=GCalc.ExecCalc.TotCpt ;
      CumulVersSolde(TotCpt1[0]) ;
      END ;
   if Date1=Date11 Then Fillchar(TotCpt1[1],SizeOf(TotCpt1[1]),#0) ;
   If QuelAN=SansAN Then
      BEGIN
      TotCpt1[0].TotDebit:=0 ; TotCpt1[0].TotCredit:=0 ;
      TotCpt1[1].TotDebit:=0 ; TotCpt1[1].TotCredit:=0 ;
      END ;
if Date1=Date11 Then Fillchar(TotCpt1[1],SizeOf(TotCpt1[1]),#0) ;
if (FFormat.Value='EDI') then
  BEGIN
  Tot[1].TotDebit:= Arrondi(TotCpt1[0].TotDebit+TotCpt1[1].TotDebit,2) ;
  Tot[1].TotCredit:=Arrondi(TotCpt1[0].TotCredit+TotCpt1[1].TotCredit,2) ;
  Tot[2].TotDebit:= TotCpt2[1].TotDebit ;
  Tot[2].TotCredit:=TotCpt2[1].TotCredit ;
  Tot[3].TotDebit:= Arrondi(Tot[2].TotDebit+Tot[1].TotDebit,2) ;
  Tot[3].TotCredit:=Arrondi(Tot[2].TotCredit+Tot[1].TotCredit,2) ;
  END else
  BEGIN
  CumulVersSolde(TotCpt2[0]) ;
  if ((Query.FindField('G_COLLECTIF').AsString='X') and PrColl) then TotCpt1[0]:=TotCpt1S[0] ;
  Tot[1].TotDebit:= Arrondi(TotCpt1[0].TotDebit,2) ;
  Tot[1].TotCredit:=Arrondi(TotCpt1[0].TotCredit,2) ;
  Tot[2].TotDebit:= TotCpt2[1].TotDebit+TotCpt1[1].TotDebit ;
  Tot[2].TotCredit:=TotCpt2[1].TotCredit+TotCpt1[1].TotCredit ;
  Tot[3].TotDebit:= Arrondi(Tot[2].TotDebit+Tot[1].TotDebit,2) ;
  Tot[3].TotCredit:=Arrondi(Tot[2].TotCredit+Tot[1].TotCredit,2) ;
  END ;
SNDCompte:=Tot[1].TotDebit ;
SNCCompte:=Tot[1].TotCredit ;

if (Query.FindField('G_COLLECTIF').AsString='X') and PrColl Then
   BEGIN
   Tot[4]:=TotCpt2S[1] ;
   CumulVersSolde(TotCpt2S[1]) ;
   END Else
   BEGIN
   Tot[4]:=Tot[3] ;
   CumulVersSolde(Tot[4]) ;
   END ;

SFDCompte:=Tot[4].TotDebit ;
SFCCompte:=Tot[4].TotCredit ;

If ((TotCpt2[1].TotDebit)>0) then DPTotCompte:=(TotCpt2[1].TotDebit)
   else DNTotCompte:=Abs(TotCpt2[1].TotDebit) ;
If (TotCpt2[1].TotCredit>0) then CPTotCompte:=(TotCpt2[1].TotCredit)
   else CNTotCompte:=Abs(TotCpt2[1].TotCredit) ;

END ;

procedure TFExpMvt.RecupExo(Var Exo : TExoDate ; var Date1,Date2,Date11,Date21,DateDeb,DateFin : TDateTime) ;
var MonoExo : boolean ;
BEGIN
Date1:=StrToDate(FDateEcr1.Text) ; Date2:=StrToDate(FDateEcr2.Text) ;
DateDeb:=Date1 ; DateFin:=Date2 ;
MonoExo:=(Exercice.ItemIndex>=0) ;
if MonoExo then
   BEGIN
   Exo.Code:=Exercice.Value ;
   QuelExoDate(Date1,Date2,MonoExo,Exo) ;
   END ;
IntervalleDateBalGL1(MonoExo,Exo,Date1,Date11,Date21) ;
{
if V_PGI.EnCours.Deb=Date1 then
   BEGIN
   Date11:=StrToDate(FDateEcr1.Text) ; Date21:=Date11 ;
   END else
   BEGIN
   Date1:=V_PGI.EnCours.Deb ; Date11:=StrToDate(FDateEcr1.Text)-1 ; Date21:=Date11+1 ;
   END ;
}
END ;


function SQLExiste(PreE,PreJ,PreB,TOTDEBIT,TOTCREDIT,Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState ;
                   LDate1,LDate2 : TDateTime ; CodeExo : String) : String ;
Var St : String ;
BEGIN
St:='' ;
if Valeur='NOR' then
   BEGIN
   St:='('+TOTDEBIT+'<>0 OR '+TOTCREDIT+'<>0 OR '+PreB+'TOTDEBANO<>0 OR '+PreB+'TOTCREANO<>0) ';
   If V_PGI.Controleur Then
      St:='('+St+'Or '+WhatExiste(PreE,PreJ,PreB,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,1,'')+')' ;
   END ;
if Valeur='NSS' then
   BEGIN
   St:='(('+TOTDEBIT+'<>0 OR '+TOTCREDIT+'<>0 OR '+PreB+'TOTDEBANO<>0 OR '+PreB+'TOTCREANO<>0) ';
   St:=St+'Or '+WhatExiste(PreE,PreJ,PreB,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,1,'')+')' ;
   END ;
if Valeur='SSI' then
   BEGIN
   St:=WhatExiste(PreE,PreJ,PreB,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,1,'') ;
   END ;
if Valeur='PRE' then
   BEGIN
   St:=WhatExiste(PreE,PreJ,PreB,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,1,'') ;
   END ;
Result:=St ;
END ;

function TFExpMvt.WhereSQL : String;
Var TOTDEBIT,TOTCREDIT,St : String ;
    PreB,PreE,PreJ : String ;
    Exo : TExoDate ;
    ts : Integer ;
    me : Boolean ;
    Valeur : String ;
    Date1,Date2,Date11,Date21 : TDateTime ;
BEGIN
PreB:='G_' ; PreJ:='GENERAL' ;
PreE:='E_' ;
TOTDEBIT:=PreB+'TOTDEBE' ; TOTCREDIT:=PreB+'TOTCREE' ;
RecupExo(Exo,Date1,Date2,Date11,Date21,DateDeb,DateFin) ;
me:=(Exercice.ItemIndex>=0) ;
Ts:=3 ;   // Mouvementés sur la période...
If Me Then
   BEGIN
   If Exo.Code=VH^.Precedent.Code Then BEGIN TOTDEBIT:=PreB+'TOTDEBP' ; TOTCREDIT:=PreB+'TOTCREP' ; END ;
   If Exo.Code=VH^.EnCours.Code   Then BEGIN TOTDEBIT:=PreB+'TOTDEBE' ; TOTCREDIT:=PreB+'TOTCREE' ; END ;
   If Exo.Code=VH^.Suivant.Code   Then BEGIN TOTDEBIT:=PreB+'TOTDEBS' ; TOTCREDIT:=PreB+'TOTCRES' ; END ;
   END ;
Valeur:=FTypeEcr.Value ;
//Type ttTypePiece  = (tpReel,tpSim,tpPrev,tpSitu,tpRevi) ;
//TS = 0 : Exo, 1 = Tous, 2 = Cpt non soldés, 3 = Période
Case TS of
  0 : BEGIN
      if (Valeur<>'TOU') Then
         BEGIN
         St:=SQLExiste(PreE,PreJ,PreB,TOTDEBIT,TOTCREDIT,Valeur,V_PGI.Controleur,AvecRevision.State,Date1,Date2,Exo.Code) ;
         END Else
         BEGIN
         St:=WhatExiste(PreE,PreJ,PreB,Valeur,V_PGI.Controleur,AvecRevision.State,Date1,Date2,Exo.Code,1,'') ;
         END ;
      Result:=St ;
      END ;
  1 : BEGIN Result:=PreB+PreJ+'='+PreB+PreJ ; END ;
  2 : BEGIN
      if (Valeur<>'TOU') Then
         BEGIN
         if Valeur='NOR' then BEGIN Result:=PreB+'TOTALDEBIT<>'+PreB+'TOTALCREDIT ' ; END ;
         if Valeur='NSS' then
            BEGIN
            Result:=PreB+'TOTALDEBIT<>'+PreB+'TOTALCREDIT ' ;
            Result:=Result+' Or (('+PreB+'TOTALDEBIT='+PreB+'TOTALCREDIT '+') And (' ;
            Result:=Result+WhatExiste(PreE,PreJ,PreB,Valeur,V_PGI.Controleur,AvecRevision.State,Date1,Date2,Exo.Code,1,'')+'))' ;
            END ;
         END Else
         BEGIN
         Result:=PreB+'TOTALDEBIT<>'+PreB+'TOTALCREDIT ' ;
         Result:=Result+' Or (('+PreB+'TOTALDEBIT='+PreB+'TOTALCREDIT '+') And (' ;
         Result:=Result+WhatExiste(PreE,PreJ,PreB,Valeur,V_PGI.Controleur,AvecRevision.State,Date1,Date2,Exo.Code,1,'')+'))' ;
         END ;
      END ;
  3 : BEGIN
      if (Valeur<>'TOU') Then
         BEGIN
         St:=SQLExiste(PreE,PreJ,PreB,TOTDEBIT,TOTCREDIT,Valeur,V_PGI.Controleur,AvecRevision.State,DateDeb,DateFin,Exo.Code) ;
         END Else
         BEGIN
         St:=WhatExiste(PreE,PreJ,PreB,Valeur,V_PGI.Controleur,AvecRevision.State,DateDeb,DateFin,Exo.Code,1,'') ;
         END ;
      Result:=St ;
      END ;
 end ;
END;

procedure TFExpMvt.ZoneVisible(BalSISCO : Boolean);
Var i : Integer ;
    OkEcr : Boolean ;
BEGIN
OkEcr:=(Lequel='FEC') or (Lequel='FOD') or (Lequel='FBE') ;
for i:=0 to PGene.ControlCount-1 do
  BEGIN
  Case PGene.Controls[i].Tag of
    1  : TControl(PGene.Controls[i]).Visible:=(Lequel='FBA') ;
    10 : TControl(PGene.Controls[i]).Visible:=OkEcr Or BalSISCO ;
    END ;
  END ;
SiscoDefaut ;
END ;

procedure TFExpMvt.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
TJal.Free ; VideListe(TDev) ; TDev.Free ;
ListeNumJal.Clear ; ListeNumJal.Free ;
end;

procedure TFExpMvt.VisiblePourBalance ;
var OkEcr : boolean ;
    Q : TQuery ;
BEGIN
OkEcr:=(Lequel='FEC') or (Lequel='FOD') or (Lequel='FBE') ;
FExport.Visible:= OkEcr ;
if OkEcr and (Not V_PGI.Controleur) then
  BEGIN
  FTypeEcr.Items.Clear ;
  FTypeEcr.Values.Clear ;
  FTypeEcr.Items.add(Traduirememoire('<<Tous>>')) ; FTypeEcr.Values.add('') ;
  Q:=OpenSQL('SELECT CO_CODE,CO_LIBELLE FROM COMMUN WHERE CO_TYPE="QFP" AND CO_CODE<>"R"',True) ;
  While not Q.Eof do
    BEGIN
    FTypeEcr.Values.Add(Q.FindField('CO_CODE').AsString) ;
    FTypeEcr.Items.Add(Q.FindField('CO_LIBELLE').AsString) ;
    Q.Next ;
    END ;
  Ferme(Q) ;
  END else if (Lequel='FBA') then BEGIN FTypeEcr.Vide:=False ; FTypeEcr.DataType:='ttQualPieceCrit' ; FTypeEcr.ItemIndex:=0 ; FTypeEcr.Width:=233 ; END ;
ZoneVisible((FFormat.Value='SIS') And (Lequel='FBA')) ;
if AvecRevision.Visible then AvecRevision.Visible:=V_PGI.Controleur ;
if AvecRevision.Visible then
  BEGIN
  AvecRevision.Checked:=TRUE ; AvecRevision.State:=cbGrayed ;
  END Else
  BEGIN
  AvecRevision.Checked:=FALSE ; AvecRevision.State:=cbUnchecked ;
  END ;
END ;

{================= Exportation du Lettrage des mouvements =====================}

procedure TFExpMvt.ExporteLettre(ListeMvtL : TStringList) ;
var Nb,i : integer ;
    St,SQL : String ;
    QL : TQuery ;
    BlocML : TMonoLigne ;
    Aux,Gen,CodeLettre : String ;
BEGIN
If ListeMvtL=Nil Then Exit ; If ListeMvtL.Count<=0 Then Exit ; If Lequel<>'FEC' Then Exit ;
Application.ProcessMessages ;
//QL:=TQuery.Create(Application) ; QL.DataBaseName:='SOC' ;
SQL:='SELECT * FROM ECRITURE WHERE E_GENERAL=:GENE'
    +' AND E_AUXILIAIRE=:AUXI AND E_LETTRAGE=:LETT'
    +' ORDER BY E_AUXILIAIRE, E_GENERAL, E_ETATLETTRAGE, E_LETTRAGE, E_DATECOMPTABLE, E_DATEECHEANCE, E_NUMECHE' ;
QL:=PrepareSQL(SQL,TRUE) ;
//PourVisu:=True ; ConstruitRequete('L',False,FALSE) ; PourVisu:=False ;
Fillchar(BlocML,SizeOf(BlocML),#0) ;
AssignFile(Fichier,FileName.Text) ;
{$I-} Append(Fichier) ; {$I+}
Nb:=ListeMvtL.Count-1 ;
InitMove(Nb,'') ;
MoveCur(False) ;
//QEcr.First ;
For i:=0  To ListeMvtL.Count-1 Do
  BEGIN
  Aux:=Trim(Copy(ListeMvtL[i],1,17)) ;
  Gen:=Trim(Copy(ListeMvtL[i],18,17)) ;
  CodeLettre:=Trim(Copy(ListeMvtL[i],35,4)) ;
  QL.Params[0].AsString:=Gen ;
  QL.Params[1].AsString:=Aux ;
  QL.Params[2].AsString:=CodeLettre ;
  QL.Open ;
  While not QL.Eof do
    BEGIN
    EcrireLigne(QL,'E',St,BlocML,NIL) ;
    If (FFormat.Value='CGN') Or (FFormat.Value='CGE') Then St[31]:='L' Else St[27]:='L' ;
    Writeln(Fichier,St) ;
    QL.Next ;
    END ;
  QL.Close ;
  MoveCur(False) ;
  END ;
FiniMove ;
Ferme(QL) ;
CloseFile(Fichier) ;
END ;

procedure TFExpMvt.BZoomClick(Sender: TObject);
begin
Exit ;
(*
If QEcr.State<>dsInactive Then
   BEGIN
   QEcr.First ;
//   if RecordsCount(QEcr)>0 then
   If Not QEcr.EOF Then
     BEGIN
//     ListeExportes(QEcr,Prefixe) ;
     //VisuPiecesGenere(TPiece,EcrGen,11) ;
     END ;
   END ;
*)
end;

function TFExpMvt.AjouteInfosCompte(Q : TQuery ; P : String) : String ;
var TGGene : TGGeneral ;
    TGSect : TGSection ;
    S,St : String ;
    Lett,Point,Coll : String1 ;
    Ventil  : String[2] ;
    QAux : TQuery ;
    Dom,Etab,Guichet,NumCpt,Cle : String ;
begin
exit ;
Result:='' ;
if CBInfosComptes.State=cbUnchecked then Exit ;
Lett:='-' ; Ventil:='  ' ; Point:='-' ; Coll:='-' ; St:='' ;
if (Q.FindField(P+'_GENERAL').AsString<>'') then
  BEGIN
  TGGene:=TGGeneral.Create(Q.FindField(P+'_GENERAL').AsString) ;
  if TGGene.Lettrable then Lett:='X' ;
  if TGGene.Pointable then Point:='X' ;
  if TGGene.Collectif then Coll:='X' ;
  if TGGene.Ventilable[1] then Ventil:='A1' else
   if TGGene.Ventilable[1] then Ventil:='A2' else
    if TGGene.Ventilable[1] then Ventil:='A3' else
     if TGGene.Ventilable[1] then Ventil:='A4' else
      if TGGene.Ventilable[1] then Ventil:='A5' ;
  S:=SepLigneIE+Format_String(TGGene.Libelle,35)+Format_String(TGGene.NatureGene,3)+Lett+Coll+Point+Ventil ;
  St:=St+S ;
  TGGene.Free ;
  END else St:=St+SepLigneIE+Format_String(' ',43) ;
if P='E' then
  BEGIN
  if (Q.FindField(P+'_AUXILIAIRE').AsString<>'') then
    BEGIN
    QAux:=OpenSQL('SELECT * FROM TIERS WHERE T_AUXILIAIRE="'+Q.FindField(P+'_AUXILIAIRE').AsString+'"',True) ;
    if QAux.Eof then BEGIN St:=St+SepLigneIE+Format_String(' ',52)+Format_String(' ',149)+SepLigneIE+Format_String(' ',47)+SepLigneIE+Format_String(' ',60) ; Ferme(QAux) ; Exit ; END ;
    St:=St+SepLigneIE+Format_String(QAux.FindField('T_LIBELLE').AsString,35)+Format_String(QAux.FindField('T_NATUREAUXI').AsString,3)+QAux.FindField('T_LETTRABLE').AsString+Format_String(QAux.FindField('T_COLLECTIF').AsString,17)
          +Format_String(QAux.FindField('T_ADRESSE1').AsString,35)+Format_String(QAux.FindField('T_ADRESSE2').AsString,35)+Format_String(QAux.FindField('T_ADRESSE3').AsString,35)
          +Format_String(QAux.FindField('T_CODEPOSTAL').AsString,9)+Format_String(QAux.FindField('T_VILLE').AsString,35) ;
    Ferme(QAux) ;
    ChercheRIBAux(Q.FindField(P+'_AUXILIAIRE').AsString,Dom,Etab,Guichet,NumCpt,Cle) ;
    St:=St+SepLigneIE+Format_String(Dom,24)
          +Format_String(Etab,5)+Format_String(Guichet,5)
          +Format_String(NumCpt,11)+Format_String(Cle,2) ;
    Ferme(QAux) ;
    QAux:=OpenSQL('SELECT * FROM CONTACT WHERE C_AUXILIAIRE="'+Q.FindField(P+'_AUXILIAIRE').AsString+'" AND C_PRINCIPAL="X"',True) ;
    if QAux.Eof then St:=St+SepLigneIE+Format_String(' ',60)
                else St:=St+SepLigneIE+Format_String(QAux.FindField('C_NOM').AsString,35)+Format_String(QAux.FindField('C_TELEPHONE').AsString,25) ;
    Ferme(QAux) ;
    END else St:=St+SepLigneIE+Format_String(' ',52)+Format_String(' ',149)+SepLigneIE+Format_String(' ',47)+SepLigneIE+Format_String(' ',60) ;
  END else
  if (Q.FindField(P+'_SECTION').AsString<>'') then
    BEGIN
    TGSect:=TGSection.Create(Q.FindField(P+'_SECTION').AsString,AxeToTz(Q.FindField(P+'_AXE').AsString)) ;
    St:=St+SepLigneIE+Format_String(TGSect.Libelle,35)+Format_String(' ',6)+TGSect.Axe+Format_String(' ',9)+Format_String(' ',149)+SepLigneIE+Format_String(' ',47)+SepLigneIE+Format_String(' ',60) ;
    TGSect.Free ;
    END else St:=St+SepLigneIE+Format_String(' ',52)+Format_String(' ',149)+SepLigneIE+Format_String(' ',47)+SepLigneIE+Format_String(' ',60) ;
Result:=St ;
end;

procedure TFExpMvt.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFExpMvt.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FreeAndNil(ObjFiltre); //SG6 10/11/04 Gestion des filtres FQ 14826
ListeCptCegid.Free ;
if IsInside(Self) then Action:=caFree ;
end;

procedure TFExpMvt.FormCreate(Sender: TObject);
begin
ListeCptCegid:=TStringList.Create ;

//SG6 10/11/04 Gestion des Filtres FQ 14826
Composants.PopupF   := POPF;
Composants.Filtres  := FFILTRES;
Composants.Filtre   := BFILTRE;
Composants.PageCtrl := Pages;
ObjFiltre := TObjFiltre.Create(Composants,'');

end;

procedure TFExpMvt.FNatureGeneChange(Sender: TObject);
begin
If FNatureGene.ItemIndex=0 Then BEGIN CPTEDEBUT.ZoomTable:=tzGeneral      ; CPTEFIN.ZoomTable:=tzGeneral      ; END Else
If FNatureGene.Value='BQE' Then BEGIN CPTEDEBUT.ZoomTable:=tzGbanque      ; CPTEFIN.ZoomTable:=tzGbanque      ; END Else
If FNatureGene.Value='CAI' Then BEGIN CPTEDEBUT.ZoomTable:=tzGcaisse      ; CPTEFIN.ZoomTable:=tzGcaisse      ; END Else
If FNatureGene.Value='CHA' Then BEGIN CPTEDEBUT.ZoomTable:=tzGcharge      ; CPTEFIN.ZoomTable:=tzGcharge      ; END Else
If FNatureGene.Value='COC' Then BEGIN CPTEDEBUT.ZoomTable:=TzGCollClient  ; CPTEFIN.ZoomTable:=TzGCollClient  ; END Else
If FNatureGene.Value='COD' Then BEGIN CPTEDEBUT.ZoomTable:=tzGCollDivers  ; CPTEFIN.ZoomTable:=tzGCollDivers  ; END Else
If FNatureGene.Value='COF' Then BEGIN CPTEDEBUT.ZoomTable:=tzGCollFourn   ; CPTEFIN.ZoomTable:=tzGCollFourn   ; END Else
If FNatureGene.Value='COS' Then BEGIN CPTEDEBUT.ZoomTable:=tzGCollSalarie ; CPTEFIN.ZoomTable:=tzGCollSalarie ; END Else
If FNatureGene.Value='DIV' Then BEGIN CPTEDEBUT.ZoomTable:=tzGDivers      ; CPTEFIN.ZoomTable:=tzGDivers      ; END Else
If FNatureGene.Value='EXT' Then BEGIN CPTEDEBUT.ZoomTable:=tzGextra       ; CPTEFIN.ZoomTable:=tzGextra       ; END Else
If FNatureGene.Value='IMO' Then BEGIN CPTEDEBUT.ZoomTable:=tzGimmo        ; CPTEFIN.ZoomTable:=tzGimmo        ; END Else
If FNatureGene.Value='PRO' Then BEGIN CPTEDEBUT.ZoomTable:=tzGproduit     ; CPTEFIN.ZoomTable:=tzGproduit     ; END Else
If FNatureGene.Value='TIC' Then BEGIN CPTEDEBUT.ZoomTable:=tzGTIC         ; CPTEFIN.ZoomTable:=tzGTIC         ; END Else
If FNatureGene.Value='TID' Then BEGIN CPTEDEBUT.ZoomTable:=tzGTID         ; CPTEFIN.ZoomTable:=tzGTID         ; END ;
end;

procedure TFExpMvt.SCENARIOEXPORTKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if ((SCENARIOEXPORT.Focused) and (Key=VK_F6)) Then
  BEGIN
  Key:=0 ;
  ParamSupExport('-','FEC',FFormat.Value,ScenarioExport.Text,tamodif,0,TRUE);
  END ;

end;

procedure TFExpMvt.SCENARIOEXPORTChange(Sender: TObject);
begin
Exit ;
FFormatChange(Nil) ;
end;

procedure TFExpMvt.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  //Sg6 13/01/05 FQ 15242
  if key = VK_ESCAPE then BFermeClick(nil);
end;

end.

