{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 31/03/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par BALBUDECSEG_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit QRBLecSG;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Hctrls, Spin, Menus, hmsgbox, HQuickrp, StdCtrls,
  Buttons, ExtCtrls, Mask, Hcompte, ComCtrls, UtilEdt, HEnt1, Ent1,
  CpteUtil, EdtLegal, CritEDT, HQry, HSysMenu,UtilEdt1, QRRupt, Calcole, 
  Budsect, HTB97, HPanel, UiUtil ;

procedure BalBudecSecGen ;
procedure BalBudecSecGenRupture ;
procedure BalBudecSecGenNature ;
procedure BalBudecSecGenZoom(Crit : TCritEdt) ;

type
  TFQRBLecSG = class(TFQR)
    TitreColCpt: TQRLabel;
    Col1: TQRLabel;
    Col2: TQRLabel;
    col3: TQRLabel;
    Col4: TQRLabel;
    Col5: TQRLabel;
    Col6: TQRLabel;
    TITRE1REP: TQRLabel;
    REPORTCOL1: TQRLabel;
    REPORTCOL2: TQRLabel;
    REPORTCOL3: TQRLabel;
    REPORTCOL4: TQRLabel;
    REPORTCOL5: TQRLabel;
    REPORTCOL6: TQRLabel;
    BSecondaire: TQRBand;
    IniGenCum5: TQRLabel;
    IniGenCum4: TQRLabel;
    IniGenCum2: TQRLabel;
    IniGenCum3: TQRLabel;
    IniGenCum6: TQRLabel;
    IniGenCum1: TQRLabel;
    QRLabel33: TQRLabel;
    TotCum1: TQRLabel;
    TotCum2: TQRLabel;
    TotCum3: TQRLabel;
    TotCum4: TQRLabel;
    TotCum5: TQRLabel;
    TotCum6: TQRLabel;
    TITRE2REP: TQRLabel;
    REPORTCOL11: TQRLabel;
    REPORTCOL12: TQRLabel;
    REPORTCOL13: TQRLabel;
    REPORTCOL14: TQRLabel;
    REPORTCOL15: TQRLabel;
    REPORTCOL16: TQRLabel;
    REPORTCOL17: TQRLabel;
    MsgBox: THMsgBox;
    Bevel3: TBevel;
    BMulti: TQRBand;
    DLMulti: TQRDetailLink;
    TFJournal: THLabel;
    FJournal: THValComboBox;
    FPeriode1: THValComboBox;
    FPeriode2: THValComboBox;
    TFExercice_: TLabel;
    FExercice_: THValComboBox;
    FValide: TCheckBox;
    Col7: TQRLabel;
    Col8: TQRLabel;
    Col9: TQRLabel;
    Col10: TQRLabel;
    IniGenCum7: TQRLabel;
    IniGenCum8: TQRLabel;
    IniGenCum9: TQRLabel;
    IniGenCum10: TQRLabel;
    REPORTCOL7: TQRLabel;
    REPORTCOL8: TQRLabel;
    REPORTCOL9: TQRLabel;
    REPORTCOL10: TQRLabel;
    TotCum7: TQRLabel;
    TotCum8: TQRLabel;
    TotCum9: TQRLabel;
    TotCum10: TQRLabel;
    IniGen: TQRLabel;
    RevGen: TQRLabel;
    RevGenCum1: TQRLabel;
    RevGenCum2: TQRLabel;
    RevGenCum3: TQRLabel;
    RevGenCum4: TQRLabel;
    RevGenCum5: TQRLabel;
    RevGenCum6: TQRLabel;
    RevGenCum7: TQRLabel;
    RevGenCum8: TQRLabel;
    RevGenCum9: TQRLabel;
    RevGenCum10: TQRLabel;
    QIniGen: TQRLabel;
    QIniGenCum1: TQRLabel;
    QIniGenCum2: TQRLabel;
    QIniGenCum3: TQRLabel;
    QIniGenCum4: TQRLabel;
    QIniGenCum5: TQRLabel;
    QIniGenCum6: TQRLabel;
    QIniGenCum7: TQRLabel;
    QIniGenCum8: TQRLabel;
    QIniGenCum9: TQRLabel;
    QIniGenCum10: TQRLabel;
    QRevGen: TQRLabel;
    QRevGenCum1: TQRLabel;
    QRevGenCum2: TQRLabel;
    QRevGenCum3: TQRLabel;
    QRevGenCum4: TQRLabel;
    QRevGenCum5: TQRLabel;
    QRevGenCum6: TQRLabel;
    QRevGenCum7: TQRLabel;
    QRevGenCum8: TQRLabel;
    QRevGenCum9: TQRLabel;
    QRevGenCum10: TQRLabel;
    BS_BUDSECT: TQRLabel;
    TRJournal: TQRLabel;
    RJournal: TQRLabel;
    QRLabel3: TQRLabel;
    RExercice_: TQRLabel;
    QRLabel17: TQRLabel;
    RValide: TQRLabel;
    TFRESOL: THLabel;
    FRESOL: THValComboBox;
    TFNatBud: THLabel;
    FNatBud: THValComboBox;
    TFTaux: THLabel;
    TRNatBud: TQRLabel;
    RNatBud: TQRLabel;
    TRTaux: TQRLabel;
    RTaux: TQRLabel;
    REPORTCOL18: TQRLabel;
    REPORTCOL19: TQRLabel;
    REPORTCOL20: TQRLabel;
    BS_LIBELLE: TQRLabel;
    FQte: TCheckBox;
    QRLabel1: TQRLabel;
    RQte: TQRLabel;
    BFPrimaire: TQRBand;
    BS_BUDSECT_: TQRLabel;
    TotSec1: TQRLabel;
    TotSec2: TQRLabel;
    TotSec3: TQRLabel;
    TotSec4: TQRLabel;
    TotSec5: TQRLabel;
    TotSec6: TQRLabel;
    TotSec7: TQRLabel;
    TotSec8: TQRLabel;
    TotSec9: TQRLabel;
    TotSec10: TQRLabel;
    DLGen: TQRDetailLink;
    TFJokerGene: THLabel;
    TFGene: THLabel;
    TFaS: TLabel;
    FJokerGene: TEdit;
    FGene1: THCpteEdit;
    FGene2: THCpteEdit;
    QSecond: TQuery;
    SSecond: TDataSource;
    TRGeneral: TQRLabel;
    RGeneral1: TQRLabel;
    TRaG: TQRLabel;
    RGeneral2: TQRLabel;
    FLigneRupt: TCheckBox;
    FOnlySecRupt: TCheckBox;
    BHCpt: TQRBand;
    BG_BUDGENE: TQRDBText;
    BG_LIBELLE: TQRLabel;
    DLHGen: TQRDetailLink;
    IRGen: TQRLabel;
    IRGenCum1: TQRLabel;
    IRGenCum2: TQRLabel;
    IRGenCum3: TQRLabel;
    IRGenCum4: TQRLabel;
    IRGenCum5: TQRLabel;
    IRGenCum6: TQRLabel;
    IRGenCum7: TQRLabel;
    IRGenCum8: TQRLabel;
    IRGenCum9: TQRLabel;
    IRGenCum10: TQRLabel;
    FTotIniRev: TCheckBox;
    BRappelLibreHaut: TQRBand;
    TCodRupt: TQRLabel;
    TLibRupt: TQRLabel;
    DLHLibre: TQRDetailLink;
    IniLibre4: TQRLabel;
    IniLibre: TQRLabel;
    IniLibre5: TQRLabel;
    IniLibre6: TQRLabel;
    IniLibre1: TQRLabel;
    IniLibre2: TQRLabel;
    IniLibre3: TQRLabel;
    IniLibre7: TQRLabel;
    IniLibre8: TQRLabel;
    IniLibre9: TQRLabel;
    RevLibre4: TQRLabel;
    RevLibre: TQRLabel;
    RevLibre5: TQRLabel;
    RevLibre6: TQRLabel;
    RevLibre1: TQRLabel;
    RevLibre2: TQRLabel;
    RevLibre3: TQRLabel;
    RevLibre7: TQRLabel;
    RevLibre8: TQRLabel;
    RevLibre9: TQRLabel;
    QIniLibre4: TQRLabel;
    QIniLibre: TQRLabel;
    QIniLibre5: TQRLabel;
    QIniLibre6: TQRLabel;
    QIniLibre1: TQRLabel;
    QIniLibre2: TQRLabel;
    QIniLibre3: TQRLabel;
    QIniLibre7: TQRLabel;
    QIniLibre8: TQRLabel;
    QIniLibre9: TQRLabel;
    QRevLibre4: TQRLabel;
    QRevLibre: TQRLabel;
    QRevLibre5: TQRLabel;
    QRevLibre6: TQRLabel;
    QRevLibre1: TQRLabel;
    QRevLibre2: TQRLabel;
    QRevLibre3: TQRLabel;
    QRevLibre7: TQRLabel;
    QRevLibre8: TQRLabel;
    QRevLibre9: TQRLabel;
    IRLibre4: TQRLabel;
    IRLibre: TQRLabel;
    IRLibre5: TQRLabel;
    IRLibre6: TQRLabel;
    IRLibre1: TQRLabel;
    IRLibre2: TQRLabel;
    IRLibre3: TQRLabel;
    IRLibre7: TQRLabel;
    IRLibre8: TQRLabel;
    IRLibre9: TQRLabel;
    IniLibre10: TQRLabel;
    RevLibre10: TQRLabel;
    QIniLibre10: TQRLabel;
    QRevLibre10: TQRLabel;
    IRLibre10: TQRLabel;
    BHRecap: TQRBand;
    BG_BUDGENE2: TQRDBText;
    BG_LIBELLE2: TQRLabel;
    BRecap: TQRBand;
    IniRecap5: TQRLabel;
    IniRecap4: TQRLabel;
    IniRecap2: TQRLabel;
    IniRecap3: TQRLabel;
    IniRecap6: TQRLabel;
    IniRecap1: TQRLabel;
    IniRecap7: TQRLabel;
    IniRecap8: TQRLabel;
    IniRecap9: TQRLabel;
    IniRecap10: TQRLabel;
    IniRecap: TQRLabel;
    RevRecap: TQRLabel;
    RevRecap1: TQRLabel;
    RevRecap2: TQRLabel;
    RevRecap3: TQRLabel;
    RevRecap4: TQRLabel;
    RevRecap5: TQRLabel;
    RevRecap6: TQRLabel;
    RevRecap7: TQRLabel;
    RevRecap8: TQRLabel;
    RevRecap9: TQRLabel;
    RevRecap10: TQRLabel;
    QIniRecap: TQRLabel;
    QIniRecap1: TQRLabel;
    QIniRecap2: TQRLabel;
    QIniRecap3: TQRLabel;
    QIniRecap4: TQRLabel;
    QIniRecap5: TQRLabel;
    QIniRecap6: TQRLabel;
    QIniRecap7: TQRLabel;
    QIniRecap8: TQRLabel;
    QIniRecap9: TQRLabel;
    QIniRecap10: TQRLabel;
    QRevRecap: TQRLabel;
    QRevRecap1: TQRLabel;
    QRevRecap2: TQRLabel;
    QRevRecap3: TQRLabel;
    QRevRecap4: TQRLabel;
    QRevRecap5: TQRLabel;
    QRevRecap6: TQRLabel;
    QRevRecap7: TQRLabel;
    QRevRecap8: TQRLabel;
    QRevRecap9: TQRLabel;
    QRevRecap10: TQRLabel;
    IRRecap: TQRLabel;
    IRRecap1: TQRLabel;
    IRRecap2: TQRLabel;
    IRRecap3: TQRLabel;
    IRRecap4: TQRLabel;
    IRRecap5: TQRLabel;
    IRRecap6: TQRLabel;
    IRRecap7: TQRLabel;
    IRRecap8: TQRLabel;
    IRRecap9: TQRLabel;
    IRRecap10: TQRLabel;
    DLRecap: TQRDetailLink;
    DLHRecap: TQRDetailLink;
    TFSoldeFormate: THLabel;
    FSoldeFormate: THValComboBox;
    FTaux: THNumEdit;
    TRResol: TQRLabel;
    RResol: TQRLabel;
    FAvecCptSecond: TCheckBox;
    BLibreHaut: TQRBand;
    TRappelLibRupt: TQRLabel;
    DHLRappelLibre: TQRDetailLink;
    TRappelCodRupt: TQRLabel;
    FQueTotBud: TCheckBox;
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BSecondaireBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure DLMultiNeedData(var MoreData: Boolean);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure FJournalChange(Sender: TObject);
    procedure FExerciceChange(Sender: TObject);
    procedure FExercice_Change(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FPeriode1Change(Sender: TObject);
    procedure FPeriode2Change(Sender: TObject);
    procedure BMultiBeforePrint(var PrintBand: Boolean;  var Quoi: string);
    procedure FTauxExit(Sender: TObject);
    procedure BSecondaireAfterPrint(BandPrinted: Boolean);
    procedure BFPrimaireBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFPrimaireAfterPrint(BandPrinted: Boolean);
    procedure FGene1Change(Sender: TObject);
    procedure QSecondAfterOpen(DataSet: TDataSet);
    procedure FSansRuptClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BHCptBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure DLHGenNeedData(var MoreData: Boolean);
    procedure FNatBudChange(Sender: TObject);
    procedure DLHLibreNeedData(var MoreData: Boolean);
    procedure BHRecapBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure DLHRecapNeedData(var MoreData: Boolean);
    procedure DLRecapNeedData(var MoreData: Boolean);
    procedure FPlanRupturesChange(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure BRecapBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BMultiAfterPrint(BandPrinted: Boolean);
    procedure DHLRappelLibreNeedData(var MoreData: Boolean);
    procedure FQueTotBudClick(Sender: TObject);
    procedure FTotIniRevClick(Sender: TObject);
  private    { Déclarations privées }
    OkEnTeteGen, OkEnteteLibre,
    OkEnTeteRecap,OkRappelEnteteLibre : Boolean ;
    DebitPos                   : Boolean ;
    DansTotal                  : Boolean ;
    StReportGen,StRappelLibTableLibre : String ;
    QBud                       : TQuery ;
    ExoDateDeb, ExoDateFin     : TDateTime ;
    LMulti, LRecap             : TStringList ;
    TotEdt,TotSec, TotIR       : TabTot12 ;
    TotGen                     : TMontTot ;
    TabDate                    : TTabDate12 ;
    AxeJal                     : String3 ;
    Qr1BS_AXE, Qr1BS_BUDSECT, Qr1BS_RUB,
    Qr1BS_LIBELLE, Qr1BS_SECTIONTRIE    : TStringField ;
    QR12BG_BUDGENE, QR12BG_LIBELLE,
    QR12BG_RUB                     : TStringField ;
    FLoading                       : Boolean ;
    ListeCodesRupture              : TStringList ;
    QuelEtatBud                    : TQuelEtatBud ;
    Function  QuoiCpt(i : Integer) : String ;
    Function  QuoiCptRecap(St1,St2 : String) : String ;
    Procedure BalBudGenZoom(Quoi : String) ;
    procedure InfosJournal ;
    procedure DebToDate ;
    procedure FinToDate ;
    Function OkPourRupt : Boolean ;
    procedure MontantIR ;
//    procedure StMontantsBudEcart(M : TabTot12 ; var AffM : array of String) ;
  public
    procedure GenereSQL ; Override ;
    procedure GenereSQL2 ;
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
  end;

Procedure PrepCalcBud(CritEdt : TCritEdt ; Var QMvt : TQuery)  ;
Procedure InitCalcBud(C1, C2 : String ; Var Tot1 : TMontTot ; TabD : TTabDate12 ; CritEdt : TCritEdt ; Q : TQuery) ;
Procedure CalculBud(i : Byte ; Var Tot1 : TMontTot ; NatB : String ; D,C : Double ; Dec : Byte) ;

implementation

{$R *.DFM}

procedure BalBudecSecGen ;
var QR : TFQRBLecSG ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQRBLecSG.Create(Application) ;
Edition.Etat:=etBlBudEcartAnaGen ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalBudGenZoom ;
QR.QuelEtatBud:=Normal ;
QR.InitType (nbBudSec,neBalBud,msSecBAna,'QRBLecSGNor','',TRUE,FALSE,TRUE,Edition) ;
if PP=Nil then
   BEGIN
   try
    QR.ShowModal ;
    finally
    QR.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END Else
   BEGIN
   InitInside(QR,PP) ;
   QR.Show ;
   END ;
END ;

procedure BalBudecSecGenRupture ;
var QR : TFQRBLecSG ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQRBLecSG.Create(Application) ;
Edition.Etat:=etBlBudEcartAnaGen ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalBudGenZoom ;
QR.QuelEtatBud:=Rupture ;
QR.InitType (nbBudSec,neBalBud,msSecBAna,'QRBLecSGRup','',TRUE,FALSE,TRUE,Edition) ;
if PP=Nil then
   BEGIN
   try
    QR.ShowModal ;
    finally
    QR.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END Else
   BEGIN
   InitInside(QR,PP) ;
   QR.Show ;
   END ;
END ;

procedure BalBudecSecGenNature ;
var QR : TFQRBLecSG ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQRBLecSG.Create(Application) ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalBudGenZoom ;
QR.QuelEtatBud:=Nature ;
QR.InitType (nbBudSec,neBalBud,msSecBAna,'QRBLecSGNat','',TRUE,FALSE,TRUE,Edition) ;
if PP=Nil then
   BEGIN
   try
    QR.ShowModal ;
    finally
    QR.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END Else
   BEGIN
   InitInside(QR,PP) ;
   QR.Show ;
   END ;
END ;

procedure BalBudecSecGenZoom(Crit : TCritEdt) ;
var QR : TFQRBLecSG ;
    Edition : TEdition ;
BEGIN
QR:=TFQRBLecSG.Create(Application) ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.BalBudGenZoom ;
 QR.QuelEtatBud:=Normal ;
 QR.InitType (nbBudSec,neBalBud,msSecBAna,'QRBLecSGNor','',FALSE,TRUE,TRUE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

Function TFQRBLecSG.QuoiCpt(i : Integer) : String ;
BEGIN
Case i of
   0 : Result:=Q.FindField('BS_BUDSECT').AsString+' '+Q.FindField('BS_LIBELLE').AsString+'@'+'9'+Qr1BS_AXE.AsString;
   1 : Result:=Q.FindField('BS_BUDSECT').AsString+' '+Q.FindField('BS_LIBELLE').AsString+'@'+'9'+Qr1BS_AXE.AsString ;
   2 : Result:=QSecond.FindField('BG_BUDGENE').AsString+' '+QSecond.FindField('BG_LIBELLE').AsString+'@'+'8' ;
  End ;
END ;

Function TFQRBLecSG.QuoiCptRecap(St1,St2 : String) : String ;
BEGIN
Result:=St1+St2+'@'+'8' ;
END ;


Procedure TFQRBLecSG.BalBudGenZoom(Quoi : String) ;
Var Lp,i: Integer ;
BEGIN
Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
i:=StrToInt(Copy(Quoi,Lp+1,1)) ;
ZoomEdt(i,Quoi) ;
END ;

procedure TFQRBLecSG.FormShow(Sender: TObject);
begin
HelpContext:=15334000 ;
Standards.HelpContext:=15334100 ;
Avances.HelpContext:=15334200 ;
Mise.HelpContext:=15334300 ;
Option.HelpContext:=15334400 ;
TabRuptures.HelpContext:=15334500 ;
Floading:=FALSE ;
TabSup.TabVisible:=False ;
If FJournal.Values.Count>0 Then FJournal.Value:=FJournal.Values[0] ;
FNatBud.ItemIndex:=0 ; FSoldeFormate.ItemIndex:=0 ;
//FRESOL.Value:='C' ;
  inherited;
InitResolution(FRESOL) ;
FLigneCpt.Checked:=True ;
TFGen.Parent:=Avances ; TFGenJoker.Parent:=Avances ;
FJoker.Parent:=Avances ; TFaG.Parent:=Avances ;
FCpte1.Parent:=Avances ; FCpte2.Parent:=Avances ;
FCpte1.TabOrder:=0 ; FCpte2.TabOrder:=1 ; FJoker.TabOrder:=2 ;
DebToDate ; FinToDate  ;
FPlansCo.Enabled:=False ;
(* Plus de menu
FGroupChoixRupt.Visible:=(QuelEtatBud=Nature) or (QuelEtatBud=Rupture) ;
FLigneRupt.Checked:=(QuelEtatBud=Nature) or (QuelEtatBud=Rupture) ;
FLigneRupt.Visible:=(QuelEtatBud=Nature) or (QuelEtatBud=Rupture) ;
FOnlySecRupt.Visible:=(QuelEtatBud=Rupture) ;
FAvecCptSecond.Visible:=(QuelEtatBud=Rupture) or (QuelEtatBud=Nature);
FRuptures.Enabled:=(QuelEtatBud=Rupture);
FTablesLibres.Enabled:=(QuelEtatBud=Nature);
FNatBudChange(Nil) ;
*)
FNatBudChange(Nil) ;
FRuptures.Enabled:=True ;
FRuptures.Checked:=True ;
//If (Not EstSerie(S7)) Then FGroupChoixRupt.Enabled:=FALSE ;
If (EstSerie(S3)) Then FGroupChoixRupt.Visible:=FALSE ;
end;

Procedure TFQRBLecSG.FinirPrint ;
BEGIN
Inherited ;
if (CritEdt.Rupture=rLibres) then VideGroup(LMulti) ;
If (CritEdt.Rupture=rRuptures) then VideRupt(LMulti) ;
if CritEdt.Balbud.FormatPrint.PrSepMontant then FreeLesLignes(Self) ;
//if (CritEdt.Rupture=rLibres) then if CritEdt.BalBud.AvecCptSecond then VideRecap(LRecap) ;
if (CritEdt.Rupture=rLibres) Or (CritEdt.Rupture=rRuptures) then if CritEdt.BalBud.AvecCptSecond then VideRecap(LRecap) ;
END ;

procedure TFQRBLecSG.InitDivers ;
BEGIN
Inherited ;
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
{ Calcul des différentes fourchettes de dates }
//BFPrimaire.Frame.DrawBottom:=CritEdt.BalBud.FormatPrint.PrSepCompte[1] ;
BMulti.Frame.DrawBottom:=FLigneRupt.checked ;
CritEdt.BalBud.FormatPrint.ColMin:=1 ;
CritEdt.BalBud.FormatPrint.ColMax:=10 ;
OkEnTeteGen:=True ;
OkRappelEnteteLibre:=FALSE ;
StRappelLibTableLibre:='' ;
If CritEdt.Rupture=rLibres then
   BEGIN
   if CritEdt.BalBud.RuptOnly=Sur then OkEnTeteGen:=False ;
   END Else
If CritEdt.Rupture=rRuptures then
   BEGIN
   if CritEdt.BalBud.RuptOnly=Sur then OkEnTeteGen:=False ;
   If CritEdt.BalBud.AvecCptSecond then OkEnTeteGen:=False ;
   END ;
END ;

procedure TFQRBLecSG.GenereSQL ;
{ Construction de la requête SQL en fonction du multicritère }
Var St,St1,St2 : String ;
    i : Integer ;
BEGIN
Inherited ;
Q.Close ; Q.SQL.Clear ;
Q.SQL.Add(' Select * ') ;
Q.SQL.Add(' From BUDSECT S Where ') ;
Q.SQL.Add(ExistBud(AxeToFbBud(CritEdt.BalBud.Axe),CritEdt.BalBud.MvtSur,CritEdt.BalBud.Journal,CritEdt.BalBud.Axe,True)) ;
(*
Q.SQL.Add(' Exists  (select BE_AXE, BE_BUDSECT from BUDECR  ') ;
Q.SQL.Add(' where BE_BUDSECT=S.BS_BUDSECT and BE_AXE="'+CritEdt.BalBud.Axe+'" ') ;
Q.SQL.Add(' and BE_BUDJAL="'+CritEdt.BalBud.Journal+'" ) ') ;
*)
Q.SQL.Add(' and BS_AXE="'+CritEdt.BalBud.Axe+'" ') ;
if CritEdt.Joker then Q.SQL.Add(' AND BS_BUDSECT like "'+TraduitJoker(CritEdt.Cpt1)+'" ') Else
   BEGIN
   if CritEdt.Cpt1<>'' then Q.SQL.Add(' AND BS_BUDSECT>="'+CritEdt.Cpt1+'" ') ;
   if CritEdt.Cpt2<>'' then Q.SQL.Add(' AND BS_BUDSECT<="'+CritEdt.Cpt2+'" ') ;
   END ;
(*
If (CritEdt.Rupture=rLibres) then
   BEGIN
   St:=WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,fbBudgen,CritEdt.BalBud.OnlyCptAssocie) ;
   If St<>'' Then Q.SQL.Add(St) ;
   Q.SQL.Add('Order By '+OrderLibre(CritEdt.LibreTrie)+'BS_BUDSECT ') ;
   END Else Q.SQL.Add(' Order By BS_AXE, BS_BUDSECT') ;
*)
if (CritEdt.Rupture=rRuptures) And (CritEdt.BalBud.RuptOnly=Sur) then
   BEGIN
   St1:=CritEdt.BalBud.CodeRupt1 ; St2:=CritEdt.BalBud.CodeRupt2 ;
   i:=Pos('x',St1) ; If i>0 Then Delete(St1,i,1) ;
   If St1<>'' Then Q.SQL.Add(' AND BS_SECTIONTRIE>="'+St1+'" ') ;
   If St2<>'' Then Q.SQL.Add(' AND BS_SECTIONTRIE<="'+St2+'" ') ;
   END ;
Case CritEdt.Rupture Of
  rLibres : BEGIN
            St:=WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,AxeToFbBud(CritEdt.BalBud.Axe),CritEdt.BalBud.OnlyCptAssocie) ;
            If St<>'' Then Q.SQL.Add(St) ;
            Q.SQL.Add('Order By '+OrderLibre(CritEdt.LibreTrie)+'BS_BUDSECT ') ;
            END ;
  rRuptures : Q.SQL.Add(' Order By BS_AXE, BS_SECTIONTRIE') ;
  Else Q.SQL.Add(' Order By BS_AXE, BS_BUDSECT') ;
  END ;
ChangeSQL(Q) ; Q.Open ;
GenereSQL2 ;
END ;

procedure TFQRBLecSG.GenereSQL2 ;
BEGIN
QSecond.Close ; QSecond.SQL.Clear ;
QSecond.SQL.Add(' Select * ') ;
QSecond.SQL.Add(' From BUDGENE G Where ') ;
QSecond.SQL.Add(ExistBud(fbBudgen,CritEdt.BalBud.MvtSur,CritEdt.BalBud.Journal,CritEdt.BalBud.Axe,False)) ;
(*
QSecond.SQL.Add(' Exists  (select BE_BUDGENE, BE_AXE, BE_BUDSECT from BUDECR  ') ;
QSecond.SQL.Add(' where BE_BUDSECT=:BS_BUDSECT and BE_BUDGENE=G.BG_BUDGENE and BE_AXE="'+CritEdt.BalBud.Axe+'" ') ;
QSecond.SQL.Add(' and BE_BUDJAL="'+CritEdt.BalBud.Journal+'" ) ') ;
*)
if CritEdt.SJoker then QSecond.SQL.Add(' AND BG_BUDGENE like "'+TraduitJoker(CritEdt.SCpt1)+'" ') Else
   BEGIN
   if CritEdt.SCpt1<>'' then QSecond.SQL.Add(' AND BG_BUDGENE>="'+CritEdt.SCpt1+'" ') ;
   if CritEdt.SCpt2<>'' then QSecond.SQL.Add(' AND BG_BUDGENE<="'+CritEdt.SCpt2+'" ') ;
   END ;
QSecond.SQL.Add(' Order By BG_BUDGENE') ;
ChangeSQL(QSecond) ; //QSecond.prepare ;
PrepareSQLODBC(QSecond) ;
QSecond.Open ;
END ;

procedure TFQRBLecSG.RenseigneCritere ;
BEGIN
Inherited ;
If OkZoomEdt Then Exit ;
RJournal.Caption:=FJournal.Text ;
RExercice_.Caption:=FExercice_.Text ;
CaseACocher(FValide,RValide) ;
CaseACocher(FQte,RQte) ;
RNatBud.Caption:=FNatBud.Text ;
RTaux.Caption:=FTaux.Text ;
if CritEdt.SJoker then
   BEGIN
   TRGeneral.Caption:=MsgBox.Mess[9] ;
   RGeneral1.Caption:=CritEdt.SCpt1 ;
   END Else
   BEGIN
   TRGeneral.Caption:=MsgBox.Mess[8] ;
   RGeneral1.Caption:=CritEdt.LSCpt1 ; RGeneral2.Caption:=CritEdt.LScpt2 ;
   END ;
RGeneral2.Visible:=Not CritEdt.SJoker ; TRaG.Visible:=Not CritEdt.SJoker ;
AfficheBudgetEn(RDevises,CritEdt) ;
END ;

procedure TFQRBLecSG.ChoixEdition ;
{ Initialisation des options d'édition }
BEGIN
Inherited ;
(*
if (CritEdt.Rupture=rLibres) then
   BEGIN
   DLMulti.PrintBefore:=False ;
   ChargeGroup(LMulti,['D00','D01','D02','D03','D04','D05','D06','D07','D08','D09']) ;
   if CritEdt.BalBud.AvecCptSecond then ChargeRecap(LRecap) ;
   END Else
   BEGIN
   END ;
*)
Case CritEdt.Rupture Of
  rLibres : BEGIN
            DLMulti.PrintBefore:=FALSE ;
            ChargeGroup(LMulti,['D00','D01','D02','D03','D04','D05','D06','D07','D08','D09']) ;
            if CritEdt.BalBud.AvecCptSecond then ChargeRecap(LRecap) ;
            END ;
  rRuptures : BEGIN
              if CritEdt.BalBud.AvecCptSecond then ChargeRecap(LRecap) ;
              END ;
  END ;

RetailleBandesBudEcarts(Self,CritEdt,Msgbox.Mess[3],BSecondaire,BMulti,BRecap,BFinEtat) ;

If CritEdt.BalBud.Resol='C' then CritEdt.Decimale:=V_PGI.OkDecV else
If CritEdt.BalBud.Resol='F' then CritEdt.Decimale:=0 else
If CritEdt.BalBud.Resol='K' then CritEdt.Decimale:=0 else
If CritEdt.BalBud.Resol='M' then CritEdt.Decimale:=0 ;

ChangeMask(Formateur,CritEdt.Decimale,CritEdt.Symbole);
CritEdt.FormatMontant:=Formateur.Masks.PositiveMask;

if CritEdt.Balbud.FormatPrint.PrSepMontant then AffecteLigne(Self,EntetePage,CritEdt.BalBud.FormatPrint) ;
END ;

procedure TFQRBLecSG.RecupCritEdt ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  BalBud.Journal:=FJournal.Value ; BalBud.NatureBud:='' ;
  BalBud.Axe:=AxeJal ;
  If FNatBud.Value<>'' then BalBud.NatureBud:=FNatBud.Value ;
//  BalBud.Exo1:=FExercice.Value ; BalBud.Exo2:=FExercice_.Value ;
  BalBud.Exo1:=QUELEXODTBud(ExoDateDeb) ; BalBud.Exo2:=QUELEXODTBud(ExoDateFin) ;
  Date1:=StrToDate(FDateCompta1.Text)    ; Date2:=StrToDate(FDateCompta2.Text) ;
  DateDeb:=ExoDateDeb ; DateFin:=ExoDateFin ;
  if FValide.State=cbGrayed then Valide:='g' ;
  if FValide.State=cbChecked then Valide:='X' ;
  if FValide.State=cbUnChecked then Valide:='-' ;
  SJoker:=FJokerGene.Visible ;
  If SJoker Then BEGIN SCpt1:=FJokerGene.Text ; SCpt2:=FJokerGene.Text ; END
            Else BEGIN
                 SCpt1:=FGene1.Text  ; SCpt2:=FGene2.Text ;
                 PositionneFourchetteSt(FGene1,FGene2,CritEdt.LSCpt1,CritEdt.LSCpt2) ;
                 END ;
  BalBud.Resol:=FResol.Value[1] ; BalBud.Taux:=StrToFloat(FTaux.Text) ;
  BalBud.Qte:=FQte.Checked ; BalBud.TotIniRev:=FTotIniRev.Checked and FTotIniRev.Enabled ;
  { Rupture }
  BalBud.CodeRupt1:='' ; BalBud.CodeRupt2:='' ; BalBud.PlanRupt:='' ;
  If Not FSansRupt.Checked Then BalBud.PlanRupt:=FPlanRuptures.Value ;
  BALBud.RuptOnly:=QuelleTypeRupt(0,FSAnsRupt.Checked,FAvecRupt.Checked,FSurRupt.Checked) ;
  If BalBud.RuptOnly=Sur Then
     BEGIN
     BalBud.CodeRupt1:=FCodeRupt1.Text ; BalBud.CodeRupt2:=FCodeRupt2.Text ;
     END ;
  BalBud.OnlySecRupt:=FOnlySecRupt.checked ;
  BalBud.AvecCptSecond:=FAvecCptSecond.checked ;
//  BalBud.SoldeFormate:=FSoldeFormate.Value ;
  If DebitPos Then BalBud.SoldeFormate:='PD' Else BalBud.SoldeFormate:='PC' ;
  BalBud.OnlyCptAssocie:=(Rupture<>rRien) and FOnlyCptAssocie.Checked ;
  BalBud.QueTotalBud:=FQueTotBud.Checked ;
  With BalBud.FormatPrint Do
     BEGIN
     END ;
  END ;
END ;

function TFQRBLecSG.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK and OkPourRupt ;
if Result then PrepCalcBud(CritEdt,QBUD)  ;
END ;

Function TFQRBLecSG.OkPourRupt : Boolean ;
BEGIN
Result:=True ;
if (CritEdt.Rupture=rRuptures) then
   BEGIN
   Result:=False ;
   DLMulti.PrintBefore:=True ;
   if (CritEdt.BalBud.RuptOnly=Sur) then
      BEGIN
      ChargeRupt(LMulti,'RU'+Char(CritEdt.BalBud.Axe[2]),CritEdt.BalBud.PlanRupt,CritEdt.BalBud.CodeRupt1 ,CritEdt.BalBud.CodeRupt2) ;
      NiveauRupt(LMulti) ;
      END Else
      BEGIN
      ChargeRupt(LMulti,'RU'+Char(CritEdt.BalBud.Axe[2]),CritEdt.BalBud.PlanRupt,'','') ;
      END ;
   (**)
   Case SectionRetrieBud(CritEdt.BalBud.PlanRupt,CritEdt.BalBud.Axe,ListeCodesRupture) of
     srOk              : Result:=True ;
     srNonStruct       : NumErreurCrit:=9 ;
     srPasEnchainement : NumErreurCrit:=10 ;
     End ;
   END ;
END ;

procedure TFQRBLecSG.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
Fillchar(TotSec,SizeOf(TotSec),#0) ;
BS_BUDSECT.Caption:=MsgBox.Mess[14]+'  '+Qr1BS_BUDSECT.AsString ;
BS_LIBELLE.Caption:=Qr1BS_LIBELLE.AsString ;
InitReport([2],CritEdt.BalBud.FormatPrint.Report) ;
If CritEdt.Rupture=rRuptures then
   BEGIN
   if CritEdt.BalBud.RuptOnly=Avec then
      BEGIN
      (*
      if CritEdt.BalBud.OnlySecRupt then PrintBand:=(Qr1BS_SECTIONTRIE.AsString<>'') ;
      *)
      //Rony 06/11/97
      OkEnTeteGen:=PrintBand ;
      END Else PrintBand:=(CritEdt.BalBud.RuptOnly<>Sur) ;
   END ;
if CritEdt.Rupture=rLibres then PrintBand:=(CritEdt.BalBud.RuptOnly<>Sur) ;
If PrintBand then Quoi:=QuoiCpt(0) ;
end;

procedure TFQRBLecSG.BSecondaireBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var RealPer, EcarPer, RealCum, EcarCum : TabTot12 ;
    I                : Byte ; {}
    C1, C2           : TComponent ;
    TD               : TabloExt ;
    Compte1, Compte2, Lexo   : String ;
    Pourcentage : double ;
    MReport          : TabTRep ;
    FinTemp          : TDateTime ;
    TotBud           : TMontTot ;
    TotMulti         : TabMont77 ;
    LaRupture              : TRuptInf ;
    St                     : String ;
//    Trouve           : Boolean ;
begin
  inherited;
StReportGen:=Qr12BG_BUDGENE.AsString ;
Fillchar(TotBud,SizeOf(TotBud),#0) ;
Fillchar(TotMulti,SizeOf(TotMulti),#0) ;
Fillchar(MReport,SizeOf(MReport),#0) ;
Fillchar(TotIR,SizeOf(TotIR),#0) ;
InitReport([3],CritEdt.BalBud.FormatPrint.Report) ;
InitCalcBud(Qr1BS_BUDSECT.AsString, Qr12BG_BUDGENE.AsString,TotBud,TabDate,CritEdt,QBud) ;
TotGen:=TotBud ;
{Pour Total Cpt Section}
AlimTotEdtEcartBudget(TotSec,TotBud,CritEdt) ;
(*
TotSec[0].TotDebit:= Arrondi(TotSec[0].TotDebit+(TotBud[0][0].totDebit+TotBud[1][0].totDebit), CritEdt.Decimale) ;
TotSec[0].TotCredit:= Arrondi(TotSec[0].TotCredit+(TotBud[0][0].totCredit+TotBud[1][0].totCredit), CritEdt.Decimale) ;
TotSec[4].TotDebit:= Arrondi(TotSec[4].TotDebit+(TotBud[0][1].totDebit+TotBud[1][1].totDebit), CritEdt.Decimale) ;
TotSec[4].TotCredit:= Arrondi(TotSec[4].TotCredit+(TotBud[0][1].totCredit+TotBud[1][1].totCredit), CritEdt.Decimale) ;
TotSec[8].TotDebit:= Arrondi(TotSec[8].TotDebit+(TotBud[0][2].totDebit+TotBud[1][2].totDebit), CritEdt.Decimale) ;
TotSec[8].TotCredit:= Arrondi(TotSec[8].TotCredit+(TotBud[0][2].totCredit+TotBud[1][2].totCredit), CritEdt.Decimale) ;
*)
{Pour Total Initial + Révisé}
AlimTotEdtEcartBudget(TotIR,TotBud,CritEdt) ;
(*
TotIR[0].TotDebit:= Arrondi(TotIR[0].TotDebit+(TotBud[0][0].totDebit+TotBud[1][0].totDebit), CritEdt.Decimale) ;
TotIR[0].TotCredit:= Arrondi(TotIR[0].TotCredit+(TotBud[0][0].totCredit+TotBud[1][0].totCredit), CritEdt.Decimale) ;
TotIR[4].TotDebit:= Arrondi(TotIR[4].TotDebit+(TotBud[0][1].totDebit+TotBud[1][1].totDebit), CritEdt.Decimale) ;
TotIR[4].TotCredit:= Arrondi(TotIR[4].TotCredit+(TotBud[0][1].totCredit+TotBud[1][1].totCredit), CritEdt.Decimale) ;
TotIR[8].TotDebit:= Arrondi(TotIR[8].TotDebit+(TotBud[0][2].totDebit+TotBud[1][2].totDebit), CritEdt.Decimale) ;
TotIR[8].TotCredit:= Arrondi(TotIR[8].TotCredit+(TotBud[0][2].totCredit+TotBud[1][2].totCredit), CritEdt.Decimale) ;
*)
{Pour Total Edition}
if (CritEdt.BalBud.RuptOnly<>Sur) then {En SurRupture le Total Edt est calculé sur le total de chaque rupture}
   BEGIN
   AlimTotEdtEcartBudget(TotEdt,TotBud,CritEdt) ;
   (*
   TotEdt[0].TotDebit:= Arrondi(TotEdt[0].TotDebit+(TotBud[0][0].totDebit+TotBud[1][0].totDebit), CritEdt.Decimale) ;
   TotEdt[0].TotCredit:= Arrondi(TotEdt[0].TotCredit+(TotBud[0][0].totCredit+TotBud[1][0].totCredit), CritEdt.Decimale) ;
   TotEdt[4].TotDebit:= Arrondi(TotEdt[4].TotDebit+(TotBud[0][1].totDebit+TotBud[1][1].totDebit), CritEdt.Decimale) ;
   TotEdt[4].TotCredit:= Arrondi(TotEdt[4].TotCredit+(TotBud[0][1].totCredit+TotBud[1][1].totCredit), CritEdt.Decimale) ;
   TotEDT[8].TotDebit:= Arrondi(TotEdt[8].TotDebit+(TotBud[0][2].totDebit+TotBud[1][2].totDebit), CritEdt.Decimale) ;
   TotEdt[8].TotCredit:= Arrondi(TotEdt[8].TotCredit+(TotBud[0][2].totCredit+TotBud[1][2].totCredit), CritEdt.Decimale) ;
   *)
   END ;
{Pour Report}
MReport[0+1].TotDebit:= Arrondi(MReport[0+1].TotDebit+(TotBud[0][0].totDebit+TotBud[1][0].totDebit), CritEdt.Decimale) ;
MReport[0+1].TotCredit:= Arrondi(MReport[0+1].TotCredit+(TotBud[0][0].totCredit+TotBud[1][0].totCredit), CritEdt.Decimale) ;
MReport[4+1].TotDebit:= Arrondi(MReport[4+1].TotDebit+(TotBud[0][1].totDebit+TotBud[1][1].totDebit), CritEdt.Decimale) ;
MReport[4+1].TotCredit:= Arrondi(MReport[4+1].TotCredit+(TotBud[0][1].totCredit+TotBud[1][1].totCredit), CritEdt.Decimale) ;
MReport[8+1].TotDebit:= Arrondi(MReport[8+1].TotDebit+(TotBud[0][2].totDebit+TotBud[1][2].totDebit), CritEdt.Decimale) ;
MReport[8+1].TotCredit:= Arrondi(MReport[8+1].TotCredit+(TotBud[0][2].totCredit+TotBud[1][2].totCredit), CritEdt.Decimale) ;
{Pour Rupture : Nature ou Rupture}
{ Initial}
TotMulti[0]:= Arrondi(TotMulti[0]+TotBud[0][0].totDebit, CritEdt.Decimale) ;
TotMulti[1]:= Arrondi(TotMulti[1]+TotBud[0][0].totCredit, CritEdt.Decimale) ;
TotMulti[8]:= Arrondi(TotMulti[8]+TotBud[0][1].totDebit, CritEdt.Decimale) ;
TotMulti[9]:= Arrondi(TotMulti[9]+TotBud[0][1].totCredit, CritEdt.Decimale) ;
TotMulti[16]:= Arrondi(TotMulti[16]+TotBud[0][2].totDebit, CritEdt.Decimale) ;
TotMulti[17]:= Arrondi(TotMulti[17]+TotBud[0][2].totCredit, CritEdt.Decimale) ;
{ Revise }
TotMulti[20]:= Arrondi(TotMulti[20]+TotBud[1][0].totDebit, CritEdt.Decimale) ;
TotMulti[21]:= Arrondi(TotMulti[21]+TotBud[1][0].totCredit, CritEdt.Decimale) ;
TotMulti[28]:= Arrondi(TotMulti[28]+TotBud[1][1].totDebit, CritEdt.Decimale) ;
TotMulti[29]:= Arrondi(TotMulti[29]+TotBud[1][1].totCredit, CritEdt.Decimale) ;
TotMulti[36]:= Arrondi(TotMulti[36]+TotBud[1][2].totDebit, CritEdt.Decimale) ;
TotMulti[37]:= Arrondi(TotMulti[37]+TotBud[1][2].totCredit, CritEdt.Decimale) ;

{ Période }
     { Mvt init }
Compte1:='S/G'+CritEdt.BalBud.Journal+QR1BS_RUB.AsString+':'+QR12BG_RUB.AsString ;
Compte2:='' ;
LExo:=QUELEXODTBud(CritEdt.Date1) ; if (Lexo<>QUELEXODTBud(CritEdt.Date2)) then LExo:='' ; FinTemp:=FinDeMois(CritEdt.Date2) ;
GetCumul('RUBREA',Compte1,Compte2,CritEdt.BalBud.SANbud,CritEdt.Etab,CritEDT.DeviseAffichee,LExo,CritEdt.Date1,FinTemp,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
RealPer[0].TotDebit:=TD[3] ; RealPer[0].TotCredit:=TD[2] ;
EcarPer[0].TotDebit:= Arrondi((RealPer[0].TotDebit-TotGen[0][0].TotDebit), CritEdt.Decimale) ;
EcarPer[0].TotCredit:=Arrondi((RealPer[0].TotCredit-TotGen[0][0].TotCredit), CritEdt.Decimale) ;

Reevaluation(RealPer[0].TotDebit,RealPer[0].TotCredit,CritEdt.BalBud.Resol, CritEdt.BalBud.Taux) ;
Reevaluation(EcarPer[0].TotDebit,EcarPer[0].TotCredit,CritEdt.BalBud.Resol, CritEdt.BalBud.Taux) ;

IniGenCum1.Caption:=PrintSoldeFormate(TotGen[0][0].TotDebit, TotGen[0][0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
IniGenCum2.Caption:=PrintSoldeFormate(RealPer[0].TotDebit, RealPer[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
//IniGenCum3.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Abs(EcarPer[0].TotDebit-EcarPer[0].TotCredit) , CritEDT.AfficheSymbole ) ;
IniGenCum3.Caption:=PrintEcart(EcarPer[0].TotDebit,EcarPer[0].TotCredit,CritEDT.Decimale,DebitPos) ;
(*
if (IniGenCum1.Caption<>'')and(IniGenCum3.Caption<>'') then Pourcentage:=(Abs(EcarPer[0].TotDebit-EcarPer[0].TotCredit)*100)/(TotGen[0][0].TotDebit-TotGen[0][0].TotCredit)
                                                       Else Pourcentage:=0 ;
*)
if (IniGenCum1.Caption<>'')and(IniGenCum3.Caption<>'') then
   BEGIN
   Pourcentage:=((EcarPer[0].TotDebit-EcarPer[0].TotCredit)*100)/(TotGen[0][0].TotDebit-TotGen[0][0].TotCredit) ;
   Pourcentage:=Abs(Pourcentage) ;
   If Pos('-',IniGenCum3.Caption)>0 Then Pourcentage:=-1*Pourcentage ;
   END Else Pourcentage:=0 ;
IniGenCum4.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Pourcentage , CritEDT.AfficheSymbole ) ;
     { Mvt Revi }
RealPer[1].TotDebit:=RealPer[0].TotDebit ;
RealPer[1].TotCredit:=RealPer[0].TotCredit ;
EcarPer[1].TotDebit:= Arrondi((RealPer[0].TotDebit-TotGen[1][0].TotDebit), CritEdt.Decimale) ;
EcarPer[1].TotCredit:=Arrondi((RealPer[0].TotCredit-TotGen[1][0].TotCredit), CritEdt.Decimale) ;

RevGenCum1.Caption:=PrintSoldeFormate(TotGen[1][0].TotDebit, TotGen[1][0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
RevGenCum2.Caption:=PrintSoldeFormate(RealPer[1].TotDebit, RealPer[1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
//RevGenCum3.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Abs(EcarPer[1].TotDebit-EcarPer[1].TotCredit), CritEDT.AfficheSymbole ) ;
RevGenCum3.Caption:=PrintEcart(EcarPer[1].TotDebit,EcarPer[1].TotCredit,CritEDT.Decimale,DebitPos) ;
(*
if (RevGenCum1.Caption<>'') and (RevGenCum3.Caption<>'') then Pourcentage:=((EcarPer[1].TotDebit-EcarPer[1].TotCredit)*100)/(TotGen[1][0].TotDebit-TotGen[1][0].TotCredit)
                                                         Else Pourcentage:=0 ;
*)
if (RevGenCum1.Caption<>'') and (RevGenCum3.Caption<>'') then
  BEGIN
  Pourcentage:=((EcarPer[1].TotDebit-EcarPer[1].TotCredit)*100)/(TotGen[1][0].TotDebit-TotGen[1][0].TotCredit) ;
  Pourcentage:=Abs(Pourcentage) ;
  If Pos('-',RevGenCum3.Caption)>0 Then Pourcentage:=-1*Pourcentage ;
  END Else Pourcentage:=0 ;
RevGenCum4.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Pourcentage, CritEDT.AfficheSymbole ) ;
     { Qt Ini }
     { -- }
     {Report : réalisé}
MReport[2].TotDebit:=Arrondi(RealPer[0].TotDebit,CritEdt.Decimale) ;
MReport[2].TotCredit:=Arrondi(RealPer[0].TotCredit,CritEdt.Decimale) ;
{ Cumul }
     { init }
(* GP le 26/02/99
LExo:=QUELEXODTBud(CritEdt.Date1) ; FinTemp:=FinDeMois(CritEdt.DateFin) ;
GetCumul('RUBREA',Compte1,Compte2,CritEdt.BalBud.SANbud,CritEdt.Etab,CritEDT.DeviseAffichee,LExo,CritEdt.DateDeb,FinTemp,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
*)
LExo:=QUELEXODTBud(CritEdt.DateDeb) ; if (Lexo<>QUELEXODTBud(CritEdt.Date2)) then LExo:='' ; FinTemp:=FinDeMois(CritEdt.Date2) ;
GetCumul('RUBREA',Compte1,Compte2,CritEdt.BalBud.SANbud,CritEdt.Etab,CritEDT.DeviseAffichee,LExo,CritEdt.DateDeb,FinTemp,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
RealCum[0].TotDebit:=TD[3] ; RealCum[0].TotCredit:=TD[2] ;
EcarCum[0].TotDebit:= Arrondi((RealCum[0].TotDebit-TotGen[0][1].TotDebit), CritEdt.Decimale) ;
EcarCum[0].TotCredit:=Arrondi((RealCum[0].TotCredit-TotGen[0][1].TotCredit), CritEdt.Decimale) ;

Reevaluation(RealCum[0].TotDebit,RealCum[0].TotCredit,CritEdt.BalBud.Resol, CritEdt.BalBud.Taux) ;
Reevaluation(EcarCum[0].TotDebit,EcarCum[0].TotCredit,CritEdt.BalBud.Resol, CritEdt.BalBud.Taux) ;

IniGenCum5.Caption:=PrintSoldeFormate(TotGen[0][1].TotDebit, TotGen[0][1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
IniGenCum6.Caption:=PrintSoldeFormate(RealCum[0].TotDebit, RealCum[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
//IniGenCum7.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Abs(EcarCum[0].TotDebit-EcarCum[0].TotCredit), CritEDT.AfficheSymbole ) ;
IniGenCum7.Caption:=PrintEcart(EcarCum[0].TotDebit,EcarCum[0].TotCredit,CritEDT.Decimale,DebitPos ) ;
(*
if (IniGenCum5.Caption<>'') and (IniGenCum7.Caption<>'') then Pourcentage:=(Abs(EcarCum[0].TotDebit-EcarCum[0].TotCredit)*100)/(TotGen[0][1].TotDebit-TotGen[0][1].TotCredit)
                                                         Else Pourcentage:=0 ;
*)
if (IniGenCum5.Caption<>'') and (IniGenCum7.Caption<>'') then
   BEGIN
   Pourcentage:=((EcarCum[0].TotDebit-EcarCum[0].TotCredit)*100)/(TotGen[0][1].TotDebit-TotGen[0][1].TotCredit) ;
   Pourcentage:=Abs(Pourcentage) ;
   If Pos('-',IniGenCum7.Caption)>0 Then Pourcentage:=-1*Pourcentage ;
   END Else Pourcentage:=0 ;
IniGenCum8.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Pourcentage , CritEDT.AfficheSymbole ) ;
     { Revi }
RealCum[1].TotDebit:=RealCum[0].TotDebit ;
RealCum[1].TotCredit:=RealCum[0].TotCredit ;
EcarCum[1].TotDebit:= Arrondi((RealCum[0].TotDebit-TotGen[1][1].TotDebit), CritEdt.Decimale) ;
EcarCum[1].TotCredit:=Arrondi((RealCum[0].TotCredit-TotGen[1][1].TotCredit), CritEdt.Decimale) ;

RevGenCum5.Caption:=PrintSoldeFormate(TotGen[1][1].TotDebit, TotGen[1][1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
RevGenCum6.Caption:=PrintSoldeFormate(RealCum[1].TotDebit, RealCum[1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
//RevGenCum7.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Abs(EcarCum[1].TotDebit-EcarCum[1].TotCredit), CritEDT.AfficheSymbole ) ;
RevGenCum7.Caption:=PrintEcart(EcarCum[1].TotDebit,EcarCum[1].TotCredit,CritEDT.Decimale,DebitPos) ;
(*
if (RevGenCum5.Caption<>'') and (RevGenCum7.Caption<>'') then Pourcentage:=((EcarCum[1].TotDebit-EcarCum[1].TotCredit)*100)/(TotGen[1][1].TotDebit-TotGen[1][1].TotCredit)
                                                         Else Pourcentage:=0 ;
*)
if (RevGenCum5.Caption<>'') and (RevGenCum7.Caption<>'') then
   BEGIN
   Pourcentage:=((EcarCum[1].TotDebit-EcarCum[1].TotCredit)*100)/(TotGen[1][1].TotDebit-TotGen[1][1].TotCredit) ;
   Pourcentage:=Abs(Pourcentage) ;
   If Pos('-',RevGenCum7.Caption)>0 Then Pourcentage:=-1*Pourcentage ;
   END Else Pourcentage:=0 ;
RevGenCum8.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Pourcentage, CritEDT.AfficheSymbole ) ;

     {Report : réalisé}
MReport[6].TotDebit:=Arrondi(RealCum[0].TotDebit,CritEdt.Decimale) ;
MReport[6].TotCredit:=Arrondi(RealCum[0].TotCredit,CritEdt.Decimale) ;
{ Annuel }
     { Init }
IniGenCum9.Caption:=PrintSoldeFormate(TotGen[0][2].TotDebit, TotGen[0][2].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
     { Revi }
RevGenCum9.Caption:=PrintSoldeFormate(TotGen[1][2].TotDebit, TotGen[1][2].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;

{--------------------------------------------}
{ Calcule du réalisé en Total Edition }
if (CritEdt.BalBud.RuptOnly<>Sur) then {En SurRupture le Total Edt du réalisé est calculé sur le total de chaque rupture}
   BEGIN
   { CORRECTION GP POUR LA X° FOIS
   TotEdt[1].TotDebit:= Arrondi(TotEdt[1].TotDebit+(RealPer[0].TotDebit+RealPer[1].TotDebit), CritEdt.Decimale) ;
   TotEdt[1].TotCredit:= Arrondi(TotEdt[1].TotCredit+(RealPer[0].TotCredit+RealPer[1].TotCredit), CritEdt.Decimale) ;
   TotEdt[5].TotDebit:= Arrondi(TotEdt[5].TotDebit+(RealCum[0].TotDebit+RealCum[1].TotDebit), CritEdt.Decimale) ;
   TotEdt[5].TotCredit:= Arrondi(TotEdt[5].TotCredit+(RealCum[0].TotCredit+RealCum[1].TotCredit), CritEdt.Decimale) ;
   }
   TotEdt[1].TotDebit:= Arrondi(TotEdt[1].TotDebit+(RealPer[0].TotDebit), CritEdt.Decimale) ;
   TotEdt[1].TotCredit:= Arrondi(TotEdt[1].TotCredit+(RealPer[0].TotCredit), CritEdt.Decimale) ;
   TotEdt[5].TotDebit:= Arrondi(TotEdt[5].TotDebit+(RealCum[0].TotDebit), CritEdt.Decimale) ;
   TotEdt[5].TotCredit:= Arrondi(TotEdt[5].TotCredit+(RealCum[0].TotCredit), CritEdt.Decimale) ;
   END ;
{ Calcule du réalisé en Compte Primaire }
{ CORRECTION GP POUR LA X° FOIS
TotSec[1].TotDebit:= Arrondi(TotSec[1].TotDebit+(RealPer[0].TotDebit+RealPer[1].TotDebit), CritEdt.Decimale) ;
TotSec[1].TotCredit:= Arrondi(TotSec[1].TotCredit+(RealPer[0].TotCredit+RealPer[1].TotCredit), CritEdt.Decimale) ;
TotSec[5].TotDebit:= Arrondi(TotSec[5].TotDebit+(RealCum[0].TotDebit+RealCum[1].TotDebit), CritEdt.Decimale) ;
TotSec[5].TotCredit:= Arrondi(TotSec[5].TotCredit+(RealCum[0].TotCredit+RealCum[1].TotCredit), CritEdt.Decimale) ;
}
TotSec[1].TotDebit:= Arrondi(TotSec[1].TotDebit+(RealPer[0].TotDebit), CritEdt.Decimale) ;
TotSec[1].TotCredit:= Arrondi(TotSec[1].TotCredit+(RealPer[0].TotCredit), CritEdt.Decimale) ;
TotSec[5].TotDebit:= Arrondi(TotSec[5].TotDebit+(RealCum[0].TotDebit), CritEdt.Decimale) ;
TotSec[5].TotCredit:= Arrondi(TotSec[5].TotCredit+(RealCum[0].TotCredit), CritEdt.Decimale) ;
{ Calcule du réalisé en Total Ini+Rev }
{ Réalisé Période }
TotIR[1].TotDebit:= Arrondi(TotIR[1].TotDebit+RealPer[0].TotDebit, CritEdt.Decimale) ;
TotIR[1].TotCredit:=Arrondi(TotIR[1].TotCredit+RealPer[0].TotCredit, CritEdt.Decimale) ;
{ Réalisé Cumulé }
TotIR[5].TotDebit:= Arrondi(TotIR[5].TotDebit+RealCum[0].TotDebit, CritEdt.Decimale) ;
TotIR[5].TotCredit:=Arrondi(TotIR[5].TotCredit+RealCum[0].TotCredit, CritEdt.Decimale) ;
{ Calcule du réalisé en Zone libre ou Rupture }
TotMulti[2]:= Arrondi(TotMulti[2]+RealPer[0].TotDebit, CritEdt.Decimale) ;
TotMulti[3]:= Arrondi(TotMulti[3]+RealPer[0].TotCredit, CritEdt.Decimale) ;
TotMulti[10]:= Arrondi(TotMulti[10]+RealCum[0].TotDebit, CritEdt.Decimale) ;
TotMulti[11]:= Arrondi(TotMulti[11]+RealCum[0].TotCredit, CritEdt.Decimale) ;
{Révisé}
TotMulti[22]:= Arrondi(TotMulti[22]+RealPer[1].TotDebit, CritEdt.Decimale) ;
TotMulti[23]:= Arrondi(TotMulti[23]+RealPer[1].TotCredit, CritEdt.Decimale) ;
TotMulti[30]:= Arrondi(TotMulti[30]+RealCum[1].TotDebit, CritEdt.Decimale) ;
TotMulti[31]:= Arrondi(TotMulti[31]+RealCum[1].TotCredit, CritEdt.Decimale) ;
{Valeurs aux Quantitées}
for i:=1 to 10 do
    BEGIN
    C1:=FindComponent('QiniGenCum'+InttoStr(i)) ;
    C2:=FindComponent('QRevGenCum'+InttoStr(i)) ;
    TQRLabel(C1).Caption:='' ;
    TQRLabel(C2).Caption:='' ;
    END ;
IniGenCum10.Caption:='' ;
RevGenCum10.Caption:='' ;
AddReportBAL([1], CritEdt.Balbud.FormatPrint.Report, MReport, CritEdt.Decimale) ;
if (CritEdt.Rupture=rLibres) then
   BEGIN
   AddGroupLibre(LMulti,Q,AxeToFbBud(CritEdt.BalBud.Axe),CritEdt.LibreTrie,TotMulti) ;
   if (CritEdt.BalBud.RuptOnly=Sur) then PrintBand:=False ;
   if CritEdt.BalBud.AvecCptSecond then
      BEGIN
      if DansRuptLibre(Q,AxeToFbBud(CritEdt.BalBud.Axe),CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie)
         then AddRecap(LRecap, [QR12BG_BUDGENE.AsString], [QR12BG_LIBELLE.AsString],TotMulti) ;

      (*
      Trouve:=False ;
      for i:=0 to 9 do
          BEGIN
          if Q.findField('BS_TABLE'+IntToStr(i)).AsString<>'' then Trouve:=True ;
          if Trouve then Break ;
          END ;
      if Trouve then AddRecap(LRecap, [QR12BG_BUDGENE.AsString], [QR12BG_LIBELLE.AsString],TotMulti) ;
      *)
      END ;
   END Else
   BEGIN
   if (CritEdt.Rupture=rRuptures) then
      BEGIN
      if (CritEdt.BalBud.RuptOnly=Avec) then
         BEGIN
         (*
         if CritEdt.BalBud.OnlySecRupt then PrintBand:=(Qr1BS_SECTIONTRIE.AsString<>'') ;
         *)
         END Else
         BEGIN
         (*
         if (CritEdt.BalBud.RuptOnly=Sur) then
            if CritEdt.BalBud.AvecCptSecond then PrintBand:=(Qr1BS_SECTIONTRIE.AsString<>'')
                                            Else PrintBand:=False ;
         *)
         if (CritEdt.BalBud.RuptOnly=Sur) then
            BEGIN
            if CritEdt.BalBud.AvecCptSecond then
               BEGIN
//               PrintBand:=(Qr1BS_SECTIONTRIE.AsString<>'') ;
               PrintBand:=FALSE ;
               DansQuelleRupt(LMulti,Qr1BS_SECTIONTRIE.AsString,LaRupture) ;
               If LaRupture.CodRupt<>'' Then
                  BEGIN
                  St:=' ('+LaRupture.CodRupt+' '+LaRupture.LibRupt+')' ;
                  AddRecap(LRecap, [QR12BG_BUDGENE.AsString], [QR12BG_LIBELLE.AsString+St],TotMulti) ;
                  END ;
               END Else PrintBand:=False ;
            END ;
         END ;
      If CritEdt.BalBud.RuptOnly<>Sans then AddRupt(LMulti,Qr1BS_SECTIONTRIE.AsString,TotMulti) ;
      END ;
   END ;
OkEnTeteGen:=PrintBand ;
If PrintBand then MontantIR ;
end;

procedure TFQRBLecSG.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var StMontants : array[0..9] of String ;
    I          : Byte ;
begin
  inherited;
Fillchar(StMontants,SizeOf(StMontants),#0) ;
StPourMontantsBudgetEcart(TotEdt,CritEdt,DebitPos,StMontants) ;
For i:=1 to 10 do
    BEGIN
    TQRLabel(FindComponent('TotCum'+IntToStr(i))).Caption:=StMontants[i-1] ;
    END ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFQRBLecSG.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
Titre1Rep.Caption:=Titre2Rep.Caption ;
ReportCol1.Caption:=ReportCol11.Caption ;
ReportCol2.Caption:=ReportCol12.Caption ;
ReportCol3.Caption:=ReportCol13.Caption ;
ReportCol4.Caption:=ReportCol14.Caption ;
ReportCol5.Caption:=ReportCol15.Caption ;
ReportCol6.Caption:=ReportCol16.Caption ;
ReportCol7.Caption:=ReportCol17.Caption ;
ReportCol8.Caption:=ReportCol18.Caption ;
ReportCol9.Caption:=ReportCol19.Caption ;
ReportCol10.Caption:=ReportCol20.Caption ;
end;

procedure TFQRBLecSG.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var MReport : TabTRep ;
    EcartPer, EcartCum, Pourcentage : Double ;
begin
  inherited;
FillChar(Mreport,Sizeof(Mreport),#0) ;
Case QuelReportBAL(CritEdt.Balbud.FormatPrint.Report,MReport) of
  1 : Titre2Rep.Caption:=MsgBox.Mess[1] ;
 end ;
EcartPer:=Arrondi((MReport[1].TotDebit-MReport[1].TotCredit)-(MReport[2].TotDebit-MReport[2].TotCredit), CritEdt.Decimale) ;
EcartCum:=Arrondi((MReport[5].TotDebit-MReport[5].TotCredit)-(MReport[6].TotDebit-MReport[6].TotCredit), CritEdt.Decimale) ;
{Total Report en période}
ReportCol11.Caption:=PrintSoldeFormate(MReport[1].TotDebit, MReport[1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
ReportCol12.Caption:=PrintSoldeFormate(MReport[2].TotDebit, MReport[2].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
ReportCol13.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Abs(EcartPer), CritEDT.AfficheSymbole ) ;
if (ReportCol11.Caption<>'') and (ReportCol13.Caption<>'') then Pourcentage:=(EcartPer*100)/(MReport[1].TotDebit-MReport[1].TotCredit)
                                                   Else Pourcentage:=0 ;
ReportCol14.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Abs(Pourcentage) , CritEDT.AfficheSymbole ) ;
{Total Report en Cumulé}
ReportCol15.Caption:=PrintSoldeFormate(MReport[5].TotDebit, MReport[5].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
ReportCol16.Caption:=PrintSoldeFormate(MReport[6].TotDebit, MReport[6].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
ReportCol17.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Abs(EcartCum), CritEDT.AfficheSymbole ) ;
if (ReportCol15.Caption<>'') and (ReportCol17.Caption<>'') then Pourcentage:=(EcartCum*100)/(MReport[5].TotDebit-MReport[5].TotCredit)
                                                   Else Pourcentage:=0 ;
ReportCol18.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Abs(Pourcentage) , CritEDT.AfficheSymbole ) ;
{Total Report en Annuel}
ReportCol19.Caption:=PrintSoldeFormate(MReport[9].TotDebit, MReport[9].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
{Total Report en Extrapolé}
ReportCol20.Caption:='' ;

end;

procedure TFQRBLecSG.BMultiBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.BalBud.RuptOnly<>Sans) ;
end;

procedure TFQRBLecSG.DLMultiNeedData(var MoreData: Boolean);
var TotMulti         : TabMont77 ;
    Librupt, Lib1, CodRupt, stCode,SauveCodRupt : String ;
    Quellerupt       : Integer ;
    Chaque,i         : Byte ;
    Col              : TColor ;
    TotMontLibre     : TabTot12 ;
    StMontLibre      : array[0..9] of String ;
    LibRuptInf : Array[1..10] Of TRuptInf ;
begin
  inherited;
MoreData:=False ;
If (CritEdt.BalBud.RuptOnly<>Sans) then
   BEGIN
   Case CritEdt.Rupture of
     rRuptures : MoreData:=PrintRupt(LMulti,Qr1BS_SECTIONTRIE.AsString,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotMulti) ;
     rLibres   : BEGIN
                 MoreData:=PrintGroupLibre(LMulti,Q,AxeToFbBud(CritEdt.BalBud.Axe),CritEdt.LibreTrie,CodRupt,LibRupt,Lib1,TotMulti,Quellerupt,Col,LibRuptInf) ;
                 If MoreData then
                    BEGIN
                    StCode:=CodRupt ;
                    Delete(StCode,1,Quellerupt+2) ;
                    MoreData:=DansChoixCodeLibre(StCode,Q,AxeToFbBud(CritEdt.BalBud.Axe),CritEdt.LibreCodes1,CritEdt.LibreCodes2, CritEdt.LibreTrie) ;
                    StRappelLibTableLibre:='' ;
                    END ;
                 BMulti.Font.Color:=Col ; BLibreHaut.Font.Color:=Col ; BRappelLibreHaut.Font.Color:=Col ;
                 END ;
     End ;
   END ;
OkEnteteLibre:=MoreData ;
if MoreData then
   BEGIN
   {Affichage de la bande Rupture ou Nature libre}
   TLibRupt.Caption:='' ; TLibRupt.Left:=BG_LIBELLE.Left ;
   BRappelLibreHaut.Height:=HauteurBandeRuptIni ;
   if CritEdt.Rupture=rLibres then
      BEGIN
      SauveCodRupt:=CodRupt ;
      StRappelLibTableLibre:=AlimStRuptSup(QuelleRupt,CodRupt,Lib1,LibRuptInf) ;
      insert(MsgBox.Mess[6]+' ',CodRupt,Quellerupt+2) ;
      TCodRupt.Caption:=CodRupt+' '+Lib1 ;
//      AlimRuptSup(HauteurBandeRuptIni,QuelleRupt,TCodRupt.Width,BRappelLibreHaut,LibRuptInf,Self) ;
      AlimRuptSup(HauteurBandeRuptIni,QuelleRupt,TitreColCpt.Width,BRappelLibreHaut,LibRuptInf,Self) ;
//      StRappelLibTableLibre:=' ('+StCode+' '+Lib1+')'
      END Else BEGIN TCodRupt.Caption:=CodRupt ; TLibRupt.Caption:=LibRupt ; TLibRupt.Left:=BS_LIBELLE.Left ; END ;
   TRappelCodRupt.Caption:=TCodRupt.Caption ;
   TRappelLibRupt.Caption:=TLibRupt.Caption ;
   TRappelLibRupt.Left:=TLibRupt.Left ;
   if (CritEdt.Rupture=rLibres) And (CritEdt.BalBud.RuptOnly=Sur) And CritEdt.BalBud.AvecCptSecond then TCodRupt.Caption:=SauveCodRupt+' '+Lib1 ;
   For Chaque:=1 to 5 do { Ini, Rév, Ini+Rev, QIni, QRev }
       BEGIN
       for i:=0 to 9 do
           BEGIN
           Case Chaque of
              1 :  BEGIN { Initial }
                   TotMontLibre[i].TotDebit:=TotMulti[i*2] ; TotMontLibre[i].TotCredit:=TotMulti[i*2+1] ;
                   END ;
              2 :  BEGIN { Révisé }
                   TotMontLibre[i].TotDebit:=TotMulti[(i*2)+20] ; TotMontLibre[i].TotCredit:=TotMulti[(i*2)+(1+20)] ;
                   END ;
              3 :  BEGIN { Total Ini+Rév }
                   { Pas de total en Réalisé période (i==1) ni en Réalisé cumulé (i==5) }
                   if (i=1) or (i=5) then BEGIN TotMontLibre[i].TotDebit:=TotMulti[i*2] ; TotMontLibre[i].TotCredit:=TotMulti[i*2+1] END
                                     Else BEGIN TotMontLibre[i].TotDebit:=TotMulti[i*2]+TotMulti[(i*2)+20] ; TotMontLibre[i].TotCredit:=TotMulti[i*2+1]+TotMulti[(i*2)+(1+20)] ; END ;
                   if (CritEdt.BalBud.RuptOnly=Sur)
                      or CritEdt.BalBud.OnlyLibre then
                      BEGIN
                      {Total Edition en Nature uniquement}
                      TotEdt[i].TotDebit:=Arrondi(TotEdt[i].TotDebit+TotMontLibre[i].TotDebit, CritEdt.Decimale) ;
                      TotEdt[i].TotCredit:=Arrondi(TotEdt[i].TotCredit+TotMontLibre[i].TotCredit, CritEdt.Decimale) ;
                      END ;
                   END ;
              End ;
           END ;
       Fillchar(StMontLibre,SizeOf(StMontLibre),#0) ;
       StPourMontantsBudgetEcart(TotMontLibre,CritEdt,DebitPos,StMontLibre) ;
       For i:=1 to 10 do
           BEGIN
           Case Chaque of
              1 : TQRLabel(FindComponent('IniLibre'+IntToStr(i))).Caption:=StMontLibre[i-1] ;
              2 : TQRLabel(FindComponent('RevLibre'+IntToStr(i))).Caption:=StMontLibre[i-1] ;
              3 : TQRLabel(FindComponent('IRLibre'+IntToStr(i))).Caption:=StMontLibre[i-1] ;
              4 : TQRLabel(FindComponent('QIniLibre'+IntToStr(i))).Caption:='' ;
              5 : TQRLabel(FindComponent('QRevLibre'+IntToStr(i))).Caption:='' ;
              End ;
           END ;
       END ;
   END ;
end;

procedure TFQRBLecSG.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr1BS_AXE    :=TStringField(Q.FindField('BS_AXE'));
Qr1BS_BUDSECT :=TStringField(Q.FindField('BS_BUDSECT'));
Qr1BS_RUB     :=TStringField(Q.FindField('BS_RUB'));
Qr1BS_LIBELLE :=TStringField(Q.FindField('BS_LIBELLE'));
Qr1BS_SECTIONTRIE  :=TStringField(Q.FindField('BS_SECTIONTRIE'));
end;

procedure TFQRBLecSG.FJournalChange(Sender: TObject);
Var St : String ;
     A : Char ;
begin
  inherited;
if FJournal.Value='' then Exit ;
InfosJournal ;
If Not QRLoading Then BEGIN FCpte1.Clear ; FCpte2.Clear ; FJoker.Clear ; END ;
ST:=AxeJal ; A:=St[2] ;
Case A of
  '1' : BEGIN FCpte1.ZoomTable:=tzBudSec1 ; FCpte2.ZoomTable:=FCpte1.ZoomTable ; FPlanRuptures.Datatype:='ttRuptSect1' ; END ;
  '2' : BEGIN FCpte1.ZoomTable:=tzBudSec2 ; FCpte2.ZoomTable:=FCpte1.ZoomTable ; FPlanRuptures.Datatype:='ttRuptSect2' ; END ;
  '3' : BEGIN FCpte1.ZoomTable:=tzBudSec3 ; FCpte2.ZoomTable:=FCpte1.ZoomTable ; FPlanRuptures.Datatype:='ttRuptSect3' ; END ;
  '4' : BEGIN FCpte1.ZoomTable:=tzBudSec4 ; FCpte2.ZoomTable:=FCpte1.ZoomTable ; FPlanRuptures.Datatype:='ttRuptSect4' ; END ;
  '5' : BEGIN FCpte1.ZoomTable:=tzBudSec5 ; FCpte2.ZoomTable:=FCpte1.ZoomTable ; FPlanRuptures.Datatype:='ttRuptSect5' ; END ;
  end ;
If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
end;

procedure TFQRBLecSG.InfosJournal ;
Var QJal  : TQuery ;
BEGIN
QJal:=OpenSQL('Select BJ_EXODEB, BJ_EXOFIN, BJ_PERDEB, BJ_PERFIN, BJ_AXE, BJ_NATJAL From BUDJAL Where BJ_BUDJAL="'+FJournal.Value+'"',True) ;
if Not QJal.EOF then
   BEGIN
   FEXERCICE.Value:=QJal.FindField('BJ_EXODEB').AsString ; FEXERCICE_.Value:=QJal.FindField('BJ_EXOFIN').AsString ;
   FPeriode1.Value:=QJal.FindField('BJ_PERDEB').AsString ; FPeriode2.Value:=QJal.FindField('BJ_PERFIN').AsString ;
   FDateCompta1.Text:=QJal.FindField('BJ_PERDEB').AsString ; FDateCompta2.Text:=QJal.FindField('BJ_PERFIN').AsString ;
   ExoDateDeb:=QJal.FindField('BJ_PERDEB').AsDateTime ; ExoDateFin:=QJal.FindField('BJ_PERFIN').AsDateTime ;
   AxeJal:=QJal.FindField('BJ_AXE').AsString ;
   DebitPos:=QJal.FindField('BJ_NATJAL').AsString='CHA' ;
   END ;
Ferme(QJal) ;
END ;

procedure TFQRBLecSG.FExerciceChange(Sender: TObject);
begin
if FExercice.Value>FExercice_.Value then FExercice_.Value:=FExercice.Value ;
ListePeriode(FExercice.Value,FPeriode1.Items,FPeriode1.Values,True) ;
FPeriode1.ItemIndex:=0 ; DebToDate ;
end;

procedure TFQRBLecSG.FExercice_Change(Sender: TObject);
begin
  inherited;
if FExercice_.Value<FExercice.Value then FExercice.Value:=FExercice_.Value ;
ListePeriode(FExercice_.Value,FPeriode2.Items,FPeriode2.Values,False) ;
FPeriode2.ItemIndex:=FPeriode2.Items.Count-1 ; FinToDate ;
end;

procedure TFQRBLecSG.FPeriode1Change(Sender: TObject);
begin
  inherited;
if StrToDate(FPeriode1.Value)>StrToDate(FPeriode2.Value) then FPeriode2.Value:=DateToStr(FinDeMois(StrToDate(FPeriode1.Value))) ;
FDateCompta1.Text:=FPeriode1.Value ;
end;

procedure TFQRBLecSG.FPeriode2Change(Sender: TObject);
begin
  inherited;
if StrToDate(FPeriode2.Value)<StrToDate(FPeriode1.Value) then FPeriode1.Value:=DateToStr(DebutDeMois(StrToDate(FPeriode2.Value))) ;
FDateCompta2.Text:=FPeriode2.Value ;
end;

procedure TFQRBLecSG.DebToDate ;
BEGIN
FDateCompta1.Text:=FPeriode1.Value ;
END ;

procedure TFQRBLecSG.FinToDate ;
BEGIN
FDateCompta2.Text:=FPeriode2.Value ;
END ;

procedure TFQRBLecSG.FTauxExit(Sender: TObject);
begin
  inherited;
if FTaux.Text='' then Ftaux.Text:='0' ;
end;

procedure TFQRBLecSG.BSecondaireAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Fillchar(TotGen,SizeOf(TotGen),#0) ;
InitReport([3],CritEdt.BalBud.FormatPrint.Report) ;
end;

procedure TFQRBLecSG.BFPrimaireBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var StMontants : array[0..9] of String ;
    I          : Byte ;
begin
  inherited;
if CritEdt.Rupture=rRuptures then
   BEGIN
   if CritEdt.BalBud.RuptOnly=Avec then
      BEGIN
      (*
      if CritEdt.BalBud.OnlySecRupt then PrintBand:=(Qr1BS_SECTIONTRIE.AsString<>'') ;
      *)
      END Else PrintBand:=(CritEdt.BalBud.RuptOnly<>Sur) ;
   END ;
if CritEdt.Rupture=rLibres then PrintBand:=(CritEdt.BalBud.RuptOnly<>Sur) ;
BS_BUDSECT_.Caption:=MsgBox.Mess[15]+'  '+QR1BS_BUDSECT.AsString ;
if PrintBand then
   BEGIN
   Quoi:=QuoiCpt(1) ;
   Fillchar(StMontants,SizeOf(StMontants),#0) ;
   StPourMontantsBudgetEcart(TotSec,CritEdt,DebitPos,StMontants) ;
   For i:=1 to 10 do
       BEGIN
       TQRLabel(FindComponent('TotSec'+IntToStr(i))).Caption:=StMontants[i-1] ;
       END ;
    END ;
end;

procedure TFQRBLecSG.BFPrimaireAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([2],CritEdt.BalBud.FormatPrint.Report) ;
end;

procedure TFQRBLecSG.FGene1Change(Sender: TObject);
Var AvecJokerS : Boolean ;
begin
  inherited;
AvecJokerS:=Joker(FGene1, FGene2, FJokerGene ) ;
TFaS.Visible:=Not AvecJokerS ;
TFGene.Visible:=Not AvecJokerS ;
TFJokerGene.Visible:=AvecJokerS ;
end;

procedure TFQRBLecSG.QSecondAfterOpen(DataSet: TDataSet);
begin
  inherited;
QR12BG_BUDGENE       :=TStringField(QSecond.FindField('BG_BUDGENE')) ;
QR12BG_RUB           :=TStringField(QSecond.FindField('BG_RUB')) ;
QR12BG_LIBELLE       :=TStringField(QSecond.FindField('BG_LIBELLE')) ;
end;

Procedure PrepCalcBud(CritEdt : TCritEdt ; Var QMvt : TQuery)  ;
BEGIN
QMvt:=TQuery.Create(Application) ; QMvt.DataBaseName:='SOC' ; QMvt.SQL.Clear ;
QMvt.SQL.Clear ;
QMvt.SQL.Add('Select BE_BUDGENE, BE_BUDSECT, BE_EXERCICE, BE_DATECOMPTABLE, BE_NUMEROPIECE,') ;
QMvt.SQL.Add(       'BE_REFINTERNE, BE_ETABLISSEMENT, BE_LIBELLE, BE_VALIDE,') ;
QMvt.SQL.Add(       'BE_BUDJAL,BE_NATUREBUD, BE_QUALIFPIECE, ') ;
Case CritEdt.Monnaie of
  0 : BEGIN QMvt.SQL.Add('BE_DEBIT as DEBIT,BE_CREDIT as CREDIT') ; END ;
//  1 : BEGIN QEcr.SQL.Add('E_DEBITDEV as DEBIT,E_CREDITDEV as CREDIT') ; END ;
//  2 : BEGIN QMvt.SQL.Add('BE_DEBITEURO as DEBIT,BE_CREDITEURO as CREDIT') ; END ;
end ;
QMvt.SQL.Add(' From BUDECR ') ;
QMvt.SQL.Add(' Where BE_BUDSECT=:C1 and BE_BUDGENE=:C2 ') ;
QMvt.SQL.Add(' And BE_AXE="'+CritEdt.BalBud.AXE+'" ') ;
QMvt.SQL.Add(' And BE_BUDJAL="'+CritEdt.BalBud.Journal+'" ') ;
If CritEdt.BalBud.NatureBud<>'' then QMvt.SQL.Add('And BE_NATUREBUD="'+CritEdt.BalBud.NatureBud+'" ') ;
QMvt.SQL.Add('And BE_DATECOMPTABLE>="'+usdatetime(CritEdt.DateDeb)+'" And BE_DATECOMPTABLE<="'+usdatetime(CritEdt.DateFin)+'" ') ;
QMvt.SQL.Add('And BE_EXERCICE>="'+CritEdt.BalBud.Exo1+'" And BE_EXERCICE<="'+CritEdt.BalBud.Exo2+'" ') ;
QMvt.SQL.Add(TraduitNatureEcr(CritEdt.Qualifpiece, 'BE_QUALIFPIECE', true, CritEdt.ModeRevision)) ;
if CritEdt.Etab<>'' then QMvt.SQL.Add(' And BE_ETABLISSEMENT="'+CritEdt.Etab+'"') ;
if CritEdt.Valide<>'g' then QMvt.SQL.Add(' And BE_VALIDE="'+CritEdt.Valide+'" ') ;
If CritEDT.SQLPlus<>'' Then QMvt.SQL.Add(CritEDT.SQLPlus) ;
QMvt.SQL.Add(' Order By BE_BUDSECT,BE_AXE,BE_BUDGENE,BE_EXERCICE,BE_DATECOMPTABLE,BE_BUDJAL,BE_NUMEROPIECE,BE_QUALIFPIECE ') ;
ChangeSql(QMvt) ;
END ;


Procedure InitCalcBud(C1, C2 : String ; Var Tot1 : TMontTot ; TabD : TTabDate12 ; CritEdt : TCritEdt ; Q : TQuery) ;
Var DateCompta : TDateTime ;
    LaNatBud : String ;
    DEBIT, CREDIT : Double ;
BEGIN
Q.Close ;
Q.PAramByName('C1').AsString:=C1 ;
Q.PAramByName('C2').AsString:=C2 ;
//Q.Prepare ;
PrepareSQLODBC(Q) ;
Q.Open ;
Fillchar(Tot1,SizeOf(Tot1),#0) ;
If Q.Eof then Exit ;
While Not Q.eof do
      BEGIN
      DateCompta:=Q.FindField('BE_DATECOMPTABLE').AsDateTime ;
      LaNatBud:=Q.FindField('BE_NATUREBUD').AsString ;
      DEBIT:=Q.FindField('DEBIT').AsFloat ;
      CREDIT:=Q.FindField('CREDIT').AsFloat ;
      Reevaluation(DEBIT, CREDIT,CritEdt.BalBud.Resol, CritEdt.BalBud.Taux) ;
      if (DateCompta>=Critedt.Date1) and (DateCompta<Critedt.Date2) then
         BEGIN  CalculBud(0,Tot1, LaNatBud, DEBIT, CREDIT, CritEdt.Decimale) END ; {Calculs Période}
(* GP le 26/02/98
      if (DateCompta>=Critedt.Date1) and (DateCompta<Critedt.DateFin) then
         BEGIN  CalculBud(1,Tot1, LaNatBud, DEBIT, CREDIT, CritEdt.Decimale) END ; {Calculs Cumulé}
*)
      if (DateCompta>=Critedt.DateDeb) and (DateCompta<Critedt.Date2) then
         BEGIN  CalculBud(1,Tot1, LaNatBud, DEBIT, CREDIT, CritEdt.Decimale) END ; {Calculs Cumulé}
      if (DateCompta>=Critedt.DateDeb) and (DateCompta<Critedt.DateFin) then
         BEGIN  CalculBud(2,Tot1, LaNatBud, DEBIT, CREDIT, CritEdt.Decimale) END ; {Calculs Annuel}
      Q.Next ;
      END ;
END ;

Procedure CalculBud(i : Byte ; Var Tot1 : TMontTot ; NatB : String ; D,C : Double ; Dec : Byte) ;
Var NB : String ;
BEGIN
NB:=NatB ;
If NB='INI' then
   BEGIN
   Tot1[0][i].TotDebit:=  Arrondi(Tot1[0][i].TotDebit+D, Dec) ;
   Tot1[0][i].TotCredit:= Arrondi(Tot1[0][i].TotCredit+C, Dec) ;
   Tot1[0][12].TotDebit:= Arrondi(Tot1[0][12].TotDebit+D, Dec) ;
   Tot1[0][12].TotCredit:=Arrondi(Tot1[0][12].TotCredit+C, Dec) ;
   END Else
If Copy(NB,1,2)='DM' then
   BEGIN
   Tot1[1][i].TotDebit:=  Arrondi(Tot1[1][i].TotDebit+D, Dec) ;
   Tot1[1][i].TotCredit:= Arrondi(Tot1[1][i].TotCredit+C, Dec) ;
   Tot1[1][12].TotDebit:= Arrondi(Tot1[1][12].TotDebit+D, Dec) ;
   Tot1[1][12].TotCredit:=Arrondi(Tot1[1][12].TotCredit+C, Dec) ;
   END ;
END ;

procedure TFQRBLecSG.FormCreate(Sender: TObject);
begin
  inherited;
ListeCodesRupture:=TStringList.Create ;
end;

procedure TFQRBLecSG.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
ListeCodesRupture.Free ;
end;

procedure TFQRBLecSG.BHCptBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
BG_BUDGENE.Caption:=Qr12BG_BUDGENE.AsString ;
BG_LIBELLE.Caption:=Qr12BG_LIBELLE.AsString ;
If PrintBand then Quoi:=QuoiCpt(2) ;
if ((CritEdt.Rupture=rLibres) Or (CritEdt.Rupture=rRuptures)) And (CritEdt.BalBud.AvecCptSecond)
   then Quoi:=QuoiCptRecap(BG_BUDGENE2.Caption,BG_LIBELLE2.Caption) ;
end;

procedure TFQRBLecSG.DLHGenNeedData(var MoreData: Boolean);
begin
  inherited;
MoreData:=OkEnTeteGen ;
OkEnTeteGen:=False ;
end;

procedure TFQRBLecSG.MontantIR ;
Var StMontants : array[0..9] of String ;
    I          : Byte ;
BEGIN
Fillchar(StMontants,SizeOf(StMontants),#0) ;
StPourMontantsBudgetEcart(TotIR,CritEdt,DebitPos,StMontants) ;
For i:=1 to 10 do
    BEGIN
    TQRLabel(FindComponent('IRGenCum'+IntToStr(i))).Caption:=StMontants[i-1] ;
    END ;
END ;

procedure TFQRBLecSG.FNatBudChange(Sender: TObject);
begin
  inherited;
FTotIniRev.Enabled:=(FNatBud.Value='') ; //FTotIniRev.Checked:=(FNatBud.Value='') ;
end;

procedure TFQRBLecSG.DLHLibreNeedData(var MoreData: Boolean);
begin
  inherited;
MoreData:=OkEnteteLibre ;
OkEnteteLibre:=False ;
if MoreData And ((CritEdt.Rupture=rLibres) Or (CritEdt.Rupture=rRuptures)) then
   if CritEdt.BalBud.AvecCptSecond then OkRappelEnteteLibre:=TRUE ;
end;
(*
procedure TFQRBLecSG.StMontantsBudEcart(M : TabTot12 ; var AffM : array of String) ;
Var EcartPer, EcartCum, Pourcentage : Double ;
BEGIN
EcartPer:=Arrondi((M[0].TotDebit-M[0].TotCredit)-(M[1].TotDebit-M[1].TotCredit), CritEdt.Decimale) ;
EcartCum:=Arrondi((M[4].TotDebit-M[4].TotCredit)-(M[5].TotDebit-M[5].TotCredit), CritEdt.Decimale) ;
{Total compte en période}
AffM[0]:=PrintSoldeFormate(M[0].TotDebit, M[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
AffM[1]:=PrintSoldeFormate(M[1].TotDebit, M[1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;

AffM[2]:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Abs(EcartPer), CritEDT.AfficheSymbole ) ;
if (AffM[0]<>'') and (AffM[2]<>'') then Pourcentage:=EcartPer*100/(M[0].TotDebit-M[0].TotCredit)
                                   Else Pourcentage:=0 ;
AffM[3]:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Abs(Pourcentage) , CritEDT.AfficheSymbole ) ;
{Total compte en Cumulé}
AffM[4]:=PrintSoldeFormate(M[4].TotDebit, M[4].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
AffM[5]:=PrintSoldeFormate(M[5].TotDebit, M[5].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
AffM[6]:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Abs(EcartCum), CritEDT.AfficheSymbole ) ;
if (AffM[4]<>'') and (AffM[6]<>'') then Pourcentage:=Abs(EcartCum)*100/(M[4].TotDebit-M[4].TotCredit)
                                   Else Pourcentage:=0 ;
AffM[7]:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Abs(Pourcentage) , CritEDT.AfficheSymbole ) ;
{Total compte en Annuel}
AffM[8]:=PrintSoldeFormate(M[8].TotDebit, M[8].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
{Total compte en Extrapolé}
AffM[9]:='' ;
END ;
*)
procedure TFQRBLecSG.BHRecapBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
If PrintBand then Quoi:=QuoiCpt(2) ;
if ((CritEdt.Rupture=rLibres) Or (CritEdt.Rupture=rRuptures)) And (CritEdt.BalBud.AvecCptSecond)
   then Quoi:=QuoiCptRecap(BG_BUDGENE2.Caption,BG_LIBELLE2.Caption) ;
end;

procedure TFQRBLecSG.DLHRecapNeedData(var MoreData: Boolean);
begin
  inherited;
MoreData:=OkEnTeteRecap ;
OkEnTeteRecap:=False ;
end;

procedure TFQRBLecSG.DLRecapNeedData(var MoreData: Boolean);
var TotRecap         : TabMont77 ;
    LibRecap, CodRecap : String ;
    Chaque,i         : Byte ;
    TotMontRecap     : TabTot12 ;
    StMontRecap      : array[0..9] of String ;
begin
  inherited;
MoreData:=False ;
if (CritEdt.Rupture=rLibres) Or (CritEdt.Rupture=rRuptures)then
   BEGIN
   if CritEdt.BalBud.AvecCptSecond then MoreData:=PrintRecap(LRecap,CodRecap,LibRecap,TotRecap) ;
   if MoreData then
      BEGIN
      { Affichage sur la Bande }
      BG_BUDGENE2.Caption:=CodRecap ;
      If (CritEdt.Rupture=rRuptures) Then BG_LIBELLE2.Caption:=LibRecap Else BG_LIBELLE2.Caption:=LibRecap+StRappelLibTableLibre ;
      For Chaque:=1 to 5 do { Ini, Rév, Ini+Rev, QIni, QRev }
          BEGIN
          for i:=0 to 9 do
              BEGIN
              Case Chaque of
                 1 :  BEGIN { Initial }
                      TotMontRecap[i].TotDebit:=TotRecap[i*2] ; TotMontRecap[i].TotCredit:=TotRecap[i*2+1] ;
                      END ;
                 2 :  BEGIN { Révisé }
                      TotMontRecap[i].TotDebit:=TotRecap[(i*2)+20] ; TotMontRecap[i].TotCredit:=TotRecap[(i*2)+(1+20)] ;
                      END ;
                 3 :  BEGIN { Total Ini+Rév }
//                      TotMontRecap[i].TotDebit:=TotRecap[i*2]+TotRecap[(i*2)+20] ; TotMontRecap[i].TotCredit:=TotRecap[i*2+1]+TotRecap[(i*2)+(1+20)] ;
                      if (i=1) or (i=5) then
                         BEGIN
                         TotMontRecap[i].TotDebit:=TotRecap[i*2] ; TotMontRecap[i].TotCredit:=TotRecap[i*2+1]
                         END Else
                         BEGIN
                         TotMontRecap[i].TotDebit:=TotRecap[i*2]+TotRecap[(i*2)+20] ; TotMontRecap[i].TotCredit:=TotRecap[i*2+1]+TotRecap[(i*2)+(1+20)] ;
                         END ;
                      END ;
                 End ;
              END ;
          Fillchar(StMontRecap,SizeOf(StMontRecap),#0) ;
          StPourMontantsBudgetEcart(TotMontRecap,CritEdt,DebitPos,StMontRecap) ;
          For i:=1 to 10 do
              BEGIN
              Case Chaque of
                 1 : TQRLabel(FindComponent('IniRecap'+IntToStr(i))).Caption:=StMontRecap[i-1] ;
                 2 : TQRLabel(FindComponent('RevRecap'+IntToStr(i))).Caption:=StMontRecap[i-1] ;
                 3 : TQRLabel(FindComponent('IRRecap'+IntToStr(i))).Caption:=StMontRecap[i-1] ;
                 4 : TQRLabel(FindComponent('QIniRecap'+IntToStr(i))).Caption:='' ;
                 5 : TQRLabel(FindComponent('QRevRecap'+IntToStr(i))).Caption:='' ;
                 End ;
              END ;
          END ;
      END ;
   END ;
OkEnTeteRecap:=MoreData ;
end;

procedure TFQRBLecSG.FPlanRupturesChange(Sender: TObject);
var Q1  : TQuery ; I : Integer ; St,St1 : String ;
begin
  inherited;
If QRLoading Then Exit ;
ST1:=AxeJal ; If St1='' Then Exit ; If Length(St1)<2 Then Exit ;
if FRuptures.Checked then  { Rempli la Fourchette des codes ruptures en 'Sur Rupture'  }
   BEGIN
   St:='' ;
   FCodeRupt1.Clear ; FCodeRupt2.clear ;
   (*
   Q1:=OpenSQL('Select RU_CLASSE from Rupture where RU_NATURERUPT="RU'+Char(FNatureCpt.Value[2])+'" and RU_PLANRUPT="'+FPlanRuptures.Value+'" '
             +'Order by RU_CLASSE ',TRUE) ;
   *)
   Q1:=OpenSQL('Select RU_CLASSE from Rupture where RU_NATURERUPT="RU'+Copy(St1,2,1)+'" and RU_PLANRUPT="'+FPlanRuptures.Value+'" '
             +'Order by RU_CLASSE ',TRUE) ;
   Q1.First ;
   i:=0 ;
   if ListeCodesRupture<>Nil then ListeCodesRupture.Clear ;
   While Not Q1.Eof Do
       BEGIN
       FCodeRupt1.Items.Add(Q1.Fields[0].AsString) ;
       FCodeRupt2.Items.Add(Q1.Fields[0].AsString) ;
       (* GP le 19/12/2000
       FCodeRupt1.Items[i]:=Q1.Fields[0].AsString ;
       FCodeRupt2.Items[i]:=Q1.Fields[0].AsString ;
       *)
       ListeCodesRupture.Add(Q1.Fields[0].AsString) ;
       Q1.Next ;
       Inc(i) ;
       END ;
   FCodeRupt1.ItemIndex:=0 ; FCodeRupt2.ItemIndex:=i-1 ;
   Ferme(Q1) ;
   END ;
end;

procedure TFQRBLecSG.FSansRuptClick(Sender: TObject);
begin
  inherited;
if Not FSurRupt.Checked then FAvecCptSecond.Checked:=False ;
FAvecCptSecond.Enabled:=FSurRupt.Checked ;
// Plus de menu if (QuelEtatBud=Nature) then FTablesLibres.Checked:=True ;
FRupturesClick(Nil);
end;

procedure TFQRBLecSG.FRupturesClick(Sender: TObject);
begin
  inherited;
FCodeRupt1.Enabled:=FSurRupt.Checked ;
FCodeRupt2.Enabled:=FCodeRupt1.Enabled ;
//if FCodeRupt1.Enabled then FPlanRupturesChange(sender) ;
if FRuptures.Checked then If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
(*
FOnlySecRupt.Enabled:=FAvecRupt.Checked and FRuptures.Checked ;
FOnlySecRupt.checked:=FAvecRupt.Checked and FRuptures.Checked ;
*)
//FAvecCptSecond.checked:=FSurRupt.Checked ;
FAvecCptSecond.Enabled:=FSurRupt.Checked ;
If Not FSurRupt.Checked Then FAvecCptSecond.checked:=FALSE ;
end;

procedure TFQRBLecSG.BNouvRechClick(Sender: TObject);
begin
  inherited;
FJournal.ItemIndex:=0 ;
//FRESOL.ItemIndex:=0 ;
InitResolution(FRESOL) ;

end;

procedure TFQRBLecSG.BRecapBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
if ((CritEdt.Rupture=rLibres) Or (CritEdt.Rupture=rRuptures)) And (CritEdt.BalBud.AvecCptSecond)
   then Quoi:=QuoiCptRecap(BG_BUDGENE2.Caption,BG_LIBELLE2.Caption) ;
end;

procedure TFQRBLecSG.BMultiAfterPrint(BandPrinted: Boolean);
begin
  inherited;
if (CritEdt.Rupture=rLibres) Or (CritEdt.Rupture=rRuptures) then
   if CritEdt.BalBud.AvecCptSecond then BEGIN VideRecap(LRecap) ; ChargeRecap(LRecap) ; END ;
end;

procedure TFQRBLecSG.DHLRappelLibreNeedData(var MoreData: Boolean);
begin
  inherited;
MoreData:=OkRappelEnteteLibre ;
OkRappelEnteteLibre:=False ;
end;

procedure TFQRBLecSG.FQueTotBudClick(Sender: TObject);
begin
  inherited;
FTotIniRev.Enabled:=Not FQueTotBud.Checked ; If FQueTotBud.Checked Then FTotIniRev.Checked:=FALSE ;
end;

procedure TFQRBLecSG.FTotIniRevClick(Sender: TObject);
begin
  inherited;
FQueTotBud.Enabled:=Not (FTotIniRev.Checked) ; If FTotIniRev.Checked Then FQueTotBud.Checked:=FALSE ;
end;

end.
