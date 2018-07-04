{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 31/03/2005
Modifié le ... : 24/01/2006
Description .. : Remplacé en eAGL par BALBUDTEGEN_TOF.PAS
Suite ........ : Lek: Remplace le nom de la procédure pour que le mode 
Suite ........ : 2/3 puisse appeler les deux procédures.
Suite ........ : BalBudteGen --> BalBudteGenQR
Mots clefs ... :
*****************************************************************}
unit QRBLteG;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Hctrls, Spin, Menus, hmsgbox, HQuickrp, StdCtrls,
  Buttons, ExtCtrls, Mask, Hcompte, ComCtrls, UtilEdt, HEnt1, Ent1,
  CpteUtil, EdtLegal, CritEDT, HQry, HSysMenu,UtilEdt1, QRRupt, Calcole,
  HTB97, HPanel, UiUtil, tCalcCum,
  uLibWindows ; // TraductionTHMultiValComboBox  // FQ 16536 SBO 18/10/2005

procedure BalBudteGenQR ;
procedure BalBudteGenNature ;
procedure BalBudteGenZoom(Crit : TCritEdt) ;

Type TabTot12 = Array[0..12] Of TabDC ;
Type TMontTot = Array[0..2] Of TabTot12 ;
type
  TFQRBLteG = class(TFQR)
    TitreColCpt: TQRLabel;
    Col1: TQRLabel;
    Col2: TQRLabel;
    col3: TQRLabel;
    Col4: TQRLabel;
    Col5: TQRLabel;
    Col6: TQRLabel;
    Col13: TQRLabel;
    TITRE1REP: TQRLabel;
    REPORTCOL1: TQRLabel;
    REPORTCOL2: TQRLabel;
    REPORTCOL3: TQRLabel;
    REPORTCOL4: TQRLabel;
    REPORTCOL5: TQRLabel;
    REPORTCOL6: TQRLabel;
    REPORTTOTAL1: TQRLabel;
    TITRE2REP: TQRLabel;
    REPORTCOL13: TQRLabel;
    REPORTCOL14: TQRLabel;
    REPORTCOL15: TQRLabel;
    REPORTCOL16: TQRLabel;
    REPORTCOL17: TQRLabel;
    REPORTCOL18: TQRLabel;
    REPORTCOL19: TQRLabel;
    MsgBox: THMsgBox;
    Bevel3: TBevel;
    BLibre: TQRBand;
    DLLibre: TQRDetailLink;
    TFJournal: THLabel;
    FJournal: THValComboBox;
    FPeriode1: THValComboBox;
    TFExercice_: TLabel;
    FExercice_: THValComboBox;
    FValide: TCheckBox;
    Col7: TQRLabel;
    Col8: TQRLabel;
    Col9: TQRLabel;
    Col10: TQRLabel;
    Col11: TQRLabel;
    Col12: TQRLabel;
    REPORTCOL7: TQRLabel;
    REPORTCOL8: TQRLabel;
    REPORTCOL9: TQRLabel;
    REPORTCOL10: TQRLabel;
    REPORTCOL11: TQRLabel;
    REPORTCOL12: TQRLabel;
    TRJournal: TQRLabel;
    RJournal: TQRLabel;
    QRLabel3: TQRLabel;
    RExercice_: TQRLabel;
    QRLabel17: TQRLabel;
    RValide: TQRLabel;
    TFRESOL: THLabel;
    FRESOL: THValComboBox;
    TFNatBud: THLabel;
    TFTaux: THLabel;
    TRNatBud: TQRLabel;
    RNatBud: TQRLabel;
    TRTaux: TQRLabel;
    RTaux: TQRLabel;
    REPORTCOL20: TQRLabel;
    REPORTCOL21: TQRLabel;
    REPORTCOL22: TQRLabel;
    REPORTCOL23: TQRLabel;
    REPORTCOL24: TQRLabel;
    REPORTTOTAL2: TQRLabel;
    FRealise: TCheckBox;
    QRLabel1: TQRLabel;
    RRealise: TQRLabel;
    BudTotal: TQRLabel;
    BudTotal1: TQRLabel;
    BudTotal2: TQRLabel;
    BudTotal3: TQRLabel;
    BudTotal4: TQRLabel;
    BudTotal5: TQRLabel;
    BudTotal6: TQRLabel;
    BudTotal7: TQRLabel;
    BudTotal8: TQRLabel;
    BudTotal9: TQRLabel;
    BudTotal10: TQRLabel;
    ReaTotal: TQRLabel;
    ReaTotal1: TQRLabel;
    ReaTotal2: TQRLabel;
    ReaTotal3: TQRLabel;
    ReaTotal4: TQRLabel;
    ReaTotal5: TQRLabel;
    ReaTotal6: TQRLabel;
    ReaTotal7: TQRLabel;
    ReaTotal8: TQRLabel;
    ReaTotal9: TQRLabel;
    ReaTotal10: TQRLabel;
    EcaTotal: TQRLabel;
    EcaTotal1: TQRLabel;
    EcaTotal2: TQRLabel;
    EcaTotal3: TQRLabel;
    EcaTotal4: TQRLabel;
    EcaTotal5: TQRLabel;
    EcaTotal6: TQRLabel;
    EcaTotal7: TQRLabel;
    EcaTotal8: TQRLabel;
    EcaTotal9: TQRLabel;
    EcaTotal10: TQRLabel;
    BudTotal11: TQRLabel;
    ReaTotal11: TQRLabel;
    EcaTotal11: TQRLabel;
    EcaTotal12: TQRLabel;
    ReaTotal12: TQRLabel;
    BudTotal12: TQRLabel;
    TotTotalB: TQRLabel;
    TotTotalR: TQRLabel;
    TotTotalE: TQRLabel;
    BudPositif: TCheckBox;
    IniGenCum5: TQRLabel;
    IniGenCum4: TQRLabel;
    IniGenCum2: TQRLabel;
    IniGenCum3: TQRLabel;
    IniGenCum6: TQRLabel;
    IniGenCum1: TQRLabel;
    IniGenCum7: TQRLabel;
    IniGenCum8: TQRLabel;
    IniGenCum9: TQRLabel;
    IniGenCum10: TQRLabel;
    IniGenCum11: TQRLabel;
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
    RevGenCum11: TQRLabel;
    ReaGen: TQRLabel;
    ReaGenCum1: TQRLabel;
    ReaGenCum2: TQRLabel;
    ReaGenCum3: TQRLabel;
    ReaGenCum4: TQRLabel;
    ReaGenCum5: TQRLabel;
    ReaGenCum6: TQRLabel;
    ReaGenCum7: TQRLabel;
    ReaGenCum8: TQRLabel;
    ReaGenCum9: TQRLabel;
    ReaGenCum10: TQRLabel;
    ReaGenCum11: TQRLabel;
    EcaGen: TQRLabel;
    EcaGenCum1: TQRLabel;
    EcaGenCum2: TQRLabel;
    EcaGenCum3: TQRLabel;
    EcaGenCum4: TQRLabel;
    EcaGenCum5: TQRLabel;
    EcaGenCum6: TQRLabel;
    EcaGenCum7: TQRLabel;
    EcaGenCum8: TQRLabel;
    EcaGenCum9: TQRLabel;
    EcaGenCum10: TQRLabel;
    EcaGenCum11: TQRLabel;
    IniGenCum12: TQRLabel;
    RevGenCum12: TQRLabel;
    ReaGenCum12: TQRLabel;
    EcaGenCum12: TQRLabel;
    IniGenTotal: TQRLabel;
    RevGenTotal: TQRLabel;
    ReaGenTotal: TQRLabel;
    EcaGenTotal: TQRLabel;
    BHCpt: TQRBand;
    HGene: TQRGroup;
    BG_BUDGENE: TQRDBText;
    TFBG_LIBELLE: TQRLabel;
    IniLibre: TQRLabel;
    IniLibre1: TQRLabel;
    IniLibre2: TQRLabel;
    IniLibre3: TQRLabel;
    IniLibre4: TQRLabel;
    IniLibre5: TQRLabel;
    IniLibre6: TQRLabel;
    IniLibre7: TQRLabel;
    IniLibre8: TQRLabel;
    IniLibre9: TQRLabel;
    IniLibre10: TQRLabel;
    IniLibre11: TQRLabel;
    RevLibre: TQRLabel;
    RevLibre1: TQRLabel;
    RevLibre2: TQRLabel;
    RevLibre3: TQRLabel;
    RevLibre4: TQRLabel;
    RevLibre5: TQRLabel;
    RevLibre6: TQRLabel;
    RevLibre7: TQRLabel;
    RevLibre8: TQRLabel;
    RevLibre9: TQRLabel;
    RevLibre10: TQRLabel;
    RevLibre11: TQRLabel;
    ReaLibre: TQRLabel;
    ReaLibre1: TQRLabel;
    ReaLibre2: TQRLabel;
    ReaLibre3: TQRLabel;
    ReaLibre4: TQRLabel;
    ReaLibre5: TQRLabel;
    ReaLibre6: TQRLabel;
    ReaLibre7: TQRLabel;
    ReaLibre8: TQRLabel;
    ReaLibre9: TQRLabel;
    ReaLibre10: TQRLabel;
    ReaLibre11: TQRLabel;
    EcaLibre: TQRLabel;
    EcaLibre1: TQRLabel;
    EcaLibre2: TQRLabel;
    EcaLibre3: TQRLabel;
    EcaLibre4: TQRLabel;
    EcaLibre5: TQRLabel;
    EcaLibre6: TQRLabel;
    EcaLibre7: TQRLabel;
    EcaLibre8: TQRLabel;
    EcaLibre9: TQRLabel;
    EcaLibre10: TQRLabel;
    EcaLibre11: TQRLabel;
    IniLibre12: TQRLabel;
    RevLibre12: TQRLabel;
    ReaLibre12: TQRLabel;
    EcaLibre12: TQRLabel;
    IniLibreTotal: TQRLabel;
    RevLibreTotal: TQRLabel;
    ReaLibreTotal: TQRLabel;
    EcaLibreTotal: TQRLabel;
    BLibreHaut: TQRBand;
    TCodRupt: TQRDBText;
    TLibRupt: TQRLabel;
    DLHLibre: TQRDetailLink;
    TFSoldeFormate: THLabel;
    FSoldeFormate: THValComboBox;
    FTaux: THNumEdit;
    TRResol: TQRLabel;
    RResol: TQRLabel;
    FPeriode2: THValComboBox;
    FNATBUD: THMultiValComboBox;   // FQ 16536 SBO 18/10/2005
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure DLLibreNeedData(var MoreData: Boolean);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure FJournalChange(Sender: TObject);
    procedure FExerciceChange(Sender: TObject);
    procedure FExercice_Change(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FPeriode1Change(Sender: TObject);
    procedure FPeriode2Change(Sender: TObject);
    procedure BLibreBeforePrint(var PrintBand: Boolean;  var Quoi: string);
    procedure BHCptBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure DLHLibreNeedData(var MoreData: Boolean);
    procedure FSansRuptClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
  private    { Déclarations privées }
    QBud                      : TQuery ;
    DebitPos                  : Boolean ;
    ExoDebJal, ExoFinJal, AxeJal : String3 ;
    DatDebJal, DatFinJal      : TDateTime ;
    LLibre                    : TStringList ;
    TotGen, TotEdt            : TMontTot ;
    TabDate                   : TTabDate12 ;
    Qr1BG_BUDGENE, Qr1BG_RUB, Qr1BG_LIBELLE : TStringField ;
    OkEnteteLibre,
    Affiche                      : Boolean ;
    QuelEtatBud                  : TQuelEtatBud ;
    NbMois                       : Word ;
    Function  QuoiCpt(i : Integer) : String ;
    Procedure BalBudGenZoom(Quoi : String) ;
//    Procedure Calculs(var Report1,Report2 : TReport; i : integer; Bool1, Bool2 : Boolean ) ;
    procedure CalculDateBud ;
    procedure InfosJournal ;
    procedure DebToDate ;
    procedure FinToDate ;
    Function  FourchetteExoOk : Boolean ;
  public
    procedure GenereSQL ; Override ;
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
  end;

Procedure PrepCalcBud(CritEdt : TCritEdt ; Var QMvt : TQuery)  ;
Procedure InitCalcBud(C1 : String ; Var Tot1 : TMontTot ; TabD : TTabDate12 ; CritEdt : TCritEdt ; Q : TQuery) ;
Procedure CalculBud(i : Byte ; Var Tot1 : TMontTot ; NatB : String ; D,C : Double ; Dec : Byte) ;

implementation

var ChargeFiltre : Boolean;

{$R *.DFM}

procedure BalBudteGenQR ;
var QR : TFQRBLteG ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQRBLteG.Create(Application) ;
Edition.Etat:=etBlBudgeteGen ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalBudGenZoom ;
QR.QuelEtatBud:=Normal ;
QR.InitType (nbBudGen,neBalBud,msGenBEcr,'QRBLteGNor','',TRUE,FALSE,FALSE,Edition) ;
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

procedure BalBudteGenNature ;
var QR : TFQRBLteG ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQRBLteG.Create(Application) ;
Edition.Etat:=etBlBudgeteGen ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalBudGenZoom ;
QR.QuelEtatBud:=Nature ;
QR.InitType (nbBudGen,neBalBud,msGenBEcr,'QRBLteGNat','',TRUE,FALSE,FALSE,Edition) ;
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

procedure BalBudteGenZoom(Crit : TCritEdt) ;
var QR : TFQRBLteG ;
    Edition : TEdition ;
BEGIN
QR:=TFQRBLteG.Create(Application) ;
Edition.Etat:=etBlBudgeteGen ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.BalBudGenZoom ;
 QR.QuelEtatBud:=Normal ;
 QR.InitType (nbBudGen,neGlBud,msGenBEcr,'QRBLteGNor','',FALSE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

Function TFQRBLteG.QuoiCpt(i : Integer) : String ;
BEGIN
Case i of
   0 : Result:=Q.FindField('BG_BUDGENE').AsString+' '+Q.FindField('BG_LIBELLE').AsString ;
  End ;
END ;

Procedure TFQRBLteG.BalBudGenZoom(Quoi : String) ;
BEGIN
ZoomEdt(8,Quoi) ;
END ;

procedure TFQRBLteG.FormShow(Sender: TObject);
begin
HelpContext:=15321000 ;
Standards.HelpContext:=15321100 ;
Avances.HelpContext:=15321200 ;
Mise.HelpContext:=15321300 ;
Option.HelpContext:=15321400 ;
TabRuptures.HelpContext:=15321500 ;
TabSup.TabVisible:=False ;
If FJournal.Values.Count>0 Then FJournal.Value:=FJournal.Values[0] ;
// FQ 16536 SBO 18/10/2005 Mise en place des options de révision dans les balances budgétaires
FNatBud.Text := TraduireMemoire('<<Tous>>') ;
FSoldeFormate.ItemIndex:=0 ;
  inherited;
//FRESOL.Value:='C'      ;
InitResolution(FRESOL) ;
FLigneCpt.Checked:=True ;
TFGen.Parent:=Avances  ; TFGenJoker.Parent:=Avances ;
FJoker.Parent:=Avances ; TFaG.Parent:=Avances ;
FCpte1.Parent:=Avances ; FCpte2.Parent:=Avances ;
FCpte1.TabOrder:=0 ; FCpte2.TabOrder:=1 ; FJoker.TabOrder:=2 ;
DebToDate ; FinToDate  ;
FOnlyCptAssocie.Enabled:=False ;
If (EstSerie(S3)) Then FGroupChoixRupt.Visible:=FALSE ;
AvecRevision.Visible := False ; // FQ 16536 SBO 18/10/2005
end;

Procedure TFQRBLteG.FinirPrint ;
BEGIN
Inherited ;
if (CritEdt.Rupture=rLibres) then VideGroup(LLibre) ;
if CritEdt.Balbud.FormatPrint.PrSepMontant then FreeLesLignes(Self) ;
END ;

procedure TFQRBLteG.InitDivers ;
Var i : Byte ;
    St : String ;
BEGIN
Inherited ;
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
For i:=1 to 12 do TQRLabel(FindComponent('Col'+InttoStr(i))).Caption:='' ;

{ Calcul des différentes fourchettes de dates }
CalculDateBud ;
{ Titres des colonnes de montants }
For i:=1 to NbMois do
    BEGIN
    St:=FormatDatetime('mmm yy',TabDate[i]) ; TQRLabel(FindComponent('Col'+InttoStr(i))).Caption:=FirstMajuscule(St) ;
    END ;
//BDetail.Frame.DrawBottom:=CritEdt.BalBud.FormatPrint.PrSepCompte[1] ;
CritEdt.BalBud.FormatPrint.ColMin:=1 ;
CritEdt.BalBud.FormatPrint.ColMax:=13 ;
END ;

procedure TFQRBLteG.CalculDateBud ;
Var i : Integer ;
    LaDate : TDateTime ;
    Pm, PA : Word ;
BEGIN
Fillchar(TabDate,SizeOf(TabDate),#0) ;
LaDate:=StrToDate(FPeriode1.Value) ;
NbMois:=0 ;
NOMBREMOIS (CritEdt.Date1, CritEdt.Date2, Pm, PA, NbMois) ;
If NbMois>13 then NbMois:=12 ;
For i:=1 to NbMois do
    BEGIN
    TabDate[i]:=PlusMois(LaDate,i-1) ;
    END ;
END ;

procedure TFQRBLteG.GenereSQL ;
{ Construction de la requête SQL en fonction du multicritère }
Var St : String ;
BEGIN
Inherited ;
Q.Close ; Q.SQL.Clear ;
Q.SQL.Add(' Select * ') ;
Q.SQL.Add(' From BUDGENE G Where ') ;
//Q.SQL.Add(' BG_BUDGENE <> "'+w_w+'" ') ;
Q.SQL.Add(ExistBud(fbBudgen,CritEdt.BalBud.MvtSur,CritEdt.BalBud.Journal,CritEdt.BalBud.Axe,True)) ;
if CritEdt.Joker then Q.SQL.Add(' AND BG_BUDGENE like "'+TraduitJoker(CritEdt.Cpt1)+'" ') Else
   BEGIN
   if CritEdt.Cpt1<>'' then Q.SQL.Add(' AND BG_BUDGENE>="'+CritEdt.Cpt1+'" ') ;
   if CritEdt.Cpt2<>'' then Q.SQL.Add(' AND BG_BUDGENE<="'+CritEdt.Cpt2+'" ') ;
   END ;
If CritEdt.Rupture=rLibres then
   BEGIN
   St:=WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,fbBudgen,CritEdt.BalBud.OnlyCptAssocie) ;
   If St<>'' Then Q.SQL.Add(St) ;
   Q.SQL.Add('Order By '+OrderLibre(CritEdt.LibreTrie)+'BG_BUDGENE ') ;
   END Else Q.SQL.Add(' Order By BG_BUDGENE') ;
ChangeSQL(Q) ; Q.Open ;
END ;

procedure TFQRBLteG.RenseigneCritere ;
BEGIN
Inherited ;
RJournal.Caption:=FJournal.Text ;
RExercice_.Caption:=FExercice_.Text ;
CaseACocher(FValide,RValide) ;
CaseACocher(FRealise,RRealise) ;
RNatBud.Caption := CritEdt.BalBud.NatureLib ;   // FQ 16536 SBO 18/10/2005 Mise en place des options de révision dans les balances budgétaires
RTaux.Caption:=FloatToStr(FTaux.Value) ;
RResol.Caption:=FResol.Text ;
AfficheBudgetEn(RDevises,CritEdt) ;
END ;

procedure TFQRBLteG.ChoixEdition ;
{ Initialisation des options d'édition }
BEGIN
Inherited ;
if FinDeMois(PlusMois(CritEdt.Date1,11))<CritEdt.Date2 then
   BEGIN
   MsgRien.Execute(8,Caption,'') ; FDateCompta2.Text:=FPeriode2.Value ;
   END ;
if CritEdt.Rupture=rLibres then
   BEGIN
   ChargeGroup(LLibre,['B00','B01','B02','B03','B04','B05','B06','B07','B08','B09']) ;
   END ;

RetailleBandesBudgetes(Self,CritEdt,Msgbox.Mess[3],BDetail,Blibre,nil,nil,BFinEtat) ;

If CritEdt.BalBud.Resol='C' then CritEdt.Decimale:=V_PGI.OkDecV else
If CritEdt.BalBud.Resol='F' then CritEdt.Decimale:=0 else
If CritEdt.BalBud.Resol='K' then CritEdt.Decimale:=0 else
If CritEdt.BalBud.Resol='M' then CritEdt.Decimale:=0 ;
if CritEdt.Balbud.FormatPrint.PrSepMontant then AffecteLigne(Self,EntetePage,CritEdt.BalBud.FormatPrint) ;
END ;

procedure TFQRBLteG.RecupCritEdt ;
var lStWhere : String ;
    lStLib   : String ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  BalBud.Journal:=FJournal.Value ;
  BalBud.Axe:=AxeJal ;

  // FQ 16536 SBO 18/10/2005 Mise en place des options de révision dans les balances budgétaires
  TraductionTHMultiValComboBox( FNatBud, lStWhere, lStLib, 'BE_NATUREBUD' ) ;
  BalBud.NatureBud := lStWhere ;
  BalBud.NatureLib := lStLib ;
  // Fin FQ 16536 SBO 18/10/2005 Mise en place des options de révision dans les balances budgétaires

  BalBud.Exo1:=QUELEXODTBud(DatDebJal) ; BalBud.Exo2:=QUELEXODTBud(DatFinJal) ;
  Date1:=StrToDate(FDateCompta1.Text)    ; Date2:=StrToDate(FDateCompta2.Text) ;
  DateDeb:=DAte1 ; DateFin:=DAte2 ;
  if FValide.State=cbGrayed then Valide:='g' ;
  if FValide.State=cbChecked then Valide:='X' ;
  if FValide.State=cbUnChecked then Valide:='-' ;
  BalBud.Resol:=FResol.Value[1] ; BalBud.Taux:=FTaux.Value ;
  BalBud.Realise:=FRealise.Checked ;
  BalBud.RuptOnly:=QuelleTypeRupt(0,FSAnsRupt.Checked,FAvecRupt.Checked,FSurRupt.Checked) ;
  BalBud.OnlyCptAssocie:=(Rupture<>rRien) and FOnlyCptAssocie.Checked ;
//  BalBud.SoldeFormate:=FSoldeFormate.Value ;
  If DebitPos Then BalBud.SoldeFormate:='PD' Else BalBud.SoldeFormate:='PC' ;
///  BalBud.Mvtsur:=
  With BalBud.FormatPrint Do
     BEGIN
     END ;
  END ;
END ;

function TFQRBLteG.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ;
if Result then result:=FourchetteExoOk ;
if Result then PrepCalcBud(CritEdt,QBUD)  ;
END ;

Function TFQRBLteG.FourchetteExoOk : Boolean ;
BEGIN
Result:=(CritEdt.BalBud.Exo1>=ExoDebJal) and (CritEdt.BalBud.Exo2<=ExoFinJal) ;
If Not Result then NumErreurCrit:=9 ;
END ;

{Procedure TFQRBLteG.Calculs(var Report1,Report2 : TReport; i : integer; Bool1, Bool2 : Boolean ) ;
BEGIN
END ;}

procedure TFQRBLteG.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var Realise, Ecar   : TabTot12 ;
    I              : Byte ;
    C1, C2, C3, C4 : TComponent ;
    TD             : TabloExt ;
    LeCompte, Lexo : String ;
    FinTemp,DateTemp : TDateTime ;
    MReport        : TabTRep ;
    TotBud         :  TMontTot ;
    TotLibre       : Array[0..77] of Double ;
begin
  inherited;
Fillchar(realise,SizeOf(realise),#0) ;
Fillchar(Ecar,SizeOf(Ecar),#0) ;
Fillchar(TotGen,SizeOf(TotGen),#0) ;
Fillchar(MReport,SizeOf(MReport),#0) ;
Fillchar(TotBud,SizeOf(TotBud),#0) ;
Fillchar(TotLibre,SizeOf(TotLibre),#0) ;
PrintBand:=Affiche ;
If Not PrintBand then Exit ;
InitCalcBud(Qr1BG_BUDGENE.AsString,TotBud,TabDate,CritEdt,QBud) ;
TotGen:=TotBud ;
For i:=0 to 12 do
    BEGIN
    if (CritEdt.BalBud.RuptOnly<>Sur) then
       BEGIN
       TotEdt[0][i].TotDebit:=Arrondi(TotEdt[0][i].TotDebit+(TotBud[0][i].TotDebit+TotBud[1][i].TotDebit), CritEdt.Decimale) ;
       TotEdt[0][i].TotCredit:=Arrondi(TotEdt[0][i].TotCredit+(TotBud[0][i].TotCredit+TotBud[1][i].TotCredit), CritEdt.Decimale) ;
       END ;
    MReport[i+1].TotDebit:=Arrondi(MReport[i+1].TotDebit+(TotBud[0][i].TotDebit+TotBud[1][i].TotDebit), CritEdt.Decimale) ;
    MReport[i+1].TotCredit:=Arrondi(MReport[i+1].TotCredit+(TotBud[0][i].TotCredit+TotBud[1][i].TotCredit), CritEdt.Decimale) ;
    (* GP le 25/06/99
    TotLibre[i*2]:=Arrondi(TotLibre[i*2]+(TotBud[0][i].TotDebit+TotBud[1][i].TotDebit), CritEdt.Decimale) ;
    TotLibre[i*2+1]:=Arrondi(TotLibre[i*2+1]+(TotBud[0][i].TotCredit+TotBud[1][i].TotCredit), CritEdt.Decimale) ;
    *)
    TotLibre[i*2]:=Arrondi(TotLibre[i*2]+(TotBud[0][i].TotDebit), CritEdt.Decimale) ;
    TotLibre[i*2+1]:=Arrondi(TotLibre[i*2+1]+(TotBud[0][i].TotCredit), CritEdt.Decimale) ;
    TotLibre[(i*2)+26]:=Arrondi(TotLibre[(i*2)+26]+(TotBud[1][i].TotDebit), CritEdt.Decimale) ;
    TotLibre[(i*2)+1+26]:=Arrondi(TotLibre[(i*2)+26+1]+(TotBud[1][i].TotCredit), CritEdt.Decimale) ;
    END ;
if critEdt.BalBud.Realise then
   BEGIN
   { Calcul réalisé }
   LeCompte:='CBG'+CritEdt.BalBud.Journal+Qr1BG_RUB.AsString ;
   LExo:=QUELEXODTBud(TabDate[1]) ; DateTemp:=PlusMois(TabDate[1],1) ; if (Lexo<>QUELEXODTBud(DateTemp)) then LExo:='' ; FinTemp:=FinDeMois(TabDate[1]) ;
   GetCumul('RUBREA',LeCompte,LeCompte,'SAN',CritEdt.Etab,CritEDT.DeviseAffichee,LExo,TabDate[1],FinTemp,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
   realise[0].TotDebit:=TD[3] ; realise[0].TotCredit:=TD[2] ;

   Lexo:=QUELEXODTBud(TabDate[2]) ; DateTemp:=PlusMois(TabDate[2],1) ; if (LExo<>QUELEXODTBud(DateTemp)) then LExo:='' ; FinTemp:=FinDeMois(TabDate[2]) ;
   GetCumul('RUBREA',LeCompte,LeCompte,'SAN',CritEdt.Etab,CritEDT.DeviseAffichee,LExo,TabDate[2],FinTemp,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
   realise[1].TotDebit:=TD[3] ; realise[1].TotCredit:=TD[2] ;

   LExo:=QUELEXODTBud(TabDate[3]) ; DateTemp:=PlusMois(TabDate[3],1) ; if (LExo<>QUELEXODTBud(DateTemp)) then LExo:='' ; FinTemp:=FinDeMois(TabDate[3]) ;
   GetCumul('RUBREA',LeCompte,LeCompte,'SAN',CritEdt.Etab,CritEDT.DeviseAffichee,LExo,TabDate[3],FinTemp,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
   realise[2].TotDebit:=TD[3] ; realise[2].TotCredit:=TD[2] ;

   LExo:=QUELEXODTBud(TabDate[4]) ; DateTemp:=PlusMois(TabDate[4],1) ; if (LExo<>QUELEXODTBud(DateTemp)) then LExo:='' ; FinTemp:=FinDeMois(TabDate[4]) ;
   GetCumul('RUBREA',LeCompte,LeCompte,'SAN',CritEdt.Etab,CritEDT.DeviseAffichee,LExo,TabDate[4],FinTemp,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
   realise[3].TotDebit:=TD[3] ; realise[3].TotCredit:=TD[2] ;

   LExo:=QUELEXODTBud(TabDate[5]) ; DateTemp:=PlusMois(TabDate[5],1) ; if (LExo<>QUELEXODTBud(DateTemp)) then LExo:='' ; FinTemp:=FinDeMois(TabDate[5]) ;
   GetCumul('RUBREA',LeCompte,LeCompte,'SAN',CritEdt.Etab,CritEDT.DeviseAffichee,LExo,TabDate[5],FinTemp,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
   realise[4].TotDebit:=TD[3] ; realise[4].TotCredit:=TD[2] ;

   LExo:=QUELEXODTBud(TabDate[6]) ; DateTemp:=PlusMois(TabDate[6],1) ; if (LExo<>QUELEXODTBud(DateTemp)) then LExo:='' ; FinTemp:=FinDeMois(TabDate[6]) ;
   GetCumul('RUBREA',LeCompte,LeCompte,'SAN',CritEdt.Etab,CritEDT.DeviseAffichee,LExo,TabDate[6],FinTemp,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
   realise[5].TotDebit:=TD[3] ; realise[5].TotCredit:=TD[2] ;

   LExo:=QUELEXODTBud(TabDate[7]) ; DateTemp:=PlusMois(TabDate[7],1) ; if (LExo<>QUELEXODTBud(DateTemp)) then LExo:='' ; FinTemp:=FinDeMois(TabDate[7]) ;
   GetCumul('RUBREA',LeCompte,LeCompte,'SAN',CritEdt.Etab,CritEDT.DeviseAffichee,LExo,TabDate[7],FinTemp,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
   realise[6].TotDebit:=TD[3] ; realise[6].TotCredit:=TD[2] ;

   LExo:=QUELEXODTBud(TabDate[8]) ; DateTemp:=PlusMois(TabDate[8],1) ; if (LExo<>QUELEXODTBud(DateTemp)) then LExo:='' ; FinTemp:=FinDeMois(TabDate[8]) ;
   GetCumul('RUBREA',LeCompte,LeCompte,'SAN',CritEdt.Etab,CritEDT.DeviseAffichee,LExo,TabDate[8],FinTemp,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
   realise[7].TotDebit:=TD[3] ; realise[7].TotCredit:=TD[2] ;

   LExo:=QUELEXODTBud(TabDate[9]) ; DateTemp:=PlusMois(TabDate[9],1) ; if (LExo<>QUELEXODTBud(DateTemp)) then LExo:='' ; FinTemp:=FinDeMois(TabDate[9]) ;
   GetCumul('RUBREA',LeCompte,LeCompte,'SAN',CritEdt.Etab,CritEDT.DeviseAffichee,LExo,TabDate[9],FinTemp,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
   realise[8].TotDebit:=TD[3] ; realise[8].TotCredit:=TD[2] ;

   LExo:=QUELEXODTBud(TabDate[10]) ; DateTemp:=PlusMois(TabDate[10],1) ; if (LExo<>QUELEXODTBud(DateTemp)) then LExo:='' ; FinTemp:=FinDeMois(TabDate[10]) ;
   GetCumul('RUBREA',LeCompte,LeCompte,'SAN',CritEdt.Etab,CritEDT.DeviseAffichee,LExo,TabDate[10],FinTemp,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
   realise[9].TotDebit:=TD[3] ; realise[9].TotCredit:=TD[2] ;

   LExo:=QUELEXODTBud(TabDate[11]) ; DateTemp:=PlusMois(TabDate[11],1) ; if (LExo<>QUELEXODTBud(DateTemp)) then LExo:='' ; FinTemp:=FinDeMois(TabDate[11]) ;
   GetCumul('RUBREA',LeCompte,LeCompte,'SAN',CritEdt.Etab,CritEDT.DeviseAffichee,LExo,TabDate[11],FinTemp,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
   realise[10].TotDebit:=TD[3] ; realise[10].TotCredit:=TD[2] ;
   (*Prévoir boucle*)

   LExo:=QUELEXODTBud(TabDate[12]) ; DateTemp:=PlusMois(TabDate[12],1) ; if (LExo<>QUELEXODTBud(DateTemp)) then LExo:='' ; FinTemp:=FinDeMois(TabDate[12]) ;
   GetCumul('RUBREA',LeCompte,LeCompte,'SAN',CritEdt.Etab,CritEDT.DeviseAffichee,LExo,TabDate[12],FinTemp,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
   realise[11].TotDebit:=TD[3] ; realise[11].TotCredit:=TD[2] ;
(*
   LExo:=QUELEXODTBud(CritEdt.Date2) ; FinTemp:=FinDeMois(CritEdt.Date2) ;
   GetCumul('RUBREA',LeCompte,LeCompte,'SAN',CritEdt.Etab,CritEDT.DeviseAffichee,LExo,CritEdt.Date2,FinTemp,false,TRUE,Nil,TD) ;
//   GetCumul('RUBREA',LeCompte,LeCompte,'SAN',CritEdt.Etab,CritEDT.DeviseAffichee,LExo,TabDate[12],FinTemp,false,TRUE,Nil,TD) ;
   realise[11].TotDebit:=TD[1] ; realise[11].TotCredit:=TD[2] ;
*)
   For i:=0 to 11 do
       BEGIN
       Reevaluation(Realise[i].TotDebit,Realise[i].TotCredit,CritEdt.BalBud.Resol, CritEdt.BalBud.Taux) ;
       Realise[12].TotDebit:=Arrondi(Realise[12].TotDebit+Realise[i].TotDebit, CritEdt.Decimale) ;
       Realise[12].TotCredit:=Arrondi(Realise[12].TotCredit+Realise[i].TotCredit, CritEdt.Decimale) ;
       END ;

   { Calcul de l'écart }
   for i:=0 to 12 do
       BEGIN
       Ecar[i].TotDebit:=Arrondi(Realise[i].TotDebit-(TotGen[0][i].TotDebit+TotGen[1][i].TotDebit), CritEdt.Decimale) ;
       Ecar[i].TotCredit:=Arrondi(Realise[i].TotCredit-(TotGen[0][i].TotCredit+TotGen[1][i].TotCredit), CritEdt.Decimale) ;
//       Reevaluation(Ecar[i].TotDebit,Ecar[i].TotCredit,CritEdt.BalBud.Resol, CritEdt.BalBud.Taux) ;
       END ;
   { Calcul du réalisé ...}
   for i:=0 to 12 do
       BEGIN
       if (CritEdt.BalBud.RuptOnly<>Sur) then
          BEGIN {... En Total Edition : En ruptue il se fait sur le PrintGroup }
          TotEdt[1][i].TotDebit:= Arrondi(TotEdt[1][i].TotDebit+Realise[i].TotDebit, CritEdt.Decimale) ;
          TotEdt[1][i].TotCredit:=Arrondi(TotEdt[1][i].TotCredit+Realise[i].TotCredit, CritEdt.Decimale) ;
          END ;
       { ... en Nature libre }
       TotLibre[(i*2)+52]:=Arrondi(TotLibre[(i*2)+52]+Realise[i].TotDebit, CritEdt.Decimale) ;
       TotLibre[(i*2)+(1+52)]:=Arrondi(TotLibre[(i*2)+(1+52)]+Realise[i].TotCredit, CritEdt.Decimale) ;
       END ;
   END ;
For i:=0 to 11 do
    BEGIN
    C1:=FindComponent('IniGenCum'+InttoStr(i+1)) ;
    C2:=FindComponent('RevGenCum'+InttoStr(i+1)) ;
    C3:=FindComponent('ReaGenCum'+InttoStr(i+1)) ;
    C4:=FindComponent('EcaGenCum'+InttoStr(i+1)) ;
    TQRLabel(C1).Caption:=PrintSoldeFormate(TotGen[0][i].TotDebit, TotGen[0][i].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
    TQRLabel(C2).Caption:=PrintSoldeFormate(TotGen[1][i].TotDebit, TotGen[1][i].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
    TQRLabel(C3).Caption:=PrintSoldeFormate(realise[i].TotDebit, realise[i].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
//    TQRLabel(C4).Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Arrondi(Ecar[i].TotDebit-Ecar[i].TotCredit, CritEdt.Decimale), CritEDT.AfficheSymbole ) ;
    TQRLabel(C4).Caption:=PrintEcart(Ecar[i].TotDebit,Ecar[i].TotCredit,CritEdt.Decimale,DebitPos) ;
    END ;
IniGenTotal.Caption:=PrintSoldeFormate(TotGen[0][12].TotDebit, TotGen[0][12].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
RevGenTotal.Caption:=PrintSoldeFormate(TotGen[1][12].TotDebit, TotGen[1][12].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
ReaGenTotal.Caption:=PrintSoldeFormate(realise[12].TotDebit, realise[12].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
//EcaGenTotal.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Arrondi(Ecar[12].TotDebit-Ecar[12].TotCredit, CritEdt.Decimale) , CritEDT.AfficheSymbole ) ;
EcaGenTotal.Caption:=PrintEcart(Ecar[12].TotDebit,Ecar[12].TotCredit,CritEdt.Decimale,DebitPos) ;

AddReportBAL([1], CritEdt.Balbud.FormatPrint.Report, MReport, CritEdt.Decimale) ;
if CritEdt.Rupture=rLibres then
   BEGIN
   PrintBand:=(CritEdt.BalBud.RuptOnly<>Sur) ;
   AddGroupLibre(LLibre,Q,fbBudgen,CritEdt.LibreTrie,TotLibre) ;
   END ;

If PrintBand Then Quoi:=QuoiCpt(0) ;
end;

procedure TFQRBLteG.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var I          : Byte ;
    C1, C2, C3 : TComponent ;
begin
  inherited;
For i:=0 to 11 do
    BEGIN
    { En Budgété }
    C1:=FindComponent('BudTotal'+InttoStr(i+1)) ;
    { En Réalisé }
    C2:=FindComponent('ReaTotal'+InttoStr(i+1)) ;
    { Ecart : Réalisé - Budgété }
    C3:=FindComponent('EcaTotal'+InttoStr(i+1)) ;
    TQRLabel(C1).Caption:=PrintSoldeFormate(TotEdt[0][i].TotDebit, TotEdt[0][i].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
    TQRLabel(C2).Caption:=PrintSoldeFormate(TotEdt[1][i].TotDebit, TotEdt[1][i].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
//    TQRLabel(C3).Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, (TotEdt[1][i].TotDebit-TotEdt[0][i].TotDebit)-(TotEdt[1][i].TotCredit-TotEdt[0][i].TotCredit) , CritEDT.AfficheSymbole ) ;
    TQRLabel(C3).Caption:=PrintEcart(TotEdt[1][i].TotDebit-TotEdt[0][i].TotDebit,TotEdt[1][i].TotCredit-TotEdt[0][i].TotCredit,CritEDT.Decimale,DebitPos) ;
    END ;
TotTotalB.Caption:=PrintSoldeFormate(TotEdt[0][12].TotDebit, TotEdt[0][12].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
TotTotalR.Caption:=PrintSoldeFormate(TotEdt[1][12].TotDebit, TotEdt[1][12].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
//TotTotalE.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, (TotEdt[1][12].TotDebit-TotEdt[0][12].TotDebit)-(TotEdt[1][12].TotCredit-TotEdt[0][12].TotCredit) , CritEDT.AfficheSymbole ) ;
TotTotalE.Caption:=PrintEcart(TotEdt[1][12].TotDebit-TotEdt[0][12].TotDebit,TotEdt[1][12].TotCredit-TotEdt[0][12].TotCredit,CritEDT.Decimale,DebitPos) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFQRBLteG.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
Titre1Rep.Caption:=Titre2Rep.Caption ;
ReportCol1.Caption:=ReportCol13.Caption ;
ReportCol2.Caption:=ReportCol14.Caption ;
ReportCol3.Caption:=ReportCol15.Caption ;
ReportCol4.Caption:=ReportCol16.Caption ;
ReportCol5.Caption:=ReportCol17.Caption ;
ReportCol6.Caption:=ReportCol18.Caption ;
ReportCol7.Caption:=ReportCol19.Caption ;
ReportCol8.Caption:=ReportCol20.Caption ;
ReportCol9.Caption:=ReportCol21.Caption ;
ReportCol10.Caption:=ReportCol22.Caption ;
ReportCol11.Caption:=ReportCol23.Caption ;
ReportCol12.Caption:=ReportCol24.Caption ;
ReportTotal1.Caption:=ReportTotal2.Caption ;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
end;

procedure TFQRBLteG.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var MReport : TabTRep ;
begin
  inherited;
FillChar(Mreport,Sizeof(Mreport),#0) ;
Case QuelReportBAL(CritEdt.Balbud.FormatPrint.Report,MReport) of
  1 : Titre2Rep.Caption:=MsgBox.Mess[1] ;
 end ;
ReportCol13.Caption :=PrintSoldeFormate(MReport[1].TotDebit,MReport[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
ReportCol14.Caption :=PrintSoldeFormate(MReport[2].TotDebit,MReport[2].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
ReportCol15.Caption :=PrintSoldeFormate(MReport[3].TotDebit,MReport[3].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
ReportCol16.Caption :=PrintSoldeFormate(MReport[4].TotDebit,MReport[4].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
ReportCol17.Caption :=PrintSoldeFormate(MReport[5].TotDebit,MReport[5].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
ReportCol18.Caption :=PrintSoldeFormate(MReport[6].TotDebit,MReport[6].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
ReportCol19.Caption :=PrintSoldeFormate(MReport[7].TotDebit,MReport[7].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
ReportCol20.Caption :=PrintSoldeFormate(MReport[8].TotDebit,MReport[8].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
ReportCol21.Caption :=PrintSoldeFormate(MReport[9].TotDebit,MReport[9].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
ReportCol22.Caption :=PrintSoldeFormate(MReport[10].TotDebit,MReport[10].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
ReportCol23.Caption :=PrintSoldeFormate(MReport[11].TotDebit,MReport[11].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
ReportCol24.Caption :=PrintSoldeFormate(MReport[12].TotDebit,MReport[12].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
ReportTotal2.Caption:=PrintSoldeFormate(MReport[13].TotDebit,MReport[13].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
TITREREPORTB.Caption:='' ;
end;

procedure TFQRBLteG.BLibreBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.Rupture=rLibres) ;
end;

procedure TFQRBLteG.DLLibreNeedData(var MoreData: Boolean);
var TotLibre          : Array[0..77] of Double ;
    Librupt, Lib1, CodRupt, Stcode : String ;
    Quellerupt, i    : Integer ;
    Col              : TColor ;
    C1, C2, C3, C4   : TComponent ;
    CalIni, CalRev, CalRea : Double ;
    LibRuptInf : Array[1..10] Of TRuptInf ;
begin
  inherited;
MoreData:=False ;
if (CritEdt.Rupture=rLibres) then
   BEGIN
   MoreData:=PrintGroupLibre(LLibre,Q,fbBudGen,CritEdt.LibreTrie,CodRupt,LibRupt,Lib1,TotLibre,Quellerupt,Col,LibRuptInf) ;
   BLibre.Font.Color:=Col ; BLibreHaut.Font.Color:=Col ;
   If MoreData then
      BEGIN
      StCode:=CodRupt ;
      Delete(StCode,1,Quellerupt+2) ;
      MoreData:=DansChoixCodeLibre(StCode,Q,fbBudGen,CritEdt.LibreCodes1,CritEdt.LibreCodes2, CritEdt.LibreTrie) ;
      END ;
   if MoreData then
      BEGIN
      { Total Editon en Nature libre Uniquement}
      if (CritEdt.BalBud.RuptOnly=Sur) then
         BEGIN
         For i:=0 to 12 do
             BEGIN
             TotEdt[0][i].TotDebit:=Arrondi(TotEdt[0][i].TotDebit+(TotLibre[i*2]+TotLibre[(i*2)+26]), CritEdt.Decimale) ;
             TotEdt[0][i].TotCredit:=Arrondi(TotEdt[0][i].TotCredit+(TotLibre[i*2+1]+TotLibre[(i*2)+(1+26)]), CritEdt.Decimale) ;
             TotEdt[1][i].TotDebit:= Arrondi(TotEdt[1][i].TotDebit+TotLibre[(i*2)+52], CritEdt.Decimale) ;
             TotEdt[1][i].TotCredit:=Arrondi(TotEdt[1][i].TotCredit+TotLibre[(i*2)+(1+52)], CritEdt.Decimale) ;
             END ;
         END ;
      { Affichage sur la bande }
//      TCodRupt.Caption:=MsgBox.Mess[6]+' '+CodRupt+' '+Lib1 ;
      TLibRupt.Caption:='' ;

      insert(MsgBox.Mess[6]+' ',CodRupt,Quellerupt+2) ;
      TCodRupt.Caption:=CodRupt+' '+Lib1 ;


      for i:=0 to 11 do
          BEGIN
          C1:=FindComponent('IniLibre'+InttoStr(i+1)) ;
          C2:=FindComponent('RevLibre'+InttoStr(i+1)) ;
          C3:=FindComponent('ReaLibre'+InttoStr(i+1)) ;
          TQRLabel(C1).Caption:=PrintSoldeFormate(TotLibre[i*2], TotLibre[i*2+1], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
          TQRLabel(C2).Caption:=PrintSoldeFormate(TotLibre[(i*2)+26], TotLibre[(i*2)+(1+26)], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
          TQRLabel(C3).Caption:=PrintSoldeFormate(TotLibre[(i*2)+52], TotLibre[(i*2)+(1+52)], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
          END ;
      IniLibreTotal.Caption:=PrintSoldeFormate(TotLibre[24], TotLibre[25], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
      RevLibreTotal.Caption:=PrintSoldeFormate(TotLibre[50], TotLibre[51], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
      ReaLibreTotal.Caption:=PrintSoldeFormate(TotLibre[76], TotLibre[77], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
      { Calcul de l'écart }
      for i:=0 to 11 do
          BEGIN
          { (Ini+Rev)-Rea }
          CalIni:=Arrondi(TotLibre[i*2]-TotLibre[i*2+1],CritEdt.Decimale) ;
          CalRev:=Arrondi(TotLibre[(i*2)+26]-TotLibre[(i*2)+(1+26)],CritEdt.Decimale) ;
          CalRea:=Arrondi(TotLibre[(i*2)+52]-TotLibre[(i*2)+(1+52)],CritEdt.Decimale) ;
          C4:=FindComponent('EcaLibre'+InttoStr(i+1)) ;
//          TQRLabel(C4).Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, CalRea-(CalIni+CalRev) , CritEDT.AfficheSymbole ) ;
          TQRLabel(C4).Caption:=PrintEcart(CalRea,(CalIni+CalRev),CritEDT.Decimale,DebitPos ) ;
          END ;
      CalIni:=Arrondi(TotLibre[24]-TotLibre[25],CritEdt.Decimale) ;
      CalRev:=Arrondi(TotLibre[50]-TotLibre[51],CritEdt.Decimale) ;
      CalRea:=Arrondi(TotLibre[76]-TotLibre[77],CritEdt.Decimale) ;
//      EcaLibreTotal.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, CalRea-(CalIni+CalRev) , CritEDT.AfficheSymbole ) ;
      EcaLibreTotal.Caption:=PrintEcart(CalRea,(CalIni+CalRev),CritEDT.Decimale,DebitPos) ;
      END ;
   END ;
OkEnteteLibre:=MoreData ;
end;

procedure TFQRBLteG.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr1BG_BUDGENE :=TStringField(Q.FindField('BG_BUDGENE'));
Qr1BG_LIBELLE :=TStringField(Q.FindField('BG_LIBELLE'));
Qr1BG_RUB:=TStringField(Q.FindField('BG_RUB'));
end;

procedure TFQRBLteG.FJournalChange(Sender: TObject);
begin
  inherited;
if FJournal.Value='' then Exit ;
//YMO 07/10/2005 FQ 15151 Quand le change est provoqué par un nouveau filtre,
//pas de maj sur les zones
If Not ChargeFiltre then InfosJournal ;
end;

procedure TFQRBLteG.InfosJournal ;
Var QJal  : TQuery ;
BEGIN
QJal:=OpenSQL('Select BJ_EXODEB, BJ_EXOFIN, BJ_PERDEB, BJ_PERFIN, BJ_AXE, BJ_NATJAL from BUDJAL Where BJ_BUDJAL="'+FJournal.Value+'"',True) ;
if Not QJal.EOF then
   BEGIN
   FEXERCICE.Value:=QJal.FindField('BJ_EXODEB').AsString ; FEXERCICE_.Value:=QJal.FindField('BJ_EXOFIN').AsString ;
   ExoDebJal:=QJal.FindField('BJ_EXODEB').AsString ; ExoFinJal:=QJal.FindField('BJ_EXOFIN').AsString ;
   FPeriode1.Value:=QJal.FindField('BJ_PERDEB').AsString ; FPeriode2.Value:=QJal.FindField('BJ_PERFIN').AsString ;
   FDateCompta1.Text:=QJal.FindField('BJ_PERDEB').AsString ; FDateCompta2.Text:=QJal.FindField('BJ_PERFIN').AsString ;
   DatDebJal:=QJal.FindField('BJ_PERDEB').AsDateTime ; DatFinJal:=QJal.FindField('BJ_PERFIN').AsDateTime ;
   AxeJal:=QJal.FindField('BJ_AXE').AsString ;
   DebitPos:=QJal.FindField('BJ_NATJAL').AsString='CHA' ;
   END ;
Ferme(QJal) ;
END ;

procedure TFQRBLteG.FExerciceChange(Sender: TObject);
begin
//  inherited;
if FExercice.Value>FExercice_.Value then FExercice_.Value:=FExercice.Value ;
ListePeriode(FExercice.Value,FPeriode1.Items,FPeriode1.Values,True) ;
FPeriode1.ItemIndex:=0 ; DebToDate ;
end;

procedure TFQRBLteG.FExercice_Change(Sender: TObject);
begin
  inherited;
if FExercice_.Value<FExercice.Value then FExercice.Value:=FExercice_.Value ;
ListePeriode(FExercice_.Value,FPeriode2.Items,FPeriode2.Values,False) ;
FPeriode2.ItemIndex:=FPeriode2.Items.Count-1 ; FinToDate ;
end;

procedure TFQRBLteG.FPeriode1Change(Sender: TObject);
begin
  inherited;
//YMO 06/10/2005 FQ 15151 Quand le change est provoqué par un nouveau filtre,
//les dates ne sont pas encore renseignées quand cet évènement est déclenché
if isvalidDate(FPeriode1.Value)
and isvalidDate(FPeriode2.Value)
and (StrToDate(FPeriode1.Value)>StrToDate(FPeriode2.Value))
then FPeriode2.Value:=DateToStr(FinDeMois(StrToDate(FPeriode1.Value))) ;

FDateCompta1.Text:=FPeriode1.Value ;
end;

procedure TFQRBLteG.FPeriode2Change(Sender: TObject);
begin
  inherited;
//YMO 06/10/2005 FQ 15151 Quand le change est provoqué par un nouveau filtre,
//les dates ne sont pas encore renseignées quand cet évènement est déclenché
if isvalidDate(FPeriode1.Value)
and isvalidDate(FPeriode2.Value)
and (StrToDate(FPeriode2.Value)<StrToDate(FPeriode1.Value))
then FPeriode1.Value:=DateToStr(DebutDeMois(StrToDate(FPeriode2.Value))) ;

FDateCompta2.Text:=FPeriode2.Value ;
end;

procedure TFQRBLteG.DebToDate ;
BEGIN
FDateCompta1.Text:=FPeriode1.Value ;
END ;

procedure TFQRBLteG.FinToDate ;
BEGIN
FDateCompta2.Text:=FPeriode2.Value ;
END ;

Procedure PrepCalcBud(CritEdt : TCritEdt ; Var QMvt : TQuery)  ;
BEGIN
QMvt:=TQuery.Create(Application) ; QMvt.DataBaseName:='SOC' ; QMvt.SQL.Clear ;
QMvt.SQL.Clear ;
QMvt.SQL.Add('Select BE_BUDGENE, BE_BUDSECT, BE_EXERCICE, BE_DATECOMPTABLE, BE_NUMEROPIECE,') ;
QMvt.SQL.Add(       'BE_REFINTERNE, BE_ETABLISSEMENT, BE_LIBELLE, BE_VALIDE,') ;
QMvt.SQL.Add(       'BE_BUDJAL,BE_NATUREBUD, BE_QUALIFPIECE, ') ;
Case CritEdt.Monnaie of
  0 : BEGIN QMvt.SQL.Add('BE_DEBIT DEBIT,BE_CREDIT CREDIT') ; END ;
//  1 : BEGIN QMvt.SQL.Add('E_DEBITDEV as DEBIT,E_CREDITDEV as CREDIT') ; END ;
//  2 : BEGIN QMvt.SQL.Add('BE_DEBITEURO DEBIT,BE_CREDITEURO CREDIT') ; END ;
end ;
QMvt.SQL.Add(' From BUDECR ') ;
QMvt.SQL.Add(' Where BE_BUDGENE=:C1 ') ;
//QMvt.SQL.Add(' And BE_AXE="'+CritEdt.BalBud.AXE+'" ') ;
QMvt.SQL.Add(' And BE_BUDJAL="'+CritEdt.BalBud.Journal+'" ') ;

  // FQ 16536 SBO 18/10/2005 Mise en place des options de révision dans les balances budgétaires
  If CritEdt.BalBud.NatureBud<>'' then
    QMvt.SQL.Add('AND ' + CritEdt.BalBud.NatureBud + ' ') ;
  // fin FQ 16536 SBO 18/10/2005 Mise en place des options de révision dans les balances budgétaires

QMvt.SQL.Add('And BE_DATECOMPTABLE>="'+usdatetime(CritEdt.Date1)+'" And BE_DATECOMPTABLE<="'+usdatetime(CritEdt.Date2)+'" ') ;
QMvt.SQL.Add('And BE_EXERCICE>="'+CritEdt.BalBud.Exo1+'" And BE_EXERCICE<="'+CritEdt.BalBud.Exo2+'" ') ;
if CritEdt.Etab<>'' then QMvt.SQL.Add(' And BE_ETABLISSEMENT="'+CritEdt.Etab+'"') ;
if CritEdt.Valide<>'g' then QMvt.SQL.Add(' And BE_VALIDE="'+CritEdt.Valide+'" ') ;
If CritEDT.SQLPlus<>'' Then QMvt.SQL.Add(CritEDT.SQLPlus) ;
QMvt.SQL.Add(' Order By BE_BUDGENE,BE_BUDJAL,BE_EXERCICE,BE_DATECOMPTABLE,BE_BUDSECT,BE_AXE,BE_NUMEROPIECE,BE_QUALIFPIECE') ;
ChangeSql(QMvt) ;
END ;


Procedure InitCalcBud(C1 : String ; Var Tot1 : TMontTot ; TabD : TTabDate12 ; CritEdt : TCritEdt ; Q : TQuery) ;
Var DateCompta : TDateTime ;
    LaNatBud : String ;
    DEBIT, CREDIT : Double ;
BEGIN
Q.Close ;
Q.PAramByName('C1').AsString:=C1 ;
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
      Reevaluation(DEBIT,CREDIT,CritEdt.BalBud.Resol, CritEdt.BalBud.Taux) ;
      if (DateCOMPTA>=TabD[1]) And (DateCompta<=FinDeMois(TabD[1]))
         then CalculBud(0,Tot1, LaNatBud,DEBIT, CREDIT, CritEdt.Decimale) ;
      if (DateCompta>=TabD[2]) And (DateCompta<=FinDeMois(TabD[2]))
         then CalculBud(1,Tot1, LaNatBud,DEBIT, CREDIT, CritEdt.Decimale) ;
      if (DateCompta>=TabD[3]) And (DateCompta<=FinDeMois(TabD[3]))
         then CalculBud(2,Tot1, LaNatBud,DEBIT, CREDIT, CritEdt.Decimale) ;
      if (DateCompta>=TabD[4]) And (DateCompta<=FinDeMois(TabD[4]))
         then CalculBud(3,Tot1, LaNatBud,DEBIT, CREDIT, CritEdt.Decimale) ;
      if (DateCompta>=TabD[5]) And (DateCompta<=FinDeMois(TabD[5]))
         then CalculBud(4,Tot1, LaNatBud,DEBIT, CREDIT, CritEdt.Decimale) ;
      if (DateCompta>=TabD[6]) And (DateCompta<=FinDeMois(TabD[6]))
         then CalculBud(5,Tot1, LaNatBud,DEBIT, CREDIT, CritEdt.Decimale) ;
      if (DateCompta>=TabD[7]) And (DateCompta<=FinDeMois(TabD[7]))
         then CalculBud(6,Tot1, LaNatBud,DEBIT, CREDIT, CritEdt.Decimale) ;
      if (DateCompta>=TabD[8]) And (DateCompta<=FinDeMois(TabD[8]))
         then CalculBud(7,Tot1, LaNatBud,DEBIT, CREDIT, CritEdt.Decimale) ;
      if (DateCompta>=TabD[9]) And (DateCompta<=FinDeMois(TabD[9]))
         then CalculBud(8,Tot1, LaNatBud,DEBIT, CREDIT, CritEdt.Decimale) ;
      if (DateCompta>=TabD[10]) And (DateCompta<=FinDeMois(TabD[10]))
         then CalculBud(9,Tot1, LaNatBud,DEBIT, CREDIT, CritEdt.Decimale) ;
      if (DateCompta>=TabD[11]) And (DateCompta<=FinDeMois(TabD[11]))
         then CalculBud(10,Tot1, LaNatBud,DEBIT, CREDIT, CritEdt.Decimale) ;
      if (DateCompta>=TabD[12]) And (DateCompta<=FinDeMois(TabD[12]))
         then CalculBud(11,Tot1, LaNatBud,DEBIT, CREDIT, CritEdt.Decimale) ;
      Q.Next ;
      END ;
END ;

Procedure CalculBud(i : Byte ; Var Tot1 : TMontTot ; NatB : String ; D,C : Double ; Dec : Byte) ;
Var NB : String ;
BEGIN
NB:=NatB ;
If (NB='INI') or (NB='ANO') then         {FP FQ16059}
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

procedure TFQRBLteG.BHCptBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
If CritEdt.Rupture=rLibres then
   if CritEdt.BalBud.OnlyCptAssocie then
      BEGIN
      PrintBand:=DansRuptLibre(Q,fbBudgen,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
      END ;
Affiche:=PrintBand ;
{ La valeur de Affiche doit etre affecté avant le test en sur rupture,
  sinon il n'y a pas de calcul sur 'BEFORE PRINT' de BDETAIL }
if CritEdt.BalBud.RuptOnly=Sur then PrintBand:=False ;//
If PrintBand then
   BEGIN
   BG_BUDGENE.Caption:=MsgBox.Mess[0]+'  '+Qr1BG_BUDGENE.AsString ;
   TFBG_LIBELLE.Caption:=Qr1BG_LIBELLE.AsString ;
   Quoi:=QuoiCpt(0) ;
   END ;
end;

procedure TFQRBLteG.DLHLibreNeedData(var MoreData: Boolean);
begin
  inherited;
MoreData:=OkEnteteLibre ;
OkEnteteLibre:=False ;
end;

procedure TFQRBLteG.FSansRuptClick(Sender: TObject);
begin
  inherited;
if QuelEtatBud=Nature then
   BEGIN
   if Not FSansRupt.Checked then FTablesLibres.Checked:=True ;
   FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
   FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
   END ;
end;

procedure TFQRBLteG.BNouvRechClick(Sender: TObject);
begin
  inherited;
FJournal.ItemIndex:=0 ;
//FRESOL.ItemIndex:=0 ;
InitResolution(FRESOL) ;
end;


procedure TFQRBLteG.FFiltresChange(Sender: TObject);
begin
ChargeFiltre:=True;
  inherited;
ChargeFiltre:=False;
end;

end.
