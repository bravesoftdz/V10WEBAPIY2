{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 31/03/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par BALBUDECGEN_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit QRBLecG;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Hctrls, Spin, Menus, hmsgbox, HQuickrp, StdCtrls,
  Buttons, ExtCtrls, Mask, Hcompte, ComCtrls, UtilEdt, HEnt1, Ent1,
  CpteUtil, EdtLegal, CritEDT, HQry, HSysMenu,UtilEdt1, QRRupt, Calcole, SAISUTIL,
  HTB97, HPanel, UiUtil ;

procedure BalBudecGen ;
procedure BalBudecGenNature ;
procedure BalBudecGenZoom(Crit : TCritEdt) ;

type
  TFQRBLecG = class(TFQR)
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
    BLibre: TQRBand;
    DLLibre: TQRDetailLink;
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
    REPORTCOL7: TQRLabel;
    REPORTCOL8: TQRLabel;
    REPORTCOL9: TQRLabel;
    REPORTCOL10: TQRLabel;
    TotCum7: TQRLabel;
    TotCum8: TQRLabel;
    TotCum9: TQRLabel;
    TotCum10: TQRLabel;
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
    FQte: TCheckBox;
    QRLabel1: TQRLabel;
    RQte: TQRLabel;
    IniGenCum5: TQRLabel;
    IniGenCum4: TQRLabel;
    IniGenCum2: TQRLabel;
    IniGenCum3: TQRLabel;
    IniGenCum6: TQRLabel;
    IniGenCum1: TQRLabel;
    IniGenCum7: TQRLabel;
    IniGen: TQRLabel;
    RevGen: TQRLabel;
    RevGenCum1: TQRLabel;
    RevGenCum2: TQRLabel;
    RevGenCum3: TQRLabel;
    RevGenCum4: TQRLabel;
    RevGenCum5: TQRLabel;
    RevGenCum6: TQRLabel;
    RevGenCum7: TQRLabel;
    QIniGen: TQRLabel;
    QIniGenCum1: TQRLabel;
    QIniGenCum2: TQRLabel;
    QIniGenCum3: TQRLabel;
    QIniGenCum4: TQRLabel;
    QIniGenCum5: TQRLabel;
    QIniGenCum6: TQRLabel;
    QIniGenCum7: TQRLabel;
    QRevGen: TQRLabel;
    QRevGenCum1: TQRLabel;
    QRevGenCum2: TQRLabel;
    QRevGenCum3: TQRLabel;
    QRevGenCum4: TQRLabel;
    QRevGenCum5: TQRLabel;
    QRevGenCum6: TQRLabel;
    QRevGenCum7: TQRLabel;
    IniGenCum9: TQRLabel;
    RevGenCum9: TQRLabel;
    QIniGenCum9: TQRLabel;
    QRevGenCum9: TQRLabel;
    IniGenCum10: TQRLabel;
    RevGenCum10: TQRLabel;
    QIniGenCum10: TQRLabel;
    QRevGenCum10: TQRLabel;
    IniGenCum8: TQRLabel;
    RevGenCum8: TQRLabel;
    QIniGenCum8: TQRLabel;
    QRevGenCum8: TQRLabel;
    BHCpt: TQRBand;
    BG_BUDGENE: TQRDBText;
    BG_LIBELLE: TQRLabel;
    HGene: TQRGroup;
    FTotIniRev: TCheckBox;
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
    BLibreHaut: TQRBand;
    TCodRupt: TQRDBText;
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
    TFSoldeFormate: THLabel;
    FSoldeFormate: THValComboBox;
    FTaux: THNumEdit;
    TRResol: TQRLabel;
    RResol: TQRLabel;
    FQueTotBud: TCheckBox;
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
    procedure FNatBudChange(Sender: TObject);
    procedure DLHLibreNeedData(var MoreData: Boolean);
    procedure FSansRuptClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure FQueTotBudClick(Sender: TObject);
    procedure FTotIniRevClick(Sender: TObject);
    procedure GenereSQL ; Override ;
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
  private    { Déclarations privées }
    QBUD                      : TQuery ;
    DebitPos                  : Boolean ;
    ExoDateDeb, ExoDateFin    : TDateTime ;
    LLibre                    : TStringList ;
    TotEdt, TotIR             : TabTot12 ; { Total de l'édition, total Initial +Révisé }
    TotGen                    : TMontTot ; { Total du compte en cours de traitement }
    TabDate                   : TTabDate12 ;
    Qr1BG_BUDGENE, Qr1BG_RUB, Qr1BG_LIBELLE : TStringField ;
    OkEnteteLibre                : Boolean ;
    QuelEtatBud                  : TQuelEtatBud ;
    Affiche                      : Boolean ;
    AxeJal                       : String3 ;
    Function  QuoiCpt(i : Integer) : String ;
    Procedure BalBudGenZoom(Quoi : String) ;
    procedure InfosJournal ;
    procedure DebToDate ;
    procedure FinToDate ;
    procedure MontantIR ;
//    procedure StMontantsBudEcart(M : TabTot12 ; var AffM : array of String) ;
  public    { Déclarations publiques }
  end;

Procedure PrepCalcBud(CritEdt : TCritEdt ; Var QMvt : TQuery)  ;
Procedure InitCalcBud(C1 : String ; Var Tot1 : TMontTot ; TabD : TTabDate12 ; CritEdt : TCritEdt ; Q : TQuery) ;
Procedure CalculBud(i : Byte ; Var Tot1 : TMontTot ; NatB : String ; D,C : Double ; Dec : Byte) ;

implementation

{$R *.DFM}

Function TFQRBLecG.QuoiCpt(i : Integer) : String ;
BEGIN
Case i of
   0 : Result:=Q.FindField('BG_BUDGENE').AsString+' '+Q.FindField('BG_LIBELLE').AsString ;
  End ;
END ;

Procedure TFQRBLecG.BalBudGenZoom(Quoi : String) ;
BEGIN
ZoomEdt(8,Quoi) ;
END ;

procedure BalBudecGen ;
var QR : TFQRBLecG ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQRBLecG.Create(Application) ;
Edition.Etat:=etBlBudEcartGen ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalBudGenZoom ;
QR.QuelEtatBud:=Normal ;
QR.InitType (nbBudGen,neBalBud,msGenBEcr,'QRBLecGNor','',TRUE,FALSE,FALSE,Edition) ;
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

procedure BalBudecGenNature ;
var QR : TFQRBLecG ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQRBLecG.Create(Application) ;
Edition.Etat:=etBlBudEcartGen ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalBudGenZoom ;
QR.QuelEtatBud:=Nature ;
QR.InitType (nbBudGen,neBalBud,msGenBEcr,'QRBLecGNat','',TRUE,FALSE,FALSE,Edition) ;
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

procedure BalBudecGenZoom(Crit : TCritEdt) ;
var QR : TFQRBLecG ;
    Edition : TEdition ;
BEGIN
QR:=TFQRBLecG.Create(Application) ;
Edition.Etat:=etBlBudEcartGen ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.BalBudGenZoom ;
 QR.QuelEtatBud:=Normal ;
 QR.InitType (nbBudGen,neGlBud,msGenBEcr,'QRBLecGNor','',FALSE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFQRBLecG.FormShow(Sender: TObject);
begin
HelpContext:=15331000 ;
Standards.HelpContext:=15331100 ;
Avances.HelpContext:=15331200 ;
Mise.HelpContext:=15331300 ;
Option.HelpContext:=15331400 ;
TabRuptures.HelpContext:=15331500 ;
TabSup.TabVisible:=False ;
If FJournal.Values.Count>0 Then FJournal.Value:=FJournal.Values[0] ;
FNatBud.ItemIndex:=0 ; FSoldeFormate.ItemIndex:=0 ;
  inherited;
InitResolution(FRESOL) ;
//FRESOL.Value:='C' ;
FLigneCpt.Checked:=True ;
TFGen.Parent:=Avances ; TFGenJoker.Parent:=Avances ;
FJoker.Parent:=Avances ; TFaG.Parent:=Avances ;
FCpte1.Parent:=Avances ; FCpte2.Parent:=Avances ;
FCpte1.TabOrder:=0 ; FCpte2.TabOrder:=1 ; FJoker.TabOrder:=2 ;

DebToDate ; FinToDate ;
FNatBudChange(Nil) ;
FOnlyCptAssocie.Enabled:=False ;
//If (Not EstSerie(S7)) Then FGroupChoixRupt.Enabled:=FALSE ;
If (EstSerie(S3)) Then FGroupChoixRupt.Visible:=FALSE ;
end;

Procedure TFQRBLecG.FinirPrint ;
BEGIN
Inherited ;
if (CritEdt.Rupture=rLibres) then VideGroup(LLibre) ;
if CritEdt.Balbud.FormatPrint.PrSepMontant then FreeLesLignes(Self) ;
END ;

procedure TFQRBLecG.InitDivers ;
BEGIN
Inherited ;
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
{ Calcul des différentes fourchettes de dates }
CritEdt.BalBud.FormatPrint.ColMin:=1 ;
CritEdt.BalBud.FormatPrint.ColMax:=10 ;
//BDetail.Frame.DrawBottom:=CritEdt.BalBud.FormatPrint.PrSepCompte[1] ;
END ;

procedure TFQRBLecG.GenereSQL ;
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
If (CritEdt.Rupture=rLibres) then
   BEGIN
   St:=WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,fbBudgen,CritEdt.BalBud.OnlyCptAssocie) ;
   If St<>'' Then Q.SQL.Add(St) ;
   Q.SQL.Add('Order By '+OrderLibre(CritEdt.LibreTrie)+'BG_BUDGENE ') ;
   END Else Q.SQL.Add(' Order By BG_BUDGENE') ;
ChangeSQL(Q) ; Q.Open ;
END ;

procedure TFQRBLecG.RenseigneCritere ;
BEGIN
Inherited ;
RJournal.Caption:=FJournal.Text ;
RExercice_.Caption:=FExercice_.Text ;
CaseACocher(FValide,RValide) ;
CaseACocher(FQte,RQte) ;
RNatBud.Caption:=FNatBud.Text ;
RTaux.Caption:=FloatToStr(FTaux.Value) ;
RResol.Caption:=FResol.Text ;
AfficheBudgetEn(RDevises,CritEdt) ;
END ;

procedure TFQRBLecG.ChoixEdition ;
{ Initialisation des options d'édition }
BEGIN
Inherited ;
if (CritEdt.Rupture=rLibres) then
   BEGIN
   ChargeGroup(LLibre,['B00','B01','B02','B03','B04','B05','B06','B07','B08','B09']) ;
   END ;

RetailleBandesBudEcarts(Self,CritEdt,Msgbox.Mess[3],BDetail,BLibre,Nil,BFinEtat) ;

If CritEdt.BalBud.Resol='C' then CritEdt.Decimale:=V_PGI.OkDecV else
If CritEdt.BalBud.Resol='F' then CritEdt.Decimale:=0 else
If CritEdt.BalBud.Resol='K' then CritEdt.Decimale:=0 else
If CritEdt.BalBud.Resol='M' then CritEdt.Decimale:=0 ;

ChangeMask(Formateur,CritEdt.Decimale,CritEdt.Symbole);
CritEdt.FormatMontant:=Formateur.Masks.PositiveMask;

if CritEdt.Balbud.FormatPrint.PrSepMontant then AffecteLigne(Self,EntetePage,CritEdt.BalBud.FormatPrint) ;
END ;

procedure TFQRBLecG.RecupCritEdt ;
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
  BalBud.Resol:=FResol.Value[1] ; BalBud.Taux:=FTaux.Value ;
  BalBud.Qte:=FQte.Checked ; BalBud.TotIniRev:=FTotIniRev.Checked and FTotIniRev.Enabled ;
  BalBud.RuptOnly:=QuelleTypeRupt(0,FSAnsRupt.Checked,FAvecRupt.Checked,FSurRupt.Checked) ;
  BalBud.OnlyCptAssocie:=(Rupture<>rRien) and FOnlyCptAssocie.Checked ;
//  BalBud.SoldeFormate:=FSoldeFormate.Value ;
  If DebitPos Then BalBud.SoldeFormate:='PD' Else BalBud.SoldeFormate:='PC' ;
  BalBud.QueTotalBud:=FQueTotBud.Checked ;
  With BalBud.FormatPrint Do
     BEGIN
     END ;
  END ;
END ;

function TFQRBLecG.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ;
if Result then PrepCalcBud(CritEdt,QBUD)  ;
END ;

procedure TFQRBLecG.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var RealPer, EcarPer, RealCum, EcarCum : TabTot12 ;
    I                : Byte ; {}
    C1, C2           : TComponent ;
    TD               : TabloExt ;
    LeCompte, Lexo   : String ;
    Pourcentage      : double ;
    MReport          : TabTRep ;
    FinTemp          : TDateTime ;
    TotBud           : TMontTot ;
    TotLibre         : TabMont77 ;
begin
  inherited;
Fillchar(TotBud,SizeOf(TotBud),#0) ;     { Init le total Budgété de chaque compte }
Fillchar(TotGen,SizeOf(TotGen),#0) ;     { Init le total du compte budgetaire }
Fillchar(MReport,SizeOf(MReport),#0) ;   { Init le report }
Fillchar(TotLibre,SizeOf(TotLibre),#0) ; { Init les montants libres }
Fillchar(TotIR,SizeOf(TotIR),#0) ;       { Init le total de Initial +Réalisé }
(**)
PrintBand:=Affiche ;
If Not PrintBand then Exit ;
InitCalcBud(Qr1Bg_BUDGENE.AsString,TotBud,TabDate,CritEdt,QBud) ; { Calcul les montants du budgété du compte budgétaire }
{ Affectation du calcul des montants budgétes pour le total du Compte budgétaire aux cases suivantes du tableau... }
{ 0 : Budgété de la période, 4 : Budgété Cumulé, 8 : Budgété annuel }
TotGen:=TotBud ;
{ Idem pour... }
{...Total Initial + Révisé}
AlimTotEdtEcartBudget(TotIR,TotBud,CritEdt) ;
(*
TotIR[0].TotDebit:= Arrondi(TotIR[0].TotDebit+(TotBud[0][0].totDebit+TotBud[1][0].totDebit), CritEdt.Decimale) ;
TotIR[0].TotCredit:= Arrondi(TotIR[0].TotCredit+(TotBud[0][0].totCredit+TotBud[1][0].totCredit), CritEdt.Decimale) ;
TotIR[4].TotDebit:= Arrondi(TotIR[4].TotDebit+(TotBud[0][1].totDebit+TotBud[1][1].totDebit), CritEdt.Decimale) ;
TotIR[4].TotCredit:= Arrondi(TotIR[4].TotCredit+(TotBud[0][1].totCredit+TotBud[1][1].totCredit), CritEdt.Decimale) ;
TotIR[8].TotDebit:= Arrondi(TotIR[8].TotDebit+(TotBud[0][2].totDebit+TotBud[1][2].totDebit), CritEdt.Decimale) ;
TotIR[8].TotCredit:= Arrondi(TotIR[8].TotCredit+(TotBud[0][2].totCredit+TotBud[1][2].totCredit), CritEdt.Decimale) ;
*)
{...Total Edition}
if (CritEdt.BalBud.RuptOnly<>Sur) then
   BEGIN
   {Pour Total Edition : En rupture le total est fait sur le PrintGroup}
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
{...Report}
(*
MReport[0+1].TotDebit:= Arrondi(MReport[0+1].TotDebit+(TotBud[0][0].totDebit+TotBud[1][0].totDebit), CritEdt.Decimale) ;
MReport[0+1].TotCredit:= Arrondi(MReport[0+1].TotCredit+(TotBud[0][0].totCredit+TotBud[1][0].totCredit), CritEdt.Decimale) ;
MReport[4+1].TotDebit:= Arrondi(MReport[4+1].TotDebit+(TotBud[0][1].totDebit+TotBud[1][1].totDebit), CritEdt.Decimale) ;
MReport[4+1].TotCredit:= Arrondi(MReport[4+1].TotCredit+(TotBud[0][1].totCredit+TotBud[1][1].totCredit), CritEdt.Decimale) ;
MReport[8+1].TotDebit:= Arrondi(MReport[8+1].TotDebit+(TotBud[0][2].totDebit+TotBud[1][2].totDebit), CritEdt.Decimale) ;
MReport[8+1].TotCredit:= Arrondi(MReport[8+1].TotCredit+(TotBud[0][2].totCredit+TotBud[1][2].totCredit), CritEdt.Decimale) ;
*)
{...Rupture : Nature libre }
{ Revise }
TotLibre[0]:= Arrondi(TotLibre[0]+TotBud[0][0].totDebit, CritEdt.Decimale) ;
TotLibre[1]:= Arrondi(TotLibre[1]+TotBud[0][0].totCredit, CritEdt.Decimale) ;
TotLibre[8]:= Arrondi(TotLibre[8]+TotBud[0][1].totDebit, CritEdt.Decimale) ;
TotLibre[9]:= Arrondi(TotLibre[9]+TotBud[0][1].totCredit, CritEdt.Decimale) ;
TotLibre[16]:= Arrondi(TotLibre[16]+TotBud[0][2].totDebit, CritEdt.Decimale) ;
TotLibre[17]:= Arrondi(TotLibre[17]+TotBud[0][2].totCredit, CritEdt.Decimale) ;
{ Revise }
TotLibre[20]:= Arrondi(TotLibre[20]+TotBud[1][0].totDebit, CritEdt.Decimale) ;
TotLibre[21]:= Arrondi(TotLibre[21]+TotBud[1][0].totCredit, CritEdt.Decimale) ;
TotLibre[28]:= Arrondi(TotLibre[28]+TotBud[1][1].totDebit, CritEdt.Decimale) ;
TotLibre[29]:= Arrondi(TotLibre[29]+TotBud[1][1].totCredit, CritEdt.Decimale) ;
TotLibre[36]:= Arrondi(TotLibre[36]+TotBud[1][2].totDebit, CritEdt.Decimale) ;
TotLibre[37]:= Arrondi(TotLibre[37]+TotBud[1][2].totCredit, CritEdt.Decimale) ;

{ Calcul du Réalisé pour... }
{ ...La Période }
     { Mvt init }
LeCompte:='CBG'+CritEdt.BalBud.Journal+Qr1BG_RUB.AsString ;

LExo:=QUELEXODTBud(CritEdt.Date1) ; if (Lexo<>QUELEXODTBud(CritEdt.Date2)) then LExo:='' ; FinTemp:=FinDeMois(CritEdt.Date2) ;
GetCumul('RUBREA',LeCompte,LeCompte,CritEdt.BalBud.SANbud,CritEdt.Etab,CritEDT.DeviseAffichee,LExo,CritEdt.Date1,FinTemp,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
RealPer[0].TotDebit:=TD[3] ; RealPer[0].TotCredit:=TD[2] ;
EcarPer[0].TotDebit:= Arrondi((RealPer[0].TotDebit-TotGen[0][0].TotDebit), CritEdt.Decimale) ;
EcarPer[0].TotCredit:=Arrondi((RealPer[0].TotCredit-TotGen[0][0].TotCredit), CritEdt.Decimale) ;

Reevaluation(RealPer[0].TotDebit,RealPer[0].TotCredit,CritEdt.BalBud.Resol, CritEdt.BalBud.Taux) ;
Reevaluation(EcarPer[0].TotDebit,EcarPer[0].TotCredit,CritEdt.BalBud.Resol, CritEdt.BalBud.Taux) ;

IniGenCum1.Caption:=PrintSoldeFormate(TotGen[0][0].TotDebit, TotGen[0][0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
IniGenCum2.Caption:=PrintSoldeFormate(RealPer[0].TotDebit, RealPer[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
//IniGenCum3.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Abs(EcarPer[0].TotDebit-EcarPer[0].TotCredit) , CritEDT.AfficheSymbole ) ;
IniGenCum3.Caption:=PrintEcart(EcarPer[0].TotDebit,EcarPer[0].TotCredit,CritEDT.Decimale,DebitPos) ;
if (IniGenCum1.Caption<>'')and(IniGenCum3.Caption<>'') then
   BEGIN
   Pourcentage:=((EcarPer[0].TotDebit-EcarPer[0].TotCredit)*100)/(TotGen[0][0].TotDebit-TotGen[0][0].TotCredit) ;
   Pourcentage:=Abs(Pourcentage) ;
   If Pos('-',IniGenCum3.Caption)>0 Then Pourcentage:=-1*Pourcentage ;
   END Else Pourcentage:=0 ;
IniGenCum4.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,Pourcentage , CritEDT.AfficheSymbole ) ;
     { Mvt Revi }
RealPer[1].TotDebit:=RealPer[0].TotDebit ;
RealPer[1].TotCredit:=RealPer[0].TotCredit ;
EcarPer[1].TotDebit:= Arrondi((RealPer[1].TotDebit-TotGen[1][0].TotDebit), CritEdt.Decimale) ;
EcarPer[1].TotCredit:=Arrondi((RealPer[1].TotCredit-TotGen[1][0].TotCredit), CritEdt.Decimale) ;

RevGenCum1.Caption:=PrintSoldeFormate(TotGen[1][0].TotDebit, TotGen[1][0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
RevGenCum2.Caption:=PrintSoldeFormate(RealPer[1].TotDebit, RealPer[1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
//RevGenCum3.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Abs(EcarPer[1].TotDebit-EcarPer[1].TotCredit), CritEDT.AfficheSymbole ) ;
RevGenCum3.Caption:=PrintEcart(EcarPer[1].TotDebit,EcarPer[1].TotCredit,CritEDT.Decimale,DebitPos) ;
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

{ ...Le Cumulé }
     { init }
(*
LExo:=QUELEXODTBud(CritEdt.Date1) ; FinTemp:=FinDeMois(CritEdt.DateFin) ;
GetCumul('RUBREA',LeCompte,LeCompte,CritEdt.BalBud.SANbud,CritEdt.Etab,CritEDT.DeviseAffichee,LExo,CritEdt.DateDeb,FinTemp,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
*)
LExo:=QUELEXODTBud(CritEdt.DateDeb) ; if (Lexo<>QUELEXODTBud(CritEdt.Date2)) then LExo:='' ; FinTemp:=FinDeMois(CritEdt.Date2) ;
GetCumul('RUBREA',LeCompte,LeCompte,CritEdt.BalBud.SANbud,CritEdt.Etab,CritEDT.DeviseAffichee,LExo,CritEdt.DateDeb,FinTemp,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
RealCum[0].TotDebit:=TD[3] ; RealCum[0].TotCredit:=TD[2] ;
EcarCum[0].TotDebit:= Arrondi((RealCum[0].TotDebit-TotGen[0][1].TotDebit), CritEdt.Decimale) ;
EcarCum[0].TotCredit:=Arrondi((RealCum[0].TotCredit-TotGen[0][1].TotCredit), CritEdt.Decimale) ;

Reevaluation(RealCum[0].TotDebit,RealCum[0].TotCredit,CritEdt.BalBud.Resol, CritEdt.BalBud.Taux) ;
Reevaluation(EcarCum[0].TotDebit,EcarCum[0].TotCredit,CritEdt.BalBud.Resol, CritEdt.BalBud.Taux) ;

IniGenCum5.Caption:=PrintSoldeFormate(TotGen[0][1].TotDebit, TotGen[0][1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
IniGenCum6.Caption:=PrintSoldeFormate(RealCum[0].TotDebit, RealCum[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
//IniGenCum7.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Abs(EcarCum[0].TotDebit-EcarCum[0].TotCredit), CritEDT.AfficheSymbole ) ;
IniGenCum7.Caption:=PrintEcart(EcarCum[0].TotDebit,EcarCum[0].TotCredit,CritEDT.Decimale,DebitPos ) ;
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
{ ... L' Annuel }
     { Init }
IniGenCum9.Caption:=PrintSoldeFormate(TotGen[0][2].TotDebit, TotGen[0][2].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
     { Revi }
RevGenCum9.Caption:=PrintSoldeFormate(TotGen[1][2].TotDebit, TotGen[1][2].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;

{--------------------------------------------}
{ Calcule du réalisé en Total Ini+Rev }
TotIR[1].TotDebit:= Arrondi(TotIR[1].TotDebit+RealPer[0].TotDebit, CritEdt.Decimale) ;
TotIR[1].TotCredit:= Arrondi(TotIR[1].TotCredit+RealPer[0].TotCredit, CritEdt.Decimale) ;
TotIR[5].TotDebit:= Arrondi(TotIR[5].TotDebit+RealCum[0].TotDebit, CritEdt.Decimale) ;
TotIR[5].TotCredit:= Arrondi(TotIR[5].TotCredit+RealCum[0].TotCredit, CritEdt.Decimale) ;
(* Pas de total init+rev...
...En Réalisé période
TotIR[1].TotDebit:= Arrondi(TotIR[1].TotDebit+(RealPer[0].TotDebit+RealPer[1].TotDebit), CritEdt.Decimale) ;
TotIR[1].TotCredit:= Arrondi(TotIR[1].TotCredit+(RealPer[0].TotCredit+RealPer[1].TotCredit), CritEdt.Decimale) ;
...En Réalisé cumulé
TotIR[5].TotDebit:= Arrondi(TotIR[5].TotDebit+(RealCum[0].TotDebit+RealCum[1].TotDebit), CritEdt.Decimale) ;
TotIR[5].TotCredit:= Arrondi(TotIR[5].TotCredit+(RealCum[0].TotCredit+RealCum[1].TotCredit), CritEdt.Decimale) ;
*)
{ Calcule du réalisé en Total Edition }
if (CritEdt.BalBud.RuptOnly<>Sur) then
   BEGIN
   { en rupture, le total est fait sur le PrintGroup }
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
{ Calcule du réalisé en Zone libre }
{Initial}
TotLibre[2]:= Arrondi(TotLibre[2]+RealPer[0].TotDebit, CritEdt.Decimale) ;
TotLibre[3]:= Arrondi(TotLibre[3]+RealPer[0].TotCredit, CritEdt.Decimale) ;
TotLibre[10]:= Arrondi(TotLibre[10]+RealCum[0].TotDebit, CritEdt.Decimale) ;
TotLibre[11]:= Arrondi(TotLibre[11]+RealCum[0].TotCredit, CritEdt.Decimale) ;
{Révisé}
TotLibre[22]:= Arrondi(TotLibre[22]+RealPer[1].TotDebit, CritEdt.Decimale) ;
TotLibre[23]:= Arrondi(TotLibre[23]+RealPer[1].TotCredit, CritEdt.Decimale) ;
TotLibre[30]:= Arrondi(TotLibre[30]+RealCum[1].TotDebit, CritEdt.Decimale) ;
TotLibre[31]:= Arrondi(TotLibre[29]+RealCum[1].TotCredit, CritEdt.Decimale) ;

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
   PrintBand:=(CritEdt.BalBud.RuptOnly<>Sur);
   AddGroupLibre(LLibre,Q,fbBudGen,CritEdt.LibreTrie,TotLibre) ;
   END ;
MontantIR ;
end;

procedure TFQRBLecG.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
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

procedure TFQRBLecG.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
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

procedure TFQRBLecG.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
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

procedure TFQRBLecG.BLibreBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
PrintBand:=CritEdt.Rupture=rLibres ;
end;

procedure TFQRBLecG.DLLibreNeedData(var MoreData: Boolean);
var TotLibre         : TabMont77 ;
    Librupt, CodRupt,Lib1, Stcode : String ;
    Quellerupt       : Integer ;
    Chaque,i         : Byte ;
    Col              : TColor ;
    TotMontLibre     : TabTot12 ;
    StMontLibre      : array[0..9] of String ;
    LibRuptInf : Array[1..10] Of TRuptInf ;
begin
  inherited;
MoreData:=False ;
if (CritEdt.Rupture=rLibres) then
   BEGIN
   MoreData:=PrintGroupLibre(LLibre,Q,fbBudGen,CritEdt.LibreTrie,CodRupt,LibRupt,Lib1,TotLibre,Quellerupt,Col,LibRuptInf) ;
   If MoreData then
      BEGIN
      StCode:=CodRupt ;
      Delete(StCode,1,Quellerupt+2) ;
      MoreData:=DansChoixCodeLibre(StCode,Q,fbBudGen,CritEdt.LibreCodes1,CritEdt.LibreCodes2, CritEdt.LibreTrie) ;
      END ;
   if MoreData then
      BEGIN
      { Affichage sur la Bande Nature }
      BLibre.Font.Color:=Col ; BLibreHaut.Font.Color:=Col ;
      TLibRupt.Caption:='' ;
      insert(MsgBox.Mess[6]+' ',CodRupt,Quellerupt+2) ;
      TCodRupt.Caption:=CodRupt+' '+Lib1 ;
      For Chaque:=1 to 5 do
          BEGIN
          for i:=0 to 9 do
              BEGIN
              Case Chaque of
                 1 :  BEGIN { Initial }
                      TotMontLibre[i].TotDebit:=TotLibre[i*2] ; TotMontLibre[i].TotCredit:=TotLibre[i*2+1] ;
                      END ;
                 2 :  BEGIN { Révisé }
                      TotMontLibre[i].TotDebit:=TotLibre[(i*2)+20] ; TotMontLibre[i].TotCredit:=TotLibre[(i*2)+(1+20)] ;
                      END ;
                 3 :  BEGIN { Total Ini+Rév }
                      { Pas de total en Réalisé période (i==1) ni en Réalisé cumulé (i==5) }
                      if (i=1) or (i=5) then BEGIN TotMontLibre[i].TotDebit:=TotLibre[i*2] ; TotMontLibre[i].TotCredit:=TotLibre[i*2+1] END
                                        Else BEGIN TotMontLibre[i].TotDebit:=TotLibre[i*2]+TotLibre[(i*2)+20] ; TotMontLibre[i].TotCredit:=TotLibre[i*2+1]+TotLibre[(i*2)+(1+20)] ; END ;
                      if (CritEdt.BalBud.RuptOnly=Sur) then
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
          For i:=1 to 10 do { Les 10 colonnes de l'état }
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
   END ;
OkEnteteLibre:=MoreData ;
end;

procedure TFQRBLecG.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr1BG_BUDGENE :=TStringField(Q.FindField('BG_BUDGENE'));
Qr1BG_LIBELLE :=TStringField(Q.FindField('BG_LIBELLE'));
Qr1BG_RUB:=TStringField(Q.FindField('BG_RUB'));
end;

procedure TFQRBLecG.FJournalChange(Sender: TObject);
begin
  inherited;
if FJournal.Value='' then Exit ;
InfosJournal ;
end;

procedure TFQRBLecG.InfosJournal ;
Var QJal  : TQuery ;
BEGIN
QJal:=OpenSQL('Select BJ_EXODEB, BJ_EXOFIN, BJ_PERDEB, BJ_PERFIN, BJ_AXE, BJ_NATJAL from BUDJAL Where BJ_BUDJAL="'+FJournal.Value+'"',True) ;
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

procedure TFQRBLecG.FExerciceChange(Sender: TObject);
begin
if FExercice.Value>FExercice_.Value then FExercice_.Value:=FExercice.Value ;
ListePeriode(FExercice.Value,FPeriode1.Items,FPeriode1.Values,True) ;
FPeriode1.ItemIndex:=0 ; DebToDate ;
end;

procedure TFQRBLecG.FExercice_Change(Sender: TObject);
begin
  inherited;
if FExercice_.Value<FExercice.Value then FExercice.Value:=FExercice_.Value ;
ListePeriode(FExercice_.Value,FPeriode2.Items,FPeriode2.Values,False) ;
FPeriode2.ItemIndex:=FPeriode2.Items.Count-1 ; FinToDate ;
end;

procedure TFQRBLecG.FPeriode1Change(Sender: TObject);
begin
  inherited;
if StrToDate(FPeriode1.Value)>StrToDate(FPeriode2.Value) then FPeriode2.Value:=DateToStr(FinDeMois(StrToDate(FPeriode1.Value))) ;
FDateCompta1.Text:=FPeriode1.Value ;
end;


procedure TFQRBLecG.FPeriode2Change(Sender: TObject);
begin
  inherited;
if StrToDate(FPeriode2.Value)<StrToDate(FPeriode1.Value) then FPeriode1.Value:=DateToStr(DebutDeMois(StrToDate(FPeriode2.Value))) ;
FDateCompta2.Text:=FPeriode2.Value ;
end;

procedure TFQRBLecG.DebToDate ;
BEGIN
FDateCompta1.Text:=FPeriode1.Value ;
END ;

procedure TFQRBLecG.FinToDate ;
BEGIN
FDateCompta2.Text:=FPeriode2.Value ;
END ;

Procedure PrepCalcBud(CritEdt : TCritEdt ; Var QMvt : TQuery)  ;
BEGIN
QMvt:=TQuery.Create(Application) ; QMvt.DataBaseName:='SOC' ; QMvt.SQL.Clear ;
QMvt.SQL.Clear ;
QMvt.SQL.Add('Select BE_BUDGENE, BE_EXERCICE, BE_DATECOMPTABLE, BE_NUMEROPIECE,') ;
QMvt.SQL.Add(       'BE_REFINTERNE, BE_ETABLISSEMENT, BE_LIBELLE, BE_VALIDE,') ;
QMvt.SQL.Add(       'BE_BUDJAL,BE_NATUREBUD, BE_QUALIFPIECE, BE_QTE1, BE_QTE2, ') ;
Case CritEdt.Monnaie of
  0 : BEGIN QMvt.SQL.Add('BE_DEBIT DEBIT,BE_CREDIT CREDIT') ; END ;
//  1 : BEGIN QEcr.SQL.Add('E_DEBITDEV DEBIT,E_CREDITDEV CREDIT') ; END ;
//  2 : BEGIN QMvt.SQL.Add('BE_DEBITEURO as DEBIT,BE_CREDITEURO as CREDIT') ;END ;
end ;
QMvt.SQL.Add(' From BUDECR ') ;
QMvt.SQL.Add(' Where BE_BUDGENE=:C1') ;
QMvt.SQL.Add(' And BE_BUDJAL="'+CritEdt.BalBud.Journal+'" ') ;
If CritEdt.BalBud.NatureBud<>'' then QMvt.SQL.Add('And BE_NATUREBUD="'+CritEdt.BalBud.NatureBud+'" ') ;
QMvt.SQL.Add('And BE_DATECOMPTABLE>="'+usdatetime(CritEdt.DateDeb)+'" And BE_DATECOMPTABLE<="'+usdatetime(CritEdt.DateFin)+'" ') ;
QMvt.SQL.Add('And BE_EXERCICE>="'+CritEdt.BalBud.Exo1+'" And BE_EXERCICE<="'+CritEdt.BalBud.Exo2+'" ') ;
QMvt.SQL.Add(TraduitNatureEcr(CritEdt.Qualifpiece, 'BE_QUALIFPIECE', true, CritEdt.ModeRevision)) ;
if CritEdt.Etab<>'' then QMvt.SQL.Add(' And BE_ETABLISSEMENT="'+CritEdt.Etab+'"') ;
if CritEdt.Valide<>'g' then QMvt.SQL.Add(' And BE_VALIDE="'+CritEdt.Valide+'" ') ;
If CritEDT.SQLPlus<>'' Then QMvt.SQL.Add(CritEDT.SQLPlus) ;
QMvt.SQL.Add(' Order By BE_BUDGENE,BE_BUDJAL,BE_EXERCICE,BE_DATECOMPTABLE,BE_BUDSECT,BE_AXE,BE_NUMEROPIECE,BE_QUALIFPIECE ') ;
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
      Reevaluation(DEBIT, CREDIT,CritEdt.BalBud.Resol, CritEdt.BalBud.Taux) ;


      if (DateCompta>=Critedt.Date1) and (DateCompta<Critedt.Date2) then
         BEGIN  CalculBud(0,Tot1, LaNatBud, DEBIT, CREDIT, CritEdt.Decimale) END ; {Calculs Période}
(* GP le 26/02/99
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

procedure TFQRBLecG.BHCptBeforePrint(var PrintBand: Boolean; var Quoi: string);
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
if PrintBand then
   BEGIN
   BG_BUDGENE.Caption:=MsgBox.Mess[8]+'  '+Qr1BG_BUDGENE.AsString ;
   BG_LIBELLE.Caption:=Qr1BG_LIBELLE.AsString ;
   Quoi:=QuoiCpt(0) ;
   END ;
end;

procedure TFQRBLecG.FNatBudChange(Sender: TObject);
begin
  inherited;
FTotIniRev.Enabled:=(FNatBud.Value='') ; //FTotIniRev.Checked:=(FNatBud.Value='') ;
end;

procedure TFQRBLecG.MontantIR ;
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

procedure TFQRBLecG.DLHLibreNeedData(var MoreData: Boolean);
begin
  inherited;
MoreData:=OkEnteteLibre ;
OkEnteteLibre:=False ;
end;

(*
procedure TFQRBLecG.StMontantsBudEcart(M : TabTot12 ; var AffM : array of String) ;
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

procedure TFQRBLecG.FSansRuptClick(Sender: TObject);
begin
  inherited;
if QuelEtatBud=Nature then
   BEGIN
   if Not FSansRupt.Checked then FTablesLibres.Checked:=True ;
   FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
   FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
   END ;
end;

procedure TFQRBLecG.BNouvRechClick(Sender: TObject);
begin
  inherited;
FJournal.ItemIndex:=0 ;
InitResolution(FRESOL) ;
//FRESOL.ItemIndex:=0 ;

end;

procedure TFQRBLecG.FQueTotBudClick(Sender: TObject);
begin
  inherited;
FTotIniRev.Enabled:=Not FQueTotBud.Checked ; If FQueTotBud.Checked Then FTotIniRev.Checked:=FALSE ;
end;


procedure TFQRBLecG.FTotIniRevClick(Sender: TObject);
begin
  inherited;
FQueTotBud.Enabled:=Not (FTotIniRev.Checked) ; If FTotIniRev.Checked Then FQueTotBud.Checked:=FALSE ;
end;

end.
