unit QRBLGESE;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, UtilEdt, HEnt1, Ent1, QRRupt, CpteUtil, EdtLegal,
  CritEDT, Menus, HSysMenu, Calcole, HTB97, HPanel, UiUtil, tCalcCum ;

procedure BLGESE ;
procedure BLGESEZoom(Crit : TCritEdt) ;

type
  TFQRBLGESE = class(TFQR)
    TFJokerSect: THLabel;
    FSection1: THCpteEdit;
    TFaS: TLabel;
    FSection2: THCpteEdit;
    FJokerSect: TEdit;
    TFSection: THLabel;
    HLabel3: THLabel;
    FAxe: THValComboBox;
    HLabel9: THLabel;
    FExceptionS: TEdit;
    TRAxe: TQRLabel;
    RAxe: TQRLabel;
    TRSection: TQRLabel;
    RSection1: TQRLabel;
    TRaS: TQRLabel;
    RSection2: TQRLabel;
    QRLabel6: TQRLabel;
    RExceptionS: TQRLabel;
    TitreColCpt: TQRLabel;
    TitreColLibelle: TQRLabel;
    TTitreColAvant: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel10: TQRLabel;
    TTitreColSelection: TQRLabel;
    QRLabel8: TQRLabel;
    T_DEBIT: TQRLabel;
    TTitreColApres: TQRLabel;
    T_CREDIT: TQRLabel;
    QRLabel9: TQRLabel;
    TTitreColSolde: TQRLabel;
    QRLabel13: TQRLabel;
    REPORT1DEBITanv: TQRLabel;
    REPORT1CREDITanv: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    REPORT1DEBITsum: TQRLabel;
    REPORT1CREDITsum: TQRLabel;
    REPORT1DEBITsol: TQRLabel;
    REPORT1CREDITsol: TQRLabel;
    G_GENERAL: TQRDBText;
    G_LIBELLE: TQRDBText;
    BDSect: TQRBand;
    SOLDEdeb: TQRLabel;
    ANvDEBIT: TQRLabel;
    ANvCREDIT: TQRLabel;
    DEBITsum: TQRLabel;
    CREDITsum: TQRLabel;
    SOLDEcre: TQRLabel;
    S_SECTION: TQRDBText;
    S_LIBELLE: TQRDBText;
    DEBIT: TQRLabel;
    CREDIT: TQRLabel;
    BTOTGENE: TQRBand;
    QRDBText1: TQRDBText;
    QRDBText2: TQRDBText;
    TOT1DEBITanv: TQRLabel;
    TOT1CREDITanv: TQRLabel;
    TOT1DEBITsum: TQRLabel;
    TOT1CREDITsum: TQRLabel;
    TOT1SOLdeb: TQRLabel;
    TOT1SOLcre: TQRLabel;
    LeSoldeD: TQRLabel;
    LeSoldeC: TQRLabel;
    QRLabel14: TQRLabel;
    TOT1DEBIT: TQRLabel;
    TOT1CREDIT: TQRLabel;
    QRLabel33: TQRLabel;
    TOT2DEBITanv: TQRLabel;
    TOT2CREDITanv: TQRLabel;
    TOT2DEBIT: TQRLabel;
    TOT2CREDIT: TQRLabel;
    TOT2DEBITsum: TQRLabel;
    TOT2CREDITsum: TQRLabel;
    TOT2SOLdeb: TQRLabel;
    TOT2SOLcre: TQRLabel;
    QRBand1: TQRBand;
    Trait6: TQRLigne;
    Trait0: TQRLigne;
    Trait5: TQRLigne;
    Trait4: TQRLigne;
    Trait3: TQRLigne;
    Trait2: TQRLigne;
    Ligne1: TQRLigne;
    REPORT2DEBITanv: TQRLabel;
    REPORT2CREDITanv: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    REPORT2DEBITsum: TQRLabel;
    REPORT2CREDITsum: TQRLabel;
    REPORT2DEBITsol: TQRLabel;
    REPORT2CREDITsol: TQRLabel;
    QSec: TQuery;
    SSec: TDataSource;
    DL: TQRDetailLink;
    MsgLibel: THMsgBox;
    FCol1: TCheckBox;
    FCol2: TCheckBox;
    FCol3: TCheckBox;
    FCol4: TCheckBox;
    DLRupt: TQRDetailLink;
    BRupt: TQRBand;
    DetDebRupt: TQRLabel;
    CumCreRupt: TQRLabel;
    CumDebRupt: TQRLabel;
    TCodRupt: TQRLabel;
    SumDebRupt: TQRLabel;
    SumCreRupt: TQRLabel;
    DetCreRupt: TQRLabel;
    SolCreRupt: TQRLabel;
    SolDebRupt: TQRLabel;
    FLigneRupt: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BTOTGENEBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BTOTGENEAfterPrint(BandPrinted: Boolean);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure FSection1Change(Sender: TObject);
    procedure BDSectBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure QSecAfterOpen(DataSet: TDataSet);
    procedure FAxeChange(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure DLRuptNeedData(var MoreData: Boolean);
    procedure BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FSansRuptClick(Sender: TObject);
    procedure FPlanRupturesChange(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure FJokerSectKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private    { Déclarations privées }
    Qr11G_GENERAL, Qr11G_LIBELLE,
    Qr11G_CORRESP1, Qr11G_CORRESP2         : TStringField;
    Qr12S_AXE, Qr12S_SECTION               : TStringField;
    Qr12S_LIBELLE                          : TStringField;
//    QBal,QBalCum                           : TQuery ;
    LRupt                                 : TStringList ;
    Affiche                                : Boolean ;
    StReportGen                            : String ;
    TotCpt, TotCptGene, TotEdt : TabTot ;
    procedure GenereSQLSub ;
    Function  QuoiCpt(i : Integer) : String ;
    Procedure BalGeSeZoom(Quoi : String) ;
    Function  ColonnesOK : Boolean ;
    Procedure CalcTotalEdt(AnVD, AnVC, D, C : Double );
  public    { Déclarations publiques }
    procedure GenereSQL ; Override ;
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
  end;

implementation

{$R *.DFM}

procedure BLGESE ;
var QR : TFQRBLGESE ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQRBLGESE.Create(Application) ;
Edition.Etat:=etBalGenAna ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalGeSeZoom ;
QR.InitType (nbGen,neBal,msGenAna,'QRBAGESE','',TRUE,FALSE,TRUE,Edition) ;
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

procedure BLGESEZoom(Crit : TCritEdt) ;
var QR : TFQRBLGESE ;
    Edition : TEdition ;
BEGIN
QR:=TFQRBLGESE.Create(Application) ;
Edition.Etat:=etBalGenAna ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.BalGeSeZoom ;
 QR.CritEdt:=Crit ;
 QR.InitType (nbGen,neBal,msGenAna,'QRBAGESE','',FALSE,TRUE,TRUE,Edition) ;
 finally
 QR.Free ;
 end ;
END ;

Function TFQRBLGESE.QuoiCpt(i : Integer) : String ;
BEGIN
Case i Of
  0 : Result:=Qr11G_General.AsString+'#'+Qr11G_LIBELLE.AsString+'@'+'1' ;
  1 : Result:=Qr11G_General.AsString+'#'+Qr11G_LIBELLE.AsString+'@'+'1' ;
  2 : Result:=Qr12S_SECTION.AsString+' '+Qr11G_General.AsString+'#'+S_LIBELLE.Caption+'@'+'3'+Qr12S_AXE.AsString ;
 end ;
END ;

Procedure TFQRBLGESE.BalGeSeZoom(Quoi : String) ;
Var Lp,i: Integer ;
BEGIN
Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
i:=StrToInt(Copy(Quoi,Lp+1,1)) ;
ZoomEdt(i,Quoi) ;
END ;

procedure TFQRBLGESE.FinirPrint ;
begin
Inherited ;
//QBal.Free ; QSec.Close ; QBalCum.Free ;
QSec.Close ;
If GCalc<>NIL Then BEGIN GCalc.Free ; GCalc:=NIL ; END ;
if CritEdt.Rupture<>rRien then VideRupt(LRupt) ;
end;

procedure TFQRBLGESE.InitDivers ;
BEGIN
Inherited ;
{ Init du Total Général }
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
Fillchar(TotCpt,SizeOf(TotCptGene),#0) ;
{ Labels sur les bandes }
DateBandeauBalance(TTitreColAvant,TTitreColSelection,TTitreColApres,TTitreColSolde,CritEdt,
                   MsgLibel.Mess[0],MsgLibel.Mess[1],MsgLibel.Mess[2]) ;
ChangeFormatCompte(fbGene,AxeToFb(CritEdt.Bal.Axe),Self,G_GENERAL.Width,1,2,G_GENERAL.Name) ;
BTOTGENE.Frame.DrawBottom:=CritEDT.BAL.FormatPrint.PrSepCompte ;
BRupt.Frame.DrawTop:=CritEDT.BAL.FormatPrint.PrSepRupt  ;
BRupt.Frame.DrawBottom:=BRupt.Frame.DrawTop ;
END ;

procedure TFQRBLGESE.GenereSQL ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
AlimSQLMul(Q,0) ; Q.Close ; ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
Q.Open ;
GenereSQLSub ;
END;

Procedure TFQRBLGESE.GenereSQLSUB ;
BEGIN
Inherited ;
AlimSQLMul(QSec,1) ; ChangeSQL(QSec) ; //QSec.Prepare ;
PrepareSQLODBC(QSec) ;
QSec.Open ;
END ;

procedure TFQRBLGESE.RenseigneCritere ;
{ Récupération des champs du multicritère dans l'entête d'état }
BEGIN
Inherited ;
If OkZoomEdt Then Exit ;
RAxe.Caption:=FAxe.Text ;
if CritEdt.SJoker then
   BEGIN
   TRSection.Caption:=MsgLibel.Mess[10] ; RSection1.Caption:=FJokerSect.Text ;
   END Else
   BEGIN
   TRSection.Caption:=MsgLibel.Mess[9] ;
   RSection1.Caption:=CritEdt.LSCpt1 ; RSection2.Caption:=CritEdt.LScpt2 ;
   END ;

RSection2.Visible:=Not CritEdt.SJoker ; TRaS.Visible:=Not CritEdt.SJoker ;
RExceptionS.Caption:=FExceptionS.Text ;
(*CaseACocher(FSolde, RSolde) ;
CaseACocher(FCol1, RCol1) ; CaseACocher(FCol2, RCol2) ;
CaseACocher(FCol3, RCol3) ; CaseACocher(FCol4, RCol4) ;*)
END ;

procedure TFQRBLGESE.ChoixEdition ;
{ Initialisation des options d'édition }
BEGIN
Inherited ;
DLRupt.PrintBefore:=TRUE ;
Case CritEdt.Rupture of
  rLibres   : BEGIN
              DLRupt.PrintBefore:=FALSE ;
              ChargeGroup(LRupt,['G00','G01','G02','G03','G04','G05','G06','G07','G08','G09']) ;
              END ;
  rRuptures : BEGIN
              ChargeRupt(LRupt, 'RUG',CritEdt.Bal.PlanRupt,CritEdt.Bal.CodeRupt1,CritEdt.Bal.CodeRupt2) ;
              NiveauRupt(LRupt);
              END  ;
   rCorresp : BEGIN
              ChargeRuptCorresp(LRupt, CritEdt.Bal.PlanRupt,CritEdt.Bal.CodeRupt1,CritEdt.Bal.CodeRupt2, False) ;
              NiveauRupt(LRupt);
              END ;
  End ;
END ;

procedure TFQRBLGESE.RecupCritEdt ;
var st : String ;
    NonLibres : Boolean ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  Composite:=TRUE ;
  SJoker:=FJokerSect.Visible ;
  If SJoker Then BEGIN SCpt1:=FJokerSect.Text ; SCpt2:=FJokerSect.Text ; END
            Else BEGIN
                 SCpt1:=FSection1.Text  ; SCpt2:=FSection2.Text ;
                 PositionneFourchetteST(FSection1,FSection2,LSCpt1,LSCpt2) ;
                 END ;
  Bal.Axe:=FAxe.Value ;
  Bal.SSauf:=Trim(FExceptionS.Text) ;
  st:=Bal.SSauf ;
  If Bal.SSauf<>'' then St:=' And '+AnalyseCompte(FExceptionS.Text,AxeToFb(FAxe.Value),True,FALSE) ;
  Bal.SSqlSauf:=St ;
  BAL.RuptOnly:=QuelleTypeRupt(0,FSAnsRupt.Checked,FAvecRupt.Checked,FSurRupt.Checked) ;
  NonLibres:=((Rupture=rRuptures) or (Rupture=rCorresp)) ;
  if NonLibres then BAL.PlanRupt:=FPlanRuptures.Value ;
  if NonLibres then BEGIN BAL.CodeRupt1:=FCodeRupt1.Text ; BAL.CodeRupt2:=FCodeRupt2.Text ; END ;
  If (CritEdt.Rupture=rCorresp) Then Bal.PlansCorresp:=StrToInt(Copy(FPlanRuptures.Value,3,1)) ;
  Bal.OnlyCptAssocie:=(Rupture<>rRien) and FOnlyCptAssocie.Checked ;
  With Bal.FormatPrint Do
    BEGIN
    PrSepRupt:=FLigneRupt.Checked ;
    PutTabCol(1,TabColl[1],TitreColCpt.Tag,TRUE) ;
    PutTabCol(2,TabColl[2],TitreColLibelle.Tag,TRUE) ;
    PutTabCol(3,TabColl[3],TTitreColAvant.Tag,FCol1.Checked) ;
    PutTabCol(4,TabColl[4],TTitreColSelection.Tag,FCol2.Checked) ;
    PutTabCol(5,TabColl[5],TTitreColApres.Tag,FCol3.Checked) ;
    PutTabCol(6,TabColl[6],TTitreColSolde.Tag,FCol4.Checked) ;
    END ;
  END ;
END ;

Function TFQRBLGESE.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK and ColonnesOK;
If Result Then
   BEGIN
//   QBal:=PrepareTotCptM(fbGene,AxeToFb(CritEdt.Bal.Axe),QuelTypeEcr,Dev,Etab,Exo) ;
   Gcalc:=TGCalculCum.create(Deux,fbGene,AxeToFb(CritEdt.Bal.Axe),QuelTypeEcr,Dev,Etab,Exo,DevEnp,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE) ;
   GCalc.initCalcul('','','',CritEdt.Bal.Axe,CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                    CritEdt.Date1,CritEdt.Bal.Date11,TRUE) ;
//   QBal:=PrepareTotCptMCum(QBalCum,fbGene,AxeToFb(CritEdt.Bal.Axe),QuelTypeEcr,Dev,Etab,Exo) ;
   END ;
END ;

Function TFQRBLGESE.ColonnesOK : Boolean ;
BEGIN
If OkZoomEdt then begin Result:=True ; exit ; end ;
Result:=(FCol1.Checked or FCol2.Checked  or FCol3.Checked or FCol4.Checked) ;
If Not Result then NumErreurCrit:=7 ;
END ;

procedure TFQRBLGESE.FormShow(Sender: TObject);
begin
//FAxe.ItemIndex:=0 ;
If FAxe.Values.Count>0 Then FAxe.Value:=FAxe.Values[0] ;
{$IFDEF CCS3}
FAxe.Visible:=FALSE ; HLabel3.Visible:=FALSE ;
{$ENDIF}
HelpContext:=7466000 ;
//Standards.HelpContext:=7466010 ;
//Avances.HelpContext:=7466020 ;
//Mise.HelpContext:=7466030 ;
//Option.HelpContext:=7466040 ;
//TabRuptures.HelpContext:=7466050 ;

  inherited;
TabSup.TabVisible:=False;
FSection1.Text:='' ; FSection2.Text:='' ;
FCol1.Checked:=TRUE ; FCol2.Checked:=TRUE ; FCol3.Checked:=TRUE ; FCol4.Checked:=TRUE ;
FOnlyCptAssocie.Checked:=False ; FOnlyCptAssocie.Enabled:=False ;
FSection1.SynJoker:=TRUE ; FSection2.SynJoker:=TRUE ;
end;

procedure TFQRBLGESE.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TITREREPORTH.Left:=TITREREPORTB.Left ;
TITREREPORTH.Width:=TITREREPORTB.Width ;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
Report1DebitAnv.Caption:=Report2DebitAnv.Caption ;
Report1CreditAnv.Caption:=Report2CreditAnv.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
Report1DebitSum.Caption:=Report2DebitSum.Caption ;
Report1CreditSum.Caption:=Report2CreditSum.Caption ;
Report1DebitSol.Caption:=Report2DebitSol.Caption ;
Report1CreditSol.Caption:=Report2CreditSol.Caption ;
end;

procedure TFQRBLGESE.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
StReportGen:=G_GENERAL.Caption ;
InitReport([2],CritEdt.Bal.FormatPrint.Report) ;
Case CritEdt.Bal.TypCpt of  { 0 : Mvt sur L'exo,    3 : MVt sur la Période }
  0 :  PrintBand:=True ;
  3 :  PrintBand:= TRUE ;
 end ;
Case CritEdt.Rupture of
  rLibres    : if CritEdt.Bal.OnlyCptAssocie then PrintBand:=DansRuptLibre(Q,fbGene,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
  rRuptures  : if CritEdt.Bal.OnlyCptAssocie then PrintBand:=DansRupt(LRupt,Qr11G_GENERAL.AsString) ;
  rCorresp   : if CritEdt.Bal.OnlyCptAssocie then
                  if CritEDT.Bal.PlansCorresp=1 then PrintBand:=DansRuptCorresp(LRupt,Qr11G_CORRESP1.AsString) Else
                  if CritEDT.Bal.PlansCorresp=2 then PrintBand:=DansRuptCorresp(LRupt,Qr11G_CORRESP2.AsString) ;
//                 if CritEDT.Bal.PlansCorresp=1 then PrintBand:=(Qr11G_CORRESP1.AsString<>'') Else
//                 if CritEDT.Bal.PlansCorresp=2 then PrintBand:=(Qr11G_CORRESP2.AsString<>'') ;
  End;
Affiche:=PrintBand ; { Pour ne pas afficher les sous comptes }
if CritEdt.Bal.RuptOnly=Sur then PrintBand:=False ;
If PrintBand then Quoi:=Quoicpt(0) ;
end;

procedure TFQRBLGESE.BTOTGENEBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
if CritEdt.Bal.RuptOnly=Sur then PrintBand:=False Else PrintBand:=Affiche ;
if PrintBand then
   BEGIN
   Quoi:=Quoicpt(1) ;
   TOT1DEBITanv.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGene[1].TotDebit, CritEdt.AfficheSymbole ) ;
   TOT1CREDITanv.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGene[1].TotCredit,CritEdt.AfficheSymbole ) ;
   TOT1DEBIT.Caption :=   AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGene[2].TotDebit, CritEdt.AfficheSymbole ) ;
   TOT1CREDIT.Caption:=   AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGene[2].TotCredit, CritEdt.AfficheSymbole ) ;
   TOT1DEBITsum.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGene[3].TotDebit, CritEdt.AfficheSymbole ) ;
   TOT1CREDITsum.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGene[3].TotCredit, CritEdt.AfficheSymbole ) ;

   TOT1SOLdeb.Caption :=  AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGene[4].TotDebit, CritEdt.AfficheSymbole ) ;
   TOT1SOLcre.Caption:=   AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGene[4].TotCredit, CritEdt.AfficheSymbole ) ;
   CumulVersSolde(TotCptGene[4]) ;
   if TotCptGene[4].TotDebit=0 then
      BEGIN
      LeSoldeD.Caption:='' ;
      LeSoldeC.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGene[4].TotCredit,CritEdt.AfficheSymbole ) ;
      END else
      BEGIN
      LeSoldeD.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGene[4].TotDebit,CritEdt.AfficheSymbole ) ;
      LeSoldeC.Caption:='' ;
      END ;
   END ;
end;

procedure TFQRBLGESE.BTOTGENEAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Fillchar(TotCptGene,SizeOf(TotCptGene),#0) ;
InitReport([2],CritEdt.Bal.FormatPrint.Report) ;
end;

Procedure TFQRBLGESE.CalcTotalEdt(AnVD, AnVC, D, C : Double );
Var SumD, SumC, TotSoldeD, TotSoldeC : Double ;
BEGIN
SumD:=Arrondi(AnVD+D, CritEdt.Decimale) ;
SumC:=Arrondi(AnVC+C, CritEdt.Decimale) ;
CalCulSolde(SumD,SumC,TotSoldeD,TotSoldeC) ;
TotEdt[1].TotDebit:= Arrondi(TotEdt[1].TotDebit  + AnVD, CritEdt.Decimale) ;
TotEdt[1].TotCredit:=Arrondi(TotEdt[1].TotCredit + AnVC, CritEdt.Decimale) ;
TotEdt[2].TotDebit:= Arrondi(TotEdt[2].TotDebit  + D, CritEdt.Decimale) ;
TotEdt[2].TotCredit:=Arrondi(TotEdt[2].TotCredit + C, CritEdt.Decimale) ;
TotEdt[3].TotDebit:= Arrondi(TotEdt[3].TotDebit  + SumD, CritEdt.Decimale) ;
TotEdt[3].TotCredit:=Arrondi(TotEdt[3].TotCredit + SumC, CritEdt.Decimale) ;
TotEdt[4].TotDebit:= Arrondi(TotEdt[4].TotDebit + TotSoldeD, CritEdt.Decimale) ;
TotEdt[4].TotCredit:=Arrondi(TotEdt[4].TotCredit + TotSoldeC, CritEdt.Decimale) ;
END ;

procedure TFQRBLGESE.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TOT2DEBITanv.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[1].TotDebit,CritEdt.AfficheSymbole ) ;
TOT2CREDITanv.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotEdt[1].TotCredit,CritEdt.AfficheSymbole ) ;

TOT2DEBIT.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[2].TotDebit,CritEdt.AfficheSymbole ) ;
TOT2CREDIT.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[2].TotCredit,CritEdt.AfficheSymbole ) ;

TOT2DEBITsum.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Totedt[3].TotDebit,CritEdt.AfficheSymbole ) ;
TOT2CREDITsum.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[3].TotCredit,CritEdt.AfficheSymbole ) ;

TOT2SOLdeb.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[4].TotDebit,CritEdt.AfficheSymbole ) ;
TOT2SOLcre.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[4].TotCredit,CritEdt.AfficheSymbole ) ;
BOTTOMREPORT.enabled:=FALSE ;

end;

procedure TFQRBLGESE.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var MReport : TabTRep ;
begin
  inherited;
TITREREPORTB.Left:=TitreColCpt.Left ;
TITREREPORTB.Width:=TitreColCpt.Width+TitreColLibelle.Width+1 ;
Case QuelReportBAL(CritEdt.Bal.FormatPrint.Report,MReport) Of
  1 : TITREREPORTB.Caption:=MsgLibel.Mess[7] ;
  2 : TITREREPORTB.Caption:=MsgLibel.Mess[8]+' '+StReportGen ;
  END ;
Report2DebitAnv.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[1].TotDebit, CritEdt.AfficheSymbole ) ;
Report2CreditAnv.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[1].TotCredit, CritEdt.AfficheSymbole ) ;
Report2Debit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[2].TotDebit, CritEdt.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[2].TotCredit, CritEdt.AfficheSymbole ) ;
Report2DebitSum.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[3].TotDebit, CritEdt.AfficheSymbole ) ;
Report2CreditSum.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[3].TotCredit, CritEdt.AfficheSymbole ) ;
Report2DebitSol.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[4].TotDebit, CritEdt.AfficheSymbole ) ;
Report2CreditSol.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[4].TotCredit, CritEdt.AfficheSymbole ) ;
end;

procedure TFQRBLGESE.FSection1Change(Sender: TObject);
Var SAvecJoker : Boolean ;
begin
  inherited;
SAvecJoker:=Joker(FSection1, FSection2, FJokerSect ) ;
TFaS.Visible:=Not SAvecJoker ;
TFSection.Visible:=Not SAvecJoker ;
TFJokerSect.Visible:=SAvecJoker ;
end;

procedure TFQRBLGESE.BDSectBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var TotCpt1, TotCpt2, Tot : TabTot ;
    MReport               : TabTRep ;
    CptRupt               : String ;
begin
  inherited;
Fillchar(Tot,SizeOf(Tot),#0) ; Fillchar(TotCpt1,SizeOf(TotCpt1),#0) ; Fillchar(TotCpt2,SizeOf(TotCpt2),#0) ;
Fillchar(Mreport,SizeOf(MReport),#0) ;
PrintBand:=Affiche ;
if PrintBand then
   BEGIN
   (*
   ExecuteTotCptM(QBal,QG_GENERAL.AsString, QSecS_SECTION.AsString,CritEdt.Date1,CritEdt.Bal.Date11,
                 CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,TotCpt1,FALSE) ;
   ExecuteTotCptM(QBal,QG_GENERAL.AsString, QSecS_SECTION.AsString,CritEdt.Bal.Date21,CritEdt.Date2,
                 CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,TotCpt2,FALSE) ;
   ExecuteTotCptMCum(QBal,QBalCum,Qr11G_GENERAL.AsString, Qr12S_SECTION.AsString,'',CritEdt.Bal.Axe,CritEdt.Date1,CritEdt.Bal.Date11,
                 CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,TotCpt1,FALSE) ;
   ExecuteTotCptMCum(QBal,QBalCum,Qr11G_GENERAL.AsString, Qr12S_SECTION.AsString,'',CritEdt.Bal.Axe,CritEdt.Bal.Date21,CritEdt.Date2,
                 CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,TotCpt2,FALSE) ;
   *)
   GCAlc.ReInitCalcul(Qr11G_GENERAL.AsString, Qr12S_SECTION.AsString,CritEdt.Date1,CritEdt.Bal.Date11) ;
   GCalc.Calcul ; TotCpt1:=GCalc.ExecCalc.TotCpt ;

   GCAlc.ReInitCalcul(Qr11G_GENERAL.AsString, Qr12S_SECTION.AsString,CritEdt.Bal.Date21,CritEdt.Date2) ;
   GCalc.Calcul ; TotCpt2:=GCalc.ExecCalc.TotCpt ;
   CumulVersSolde(TotCpt1[0]) ;
   if CritEdt.Date1=CritEdt.Bal.Date11 Then Fillchar(TotCpt1[1],SizeOf(TotCpt1[1]),#0) ;
   Tot[1].TotDebit:= TotCpt1[0].TotDebit  + TotCpt1[1].TotDebit ;
   Tot[1].TotCredit:=TotCpt1[0].TotCredit + TotCpt1[1].TotCredit ;
   Tot[2].TotDebit:= TotCpt2[1].TotDebit ;
   Tot[2].TotCredit:=TotCpt2[1].TotCredit ;
   Tot[3].TotDebit:= Tot[1].TotDebit+Tot[2].TotDebit ;
   Tot[3].TotCredit:=Tot[1].TotCredit+Tot[2].TotCredit ;
   MReport[1].TotDebit:=TotCpt1[0].TotDebit+TotCpt1[1].TotDebit ;
   MReport[1].TotCredit:=TotCpt1[0].TotCredit+TotCpt1[1].TotCredit ;
   MReport[2].TotDebit:=TotCpt2[1].TotDebit ;
   MReport[2].TotCredit:=TotCpt2[1].TotCredit ;
   MReport[3].TotDebit:=Tot[2].TotDebit+Tot[1].TotDebit ;
   MReport[3].TotCredit:=Tot[2].TotCredit+Tot[1].TotCredit ;
   (**  Rony 24/10/97 Passé au moule de CalcTotalEdt() en non Sur Rupture ou pas ...
   TotEdt[1].totDebit:=Arrondi(TotEdt[1].totDebit + Tot[1].TotDebit,CritEdt.Decimale) ;
   TotEdt[1].totCredit:=Arrondi(TotEdt[1].totCredit + Tot[1].TotCredit,CritEdt.Decimale) ;
   TotEdt[2].totDebit:=Arrondi(TotEdt[2].totDebit + Tot[2].TotDebit,CritEdt.Decimale) ;
   TotEdt[2].totCredit:=Arrondi(TotEdt[2].totCredit + Tot[2].TotCredit,CritEdt.Decimale) ;
   TotEdt[3].totDebit:=Arrondi(TotEdt[3].totDebit + Tot[3].TotDebit,CritEdt.Decimale) ;
   TotEdt[3].totCredit:=Arrondi(TotEdt[3].totCredit + Tot[3].TotCredit,CritEdt.Decimale) ;
   *)
   TotCptGene[1].totDebit:=Arrondi(TotCptGene[1].totDebit + Tot[1].TotDebit,CritEdt.Decimale) ;
   TotCptGene[1].totCredit:=Arrondi(TotCptGene[1].totCredit + Tot[1].TotCredit,CritEdt.Decimale) ;
   TotCptGene[2].totDebit:=Arrondi(TotCptGene[2].totDebit + Tot[2].TotDebit,CritEdt.Decimale) ;
   TotCptGene[2].totCredit:=Arrondi(TotCptGene[2].totCredit + Tot[2].TotCredit,CritEdt.Decimale) ;
   TotCptGene[3].totDebit:=Arrondi(TotCptGene[3].totDebit + Tot[3].TotDebit,CritEdt.Decimale) ;
   TotCptGene[3].totCredit:=Arrondi(TotCptGene[3].totCredit + Tot[3].TotCredit,CritEdt.Decimale) ;
   Tot[4]:=Tot[3] ;
   CumulVersSolde(Tot[4]) ;
   if Tot[4].TotDebit=0 then
      BEGIN
      MReport[4].TotCredit:=Tot[4].TotCredit ;
      TotCptGene[4].TotCredit:=Arrondi(TotCptGene[4].TotCredit+Tot[4].TotCredit,CritEdt.Decimale) ;
      SOLDEdeb.Caption:='' ;
      SOLDEcre.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[4].TotCredit,CritEdt.AfficheSymbole) ;
      END else
      BEGIN
      MReport[4].TotDebit:=Tot[4].TotDebit ;
      TotCptGene[4].TotDebit:=Arrondi(TotCptGene[4].TotDebit+Tot[4].TotDebit,CritEdt.Decimale) ;
      SOLDEdeb.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[4].TotDebit,CritEdt.AfficheSymbole ) ;
      SOLDEcre.Caption:='' ;
      END ;
   AddReportBAL([1,2],CritEdt.Bal.FormatPrint.Report,MReport,CritEdt.Decimale) ;
   Case CritEdt.Rupture of
     rLibres   : AddGroupLibre(LRupt,Q,fbGene,CritEdt.LibreTrie,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,0]) ;
     rRuptures : AddRupt(LRupt,Qr11G_GENERAL.AsString,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,0]) ;
     rCorresp  : BEGIN
                 Case CritEDT.Bal.PlansCorresp Of
                   1 : If Qr11G_CORRESP1.AsString<>'' Then CptRupt:=Qr11G_CORRESP1.AsString+Qr11G_GENERAL.AsString
                                                      Else CptRupt:='.'+Qr11G_GENERAL.AsString ;
                   2 : If Qr11G_CORRESP2.AsString<>'' Then CptRupt:=Qr11G_CORRESP2.AsString+Qr11G_GENERAL.AsString
                                                      Else CptRupt:='.'+Qr11G_GENERAL.AsString ;
                   Else CptRupt:=Qr11G_GENERAL.AsString ;
                  End ;
                 AddRuptCorres(LRupt,CptRupt,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,0]) ;
                 END ;
     End ;
   If (CritEdt.Bal.RuptOnly=Sur) then
      BEGIN
      PrintBand:=False ;
      IF CritEdt.Rupture=rLibres then CalcTotalEdt(Tot[1].TotDebit, Tot[1].TotCredit, Tot[2].TotDebit, Tot[2].TotCredit) ;
      END Else CalcTotalEdt(Tot[1].TotDebit, Tot[1].TotCredit, Tot[2].TotDebit, Tot[2].TotCredit) ;
   If PrintBand then Quoi:=Quoicpt(2) ;
   End ;
AnvDebit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[1].TotDebit,CritEdt.AfficheSymbole ) ;
AnvCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[1].TotCredit,CritEdt.AfficheSymbole ) ;

debit.caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[2].TotDebit,CritEdt.AfficheSymbole ) ;
credit.caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[2].TotCredit,CritEdt.AfficheSymbole ) ;

DEBITsum.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[3].TotDebit,CritEdt.AfficheSymbole ) ;
CREDITsum.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[3].TotCredit,CritEdt.AfficheSymbole ) ;
end;

procedure TFQRBLGESE.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr11G_GENERAL         :=TStringField(Q.FindField('G_GENERAL'));
Qr11G_LIBELLE         :=TStringField(Q.FindField('G_LIBELLE'));
If CritEDT.Rupture=rCorresp then
   BEGIN
   Qr11G_CORRESP1         :=TStringField(Q.FindField('G_CORRESP1'));
   Qr11G_CORRESP2         :=TStringField(Q.FindField('G_CORRESP2'));
   END ;
end;

procedure TFQRBLGESE.QSecAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr12S_AXE    :=TStringField(QSec.FindField('S_AXE'));
Qr12S_SECTION:=TStringField(QSec.FindField('S_SECTION'));
Qr12S_LIBELLE:=TStringField(QSec.FindField('S_LIBELLE'));
end;

procedure TFQRBLGESE.FAxeChange(Sender: TObject);
Var St : String ;
     A : Char ;
begin
  inherited;
If QRLoading then Exit ;
FSection1.Clear ; FSection2.Clear ;
IF FAxe.Value='' then Exit ;
ST:=FAxe.Value ; A:=St[2] ;
Case A of
   '1' : BEGIN
         FSection1.ZoomTable:=tzSection ; FSection2.ZoomTable:=tzSection ;
         FCpte1.ZoomTable:=tzGVentil1 ; FCpte2.ZoomTable:=tzGVentil1 ;
         END ;
   '2' : BEGIN
         FSection1.ZoomTable:=tzSection2 ; FSection2.ZoomTable:=tzSection2 ;
         FCpte1.ZoomTable:=tzGVentil2 ; FCpte2.ZoomTable:=tzGVentil2 ;
         END ;
   '3' : BEGIN
         FSection1.ZoomTable:=tzSection3 ; FSection2.ZoomTable:=tzSection3 ;
         FCpte1.ZoomTable:=tzGVentil3 ; FCpte2.ZoomTable:=tzGVentil3 ;
         END ;
   '4' : BEGIN
         FSection1.ZoomTable:=tzSection4 ; FSection2.ZoomTable:=tzSection4 ;
         FCpte1.ZoomTable:=tzGVentil4 ; FCpte2.ZoomTable:=tzGVentil4 ;
         END ;
   '5' : BEGIN
         FSection1.ZoomTable:=tzSection5 ; FSection2.ZoomTable:=tzSection5 ;
         FCpte1.ZoomTable:=tzGVentil5 ; FCpte2.ZoomTable:=tzGVentil5 ;
         END ;
  end ;
end;

procedure TFQRBLGESE.BNouvRechClick(Sender: TObject);
begin
If FJokerSect.Visible then FJokerSect.Text:='' ;
  inherited;
end;

procedure TFQRBLGESE.DLRuptNeedData(var MoreData: Boolean);
Var TotRupt : Array[0..12] of Double ;
    CodRupt, LibRupt, Lib1, CptRupt, stcode : String ;
    QuelleRupt : Integer ;
    SumD, SumC, TotSoldeD, TotSoldeC : Double ;
    Col : TColor ;
    OkOk, DansTotal, AddTotEdt : Boolean ;
    LibRuptInf : Array[1..10] Of TRuptInf ;
begin
  inherited;
MoreData:=False ;
TotSoldeD:=0 ; TotSoldeC:=0 ;
case CritEdt.Rupture of
  rLibres   : BEGIN
              MoreData:=PrintGroupLibre(LRupt,Q,fbGene,CritEdt.LibreTrie,CodRupt,LibRupt,Lib1,TotRupt,Quellerupt,Col,LibRuptInf) ;
              If MoreData then
                 BEGIN
                 StCode:=CodRupt ;
                 Delete(StCode,1,Quellerupt+2) ;
                 MoreData:=DansChoixCodeLibre(StCode,Q,fbGene,CritEdt.LibreCodes1,CritEdt.LibreCodes2, CritEdt.LibreTrie) ;
                 BRupt.Font.Color:=Col ;
                 END ;
              END ;
  rRuptures : MoreData:=PrintRupt(LRupt,Qr11G_GENERAL.AsString,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt) ;
  rCorresp  : BEGIN
              OkOk:=TRUE ;
              Case CritEDT.Bal.PlansCorresp  Of
                1 : If Qr11G_CORRESP1.AsString<>'' Then CptRupt:=Qr11G_CORRESP1.AsString+Qr11G_GENERAL.AsString
                                                   Else CptRupt:='.'+Qr11G_GENERAL.AsString ;
                2 : If Qr11G_CORRESP2.AsString<>'' Then CptRupt:=Qr11G_CORRESP2.AsString+Qr11G_GENERAL.AsString
                                                   Else CptRupt:='.'+Qr11G_GENERAL.AsString ;
                Else OkOk:=FALSE ;
                End ;
              If OkOk Then MoreData:=PrintRupt(LRupt,CptRupt,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt) Else MoreData:=FALSE ;
              END ;
  End ;
If MoreData then
   BEGIN
   SumD:=TotRupt[1]+TotRupt[3] ;
   SumC:=TotRupt[2]+TotRupt[4] ;
   TCodRupt.Width:=TitreColCpt.Width+TitreColLibelle.Width+1 ;
   TCodRupt.Caption:='' ;
   if CritEdt.Rupture=rLibres then
      BEGIN
      insert(MsgLibel.Mess[12]+' ',CodRupt,Quellerupt+2) ;
      TCodRupt.Caption:=CodRupt+' '+Lib1 ;
      END Else TCodRupt.Caption:=CodRupt+'   '+LibRupt ;
   CalCulSolde(SumD,SumC,TotSoldeD,TotSoldeC) ;
   CumDebRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotRupt[1], CritEDT.AfficheSymbole) ;
   CumCreRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotRupt[2], CritEDT.AfficheSymbole) ;
   DetDebRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotRupt[3], CritEDT.AfficheSymbole) ;
   DetCreRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotRupt[4], CritEDT.AfficheSymbole) ;
   SumDebRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumD  , CritEDT.AfficheSymbole) ;
   SumCreRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumC  , CritEDT.AfficheSymbole) ;
   SolDebRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotSoldeD, CritEDT.AfficheSymbole) ;
   SolCreRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotSoldeC, CritEDT.AfficheSymbole) ;
   if CritEdt.Bal.RuptOnly=Sur then
      BEGIN
      TCodRupt.Tag:=1 ;
      AddTotEdt:=False ;
      if (CritEdt.Rupture=rLibres) then AddTotEdt:=False else
      if (CritEdt.Rupture=rRuptures) or (CritEdt.Rupture=rCorresp) then AddTotEdt:=DansTotal ;
      if AddTotEdt then CalcTotalEdt(TotRupt[1], TotRupt[2], TotRupt[3], TotRupt[4]) ;
      END Else TCodRupt.Tag:=2 ;
   END ;
end;

procedure TFQRBLGESE.BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.Rupture<>rRien) ;
end;

procedure TFQRBLGESE.FSansRuptClick(Sender: TObject);
begin
  inherited;
FLigneRupt.Enabled:=Not FSansRupt.Checked ;
FLigneRupt.checked:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
FRupturesClick(Nil) ;
end;

procedure TFQRBLGESE.FPlanRupturesChange(Sender: TObject);
begin
  inherited;
If QRLoading Then Exit ;
if FPlansCo.Checked then CorrespToCodes(FPlanRuptures,FCodeRupt1,FCodeRupt2) Else
if FRuptures.Checked then RuptureToCodes(FPlanRuptures,FCodeRupt1,FCodeRupt2,FbGene) ;
FCodeRupt1.ItemIndex:=0 ; FCodeRupt2.ItemIndex:=FCodeRupt2.Items.Count-1 ;
end;

procedure TFQRBLGESE.FRupturesClick(Sender: TObject);
begin
  inherited;
if FPlansCo.Checked then
   BEGIN
   FGroupRuptures.Caption:=' '+MsgLibel.Mess[15] ;
   CorrespToCombo(FPlanRuptures,FbGene) ;
   END Else
   BEGIN
   FPlanRuptures.Reload ;
   FGroupRuptures.Caption:=' '+MsgLibel.Mess[14] ;
   END ;
If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
end;

procedure TFQRBLGESE.FJokerSectKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var CtrlF5 : Boolean ;
BEGIN
CtrlF5:=(Shift=[ssCtrl]) And (Key=VK_F5) ;
If CtrlF5 Then GereZoneJoker(Sender,'FJokerSect',FSection1) ;
  inherited;
end;

end.
