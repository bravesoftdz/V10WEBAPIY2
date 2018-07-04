unit QRBaAuGe;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  StdCtrls, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, UtilEdt, HEnt1, Ent1, QRRupt, CpteUtil,
{$IFNDEF CCMP}
  EdtLegal,
{$ENDIF}
  CritEDT, Menus, HSysMenu, Calcole, HTB97, HPanel, UiUtil, tCalcCum ;

procedure BalAuGe ;
procedure BalAuGeZoom(Crit : TCritEdt) ;

type
  TFBalAuGe = class(TFQR)
    HLabel3: THLabel;
    FGen1: THCpteEdit;
    FGen2: THCpteEdit;
    FGenJoker: TEdit;
    HLabel9: THLabel;
    FExcepGen: TEdit;
    HLabel10: THLabel;
    RGen: TQRLabel;
    RGen1: TQRLabel;
    TRaG: TQRLabel;
    RGen2: TQRLabel;
    QRLabel8: TQRLabel;
    TitreColCpt: TQRLabel;
    TitreColLibelle: TQRLabel;
    TTitreColAvant: TQRLabel;
    TLE_DEBIT: TQRLabel;
    TLE_CREDIT: TQRLabel;
    TTitreColSelection: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    TTitreColApres: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel21: TQRLabel;
    QRLabel16: TQRLabel;
    TTitreColSolde: TQRLabel;
    QRLabel25: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    REPORT3DEBIT: TQRLabel;
    REPORT3CREDIT: TQRLabel;
    REPORT4DEBIT: TQRLabel;
    REPORT4CREDIT: TQRLabel;
    HT_AUXILIAIRE: TQRDBText;
    HT_LIBELLE: TQRDBText;
    BDGene: TQRBand;
    CumDebitGen: TQRLabel;
    CumCreditGen: TQRLabel;
    DetDebitGen: TQRLabel;
    DetCreditGen: TQRLabel;
    SumDebitGen: TQRLabel;
    SumCreditGen: TQRLabel;
    SolDebitGen: TQRLabel;
    SolCreditGen: TQRLabel;
    GENERAL: TQRDBText;
    GLIBELLE: TQRDBText;
    BFCpteAux: TQRBand;
    CumDebitAux: TQRLabel;
    CumCreditAux: TQRLabel;
    SumDebitAux: TQRLabel;
    SumCreditAux: TQRLabel;
    SolDebitAux: TQRLabel;
    SolCreditAux: TQRLabel;
    LeSoldeD: TQRLabel;
    LeSoldeC: TQRLabel;
    QRLabel17: TQRLabel;
    DetDebitAux: TQRLabel;
    DetCreditAux: TQRLabel;
    FT_AUXILIAIRE: TQRDBText;
    FT_LIBELLE: TQRDBText;
    QRLabel33: TQRLabel;
    TotCumDebit: TQRLabel;
    TotCumCredit: TQRLabel;
    TotDetDebit: TQRLabel;
    TotDetCredit: TQRLabel;
    TotSumDebit: TQRLabel;
    TotSumCredit: TQRLabel;
    TotSolDebit: TQRLabel;
    TotSolCredit: TQRLabel;
    QRBand1: TQRBand;
    Trait6: TQRLigne;
    Trait0: TQRLigne;
    Trait5: TQRLigne;
    Trait4: TQRLigne;
    Trait3: TQRLigne;
    Trait2: TQRLigne;
    Ligne1: TQRLigne;
    REPORT5DEBIT: TQRLabel;
    REPORT5CREDIT: TQRLabel;
    REPORT6DEBIT: TQRLabel;
    REPORT6CREDIT: TQRLabel;
    REPORT7DEBIT: TQRLabel;
    REPORT7CREDIT: TQRLabel;
    REPORT8DEBIT: TQRLabel;
    REPORT8CREDIT: TQRLabel;
    SGen: TDataSource;
    QGen: TQuery;
    MsgLibel: THMsgBox;
    QRDLGene: TQRDetailLink;
    TFaS: TLabel;
    FCol1: TCheckBox;
    FCol2: TCheckBox;
    FCol3: TCheckBox;
    FCol4: TCheckBox;
    RExcepAux: TQRLabel;
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
    DLRupt: TQRDetailLink;
    FLigneRupt: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure BDGeneBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FGen1Change(Sender: TObject);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean;      var Quoi: string);
    procedure BFCpteAuxBeforePrint(var PrintBand: Boolean;      var Quoi: string);
    procedure BFCpteAuxAfterPrint(BandPrinted: Boolean);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean;      var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure QGenAfterOpen(DataSet: TDataSet);
    procedure FNatureCptChange(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure DLRuptNeedData(var MoreData: Boolean);
    procedure BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FRupturesClick(Sender: TObject);
    procedure FSansRuptClick(Sender: TObject);
  private    { Déclarations privées }
    Qr11T_AUXILIAIRE, Qr11T_LIBELLE,
    Qr11T_CORRESP1, Qr11T_CORRESP2 : TStringField ;
    Qr12G_GENERAL               : TStringField ;
    Qr12G_LIBELLE               : TStringField ;
    Affiche                     : Boolean ;
    StReportAux                 : string ;
    TotAuxi, TotEdt             : TabTot ;
    LRupt                      : TStringList ;
    Function  QuoiCpt(i : Integer) : String ;
    Procedure BalAuGeZoom(Quoi : String) ;
    procedure GenereSQLSub ;
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

procedure BalAuGe ;
var QR : TFBalAuGe ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:= TFBalAuGe.Create(Application) ;
Edition.Etat:=etBalAuxGen ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalAuGeZoom ;
QR.InitType (nbAux,neBal,msAuxEcr,'QRBAAUGE','',TRUE,FALSE,TRUE,Edition) ;
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
END;

procedure BalAuGeZoom(Crit : TCritEdt) ;
var QR : TFBalAuGe ;
    Edition : TEdition ;
BEGIN
QR:=TFBalAuGe.Create(Application) ;
Edition.Etat:=etBalAuxGen ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.BalAuGeZoom ;
 QR.CritEdt:=Crit ;
 QR.InitType (nbAux,neBal,msAuxEcr,'QRBAAUGE','',FALSE,TRUE,TRUE,Edition) ;
 finally
 QR.Free ;
 end ;
END ;

Function TFBalAuGe.QuoiCpt(i : Integer) : String ;
BEGIN
Case i Of
  0 : Result:=Qr11T_AUXILIAIRE.AsString+'#'+Qr11T_LIBELLE.AsString+'@'+'2' ;
  1 : Result:=Qr11T_AUXILIAIRE.AsString+'#'+Qr11T_LIBELLE.AsString+'@'+'2' ;
  2 : Result:=Qr12G_General.AsString+' '+Qr11T_AUXILIAIRE.AsString+'#'+GLIBELLE.Caption+'@'+'2' ;
 end ;
END ;

Procedure TFBalAuGe.BalAuGeZoom(Quoi : String) ;
Var Lp,i: Integer ;
BEGIN
Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
i:=StrToInt(Copy(Quoi,Lp+1,1)) ;
ZoomEdt(i,Quoi) ;
END ;

procedure TFBalAuGe.FinirPrint ;
begin
Inherited ;
QGen.Close ;
If GCalc<>NIL Then BEGIN GCalc.Free ; GCalc:=NIL ; END ;
if CritEdt.Rupture<>rRien then VideRupt(LRupt) ;
end;

procedure TFBalAuGe.GenereSQL ;
BEGIN
InHerited ;
AlimSQLMul(Q,0) ; Q.Close ; ChangeSQL(Q) ;
//Q.Prepare ;
PrepareSQLODBC(Q) ;
Q.Open ;
GenereSQLSub ;
END ;

procedure TFBalAuGe.GenereSQLSub ;
BEGIN
AlimSQLMul(QGen,1) ; ChangeSQL(QGen) ;
//QGen.Prepare ;
PrepareSQLODBC(QGen) ;
QGen.Open ;
END ;

procedure TFBalAuGe.RenseigneCritere;
BEGIN
Inherited ;
If OkZoomEdt Then Exit ;
if CritEdt.SJoker then
   BEGIN
   RGen.Caption:=MsgLibel.Mess[10] ;
   RGen1.Caption:=FGenJoker.Text ;
   END else
   BEGIN
   RGen.Caption:=MsgLibel.Mess[9] ;
   RGen1.Caption:=CritEdt.LSCpt1 ; RGen2.Caption:=CritEdt.LSCpt2 ;
   END ;
RGen2.Visible:=Not CritEdt.SJoker ; TRaG.Visible:=Not CritEdt.SJoker ;
RExcepGen.Caption:=FExcepGen.Text ;
RExcepAux.Caption:=FExcep.Text ;
END;

procedure TFBalAuGe.InitDivers ;
BEGIN
Inherited ;
{ Initialise les tableaux de montants }
Fillchar(TotAuxi,SizeOf(TotAuxi),#0) ;
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
{ Labels sur les bandes }
DateBandeauBalance(TTitreColAvant,TTitreColSelection,TTitreColApres,TTitreColSolde,CritEdt,
                   MsgLibel.Mess[0],MsgLibel.Mess[1],MsgLibel.Mess[2]) ;
ChangeFormatCompte(fbAux,fbGene,Self,HT_AUXILIAIRE.Width,1,2,HT_AUXILIAIRE.Name) ;
BFCpteAux.Frame.DrawBottom:=CritEDT.BAL.FormatPrint.PrSepCompte ;
BRupt.Frame.DrawTop:=CritEDT.BAL.FormatPrint.PrSepRupt  ;
BRupt.Frame.DrawBottom:=BRupt.Frame.DrawTop ;
END ;

procedure TFBalAuGe.ChoixEdition ;
{ Initialisation des options d'édition }
BEGIN
Inherited;
DLRupt.PrintBefore:=TRUE ;
Case CritEdt.Rupture of
  rLibres   : BEGIN
              DLRupt.PrintBefore:=FALSE ;
              ChargeGroup(LRupt,['T00','T01','T02','T03','T04','T05','T06','T07','T08','T09']) ;
              END ;
  rRuptures : BEGIN
              ChargeRupt(LRupt, 'RUT', CritEdt.Bal.PlanRupt,CritEdt.Bal.CodeRupt1,CritEdt.Bal.CodeRupt2) ;
              NiveauRupt(LRupt);
              END  ;
  rCorresp  : BEGIN
              ChargeRuptCorresp(LRupt, CritEdt.Bal.PlanRupt,CritEdt.Bal.CodeRupt1,CritEdt.Bal.CodeRupt2, False) ;
              NiveauRupt(LRupt);
              END  ;
  End ;
END ;

procedure TFBalAuGe.RecupCritEdt ;
Var St        : String ;
    NonLibres : Boolean ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  SJoker:=FGenJoker.Visible ; Composite:=True ;
  if SJoker Then BEGIN SCpt1:=FGenJoker.Text ; SCpt2:=FGenJoker.Text ; END
            Else BEGIN
                 SCpt1:=FGen1.Text ; SCpt2:=FGen2.Text ;
                 PositionneFourchetteSt(FGen1,FGen2,CritEdt.LSCpt1,CritEdt.LSCpt2) ;
                 END ;
  Bal.SSauf:=Trim(FExcepGen.Text) ; st:=Bal.SSauf ;
  If Bal.SSauf<>'' Then St:=' And '+AnalyseCompte(FExcepGen.Text,fbGene,True,False) ;
  Bal.SSqlSauf:=St ;
  BAL.RuptOnly:=QuelleTypeRupt(0,FSAnsRupt.Checked,FAvecRupt.Checked,FSurRupt.Checked) ;
  NonLibres:=((Rupture=rRuptures) or (Rupture=rCorresp)) ;
  if NonLibres then BAL.PlanRupt:=FPlanRuptures.Value ;
  if NonLibres then BEGIN BAL.CodeRupt1:=FCodeRupt1.Text ; BAL.CodeRupt2:=FCodeRupt2.Text ; END ;
  If (CritEdt.Rupture=rCorresp) Then Bal.PlansCorresp:=StrToInt(Copy(FPlanRuptures.Value,3,1)) ;
  Bal.OnlyCptAssocie:=(CritEdt.Rupture<>rRien) and FOnlyCptAssocie.Checked ;
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

function TFBalAuGe.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK and ColonnesOK;
If Result Then
   BEGIN
   Gcalc:=TGCalculCum.create(Deux,fbAux, fbGene,QuelTypeEcr,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE) ;
   GCalc.initCalcul('','','','',CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                    CritEdt.Date1,CritEdt.Bal.Date11,TRUE) ;
   END ;
END ;

Function TFBalAuGe.ColonnesOK : Boolean ;
BEGIN
If OkZoomEdt then begin Result:=True ; exit ; end ;
Result:=(FCol1.Checked or FCol2.Checked  or FCol3.Checked or FCol4.Checked) ;
If Not Result then NumErreurCrit:=7 ;
END ;

procedure TFBalAuGe.FormShow(Sender: TObject);
begin
HelpContext:=7460000 ;
//Standards.HelpContext:=7460010 ;
//Avances.HelpContext:=7460020 ;
//Mise.HelpContext:=7460030 ;
//Option.HelpContext:=7460040 ;
//TabRuptures.HelpContext:=7460050 ;
  inherited;
TabSup.TabVisible:=False;
FGen1.Text:='' ; FGen2.Text:='' ;
FOnlyCptAssocie.Checked:=False ; FOnlyCptAssocie.Enabled:=False ;
end;

procedure TFBalAuGe.BDGeneBeforePrint(var PrintBand: Boolean;  var Quoi: string);
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
   GCAlc.ReInitCalcul(Qr11T_AUXILIAIRE.AsString, Qr12G_GENERAL.AsString,CritEdt.Date1,CritEdt.Bal.Date11) ;
   GCalc.Calcul ; TotCpt1:=GCalc.ExecCalc.TotCpt ;

   GCAlc.ReInitCalcul(Qr11T_AUXILIAIRE.AsString, Qr12G_GENERAL.AsString,CritEdt.Bal.Date21,CritEdt.Date2) ;
   GCalc.Calcul ; TotCpt2:=GCalc.ExecCalc.TotCpt ;
   CumulVersSolde(TotCpt1[0]) ;
   if CritEdt.Date1=CritEdt.Bal.Date11 Then Fillchar(TotCpt1[1],SizeOf(TotCpt1[1]),#0) ;

   Tot[1].TotDebit:= Arrondi(TotCpt1[0].TotDebit+TotCpt1[1].TotDebit,CritEdt.Decimale) ;
   Tot[1].TotCredit:=Arrondi(TotCpt1[0].TotCredit+TotCpt1[1].TotCredit,CritEdt.Decimale) ;
   Tot[2].TotDebit:= TotCpt2[1].TotDebit ;
   Tot[2].TotCredit:=TotCpt2[1].TotCredit ;
   Tot[3].TotDebit:= Arrondi(Tot[2].TotDebit+Tot[1].TotDebit,CritEdt.Decimale) ;
   Tot[3].TotCredit:=Arrondi(Tot[2].TotCredit+Tot[1].TotCredit,CritEdt.Decimale) ;

   MReport[1].TotDebit:= Arrondi(TotCpt1[0].TotDebit+TotCpt1[1].TotDebit,CritEdt.Decimale) ;
   MReport[1].TotCredit:=Arrondi(TotCpt1[0].TotCredit+TotCpt1[1].TotCredit,CritEdt.Decimale) ;
   MReport[2].TotDebit:= TotCpt2[1].TotDebit ;
   MReport[2].TotCredit:=TotCpt2[1].TotCredit ;
   MReport[3].TotDebit:= Arrondi(Tot[2].TotDebit+Tot[1].TotDebit,CritEdt.Decimale) ;
   MReport[3].TotCredit:=Arrondi(Tot[2].TotCredit+Tot[1].TotCredit,CritEdt.Decimale) ;
   (* Rony 24/10/97 Passé au moule de CalcTotalEdt() en non Sur Rupture ou pas ...
   TotEdt[1].totDebit:= Arrondi(TotEdt[1].totDebit  + Tot[1].TotDebit, CritEdt.Decimale) ;
   TotEdt[1].totCredit:=Arrondi(TotEdt[1].totCredit + Tot[1].TotCredit,CritEdt.Decimale) ;
   TotEdt[2].totDebit:= Arrondi(TotEdt[2].totDebit  + Tot[2].TotDebit, CritEdt.Decimale) ;
   TotEdt[2].totCredit:=Arrondi(TotEdt[2].totCredit + Tot[2].TotCredit,CritEdt.Decimale) ;
   TotEdt[3].totDebit:= Arrondi(TotEdt[3].totDebit  + Tot[3].TotDebit, CritEdt.Decimale) ;
   TotEdt[3].totCredit:=Arrondi(TotEdt[3].totCredit + Tot[3].TotCredit,CritEdt.Decimale) ;
   *)
   TotAuxi[1].totDebit:= Arrondi(TotAuxi[1].totDebit  + Tot[1].TotDebit, CritEdt.Decimale) ;
   TotAuxi[1].totCredit:=Arrondi(TotAuxi[1].totCredit + Tot[1].TotCredit,CritEdt.Decimale) ;
   TotAuxi[2].totDebit:= Arrondi(TotAuxi[2].totDebit  + Tot[2].TotDebit, CritEdt.Decimale) ;
   TotAuxi[2].totCredit:=Arrondi(TotAuxi[2].totCredit + Tot[2].TotCredit,CritEdt.Decimale) ;
   TotAuxi[3].totDebit:= Arrondi(TotAuxi[3].totDebit  + Tot[3].TotDebit, CritEdt.Decimale) ;
   TotAuxi[3].totCredit:=Arrondi(TotAuxi[3].totCredit + Tot[3].TotCredit,CritEdt.Decimale) ;
   Tot[4]:=Tot[3] ;
   CumulVersSolde(Tot[4]) ;
   if Tot[4].TotDebit=0 then
      BEGIN
      MReport[4].TotCredit:=Tot[4].TotCredit ;
      TotAuxi[4].TotCredit:=Arrondi(TotAuxi[4].TotCredit+Tot[4].TotCredit,CritEdt.Decimale) ;
      SolDebitGen.Caption:='' ;
      SolCreditGen.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[4].TotCredit,CritEdt.AfficheSymbole) ;
      END else
      BEGIN
      MReport[4].TotDebit:=Tot[4].TotDebit ;
      TotAuxi[4].TotDebit:=Arrondi(TotAuxi[4].TotDebit+Tot[4].TotDebit,CritEdt.Decimale) ;
      SolDebitGen.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[4].TotDebit,CritEdt.AfficheSymbole ) ;
      SolCreditGen.Caption:='' ;
      END ;
   AddReportBAL([1,2], CritEdt.Bal.FormatPrint.Report, MReport, CritEdt.Decimale) ;
   Case CritEdt.Rupture of
     rLibres   : AddGroupLibre(LRupt,Q,fbAux,CritEdt.LibreTrie,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,0]) ;
     rRuptures : AddRupt(LRupt,Qr11T_AUXILIAIRE.AsString,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,0]) ;
     rCorresp  : BEGIN
                 Case CritEDT.Bal.PlansCorresp Of
                   1 : If Qr11T_CORRESP1.AsString<>'' Then CptRupt:=Qr11T_CORRESP1.AsString+Qr11T_AUXILIAIRE.AsString
                                                      Else CptRupt:='.'+Qr11T_AUXILIAIRE.AsString ;
                   2 : If Qr11T_CORRESP2.AsString<>'' Then CptRupt:=Qr11T_CORRESP2.AsString+Qr11T_AUXILIAIRE.AsString
                                                      Else CptRupt:='.'+Qr11T_AUXILIAIRE.AsString ;
                   Else CptRupt:=Qr11T_AUXILIAIRE.AsString ;
                   End ;
                 AddRuptCorres(LRupt,CptRupt,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,0]) ;
                 END ;
     End ;
   If CritEdt.Bal.RuptOnly=Sur then
      BEGIN
      PrintBand:=False ;
      IF CritEdt.Rupture=rLibres then CalcTotalEdt(Tot[1].TotDebit, Tot[1].TotCredit, Tot[2].TotDebit, Tot[2].TotCredit) ;
      END Else CalcTotalEdt(Tot[1].TotDebit, Tot[1].TotCredit, Tot[2].TotDebit, Tot[2].TotCredit) ;
   If PrintBand then Quoi:=Quoicpt(2) ;
   END ;
CumDebitGen.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[1].TotDebit, CritEdt.AfficheSymbole ) ;
CumCreditGen.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[1].TotCredit,CritEdt.AfficheSymbole ) ;
DetDebitGen.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[2].TotDebit, CritEdt.AfficheSymbole ) ;
DetCreditGen.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[2].TotCredit,CritEdt.AfficheSymbole ) ;
SumDebitGen.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[3].TotDebit, CritEdt.AfficheSymbole ) ;
SumCreditGen.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[3].TotCredit,CritEdt.AfficheSymbole ) ;
end;

procedure TFBalAuGe.FGen1Change(Sender: TObject);
Var GAvecJoker : Boolean ;
begin
  inherited;
GAvecJoker:=Joker(FGen1, FGen2, FGenJoker) ;
TFaS.Visible:=Not GAvecJoker ;
TFGen.Visible:=Not GAvecJoker ;
TFGenJoker.Visible:=GAvecJoker ;
end;

Procedure TFBalAuGe.CalcTotalEdt(AnVD, AnVC, D, C : Double );
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

procedure TFBalAuGe.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TotCumDebit.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[1].TotDebit, CritEdt.AfficheSymbole ) ;
TotCumCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[1].TotCredit,CritEdt.AfficheSymbole ) ;
TotDetDebit.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[2].TotDebit, CritEdt.AfficheSymbole ) ;
TotDetCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[2].TotCredit,CritEdt.AfficheSymbole ) ;
TotSumDebit.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[3].TotDebit, CritEdt.AfficheSymbole ) ;
TotSumCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[3].TotCredit,CritEdt.AfficheSymbole ) ;
TotSolDebit.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[4].TotDebit, CritEdt.AfficheSymbole ) ;
TotSolCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[4].TotCredit,CritEdt.AfficheSymbole ) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFBalAuGe.BFCpteAuxBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
if CritEdt.Bal.RuptOnly=Sur then PrintBand:=False Else PrintBand:=Affiche ;
if PrintBand then
   BEGIN
   Quoi:=Quoicpt(1) ;
   CumDebitAux.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotAuxi[1].TotDebit, CritEdt.AfficheSymbole ) ;
   CumCreditAux.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotAuxi[1].TotCredit,CritEdt.AfficheSymbole ) ;
   DetDebitAux.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotAuxi[2].TotDebit, CritEdt.AfficheSymbole ) ;
   DetCreditAux.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotAuxi[2].TotCredit,CritEdt.AfficheSymbole ) ;
   SumDebitAux.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotAuxi[3].TotDebit, CritEdt.AfficheSymbole ) ;
   SumCreditAux.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotAuxi[3].TotCredit,CritEdt.AfficheSymbole ) ;
   SolDebitAux.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotAuxi[4].TotDebit, CritEdt.AfficheSymbole ) ;
   SolCreditAux.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotAuxi[4].TotCredit,CritEdt.AfficheSymbole ) ;
   CumulVersSolde(TotAuxi[4]) ;
   if TotAuxi[4].TotDebit=0 then
      BEGIN
      LeSoldeD.Caption:='' ;
      LeSoldeC.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotAuxi[4].TotCredit,CritEdt.AfficheSymbole ) ;
      END else
      BEGIN
      LeSoldeD.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotAuxi[4].TotDebit,CritEdt.AfficheSymbole ) ;
      LeSoldeC.Caption:='' ;
      END ;
   END ;
end;

procedure TFBalAuGe.BFCpteAuxAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Fillchar(TotAuxi,SizeOf(TotAuxi),#0) ;
InitReport([2],CritEdt.Bal.FormatPrint.Report) ;
end;

procedure TFBalAuGe.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
Case CritEdt.Bal.TypCpt of  { 0 : Mvt sur L'exo,    3 : MVt sur la Période }
  0 :  PrintBand:=True ;
  3 :  PrintBand:= TRUE ;
 end ;
If Not PrintBand then Exit ;
StReportAux:=HT_AUXILIAIRE.Caption ;
InitReport([2],CritEdt.Bal.FormatPrint.Report) ;
Case CritEdt.Rupture of
  rLibres    : if CritEdt.Bal.OnlyCptAssocie then PrintBand:=DansRuptLibre(Q,fbAux,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
  rRuptures  : if CritEdt.Bal.OnlyCptAssocie then PrintBand:=DansRupt(LRupt,Qr11T_AUXILIAIRE.AsString) ;
  rCorresp   : if CritEdt.Bal.OnlyCptAssocie then
                  if CritEDT.Bal.PlansCorresp=1 then PrintBand:=DansRuptCorresp(LRupt,Qr11T_CORRESP1.AsString) Else
                  if CritEDT.Bal.PlansCorresp=2 then PrintBand:=DansRuptCorresp(LRupt,Qr11T_CORRESP2.AsString) ;
//                 if CritEDT.Bal.PlansCorresp=1 then PrintBand:=(Qr11T_CORRESP1.AsString<>'') Else
//                 if CritEDT.Bal.PlansCorresp=2 then PrintBand:=(Qr11T_CORRESP2.AsString<>'') ;
  End;
Affiche:=PrintBand ; { Pour ne pas afficher les sous comptes }
if CritEdt.Bal.RuptOnly=Sur then PrintBand:=False ;
If PrintBand then Quoi:=Quoicpt(0) ;
end;

procedure TFBalAuGe.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TITREREPORTH.Left:=TITREREPORTB.Left ;
TITREREPORTH.Width:=TITREREPORTB.Width ;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
Report1Debit.Caption :=Report5Debit.Caption ;
Report1Credit.Caption:=Report5Credit.Caption ;
Report2Debit.Caption :=Report6Debit.Caption ;
Report2Credit.Caption:=Report6Credit.Caption ;
Report3Debit.Caption :=Report7Debit.Caption ;
Report3Credit.Caption:=Report7Credit.Caption ;
Report4Debit.Caption :=Report8Debit.Caption ;
Report4Credit.Caption:=Report8Credit.Caption ;
end;

procedure TFBalAuGe.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var MReport : TabTRep ;
begin
  inherited;
TITREREPORTB.Left:=TitreColCpt.Left ;
TITREREPORTB.Width:=TitreColCpt.Width+TitreColLibelle.Width+1 ;
Case QuelReportBAL(CritEdt.Bal.FormatPrint.Report,MReport) of
  1 : TITREREPORTB.Caption:=MsgLibel.Mess[7] ;
  2 : TITREREPORTB.Caption:=MsgLibel.Mess[8]+' '+StReportAux ;
  END ;
Report5Debit.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[1].TotDebit,  CritEdt.AfficheSymbole ) ;
Report5Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[1].TotCredit, CritEdt.AfficheSymbole ) ;
Report6Debit.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[2].TotDebit,  CritEdt.AfficheSymbole ) ;
Report6Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[2].TotCredit, CritEdt.AfficheSymbole ) ;
Report7Debit.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[3].TotDebit,  CritEdt.AfficheSymbole ) ;
Report7Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[3].TotCredit, CritEdt.AfficheSymbole ) ;
Report8Debit.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[4].TotDebit,  CritEdt.AfficheSymbole ) ;
Report8Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[4].TotCredit, CritEdt.AfficheSymbole ) ;
end;

procedure TFBalAuGe.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr11T_AUXILIAIRE:=TStringField(Q.FindField('T_AUXILIAIRE'));
Qr11T_LIBELLE:=TStringField(Q.FindField('T_LIBELLE'));
If CritEDT.Rupture=rCorresp then
   BEGIN
   Qr11T_CORRESP1         :=TStringField(Q.FindField('T_CORRESP1'));
   Qr11T_CORRESP2         :=TStringField(Q.FindField('T_CORRESP2'));
   END ;
end;

procedure TFBalAuGe.QGenAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr12G_GENERAL         :=TStringField(QGen.FindField('G_GENERAL'));
Qr12G_LIBELLE         :=TStringField(QGen.FindField('G_LIBELLE'));
end;

procedure TFBalAuGe.FNatureCptChange(Sender: TObject);
begin
  inherited;
If Not QRLoading Then BEGIN FGen1.Clear ;   FGen2.Clear ;   FGenJoker.Clear ; END ;
Case FNatureCpt.ItemIndex of
  0     : BEGIN FGen1.ZoomTable:=tzGCollectif ; FGen2.ZoomTable:=tzGCollectif ; END ;
  1,2,4 : BEGIN FGen1.ZoomTable:=tzGCollDivers ; FGen2.ZoomTable:=tzGCollDivers ; END ;
  3     : BEGIN FGen1.ZoomTable:=tzGCollClient ; FGen2.ZoomTable:=tzGCollClient ; END ;
  5     : BEGIN FGen1.ZoomTable:=tzGCollFourn ; FGen2.ZoomTable:=tzGCollFourn ; END ;
  6     : BEGIN FGen1.ZoomTable:=tzGCollSalarie ; FGen2.ZoomTable:=tzGCollSalarie ; END ;
  End ;
end;

procedure TFBalAuGe.BNouvRechClick(Sender: TObject);
begin
  inherited;
If FGenJoker.Visible then FGenJoker.Text:='' ;
end;

procedure TFBalAuGe.DLRuptNeedData(var MoreData: Boolean);
Var TotRupt : Array[0..12] of Double ;
    CodRupt, LibRupt, Lib1 ,CptRupt, Stcode : String ;
    QuelleRupt : Integer ;
    SumD, SumC, TotSoldeD, TotSoldeC : Double ;
    Col : TColor ;
    OkOk, DansTotal, AddTotEdt : Boolean ;
    LibRuptInf : Array[1..10] Of TRuptInf ;
begin
  inherited;
MoreData:=False ;
TotSoldeD:=0 ; TotSoldeC:=0 ;
Case CritEdt.Rupture of
  rLibres   : BEGIN
              MoreData:=PrintGroupLibre(LRupt,Q,fbAux,CritEdt.LibreTrie,CodRupt,LibRupt,Lib1,TotRupt,Quellerupt,Col,LibRuptInf) ;
              If MoreData then
                 BEGIN
                 StCode:=CodRupt ;
                 Delete(StCode,1,Quellerupt+2) ;
                 MoreData:=DansChoixCodeLibre(StCode,Q,fbAux,CritEdt.LibreCodes1,CritEdt.LibreCodes2, CritEdt.LibreTrie) ;
                 BRupt.Font.Color:=Col ;
                 END ;
              END ;
  rRuptures : MoreData:=PrintRupt(LRupt,Qr11T_AUXILIAIRE.AsString,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt) ;
  rCorresp  : BEGIN
              OkOk:=TRUE ;
              Case CritEDT.Bal.PlansCorresp  Of
                1 : If Qr11T_CORRESP1.AsString<>'' Then CptRupt:=Qr11T_CORRESP1.AsString+Qr11T_AUXILIAIRE.AsString
                                                   Else CptRupt:='.'+Qr11T_AUXILIAIRE.AsString ;
                2 : If Qr11T_CORRESP2.AsString<>'' Then CptRupt:=Qr11T_CORRESP2.AsString+Qr11T_AUXILIAIRE.AsString
                                                   Else CptRupt:='.'+Qr11T_AUXILIAIRE.AsString ;
                Else OkOk:=FALSE ;
                END ;
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
      AddTotEdt:=False ;
      if (CritEdt.Rupture=rLibres) then AddTotEdt:=False else
      if (CritEdt.Rupture=rRuptures) or (CritEdt.Rupture=rCorresp) then AddTotEdt:=DansTotal ;
      if AddTotEdt then CalcTotalEdt(TotRupt[1], TotRupt[2], TotRupt[3], TotRupt[4]) ;
      END ;
   END ;
end;

procedure TFBalAuGe.BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
PrintBand:=CritEdt.Rupture<>rRien ;
end;

procedure TFBalAuGe.FRupturesClick(Sender: TObject);
begin
  inherited;
If FPlansCo.Checked then FGroupRuptures.Caption:=' '+MsgLibel.Mess[15] ;
If FRuptures.Checked then FGroupRuptures.Caption:=' '+MsgLibel.Mess[14] ;
end;

procedure TFBalAuGe.FSansRuptClick(Sender: TObject);
begin
  inherited;
FLigneRupt.Enabled:=Not FSansRupt.Checked ;
FLigneRupt.checked:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
FRupturesClick(Nil) ;
end;

end.
  
