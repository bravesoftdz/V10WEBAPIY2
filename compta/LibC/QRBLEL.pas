unit QRBLEL;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, StdCtrls, hmsgbox, HQuickrp, DB, DBTables, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, UtilEdt, HEnt1, Ent1, QRRupt, CpteUtil, EdtLegal, CritEDT,
  Menus, HSysMenu, CalcOle, UtilEdt1, HTB97, HPanel, UiUtil, tCalcCum ;

procedure BalanceLibreGene ;
procedure BalanceLibreGeneZoom(Crit : TCritEdt) ;

type
  TFBLEL = class(TFQR)
    FLigneRupt: TCheckBox;
    TitreColCpt: TQRLabel;
    TitreColLibelle: TQRLabel;
    TTitreColAvant: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel30: TQRLabel;
    TTitreColSelection: TQRLabel;
    QRLabel38: TQRLabel;
    QRLabel39: TQRLabel;
    TTitreColApres: TQRLabel;
    QRLabel41: TQRLabel;
    QRLabel42: TQRLabel;
    TTitreColSolde: TQRLabel;
    QRLabel44: TQRLabel;
    QRLabel45: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    REPORT3DEBIT: TQRLabel;
    REPORT3CREDIT: TQRLabel;
    REPORT4DEBIT: TQRLabel;
    REPORT4CREDIT: TQRLabel;
    NT_NATURE: TQRDBText;
    NT_LIBELLE: TQRDBText;
    CumDebit: TQRLabel;
    CumCredit: TQRLabel;
    DetDebit: TQRLabel;
    DetCredit: TQRLabel;
    SumDebit: TQRLabel;
    SumCredit: TQRLabel;
    SolDebit: TQRLabel;
    SolCredit: TQRLabel;
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
    MsgBox: THMsgBox;
    REPORT5DEBIT: TQRLabel;
    REPORT5CREDIT: TQRLabel;
    REPORT6DEBIT: TQRLabel;
    REPORT6CREDIT: TQRLabel;
    REPORT7DEBIT: TQRLabel;
    REPORT7CREDIT: TQRLabel;
    REPORT8DEBIT: TQRLabel;
    REPORT8CREDIT: TQRLabel;
    FCol1: TCheckBox;
    FCol2: TCheckBox;
    FCol3: TCheckBox;
    FCol4: TCheckBox;
    FCollectif: TCheckBox;
    FAvecAN: TCheckBox;
    FBilGes: TCheckBox;
    FOnlyClo: TCheckBox;
    TFGenJoker3: THLabel;
    TFGen3: THLabel;
    TFGen4: TLabel;
    FJoker3: TEdit;
    FCpte3: THCpteEdit;
    FCpte4: THCpteEdit;
    TitreDates: TLabel;
    HLabel12: THLabel;
    TFDateCpta2: TLabel;
    Label2: TLabel;
    HLabel13: THLabel;
    HLabel14: THLabel;
    Label3: TLabel;
    Label4: TLabel;
    HLabel15: THLabel;
    TFPeriodicite: TLabel;
    FP1: TMaskEdit;
    FP5: TMaskEdit;
    FP6: TMaskEdit;
    FP2: TMaskEdit;
    FP3: TMaskEdit;
    FP7: TMaskEdit;
    FP8: TMaskEdit;
    FP4: TMaskEdit;
    FPeriodicite: THValComboBox;
    HLabel3: THLabel;
    FExcept3: TEdit;
    CTypeCpte1: THValComboBox;
    HLabel9: THLabel;
    CTypeCpte3: THValComboBox;
    HLabel10: THLabel;
    Q1: TQuery;
    S1: TDataSource;
    BSecond: TQRBand;
    QRD: TQRDetailLink;
    BFDetail: TQRBand;
    CumDebit1: TQRLabel;
    CumCredit1: TQRLabel;
    DetDebit1: TQRLabel;
    DetCredit1: TQRLabel;
    SumDebit1: TQRLabel;
    SumCredit1: TQRLabel;
    SolCredit1: TQRLabel;
    SolDebit1: TQRLabel;
    NT_LIBELLE0: TQRDBText;
    NT_NATURE0: TQRDBText;
    NT_NATURE1: TQRLabel;
    RCpteB: TQRLabel;
    RCpte3: TQRLabel;
    TRa3: TQRLabel;
    RCpte4: TQRLabel;
    RExcepGen3: TQRLabel;
    QRLabel6: TQRLabel;
    Rjoker3: TQRLabel;
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FormShow(Sender: TObject);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure FExerciceChange(Sender: TObject);
    procedure FDateCompta1Change(Sender: TObject);
    procedure FCpte3Change(Sender: TObject);
    procedure FP1Exit(Sender: TObject);
    procedure TypeChange(Sender: TObject);
    procedure Q1AfterOpen(DataSet: TDataSet);
    procedure FPeriodiciteChange(Sender: TObject);
    procedure BFDetailBeforePrint(var PrintBand: Boolean;
      var Quoi: string);
    procedure BSecondBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure Q1BeforeOpen(DataSet: TDataSet);
  private { Déclarations privées }
    // Rony 10/03/97 * QBal, QBalC,QBalCum : TQuery ;
    QBalC       : TQuery ;
    TotEdt,TotNat        : TabTot ;
    QR11NT_NATURE        : TStringField ;
    QR11NT_LIBELLE       : TStringField ;
    QR12NT_NATURE        : TStringField ;
    QR12NT_LIBELLE       : TStringField ;
    FLoading, IsLegal : Boolean ;
    TabDate                        : TTabDate4 ;
    procedure GenereSQL ; Override ;
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
    Procedure BalLibGeneZoom(Quoi : String) ;
    Function  ColonnesOK : Boolean ;
    procedure ChargeComboCptes ;
    procedure SaisieDates ;
    procedure CalculDetail(var PrintBand : Boolean ; var Quoi : String) ;
  public  { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Uses Filtre ;

Procedure TFBLEL.BalLibGeneZoom(Quoi : String) ;
BEGIN
ZoomEdt(1,Quoi) ;
END ;

procedure BalanceLibreGene ;
var QR : TFBLEL ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:= TFBLEL.Create(Application) ;
Edition.Etat:=etBalLibEcr ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalLibGeneZoom ;
QR.IsLegal:=FALSE ;
QR.ChargeComboCptes ;
QR.InitType (nbLibEcr,neBalLib,msRien,'QRBLEL','',False,False,False,Edition) ;
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

procedure BalanceLibreGeneZoom(Crit : TCritEdt) ;
var QR : TFBLEL ;
    Edition : TEdition ;
BEGIN
QR:=TFBLEL.Create(Application) ;
Edition.Etat:=etBalLibEcr ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.BalLibGeneZoom ;
 QR.CritEdt:=Crit ;
 QR.IsLegal:=FALSE ;
 QR.ChargeComboCptes ;
 QR.InitType (nbLibEcr,neBalLib,msRien,'QRBLEL','',False,True,False,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFBLEL.GenereSQL ;
BEGIN
InHerited ;
AlimSQLMulNat(Q,0,FCpte1.Zoomtable,tzRien) ;
ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
Q.Open ;
if CTypeCpte3.Value<>'' then
  BEGIN
  AlimSQLMulNat(Q1,1,FCpte1.Zoomtable,FCpte3.ZoomTable) ;
  ChangeSQL(Q1) ; //Q1.Prepare ;
  PrepareSQLODBC(Q1) ;
  Q1.Open ;
  END ;
END ;

procedure TFBLEL.RenseigneCritere;
BEGIN
InHerited ;
If OkZoomEdt Then Exit ;
QRLabel7.Visible:=False ; RNatureEcr.Visible:=False ;
QRLabel12.Visible:=False ; RNatureCpt.Visible:=False ;
if CritEdt.SJoker then
   BEGIN
   //RAux.Caption:=MsgLibel.Mess[10] ;
   RCpte1.Caption:=FJoker.Text ;
   RCpte3.Caption:=FJoker3.Text ;
   END Else
   BEGIN
   //RAux.Caption:=MsgLibel.Mess[9] ;
   if CTypeCpte3.ItemIndex=0 then
     BEGIN
     RCpteB.Visible:=False ; RJoker3.Visible:=False ;
     RCpte3.Visible:=False ; RCpte4.Visible:=False ;
     TRa3.Visible:=False ;
     END else
     BEGIN
     RCpteB.Visible:=True ; RJoker3.Visible:=True ;
     RCpte3.Visible:=True ; RCpte4.Visible:=True ;
     TRa3.Visible:=True ;
     RCpte3.Caption:=CritEdt.LSCPT1 ; RCpte4.Caption:=CritEdt.LSCpt2 ;
     END ;
   END ;
FCpte4.Visible:=Not CritEdt.SJoker ;
TFGen4.Visible:=Not CritEdt.SJoker ;
RExcepGen3.Caption:=FExcept3.Text ;
END;

procedure TFBLEL.RecupCritEdt ;
var St : String ;
    i : integer ;
BEGIN
Inherited ;
With CritEDT Do
  BEGIN
  for i:=1 to 4 do BalLib.NbCol[i]:=TCheckBox(FindComponent('FCol'+IntToStr(i))).Checked ;
  SJoker:=FJoker3.Visible ;
  Composite:=(CTypeCpte3.Value<>'') ;
  if Composite then
    BEGIN
    CumDebit.Parent:=BSecond ;
    NT_NATURE.Font.Color:=BFDetail.Font.Color ;
    NT_NATURE.Font.Size:=BFDetail.Font.Size ;
    NT_NATURE.Font.Name:=BFDetail.Font.Name ;
    NT_NATURE.Font.Style:=BFDetail.Font.Style ;
    END else
    BEGIN
    CumDebit.Parent:=BDetail ;
    NT_NATURE.Font.Color:=clWindowText ;
    NT_NATURE.Font.Name:='Arial' ;
    NT_NATURE.Font.Size:=8 ;
    NT_NATURE.Font.Style:=NT_NATURE.Font.Style-[fsBold] ;
    END ;
  NT_LIBELLE.Font.Color:=NT_NATURE.Font.Color ;
  NT_LIBELLE.Font.Name:=NT_NATURE.Font.Name ;
  NT_LIBELLE.Font.Size:=NT_NATURE.Font.Size ;
  NT_LIBELLE.Font.Style:=NT_NATURE.Font.Style ;
  CumCredit.Parent:=CumDebit.Parent ;
  DetDebit.Parent:=CumDebit.Parent ; DetCredit.Parent:=CumDebit.Parent ;
  SumDebit.Parent:=CumDebit.Parent ; SumCredit.Parent:=CumDebit.Parent ;
  SolDebit.Parent:=CumDebit.Parent ; SolCredit.Parent:=CumDebit.Parent ;
  if SJoker Then BEGIN SCpt1:=FJoker3.Text ; SCpt2:=FJoker3.Text ; END
            Else BEGIN
                 SCpt1:=FCpte3.Text ; SCpt2:=FCpte4.Text ;
                 PositionneFourchetteSt(FCpte3,FCpte4,CritEdt.LSCpt1,CritEdt.LSCpt2) ;
                 END ;
  BalLib.SSauf:=Trim(FExcept3.Text) ; st:=BalLib.SSauf ;
  If BalLib.SSauf<>'' Then St:=' And '+AnalyseCompte(FExcept3.Text,fbGene,True,False) ;
  BalLib.SSqlSauf:=St ;
  With BalLib.FormatPrint Do
    BEGIN
    PutTabCol(1,TabColl[1],TitreColCpt.Tag,TRUE) ;
    PutTabCol(2,TabColl[2],TitreColLibelle.Tag,TRUE) ;
    PutTabCol(3,TabColl[3],TTitreColAvant.Tag,FCol1.Checked) ;
    PutTabCol(4,TabColl[4],TTitreColSelection.Tag,FCol2.Checked) ;
    PutTabCol(5,TabColl[5],TTitreColApres.Tag,FCol3.Checked) ;
    PutTabCol(6,TabColl[6],TTitreColSolde.Tag,FCol4.Checked) ;
    END ;
  END ;
END ;

Function TFBLEL.CritOk : Boolean ;
var TypCalc : TTypeCalc ;
BEGIN
Result:=Inherited CritOK ;
If Result And (Not OkZoomEdt) Then Result:=ColonnesOK;
if (CTypeCpte3.ItemIndex-1<=CTypeCpte1.ItemIndex) and (CTypeCpte3.ItemIndex<>0) then NumErreurCrit:=9 ;
Result:=Result and (CTypeCpte3.ItemIndex-1>CTypeCpte1.ItemIndex) or (CTypeCpte3.ItemIndex=0) ;
If Result Then
   BEGIN
//   QBal:=PrepareTotCpt(fbGene,QuelTypeEcr,Dev,Etab,Exo) ;
// rony 10/03/97 * QBal:=PrepareTotCptCum(QBalCum,fbGene,QuelTypeEcr,Dev,Etab,Exo) ;
   QBalC:=PrepareTotCptSolde(QuelTypeEcr,Dev,Etab,Exo,DevEnP) ;
   // Rony 10/07/97 Gcalc:=TGCalculCum.create(Un,fbGene,fbGene,QuelTypeEcr,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE) ;
(*
   if NET=neClo then
      BEGIN
      if FOnlyClo.Checked then Gcalc:=TGCalculCum.create(Un,fbGene,fbGene,[tpCloture],Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE)
                          Else Gcalc:=TGCalculCum.create(Un,fbGene,fbGene,[tpReel,tpCloture],Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE)
      END Else
*)
   TypCalc:=UnNatEcr ;
   if CritEdt.Composite then TypCalc:=DeuxNatEcr ;
   GCalc:=TGCalculCum.CreateNat(TypCalc,fbGene,fbGene,QuelTypeEcr,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE,FCpte1.ZoomTable,FCpte3.ZoomTable) ;
   GCalc.InitCalcul('','','','',CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                    CritEdt.Date1,CritEdt.BalLib.Date11,TRUE) ;
   END ;
InitReport([1],CritEDT.BaLLib.FormatPrint.Report) ;
END ;

Function TFBLEL.ColonnesOK : Boolean ;
BEGIN
If OkZoomEdt then begin Result:=True ; exit ; end ;
Result:=(FCol1.Checked or FCol2.Checked  or FCol3.Checked or FCol4.Checked) ;
If Not Result then NumErreurCrit:=7 ;
END ;

Procedure TFBLEL.FinirPrint ;
BEGIN
Inherited ;
If GCalc<>NIL Then BEGIN GCalc.Free ; GCalc:=NIL ; END ;
If QBalC<>NIL Then BEGIN QBalC.Free ; QBalC:=NIL ; END ;
END ;

procedure TFBLEL.CalculDetail(var PrintBand : Boolean ; var Quoi : String) ;
var TotCpt1,TotCpt2,tot,
    TotCpt1M,TotCpt1S,
    TotCpt2M,TotCpt2S   : TabTot ;
    MReport             : TabTRep ;
    i : integer ;
    F,FLib : TStringField ;
begin
Fillchar(Tot,SizeOf(Tot),#0) ; Fillchar(MReport,SizeOf(MReport),#0) ;
Fillchar(TotCpt1,SizeOf(TotCpt1),#0) ;
Fillchar(TotCpt2,SizeOf(TotCpt2),#0) ;
Fillchar(TotCpt1M,SizeOf(TotCpt1M),#0) ; Fillchar(TotCpt1S,SizeOf(TotCpt1S),#0) ;
Fillchar(TotCpt2M,SizeOf(TotCpt2M),#0) ; Fillchar(TotCpt2S,SizeOf(TotCpt2S),#0) ;
if True then
   BEGIN
   if not CritEdt.Composite then BEGIN F:=Qr11NT_NATURE ; FLib:=Qr11NT_LIBELLE ; END else BEGIN F:=Qr12NT_NATURE ; FLib:=Qr12NT_LIBELLE ; END ;
   Quoi:=F.AsString+' '+FLib.AsString ;
   for i:=0 to 3 do
     BEGIN
     if CritEdt.BalLib.NbCol[i+1] then
       BEGIN
       if CritEdt.Composite then GCalc.ReInitCalcul(Qr11NT_NATURE.AsString,Qr12NT_NATURE.AsString,StrToDate(TMaskEdit(FindComponent('FP'+IntToStr(i+1))).Text),StrToDate(TMaskEdit(FindComponent('FP'+IntToStr(5+i))).Text))
                            else GCalc.ReInitCalcul(Qr11NT_NATURE.AsString,'',StrToDate(TMaskEdit(FindComponent('FP'+IntToStr(i+1))).Text),StrToDate(TMaskEdit(FindComponent('FP'+IntToStr(5+i))).Text)) ;
       GCalc.Calcul ; TotCpt1:=GCalc.ExecCalc.TotCpt ;
       Tot[i].TotDebit:=Arrondi(TotCpt1[1].TotDebit,CritEDT.Decimale) ;
       Tot[i].TotCredit:=Arrondi(TotCpt1[1].TotCredit,CritEDT.Decimale) ;
       MReport[i+1].TotDebit:= Arrondi(Tot[i].TotDebit,CritEDT.Decimale) ;
       MReport[i+1].TotCredit:=Arrondi(Tot[i].TotCredit,CritEDT.Decimale) ;
       TotEdt[i].totDebit:= Arrondi(TotEdt[i].TotDebit + Tot[i].TotDebit,CritEDT.Decimale) ;
       TotEdt[i].totCredit:=Arrondi(TotEdt[i].TotCredit + Tot[i].TotCredit,CritEDT.Decimale) ;
       if CritEdt.Composite then
         BEGIN
         TotNat[i].totDebit:= Arrondi(TotNat[i].totDebit + Tot[i].TotDebit,CritEDT.Decimale) ;
         TotNat[i].totCredit:=Arrondi(TotNat[i].totCredit + Tot[i].TotCredit,CritEDT.Decimale) ;
         END ;
       case i of
         0: BEGIN
            CumDebit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,Tot[i].TotDebit,CritEDT.AfficheSymbole ) ;
            CumCredit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,Tot[i].TotCredit,CritEDT.AfficheSymbole ) ;
            END ;
         1: BEGIN
            DetDebit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,Tot[i].TotDebit,CritEDT.AfficheSymbole ) ;
            DetCredit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,Tot[i].TotCredit,CritEDT.AfficheSymbole ) ;
            END ;
         2: BEGIN
            SumDebit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,Tot[i].TotDebit,CritEDT.AfficheSymbole ) ;
            SumCredit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,Tot[i].TotCredit,CritEDT.AfficheSymbole ) ;
            END ;
         3: BEGIN
            SolDebit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,Tot[i].TotDebit,CritEDT.AfficheSymbole ) ;
            SolCredit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,Tot[i].TotCredit,CritEDT.AfficheSymbole ) ;
            END ;
         END ;
       END ;
     END ;
   AddReportBAL([1],CritEDT.BalLib.FormatPrint.Report,MReport,CritEDT.Decimale) ;
   END ;
END ;

procedure TFBLEL.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
inherited ;
if not CritEdt.Composite then CalculDetail(PrintBand,Quoi)
                         else Fillchar(TotNat,SizeOf(TotNat),#0) ;
end;

procedure TFBLEL.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TotCumDebit.Caption :=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotEdt[0].TotDebit,CritEDT.AfficheSymbole ) ;
TotCumCredit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotEdt[0].TotCredit,CritEDT.AfficheSymbole ) ;
TotDetDebit.Caption :=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotEdt[1].TotDebit,CritEDT.AfficheSymbole ) ;
TotDetCredit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotEdt[1].TotCredit,CritEDT.AfficheSymbole ) ;
TotSumDebit.Caption :=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotEdt[2].TotDebit,CritEDT.AfficheSymbole ) ;
TotSumCredit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotEdt[2].TotCredit,CritEDT.AfficheSymbole ) ;
TotSolDebit.Caption :=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotEdt[3].TotDebit,CritEDT.AfficheSymbole ) ;
TotSolCredit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotEdt[3].TotCredit,CritEDT.AfficheSymbole ) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFBLEL.InitDivers ;
BEGIN
InHerited ;
CalculDateTiers(neBalLib,Ventile,CritEdt,TabDate,FP1,FP2,FP3,FP4,FP5,FP6,FP7,FP8) ;
TTitreColAvant.Caption:=MsgBox.Mess[1]+' '+FP1.Text+MsgBox.Mess[0]+' '+FP5.Text ;
TTitreColSelection.Caption:=MsgBox.Mess[1]+' '+FP2.Text+MsgBox.Mess[0]+' '+FP6.Text ;
TTitreColApres.Caption:=MsgBox.Mess[1]+' '+FP3.Text+MsgBox.Mess[0]+' '+FP7.Text ;
TTitreColSolde.Caption:=MsgBox.Mess[1]+' '+FP4.Text+MsgBox.Mess[0]+' '+FP8.Text ;
{ Initialise les tableaux de montants }
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
{ Labels sur les bandes }
ChangeFormatCompte(fbGene,fbGene,Self,NT_NATURE.Width,1,2,NT_NATURE.Name) ;
END ;

procedure TFBLEL.TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
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

procedure TFBLEL.BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var MReport : TabTRep ;
begin
  inherited;
Case QuelReportBAL(CritEDT.BalLib.FormatPrint.Report,MReport) of
  1 : TITREREPORTB.Caption:=MsgBox.Mess[7] ;
 end ;
Report5Debit.Caption :=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, MReport[1].TotDebit,  CritEDT.AfficheSymbole ) ;
Report5Credit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, MReport[1].TotCredit, CritEDT.AfficheSymbole ) ;
Report6Debit.Caption :=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, MReport[2].TotDebit,  CritEDT.AfficheSymbole ) ;
Report6Credit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, MReport[2].TotCredit, CritEDT.AfficheSymbole ) ;
Report7Debit.Caption :=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, MReport[3].TotDebit,  CritEDT.AfficheSymbole ) ;
Report7Credit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, MReport[3].TotCredit, CritEDT.AfficheSymbole ) ;
Report8Debit.Caption :=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, MReport[4].TotDebit,  CritEDT.AfficheSymbole ) ;
Report8Credit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, MReport[4].TotCredit, CritEDT.AfficheSymbole ) ;
end;

procedure TFBLEL.FormShow(Sender: TObject);
begin
  inherited;
Floading:=FALSE ;
//TabSup.TabVisible:=False;
CTypeCpte1.Value:=CTypeCpte1.Values[0] ; CTypeCpte3.ItemIndex:=0 ; TypeChange(CTypeCpte3) ;
SaisieDates ;
FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
FCollectif.Checked:=FALSE ; FLigneRupt.Checked:=FALSE ;
// RL 17/10/97 If V_PGI.CergEF Then FGroupRuptures.Enabled:=FALSE ;
FGroupChoixRupt.Visible:=False ; FSansRuptClick(Nil);
TFGenJoker.Visible:=False ;
FCol1.Checked:=TRUE ; FCol2.Checked:=TRUE ; FCol3.Checked:=TRUE ; FCol4.Checked:=TRUE ;
//FJoker.MaxLength:=V_PGI.Cpta[fbGene].lg ;
{ JP 23/08/04 : FQ 14242 : Nouvelle gestion des filtres
 ChargeFiltrePourRupture ; }

FOnlyCptAssocie.Checked:=False ;
FOnlyClo.Visible:=False ; FBilGes.Enabled:=true ;

//TFGen.Enabled:=False ; TFGenJoker.Enabled:=False ; FJoker.Enabled:=False ;
//FCpte1.Enabled:=False ; FCpte2.Enabled:=False ;
//HLabel2.Enabled:=False ; FNatureCpt.Enabled:=False ;
//TSelectCpte.Enabled:=False  ; FSelectCpte.Enabled:=False ;
FSelectCpte.Value:='EXO' ;
{ Zone Dates (de, à) }
HLabel6.Enabled:=False ; Label7.Enabled:=False ;
FDateCompta1.Enabled:=False ; FDateCompta2.Enabled:=False ;
//HLabel5.Visible:=False ; FNatureEcr.Visible:=False ;
//HLabel1.Visible:=False ; FExcep.Visible:=False ;

{ Zone Ruptures }
FGroupChoixRupt.enabled:=False ;
FSansRupt.Enabled:=False ; FAvecRupt.Enabled:=False ; FSurRupt.Enabled:=False ;

AvecRevision.Visible:=False ; OnCum.Visible:=False ;

FCollectif.Checked:=False ; FCollectif.Enabled:=False ;
Avance.Checked:=False ; Avance.Visible:=False ;
FOnlyCptAssocie.Checked:=False ; FOnlyCptAssocie.Visible:=False ;
FGroupRuptures.Enabled:=False ;
QRP.ReportTitle:=Caption ;
end;

procedure TFBLEL.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr11NT_NATURE          :=TStringField(Q.FindField('NT_NATURE'));
Qr11NT_LIBELLE         :=TStringField(Q.FindField('NT_LIBELLE'));
end;

procedure TFBLEL.Q1AfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr12NT_NATURE          :=TStringField(Q1.FindField('NT_NATURE'));
Qr12NT_LIBELLE         :=TStringField(Q1.FindField('NT_LIBELLE'));
end;

procedure TFBLEL.FExerciceChange(Sender: TObject);
begin
FLoading:=TRUE ;
  inherited;
FLoading:=FALSE ;
end;

procedure TFBLEL.FDateCompta1Change(Sender: TObject);
begin
  inherited;
If FLoading Then Exit ;
end;

procedure TFBLEL.FCpte3Change(Sender: TObject);
begin
AvecJoker:=Joker(FCpte3, FCpte4, FJoker3) ;
TFGen3.Visible:=Not AvecJoker ;
TFGen4.Visible:=Not AvecJoker ;
TFGenJoker3.Visible:=AvecJoker ;
end;

procedure TFBLEL.FP1Exit(Sender: TObject);
Var St : String ;
    A  : Char   ;
begin
  inherited;
St:=TMaskEdit(Sender).Name ; A:=St[3] ;
if (TMaskEdit(Sender).Text<>'  /  /    ') and IsValidDate(TMaskEdit(Sender).Text)  then
   BEGIN
   FP2.Text:=DateToStr(StrToDate(FP5.Text)+1) ;
   FP3.Text:=DateToStr(StrToDate(FP6.Text)+1) ;
   FP4.Text:=DateToStr(StrToDate(FP7.Text)+1) ;
   FP7.Text:=DateToStr(StrToDate(FP4.Text)-1) ;
   If A='8' then FDateCompta1.Text:=DateToStr(StrToDate(FP8.Text)+1) ;
   END else
   BEGIN
   MsgRien.Execute(8,'',' '+TMaskEdit(Sender).Text) ;
   TMaskEdit(Sender).Text:=DateToStr(VH^.ExoV8.Fin);
   TMaskEdit(Sender).SetFocus ;
   END ;
end;

procedure TFBLEL.ChargeComboCptes ;
var Q : TQuery ;
BEGIN
CTypeCpte1.Items.CLear ; CTypeCpte1.Values.CLear ;
CTypeCpte3.Items.CLear ; CTypeCpte3.Values.CLear ;
CTypeCpte3.Values.Add('') ; CTypeCpte3.Items.Add('Aucun') ;
Q:=OpenSQL('Select CO_CODE,CO_LIBELLE From COMMUN Where CO_TYPE="NAT" And CO_CODE Like "E%" ORDER BY CO_CODE',True) ;
While not Q.Eof do
  BEGIN
  CTypeCpte1.Values.Add(Q.Fields[0].AsString) ; CTypeCpte1.Items.Add(Q.Fields[1].AsString) ;
  CTypeCpte3.Values.Add(Q.Fields[0].AsString) ; CTypeCpte3.Items.Add(Q.Fields[1].AsString) ;
  Q.Next ;
  END ;
Ferme(Q) ;

END ;

procedure TFBLEL.TypeChange(Sender: TObject);
var Num : byte ;
    Ok : boolean ;
    T : THValCombobox ;
    L1,L2 : TLabel ;
    C1,C2 : THCpteEdit ;
begin
T:=THValComboBox(Sender) ;
if T=CTypeCpte1 then
  BEGIN
  L1:=TFGen ; L2:=TFAG ;
  C1:=FCpte1 ; C2:=FCpte2 ;
  END else
  BEGIN
  L1:=TFGen3 ; L2:=TFGen4 ;
  C1:=FCpte3 ; C2:=FCpte4 ;
  END ;
Ok:=(T.Value='') ;
C1.ZoomTable:=tzRien ;
L1.Enabled:=Not Ok ; C1.Enabled:=Not Ok ;
L2.Enabled:=Not Ok ; C2.Enabled:=Not Ok ;
if not Ok then
  BEGIN
  Num:=StrToInt(Copy(T.Value,Length(T.Value),1)) ;
  //L1.Caption:=MsgQR.Mess[6]+' '+IntToStr(Num+1) ;
  Case Num Of
    0 : C1.ZoomTable:=tzNatEcrE0 ;
    1 : C1.ZoomTable:=tzNatEcrE1 ;
    2 : C1.ZoomTable:=tzNatEcrE2 ;
    3 : C1.ZoomTable:=tzNatEcrE3 ;
    END ;
  END ; //else L1.Caption:=MsgQR.Mess[6]+'   ' ;

C2.ZoomTable:=C1.ZoomTable ;
END ;

procedure TFBLEL.SaisieDates ;
Var i : byte ;
BEGIN
For i:=1 to 4 do
  BEGIN
  TMaskEdit(FindComponent('FP'+InttoStr(i))).Text:=StDate1900 ;
  TMaskEdit(FindComponent('FP'+InttoStr(i))).Enabled:=False ;
  END ;
For i:=5 to 8 do
  BEGIN
  TMaskEdit(FindComponent('FP'+InttoStr(i))).Enabled:=True ;
  TMaskEdit(FindComponent('FP'+InttoStr(i))).Text:=stDate2099 ;
  END ;
FP1.Text:=DateToStr(StrToDate(FDateCompta1.Text)) ;
FPeriodicite.Value:=FPeriodicite.Values[0] ;
END ;

procedure TFBLEL.FPeriodiciteChange(Sender: TObject);
Var i : Integer ;
    TabD : TTabDate8 ;
begin
  inherited;
If QRLoading then Exit ;
PeriodiciteChangeTiers(neBalLib,Ventile,FPeriodicite.ItemIndex,FP8,FP1,False,TabD) ;
For i:=1 to 8 do TMaskEdit(FindComponent('FP'+InttoStr(i))).text:=DateToStr(TabD[i] ) ;
end;

procedure TFBLEL.BFDetailBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
inherited ;
if not CritEdt.Composite then BEGIN PrintBand:=False ; Exit ; END ;
NT_NATURE1.Caption:=MsgBox.Mess[22]+' '+QR11NT_NATURE.AsString+' '+QR11NT_LIBELLE.AsString ;
CumDebit1.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotNat[0].TotDebit,CritEDT.AfficheSymbole ) ;
CumCredit1.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotNat[0].TotCredit,CritEDT.AfficheSymbole ) ;
DetDebit1.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotNat[1].TotDebit,CritEDT.AfficheSymbole ) ;
DetCredit1.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotNat[1].TotCredit,CritEDT.AfficheSymbole ) ;
SumDebit1.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotNat[2].TotDebit,CritEDT.AfficheSymbole ) ;
SumCredit1.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotNat[2].TotCredit,CritEDT.AfficheSymbole ) ;
SolDebit1.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotNat[3].TotDebit,CritEDT.AfficheSymbole ) ;
SolCredit1.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotNat[3].TotCredit,CritEDT.AfficheSymbole ) ;
end;

procedure TFBLEL.BSecondBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
if CritEdt.Composite and not Q1.Findfield('NT_NATURE').IsNull then CalculDetail(PrintBand,Quoi) else PrintBand:=False ;
end;

procedure TFBLEL.Q1BeforeOpen(DataSet: TDataSet);
begin
  inherited;
Q1.Params[0].AsString:=Q.FindField('NT_NATURE').AsString ;
end;

end.
