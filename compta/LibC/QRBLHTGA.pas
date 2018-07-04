unit QRBLHTGA;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, HSysMenu, Menus, hmsgbox, HQuickrp, DB, DBTables, StdCtrls, Buttons,
  Hctrls, ExtCtrls, Mask, Hcompte, ComCtrls,CritEdt, Ent1, HEnt1,CpteUtil,
  QRRupt,CalcOle, UtilEdt, HTB97, HPanel, UiUtil ;

type
  TFQRBLHTGA = class(TFQR)
    TFAux: THLabel;
    FAuxJoker: TEdit;
    TFaA: TLabel;
    FAux2: THCpteEdit;
    HLabel3: THLabel;
    FNatCpteAux: THValComboBox;
    FExcepAux: TEdit;
    HLabel9: THLabel;
    FCol1: TCheckBox;
    FCol2: TCheckBox;
    FCol3: TCheckBox;
    FCol4: TCheckBox;
    FLigneRupt: TCheckBox;
    RAux: TQRLabel;
    RAux1: TQRLabel;
    TRaA: TQRLabel;
    RAux2: TQRLabel;
    QAux: TQuery;
    SAux: TDataSource;
    HG_GENERAL: TQRDBText;
    HG_LIBELLE: TQRDBText;
    BDAux: TQRBand;
    AUXILIAIRE: TQRDBText;
    TLIBELLE: TQRDBText;
    CumDebitAux: TQRLabel;
    CumCreditAux: TQRLabel;
    DetDebitAux: TQRLabel;
    DetCreditAux: TQRLabel;
    SumDebitAux: TQRLabel;
    SumCreditAux: TQRLabel;
    SolDebitAux: TQRLabel;
    SolCreditAux: TQRLabel;
    QRLabel1: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    REPORT3DEBIT: TQRLabel;
    REPORT3CREDIT: TQRLabel;
    REPORT4DEBIT: TQRLabel;
    REPORT4CREDIT: TQRLabel;
    BFCpteGene: TQRBand;
    LeSoldeD: TQRLabel;
    LeSoldeC: TQRLabel;
    QRLabel33: TQRLabel;
    FG_GENERAL: TQRDBText;
    FG_LIBELLE: TQRDBText;
    CumDebitGen: TQRLabel;
    CumCreditGen: TQRLabel;
    DetDebitGen: TQRLabel;
    DetCreditGen: TQRLabel;
    SumDebitGen: TQRLabel;
    SumCreditGen: TQRLabel;
    SolDebitGen: TQRLabel;
    SolCreditGen: TQRLabel;
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
    TitreColCpt: TQRLabel;
    TitreColLibelle: TQRLabel;
    TTitreColAvant: TQRLabel;
    TLE_DEBIT: TQRLabel;
    TLE_CREDIT: TQRLabel;
    TTitreColSelection: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    TTitreColApres: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel21: TQRLabel;
    TTitreColSolde: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel30: TQRLabel;
    MsgLibel: THMsgBox;
    RExcepAux: TQRLabel;
    QRLabel36: TQRLabel;
    DLRupt: TQRDetailLink;
    FAux1: THCpteEdit;
    QRLabel38: TQRLabel;
    TotCumDebit: TQRLabel;
    TotCumCredit: TQRLabel;
    TotDetDebit: TQRLabel;
    TotDetCredit: TQRLabel;
    TotSumDebit: TQRLabel;
    TotSumCredit: TQRLabel;
    TotSolDebit: TQRLabel;
    TotSolCredit: TQRLabel;
    TFAuxJoker: THLabel;
    QRLabel2: TQRLabel;
    REPORT5DEBIT: TQRLabel;
    REPORT5CREDIT: TQRLabel;
    REPORT6DEBIT: TQRLabel;
    REPORT6CREDIT: TQRLabel;
    REPORT7CREDIT: TQRLabel;
    REPORT7DEBIT: TQRLabel;
    REPORT8DEBIT: TQRLabel;
    REPORT8CREDIT: TQRLabel;
    QRDLAux: TQRDetailLink;
    procedure FormShow(Sender: TObject);
    procedure BDAuxBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean;
      var Quoi: string);
    procedure BFCpteGeneBeforePrint(var PrintBand: Boolean;
      var Quoi: string);
    procedure BFCpteGeneAfterPrint(BandPrinted: Boolean);
    procedure FAux1Change(Sender: TObject);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean;
      var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean;
      var Quoi: string);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure QAuxAfterOpen(DataSet: TDataSet);
    procedure FNatureCptChange(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure DLRuptNeedData(var MoreData: Boolean);
    procedure BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FSansRuptClick(Sender: TObject);
    procedure FPlanRupturesChange(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure FNatCpteAuxChange(Sender: TObject);
  private
     Qr11G_GENERAL, Qr11G_LIBELLE,
    Qr11G_COLLECTIF              : TStringField;
    Qr11G_CORRESP1, Qr11G_CORRESP2  : TStringField;
    Qr12T_AUXILIAIRE, Qr12T_LIBELLE : TStringField;
    Affiche               : Boolean ;
    StReportGen           : string ;
    TotCptGen, TotEdt     : TabTot ;
    LRupt                 : TStringList ;
    procedure GenereSQL ; Override ;
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
    Function  QuoiCpt(i : Integer) : String ;
    procedure GenereSQLSub ;
    Function  ColonnesOK : Boolean ;
    Procedure CalcTotalEdt(AnVD, AnVC, D, C : Double );
  public
    { Déclarations publiques }
  end;

procedure BalGeAuHT ;

implementation

{$R *.DFM}

procedure BalGeAuHT ;
var QR : TFQRBLHTGA ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
QR:= TFQRBLHTGA.Create(Application) ;
Edition.Etat:=etBalHTGA ;
QR.InitType (nbGen,neBal,msGenEcr,'QRBLHTGA','',TRUE,FALSE,TRUE,Edition) ;
PP:=FindInsidePanel ;
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

Function TFQRBLHTGA.QuoiCpt(i : Integer) : String ;
BEGIN
Inherited ;
Case i Of
  0 : Result:=Qr11G_GENERAL.AsString+'#'+Qr11G_LIBELLE.AsString+'@'+'1' ;
  1 : Result:=Qr11G_General.AsString+'#'+Qr11G_LIBELLE.AsString+'@'+'1' ;
  2 : Result:=Qr12T_AUXILIAIRE.AsString+' '+Qr11G_GENERAL.AsString+'#'+TLIBELLE.Caption+'@'+'2' ;
 end ;
END ;

procedure TFQRBLHTGA.FinirPrint ;
begin
Inherited ;
QAux.Close ; GCalc.Free ;
if CritEdt.Rupture<>rRien then VideRupt(LRupt) ;
end;

procedure TFQRBLHTGA.GenereSQL ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
AlimSQLMul(Q,0) ; Q.Close ; ChangeSQL(Q) ; Q.Prepare ; Q.Open ;
GenereSQLSub ;
END ;

procedure TFQRBLHTGA.GenereSQLSub ;
var St,St2 : String ;
    StTable,Where1,PreB1,PreB2,PreE,PreJ1,PreJ2,Table1,Table2 : String ;
    SetType : SetttTypePiece ;
BEGIN
Inherited ;
// Genre WhatExisteMul ; TModeSelect : msAuxEcr :
PreJ1:='GENERAL' ; PreJ2:='AUXILIAIRE' ; PreB1:='G_' ; PreB2:='T_' ; PreE:='E_' ;
Table1:='GENERAUX' ; Table2:='TIERS' ;
if PreE='Y_' then StTable:='ANALYTIQ' else If PreE='BE_' Then StTable:='BUDECR' Else StTable:='ECRITURE' ;
Where1:=':'+PreB1+PreJ1 ;
QAux.SQL.Clear ;
QAux.SQL.Add('SELECT '+PreB2+PreJ2+','+PreB2+'LIBELLE FROM '+Table2+' Q1 ') ;
QAux.SQL.Add(' WHERE ') ;
St:=' (Exists (Select '+PreE+PreJ1+','+PreE+PreJ2+','+PreE+'EXERCICE,'+PreE+'DATECOMPTABLE From '+StTable+' ' ;
//St:=St+' Where ('+PreE+PreJ1+'='+Where1+' And '+PreE+PreJ2+'=Q1.'+Where2+' ' ;
St:=St+' Where '+PreE+'EXERCICE="'+CritEDT.Exo.Code+'" ';
St:=St+' And '+PreE+'DATECOMPTABLE>="'+UsDateTime(CritEdt.Date1)+'" And '+PreE+'DATECOMPTABLE<="'+UsDateTime(CritEDT.Date2)+'" ' ;
SetType:=WhatTypeEcr(CritEdt.QualifPiece,V_PGI.Controleur,AvecRevision.State) ;
St2:=WhereSupp(PreE,SetType) ;
If St2<>'' Then St:=St+St2 ;
if CritEDT.SJoker then St:=St+'AND '+PreE+PreJ2+' like "'+TraduitJoker(CritEdt.SCpt1)+'" ' Else
  BEGIN
  if CritEDT.SCpt1<>'' then St:=St+' And '+PreE+PreJ2+'>="'+CritEdt.SCpt1+'" ';
  if CritEDT.SCpt2<>'' then St:=St+' And '+PreE+PreJ2+'<="'+CritEdt.SCpt2+'" ' ;
  END ;
St:=St+' And ('+PreE+'CONTREPARTIEGEN'+'='+Where1+' And '+PreE+PreJ2+'=Q1.'+PreB2+PreJ2+' ' ;
St:=St+')))' ;
QAux.SQL.Add(St) ;
St:='' ;
if CritEDT.SJoker then St:=St+'AND '+PreB2+PreJ2+' like "'+TraduitJoker(CritEdt.SCpt1)+'" ' Else
  BEGIN
  if CritEDT.SCpt1<>'' then St:=St+' And '+PreB2+PreJ2+'>="'+CritEdt.SCpt1+'" ';
  if CritEDT.SCpt2<>'' then St:=St+' And '+PreB2+PreJ2+'<="'+CritEdt.SCpt2+'" ' ;
  END ;
if FNatCpteAux.Value<>'' then St:=St+' And '+PreB2+'NATUREAUXI="'+FNatCpteAux.Value+'" ' ;
if St<>'' then QAux.SQL.Add(St) ;
QAux.SQL.Add(' ORDER BY '+PreB2+PreJ2) ;
// Fin genre WhatExisteMul
ChangeSQL(QAux) ; QAux.Prepare ; QAux.Open ;
END ;

procedure TFQRBLHTGA.RenseigneCritere;
BEGIN
Inherited ;
If OkZoomEdt Then Exit ;
if CritEdt.SJoker then
   BEGIN
   RAux.Caption:=MsgLibel.Mess[10] ;
   RAux1.Caption:=FAuxJoker.Text
   END Else
   BEGIN
   RAux.Caption:=MsgLibel.Mess[9] ;
   RAux1.Caption:=CritEdt.LSCPT1 ; RAux2.Caption:=CritEdt.LSCpt2 ;
   END ;
RAux2.Visible:=Not CritEdt.SJoker ;
TRaA.Visible:=Not CritEdt.SJoker ;

RExcepAux.Caption:=FExcepAux.Text ;
(* Rony 17/03/97
CaseACocher(FCol1,RCol1)            ; CaseACocher(FCol2,RCol2) ;
CaseACocher(FCol3 ,RCol3)           ; CaseACocher(FCol4,RCol4) ;
CaseACocher(FCollectif,RCollectif) ; *)
END;

procedure TFQRBLHTGA.ChoixEdition ;
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
              END ;
   rCorresp : BEGIN
              ChargeRuptCorresp(LRupt, CritEdt.Bal.PlanRupt,CritEdt.Bal.CodeRupt1,CritEdt.Bal.CodeRupt2, False) ;
              NiveauRupt(LRupt);
              END ;
  End ;
END ;

procedure TFQRBLHTGA.InitDivers ;
BEGIN
Inherited ;
{ Initialise les tableaux de montants }
Fillchar(TotCptGen,SizeOf(TotCptGen),#0) ;
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
{ Labels sur les bandes }
DateBandeauBalance(TTitreColAvant,TTitreColSelection,TTitreColApres,TTitreColSolde,CritEdt,
                   MsgLibel.Mess[0],MsgLibel.Mess[1],MsgLibel.Mess[2]) ;
ChangeFormatCompte(fbGene,fbAux,Self,HG_GENERAL.Width,1,2,HG_GENERAL.Name) ;
BFCpteGene.Frame.DrawBottom:=CritEDT.BAL.FormatPrint.PrSepCompte ;
BRupt.Frame.DrawTop:=CritEDT.BAL.FormatPrint.PrSepRupt  ;
BRupt.Frame.DrawBottom:=BRupt.Frame.DrawTop ;
END ;

procedure ModifContrePartie(var Q : TQuery) ;
var i,p : integer ;
    St : string ;
BEGIN
for i:=0 to Q.SQL.Count-1 do
  BEGIN
  St:=Q.SQL[i] ;
  p:=Pos('GENERAL=:CPTE1 ',St) ;
  if p<>0 then St:=Copy(St,1,p-1)+'CONTREPARTIEGEN=:CPTE1 '+Copy(St,p+15,Length(St)-(p+14)) ;
  Q.SQL[i]:=St ;
  END ;
// Pas de ChangeSQL car déja fait dans TGCalculCum.Create ;
Q.Prepare ;
END ;

procedure TFQRBLHTGA.RecupCritEdt ;
var st : String ;
    NonLibres : Boolean ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  SJoker:=FAuxJoker.Visible ; Composite:=True ;
  if SJoker Then BEGIN SCpt1:=FAuxJoker.Text ; SCpt2:=FAuxJoker.Text ; END
            Else BEGIN
                 SCpt1:=FAux1.Text ; SCpt2:=FAux2.Text ;
                 PositionneFourchetteST(FAux1,FAux2,LSCpt1,LSCpt2) ;
                 END ;
  Bal.SSauf:=Trim(FExcepAux.Text) ; st:=Bal.SSauf ;
  If Bal.SSauf<>'' Then St:=' And '+AnalyseCompte(FExcepAux.Text,fbAux,True,False) ;
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

Function TFQRBLHTGA.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK and ColonnesOK ;
If Result Then
   BEGIN
   Gcalc:=TGCalculCum.create(Deux,fbGene,fbAux,QuelTypeEcr,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE) ;
   ModifContrePartie(GCalc.Q) ;
   GCalc.initCalcul('','','','',CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                    CritEdt.Date1,CritEdt.Bal.Date11,TRUE) ;
   END ;
END ;

Function TFQRBLHTGA.ColonnesOK : Boolean ;
BEGIN
If OkZoomEdt then begin Result:=True ; exit ; end ;
Result:=(FCol1.Checked or FCol2.Checked  or FCol3.Checked or FCol4.Checked) ;
If Not Result then NumErreurCrit:=7 ;
END ;

procedure TFQRBLHTGA.FormShow(Sender: TObject);
begin
{
HelpContext:=7457010 ;
Standards.HelpContext:=7457010 ;
Avances.HelpContext:=7457020 ;
Mise.HelpContext:=7457030 ;
Option.HelpContext:=7457040 ;
}
FAux1.Text:='' ; FAux2.Text:='' ;
  inherited;
FNatureCpt.DataType:='ttNatJalBud' ;
TabSup.TabVisible:=False;
FOnlyCptAssocie.Checked:=False ; FOnlyCptAssocie.Enabled:=False ;
FCol1.Checked:=False ; FCol1.Enabled:=False ;
FCol2.Checked:=False ; FCol2.Enabled:=False ;
end;

procedure TFQRBLHTGA.BDAuxBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var TotCpt1, TotCpt2, Tot : TabTot ;
    MReport               : TabTRep ;
    CptRupt : String ;
begin
  inherited;
Fillchar(Tot,SizeOf(Tot),#0) ; Fillchar(TotCpt1,SizeOf(TotCpt1),#0) ; Fillchar(TotCpt2,SizeOf(TotCpt2),#0) ;
Fillchar(Mreport,SizeOf(MReport),#0) ;
PrintBand:=Affiche ;
if PrintBand then
   BEGIN
   (*
   ExecuteTotCptM(QBal,QG_GENERAL.AsString, QAuxT_AUXILIAIRE.AsString,CritEdt.Date1,CritEdt.Bal.Date11,
                 CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,TotCpt1,FALSE) ;
   ExecuteTotCptM(QBal,QG_GENERAL.AsString, QAuxT_AUXILIAIRE.AsString,CritEdt.Bal.Date21,CritEdt.Date2,
                 CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,TotCpt2,FALSE) ;
   ExecuteTotCptMCum(QBal,QBalCum,QG_GENERAL.AsString, QAuxT_AUXILIAIRE.AsString,'','',CritEdt.Date1,CritEdt.Bal.Date11,
                 CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,TotCpt1,FALSE) ;
   ExecuteTotCptMCum(QBal,QBalCum,QG_GENERAL.AsString, QAuxT_AUXILIAIRE.AsString,'','',CritEdt.Bal.Date21,CritEdt.Date2,
                 CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,TotCpt2,FALSE) ;
   *)
   GCAlc.ReInitCalcul(Qr11G_GENERAL.AsString, Qr12T_AUXILIAIRE.AsString,CritEdt.Date1,CritEdt.Bal.Date11) ;
   GCalc.Calcul ; TotCpt1:=GCalc.ExecCalc.TotCpt ;

   GCAlc.ReInitCalcul(Qr11G_GENERAL.AsString, Qr12T_AUXILIAIRE.AsString,CritEdt.Bal.Date21,CritEdt.Date2) ;
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
   (**  Rony 24/10/97 Passé au moule de CalcTotalEdt() en non Sur Rupture ou pas ...
   TotEdt[1].totDebit:=Arrondi(TotEdt[1].totDebit + Tot[1].TotDebit,CritEdt.Decimale) ;
   TotEdt[1].totCredit:=Arrondi(TotEdt[1].totCredit + Tot[1].TotCredit,CritEdt.Decimale) ;
   TotEdt[2].totDebit:=Arrondi(TotEdt[2].totDebit + Tot[2].TotDebit,CritEdt.Decimale) ;
   TotEdt[2].totCredit:=Arrondi(TotEdt[2].totCredit + Tot[2].TotCredit,CritEdt.Decimale) ;
   TotEdt[3].totDebit:=Arrondi(TotEdt[3].totDebit + Tot[3].TotDebit,CritEdt.Decimale) ;
   TotEdt[3].totCredit:=Arrondi(TotEdt[3].totCredit + Tot[3].TotCredit,CritEdt.Decimale) ;
   **)
   TotCptGen[1].totDebit:=Arrondi(TotCptGen[1].totDebit + Tot[1].TotDebit,CritEdt.Decimale) ;
   TotCptGen[1].totCredit:=Arrondi(TotCptGen[1].totCredit + Tot[1].TotCredit,CritEdt.Decimale) ;
   TotCptGen[2].totDebit:=Arrondi(TotCptGen[2].totDebit + Tot[2].TotDebit,CritEdt.Decimale) ;
   TotCptGen[2].totCredit:=Arrondi(TotCptGen[2].totCredit + Tot[2].TotCredit,CritEdt.Decimale) ;
   TotCptGen[3].totDebit:=Arrondi(TotCptGen[3].totDebit + Tot[3].TotDebit,CritEdt.Decimale) ;
   TotCptGen[3].totCredit:=Arrondi(TotCptGen[3].totCredit + Tot[3].TotCredit,CritEdt.Decimale) ;
   Tot[4]:=Tot[3] ;
   CumulVersSolde(Tot[4]) ;
   if Tot[4].TotDebit=0 then
      BEGIN
      MReport[4].TotCredit:=Tot[4].TotCredit ;
      TotCptGen[4].TotCredit:=Arrondi(TotCptGen[4].TotCredit+Tot[4].TotCredit,CritEdt.Decimale) ;
      SolDebitAux.Caption:='' ;
      SolCreditAux.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[4].TotCredit,CritEdt.AfficheSymbole) ;
      END else
      BEGIN
      MReport[4].TotDebit:=Tot[4].TotDebit ;
      TotCptGen[4].TotDebit:=Arrondi(TotCptGen[4].TotDebit+Tot[4].TotDebit,CritEdt.Decimale) ;
      SolDebitAux.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[4].TotDebit,CritEdt.AfficheSymbole ) ;
      SolCreditAux.Caption:='' ;
      END ;
   AddReportBAL([1,2], CritEdt.Bal.FormatPrint.Report, MReport, CritEdt.Decimale) ;
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
   If PrintBand then Quoi:=QuoiCpt(2) ;
   END ;
CumDebitAux.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[1].TotDebit, CritEdt.AfficheSymbole ) ;
CumCreditAux.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[1].TotCredit,CritEdt.AfficheSymbole ) ;
DetDebitAux.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[2].TotDebit, CritEdt.AfficheSymbole ) ;
DetCreditAux.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[2].TotCredit,CritEdt.AfficheSymbole ) ;
SumDebitAux.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[3].TotDebit, CritEdt.AfficheSymbole ) ;
SumCreditAux.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[3].TotCredit,CritEdt.AfficheSymbole ) ;
end;

procedure TFQRBLHTGA.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TotCumDebit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[1].TotDebit,CritEdt.AfficheSymbole ) ;
TotCumCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotEdt[1].TotCredit,CritEdt.AfficheSymbole ) ;
TotDetDebit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[2].TotDebit,CritEdt.AfficheSymbole ) ;
TotDetCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[2].TotCredit,CritEdt.AfficheSymbole ) ;
TotSumDebit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[3].TotDebit,CritEdt.AfficheSymbole ) ;
TotSumCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[3].TotCredit,CritEdt.AfficheSymbole ) ;
TotSolDebit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[4].TotDebit,CritEdt.AfficheSymbole ) ;
TotSolCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[4].TotCredit,CritEdt.AfficheSymbole ) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFQRBLHTGA.BFCpteGeneBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
if CritEdt.Bal.RuptOnly=Sur then PrintBand:=False Else PrintBand:=Affiche ;
if PrintBand then
   BEGIN
   Quoi:=QuoiCpt(1) ;
   CumDebitGen.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGen[1].TotDebit, CritEdt.AfficheSymbole ) ;
   CumCreditGen.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGen[1].TotCredit,CritEdt.AfficheSymbole ) ;
   DetDebitGen.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGen[2].TotDebit, CritEdt.AfficheSymbole ) ;
   DetCreditGen.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGen[2].TotCredit,CritEdt.AfficheSymbole ) ;
   SumDebitGen.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGen[3].TotDebit, CritEdt.AfficheSymbole ) ;
   SumCreditGen.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGen[3].TotCredit,CritEdt.AfficheSymbole ) ;

   SolDebitGen.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGen[4].TotDebit, CritEdt.AfficheSymbole ) ;
   SolCreditGen.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGen[4].TotCredit,CritEdt.AfficheSymbole ) ;
   CumulVersSolde(TotCptGen[4]) ;
   if TotCptGen[4].TotDebit=0 then
      BEGIN
      LeSoldeD.Caption:='' ;
      LeSoldeC.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGen[4].TotCredit,CritEdt.AfficheSymbole ) ;
      END else
      BEGIN
      LeSoldeD.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptGen[4].TotDebit,CritEdt.AfficheSymbole ) ;
      LeSoldeC.Caption:='' ;
      END ;
   END ;
end;

procedure TFQRBLHTGA.BFCpteGeneAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Fillchar(TotCptGen,SizeOf(TotCptGen),#0) ;
InitReport([2],CritEdt.Bal.FormatPrint.Report) ;
end;

procedure TFQRBLHTGA.FAux1Change(Sender: TObject);
Var XAvecJoker : Boolean ;
begin
  inherited;
XAvecJoker:=Joker(FAux1, FAux2, FAuxJoker) ;
TFaA.Visible:=Not XAvecJoker ;
TFAux.Visible:=Not XAvecJoker ;
TFAuxJoker.Visible:=XAvecJoker ;
end;

procedure TFQRBLHTGA.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
StReportGen:=HG_GENERAL.Caption ;
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
If PrintBand then Quoi:=QuoiCpt(0) ;
end;

procedure TFQRBLHTGA.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
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

procedure TFQRBLHTGA.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var MReport : TabTRep ;
begin
  inherited;
TITREREPORTB.Left:=TitreColCpt.Left ;
TITREREPORTB.Width:=TitreColCpt.Width+TitreColLibelle.Width+1 ;
Case QuelReportBAL(CritEdt.Bal.FormatPrint.Report,MReport) Of
  1 : TITREREPORTB.Caption:=MsgLibel.Mess[7] ;
  2 : TITREREPORTB.Caption:=MsgLibel.Mess[8]+' '+StReportGen ;
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

procedure TFQRBLHTGA.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr11G_GENERAL         :=TStringField(Q.FindField('G_GENERAL'));
Qr11G_LIBELLE         :=TStringField(Q.FindField('G_LIBELLE'));
Qr11G_COLLECTIF       :=TStringField(Q.FindField('G_COLLECTIF'));
If CritEDT.Rupture=rCorresp then
   BEGIN
   Qr11G_CORRESP1         :=TStringField(Q.FindField('G_CORRESP1'));
   Qr11G_CORRESP2         :=TStringField(Q.FindField('G_CORRESP2'));
   END ;
end;

procedure TFQRBLHTGA.QAuxAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr12T_AUXILIAIRE:=TStringField(QAux.FindField('T_AUXILIAIRE'));
Qr12T_LIBELLE:=TStringField(QAux.FindField('T_LIBELLE'));
end;

procedure TFQRBLHTGA.FNatureCptChange(Sender: TObject);
var CJ,C1,C2 : String ;
begin
C1:=FCpte1.Text ; C2:=FCpte2.Text ;
if FJoker.Visible then FJoker.Text:=CJ ;
  inherited;
FCpte1.Clear ; FCpte2.Clear ; FJoker.Clear ;
Case FNatureCpt.ItemIndex of
  0     : BEGIN FCpte1.ZoomTable:=tzGCharge  ; FCpte2.ZoomTable:=tzGProduit ; END ;
  1     : BEGIN FCpte1.ZoomTable:=tzGCharge  ; FCpte2.ZoomTable:=tzGCharge  ; END ;
  2     : BEGIN FCpte1.ZoomTable:=tzGProduit ; FCpte2.ZoomTable:=tzGProduit ; END ;
  End ;
FCpte1.Text:=C1 ; FCpte2.Text:=C2 ;
if FJoker.Visible then FJoker.Text:=CJ ;
end;

procedure TFQRBLHTGA.BNouvRechClick(Sender: TObject);
begin
  inherited;
If FAuxJoker.Visible then FAuxJoker.Text:='' ;
end;

procedure TFQRBLHTGA.DLRuptNeedData(var MoreData: Boolean);
Var TotRupt : Array[0..12] of Double ;
    CodRupt, LibRupt,Lib1, CptRupt, Stcode : String ;
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
   TCodRupt.Width:=TitreColCpt.Width+TitreColLibelle.Width+1 ;
   TCodRupt.Caption:='' ;
   if CritEdt.Rupture=rLibres then
      BEGIN
      insert(MsgLibel.Mess[12]+' ',CodRupt,Quellerupt+2) ;
      TCodRupt.Caption:=CodRupt+' '+Lib1 ;
      END Else TCodRupt.Caption:=CodRupt+'   '+LibRupt ;
   SumD:=TotRupt[1]+TotRupt[3] ;
   SumC:=TotRupt[2]+TotRupt[4] ;
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

procedure TFQRBLHTGA.BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
PrintBand:=CritEdt.Rupture<>rRien ;
end;

procedure TFQRBLHTGA.FSansRuptClick(Sender: TObject);
begin
  inherited;
FLigneRupt.Enabled:=Not FSansRupt.Checked ;
FLigneRupt.checked:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
FRupturesClick(Nil) ;
end;

procedure TFQRBLHTGA.FPlanRupturesChange(Sender: TObject);
begin
  inherited;
If QRLoading Then Exit ;
if FPlansCo.Checked then CorrespToCodes(FPlanRuptures,FCodeRupt1,FCodeRupt2) Else
if FRuptures.Checked then RuptureToCodes(FPlanRuptures,FCodeRupt1,FCodeRupt2,FbGene) ;
FCodeRupt1.ItemIndex:=0 ; FCodeRupt2.ItemIndex:=FCodeRupt2.Items.Count-1 ;
end;

procedure TFQRBLHTGA.FRupturesClick(Sender: TObject);
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
FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
end;

Procedure TFQRBLHTGA.CalcTotalEdt(AnVD, AnVC, D, C : Double );
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

procedure TFQRBLHTGA.FNatCpteAuxChange(Sender: TObject);
var CJ,C1,C2 : String ;
begin
C1:=FAux1.Text ; C2:=FAux2.Text ;
CJ:=FAuxJoker.Text ;
  inherited;
FAux1.Clear ; FAux2.Clear ; FAuxJoker.Clear ;
Case FNatCpteAux.ItemIndex of
  0     : BEGIN FAux1.ZoomTable:=tzTiers    ; FAux2.ZoomTable:=tzTiers    ; END ;
  1     : BEGIN FAux1.ZoomTable:=tzTClient  ; FAux2.ZoomTable:=tzTClient  ; END ;
  2     : BEGIN FAux1.ZoomTable:=tzTDivers  ; FAux2.ZoomTable:=tzTDivers  ; END ;
  3     : BEGIN FAux1.ZoomTable:=tzTFourn   ; FAux2.ZoomTable:=tzTFourn   ; END ;
  4     : BEGIN FAux1.ZoomTable:=tzTSalarie ; FAux2.ZoomTable:=tzTSalarie ; END ;
  End ;
(* En CollDivers Zoom sur tzTCrediteur ET tzTDebiteur ET tzTDivers *)
FAux1.Text:=C1 ; FAux2.Text:=C2 ;
if FAuxJoker.Visible then FAuxJoker.Text:=CJ ;
end;

end.
