unit QRBlAna ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, UtilEdt, CritEdt, EdtLegal, QrRupt, Ent1, HEnt1,
  CpteUtil, CPSection_TOM, HSysMenu, Menus, HTB97, HPanel, UiUtil, tCalcCum  ;

procedure BLAnal ;
procedure BLAnalZoom(Crit : TCritEdt) ;

type
  TFQRBLANA = class(TFQR)
    TitreColCpt: TQRLabel;
    TitreColLibelle: TQRLabel;
    TTitreColAvant: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel16: TQRLabel;
    TTitreColSelection: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel9: TQRLabel;
    TTitreColApres: TQRLabel;
    T_DEBIT: TQRLabel;
    T_CREDIT: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    TTitreColSolde: TQRLabel;
    REPORT1DEBITanv: TQRLabel;
    REPORT1CREDITanv: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    REPORT1DEBITsum: TQRLabel;
    REPORT1CREDITsum: TQRLabel;
    REPORT1DEBITsol: TQRLabel;
    REPORT1CREDITsol: TQRLabel;
    S_SECTION: TQRDBText;
    S_LIBELLE: TQRDBText;
    ANvDEBIT: TQRLabel;
    ANvCREDIT: TQRLabel;
    DEBIT: TQRLabel;
    CREDIT: TQRLabel;
    DEBITsum: TQRLabel;
    CREDITsum: TQRLabel;
    SOLDEdeb: TQRLabel;
    SOLDEcre: TQRLabel;
    BRUPT: TQRBand;
    RDEBIT: TQRLabel;
    RCREDIT: TQRLabel;
    TCodRupt: TQRLabel;
    RSOLDEdeb: TQRLabel;
    RSOLDEcre: TQRLabel;
    RDEBITsum: TQRLabel;
    RCREDITsum: TQRLabel;
    ANvRDEBIT: TQRLabel;
    ANvRCREDIT: TQRLabel;
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
    Trait2: TQRLigne;
    Trait3: TQRLigne;
    Trait4: TQRLigne;
    Trait5: TQRLigne;
    Trait0: TQRLigne;
    Trait6: TQRLigne;
    Ligne1: TQRLigne;
    REPORT2DEBITanv: TQRLabel;
    REPORT2CREDITanv: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    REPORT2CREDITsum: TQRLabel;
    REPORT2DEBITsum: TQRLabel;
    REPORT2DEBITsol: TQRLabel;
    REPORT2CREDITsol: TQRLabel;
    MsgLibel: THMsgBox;
    DLrupt: TQRDetailLink;
    FCol1: TCheckBox;
    FCol2: TCheckBox;
    FCol3: TCheckBox;
    FCol4: TCheckBox;
    FLigneRupt: TCheckBox;
    FAvecAN: TCheckBox;
    FSansANO: TCheckBox;
    BRuptTRI: TQRBand;
    DetDebRuptTRI: TQRLabel;
    CumCreRuptTRI: TQRLabel;
    CumDebRuptTRI: TQRLabel;
    TCodRuptTRI: TQRLabel;
    SumDebRuptTRI: TQRLabel;
    SumCreRuptTRI: TQRLabel;
    DetCreRuptTRI: TQRLabel;
    SolCreRuptTRI: TQRLabel;
    SolDebRuptTRI: TQRLabel;
    DLRUPTTRI: TQRDetailLink;
    procedure FormShow(Sender: TObject);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean;      var Quoi: string);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BRUPTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean;      var Quoi: string);
    procedure DLruptNeedData(var MoreData: Boolean);
    procedure FNatureCptChange(Sender: TObject);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean;      var Quoi: string);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure FExerciceChange(Sender: TObject);
    procedure FDateCompta1Change(Sender: TObject);
    procedure FSelectCpteChange(Sender: TObject);
    procedure FAvecANClick(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure FSansRuptClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DLRUPTTRINeedData(var MoreData: Boolean);
    procedure BRuptTRIBeforePrint(var PrintBand: Boolean;
      var Quoi: String);
    procedure BDetailAfterPrint(BandPrinted: Boolean);
    procedure BRUPTAfterPrint(BandPrinted: Boolean);
    procedure GenereSQL ; Override ;
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
  private    { Déclarations privées }
    LRupt,LRuptTri             : TStringList ;
    ListeCodesRupture : TStringList ;
    // Rony 17/03/97 QBal,QBalCum  : TQuery ;
    TotEdt     : TabTot ;
    DansTotal         : Boolean ;
    Qr11S_AXE         : TStringField ;
    Qr11S_SECTION     : TStringField ;
    Qr11S_LIBELLE     : TStringField ;
    Qr11S_SECTIONTRIE : TStringField ;
    Qr11S_CORRESP1, Qr11S_CORRESP2: TStringField ;
    Qr11STTS : tVariantField ;
    FLoading            : Boolean ;
    Procedure BalSectionZoom(Quoi : String) ;
    Procedure CalcTotalEdt(AnVD, AnVC, D, C : Double );
    Function  OkPourRupt : Boolean ;
    Function  ColonnesOK : Boolean ;
  public    { Déclarations publiques }
  end;

implementation

{$R *.DFM}
Uses Filtre ;

procedure BLAnal ;
var QR : TFQRBLANA ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQRBLANA.Create(Application) ;
Edition.Etat:=etBalAna ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalSectionZoom ;
QR.InitType (nbSec,neBal,msSecAna,'QRBALANA','',TRUE,FALSE,FALSE,Edition) ;
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

procedure BLAnalZoom(Crit : TCritEdt) ;
var QR : TFQRBLANA ;
    Edition : TEdition ;
BEGIN
QR:=TFQRBLANA.Create(Application) ;
Edition.Etat:=etBalAna ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.BalSectionZoom ;
 QR.CritEdt:=Crit ;
 QR.FPlanRuptures.Enabled:=False ; // Simon
 QR.InitType (nbSec,neBal,msSecAna,'QRBALANA','',FALSE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

Procedure TFQRBLANA.BalSectionZoom(Quoi : String) ;
BEGIN
Inherited ;
ZoomEdt(3,Quoi) ;
END ;

Procedure TFQRBLANA.FinirPrint ;
BEGIN
Inherited ;
if CritEdt.Rupture<>rRien then VideRupt(LRupt) ;
If CritEdt.Bal.STTS<>'' Then VideRupt(LRuptTRI) ;
If GCalc<>NIL Then BEGIN GCalc.Free ; GCalc:=NIL ; END ;
(* QBal.Free ; Rony 17/03/97
QBalCum.Free ; *)
If OkMajEdt Then
   MajEdition('BA'+Char(CritEdt.Bal.Axe[2]),CritEdt.Exo.Code,DateToStr(CritEdt.DateDeb),DateToStr(CritEdt.DateFin),
                 '',TotEdt[3].TotDebit,TotEdt[3].TotCredit,TotEdt[4].TotDebit,TotEdt[4].TotCredit) ;
END ;

procedure TFQRBLANA.InitDivers ;
BEGIN
Inherited ;
{ Init du Total Général }
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
{ Labels sur les bandes }
DateBandeauBalance(TTitreColAvant,TTitreColSelection,TTitreColApres,TTitreColSolde,CritEdt,
                   MsgLibel.Mess[0],MsgLibel.Mess[1],MsgLibel.Mess[2]) ;
BRupt.Frame.DrawTop:=CritEDT.BAL.FormatPrint.PrSepRupt  ;
BRupt.Frame.DrawBottom:=BRupt.Frame.DrawTop ;
ChangeFormatCompte(AxeToFb(CritEdt.Bal.Axe),AxeToFb(CritEdt.Bal.Axe),Self,S_SECTION.Width,1,2,S_SECTION.Name) ;
END ;

procedure TFQRBLANA.GenereSQL ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
//GenereSQLBase ;
If CritEdt.Bal.STTS<>'' Then GenereSQLBaseTS
                        Else GenereSQLBase ;
END;

procedure TFQRBLANA.RenseigneCritere ;
{ Récupération des champs du multicritère dans l'entête d'état }
BEGIN
Inherited ;
(*if FSansRupt.Checked=True then BEGIN RSansRupt.Caption:='þ' ; RAvecRupt.Caption:='o' ; RSurRupt.Caption:='o' ; END Else
if FAvecRupt.Checked=True then BEGIN RSansRupt.Caption:='o' ; RAvecRupt.Caption:='þ' ; RSurRupt.Caption:='o' ; END Else
if FSurRupt.Checked=True  then BEGIN RSansRupt.Caption:='o' ; RAvecRupt.Caption:='o' ; RSurRupt.Caption:='þ' ; END ;
If FPlanRupt.Enabled=True then BEGIN RPlanRupt.Caption:=FPlanRupt.Text ; RTriRupt1.Caption:=FTriRupt1.Text ; RTriRupt2.Caption:=FTriRupt2.Text ; END ;
if not CritEdt.Bal.RuptOnly Then begin RTriRupt1.Caption:='' ; RTriRupt2.Caption:='' ; End ;
CaseACocher(FCol1, RCol1) ; CaseACocher(FCol2, RCol2) ;
CaseACocher(FCol3, RCol3) ; CaseACocher(FCol4, RCol4) ;
CaseACocher(FLigneRupt, RLigneRupt) ;*)
END ;

procedure TFQRBLANA.ChoixEdition ;
BEGIN
Inherited ;
If CritEdt.Bal.STTS<>'' Then
  BEGIN
  DLRuptTRI.PrintBefore:=FALSE ;
  ChargeGroup(LRuptTRI,['TRI']) ;
  END  ;
DLRupt.PrintBefore:=FALSE ;
Case CritEdt.Rupture of
  rLibres   : BEGIN
              DLrupt.PrintBefore:=FALSE ;
              ChargeGroup(LRupt,['S00','S01','S02','S03','S04','S05','S06','S07','S08','S09']) ;
              END ;
  rCorresp : BEGIN
             DLrupt.PrintBefore:=TRUE ;
             ChargeRuptCorresp(LRupt,CritEdt.Bal.PlanRupt,CritEdt.Bal.CodeRupt1 ,CritEdt.Bal.CodeRupt2, False) ;
             NiveauRupt(LRupt) ;
             END ;
  End ;
END ;

procedure TFQRBLANA.RecupCritEdt ;
Var NonLibres : Boolean ;
BEGIN
InHerited ;
With CritEdt Do
  BEGIN
  Bal.CodeRupt1:='' ; Bal.CodeRupt2:='' ; Bal.PlanRupt:='' ;
  NonLibres:=((Rupture=rRuptures) or (Rupture=rCorresp)) ;
  if NonLibres then BAL.PlanRupt:=FPlanRuptures.Value ;

  BAL.RuptOnly:=QuelleTypeRupt(0,FSAnsRupt.Checked,FAvecRupt.Checked,FSurRupt.Checked) ;
  if NonLibres then BEGIN Bal.CodeRupt1:=FCodeRupt1.Text ; Bal.CodeRupt2:=FCodeRupt2.Text ; END ;
  Bal.Axe:=FNatureCpt.Value ;
  Bal.QuelAN:=FQuelAN(FAvecAN.Checked) ;
  If Bal.QuelAN=SansAN Then FCol1.Checked:=FALSE ;
  Bal.SansAnoANA:=FSansANO.Checked ;
  If (CritEdt.Rupture=rCorresp) Then Bal.PlansCorresp:=FPlanRuptures.ItemIndex+1 ;
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

Function TFQRBLANA.CritOK : Boolean ;
BEGIN
Result:=Inherited CritOK and OkPourRupt and ColonnesOK ;
If Result Then
   BEGIN
//   QBal:=PrepareTotCpt(AxeToFb(CritEdt.Bal.Axe),QuelTypeEcr,Dev,Etab,Exo) ;
// Rony 17/03/97 QBal:=PrepareTotCptCum(QBalCum,AxeToFb(CritEdt.Bal.Axe),QuelTypeEcr,Dev,Etab,Exo,DevEnP) ;
   Gcalc:=TGCalculCum.create(Un,AxeToFb(CritEdt.Bal.Axe),AxeToFb(CritEdt.Bal.Axe),QuelTypeEcr,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE,FALSE,'',CritEdt.FiltreSup,CritEdt.Bal.STTS) ;
   GCalc.initCalcul('','','',CritEdt.Bal.Axe,CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                    CritEdt.Date1,CritEdt.Bal.Date11,TRUE) ;



   END ;
InitReport([1],CritEDT.BAL.FormatPrint.Report) ;
END ;

Function TFQRBLANA.ColonnesOK : Boolean ;
BEGIN
If OkZoomEdt then begin Result:=True ; exit ; end ;
Result:=(FCol1.Checked or FCol2.Checked  or FCol3.Checked or FCol4.Checked) ;
If Not Result then NumErreurCrit:=9 ;
END ;

Procedure TFQRBLANA.CalcTotalEdt(AnVD, AnVC, D, C : Double );
Var SumD, SumC, TotSoldeD, TotSoldeC : Double ;
BEGIN
  inherited;
SumD:=AnVD+D ;
SumC:=AnVC+C ;
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

procedure TFQRBLANA.FormShow(Sender: TObject);
begin
HelpContext:=7451000 ;
//Standards.HelpContext:=7451010 ;
//Avances.HelpContext:=7451020 ;
//Mise.HelpContext:=7451030 ;
//Option.HelpContext:=7451040 ;
//TabRuptures.HelpContext:=7451050 ;
  inherited;
{$IFDEF CCS3}
FNatureCpt.Visible:=FALSE ; HLabel2.Visible:=FALSE ;
{$ENDIF}
If EstSerie(S3) Then FSansAno.Visible:=FALSE ;
Floading:=FALSE ;
TabSup.TabVisible:=False;
TFGenJoker.Visible:=False ; TFGen.Visible:=True ;
If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
FSansRuptClick(Nil);
FCol1.Checked:=TRUE ; FCol2.Checked:=TRUE ; FCol3.Checked:=TRUE ; FCol4.Checked:=TRUE ;
FOnlyCptAssocie.Enabled:=False ;
{ JP 23/08/04 : FQ 14242 : Nouvelle gestion des filtrs
ChargeFiltrePourRupture ;}
end;

procedure TFQRBLANA.FormCreate(Sender: TObject);
begin
  inherited;
ListeCodesRupture:=TStringList.Create ;
end;

procedure TFQRBLANA.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
ListeCodesRupture.Free ;
end;

procedure TFQRBLANA.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
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

procedure TFQRBLANA.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var TotCpt1, TotCpt2, Tot : TabTot ;
    MReport               : TabTRep ;
    CptRupt     : String ;
    VCS : Variant ;
begin
  inherited;
VCS:=RecupVCS(Qr11STTS) ;
PrintBand:=True ;
Fillchar(Tot,SizeOf(Tot),#0) ; Fillchar(TotCpt1,SizeOf(TotCpt1),#0) ; Fillchar(TotCpt2,SizeOf(TotCpt2),#0) ;
Fillchar(MReport,SizeOf(MReport),#0) ;
Case CritEdt.Bal.TypCpt of  { 0 : Mvt sur L'exo,    3 : MVt sur la Période }
  0 :  PrintBand:=True ;
  3 :  PrintBand:= TRUE ;
 end ;
if PrintBand then
   BEGIN
   Case CritEdt.Rupture of
     rLibres    : if CritEdt.Bal.OnlyCptAssocie then PrintBand:=DansRuptLibre(Q,AxeToFb(CritEdt.Bal.Axe),CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
     rRuptures  : if CritEdt.Bal.OnlyCptAssocie then PrintBand:=(Qr11S_SECTIONTRIE.AsString<>'') ;
     rCorresp   : if CritEdt.Bal.OnlyCptAssocie then
                    if CritEDT.Bal.PlansCorresp=1 then PrintBand:=DansRuptCorresp(LRupt,Qr11S_CORRESP1.AsString) Else
                    if CritEDT.Bal.PlansCorresp=2 then PrintBand:=DansRuptCorresp(LRupt,Qr11S_CORRESP2.AsString) ;
//                    if CritEDT.Bal.PlansCorresp=1 then PrintBand:=(Qr11S_CORRESP1.AsString<>'') Else
//                    if CritEDT.Bal.PlansCorresp=2 then PrintBand:=(Qr11S_CORRESP2.AsString<>'') ;
     End;
   END ;
if PrintBand then
   BEGIN
   Quoi:=Qr11S_SECTION.AsString+' '+'#'+Qr11S_LIBELLE.AsString+'@'+'3'+Qr11S_AXE.AsString ;
   (*
   ExecuteTotCpt(QBal,QS_SECTION.AsString,CritEdt CritEdt.Date1,CritEdt.Bal.Date11,
                 CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,TotCpt1,FALSE) ;
   ExecuteTotCpt(QBal,QS_SECTION.AsString, CritEdt.Bal.Date21,CritEdt.Date2,
                 CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,TotCpt2,FALSE) ;
   *)
   (*           Rony 17/03/97
   ExecuteTotCptCum(QBal,QBalCum,QS_SECTION.AsString,CritEdt.Bal.Axe,CritEdt.Date1,CritEdt.Bal.Date11,
                 CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,TotCpt1,FALSE) ;
   ExecuteTotCptCum(QBal,QBalCum,QS_SECTION.AsString,CritEdt.Bal.Axe,CritEdt.Bal.Date21,CritEdt.Date2,
                 CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,TotCpt2,FALSE) ;
   *)

   If CritEdt.Bal.QuelAN<>SansAN Then
      BEGIN
      If VH^.UseTC Then GCAlc.ReInitCalculVCS(Qr11S_SECTION.AsString,'',CritEDT.BAL.Date21,CritEDT.Date2,VCS)
                      Else GCAlc.ReInitCalcul(Qr11S_SECTION.AsString, '',CritEdt.Date1,CritEdt.Bal.Date11) ;
      GCalc.Calcul ; TotCpt1:=GCalc.ExecCalc.TotCpt ;
      END ;

   GCAlc.ReInitCalculVCS(Qr11S_SECTION.AsString, '',CritEdt.Bal.Date21,CritEdt.Date2,VCS) ;
   GCalc.Calcul ; TotCpt2:=GCalc.ExecCalc.TotCpt ;

   CumulVersSolde(TotCpt1[0]) ;

   If CritEdt.Date1=CritEdt.Bal.Date11 Then Fillchar(TotCpt1[1],SizeOf(TotCpt1[1]),#0) ;
   If CritEdt.Bal.QuelAN=SansAN Then
      BEGIN
      TotCpt1[0].TotDebit:=0 ; TotCpt1[0].TotCredit:=0 ;
      TotCpt1[1].TotDebit:=0 ; TotCpt1[1].TotCredit:=0 ;
      END ;
   If CritEdt.Bal.SansAnoAna Then
     BEGIN
     TotCpt1[0].TotDebit:=0 ; TotCpt1[0].TotCredit:=0 ;
     END ;
   Tot[1].TotDebit:= Arrondi(TotCpt1[0].TotDebit  + TotCpt1[1].TotDebit,CritEdt.Decimale) ;  { Anv + cumul }
   Tot[1].TotCredit:=Arrondi(TotCpt1[0].TotCredit + TotCpt1[1].TotCredit,CritEdt.Decimale) ;
   Tot[2].TotDebit:= Arrondi(TotCpt2[1].TotDebit,CritEdt.Decimale) ;                         { Montant }
   Tot[2].TotCredit:=Arrondi(TotCpt2[1].TotCredit,CritEdt.Decimale) ;
   Tot[3].TotDebit:= Arrondi(Tot[1].TotDebit+Tot[2].TotDebit,CritEdt.Decimale) ;             { Anv&Cum + Montant }
   Tot[3].TotCredit:=Arrondi(Tot[1].TotCredit+Tot[2].TotCredit,CritEdt.Decimale) ;
   MReport[1].TotDebit:= Arrondi(TotCpt1[0].TotDebit+TotCpt1[1].TotDebit,CritEdt.Decimale) ;
   MReport[1].TotCredit:=Arrondi(TotCpt1[0].TotCredit+TotCpt1[1].TotCredit,CritEdt.Decimale) ;
   MReport[2].TotDebit:= Arrondi(TotCpt2[1].TotDebit,CritEdt.Decimale) ;
   MReport[2].TotCredit:=Arrondi(TotCpt2[1].TotCredit,CritEdt.Decimale) ;
   MReport[3].TotDebit:= Arrondi(Tot[2].TotDebit+Tot[1].TotDebit,CritEdt.Decimale) ;
   MReport[3].TotCredit:=Arrondi(Tot[2].TotCredit+Tot[1].TotCredit,CritEdt.Decimale) ;
   Tot[4]:=Tot[3] ;
   CumulVersSolde(Tot[4]) ;
   if Tot[4].TotDebit=0 then
      BEGIN
      MReport[4].TotCredit:=Tot[4].TotCredit ;
      SOLDEdeb.Caption:='' ;
      SOLDEcre.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[4].TotCredit,CritEdt.AfficheSymbole) ;
      END else
      BEGIN
      MReport[4].TotDebit:=Tot[4].TotDebit ;
      SOLDEdeb.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[4].TotDebit,CritEdt.AfficheSymbole ) ;
      SOLDEcre.Caption:='' ;
      END ;
   AnvDebit.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[1].TotDebit,CritEdt.AfficheSymbole ) ;
   AnvCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[1].TotCredit,CritEdt.AfficheSymbole ) ;

   debit.caption:=    AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[2].TotDebit,CritEdt.AfficheSymbole ) ;
   credit.caption:=   AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[2].TotCredit,CritEdt.AfficheSymbole ) ;

   DEBITsum.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[3].TotDebit,CritEdt.AfficheSymbole ) ;
   CREDITsum.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[3].TotCredit,CritEdt.AfficheSymbole ) ;

   Case CritEdt.Rupture of
     rLibres   : AddGroupLibre(LRupt,Q,AxeToFb(CritEdt.Bal.Axe),CritEdt.LibreTrie,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,0]) ;
     rRuptures : AddRupt(LRupt,Qr11S_SECTIONTRIE.AsString,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,0]) ;
     rCorresp  : BEGIN
                 Case CritEDT.Bal.PlansCorresp Of
                   1 : If Qr11S_CORRESP1.AsString<>'' Then CptRupt:=Qr11S_CORRESP1.AsString+Qr11S_SECTION.AsString
                                                      Else CptRupt:='.'+Qr11S_SECTION.AsString ;
                   2 : If Qr11S_CORRESP2.AsString<>'' Then CptRupt:=Qr11S_CORRESP2.AsString+Qr11S_SECTION.AsString
                                                      Else CptRupt:='.'+Qr11S_SECTION.AsString ;
                   Else CptRupt:=Qr11S_SECTION.AsString ;
                  End ;
                 AddRuptCorres(LRupt,CptRupt,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,0]) ;
                 END ;
     End ;

   If CritEdt.Bal.STTS<>'' Then AddGroup(LRuptTri,[Qr11STTS],[Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit]) ;

   If (CritEdt.Bal.RuptOnly=Sur) then
      BEGIN
      PrintBand:=False ;
      IF CritEdt.Rupture=rLibres then CalcTotalEdt(Tot[1].TotDebit, Tot[1].TotCredit, Tot[2].TotDebit, Tot[2].TotCredit) ;
      END Else CalcTotalEdt(Tot[1].TotDebit, Tot[1].TotCredit, Tot[2].TotDebit, Tot[2].TotCredit) ;

   AddReportBAL([1],CritEdt.Bal.FormatPrint.Report,MReport,CritEdt.Decimale) ;
   END ;
end;

procedure TFQRBLANA.BRUPTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.Bal.RuptOnly<>Sans) ;
end;

procedure TFQRBLANA.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TOT2DEBITanv.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[1].TotDebit,CritEdt.AfficheSymbole ) ;
TOT2CREDITanv.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[1].TotCredit,CritEdt.AfficheSymbole ) ;

TOT2DEBIT.Caption:=    AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[2].TotDebit,CritEdt.AfficheSymbole ) ;
TOT2CREDIT.Caption:=   AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[2].TotCredit,CritEdt.AfficheSymbole ) ;

TOT2DEBITsum.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[3].TotDebit,CritEdt.AfficheSymbole ) ;
TOT2CREDITsum.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[3].TotCredit,CritEdt.AfficheSymbole ) ;

TOT2SOLdeb.Caption:=   AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[4].TotDebit,CritEdt.AfficheSymbole ) ;
TOT2SOLcre.Caption:=   AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[4].TotCredit,CritEdt.AfficheSymbole ) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFQRBLANA.DLruptNeedData(var MoreData: Boolean);
Var CodRupt,LibRupt,Lib1  : String ;
    SumD, SumC, TotSoldeD, TotSoldeC : Double ;
    TotRupt    : Array[0..12] of Double ;
    CptRupt, StCode     : String ;
    Quellerupt : Integer ;
    OkOk,AddTotEdt : Boolean ;
    Col        : TColor ;
    LibRuptInf : Array[1..10] Of TRuptInf ;
begin
  inherited;
MoreData:=False ;
Case CritEdt.Rupture of
  rLibres   : BEGIN
              MoreData:=PrintGroupLibre(LRupt,Q,AxeToFb(CritEdt.Bal.Axe),CritEdt.LibreTrie,CodRupt,LibRupt,Lib1,TotRupt,Quellerupt,Col,LibRuptInf,CritEdt.Bal.STTS) ;
              If MoreData then
                 BEGIN
                 StCode:=CodRupt ;
                 Delete(StCode,1,Quellerupt+2) ;
                 MoreData:=DansChoixCodeLibre(StCode,Q,AxeToFb(CritEdt.Bal.Axe),CritEdt.LibreCodes1,CritEdt.LibreCodes2, CritEdt.LibreTrie) ;
                 BRupt.Font.Color:=Col ;
                 END ;
              END ;
  rRuptures : MoreData:=PrintRupt2(LRupt,Qr11S_SECTIONTRIE.AsString,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt,[Qr11S_SECTIONTRIE],Q,CritEdt.Bal.STTS) ;
//  rRuptures : MoreData:=PrintRupt2(LRupt,Qr11S_SECTIONTRIE.AsString,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt,Q,[Qr11S_SECTIONTRIE]) ;
  rCorresp  : BEGIN
              OkOk:=TRUE ;
              Case CritEDT.Bal.PlansCorresp  Of
                1 : BEGIN
                    If Qr11S_CORRESP1.AsString<>'' Then CptRupt:=Qr11S_CORRESP1.AsString+Qr11S_SECTION.AsString
                                                   Else CptRupt:='.'+Qr11S_SECTION.AsString ;
                    If OkOk Then MoreData:=PrintRupt2(LRupt,CptRupt,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt,[Qr11S_CORRESP2],Q,CritEdt.Bal.STTS) Else MoreData:=FALSE ;
                    END ;
                2 : BEGIN
                    If Qr11S_CORRESP2.AsString<>'' Then CptRupt:=Qr11S_CORRESP2.AsString+Qr11S_SECTION.AsString
                                                   Else CptRupt:='.'+Qr11S_SECTION.AsString ;
                    If OkOk Then MoreData:=PrintRupt2(LRupt,CptRupt,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt,[Qr11S_CORRESP2],Q,CritEdt.Bal.STTS) Else MoreData:=FALSE ;
                    END ;
                Else OkOk:=FALSE ;
                End ;
              END ;
  End ;

If MoreData then
   BEGIN
   TCodRupt.Width:=TitreColCpt.Width+TitreColLibelle.Width+1 ;
   TCodRupt.Caption:='' ;
   BRupt.Height:=HauteurBandeRuptIni ;
   if CritEdt.Rupture=rLibres then
      BEGIN
      insert(MsgLibel.Mess[13]+' ',CodRupt,Quellerupt+2) ;
      TCodRupt.Caption:=CodRupt+' '+Lib1 ;
      AlimRuptSup(HauteurBandeRuptIni,QuelleRupt,TCodRupt.Width,BRupt,LibRuptInf,Self) ;
      END Else TCodRupt.Caption:=CodRupt+'   '+LibRupt ;
   SumD:=TotRupt[1]+TotRupt[3] ;
   SumC:=TotRupt[2]+TotRupt[4] ;
   CalCulSolde(SumD,SumC,TotSoldeD,TotSoldeC) ;
   ANvRDEBIT.Caption:=AfficheMontant (CritEdt.FormatMontant, CritEdt.Symbole,TotRupt[1], CritEDT.AfficheSymbole ) ;
   ANvRCREDIT.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotRupt[2], CritEDT.AfficheSymbole ) ;
   RDEBIT.Caption:=AfficheMontant    (CritEdt.FormatMontant, CritEdt.Symbole,TotRupt[3], CritEDT.AfficheSymbole ) ;
   RCREDIT.Caption:=AfficheMontant   (CritEdt.FormatMontant, CritEdt.Symbole,TotRupt[4], CritEDT.AfficheSymbole ) ;
   RDEBITsum.Caption:=AfficheMontant (CritEdt.FormatMontant, CritEdt.Symbole,SumD, CritEDT.AfficheSymbole ) ;
   RCREDITsum.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,SumC, CritEDT.AfficheSymbole ) ;
   RSOLDEdeb.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotSoldeD, CritEDT.AfficheSymbole) ;
   RSOLDEcre.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotSoldeC, CritEDT.AfficheSymbole) ;
if CritEdt.Bal.RuptOnly=Sur then
   BEGIN
   AddTotEdt:=False ;
   if (CritEdt.Rupture=rLibres) then AddTotEdt:=False else
   if (CritEdt.Rupture=rRuptures) or (CritEdt.Rupture=rCorresp) then AddTotEdt:=DansTotal ;
   if AddTotEdt then CalcTotalEdt(TotRupt[1], TotRupt[2], TotRupt[3], TotRupt[4]) ;
   END ;
(*  Rony 21/10/97 Avant Grosse MAJ
   if (CritEdt.Bal.RuptOnly=Sur) and DansTotal then
      BEGIN
      CalcTotalEdt(TotRupt[1], TotRupt[2], TotRupt[3], TotRupt[4]) ;
      END ;
   *)
   END ;
end;

procedure TFQRBLANA.FNatureCptChange(Sender: TObject);
begin
//FCpte1.Clear ; FCpte2.Clear ;
  inherited;
if FPlansCo.checked then
   BEGIN CorrespToCombo(FPlanRuptures,AxeToFb(FNatureCpt.Value)) ; Exit ; END ;
Case FNatureCpt.ItemIndex of
   0 : FPlanRuptures.Datatype:='ttRuptSect1' ;
   1 : FPlanRuptures.Datatype:='ttRuptSect2' ;
   2 : FPlanRuptures.Datatype:='ttRuptSect3' ;
   3 : FPlanRuptures.Datatype:='ttRuptSect4' ;
   4 : FPlanRuptures.Datatype:='ttRuptSect5' ;
  end ;
If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
end;

procedure TFQRBLANA.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var MReport : TabTRep ;
begin
  inherited;
Case QuelReportBAL(CritEdt.Bal.FormatPrint.Report,MReport) Of
  1 : TITREREPORTB.Caption:=MsgLibel.Mess[7] ;
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

procedure TFQRBLANA.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr11S_AXE         :=TStringField(Q.FindField('S_AXE'));
Qr11S_SECTION     :=TStringField(Q.FindField('S_SECTION'));
Qr11S_LIBELLE     :=TStringField(Q.FindField('S_LIBELLE'));
If CritEdt.Rupture=rRuptures Then Qr11S_SECTIONTRIE :=TStringField(Q.FindField('S_SECTIONTRIE'));
if CritEdt.Rupture=rCorresp then
   BEGIN
   Qr11S_CORRESP1         :=TStringField(Q.FindField('S_CORRESP1'));
   Qr11S_CORRESP2         :=TStringField(Q.FindField('S_CORRESP2'));
   END ;
If CritEdt.Bal.STTS<>'' Then
  BEGIN
  Qr11STTS:=TVariantField(Q.FindField(CritEdt.Bal.STTS)) ;
  END ;
end;

Function TFQRBLANA.OkPourRupt : Boolean ;
Var i : Integer ;
BEGIN
Result:=True ;
if (CritEdt.Rupture=rRuptures) then
   BEGIN
   Result:=False ; DLrupt.PrintBefore:=TRUE ;
   if ListeCodesRupture<>Nil then ListeCodesRupture.Clear ;
   For i:=0 to FCodeRupt1.Items.Count-1 do ListeCodesRupture.Add(FCodeRupt1.Items[i]) ;
   ChargeRupt(LRupt,'RU'+Char(CritEdt.Bal.Axe[2]),CritEdt.Bal.PlanRupt,CritEdt.Bal.CodeRupt1 ,CritEdt.Bal.CodeRupt2) ;
   NiveauRupt(LRupt) ;
   Case SectionRetrie(CritEdt.Bal.PlanRupt,CritEdt.Bal.Axe,ListeCodesRupture) of
     srOk              : Result:=True ;
     srNonStruct       : NumErreurCrit:=7 ; //MsgRien.Execute(7,'','') ;
     srPasEnchainement : NumErreurCrit:=8 ; // MsgRien.Execute(8,'','') ;
     End ;

   END ;
END ;

procedure TFQRBLANA.FExerciceChange(Sender: TObject);
begin
FLoading:=TRUE ;
  inherited;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
FLoading:=FALSE ;
end;

procedure TFQRBLANA.FDateCompta1Change(Sender: TObject);
begin
  inherited;
If FLoading Then Exit ;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
end;

procedure TFQRBLANA.FSelectCpteChange(Sender: TObject);
begin
  inherited;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
end;

procedure TFQRBLANA.FAvecANClick(Sender: TObject);
begin
  inherited;
FCol1.Checked:=FAvecAN.Checked ;
end;

procedure TFQRBLANA.FRupturesClick(Sender: TObject);
begin
  inherited;
If FPlansCo.Checked then FGroupRuptures.Caption:=' '+MsgLibel.Mess[11] ;
If FRuptures.Checked then FGroupRuptures.Caption:=' '+MsgLibel.Mess[10] ;
end;

procedure TFQRBLANA.FSansRuptClick(Sender: TObject);
begin
  inherited;
FLigneRupt.Enabled:=Not FSansRupt.Checked ;
FLigneRupt.checked:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
FRupturesClick(Nil) ;
end;


procedure TFQRBLANA.DLRUPTTRINeedData(var MoreData: Boolean);
Var CodTri,LibTri : String ;
    TotTri : Array[0..12] of Double ;
    NumR : Integer ;
    SumD, SumC, TotSoldeD, TotSoldeC : Double ;
begin
  inherited;
If CritEdt.Bal.STTS='' Then  BEGIN MoreData:=FALSE ; Exit ; END ;
MoreData:=PrintGroup(LRuptTri,Q,[Qr11STTS],CodTri,LibTri,TotTri,NumR);

if MoreData then
   BEGIN
   TCodRuptTRI.Width:=TitreColCpt.Width+TitreColLibelle.Width+1 ;
   TCodRuptTRI.Caption:='TOTAL '+GetNomChampSTTS+' '+CodTri ;
   BRuptTRI.Height:=HauteurBandeRuptIni ;
   SumD:=TotTRI[0]+TotTRI[2] ;
   SumC:=TotTRI[1]+TotTRI[3] ;

   CalCulSolde(SumD,SumC,TotSoldeD,TotSoldeC) ;
   CumDebRuptTRI.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotTRI[0], CritEDT.AfficheSymbole) ;
   CumCreRuptTRI.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotTRI[1], CritEDT.AfficheSymbole) ;
   DetDebRuptTRI.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotTRI[2], CritEDT.AfficheSymbole) ;
   DetCreRuptTRI.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotTRI[3], CritEDT.AfficheSymbole) ;
   SumDebRuptTRI.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumD  , CritEDT.AfficheSymbole) ;
   SumCreRuptTRI.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumC  , CritEDT.AfficheSymbole) ;
   SolDebRuptTRI.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotSoldeD, CritEDT.AfficheSymbole) ;
   SolCreRuptTRI.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotSoldeC, CritEDT.AfficheSymbole) ;
   If CritEdt.SautPageTRI Then
     BEGIN
     If CritEdt.BAL.RuptOnly=Sur Then Brupt.ForceNewPage:=TRUE Else BDetail.ForceNewPage:=TRUE ;
     END ;
   END ;
end;


procedure TFQRBLANA.BRuptTRIBeforePrint(var PrintBand: Boolean;
  var Quoi: String);
begin
  inherited;
PrintBand:=CritEdt.Bal.STTS<>'' ;
If PrintBand Then
  BEGIN
  if CritEdt.Rupture<>rRien then
    BEGIN
    ReinitRupt(LRupt) ;
    (*
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
                  END  ;
      End ;
      *)
    END ;
  END ;
end;

procedure TFQRBLANA.BDetailAfterPrint(BandPrinted: Boolean);
begin
  inherited;
BDetail.ForceNewPage:=FALSE ;
end;

procedure TFQRBLANA.BRUPTAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Brupt.ForceNewPage:=FALSE ;
end;

end.
