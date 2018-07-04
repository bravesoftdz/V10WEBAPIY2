unit QRBLSEGE;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, UtilEDT, CritEDT, QRRupt, CpteUtil, EdtLegal,
  Ent1, HEnt1, Menus, HSysMenu, Calcole, HTB97, HPanel, UiUtil, tCalcCum ;

procedure BLSEGE ;
procedure BLSEGEZoom(Crit : TCritEdt) ;

type
  TFQRBLSEGE = class(TFQR)
    TFGeneral: THLabel;
    FGeneral1: THCpteEdit;
    FGeneral2: THCpteEdit;
    TFJokerGene: THLabel;
    FJokerGene: TEdit;
    HLabel3: THLabel;
    FExceptionG: TEdit;
    TRGeneral: TQRLabel;
    RGeneral1: TQRLabel;
    TRaG: TQRLabel;
    RGeneral2: TQRLabel;
    QRLabel9: TQRLabel;
    RExceptionG: TQRLabel;
    TitreColCpt: TQRLabel;
    TitreColLibelle: TQRLabel;
    TTitreColAvant: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel10: TQRLabel;
    TTitreColSelection: TQRLabel;
    TTitreColApres: TQRLabel;
    T_DEBIT: TQRLabel;
    T_CREDIT: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel8: TQRLabel;
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
    S_SECTION: TQRDBText;
    S_LIBELLE: TQRDBText;
    BSubDetail: TQRBand;
    SOLDEdeb: TQRLabel;
    ANvDEBIT: TQRLabel;
    ANvCREDIT: TQRLabel;
    DEBITsum: TQRLabel;
    CREDITsum: TQRLabel;
    SOLDEcre: TQRLabel;
    G_GENERAL: TQRDBText;
    G_LIBELLE: TQRDBText;
    DEBIT: TQRLabel;
    CREDIT: TQRLabel;
    BTOTSEC: TQRBand;
    S_SECTION_: TQRDBText;
    S_LIBELLE_: TQRDBText;
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
    REPORT2DEBITsum: TQRLabel;
    REPORT2CREDITsum: TQRLabel;
    REPORT2DEBITsol: TQRLabel;
    REPORT2CREDITsol: TQRLabel;
    QGene: TQuery;
    QGeneG_GENERAL: TStringField;
    QGeneG_LIBELLE: TStringField;
    SGene: TDataSource;
    DLGENE: TQRDetailLink;
    MsgLibel: THMsgBox;
    FCol1: TCheckBox;
    FCol2: TCheckBox;
    FCol3: TCheckBox;
    FCol4: TCheckBox;
    TFaS: TLabel;
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
    DLMulti: TQRDetailLink;
    FLigneRupt: TCheckBox;
    FAvecCptSecond: TCheckBox;
    BRecap: TQRBand;
    RSolDebRupt: TQRLabel;
    RCumDebRupt: TQRLabel;
    RCumCreRupt: TQRLabel;
    RSumDebRupt: TQRLabel;
    RSumCreRupt: TQRLabel;
    RSolCreRupt: TQRLabel;
    RDetDebRupt: TQRLabel;
    RDetCreRupt: TQRLabel;
    RG_GENERAL: TQRLabel;
    RG_LIBELLE: TQRLabel;
    DLRecap: TQRDetailLink;
    BHRecap: TQRBand;
    HDetDebRupt: TQRLabel;
    HCumCreRupt: TQRLabel;
    HCumDebRupt: TQRLabel;
    HTCodRupt: TQRLabel;
    HSumDebRupt: TQRLabel;
    HSumCreRupt: TQRLabel;
    HDetCreRupt: TQRLabel;
    HSolCreRupt: TQRLabel;
    HSolDebRupt: TQRLabel;
    DLHRecap: TQRDetailLink;
    procedure FormShow(Sender: TObject);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BSubDetailBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BTOTSECBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BTOTSECAfterPrint(BandPrinted: Boolean);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure QGeneAfterOpen(DataSet: TDataSet);
    procedure FNatureCptChange(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure FGeneral1Change(Sender: TObject);
    procedure BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure DLMultiNeedData(var MoreData: Boolean);
    procedure FSansRuptClick(Sender: TObject);
    procedure FPlanRupturesChange(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DLRecapNeedData(var MoreData: Boolean);
    procedure BRuptAfterPrint(BandPrinted: Boolean);
    procedure BRecapBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure DLHRecapNeedData(var MoreData: Boolean);
  private { Déclarations privées }
    OkEnTeteGen, OkEnTeteRecap,OkRappelEnteteLibre : Boolean ;
    Qr11S_AXE, Qr11S_SECTIONTRIE   : TStringField;
    Qr11S_SECTION                  : TStringField;
    Qr11S_LIBELLE,
    Qr11S_CORRESP1, Qr11S_CORRESP2 : TStringField;
    Qr12G_GENERAL                  : TStringField;
    Qr12G_LIBELLE                  : TStringField;
    Affiche                        : Boolean ;
    StReportSec,StRappelLibTableLibre : String ;
    LRupt, LRecap                  : TStringList ;
    ListeCodesRupture              : TStringList ;
    TotCptSect, TotEdt : TabTot ;
    procedure GenereSQLSub ;
    Procedure BalSeGeZoom(Quoi : String) ;
    Function  QuoiCpt(i : Integer) : String ;
    Function  QuoiCptRecap(St1,St2 : String) : String ;
    Function  OkPourRupt : Boolean ;
    Function  ColonnesOK : Boolean ;
    Procedure CalcTotalEdt(AnVD, AnVC, D, C : Double );
public  { Déclarations publiques }
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

Uses CPSection_TOM ;

procedure BLSEGE ;
var QR : TFQRBLSEGE ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQRBLSEGE.Create(Application) ;
Edition.Etat:=etBalAnaGen ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalSeGeZoom ;
QR.InitType(nbSec,neBal,msSecAna,'QRBASEGE','',TRUE,FALSE,TRUE,Edition) ;
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

procedure BLSEGEZoom(Crit : TCritEdt) ;
var QR : TFQRBLSEGE ;
    Edition : TEdition ;
BEGIN
QR:=TFQRBLSEGE.Create(Application) ;
Edition.Etat:=etBalAnaGen ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.BalSeGeZoom ;
 QR.CritEdt:=Crit ;
 QR.InitType (nbSec,neBal,msSecAna,'QRBASEGE','',FALSE,TRUE,TRUE,Edition) ;
 finally
 QR.Free ;
 end ;
END ;

Function TFQRBLSEGE.QuoiCpt(i : Integer) : String ;
BEGIN
Case i Of
  0 : Result:=Qr11S_SECTION.AsString+'#'+Qr11S_LIBELLE.AsString+'@'+'3'+Qr11S_AXE.AsString ;
  1 : Result:=Qr11S_SECTION.AsString+'#'+Qr11S_LIBELLE.AsString+'@'+'3'+Qr11S_AXE.AsString ;
  2 : Result:=Qr12G_General.AsString+' '+Qr11S_SECTION.AsString+'#'+G_LIBELLE.Caption+'@'+'1' ;
 end ;
END ;

Function TFQRBLSEGE.QuoiCptRecap(St1,St2 : String) : String ;
BEGIN
Result:=St1+St2+'@'+'1' ;
END ;


Procedure TFQRBLSEGE.BalSeGeZoom(Quoi : String) ;
Var Lp,i: Integer ;
BEGIN
Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
i:=StrToInt(Copy(Quoi,Lp+1,1)) ;
ZoomEdt(i,Quoi) ;
END ;

procedure TFQRBLSEGE.InitDivers ;
BEGIN
Inherited ;
{ Init du Total Général }
Fillchar(TotCptSect,SizeOf(TotCptSect),#0) ;
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
{ Labels sur les bandes }
DateBandeauBalance(TTitreColAvant,TTitreColSelection,TTitreColApres,TTitreColSolde,CritEdt,
                   MsgLibel.Mess[0],MsgLibel.Mess[1],MsgLibel.Mess[2]) ;
ChangeFormatCompte(AxeToFb(FNatureCpt.Value),fbGene,Self,S_SECTION.Width,1,2,S_SECTION.Name) ;
BTOTSEC.Frame.DrawBottom:=CritEDT.BAL.FormatPrint.PrSepCompte ;
BRupt.Frame.DrawTop:=CritEDT.BAL.FormatPrint.PrSepRupt  ;
BRupt.Frame.DrawBottom:=BRupt.Frame.DrawTop ;
OkEnTeteGen:=True ;
OkRappelEnteteLibre:=FALSE ;
StRappelLibTableLibre:='' ;
If CritEdt.Rupture=rLibres then
   BEGIN
   if CritEdt.Bal.RuptOnly=Sur then OkEnTeteGen:=False ;
   If CritEdt.Bal.AvecCptSecond then OkEnTeteGen:=False ;
   END else
If CritEdt.Rupture=rRuptures then
   BEGIN
   if CritEdt.Bal.RuptOnly=Sur then OkEnTeteGen:=False ;
   If CritEdt.Bal.AvecCptSecond then OkEnTeteGen:=False ;
   END ;
END ;

procedure TFQRBLSEGE.GenereSQL ;
BEGIN
Inherited ;
AlimSQLMul(Q,0) ; Q.Close ; ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
Q.Open ;
GenereSQLSub ;
END;

Procedure TFQRBLSEGE.GenereSQLSub ;
BEGIN
Inherited ;
AlimSQLMul(QGene,1) ; ChangeSQL(QGene) ; //QGene.Prepare ;
PrepareSQLODBC(QGene) ;
QGene.Open ;
END ;

procedure TFQRBLSEGE.RenseigneCritere ;
{ Récupération des champs du multicritère dans l'entête d'état }
BEGIN
Inherited ;
If OkZoomEdt Then Exit ;
if CritEdt.SJoker then
   BEGIN
   TRGeneral.Caption:=MsgLibel.Mess[9] ;
   RGeneral1.Caption:=FJokerGene.Text ;
   END else
   BEGIN
   TRGeneral.Caption:=MsgLibel.Mess[10] ;
   RGeneral1.Caption:=CritEdt.LSCpt1 ; RGeneral2.Caption:=CritEdt.LSCpt2 ;
   END ;
RGeneral2.Visible:=Not CritEdt.SJoker ; TRaG.Visible:=Not CritEdt.SJoker ;
RExceptionG.Caption:=FExceptionG.Text ;
(* Rony 17/03/97
CaseACocher(FCol1, RCol1) ;
CaseACocher(FCol2, RCol2) ;
CaseACocher(FCol3, RCol3) ;
CaseACocher(FCol4, RCol4) ;*)
END ;

procedure TFQRBLSEGE.ChoixEdition ;
BEGIN
Inherited ;
Case CritEdt.Rupture of          
  rLibres  : BEGIN
             DLMulti.PrintBefore:=FALSE ;
             ChargeGroup(LRupt,['S00','S01','S02','S03','S04','S05','S06','S07','S08','S09']) ;
             if CritEdt.Bal.AvecCptSecond then ChargeRecap(LRecap) ;
             END ;
  rCorresp : BEGIN
             DLMulti.PrintBefore:=TRUE ;
             ChargeRuptCorresp(LRupt,CritEdt.Bal.PlanRupt,CritEdt.Bal.CodeRupt1 ,CritEdt.Bal.CodeRupt2, False) ;
             NiveauRupt(LRupt) ;
             END ;
  rRuptures: BEGIN
             if CritEdt.Bal.AvecCptSecond then ChargeRecap(LRecap) ;
             END ;
 End ;
END ;

procedure TFQRBLSEGE.RecupCritEdt ;
var st : String ;
    NonLibres : Boolean ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  SJoker:=FJokerGene.Visible ; Composite:=True ;
  If SJoker Then BEGIN SCpt1:=FJokerGene.Text ; SCpt2:=FJokerGene.Text ; END
            Else BEGIN
                 SCpt1:=FGeneral1.Text  ; SCpt2:=FGeneral2.Text ;
                 PositionneFourchetteSt(FGeneral1,FGeneral2,CritEdt.LSCpt1,CritEdt.LSCpt2) ;
                 END ;
  Bal.Axe:=FNatureCpt.Value ;
  Bal.SSauf:=Trim(FExceptionG.Text) ; st:=Bal.SSauf ;
  If Bal.SSauf<>'' then St:=' And '+AnalyseCompte(FExceptionG.Text,fbGene,True,False) ;
  Bal.SSqlSauf:=St ;
  BAL.RuptOnly:=QuelleTypeRupt(0,FSAnsRupt.Checked,FAvecRupt.Checked,FSurRupt.Checked) ;
  NonLibres:=((Rupture=rRuptures) or (Rupture=rCorresp)) ;
  if NonLibres then BAL.PlanRupt:=FPlanRuptures.Value ;
  if NonLibres then BEGIN BAL.CodeRupt1:=FCodeRupt1.Text ; BAL.CodeRupt2:=FCodeRupt2.Text ; END ;
  If (CritEdt.Rupture=rCorresp) Then Bal.PlansCorresp:=StrToInt(Copy(FPlanRuptures.Value,3,1)) ;
  Bal.OnlyCptAssocie:=(Rupture<>rRien) and FOnlyCptAssocie.Checked ;
  Bal.AvecCptSecond:=FAvecCptSecond.checked ;
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

Function TFQRBLSEGE.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK and ColonnesOK and OkPourRupt;
If Result Then
   BEGIN
   Gcalc:=TGCalculCum.create(Deux,AxeToFb(CritEdt.Bal.Axe),fbGene,QuelTypeEcr,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE) ;
   GCalc.initCalcul('','','',CritEdt.Bal.Axe,CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                    CritEdt.Date1,CritEdt.Bal.Date11,TRUE) ;
   END ;
END ;

Function TFQRBLSEGE.ColonnesOK : Boolean ;
BEGIN
If OkZoomEdt then begin Result:=True ; exit ; end ;
Result:=(FCol1.Checked or FCol2.Checked  or FCol3.Checked or FCol4.Checked) ;
If Not Result then NumErreurCrit:=7 ;
END ;

Function TFQRBLSEGE.OkPourRupt : Boolean ;
Var i : Integer ;
BEGIN
Result:=True ;
if (CritEdt.Rupture=rRuptures) then
   BEGIN
   Result:=False ; DLMulti.PrintBefore:=TRUE ;
   if ListeCodesRupture<>Nil then ListeCodesRupture.Clear ;
   For i:=0 to FCodeRupt1.Items.Count-1 do ListeCodesRupture.Add(FCodeRupt1.Items[i]) ;
   ChargeRupt(LRupt,'RU'+Char(CritEdt.Bal.Axe[2]),CritEdt.Bal.PlanRupt,CritEdt.Bal.CodeRupt1 ,CritEdt.Bal.CodeRupt2) ;
   NiveauRupt(LRupt) ;
   Case SectionRetrie(CritEdt.Bal.PlanRupt,CritEdt.Bal.Axe,ListeCodesRupture) of
     srOk              : Result:=True ;
     srNonStruct       : NumErreurCrit:=7 ;
     srPasEnchainement : NumErreurCrit:=8 ;
     End ;
   END ;
END ;

procedure TFQRBLSEGE.FormShow(Sender: TObject);
begin
FGeneral1.Text:='' ; FGeneral2.Text:='' ;
HelpContext:=7469000 ;
//Standards.HelpContext:=7469010 ;
//Avances.HelpContext:=7469020 ;
//Mise.HelpContext:=7469030 ;
//Option.HelpContext:=7469040 ;
//TabRuptures.HelpContext:=7469050 ;

  inherited;
{$IFDEF CCS3}
FNatureCpt.Visible:=FALSE ; HLabel2.Visible:=FALSE ;
{$ENDIF}
TabSup.TabVisible:=False;
//FAxe.ItemIndex:=0 ;
FCol1.Checked:=TRUE ; FCol2.Checked:=TRUE ; FCol3.Checked:=TRUE ; FCol4.Checked:=TRUE ;
FOnlyCptAssocie.Checked:=False ; FOnlyCptAssocie.Enabled:=False ;
end;

procedure TFQRBLSEGE.FormCreate(Sender: TObject);
begin
  inherited;
ListeCodesRupture:=TStringList.Create ;
end;

procedure TFQRBLSEGE.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
ListeCodesRupture.Free ;
end;

procedure TFQRBLSEGE.FinirPrint ;
begin
Inherited ;
QGene.Close ;
If GCalc<>NIL Then BEGIN GCalc.Free ; GCalc:=NIL ; END ;
if CritEdt.Rupture<>rRien then VideRupt(LRupt) ;
if (CritEdt.Rupture=rLibres) Or (CritEdt.Rupture=rRuptures)then if CritEdt.Bal.AvecCptSecond then VideRecap(LRecap) ;
end;

procedure TFQRBLSEGE.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
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

procedure TFQRBLSEGE.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
StReportSec:=S_SECTION.Caption ;
InitReport([2],CritEdt.Bal.FormatPrint.Report) ;
Case CritEdt.Bal.TypCpt of  { 0 : Mvt sur L'exo,    3 : MVt sur la Période }
  0 :  PrintBand:=True ;
  3 :  PrintBand:= TRUE ;
 end ;
Case CritEdt.Rupture of
  rLibres    : if CritEdt.Bal.OnlyCptAssocie then PrintBand:=DansRuptLibre(Q,AxeToFb(CritEdt.Bal.Axe),CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
  rRuptures  : if CritEdt.Bal.OnlyCptAssocie then PrintBand:=(Qr11S_SECTIONTRIE.AsString<>'') ;
  rCorresp   : if CritEdt.Bal.OnlyCptAssocie then
                  if CritEDT.Bal.PlansCorresp=1 then PrintBand:=DansRuptCorresp(LRupt,Qr11S_CORRESP1.AsString) Else
                  if CritEDT.Bal.PlansCorresp=2 then PrintBand:=DansRuptCorresp(LRupt,Qr11S_CORRESP2.AsString) ;
 //                if CritEDT.Bal.PlansCorresp=1 then PrintBand:=(Qr11S_CORRESP1.AsString<>'') Else
//                 if CritEDT.Bal.PlansCorresp=2 then PrintBand:=(Qr11S_CORRESP2.AsString<>'') ;
  End;
Affiche:=PrintBand ; { Pour ne pas afficher les sous comptes }
if CritEdt.Bal.RuptOnly=Sur then PrintBand:=False ;
If PrintBand then Quoi:=Quoicpt(0) ;
end;

procedure TFQRBLSEGE.BSubDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var TotCpt1, TotCpt2, Tot : TabTot ;
    MReport               : TabTRep ;
    CptRupt,St            : String ;
    LaRupture             : TRuptInf ;
begin
  inherited;
Fillchar(Tot,SizeOf(Tot),#0) ; Fillchar(TotCpt1,SizeOf(TotCpt1),#0) ; Fillchar(TotCpt2,SizeOf(TotCpt2),#0) ;
Fillchar(Mreport,SizeOf(MReport),#0) ;
PrintBand:=Affiche ;
if PrintBand then
   BEGIN
   Quoi:=Quoicpt(2) ;
   GCAlc.ReInitCalcul(Qr11S_SECTION.AsString, Qr12G_GENERAL.AsString,CritEdt.Date1,CritEdt.Bal.Date11) ;
   GCalc.Calcul ; TotCpt1:=GCalc.ExecCalc.TotCpt ;

   GCAlc.ReInitCalcul(Qr11S_SECTION.AsString, Qr12G_GENERAL.AsString,CritEdt.Bal.Date21,CritEdt.Date2) ;
   GCalc.Calcul ; TotCpt2:=GCalc.ExecCalc.TotCpt ;
   CumulVersSolde(TotCpt1[0]) ;
   If CritEdt.Date1=CritEdt.Bal.Date11 Then Fillchar(TotCpt1[1],SizeOf(TotCpt1[1]),#0) ;
   Tot[1].TotDebit:= TotCpt1[0].TotDebit  + TotCpt1[1].TotDebit ;  { Anv + cumul }
   Tot[1].TotCredit:=TotCpt1[0].TotCredit + TotCpt1[1].TotCredit ;
   Tot[2].TotDebit:= TotCpt2[1].TotDebit ;                         { Montant }
   Tot[2].TotCredit:=TotCpt2[1].TotCredit ;
   Tot[3].TotDebit:= Tot[1].TotDebit+Tot[2].TotDebit ;             { Anv&Cum + Montant }
   Tot[3].TotCredit:=Tot[1].TotCredit+Tot[2].TotCredit ;
   MReport[1].TotDebit:=TotCpt1[0].TotDebit+TotCpt1[1].TotDebit ;
   MReport[1].TotCredit:=TotCpt1[0].TotCredit+TotCpt1[1].TotCredit ;
   MReport[2].TotDebit:=TotCpt2[1].TotDebit ;
   MReport[2].TotCredit:=TotCpt2[1].TotCredit ;
   MReport[3].TotDebit:=Tot[2].TotDebit+Tot[1].TotDebit ;
   MReport[3].TotCredit:=Tot[2].TotCredit+Tot[1].TotCredit ;
   (* Rony 24/10/97 Passé au moule de CalcTotalEdt() en non Sur Rupture ou pas....
   TotEdt[1].totDebit:=Arrondi(TotEdt[1].totDebit + Tot[1].TotDebit,CritEdt.Decimale) ;
   TotEdt[1].totCredit:=Arrondi(TotEdt[1].totCredit + Tot[1].TotCredit,CritEdt.Decimale) ;
   TotEdt[2].totDebit:=Arrondi(TotEdt[2].totDebit + Tot[2].TotDebit,CritEdt.Decimale) ;
   TotEdt[2].totCredit:=Arrondi(TotEdt[2].totCredit + Tot[2].TotCredit,CritEdt.Decimale) ;
   TotEdt[3].totDebit:=Arrondi(TotEdt[3].totDebit + Tot[3].TotDebit,CritEdt.Decimale) ;
   TotEdt[3].totCredit:=Arrondi(TotEdt[3].totCredit + Tot[3].TotCredit,CritEdt.Decimale) ;
   *)
   TotCptSect[1].totDebit:=Arrondi(TotCptSect[1].totDebit + Tot[1].TotDebit,CritEdt.Decimale) ;
   TotCptsect[1].totCredit:=Arrondi(TotCptSect[1].totCredit + Tot[1].TotCredit,CritEdt.Decimale) ;
   TotCptSect[2].totDebit:=Arrondi(TotCptSect[2].totDebit + Tot[2].TotDebit,CritEdt.Decimale) ;
   TotCptSect[2].totCredit:=Arrondi(TotCptSect[2].totCredit + Tot[2].TotCredit,CritEdt.Decimale) ;
   TotCptSect[3].totDebit:=Arrondi(TotCptSect[3].totDebit + Tot[3].TotDebit,CritEdt.Decimale) ;
   TotCptSect[3].totCredit:=Arrondi(TotCptSect[3].totCredit + Tot[3].TotCredit,CritEdt.Decimale) ;
   Tot[4]:=Tot[3] ;
   CumulVersSolde(Tot[4]) ;
   if Tot[4].TotDebit=0 then
      BEGIN
      MReport[4].TotCredit:=Tot[4].TotCredit ;
      TotCptSect[4].TotCredit:=Arrondi(TotCptSect[4].TotCredit+Tot[4].TotCredit,CritEdt.Decimale) ;
      SOLDEdeb.Caption:='' ;
      SOLDEcre.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[4].TotCredit,CritEdt.AfficheSymbole) ;
      END else
      BEGIN
      MReport[4].TotDebit:=Tot[4].TotDebit ;
      TotCptSect[4].TotDebit:=Arrondi(TotCptSect[4].TotDebit+Tot[4].TotDebit,CritEdt.Decimale) ;
      SOLDEdeb.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[4].TotDebit,CritEdt.AfficheSymbole ) ;
      SOLDEcre.Caption:='' ;
      END ;
   AddReportBAL([1,2],CritEdt.Bal.FormatPrint.Report,MReport,CritEdt.Decimale) ;
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

   If (CritEdt.Bal.RuptOnly=Sur) then
      BEGIN
      PrintBand:=False  ;
      IF CritEdt.Rupture=rLibres then CalcTotalEdt(Tot[1].TotDebit, Tot[1].TotCredit, Tot[2].TotDebit, Tot[2].TotCredit) ;

//      if (CritEdt.Rupture=rRuptures) And CritEdt.Bal.AvecCptSecond then
      if ((CritEdt.Rupture=rLibres) Or (CritEdt.Rupture=rRuptures)) And CritEdt.Bal.AvecCptSecond then
         BEGIN
         if (CritEdt.Rupture=rRuptures) Then
           BEGIN
           DansQuelleRupt(LRupt,Qr11S_SECTIONTRIE.AsString,LaRupture) ;
           If LaRupture.CodRupt<>'' Then
              BEGIN
  //            St:=' ('+LaRupture.CodRupt+' '+LaRupture.LibRupt+')' ;
              St:=' ('+LaRupture.LibRupt+')' ;
              AddRecap(LRecap, [Qr12G_GENERAL.AsString],
                      [QR12G_LIBELLE.AsString+St],[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,0]) ;
              END ;
           END ;
         if (CritEdt.Rupture=rLibres) then
            BEGIN
            If DansRuptLibre(Q,AxeToFb(CritEdt.Bal.Axe),CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) Then
               BEGIN
   //            St:=' ('+LaRupture.CodRupt+' '+LaRupture.LibRupt+')' ;
//               St:=' ('+LaRupture.LibRupt+')' ;
               St:='' ;
               AddRecap(LRecap, [Qr12G_GENERAL.AsString],
                       [QR12G_LIBELLE.AsString+St],[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,0]) ;
               END ;
            END ;
         END Else PrintBand:=False ;


      END Else CalcTotalEdt(Tot[1].TotDebit, Tot[1].TotCredit, Tot[2].TotDebit, Tot[2].TotCredit) ;
   if PrintBand then Quoi:=Quoicpt(2) ;
   END ;
AnvDebit.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[1].TotDebit,CritEdt.AfficheSymbole ) ;
AnvCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[1].TotCredit,CritEdt.AfficheSymbole ) ;

debit.caption:=    AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[2].TotDebit,CritEdt.AfficheSymbole ) ;
credit.caption:=   AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[2].TotCredit,CritEdt.AfficheSymbole ) ;

DEBITsum.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[3].TotDebit,CritEdt.AfficheSymbole ) ;
CREDITsum.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[3].TotCredit,CritEdt.AfficheSymbole ) ;
end;

procedure TFQRBLSEGE.BTOTSECBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
if CritEdt.Bal.RuptOnly=Sur then PrintBand:=False Else PrintBand:=Affiche ;
if PrintBand then
   BEGIN
   Quoi:=Quoicpt(1) ;
   TOT1DEBITanv.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptSect[1].TotDebit, CritEdt.AfficheSymbole ) ;
   TOT1CREDITanv.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptSect[1].TotCredit,CritEdt.AfficheSymbole ) ;
   TOT1DEBIT.Caption :=   AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptSect[2].TotDebit, CritEdt.AfficheSymbole ) ;
   TOT1CREDIT.Caption:=   AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptSect[2].TotCredit, CritEdt.AfficheSymbole ) ;
   TOT1DEBITsum.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptSect[3].TotDebit, CritEdt.AfficheSymbole ) ;
   TOT1CREDITsum.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptSect[3].TotCredit, CritEdt.AfficheSymbole ) ;

   TOT1SOLdeb.Caption :=  AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptSect[4].TotDebit, CritEdt.AfficheSymbole ) ;
   TOT1SOLcre.Caption:=   AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptSect[4].TotCredit, CritEdt.AfficheSymbole ) ;
   CumulVersSolde(TotCptSect[4]) ;
   if TotCptSect[4].TotDebit=0 then
      BEGIN
      LeSoldeD.Caption:='' ;
      LeSoldeC.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptSect[4].TotCredit,CritEdt.AfficheSymbole ) ;
      END else
      BEGIN
      LeSoldeD.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptSect[4].TotDebit,CritEdt.AfficheSymbole ) ;
      LeSoldeC.Caption:='' ;
      END ;
   END ;
end;

procedure TFQRBLSEGE.BTOTSECAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Fillchar(TotCptSect,SizeOf(TotCptSect),#0) ;
InitReport([2],CritEdt.Bal.FormatPrint.Report) ;
end;

Procedure TFQRBLSEGE.CalcTotalEdt(AnVD, AnVC, D, C : Double );
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

procedure TFQRBLSEGE.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TOT2DEBITanv.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[1].TotDebit,CritEdt.AfficheSymbole ) ;
TOT2CREDITanv.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotEdt[1].TotCredit,CritEdt.AfficheSymbole ) ;

TOT2DEBIT.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[2].TotDebit,CritEdt.AfficheSymbole ) ;
TOT2CREDIT.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[2].TotCredit,CritEdt.AfficheSymbole ) ;

TOT2DEBITsum.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[3].TotDebit,CritEdt.AfficheSymbole ) ;
TOT2CREDITsum.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[3].TotCredit,CritEdt.AfficheSymbole ) ;

TOT2SOLdeb.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[4].TotDebit,CritEdt.AfficheSymbole ) ;
TOT2SOLcre.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[4].TotCredit,CritEdt.AfficheSymbole ) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFQRBLSEGE.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var MReport : TabTRep ;
begin
  inherited;
TITREREPORTB.Left:=TitreColCpt.Left ;
TITREREPORTB.Width:=TitreColCpt.Width+TitreColLibelle.Width+1 ;
Case QuelReportBAL(CritEdt.Bal.FormatPrint.Report,MReport) Of
  1 : TITREREPORTB.Caption:=MsgLibel.Mess[7]  ;
  2 : TITREREPORTB.Caption:=MsgLibel.Mess[8]+' '+StReportSec ;
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

procedure TFQRBLSEGE.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr11S_AXE:=TStringField(Q.FindField('S_AXE'));
Qr11S_SECTION:=TStringField(Q.FindField('S_SECTION'));
Qr11S_LIBELLE:=TStringField(Q.FindField('S_LIBELLE'));
Qr11S_SECTIONTRIE:=TStringField(Q.FindField('S_SECTIONTRIE'));
if CritEdt.Rupture=rCorresp then
   BEGIN
   Qr11S_CORRESP1         :=TStringField(Q.FindField('S_CORRESP1'));
   Qr11S_CORRESP2         :=TStringField(Q.FindField('S_CORRESP2'));
   END ;
end;

procedure TFQRBLSEGE.QGeneAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr12G_GENERAL:=TStringField(QGene.FindField('G_GENERAL'));
Qr12G_LIBELLE:=TStringField(QGene.FindField('G_LIBELLE'));

end;

procedure TFQRBLSEGE.FNatureCptChange(Sender: TObject);
begin
  inherited;
If QRLoading then Exit ;
FGeneral1.Clear ;  FGeneral2.Clear ; FJokerGene.Clear ;
if FPlansCo.checked then
   BEGIN CorrespToCombo(FPlanRuptures,AxeToFb(FNatureCpt.Value)) ; Exit ; END ;
case FNatureCpt.ItemIndex of
  0 : BEGIN FGeneral1.ZoomTable:=tzGVentil1 ; FGeneral2.ZoomTable:=tzGVentil1 ; END ;
  1 : BEGIN FGeneral1.ZoomTable:=tzGVentil2 ; FGeneral2.ZoomTable:=tzGVentil2 ; END ;
  2 : BEGIN FGeneral1.ZoomTable:=tzGVentil3 ; FGeneral2.ZoomTable:=tzGVentil3 ; END ;
  3 : BEGIN FGeneral1.ZoomTable:=tzGVentil4 ; FGeneral2.ZoomTable:=tzGVentil4 ; END ;
  4 : BEGIN FGeneral1.ZoomTable:=tzGVentil5 ; FGeneral2.ZoomTable:=tzGVentil5 ; END ;
  End ;
Case FNatureCpt.ItemIndex of
   0 : FPlanRuptures.Datatype:='ttRuptSect1' ;
   1 : FPlanRuptures.Datatype:='ttRuptSect2' ;
   2 : FPlanRuptures.Datatype:='ttRuptSect3' ;
   3 : FPlanRuptures.Datatype:='ttRuptSect4' ;
   4 : FPlanRuptures.Datatype:='ttRuptSect5' ;
  end ;
If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;

end;

procedure TFQRBLSEGE.BNouvRechClick(Sender: TObject);
begin
  inherited;
If FJokerGene.Visible then FJokerGene.Text:='' ;

end;

procedure TFQRBLSEGE.FGeneral1Change(Sender: TObject);
Var AvecJoker : Boolean ;
begin
  inherited;
AvecJoker:=Joker(FGeneral1, FGeneral2, FJokerGene ) ;
TFaS.Visible:=Not AvecJoker ;
TFGeneral.Visible:=Not AvecJoker ;
TFJokerGene.Visible:=AvecJoker ;
end;

procedure TFQRBLSEGE.BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.Rupture<>rRien) ;
end;

procedure TFQRBLSEGE.DLMultiNeedData(var MoreData: Boolean);
Var TotRupt : Array[0..12] of Double ;
    CodRupt, LibRupt, Lib1, CptRupt, Stcode, CodeRupture : String ;
    QuelleRupt : Integer ;
    SumD, SumC, TotSoldeD, TotSoldeC : Double ;
    Col : TColor ;
    OkOk, DansTotal, AddTotEdt : Boolean ;
    LibRuptInf : Array[1..10] Of TRuptInf ;
begin
  inherited;
MoreData:=False ; Lib1:='' ;
TotSoldeD:=0 ; TotSoldeC:=0 ; CodeRupture:='' ;
Case CritEdt.Rupture of
  rLibres   : BEGIN
              MoreData:=PrintGroupLibre(LRupt,Q,AxeToFb(CritEdt.Bal.Axe),CritEdt.LibreTrie,CodRupt,LibRupt,Lib1,TotRupt,Quellerupt,Col,LibRuptInf) ;
              If MoreData then
                 BEGIN
                 StCode:=CodRupt ; CodeRupture:=CodRupt ;
                 Delete(StCode,1,Quellerupt+2) ;
                 MoreData:=DansChoixCodeLibre(StCode,Q,AxeToFb(CritEdt.Bal.Axe),CritEdt.LibreCodes1,CritEdt.LibreCodes2, CritEdt.LibreTrie) ;
                 BRupt.Font.Color:=Col ;
                 END ;
              END ;
  rRuptures : BEGIN
              MoreData:=PrintRupt(LRupt,Qr11S_SECTIONTRIE.AsString,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt) ;
              CodeRupture:=CodRupt ;
              END ;
  rCorresp  : BEGIN
              OkOk:=TRUE ;
              Case CritEDT.Bal.PlansCorresp  Of
                1 : If Qr11S_CORRESP1.AsString<>'' Then CptRupt:=Qr11S_CORRESP1.AsString+Qr11S_SECTION.AsString
                                                   Else CptRupt:='.'+Qr11S_SECTION.AsString ;
                2 : If Qr11S_CORRESP2.AsString<>'' Then CptRupt:=Qr11S_CORRESP2.AsString+Qr11S_SECTION.AsString
                                                   Else CptRupt:='.'+Qr11S_SECTION.AsString ;
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
   HTCodRupt.Width:=TCodRupt.Width ;
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

   if ((CritEdt.Rupture=rLibres) Or (CritEdt.Rupture=rRuptures)) And CritEdt.Bal.AvecCptSecond then
      BEGIN
      StRappelLibTableLibre:=Lib1 ;
      HTCodRupt.Width:=TCodRupt.Width ;
      HTCodRupt.Caption  :=TCodRupt.Caption ;
      HCumDebRupt.Caption:=CumDebRupt.Caption ;
      HCumCreRupt.Caption:=CumCreRupt.Caption ;
      HDetDebRupt.Caption:=DetDebRupt.Caption ;
      HDetCreRupt.Caption:=DetCreRupt.Caption ;
      HSumDebRupt.Caption:=SumDebRupt.Caption ;
      HSumCreRupt.Caption:=SumCreRupt.Caption ;
      HSolDebRupt.Caption:=SolDebRupt.Caption ;
      HSolCreRupt.Caption:=SolCreRupt.Caption ;
      END ;

   if CritEdt.Bal.RuptOnly=Sur then
      BEGIN
      TCodRupt.Tag:=1 ;
      AddTotEdt:=False ;
      if (CritEdt.Rupture=rLibres) then AddTotEdt:=False else
      if (CritEdt.Rupture=rRuptures) or (CritEdt.Rupture=rCorresp) then AddTotEdt:=DansTotal ;
      if AddTotEdt then CalcTotalEdt(TotRupt[1], TotRupt[2], TotRupt[3], TotRupt[4]) ;
      END Else TCodRupt.Tag:=2 ;
   END ;
OkEnTeteRecap:=MoreData ;

if MoreData And ((CritEdt.Rupture=rLibres) Or (CritEdt.Rupture=rRuptures)) And CritEdt.Bal.AvecCptSecond then
   If (Not IsRuptDetail(LRupt,CodeRupture)) Then OkEnTeteRecap:=FALSE ;
end;

procedure TFQRBLSEGE.FSansRuptClick(Sender: TObject);
begin
  inherited;
FLigneRupt.Enabled:=Not FSansRupt.Checked ;
FLigneRupt.checked:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
if Not FSurRupt.Checked then FAvecCptSecond.Checked:=False ;
FAvecCptSecond.Enabled:=FSurRupt.Checked ;
//FAvecCptSecond.Enabled:=FALSE ;
FRupturesClick(Nil) ;
end;

procedure TFQRBLSEGE.FPlanRupturesChange(Sender: TObject);
begin
  inherited;
if FPlansCo.Checked then CorrespToCodes(FPlanRuptures,FCodeRupt1,FCodeRupt2) Else
if FRuptures.Checked then RuptureToCodes(FPlanRuptures,FCodeRupt1,FCodeRupt2,AxeToFb(FNatureCpt.Value)) ;
FCodeRupt1.ItemIndex:=0 ; FCodeRupt2.ItemIndex:=FCodeRupt2.Items.Count-1 ;
end;

procedure TFQRBLSEGE.FRupturesClick(Sender: TObject);
begin
  inherited;
If FPlansCo.Checked then
   BEGIN
   FGroupRuptures.Caption:=' '+MsgLibel.Mess[15] ;
   CorrespToCombo(FPlanRuptures,AxeToFb(FNatureCpt.Value)) ;
   END Else
   BEGIN
   FPlanRuptures.Reload ; If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
   FGroupRuptures.Caption:=' '+MsgLibel.Mess[14] ;
   END ;
If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
FAvecCptSecond.Enabled:=FSurRupt.Checked ;
//FAvecCptSecond.Enabled:=FALSE ;
If Not FSurRupt.Checked Then FAvecCptSecond.checked:=FALSE ;
end;


procedure TFQRBLSEGE.DLRecapNeedData(var MoreData: Boolean);
var TotRecap         : Array[0..77] of Double ;
    LibRecap, CodRecap : String ;
    SumD, SumC, TotSoldeD, TotSoldeC : Double ;
begin
  inherited;
MoreData:=False ;
TotSoldeD:=0 ; TotSoldeC:=0 ;
if (CritEdt.Rupture=rLibres) Or (CritEdt.Rupture=rRuptures)then
   BEGIN
   if CritEdt.Bal.AvecCptSecond then
      BEGIN
      MoreData:=PrintRecap(LRecap,CodRecap,LibRecap,TotRecap) ;
      if MoreData then
         BEGIN
         { Affichage sur la Bande }
         RG_GENERAL.Caption:=CodRecap ;
         If (CritEdt.Rupture=rRuptures) Then RG_LIBELLE.Caption:=LibRecap Else RG_LIBELLE.Caption:=LibRecap+' ('+StRappelLibTableLibre+')' ;
         SumD:=TotRecap[1]+TotRecap[3] ;
         SumC:=TotRecap[2]+TotRecap[4] ;
         (*
         TCodRupt.Width:=TitreColCpt.Width+TitreColLibelle.Width+1 ;
         TCodRupt.Caption:='' ;
         if CritEdt.Rupture=rLibres then
            BEGIN
            insert(MsgLibel.Mess[12]+' ',CodRupt,Quellerupt+2) ;
            TCodRupt.Caption:=CodRupt+' '+Lib1 ;
            END Else TCodRupt.Caption:=CodRupt+'   '+LibRupt ;
         *)
         CalCulSolde(SumD,SumC,TotSoldeD,TotSoldeC) ;
         RCumDebRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotRecap[1], CritEDT.AfficheSymbole) ;
         RCumCreRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotRecap[2], CritEDT.AfficheSymbole) ;
         RDetDebRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotRecap[3], CritEDT.AfficheSymbole) ;
         RDetCreRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotRecap[4], CritEDT.AfficheSymbole) ;
         RSumDebRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumD  , CritEDT.AfficheSymbole) ;
         RSumCreRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumC  , CritEDT.AfficheSymbole) ;
         RSolDebRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotSoldeD, CritEDT.AfficheSymbole) ;
         RSolCreRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotSoldeC, CritEDT.AfficheSymbole) ;
         END ;
      END ;
   END ;
//OkEnTeteRecap:=MoreData ;
end;

procedure TFQRBLSEGE.BRuptAfterPrint(BandPrinted: Boolean);
begin
  inherited;
if (CritEdt.Rupture=rLibres) Or (CritEdt.Rupture=rRuptures) then
   if CritEdt.Bal.AvecCptSecond then BEGIN VideRecap(LRecap) ; ChargeRecap(LRecap) ; END ;
end;

procedure TFQRBLSEGE.BRecapBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
if ((CritEdt.Rupture=rLibres) Or (CritEdt.Rupture=rRuptures)) And (CritEdt.Bal.AvecCptSecond)
   then Quoi:=QuoiCptRecap(RG_GENERAL.Caption,RG_LIBELLE.Caption) ;
end;

procedure TFQRBLSEGE.DLHRecapNeedData(var MoreData: Boolean);
begin
  inherited;
If Not CritEdt.Bal.AvecCptSecond Then BEGIN MoreData:=FALSE ; Exit ;END ;
MoreData:=OkEnTeteRecap ;
OkEnTeteRecap:=False ;
end;

end.
