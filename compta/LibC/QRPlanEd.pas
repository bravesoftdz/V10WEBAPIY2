unit QRPlanEd;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, HSysMenu, Menus, hmsgbox, HQuickrp, DB, DBTables, StdCtrls, Buttons,
  Hctrls, ExtCtrls, Mask, Hcompte, ComCtrls, Ent1, UtilEdt, CritEdt, CpteUtil, Hent1,
  UtilEdt1,CalcOle,SaisUtil,QRRupt, HTB97, Spin ;

procedure PlanEdition ;

type
  TFQRPlanEd = class(TFQR)
    FPlanEdt: THValComboBox;
    HLabel3: THLabel;
    TFGeneJoker: THLabel;
    TFGene: THLabel;
    TFaB: TLabel;
    FGeneJoker: TEdit;
    FGene1: THCpteEdit;
    FGene2: THCpteEdit;
    FRESOL: THValComboBox;
    TFRESOL: THLabel;
    FChoixMontant: THValComboBox;
    TFChoixMontant: THLabel;
    Bevel3: TBevel;
    RResol: TQRLabel;
    TRResol: TQRLabel;
    TitreColCpt: TQRLabel;
    Col15: TQRLabel;
    Col2: TQRLabel;
    Col3: TQRLabel;
    Col16: TQRLabel;
    Col4: TQRLabel;
    Col17: TQRLabel;
    Col18: TQRLabel;
    Col5: TQRLabel;
    Col6: TQRLabel;
    Col19: TQRLabel;
    Col20: TQRLabel;
    Col7: TQRLabel;
    Col8: TQRLabel;
    Col21: TQRLabel;
    Col9: TQRLabel;
    Col22: TQRLabel;
    Col10: TQRLabel;
    Col23: TQRLabel;
    Col11: TQRLabel;
    Col24: TQRLabel;
    Col1: TQRLabel;
    Col14: TQRLabel;
    Col12: TQRLabel;
    Col25: TQRLabel;
    Col13: TQRLabel;
    RTotalGen: TQRLabel;
    RTotEdtRea: TQRLabel;
    RTotEdtBud: TQRLabel;
    TotGenBud1: TQRLabel;
    TotGenRea1: TQRLabel;
    TotGenRea2: TQRLabel;
    TotGenBud2: TQRLabel;
    TotGenBud3: TQRLabel;
    TotGenRea3: TQRLabel;
    TotGenRea4: TQRLabel;
    TotGenBud4: TQRLabel;
    TotGenBud5: TQRLabel;
    TotGenRea5: TQRLabel;
    TotGenRea6: TQRLabel;
    TotGenBud6: TQRLabel;
    TotGenRea7: TQRLabel;
    TotGenBud7: TQRLabel;
    TotGenRea8: TQRLabel;
    TotGenBud8: TQRLabel;
    TotGenRea9: TQRLabel;
    TotGenBud9: TQRLabel;
    TotGenRea10: TQRLabel;
    TotGenBud10: TQRLabel;
    TotGenRea11: TQRLabel;
    TotGenBud11: TQRLabel;
    TotGenRea12: TQRLabel;
    TotGenBud12: TQRLabel;
    TotGenRea13: TQRLabel;
    TotGenBud13: TQRLabel;
    BSecondaire: TQRBand;
    MontantBud5: TQRLabel;
    MontantBud4: TQRLabel;
    MontantBud2: TQRLabel;
    MontantBud3: TQRLabel;
    MontantBud6: TQRLabel;
    MontantBud1: TQRLabel;
    MontantBud7: TQRLabel;
    MontantBud8: TQRLabel;
    MontantBud9: TQRLabel;
    IniGenBud: TQRLabel;
    RevGenRea: TQRLabel;
    MontantRea1: TQRLabel;
    MontantRea2: TQRLabel;
    MontantRea3: TQRLabel;
    MontantRea4: TQRLabel;
    MontantRea5: TQRLabel;
    MontantRea6: TQRLabel;
    MontantRea7: TQRLabel;
    MontantRea8: TQRLabel;
    MontantRea9: TQRLabel;
    MontantRea10: TQRLabel;
    MontantBud10: TQRLabel;
    MontantRea11: TQRLabel;
    MontantBud11: TQRLabel;
    MontantRea12: TQRLabel;
    MontantBud12: TQRLabel;
    MontantRea13: TQRLabel;
    MontantBud13: TQRLabel;
    FDATEDEF: TEdit;
    BChoixPer: TBitBtn;
    FNbDates: TEdit;
    FLISTEDATES: TEdit;
    MsgBox: THMsgBox;
    HLabel9: THLabel;
    COMPTEH: TQRLabel;
    LIBELLEH: TQRLabel;
    BFPrimaire: TQRBand;
    COMPTEB: TQRLabel;
    TotEdtRea1: TQRLabel;
    TotEdtRea2: TQRLabel;
    TotEdtRea3: TQRLabel;
    TotEdtRea4: TQRLabel;
    TotEdtRea5: TQRLabel;
    TotEdtRea6: TQRLabel;
    TotEdtRea7: TQRLabel;
    TotEdtRea8: TQRLabel;
    TotEdtRea9: TQRLabel;
    TotEdtBud1: TQRLabel;
    TotEdtBud2: TQRLabel;
    TotEdtBud3: TQRLabel;
    TotEdtBud4: TQRLabel;
    TotEdtBud5: TQRLabel;
    TotEdtBud6: TQRLabel;
    TotEdtBud7: TQRLabel;
    TotEdtBud8: TQRLabel;
    TotEdtBud9: TQRLabel;
    TotCpteRea: TQRLabel;
    TotCpteBud: TQRLabel;
    TotEdtRea10: TQRLabel;
    TotEdtBud10: TQRLabel;
    TotEdtRea11: TQRLabel;
    TotEdtBud11: TQRLabel;
    TotEdtRea12: TQRLabel;
    TotEdtBud12: TQRLabel;
    TotEdtRea13: TQRLabel;
    TotEdtBud13: TQRLabel;
    Col26: TQRLabel;
    HM: THMsgBox;
    DLGen: TQRDetailLink;
    BRecap: TQRBand;
    TotRecapRea: TQRLabel;
    TotRRecap1: TQRLabel;
    TotRRecap2: TQRLabel;
    TotRRecap3: TQRLabel;
    TotRRecap4: TQRLabel;
    TotRRecap5: TQRLabel;
    TotRRecap6: TQRLabel;
    TotRRecap7: TQRLabel;
    TotRRecap8: TQRLabel;
    TotRRecap9: TQRLabel;
    TotRRecap10: TQRLabel;
    TotRRecap11: TQRLabel;
    TotRRecap12: TQRLabel;
    TotRRecap13: TQRLabel;
    BFAnaGene: TQRBand;
    TotAnaGene1: TQRLabel;
    TotAnaGene2: TQRLabel;
    TFAnaGene: TQRLabel;
    TotAnaGene3: TQRLabel;
    TotAnaGene4: TQRLabel;
    TotAnaGene5: TQRLabel;
    TotAnaGene6: TQRLabel;
    TotAnaGene7: TQRLabel;
    TotAnaGene8: TQRLabel;
    TotAnaGene9: TQRLabel;
    TotAnaGene10: TQRLabel;
    TotAnaGene11: TQRLabel;
    TotAnaGene12: TQRLabel;
    TotAnaGene13: TQRLabel;
    TRPlanEdt: TQRLabel;
    RPlanEdt: TQRLabel;
    PD_NATUREEDT: TEdit;
    FDetail1: TCheckBox;
    FDetail2: TCheckBox;
    FQueTotRea: TCheckBox;
    FDetail3: TCheckBox;
    FQuelSPlan: TSpinEdit;
    HLabel10: THLabel;
    DLAnaGene: TQRDetailLink;
    DLRecap: TQRDetailLink;
    FRecap: TCheckBox;
    QROVERLAY: TQRBand;
    Trait8: TQRLigne;
    Trait0: TQRLigne;
    Trait4: TQRLigne;
    Trait3: TQRLigne;
    Trait2: TQRLigne;
    Trait1: TQRLigne;
    Trait5: TQRLigne;
    Trait6: TQRLigne;
    Trait7: TQRLigne;
    Trait9: TQRLigne;
    QRLigne1: TQRLigne;
    Ligne1: TQRLigne;
    QRLigne2: TQRLigne;
    QRLigne3: TQRLigne;
    QRLigne4: TQRLigne;
    QRLigne5: TQRLigne;
    QRLigne6: TQRLigne;
    procedure FormShow(Sender: TObject);
    procedure FPlanEdtChange(Sender: TObject);
    procedure FGene1Change(Sender: TObject);
    procedure BChoixPerClick(Sender: TObject);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFPrimaireBeforePrint(var PrintBand: Boolean;
      var Quoi: string);
    procedure BSecondaireBeforePrint(var PrintBand: Boolean;
      var Quoi: string);
    procedure BHCptBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean;
      var Quoi: string);
    procedure BHRecapBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure DLAnaGeneNeedData(var MoreData: Boolean);
    procedure PD_NATUREEDTChange(Sender: TObject);
    procedure BMultiBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BRecapBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFAnaGeneBeforePrint(var PrintBand: Boolean;
      var Quoi: String);
    procedure DLRecapNeedData(var MoreData: Boolean);
    procedure BRecapAfterPrint(BandPrinted: Boolean);
    procedure FRecapClick(Sender: TObject);
  private
    TotEdt : TabTot ;
    TotEdtRea,TotEdtBud          : TabTot12 ;
    TotCptRea,TotCptBud          : TabTot12 ;
    QrLD_RUBRIQUE,QrLD_TYPEDETAIL,QrLD_POURBUD,QrLD_TOTALISATION,QrLD_FORMAT : TStringField ;
    AxeRub          : String3 ;
    FamRub          : String ;
    NbColAff        : integer ;
    ListeAnaGene    : TStringList ;
    ListeTotAnaGene : TStringList ;
    LRecap  : TStringList ;
    OkEnTeteRecap   : boolean ;
    Function  QuoiCpt(i : Integer) : String ;
    procedure GenereSQL ; Override ;
    Procedure PlanEdtZoom(Quoi : String) ;
    Procedure ChargeCbPlan ;
    procedure InitDivers ; Override ;
    Procedure FinirPrint ; override ;
    procedure RecupCritEdt ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    function  CritOk : Boolean ; Override ;
    procedure AlimMontant(Var TotBud : TMontTot ; Var TotMulti : TabMont77) ;
    procedure AfficheBudOuRea(QRB : TQRBand) ;
    function  InfosRubrique : string ;
    procedure MontantAnaGene(Indice  : Integer) ;
    function  PrintAnaGene (var Cpt1,cpt2 : String17 ; var Tot : TabMont77) : boolean ;
    procedure RempliLesDates ;
    Procedure PositionnePlan ;
    procedure ReTrieDetail ;
    procedure AddDetail(Liste : TStringList ; Cpt : String ; FTotAG : TFTotAnaGene) ;
    procedure MajListeTot(Liste : TStringList) ;
    procedure AddRecapSections ;
  public
    { Déclarations publiques }
  end;

implementation

uses LigneEdt,DateEdt ;

const MaxTabCol    = 15 ;

{$R *.DFM}

procedure PlanEdition ;
var QR : TFQRPlanEd ;
    Edition : TEdition ;
BEGIN
QR:=TFQRPlanEd.Create(Application) ;
Edition.Etat:=etBudMulti ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.PlanEdtZoom ;
 //QR.QuelEtatBud:=Normal ;
 QR.InitType (nbBudSec,nePlanEdt,msSecBAna,'TFQRPlanEd','',TRUE,FALSE,TRUE,Edition) ;
 QR.ShowModal ;
 finally
 QR.Free ;
 end ;
SourisNormale ;
END ;

Procedure TFQRPlanEd.PlanEdtZoom(Quoi : String) ;
Var Lp,i: Integer ;
BEGIN
Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
i:=StrToInt(Copy(Quoi,Lp+1,1)) ;
ZoomEdt(i,Quoi) ;
END ;

Function TFQRPlanEd.QuoiCpt(i : Integer) : String ;
BEGIN
Result:='' ;
Case i of
   0 : Result:=QrLD_RUBRIQUE.AsString+' '+InfosRubrique+'@'+'9';
  End ;
END ;

Procedure TFQRPlanEd.ChargeCbPlan ;
Var QLoc : TQuery ;
    X : TInfoPlanEdt ;
BEGIN
FPlanEdt.Values.Clear ; FPlanEdt.Items.Clear ;
QLoc:=OpenSql('Select * From PLANEDT',True) ;
While Not QLoc.Eof do
   BEGIN
   X:=TInfoPlanEdt.Create ;
   X.Nat:=QLoc.FindField('PD_NATUREEDT').AsString ;
   X.Fam:=QLoc.FindField('PD_FAMILLES').AsString ;
   X.Jal:=QLoc.FindField('PD_BUDJAL').AsString ;
   FPlanEdt.Values.AddObject(QLoc.FindField('PD_PLANEDT').AsString,X) ;
   FPlanEdt.Items.Add(QLoc.FindField('PD_LIBELLE').AsString) ;
   QLoc.Next ;
   END ;
Ferme(QLoc) ;
if FPlanEdt.Values.Count>0 then FPlanEdt.Value:=FPlanEdt.Values[0] ;
END ;

procedure TFQRPlanEd.FormShow(Sender: TObject);
begin
ChargeCbPlan ;
InitResolution(FRESOL) ;
  inherited;
TabSup.TabVisible:=False ;
FChoixMontant.ItemIndex:=0 ;
FSelectCpte.Visible:=False ; TSelectCpte.Visible:=False ;
FQuelSPlan.MaxValue:=MaxSousPlan ;
end;

function TFQRPlanEd.InfosRubrique : string ;
Var QLoc : TQuery ;
    Sql,StAnd : String ;
BEGIN
Result:='' ;
if FPLANEDT.Value='' then Exit ;
if CritEdt.PlanEdt.Famille='BUD' then StAnd:='And RB_NATRUB="'+CritEdt.PlanEdt.Nature+'"'
                                 else StAnd:='' ;
Sql:='Select RB_LIBELLE,RB_AXE,RB_FAMILLES from RUBRIQUE Where RB_RUBRIQUE="'+QrLD_RUBRIQUE.AsString+'"';
QLoc:=OpenSql(Sql,True) ;
if not QLoc.Eof then Result:=QLoc.Fields[0].AsString ;
AxeRub:=QLoc.Fields[1].AsString ;
FamRub:=QLoc.Fields[2].AsString ;
Ferme(QLoc) ;
END ;

procedure TFQRPlanEd.FPlanEdtChange(Sender: TObject);
Var Q       : TQuery ;
    T       : TInfoPlanEdt ;
    Edition : TEdition ;
begin
if QRLoading then Exit ;
T:=TInfoPlanEdt(FPlanEdt.Values.Objects[FPlanEdt.ItemIndex]) ;
Edition.Etat:=etBudMulti ;
if T.Nat='CPT' then
  BEGIN
  InitType (nbGen,nePlanEdt,msGenEcr,'TFQRPlanEd','',TRUE,FALSE,TRUE,Edition) ;
  FQueTotRea.Checked:=True ; FQueTotRea.Enabled:=False ;
  FDetail3.Checked:=False  ; FDetail3.Enabled:=False ;
  FDetail2.Checked:=False  ; FDetail2.Enabled:=False ;
  FDetail1.Checked:=False  ; FDetail1.Enabled:=False ;
  FRecap.checked:=False ; FRecap.Enabled:=False ; FQuelSPlan.Enabled:=False ;
  END else
  BEGIN
  FQueTotRea.Checked:=False ; FQueTotRea.Enabled:=True ;
  FDetail3.State:=cbGrayed  ; FDetail3.Enabled:=True ;
  FDetail2.State:=cbGrayed  ; FDetail2.Enabled:=True ;
  FDetail1.State:=cbGrayed  ; FDetail1.Enabled:=True ;
  FRecap.Enabled:=True ; FQuelSPlan.Enabled:=FRecap.Checked ;
  if (T.Fam='CBG') then InitType (nbBudGen,nePlanEdt,msGenBEcr,'TFQRPlanEd','',TRUE,FALSE,TRUE,Edition) else
    if (T.Fam='G/S') then InitType (nbBudGen,nePlanEdt,msGenBAna,'TFQRPlanEd','',TRUE,FALSE,TRUE,Edition) else
      if (T.Fam='CBS') then InitType (nbBudSec,nePlanEdt,msRien,'TFQRPlanEd','',TRUE,FALSE,TRUE,Edition) else
        if (T.Fam='S/G') then InitType (nbBudSec,nePlanEdt,msSecBAna,'TFQRPlanEd','',TRUE,FALSE,TRUE,Edition) ;
  END ;
Q:=OpenSQL('SELECT PD_DATEDEF,PD_NBDATEDEF FROM PLANEDT WHERE PD_PLANEDT="'+FPlanEdt.Value+'" AND PD_NATUREEDT="'+T.Nat+'"',True) ;
FDATEDEF.Text:=Q.Fields[0].AsString ;
FNBDATES.Text:=IntToStr(Q.Fields[1].AsInteger) ;
PD_NATUREEDT.Text:=T.Nat ;
Ferme(Q) ;
RempliLesDates ;
end;

procedure TFQRPlanEd.RempliLesDates ;
var i : integer ;
    Per,St,StD1,StD2 : String ;
BEGIN
i:=1 ; St:=FDateDef.Text ;
FListeDates.Text:='' ; NbColAff:=0 ;
while i<=23 do
  BEGIN
  Per:=ReadTokenSt(St) ;
  if Per<>'' then
    BEGIN
    GetCorrespDates(Per,StD1,StD2) ;
    FListeDates.Text:=FListeDates.Text+StD1+';'+StD2+';' ;
    Inc(NbColAff) ;
    END ;
  Inc(i) ;
  END ;
END ;

procedure TFQRPlanEd.InitDivers ;
var St,S,SD :  string ;
    i,NumM  : integer ;
    DD      : TDateTime ;
BEGIN
Inherited ;
{ Calcul des différentes fourchettes de dates }
//CalculDateTiers(nePlanEdt,Age,CritEdt,TabDate,FP1,FP2,FP3,FP4,FP5,FP6,FP7,FP8) ;
{ Titres des colonnes de montants }
BFPrimaire.Frame.DrawBottom:=CritEdt.PlanEdt.FormatPrint.PrSepCompte[1] ;
//BFSecondaire.Frame.DrawBottom:=CritEdt.PlanEdt.FormatPrint.PrSepCompte[1] ;
//BFSecondaire.Frame.DrawTop:=CritEdt.PlanEdt.FormatPrint.PrSepCompte[1] ;
Fillchar(TotEdtRea,SizeOf(TotEdtRea),#0) ; Fillchar(TotEdtBud,SizeOf(TotEdtBud),#0) ;
Fillchar(TotCptRea,SizeOf(TotCptRea),#0) ; Fillchar(TotCptBud,SizeOf(TotCptBud),#0) ;
Fillchar(CritEdt.PlanEdt.PeriodeDates,SizeOf(CritEdt.PlanEdt.PeriodeDates),#0) ;
AxeRub:='A1' ;
NbColAff:=StrToInt(FNbDates.Text) ;
S:=FListeDates.Text ;
SD:=ReadTokenSt(S) ; i:=1 ;St:='' ; NumM:=0 ;
While SD<>'' do
  BEGIN
  DD:=StrToDate(SD) ;
  if (i mod 2=0) then
    BEGIN
    // Fin de période ou date de fin
    NumM:=1 ;
    if FINDEMOIS(DD)=DD then BEGIN SD:=FormatDateTime('mmm yyyy',DD) ; NumM:=4 ; END ;
    END else
    BEGIN
    // Début de période ou date de début
    if DEBUTDEMOIS(DD)=DD then BEGIN SD:=FormatDateTime('mmm yyyy',DD) ; NumM:=3 ; END ;
    END ;
  TQRLabel(FindComponent('Col'+IntToStr(i))).Caption:=MsgBox.Mess[NumM]+' '+SD ;
  CritEdt.PlanEdt.PeriodeDates[i]:=DD ;
  SD:=ReadTokenSt(S) ; inc(i) ;
  END ;
CritEdt.PlanEdt.FormatPrint.ColMin:=1 ;
CritEdt.PlanEdt.FormatPrint.ColMax:=(i div 2)+2 ;
BRecap.ForceNewPage:=CritEdt.PlanEdt.FormatPrint.PrSepCompte[10] ;
END ;

Procedure TFQRPlanEd.FinirPrint ;
BEGIN
  inherited ;
if ListeAnaGene<>nil then BEGIN ListeAnaGene.Clear ; ListeAnaGene.Free ; END ;
if ListeTotAnaGene<>nil then BEGIN ListeTotAnaGene.Clear ; ListeTotAnaGene.Free ; END ;
if FComposite or (CritEdt.PlanEdt.Famille='CBS') then VideRecap(LRecap) ;
END ;

procedure TFQRPlanEd.RecupCritEdt ;
var i : integer ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  SJoker:=FGeneJoker.Visible ;
  If SJoker Then BEGIN SCpt1:=FGeneJoker.Text ; SCpt2:=FGeneJoker.Text ; END
            Else BEGIN
                 SCpt1:=FGene1.Text  ; SCpt2:=FGene2.Text ;
                 PositionneFourchetteSt(FGene1,FGene2,CritEdt.LSCpt1,CritEdt.LSCpt2) ;
                 END ;
  PlanEdt.Resol:=FResol.Value[1] ;
  Monoexo:=False ;
  //PlanEdt.Taux:=StrToFloat(FTaux.Text) ;
  //PlanEdt.Qte:=FQte.Checked ;
  //PlanEdt.AnaGene:=FAnaGene.Checked ;
  { Rupture }
  PlanEdt.QueTotalRea:=FQueTotRea.Checked ;
// Si variable dir debit pos ou crédit Pos
  //DebitPos:=(FDebitPos.ItemIndex=0) ;
  PlanEdt.Journal:=TInfoPlanEdt(FPLANEDT.Values.Objects[FPLANEDT.ItemIndex]).Jal ;
  PlanEdt.Nature:=TInfoPlanEdt(FPLANEDT.Values.Objects[FPLANEDT.ItemIndex]).Nat ;
  PlanEdt.Famille:=TInfoPlanEdt(FPLANEDT.Values.Objects[FPLANEDT.ItemIndex]).Fam ;

  PlanEdt.PasDeRubrique:=False ; //Not FAvecRub.Checked ;
  PlanEdt.SoldeFormate:=FChoixMontant.Value ;
  case FDetail1.State of
    cbgrayed    : PlanEdt.Detail1:='' ;
    cbChecked   : PlanEdt.Detail1:='X' ;
    cbUnChecked : PlanEdt.Detail1:='-' ;
    end ;
  case FDetail2.State of
    cbgrayed    : PlanEdt.Detail2:='' ;
    cbChecked   : PlanEdt.Detail2:='X' ;
    cbUnChecked : PlanEdt.Detail2:='-' ;
    end ;
  case FDetail3.State of
    cbgrayed    : PlanEdt.Detail3:='' ;
    cbChecked   : PlanEdt.Detail3:='X' ;
    cbUnChecked : PlanEdt.Detail3:='-' ;
    end ;
  PlanEdt.FormatPrint.PrSepCompte[10]:=FRecap.Checked ;
  PlanEdt.SousPlan:=FQuelSPlan.Value ;
// Total
  TQRLabel(FindComponent('Col'+IntToStr((StrToInt(FNbDates.Text)*2)+1))).Caption:=MsgBox.Mess[2] ;
  TQRLabel(FindComponent('Col'+IntToStr((StrToInt(FNbDates.Text)*2)+2))).Caption:='' ;
//If DebitPos Then PlanEdt.SoldeFormate:='PD' Else PlanEdt.SoldeFormate:='PC' ;
  for i:=1 to 14 do
      BEGIN
      PlanEdt.FormatPrint.TabColl[i].OkAff:=(i<=(StrToInt(FNbDates.Text)+2)) ;
      PremTabColEdt[i].OkAff:=(i<=(StrToInt(FNbDates.Text)+2)) ;
      END ;

  END ;
END ;

procedure TFQRPlanEd.RenseigneCritere ;
BEGIN
Inherited ;
//CaseACocher(FQte,RQte) ;
//RChoixMontant.Caption:=FChoixMontant.Text ;
RResol.Caption:=FResol.Text ;
RPlanEdt.Caption:=FPlanEdt.Text ;
RDateCompta1.Caption:=DateToStr(CritEdt.PlanEdt.PeriodeDates[1]);
RDateCompta2.Caption:=DatetoStr(CritEdt.PlanEdt.PeriodeDates[NbColAff*2]);
//RTaux.Caption:=FTaux.Text ;
{
if CritEdt.SJoker then
   BEGIN
   TRGeneral.Caption:=MsgBox.Mess[5] ;
   RGeneral1.Caption:=CritEdt.SCpt1 ;
   END Else
   BEGIN
   TRGeneral.Caption:=MsgBox.Mess[4] ;
   RGeneral1.Caption:=CritEdt.LSCpt1 ; RGeneral2.Caption:=CritEdt.LScpt2 ;
   END ;
//General2.Visible:=Not CritEdt.SJoker ; TRaG.Visible:=Not CritEdt.SJoker ;
//RGene2.Visible:=Not CritEdt.SJoker ; TRaC.Visible:=Not CritEdt.SJoker ;
}
END ;

procedure TFQRPlanEd.ChoixEdition ;
{ Initialisation des options d'édition }
BEGIN
Inherited ;
//DLRupt.PrintBefore:=TRUE ;
CritEdt.Decimale:=InitDecimaleResol(CritEdt.PlanEdt.Resol,CritEdt.Decimale) ;
ChangeMask(Formateur,CritEdt.Decimale,CritEdt.Symbole);
CritEdt.FormatMontant:=Formateur.Masks.PositiveMask;
if FComposite or (CritEdt.PlanEdt.Famille='CBS') then ChargeRecap(LRecap) ;
//ChargeRecap(LRecapG) ;
END ;

function TFQRPlanEd.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ;
ListeAnaGene:=TStringList.Create ;
ListeTotAnaGene:=TStringList.Create ;
If Result Then Fillchar(TotEdt,SizeOf(Totedt),#0) ;
END ;

procedure TFQRPlanEd.FGene1Change(Sender: TObject);
Var AvecJokerS : Boolean ;
begin
  inherited;
AvecJokerS:=Joker(FGene1, FGene2, FGeneJoker) ;
TFaB.Visible:=Not AvecJokerS ;
TFGene.Visible:=Not AvecJokerS ;
TFGeneJoker.Visible:=AvecJokerS ;
end;

procedure TFQRPlanEd.BChoixPerClick(Sender: TObject);
var StDate,Dat,StPer,NbDat,Per,ListeDates : String ;
    NbD,NumPer : Integer ;
    OkDebPer,OkFinPer : boolean ;
begin
ListeDates:=FLISTEDATES.Text ;
StPer:=FDATEDEF.Text ;
Dat:=StPer ;
StDate:='' ; NbD:=0 ;
while Dat<>'' do
  BEGIN
  if (StDate='') then StDate:=ReadTokenSt(Dat) else ReadTokenSt(Dat) ;
  Inc(NbD) ;
  END ;
NbDat:=FNbDates.Text ;
DecodePeriode(Dat,Per,OkDebPer,OkFinPer,NumPer) ;
ChoixDateEdt(NbD,StPer,ListeDates) ;
FDATEDEF.Text:=StPer ; FNbDates.Text:=IntToStr(NbD) ;
FLISTEDATES.Text:=ListeDates ;
end;

procedure TFQRPlanEd.GenereSQL ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
//Inherited ;
FComposite:=Pos('/',CritEdt.PlanEdt.Famille)>0 ;
Q.Close ; Q.SQL.Clear ;
Q.SQL.Add(' Select LD_RUBRIQUE,LD_TYPEDETAIL,LD_POURBUD,LD_TOTALISATION,LD_FORMAT FROM LIGNEEDT') ;
Q.SQL.Add(' Where LD_PLANEDT="'+FPlanEdt.Value+'" AND LD_NATLIGNEEDT="'+CritEdt.PlanEdt.Nature+'" ORDER BY LD_CHRONO') ;
ChangeSQL(Q) ; Q.Open ;
end ;

procedure TFQRPlanEd.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
QrLD_RUBRIQUE:=TStringField(Q.FindField('LD_RUBRIQUE')) ;
QrLD_TYPEDETAIL:=TStringField(Q.FindField('LD_TYPEDETAIL')) ;
QrLD_POURBUD:=TStringField(Q.FindField('LD_POURBUD')) ;
QrLD_TOTALISATION:=TStringField(Q.FindField('LD_TOTALISATION')) ;
QrLD_FORMAT:=TStringField(Q.FindField('LD_FORMAT')) ;
end;

procedure TFQRPlanEd.MontantAnaGene(Indice  : Integer) ;
var i,j : Integer ;
    St,StG,StS : String ;
    MD,MC : Double ;
    Find : boolean ;
    FTotAG : TFTotAnaGene ;
BEGIN
if (ListeAnaGene=nil) then Exit ;
for i:=0 to ListeAnaGene.Count-1 do
  BEGIN
  Find:=False ;
  St:=ListeAnaGene[i] ;
  StS:=ReadTokenSt(St) ;
  // Simple : 1 seul compte
  if (St='') then BEGIN St:=StS ; StS:='' ; END ;
  StG:=Copy(St,1,Pos(':',St)-1) ; Delete(St,1,Pos(':',St)) ;
  Delete(St,1,Pos(':',St)) ;Delete(St,1,Pos(':',St)) ;Delete(St,1,Pos(':',St)) ;
  // TC et TD en 4 et 5
  MC:=Valeur(Copy(St,1,Pos(':',St)-1)) ; Delete(St,1,Pos(':',St)) ;
  MD:=Valeur(Copy(St,1,Pos(':',St)-1)) ; Delete(St,1,Pos(':',St)) ;
  for j:=0 to ListeTotAnaGene.Count-1 do
    BEGIN
    FTotAG:=TFTotAnaGene(ListeTotAnaGene.Objects[j]) ;
    if (FTotAG.Sect=StS) and (FTotAG.Gene=StG) then
      BEGIN
      Find:=True ;
      FTotAG.TotAG[Indice]:=FTotAG.TotAG[Indice]+MD ;
      FTotAG.TotAG[MaxTabCol+Indice]:=FTotAG.TotAG[MaxTabCol+Indice]+MC ;
      ListeTotAnaGene.Objects[j]:=FTotAG ;
      Break ;
      END ;
    END ;
  if not Find then
    BEGIN
    FTotAG:=TFTotAnaGene.Create ;
    FTotAG.Sect:=StS ; FTotAG.Gene:=StG ;
    Fillchar(FTotAG.TotAG,SizeOf(FTotAG.TotAG),#0) ;
    FTotAG.TotAG[Indice]:=MD ;
    FTotAG.TotAG[MaxTabCol+Indice]:=MC ;
    ListeTotAnaGene.AddObject('',FTotAG) ;
    END ;
  END ;
ListeAnaGene.Clear ;
END ;

procedure TFQRPlanEd.BDetailBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
Fillchar(TotCptRea,SizeOf(TotCptRea),#0) ; Fillchar(TotCptBud,SizeOf(TotCptBud),#0) ;
COMPTEH.Caption:=QrLD_RUBRIQUE.AsString ;
LIBELLEH.Caption:=InfosRubrique ;
Quoi:=QuoiCpt(0) ;
end;

procedure TFQRPlanEd.BFPrimaireBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
  var i : integer ;
begin
  inherited;
if FComposite and (QrLD_TOTALISATION.AsString='-') then BEGIN PrintBand:=False ; Exit ; END ;
AfficheBudOuRea(BFPrimaire) ;
For i:=1 To NbColAff Do
  BEGIN
  TQRLabel(FindComponent('TotEdtRea'+IntToStr(i))).Caption:=PrintSoldeFormate(TotCptRea[i].TotDebit, TotCptRea[i].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.PlanEdt.SoldeFormate) ;
  TQRLabel(FindComponent('TotEdtBud'+IntToStr(i))).Caption:=PrintSoldeFormate(TotCptBud[i].TotDebit, TotCptBud[i].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.PlanEdt.SoldeFormate) ;
  END ;
// totaux
TQRLabel(FindComponent('TotEdtRea'+IntToStr(NbColAff+1))).Caption:=PrintSoldeFormate(TotCptRea[0].TotDebit, TotCptRea[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.PlanEdt.SoldeFormate) ;
TQRLabel(FindComponent('TotEdtBud'+IntToStr(NbColAff+1))).Caption:=PrintSoldeFormate(TotCptBud[0].TotDebit, TotCptBud[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.PlanEdt.SoldeFormate) ;
CompteB.Caption:=MsgBox.Mess[2]+' '+QrLD_RUBRIQUE.AsString ;
Quoi:=QuoiCpt(0) ;
end;

procedure TFQRPlanEd.AddRecapSections ;
var i,j        : Integer ;
    FTotAG     : TFTotAnaGene ;
    Debut,Lg   : Integer ;
    Cod,St,Lib : String ;
    fb         : TFichierBase ;
begin
if not ((CritEdt.PlanEdt.Famille='S/G') or (CritEdt.PlanEdt.Famille='G/S') or (CritEdt.PlanEdt.Famille='CBS')) then Exit ;
if VH^.SousPlanAxe[fb,CritEdt.PlanEdt.SousPlan].ListeSP=nil then Exit ;
fb:=fbAxe1 ;
if AxeRub='A2' then fb:=fbAxe2 else
 if AxeRub='A3' then fb:=fbAxe3 else
  if AxeRub='A4' then fb:=fbAxe4 else
   if AxeRub='A5' then fb:=fbAxe5 ;
Debut:=VH^.SousPlanAxe[fb,CritEdt.PlanEdt.SousPlan].Debut ;
Lg:=VH^.SousPlanAxe[fb,CritEdt.PlanEdt.SousPlan].Longueur ;
for i:=0 to ListeTotAnaGene.Count-1 do
  BEGIN
  FTotAG := TFTotAnaGene(ListeTotAnaGene.Objects[i]) ;
  if (CritEdt.PlanEdt.Famille='S/G') or (CritEdt.PlanEdt.Famille='CBS') then Cod:=Copy(FTotAG.Sect,Debut,Lg)
   else if (CritEdt.PlanEdt.Famille='G/S') then Cod:=Copy(FTotAG.Gene,Debut,Lg) ;
  for j:=0 to VH^.SousPlanAxe[fb,CritEdt.PlanEdt.SousPlan].ListeSP.Count-1 Do
     BEGIN
     St:=VH^.SousPlanAxe[fb,CritEdt.PlanEdt.SousPlan].ListeSP[j] ;
     if ReadTokenSt(St)=Cod then BEGIN Lib:=ReadTokenSt(St) ; Break ; END ;
     END ;
  if   ((CritEdt.PlanEdt.Famille='S/G') and (FTotAG.Sect<>'') and (FTotAG.Gene<>''))
    or ((CritEdt.PlanEdt.Famille='CBS') and (FTotAG.Sect<>'') and (FTotAG.Gene=''))
    or ((CritEdt.PlanEdt.Famille='G/S') and (FTotAG.Gene<>'') and (FTotAG.Sect<>'')) then AddRecap(LRecap, [Cod], [Lib],FTotAG.TotAG) ;
  END ;
end;

procedure TFQRPlanEd.AlimMontant(Var TotBud : TMontTot ; Var TotMulti : TabMont77) ;
var i  : integer ;
    C,C1,C2 : TComponent ;
    TMR,TMB : TQrLABEL ;
    Compte1,Compte2,LExo : String ;
    TD      : TabloExt ;
    RealPer : TabTot12 ;
    OkMR,OkMB : Boolean ;
    XX : Double ;
    ListeDates : String ;
    TypeCumRea,TypeCumBud : String ;
BEGIN
Fillchar(TotBud,SizeOf(TotBud),#0) ;
Fillchar(RealPer,SizeOf(RealPer),#0) ;
TMB:=NIL ; TMR:=NIL ;
ListeDates:=FListeDates.Text ;
If CritEdt.PlanEdt.Nature='CPT' Then
  BEGIN
{  TypeCumRea:='TOG' ;//: Rubrique de totalisation. Calcul sur rubrique comptable
  TypeCumRea:='TOB' ;//: Rubrique de totalisation. Calcul sur rubrique budgétaire
  TypeCumRea:='TOR' ;//: Rubrique de totalisation. Calcul sur rubrique budgétaire réalisé
  TypeCumRea:='BUDGET:G/AREA' ;//: Calcul sur couple compte budgétaire/section budgétaire pour le réalisé
  TypeCumRea:='BUDGET:A/GREA' ;//: Calcul sur couple section budgétaire/compte budgétaire pour le réalisé
  TypeCumRea:='BUDGET:BGEREA' ;//: Calcul sur compte budgétaire pour le réalisé
  TypeCumRea:='BUDGET:BSEREA' ;//: Calcul sur section budgétaire pour le réalisé
}
  TypeCumRea:='RUB' ; TypeCumBud:='RUB' ;
  END else
  BEGIN
//  if LaFam='CBG' then
//  'CBS','S/G','G/S'
  TypeCumRea:='RUBREA' ; TypeCumBud:='RUBBUD' ;
  END ;
for i:=1 To NbColAff do
  BEGIN
  OkMR:=FALSE ; OkMB:=FALSE ;
  C:=FindComponent('Col'+IntToStr(i)) ;
  If C<>NIL Then
     BEGIN
     C1:=FindComponent('MontantRea'+IntToStr(i)) ;
     If C1<>NIL Then BEGIN TMR:=TQRLabel(C1) ; OkMR:=TRUE ; END ;
     C2:=FindComponent('MontantBud'+IntToStr(i)) ;
     If C2<>NIL Then BEGIN TMB:=TQRLabel(C2) ; OkMB:=TRUE ; END ;
     END Else continue ;
  If C<>NIL Then
     BEGIN
     If CritEdt.Monoexo Then LExo:=CritEdt.PlanEdt.Exo1 Else LExo:='' ;
     CritEdt.Date1:=CritEdt.PlanEdt.PeriodeDates[(2*i)-1] ;
     CritEdt.PlanEdt.Date21:=CritEdt.PlanEdt.PeriodeDates[(2*i)] ;
     If not CritEdt.Monoexo Then LExo:=QUELEXODT(CritEdt.Date1) ;
     Compte1:=QrLD_RUBRIQUE.AsString ;
      If CritEdt.PlanEdt.PasDeRubrique Then
        BEGIN
        //Compte1:=Qr1BS_AXE.AsString+QR1BS_BUDSECT.AsString ;
        //Compte2:=QR12BG_BUDGENE.AsString ;
        XX:=GetCumul(TypeCumRea,Compte1,Compte2,CritEdt.PlanEdt.SANbud,CritEdt.Etab,CritEdt.PlanEdt.Journal,LExo,CritEdt.Date1,CritEdt.PlanEdt.Date21,false,FALSE,ListeAnaGene,TD,FALSE) ;
        // 4 : TC ; 5 : TD
        TD[4]:=0 ; TD[5]:=0 ;
        If XX>0 Then TD[5]:=XX Else TD[4]:=Abs(XX) ;
        RealPer[0].TotDebit:=TD[4] ; RealPer[0].TotCredit:=TD[5] ;
        END else
        BEGIN
        //RubCpt:=QR12BG_RUB.AsString ;
        //RubSect:=QR1BS_RUB.AsString ;
        //Compte1:='S/G'+CritEdt.PlanEdt.Journal+RubSect+':'+RubCpt ;
        GetCumul(TypeCumRea,Compte1,'',CritEdt.PlanEdt.SANbud,CritEdt.Etab,CritEDT.DeviseAffichee,LExo,CritEdt.Date1,CritEdt.PlanEdt.Date21,false,FALSE,ListeAnaGene,TD,FALSE) ;
        RealPer[0].TotDebit:=TD[3] ; RealPer[0].TotCredit:=TD[2] ;
        END ;
     Reevaluation(RealPer[0].TotDebit,RealPer[0].TotCredit,CritEdt.PlanEdt.Resol,CritEdt.PlanEdt.Taux) ;
     RealPer[2].TotDebit:=RealPer[2].TotDebit+RealPer[0].TotDebit ; RealPer[2].TotCredit:=RealPer[2].TotCredit+RealPer[0].TotCredit ;
     MontantAnaGene(i) ;
     If OkMR Then TMR.Caption:=PrintSoldeFormate(RealPer[0].TotDebit, RealPer[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.PlanEdt.SoldeFormate) ;
     TotMulti[(MaxTabCol*2)+i]:= Arrondi(TotMulti[(MaxTabCol*2)+i]+RealPer[0].TotDebit, CritEdt.Decimale) ;
     TotMulti[(MaxTabCol*3)+i]:= Arrondi(TotMulti[(MaxTabCol*3)+i]+RealPer[0].TotCredit, CritEdt.Decimale) ;
     // Totalisation
     TotMulti[(MaxTabCol*2)]:= Arrondi(TotMulti[(MaxTabCol*2)]+RealPer[0].TotDebit, CritEdt.Decimale) ;
     TotMulti[(MaxTabCol*3)]:= Arrondi(TotMulti[(MaxTabCol*3)]+RealPer[0].TotCredit, CritEdt.Decimale) ;

     if (QrLD_POURBUD.AsString<>'BUD') then
       BEGIN
       TotEdtRea[i].TotDebit := Arrondi(TotEdtRea[i].TotDebit+RealPer[0].TotDebit, CritEdt.Decimale) ;
       TotEdtRea[i].TotCredit:= Arrondi(TotEdtRea[i].TotCredit+RealPer[0].TotCredit, CritEdt.Decimale) ;
       // Totalisation
       TotEdtRea[0].TotDebit := Arrondi(TotEdtRea[0].TotDebit+RealPer[0].TotDebit, CritEdt.Decimale) ;
       TotEdtRea[0].TotCredit:= Arrondi(TotEdtRea[0].TotCredit+RealPer[0].TotCredit, CritEdt.Decimale) ;

       TotCptRea[i].TotDebit := Arrondi(TotCptRea[i].TotDebit+RealPer[0].TotDebit, CritEdt.Decimale) ;
       TotCptRea[i].TotCredit:= Arrondi(TotCptRea[i].TotCredit+RealPer[0].TotCredit, CritEdt.Decimale) ;
       // Totalisation
       TotCptRea[0].TotDebit := Arrondi(TotCptRea[0].TotDebit+RealPer[0].TotDebit, CritEdt.Decimale) ;
       TotCptRea[0].TotCredit:= Arrondi(TotCptRea[0].TotCredit+RealPer[0].TotCredit, CritEdt.Decimale) ;
       END ;
     If CritEdt.PlanEdt.PasDeRubrique Then
        BEGIN
        XX:=GetCumul(TypeCumBud,Compte1,Compte2,CritEdt.PlanEdt.SANbud,CritEdt.Etab,CritEdt.PlanEdt.Journal,LExo,CritEdt.Date1,CritEdt.PlanEdt.Date21,false,TRUE,Nil,TD,FALSE) ;
        TD[4]:=0 ; TD[5]:=0 ;
        If XX>0 Then TD[5]:=XX Else TD[4]:=Abs(XX) ;
        RealPer[1].TotDebit:=TD[4] ; RealPer[1].TotCredit:=TD[5] ;
        END Else
        BEGIN
        GetCumul(TypeCumBud,Compte1,'',CritEdt.PlanEdt.SANbud,CritEdt.Etab,
                 CritEdt.PlanEdt.Journal,LExo,CritEdt.Date1,CritEdt.PlanEdt.Date21,false,TRUE,Nil,TD,FALSE) ;
        RealPer[1].TotDebit:=TD[1] ; RealPer[1].TotCredit:=TD[2] ;
        END ;
     Reevaluation(RealPer[1].TotDebit,RealPer[1].TotCredit,CritEdt.PlanEdt.Resol,CritEdt.PlanEdt.Taux) ;
     RealPer[3].TotDebit:=RealPer[3].TotDebit+RealPer[1].TotDebit ; RealPer[3].TotCredit:=RealPer[3].TotCredit+RealPer[1].TotCredit ;
     If OkMB Then TMB.Caption:=PrintSoldeFormate(RealPer[1].TotDebit, RealPer[1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.PlanEdt.SoldeFormate) ;
     TotMulti[i]:= Arrondi(TotMulti[i]+RealPer[1].TotDebit, CritEdt.Decimale) ;
     TotMulti[MaxTabCol+i]:= Arrondi(TotMulti[MaxTabCol+i]+RealPer[1].TotCredit, CritEdt.Decimale) ;
     //Totalisation
     TotMulti[0]:= Arrondi(TotMulti[0]+RealPer[1].TotDebit, CritEdt.Decimale) ;
     TotMulti[MaxTabCol]:= Arrondi(TotMulti[MaxTabCol]+RealPer[1].TotCredit, CritEdt.Decimale) ;
     if (QrLD_POURBUD.AsString<>'REA') then
       BEGIN
       TotEdtBud[i].TotDebit:= Arrondi(TotEdtBud[i].TotDebit+RealPer[1].TotDebit, CritEdt.Decimale) ;
       TotEdtBud[i].TotCredit:= Arrondi(TotEdtBud[i].TotCredit+RealPer[1].TotCredit, CritEdt.Decimale) ;
       //Totalisation
       TotEdtBud[0].TotDebit:= Arrondi(TotEdtBud[0].TotDebit+RealPer[1].TotDebit, CritEdt.Decimale) ;
       TotEdtBud[0].TotCredit:= Arrondi(TotEdtBud[0].TotCredit+RealPer[1].TotCredit, CritEdt.Decimale) ;

       TotCptBud[i].TotDebit:= Arrondi(TotCptBud[i].TotDebit+RealPer[1].TotDebit, CritEdt.Decimale) ;
       TotCptBud[i].TotCredit:= Arrondi(TotCptBud[i].TotCredit+RealPer[1].TotCredit, CritEdt.Decimale) ;
       //Totalisation
       TotCptBud[0].TotDebit:= Arrondi(TotCptBud[0].TotDebit+RealPer[1].TotDebit, CritEdt.Decimale) ;
       TotCptBud[0].TotCredit:= Arrondi(TotCptBud[0].TotCredit+RealPer[1].TotCredit, CritEdt.Decimale) ;
       END ;
     END ;
  END ;
TQRLabel(Findcomponent('MontantRea'+IntToStr(NbColAff+1))).Caption:=PrintSoldeFormate(RealPer[2].TotDebit, RealPer[2].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.PlanEdt.SoldeFormate) ;
TQRLabel(Findcomponent('MontantBud'+IntToStr(NbColAff+1))).Caption:=PrintSoldeFormate(RealPer[3].TotDebit, RealPer[3].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.PlanEdt.SoldeFormate) ;
END ;

procedure TFQRPlanEd.AfficheBudOuRea(QRB : TQRBand) ;
var T : TQRLabel ;
    HBud,X,i : integer ;
    Leq : String3 ;
BEGIN
X:=QRB.Height ;
T:=TotCpteRea ;
if QRB=BSecondaire then T:=RevGenRea else
  if QRB=BFinEtat  then T:=RTotEdtRea ;
Leq:=QrLD_POURBUD.AsString ;
if CritEdt.PlanEdt.QueTotalRea then Leq:='REA' ;
if (Leq<>'BRE') and (X>=(T.Top+T.Height)) then X:=X-T.Height ;
if (Leq='BRE') and (X<=(T.Top+T.Height)) then X:=X+T.Height ;
if (Leq='BUD') then HBud:=T.Top else HBud:=T.Top+T.Height;
for i:=0 to QRB.ControlCount-1 do
  BEGIN
  if QRB.Controls[i] is TQRLabel then
    BEGIN
    T:=TQRLabel(QRB.Controls[i]) ;
    if (UpperCase(T.Name)<>'COMPTEB') and (UpperCase(T.Name)<>'RTOTALGEN') then
      BEGIN
      if (Pos('BUD',UpperCase(T.Name))<>0) then T.Top:=HBud ;
      T.Visible:=(Pos('REA',UpperCase(T.Name))<>0) and ((Leq<>'BUD')) ;
      if not T.Visible then T.Visible:=(Pos('BUD',UpperCase(T.Name))<>0) and ((Leq='BUD') or (Leq='BRE')) ;
      END ;
    END ;
  END ;
QRB.Height:=X ;
END ;

procedure TFQRPlanEd.BSecondaireBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
//var LaRupture : TRuptInf ;
var    TotMulti : TabMont77 ;
       TotBud   : TMontTot ;
       CodeRecap, LibRecap : String ;
begin
  inherited;
Fillchar(TotMulti,SizeOf(TotMulti),#0) ;
AlimMontant(TotBud,TotMulti) ;
AddRecapSections ;
ReTrieDetail ;
AfficheBudOuRea(BSecondaire) ;
end;

procedure TFQRPlanEd.BHCptBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
PrintBand:=FComposite ;
end;

procedure TFQRPlanEd.BFinEtatBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
var i : integer ;
begin
  inherited;
AfficheBudOuRea(BFinEtat) ;
For i:=1 To NbColAff Do
  BEGIN
  TQRLabel(FindComponent('TotGenRea'+IntToStr(i))).Caption:=PrintSoldeFormate(TotEdtRea[i].TotDebit, TotEdtRea[i].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.PlanEdt.SoldeFormate) ;
  TQRLabel(FindComponent('TotGenBud'+IntToStr(i))).Caption:=PrintSoldeFormate(TotEdtBud[i].TotDebit, TotEdtBud[i].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.PlanEdt.SoldeFormate) ;
  END ;
//Totaux
TQRLabel(FindComponent('TotGenRea'+IntToStr(NbColAff+1))).Caption:=PrintSoldeFormate(TotEdtRea[0].TotDebit, TotEdtRea[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.PlanEdt.SoldeFormate) ;
TQRLabel(FindComponent('TotGenBud'+IntToStr(NbColAff+1))).Caption:=PrintSoldeFormate(TotEdtBud[0].TotDebit, TotEdtBud[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.PlanEdt.SoldeFormate) ;
end;

procedure TFQRPlanEd.BHRecapBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
PrintBand:=FComposite ;
end;

procedure TFQRPlanEd.MajListeTot(Liste : TStringList) ;
var i : integer ;
BEGIN
if Liste.Count=0 then Exit ;
ListeTotAnaGene.Clear ;
for i:=0 to Liste.Count-1 do ListeTotAnaGene.AddObject(Liste[i],Liste.Objects[i]) ;
END ;

procedure TFQRPlanEd.AddDetail(Liste : TStringList ; Cpt : String ; FTotAG : TFTotAnaGene) ;
var i,Ind    : integer ;
    FTotD : TFTotAnaGene ;
BEGIN
Ind:=Liste.IndexOf(Cpt) ;
if Ind<>-1 then
   BEGIN
   FTotD:=TFTotAnaGene(Liste.Objects[Ind]) ;
   FTotD.Sect:='' ; FTotD.Gene:=Cpt ;
   for i:=Low(FTotD.TotAG) to MaxTabCol do
     BEGIN
     FTotD.TotAG[i]:=FTotD.TotAG[i]+FTotAG.TotAG[i] ;
     FTotD.TotAG[MaxTabCol+i]:=FTotD.TotAG[MaxTabCol+i]+FTotAG.TotAG[MaxTabCol+i] ;
     END ;
   Liste.Objects[Ind]:=FTotD ;
   END else
   BEGIN
   FTotD:=TFTotAnaGene.Create ;
   FTotD.Sect:='' ; FTotD.Gene:=Cpt ;
   FTotD.TotAG:=FTotAG.TotAG ;
   Liste.AddObject(Cpt,FTotD) ;
   END ;
END ;

procedure TFQRPlanEd.ReTrieDetail ;
var FTotAG : TFTotAnaGene ;
    i : integer ;
    Tot       : TabMont77 ;
    Cpt1,Cpt2 : String17 ;
    Liste : TStringList ;
BEGIN
if (ListeTotAnaGene=nil) or (ListeTotAnaGene.Count=0) then Exit ;
if (Copy(QrLD_TYPEDETAIL.AsString,6,1)='X') and ((Copy(QrLD_TYPEDETAIL.AsString,4,1)='-') and (Copy(QrLD_TYPEDETAIL.AsString,5,1)='-')) then Exit ;
Liste:=TStringList.Create ;
for i:=0 to ListeTotAnaGene.Count-1 do
  BEGIN
  FTotAG := TFTotAnaGene(ListeTotAnaGene.Objects[i]) ;
  Tot:=FTotAG.TotAG ;
  Cpt1:=FTotAG.Sect ; Cpt2:=FTotAG.Gene ;
  // Détail sur premier compte
  if ((CritEdt.PlanEdt.Detail1='X') or ((CritEdt.PlanEdt.Detail1='') and (Copy(QrLD_TYPEDETAIL.AsString,4,1)='X'))) and (Cpt1<>'') then AddDetail(Liste,Cpt1,FTotAG) ;
  // Détail sur deuxième compte
  if ((CritEdt.PlanEdt.Detail2='X') or ((CritEdt.PlanEdt.Detail2='') and (Copy(QrLD_TYPEDETAIL.AsString,5,1)='X'))) and (Cpt2<>'') then AddDetail(Liste,Cpt2,FTotAG) ;
  // Détail sur couple
  if ((CritEdt.PlanEdt.Detail3='X') or ((CritEdt.PlanEdt.Detail3='') and (Copy(QrLD_TYPEDETAIL.AsString,6,1)='X'))) and (Cpt1<>'') and (Cpt2<>'') then AddDetail(Liste,Cpt1+'/'+Cpt2,FTotAG) ;
  END ;
MajListeTot(Liste) ;
Liste.Clear ; Liste.Free ;
END ;

procedure TFQRPlanEd.DLAnaGeneNeedData(var MoreData: Boolean);
var    Cpt1,Cpt2 : String17 ;
       Tot       : TabMont77 ;
       i         : integer ;
       LaFam,St : String ;
begin
  inherited;
Cpt1:='' ; Cpt2:='' ;
Fillchar(Tot,SizeOf(Tot),#0) ; Fillchar(Tot,SizeOf(Tot),#0) ;
MoreData:=PrintAnaGene(Cpt1,Cpt2,Tot) ;
St:='' ;
if MoreData then
  BEGIN
  LaFam:=CritEdt.PlanEdt.Famille ;
  if (Cpt1<>'') and (Cpt2<>'') then St:=Cpt1+'/'+Cpt2 else
   if (Cpt1<>'') and (Cpt2='') then St:=Cpt1 else
    if (Cpt1='') and (Cpt2<>'') then St:=Cpt2 ;
  TFAnaGene.Caption:=MsgBox.Mess[2]+' '+St ;
  for i:=1 to NbColAff do
    BEGIN
    Tot[0]:=Tot[0]+Tot[i] ; Tot[MaxTabCol]:=Tot[MaxTabCol]+Tot[MaxTabCol+i] ;
    TQRLabel(FindComponent('TotAnaGene'+IntToStr(i))).Caption:=PrintSoldeFormate(Tot[i], Tot[MaxTabCol+i], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
    END ;
  TQRLabel(FindComponent('TotAnaGene'+IntToStr(NbColAff+1))).Caption:=PrintSoldeFormate(Tot[0], Tot[MaxTabCol], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
  END ;
end;

function TFQRPlanEd.PrintAnaGene (var Cpt1,Cpt2 : String17 ; var Tot : TabMont77) : boolean ;
var FTotAG : TFTotAnaGene ;
BEGIN
Result:=False ;
if (ListeTotAnaGene=nil) or (ListeTotAnaGene.Count=0) then Exit ;
FTotAG := TFTotAnaGene(ListeTotAnaGene.Objects[0]) ;
Tot:=FTotAG.TotAG ;
Cpt1:=FTotAG.Sect ; Cpt2:=FTotAG.Gene ;
ListeTotAnaGene.Delete(0) ;
Result:=True ;
END ;

Procedure TFQRPlanEd.PositionnePlan ;
Var i : Integer ;
BEGIN
for i:=0 to FPLANEDT.Values.Count-1 do
   if (i<>FPLANEDT.ItemIndex) and (FPLANEDT.Values[i]=FPLANEDT.Value) And
      (TInfoPlanEdt(FPLANEDT.Values.Objects[i]).Nat=PD_NATUREEDT.Text) then BEGIN FPLANEDT.ItemIndex:=i ; Break ; END ;
END ;

procedure TFQRPlanEd.PD_NATUREEDTChange(Sender: TObject);
begin
  inherited;
PositionnePlan ;
end;

procedure TFQRPlanEd.BMultiBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
PrintBand:=FComposite ;
end;

procedure TFQRPlanEd.BRecapBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
PrintBand:=((FComposite or (CritEdt.PlanEdt.Famille='CBS')) and CritEdt.PlanEdt.FormatPrint.PrSepCompte[10]) ;
end;

procedure TFQRPlanEd.BFAnaGeneBeforePrint(var PrintBand: Boolean;
  var Quoi: String);
begin
  inherited;
if (CritEdt.PlanEdt.Detail1='-') and (CritEdt.PlanEdt.Detail2='-') and (CritEdt.PlanEdt.Detail3='-') then PrintBand:=False ;
if PrintBand and (Copy(QrLD_TYPEDETAIL.AsString,4,1)='-') and ((Copy(QrLD_TYPEDETAIL.AsString,5,1)='-') and (Copy(QrLD_TYPEDETAIL.AsString,6,1)='-')) then PrintBand:=False ;
end;

procedure TFQRPlanEd.DLRecapNeedData(var MoreData: Boolean);
var TotRecap         : TabMont77 ;
    LibRecap, CodRecap : String ;
    Chaque,i         : Byte ;
    StM              : String ;
    D,C              : Double ;
begin
  inherited;
MoreData:=False ;
if FComposite or (CritEdt.PlanEdt.Famille='CBS') then
   BEGIN
   MoreData:=PrintRecap(LRecap,CodRecap,LibRecap,TotRecap) ;
   if MoreData then
      BEGIN
      { Affichage sur la Bande }
      TotRecapRea.Caption:=CodRecap+' '+LibRecap ;
      for i:=1 to NbColAff+1 do
        BEGIN
        if i<NbColAff+1 then
           BEGIN
           D:=TotRecap[i] ; C:=TotRecap[MaxTabCol+i] ;
           END else
           BEGIN
           D:=TotRecap[0] ; C:=TotRecap[MaxTabCol] ;
           END ;
        StM:=PrintSoldeFormate(D,C, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
        TQRLabel(FindComponent('TotRRecap'+IntToStr(i))).Caption:=StM;
        END ;
     END ;
   END ;
OkEnTeteRecap:=MoreData ;
end;

procedure TFQRPlanEd.BRecapAfterPrint(BandPrinted: Boolean);
begin
  inherited;
BRecap.ForceNewPage:=False ;
end;

procedure TFQRPlanEd.FRecapClick(Sender: TObject);
begin
  inherited;
FQuelSPlan.Enabled:=FRecap.Checked ;
end;

end.
