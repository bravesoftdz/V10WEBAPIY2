unit QRCatSG;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, HQuickrp, HSysMenu, Menus, hmsgbox, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  StdCtrls, Buttons,
  Ent1, Hent1, Hctrls, ExtCtrls, Mask, Hcompte, ComCtrls, CritEdt,UtilEdt,Choix,
  BudSect,QRRupt,Calcole,CpteUtil,UtilEdt1,SaisUtil, HTB97, HPanel, UiUtil ;

procedure EcartParCategorie ;

type TFTotAnaGene = Class
    Sect,Gene : String17 ;
    TotAG : TabMont77 ;
    END ;
type
  TFQRCatSG = class(TFQR)
    TFCategorie: THLabel;
    HLabel3: THLabel;
    FPeriode1: THValComboBox;
    FPeriode2: THValComboBox;
    Label1: TLabel;
    TFExercice_: TLabel;
    FExercice_: THValComboBox;
    FValide: TCheckBox;
    FTaux: THNumEdit;
    TFTaux: THLabel;
    FQte: TCheckBox;
    FQueTotRea: TCheckBox;
    Bevel3: TBevel;
    FRESOL: THValComboBox;
    TFRESOL: THLabel;
    BS_BUDSECT: TQRLabel;
    BS_LIBELLE: TQRLabel;
    TitreColCpt: TQRLabel;
    Col1: TQRLabel;
    Col2: TQRLabel;
    Col3: TQRLabel;
    Col5: TQRLabel;
    Col4: TQRLabel;
    BHCpt: TQRBand;
    BG_BUDGENE: TQRDBText;
    BG_LIBELLE: TQRLabel;
    RCategorie: TQRLabel;
    TRCategorie: TQRLabel;
    QRLabel3: TQRLabel;
    RExercice_: TQRLabel;
    TRChoixMontant: TQRLabel;
    RChoixMontant: TQRLabel;
    RResol: TQRLabel;
    TRResol: TQRLabel;
    QRLabel1: TQRLabel;
    TRTaux: TQRLabel;
    RTaux: TQRLabel;
    RValide: TQRLabel;
    QRLabel17: TQRLabel;
    RQte: TQRLabel;
    TRGeneral: TQRLabel;
    RGeneral1: TQRLabel;
    TRaG: TQRLabel;
    RGeneral2: TQRLabel;
    ZoomJAL: TToolbarButton97;
    FListeCateg: TEdit;
    TFJokerGene: THLabel;
    TFGene: THLabel;
    TFaB: TLabel;
    FJokerGene: TEdit;
    FGene1: THCpteEdit;
    FGene2: THCpteEdit;
    Col7: TQRLabel;
    Col6: TQRLabel;
    Col8: TQRLabel;
    Col9: TQRLabel;
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
    IniGen: TQRLabel;
    RevGen: TQRLabel;
    MontantRea1: TQRLabel;
    MontantRea2: TQRLabel;
    MontantRea3: TQRLabel;
    MontantRea4: TQRLabel;
    MontantRea5: TQRLabel;
    MontantRea6: TQRLabel;
    MontantRea7: TQRLabel;
    MontantRea8: TQRLabel;
    MontantRea9: TQRLabel;
    BFPrimaire: TQRBand;
    BS_BUDSECT_: TQRLabel;
    TotEdtRea1: TQRLabel;
    TotEdtRea2: TQRLabel;
    TotEdtRea3: TQRLabel;
    TotEdtRea4: TQRLabel;
    TotEdtRea5: TQRLabel;
    TotEdtRea6: TQRLabel;
    TotEdtRea7: TQRLabel;
    TotEdtRea8: TQRLabel;
    TotEdtRea9: TQRLabel;
    MsgBox: THMsgBox;
    QSecond: TQuery;
    DLGen: TQRDetailLink;
    DLHGen: TQRDetailLink;
    SSecond: TDataSource;
    DLMulti: TQRDetailLink;
    BRappelLibreHaut: TQRBand;
    TCodRupt: TQRLabel;
    TLibRupt: TQRLabel;
    BLibreHaut: TQRBand;
    TRappelLibRupt: TQRLabel;
    TRappelCodRupt: TQRLabel;
    QRLabel2: TQRLabel;
    TotEdtBud1: TQRLabel;
    TotEdtBud2: TQRLabel;
    TotEdtBud3: TQRLabel;
    TotEdtBud4: TQRLabel;
    TotEdtBud5: TQRLabel;
    TotEdtBud6: TQRLabel;
    TotEdtBud7: TQRLabel;
    TotEdtBud8: TQRLabel;
    TotEdtBud9: TQRLabel;
    QRLabel20: TQRLabel;
    TotBudSect: TQRLabel;
    FAvecRub: TCheckBox;
    DHLRappelLibre: TQRDetailLink;
    DLHLibre: TQRDetailLink;
    FLigneRupt: TCheckBox;
    QRLabel22: TQRLabel;
    QRLabel23: TQRLabel;
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
    BHRecap: TQRBand;
    BG_BUDGENE2: TQRDBText;
    BG_LIBELLE2: TQRLabel;
    BRecap: TQRBand;
    TotBRecap5: TQRLabel;
    TotBRecap4: TQRLabel;
    TotBRecap2: TQRLabel;
    TotBRecap3: TQRLabel;
    TotBRecap6: TQRLabel;
    TotBRecap1: TQRLabel;
    TotBRecap7: TQRLabel;
    TotBRecap8: TQRLabel;
    TotBRecap9: TQRLabel;
    TotBudRecap: TQRLabel;
    QRLabel37: TQRLabel;
    TotRRecap1: TQRLabel;
    TotRRecap2: TQRLabel;
    TotRRecap3: TQRLabel;
    TotRRecap4: TQRLabel;
    TotRRecap5: TQRLabel;
    TotRRecap6: TQRLabel;
    TotRRecap7: TQRLabel;
    TotRRecap8: TQRLabel;
    TotRRecap9: TQRLabel;
    FAvecCptSecond: TCheckBox;
    DLHRecap: TQRDetailLink;
    DLRecap: TQRDetailLink;
    BMulti: TQRBand;
    QReaLibre4: TQRLabel;
    QRea: TQRLabel;
    QReaLibre5: TQRLabel;
    QReaLibre6: TQRLabel;
    QReaLibre1: TQRLabel;
    QReaLibre2: TQRLabel;
    QReaLibre3: TQRLabel;
    QReaLibre7: TQRLabel;
    QReaLibre8: TQRLabel;
    QReaLibre9: TQRLabel;
    QBudLibre4: TQRLabel;
    QBud: TQRLabel;
    QBudLibre5: TQRLabel;
    QBudLibre6: TQRLabel;
    QBudLibre1: TQRLabel;
    QBudLibre2: TQRLabel;
    QBudLibre3: TQRLabel;
    QBudLibre7: TQRLabel;
    QBudLibre8: TQRLabel;
    QBudLibre9: TQRLabel;
    FAnaGene: TCheckBox;
    BFAnaGene: TQRBand;
    TotAnaGene1: TQRLabel;
    TotAnaGene2: TQRLabel;
    TFAnaGene: TQRLabel;
    DLAnaGene: TQRDetailLink;
    TotAnaGene3: TQRLabel;
    TotAnaGene4: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel29: TQRLabel;
    QRLabel30: TQRLabel;
    QRLabel31: TQRLabel;
    TotAnaGene5: TQRLabel;
    TotAnaGene6: TQRLabel;
    TotAnaGene7: TQRLabel;
    TotAnaGene8: TQRLabel;
    TotAnaGene9: TQRLabel;
    FColonnage: TRadioGroup;
    FChoixMontant: THValComboBox;
    TFChoixMontant: THLabel;
    Col10: TQRLabel;
    Col11: TQRLabel;
    Col13: TQRLabel;
    Col12: TQRLabel;
    MontantRea10: TQRLabel;
    MontantBud10: TQRLabel;
    MontantRea11: TQRLabel;
    MontantBud11: TQRLabel;
    MontantRea12: TQRLabel;
    MontantBud12: TQRLabel;
    MontantRea13: TQRLabel;
    MontantBud13: TQRLabel;
    TotEdtRea10: TQRLabel;
    TotEdtBud10: TQRLabel;
    TotEdtRea11: TQRLabel;
    TotEdtBud11: TQRLabel;
    TotEdtRea12: TQRLabel;
    TotEdtBud12: TQRLabel;
    TotEdtRea13: TQRLabel;
    TotEdtBud13: TQRLabel;
    TotAnaGene10: TQRLabel;
    TotAnaGene11: TQRLabel;
    TotAnaGene12: TQRLabel;
    TotAnaGene13: TQRLabel;
    TotRRecap10: TQRLabel;
    TotBRecap10: TQRLabel;
    TotRRecap11: TQRLabel;
    TotBRecap11: TQRLabel;
    TotRRecap12: TQRLabel;
    TotBRecap12: TQRLabel;
    TotRRecap13: TQRLabel;
    TotBRecap13: TQRLabel;
    QReaLibre10: TQRLabel;
    QBudLibre10: TQRLabel;
    QReaLibre11: TQRLabel;
    QBudLibre11: TQRLabel;
    QReaLibre12: TQRLabel;
    QBudLibre12: TQRLabel;
    QReaLibre13: TQRLabel;
    QBudLibre13: TQRLabel;
    TotGenRea10: TQRLabel;
    TotGenBud10: TQRLabel;
    TotGenRea11: TQRLabel;
    TotGenBud11: TQRLabel;
    TotGenRea12: TQRLabel;
    TotGenBud12: TQRLabel;
    TotGenRea13: TQRLabel;
    TotGenBud13: TQRLabel;
    FCategorie: THValComboBox;
    procedure ZoomJALClick(Sender: TObject);
    procedure BHCptBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure QSecondAfterOpen(DataSet: TDataSet);
    procedure FGene1Change(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FExerciceChange(Sender: TObject);
    procedure FExercice_Change(Sender: TObject);
    procedure FPeriode1Change(Sender: TObject);
    procedure FPeriode2Change(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure BFPrimaireBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BSecondaireBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FCategorieChange(Sender: TObject);
    procedure DLMultiNeedData(var MoreData: Boolean);
    procedure FPlanRupturesChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BMultiBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BMultiAfterPrint(BandPrinted: Boolean);
    procedure DHLRappelLibreNeedData(var MoreData: Boolean);
    procedure DLHLibreNeedData(var MoreData: Boolean);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean;
      var Quoi: string);
    procedure FSansRuptClick(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure DLRecapNeedData(var MoreData: Boolean);
    procedure DLHRecapNeedData(var MoreData: Boolean);
    procedure DLHGenNeedData(var MoreData: Boolean);
    procedure BRecapBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FAnaGeneClick(Sender: TObject);
    procedure BFAnaGeneBeforePrint(var PrintBand: Boolean;
      var Quoi: string);
    procedure DLAnaGeneNeedData(var MoreData: Boolean);
    procedure FColonnageClick(Sender: TObject);
  private
    QuelEtatBud                  : TQuelEtatBud ;
    AxeJal                       : String3 ;
    Floading,DebitPos            : Boolean ;
    DansTotal                    : Boolean ;
    OkEnteteLibre,OkEnTeteGen,OkEnteteRecap,
    OkRappelEnteteLibre          : Boolean ;
    ExoDateDeb,ExoDateFin        : TDateTime ;
    Qr1BS_AXE,QR1BS_BUDSECT,
    Qr1BS_LIBELLE,QR1BS_RUB,
    Qr1BS_SECTIONTRIE            : TStringField ;
    Qr12BG_BUDGENE,Qr12BG_LIBELLE,QR12BG_RUB: TStringField ;
    TotEdtRea,TotEdtBud          : TabTot12 ;
    TotCptRea,TotCptBud          : TabTot12 ;
    TotAGRea                     : TabTot12 ;
    NbColAff                     : Integer ;
    LMulti, LRecap               : TStringList ;
    ListeCodesRupture            : TStringList ;
    ListeAnaGene                 : TStringList ;
    ListeTotAnaGene              : TStringList ;
    TabDate                      : TTabDate12 ;
    NbMois                       : Word ;
    procedure GenereSQL2 ;
    procedure InfosJournal ;
    procedure DebToDate ;
    procedure FinToDate ;
    procedure AffEntetesCol ;
    function  OkPourRupt : Boolean ;
    function  OkJal : Boolean ;
    procedure AlimMontant(Var TotBud : TMontTot ; Var TotMulti : TabMont77) ;
    Function  QuoiCpt(i : Integer) : String ;
    Function  QuoiCptRecap(St1,St2 : String) : String ;
    Procedure BalMultiBudZoom(Quoi : String) ;
    Procedure MontantAnaGene(Indice : Integer) ;
    function  PrintAnaGene (var Sect,Gene:String17 ; var Tot : TabMont77) : boolean ;
    procedure CalculDateBud ;
  public
    { Déclarations publiques }
    procedure InitDivers ; override ;
    procedure GenereSQL ; override ;
    procedure ChoixEdition ; Override ;
    procedure RenseigneCritere ; override ;
    procedure RecupCritEdt ; override ;
    Procedure FinirPrint ; override ;
    function  CritOk : Boolean ; override ;
  end;

implementation

const MaxColonnes  = 13 ;
      MaxTabCol    = 15 ;
{$R *.DFM}

procedure EcartParCategorie ;
var QR : TFQRCatSG ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
QR:=TFQRCatSG.Create(Application) ;
Edition.Etat:=etBudMulti ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalMultiBudZoom ;
QR.QuelEtatBud:=Normal ;
QR.InitType (nbBudSec,neBalBud,msSecBAna,'TFQRCatSG','',TRUE,FALSE,TRUE,Edition) ;
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
END ;

Function TFQRCatSG.QuoiCpt(i : Integer) : String ;
BEGIN
Result:='' ;
Case i of
   0 : Result:=QR1BS_BUDSECT.AsString+' '+QR1BS_LIBELLE.AsString+'@'+'9'+Qr1BS_AXE.AsString;
   1 : Result:=QR12BG_BUDGENE.AsString+' '+QR12BG_LIBELLE.AsString+'@'+'8' ;
   2 : Result:=QR1BS_BUDSECT.AsString+' '+QR1BS_LIBELLE.AsString+
               '#'+QR12BG_BUDGENE.AsString+' '+QR12BG_LIBELLE.AsString+'@'+'9'+Qr1BS_AXE.AsString;
  End ;
END ;

Function TFQRCatSG.QuoiCptRecap(St1,St2 : String) : String ;
BEGIN
Result:=St1+St2+'@'+'8' ;
END ;

Procedure TFQRCatSG.BalMultiBudZoom(Quoi : String) ;
Var Lp,i: Integer ;
BEGIN
Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
i:=StrToInt(Copy(Quoi,Lp+1,1)) ;
ZoomEdt(i,Quoi) ;
END ;

function BudJal2Categorie(JAL : String3) : String3 ;
var Q : TQuery ;
BEGIN
Result:='' ;
Q:=OpenSQL('SELECT BJ_CATEGORIE FROM BUDJAL WHERE BJ_BUDJAL="'+JAL+'"',TRue) ;
if Not Q.Eof then Result:=Q.Fields[0].AsString ; 
Ferme(Q) ;
END ;

procedure TFQRCatSG.RecupCritEdt ;
var i : integer ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  BalBud.ColonnageJal:=(FColonnage.ItemIndex=0) ;
  if BalBud.ColonnageJal then
    BEGIN
    BalBud.Categorie:=FCategorie.Value ;
    BalBud.Journal:='' ;
    END else
    BEGIN
    BalBud.Categorie:=BudJal2Categorie(FCategorie.Value) ;
    BalBud.Journal:=FCategorie.Value ;
    END ;
  BalBud.NatureBud:='' ;
  BalBud.Axe:=AxeJal ;
  BalBud.Exo1:=FExercice.Value ; BalBud.Exo2:=FExercice_.Value ;
  Date1:=StrToDate(FDateCompta1.Text)    ; Date2:=StrToDate(FDateCompta2.Text) ;
  BalBud.Date21:=FinDeMois(CritEdt.Date2) ;
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
  BalBud.Qte:=FQte.Checked ;
  BalBud.AnaGene:=FAnaGene.Checked ;
  { Rupture }
  BalBud.CodeRupt1:='' ; BalBud.CodeRupt2:='' ; BalBud.PlanRupt:='' ;
  If Not FSansRupt.Checked Then BalBud.PlanRupt:=FPlanRuptures.Value ;
  BaLBud.RuptOnly:=QuelleTypeRupt(0,FSAnsRupt.Checked,FAvecRupt.Checked,FSurRupt.Checked) ;
  If BalBud.RuptOnly=Sur Then
     BEGIN
     BalBud.CodeRupt1:=FCodeRupt1.Text ; BalBud.CodeRupt2:=FCodeRupt2.Text ;
     END ;
//  BalBud.OnlySecRupt:=FOnlySecRupt.checked ;
//  BalBud.AvecCptSecond:=FAvecCptSecond.checked ;
//  BalBud.SoldeFormate:=FSoldeFormate.Value ;
  BalBud.AvecCptSecond:=FAvecCptSecond.checked ;
  BalBud.OnlyCptAssocie:=(Rupture<>rRien) and FOnlyCptAssocie.Checked ;
  BalBud.QueTotalRea:=FQueTotRea.Checked ;
// Si variable dir debit pos ou crédit Pos
  //DebitPos:=(FDebitPos.ItemIndex=0) ;
  BalBud.SoldeFormate:=FChoixMontant.Value ;
  //If DebitPos Then BalBud.SoldeFormate:='PD' Else BalBud.SoldeFormate:='PC' ;
  Balbud.PasDeRubrique:=Not FAvecRub.Checked ;
  AffEntetesCol ;
  // +1 pour la 1ère colonne non affichée
  if NbColAff+1<=MaxColonnes then
// MaxColonnes+1 : pour Colonne budget
    for i:=1 to MaxColonnes+1 do
      BEGIN
// NbColAff +2 -> +1 : Budget +1 : Total
      BalBud.FormatPrint.TabColl[i].OkAff:=(i<=NbColAff+2) ;
      PremTabColEdt[i].OkAff:=(i<=NbColAff+2) ;
      END ;
  END ;
END ;

function TFQRCatSG.OkJal : Boolean ;
var s,Jal : String ;
    QJal  : TQuery ;
BEGIN
Result:=False ;
s:=FListeCateg.Text ; if s='' then BEGIN NumErreurCrit:=11 ; Exit ; END ;
Jal:=ReadTokenSt(s) ;
AxeJal:='' ;
QJal:=TQuery.Create(Application) ; QJal.DataBaseName:='SOC' ;
QJal.SQL.Add('Select BJ_AXE From BUDJAL Where BJ_BUDJAL=:JAL') ;
ChangeSQL(QJal) ; QJal.Prepare ;
while Jal<>'' do
  BEGIN
  QJal.Params[0].AsString:=Jal ;
  QJal.Open ;
  if not QJal.Eof then
    if (AxeJal='') then
      BEGIN
      if not QJal.Eof then
        BEGIN
        AxeJal:=QJal.Fields[0].AsString ;
        END ;
      END else
      BEGIN
      if AxeJal<>QJal.Fields[0].AsString then
        BEGIN
        NumErreurCrit:=7 ;
        QJal.CLose ; QJal.Free ;
        Exit ;
        END ;
      END ;
  QJal.Close ;
  Jal:=ReadTokenSt(s) ;
  END ;
QJal.Free ;
Result:=True ;
END ;

function TFQRCatSG.CritOk : Boolean ;
BEGIN
Result:=OkJal and Inherited CritOK and OkPourRupt ;
END ;

Function TFQRCatSG.OkPourRupt : Boolean ;
BEGIN
Result:=True ;
ListeTotAnaGene:=nil ; ListeAnaGene:=nil ;
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
   END else
if (CritEdt.Rupture=rRien) then
   BEGIN
   ListeAnaGene:=TStringList.Create ;
   ListeTotAnaGene:=TStringList.Create ;
   END ;
END ;

procedure TFQRCatSG.InitDivers ;
BEGIN
Inherited ;
Fillchar(TotEdtRea,SizeOf(TotEdtRea),#0) ; Fillchar(TotEdtBud,SizeOf(TotEdtBud),#0) ;
Fillchar(TotCptRea,SizeOf(TotCptRea),#0) ; Fillchar(TotCptBud,SizeOf(TotCptBud),#0) ;
Fillchar(TotAGRea,SizeOf(TotAGRea),#0) ;
BFPrimaire.Frame.DrawBottom:=CritEdt.BalBud.FormatPrint.PrSepCompte[1] ;
BHCpt.Frame.DrawTop:=CritEdt.BalBud.AnaGene ;
BMulti.Frame.DrawBottom:=FLigneRupt.checked ;
CritEdt.BalBud.FormatPrint.ColMin:=1 ;
CritEdt.BalBud.FormatPrint.ColMax:=MaxColonnes+1 ;
if CritEdt.BalBud.QueTotalRea then
  BEGIN
  if (BSecondaire.Height>(IniGen.Height*2)) then
    BEGIN
    IniGen.Caption:='' ;
    BSecondaire.Height:=BSecondaire.Height-IniGen.Height ;
    TotBudSect.Caption:='' ;
    BFPrimaire.Height:=BFPrimaire.Height-TotBudSect.Height ;
    TotBudRecap.Caption:='' ;
    BRecap.Height:=BRecap.Height-TotBudRecap.Height ;
    QBud.Caption:='' ;
    BMulti.Height:=BMulti.Height-QBud.Height ;
    RTotEdtBud.Caption:='' ;
    BFinEtat.Height:=BFinEtat.Height-RTotEdtBud.Height ;
    END ;
  END else
  if (BSecondaire.Height<(IniGen.Height*2)) then
    BEGIN
    IniGen.Caption:=MsgBox.Mess[9]+' '+MsgBox.Mess[10] ;
    BSecondaire.Height:=BSecondaire.Height+IniGen.Height ;
    TotBudSect.Caption:=MsgBox.Mess[6]+' '+MsgBox.Mess[10] ;
    BFPrimaire.Height:=BFPrimaire.Height+TotBudSect.Height ;
    TotBudRecap.Caption:=MsgBox.Mess[9]+' '+MsgBox.Mess[10] ; ;
    BRecap.Height:=BRecap.Height+TotBudRecap.Height ;
    QBud.Caption:=MsgBox.Mess[6]+' '+MsgBox.Mess[10] ;
    BMulti.Height:=BMulti.Height+QBud.Height ;
    RTotEdtBud.Caption:=MsgBox.Mess[6]+' '+MsgBox.Mess[10] ;
    BFinEtat.Height:=BFinEtat.Height+RTotEdtBud.Height ;
    END ;
{ Calcul des différentes fourchettes de dates }
OkEnTeteGen:=True ;
OkRappelEnteteLibre:=FALSE ;
If CritEdt.Rupture=rRuptures then
   BEGIN
   if CritEdt.BalBud.RuptOnly=Sur then OkEnTeteGen:=False ;
   If CritEdt.BalBud.AvecCptSecond then OkEnTeteGen:=False ;
   END ;
END ;

Procedure TFQRCatSG.FinirPrint ;
BEGIN
  inherited ;
if (CritEdt.Rupture=rRuptures) then VideRupt(LMulti) ;
if (CritEdt.Rupture=rLibres) Or (CritEdt.Rupture=rRuptures) then if CritEdt.BalBud.AvecCptSecond then VideRecap(LRecap) ;
If (CritEdt.Rupture=rRien) then
   BEGIN
   if ListeAnaGene<>nil then BEGIN ListeAnaGene.Clear ; ListeAnaGene.Free ; END ;
   if ListeTotAnaGene<>nil then BEGIN ListeTotAnaGene.Clear ; ListeTotAnaGene.Free ; END ;
   END ;
//if CritEdt.Balbud.FormatPrint.PrSepMontant then FreeLesLignes(Self) ;
END ;

procedure TFQRCatSG.GenereSQL ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
Q.Close ; Q.SQL.Clear ;
Q.SQL.Add(' Select BS_AXE,BS_BUDSECT,BS_LIBELLE,BS_SECTIONTRIE,BS_RUB') ;
Q.SQL.Add(' From BUDSECT S Where ') ;
Q.SQL.Add(ExistBud(AxeToFbBud(CritEdt.BalBud.Axe),CritEdt.BalBud.MvtSur,CritEdt.BalBud.Categorie,CritEdt.BalBud.Axe,True)) ;
(*
Q.SQL.Add(' Exists  (select BE_BUDGENE, BE_AXE, BE_BUDSECT from BUDECR  ') ;
Q.SQL.Add(' where BE_BUDSECT=:BS_BUDSECT and BE_BUDGENE=G.BG_BUDGENE and BE_AXE="'+CritEdt.BalBud.Axe+'" ') ;
Q.SQL.Add(' and BE_BUDJAL="'+CritEdt.BalBud.Journal+'" ) ') ;
*)
Q.SQL.Add(' AND BS_CATEGORIE="'+CritEdt.BalBud.Categorie+'" ') ;
if CritEdt.Joker then Q.SQL.Add(' AND BS_BUDSECT like "'+TraduitJoker(CritEdt.Cpt1)+'" ') Else
   BEGIN
   if CritEdt.Cpt1<>'' then Q.SQL.Add(' AND BS_BUDSECT>="'+CritEdt.Cpt1+'" ') ;
   if CritEdt.Cpt2<>'' then Q.SQL.Add(' AND BS_BUDSECT<="'+CritEdt.Cpt2+'" ') ;
   END ;
Q.SQL.Add(' Order By BS_AXE,BS_BUDSECT') ;
ChangeSQL(Q) ; Q.Open ;
GenereSQL2 ;
END ;

procedure TFQRCatSG.GenereSQL2 ;
BEGIN
QSecond.Close ; QSecond.SQL.Clear ;
QSecond.SQL.Add(' Select BG_BUDGENE,BG_LIBELLE,BG_RUB ') ;
QSecond.SQL.Add(' From BUDGENE G Where ') ;
//QSecond.SQL.Add('BG_BUDGENE<>""') ;
QSecond.SQL.Add(ExistBud(fbBudgen,CritEdt.BalBud.MvtSur,CritEdt.BalBud.Categorie,CritEdt.BalBud.Axe,False)) ;
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
ChangeSQL(QSecond) ; QSecond.prepare ; QSecond.Open ;
END ;

procedure TFQRCatSG.ChoixEdition ;
BEGIN
  inherited ;
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
If CritEdt.BalBud.Resol='C' then CritEdt.Decimale:=V_PGI.OkDecV else
 If CritEdt.BalBud.Resol='F' then CritEdt.Decimale:=0 else
  If CritEdt.BalBud.Resol='K' then CritEdt.Decimale:=0 else
   If CritEdt.BalBud.Resol='M' then CritEdt.Decimale:=0 ;

ChangeMask(Formateur,CritEdt.Decimale,CritEdt.Symbole);
CritEdt.FormatMontant:=Formateur.Masks.PositiveMask;
//if CritEdt.Balbud.FormatPrint.PrSepMontant then AffecteLigne(Self,EntetePage,CritEdt.BalBud.FormatPrint) ;
END ;

procedure TFQRCatSG.RenseigneCritere ;
BEGIN
  inherited ;
If OkZoomEdt Then Exit ;
if CritEdt.BalBud.ColonnageJal then TRCategorie.Caption:=MsgBox.Mess[7]
                               else TRCategorie.Caption:=MsgBox.Mess[8] ;
RCategorie.Caption:=FCategorie.Text ;
RExercice_.Caption:=FExercice_.Text ;
CaseACocher(FValide,RValide) ;
CaseACocher(FQte,RQte) ;
RChoixMontant.Caption:=FChoixMontant.Text ;
RResol.Caption:=FResol.Text ;
RTaux.Caption:=FTaux.Text ;
if CritEdt.SJoker then
   BEGIN
   TRGeneral.Caption:=MsgBox.Mess[5] ;
   RGeneral1.Caption:=CritEdt.SCpt1 ;
   END Else
   BEGIN
   TRGeneral.Caption:=MsgBox.Mess[4] ;
   RGeneral1.Caption:=CritEdt.LSCpt1 ; RGeneral2.Caption:=CritEdt.LScpt2 ;
   END ;
RGeneral2.Visible:=Not CritEdt.SJoker ; TRaG.Visible:=Not CritEdt.SJoker ;
END ;

procedure TFQRCatSG.InfosJournal ;
Var QJal : TQuery ;
BEGIN
QJal:=OpenSQL('Select BJ_EXODEB, BJ_EXOFIN, BJ_PERDEB, BJ_PERFIN, BJ_AXE, BJ_NATJAL From BUDJAL Where BJ_BUDJAL="'+FCategorie.Value+'"',True) ;
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

procedure TFQRCatSG.CalculDateBud ;
Var i : Integer ;
    LaDate : TDateTime ;
    Pm, PA : Word ;
BEGIN
Fillchar(TabDate,SizeOf(TabDate),#0) ;
LaDate:=StrToDate(FPeriode1.Value) ;
NbMois:=0 ;
NOMBREMOIS (CritEdt.Date1, CritEdt.Date2, Pm, PA, NbMois) ;
If NbMois>MaxColonnes-1 then NbMois:=MaxColonnes-1 ;
For i:=1 to NbMois do
    BEGIN
    TabDate[i]:=PlusMois(LaDate,i-1) ;
    END ;
END ;

procedure TFQRCatSG.AffEntetesCol ;
//GP
var i : integer ;
    St,s,s1 : String ;
    C : TComponent ;
begin
NbColAff:=0 ;
if CritEdt.BalBud.ColonnageJal then
  BEGIN
  s := FListeCateg.Text ;
  s1:=ReadTokenSt(s) ; 
  for i:=1 to MaxColonnes-1 do
    BEGIN
    C:=FindComponent('Col'+IntToStr(i)) ;
    If C<>NIL Then
      BEGIN
      if s1<>'' then
         BEGIN
         TQRLabel(C).Caption:=RechDom('ttBudJal',s1,False) ;
         TQRLabel(C).SynData:=s1 ;
         Inc(NbColAff) ;
         END else Break ;
      s1:=ReadTokenSt(s) ;
      END ;
    END ;
  END ;
if not CritEdt.BalBud.ColonnageJal then
  BEGIN
  CalculDateBud ;
  { Titres des colonnes de montants }
  for i:=1 to NbMois do
     BEGIN
     C:=FindComponent('Col'+IntToStr(i)) ;
     if C<>nil then
        BEGIN
        TQRLabel(C).SynData:=DateToStr(TabDate[i])+';'+DateToStr(FinDeMois(TabDate[i])) ;
        St:=FormatDatetime('mmm yy',TabDate[i]) ;
        TQRLabel(C).Caption:=FirstMajuscule(St) ;
        Inc(NbColAff) ;
        END ;
     END ;
  END ;
// Total
TQRLabel(FindComponent('Col'+IntToStr(NbColAff+1))).Caption:=MsgBox.Mess[6] ;
end ;

procedure TFQRCatSG.ZoomJALClick(Sender: TObject);
var StW    : string ;
    ListeC : String ;
begin
StW:='' ;
if FCategorie.Value<>'' then StW:='BJ_CATEGORIE="'+FCategorie.Value+'"' ;
ListeC:=ChoisirMulti(HMCrit.Mess[10],'BUDJAL','BJ_BUDJAL','BJ_LIBELLE',StW,'BJ_BUDJAL',FListeCateg.Text) ;
if ListeC<>'???' then FListeCateg.Text:=ListeC ;
end;

procedure TFQRCatSG.BHCptBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
BG_BUDGENE.Caption:=MsgBox.Mess[1]+'  '+Qr12BG_BUDGENE.AsString ;
BG_LIBELLE.Caption:=Qr12BG_LIBELLE.AsString ;
Quoi:=QuoiCpt(1) ;
end;

procedure TFQRCatSG.BDetailBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
Fillchar(TotCptRea,SizeOf(TotCptRea),#0) ; Fillchar(TotCptBud,SizeOf(TotCptBud),#0) ;
BS_BUDSECT.Caption:=MsgBox.Mess[0]+'  '+Qr1BS_BUDSECT.AsString ;
BS_LIBELLE.Caption:=Qr1BS_LIBELLE.AsString ;
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
Quoi:=QuoiCpt(0) ;
end;

procedure TFQRCatSG.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr1BS_AXE         :=TStringField(Q.FindField('BS_AXE'));
QR1BS_BUDSECT     :=TStringField(Q.FindField('BS_BUDSECT')) ;
QR1BS_LIBELLE     :=TStringField(Q.FindField('BS_LIBELLE')) ;
QR1BS_SECTIONTRIE :=TStringField(Q.FindField('BS_SECTIONTRIE')) ;
QR1BS_RUB         :=TStringField(Q.FindField('BS_RUB')) ;
end;

procedure TFQRCatSG.QSecondAfterOpen(DataSet: TDataSet);
begin
  inherited;
QR12BG_BUDGENE:=TStringField(QSecond.FindField('BG_BUDGENE')) ;
QR12BG_LIBELLE:=TStringField(QSecond.FindField('BG_LIBELLE')) ;
QR12BG_RUB    :=TStringField(QSecond.FindField('BG_RUB')) ;
end;

procedure TFQRCatSG.FGene1Change(Sender: TObject);
Var AvecJokerS : Boolean ;
begin
  inherited;
AvecJokerS:=Joker(FGene1, FGene2, FJokerGene ) ;
TFaB.Visible:=Not AvecJokerS ;
TFGene.Visible:=Not AvecJokerS ;
TFJokerGene.Visible:=AvecJokerS ;
end;

procedure TFQRCatSG.FormShow(Sender: TObject);
begin
  inherited;
HelpContext:=15340000 ;
Standards.HelpContext:=15340100 ;
Avances.HelpContext:=15340200 ;
Mise.HelpContext:=15340300 ;
Option.HelpContext:=15340400 ;
TabRuptures.HelpContext:=15340500 ;
Floading:=FALSE ;
FRuptures.Enabled:=True ; FRuptures.Checked:=True ;
TabSup.TabVisible:=False ;
TabRuptures.Visible:=False ;
If FCategorie.Values.Count>0 Then FCategorie.Value:=FCategorie.Values[0] ;
InitResolution(FRESOL) ;
FLigneCpt.Checked:=True ;
TFGen.Parent:=Avances ; TFGenJoker.Parent:=Avances ;
FJoker.Parent:=Avances ; TFaG.Parent:=Avances ;
FCpte1.Parent:=Avances ; FCpte2.Parent:=Avances ;
FCpte1.TabOrder:=0 ; FCpte2.TabOrder:=1 ; FJoker.TabOrder:=2 ;
DebToDate ; FinToDate  ; //FSoldeFormate.ItemIndex:=0 ;
FPlansCo.Enabled:=False ;
//FNatBudChange(Nil) ;
end;

procedure TFQRCatSG.DebToDate ;
BEGIN
FDateCompta1.Text:=FPeriode1.Value ;
END ;

procedure TFQRCatSG.FinToDate ;
BEGIN
FDateCompta2.Text:=FPeriode2.Value ;
END ;

procedure TFQRCatSG.FExerciceChange(Sender: TObject);
begin
  inherited ;
if FExercice.Value>FExercice_.Value then FExercice_.Value:=FExercice.Value ;
ListePeriode(FExercice.Value,FPeriode1.Items,FPeriode1.Values,True) ;
FPeriode1.ItemIndex:=0 ; DebToDate ;
end;

procedure TFQRCatSG.FExercice_Change(Sender: TObject);
begin
  inherited;
if FExercice_.Value<FExercice.Value then FExercice.Value:=FExercice_.Value ;
ListePeriode(FExercice_.Value,FPeriode2.Items,FPeriode2.Values,False) ;
FPeriode2.ItemIndex:=FPeriode2.Items.Count-1 ; FinToDate ;
end;

procedure TFQRCatSG.FPeriode1Change(Sender: TObject);
begin
  inherited;
if StrToDate(FPeriode1.Value)>StrToDate(FPeriode2.Value) then FPeriode2.Value:=DateToStr(FinDeMois(StrToDate(FPeriode1.Value))) ;
FDateCompta1.Text:=FPeriode1.Value ;
end;

procedure TFQRCatSG.FPeriode2Change(Sender: TObject);
begin
  inherited;
if StrToDate(FPeriode2.Value)<StrToDate(FPeriode1.Value) then FPeriode1.Value:=DateToStr(DebutDeMois(StrToDate(FPeriode2.Value))) ;
FDateCompta2.Text:=FPeriode2.Value ;
end;

procedure TFQRCatSG.BNouvRechClick(Sender: TObject);
begin
  inherited;
FCategorie.ItemIndex:=0 ;
InitResolution(FRESOL) ;
end;

procedure TFQRCatSG.BFPrimaireBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var i : Integer ;
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
BS_BUDSECT_.Caption:=MsgBox.Mess[3]+'  '+QR1BS_BUDSECT.AsString ;
For i:=1 To NbColAff Do
  BEGIN
  TQRLabel(FindComponent('TotEdtRea'+IntToStr(i))).Caption:=PrintSoldeFormate(TotCptRea[i].TotDebit, TotCptRea[i].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
  TQRLabel(FindComponent('TotEdtBud'+IntToStr(i))).Caption:=PrintSoldeFormate(TotCptBud[i].TotDebit, TotCptBud[i].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
  END ;
// totaux
TQRLabel(FindComponent('TotEdtRea'+IntToStr(NbColAff+1))).Caption:=PrintSoldeFormate(TotCptRea[0].TotDebit, TotCptRea[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
TQRLabel(FindComponent('TotEdtBud'+IntToStr(NbColAff+1))).Caption:=PrintSoldeFormate(TotCptBud[0].TotDebit, TotCptBud[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
Quoi:=QuoiCpt(0) ;
end;

procedure TFQRCatSG.AlimMontant(Var TotBud : TMontTot ; Var TotMulti : TabMont77) ;
var i  : integer ;
    St,Jal,RubCpt,RubSect : String ;
    C,C1,C2 : TComponent ;
    T,TMR,TMB : TQrLABEL ;
    Compte1,Compte2,LExo : String ;
    TD      : TabloExt ;
    RealPer : TabTot12 ;
    OkMR,OkMB : Boolean ;
    XX : Double ;
BEGIN
Fillchar(TotBud,SizeOf(TotBud),#0) ;
Fillchar(RealPer,SizeOf(RealPer),#0) ;
T:=NIL ; TMB:=NIL ; TMR:=NIL ;
for i:=1 To NbColAff do
  BEGIN
  OkMR:=FALSE ; OkMB:=FALSE ;
  C:=FindComponent('Col'+IntToStr(i)) ;
  If C<>NIL Then
     BEGIN
     T:=TQRLabel(C) ;
     C1:=FindComponent('MontantRea'+IntToStr(i)) ;
     If C1<>NIL Then BEGIN TMR:=TQRLabel(C1) ; OkMR:=TRUE ; END ;
     C2:=FindComponent('MontantBud'+IntToStr(i)) ;
     If C2<>NIL Then BEGIN TMB:=TQRLabel(C2) ; OkMB:=TRUE ; END ;
     END Else continue ;
  If C<>NIL Then
     BEGIN
     If CritEdt.Monoexo Then LExo:=CritEdt.BalBud.Exo1 Else LExo:='' ;
     if CritEdt.BalBud.ColonnageJal then
        BEGIN
        Jal:=T.SynData ;
        END else
        BEGIN
        Jal:=CritEdt.BalBud.Journal ;
        St:=T.SynData ;
        CritEdt.Date1:=StrToDate(ReadTokenSt(St)) ;
        CritEdt.BalBud.Date21:=StrToDate(St) ;
        If not CritEdt.Monoexo Then LExo:=QUELEXODT(CritEdt.Date1) ;
        END ;
     If Jal='' Then Continue ;
     If CritEdt.BalBud.PasDeRubrique Then
        BEGIN
        Compte1:=Qr1BS_AXE.AsString+QR1BS_BUDSECT.AsString ;
        Compte2:=QR12BG_BUDGENE.AsString ;
        XX:=GetCumul('BUDGET:A/GREA',Compte1,Compte2,CritEdt.BalBud.SANbud,CritEdt.Etab,Jal,LExo,CritEdt.Date1,CritEdt.BalBud.Date21,false,FALSE,ListeAnaGene,TD,CritEdt.Monnaie=2) ;
        // 4 : TC ; 5 : TD
        TD[4]:=0 ; TD[5]:=0 ;
        If XX>0 Then TD[5]:=XX Else TD[4]:=Abs(XX) ;
        END Else
        BEGIN
        RubCpt:=QR12BG_RUB.AsString ;
        RubSect:=QR1BS_RUB.AsString ;
        Compte1:='S/G'+CritEdt.BalBud.Journal+RubSect+':'+RubCpt ;
        GetCumul('RUBREA',Compte1,'',CritEdt.BalBud.SANbud,CritEdt.Etab,CritEDT.DeviseAffichee,LExo,CritEdt.Date1,CritEdt.BalBud.Date21,false,FALSE,ListeAnaGene,TD,CritEdt.Monnaie=2) ;
        END ;
     Reevaluation(TD[5],TD[4],CritEdt.BalBud.Resol,CritEdt.BalBud.Taux) ;
     RealPer[0].TotDebit:=TD[5] ; RealPer[0].TotCredit:=TD[4] ;
     RealPer[2].TotDebit:=RealPer[2].TotDebit+RealPer[0].TotDebit ; RealPer[2].TotCredit:=RealPer[2].TotCredit+RealPer[0].TotCredit ;
     MontantAnaGene(i) ;
     If OkMR Then TMR.Caption:=PrintSoldeFormate(RealPer[0].TotDebit, RealPer[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
     TotMulti[(15*2)+i]:= Arrondi(TotMulti[(15*2)+i]+TD[1], CritEdt.Decimale) ;
     TotMulti[(15*3)+i]:= Arrondi(TotMulti[(15*3)+i]+TD[2], CritEdt.Decimale) ;
     // Totalisation
     TotMulti[(15*2)]:= Arrondi(TotMulti[(15*2)]+TD[1], CritEdt.Decimale) ;
     TotMulti[(15*3)]:= Arrondi(TotMulti[(15*3)]+TD[2], CritEdt.Decimale) ;

     TotEdtRea[i].TotDebit:= Arrondi(TotEdtRea[i].TotDebit+TD[1], CritEdt.Decimale) ;
     TotEdtRea[i].TotCredit:= Arrondi(TotEdtRea[i].TotCredit+TD[2], CritEdt.Decimale) ;
     // Totalisation
     TotEdtRea[0].TotDebit:= Arrondi(TotEdtRea[0].TotDebit+TD[1], CritEdt.Decimale) ;
     TotEdtRea[0].TotCredit:= Arrondi(TotEdtRea[0].TotCredit+TD[2], CritEdt.Decimale) ;

     TotCptRea[i].TotDebit:= Arrondi(TotCptRea[i].TotDebit+TD[1], CritEdt.Decimale) ;
     TotCptRea[i].TotCredit:= Arrondi(TotCptRea[i].TotCredit+TD[2], CritEdt.Decimale) ;
     // Totalisation
     TotCptRea[0].TotDebit:= Arrondi(TotCptRea[0].TotDebit+TD[1], CritEdt.Decimale) ;
     TotCptRea[0].TotCredit:= Arrondi(TotCptRea[0].TotCredit+TD[2], CritEdt.Decimale) ;
     If CritEdt.BalBud.PasDeRubrique Then
        BEGIN
        XX:=GetCumul('BUDGET:A/G',Compte1,Compte2,CritEdt.BalBud.SANbud,CritEdt.Etab,Jal,LExo,CritEdt.Date1,CritEdt.BalBud.Date21,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
        TD[4]:=0 ; TD[5]:=0 ;
        If XX>0 Then TD[5]:=XX Else TD[4]:=Abs(XX) ;
        END Else
        BEGIN
        GetCumul('RUBBUD',Compte1,'',CritEdt.BalBud.SANbud,CritEdt.Etab,
                 CritEDT.DeviseAffichee,LExo,CritEdt.Date1,CritEdt.BalBud.Date21,false,TRUE,Nil,TD,CritEdt.Monnaie=2) ;
        END ;
     Reevaluation(TD[5],TD[4],CritEdt.BalBud.Resol,CritEdt.BalBud.Taux) ;
     RealPer[1].TotDebit:=TD[5] ; RealPer[1].TotCredit:=TD[4] ;
     RealPer[3].TotDebit:=RealPer[3].TotDebit+RealPer[1].TotDebit ; RealPer[3].TotCredit:=RealPer[3].TotCredit+RealPer[1].TotCredit ;
     If OkMB Then TMB.Caption:=PrintSoldeFormate(RealPer[1].TotDebit, RealPer[1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
     TotMulti[i]:= Arrondi(TotMulti[i]+TD[1], CritEdt.Decimale) ;
     TotMulti[MaxTabCol+i]:= Arrondi(TotMulti[MaxTabCol+i]+TD[2], CritEdt.Decimale) ;
     //Totalisation
     TotMulti[0]:= Arrondi(TotMulti[0]+TD[1], CritEdt.Decimale) ;
     TotMulti[MaxTabCol]:= Arrondi(TotMulti[MaxTabCol]+TD[2], CritEdt.Decimale) ;
     TotEdtBud[i].TotDebit:= Arrondi(TotEdtBud[i].TotDebit+TD[1], CritEdt.Decimale) ;
     TotEdtBud[i].TotCredit:= Arrondi(TotEdtBud[i].TotCredit+TD[2], CritEdt.Decimale) ;
     //Totalisation
     TotEdtBud[0].TotDebit:= Arrondi(TotEdtBud[0].TotDebit+TD[1], CritEdt.Decimale) ;
     TotEdtBud[0].TotCredit:= Arrondi(TotEdtBud[0].TotCredit+TD[2], CritEdt.Decimale) ;

     TotCptBud[i].TotDebit:= Arrondi(TotCptBud[i].TotDebit+TD[1], CritEdt.Decimale) ;
     TotCptBud[i].TotCredit:= Arrondi(TotCptBud[i].TotCredit+TD[2], CritEdt.Decimale) ;
     //Totalisation
     TotCptBud[0].TotDebit:= Arrondi(TotCptBud[0].TotDebit+TD[1], CritEdt.Decimale) ;
     TotCptBud[0].TotCredit:= Arrondi(TotCptBud[0].TotCredit+TD[2], CritEdt.Decimale) ;
     END ;
  END ;
TQRLabel(Findcomponent('MontantRea'+IntToStr(NbColAff+1))).Caption:=PrintSoldeFormate(RealPer[2].TotDebit, RealPer[2].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
TQRLabel(Findcomponent('MontantBud'+IntToStr(NbColAff+1))).Caption:=PrintSoldeFormate(RealPer[3].TotDebit, RealPer[3].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
END ;

procedure TFQRCatSG.BSecondaireBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var LaRupture : TRuptInf ;
    TotMulti : TabMont77 ;
    St : String ;
    TotBud : TMontTot ;
begin
  inherited;
Fillchar(TotMulti,SizeOf(TotMulti),#0) ;
AlimMontant(TotBud,TotMulti) ;
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
Quoi:=QuoiCpt(2) ;
OkEnTeteGen:=PrintBand ;
end;

procedure TFQRCatSG.FCategorieChange(Sender: TObject);
var Q1 : TQuery ;
    s : string ;
    i : integer ;
begin
  inherited;
if FListeCateg.Text<>'' then exit ;
Q1:=OpenSQL('SELECT BJ_BUDJAL FROM BUDJAL WHERE BJ_CATEGORIE="'+FCategorie.Value+'"',True) ;
s:='' ; i:=0 ;
While not Q1.Eof do
  BEGIN
  s:=s+Q1.Fields[0].AsString+';' ;
  Q1.Next ;
  inc(i) ;
  if i>MaxColonnes-1 then Break ;
  END ;
Ferme(Q1) ;
FListeCateg.Text:=s ;
end;

procedure TFQRCatSG.DLMultiNeedData(var MoreData: Boolean);
var TotMulti         : TabMont77 ;
    Librupt, CodRupt : String ;
    Chaque,i         : Byte ;
//    TotMontLibre     : TabTot12 ;
    D,C : Double ;
    StM : String ;
begin
  inherited;
MoreData:=False ;
If (CritEdt.BalBud.RuptOnly<>Sans) then
   BEGIN
   Case CritEdt.Rupture of
     rRuptures : MoreData:=PrintRupt(LMulti,Qr1BS_SECTIONTRIE.AsString,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotMulti) ;
     End ;
   END ;
OkEnteteLibre:=MoreData ;
if MoreData then
   BEGIN
   {Affichage de la bande Rupture ou Nature libre}
   TLibRupt.Caption:='' ; TLibRupt.Left:=BG_LIBELLE.Left ;
   BRappelLibreHaut.Height:=HauteurBandeRuptIni ;
   if CritEdt.Rupture<>rLibres then
      BEGIN
      TCodRupt.Caption:=CodRupt ;
      TLibRupt.Caption:=LibRupt ;
      TLibRupt.Left:=BS_LIBELLE.Left ;
      END ;
   TRappelCodRupt.Caption:=TCodRupt.Caption ;
   TRappelLibRupt.Caption:=TLibRupt.Caption ;
   TRappelLibRupt.Left:=TLibRupt.Left ;
   For Chaque:=1 to 2 do { Réa,Budgété }
       BEGIN
       for i:=1 to NbColAff+1 do
           BEGIN
           Case Chaque of
              1 :  BEGIN { Réalisé }
                   if i<NbColAff+1 then
                     BEGIN
                     D:=TotMulti[(15*2)+i] ; C:=TotMulti[(15*3)+i] ;
                     END else
                     BEGIN
                     D:=TotMulti[(15*2)] ; C:=TotMulti[(15*3)] ;
                     END ;
                   StM:=PrintSoldeFormate(D,C, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
                   TQRLabel(FindComponent('QReaLibre'+IntToStr(i))).Caption:=StM ;
                   END ;
              2 :  BEGIN { Budgété }
                   if i<NbColAff+1 then
                     BEGIN
                     D:=TotMulti[i] ; C:=TotMulti[MaxTabCol+i] ;
                     END else
                     BEGIN
                     D:=TotMulti[0] ; C:=TotMulti[MaxTabCol] ;
                     END ;
                   StM:=PrintSoldeFormate(D,C, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
                   TQRLabel(FindComponent('QBudLibre'+IntToStr(i))).Caption:=StM;
                   END ;
              End ;
           END ;
       END ;
   END ;
end;

procedure TFQRCatSG.FPlanRupturesChange(Sender: TObject);
var Q1  : TQuery ; I : Integer ; St : String ;
begin
  inherited;
If QRLoading Then Exit ;
if FRuptures.Checked then  { Rempli la Fourchette des codes ruptures en 'Sur Rupture'  }
   BEGIN
   St:='' ;
   FCodeRupt1.Clear ; FCodeRupt2.clear ;
   Q1:=OpenSQL('Select RU_CLASSE from Rupture where RU_NATURERUPT="RU'+Char(FNatureCpt.Value[2])+'" and RU_PLANRUPT="'+FPlanRuptures.Value+'" '
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

procedure TFQRCatSG.FormCreate(Sender: TObject);
begin
  inherited;
ListeCodesRupture:=TStringList.Create ;
end;

procedure TFQRCatSG.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
ListeCodesRupture.Free ;
end;

procedure TFQRCatSG.BMultiBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.BalBud.RuptOnly<>Sans) ;
end;

procedure TFQRCatSG.BMultiAfterPrint(BandPrinted: Boolean);
begin
  inherited;
if (CritEdt.Rupture=rRuptures) then
   if CritEdt.BalBud.AvecCptSecond then BEGIN VideRecap(LRecap) ; ChargeRecap(LRecap) ; END ;
end;

procedure TFQRCatSG.DHLRappelLibreNeedData(var MoreData: Boolean);
begin
  inherited;
MoreData:=OkRappelEnteteLibre ;
OkRappelEnteteLibre:=False ;
end;

procedure TFQRCatSG.DLHLibreNeedData(var MoreData: Boolean);
begin
  inherited;
MoreData:=OkEnteteLibre ;
OkEnteteLibre:=False ;
if MoreData And (CritEdt.Rupture=rRuptures) then
   if CritEdt.BalBud.AvecCptSecond then OkRappelEnteteLibre:=TRUE ;
end;

procedure TFQRCatSG.BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var i : Integer ;
begin
  inherited;
For i:=1 To NbColAff Do
  BEGIN
  TQRLabel(FindComponent('TotGenRea'+IntToStr(i))).Caption:=PrintSoldeFormate(TotEdtRea[i].TotDebit, TotEdtRea[i].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
  TQRLabel(FindComponent('TotGenBud'+IntToStr(i))).Caption:=PrintSoldeFormate(TotEdtBud[i].TotDebit, TotEdtBud[i].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
  END ;
//Totaux
TQRLabel(FindComponent('TotGenRea'+IntToStr(NbColAff+1))).Caption:=PrintSoldeFormate(TotEdtRea[0].TotDebit, TotEdtRea[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
TQRLabel(FindComponent('TotGenBud'+IntToStr(NbColAff+1))).Caption:=PrintSoldeFormate(TotEdtBud[0].TotDebit, TotEdtBud[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
end;

procedure TFQRCatSG.FSansRuptClick(Sender: TObject);
begin
  inherited;
if Not FSurRupt.Checked then FAvecCptSecond.Checked:=False ;
FAvecCptSecond.Enabled:=FSurRupt.Checked ;
if FSurRupt.Checked or FAvecRupt.Checked then
  BEGIN
  FAnaGene.Checked:=False ; FAnaGene.Enabled:=False ;
  END else FAnaGene.Enabled:=True ;
end;

procedure TFQRCatSG.FRupturesClick(Sender: TObject);
begin
  inherited;
FAvecCptSecond.Enabled:=FSurRupt.Checked ;
If Not FSurRupt.Checked Then FAvecCptSecond.checked:=FALSE ;
if FRuptures.Checked then If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
end;

procedure TFQRCatSG.DLRecapNeedData(var MoreData: Boolean);
var TotRecap         : TabMont77 ;
    LibRecap, CodRecap : String ;
    Chaque,i         : Byte ;
    StM              : String ;
    D,C              : Double ;
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
      If (CritEdt.Rupture=rRuptures) Then BG_LIBELLE2.Caption:=LibRecap
         (*Else BG_LIBELLE2.Caption:=LibRecap+StRappelLibTableLibre*) ;
   For Chaque:=1 to 2 do { Réa,Budgété }
       BEGIN
       for i:=1 to NbColAff+1 do
           BEGIN
           Case Chaque of
              1 :  BEGIN { Réalisé }
                   if i<NbColAff+1 then
                     BEGIN
                     D:=TotRecap[(15*2)+i] ; C:=TotRecap[(15*3)+i] ;
                     END else
                     BEGIN
                     D:=TotRecap[(15*2)] ; C:=TotRecap[(15*3)] ;
                     END ;
                   StM:=PrintSoldeFormate(D,C, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
                   TQRLabel(FindComponent('TotRRecap'+IntToStr(i))).Caption:=StM ;
                   END ;
              2 :  BEGIN { Budgété }
                   if i<NbColAff+1 then
                     BEGIN
                     D:=TotRecap[i] ; C:=TotRecap[MaxTabCol+i] ;
                     END else
                     BEGIN
                     D:=TotRecap[0] ; C:=TotRecap[MaxTabCol] ;
                     END ;
                   StM:=PrintSoldeFormate(D,C, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
                   TQRLabel(FindComponent('TotBRecap'+IntToStr(i))).Caption:=StM;
                   END ;
              End ;
           END ;
       END ;
      END ;
   END ;
OkEnTeteRecap:=MoreData ;
end;

procedure TFQRCatSG.DLHRecapNeedData(var MoreData: Boolean);
begin
  inherited;
MoreData:=OkEnTeteRecap ;
OkEnTeteRecap:=False ;
end;

procedure TFQRCatSG.DLHGenNeedData(var MoreData: Boolean);
begin
  inherited;
MoreData:=OkEnTeteGen ;
OkEnTeteGen:=False ;
end;

procedure TFQRCatSG.BRecapBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
if ((CritEdt.Rupture=rLibres) Or (CritEdt.Rupture=rRuptures)) And (CritEdt.BalBud.AvecCptSecond)
   then Quoi:=QuoiCptRecap(BG_BUDGENE2.Caption,BG_LIBELLE2.Caption) ;
end;

procedure TFQRCatSG.FAnaGeneClick(Sender: TObject);
begin
  inherited;
FGroupChoixRupt.Enabled:=not FAnaGene.Checked ;
end;

procedure TFQRCatSG.BFAnaGeneBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
PrintBand:=CritEdt.BalBud.AnaGene ;
Quoi:=QuoiCpt(0) ;
end;

procedure TFQRCatSG.MontantAnaGene(Indice  : Integer) ;
var i,j : Integer ;
    St,StG,StS : String ;
    StTot : String ;
    MD,MC : Double ;
    MDTot,MCTot : Double ;
    Find : boolean ;
    FTotAG : TFTotAnaGene ;
BEGIN
if ListeAnaGene=nil then Exit ;
for i:=0 to ListeAnaGene.Count-1 do
  BEGIN
  Find:=False ;
  St:=ListeAnaGene[i] ;
  StS:=ReadTokenSt(St) ;
  StG:=Copy(St,1,Pos(':',St)-1) ; Delete(St,1,Pos(':',St)) ;
  Delete(St,1,Pos(':',St)) ;Delete(St,1,Pos(':',St)) ;Delete(St,1,Pos(':',St)) ;
  // TC et TD en 4 et 5
  MC:=StrToFloat(Copy(St,1,Pos(':',St)-1)) ; Delete(St,1,Pos(':',St)) ;
  MD:=StrToFloat(Copy(St,1,Pos(':',St)-1)) ; Delete(St,1,Pos(':',St)) ;
  for j:=0 to ListeTotAnaGene.Count-1 do
    BEGIN
    FTotAG:=TFTotAnaGene(ListeTotAnaGene.Objects[j]) ;
    if (FTotAG.Sect=StS) and (FTotAG.Gene=StG) then
      BEGIN
      Find:=True ;
      MDTot:=StrToFloat(Copy(StTot,1,Pos(':',StTot)-1)) ; Delete(StTot,1,Pos(':',StTot)) ;
      MCTot:=StrToFloat(Copy(StTot,1,Pos(':',StTot)-1)) ; Delete(StTot,1,Pos(':',StTot)) ;
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

function TFQRCatSG.PrintAnaGene (var Sect,Gene : String17 ; var Tot : TabMont77) : boolean ;
var FTotAG : TFTotAnaGene ;
BEGIN
Result:=False ;
if (ListeTotAnaGene=nil) or (ListeTotAnaGene.Count=0) then Exit ;
FTotAG := TFTotAnaGene(ListeTotAnaGene.Objects[0]) ;
Tot:=FTotAG.TotAG ;
Sect:=FTotAG.Sect ; Gene:=FTotAG.Gene ;
ListeTotAnaGene.Objects[0].Free ;
ListeTotAnaGene.Delete(0) ;
Result:=True ;
END ;

procedure TFQRCatSG.DLAnaGeneNeedData(var MoreData: Boolean);
var    Sect,Gene : String17 ;
       Tot       : TabMont77 ;
       i         : integer ;
begin
  inherited;
Sect:='' ; Gene:='' ;
Fillchar(Tot,SizeOf(Tot),#0) ; Fillchar(Tot,SizeOf(Tot),#0) ;
MoreData:=PrintAnaGene (Sect,Gene,Tot) ;
if MoreData then
  BEGIN
  TFAnaGene.Caption:=MsgBox.Mess[6]+' '+Sect+' / '+ Gene ;
  for i:=1 to NbColAff do
    BEGIN
    Tot[0]:=Tot[0]+Tot[i] ; Tot[MaxTabCol]:=Tot[MaxTabCol]+Tot[MaxTabCol+i] ;
    TQRLabel(FindComponent('TotAnaGene'+IntToStr(i))).Caption:=PrintSoldeFormate(Tot[i], Tot[MaxTabCol+i], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
    END ;
  TQRLabel(FindComponent('TotAnaGene'+IntToStr(NbColAff+1))).Caption:=PrintSoldeFormate(Tot[0], Tot[MaxTabCol], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.BalBud.SoldeFormate) ;
  END ;
end;

procedure TFQRCatSG.FColonnageClick(Sender: TObject);
var Val : String3 ;
begin
  inherited;
Val:=FCategorie.Value ;
Case FColonnage.ItemIndex of
//Journal
  0 : BEGIN
      ZoomJal.Visible:=True ;
      FCategorie.DataType:='ttCatJalBud' ;
      TFCategorie.Caption:=MsgBox.Mess[7] ;
      END ;
//Période
  1 : BEGIN
      ZoomJal.Visible:=False ;
      FCategorie.DataType:='ttBudJal' ;
      TFCategorie.Caption:=MsgBox.Mess[8] ;
      END ;
  END ;
FCategorie.Value:=Val ;
end;

end.
