unit QrBalAge;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Hctrls, Spin, Menus, hmsgbox, HQuickrp, StdCtrls,
  Buttons, ExtCtrls, Mask, Hcompte, ComCtrls, UtilEdt, HEnt1, Ent1,
  CpteUtil,
  EdtLegal,
  CritEDT, HQry, HSysMenu,UtilEdt1, QRRupt, HTB97, HPanel, UiUtil, tCalcCum ;

procedure BalanceAge ;
procedure BalanceAgeZoom(Crit : TCritEdt) ;

type
  TFQRBalAge = class(TFQR)
    HLabel3: THLabel;
    FChoixEcart: TRadioGroup;
    HLabel9: THLabel;
    FEcart: TSpinEdit;
    FColl: THLabel;
    FColl1: THCpteEdit;
    TFaC: TLabel;
    FColl2: THCpteEdit;
    FCollJoker: TEdit;
    TFCollJoker: THLabel;
    TFPaie: THLabel;
    E_MODEPAIE: THMultiValComboBox;
    TFSens: THLabel;
    FSens: THValComboBox;
    HLabel12: THLabel;
    FTypePar: TRadioGroup;
    RColl: TQRLabel;
    RColl1: TQRLabel;
    TRaC: TQRLabel;
    RColl2: TQRLabel;
    QRLabel10: TQRLabel;
    RE_MODEPAIE: TQRLabel;
    RSens: TQRLabel;
    TRSens: TQRLabel;
    TRSituation: TQRLabel;
    RParAux: TQRLabel;
    QRLabel9: TQRLabel;
    RParPaie: TQRLabel;
    QRLabel14: TQRLabel;
    TitreColCpt: TQRLabel;
    TitreColLibelle: TQRLabel;
    Col11: TQRLabel;
    Col1: TQRLabel;
    Col2: TQRLabel;
    Col22: TQRLabel;
    col3: TQRLabel;
    col33: TQRLabel;
    Col44: TQRLabel;
    Col4: TQRLabel;
    col5: TQRLabel;
    Col55: TQRLabel;
    Col66: TQRLabel;
    Col6: TQRLabel;
    Col7: TQRLabel;
    TITRE1REP: TQRLabel;
    REPORTCOL1: TQRLabel;
    REPORTCOL2: TQRLabel;
    REPORTCOL3: TQRLabel;
    REPORTCOL4: TQRLabel;
    REPORTCOL5: TQRLabel;
    REPORTCOL6: TQRLabel;
    REPORTTOTAL1: TQRLabel;
    BFCompteAux: TQRBand;
    TOTCLEF: TQRDBText;
    AuxCum1: TQRLabel;
    AuxCum2: TQRLabel;
    AuxCum4: TQRLabel;
    AuxCum3: TQRLabel;
    AuxCum0: TQRLabel;
    AuxTotal: TQRLabel;
    AuxCum5: TQRLabel;
    TOTLIB: TQRDBText;
    BMdp: TQRBand;
    QRLabel8: TQRLabel;
    MdpSolde5: TQRLabel;
    MdpSolde4: TQRLabel;
    MdpSolde3: TQRLabel;
    MdpSolde2: TQRLabel;
    MdpSolde1: TQRLabel;
    MdpSolde0: TQRLabel;
    MdpTotal: TQRLabel;
    QRLabel33: TQRLabel;
    TotCum5: TQRLabel;
    TotCum4: TQRLabel;
    TotCum3: TQRLabel;
    TotCum2: TQRLabel;
    TotCum1: TQRLabel;
    TotCum0: TQRLabel;
    TotTotal: TQRLabel;
    QROVERLAY: TQRBand;
    Trait7: TQRLigne;
    Trait6: TQRLigne;
    Trait5: TQRLigne;
    Trait1: TQRLigne;
    Trait2: TQRLigne;
    Trait3: TQRLigne;
    Trait4: TQRLigne;
    Trait0: TQRLigne;
    Trait8: TQRLigne;
    Ligne1: TQRLigne;
    TITRE2REP: TQRLabel;
    REPORTCOL7: TQRLabel;
    REPORTCOL8: TQRLabel;
    REPORTCOL9: TQRLabel;
    REPORTCOL10: TQRLabel;
    REPORTCOL11: TQRLabel;
    REPORTCOL12: TQRLabel;
    REPORTTOTAL2: TQRLabel;
    BSubDetail: TQRBand;
    SEcr: TDataSource;
    QEcr: TQuery;
    MsgBox: THMsgBox;
    GMdp: TQRGroup;
    QRDLAuxF: TQRDetailLink;
    TitreDates: TLabel;
    HLabel13: THLabel;
    HLabel14: THLabel;
    HLabel15: THLabel;
    HLabel16: THLabel;
    TFPeriodicite: TLabel;
    TFDateCpta2: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    FPeriodicite: THValComboBox;
    FP4: TMaskEdit;
    FP3: TMaskEdit;
    FP2: TMaskEdit;
    FP1: TMaskEdit;
    FP8: TMaskEdit;
    FP7: TMaskEdit;
    FP6: TMaskEdit;
    FP5: TMaskEdit;
    FTriPar: TRadioGroup;
    TFChoixMontant: THLabel;
    Bevel3: TBevel;
    BRupt: TQRBand;
    Libre2: TQRLabel;
    TCodRupt: TQRLabel;
    Libre1: TQRLabel;
    Libre0: TQRLabel;
    Libre5: TQRLabel;
    Libre4: TQRLabel;
    Libre3: TQRLabel;
    LibreTotal: TQRLabel;
    DLRupt: TQRDetailLink;
    FLigneRupt: TCheckBox;
    FChoixMontant: THValComboBox;
    procedure FormShow(Sender: TObject);
    procedure FChoixEcartClick(Sender: TObject);
    procedure FDateCompta1Exit(Sender: TObject);
    procedure FP1Exit(Sender: TObject);
    procedure PeriodesEnter(Sender: TObject);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BSubDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFCompteAuxBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BMdpBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FPeriodiciteChange(Sender: TObject);
    procedure FTypeParClick(Sender: TObject);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QEcrAfterOpen(DataSet: TDataSet);
    procedure FNatureCptChange(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure FColl1Change(Sender: TObject);
    procedure DLRuptNeedData(var MoreData: Boolean);
    procedure BRuptBeforePrint(var PrintBand: Boolean;  var Quoi: string);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure FSansRuptClick(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure FPlanRupturesChange(Sender: TObject);
  private    { Déclarations privées }
    LMdp, LRupt                    : TStringList ;
    TotEdt, TotAux, TotMdp         : TabTot ;
    TabDate                        : TTabDate4 ;
    Qr1T_AUXILIAIRE, Qr1T_LIBELLE,
    Qr1T_CORRESP1,Qr1T_CORRESP2    : TStringField ;
    QR2E_MODEPAIE                  : TStringField ;
    QR2E_DATECOMPTABLE             : TDateTimeField ;
    QR2COUVERTURE,
    QR2DEBIT, QR2CREDIT            : TFloatField ;
    FLoadColl, Affiche             : Boolean ;
    procedure GenereSQLSub ;
    Function  QuoiCpt(i : Integer) : String ;
    Procedure BalAgeZoom(Quoi : String) ;
    procedure SaisieDates ;
    Procedure Calculs(var Report1,Report2 : TReport; i : integer; Bool1, Bool2 : Boolean ) ;
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

Function TFQRBalAge.QuoiCpt(i : Integer) : String ;
BEGIN
Case CritEdt.GlV.TypePar of
  0 : BEGIN
      Case i Of
        0 : Result:=Q.FindField('T_AUXILIAIRE').AsString+' '+Q.FindField('T_LIBELLE').AsString+'#'+' '+'@'+'2' ;
        end ;
      END ;
  1 : BEGIN
      Case i Of
        0 : Result:=Q.FindField('MP_MODEPAIE').AsString+' '+Q.FindField('MP_LIBELLE').AsString+'#'+' '+'@'+'99' ;
        end ;
      END ;
  END ;
END ;

Procedure TFQRBalAge.BalAgeZoom(Quoi : String) ;
Var Lp,i: Integer ;
BEGIN
Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
i:=StrToInt(Copy(Quoi,Lp+1,1)) ;
ZoomEdt(i,Quoi) ;
END ;

procedure BalanceAge ;
var QR : TFQRBalAge ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQRBalAge.Create(Application) ;
Edition.Etat:=etBalAge ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalAgeZoom ;
QR.InitType (nbAux,neGlV,msRien,'QRBALAGEE','',TRUE,FALSE,FALSE,Edition) ;
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

procedure BalanceAgeZoom(Crit : TCritEdt) ;
var QR : TFQRBalAge ;
    Edition : TEdition ;
BEGIN
QR:=TFQRBalAge.Create(Application) ;
Edition.Etat:=etBalAge ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.BalAgeZoom ;
 QR.InitType (nbAux,neGlV,msRien,'QRBALAGEE','',FALSE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

Procedure TFQRBalAge.FinirPrint ;
Var Solde   : Double ;
BEGIN
Inherited ;
QEcr.Close ;
if CritEdt.Rupture<>rRien then VideRupt(LRupt) ;
//if (CritEdt.GlV.TypePar=0) and (CritEdt.Rupture=rLibres) then VideGroup(LRupt) ;
if OkMajEdt Then
   BEGIN
   Solde:=TotEdt[6].TotDebit-TotEdt[6].TotCredit ;
   if Solde<0
      then MajEdition('BLA', '', DateToStr(CritEdt.Date1), DateToStr(CritEdt.Date1),
                      '', TotEdt[6].TotDebit, TotEdt[6].TotCredit, Solde, 0)
      else MajEdition('BLA', '', DateToStr(CritEdt.Date1), DateToStr(CritEdt.Date1),
                      '', TotEdt[6].TotDebit, TotEdt[6].TotCredit, 0, Solde) ;
   END ;
END ;

procedure TFQRBalAge.InitDivers ;
BEGIN
Inherited ;
{ Calcul des différentes fourchettes de dates }
CalculDateTiers(neBal,Age,CritEdt,TabDate,FP1,FP2,FP3,FP4,FP5,FP6,FP7,FP8) ;
{ Titres des colonnes de montants }
Col1.Caption:=MsgBox.Mess[7] ;                              Col11.Caption:=DateToStr(TabDate[4]-1) ;
Col2.Caption:=MsgBox.Mess[4]+DateToStr(TabDate[4]) ;        Col22.Caption:=MsgBox.Mess[5]+DateToStr(TabDate[3]-1) ;
Col3.Caption:=MsgBox.Mess[4]+DateToStr(TabDate[3]) ;        Col33.Caption:=MsgBox.Mess[5]+DateToStr(TabDate[2]-1) ;
Col4.Caption:=MsgBox.Mess[4]+DateToStr(TabDate[2]) ;        Col44.Caption:=MsgBox.Mess[5]+DateToStr(TabDate[1]-1) ;
Col5.Caption:=MsgBox.Mess[4]+DateToStr(TabDate[1]) ;        Col55.Caption:=MsgBox.Mess[5]+DateToStr(CritEdt.Date1-1) ;
Col6.Caption:=MsgBox.Mess[3] ;                              Col66.Caption:=DateToStr(CritEdt.Date1) ; // Rony 9/04/97 Col66.Caption:=Format_String(' ',1) ;
BFCompteAux.Frame.DrawTop:=CritEdt.GlV.FormatPrint.PrSepCompte[1] ;
BRupt.Frame.DrawTop:=CritEdt.GlV.FormatPrint.PrSepCompte[3] ;
BRupt.Frame.DrawBottom:=BRupt.Frame.DrawTop ;
//BMdp.Frame.DrawTop:=CritEdt.FormatPrint.PrSepCompte[2] ; en attente
{ Edition par auxiliaire ou par mode de paiement }
Case CritEdt.GlV.TypePar of
  0 : BEGIN
      TitreColCpt.Caption:=MsgBox.Mess[8] ;
      TOTCLEF.DataField:='T_AUXILIAIRE' ;
      TOTLIB.DataField:='T_LIBELLE' ;
      END ;
  1 : BEGIN
      TitreColCpt.Caption:=MsgBox.Mess[9] ;
      TOTCLEF.DataField:='MP_MODEPAIE' ;
      TOTLIB.DataField:='MP_LIBELLE' ;
      END ;
 end ;
if Not E_MODEPAIE.Tous then ChargeRecapMPTiers(LMdp, E_MODEPAIE);
If CritEdt.GLV.TypePAR=0 Then
   BEGIN
   ChangeFormatCompte(fbAux,fbAux,Self,TOTCLEF.Width,1,2,TOTCLEF.Name) ;
   TotCleF.Alignment:=taLeftJustify ;
   END Else TotCleF.Alignment:=taCenter ;
END ;

procedure TFQRBalAge.GenereSQL ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
GenereSQLTiers(neBal,Age,CritEdt,Q,FALSE,AvecRevision.State) ;
GenereSQLSub ;
END ;

procedure TFQRBalAge.GenereSQLSub ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
{ Construction de la clause Select de la SQL }
GenereSQLSUBTiers(neBal,Age,CritEdt,QEcr,FALSE) ;
END ;

procedure TFQRBalAge.RenseigneCritere ;
Var St11,St22 : String ;
BEGIN
Inherited ;
IF CritEdt.SJoker then
   BEGIN
   RColl.Caption:=MsgBox.Mess[18] ;
   RColl1.Caption:=FCollJoker.Text ;
   END Else
   BEGIN
   RColl.Caption:=MsgBox.Mess[17] ;
   //RColl1.Caption:=FColl1.Text ; RColl2.Caption:=FColl2.Text ;
   PositionneFourchetteST(FColl1,FColl2,St11,St22) ;
   RColl1.Caption:=St11 ;
   RColl2.Caption:=St22 ;
   END ;
RColl2.Visible:=Not CritEdt.SJoker ; TRaC.Visible:=Not CritEdt.SJoker ;
RSens.Caption:=FSens.Text ; RE_MODEPAIE.Caption:=E_MODEPAIE.Text ;
Case FTypePar.ItemIndex of
  0 : BEGIN RParAux.Caption:='þ' ;    RParPaie.Caption:='o'  ; END ;
  1 : BEGIN RParAux.Caption:='o'  ;   RParPaie.Caption:='þ' ; END ;
end ;
END ;

procedure TFQRBalAge.ChoixEdition ;
{ Initialisation des options d'édition }
BEGIN
Inherited ;
DLRupt.PrintBefore:=TRUE ;
Case CritEdt.Rupture of
  rLibres   : BEGIN
              DLRupt.PrintBefore:=FALSE ;
              ChargeGroup(LRupt,['T00','T01','T02','T03','T04','T05','T06','T07','T08','T09']) ;
              END ;
  rRuptures : BEGIN
              ChargeRupt(LRupt, 'RUT', CritEdt.GlV.PlanRupt,CritEdt.GlV.CodeRupt1,CritEdt.GlV.CodeRupt2) ;
              NiveauRupt(LRupt);
              END ;
  rCorresp  : BEGIN
              ChargeRuptCorresp(LRupt, CritEdt.GlV.PlanRupt,CritEdt.GlV.CodeRupt1,CritEdt.GlV.CodeRupt2, False) ;
              NiveauRupt(LRupt);
              END ;
  End ;
END ;

procedure TFQRBalAge.RecupCritEdt ;
Var NonLibres : Boolean ;
BEGIN
Inherited ;
With CritEDT Do
  BEGIN
  SJoker:=FCollJoker.Visible ;
  if SJoker Then
     BEGIN
     SCpt1:=FCollJoker.Text ; SCpt2:=FCollJoker.Text ;
     LSCpt1:=SCpt1 ; LSCpt2:=SCpt2 ;
     END Else
     BEGIN
     SCpt1:=FColl1.Text ; SCpt2:=FColl2.Text ;
     PositionneFourchetteSt(FColl1,FColl2,CritEdt.LSCpt1,CritEdt.LSCpt2) ;
     END ;
  GlV.TriePar:=FTriPar.ItemIndex ;
  GlV.TypePar:=FTypePar.ItemIndex ;
  GlV.Sens:=FSens.ItemIndex ;
  GlV.Ecart:=FEcart.Value ;
  GlV.ChoixEcart:=FChoixEcart.ItemIndex ;
  GlV.SQLModePaie:=E_MODEPAIE.GetSQLValue ;
  GlV.RuptOnly:=QuelleTypeRupt(0,FSAnsRupt.Checked,FAvecRupt.Checked,FSurRupt.Checked) ;
  NonLibres:=((Rupture=rRuptures) or (Rupture=rCorresp)) ;
  If NonLibres Then GlV.PlanRupt:=FPlanRuptures.Value ;
  GlV.RuptOnly:=QuelleTypeRupt(0,FSAnsRupt.Checked,FAvecRupt.Checked,FSurRupt.Checked) ;
  If NonLibres Then BEGIN GlV.CodeRupt1:=FCodeRupt1.Text ; GlV.CodeRupt2:=FCodeRupt2.Text ; END ;
  GlV.OnlyCptAssocie:=(Rupture<>rRien) and FOnlyCptAssocie.Checked ;
  If (CritEdt.Rupture=rCorresp) Then GLV.PlansCorresp:=FPlanRuptures.ItemIndex+1 ;
  GLV.FormatPrint.PrSepCompte[3]:=FLigneRupt.Checked ;
  GLV.SoldeFormate:=FChoixMontant.Value ;
  END ;
END ;

function TFQRBalAge.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ;
If Result Then
   BEGIN
   if CritEdt.GLV.ChoixEcart=1 then Result:=ValidPeriodeTiers(FP1,FP2,FP3,FP4,FP5,FP6,FP7,FP8) ;
   Fillchar(TotEdt,SizeOf(Totedt),#0) ;
   END ;
END ;

Procedure TFQRBalAge.Calculs(var Report1,Report2 : TReport; i : integer; Bool1, Bool2 : Boolean ) ;
{ Incrémentations de toutes les variables de l'état }
Var TotRupt : Array[0..13] of double ;
    TotTemp : TabTot ;
    CptRupt : String ;
BEGIN
Fillchar(TotTemp,SizeOf(TotTemp),#0) ;
if QR2DEBIT.AsFloat<>0 then
   BEGIN
   if Bool1 then
      BEGIN
      Report1.TotDebit:= Arrondi(QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat,CritEdt.Decimale) ;
      Report2.TotDebit:=Report1.TotDebit ;
      TotAux[i].TotDebit:= Arrondi(TotAux[i].TotDebit+QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      { Gestion de la totalisation de l'édition en Non Sur Rupture }
      if CritEdt.GlV.RuptOnly<>Sur then TotEdt[i].TotDebit:= Arrondi(TotEdt[i].TotDebit+QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale)
                                   Else IF CritEdt.Rupture=rLibres then TotEdt[i].TotDebit:= Arrondi(TotEdt[i].TotDebit+QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      TotTemp[i].TotDebit:= Arrondi(TotTemp[i].TotDebit+QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      if Bool2 then
         TotAux[6].TotDebit:= Arrondi(TotAux[6].TotDebit+QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      if CritEdt.GlV.RuptOnly<>Sur then TotEdt[6].TotDebit:= Arrondi(Totedt[6].TotDebit+QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale)
                                   Else IF CritEdt.Rupture=rLibres then TotEdt[6].TotDebit:= Arrondi(Totedt[6].TotDebit+QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      TotTemp[6].TotDebit:= Arrondi(TotTemp[6].TotDebit+QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      END else
      BEGIN
      TotMdp[i].TotDebit:= Arrondi(TotMdp[i].TotDebit+QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      TotMdp[6].TotDebit:= Arrondi(TotMdp[6].TotDebit+QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      END ;
   END else
   BEGIN
   if Bool1 then
      BEGIN
      Report1.TotCredit:= Arrondi(QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat,CritEdt.Decimale) ;
      Report2.TotCredit:=Report1.TotCredit ;
      TotAux[i].TotCredit:= Arrondi(TotAux[i].TotCredit+QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      if CritEdt.GlV.RuptOnly<>Sur then TotEdt[i].TotCredit:= Arrondi(Totedt[i].TotCredit+QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale)
                                   Else IF CritEdt.Rupture=rLibres then TotEdt[i].TotCredit:= Arrondi(Totedt[i].TotCredit+QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      TotTemp[i].TotCredit:= Arrondi(TotTemp[i].TotCredit+QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      if Bool2 then
         TotAux[6].TotCredit:= Arrondi(TotAux[6].TotCredit+QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      if CritEdt.GlV.RuptOnly<>Sur then TotEdt[6].TotCredit:= Arrondi(TotEdt[6].TotCredit+QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale)
                                   Else IF CritEdt.Rupture=rLibres then TotEdt[6].TotCredit:= Arrondi(TotEdt[6].TotCredit+QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      TotTemp[6].TotCredit:= Arrondi(TotTemp[6].TotCredit+QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      END else
      BEGIN
      TotMdp[i].TotCredit:= Arrondi(TotMdp[i].TotCredit+QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      TotMdp[6].TotCredit:= Arrondi(TotMdp[6].TotCredit+QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      END ;
   END ;
TotRupt[0]:=Arrondi(TotTemp[0].TotDebit,CritEdt.Decimale) ;
TotRupt[1]:=Arrondi(TotTemp[0].TotCredit,CritEdt.Decimale) ;

TotRupt[2]:=Arrondi(TotTemp[1].TotDebit,CritEdt.Decimale) ;
TotRupt[3]:=Arrondi(TotTemp[1].TotCredit,CritEdt.Decimale) ;

TotRupt[4]:=Arrondi(TotTemp[2].TotDebit,CritEdt.Decimale) ;
TotRupt[5]:=Arrondi(TotTemp[2].TotCredit,CritEdt.Decimale) ;

TotRupt[6]:=Arrondi(TotTemp[3].TotDebit,CritEdt.Decimale) ;
TotRupt[7]:=Arrondi(TotTemp[3].TotCredit,CritEdt.Decimale) ;

TotRupt[8]:=Arrondi(TotTemp[4].TotDebit,CritEdt.Decimale) ;
TotRupt[9]:=Arrondi(TotTemp[4].TotCredit,CritEdt.Decimale) ;

TotRupt[10]:=Arrondi(TotTemp[5].TotDebit,CritEdt.Decimale) ;
TotRupt[11]:=Arrondi(TotTemp[5].TotCredit,CritEdt.Decimale) ;

TotRupt[12]:=Arrondi(TotTemp[6].TotDebit,CritEdt.Decimale) ;
TotRupt[13]:=Arrondi(TotTemp[6].TotCredit,CritEdt.Decimale) ;
Case CritEdt.Rupture of
  rLibres   : AddGroupLibre(LRupt,Q,fbAux,CritEdt.LibreTrie,TotRupt) ;
  rRuptures : AddRupt(LRupt,Qr1T_AUXILIAIRE.AsString,TotRupt) ;
  rCorresp  : BEGIN
              Case CritEDT.GlV.PlansCorresp Of
                1 : If Qr1T_CORRESP1.AsString<>'' Then CptRupt:=Qr1T_CORRESP1.AsString+Qr1T_AUXILIAIRE.AsString
                                                  Else CptRupt:='.'+Qr1T_AUXILIAIRE.AsString ;
                2 : If Qr1T_CORRESP2.AsString<>'' Then CptRupt:=Qr1T_CORRESP2.AsString+Qr1T_AUXILIAIRE.AsString
                                                  Else CptRupt:='.'+Qr1T_AUXILIAIRE.AsString ;
                Else CptRupt:=Qr1T_AUXILIAIRE.AsString ;
                End ;
              AddRuptCorres(LRupt,CptRupt,TotRupt) ;
              END ;
  End ;
END ;

procedure TFQRBalAge.SaisieDates ;
BEGIN { Report de la date d'arrêtée - 1 jours , et accées au zone de saisie de dates }
FP8.Text:=DateToStr(StrToDatetime(FDateCompta1.Text)-1) ; ;
FP1.Enabled:=True ; FP2.Enabled:=False ; FP3.Enabled:=False ; FP4.Enabled:=False ;
FP5.Enabled:=True ; FP6.Enabled:=True  ; FP7.Enabled:=True  ; FP8.Enabled:=True ;
If FPeriodicite.Values.Count>0 Then FPeriodicite.Value:=FPeriodicite.Values[0] ;
END ;


procedure TFQRBalAge.FormShow(Sender: TObject);
begin
HelpContext:=7547000 ;
{Standards.HelpContext:=7547010 ;
Avances.HelpContext:=7547020 ;
Mise.HelpContext:=7547030 ;
Option.HelpContext:=7547040 ;}

FSens.ItemIndex:=2 ;
E_MODEPAIE.Text:=Traduirememoire('<<Tous>>') ;
FDateCompta1.Text:=DateToStr(V_PGI.DateEntree) ;
  inherited;
FloadColl:=FALSE ;
If FFiltres.Text='' then FDateCompta1.Text:=DateToStr(V_PGI.DateEntree) ;
TabSup.TabVisible:=(FChoixEcart.ItemIndex=1) ;
FExercice.Visible:=False ;
FSelectCpte.Visible:=False ; TSelectCpte.Visible:=False ;
FOnlyCptAssocie.Enabled:=False ; FLigneRupt.Enabled:=False ;
FChoixMontant.ItemIndex:=0 ;
end;

procedure TFQRBalAge.FChoixEcartClick(Sender: TObject);
begin
  inherited;
FEcart.Enabled:=(FChoixEcart.ItemIndex=0) ;
TabSup.TabVisible:=(FChoixEcart.ItemIndex=1) ;
If (FChoixEcart.ItemIndex=1) then SaisieDates ;
end;

procedure TFQRBalAge.FDateCompta1Exit(Sender: TObject);
begin
  inherited;
If FChoixEcart.ItemIndex=1 then SaisieDates ;
end;

procedure TFQRBalAge.FP1Exit(Sender: TObject);
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
   MsgRien.Execute(7,'',' '+TMaskEdit(Sender).Text) ;
   TMaskEdit(Sender).Text:=DateToStr(VH^.ExoV8.Fin);
   TMaskEdit(Sender).SetFocus ;
   END ;
end;

procedure TFQRBalAge.PeriodesEnter(Sender: TObject);
begin
  inherited;
FPeriodicite.SetFocus ;
end;

procedure TFQRBalAge.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
Fillchar(TotAux,SizeOf(TotAux),#0) ;
Fillchar(TotMdp,SizeOf(TotMdp),#0) ;
PrintBand:=False ;
end;

procedure TFQRBalAge.BSubDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var MReport : TabTRep ;
    OkCalcul : Boolean ;
begin
  inherited;
Fillchar(MReport,SizeOf(MReport),#0) ;
if (Not E_MODEPAIE.Tous) And (FTypePar.ItemIndex=0) then PrintBand:=MDPRetenuTiers(LMdp, QR2E_MODEPAIE.AsString) ;
OkCalcul:=True ;
Case CritEdt.Rupture of
  rLibres    : if CritEdt.GLV.OnlyCptAssocie then OkCalcul:=DansRuptLibre(Q,fbAux,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
  rRuptures  : if CritEdt.GlV.OnlyCptAssocie then OkCalcul:=DansRupt(LRupt,Qr1T_AUXILIAIRE.AsString) ;
  rCorresp   : if CritEdt.GlV.OnlyCptAssocie then
                  if CritEDT.GLV.PlansCorresp=1 then OkCalcul:=DansRuptCorresp(LRupt,Qr1T_CORRESP1.AsString) Else
                  if CritEDT.GLV.PlansCorresp=2 then OkCalcul:=DansRuptCorresp(LRupt,Qr1T_CORRESP2.AsString) ;
 //                if CritEDT.GlV.PlansCorresp=1 then OkCalcul:=(Qr1T_CORRESP1.AsString<>'') Else
 //                if CritEDT.GlV.PlansCorresp=2 then OkCalcul:=(Qr1T_CORRESP2.AsString<>'') ;
  End;
Affiche:=OkCalcul ;
IF OkCalcul then
   BEGIN
   if QR2E_DateCOMPTABLE.AsDateTime<TabDate[4]
      then Calculs(MReport[1], MReport[7], 5, PrintBand, True) ;
   if (QR2E_DateCOMPTABLE.AsDateTime>=TabDate[4]) And (QR2E_DateCOMPTABLE.AsDateTime<TabDate[3])
      then Calculs(MReport[2], MReport[7], 4, PrintBand, True) ;
   if (QR2E_DateCOMPTABLE.AsDateTime>=TabDate[3]) And (QR2E_DateCOMPTABLE.AsDateTime<TabDate[2])
      then Calculs(MReport[3], MReport[7], 3, PrintBand, True) ;
   if (QR2E_DateCOMPTABLE.AsDateTime>=TabDate[2]) And (QR2E_DateCOMPTABLE.AsDateTime<TabDate[1])
      then Calculs(MReport[4], MReport[7], 2, PrintBand, True) ;
   if (QR2E_DateCOMPTABLE.AsDateTime>=TabDate[1]) And (QR2E_DateCOMPTABLE.AsDateTime<CritEdt.Date1)
      then Calculs(MReport[5] ,MReport[7], 1, PrintBand, True) ;
   if QR2E_DateCOMPTABLE.AsDateTime>=CritEdt.Date1
      then case CritEdt.GlV.TypePar of
             0 : Calculs(MReport[6], MReport[7], 0, PrintBand, True) ;
             1 : Calculs(MReport[6], MReport[7], 0, PrintBand, False) ;
            end ;
   AddReportBAL([1], CritEdt.GlV.FormatPrint.Report, MReport, CritEdt.Decimale) ;
   END ;
PrintBand:=False ;
end;

procedure TFQRBalAge.BFCompteAuxBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=Affiche ;
if CritEdt.GlV.RuptOnly=Sur then PrintBand:=False ;
If PrintBand then
   BEGIN
   Quoi:=QuoiCpt(0) ;
   AuxCum5.Caption:=PrintSoldeFormate(TotAux[5].TotDebit, TotAux[5].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
   AuxCum4.Caption:=PrintSoldeFormate(TotAux[4].TotDebit, TotAux[4].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
   AuxCum3.Caption:=PrintSoldeFormate(TotAux[3].TotDebit, TotAux[3].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
   AuxCum2.Caption:=PrintSoldeFormate(TotAux[2].TotDebit, TotAux[2].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
   AuxCum1.Caption:=PrintSoldeFormate(TotAux[1].TotDebit, TotAux[1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
   AuxCum0.Caption:=PrintSoldeFormate(TotAux[0].TotDebit, TotAux[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
   AuxTotal.Caption:=PrintSoldeFormate(TotAux[6].TotDebit, TotAux[6].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
   END ;
end;

procedure TFQRBalAge.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TotCum0.Caption:=PrintSoldeFormate(TotEdt[0].TotDebit, TotEdt[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
TotCum1.Caption:=PrintSoldeFormate(TotEdt[1].TotDebit, TotEdt[1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
TotCum2.Caption:=PrintSoldeFormate(TotEdt[2].TotDebit, TotEdt[2].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
TotCum3.Caption:=PrintSoldeFormate(TotEdt[3].TotDebit, TotEdt[3].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
TotCum4.Caption:=PrintSoldeFormate(TotEdt[4].TotDebit, TotEdt[4].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
TotCum5.Caption:=PrintSoldeFormate(TotEdt[5].TotDebit, Totedt[5].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
TotTotal.Caption:=PrintSoldeFormate(TotEdt[6].TotDebit, TotEdt[6].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFQRBalAge.BMdpBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=(TotMdp[6].TotDebit<>0) Or (TotMdp[6].TotCredit<>0) ;
MdpSolde5.Caption:=PrintSoldeFormate(TotMdp[5].TotDebit, TotMdp[5].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
MdpSolde4.Caption:=PrintSoldeFormate(TotMdp[4].TotDebit, TotMdp[4].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
MdpSolde3.Caption:=PrintSoldeFormate(TotMdp[3].TotDebit, TotMdp[3].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
MdpSolde2.Caption:=PrintSoldeFormate(TotMdp[2].TotDebit, TotMdp[2].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
MdpSolde1.Caption:=PrintSoldeFormate(TotMdp[1].TotDebit, TotMdp[1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
MdpSolde0.Caption:=PrintSoldeFormate(TotMdp[0].TotDebit, TotMdp[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
MdpTotal.Caption:= PrintSoldeFormate(TotMdp[6].TotDebit, TotMdp[6].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, CritEdt.GLV.SoldeFormate) ;
end;

procedure TFQRBalAge.FPeriodiciteChange(Sender: TObject);
Var i : Integer ;
    TabD : TTabDate8 ;
begin
  inherited;
If QRLoading then Exit ;
PeriodiciteChangeTiers(neBal,Age,FPeriodicite.ItemIndex,FP8,FP1,TRUE,TabD) ;
For i:=1 to 8 do TMaskEdit(FindComponent('FP'+InttoStr(i))).text:=DateToStr(TabD[i]) ;
end;

procedure TFQRBalAge.FTypeParClick(Sender: TObject);
begin
  inherited;
if FTypePar.ItemIndex=1 then
   BEGIN
   E_MODEPAIE.Text:=Traduirememoire('<<Tous>>') ;
   FTriPar.ItemIndex:=0 ;
   FSansRupt.Checked:=True ;
   END ;
E_MODEPAIE.Enabled:=(FTypePar.ItemIndex=0) ;
FTriPar.Enabled:=(FTypePar.ItemIndex=0) ;
FGroupChoixRupt.Enabled:=(FTypePar.ItemIndex=0) ;
end;

procedure TFQRBalAge.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
Titre1Rep.Caption:=Titre2Rep.Caption ;
ReportCol1.Caption:=ReportCol7.Caption ;
ReportCol2.Caption:=ReportCol8.Caption ;
ReportCol3.Caption:=ReportCol9.Caption ;
ReportCol4.Caption:=ReportCol10.Caption ;
ReportCol5.Caption:=ReportCol11.Caption ;
ReportCol6.Caption:=ReportCol12.Caption ;
ReportTotal1.Caption:=ReportTotal2.Caption ;
end;

procedure TFQRBalAge.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var MReport : TabTRep ;
begin
  inherited;
FillChar(Mreport,Sizeof(Mreport),#0) ;
Case QuelReportBAL(CritEdt.GlV.FormatPrint.Report,MReport) of
  1 : Titre2Rep.Caption:=MsgBox.Mess[16] ;
 end ;
ReportCol7.Caption  :=PrintSoldeFormate(MReport[1].TotDebit,MReport[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.GLV.SoldeFormate) ;
ReportCol8.Caption  :=PrintSoldeFormate(MReport[2].TotDebit,MReport[2].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.GLV.SoldeFormate) ;
ReportCol9.Caption  :=PrintSoldeFormate(MReport[3].TotDebit,MReport[3].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.GLV.SoldeFormate) ;
ReportCol10.Caption :=PrintSoldeFormate(MReport[4].TotDebit,MReport[4].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.GLV.SoldeFormate) ;
ReportCol11.Caption :=PrintSoldeFormate(MReport[5].TotDebit,MReport[5].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.GLV.SoldeFormate) ;
ReportCol12.Caption :=PrintSoldeFormate(MReport[6].TotDebit,MReport[6].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.GLV.SoldeFormate) ;
ReportTotal2.Caption:=PrintSoldeFormate(MReport[7].TotDebit,MReport[7].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.GLV.SoldeFormate) ;
end;

procedure TFQRBalAge.QEcrAfterOpen(DataSet: TDataSet);
begin
  inherited;
QR2E_MODEPAIE      :=TStringField(QEcr.FindField('E_MODEPAIE')) ;
QR2E_DATECOMPTABLE :=TDateTimeField(QEcr.FindField('E_DATECOMPTABLE')) ;
QR2COUVERTURE      :=TFloatField(QEcr.FindField('COUVERTURE')) ;
QR2DEBIT           :=TFloatField(QEcr.FindField('DEBIT')) ;
QR2CREDIT          :=TFloatField(QEcr.FindField('CREDIT')) ;
ChgMaskChamp(Qr2DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
ChgMaskChamp(Qr2CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
end;

procedure TFQRBalAge.FNatureCptChange(Sender: TObject);
begin
  inherited;
If QRLoading then Exit ; 
FColl1.clear ; FColl2.clear ; FCollJoker.clear ;
ChangeColl(FNatureCpt.Value,FColl1, FColl2) ;
end;

procedure TFQRBalAge.BNouvRechClick(Sender: TObject);
begin
  inherited;
If FCollJoker.Visible then FCollJoker.Text:='' ;

end;

procedure TFQRBalAge.FColl1Change(Sender: TObject);
Var AvecJokerS : Boolean ;
begin
  inherited;
AvecJokerS:=Joker(FColl1, FColl2, FCollJoker) ;
TFaC.Visible:=Not AvecJokerS ;
FColl.Visible:=Not AvecJokerS ;
TFCollJoker.Visible:=AvecJokerS ;
End ;

procedure TFQRBalAge.DLRuptNeedData(var MoreData: Boolean);
var TotRupt          : Array[0..13] of Double ;
    Librupt, CodRupt, Lib1, Cptrupt, StCode : String ;
    Quellerupt, i    : Integer ;
    Col              : TColor ;
    OkOk, DansTotal, AddTotEdt : Boolean ;
    LibRuptInf : Array[1..10] Of TRuptInf ;
begin
  inherited;
MoreData:=false ;
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
  rRuptures : MoreData:=PrintRupt(LRupt,Qr1T_AUXILIAIRE.AsString,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt) ;
  rCorresp  : BEGIN
              OkOk:=TRUE ;
              Case CritEDT.GlV.PlansCorresp  Of
                1 : If Qr1T_CORRESP1.AsString<>'' Then CptRupt:=Qr1T_CORRESP1.AsString+Qr1T_AUXILIAIRE.AsString
                                                   Else CptRupt:='.'+Qr1T_AUXILIAIRE.AsString ;
                2 : If Qr1T_CORRESP2.AsString<>'' Then CptRupt:=Qr1T_CORRESP2.AsString+Qr1T_AUXILIAIRE.AsString
                                                   Else CptRupt:='.'+Qr1T_AUXILIAIRE.AsString ;
                Else OkOk:=FALSE ;
                END ;
              If OkOk Then MoreData:=PrintRupt(LRupt,CptRupt,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt) Else MoreData:=FALSE ;
              END ;
  End ;
if MoreData then
   BEGIN
   TCodRupt.Width:=TitreColCpt.Width+TitreColLibelle.Width+1 ;
   TCodRupt.Caption:='' ;
   BRupt.Height:=HauteurBandeRuptIni ;
   if CritEdt.Rupture=rLibres then
      BEGIN
      insert(MsgBox.Mess[19]+' ',CodRupt,Quellerupt+2) ;
      TCodRupt.Caption:=CodRupt+' '+Lib1 ;
      AlimRuptSup(HauteurBandeRuptIni,QuelleRupt,TCodRupt.Width,BRupt,LibRuptInf,Self) ;
      END Else TCodRupt.Caption:=CodRupt+'   '+LibRupt ;
   Libre5.Caption:=PrintSoldeFormate(TotRupt[10], TotRupt[11], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.GLV.SoldeFormate) ;
   Libre4.Caption:=PrintSoldeFormate(TotRupt[8], TotRupt[9], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.GLV.SoldeFormate) ;
   Libre3.Caption:=PrintSoldeFormate(TotRupt[6], TotRupt[7], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.GLV.SoldeFormate) ;
   Libre2.Caption:=PrintSoldeFormate(TotRupt[4], TotRupt[5], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.GLV.SoldeFormate) ;
   Libre1.Caption:=PrintSoldeFormate(TotRupt[2], TotRupt[3], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.GLV.SoldeFormate) ;
   Libre0.Caption:=PrintSoldeFormate(TotRupt[0], TotRupt[1], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.GLV.SoldeFormate) ;
   LibreTotal.Caption:=PrintSoldeFormate(TotRupt[12], TotRupt[13], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole, Critedt.GLV.SoldeFormate) ;
   if CritEdt.GlV.RuptOnly=Sur then
      BEGIN
      AddTotEdt:=False ;
      if (CritEdt.Rupture=rLibres) then AddTotEdt:=False else
      if (CritEdt.Rupture=rRuptures) or (CritEdt.Rupture=rCorresp) then AddTotEdt:=DansTotal ;
      if AddTotEdt then
         BEGIN
         For i:=0 to 6 do
             BEGIN
             TotEdt[i].TotDebit:=Arrondi(TotEdt[i].TotDebit+TotRupt[i*2], CritEdt.Decimale) ;
             TotEdt[i].TotCredit:=Arrondi(TotEdt[i].TotCredit+TotRupt[i*2+1], CritEdt.Decimale) ;
             END ;
         END ;
      END ;
   END ;
end;

procedure TFQRBalAge.BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.Rupture<>rRien) ;
end;

procedure TFQRBalAge.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
If CritEDT.Rupture<>rRien then
   BEGIN
   Qr1T_AUXILIAIRE  :=TStringField(Q.FindField('T_AUXILIAIRE'));
   Qr1T_LIBELLE     :=TStringField(Q.FindField('T_LIBELLE'));
   If CritEDT.Rupture=rCorresp then
      BEGIN
      Qr1T_CORRESP1         :=TStringField(Q.FindField('T_CORRESP1'));
      Qr1T_CORRESP2         :=TStringField(Q.FindField('T_CORRESP2'));
      END ;
   END ;
end;

procedure TFQRBalAge.FSansRuptClick(Sender: TObject);
begin
  inherited;
FLigneRupt.Enabled:=Not FSansRupt.Checked ;
FLigneRupt.checked:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
FRupturesClick(Nil) ;
end;

procedure TFQRBalAge.FRupturesClick(Sender: TObject);
begin
  inherited;
If FPlansCo.Checked then
   BEGIN
   FGroupRuptures.Caption:=' '+MsgBox.Mess[23] ;
   CorrespToCombo(FPlanRuptures,fbAux) ;
   END Else
   BEGIN
   FPlanRuptures.Reload    ;
   FGroupRuptures.Caption:=' '+MsgBox.Mess[22] ;
   END ;
If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
end;

procedure TFQRBalAge.FPlanRupturesChange(Sender: TObject);
begin
  inherited;
If QRLoading Then Exit ;
if FPlansCo.Checked then CorrespToCodes(FPlanRuptures,FCodeRupt1,FCodeRupt2) Else
if FRuptures.Checked then RuptureToCodes(FPlanRuptures,FCodeRupt1,FCodeRupt2,FbAux) ;
FCodeRupt1.ItemIndex:=0 ; FCodeRupt2.ItemIndex:=FCodeRupt2.Items.Count-1 ;
end;


end.
