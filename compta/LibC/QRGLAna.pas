unit QRGLANA;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, Ent1, HEnt1, QrRupt, CritEdt, UtilEdt, HQry,
  EdtLegal, CpteUtil, ParamDBG, HSysMenu, Menus, HTB97, HPanel, UiUtil, tCalcCum ;

procedure GLAnal ;
procedure GLAnalZoom(Crit : TCritEdt) ;

type
  TFQRGLANA = class(TFQR)
    HLabel9: THLabel;
    FNumPiece1: TMaskEdit;
    Label2: TLabel;
    FNumPiece2: TMaskEdit;
    HLabel10: THLabel;
    FRefInterne: TEdit;
    FValide: TCheckBox;
    FTypeAnal: TCheckBox;
    FSautPage: TCheckBox;
    FLigneCptPied: TCheckBox;
    TRLegende: TQRLabel;
    RLegende: TQRLabel;
    QRLabel13: TQRLabel;
    RNumPiece1: TQRLabel;
    QRLabel14: TQRLabel;
    RNumPiece2: TQRLabel;
    TRRefInterne: TQRLabel;
    RRefInterne: TQRLabel;
    QRLabel17: TQRLabel;
    RValide: TQRLabel;
    TValide: TQRLabel;
    TTYPEANALYTIQUE: TQRLabel;
    T_NUMLIGNE: TQRLabel;
    T_GENERAL: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel21: TQRLabel;
    QRLabel10: TQRLabel;
    T_DEBIT: TQRLabel;
    T_CREDIT: TQRLabel;
    QRLabel16: TQRLabel;
    TITRE1REP: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    REPORT1SOLDE: TQRLabel;
    S_SECTION: TQRLabel;
    QRLabel25: TQRLabel;
    TCumul: TQRLabel;
    ANvDEBIT: TQRLabel;
    CumDEBIT: TQRLabel;
    CumCREDIT: TQRLabel;
    ANvCREDIT: TQRLabel;
    ANvSolde: TQRLabel;
    CumSolde: TQRLabel;
    BMVT: TQRBand;
    Y_DATECOMPTABLE: TQRDBText;
    Y_REFINTERNE: TQRDBText;
    Y_VALIDE: TQRDBText;
    Y_LIBELLE: TQRDBText;
    Y_TYPEANALYTIQUE: TQRDBText;
    SOLDE: TQRLabel;
    Y_JOURNAL: TQRDBText;
    CREDIT: TQRDBText;
    DEBIT: TQRDBText;
    E_PIECELIGNE: TQRLabel;
    BTOTM: TQRBand;
    TTOTMENS: TQRLabel;
    TOTMENSD: TQRLabel;
    TOTMENSC: TQRLabel;
    TOTMENSS: TQRLabel;
    BSecF: TQRBand;
    S_SECTION_: TQRLabel;
    TOT1SOLDE: TQRLabel;
    TTOTQte1: TQRLabel;
    TTOTQte2: TQRLabel;
    TOTQte1: TQRDBCalc;
    TOTQte2: TQRDBCalc;
    TOT1DEBIT: TQRLabel;
    TOT1CREDIT: TQRLabel;
    BRUPT: TQRBand;
    DEBITRUPT: TQRLabel;
    CREDITRUPT: TQRLabel;
    CodR: TQRLabel;
    SOLDERUPT: TQRLabel;
    QRLabel33: TQRLabel;
    TOT2DEBIT: TQRLabel;
    TOT2CREDIT: TQRLabel;
    TOT2SOLDE: TQRLabel;
    QRBand1: TQRBand;
    Trait4: TQRLigne;
    Trait3: TQRLigne;
    Trait2: TQRLigne;
    Trait1: TQRLigne;
    Trait0: TQRLigne;
    Ligne1: TQRLigne;
    TITRE2REP: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    REPORT2SOLDE: TQRLabel;
    QAnal: TQuery;
    SAna: TDataSource;
    DLMVT: TQRDetailLink;
    DLMOIS: TQRDetailLink;
    DLrupt: TQRDetailLink;
    MsgLibel: THMsgBox;
    TitreTotDesMvts: TQRLabel;
    TotDesMvtsDebit: TQRLabel;
    TotDesMvtsCredit: TQRLabel;
    TotDesMvtsSolde: TQRLabel;
    FSoldeProg: TCheckBox;
    FTotMens: TCheckBox;
    FQtes: TCheckBox;
    FLigneRupt: TCheckBox;
    FAvecAN: TCheckBox;
    BRupTete: TQRBand;
    CodeRuptAu: TQRLabel;
    AnoRupt: TQRLabel;
    CumRupt: TQRLabel;
    AnoDebitRupt: TQRLabel;
    CumDebitRupt: TQRLabel;
    AnoCreditRupt: TQRLabel;
    CumCreditRupt: TQRLabel;
    AnoSoldeRupt: TQRLabel;
    CumSoldeRupt: TQRLabel;
    DLRuptH: TQRDetailLink;
    FSansANO: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean;      var Quoi: string);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BMVTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BMVTAfterPrint(BandPrinted: Boolean);
    procedure BTOTMBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BSecFBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean;      var Quoi: string);
    procedure DLMOISNeedData(var MoreData: Boolean);
    procedure DLruptNeedData(var MoreData: Boolean);
    procedure BRUPTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FNatureCptChange(Sender: TObject);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QAnalAfterOpen(DataSet: TDataSet);
    procedure BRUPTAfterPrint(BandPrinted: Boolean);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure FSelectCpteChange(Sender: TObject);
    procedure FExerciceChange(Sender: TObject);
    procedure FDateCompta1Change(Sender: TObject);
    procedure BSecFAfterPrint(BandPrinted: Boolean);
    procedure DLRuptHNeedData(var MoreData: Boolean);
    procedure BRupTeteBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FSansRuptClick(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private    { Déclarations privées }
    TotCptSect, TotEdt, TotDesMvts,TotAnCRupt : TabTot ;
    Affiche,AvecQtes, DansTotal,
    IsECC, OkEnteteRup                      : Boolean ;
    StReportSec                             : String ;
    Lmois, Lrupt                            : TStringList ;
    ListeCodesRupture                       : TStringList ;
    Qr1S_AXE                                : TStringField ;
    Qr1S_SECTION                            : TStringField ;
    Qr1S_LIBELLE                            : TStringField ;
    Qr1S_SAUTPAGE                           : TStringField ;
    Qr1S_SOLDEPROGRESSIF                    : TStringField ;
    Qr1S_TOTAUXMENSUELS, Qr1S_SECTIONTRIE   : TStringField ;
    Qr1S_CORRESP1, Qr1S_CORRESP2            : TStringField ;
    FDebit,FCredit                          : TFloatField ;
    FY_Section                              : TStringField ;
    FMois                                   : String ;
    FY_Valide,FY_TypeAnalytique,
    FY_QUALIFPIECE,FY_JOURNAL,FY_EXERCICE   : TStringField ;
    FY_NUMEROPIECE,FY_NUMLIGNE              : TIntegerField ;
    FY_DATECOMPTABLE                        : TDateTimeField ;
    FLoading            : Boolean ;
    Procedure GenereSQLSub ;
    Procedure RequeteSelectAnal ;
    Function OkPourRupt : Boolean ;
    Function QuoiAna(i : Integer) : String ;
    Function QuoiMvt : String ;
    Procedure GLAnaZoom(Quoi : String) ;
 public    { Déclarations publiques }
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure GenereSQL ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
  end;


implementation

{$R *.DFM}

Uses CPSection_TOM ;

procedure GLAnal ;
var QR : TFQRGLANA ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQRGLANA.Create(Application) ;
Edition.Etat:=etGlAna ;
QR.QRP.QRPrinter.OnSynZoom:=QR.GLAnaZoom ;
QR.InitType (nbSec,neGL,msSecAna,'QRGLANA','QR_GLANA',TRUE,FALSE,FALSE,Edition) ;
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

procedure GLAnalZoom(Crit : TCritEdt) ;
var QR : TFQRGLANA ;
    Edition : TEdition ;
BEGIN
QR:=TFQRGLANA.Create(Application) ;
Edition.Etat:=etGlAna ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.GLAnaZoom ;
 QR.CritEdt:=Crit ;
 QR.InitType (nbSec,neGL,msSecAna,'QRGLANA','',FALSE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

Function TFQRGLANA.QuoiAna(i : Integer) : String ;
BEGIN
Inherited ;
Case i Of
  1 : Result:=Qr1S_SECTION.AsString+' '+'#'+Qr1S_LIBELLE.AsString+'@'+'3'+Qr1S_AXE.AsString ;
  2 : Result:=Qr1S_SECTION.AsString+' '+'#'+Qr1S_LIBELLE.AsString+' '+
              TOT1SOLDE.Caption+'@'+'3'+Qr1S_AXE.AsString ;
  END ;
END ;

Function TFQRGLANA.QuoiMvt : String ;
BEGIN
Inherited ;
Result:=FY_SECTION.AsString+' '+Qr1S_LIBELLE.AsString+' '+Solde.Caption+
        '#'+FY_JOURNAL.AsString+' N° '+IntToStr(FY_NUMEROPIECE.AsInteger)+' '+DateToStr(FY_DateComptable.AsDAteTime)+'-'+
        PrintSolde(FDEBIT.AsFloat,FCredit.AsFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)+
        '@'+'7;'+FY_JOURNAL.AsString+';'+UsDateTime(FY_DATECOMPTABLE.AsDateTime)+';'+FY_NUMEROPIECE.AsString+';'+FY_EXERCICE.asString+';'+
        IntToStr(FY_NumLigne.AsInteger)+';'+Qr1S_Axe.AsString+';' ;
END ;

Procedure TFQRGLANA.GLAnaZoom(Quoi : String) ;
Var Lp,i: Integer ;
BEGIN
Inherited ;
Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
i:=StrToInt(Copy(Quoi,Lp+1,1)) ;
If (i=5) Then
   BEGIN
   Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
   If QRP.QrPrinter.FSynShiftDblClick Then i:=6 ;
   END ;
If (i=7) Then Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
ZoomEdt(i,Quoi) ;
END ;



procedure TFQRGLANA.FinirPrint;
BEGIN
Inherited ;
if CritEdt.Rupture<>rRien then VideRupt(Lrupt) ;
VideGroup(Lmois) ;
QAnal.Close ;
If GCalc<>NIL Then BEGIN GCalc.Free ; GCalc:=NIL ; END ;
If OkMajEdt Then
      MajEdition('GA'+Char(FNatureCpt.ItemIndex+1),CritEdt.Exo.Code,DateToStr(CritEdt.DateDeb),DateToStr(CritEdt.DateFin),
                 '',TotEdt[3].TotDebit,TotEdt[3].TotCredit,TotEdt[4].TotDebit,TotEdt[4].TotCredit) ;
END ;

procedure TFQRGLANA.InitDivers ;
BEGIN
Inherited ;
{ Initialisation  }
TValide.Caption:='' ; TTYPEANALYTIQUE.Caption:='' ;
{ Init du Total Général }
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
Fillchar(TotAncRupt,SizeOf(TotAncRupt),#0) ;
{ Labels sur les bandes }
BSecF.Frame.DrawTop:=CritEdt.GL.FormatPrint.PrSepCompte[2] ;
BRUPT.Frame.DrawTop:=CritEdt.GL.FormatPrint.PrSepCompte[3] ;
BRUPT.Frame.DrawBottom:=CritEdt.GL.FormatPrint.PrSepCompte[3] ;
If CritEdt.GA Then CritEdt.SQLGA:=GetParamListe(NomListeGA,Self) ;
OkEnteteRup:=(CritEdt.Rupture=rCorresp) ;
END ;

procedure TFQRGLANA.GenereSQL ;
//Var St : string ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
GenereSQLBase ;
(*
Q.Close ;
Q.Sql.Clear ;
St:='Select S_AXE, S_SECTION, S_LIBELLE, S_SOLDEPROGRESSIF, S_SAUTPAGE, S_TOTAUXMENSUELS,'
          + 'S_TOTDEBANO, S_TOTCREANO, S_TOTDEBE, S_TOTCREE, S_SECTIONTRIE ' ;
Q.SQL.Add(St) ;
{ Tables explorées par la SQL }
St:=' From SECTION Q Where ' ;
Q.SQL.Add(St) ;
{ Construction de la clause Where de la SQL }
St:=SQLModeSelect ;
If St<>'' Then Q.SQL.Add(St) ;
St:='' ;
if CritEdt.Joker then St:='AND S_SECTION like "'+TraduitJoker(CritEdt.Cpt1)+'" ' Else
   BEGIN
   if CritEdt.Cpt1<>'' then St:='AND S_SECTION>="'+CritEdt.Cpt1+'" ' ;
   if CritEdt.Cpt2<>'' then St:=St+'AND S_SECTION<="'+CritEdt.Cpt2+'" ' ;
   END ;
If St<>'' Then Q.SQL.Add(St) ;
st:='' ;
St:='AND S_AXE="'+CritEdt.GL.Axe+'" ' ;
If St<>'' Then Q.SQL.Add(St) ;
st:='' ;
if CritEdt.GL.Sauf<>'' then St:=CritEdt.GL.SqlSauf ;
if st<>'' Then Q.SQL.Add(St) ;
{ Construction de la clause Order By de la SQL }
if FPlanRupt.Enabled then St:=' Order By S_AXE, S_SECTIONTRIE '
                     Else St:=' Order By S_AXE, S_SECTION ' ;
Q.SQL.Add(St) ;
ChangeSQL(Q) ; Q.Prepare ;
Q.Open ;
*)
GenereSQLSub ;
END;

procedure TFQRGLANA.RenseigneCritere ;
{ Récupération des champs du multicritère dans l'entête d'état }
BEGIN
Inherited ;
If EstSerie(S5) Then RLegende.Caption:=MsgRien.Mess[9] ;
RNumPiece1.Caption:=IntToStr(CritEdt.GL.NumPiece1)   ; RNumPiece2.Caption:=IntToStr(CritEdt.GL.NumPiece2) ;
RRefInterne.Caption:=FRefInterne.Text ;
CaseACocher(FValide, RValide) ;
(*
Case FChoixRupt.ItemIndex of
  0 : BEGIN ROui.Caption:='þ' ; RNon.Caption:='o' ; END ;
  1 : BEGIN ROui.Caption:='o' ; RNon.Caption:='þ' ; END ;
end ;
If FPlanRupt.Enabled=True then RPlanRupt.Caption:=FPlanRupt.Text Else RPlanRupt.Caption:='' ;
CaseACocher(FSoldeProg, RSoldeProg) ;
CaseACocher(FTOTMENS, RTOTMENS) ;
CaseACocher(FSautPage, RSautPage) ;
CaseACocher(FQtes, RQtes) ;
*)
DateCumulAuGL(TCumul,CritEdt,MsgLibel.Mess[0]) ;
//TCumul.Caption:=  MsgLibel.Mess[0]+' '+DateToStr(StrToDateTime(FDateCompta1.Text)-1) ;
END ;

procedure TFQRGLANA.ChoixEdition ;
Var Top1, Top2 : Byte ;
BEGIN
Inherited ;
{ Affiche les Qtés O/N }
AvecQtes:=(FQtes.Checked=True)    ;
TTOTQte1.visible:=AvecQtes        ; TOTQte1.visible:=AvecQtes ;
TTOTQte2.visible:=AvecQtes        ; TOTQte2.visible:=AvecQtes ;
// Rony 9/07/97
//If CritEdt.GA then BMVT.Height:=Sup1.Top+Sup1.Height else BMVT.Height:=Sup1.Top ;
If AvecQtes then
   BEGIN
   Top1:=TTOTQte1.Top+TTOTQte1.Height+1 ; TitreTotDesMvts.top:=Top1 ;
   TotDesMvtsDebit.Top:=Top1 ; TotDesMvtsCredit.Top:=Top1 ; TotDesMvtsSolde.Top:=Top1 ;
   Top2:=TitreTotDesMvts.Top+TitreTotDesMvts.Height+1 ;
   S_SECTION_.Top:=Top2 ; TOT1DEBIT.Top:=Top2 ;
   TOT1CREDIT.Top:=Top2 ; TOT1SOLDE.Top:=Top2 ;
   BSecF.Height:=(Top2+TitreTotDesMvts.Height)+2 ;
   END Else
   BEGIN
   Top1:=0 ; TitreTotDesMvts.top:=Top1 ;
   TotDesMvtsDebit.Top:=Top1 ; TotDesMvtsCredit.Top:=Top1 ; TotDesMvtsSolde.Top:=Top1 ;
   Top2:=(Top1+TitreTotDesMvts.Height)+1 ;
   S_SECTION_.Top:=Top2 ; TOT1DEBIT.Top:=Top2 ;
   TOT1CREDIT.Top:=Top2 ; TOT1SOLDE.Top:=Top2 ;
   BSecF.Height:=(Top2+TitreTotDesMvts.Height)+2 ;
   END ;
{ Masque de saisie des Champs TQRDBText D et C sur la bande BMVT }
ChgMaskChamp(FDEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
ChgMaskChamp(FCREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;

Case CritEdt.Rupture of
  rLibres   : BEGIN
              DLrupt.PrintBefore:=FALSE ;
              ChargeGroup(LRupt,['S00','S01','S02','S03','S04','S05','S06','S07','S08','S09']) ;
              END ;
  rCorresp  : BEGIN
              DLrupt.PrintBefore:=TRUE ;
              ChargeRuptCorresp(LRupt, CritEdt.GL.PlanRupt, '', '', False) ;
              END ;
  End ;
ChargeGroup(Lmois,['MOIS']) ; { mettre autant de string que de groupes }
END ;

procedure TFQRGLANA.RecupCritEdt ;
Var NonLibres : Boolean ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  GL.CodeRupt1:='' ; GL.CodeRupt2:='' ;
  if FValide.State=cbGrayed then Valide:='g' ;
  if FValide.State=cbChecked then Valide:='X' ;
  if FValide.State=cbUnChecked then Valide:='-' ;
  if FTypeAnal.State=cbGrayed then GL.AnalPur:='g' ;
  if FTypeAnal.State=cbChecked then GL.AnalPur:='X' ;
  if FTypeAnal.State=cbUnChecked then GL.AnalPur:='-' ;
  if FSautPage.State=cbGrayed then SautPage:=0 ;
  if FSautPage.State=cbChecked then SautPage:=1 ;
  if FSautPage.State=cbUnChecked then SautPage:=2 ;
  if FTotMens.State=cbGrayed then TotMens:=0 ;
  if FTotMens.State=cbChecked then TotMens:=1 ;
  if FTotMens.State=cbUnChecked then TotMens:=2 ;
  if FSoldeProg.State=cbGrayed then SoldeProg:=0 ;
  if FSoldeProg.State=cbChecked then SoldeProg:=1 ;
  if FSoldeProg.State=cbUnChecked then SoldeProg:=2 ;
  If FNumPiece1.Text<>'' then GL.NumPiece1:=StrToInt(FNumPiece1.Text) else GL.NumPiece1:=0 ;
  If FNumPiece2.Text<>'' then GL.NumPiece2:=StrToInt(FNumPiece2.Text) else GL.NumPiece2:=999999999 ;

  RefInterne:=FRefInterne.text ;
  GL.Axe:=FNatureCpt.Value ;
  NonLibres:=((Rupture=rRuptures) or (Rupture=rCorresp)) ;
  if NonLibres then CritEdt.GL.PlanRupt:=FPlanRuptures.Value ;
  GL.RuptOnly:=QuelleTypeRupt(1,FSAnsRupt.Checked,FAvecRupt.Checked,FALSE) ;
  GL.QuelAN:=FQuelAN(FAvecAN.Checked) ;
  GL.SansAnoANA:=FSansANO.Checked ;
  If (CritEdt.Rupture=rCorresp) Then GL.PlansCorresp:=FPlanRuptures.ItemIndex+1 ;
  GL.OnlyCptAssocie:=(Rupture<>rRien) and FOnlyCptAssocie.Checked ;
  With GL.FormatPrint Do
       BEGIN
       PrSepCompte[2]:=FLigneCptPied.Checked ;
       PrSepCompte[3]:=FLigneRupt.Checked ;
       END ;
  END ;
END ;

Function TFQRGLANA.CritOK : Boolean ;
BEGIN
Result:=Inherited CritOK and OkPourRupt ;
If Result Then
   BEGIN
   Gcalc:=TGCalculCum.create(Un,AxeToFb(CritEdt.GL.Axe),AxeToFb(CritEdt.GL.Axe),QuelTypeEcr,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE) ;
   GCalc.initCalcul('','','',CritEdt.GL.Axe,CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                    CritEdt.Date1,CritEdt.GL.Date11,TRUE) ;
   END ;
END ;

Procedure TFQRGLANA.RequeteSelectAnal ;
BEGIN
  inherited;
QAnal.Close ;
QAnal.SQL.Clear ;
QAnal.SQL.Add(' Select Y_AXE, Y_SECTION, Y_TYPEANALYTIQUE, ') ;
QAnal.SQL.Add(' Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_REFINTERNE, ') ;
QAnal.SQL.Add(' Y_ETABLISSEMENT, Y_VALIDE, Y_QTE1, Y_QTE2, Y_LIBELLE, ') ;
QAnal.SQL.Add(' Y_POURCENTQTE1, Y_POURCENTQTE2, Y_QUALIFQTE1, Y_QUALIFQTE2, ') ;
QAnal.SQL.Add(' Y_JOURNAL, Y_NATUREPIECE, Y_QUALIFPIECE, Y_EXERCICE, ') ;
If CritEdt.SQLGA<>'' Then QAnal.SQL.Add(CritEdt.SQLGA) ;
Case Fmontant.ItemIndex of
  0 : BEGIN QAnal.SQL.Add('Y_DEBIT DEBIT, Y_CREDIT CREDIT') ; END ;
  1 : BEGIN QAnal.SQL.Add('Y_DEBITDEV DEBIT, Y_CREDITDEV CREDIT') ; END ;
  end ;
QAnal.SQL.Add(' From ANALYTIQ ') ;
END ;

Procedure TFQRGLANA.GenereSQLSub ;

BEGIN
  inherited;
{ Construction de la clause Select de la SQL }
RequeteSelectAnal ;
{ Construction de la clause Where de la SQL }
QAnal.SQL.Add(' Where ') ;
QAnal.SQL.Add(' Y_AXE=:S_AXE ') ;
QAnal.SQL.Add(' And Y_SECTION=:S_SECTION ') ;
if FExercice.ItemIndex>0 then QAnal.SQL.Add(' And Y_EXERCICE="'+CritEdt.Exo.Code+'" ') ;
QAnal.SQL.Add(TraduitNatureEcr(CritEdt.QualifPiece, 'Y_QUALIFPIECE', true, CritEdt.ModeRevision) ) ;
QAnal.SQL.Add(' And Y_DATECOMPTABLE>="'+usdatetime(CritEdt.GL.Date21)+'" ');
QAnal.SQL.Add(' And Y_DATECOMPTABLE<="'+usdatetime(CritEdt.Date2)+'" ');
QAnal.SQL.Add(' And Y_NUMEROPIECE>='+IntToStr(CritEdt.GL.NumPiece1)+' and Y_NUMEROPIECE<='+IntTostr(CritEdt.GL.NumPiece2)+' ') ;
if CritEdt.RefInterne<>'' then QAnal.SQL.Add('and UPPER(Y_REFINTERNE) like "'+TraduitJoker(CritEdt.RefInterne)+'" ' );
if CritEdt.Etab<>'' then QAnal.SQL.Add(' And Y_ETABLISSEMENT="'+CritEdt.Etab+'" ') ;
if CritEdt.Valide<>'g' then QAnal.SQL.Add(' And Y_VALIDE="'+CritEdt.Valide+'" ') ;
if CritEdt.GL.AnalPur<>'g' then QAnal.SQL.Add(' And Y_TYPEANALYTIQUE="'+CritEdt.GL.AnalPur+'" ') ;
if CritEdt.DeviseSelect<>'' then QAnal.SQL.Add(' And Y_DEVISE="'+CritEdt.DeviseSelect+'" ') ;
QAnal.SQL.Add(' And Y_ECRANOUVEAU="N" and Y_QUALIFPIECE<>"C" ') ;
(*
if CritEdt.Cptlibre1<>'' then QAnal.SQL.Add(' And Y_TABLE0>="'+CritEdt.Cptlibre1+'" ') ;
if CritEdt.Cptlibre11<>'' then QAnal.SQL.Add(' And Y_TABLE0<="'+CritEdt.Cptlibre11+'" ') ;
if CritEdt.Cptlibre2<>'' then QAnal.SQL.Add(' And Y_TABLE1>="'+CritEdt.Cptlibre2+'" ') ;
if CritEdt.Cptlibre21<>'' then QAnal.SQL.Add(' And Y_TABLE1<="'+CritEdt.Cptlibre21+'" ') ;
if CritEdt.Cptlibre3<>'' then QAnal.SQL.Add(' And Y_TABLE2>="'+CritEdt.Cptlibre3+'" ') ;
if CritEdt.Cptlibre31<>'' then QAnal.SQL.Add(' And Y_TABLE2<="'+CritEdt.Cptlibre31+'" ') ;
if CritEdt.Cptlibre4<>'' then QAnal.SQL.Add(' And Y_TABLE3>="'+CritEdt.Cptlibre4+'" ') ;
if CritEdt.Cptlibre41<>'' then QAnal.SQL.Add(' And Y_TABLE3<="'+CritEdt.Cptlibre41+'" ') ;
*)

If CritEDT.SQLPlus<>'' Then QAnal.SQL.Add(CritEDT.SQLPlus) ;
{ Construction de la clause Order By de la SQL }
QAnal.SQL.Add(' Order By Y_SECTION, Y_AXE, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE ') ;
ChangeSQL(QAnal) ; QAnal.Open ;
END ;

procedure TFQRGLANA.FormShow(Sender: TObject);
begin
HelpContext:=7421000 ;
//Standards.HelpContext:=7421010 ;
//Avances.HelpContext:=7421020 ;
//Mise.HelpContext:=7421030 ;
//Option.HelpContext:=7421040 ;
//TabRuptures.HelpContext:=7421050 ;
  inherited;
{$IFDEF CCS3}
FNatureCpt.Visible:=FALSE ; HLabel2.Visible:=FALSE ;
{$ENDIF}
If EstSerie(S3) Then
  BEGIN
  FSansAno.Visible:=FALSE ; FQtes.Checked:=FALSE ; FQtes.Visible:=FALSE ;
  END ;
Floading:=FALSE ;
TabSup.TabVisible:=False; FLigneRupt.Enabled:=False ;
If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
FSurRupt.Visible:=False ;
FCodeRupt1.Visible:=False  ; FCodeRupt2.Visible:=False ;
TFCodeRupt1.Visible:=False ; TFCodeRupt2.Visible:=False ;
FOnlyCptAssocie.Enabled:=False ;
end;

procedure TFQRGLANA.FormCreate(Sender: TObject);
begin
  inherited;
ListeCodesRupture:=TStringList.Create ;
end;

procedure TFQRGLANA.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
ListeCodesRupture.Free ;
end;

procedure TFQRGLANA.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TITRE1REP.Caption:=TITRE2REP.Caption ;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
Report1Solde.Caption:=Report2Solde.Caption ;

end;

procedure TFQRGLANA.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var D,C : Double ;
    CumulAu : TabTot ; CptRupt : String ;
begin
  inherited;
Fillchar(TotDesMvts,SizeOf(TotDesMvts),#0) ;
Fillchar(TotCptSect,SizeOf(TotCptSect),#0) ;
Fillchar(CumulAu,SizeOf(CumulAu),#0) ;
S_SECTION.Caption:=MsgLibel.Mess[8]+' '+Qr1S_SECTION.AsString+' '+Qr1S_LIBELLE.AsString ;
Quoi:=Qr1S_SECTION.AsString+' '+Qr1S_LIBELLE.AsString ;
StReportSec:=Qr1S_SECTION.AsString ;
InitReport([2],CritEdt.GL.FormatPrint.Report) ;
if PrintBand then
   BEGIN
   Case CritEdt.Rupture of
     rLibres    : if CritEdt.GL.OnlyCptAssocie then PrintBand:=DansRuptLibre(Q,AxeToFb(CritEdt.Gl.Axe),CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
     rRuptures  : if CritEdt.GL.OnlyCptAssocie then PrintBand:=(Qr1S_SECTIONTRIE.AsString<>'') ;
     rCorresp   : if CritEdt.GL.OnlyCptAssocie then
                    if CritEDT.GL.PlansCorresp=1 then PrintBand:=(Qr1S_CORRESP1.AsString<>'') Else
                    if CritEDT.GL.PlansCorresp=2 then PrintBand:=(Qr1S_CORRESP2.AsString<>'') ;
     End;
   END ;


If Not PrintBand then
   BEGIN
   Affiche:=False ;
   BDetail.forceNewPage:=False ; D:=0 ; C:=0 ;
   END else
   BEGIN
   Affiche:=True ;
   (*
   ExecuteTotCpt(QBal,QS_SECTION.AsString, CritEdt.Date1,CritEdt.GL.Date11,
                 CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,CumulAu,FALSE) ;
   *)
   If (CritEdt.GL.QuelAN<>SansAN) Then
      BEGIN
      GCAlc.ReInitCalcul(Qr1S_SECTION.AsString,'',0,0) ;
      GCalc.Calcul ; CumulAu:=GCalc.ExecCalc.TotCpt ;
      END ;
   If (CritEdt.GL.QuelAN=SansAN) Or CritEdt.GL.SansAnoAna Then BEGIN CumulAu[0].TotDebit:=0 ; CumulAu[0].TotCredit:=0 ; END ;
   If CritEdt.Date1=CritEdt.GL.Date11 Then Fillchar(CumulAu[1],SizeOf(CumulAu[1]),#0) ;
   CumulVersSolde(CumulAu[0]) ;
   D:=CumulAu[0].TotDebit+CumulAu[1].TotDebit ;
   C:=CumulAu[0].TotCredit+CumulAu[1].TotCredit ;
   Case CritEdt.SautPage of             { Saut de Page O/N Sinon, d'aprés l'info sur le compte}
       0 : BDetail.forceNewPage:=Qr1S_SAUTPAGE.AsString='X' ;
       1 : BDetail.forceNewPage:=True  ;
       2 : BDetail.forceNewPage:=False ;
      end ;
   Case CritEdt.SoldeProg of             { SI Solde Progressif, donc Init du Solde Progressif O/N Sinon, d'aprés l'info sur le compte }
       1 : Progressif(False,0,0)  ;
       0 : If Qr1S_SOLDEPROGRESSIF.AsString='X' then Progressif(False,0,0) ; { Init du Solde Progressif car Solde Prog demandé par le compte }
      end ;
   TotCptSect[1].TotDebit:= Arrondi(TotCptSect[1].TotDebit+D,CritEdt.Decimale) ;  { Anv + cumul }
   TotCptSect[1].TotCredit:=Arrondi(TotCptSect[1].TotCredit+C,CritEdt.Decimale) ;
   TotAnCRupt[1].TotDebit:= Arrondi(TotAnCRupt[1].TotDebit+D,CritEDT.Decimale) ;
   TotAnCRupt[1].TotCredit:=Arrondi(TotAnCRupt[1].TotCredit+C,CritEDT.Decimale) ;
   TotEdt[1].TotDebit:=     Arrondi(TotEdt[1].TotDebit+D,CritEdt.Decimale) ;
   TotEdt[1].TotCredit:=    Arrondi(TotEdt[1].TotCredit+C,CritEdt.Decimale) ;
   END ;

   If (CritEdt.GL.QuelAN<>SansAN) Then
      BEGIN
      If CritEdt.GL.SansAnoAna Then
        BEGIN
        AnvDebit.Caption:=MsgLibel.Mess[7] ; AnvCredit.Caption:='' ; AnvSolde.Caption:='' ;
        CumDebit.Caption:=MsgLibel.Mess[7] ; CumCredit.Caption:='' ;  CumSolde.Caption:='' ;
        END Else
        BEGIN
        AnvDebit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,CumulAu[0].TotDebit, CritEdt.AfficheSymbole ) ;
        AnvCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,CumulAu[0].TotCredit, CritEdt.AfficheSymbole) ;
        AnvSolde.Caption:=PrintSolde(CumulAu[0].TotDebit,CumulAu[0].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
        END ;
      CumDebit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,D, CritEdt.AfficheSymbole) ;
      CumCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,C, CritEdt.AfficheSymbole) ;
      CumSolde.Caption:=PrintSolde(D,C,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
      END Else
      BEGIN
      AnvDebit.Caption:=MsgLibel.Mess[7] ; AnvCredit.Caption:='' ; AnvSolde.Caption:='' ;
      CumDebit.Caption:=MsgLibel.Mess[7] ; CumCredit.Caption:='' ;  CumSolde.Caption:='' ;
      END ;

   Case CritEdt.SoldeProg of
     1 :  Progressif(True,D,C) ;
     0 :  if Qr1S_SOLDEPROGRESSIF.AsString='X' then Progressif(True,D,C) ;
     end ;
   AddReport([1,2],CritEdt.GL.FormatPrint.Report,D,C,CritEdt.Decimale) ;
   Quoi:=QuoiAna(1) ;

   Case CritEdt.Rupture of
     rLibres   : AddGroupLibre(LRupt,Q,AxeToFb(CritEdt.Gl.Axe),CritEdt.LibreTrie,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
     rRuptures : AddRupt(Lrupt,Qr1S_SECTIONTRIE.AsString,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
     rCorresp  : BEGIN
                 Case CritEDT.GL.PlansCorresp Of
                   1 : If Qr1S_CORRESP1.AsString<>'' Then CptRupt:=Qr1S_CORRESP1.AsString+FY_Section.AsString
                                                     Else CptRupt:='.'+FY_Section.AsString ;
                   2 : If Qr1S_CORRESP2.AsString<>'' Then CptRupt:=Qr1S_CORRESP2.AsString+FY_Section.AsString
                                                     Else CptRupt:='.'+FY_Section.AsString ;
                   End ;
                 AddRuptCorres(Lrupt,CptRupt,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
                 END ;
     End ;
end;

procedure TFQRGLANA.BMVTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var i : Integer ; CptRupt : String ;
begin
  inherited;
PrintBand:=Affiche and (Not FY_SECTION.IsNull) ;
//PrintBand:=Not FY_SECTION.IsNull ;
If IsECC then
   BEGIN
   DEBIT.Caption:='0' ;
   CREDIT.Caption:='0' ;
   END ;
If Not PrintBand then Exit ;
If CritEdt.GA Then For i:=1 To 7 Do AlimListe(QAnal,i,Self) ;
{ Totalisation des Sections par incrément }
TotCptSect[1].TotDebit:= Arrondi(TotCptSect[1].TotDebit + FDEBIT.AsFloat,CritEdt.Decimale) ;
TotCptSect[1].TotCredit:=Arrondi(TotCptSect[1].TotCredit + FCREDIT.AsFloat,CritEdt.Decimale) ;
{ Total Général par incrément (Fin etat) }
TotEdt[1].TotDebit:=     Arrondi(TotEdt[1].TotDebit + FDEBIT.AsFloat,CritEdt.Decimale) ;
TotEdt[1].TotCredit:=    Arrondi(TotEdt[1].TotCredit + FCREDIT.AsFloat,CritEdt.Decimale) ;
{ Total des mvts}
TotDesMvts[1].TotDebit :=Arrondi(TotDesMvts[1].TotDebit + FDEBIT.AsFloat,CritEdt.Decimale) ;
TotDesMvts[1].TotCredit:=Arrondi(TotDesMvts[1].TotCredit+FCREDIT.AsFloat,CritEdt.Decimale) ;
FMois:=FormatDateTime('mmmm yyyy',FY_DATECOMPTABLE.AsDateTime) ;
AddGroup(Lmois,[FY_DateComptable],[FDEBIT.AsFloat,FCREDIT.AsFloat]) ;

Y_VALIDE.Caption:=ValiQuali(FY_VALIDE.AsString,FY_QUALIFPIECE.AsString) ;

If FY_TYPEANALYTIQUE.AsString='X' then Y_TYPEANALYTIQUE.Caption:='A' Else Y_TYPEANALYTIQUE.Caption:='' ;
E_PIECELIGNE.Caption:=IntToStr(FY_NUMEROPIECE.AsInteger)+'/'+IntToStr(FY_NUMLIGNE.AsInteger) ;
Case CritEdt.SoldeProg of             { Affectation Du Calcul du Solde Progressif O/N Sinon, d'aprés l'info sur le compte }
    0 : If Qr1S_SOLDEPROGRESSIF.AsString='X' then SOLDE.Caption:=PrintSolde(FDebit.AsFloat+ProgressDebit, FCredit.AsFloat+ProgressCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole)
                                             Else SOLDE.Caption:=PrintSolde(FDebit.AsFloat, FCredit.AsFloat, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
    1 : SOLDE.Caption:=PrintSolde(FDebit.AsFloat+ProgressDebit, FCredit.AsFloat+ProgressCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
    2 : SOLDE.Caption:=PrintSolde(FDebit.AsFloat, FCredit.AsFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
   end ;
AddReport([1,2],CritEdt.GL.FormatPrint.Report,FDebit.AsFloat,FCredit.AsFloat,CritEdt.Decimale) ;

Case CritEdt.Rupture of
  rLibres   : AddGroupLibre(LRupt,Q,AxeToFb(CritEdt.Gl.Axe),CritEdt.LibreTrie,[1,FDEBIT.AsFloat,FCREDIT.AsFloat,0,0,0,0]) ;
  rRuptures : AddRupt(Lrupt,Qr1S_SECTIONTRIE.AsString,[1,FDEBIT.AsFloat,FCREDIT.AsFloat,0,0,0,0]) ;
  rCorresp  : BEGIN
              Case CritEDT.GL.PlansCorresp Of
                1 : If Qr1S_CORRESP1.AsString<>'' Then CptRupt:=Qr1S_CORRESP1.AsString+FY_Section.AsString
                                                  Else CptRupt:='.'+FY_Section.AsString ;
                2 : If Qr1S_CORRESP2.AsString<>'' Then CptRupt:=Qr1S_CORRESP2.AsString+FY_Section.AsString
                                                  Else CptRupt:='.'+FY_Section.AsString ;
               End ;
              AddRuptCorres(Lrupt,CptRupt,[1,FDEBIT.AsFloat,FCREDIT.AsFloat,0,0,0,0]) ;
              END ;
  End ;
Quoi:=QuoiMvt ;
end;

procedure TFQRGLANA.BMVTAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Case CritEdt.SoldeProg of             { Calcul du Solde Progressif O/N Sinon, d'aprés l'info sur le compte }
    1 : Progressif(True,FDebit.AsFloat,FCredit.AsFloat)  ;
    0 : If Qr1S_SOLDEPROGRESSIF.AsString='X' then Progressif(True,FDebit.AsFloat,FCredit.AsFloat) ;
   end ;
end;

procedure TFQRGLANA.BTOTMBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
Case CritEdt.TotMens of             { Total Mensuel O/N Sinon, d'aprés l'info sur le compte }
    0 : PrintBand:=(Qr1S_TOTAUXMENSUELS.AsString='X') ;
    1 : PrintBand:=True  ;
    2 : PrintBand:=False  ;
   end ;
end;

procedure TFQRGLANA.BSecFBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
{ Affiche et Calcul O/N les Totaux des comptes mvt sur la période en fonction du Combo de sélection }
PrintBand:=Affiche ;
S_SECTION_.Caption:=MsgLibel.Mess[8]+' '+Qr1S_SECTION.AsString+'   '+Qr1S_LIBELLE.AsString ;
TOT1DEBIT.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptSect[1].TotDebit, CritEdt.AfficheSymbole ) ;
TOT1CREDIT.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptSect[1].TotCredit, CritEdt.AfficheSymbole ) ;
TOT1SOLDE.Caption:=PrintSolde(TotCptSect[1].TotDebit, TotCptSect[1].TotCredit ,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
TotDesMvtsDebit.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotDesMvts[1].TotDebit, CritEdt.AfficheSymbole) ;
TotDesMvtsCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotDesMvts[1].TotCredit, CritEdt.AfficheSymbole) ;
TotDesMvtsSolde.Caption:=PrintSolde(TotDesMvts[1].TotDebit,TotDesMvts[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
Quoi:=QuoiAna(2) ;
end;

procedure TFQRGLANA.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TOT2DEBIT.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotEdt[1].TotDebit, CritEdt.AfficheSymbole ) ;
TOT2CREDIT.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotEdt[1].TotCredit, CritEdt.AfficheSymbole ) ;
TOT2SOLDE.Caption:=PrintSolde(TotEdt[1].TotDebit, TotEdt[1].TotCredit ,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole ) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFQRGLANA.DLMOISNeedData(var MoreData: Boolean);
Var Cod,Lib : String ;
    Tot : Array[0..12] of Double ;
    QuelleRupt : Integer ;
begin
  inherited;
MoreData:=PrintGroup(Lmois,QAnal,[FY_DAteComptable],Cod,Lib,Tot,Quellerupt) ;
TTOTMENS.Caption:=MsgLibel.Mess[5]+' '+Cod ;
TOTMENSD.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, Tot[0], CritEdt.AfficheSymbole) ;
TOTMENSC.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, Tot[1], CritEdt.AfficheSymbole) ;
TOTMENSS.Caption:=PrintSolde(Tot[0],Tot[1],CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;

(* Rony 10/04/97 -- Pas de Progerssif
Case CritEdt.SoldeProg of             { Affectation Du Calcul du Solde Progressif O/N Sinon, d'aprés l'info sur le compte }
    1 : TOTMENSS.Caption:=PrintSolde(ProgressDebit,ProgressCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
    2 : TOTMENSS.Caption:=PrintSolde(Tot[0],Tot[1],CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
    0 : If Qr1S_SOLDEPROGRESSIF.AsString='X' then TOTMENSS.Caption:=PrintSolde(ProgressDebit,ProgressCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
                                                        Else TOTMENSS.Caption:=PrintSolde(Tot[0],Tot[1],CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
   end ;
*)
end;

procedure TFQRGLANA.DLruptNeedData(var MoreData: Boolean);
Var CodRupt,LibRupt, Lib1, CptRupt, Stcode : String ;
    TotRupt                  : Array[0..12] of Double ;
    SumD, SumC               : Double ;
    OkOk                     : Boolean ;
    Quellerupt               : Integer ;
    Col                      : TColor ;
    LibRuptInf : Array[1..10] Of TRuptInf ;
begin
  inherited;
MoreData:=False ;
Case CritEdt.Rupture of
  rLibres   : BEGIN
              MoreData:=PrintGroupLibre(LRupt,Q,AxeToFb(CritEdt.Gl.Axe),CritEdt.LibreTrie,CodRupt,LibRupt,Lib1,TotRupt,Quellerupt,Col,LibRuptInf) ;
              If MoreData then
                 BEGIN
                 StCode:=CodRupt ;
                 Delete(StCode,1,Quellerupt+2) ;
                 MoreData:=DansChoixCodeLibre(StCode,Q,AxeToFb(CritEdt.GL.Axe),CritEdt.LibreCodes1,CritEdt.LibreCodes2, CritEdt.LibreTrie) ;
                 BRupt.Font.Color:=Col ;
                 END ;
              END ;
  rRuptures : MoreData:=PrintRupt(Lrupt,Qr1S_SECTIONTRIE.AsString,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt) ;
  rCorresp  : BEGIN
              OkOk:=TRUE ;
              Case CritEDT.GL.PlansCorresp  Of
                1 : If Qr1S_CORRESP1.AsString<>'' Then CptRupt:=Qr1S_CORRESP1.AsString+FY_Section.AsString
                                                  Else CptRupt:='.'+FY_Section.AsString ;
                2 : If Qr1S_CORRESP2.AsString<>'' Then CptRupt:=Qr1S_CORRESP2.AsString+FY_Section.AsString
                                                  Else CptRupt:='.'+FY_Section.AsString ;
                Else OkOk:=FALSE ;
                END ;
              If OkOk Then MoreData:=PrintRupt(Lrupt,CptRupt,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt) Else MoreData:=FALSE ;
              END ;
  End ;
if MoreData then
   BEGIN
   CodR.Caption:='' ;
   Case CritEdt.Rupture of
     rLibres   : BEGIN
                 insert(MsgLibel.Mess[12]+' ',CodRupt,Quellerupt+2) ;
                 CodR.Caption:=CodRupt+' '+Lib1 ;
                 END ;
     rRuptures, rCorresp : CodR.Caption:=CodRupt+'  '+LibRupt ;
     End ;
   SumD:=Arrondi(TotAnCRupt[1].TotDebit+ TotRupt[1],CritEDT.Decimale) ;
   SumC:=Arrondi(TotAnCRupt[1].TotCredit+TotRupt[2],CritEDT.Decimale) ;

   AnoRupt.Visible:=(CritEdt.Rupture=rCorresp) ;
   AnoDebitRupt.Visible:=(CritEdt.Rupture=rCorresp) ;
   AnoCreditRupt.Visible:=(CritEdt.Rupture=rCorresp) ;
   AnoSoldeRupt.Visible:=(CritEdt.Rupture=rCorresp) ;

   CumDebitRupt.Visible:=(CritEdt.Rupture=rCorresp) ;
   CumCreditRupt.Visible:=(CritEdt.Rupture=rCorresp) ;
   CumSoldeRupt.Visible:=(CritEdt.Rupture=rCorresp) ;
   CumRupt.Visible:=(CritEdt.Rupture=rCorresp) ;
   If (CritEdt.Rupture=rCorresp) then
      BEGIN
      OkEnteteRup:=True ;
      AnoDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[3], CritEdt.AfficheSymbole) ;
      AnoCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[4], CritEdt.AfficheSymbole) ;
      AnoSoldeRupt.Caption:=PrintSolde(TotRupt[3], TotRupt[4], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

      CumDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[5], CritEdt.AfficheSymbole) ;
      CumCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[6], CritEdt.AfficheSymbole) ;
      CumSoldeRupt.Caption:=PrintSolde(TotRupt[5], TotRupt[6], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

      SumD:=Arrondi(TotRupt[5]+ TotRupt[1],CritEdt.Decimale) ;
      SumC:=Arrondi(TotRupt[6]+TotRupt[2],CritEdt.Decimale) ;
      CodR.Top:=37 ; DebitRupt.Top:=37 ; CreditRupt.Top:=37 ; SoldeRupt.Top:=37 ;
      BRupt.Height:=54 ;
      END Else
      BEGIN
      CodR.Top:=2 ; DebitRupt.Top:=2 ; CreditRupt.Top:=2 ; SoldeRupt.Top:=2;
      BRupt.Height:=20 ;
      END;

   DebitRupt.Caption:= AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumD, CritEDT.AfficheSymbole) ;
   CreditRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumC, CritEDT.AfficheSymbole) ;
   SoldeRupt.Caption:=PrintSolde(SumD, SumC, CritEDT.Decimale, CritEDT.Symbole, CritEDT.AfficheSymbole) ;
   END ;
end;

procedure TFQRGLANA.BRUPTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.Gl.RuptOnly<>Sans) ;
end;

procedure TFQRGLANA.FNatureCptChange(Sender: TObject);
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

procedure TFQRGLANA.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var D,C : Double ;
begin
  inherited;
Case QuelReport(CritEdt.GL.FormatPrint.Report,D,C) Of
  1 : BEGIN TITREREPORTB.Caption:=MsgLibel.Mess[1] ; Titre2Rep.Caption:='' ; END ;
  2 : BEGIN TITREREPORTB.Caption:=MsgLibel.Mess[2] ; Titre2Rep.Caption:=StReportSec ; END ;
  END ;
Report2Debit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, D, CritEdt.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, C, CritEdt.AfficheSymbole ) ;
Report2Solde.Caption:=PrintSolde(D,C,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
end;

procedure TFQRGLANA.QAnalAfterOpen(DataSet: TDataSet);
begin
  inherited;
FDebit:=TFloatField(Qanal.FindField('DEBIT')) ;
FCredit:=TFloatField(Qanal.FindField('CREDIT')) ;
FY_Section:=TStringField(QAnal.FindField('Y_SECTION')) ;
FY_JOURNAL:=TStringField(QAnal.FindField('Y_JOURNAL')) ;
FY_EXERCICE:=TStringField(QAnal.FindField('Y_EXERCICE')) ;
FY_Valide:=TStringField(QAnal.FindField('Y_VALIDE')) ;
FY_QUALIFPIECE:=TStringField(QAnal.FindField('Y_QUALIFPIECE')) ;
FY_TypeAnalytique:=TStringField(QAnal.FindField('Y_TYPEANALYTIQUE')) ;
FY_NUMEROPIECE:=TIntegerField(QAnal.FindField('Y_NUMEROPIECE')) ;
FY_NUMLIGNE:=TIntegerField(QAnal.FindField('Y_NUMLIGNE')) ;
FY_DATECOMPTABLE:=TDateTimeField(QAnal.FindField('Y_DATECOMPTABLE')) ; FY_DATECOMPTABLE.Tag:=1 ;
IsECC:=(FDevises.Value<>V_PGI.DevisePivot)and(FMontant.ITemIndex=1)and(QAnal.FindField('Y_NATUREPIECE').AsString='ECC') ;
ChgMaskChamp(FDEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
ChgMaskChamp(FCREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
end;

procedure TFQRGLANA.BRUPTAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Fillchar(TotAncRupt,SizeOf(TotAncRupt),#0) ;
end;

Function TFQRGLANA.OkPourRupt : Boolean ;
BEGIN
Result:=True ;
if CritEdt.Rupture=rRuptures then
   BEGIN
   Result:=False ; DLrupt.PrintBefore:=TRUE ;
   CodesRuptSec(ListeCodesRupture,CritEdt.GL.PlanRupt,CritEdt.GL.Axe) ;
   ChargeRupt(LRupt, 'RU'+Char(CritEdt.GL.Axe[2]), CritEdt.GL.PlanRupt, '', '') ;
   Case SectionRetrie(CritEdt.GL.PlanRupt,CritEdt.GL.Axe,ListeCodesRupture) of
     srOk              : Result:=True ;
     srNonStruct       : NumErreurCrit:=7 ;
     srPasEnchainement : NumErreurCrit:=8 ;
     End ;
   END ;
END ;

procedure TFQRGLANA.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr1S_AXE            :=TStringField(Q.FindField('S_AXE'));
Qr1S_SECTION        :=TStringField(Q.FindField('S_SECTION'));
Qr1S_LIBELLE        :=TStringField(Q.FindField('S_LIBELLE'));
Qr1S_SAUTPAGE       :=TStringField(Q.FindField('S_SAUTPAGE')) ;
Qr1S_SOLDEPROGRESSIF:=TStringField(Q.FindField('S_SOLDEPROGRESSIF')) ;
Qr1S_TOTAUXMENSUELS :=TStringField(Q.FindField('S_TOTAUXMENSUELS')) ;
Qr1S_SECTIONTRIE    :=TStringField(Q.FindField('S_SECTIONTRIE'));
If CritEDT.Rupture=rCorresp then
   BEGIN
   Qr1S_CORRESP1         :=TStringField(Q.FindField('S_CORRESP1'));
   Qr1S_CORRESP2         :=TStringField(Q.FindField('S_CORRESP2'));
   END ;
end;

procedure TFQRGLANA.FSelectCpteChange(Sender: TObject);
begin
  inherited;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
end;

procedure TFQRGLANA.FExerciceChange(Sender: TObject);
begin
FLoading:=TRUE ;
  inherited;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
FLoading:=FALSE ;
end;

procedure TFQRGLANA.FDateCompta1Change(Sender: TObject);
begin
  inherited;
If FLoading Then Exit ;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
end;

procedure TFQRGLANA.BSecFAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([2],CritEdt.GL.FormatPrint.Report) ;
end;

procedure TFQRGLANA.DLRuptHNeedData(var MoreData: Boolean);
begin
  inherited;
(* Gestion En-tête de rupture/corresp
MoreData:=False ;
If OkEnteteRup then
   BEGIN
   Case CritEdt.Bal.PlansCorresp of
     1 : BEGIN MoreData:=Qr1S_CORRESP1.AsString<>'' ;  END ;
     2 : BEGIN MoreData:=Qr1S_CORRESP2.AsString<>'' ;  END ;
     End ;
   END ;
OkEnteteRup:=False ;
*)
end;

procedure TFQRGLANA.BRupTeteBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
(* En-tête de rupture/corresp
PrintBand:=False ;
Case CritEdt.Bal.PlansCorresp of
  1 : BEGIN PrintBand:=Qr1S_CORRESP1.AsString<>'' ; CodeRuptAu.Caption:=Qr1S_CORRESP1.AsString+'x' ; END ;
  2 : BEGIN PrintBand:=Qr1S_CORRESP2.AsString<>'' ; CodeRuptAu.Caption:=Qr1S_CORRESP2.AsString+'x' ; END ;
  End ;
*)
end;

procedure TFQRGLANA.FSansRuptClick(Sender: TObject);
begin
  inherited;
FLigneRupt.Enabled:=Not FSansRupt.Checked ;
FLigneRupt.checked:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
FRupturesClick(Nil) ;
end;

procedure TFQRGLANA.FRupturesClick(Sender: TObject);
begin
  inherited;
If FPlansCo.Checked then FGroupRuptures.Caption:=' '+MsgLibel.Mess[10] ;
if FRuptures.Checked then FGroupRuptures.Caption:=' '+MsgLibel.Mess[9] ;

end;

end.
