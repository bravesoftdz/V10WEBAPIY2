unit QrBroBuS;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, HSysMenu, Menus, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Buttons,
  Hctrls, ExtCtrls, Mask, Hcompte, ComCtrls, UtilEDT, CritEDT, Ent1, HEnt1,
  CpteUtil, QRRupt, HTB97, HPanel, UiUtil, tCalcCum ;

procedure BrouillardBudAna(Code : char) ;

type
  TFQrBroBudSec = class(TFQR)
    Label12: TLabel;
    HLabel10: THLabel;
    FNumPiece1: TEdit;
    FNumPiece2: TEdit;
    HLabel11: THLabel;
    FRefInterne: TEdit;
    FLignePiecePied: TCheckBox;
    TNumLigne: TQRLabel;
    TCptGen: TQRLabel;
    TS_LIBELLE: TQRLabel;
    TY_REFINTERNE: TQRLabel;
    TY_LIBELLE: TQRLabel;
    TDebit: TQRLabel;
    TCredit: TQRLabel;
    BEntetePiece: TQRBand;
    TY_JOURNAL: TQRLabel;
    TY_NUMEROPIECE: TQRLabel;
    BPiedPiece: TQRBand;
    TY_NUMEROPIECE_: TQRLabel;
    TotDebitPiece: TQRLabel;
    TotCreditPiece: TQRLabel;
    BE_Valide: TQRLabel;
    BE_BUDSECT: TQRDBText;
    BS_LIBELLE: TQRDBText;
    BE_LIBELLE: TQRDBText;
    BE_REFINTERNE: TQRDBText;
    BE_DEBIT: TQRDBText;
    BE_CREDIT: TQRDBText;
    QRLabel10: TQRLabel;
    RRefInterne: TQRLabel;
    QRLabel13: TQRLabel;
    RNumPiece1: TQRLabel;
    QRLabel14: TQRLabel;
    RNumPiece2: TQRLabel;
    GPiece: TQRGroup;
    MsgBox: THMsgBox;
    QRLabel33: TQRLabel;
    TotDebitGen: TQRLabel;
    TotCreditGen: TQRLabel;
    QRBand1: TQRBand;
    Trait0: TQRLigne;
    Trait1: TQRLigne;
    Trait2: TQRLigne;
    Trait3: TQRLigne;
    Ligne1: TQRLigne;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    Y_ETABLISSEMENT: TQRLabel;
    QEcr: TQuery;
    SEcr: TDataSource;
    BComptes: TQRBand;
    DLRecap: TQRDetailLink;
    FExercice_: THValComboBox;
    TFExercice_: TLabel;
    FPeriode1: THValComboBox;
    FPeriode2: THValComboBox;
    TFNatBud: THLabel;
    FNatBud: THValComboBox;
    TFJournal: THLabel;
    FJournal: THValComboBox;
    FValide: TCheckBox;
    FAfficheAnal: TCheckBox;
    BMois: TQRBand;
    TTOTMULTI: TQRLabel;
    TDEBMULTI: TQRLabel;
    TCREMULTI: TQRLabel;
    DLMois: TQRDetailLink;
    QRLabel3: TQRLabel;
    RExercice_: TQRLabel;
    FTotSec: TCheckBox;
    LIBRECAP: TQRDBText;
    RDEBIT: TQRDBText;
    RCREDIT: TQRDBText;
    QRLabel1: TQRLabel;
    BE_NATUREBUD: TQRLabel;
    BE_BUDGENE: TQRDBText;
    BG_LIBELLE: TQRDBText;
    TBY_DEBIT: TQRDBText;
    TBY_CREDIT: TQRDBText;
    procedure QAfterOpen(DataSet: TDataSet);
    procedure BEntetePieceBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BPiedPieceBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FormShow(Sender: TObject);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BEntetePieceAfterPrint(BandPrinted: Boolean);
    procedure BPiedPieceAfterPrint(BandPrinted: Boolean);
    procedure FExerciceChange(Sender: TObject);
    procedure FExercice_Change(Sender: TObject);
    procedure FPeriode1Change(Sender: TObject);
    procedure FPeriode2Change(Sender: TObject);
    procedure FJournalChange(Sender: TObject);
    procedure FNatureCptChange(Sender: TObject);
    procedure DLMoisNeedData(var MoreData: Boolean);
    procedure FCpte1Change(Sender: TObject);
    procedure DLRecapNeedData(var MoreData: Boolean);
  private
    { Déclarations privées }
    LMois, LRecap                   : TStringList ;
    CodeNatureEcr                   : String ;
    QR1BE_EXERCICE,
    QR1BE_BUDJAL, QR1BE_NATUREBUD, QR1BE_ETABLISSEMENT    : TStringField ;
    QR1BE_DATECOMPTABLE : TDateTimeField ;
    QR1BE_NUMEROPIECE   : TIntegerField ;
    QR1BE_BUDGENE       : TStringField ;
    QR1BE_AXE           : TStringField ;
    QR1BE_QUALIFQTE1    : TStringField ;
    QR1BE_QUALIFQTE2    : TStringField ;
    QR1BE_REFINTERNE    : TStringField ;
    QR1BE_LIBELLE       : TStringField ;
    QR1BE_BUDSECT       : TStringField ;
    QR1BE_QTE1          : TFloatField ;
    QR1BE_QTE2          : TFloatField ;
    QR1BS_LIBELLE, QR1BE_VALIDE,
    QR1BE_TYPESAISIE       : TStringField ;
    QR1DEBIT            : TFloatField ;
    QR1CREDIT           : TFloatField ;
    StReportPiece       : String ;
    TotGen, TotPiece    : TabTot ;
    Procedure BrouilZoom(Quoi : String) ;
    Function  QuoiMvt(i : Integer) : String ;
    Function  LibelTypeEcr(CodeEcr : String) : String ;
    Function  CtrlDateCreat(D1,D2 :TDateTime) : Boolean ;
    Function  RetourneCouple(A, B : String) : String ;
    procedure InfosJournal ;
    procedure DebToDate ;
    procedure FinToDate ;
  public
    { Déclarations publiques }
    procedure GenereSQL ; Override ;
    procedure RecupCritEdt ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure InitDivers ; Override ;
    function  CritOk : Boolean ; Override ;
    procedure ChoixEdition ; Override ;
    procedure FinirPrint ; Override ;
  end;

implementation

{$R *.DFM}

procedure BrouillardBudAna(Code : char) ;
var QR : TFQrBroBudSec ;
    Libelle, FILTEDT : String ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQrBroBudSec.Create(Application) ;
QR.CodeNatureEcr:=Code ;
Edition.Etat:=etBrouBudAna ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BrouilZoom ;
Case Code of
  'N' : BEGIN Libelle:=QR.MsgBox.Mess[6] ; FILTEDT:='BROBUDSECNOR' ; END ;
  'P' : BEGIN Libelle:=QR.MsgBox.Mess[7] ; FILTEDT:='BROBUDSECPRE' ; END ;
  'R' : BEGIN Libelle:=QR.MsgBox.Mess[8] ; FILTEDT:='BROBUDSECREV' ; END ;
  'S' : BEGIN Libelle:=QR.MsgBox.Mess[9] ; FILTEDT:='BROBUDSECSIM' ; END ;
  'U' : BEGIN Libelle:=QR.MsgBox.Mess[10] ; FILTEDT:='BROBUDSECSIT' ; END ;
  End ;
QR.InitType(nbBudJal,neBroBudAna,msRien,FILTEDT,'',TRUE,FALSE,FALSE,Edition) ;
QR.Caption:=QR.MsgBox.Mess[27] ;
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

Procedure TFQrBroBudSec.BrouilZoom(Quoi : String) ;
Var Lp,i : Integer ;
BEGIN
Inherited ;
Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
i:=StrToInt(Copy(Quoi,Lp+1,2)) ;
If (i=10) Then Quoi:=Copy(Quoi,Lp+4,Length(Quoi)-lp-2) ;
ZoomEdt(i,Quoi) ;
END ;

Function TFQrBroBudSec.QuoiMvt(i : Integer) : String ;
BEGIN
if i=0 then
   BEGIN
   Result:=QR1BE_BUDSECT.AsString+'   '+QR1BS_LIBELLE.AsString+
           '#'+Qr1BE_BUDJAL.AsString+' N° '+IntToStr(Qr1BE_NUMEROPIECE.AsInteger)+' '+DateToStr(Qr1BE_DateComptable.AsDAteTime)+'-'+
           PrintSolde(Qr1DEBIT.AsFloat,Qr1Credit.AsFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)+
           '@'+'10;'+Qr1BE_BUDJAL.AsString+';'+UsDateTime(Qr1BE_DATECOMPTABLE.AsDateTime)+';'+Qr1BE_NUMEROPIECE.AsString+';'+Qr1BE_EXERCICE.asString+';'+
           QR1BE_AXE.AsString+';' ;
   END ;
END ;

procedure TFQrBroBudSec.ChoixEdition ;
BEGIN
Inherited ;
ChargeGroup(Lmois,['MOIS']) ;
BE_BUDGENE.Visible:=CritEdt.BroBud.AfficheAnal ;
BG_LIBELLE.Visible:=BE_BUDGENE.Visible ;
TBY_DEBIT.Visible:=BE_BUDGENE.Visible ;
TBY_CREDIT.Visible:=BE_BUDGENE.Visible ;
if CritEdt.BroBud.AfficheAnal then
   BEGIN
   BDetail.Height:=34 ;
   END Else BDetail.Height:=15 ;
If FTotSec.checked then ChargeRecap(LRecap) ;
END ;

procedure TFQrBroBudSec.FinirPrint;
BEGIN
Inherited ;
VideGroup(LMois) ;
If FTotSec.checked then VideRecap(LRecap) ;
END ;

function  TFQrBroBudSec.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ; // and CtrlDateCreat(CritEDT.BrouAna.DateCreat1,CritEDT.BrouAna.DateCreat2) ;
if Result then
   BEGIN
   Fillchar(TotGen, SizeOf(TotGen),#0) ;
   END ;
END ;

Function TFQrBroBudSec.CtrlDateCreat(D1,D2 : TDateTime) : Boolean ;
BEGIN
//Result:=true ;
Result:=(IsValidDate(DateToStr(D1)) and IsValidDate(DateToStr(D2))and(D1 <= D2) ) ;
If Not Result then NumErreurCrit:=7 ;
END ;

procedure TFQrBroBudSec.RecupCritEDT ;
BEGIN
Inherited ;
With CritEDT Do
  BEGIN
  BroBud.Journal:=FJournal.Value ; BroBud.NatureBud:='' ;
  If FNatBud.Value<>'' then BroBud.NatureBud:=FNatBud.Value ;
  Date1:=StrToDate(FDateCompta1.Text)    ; Date2:=StrToDate(FDateCompta2.Text) ;
  DateDeb:=DAte1 ; DateFin:=DAte2 ;
  BroBud.Exo1:=FExercice.Value ; BroBud.Exo2:=FExercice_.Value ;
  BroBud.RefInterne:=FRefInterne.Text       ;
  If FNumPiece1.Text<>'' then BroBud.NumPiece1:=StrToInt(FNumPiece1.Text) else BroBud.NumPiece1:=0 ;
  If FNumPiece2.Text<>'' then BroBud.NumPiece2:=StrToInt(FNumPiece2.Text) else Brobud.NumPiece2:=999999999 ;
  if FNatureCpt.Value<>'' then BroBud.Axe:=FNatureCpt.Value ;
  BroBud.AfficheAnal:=FAfficheAnal.Checked ;
  if FValide.State=cbGrayed then Valide:='g' ;
  if FValide.State=cbChecked then Valide:='X' ;
  if FValide.State=cbUnChecked then Valide:='-' ;
  With BroBud.FormatPrint Do
       BEGIN
       PrSepCompte[1]:=FLigneCpt.Checked ;
       PrSepCompte[2]:=FLignePiecePied.Checked ;
       END ;
  END ;
END ;

procedure TFQrBroBudSec.GenereSQL ;
BEGIN
InHerited ;
Q.Close ; Q.SQL.Clear ;
Q.SQL.Add('Select  BE_EXERCICE, BE_BUDJAL, BE_DATECOMPTABLE, BE_NUMEROPIECE,  ') ;
Q.SQL.Add(       ' BE_BUDGENE, BE_AXE, BE_QUALIFQTE2 , BE_REFINTERNE, BE_LIBELLE, BE_BUDSECT, ') ;
Q.SQL.Add(       ' BE_QTE1, BE_QTE2, BE_QUALIFQTE1, BS_LIBELLE, BE_ETABLISSEMENT, ') ;
Q.SQL.Add(       ' BE_NATUREBUD, BE_VALIDE, BE_TYPESAISIE, BG_LIBELLE, ') ;
Case CritEdt.Monnaie of
  0 : BEGIN Q.SQL.Add(' BE_DEBIT DEBIT, BE_CREDIT CREDIT ') ; END ;
//  1 : BEGIN Q.SQL.Add(' BE_DEBITDEV DEBIT, BE_CREDITDEV CREDIT ')   ; END ;
//  2 : BEGIN Q.SQL.Add(' BE_DEBITEURO DEBIT, BE_CREDITEURO CREDIT ') ; END ;
 end ;
 { Tables explorées par la SQL }
Q.SQL.Add(' From BUDECR ') ;
Q.SQL.Add(       ' LEFT JOIN BUDSECT on BS_AXE=BE_AXE And BS_BUDSECT=BE_BUDSECT ') ;
Q.SQL.Add(       ' LEFT JOIN BUDGENE on BG_BUDGENE=BE_BUDGENE ') ;
{ Construction de la clause Where de la SQL }
//Q.SQL.Add(' Where (BE_TYPESAISIE="S" or BE_TYPESAISIE="SG") ') ;
Q.SQL.Add(' Where BE_DATECOMPTABLE>="'+USDateTime(CritEdt.Date1)+'" And BE_DATECOMPTABLE<="'+USDateTime(CritEdt.Date2)+'" ') ;
Q.SQL.Add(' And BE_QUALIFPIECE="'+CodeNatureEcr+'" ') ;
if CritEdt.BroBudAna.Axe<>'' then Q.SQL.Add(' And BE_AXE="'+CritEdt.BroBudAna.Axe+'" ') ;
Q.SQL.Add(' AND BE_BUDJAL="'+CritEdt.BroBudAna.Journal+'" ') ;
Q.SQL.Add(' And BE_NUMEROPIECE>='+IntToStr(CritEdt.BroBudAna.NumPiece1)+' ') ;
Q.SQL.Add(' And BE_NUMEROPIECE<='+IntToStr(CritEdt.BroBudAna.NumPiece2)+' ') ;
Q.SQL.Add('And BE_EXERCICE>="'+CritEdt.BroBudAna.Exo1+'" And BE_EXERCICE<="'+CritEdt.BroBudAna.Exo2+'" ') ;
if CritEdt.BroBudAna.RefInterne<>'' then Q.SQL.Add(' And UPPER(BE_REFINTERNE) like "'+TraduitJoker(CritEdt.BroBudAna.RefInterne)+'" ' );
if CritEdt.Etab<>'' then Q.SQL.Add(' And BE_ETABLISSEMENT="'+CritEdt.Etab+'" ') ;
if (OkV2 and (V_PGI.Confidentiel<>'1')) then Q.SQL.Add('AND BE_CONFIDENTIEL<>"1" ') ;
if CritEdt.Valide<>'g' then Q.SQL.Add(' And BE_VALIDE="'+CritEdt.Valide+'" ') ;
If CritEdt.BroBud.NatureBud<>'' Then Q.SQL.Add(' And BE_NATUREBUD="'+CritEdt.BroBud.NatureBud+'" ') ;
{ Construction de la clause Order By de la SQL }
Q.SQL.Add(' Order By BE_BUDJAL, BE_NATUREBUD, BE_QUALIFPIECE, BE_NUMEROPIECE, BE_DATECOMPTABLE, BE_BUDSECT ') ;
ChangeSQL(Q) ; Q.Open ;
END ;

procedure TFQrBroBudSec.InitDivers ;
BEGIN
InHerited ;
BEntetePiece.Frame.DrawTop:=CritEdt.BroBudAna.FormatPrint.PrSepCompte[1] ;
BEntetePiece.Frame.DrawBottom:=CritEdt.BroBudAna.FormatPrint.PrSepCompte[1] ;
BDetail.Frame.DrawTop:=False ;
BDetail.Frame.DrawBottom:=False ;
BPiedPiece.Frame.DrawTop:=CritEdt.BroBudAna.FormatPrint.PrSepCompte[2] ;
BPiedPiece.Frame.DrawBottom:=CritEdt.BroBudAna.FormatPrint.PrSepCompte[2] ;
END ;

procedure TFQrBroBudSec.RenseigneCritere ;
{ Récupération des champs du multicritère dans l'entête d'état }
BEGIN
Inherited ;
RNatureCpt.Caption:=FNatureCpt.Text ;
RExercice_.Caption:=FExercice_.Text ;
RRefInterne.Caption:=FRefInterne.Text ;
RNumPiece1.Caption:=IntToStr(CritEdt.BroBudAna.NumPiece1) ; RNumPiece2.Caption:=IntToStr(CritEdt.BroBudAna.NumPiece2) ;
RNatureEcr.Caption:=LibelTypeEcr(CodeNatureEcr) ;
END ;

Function TFQrBroBudSec.LibelTypeEcr(CodeEcr : String) : String ;
BEGIN
Result:='' ;
if CodeEcr='N' then Result:=MsgBox.Mess[6] else
if CodeEcr='P' then Result:=MsgBox.Mess[7] else
if CodeEcr='R' then Result:=MsgBox.Mess[8] else
if CodeEcr='S' then Result:=MsgBox.Mess[9] else
if CodeEcr='U' then Result:=MsgBox.Mess[10] ;
END;

procedure TFQrBroBudSec.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
QR1BE_TYPESAISIE     :=TStringField(Q.FindField('BE_TYPESAISIE')) ;
QR1BE_VALIDE         :=TStringField(Q.FindField('BE_VALIDE')) ;
QR1BE_ETABLISSEMENT  :=TStringField(Q.FindField('BE_ETABLISSEMENT')) ;
QR1BE_EXERCICE       :=TStringField(Q.FindField('BE_EXERCICE')) ;
QR1BE_BUDJAL        :=TStringField(Q.FindField('BE_BUDJAL')) ;
QR1BE_DATECOMPTABLE  :=TDateTimeField(Q.FindField('BE_DATECOMPTABLE')) ;
QR1BE_DATECOMPTABLE.Tag:=1 ;
QR1BE_NUMEROPIECE    :=TIntegerField(Q.FindField('BE_NUMEROPIECE')) ;
QR1BE_BUDSECT        :=TStringField(Q.FindField('BE_BUDSECT')) ;
QR1BE_BUDGENE        :=TStringField(Q.FindField('BE_BUDGENE')) ;
QR1BE_AXE            :=TStringField(Q.FindField('BE_AXE')) ;
QR1BE_QUALIFQTE1     :=TStringField(Q.FindField('BE_QUALIFQTE1')) ;
QR1BE_QUALIFQTE2     :=TStringField(Q.FindField('BE_QUALIFQTE2')) ;
QR1BE_REFINTERNE     :=TStringField(Q.FindField('BE_REFINTERNE')) ;
QR1BE_LIBELLE        :=TStringField(Q.FindField('BE_LIBELLE')) ;
QR1BE_QTE1           :=TFloatField(Q.FindField('BE_QTE1')) ;
QR1BE_QTE2           :=TFloatField(Q.FindField('BE_QTE2')) ;
QR1BS_LIBELLE      :=TStringField(Q.FindField('BS_LIBELLE')) ;
QR1BE_NATUREBUD    :=TStringField(Q.FindField('BE_NATUREBUD')) ;
QR1DEBIT            :=TFloatField(Q.FindField('DEBIT')) ;
QR1CREDIT           :=TFloatField(Q.FindField('CREDIT')) ;
ChgMaskChamp(QR1DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
ChgMaskChamp(QR1CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
end;

procedure TFQrBroBudSec.BEntetePieceBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var St1, St2 : String ;
begin
  inherited;
St1:='' ; St2:='' ;
TY_JOURNAL.Left:=TS_LIBELLE.Left ;
TY_JOURNAL.Width:=(TS_LIBELLE.Width+TY_REFINTERNE.Width)+1 ;
TY_NUMEROPIECE.Left:=TCptGen.Left ;
TY_NUMEROPIECE.Width:=TCptGen.Width ;
TY_NUMEROPIECE.Caption:=MsgBox.Mess[2]+' '+IntToStr(QR1BE_NUMEROPIECE.AsInteger) ;
//St1:=DateToStr(QR1BE_DATECOMPTABLE.AsDateTime)+'  '+QR1BE_AXE.AsString ;
St1:=QR1BE_AXE.AsString ;
St2:=QR1BE_BUDJAL.AsString+'   '+RechDom('ttBudJal',QR1BE_BUDJAL.AsString,False)  ;
TY_JOURNAL.Caption:=St1+'  '+St2 ;
Y_ETABLISSEMENT.Caption:=RechDom('ttEtablissement',QR1BE_ETABLISSEMENT.AsString,False) ;
StReportPiece:=QR1BE_NUMEROPIECE.AsString ;
InitReport([2],CritEdt.BroBudAna.FormatPrint.Report) ;
If FTotSec.checked Then BEGIN VideRecap(LRecap) ; ChargeRecap(LRecap) ; END ;
end;

procedure TFQrBroBudSec.BPiedPieceBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
TY_NUMEROPIECE_.Caption:=MsgBox.Mess[3]+' '+IntToStr(QR1BE_NUMEROPIECE.AsInteger) ;
TotDebitPiece.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotPiece[1].TotDebit,  CritEdt.AfficheSymbole) ;
TotCreditPiece.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotPiece[1].TotCredit, CritEdt.AfficheSymbole) ;
end;

procedure TFQrBroBudSec.BFinEtatBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
TotDebitGen.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotGen[1].TotDebit,  CritEdt.AfficheSymbole) ;
TotCreditGen.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotGen[1].TotCredit, CritEdt.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFQrBroBudSec.BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
var MReport : TabTRep ;
begin
  inherited;
Case QuelReportBAL(CritEdt.BroBudAna.FormatPrint.Report,MReport) of
  1 : BEGIN TITREREPORTB.Caption:=MsgBox.Mess[23] END ;
  2 : BEGIN TITREREPORTB.Caption:=MsgBox.Mess[24]+' '+StReportPiece ; END ;
  END ;
Report2Debit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[1].TotDebit,  CritEdt.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[1].TotCredit, CritEdt.AfficheSymbole ) ;
end;

procedure TFQrBroBudSec.BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
var MReport : TabTRep ;
    StRecap : String ;
    StDate  : String ;
begin
  inherited;
Fillchar(MReport,SizeOf(MReport),#0) ;
if QR1BE_VALIDE.AsString='X' then BE_VALIDE.Caption:='V' else BE_VALIDE.Caption:=' ' ;
BE_NATUREBUD.Caption:=QR1BE_NATUREBUD.AsString ;
MReport[1].TotDebit:=QR1DEBIT.AsFloat ;
MReport[1].TotCredit:=QR1CREDIT.AsFloat ;
TotGen[1].TotDebit:= Arrondi(TotGen[1].TotDebit+QR1DEBIT.AsFloat,CritEdt.Decimale) ;
TotGen[1].TotCredit:=Arrondi(TotGen[1].TotCredit+QR1CREDIT.AsFloat,CritEdt.Decimale) ;
TotPiece[1].TotDebit:= Arrondi(TotPiece[1].TotDebit+QR1DEBIT.AsFloat,CritEdt.Decimale) ;
TotPiece[1].TotCredit:=Arrondi(TotPiece[1].TotCredit+QR1CREDIT.AsFloat,CritEdt.Decimale) ;
AddReportBAL([1,2], CritEdt.BroBudAna.FormatPrint.Report, MReport, CritEdt.Decimale) ;
AddGroup(Lmois, [QR1BE_DATECOMPTABLE],[QR1DEBIT.AsFloat,QR1CREDIT.AsFloat]) ;
StRecap:=QR1BE_BUDGENE.AsString ;
StDate:=FormatDatetime('mmyyyy',QR1BE_DATECOMPTABLE.AsDateTime) ;
If FTotSec.checked then AddRecap(LRecap, [StRecap], [StRecap],[1,QR1DEBIT.AsFloat, QR1CREDIT.AsFloat,0,0]) ;
if PrintBand then Quoi:=QuoiMvt(0) ;
end;

procedure TFQrBroBudSec.FormShow(Sender: TObject);
begin
HelpContext:=15284000 ;
Standards.HelpContext:=15284200 ;
Avances.HelpContext:=15284100 ;
Mise.HelpContext:=15284400 ;
Option.HelpContext:=15284300 ;
TabSup.TabVisible:=False; FValide.Visible:=(CodeNatureEcr<>'S')and(CodeNatureEcr<>'R') ;
If FJournal.Values.Count>0 Then FJournal.Value:=FJournal.Values[0] ;
FNatBud.ItemIndex:=0 ;
DebToDate ; FinToDate ;
  inherited;
end;

procedure TFQrBroBudSec.TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
end;

procedure TFQrBroBudSec.BEntetePieceAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Fillchar(TotPiece, SizeOf(TotPiece),#0) ;
end;

procedure TFQrBroBudSec.BPiedPieceAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([2],CritEdt.BroBudAna.FormatPrint.Report) ;
end;

Function TFQrBroBudSec.RetourneCouple(A, B : String) : String ;
BEGIN
Result:='' ;
if (A<>'')and(B<>'') then Result:=A+' / '+B else
if (A<>'')and(B='') then Result:=A else
if (A='')and(B<>'') then Result:=B else Exit ;
END ;

procedure TFQrBroBudSec.InfosJournal ;
Var QJal  : TQuery ;
BEGIN
QJal:=OpenSQL('Select BJ_EXODEB, BJ_EXOFIN, BJ_PERDEB, BJ_PERFIN from BUDJAL Where BJ_BUDJAL="'+FJournal.Value+'"',True) ;
if Not QJal.EOF then
   BEGIN
   FEXERCICE.Value:=QJal.FindField('BJ_EXODEB').AsString ; FEXERCICE_.Value:=QJal.FindField('BJ_EXOFIN').AsString ;
   FPeriode1.Value:=QJal.FindField('BJ_PERDEB').AsString ; FPeriode2.Value:=QJal.FindField('BJ_PERFIN').AsString ;
   FDateCompta1.Text:=QJal.FindField('BJ_PERDEB').AsString ; FDateCompta2.Text:=QJal.FindField('BJ_PERFIN').AsString ;
   END ;
Ferme(QJal) ;
END ;

procedure TFQrBroBudSec.FExerciceChange(Sender: TObject);
begin
if FExercice.Value>FExercice_.Value then FExercice_.Value:=FExercice.Value ;
ListePeriode(FExercice.Value,FPeriode1.Items,FPeriode1.Values,True) ;
FPeriode1.ItemIndex:=0 ; DebToDate ;
end;

procedure TFQrBroBudSec.FExercice_Change(Sender: TObject);
begin
if FExercice_.Value<FExercice.Value then FExercice.Value:=FExercice_.Value ;
ListePeriode(FExercice_.Value,FPeriode2.Items,FPeriode2.Values,False) ;
FPeriode2.ItemIndex:=FPeriode2.Items.Count-1 ; FinToDate ;
end;

procedure TFQrBroBudSec.FPeriode1Change(Sender: TObject);
begin
  inherited;
if StrToDate(FPeriode1.Value)>StrToDate(FPeriode2.Value) then FPeriode2.Value:=FPeriode1.Value ;
FDateCompta1.Text:=FPeriode1.Value ;
end;

procedure TFQrBroBudSec.FPeriode2Change(Sender: TObject);
begin
  inherited;
if StrToDate(FPeriode2.Value)<StrToDate(FPeriode1.Value) then FPeriode1.Value:=FPeriode2.Value ;
FDateCompta2.Text:=FPeriode2.Value ;
end;

procedure TFQrBroBudSec.FJournalChange(Sender: TObject);
begin
  inherited;
InfosJournal ;
FCpte1.Text:=FJournal.Value ; FCpte2.Text:=FCpte1.Text
end;

procedure TFQrBroBudSec.DebToDate ;
BEGIN
FDateCompta1.Text:=FPeriode1.Value ;
END ;

procedure TFQrBroBudSec.FinToDate ;
BEGIN
FDateCompta2.Text:=FPeriode2.Value ;
END ;

procedure TFQrBroBudSec.FNatureCptChange(Sender: TObject);
begin
//  inherited;
end;

procedure TFQrBroBudSec.DLMoisNeedData(var MoreData: Boolean);
Var Cod,Lib : String ;
    Tot : Array[0..12] of Double ;
    NumR : Integer ;
begin
  inherited;
MoreData:=PrintGroup(Lmois,Q,[QR1BE_DATECOMPTABLE],Cod,Lib,Tot,NumR);
if MoreData then
   BEGIN
   TTOTMULTI.Caption:=Cod;
   TDEBMULTI.Caption:=AfficheMontant(CRITEdt.FormatMontant, CRITEdt.Symbole, Tot[0] , CRITEdt.AfficheSymbole) ; { + 'ANouveau' + 'Cumul au' en Débit  }
   TCREMULTI.Caption:=AfficheMontant(CRITEdt.FormatMontant, CRITEdt.Symbole, Tot[1] , CRITEdt.AfficheSymbole) ; { + 'ANouveau' + 'Cumul au' en Crédit }
   END ;
end;

procedure TFQrBroBudSec.FCpte1Change(Sender: TObject);
begin
//  inherited;

end;

procedure TFQrBroBudSec.DLRecapNeedData(var MoreData: Boolean);
Var Tot : Array[0..12] of Double ;
    Cod,Lib : String ;
begin
  inherited;
If FTotSec.checked then
   BEGIN
   MoreData:=PrintRecap(LRecap,Cod,Lib,Tot) ;
   If MoreData Then
      BEGIN
      RDEBIT.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, Tot[1],  CritEdt.AfficheSymbole ) ;
      RCREDIT.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, Tot[2],  CritEdt.AfficheSymbole ) ;
      END ;
   END Else MoreData:=False ;
end;

end.
