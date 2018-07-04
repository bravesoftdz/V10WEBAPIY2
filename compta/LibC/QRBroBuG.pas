{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 31/03/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par BROBUDGEN_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit QrBroBuG;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, HSysMenu, Menus, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Buttons,
  Hctrls, ExtCtrls, Mask, Hcompte, ComCtrls, UtilEDT, CritEDT, Ent1, HEnt1,
  CpteUtil, QRRupt, HTB97, HPanel, UiUtil, tCalcCum ;

procedure BrouillardBud(Code : char) ;

type
  TFQrBroBudGen = class(TFQR)
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
    EB_Valide: TQRLabel;
    BE_BUDGENE: TQRDBText;
    S_LIBELLE: TQRDBText;
    TBE_LIBELLE: TQRDBText;
    TBE_REFINTERNE: TQRDBText;
    TBE_DEBIT: TQRDBText;
    TBE_CREDIT: TQRDBText;
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
    BAnalytique: TQRBand;
    RDEBIT: TQRDBText;
    LIBRECAP: TQRDBText;
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
    BS_LIBELLE: TQRDBText;
    TBY_DEBIT: TQRDBText;
    TBY_CREDIT: TQRDBText;
    BE_BUDSECT: TQRDBText;
    QRLabel1: TQRLabel;
    BE_NATUREBUD: TQRLabel;
    RCREDIT: TQRDBText;
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
    QR1BG_LIBELLE, QR1BE_VALIDE,
    QR1BE_TYPESAISIE       : TStringField ;
    QR1DEBIT            : TFloatField ;
    QR1CREDIT           : TFloatField ;
    StReportPiece       : String ;
    TotGen, TotPiece    : TabTot ;
    Procedure BrouilZoom(Quoi : String) ;
    Function  QuoiMvt(i : Integer) : String ;
    Function  LibelTypeEcr(CodeEcr : String) : String ;
//    Function  CtrlDateCreat(D1,D2 :TDateTime) : Boolean ;
//    Function  RetourneCouple(A, B : String) : String ;
    procedure InfosJournal ;
    procedure DebToDate ;
    procedure FinToDate ;
  public
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

procedure BrouillardBud(Code : char) ;
var QR : TFQrBroBudGen ;
    Libelle, FILTEDT : String ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQrBroBudGen.Create(Application) ;
QR.CodeNatureEcr:=Code ;
Edition.Etat:=etBrouBud ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BrouilZoom ;
Case Code of
  'N' : BEGIN Libelle:=QR.MsgBox.Mess[6] ; FILTEDT:='BROBUDGENNOR' ; END ;
  'P' : BEGIN Libelle:=QR.MsgBox.Mess[7] ; FILTEDT:='BROBUDGENPRE' ; END ;
  'R' : BEGIN Libelle:=QR.MsgBox.Mess[8] ; FILTEDT:='BROBUDGENREV' ; END ;
  'S' : BEGIN Libelle:=QR.MsgBox.Mess[9] ; FILTEDT:='BROBUDGENSIM' ; END ;
  'U' : BEGIN Libelle:=QR.MsgBox.Mess[10] ; FILTEDT:='BROBUDGENSIT' ; END ;
  End ;
QR.InitType(nbBudJal,neBroBud,msRien,FILTEDT,'',TRUE,FALSE,FALSE,Edition) ;
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

Procedure TFQrBroBudGen.BrouilZoom(Quoi : String) ;
Var Lp,i : Integer ;
BEGIN
Inherited ;
Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
i:=StrToInt(Copy(Quoi,Lp+1,2)) ;
If (i=10) Then Quoi:=Copy(Quoi,Lp+4,Length(Quoi)-lp-2) ;
ZoomEdt(i,Quoi) ;
END ;

Function TFQrBroBudGen.QuoiMvt(i : Integer) : String ;
BEGIN
if i=0 then
   BEGIN
   Result:=QR1BE_BUDGENE.AsString+'   '+QR1BG_LIBELLE.AsString+
           '#'+Qr1BE_BUDJAL.AsString+' N° '+IntToStr(Qr1BE_NUMEROPIECE.AsInteger)+' '+DateToStr(Qr1BE_DateComptable.AsDAteTime)+'-'+
           PrintSolde(Qr1DEBIT.AsFloat,Qr1Credit.AsFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)+
           '@'+'10;'+Qr1BE_BUDJAL.AsString+';'+UsDateTime(Qr1BE_DATECOMPTABLE.AsDateTime)+';'+IntToStr(Qr1BE_NUMEROPIECE.AsInteger)+';'+Qr1BE_EXERCICE.asString+';' ;
   END ;
END ;

procedure TFQrBroBudGen.ChoixEdition ;
BEGIN
Inherited ;
ChargeGroup(Lmois,['MOIS']) ;
BE_BUDSECT.Visible:=CritEdt.BroBud.AfficheAnal ;
BS_LIBELLE.Visible:=BE_BUDSECT.Visible ;
TBY_DEBIT.Visible:=BE_BUDSECT.Visible ;
TBY_CREDIT.Visible:=BE_BUDSECT.Visible ;
if CritEdt.BroBud.AfficheAnal then
   BEGIN
   BDetail.Height:=34 ;
   END Else BDetail.Height:=15 ;
If FTotSec.checked then ChargeRecap(LRecap) ;
END ;

procedure TFQrBroBudGen.FinirPrint;
BEGIN
Inherited ;
VideGroup(LMois) ;
If FTotSec.checked then VideRecap(LRecap) ;
END ;

function  TFQrBroBudGen.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ; //and CtrlDateCreat(CritEDT.BroBud.DateCreat1,CritEDT.BrouAna.DateCreat2) ;
if Result then
   BEGIN
   Fillchar(TotGen, SizeOf(TotGen),#0) ;
   END ;
END ;

{Function TFQrBroBudGen.CtrlDateCreat(D1,D2 : TDateTime) : Boolean ;
BEGIN
//Result:=true ;
Result:=(IsValidDate(DateToStr(D1)) and IsValidDate(DateToStr(D2))and(D1 <= D2) ) ;
If Not Result then NumErreurCrit:=7 ;
END ;}

procedure TFQrBroBudGen.RecupCritEDT ;
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

procedure TFQrBroBudGen.GenereSQL ;
BEGIN
InHerited ;
Q.Close ; Q.SQL.Clear ;
Q.SQL.Add('Select  BE_EXERCICE, BE_BUDJAL, BE_DATECOMPTABLE, BE_NUMEROPIECE,  ') ;
Q.SQL.Add(       ' BE_BUDGENE, BE_AXE, BE_QUALIFQTE2 , BE_REFINTERNE, BE_LIBELLE, BE_BUDSECT, ') ;
Q.SQL.Add(       ' BE_QTE1, BE_QTE2, BE_QUALIFQTE1, BG_LIBELLE, BE_ETABLISSEMENT, ') ;
Q.SQL.Add(       ' BE_NATUREBUD, BE_VALIDE, BE_TYPESAISIE, BS_LIBELLE, ') ;
Case CritEdt.Monnaie of
  0 : BEGIN Q.SQL.Add(' BE_DEBIT DEBIT, BE_CREDIT CREDIT ')         ; END ;
//  1 : BEGIN Q.SQL.Add(' BE_DEBITDEV DEBIT, BE_CREDITDEV CREDIT ')   ; END ;
//  2 : BEGIN Q.SQL.Add(' BE_DEBITEURO DEBIT, BE_CREDITEURO CREDIT ') ; END ;
 end ;
 { Tables explorées par la SQL }
Q.SQL.Add(' From BUDECR ') ;
Q.SQL.Add(       ' LEFT JOIN BUDGENE on BG_BUDGENE=BE_BUDGENE ') ;
Q.SQL.Add(       ' LEFT JOIN BUDSECT On BS_AXE=BE_AXE and BS_BUDSECT=BE_BUDSECT ') ;
{ Construction de la clause Where de la SQL }
//Q.SQL.Add(' Where (BE_TYPESAISIE="G" or BE_TYPESAISIE="GS") ') ;
Q.SQL.Add(' Where BE_DATECOMPTABLE>="'+USDateTime(CritEdt.Date1)+'" And BE_DATECOMPTABLE<="'+USDateTime(CritEdt.Date2)+'" ') ;
Q.SQL.Add(' And BE_QUALIFPIECE="'+CodeNatureEcr+'" ') ;
if CritEdt.BroBud.Axe<>'' then Q.SQL.Add(' And BE_AXE="'+CritEdt.BroBud.Axe+'" ') ;
Q.SQL.Add(' AND BE_BUDJAL="'+CritEdt.BroBud.Journal+'" ') ;
Q.SQL.Add(' And BE_NUMEROPIECE>='+IntToStr(CritEdt.BroBud.NumPiece1)+' ') ;
Q.SQL.Add(' And BE_NUMEROPIECE<='+IntToStr(CritEdt.BroBud.NumPiece2)+' ') ;
Q.SQL.Add('And BE_EXERCICE>="'+CritEdt.BroBud.Exo1+'" And BE_EXERCICE<="'+CritEdt.BroBud.Exo2+'" ') ;
if CritEdt.BroBud.RefInterne<>'' then Q.SQL.Add(' And UPPER(BE_REFINTERNE) like "'+TraduitJoker(CritEdt.BroBud.RefInterne)+'" ' );
if CritEdt.Etab<>'' then Q.SQL.Add(' And BE_ETABLISSEMENT="'+CritEdt.Etab+'" ') ;
if (OkV2 and (V_PGI.Confidentiel<>'1')) then Q.SQL.Add('AND BE_CONFIDENTIEL<>"1" ') ;
if CritEdt.Valide<>'g' then Q.SQL.Add(' And BE_VALIDE="'+CritEdt.Valide+'" ') ;
If CritEdt.BroBud.NatureBud<>'' Then Q.SQL.Add(' And BE_NATUREBUD="'+CritEdt.BroBud.NatureBud+'" ') ;
{ Construction de la clause Order By de la SQL }
//Q.SQL.Add(' Order By BE_BUDJAL, BE_NATUREBUD, BE_QUALIFPIECE, BE_NUMEROPIECE, BE_DATECOMPTABLE, BE_BUDGENE ') ;
Q.SQL.Add(' Order By BE_NUMEROPIECE, BE_DATECOMPTABLE, BE_BUDGENE, BE_BUDJAL, BE_NATUREBUD, BE_QUALIFPIECE ') ;
ChangeSQL(Q) ; Q.Open ;
END ;

procedure TFQrBroBudGen.InitDivers ;
BEGIN
InHerited ;
BEntetePiece.Frame.DrawTop:=CritEdt.BroBud.FormatPrint.PrSepCompte[1] ;
BEntetePiece.Frame.DrawBottom:=CritEdt.BroBud.FormatPrint.PrSepCompte[1] ;
BDetail.Frame.DrawTop:=False ;
BDetail.Frame.DrawBottom:=False ;
BPiedPiece.Frame.DrawTop:=CritEdt.BroBud.FormatPrint.PrSepCompte[2] ;
BPiedPiece.Frame.DrawBottom:=CritEdt.BroBud.FormatPrint.PrSepCompte[2] ;
END ;

procedure TFQrBroBudGen.RenseigneCritere ;
{ Récupération des champs du multicritère dans l'entête d'état }
BEGIN
Inherited ;
RNatureCpt.Caption:=FNatureCpt.Text ;
RExercice_.Caption:=FExercice_.Text ;
RRefInterne.Caption:=FRefInterne.Text ;
RNumPiece1.Caption:=IntToStr(CritEdt.BroBud.NumPiece1) ; RNumPiece2.Caption:=IntToStr(CritEdt.BroBud.NumPiece2) ;
RNatureEcr.Caption:=LibelTypeEcr(CodeNatureEcr) ;
END ;

Function TFQrBroBudGen.LibelTypeEcr(CodeEcr : String) : String ;
BEGIN
Result:='' ;
if CodeEcr='N' then Result:=MsgBox.Mess[6] else
if CodeEcr='P' then Result:=MsgBox.Mess[7] else
if CodeEcr='R' then Result:=MsgBox.Mess[8] else
if CodeEcr='S' then Result:=MsgBox.Mess[9] else
if CodeEcr='U' then Result:=MsgBox.Mess[10] ;
END;

procedure TFQrBroBudGen.QAfterOpen(DataSet: TDataSet);
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
QR1BE_BUDGENE       :=TStringField(Q.FindField('BE_BUDGENE')) ;
QR1BE_BUDSECT       :=TStringField(Q.FindField('BE_BUDSECT')) ;
QR1BE_AXE            :=TStringField(Q.FindField('BE_AXE')) ;
QR1BE_QUALIFQTE1     :=TStringField(Q.FindField('BE_QUALIFQTE1')) ;
QR1BE_QUALIFQTE2     :=TStringField(Q.FindField('BE_QUALIFQTE2')) ;
QR1BE_REFINTERNE     :=TStringField(Q.FindField('BE_REFINTERNE')) ;
QR1BE_LIBELLE        :=TStringField(Q.FindField('BE_LIBELLE')) ;
QR1BE_QTE1           :=TFloatField(Q.FindField('BE_QTE1')) ;
QR1BE_QTE2           :=TFloatField(Q.FindField('BE_QTE2')) ;
QR1BG_LIBELLE      :=TStringField(Q.FindField('BG_LIBELLE')) ;
QR1BE_NATUREBUD    :=TStringField(Q.FindField('BE_NATUREBUD')) ;
QR1DEBIT            :=TFloatField(Q.FindField('DEBIT')) ;
QR1CREDIT           :=TFloatField(Q.FindField('CREDIT')) ;
ChgMaskChamp(QR1DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
ChgMaskChamp(QR1CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
end;

procedure TFQrBroBudGen.BEntetePieceBeforePrint(var PrintBand: Boolean; var Quoi: string);
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
InitReport([2],CritEdt.BroBud.FormatPrint.Report) ;
If FTotSec.checked Then BEGIN VideRecap(LRecap) ; ChargeRecap(LRecap) ; END ;
end;

procedure TFQrBroBudGen.BPiedPieceBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
TY_NUMEROPIECE_.Caption:=MsgBox.Mess[3]+' '+IntToStr(QR1BE_NUMEROPIECE.AsInteger) ;
TotDebitPiece.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotPiece[1].TotDebit,  CritEdt.AfficheSymbole) ;
TotCreditPiece.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotPiece[1].TotCredit, CritEdt.AfficheSymbole) ;
end;

procedure TFQrBroBudGen.BFinEtatBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
TotDebitGen.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotGen[1].TotDebit,  CritEdt.AfficheSymbole) ;
TotCreditGen.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotGen[1].TotCredit, CritEdt.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFQrBroBudGen.BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
var MReport : TabTRep ;
begin
  inherited;
Case QuelReportBAL(CritEdt.BroBud.FormatPrint.Report,MReport) of
  1 : BEGIN TITREREPORTB.Caption:=MsgBox.Mess[23] END ;
  2 : BEGIN TITREREPORTB.Caption:=MsgBox.Mess[24]+' '+StReportPiece ; END ;
  END ;
Report2Debit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[1].TotDebit,  CritEdt.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[1].TotCredit, CritEdt.AfficheSymbole ) ;
end;

procedure TFQrBroBudGen.BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
var MReport : TabTRep ;
    StRecap : String ;
    StDate  : String ;
begin
  inherited;
Fillchar(MReport,SizeOf(MReport),#0) ;
if QR1BE_VALIDE.AsString='X' then EB_VALIDE.Caption:='V' else EB_VALIDE.Caption:=' ' ;
BE_NATUREBUD.Caption:=QR1BE_NATUREBUD.AsString ;
TBE_REFINTERNE.Caption:=QR1BE_REFINTERNE.AsString ;
TBE_LIBELLE.Caption:=QR1BE_LIBELLE.AsString ;
MReport[1].TotDebit:=QR1DEBIT.AsFloat ;
MReport[1].TotCredit:=QR1CREDIT.AsFloat ;
TotGen[1].TotDebit:= Arrondi(TotGen[1].TotDebit+QR1DEBIT.AsFloat,CritEdt.Decimale) ;
TotGen[1].TotCredit:=Arrondi(TotGen[1].TotCredit+QR1CREDIT.AsFloat,CritEdt.Decimale) ;
TotPiece[1].TotDebit:= Arrondi(TotPiece[1].TotDebit+QR1DEBIT.AsFloat,CritEdt.Decimale) ;
TotPiece[1].TotCredit:=Arrondi(TotPiece[1].TotCredit+QR1CREDIT.AsFloat,CritEdt.Decimale) ;
AddReportBAL([1,2], CritEdt.BroBud.FormatPrint.Report, MReport, CritEdt.Decimale) ;
AddGroup(Lmois, [QR1BE_DATECOMPTABLE],[QR1DEBIT.AsFloat,QR1Credit.AsFloat]) ;
//StRecap:=QR1BE_BUDSECT.AsString ;
StRecap:=QR1BE_BUDGENE.AsString ;
StDate:=FormatDatetime('mmyyyy',QR1BE_DATECOMPTABLE.AsDateTime) ;
//If FTotSec.checked then AddRecap(LRecap, [StRecap2+StRecap], [StRecap],[1,QR1DEBIT.AsFloat, QR1CREDIT.AsFloat,0,0]) ;
If FTotSec.checked then AddRecap(LRecap, [StRecap,StDate], [StRecap, StDate],[1,QR1DEBIT.AsFloat, QR1CREDIT.AsFloat,0,0]) ;
if PrintBand then Quoi:=QuoiMvt(0) ;
end;

procedure TFQrBroBudGen.FormShow(Sender: TObject);
begin
HelpContext:=15282000 ;
Standards.HelpContext:=15282200 ;
Avances.HelpContext:=15282100 ;
Mise.HelpContext:=15282400 ;
Option.HelpContext:=15282300 ;
TabSup.TabVisible:=False; FValide.Visible:=(CodeNatureEcr<>'S')and(CodeNatureEcr<>'R') ;
If FJournal.Values.Count>0 Then FJournal.Value:=FJournal.Values[0] ;
FNatBud.ItemIndex:=0 ;
DebToDate ; FinToDate ;
  inherited;
end;

procedure TFQrBroBudGen.TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
end;

procedure TFQrBroBudGen.BEntetePieceAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Fillchar(TotPiece, SizeOf(TotPiece),#0) ;

end;

procedure TFQrBroBudGen.BPiedPieceAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([2],CritEdt.BroBud.FormatPrint.Report) ;
end;

{Function TFQrBroBudGen.RetourneCouple(A, B : String) : String ;
BEGIN
Result:='' ;
if (A<>'')and(B<>'') then Result:=A+' / '+B else
if (A<>'')and(B='') then Result:=A else
if (A='')and(B<>'') then Result:=B else Exit ;
END ;}

procedure TFQrBroBudGen.InfosJournal ;
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

procedure TFQrBroBudGen.FExerciceChange(Sender: TObject);
begin
if FExercice.Value>FExercice_.Value then FExercice_.Value:=FExercice.Value ;
ListePeriode(FExercice.Value,FPeriode1.Items,FPeriode1.Values,True) ;
FPeriode1.ItemIndex:=0 ; DebToDate ;
end;

procedure TFQrBroBudGen.FExercice_Change(Sender: TObject);
begin
if FExercice_.Value<FExercice.Value then FExercice.Value:=FExercice_.Value ;
ListePeriode(FExercice_.Value,FPeriode2.Items,FPeriode2.Values,False) ;
FPeriode2.ItemIndex:=FPeriode2.Items.Count-1 ; FinToDate ;
end;

procedure TFQrBroBudGen.FPeriode1Change(Sender: TObject);
begin
  inherited;
if StrToDate(FPeriode1.Value)>StrToDate(FPeriode2.Value) then FPeriode2.Value:=FPeriode1.Value ;
FDateCompta1.Text:=FPeriode1.Value ;
end;

procedure TFQrBroBudGen.FPeriode2Change(Sender: TObject);
begin
  inherited;
if StrToDate(FPeriode2.Value)<StrToDate(FPeriode1.Value) then FPeriode1.Value:=FPeriode2.Value ;
FDateCompta2.Text:=FPeriode2.Value ;
end;

procedure TFQrBroBudGen.FJournalChange(Sender: TObject);
begin
  inherited;
InfosJournal ;
FCpte1.Text:=FJournal.Value ; FCpte2.Text:=FCpte1.Text
end;

procedure TFQrBroBudGen.DebToDate ;
BEGIN
FDateCompta1.Text:=FPeriode1.Value ;
END ;

procedure TFQrBroBudGen.FinToDate ;
BEGIN
FDateCompta2.Text:=FPeriode2.Value ;
END ;

procedure TFQrBroBudGen.FNatureCptChange(Sender: TObject);
begin
//  inherited;

end;

procedure TFQrBroBudGen.DLMoisNeedData(var MoreData: Boolean);
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

procedure TFQrBroBudGen.FCpte1Change(Sender: TObject);
begin
//  inherited;

end;

procedure TFQrBroBudGen.DLRecapNeedData(var MoreData: Boolean);
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
