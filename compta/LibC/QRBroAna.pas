unit QrBroAna;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, HSysMenu, Menus, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Buttons,
  Hctrls, ExtCtrls, Mask, Hcompte, ComCtrls, UtilEDT, CritEDT, Ent1, HEnt1,
  CpteUtil, QRRupt, HTB97, HPanel, UiUtil, tCalcCum ;

procedure BrouillardAna(Code : char) ;

type
  TFQrBrouilAna = class(TFQR)
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
    Y_NumLV: TQRLabel;
    S_SECTION: TQRDBText;
    S_LIBELLE: TQRDBText;
    Y_LIBELLE: TQRDBText;
    Y_REFINTERNE: TQRDBText;
    Y_DEBIT: TQRDBText;
    CREDIT: TQRDBText;
    QRLabel8: TQRLabel;
    RTri: TQRLabel;
    QRLabel10: TQRLabel;
    RRefInterne: TQRLabel;
    QRLabel13: TQRLabel;
    RNumPiece1: TQRLabel;
    QRLabel14: TQRLabel;
    RNumPiece2: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel23: TQRLabel;
    RRefExt: TQRLabel;
    RAffaire: TQRLabel;
    QRLabel21: TQRLabel;
    QRLabel26: TQRLabel;
    RRefLib: TQRLabel;
    RDateRefExt: TQRLabel;
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
    BTotAxe: TQRBand;
    QRDLAXE: TQRDetailLink;
    QRSUBTOT: TQRBand;
    QRLabel2: TQRLabel;
    TotDebitGen2: TQRLabel;
    TotCreditGen2: TQRLabel;
    CodAxe: TQRLabel;
    LibAxe: TQRLabel;
    TotDebitAxe: TQRLabel;
    TotCreditAxe: TQRLabel;
    FTri: TCheckBox;
    Y_LIBELLE2: TQRDBText;
    Y_REFINTERNE2: TQRDBText;
    Y_ETABLISSEMENT: TQRLabel;
    FRefExt: TCheckBox;
    FRefLib: TCheckBox;
    FAffaire: TCheckBox;
    FDateRefExt: TCheckBox;
    procedure QAfterOpen(DataSet: TDataSet);
    procedure BEntetePieceBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BPiedPieceBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FormShow(Sender: TObject);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BEntetePieceAfterPrint(BandPrinted: Boolean);
    procedure QRDLAXENeedData(var MoreData: Boolean);
    procedure QRSUBTOTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BPiedPieceAfterPrint(BandPrinted: Boolean);
  private
    { Déclarations privées }
    HighBand                      : Boolean ;
    LAxe                                    : TStringList ;
    CodeNatureEcr                           : String ;
    QR1Y_EXERCICE, QR1Y_AFFAIRE,
    QR1Y_JOURNAL, QR1Y_REFEXTERNE,
    QR1Y_NATUREPIECE, QR1Y_ETABLISSEMENT    : TStringField ;
    QR1Y_DATECOMPTABLE, QR1Y_DATEREFEXTERNE : TDateTimeField ;
    QR1Y_NUMEROPIECE, QR1Y_NUMLIGNE,
    QR1Y_NUMVENTIL      : TIntegerField ;
    QR1Y_GENERAL        : TStringField ;
    QR1Y_AXE            : TStringField ;
    QR1Y_QUALIFQTE1     : TStringField ;
    QR1Y_QUALIFQTE2     : TStringField ;
    QR1Y_REFINTERNE     : TStringField ;
    QR1Y_LIBELLE        : TStringField ;
    QR1Y_SECTION        : TStringField ;
    QR1Y_POURCENTQTE1   : TFloatField ;
    QR1Y_POURCENTQTE2   : TFloatField ;
    QR1Y_POURCENTAGE    : TFloatField ;
    QR1Y_QTE1           : TFloatField ;
    QR1Y_QTE2           : TFloatField ;
    QR1S_LIBELLE        : TStringField ;
    QR1DEBIT            : TFloatField ;
    QR1CREDIT           : TFloatField ;
    StReportPiece       : String ;
    TotGen, TotPiece    : TabTot ;
    Procedure BrouilZoom(Quoi : String) ;
    Function  QuoiMvt : String ;
    Function  LibelTypeEcr(CodeEcr : String) : String ;
    Function CtrlDateCreat(D1,D2 :TDateTime) : Boolean ;
    Function  RetourneCouple(A, B : String) : String ;
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

procedure BrouillardAna(Code : char) ;
var QR : TFQrBrouilAna ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQrBrouilAna.Create(Application) ;
QR.CodeNatureEcr:=Code ;
Edition.Etat:=etBrouAna ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BrouilZoom ;
QR.InitType(nbJalAna,neGL,msRien,'QRBROANA','',TRUE,FALSE,FALSE,Edition) ;
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

Procedure TFQrBrouilAna.BrouilZoom(Quoi : String) ;
Var Lp,i : Integer ;
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
(*Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
ZoomEdt(5,Quoi) ;*)
END ;

Function TFQrBrouilAna.QuoiMvt : String ;
BEGIN
Result:=QR1Y_SECTION.AsString+'   '+QR1S_LIBELLE.AsString+
        '#'+Qr1Y_JOURNAL.AsString+' N° '+IntToStr(Qr1Y_NUMEROPIECE.AsInteger)+' '+DateToStr(Qr1Y_DateComptable.AsDAteTime)+'-'+
        PrintSolde(Qr1DEBIT.AsFloat,Qr1Credit.AsFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)+
        '@'+'7;'+Qr1Y_JOURNAL.AsString+';'+UsDateTime(Qr1Y_DATECOMPTABLE.AsDateTime)+';'+Qr1Y_NUMEROPIECE.AsString+';'+Qr1Y_EXERCICE.asString+';'+
        IntToStr(Qr1Y_NumLigne.AsInteger)+';'+QR1Y_AXE.AsString+';' ;
END ;

procedure TFQrBrouilAna.ChoixEdition ;
BEGIN
Inherited ;
ChargeRecap(LAxe) ;
END ;

procedure TFQrBrouilAna.FinirPrint;
BEGIN
Inherited ;
VideRecap(LAxe) ;
END ;

function  TFQrBrouilAna.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK and CtrlDateCreat(CritEDT.BrouAna.DateCreat1,CritEDT.BrouAna.DateCreat2) ;
if Result then
   BEGIN
   Fillchar(TotGen, SizeOf(TotGen),#0) ;
   END ;
END ;

Function TFQrBrouilAna.CtrlDateCreat(D1,D2 : TDateTime) : Boolean ;
BEGIN
//Result:=true ;
Result:=(IsValidDate(DateToStr(D1)) and IsValidDate(DateToStr(D2))and(D1 <= D2) ) ;
If Not Result then NumErreurCrit:=7 ;
END ;

procedure TFQrBrouilAna.RecupCritEDT ;
BEGIN
Inherited ;
With CritEDT Do
  BEGIN
  BrouAna.RefInterne:=FRefInterne.Text       ;
  If FNumPiece1.Text<>'' then BrouAna.NumPiece1:=StrToInt(FNumPiece1.Text) else BrouAna.NumPiece1:=0 ;
  If FNumPiece2.Text<>'' then BrouAna.NumPiece2:=StrToInt(FNumPiece2.Text) else BrouAna.NumPiece2:=999999999 ;
  BrouAna.Tri:=FTri.Checked ; if FNatureCpt.Value<>'' then BrouAna.Axe:=FNatureCpt.Value ;
  With BrouAna.FormatPrint Do
       BEGIN
       { En attendant un neBro }
       PrSepCompte[1]:=FLigneCpt.Checked ;
       PrSepCompte[2]:=FLignePiecePied.Checked ;
       END ;
  END ;
END ;

procedure TFQrBrouilAna.GenereSQL ;
BEGIN
InHerited ;
Q.Close ; Q.SQL.Clear ;
Q.SQL.Add('Select  Y_EXERCICE, Y_JOURNAL, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE,  ') ;
Q.SQL.Add(       ' Y_GENERAL,Y_AXE, Y_QUALIFQTE2 , Y_REFINTERNE, Y_LIBELLE, Y_SECTION, ') ;
Q.SQL.Add(       ' Y_POURCENTAGE,Y_POURCENTQTE1,Y_POURCENTQTE2, ') ;
Q.SQL.Add(       ' Y_QTE1, Y_QTE2, Y_QUALIFQTE1, S_LIBELLE, Y_NUMVENTIL, Y_ETABLISSEMENT, ') ;
Q.SQL.Add(       ' Y_AFFAIRE, Y_REFEXTERNE, Y_DATEREFEXTERNE, Y_NATUREPIECE, ') ;
Case CritEdt.Monnaie of
  0 : BEGIN Q.SQL.Add(' Y_DEBIT DEBIT, Y_CREDIT CREDIT ')         ; END ;
  1 : BEGIN Q.SQL.Add(' Y_DEBITDEV DEBIT, Y_CREDITDEV CREDIT ')   ; END ;
 end ;
 { Tables explorées par la SQL }
Q.SQL.Add(' From ANALYTIQ ') ;
Q.SQL.Add(       ' LEFT JOIN SECTION on S_AXE=Y_AXE and S_SECTION=Y_SECTION ') ;
{ Construction de la clause Where de la SQL }
Q.SQL.Add(' Where Y_TYPEANALYTIQUE="X" ') ;
Q.SQL.Add(' And Y_DATECOMPTABLE>="'+USDateTime(CritEdt.DateDeb)+'" And Y_DATECOMPTABLE<="'+USDateTime(CritEdt.DateFin)+'" ') ;
Q.SQL.Add(' And Y_QUALIFPIECE="'+CodeNatureEcr+'" ') ;
if CritEdt.BrouAna.Axe<>'' then Q.SQL.Add(' And Y_AXE="'+CritEdt.BrouAna.Axe+'" ') ;
if CritEdt.Joker then Q.SQL.Add(' AND Y_JOURNAL like "'+TraduitJoker(CritEdt.Cpt1)+'" ') Else
   BEGIN
   if CritEdt.Cpt1<>'' then Q.SQL.Add(' AND Y_JOURNAL>="'+CritEdt.Cpt1+'" ') ;
   if CritEdt.Cpt2<>'' then Q.SQL.Add(' AND Y_JOURNAL<="'+CritEdt.Cpt2+'" ') ;
   END ;
Q.SQL.Add(' And Y_NUMEROPIECE>='+IntToStr(CritEdt.BrouAna.NumPiece1)+' ') ;
Q.SQL.Add(' And Y_NUMEROPIECE<='+IntToStr(CritEdt.BrouAna.NumPiece2)+' ') ;
if CritEdt.Exo.Code<>'' then Q.SQL.Add(' And Y_EXERCICE="'+CritEdt.Exo.Code+'" ') ;
if CritEdt.BrouAna.RefInterne<>'' then Q.SQL.Add(' And UPPER(Y_REFINTERNE) like "'+TraduitJoker(CritEdt.BrouAna.RefInterne)+'" ' );
if CritEdt.Etab<>'' then Q.SQL.Add(' And Y_ETABLISSEMENT="'+CritEdt.Etab+'" ') ;
if (OkV2 and (V_PGI.Confidentiel<>'1')) then Q.SQL.Add('AND Y_CONFIDENTIEL<>"1" ') ;
{ Construction de la clause Order By de la SQL }
if CritEdt.BrouAna.Tri then Q.SQL.Add(' Order By Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_AXE, Y_NUMVENTIL ')
                       Else Q.SQL.Add(' Order By Y_NUMEROPIECE, Y_AXE, Y_NUMVENTIL  ') ;
ChangeSQL(Q) ; Q.Open ;
END ;

procedure TFQrBrouilAna.InitDivers ;
BEGIN
InHerited ;
BEntetePiece.Frame.DrawTop:=CritEdt.BrouAna.FormatPrint.PrSepCompte[1] ;
BEntetePiece.Frame.DrawBottom:=CritEdt.BrouAna.FormatPrint.PrSepCompte[1] ;
BDetail.Frame.DrawTop:=False ;
BDetail.Frame.DrawBottom:=False ;
BPiedPiece.Frame.DrawTop:=CritEdt.BrouAna.FormatPrint.PrSepCompte[2] ;
BPiedPiece.Frame.DrawBottom:=CritEdt.BrouAna.FormatPrint.PrSepCompte[2] ;
If (Not FRefExt.Checked) and (Not FDateRefExt.Checked) then
   BEGIN
   TY_REFINTERNE.Caption:=MsgBox.Mess[12]  ;
   Y_REFINTERNE.DataField:='Y_REFINTERNE' ;
   END Else
If (FRefExt.Checked) and (Not FDateRefExt.Checked) then
   BEGIN
   TY_REFINTERNE.Caption:=MsgBox.Mess[15] ;
   Y_REFINTERNE.DataField:='Y_REFEXTERNE' ;
   END Else
If (Not FRefExt.Checked) and (FDateRefExt.Checked) then
   BEGIN
   TY_REFINTERNE.Caption:=MsgBox.Mess[14] ;
   Y_REFINTERNE.DataField:='Y_DATEREFEXTERNE' ;
   END Else
If (FRefExt.Checked) and (FDateRefExt.Checked) then
   BEGIN
   TY_REFINTERNE.Caption:=MsgBox.Mess[16] ;
   Y_REFINTERNE.DataField:='' ;
   END ;


If (Not FRefLib.Checked) and (Not FAffaire.Checked) then
   BEGIN
   TY_LIBELLE.Caption:=MsgBox.Mess[13] ;
   Y_LIBELLE.DataField:='Y_LIBELLE' ;
   END Else
If (FRefLib.Checked) and (Not FAffaire.Checked) then
   BEGIN
   TY_LIBELLE.Caption:=MsgBox.Mess[17] ;
   Y_LIBELLE.DataField:='' ; //Y_LIBELLE.DataField:='Y_REFLIBRE' ;
   END Else
If (Not FRefLib.Checked) and (FAffaire.Checked) then
   BEGIN
   TY_LIBELLE.Caption:=MsgBox.Mess[18] ;
   Y_LIBELLE.DataField:='Y_AFFAIRE' ;
   END Else
If (FRefLib.Checked) and (FAffaire.Checked) then
   BEGIN
   TY_LIBELLE.Caption:=MsgBox.Mess[19] ;
   Y_LIBELLE.DataField:='' ;
   END ;

HighBand:=((FRefExt.Checked)or(FRefLib.Checked)or(FAffaire.Checked)or(FDateRefExt.Checked)) ;
if HighBand then
   BEGIN
   BEntetePiece.Height:=(TY_NUMEROPIECE.Top+TY_NUMEROPIECE.Height)+Y_REFINTERNE2.Height+3 ;
   END else
   BEGIN
   BEntetePiece.Height:=(TY_NUMEROPIECE.Top+TY_NUMEROPIECE.Height)+3 ;
   END ;
END ;

procedure TFQrBrouilAna.RenseigneCritere ;
{ Récupération des champs du multicritère dans l'entête d'état }
BEGIN
Inherited ;
RNatureCpt.Caption:=FNatureCpt.Text ;
RRefInterne.Caption:=FRefInterne.Text ;
RNumPiece1.Caption:=IntToStr(CritEdt.BrouAna.NumPiece1) ; RNumPiece2.Caption:=IntToStr(CritEdt.BrouAna.NumPiece2) ;
RNatureEcr.Caption:=LibelTypeEcr(CodeNatureEcr) ;
CaseACocher(FTri,RTri) ;
CaseACocher(FRefExt,RRefExt) ; CaseACocher(FRefLib,RRefLib) ;
CaseACocher(FAffaire,RAffaire) ; CaseACocher(FDateRefExt,RDateRefExt) ;
END ;

Function TFQrBrouilAna.LibelTypeEcr(CodeEcr : String) : String ;
BEGIN
Result:='' ;
if CodeEcr='N' then Result:=MsgBox.Mess[6] else
if CodeEcr='P' then Result:=MsgBox.Mess[7] else
if CodeEcr='R' then Result:=MsgBox.Mess[8] else
if CodeEcr='S' then Result:=MsgBox.Mess[9] else
if CodeEcr='U' then Result:=MsgBox.Mess[10] ;
END;

procedure TFQrBrouilAna.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
QR1Y_ETABLISSEMENT  :=TStringField(Q.FindField('Y_ETABLISSEMENT')) ;
QR1Y_EXERCICE       :=TStringField(Q.FindField('Y_EXERCICE')) ;
QR1Y_JOURNAL        :=TStringField(Q.FindField('Y_JOURNAL')) ;
QR1Y_DATECOMPTABLE  :=TDateTimeField(Q.FindField('Y_DATECOMPTABLE')) ;
QR1Y_NUMEROPIECE    :=TIntegerField(Q.FindField('Y_NUMEROPIECE')) ;
QR1Y_NUMLIGNE       :=TIntegerField(Q.FindField('Y_NUMLIGNE')) ;
QR1Y_NUMVENTIL      :=TIntegerField(Q.FindField('Y_NUMVENTIL')) ;
QR1Y_GENERAL        :=TStringField(Q.FindField('Y_GENERAL')) ;
QR1Y_AXE            :=TStringField(Q.FindField('Y_AXE')) ;
QR1Y_QUALIFQTE1     :=TStringField(Q.FindField('Y_QUALIFQTE1')) ;
QR1Y_QUALIFQTE2     :=TStringField(Q.FindField('Y_QUALIFQTE2')) ;
QR1Y_REFINTERNE     :=TStringField(Q.FindField('Y_REFINTERNE')) ;
QR1Y_LIBELLE        :=TStringField(Q.FindField('Y_LIBELLE')) ;
QR1Y_SECTION        :=TStringField(Q.FindField('Y_SECTION')) ;
QR1Y_POURCENTQTE1   :=TFloatField(Q.FindField('Y_POURCENTQTE1')) ;
QR1Y_POURCENTQTE2   :=TFloatField(Q.FindField('Y_POURCENTQTE2')) ;
QR1Y_POURCENTAGE    :=TFloatField(Q.FindField('Y_POURCENTAGE')) ;
QR1Y_QTE1           :=TFloatField(Q.FindField('Y_QTE1')) ;
QR1Y_QTE2           :=TFloatField(Q.FindField('Y_QTE2')) ;
QR1S_LIBELLE         :=TStringField(Q.FindField('S_LIBELLE')) ;
QR1Y_AFFAIRE        :=TStringField(Q.FindField('Y_AFFAIRE')) ;
QR1Y_REFEXTERNE     :=TStringField(Q.FindField('Y_REFEXTERNE')) ;
QR1Y_DATEREFEXTERNE :=TDateTimeField(Q.FindField('Y_DATEREFEXTERNE')) ;
QR1Y_NATUREPIECE    :=TStringField(Q.FindField('Y_NATUREPIECE')) ;
QR1DEBIT            :=TFloatField(Q.FindField('DEBIT')) ;
QR1CREDIT           :=TFloatField(Q.FindField('CREDIT')) ;
ChgMaskChamp(QR1DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
ChgMaskChamp(QR1CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
end;

procedure TFQrBrouilAna.BEntetePieceBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var St1, St2 : String ;
begin
  inherited;
St1:='' ; St2:='' ;
TY_JOURNAL.Left:=TS_LIBELLE.Left ;
TY_JOURNAL.Width:=(TS_LIBELLE.Width+TY_REFINTERNE.Width)+1 ;
TY_NUMEROPIECE.Left:=TCptGen.Left ;
TY_NUMEROPIECE.Width:=TCptGen.Width ;
TY_NUMEROPIECE.Caption:=MsgBox.Mess[2]+' '+IntToStr(QR1Y_NUMEROPIECE.AsInteger) ;
St1:=DateToStr(QR1Y_DATECOMPTABLE.AsDateTime)+'  '+QR1Y_AXE.AsString+'  '+QR1Y_GENERAL.AsString ;
St2:=QR1Y_JOURNAL.AsString+'   '+RechDom('ttJalAnalytique',QR1Y_JOURNAL.AsString,False)  ;
TY_JOURNAL.Caption:=St1+'  '+St2 ;
Y_ETABLISSEMENT.Caption:=RechDom('ttEtablissement',QR1Y_ETABLISSEMENT.AsString,False) ;
Y_REFINTERNE2.Caption:=MsgBox.Mess[12]+' : '+QR1Y_REFINTERNE.AsString ;
Y_LIBELLE2.Caption:=MsgBox.Mess[13]+' : '+QR1Y_LIBELLE.AsString ;
Y_REFINTERNE2.Visible:=HighBand ;  Y_LIBELLE2.Visible:=HighBand ;
StReportPiece:=QR1Y_NUMEROPIECE.AsString ;
InitReport([2],CritEdt.BrouAna.FormatPrint.Report) ;
end;

procedure TFQrBrouilAna.BPiedPieceBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
TY_NUMEROPIECE_.Caption:=MsgBox.Mess[3]+' '+IntToStr(QR1Y_NUMEROPIECE.AsInteger) ;
TotDebitPiece.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotPiece[1].TotDebit,  CritEdt.AfficheSymbole) ;
TotCreditPiece.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotPiece[1].TotCredit, CritEdt.AfficheSymbole) ;
end;

procedure TFQrBrouilAna.BFinEtatBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
TotDebitGen.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotGen[1].TotDebit,  CritEdt.AfficheSymbole) ;
TotCreditGen.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotGen[1].TotCredit, CritEdt.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFQrBrouilAna.BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
var MReport : TabTRep ;
begin
  inherited;
Case QuelReportBAL(CritEdt.BrouAna.FormatPrint.Report,MReport) of
  1 : BEGIN TITREREPORTB.Caption:=MsgBox.Mess[23] END ;
  2 : BEGIN TITREREPORTB.Caption:=MsgBox.Mess[24]+' '+StReportPiece ; END ;
  END ;
Report2Debit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[1].TotDebit,  CritEdt.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[1].TotCredit, CritEdt.AfficheSymbole ) ;
end;

procedure TFQrBrouilAna.BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
var MReport : TabTRep ;
    St      : String ;

begin
  inherited;
St:='' ; Fillchar(MReport,SizeOf(MReport),#0) ;
If FRefExt.Checked and FDateRefExt.Checked then
   BEGIN
   If (QR1Y_DATEREFEXTERNE.AsDateTime<>IDate1900) then St:=DateToStr(QR1Y_DATEREFEXTERNE.AsDateTime)
                                                                Else St:='' ;
   Y_REFINTERNE.Caption:=RetourneCouple(St, QR1Y_REFEXTERNE.AsString) ;
   END ;
If FRefLib.Checked and FAffaire.Checked then
   BEGIN
   //Y_LIBELLE.Caption:=RetourneCouple(QR1Y_REFLIBRE.AsString, QR1Y_AFFAIRE.AsString) ;
   Y_LIBELLE.Caption:=RetourneCouple('', QR1Y_AFFAIRE.AsString) ;
   END ;
Y_NumLV.Caption:=IntToStr(QR1Y_NUMVENTIL.AsInteger) ;
MReport[1].TotDebit:=QR1DEBIT.AsFloat ;
MReport[1].TotCredit:=QR1CREDIT.AsFloat ;
TotGen[1].TotDebit:= Arrondi(TotGen[1].TotDebit+QR1DEBIT.AsFloat,CritEdt.Decimale) ;
TotGen[1].TotCredit:=Arrondi(TotGen[1].TotCredit+QR1CREDIT.AsFloat,CritEdt.Decimale) ;
TotPiece[1].TotDebit:= Arrondi(TotPiece[1].TotDebit+QR1DEBIT.AsFloat,CritEdt.Decimale) ;
TotPiece[1].TotCredit:=Arrondi(TotPiece[1].TotCredit+QR1CREDIT.AsFloat,CritEdt.Decimale) ;
AddRecap(LAxe, [QR1Y_AXE.AsString], [QR1Y_AXE.AsString],[1,QR1DEBIT.AsFloat, QR1CREDIT.AsFloat]) ;
AddReportBAL([1,2], CritEdt.BrouAna.FormatPrint.Report, MReport, CritEdt.Decimale) ;
if PrintBand then Quoi:=QuoiMvt ;
end;

procedure TFQrBrouilAna.FormShow(Sender: TObject);
begin
  inherited;
  // YMO 15/11/2005 FQ 16951
  Avecrevision.visible := false;
  //27/10/2005 Rajout de l'aide
  HelpContext:=7385000 ;
{$IFDEF CCS3}
FNatureCpt.Visible:=FALSE ; HLabel2.Visible:=FALSE ;
{$ENDIF}
TabSup.TabVisible:=False;
end;

procedure TFQrBrouilAna.TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
end;

procedure TFQrBrouilAna.BEntetePieceAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Fillchar(TotPiece, SizeOf(TotPiece),#0) ;

end;

procedure TFQrBrouilAna.QRDLAXENeedData(var MoreData: Boolean);
Var Tot : Array[0..12] of Double ;
    Cod,Lib : String ;
begin
  inherited;
if CritEdt.NatureCpt<>'' then BEGIN MoreData:=FALSE ; exit ; END ;
MoreData:=PrintRecap(LAxe,Cod,Lib,Tot) ;
If MoreData Then
   BEGIN
   CodAxe.Caption:=MsgBox.Mess[21]+' '+Cod ;
   LibAxe.Caption:=FloatToStr(Tot[0])+' '+MsgBox.Mess[25] ;
   TotDebitAxe.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, Tot[1],  CritEdt.AfficheSymbole ) ;
   TotCreditAxe.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, Tot[2],  CritEdt.AfficheSymbole ) ;
   END ;
end;

procedure TFQrBrouilAna.QRSUBTOTBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
if CritEdt.NatureCpt<>'' then BEGIN PrintBand:=FALSE ; exit ; END ;
TotDebitGen2.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotGen[1].TotDebit,  CritEdt.AfficheSymbole) ;
TotCreditGen2.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotGen[1].TotCredit, CritEdt.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
TOPREPORT.enabled:=FALSE ;
end;

procedure TFQrBrouilAna.BPiedPieceAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([2],CritEdt.BrouAna.FormatPrint.Report) ;
end;

Function TFQrBrouilAna.RetourneCouple(A, B : String) : String ;
BEGIN
Result:='' ;
if (A<>'')and(B<>'') then Result:=A+' / '+B else
if (A<>'')and(B='') then Result:=A else
if (A='')and(B<>'') then Result:=B else Exit ;
END ;

end.
