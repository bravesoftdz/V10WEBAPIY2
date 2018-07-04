unit QRJDivis;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  StdCtrls, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, UtilEDT, CritEDT, QRRupt, CpteUtil,
  EdtLegal,
  Ent1, HEnt1, HQry, HSysMenu, Menus, Calcole, HTB97, HPanel, UiUtil, tCalcCum,
  ADODB ;

procedure JalDivisioAno ;
procedure JalDivisioClo ;
procedure JalDivisio ;
procedure JalDivisioZoom(Crit : TCritEdt) ;
procedure JalDivisioLegal(Crit : TCritEdt) ;
// GC
procedure JalDivisioPCL;
procedure JalDivisioPCLChaine( vCritEdtChaine : TCritEdtChaine );
//

Type TNatEtat = (neRien,neAno,neClo) ;

type
  TQRJalDiv = class(TFQR)
    HLabel3: THLabel;
    FRefInterne: TEdit;
    HLabel9: THLabel;
    FNumPiece1: TMaskEdit;
    Label12: TLabel;
    FNumPiece2: TMaskEdit;
    FValide: TCheckBox;
    FLignePieEntete: TCheckBox;
    FLigneJalPied: TCheckBox;
    FLignePiePied: TCheckBox;
    FSautDePage: TCheckBox;
    QRLabel8: TQRLabel;
    RTri: TQRLabel;
    QRLabel10: TQRLabel;
    RRefInterne: TQRLabel;
    QRLabel13: TQRLabel;
    RNumPiece1: TQRLabel;
    QRLabel14: TQRLabel;
    RNumPiece2: TQRLabel;
    QRLabel25: TQRLabel;
    RValide: TQRLabel;
    TE_VALIDE: TQRLabel;
    TLE_NUMLIGNE: TQRLabel;
    TLeCompte: TQRLabel;
    TLibCompte: TQRLabel;
    TLE_REFINTERNE: TQRLabel;
    TLE_MODEPAIE: TQRLabel;
    TLE_DATEECHEANCE: TQRLabel;
    TLE_LIBELLE: TQRLabel;
    TLE_DEBIT: TQRLabel;
    TLE_CREDIT: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    BHPiece: TQRBand;
    TE_NUMEROPIECE: TQRLabel;
    TE_DEVISE: TQRLabel;
    BEcrDetail: TQRBand;
    LE_DEVISE: TQRLabel;
    LE_CREDIT: TQRDBText;
    LE_DEBIT: TQRDBText;
    LE_VALIDE: TQRLabel;
    BMULTI: TQRBand;
    TTOTMULTI: TQRLabel;
    TDEBMULTI: TQRLabel;
    TCREMULTI: TQRLabel;
    BFJournal: TQRBand;
    TJ_JOURNAL_: TQRLabel;
    TotDebJal: TQRLabel;
    TotCreJal: TQRLabel;
    QRLabel33: TQRLabel;
    TotGeneDeb: TQRLabel;
    TotGeneCre: TQRLabel;
    QRBand1: TQRBand;
    Trait4: TQRLigne;
    Trait3: TQRLigne;
    Trait2: TQRLigne;
    Trait0: TQRLigne;
    Ligne1: TQRLigne;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    QJ_JOURNAL: TStringField;
    QJ_LIBELLE: TStringField;
    QJ_ABREGE: TStringField;
    QJ_NATUREJAL: TStringField;
    QEcr: TQuery;
    SEcr: TDataSource;
    MsgLibel: THMsgBox;
    DLMULTI: TQRDetailLink;
    DLPieceH: TQRDetailLink;
    DLMVT: TQRDetailLink;
    LECOMPTE: TQRLabel;
    QRLigne1: TQRLigne;
    LE_NUMLIGNE: TQRLabel;
    LE_REFINTERNE: TQRLabel;
    LE_DATEECHEANCE: TQRLabel;
    LE_LIBELLE: TQRLabel;
    LE_MODEPAIE: TQRDBText;
    TJ_JOURNAL: TQRLabel;
    LibCompte: TQRLabel;
    TE_ETABLISSEMENT: TQRLabel;
    TE_TAUXDEV: TQRLabel;
    TRLegende: TQRLabel;
    RLegende: TQRLabel;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel6: TQRLabel;
    TJ_NATUREJAL: TQRLabel;
    FTri: TCheckBox;
    FAvecJalAN: TCheckBox;
    QRLabel9: TQRLabel;
    RAN: TQRLabel;
    FVoirLibelleCompte: TCheckBox;
    TLE_DEVISE: TQRLabel;
    FNouvFct: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure BEcrDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFJournalBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFJournalAfterPrint(BandPrinted: Boolean);
    procedure BHPieceBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure DLMULTINeedData(var MoreData: Boolean);
    procedure QEcrAfterOpen(DataSet: TDataSet);
    procedure DLPieceHNeedData(var MoreData: Boolean);
    procedure BDetailAfterPrint(BandPrinted: Boolean);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure QEcrBeforeOpen(DataSet: TDataSet);
    procedure BMULTIAfterPrint(BandPrinted: Boolean);
    procedure BMULTIBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure Timer1Timer(Sender: TObject);
    procedure GenereSQL ; Override ;
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
  private    { Déclarations privées }
    NET                             : TNatEtat ;
    NumR, AncienEche                : Integer ;
    LMulti                          : TStringList ;
    TotEdt, TotJour                 : TabTot ;
    StReportJal, StReportPie        : String ;
    OKEnteteP                       : Boolean ;
    QR1J_JOURNAL, QR1J_LIBELLE,
    QR1J_ABREGE, QR1J_NATUREJAL     : TStringField ;
    QR2E_JOURNAL                    : TStringField ;
    QR2E_GENERAL                    : TStringField ;
    QR2E_AUXILIAIRE                 : TStringField ;
    QR2E_EXERCICE                   : TStringField ;
    QR2E_DATECOMPTABLE              : TDateTimeField ;
    QR2DEBIT, QR2CREDIT, QR2E_TAUXDEV,Qr2E_COTATION : TFloatField ;
    QR2E_NUMEROPIECE, QR2MULTIECHE,
    QR2E_NUMLIGNE, QR2E_NUMECHE     : TIntegerField ;
    QR2E_REFINTERNE                 : TStringField ;
    QR2E_LIBELLE                    : TStringField ;
    QR2E_NATUREPIECE                : TStringField ;
    QR2E_QUALIFPIECE                : TStringField ;
    QR2E_MODEPAIE                   : TStringField ;
    // GC
    QR2E_MODESAISIE                 : TStringField ;
    //
    QR2E_ECHE, QR2E_ETABLISSEMENT   : TStringField ;
    QR2E_DATEECHEANCE, QR2E_DATETAUXDEV : TDateTimeField ;
    QR2E_DEVISE,
    QR2E_VALIDE, QR2E_ETATLETTRAGE,
    QR2T_LIBELLE, QR2G_LIBELLE          : TStringField ;

    IsLegal                         : Boolean ;
    
    procedure GenereSQLSub ;
    Procedure JournalZoom(Quoi : String) ;
    Function  QuoiMvt : String ;
    Function  QuoiJal(i : Integer) : String ;
  public     { Déclarations publiques }

  end;

implementation

{$R *.DFM}

Uses SaisUtil,UentCommun ;

procedure JalDivisio ;
var QR : TQRJalDiv ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TQRJalDiv.Create(Application) ;
Edition.Etat:=etJalDiv ;
QR.QRP.QRPrinter.OnSynZoom:=QR.JournalZoom ;
QR.IsLegal:=FALSE ;
QR.NET:=neRien ;
QR.InitType(nbJal,neJal,msRien,'QRJDIVIS','',TRUE,FALSE,FALSE,Edition) ;
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

// GC - 05/11/2001
procedure JalDivisioPCL;
var QR : TQRJalDiv ;
    Edition : TEdition ;
    PP : THPanel ;
begin

  PP:=FindInsidePanel ;
  QR:=TQRJalDiv.Create(Application) ;
  Edition.Etat:=etJalDivPCL ;

  QR.QRP.QRPrinter.OnSynZoom:=QR.JournalZoom ;
  QR.IsLegal:=FALSE ;
  QR.NET:=neRien ;
  QR.InitType(nbJal,neJal,msRien,'QRJDIVIS','',TRUE,FALSE,FALSE,Edition) ;

  QR.Caption := TraduireMemoire('Journal des écritures');
  UpdateCaption(QR);

  if PP=Nil then
  begin
    try
      QR.ShowModal ;
    finally
      QR.Free ;
    end ;
    Screen.Cursor:=SyncrDefault ;
  end
  else
  begin
    InitInside(QR,PP) ;
    QR.Show ;
  end ;

end;

// GC - 12/12/2001
procedure JalDivisioPCLChaine( vCritEdtChaine : TCritEdtChaine );
var QR : TQRJalDiv ;
    Edition : TEdition ;
    PP : THPanel ;
begin

  PP:=FindInsidePanel ;
  QR:=TQRJalDiv.Create(Application) ;
  Edition.Etat:=etJalDivPCL ;

  QR.QRP.QRPrinter.OnSynZoom:=QR.JournalZoom ;
  QR.IsLegal:=FALSE ;
  QR.NET:=neRien ;
  QR.CritEdtChaine := vCritEdtChaine;

  QR.InitType(nbJal,neJal,msRien,'QRJDIVIS','',TRUE,FALSE,FALSE,Edition) ;

  if CtxPcl in V_Pgi.PgiContexte then
  begin
    QR.Caption := TraduireMemoire('Journal des écritures');
    UpdateCaption(QR);
  end;  

  if PP=Nil then
  begin
    try
      QR.Timer1.Enabled := True;
      // Penser au chargement du filtre passé en paramètre
      QR.QRP.QRPrinter.Copies := QR.CritEdtChaine.NombreExemplaire;
      QR.ShowModal ;
    finally
      QR.Free ;
    end ;
    Screen.Cursor:=SyncrDefault ;
  end
  else
  begin
    InitInside(QR,PP) ;
    QR.Show ;
  end ;

end;
//
procedure JalDivisioZoom(Crit : TCritEdt) ;
var QR : TQRJalDiv ;
    Edition : TEdition ;
BEGIN
QR:=TQRJalDiv.Create(Application) ;
Edition.Etat:=etJalDiv ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.JournalZoom ;
 QR.IsLegal:=FALSE ;
 QR.CritEDT:=Crit ;
 QR.NET:=neRien ;
 QR.InitType(nbJal,neJal,msRien,'QRJDIVIS','',FALSE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure JalDivisioLegal(Crit : TCritEdt) ;
var QR : TQRJalDiv ;
    Edition : TEdition ;
BEGIN
QR:=TQRJalDiv.Create(Application) ;

// GC_GP - 04/01/2002
if TRUE Or (CtxPcl in V_PGI.PgiContexte) then
begin
  QR.Caption := TraduireMemoire('Journal des écritures');
  UpdateCaption(QR);
  Edition.Etat := etJalDivPCL;
  QR.FTri.Checked := True;
  QR.FTri.Enabled := False;
  QR.CritEdt.Jal.Tri := True;
  QR.FNouvFct.Checked := True;
  QR.FVoirLibelleCompte.Checked := False;
end
else
  Edition.Etat := etJalDiv ;
// FIN GC - 04/01/2002

try
 QR.QRP.QRPrinter.OnSynZoom:=QR.JournalZoom ;
 QR.IsLegal:=TRUE ;
 QR.CritEDT:=Crit ;
 QR.NET:=neRien ;

 // GC-GP - 04/01/2002
  if TRUE Or (CtxPcl in V_PGI.PgiContexte) then
  begin
  QR.CritEdt.Jal.Tri := True;
  end;
  // FIN GC - 04/01/2002

 QR.InitType(nbJal,neJal,msRien,'QRJDIVIS','',TRUE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure JalDivisioAno ;
var QR : TQRJalDiv ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TQRJalDiv.Create(Application) ;
Edition.Etat:=etJalAno ;
QR.QRP.QRPrinter.OnSynZoom:=QR.JournalZoom ;
QR.IsLegal:=FALSE ;
QR.NET:=neAno ;
QR.InitType(nbJal,neJal,msRien,'QRJALANO','',TRUE,FALSE,FALSE,Edition) ;
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

procedure JalDivisioClo ;
var QR : TQRJalDiv ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TQRJalDiv.Create(Application) ;
Edition.Etat:=etJalClo ;
QR.QRP.QRPrinter.OnSynZoom:=QR.JournalZoom ;
QR.IsLegal:=FALSE ;
QR.NET:=neClo ;
QR.InitType(nbJal,neJal,msRien,'QRJALCLO','',TRUE,FALSE,FALSE,Edition) ;
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

Procedure TQRJalDiv.JournalZoom(Quoi : String) ;
Var Lp,i : Integer ;
BEGIN
Inherited ;
Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
i:=StrToInt(Copy(Quoi,Lp+1,1)) ;
If (i=5) Then
   BEGIN
   Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
   If QRP.QrPrinter.FSynShiftDblClick Then i:=6 ;
   If QRP.QRPrinter.FSynCtrlDblClick Then i:=11 ;
   END ;
ZoomEdt(i,Quoi) ;
(*Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
ZoomEdt(5,Quoi) ;*)
END ;

Function TQRJalDiv.QuoiJal(i : Integer) : String ;
BEGIN
Inherited ;
Case i Of
  1 : Result:=Qr1J_JOURNAL.AsString+' '+'#'+Qr1J_LIBELLE.AsString+'@'+'4' ;
  (*2 : Result:=Qr1J_JOURNAL.AsString+' '+'#'+Qr1J_LIBELLE.AsString+' '+
              TotSoldecptAux.Caption+'@'+'4' ;*)
  END ;
END ;

Function TQRJalDiv.QuoiMvt : String ;
BEGIN
Inherited ;
Result:=QR2E_GENERAL.AsString+'   '+Qr2E_AUXILIAIRE.AsString+
        '#'+Qr2E_JOURNAL.AsString+' N° '+IntToStr(Qr2E_NUMEROPIECE.AsInteger)+' '+DateToStr(Qr2E_DateComptable.AsDAteTime)+'-'+
        PrintSolde(Qr2DEBIT.AsFloat,Qr2Credit.AsFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)+
        '@'+'5;'+Qr2E_JOURNAL.AsString+';'+UsDateTime(Qr2E_DATECOMPTABLE.AsDateTime)+';'+Qr2E_NUMEROPIECE.AsString+';'+Qr2E_EXERCICE.asString+';'+
        IntToStr(Qr2E_NumLigne.AsInteger)+';' ;
END ;

procedure TQRJalDiv.FinirPrint;
Var SolD, SolC : Double ;
BEGIN
Inherited ;
QEcr.Close ;
VideGroup(LMulti) ;
If OkMajEdt Then  { Montant Général pour les éditions légales }
   BEGIN
   If TotEdt[1].TotDebit>TotEdt[1].TotCredit then
      BEGIN SolD:=Abs(TotEdt[1].TotDebit-TotEdt[1].TotCredit) ; SolC:=0 ; END
      Else BEGIN SolD:=0 ; SolC:=Abs(TotEdt[1].TotDebit-TotEdt[1].TotCredit) ; END ;

   If IsLegal then MajEditionLegal('JLD',CritEdt.Exo.Code,DateToStr(CritEdt.DateDeb),DateToStr(CritEdt.DateFin),'',TotEdt[1].TotDebit,TotEdt[1].TotCredit,SolD,SolC)
              Else MajEdition('JLD',CritEdt.Exo.Code,DateToStr(CritEdt.DateDeb),DateToStr(CritEdt.DateFin),'',TotEdt[1].TotDebit,TotEdt[1].TotCredit,SolD,SolC) ;
   END ;
END ;

function  TQRJalDiv.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ;
If Result Then
   BEGIN
   END ;
END ;

procedure TQRJalDiv.RecupCritEdt ;
Var i : Integer ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  RefInterne:=FRefInterne.Text ;
  If FNumPiece1.Text<>'' then Jal.NumPiece1:=StrToInt(FNumPiece1.Text) else Jal.NumPiece1:=0 ;
  If FNumPiece2.Text<>'' then Jal.NumPiece2:=StrToInt(FNumPiece2.Text) else Jal.NumPiece2:=999999999 ;
  if FValide.State=cbGrayed then Valide:='g' ;
  if FValide.State=cbChecked then Valide:='X' ;
  if FValide.State=cbUnChecked then Valide:='-' ;

  Jal.Tri:=Ftri.Checked ;
  Jal.AvecJalAN:=FAvecJalAN.Checked ;
  Jal.AfficheTauxEuro:=True ;

  Jal.VoirLibelleCompte := FVoirLibelleCompte.Checked;
  Jal.AvecNouvFct := FNouvFct.Checked;

  With Jal.FormatPrint Do
       BEGIN
       PrSepCompte[2]:=FLignePieEntete.Checked ;
       PrSepCompte[3]:=FLigneJalPied.Checked ;
       PrSepCompte[4]:=FLignePiePied.Checked ;
       END ;

  Jal.VoirLibelleCompte := FVoirLibelleCompte.Checked ;
  For i:=1 To 11 Do
    BEGIN
    If (i=4) And (Not CritEdt.Jal.VoirLibelleCompte) Then
      BEGIN
      CritEdt.Jal.FormatPrint.TabColl[i].OkAff:=FALSE ;
      CritEdt.Jal.FormatPrint.TabColl[i].Tag:=-1;
      END Else
      BEGIN
      CritEdt.Jal.FormatPrint.TabColl[i].OkAff:=TRUE ;
      CritEdt.Jal.FormatPrint.TabColl[i].Tag:=PremTabColEdt[i].Tag ;
      END ;
    END ;
  // FIN GC - 04/01/2002
  END ;
END ;

procedure TQRJalDiv.InitDivers ;
//Var ll : Integer ;
//var i : integer;
BEGIN
Inherited ;
OKEnteteP:=True ; AncienEche:=1 ;
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
BDetail.ForceNewPage:=FSautDePage.Checked ;
BHPiece.Frame.DrawTop:=CritEdt.Jal.FormatPrint.PrSepCompte[2] ;
BHPiece.Frame.DrawBottom:=CritEdt.Jal.FormatPrint.PrSepCompte[2] ;
BFJournal.Frame.DrawTop:=CritEdt.Jal.FormatPrint.PrSepCompte[3] ;
If (NET=neClo) Or (NET=neAno) Then
   BEGIN
   LE_MODEPAIE.Visible:=FALSE ; LE_DATEECHEANCE.Visible:=FALSE ;
   TLE_MODEPAIE.Caption:='' ; TLE_DATEECHEANCE.Caption:='' ;
   BTitre.Height:=106 ;
   // 5690
   TRLegende.Visible := False;
   RLegende.Visible := False;
   QRLabel1.Visible := False;
   QRLabel2.Visible := False;
   QRLabel3.Visible := False;
   QRLabel6.Visible := False;
   END ;

// GC - 04/01/2002
if IsLegal then
  FVoirLibelleCompte.Checked := False;
// FIN GC - 04/01/2002

END ;

procedure TQRJalDiv.GenereSQL ;
Var St, StJalAuto : String ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
Q.Sql.Clear ;
St:='Select J_JOURNAL,J_LIBELLE,J_ABREGE,J_NATUREJAL ' ;
Q.SQL.Add(St) ;
{ Tables explorées par la SQL }
St:=' From JOURNAL ' ;
Q.SQL.Add(St) ;
Q.SQL.Add(' Where J_JOURNAL<>"'+w_w+'" ') ;
{ Construction de la clause Where de la SQL }
St:='' ;
If CritEdt.Joker then
   BEGIN Q.SQL.Add(' And J_JOURNAL like "'+TraduitJoker(CritEdt.Cpt1)+'" ') ;END
   Else
   BEGIN
   if CritEdt.Cpt1<>'' then Q.SQL.Add(' And J_JOURNAL>="'+CritEdt.Cpt1+'" ') ;
   if CritEdt.Cpt2<>'' then Q.SQL.Add(' And J_JOURNAL<="'+CritEdt.Cpt2+'" ') ;
   END ;
If St<>'' Then Q.SQL.Add(St) ;
//St:=' AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ODA" AND J_NATUREJAL<>"ANA" ' ;
Case NET of
  neRien : BEGIN
           St:=' AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ODA" AND J_NATUREJAL<>"ANA" ' ;
           If Not CritEdt.Jal.AvecJalAN Then St:=St+' AND J_NATUREJAL<>"ANO" ' ;
           END ;
  neAno  : St:=' AND J_NATUREJAL="ANO" ' ;
  neClo  : St:=' AND J_NATUREJAL="CLO" ' ;
  End ;
If St<>'' Then Q.SQL.Add(St) ;
if (OkV2 and (V_PGI.Confidentiel<>'1')) then
   BEGIN
   StJalAuto:=VH^.JalAutorises ;
   if StJalAuto<>'' then
      BEGIN St:=St+' And '+AnalyseCompte(StJalAuto,fbJal,False,False) ; Q.SQL.Add(St) ; END ;
   END ;
St:=' Order By J_JOURNAL' ;
Q.SQL.Add(St) ;
ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
Q.Open ;
GenereSQLSub ;
END;

procedure TQRJalDiv.GenereSQLSUB ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
QEcr.Close ;
QEcr.SQL.Clear ;
{ Construction de la clause Select de la SQL }
QEcr.SQL.Add('Select ') ;
QEcr.SQL.Add(       'E_EXERCICE,E_JOURNAL,E_DATECOMPTABLE,E_NUMEROPIECE, ') ;
Case CritEdt.Monnaie of
  0 : QEcr.SQL.Add(' E_DEBIT DEBIT, E_CREDIT CREDIT, ') ;
  1 : QEcr.SQL.Add(' E_DEBITDEV DEBIT, E_CREDITDEV CREDIT, ') ;
End ;
QEcr.SQL.Add(       'E_NUMLIGNE,E_GENERAL,E_AUXILIAIRE, E_DATETAUXDEV, ') ;
QEcr.SQL.Add(       'E_REFINTERNE,E_LIBELLE,E_NATUREPIECE,E_QUALIFPIECE, ') ;
QEcr.SQL.Add(       'E_MODEPAIE,E_ECHE,E_DATEECHEANCE,E_DEVISE, E_TAUXDEV, E_COTATION, ') ;
// GC
QEcr.SQL.Add(       'E_MODESAISIE, ');
//
QEcr.SQL.Add(       'E_QTE1,E_QTE2,E_QUALIFQTE1,E_QUALIFQTE2, E_VALIDE, E_NUMECHE, ') ;
// GP : bug edition avec filtre sur N° pièce trop grand (à cause du *10)
//QEcr.SQL.Add(       'E_ETATLETTRAGE, E_NUMEROPIECE*10+E_NUMLIGNE as MULTIECHE, ') ;
QEcr.SQL.Add(       'E_ETATLETTRAGE, E_NUMLIGNE as MULTIECHE, ') ;
QEcr.SQL.Add(       'E_ETABLISSEMENT, G_LIBELLE, T_LIBELLE ') ;
QEcr.SQL.Add(' From ECRITURE') ;
QEcr.SQL.Add(' LEFT JOIN TIERS On T_AUXILIAIRE=E_AUXILIAIRE ') ;
QEcr.SQL.Add(' LEFT JOIN GENERAUX On G_GENERAL=E_GENERAL ') ;
{ Construction de la clause Where de la SQL }
QEcr.SQL.Add('Where ') ;
QEcr.SQL.Add(' E_JOURNAL=:J_JOURNAL ') ;
QEcr.SQL.Add(' And E_EXERCICE="'+CritEdt.Exo.Code+'" ') ;
QEcr.SQL.Add(' And E_DATECOMPTABLE>="'+USDATETIME(CritEdt.Date1)+'" ') ;
QEcr.SQL.Add(' And E_DATECOMPTABLE<="'+USDATETIME(CritEdt.Date2)+'" ') ;
QEcr.SQL.Add(' And E_NUMEROPIECE>='+IntToStr(CritEdt.Jal.NumPiece1)+' and E_NUMEROPIECE<='+IntToStr(CritEdt.Jal.NumPiece2)+' ') ;
if CritEdt.RefInterne<>'' then QEcr.SQL.Add('and UPPER(E_REFINTERNE) like "'+TraduitJoker(CritEdt.RefInterne)+'" ' );
if CritEdt.Etab<>'' then QEcr.SQL.Add(' And E_ETABLISSEMENT="'+CritEdt.Etab+'" ') ;
if CritEdt.Valide<>'g' then QEcr.SQL.Add(' And E_VALIDE="'+CritEdt.Valide+'" ') ;
if CritEdt.DeviseSelect<>'' then QEcr.SQL.Add(  ' And E_DEVISE="'+CritEdt.DeviseSelect+'" ') ;
Case NET of
     neRien : BEGIN
              QEcr.SQL.Add(TraduitNatureEcr(CritEdt.QualifPiece, 'E_QUALIFPIECE', true, CritEdt.ModeRevision)) ;
              QEcr.SQL.Add(' and E_QUALIFPIECE<>"C" ') ;
              END ;
     neAno  : QEcr.SQL.Add(' And E_QUALIFPIECE="N" And (E_ECRANOUVEAU="OAN" or E_ECRANOUVEAU="H")') ;
     neClo  : QEcr.SQL.Add(' And E_QUALIFPIECE="C" And E_ECRANOUVEAU="C" ') ;
    End ;

  { Construction de la clause Order By de la SQL }
  // GC
  if CritEdt.Jal.AvecNouvFct then
  begin
    if CritEdt.Jal.Tri then
      QEcr.SQL.Add(' Order By E_JOURNAL,E_EXERCICE,E_PERIODE,E_NUMEROPIECE,E_DATECOMPTABLE,E_NUMLIGNE,E_NUMECHE')
    else
      //QEcr.SQL.Add(' Order By E_JOURNAL,E_EXERCICE,E_PERIODE,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE');// ????
      QEcr.SQL.Add(' Order By E_JOURNAL,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE') ;
  end
  else
  begin
    if CritEdt.Jal.Tri then
      QEcr.SQL.Add(' Order By E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE')
    else
      QEcr.SQL.Add(' Order By E_JOURNAL,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE') ;
  end;

ChangeSQL(QEcr) ; QEcr.Open ;
END;

procedure TQRJalDiv.RenseigneCritere ;
{ Récupération des champs du multicritère dans l'entête d'état : OK }
BEGIN
Inherited ;
RRefInterne.Caption:=FRefInterne.Text   ;
RNumPiece1.Caption:=IntToStr(CritEdt.Jal.NumPiece1) ; RNumPiece2.Caption:=IntToStr(CritEdt.Jal.NumPiece2) ;
CaseACocher(FTri,RTri) ;
CaseACocher(FValide,RValide) ;
CaseACocher(FAvecJalAN,RAN) ;
if NET=neRien then Exit ;
QRLabel8.Visible:=False ; RTri.Visible:=False ;
QRLabel7.Visible:=False ; RNatureEcr.Visible:=False ;
QRLabel25.Visible:=False ; RValide.Visible:=False ;
QRLabel9.Visible:=False ; RAN.Visible:=False ;
END ;

procedure TQRJalDiv.ChoixEdition ;
BEGIN
Inherited ;
  ChargeGroup(Lmulti,['MOIS','PIEC','ECH']) ;
END ;

procedure TQRJalDiv.FormShow(Sender: TObject);
begin
Case NET Of
  neAno : BEGIN
          HelpContext:=7757000 ;
          //Standards.HelpContext:=7757100 ;
          //Avances.HelpContext:=7757200 ;
          //Mise.HelpContext:=7757300;
          //Option.HelpContext:=7757400 ;
          END ;
  neClo : BEGIN
          HelpContext:=7766000 ;
          //Standards.HelpContext:=7766100 ;
          //Avances.HelpContext:=7766200 ;
          //Mise.HelpContext:=7766300 ;
          //Option.HelpContext:=7766400 ;
          END ;
  END ;

// GC-GP - 04/01/2002
if TRUE Or (CtxPCL in V_PGI.PGIContexte) then
begin
  FTri.Checked := True;
  FTri.Enabled := False;
  FNouvFct.Checked := True;
  FVoirLibelleCompte.Checked := False;
end;
// FIN GC - 04/01/2002

inherited;

TabSup.TabVisible:=False;
FLignePieEntete.State:=cbChecked ;
FValide.State:=cbGrayed ;

// GC - 12/12/2001
if CritEdtChaine.Utiliser then InitEtatchaine('QRJDIVIS');

if NET=neRien then Exit ;
if NET=neAno then BEGIN Caption:=MsgLibel.Mess[17] ; FCpte1.ZoomTable:=tzJAN ; FCpte2.ZoomTable:=FCpte1.ZoomTable ; END
             Else BEGIN Caption:=MsgLibel.Mess[18] ; FCpte1.ZoomTable:=tzJCloture ; FCpte2.ZoomTable:=FCpte1.ZoomTable ; END ;
UpdateCaption(Self) ;
{ Zone Dates (De, à) }
FDateCompta1.Enabled:=False ; FDateCompta2.Enabled:=False ;
HLabel6.Enabled:=False ; Label7.Enabled:=False ;
{ Zone Type d'écriture}
HLabel5.Visible:=False ; FNatureEcr.Visible:=False ;
{ Zone fourchette de journal}
TFGen.Enabled:=False ; TFGenJoker.Enabled:=False ;
FCpte1.Enabled:=False ; FCpte2.Enabled:=False ; FJoker.Enabled:=False ;

FValide.Visible:=False ; AvecRevision.Visible:=False ;
FAvecJalAN.Visible:=False ;
if NET=neAno then FAvecJalAN.Checked:=True else FAvecJalAN.Checked:=False ;
FMonetaire.Enabled:=False ; FTri.Checked:=False ; FTri.Enabled:=False ;
QRP.ReportTitle:=Caption ;
If NET=neClo Then FExercice.Value:=VH^.Precedent.Code ;
end;

procedure TQRJalDiv.BEcrDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var St3 : String ;
    StTaux,StCotation : String ;
    MoIn : Boolean ;

begin
  inherited;
PrintBand:=Not QR2E_JOURNAL.IsNull ;
If Not PrintBand then exit ;
AncienEche:=QR2E_NUMECHE.AsInteger ;

  // GC-GP
  if TRUE Or (CtXPcl in V_Pgi.PGIContexte) then
  begin
    LE_NUMLIGNE.Caption:=IntToStr(QR2E_NUMLIGNE.AsInteger) ;
  end
  else
  begin
    if (QR2E_ECHE.asString='X') And (QR2E_ETATLETTRAGE.asString<>'RI') then
    BEGIN
      LE_NUMLIGNE.Caption:=IntToStr(QR2E_NUMLIGNE.AsInteger)+'/'+IntToStr(QR2E_NUMECHE.AsInteger) ;
    END
    else
    BEGIN
      LE_NUMLIGNE.Caption:=IntToStr(QR2E_NUMLIGNE.AsInteger) ;
    END ;
  end;


LE_REFINTERNE.Caption:=QR2E_REFINTERNE.AsString ;
LE_DATEECHEANCE.Caption:=DateToStr(QR2E_DATEECHEANCE.AsDateTime) ;
LE_LIBELLE.Caption:=QR2E_LIBELLE.AsString ;

LeCompte.Caption:=QR2E_GENERAL.ASString+'  '+QR2E_AUXILIAIRE.ASString ;
AddReport([1,2,3],CritEdt.Jal.FormatPrint.Report,QR2DEBIT.AsFloat,QR2Credit.AsFloat,CritEdt.Decimale) ;

//GC - Modofication de l' affichage de la VALIDATION
  if QR2E_VALIDE.AsString = 'X' then
    LE_VALIDE.Caption := 'V'
  else
    LE_VALIDE.Caption := ' ' ;

  if Trim(le_valide.caption) = '' then
    Le_Valide.Caption := '    ' + Copy(QR2E_DATECOMPTABLE.AsString, 1 ,2 ) + ' ' + QR2E_NATUREPIECE.AsString
  else
    Le_Valide.Caption := LE_Valide.Caption + ' ' + Copy(QR2E_DATECOMPTABLE.AsString, 1 ,2 ) + ' ' + QR2E_NATUREPIECE.AsString ;

// GC - Modification sur l' affichage de la DEVISE et du TAUX
  LE_DEVISE.Caption:=QR2E_DEVISE.AsString ;

  MoIn:=FALSE ;
  if EstMonnaieIn(QR2E_DEVISE.AsString) then
    MoIn:=TRUE ;

  if QR2E_DEVISE.AsString <> V_PGI.DevisePivot then
  begin
    StTaux := FormatFloat('##0.#####',QR2E_TAUXDEV.AsFloat) ;
    StCotation := FormatFloat('##0.#####',QR2E_COTATION.AsFloat) ;

    if Qr2E_DATECOMPTABLE.AsDateTime < V_PGI.DateDebutEuro then
      if CritEdt.Jal.AvecNouvFct then
        St3 := StTaux
      else
        St3 := StTaux + '   ' + DateToStr(QR2E_DATETAUXDEV.AsDateTime)
    else
    begin
      if MoIn then
      begin
        if CritEdt.Jal.AfficheTauxEuro then
          St3 := StCotation + ' (Eur)'
        else
          St3 := StTaux ;
      end
      else
      begin
        if CritEdt.Jal.AfficheTauxEuro then
          if CritEdt.Jal.AvecNouvFct then
            St3 := StCotation + ' (Eur)'
          else
            St3 := StCotation + ' (Eur) ' + DateToStr(QR2E_DATETAUXDEV.AsDateTime)
        else
          if CritEdt.Jal.AvecNouvFct then
            St3 := StTaux + '   ' + DateToStr(QR2E_DATETAUXDEV.AsDateTime)
          else
            St3 := StTaux;
      end ;
    end ;

    LE_DEVISE.Caption := LE_DEVISE.Caption + '   ' + St3 ;

  END;

  AddGroup(Lmulti, [QR2E_DATECOMPTABLE,QR2E_NUMEROPIECE,Qr2MULTIECHE],[QR2DEBIT.AsFloat,QR2Credit.AsFloat]) ;

{ Incrémentation des montants pour ... }
{ Chaque Journal }
TotJour[1].TotDebit:=  Arrondi(TotJour[1].TotDebit+QR2DEBIT.AsFloat, CritEdt.Decimale) ;
TotJour[1].TotCredit:= Arrondi(TotJour[1].TotCredit+QR2Credit.AsFloat, CritEdt.Decimale) ;
{ Le Total Gébéral }
TotEdt[1].TotDebit:=  Arrondi(TotEdt[1].TotDebit+QR2DEBIT.AsFloat, CritEdt.Decimale) ;
TotEdt[1].TotCredit:= Arrondi(TotEdt[1].TotCredit+QR2Credit.AsFloat, CritEdt.Decimale) ;
if (QR2E_AUXILIAIRE.asString<>'') then LibCompte.Caption:=QR2T_LIBELLE.AsString
                                  Else LibCompte.Caption:=QR2G_LIBELLE.AsString ;

if QEcr.FindField('E_ECHE').asString='X'
   then BEGIN LE_MODEPAIE.Visible:=True ;  LE_DATEECHEANCE.Visible:=True ;  END
   else BEGIN LE_MODEPAIE.Visible:=False ; LE_DATEECHEANCE.Visible:=False ; END ;
If (NET=neClo) Or (NET=neAno) Then
   BEGIN
   LE_MODEPAIE.Visible:=FALSE ; LE_DATEECHEANCE.Visible:=FALSE ;
   END ;
Quoi:=QuoiMvt ;
end;

procedure TQRJalDiv.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
  TITREREPORTH.Left     := TITREREPORTB.Left ;
  TITREREPORTH.Width    := TITREREPORTB.Width ;
  TITREREPORTH.Caption  := TITREREPORTB.Caption ;
  Report1Debit.Caption  := Report2Debit.Caption ;
  Report1Credit.Caption := Report2Credit.Caption ;
end;

procedure TQRJalDiv.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var St1, St2 : String ;
begin
  inherited;
// GC
  if CritEdt.Jal.AvecNouvFct then
    PrintBand:= not QR2E_NUMEROPIECE.IsNull ;

  { Init des Totaux par Journal }
Fillchar(TotJour,SizeOf(TotJour),#0) ;
StReportJal:=QR1J_JOURNAL.AsString ;
InitReport([2],CritEdt.Jal.FormatPrint.Report) ;
TJ_JOURNAL.Left:=TLE_NUMLIGNE.Left ;
TJ_JOURNAL.Width:=TLE_MODEPAIE.Left-TE_VALIDE.Width-1 ;
TJ_NATUREJAL.Left:=TLE_MODEPAIE.Left ;
TJ_NATUREJAL.Width:=TLE_MODEPAIE.Width+TLE_DATEECHEANCE.Width+TLE_LIBELLE.Width+2 ;
St1:=QR1J_JOURNAL.AsString+'  '+QR1J_LIBELLE.AsString ;
St2:=MsgLibel.Mess[14]+'  '+RechDom('ttNatJal',QR1J_NATUREJAL.AsString,False) ;
TJ_JOURNAL.Caption:=St1 ;
TJ_NATUREJAL.Caption:=St2 ;
QuoiJal(1) ;
end;

procedure TQRJalDiv.BFJournalBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
// GC
  if CritEdt.Jal.AvecNouvFct then
    PrintBand:= not QR2E_NUMEROPIECE.IsNull ;

  { Affiche les Totaux Par journal }
TJ_JOURNAL_.Caption:=MsgLibel.Mess[15]+'  '+QR1J_JOURNAL.AsString ;
TotDebJal.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotJour[1].TotDebit,CritEdt.AfficheSymbole ) ;
TotCreJal.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotJour[1].TotCredit,CritEdt.AfficheSymbole ) ;
QuoiJal(1) ;
end;

procedure TQRJalDiv.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TotGeneDeb.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[1].TotDebit,CritEdt.AfficheSymbole ) ;
TotGeneCre.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[1].TotCredit,CritEdt.AfficheSymbole ) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TQRJalDiv.BFJournalAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([2],CritEdt.Jal.FormatPrint.Report) ;
end;

procedure TQRJalDiv.BHPieceBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var St1, St2{, St3} : String ;
//    StTaux,StCotation : String ;
//    MoIn : Boolean ;
begin
  inherited;

PrintBand:= not QR2E_NUMEROPIECE.IsNull ;
InitReport([3],CritEdt.Jal.FormatPrint.Report) ;
If Not PrintBand then exit ;
TE_NUMEROPIECE.Width:=TLE_REFINTERNE.Left-1 ;
TE_TAUXDEV.Left:=TLE_MODEPAIE.Left ; TE_TAUXDEV.Width:=TLE_MODEPAIE.Width+TLE_DATEECHEANCE.Width+1 ;

// GC
  if (QR2E_MODESAISIE.AsString = 'BOR') or (QR2E_MODESAISIE.AsString = 'LIB') then
  begin // Mode Bordereau ou Mode Libre
    TE_NUMEROPIECE.Caption:= MsgLibel.Mess[19] + ' ' +
                             IntToStr( QR2E_NUMEROPIECE.AsInteger) + '     ' +
                             FormatDateTime('mmmm yyyy', QR2E_DATECOMPTABLE.AsDateTime) ;
  end
  else
  begin // Mode Pièce
    St1 := IntToStr(QR2E_NUMEROPIECE.AsInteger) ;
    St2 := DateToStr(QR2E_DATECOMPTABLE.AsDateTime) ;
    TE_NUMEROPIECE.Caption := MsgLibel.Mess[11] + ' ' + St1 + '     ' + St2 ;
  end;

TE_DEVISE.Caption := '';
TE_TAUXDEV.Caption := '';
// GC
(*
TE_DEVISE.Caption:=QR2E_DEVISE.AsString ;
MoIn:=FALSE ; If EstMonnaieIn(QR2E_DEVISE.AsString) Then MoIn:=TRUE ;
if QR2E_DEVISE.AsString<>V_PGI.DevisePivot then
   BEGIN
   StTaux:=formatFloat('##0.#####',QR2E_TAUXDEV.AsFloat) ;
   StCotation:=formatFloat('##0.#####',QR2E_COTATION.AsFloat) ;
   If Qr2E_DATECOMPTABLE.AsDateTime<V_PGI.DateDebutEuro Then St3:=StTaux+'   '+DateToStr(QR2E_DATETAUXDEV.AsDateTime) Else
     BEGIN
     If MoIn Then
       BEGIN
       If CritEdt.Jal.AfficheTauxEuro Then St3:=StCotation+'(Eur)' Else St3:=StTaux+'   ' ;
       END Else
       BEGIN
       If CritEdt.Jal.AfficheTauxEuro Then St3:=StCotation+'(Eur)'+DateToStr(QR2E_DATETAUXDEV.AsDateTime)
                                      Else St3:=StTaux+'   '+DateToStr(QR2E_DATETAUXDEV.AsDateTime) ;
       END ;
     END ;
   TE_TAUXDEV.Caption:=St3 ;
//   TE_TAUXDEV.Caption:=FloatToStr(QR2E_TAUXDEV.AsFloat)+'   '+DateToStr(QR2E_DATETAUXDEV.AsDateTime) ;
   END Else
   BEGIN
   TE_TAUXDEV.Caption:='';
   END ;
*)

TE_ETABLISSEMENT.Caption:=RechDom('ttEtablissement',QR2E_ETABLISSEMENT.AsString,False) ;

  StReportPie := IntToStr( QR2E_NUMEROPIECE.AsInteger ) ;
//StReportPie := QR1J_JOURNAL.AsString + ' / ' + IntToStr(QR2E_NUMEROPIECE.AsInteger) ;

If (NET=neClo) Or (NET=neAno) Then TE_TAUXDEV.Caption:='';
end;

procedure TQRJalDiv.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var D,C : Double ;
begin
  inherited;
TITREREPORTB.Left:=TLE_MODEPAIE.Left ;
TITREREPORTB.Width:=TLE_MODEPAIE.Width+TLE_DATEECHEANCE.Width+TLE_LIBELLE.Width+2 ;

  case QuelReport(CritEdt.Jal.FormatPrint.Report,D,C) of
    1 : TITREREPORTB.Caption := MsgLibel.Mess[4] ;                       // REPORT TOTAL GENERAL
    2 : TITREREPORTB.Caption := MsgLibel.Mess[5] + '  ' + StReportJal ;  // REPORT JOURNAL

    // GC
    // REPORT JOURNAL / PIECE OU FOLIO
    3 : begin
          TITREREPORTB.Caption := TraduireMemoire('REPORT JOURNAL') + '  ' + StReportJal ;
          if (QR2E_MODESAISIE.AsString = 'BOR') or (QR2E_MODESAISIE.AsString = 'LIB') then
            TITREREPORTB.Caption := TITREREPORTB.Caption + ' / ' + MsgLibel.Mess[19] + ' ' + StReportPie
          else
            TITREREPORTB.Caption := TITREREPORTB.Caption + ' / ' + MsgLibel.Mess[11] + ' ' + StReportPie;
        end;
    //3 : TITREREPORTB.Caption := MsgLibel.Mess[6] + '  ' + StReportPie ;
    else
      TITREREPORTB.Caption := '' ; //  Rony --> à cause des journaux non mouvementés en saut de page
  end ;

Report2Debit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, D, CritEdt.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, C, CritEdt.AfficheSymbole ) ;
end;

procedure TQRJalDiv.DLMULTINeedData(var MoreData: Boolean);
Var CodMulti,LibMulti : String ;
    TotMulti : Array[0..12] of Double ;
begin
  inherited;

  MoreData:=PrintGroup(Lmulti,QEcr,[QR2E_DATECOMPTABLE,QR2E_NUMEROPIECE,Qr2MULTIECHE],CodMulti,LibMulti,TotMulti,NumR);

if ((MoreData) and (Not QEcr.EOF) And (NumR=1)) then OKEnteteP:=True ;
Case NumR of  { Rupture/Mois }
   0 : BEGIN
       BMulti.Font.color:=clRed  ; TTOTMULTI.Caption:=MsgLibel.Mess[8]+' '+CodMulti;
       BMulti.Font.Size:=8 ;
       TTOTMULTI.AutoSize:=False ;
       TTOTMULTI.Left:=TLE_LIBELLE.Left ; TTOTMULTI.Tag:=TLE_LIBELLE.Tag ;
       TTOTMULTI.Width:=TLE_LIBELLE.Width ;
       TTOTMULTI.Alignment:=taRightJustify ;
       TDEBMULTI.Caption:=AfficheMontant(CRITEdt.FormatMontant, CRITEdt.Symbole, TotMulti[0] , CRITEdt.AfficheSymbole) ; { + 'ANouveau' + 'Cumul au' en Débit  }
       TCREMULTI.Caption:=AfficheMontant(CRITEdt.FormatMontant, CRITEdt.Symbole, TotMulti[1] , CRITEdt.AfficheSymbole) ; { + 'ANouveau' + 'Cumul au' en Crédit }
       END ;
   1 : BEGIN  { Rupture/Piece }
       BMulti.Frame.DrawTop:=CritEdt.Jal.FormatPrint.PrSepCompte[4] ;
       BMulti.Font.color:=clNavy  ;

       // GC
       if (QR2E_MODESAISIE.AsString = 'BOR') or (QR2E_MODESAISIE.Asstring = 'LIB') then
         TTOTMULTI.Caption := MsgLibel.Mess[20] + ' ' + CodMulti
       else
         TTOTMULTI.Caption := MsgLibel.Mess[9]  + ' ' + CodMulti ;

       BMulti.Font.Size:=8 ;
       TTOTMULTI.AutoSize:=True ; TTOTMULTI.Left:=TLE_NUMLIGNE.Left ;
       TTOTMULTI.Tag:=TLE_NUMLIGNE.Tag ;
       TTOTMULTI.Alignment:=taLeftJustify ;
       TDEBMULTI.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotMulti[0], CritEdt.AfficheSymbole) ;
       TCREMULTI.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotMulti[1], CritEdt.AfficheSymbole) ;
       END ;
   2 : BEGIN  { Rupture/Eche }
       BMulti.Frame.DrawTop:=False ;
       BMulti.Font.color:=clFuchsia  ; TTOTMULTI.Caption:=MsgLibel.Mess[16] ;
       TTOTMULTI.AutoSize:=False ;
       TTOTMULTI.Left:=TLE_LIBELLE.Left ; TTOTMULTI.Tag:=TLE_LIBELLE.Tag ;
       TTOTMULTI.Width:=TLE_LIBELLE.Width ;
       TTOTMULTI.Alignment:=taRightJustify ;
       if TotMulti[0]<>0 then TDEBMULTI.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotMulti[0], CritEdt.AfficheSymbole)
                    Else TDEBMULTI.Caption:='' ;
       if TotMulti[1]<>0 then TCREMULTI.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotMulti[1], CritEdt.AfficheSymbole)
                    Else TCREMULTI.Caption:='' ;
       END ;
   END ;
end;

procedure TQRJalDiv.QEcrAfterOpen(DataSet: TDataSet);

begin
  inherited;
QR2E_JOURNAL         :=TStringField(QEcr.FindField('E_JOURNAL')) ;
QR2E_GENERAL         :=TStringField(QEcr.FindField('E_GENERAL')) ;
QR2E_AUXILIAIRE      :=TStringField(QEcr.FindField('E_AUXILIAIRE')) ;
QR2E_EXERCICE        :=TStringField(QEcr.FindField('E_EXERCICE')) ;
QR2E_DATECOMPTABLE   :=TDateTimeField(QEcr.FindField('E_DATECOMPTABLE')) ;
QR2E_DATECOMPTABLE.Tag:=1 ;
QR2DEBIT             :=TFloatField(QEcr.FindField('DEBIT')) ;
QR2CREDIT            :=TFloatField(QEcr.FindField('CREDIT')) ;
QR2E_NUMEROPIECE     :=TIntegerField(QEcr.FindField('E_NUMEROPIECE')) ;
QR2E_NUMLIGNE        :=TIntegerField(QEcr.FindField('E_NUMLIGNE')) ;
QR2E_NUMECHE         :=TIntegerField(QEcr.FindField('E_NUMECHE')) ;
QR2MULTIECHE         :=TIntegerField(QEcr.FindField('MULTIECHE')) ;
QR2E_REFINTERNE      :=TStringField(QEcr.FindField('E_REFINTERNE')) ;
QR2E_LIBELLE         :=TStringField(QEcr.FindField('E_LIBELLE')) ;
QR2E_NATUREPIECE     :=TStringField(QEcr.FindField('E_NATUREPIECE')) ;
QR2E_QUALIFPIECE     :=TStringField(QEcr.FindField('E_QUALIFPIECE')) ;
QR2E_MODEPAIE        :=TStringField(QEcr.FindField('E_MODEPAIE')) ;
// GC
QR2E_MODESAISIE      :=TStringField(QEcr.FindField('E_MODESAISIE')) ;
//
QR2E_ECHE            :=TStringField(QEcr.FindField('E_ECHE')) ;
QR2E_DATEECHEANCE    :=TDateTimeField(QEcr.FindField('E_DATEECHEANCE')) ;
QR2E_DEVISE          :=TStringField(QEcr.FindField('E_DEVISE')) ;
QR2E_TAUXDEV         :=TFloatField(QEcr.FindField('E_TAUXDEV')) ;
QR2E_COTATION        :=TFloatField(QEcr.FindField('E_COTATION')) ;
QR2E_DATETAUXDEV     :=TDateTimeField(QEcr.FindField('E_DATETAUXDEV')) ;
QR2E_ETABLISSEMENT   :=TStringField(QEcr.FindField('E_ETABLISSEMENT')) ;
QR2E_VALIDE          :=TStringField(QEcr.FindField('E_VALIDE')) ;
QR2E_ETATLETTRAGE    :=TStringField(QEcr.FindField('E_ETATLETTRAGE')) ;
QR2T_LIBELLE         :=TStringField(QEcr.FindField('T_LIBELLE')) ;
QR2G_LIBELLE         :=TStringField(QEcr.FindField('G_LIBELLE')) ;
QR1J_JOURNAL         :=TStringField(Q.FindField('J_JOURNAL')) ;
ChgMaskChamp(QR2DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
ChgMaskChamp(QR2CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
end;

procedure TQRJalDiv.DLPieceHNeedData(var MoreData: Boolean);
begin
  inherited;
MoreData:=OkEnteteP ;
OkEnteteP:=False ;
end;

procedure TQRJalDiv.BDetailAfterPrint(BandPrinted: Boolean);
begin
  inherited;
OkEnteteP:=TRUE ;

end;

procedure TQRJalDiv.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
QR1J_JOURNAL:=TStringField(Q.FindField('J_JOURNAL')) ;
QR1J_LIBELLE:=TStringField(Q.FindField('J_LIBELLE')) ;
QR1J_ABREGE:=TStringField(Q.FindField('J_ABREGE')) ;
QR1J_NATUREJAL:=TStringField(Q.FindField('J_NATUREJAL')) ;
end;

procedure TQRJalDiv.QEcrBeforeOpen(DataSet: TDataSet);
begin
  inherited;
QEcr.PArams[0].AsString:=Q.FindField('J_JOURNAL').AsString ;
end;

procedure TQRJalDiv.BMULTIAfterPrint(BandPrinted: Boolean);
begin
  inherited;
If (NumR=1) then InitReport([3],CritEdt.Jal.FormatPrint.Report) ;
end;

procedure TQRJalDiv.BMULTIBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
if NumR=2 then
   BEGIN
   if (AncienEche<=1) then PrintBand:=False else PrintBand:=True ;
   END ;

end;

// GC - 12/12/2001
procedure TQRJalDiv.Timer1Timer(Sender: TObject);
begin
  inherited;
  //
end;

end.
