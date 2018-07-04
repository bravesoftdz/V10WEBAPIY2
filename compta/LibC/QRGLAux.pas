unit QRGLAux;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  StdCtrls, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, UtilEdt, Ent1, HEnt1, CpteUtil, QRRupt,
{$IFNDEF CCMP}
  EdtLegal,
{$ENDIF}
  HQry, CritEdt, HSysMenu, Menus, HTB97, HPanel, UiUtil, tCalcCum ;

procedure GLAuxiliaire ;
// GC - 20/12/2001
procedure GLAuxiliaireChaine ( vCritEdtChaine : TCritEdtChaine );
// GC - 20/12/2001 - FIN
procedure GLAuxiliaireZoom(Crit : TCritEdt) ;
procedure GLAuxiliaireLegal(Crit : TCritEDT) ;

type
  TFGdLivAux = class(TFQR)
    FValide: TCheckBox;
    HLabel9: THLabel;
    FNumPiece1: TMaskEdit;
    Label12: TLabel;
    FNumPiece2: TMaskEdit;
    FRefInterne: TEdit;
    HLabel10: THLabel;
    FSautPage: TCheckBox;
    FLigneCptPied: TCheckBox;
    QEcr: TQuery;
    SEcr: TDataSource;
    QRDLCpt: TQRDetailLink;
    QRDLMulti: TQRDetailLink;
    DLRupt: TQRDetailLink;
    MsgBox: THMsgBox;
    RValide: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel10: TQRLabel;
    RRefInterne: TQRLabel;
    QRLabel13: TQRLabel;
    RNumPiece1: TQRLabel;
    QRLabel14: TQRLabel;
    RNumPiece2: TQRLabel;
    TRLegende: TQRLabel;
    RLegende: TQRLabel;
    TLE_VALIDE: TQRLabel;
    TLE_DATECOMPTABLE: TQRLabel;
    TE_PIECE: TQRLabel;
    TE_REFINTERNE: TQRLabel;
    TE_LIBELLE: TQRLabel;
    TLE_DEBIT: TQRLabel;
    TLE_CREDIT: TQRLabel;
    TLE_SOLDE: TQRLabel;
    TLE_LETTRAGE: TQRLabel;
    TITRE1REP: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    REPORT1SOLDE: TQRLabel;
    T_AUXILIAIRE: TQRLabel;
    QRLabel6: TQRLabel;
    TCumul: TQRLabel;
    AnvDebit: TQRLabel;
    CumDebit: TQRLabel;
    AnvCredit: TQRLabel;
    AnvSolde: TQRLabel;
    CumSolde: TQRLabel;
    CumCredit: TQRLabel;
    BDetailEcr: TQRBand;
    LE_DATECOMPTABLE: TQRDBText;
    LE_CREDIT: TQRDBText;
    LE_DEBIT: TQRDBText;
    E_LIBELLE: TQRDBText;
    E_REFINTERNE: TQRDBText;
    LE_LETTRAGE: TQRDBText;
    LE_SOLDE: TQRLabel;
    LE_VALIDE: TQRLabel;
    E_PIECELIGECH: TQRLabel;
    BSDMulti: TQRBand;
    LibTotalMulti: TQRLabel;
    TotMultiDebit: TQRLabel;
    TotMultiCredit: TQRLabel;
    TotSoldeMulti: TQRLabel;
    BFCompteAux: TQRBand;
    T_AUXILIAIRE_: TQRLabel;
    TotSoldeCptAux: TQRLabel;
    TotCptCredit: TQRLabel;
    TotCptDebit: TQRLabel;
    BRupt: TQRBand;
    DebitRupt: TQRLabel;
    TCodRupt: TQRLabel;
    CreditRupt: TQRLabel;
    SoldeRupt: TQRLabel;
    QRLabel33: TQRLabel;
    TotGenDebit: TQRLabel;
    TotGenCredit: TQRLabel;
    TotalSoldes: TQRLabel;
    QRBand2: TQRBand;
    Trait5: TQRLigne;
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
    TE_JOURNAL: TQRLabel;
    LE_JOURNAL: TQRDBText;
    FSoldeProg: TCheckBox;
    FTotMens: TCheckBox;
    FTotEche: TCheckBox;
    Label1: TLabel;
    QRLabel8: TQRLabel;
    TotDesMvtsDebit: TQRLabel;
    TotDesMvtsCredit: TQRLabel;
    TotDesMvtsSolde: TQRLabel;
    FLigneRupt: TCheckBox;
    FAvecAN: TCheckBox;
    BRupTete: TQRBand;
    CodeRuptAu: TQRLabel;
    DLRuptH: TQRDetailLink;
    AnoRupt: TQRLabel;
    CumRupt: TQRLabel;
    AnoDebitRupt: TQRLabel;
    CumDebitRupt: TQRLabel;
    AnoCreditRupt: TQRLabel;
    CumCreditRupt: TQRLabel;
    AnoSoldeRupt: TQRLabel;
    CumSoldeRupt: TQRLabel;
    FLettrable: TCheckBox;
    QRLabel1: TQRLabel;
    RLettrable: TQRLabel;
    TotRupt: TQRLabel;
    procedure FormShow(Sender: TObject);
    procedure QEcrCalcFields(DataSet: TDataSet);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BDetailEcrAfterPrint(BandPrinted: Boolean);
    procedure BFCompteAuxBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure QRDLMultiNeedData(var MoreData: Boolean);
    procedure BSDMultiBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure DLRuptNeedData(var MoreData: Boolean);
    procedure BRuptBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BRuptAfterPrint(BandPrinted: Boolean);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BDetailEcrBeforePrint(var PrintBand: Boolean;  var Quoi: string);
    procedure QEcrAfterOpen(DataSet: TDataSet);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure FSelectCpteChange(Sender: TObject);
    procedure FExerciceChange(Sender: TObject);
    procedure FDateCompta1Change(Sender: TObject);
    procedure BFCompteAuxAfterPrint(BandPrinted: Boolean);
    procedure DLRuptHNeedData(var MoreData: Boolean);
    procedure BRupTeteBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FSansRuptClick(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure BDetailAfterPrint(BandPrinted: Boolean);
    procedure BDetailCheckForSpace;
    procedure Timer1Timer(Sender: TObject);
    procedure GenereSQL ; Override ;
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
  private    { Déclarations privées }
    DansTotGen			             : Boolean ;
    LibRupt, CodRupt, StReportAux : string ;
    Imprime, Affiche,SautPageRupture : Boolean ;
    AncienEche                    : Integer ;
//    QBal                          : TQuery ;
    TotEdt, TotCptAux,TotAnCRupt,
    TotDesMvts                    : TabTot ;
    LMulti, LRupt                 : TStringList ;
    Tot                           : Array[0..12] of Double ;
    Qr1T_AUXILIAIRE               : TStringField ;
    Qr1T_LIBELLE                  : TStringField ;
    Qr1T_SAUTPAGE                 : TStringField ;
    Qr1T_SOLDEPROGRESSIF          : TStringField ;
    Qr1T_TOTAUXMENSUELS           : TStringField ;
    QR2E_AUXILIAIRE,Qr2E_JOURNAL,
    Qr2E_EXERCICE,Qr2E_VALIDE,
    Qr2E_QUALIFPIECE              : TStringField ;
    Qr2E_NUMEROPIECE,Qr2E_NUMLIGNE,Qr2E_NUMECHE,QR2MULTIECHE : TIntegerFIeld ;
    Qr2E_DateComptable   : TDateTimeField ;
    Qr2DEBIT,Qr2CREDIT   : TFloatField ;
    FLoading             : Boolean ;
    IsLegal, OkEnteteRup : Boolean ;
    Qr1T_CORRESP1, Qr1T_CORRESP2 : TStringField ;
    Function QuoiAux(i : Integer) : String ;
    Function QuoiMvt : String ;
    Procedure GLAuxZoom(Quoi : String) ;
  public     { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Function TFGdLivAux.QuoiAux(i : Integer) : String ;
BEGIN
Inherited ;
Case i Of
  1 : Result:=Qr1T_AUXILIAIRE.AsString+' '+'#'+Qr1T_LIBELLE.AsString+'@'+'2' ;
  2 : Result:=Qr1T_AUXILIAIRE.AsString+' '+'#'+Qr1T_LIBELLE.AsString+' '+
              TotSoldecptAux.Caption+'@'+'2' ;
  END ;
END ;

Function TFGdLivAux.QuoiMvt : String ;
BEGIN
Inherited ;
Result:=Qr2E_AUXILIAIRE.AsString+' '+Qr1T_LIBELLE.AsString+' '+Le_Solde.Caption+
        '#'+Qr2E_JOURNAL.AsString+' N° '+IntToStr(Qr2E_NUMEROPIECE.AsInteger)+' '+DateToStr(Qr2E_DateComptable.AsDAteTime)+'-'+
        PrintSolde(Qr2DEBIT.AsFloat,Qr2Credit.AsFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)+
        '@'+'5;'+Qr2E_JOURNAL.AsString+';'+UsDateTime(Qr2E_DATECOMPTABLE.AsDateTime)+';'+Qr2E_NUMEROPIECE.AsString+';'+Qr2E_EXERCICE.asString+';'+
        IntToStr(Qr2E_NumLigne.AsInteger)+';' ;
END ;

Procedure TFGdLivAux.GLAuxZoom(Quoi : String) ;
Var Lp,i: Integer ;
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
END ;

procedure GLAuxiliaire ;
var QR : TFGdLivAux ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFGdLivAux.Create(Application) ;
Edition.Etat:=etGlAux ;
QR.QRP.QRPrinter.OnSynZoom:=QR.GLAuxZoom ;
QR.InitType (nbAux,neGL,msAuxEcr,'QRGLAUX','QR_GLAUX',TRUE,FALSE,FALSE,Edition) ;
QR.IsLegal:=FALSE ;
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

// GC - 20/12/2001
procedure GLAuxiliaireChaine( vCritEdtChaine : TCritEdtChaine );
var QR : TFGdLivAux ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFGdLivAux.Create(Application) ;
Edition.Etat:=etGlAux ;
QR.QRP.QRPrinter.OnSynZoom:=QR.GLAuxZoom ;
QR.CritEdtChaine := vCritEdtChaine;
QR.InitType (nbAux,neGL,msAuxEcr,'QRGLAUX','QR_GLAUX',TRUE,FALSE,FALSE,Edition) ;
QR.IsLegal:=FALSE ;
if PP=Nil then
   BEGIN
   try
    QR.Timer1.Enabled := True;
    QR.QRP.QRPrinter.Copies := QR.CritEdtChaine.NombreExemplaire;
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
// GC - 20/12/2001

procedure GLAuxiliaireZoom(Crit : TCritEDT) ;
var QR : TFGdLivAux ;
    Edition : TEdition ;
BEGIN
QR:=TFGdLivAux.Create(Application) ;
Edition.Etat:=etGlAux ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.GLAuxZoom ;
 QR.IsLegal:=FALSE ;
 QR.CritEDT:=Crit ;
 QR.InitType (nbAux,neGL,msAuxEcr,'QRGLAUX','',FALSE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure GLAuxiliaireLegal(Crit : TCritEDT) ;
var QR : TFGdLivAux ;
    Edition : TEdition ;
BEGIN
QR:=TFGdLivAux.Create(Application) ;
Edition.Etat:=etGlAux ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.GLAuxZoom ;
 QR.IsLegal:=TRUE ;
 QR.CritEDT:=Crit ;
 QR.InitType (nbAux,neGL,msAuxEcr,'QRGLAUX','',TRUE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFGdLivAux.InitDivers ;
BEGIN
InHerited ;
{ Gestion multiéchéances }
AncienEche:= 1 ; TLE_VALIDE.Caption:='' ;
{ Les tableaux de montants }
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
Fillchar(TotAncRupt,SizeOf(TotAncRupt),#0) ;
BFCompteAux.Frame.DrawTop:=CritEDT.GL.FormatPrint.PrSepCompte[2] ;
BRupt.Frame.DrawTop:=CritEdt.GL.FormatPrint.PrSepCompte[3] ;
//BRupt.Frame.DrawBottom:=BRupt.Frame.DrawTop ;
If CritEdt.GA Then CritEdt.SQLGA:=GetParamListe(NomListeGA,Self) ;
//OkEnteteRup:=(CritEdt.Rupture=rCorresp) ;
If CritEdt.SautPageRupt And (CritEdt.GL.RuptOnly=Avec) And (CritEdt.SautPage<>1) Then BDetail.ForceNewPage:=FALSE ;
SautPageRupture:=FALSE ;
END ;

Procedure TFGdLivAux.FinirPrint ;
{$IFNDEF CCMP}
var Solde : Double ;
{$ENDIF}
BEGIN
InHerited ;
VideGroup(LMulti) ; QEcr.Close ;
If GCalc<>NIL Then BEGIN GCalc.Free ; GCalc:=NIL ; END ;
if CritEdt.Rupture<>rRien then VideRupt(LRupt) ;
{$IFNDEF CCMP}
if OkMajEdt Then
   BEGIN
   Solde:=TotEdt[1].TotDebit-TotEdt[1].TotCredit ;
   if Solde<0 then
      BEGIN
      If IsLegal then MajEditionLegal('GLT', CritEDT.Exo.Code, DateToStr(CritEDT.DateDeb), DateToStr(CritEDT.DateFin),
                      '', TotEdt[1].TotDebit, TotEdt[1].TotCredit, Solde, 0)
                 Else MajEdition('GLT', CritEDT.Exo.Code, DateToStr(CritEDT.DateDeb), DateToStr(CritEDT.DateFin),
                      '', TotEdt[1].TotDebit, TotEdt[1].TotCredit, Solde, 0) ;
      END Else
      BEGIN
      If IsLegal then MajEditionLegal('GLT', CritEDT.Exo.Code, DateToStr(CritEDT.DateDeb), DateToStr(CritEDT.DateFin),
                      '', TotEdt[1].TotDebit, TotEdt[1].TotCredit, 0, Solde)
                 Else MajEdition('GLT', CritEDT.Exo.Code, DateToStr(CritEDT.DateDeb), DateToStr(CritEDT.DateFin),
                      '', TotEdt[1].TotDebit, TotEdt[1].TotCredit, 0, Solde) ;
      END ;
   END ;
{$ENDIF}
END ;

procedure TFGdLivAux.GenereSQL ;
{ Construction de la requête SQL en fonction du multicritère }
//Var St : String ;
BEGIN
InHerited ;
GenereSQLBase ;
(*
{ Construction de la clause Select de la SQL }
Q.Close ;
Q.SQL.Clear ;
St:='Select T_AUXILIAIRE,T_LIBELLE,T_COLLECTIF, T_TOTDEBANO, T_TOTCREANO, T_TOTDEBE,'+
            'T_TOTCREE, T_SAUTPAGE, T_SOLDEPROGRESSIF, T_TOTAUXMENSUELS' ;
Q.SQL.Add(St) ;
{ Tables explorées par la SQL }
St:=' From TIERS Q Where ' ;
Q.SQL.Add(St) ;
{ Construction de la clause Where de la SQL }
St:=SQLModeSelect ;
If St<>'' Then Q.SQL.Add(St) ;
St:='' ;
if CritEDT.Joker then St:='AND T_AUXILIAIRE like "'+TraduitJoker(CritEDT.Cpt1)+'" ' Else
   BEGIN
   if CritEDT.Cpt1<>'' then St:='AND T_AUXILIAIRE>="'+CritEDT.Cpt1+'" ' ;
   if CritEDT.Cpt2<>'' then St:=St+'AND T_AUXILIAIRE<="'+CritEDT.Cpt2+'" ' ;
   END ;
If St<>'' Then Q.SQL.Add(St) ;
st:='' ;
if CritEDT.NatureCpt<>'' then St:='AND T_NATUREAUXI="'+CritEDT.NatureCpt+'" ' ;
If St<>'' Then Q.SQL.Add(St) ;
st:='' ;
if CritEDT.GL.Sauf<>'' then St:=CritEDT.GL.SqlSauf ;
if st<>'' Then Q.SQL.Add(St) ;
{ Construction de la clause Order By de la SQL }
St:=' Order By T_AUXILIAIRE' ;
Q.SQL.Add(St) ;
ChangeSQL(Q) ; Q.Open ;
*)
{ Construction de la clause Select de la SQL }
QEcr.Close ;
QEcr.SQL.Clear ;
QEcr.SQL.Add('Select E_AUXILIAIRE,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,') ;
QEcr.SQL.Add(       'E_REFINTERNE,E_ETABLISSEMENT,E_LIBELLE,E_NUMECHE,E_DEVISE,E_VALIDE,') ;
QEcr.SQL.Add(       'E_NUMEROPIECE*10+E_NUMLIGNE as MULTIECHE,E_LETTRAGE,E_GENERAL,') ;
QEcr.SQL.Add(       'E_JOURNAL, E_NATUREPIECE, E_QUALIFPIECE, ') ;
If CritEdt.SQLGA<>'' Then QECR.SQL.Add(CritEdt.SQLGA) ;
Case CritEDT.Monnaie of
  0 : BEGIN QEcr.SQL.Add('E_DEBIT DEBIT,E_CREDIT CREDIT') ; END ;
  1 : BEGIN QEcr.SQL.Add('E_DEBITDEV DEBIT,E_CREDITDEV CREDIT') ; END ;
end ;
{ Tables explorées par la SQL }
QEcr.SQL.Add(' From ECRITURE') ;
{ Construction de la clause Where de la SQL }
QEcr.SQL.Add('Where E_AUXILIAIRE=:T_AUXILIAIRE ') ;
QEcr.SQL.Add('And E_DATECOMPTABLE>="'+usdatetime(CritEDT.GL.Date21)+'" And E_DATECOMPTABLE<="'+usdatetime(CritEDT.Date2)+'" ') ;
if FExercice.ItemIndex>0 then QEcr.SQL.Add(' And E_EXERCICE="'+CritEDT.Exo.Code+'" ') ;
QEcr.SQL.Add(   ' And E_NUMEROPIECE>='+IntToStr(CritEDT.GL.NumPiece1)+' ') ;
QEcr.SQL.Add(   ' And E_NUMEROPIECE<='+IntToStr(CritEDT.GL.NumPiece2)+' ') ;
QEcr.SQL.Add(TraduitNatureEcr(CritEDT.Qualifpiece, 'E_QUALIFPIECE', true, CritEdt.ModeRevision)) ;
if CritEDT.RefInterne<>'' then QEcr.SQL.Add(' And UPPER(E_REFINTERNE) like "'+TraduitJoker(CritEDT.RefInterne)+'" ' );
if CritEDT.Etab<>'' then QEcr.SQL.Add(' And E_ETABLISSEMENT="'+CritEDT.Etab+'"') ;
if CritEDT.Valide<>'g' then QEcr.SQL.Add(' And E_VALIDE="'+CritEDT.Valide+'" ') ;
if CritEDT.DeviseSelect<>'' then QEcr.SQL.Add(' And E_DEVISE="'+CritEDT.DeviseSelect+'"') ;
QEcr.SQL.Add(' And E_ECRANOUVEAU="N" and E_QUALIFPIECE<>"C" ') ;
Case CritEdt.GL.Lettrable of
  2 : QEcr.SQL.Add(' And E_LETTRAGE=""  ') ;
  1 : QEcr.SQL.Add(' And E_LETTRAGE<>"" ') ;
 End ;
{ Construction de la clause Order By de la SQL }
If CritEDT.SQLPlus<>'' Then QEcr.SQL.Add(CritEDT.SQLPlus) ;
QEcr.SQL.Add(' Order By E_AUXILIAIRE,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE') ;
ChangeSQL(QEcr) ; QEcr.Open ;
END ;

procedure TFGdLivAux.RenseigneCritere ;
{ Récupération des champs du multicritère dans l'entête d'état }
BEGIN
Inherited ;
If EstSerie(S5) Then RLegende.Caption:=MsgRien.Mess[7] ;
RRefInterne.Caption:=FRefInterne.Text ;
RNumPiece1.Caption:=IntToStr(CritEdt.GL.NumPiece1)   ; RNumPiece2.Caption:=IntToStr(CritEdt.GL.NumPiece2) ;
(*
if FSansRupt.Checked then BEGIN RSansRupt.Caption:='þ' ; RAvecRupt.Caption:='o' ; END ;
if FAvecRupt.Checked then BEGIN RSansRupt.Caption:='o' ; RAvecRupt.Caption:='þ' ; END ;
if FPlanRupt.Enabled then RPlanRupt.Caption:=FPlanRupt.Text ;
*)
CaseACocher(FValide,RValide) ; CaseACocher(FLettrable,RLettrable) ;
(*
CaseACocher(FSoldeProg,RSoldeProg) ;
CaseACocher(FSautPage,RSautPage)        ; CaseACocher(FTotMens,RTotMens) ;
CaseACocher(FLigneCptPied,RLigneCptPied);
CaseACocher(FTotEche,RTotEche)          ;
*)
DateCumulAuGL(TCumul,CritEdt,MsgBox.Mess[2]) ;
DateCumulAuGL(AnoRupt,CritEdt,MsgBox.Mess[2]) ;
END ;

procedure TFGdLivAux.ChoixEdition ;
{ Initialisation des options d'édition }
BEGIN
Inherited ;
DLRupt.PrintBefore:=TRUE ;
if CritEdt.Rupture=rLibres then
   BEGIN
   DLRupt.PrintBefore:=FALSE ;
   ChargeGroup(LRupt,['T00','T01','T02','T03','T04','T05','T06','T07','T08','T09']) ;
   END else
if CritEdt.Rupture=rRuptures then
   BEGIN
   ChargeRupt(LRupt, 'RUT', CritEdt.GL.PlanRupt, '', '') ;
   NiveauRupt(LRupt) ;
   END Else
if CritEdt.Rupture=rCorresp then
   BEGIN
   ChargeRuptCorresp(LRupt, CritEdt.GL.PlanRupt, '', '', False) ;
   NiveauRupt(LRupt) ;
   END ;
ChargeGroup(LMulti,[MsgBox.Mess[5],MsgBox.Mess[6]]) ;
ChgMaskChamp(Qr2DEBIT , CritEDT.Decimale, CritEDT.AfficheSymbole, CritEDT.Symbole, False) ;
ChgMaskChamp(Qr2CREDIT, CritEDT.Decimale, CritEDT.AfficheSymbole, CritEDT.Symbole, False) ;
//If CritEdt.GA then BDetailEcr.Height:=Sup1.Top+Sup1.Height else BDetailEcr.Height:=Sup1.Top ;
END ;

{***************** EVENEMENTS LIES A L'AFFICHAGE DES BANDES ********************}

procedure TFGdLivAux.BDetailEcrBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var i : Integer ; CptRupt : String ;
		IsECC 								: Boolean ;
begin
Inherited ;
PrintBand:=Affiche And (Not QR2E_AUXILIAIRE.IsNull) ;
if PrintBand then
   BEGIN
   IsECC	:= (FDevises.Value<>V_PGI.DevisePivot) and (FMontant.ITemIndex=1)
              and (QEcr.FindField('E_NATUREPIECE').AsString='ECC') ;
   If IsECC then
      BEGIN
      LE_DEBIT.Caption:='0' ;
      LE_CREDIT.Caption:='0' ;
      END ;
   If CritEdt.GA Then For i:=1 To 7 Do AlimListe(QEcr,i,Self) ;
   E_PIECELIGECH.Caption:=Qr2E_NUMEROPIECE.AsString+' / '+Qr2E_NUMLIGNE.AsString+'   '+Qr2E_NUMECHE.AsString ;
   AddGroup(LMulti, [Qr2E_DATECOMPTABLE, Qr2MULTIECHE], [Qr2DEBIT.AsFloat, Qr2CREDIT.AsFloat]) ;
   LE_VALIDE.Caption:=ValiQuali(QR2E_VALIDE.AsString,QR2E_QUALIFPIECE.AsString) ;
   { Affectation Du Calcul du Solde Progressif O/N Sinon, d'aprés l'info sur le compte }
   Case CritEdt.SoldeProg of
     0 : if Qr1T_SOLDEPROGRESSIF.AsString='X' then LE_SOLDE.Caption:=PrintSolde(ProgressDebit+Qr2DEBIT.asFloat,ProgressCredit+Qr2CREDIT.asFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)
                                              Else LE_SOLDE.Caption:=PrintSolde(Qr2DEBIT.asFloat,Qr2CREDIT.asFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
     1 : LE_SOLDE.Caption:=PrintSolde(ProgressDebit+Qr2DEBIT.asFloat,ProgressCredit+Qr2CREDIT.asFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
     2 : LE_SOLDE.Caption:=PrintSolde(Qr2DEBIT.asFloat,Qr2CREDIT.asFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
    end ;
   TotDesMvts[1].TotDebit :=Arrondi(TotDesMvts[1].TotDebit + Qr2DEBIT.AsFloat,CritEdt.Decimale) ;
   TotDesMvts[1].TotCredit:=Arrondi(TotDesMvts[1].TotCredit+Qr2CREDIT.AsFloat,CritEdt.Decimale) ;
   TotCptAux[1].TotDebit :=Arrondi(TotCptAux[1].TotDebit + Qr2DEBIT.AsFloat,CritEDT.Decimale) ;
   TotCptAux[1].TotCredit:=Arrondi(TotCptAux[1].TotCredit+Qr2CREDIT.AsFloat,CritEDT.Decimale) ;
   AddReport([1,2],CritEDT.GL.FormatPrint.Report,Qr2DEBIT.AsFloat,Qr2CREDIT.AsFloat,CritEDT.Decimale) ;

   Case CritEdt.Rupture of
     rLibres   : AddGroupLibre(LRupt,Q,fbAux,CritEdt.LibreTrie,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0]) ;
     rRuptures : AddRupt(LRupt,Qr2E_AUXILIAIRE.AsString,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0]) ;
     rCorresp  : BEGIN
                 Case CritEDT.GL.PlansCorresp Of
                   1 : If Qr1T_CORRESP1.AsString<>'' Then CptRupt:=Qr1T_CORRESP1.AsString+Qr2E_AUXILIAIRE.AsString
                                                     Else CptRupt:='.'+Qr2E_AUXILIAIRE.AsString ;
                   2 : If Qr1T_CORRESP2.AsString<>'' Then CptRupt:=Qr1T_CORRESP2.AsString+Qr2E_AUXILIAIRE.AsString
                                                     Else CptRupt:='.'+Qr2E_AUXILIAIRE.AsString ;
                   Else CptRupt:=Qr2E_AUXILIAIRE.AsString ;
                  End ;
                 AddRuptCorres(LRupt,CptRupt,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0])
                 END ;
     End ;
   END ;
Quoi:=QuoiMvt ;
end;

procedure TFGdLivAux.RecupCritEDT ;
Var NonLibres : Boolean ;
BEGIN
Inherited ;
With CritEDT Do
  BEGIN
  If FNumPiece1.Text<>'' then GL.NumPiece1:=StrToInt(FNumPiece1.Text) else GL.NumPiece1:=0 ;
  If FNumPiece2.Text<>'' then GL.NumPiece2:=StrToInt(FNumPiece2.Text) else GL.NumPiece2:=999999999 ;
  RefInterne:=FRefInterne.Text ;
  if FValide.State=cbGrayed then Valide:='g' ;
  if FValide.State=cbChecked then Valide:='X' ;
  if FValide.State=cbUnChecked then Valide:='-' ;
  if FSautPage.State=cbGrayed then SautPage:=0 ;
  if FSautPage.State=cbChecked then SautPage:=1 ;
  if FSautPage.State=cbUnChecked then SautPage:=2 ;
  if FTotMens.State=cbGrayed then TotMens:=0 ;
  if FTotMens.State=cbChecked then TotMens:=1 ;
  if FTotMens.State=cbUnChecked then TotMens:=2 ;
  if FSoldeProg.State=cbGrayed then SoldeProg:=0 ;
  if FSoldeProg.State=cbChecked then SoldeProg:=1 ;
  if FSoldeProg.State=cbUnChecked then SoldeProg:=2 ;
  if FLettrable.State=cbUnChecked then GL.Lettrable:=2 ;
  if FLettrable.State=cbChecked then GL.Lettrable:=1 ;
  if FLettrable.State=cbGrayed then GL.Lettrable:=0 ;
  GL.Lettrable:=0 ;
  GL.TotEche:=FTotEche.Checked ;
  QualifPiece:=FNatureEcr.Value ;
  GL.QuelAN:=FQuelAN(FAvecAN.Checked) ;
  GL.RuptOnly:=QuelleTypeRupt(1,FSAnsRupt.Checked,FAvecRupt.Checked,FALSE) ;

  NonLibres:=((Rupture=rRuptures) or (Rupture=rCorresp)) ;
  if NonLibres then GL.PlanRupt:=FPlanRuptures.Value ;
  If (CritEdt.Rupture=rCorresp) Then GL.PlansCorresp:=FPlanRuptures.ItemIndex+1 ;
  GL.OnlyCptAssocie:=(Rupture<>rRien) and FOnlyCptAssocie.Checked ;
  With GL.FormatPrint Do
     BEGIN
     PrSepCompte[2]:=FLigneCptPied.Checked ;
     PrSepCompte[3]:=FLigneRupt.Checked ;
     END ;
 END ;
END ;

function  TFGdLivAux.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ;
if Result then
   BEGIN
//  QBal:=PrepareTotCpt(fbAux,[],Dev,Etab,Exo,DevEnP) ;
   Gcalc:=TGCalculCum.create(Un,fbAux,fbaux,QuelTypeEcr,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE) ;
   GCalc.initCalcul('','','','',CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                    CritEdt.Date1,CritEdt.GL.Date11,TRUE) ;

   END ;
END ;

procedure TFGdLivAux.FormShow(Sender: TObject);
begin
HelpContext:=7418000 ;
//Standards.HelpContext:=7418010 ;
//Avances.HelpContext:=7418020 ;
//Mise.HelpContext:=7418030 ;
//Option.HelpContext:=7418040 ;
//TabRuptures.HelpContext:=7418050 ;
inherited ;
Floading:=FALSE ;
TabSup.TabVisible:=False;
TFGenJoker.Visible:=False;
If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
FRefInterne.Text:='' ; FLigneRupt.Enabled:=False ;
FValide.Checked:=True ; FValide.State:=cbGrayed ;
FJoker.MaxLength:=VH^.Cpta[fbAux].lg ;
FSurRupt.Visible:=False ;
FCodeRupt1.Visible:=False  ; FCodeRupt2.Visible:=False ;
TFCodeRupt1.Visible:=False ; TFCodeRupt2.Visible:=False ;
FOnlyCptAssocie.Enabled:=False ;
// GC - 20/12/2001
  if CritEdtChaine.Utiliser then InitEtatchaine('QRGLAUX');
// GC - 20/12/2001 - FIN
end;

procedure TFGdLivAux.QEcrCalcFields(DataSet: TDataSet);
{ Champs calculés dans le cas de ruptures mensuelles ou échéances }
begin
  inherited;
//Qr2MOIS.AsString:=FormatDateTime('mmmm yyyy',Qr2E_DATECOMPTABLE.AsDateTime) ;
//Qr2MULTIECHE.AsString:=IntToStr(Qr2E_NUMEROPIECE.AsInteger+Qr2E_NUMLIGNE.AsInteger) ;
//AncienEche:=Qr2E_NUMECHE.AsInteger ;
end;

procedure TFGdLivAux.BDetailEcrAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Case CritEdt.SoldeProg of
  1 :  Progressif(True,Qr2DEBIT.AsFloat,Qr2CREDIT.AsFloat) ;
  0 :  if Qr1T_SOLDEPROGRESSIF.AsString='X' then Progressif(True,Qr2DEBIT.AsFloat,Qr2CREDIT.AsFloat) ;
  End ;
end;

procedure TFGdLivAux.BFCompteAuxBeforePrint(var PrintBand: Boolean;  var Quoi: string);
{ Affichage D,C,S pour le compte et incrément sur le total général }
begin
  inherited;
T_AUXILIAIRE_.Caption:=MsgBox.Mess[11]+' '+Qr1T_AUXILIAIRE.AsString+' '+Qr1T_LIBELLE.AsString ;
PrintBand:=Affiche ;
If Not PrintBand then Exit ;
if PrintBand then
   BEGIN
   TotCptDebit.Caption:= AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,TotCptAux[1].TotDebit, CritEDT.AfficheSymbole) ;
   TotCptCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,TotCptAux[1].TotCredit, CritEDT.AfficheSymbole) ;
   TotSoldeCptAux.Caption:=PrintSolde(TotCptAux[1].TotDebit,TotCptAux[1].TotCredit,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
   TotDesMvtsDebit.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotDesMvts[1].TotDebit, CritEdt.AfficheSymbole) ;
   TotDesMvtsCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotDesMvts[1].TotCredit, CritEdt.AfficheSymbole) ;
   TotDesMvtsSolde.Caption:=PrintSolde(TotDesMvts[1].TotDebit,TotDesMvts[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
   TotEdt[1].TotDebit :=Arrondi(TotEdt[1].TotDebit+TotCptAux[1].TotDebit,CritEDT.Decimale) ;
   TotEdt[1].TotCredit:=Arrondi(TotEdt[1].TotCredit+TotCptAux[1].TotCredit,CritEDT.Decimale) ;
   END ;
Quoi:=QuoiAux(2) ;
end;

procedure TFGdLivAux.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
{ Affichage D,C,S pour le total général }
begin
  inherited;
TotGenDebit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,TotEdt[1].TotDebit, CritEDT.AfficheSymbole) ;
TotGenCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,TotEdt[1].TotCredit, CritEDT.AfficheSymbole) ;
TotalSoldes.Caption:=PrintSolde(TotEdt[1].TotDebit,TotEdt[1].TotCredit,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFGdLivAux.QRDLMultiNeedData(var MoreData: Boolean);
Var QuelleRupt : integer ;
begin
  inherited;
MoreData:=PrintGroup(LMulti, QEcr, [Qr2E_DATECOMPTABLE , Qr2MULTIECHE], CodRupt, LibRupt, Tot,Quellerupt) ;
TotMultiDebit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,Tot[0], CritEDT.AfficheSymbole) ;
TotMultiCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,Tot[1],CritEDT.AfficheSymbole) ;
If MoreData Then
   BEGIN
   if LibRupt=MsgBox.Mess[6] then
      BEGIN
      LibTotalMulti.Caption:=MsgBox.Mess[6]+' '+IntToStr(Qr2E_NUMEROPIECE.AsInteger)+' / '+Qr2E_NUMLIGNE.AsString ;
      BSDMulti.Font.Color:=clPurple ;
      TotSoldeMulti.Caption:='' ;
      if Qr2E_NUMECHE.AsInteger =1 then Imprime:=False
                                   else Imprime:=CritEDT.GL.TotEche ;
      END ;
   if LibRupt=MsgBox.Mess[5] then
      BEGIN
      LibTotalMulti.Caption:=MsgBox.Mess[7]+' '+CodRupt ;
      BSDMulti.Font.Color:=clRed ;
      TotSoldeMulti.Caption:=PrintSolde(Tot[0],Tot[1], CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
      (* Rony 10/04/97 -- Pas de Progerssif
      Case CritEdt.SoldeProg of
        0 : if Qr1T_SOLDEPROGRESSIF.AsString='X' then TotSoldeMulti.Caption:=PrintSolde(ProgressDebit,ProgressCredit,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)
                                                 Else TotSoldeMulti.Caption:=PrintSolde(Tot[0],Tot[1], CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
        1 : TotSoldeMulti.Caption:=PrintSolde(ProgressDebit,ProgressCredit,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
        2 : TotSoldeMulti.Caption:=PrintSolde(Tot[0],Tot[1], CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
       end ;
       *)
      Case CritEdt.TotMens of             { Total Mensuel O/N Sinon, d'aprés l'info sur le compte }
       0 : Imprime:=(Qr1T_TOTAUXMENSUELS.AsString='X') ;
       1 : Imprime:=True  ;
       2 : Imprime:=False  ;
       end;
      END ;
   END ;

//ShowMessage( FloatToStr(Tot[0])+'  '+FloatToStr(Tot[1])+ #13+FloatToStr(Tot[2])+'  '+FloatToStr(Tot[0]-Tot[1]) ) ;
end;

procedure TFGdLivAux.BSDMultiBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=Imprime And Affiche ;
end;

(*
Procedure CalculEtAfficheMontantRupture ;

AnoRupt,AnoDebitRupt,AnoCreditRupt,AnoSoldeRupt,
CumRupt,CumDebitRupt,CumCreditRupt,CumSoldeRupt,
TotRupt,DebitRupt,CreditRupt,SoldeRupt : TQrLabel ;
TotRuptM : TTTotRupt ;
BRupt : TQRBand ;
BEGIN

AnoRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
AnoDebitRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
AnoCreditRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
AnoSoldeRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;

CumDebitRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
CumCreditRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
CumSoldeRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
CumRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
Case CritEdt.Rupture Of
  rCorresp :  BEGIN
              OkEnteteRup:=True ;
              AnoDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[3], CritEdt.AfficheSymbole) ;
              AnoCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[4], CritEdt.AfficheSymbole) ;
              AnoSoldeRupt.Caption:=PrintSolde(TotRupt[3], TotRupt[4], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

              CumDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[5], CritEdt.AfficheSymbole) ;
              CumCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[6], CritEdt.AfficheSymbole) ;
              CumSoldeRupt.Caption:=PrintSolde(TotRupt[5], TotRupt[6], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

              SumD:=Arrondi(TotRupt[5]+ TotRupt[1],CritEdt.Decimale) ;
              SumC:=Arrondi(TotRupt[6]+TotRupt[2],CritEdt.Decimale) ;
              TCodRupt.Top:=37 ; DebitRupt.Top:=37 ; CreditRupt.Top:=37 ; SoldeRupt.Top:=37 ;
              BRupt.Height:=54 ;
              END ;
  rlibres  :  BEGIN
              OkEnteteRup:=True ;
              AnoDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[3], CritEdt.AfficheSymbole) ;
              AnoCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[4], CritEdt.AfficheSymbole) ;
              AnoSoldeRupt.Caption:=PrintSolde(TotRupt[3], TotRupt[4], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

              CumDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[1], CritEdt.AfficheSymbole) ;
              CumCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[2], CritEdt.AfficheSymbole) ;
              CumSoldeRupt.Caption:=PrintSolde(TotRupt[1], TotRupt[2], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

              SumD:=Arrondi(TotRupt[3]+ TotRupt[1],CritEdt.Decimale) ;
              SumC:=Arrondi(TotRupt[4]+TotRupt[2],CritEdt.Decimale) ;
              END ;
   Else BEGIN
        TCodRupt.Top:=2 ; DebitRupt.Top:=2 ; CreditRupt.Top:=2 ; SoldeRupt.Top:=2;
        If (CritEdt.Rupture<>rlibres) Then BRupt.Height:=20 ;
        END;
  END ;
If (CritEdt.GL.QuelAN=SansAN) Then
   BEGIN
   AnoDebitRupt.Caption:=MsgBox.Mess[9] ; AnoCreditRupt.Caption:='' ; AnoSoldeRupt.Caption:='' ;
   END ;
DebitRupt.Caption:= AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumD, CritEDT.AfficheSymbole) ;
CreditRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumC, CritEDT.AfficheSymbole) ;
SoldeRupt.Caption:=PrintSolde(SumD, SumC, CritEDT.Decimale, CritEDT.Symbole, CritEDT.AfficheSymbole) ;
END ;
*)
procedure TFGdLivAux.DLRuptNeedData(var MoreData: Boolean);
Var TotRupt     : TTotRupt ;
    SumD, SumC  : Double ;
    Quellerupt  : Integer ;
    OkOk        : Boolean ;
    CptRupt,Lib1, Stcode  : String ;
    Col         : TColor ;
    LibRuptInf : Array[1..10] Of TRuptInf ;
begin
  inherited;
MoreData:=False ;
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
  rRuptures : MoreData:=PrintRupt(LRupt,Qr2E_AUXILIAIRE.AsString,CodRupt,LibRupt,DansTotGen,QRP.EnRupture,TotRupt) ;
  rCorresp  : BEGIN
              OkOk:=TRUE ;
              Case CritEDT.GL.PlansCorresp  Of
                1 : If Qr1T_CORRESP1.AsString<>'' Then CptRupt:=Qr1T_CORRESP1.AsString+Qr2E_AUXILIAIRE.AsString
                                                  Else CptRupt:='.'+Qr2E_AUXILIAIRE.AsString ;
                2 : If Qr1T_CORRESP2.AsString<>'' Then CptRupt:=Qr1T_CORRESP2.AsString+Qr2E_AUXILIAIRE.AsString
                                                  Else CptRupt:='.'+Qr2E_AUXILIAIRE.AsString ;
                Else OkOk:=FALSE ;
                END ;
              If OkOk Then MoreData:=PrintRupt(LRupt,CptRupt,CodRupt,LibRupt,DansTotGen,QRP.EnRupture,TotRupt) Else MoreData:=FALSE ;
              END ;
  End ;
If MoreData Then
   BEGIN
   TCodRupt.Width:=TLE_DATECOMPTABLE.Width+TE_JOURNAL.Width+TE_PIECE.Width+TE_REFINTERNE.Width+1 ;
   TCodRupt.Caption:='' ;
   BRupt.Height:=HauteurBandeRuptIni ;
   Case CritEdt.Rupture of
     rLibres   : BEGIN
                 insert(MsgBox.Mess[15]+' ',CodRupt,Quellerupt+2) ;
                 TCodRupt.Caption:=CodRupt+' '+Lib1 ;
                 AlimRuptSup(HauteurBandeRuptIni,QuelleRupt,TCodRupt.Width,BRupt,LibRuptInf,Self) ;
                 END ;
     rRuptures, rCorresp : TCodRupt.Caption:=CodRupt+'  '+LibRupt ;
     End ;
   SumD:=Arrondi(TotAnCRupt[1].TotDebit+ TotRupt[1],CritEDT.Decimale) ;
   SumC:=Arrondi(TotAnCRupt[1].TotCredit+TotRupt[2],CritEDT.Decimale) ;
   (*
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
      TCodRupt.Top:=37 ; DebitRupt.Top:=37 ; CreditRupt.Top:=37 ; SoldeRupt.Top:=37 ;
      BRupt.Height:=54 ;
      END Else
      BEGIN
      TCodRupt.Top:=2 ; DebitRupt.Top:=2 ; CreditRupt.Top:=2 ; SoldeRupt.Top:=2;
      If (CritEdt.Rupture<>rlibres) Then BRupt.Height:=20 ;
      END;
   *)

   AnoRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
   AnoDebitRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
   AnoCreditRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
   AnoSoldeRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;

   CumDebitRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
   CumCreditRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
   CumSoldeRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
   CumRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
   Case CritEdt.Rupture Of
     rCorresp :  BEGIN
                 OkEnteteRup:=True ;
                 AnoDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[3], CritEdt.AfficheSymbole) ;
                 AnoCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[4], CritEdt.AfficheSymbole) ;
                 AnoSoldeRupt.Caption:=PrintSolde(TotRupt[3], TotRupt[4], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

                 CumDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[5], CritEdt.AfficheSymbole) ;
                 CumCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[6], CritEdt.AfficheSymbole) ;
                 CumSoldeRupt.Caption:=PrintSolde(TotRupt[5], TotRupt[6], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

                 SumD:=Arrondi(TotRupt[5]+ TotRupt[1],CritEdt.Decimale) ;
                 SumC:=Arrondi(TotRupt[6]+TotRupt[2],CritEdt.Decimale) ;
                 TCodRupt.Top:=37 ; DebitRupt.Top:=37 ; CreditRupt.Top:=37 ; SoldeRupt.Top:=37 ;
                 BRupt.Height:=54 ;
                 END ;
     rlibres  :  BEGIN
                 OkEnteteRup:=True ;
                 AnoDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[3], CritEdt.AfficheSymbole) ;
                 AnoCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[4], CritEdt.AfficheSymbole) ;
                 AnoSoldeRupt.Caption:=PrintSolde(TotRupt[3], TotRupt[4], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

                 CumDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[1], CritEdt.AfficheSymbole) ;
                 CumCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[2], CritEdt.AfficheSymbole) ;
                 CumSoldeRupt.Caption:=PrintSolde(TotRupt[1], TotRupt[2], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

                 SumD:=Arrondi(TotRupt[3]+ TotRupt[1],CritEdt.Decimale) ;
                 SumC:=Arrondi(TotRupt[4]+TotRupt[2],CritEdt.Decimale) ;
                 (*
                 TCodRupt.Top:=2 ; DebitRupt.Top:=2 ; CreditRupt.Top:=2 ; SoldeRupt.Top:=2;
                 *)
                 (*
                 TCodRupt.Top:=37 ; DebitRupt.Top:=37 ; CreditRupt.Top:=37 ; SoldeRupt.Top:=37 ;
                 BRupt.Height:=54 ;
                 *)
                 END ;
      Else BEGIN
           TCodRupt.Top:=2 ; DebitRupt.Top:=2 ; CreditRupt.Top:=2 ; SoldeRupt.Top:=2;
           If (CritEdt.Rupture<>rlibres) Then BRupt.Height:=20 ;
           END;
     END ;
   If (CritEdt.GL.QuelAN=SansAN) Then
      BEGIN
      AnoDebitRupt.Caption:=MsgBox.Mess[9] ; AnoCreditRupt.Caption:='' ; AnoSoldeRupt.Caption:='' ;
      END ;
   DebitRupt.Caption:= AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumD, CritEDT.AfficheSymbole) ;
   CreditRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumC, CritEDT.AfficheSymbole) ;
   SoldeRupt.Caption:=PrintSolde(SumD, SumC, CritEDT.Decimale, CritEDT.Symbole, CritEDT.AfficheSymbole) ;
   SautPageRupture:=CritEdt.SautPageRupt And (CritEdt.GL.RuptOnly=Avec) And (CritEdt.SautPage<>1) And SautPageRuptAFaire(CritEdt,BDetail,QuelleRupt) ;
   END ;
end;

procedure TFGdLivAux.BRuptBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.Rupture<>rRien) ;
end;

procedure TFGdLivAux.BRuptAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Fillchar(TotAncRupt,SizeOf(TotAncRupt),#0) ;
end;

procedure TFGdLivAux.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
Titre1Rep.Caption:=Titre2Rep.Caption ;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
Report1Solde.Caption:=Report2Solde.Caption ;
end;

procedure TFGdLivAux.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var D,C : Double ;
begin
  inherited;
Case QuelReport(CritEDT.GL.FormatPrint.Report,D,C) Of
  1 : BEGIN TITREREPORTB.Caption:=MsgBox.Mess[3] ; Titre2Rep.Caption:='' ; END ;
  2 : BEGIN TITREREPORTB.Caption:=MsgBox.Mess[4] ; Titre2Rep.Caption:=StReportAux ; END ;
  END ;
Report2Debit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, D, CritEDT.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, C, CritEDT.AfficheSymbole ) ;
Report2Solde.Caption:=PrintSolde(D,C,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
end;

procedure TFGdLivAux.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var CumulAu : TabTot ;
    D,C   : Double ; CptRupt : String ;
begin
  inherited;
T_AUXILIAIRE.Caption:=MsgBox.Mess[10]+' '+Qr1T_AUXILIAIRE.AsString+' '+Qr1T_LIBELLE.AsString ;
Fillchar(TotDesMvts,SizeOf(TotDesMvts),#0) ;
Fillchar(TotCptAux,SizeOf(TotCptAux),#0) ;
Fillchar(CumulAu,SizeOf(CumulAu),#0) ;
Case CritEdt.Rupture of
  rLibres    : if CritEdt.GL.OnlyCptAssocie then PrintBand:=DansRuptLibre(Q,fbAux,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
  rRuptures  : if CritEdt.GL.OnlyCptAssocie then PrintBand:=DansRupt(LRupt,Qr1T_AUXILIAIRE.AsString) ;
  rCorresp   : if CritEdt.GL.OnlyCptAssocie then
                 if CritEDT.GL.PlansCorresp=1 then PrintBand:=(Qr1T_CORRESP1.AsString<>'') Else
                 if CritEDT.GL.PlansCorresp=2 then PrintBand:=(Qr1T_CORRESP2.AsString<>'') ;
  End;
Affiche:=PrintBand ;
If Not PrintBand then Exit ;
if PrintBand then
   BEGIN
   StReportAux:=Qr1T_AUXILIAIRE.AsString ;
   InitReport([2],CritEDT.GL.FormatPrint.Report) ;
   Case CritEdt.SautPage of
     0 : BDetail.forceNewPage:=(Qr1T_SAUTPAGE.AsString='X') ;
     1 : BDetail.forceNewPage:=True ;
     2 : BDetail.forceNewPage:=False ;
    end ;
   (*
   ExecuteTotCpt(QBal,QT_AUXILIAIRE.AsString, CritEDT.Date1,CritEDT.GL.Date11,
                 CritEDT.DeviseSelect,CritEDT.Etab,CritEDT.Exo.Code,CumulAu,FALSE) ;
   *)
   If (CritEdt.GL.QuelAN<>SansAN) Then
      BEGIN
      GCAlc.ReInitCalcul(Qr1T_AUXILIAIRE.AsString,'',0,0) ;
      GCalc.Calcul ; CumulAu:=GCalc.ExecCalc.TotCpt ;
      END ;
   If CritEdt.GL.QuelAN=SansAN Then BEGIN CumulAu[0].TotDebit:=0 ; CumulAu[0].TotCredit:=0 ; END ;
   CumulVersSolde(CumulAu[0]) ;
   if CritEDT.Date1=CritEDT.GL.Date11 Then Fillchar(CumulAu[1],SizeOf(CumulAu[1]),#0) ;
   D:=CumulAu[0].TotDebit+CumulAu[1].TotDebit ;
   C:=CumulAu[0].TotCredit+CumulAu[1].TotCredit ;

   If (CritEdt.GL.QuelAN<>SansAN) Then
      BEGIN
      { A Nouveau }
      AnvDebit.Caption:= AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,CumulAu[0].TotDebit, CritEDT.AfficheSymbole) ;
      AnvCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,CumulAu[0].TotCredit, CritEDT.AfficheSymbole) ;
      AnvSolde.Caption:= PrintSolde(CumulAu[0].TotDebit,CumulAu[0].TotCredit,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;

      { A Nouveau + Cumul }
      CumDebit.Caption:= AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,D, CritEDT.AfficheSymbole) ;
      CumCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,C,CritEDT.AfficheSymbole) ;
      CumSolde.Caption:= PrintSolde(D,C,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
      { Totaux tiers et rupture }
      TotCptAux[1].TotDebit:= Arrondi(TotCptAux[1].TotDebit+D,CritEDT.Decimale) ;
      TotCptAux[1].TotCredit:=Arrondi(TotCptAux[1].TotCredit+C,CritEDT.Decimale) ;
      TotAnCRupt[1].TotDebit:= Arrondi(TotAnCRupt[1].TotDebit+D,CritEDT.Decimale) ;
      TotAnCRupt[1].TotCredit:=Arrondi(TotAnCRupt[1].TotCredit+C,CritEDT.Decimale) ;
      END Else
      BEGIN
      AnvDebit.Caption:=MsgBox.Mess[9] ; AnvCredit.Caption:='' ; AnvSolde.Caption:='' ;
      CumDebit.Caption:=MsgBox.Mess[9] ; CumCredit.Caption:='' ;  CumSolde.Caption:='' ;
      END ;


   { Ajout (ANouveau + Cuml) dans le solde progr. }
   { Init du Cumul Progressif }
   Case CritEdt.SoldeProg of
     1 : BEGIN
                 Progressif(False,0,0) ; { Si Progressif  obligatoire alors initialise Solde Progressif }
                 Progressif(True,D,C) ;
                 END ;
     0 : if Qr1T_SOLDEPROGRESSIF.AsString='X' then
                    BEGIN
                    Progressif(False,0,0) ; { Si Progressif ou non alors seulement si demandé par le compte }
                    Progressif(True,D,C) ;
                    END ;
    end ;
   AddReport([1,2],CritEDT.GL.FormatPrint.Report,D,C,CritEDT.Decimale) ;
   Case CritEdt.Rupture of
     rLibres   : AddGroupLibre(LRupt,Q,fbAux,CritEdt.LibreTrie,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
     rRuptures : AddRupt(Lrupt,Qr1T_AUXILIAIRE.AsString,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
     rCorresp  : BEGIN
                 Case CritEDT.GL.PlansCorresp Of
                   1 : If Qr1T_CORRESP1.AsString<>'' Then CptRupt:=Qr1T_CORRESP1.AsString+Qr2E_AUXILIAIRE.AsString
                                                     Else CptRupt:='.'+Qr2E_AUXILIAIRE.AsString ;
                   2 : If Qr1T_CORRESP2.AsString<>'' Then CptRupt:=Qr1T_CORRESP2.AsString+Qr2E_AUXILIAIRE.AsString
                                                     Else CptRupt:='.'+Qr2E_AUXILIAIRE.AsString ;
                  End ;
                 AddRuptCorres(LRupt,CptRupt,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
                 END ;
     End ;
   END ;
Quoi:=QuoiAux(1) ;
end;

procedure TFGdLivAux.QEcrAfterOpen(DataSet: TDataSet);
begin
  inherited;
//
QR2E_AUXILIAIRE:=TStringField(QEcr.FindField('E_AUXILIAIRE')) ;
Qr2E_JOURNAL:=TStringField(QEcr.FindField('E_JOURNAL')) ;
Qr2E_VALIDE:=TStringField(QEcr.FindField('E_VALIDE')) ;
Qr2E_QUALIFPIECE:=TStringField(QEcr.FindField('E_QUALIFPIECE')) ;
Qr2E_EXERCICE:=TStringField(QEcr.FindField('E_EXERCICE')) ;
Qr2E_NUMEROPIECE:=TIntegerField(QEcr.FindField('E_NUMEROPIECE')) ;
Qr2E_NUMLIGNE:=TIntegerField(QEcr.FindField('E_NUMLIGNE')) ;
Qr2E_NUMECHE:=TIntegerField(QEcr.FindField('E_NUMECHE')) ;
Qr2E_DateComptable:=TDateTimeField(QEcr.FindField('E_DATECOMPTABLE')) ;
Qr2E_DATECOMPTABLE.Tag:=1 ;
Qr2DEBIT:=TFloatField(QEcr.FindField('DEBIT')) ;
Qr2CREDIT:=TFloatField(QEcr.FindField('CREDIT')) ;
QR2MULTIECHE:=TIntegerField(QEcr.FindField('MULTIECHE')) ;
//IsECC:=(FDevises.Value<>V_PGI.DevisePivot)and(FMontant.ITemIndex=1)and(QEcr.FindField('E_NATUREPIECE').AsString='ECC') ;
ChgMaskChamp(Qr2DEBIT , CritEDT.Decimale, CritEDT.AfficheSymbole, CritEDT.Symbole, False) ;
ChgMaskChamp(Qr2CREDIT, CritEDT.Decimale, CritEDT.AfficheSymbole, CritEDT.Symbole, False) ;
end;


(*
FPlanRupt.visible:=(Not FSansRupt.Checked) And (Not FJoker.Visible) ;
FPlanRupt.Enabled:=FPlanRupt.visible ;
TFPlanRupt.Visible:=FPlanRupt.visible ;

*)
procedure TFGdLivAux.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr1T_AUXILIAIRE     :=TStringField(Q.FindField('T_AUXILIAIRE'));
Qr1T_LIBELLE        :=TStringField(Q.FindField('T_LIBELLE'));
Qr1T_SAUTPAGE       :=TStringField(Q.FindField('T_SAUTPAGE')) ;
Qr1T_SOLDEPROGRESSIF:=TStringField(Q.FindField('T_SOLDEPROGRESSIF')) ;
Qr1T_TOTAUXMENSUELS :=TStringField(Q.FindField('T_TOTAUXMENSUELS')) ;
If CritEDT.Rupture=rCorresp then
   BEGIN
   Qr1T_CORRESP1         :=TStringField(Q.FindField('T_CORRESP1'));
   Qr1T_CORRESP2         :=TStringField(Q.FindField('T_CORRESP2'));
   END ;
end;

procedure TFGdLivAux.FSelectCpteChange(Sender: TObject);
begin
  inherited;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
end;

procedure TFGdLivAux.FExerciceChange(Sender: TObject);
begin
FLoading:=TRUE ;
  inherited;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
FLoading:=FALSE ;
end;

procedure TFGdLivAux.FDateCompta1Change(Sender: TObject);
begin
  inherited;
If FLoading Then Exit ;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
end;

procedure TFGdLivAux.BFCompteAuxAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([2],CritEDT.GL.FormatPrint.Report) ;
end;

procedure TFGdLivAux.DLRuptHNeedData(var MoreData: Boolean);
begin
  inherited;
(* En-tête de rupture/corresp
MoreData:=False ;
if OkEnteteRup then
   If CritEdt.Rupture=rCorresp then
      BEGIN
      Case CritEdt.Bal.PlansCorresp of
        1 : BEGIN MoreData:=Qr1T_CORRESP1.AsString<>'' ;  END ;
        2 : BEGIN MoreData:=Qr1T_CORRESP2.AsString<>'' ;  END ;
        End ;
      END ;
OkEnteteRup:=False ;
*)
end;


procedure TFGdLivAux.BRupTeteBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
Case CritEdt.GL.PlansCorresp of
  1 : CodeRuptAu.Caption:=Qr1T_CORRESP1.AsString+'x' ;
  2 : CodeRuptAu.Caption:=Qr1T_CORRESP2.AsString+'x' ;
  End ;
end;

procedure TFGdLivAux.FSansRuptClick(Sender: TObject);
begin
  inherited;
FLigneRupt.Enabled:=Not FSansRupt.Checked ;
FLigneRupt.checked:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
FRupturesClick(Nil) ;
end;

procedure TFGdLivAux.FRupturesClick(Sender: TObject);
begin
  inherited;
if FPlansCo.Checked then FGroupRuptures.Caption:=' '+MsgBox.Mess[14] ;
if FRuptures.Checked then If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
end;

procedure TFGdLivAux.BDetailAfterPrint(BandPrinted: Boolean);
begin
  inherited;
If CritEdt.SautPageRupt And (CritEdt.GL.RuptOnly=Avec) And (CritEdt.SautPage<>1) Then
  BEGIN
  BDetail.ForceNewPage:=FALSE ; SautPageRupture:=FALSE ;
  END ;
end;

procedure TFGdLivAux.BDetailCheckForSpace;
begin
  inherited;
If SautPageRupture Then BDetail.ForceNewPage:=TRUE ;
end;

procedure TFGdLivAux.Timer1Timer(Sender: TObject);
begin
// GC - 20/12/2001
  inherited;
// GC - 20/12/2001 - FIN
end;

end.
