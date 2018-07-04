unit QRGLGen;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, Ent1, Hent1, CritEdt, UtilEdt, HQry, QRRupt,
  CpteUtil,
  EdtLegal,
  Menus, HSysMenu, HTB97, HPanel, UiUtil, tCalcCum, HRichEdt ;

  procedure GLGeneral ;
  // GC - 20/12/2001
  procedure GLGeneralChaine( vCritEdtChaine : TCritEdtChaine );
  // GC - 20/12/2001 - FIN
  procedure GLGeneralZoom(Crit : TCritEdt) ;
  procedure GLGeneralLegal(Crit : TCritEdt) ;

type
  TGdLivGen = class(TFQR)
    FValide: TCheckBox;
    HLabel3: THLabel;
    FNumPiece1: TMaskEdit;
    Label12: TLabel;
    FNumPiece2: TMaskEdit;
    HLabel9: THLabel;
    FRefInterne: TEdit;
    FSautPage: TCheckBox;
    FLigneCptPied: TCheckBox;
    QRLabel17: TQRLabel;
    RValide: TQRLabel;
    TRLegende: TQRLabel;
    RLegende: TQRLabel;
    QRLabel13: TQRLabel;
    RNumPiece1: TQRLabel;
    QRLabel14: TQRLabel;
    RNumPiece2: TQRLabel;
    QRLabel10: TQRLabel;
    RRefInterne: TQRLabel;
    TLE_VALIDE: TQRLabel;
    TLE_DATECOMPTABLE: TQRLabel;
    TE_JOURNAL: TQRLabel;
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
    G_GENERAL: TQRLabel;
    QRLabel3: TQRLabel;
    TCumul: TQRLabel;
    AnvDebit: TQRLabel;
    CumDebit: TQRLabel;
    CumCredit: TQRLabel;
    AnvCredit: TQRLabel;
    AnvSolde: TQRLabel;
    CumSolde: TQRLabel;
    BSDMvt: TQRBand;
    LE_DATECOMPTABLE: TQRDBText;
    LE_CREDIT: TQRDBText;
    LE_DEBIT: TQRDBText;
    E_LIBELLE: TQRDBText;
    E_REFINTERNE: TQRDBText;
    LE_LETTRAGE: TQRDBText;
    LE_SOLDE: TQRLabel;
    LE_VALIDE: TQRLabel;
    LE_JOURNAL: TQRDBText;
    E_PIECELIGECH: TQRLabel;
    BSDMulti: TQRBand;
    LibTotalMulti: TQRLabel;
    TotMultiDebit: TQRLabel;
    TotMultiCredit: TQRLabel;
    TotSoldeMulti: TQRLabel;
    BFCpteGen: TQRBand;
    G_GENERAL_: TQRLabel;
    TotGenSolde: TQRLabel;
    TotGenCredit: TQRLabel;
    TotGenDebit: TQRLabel;
    BRupt: TQRBand;
    TCodRupt: TQRLabel;
    DebitRupt: TQRLabel;
    CreditRupt: TQRLabel;
    SoldeRupt: TQRLabel;
    QRLabel33: TQRLabel;
    TotalDebit: TQRLabel;
    TotalCredit: TQRLabel;
    TotalSoldes: TQRLabel;
    QRBand1: TQRBand;
    Trait0: TQRLigne;
    Trait1: TQRLigne;
    Trait2: TQRLigne;
    Trait3: TQRLigne;
    Trait4: TQRLigne;
    Trait5: TQRLigne;
    Ligne1: TQRLigne;
    TITRE2REP: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    REPORT2SOLDE: TQRLabel;
    SEcr: TDataSource;
    QEcr: TQuery;
    QRDLMvt: TQRDetailLink;
    QRDLRupt: TQRDetailLink;
    QRDLMulti: TQRDetailLink;
    MsgBox: THMsgBox;
    BCentr: TQRBand;
    CodCentr: TQRLabel;
    QRDLCENTR: TQRDetailLink;
    TotGenDCentr: TQRLabel;
    TotGenCCentr: TQRLabel;
    TotGenSCentr: TQRLabel;
    TitreCentr: TQRLabel;
    Invisible: TQRLabel;
    TitreTotGenMvt: TQRLabel;
    TotGenMvtDebit: TQRLabel;
    TotGenMvtCredit: TQRLabel;
    TotGenMvtSolde: TQRLabel;
    TotJalNonCentral: TQRLabel;
    TotJcentDebit: TQRLabel;
    TotJcentCredit: TQRLabel;
    TotJcentSolde: TQRLabel;
    FSoldeProg: TCheckBox;
    FTotMens: TCheckBox;
    FCollectif: TCheckBox;
    FAvecAN: TCheckBox;
    FCentralisation: TCheckBox;
    FLigneRupt: TCheckBox;
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
    FPointable: TCheckBox;
    QRLabel1: TQRLabel;
    RLettrable: TQRLabel;
    QRLabel8: TQRLabel;
    RPointable: TQRLabel;
    BBil: TQRBand;
    BGes: TQRBand;
    DLBil: TQRDetailLink;
    BBilGen: TQRBand;
    BGesGen: TQRBand;
    TTotBil: TQRLabel;
    TotBilDebit: TQRLabel;
    TotBilCredit: TQRLabel;
    TotBilSolde: TQRLabel;
    TTotGes: TQRLabel;
    TotGesDebit: TQRLabel;
    TotGesCredit: TQRLabel;
    TotGesSolde: TQRLabel;
    DLGes: TQRDetailLink;
    FBilGes: TCheckBox;
    TTotBilGen: TQRLabel;
    TotGenBilDebit: TQRLabel;
    TotGenBilCredit: TQRLabel;
    TotGenBilSolde: TQRLabel;
    TTotGesGen: TQRLabel;
    TotGenGesDebit: TQRLabel;
    TotGenGesCredit: TQRLabel;
    TotGenGesSolde: TQRLabel;
    TotRupt: TQRLabel;
    FDetailCentralise: TCheckBox;
    QRRichEdit1: TQRRichEdit;
    FHRichEdit1: THRichEdit;
    FAvecMemo: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BSDMvtBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BSDMvtAfterPrint(BandPrinted: Boolean);
    procedure BFCpteGenBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QRDLRuptNeedData(var MoreData: Boolean);
    procedure BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BRuptAfterPrint(BandPrinted: Boolean);
    procedure BSDMultiBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QRDLMultiNeedData(var MoreData: Boolean);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QEcrAfterOpen(DataSet: TDataSet);
    procedure BTitreAfterPrint(BandPrinted: Boolean);
    procedure QRDLCENTRNeedData(var MoreData: Boolean);
    procedure BCentrBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BCentrAfterPrint(BandPrinted: Boolean);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure FSelectCpteChange(Sender: TObject);
    procedure FExerciceChange(Sender: TObject);
    procedure FDateCompta1Change(Sender: TObject);
    procedure BFCpteGenAfterPrint(BandPrinted: Boolean);
    procedure DLRuptHNeedData(var MoreData: Boolean);
    procedure BRupTeteBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FSansRuptClick(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure DLBilNeedData(var MoreData: Boolean);
    procedure BGesBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure DLGesNeedData(var MoreData: Boolean);
    procedure Timer1Timer(Sender: TObject);
    procedure FCentralisationClick(Sender: TObject);
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure GenereSQL ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
  private    { Déclarations privées }
    DansTotGen, IsECC, RepotCent : Boolean ;
    TotCentr : Array[0..12] Of Double ;
    StReportGen               : string ;
    //QBal                             : TQuery ;
    QBalC : TQuery ;
    TotCptGenMvt,TotEdt,TotCptGen,
    TotAnCRupt, TotJalCent   : TabTot ;
    TotGes                   : TabTot ;
    L, LRupt, LMulti, LCentr, LBilan,
    LBil,LCha,LPro,LExt : TStringList ;
    Qr1G_GENERAL             : TStringField ;
    Qr1G_LIBELLE             : TStringField ;
    Qr1G_COLLECTIF           : TStringField ;
    Qr1G_SAUTPAGE,Qr1G_SOLDEPROGRESSIF,Qr1G_TOTAUXMENSUELS,Qr1G_CENTRALISABLE : TStringField ;
    QR2E_GENERAL, QR2E_EXERCICE,
    QR2J_LIBELLE, QR2E_REFINTERNE,
    QR2E_ETABLISSEMENT, QR2E_LIBELLE,
    QR2E_DEVISE, QR2E_VALIDE, QR2E_QUALIFPIECE,
    QR2E_LETTRAGE, QR2E_JOURNAL,
    QR2J_CENTRALISABLE, QR2E_NATUREPIECE : TStringField ;
    QR2E_DATECOMPTABLE : TDateTimeField ;
    QR2E_NUMEROPIECE, QR2E_NUMLIGNE   : TIntegerField ;
    QR2DEBIT, QR2CREDIT               : TFloatField ;
    QR2E_BLOCNOTE : TBlobField;
    FLoading, IsLegal, Affiche        : Boolean ;
    OkEnteteRupt, FirstPassage, OkGes : Boolean ;
    Qr1G_CORRESP1, Qr1G_CORRESP2 : TStringField ;
    Function  QuoiGen(i : Integer) : String ;
    Function  QuoiMvt : String ;
    Procedure GLGenZoom(Quoi : String) ;
    Procedure GenereSQLSub ;
    procedure RecapPourQui(Cpt : String ; LQui : TStringList ; MO : Array of double) ;
    Procedure CalculGestion(MO : Array of double ) ;
  public     { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Function TGdLivGen.QuoiGen(i : Integer) : String ;
BEGIN
Case i Of
  1 : Result:=QR1G_GENERAL.AsString+' '+'#'+QR1G_LIBELLE.AsString+'@'+'1' ;
  2 : Result:=QR1G_GENERAL.AsString+' '+'#'+QR1G_LIBELLE.AsString+' '+
              TotGenSolde.Caption+'@'+'1' ;
  END ;
END ;

Function TGdLivGen.QuoiMvt : String ;
BEGIN
Result:=QR1G_GENERAL.AsString+' '+QR1G_LIBELLE.AsString+' '+Le_Solde.Caption+        '#'+QR2E_JOURNAL.AsString+' N° '+IntToStr(QR2E_NUMEROPIECE.AsInteger)+' '+DateToStr(QR2E_DateComptable.AsDAteTime)+' . . . '+
        PrintSolde(QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole)+
        '@'+'5;'+QR2E_JOURNAL.AsString+';'+UsDateTime(QR2E_DATECOMPTABLE.AsDateTime)+';'+QR2E_NUMEROPIECE.AsString+';'+QR2E_EXERCICE.asString+';'+
        IntToStr(QR2E_NumLigne.AsInteger)+';' ;
END ;

Procedure TGdLivGen.GLGenZoom(Quoi : String) ;
Var Lp,i: Integer ;
    St : String ;
BEGIN
Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
St:=Copy(Quoi,Lp+1,1) ;
If St='!' Then i:=100 Else i:=StrToInt(St) ;
If (i=5) Or (i=100) Then
   BEGIN
   Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
   If QRP.QrPrinter.FSynShiftDblClick Then i:=6 ;
   END ;
ZoomEdt(i,Quoi) ;
END ;

procedure GLGeneral ;
var QR : TGdLivGen ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TGdLivGen.Create(Application) ;
Edition.Etat:=etGlGen ;
QR.QRP.QRPrinter.OnSynZoom:=QR.GLGenZoom ;
QR.IsLegal:=FALSE ;
QR.InitType (nbGen,neGL,msGenEcr,'QRGLGENE','QR_GLGENE',TRUE,FALSE,FALSE,Edition) ;
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
procedure GLGeneralChaine( vCritEdtChaine : TCritEdtChaine );
var QR : TGdLivGen ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TGdLivGen.Create(Application) ;
Edition.Etat:=etGlGen ;
QR.QRP.QRPrinter.OnSynZoom:=QR.GLGenZoom ;
QR.IsLegal:=FALSE ;
QR.CRitEdtChaine := vCritEdtChaine;
QR.InitType (nbGen,neGL,msGenEcr,'QRGLGENE','QR_GLGENE',TRUE,FALSE,FALSE,Edition) ;
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
// GC - 20/12/2001 - FIN

procedure GLGeneralZoom(Crit : TCritEdt) ;
var QR : TGdLivGen ;
    Edition : TEdition ;
BEGIN
QR:=TGdLivGen.Create(Application) ;
Edition.Etat:=etGlGen ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.GLGenZoom ;
 QR.IsLegal:=FALSE ;
 QR.CritEdt:=Crit ;
 QR.InitType (nbGen,neGL,msGenEcr,'GLGENE','',FALSE,TRUE,FALSE,edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure GLGeneralLegal(Crit : TCritEdt) ;
var QR : TGdLivGen ;
    Edition : TEdition ;
BEGIN
QR:=TGdLivGen.Create(Application) ;
Edition.Etat:=etGlGen ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.GLGenZoom ;
 QR.IsLegal:=TRUE ;
 QR.CritEdt:=Crit ;
 QR.InitType (nbGen,neGL,msGenEcr,'GLGENE','',TRUE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TGdLivGen.GenereSQL ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
{ Construction de la clause Select de la SQL }
GenereSQLBase ;
GenereSQLSub ;
END ;

procedure TGdLivGen.GenereSQLSub ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
{ Construction de la clause Select de la SQL }
QEcr.Close ;
QEcr.SQL.Clear ;
QEcr.SQL.Add('Select E_GENERAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,J_LIBELLE,') ;
QEcr.SQL.Add(       'E_REFINTERNE,E_ETABLISSEMENT,E_LIBELLE,E_DEVISE,E_VALIDE,') ;
// GC - 03/02/2002
if CritEdt.GL.AvecMemoEcr then
  QEcr.SQL.Add(       'E_LETTRAGE,E_JOURNAL,J_CENTRALISABLE, E_NATUREPIECE, E_QUALIFPIECE, E_BLOCNOTE, ')
else
  QEcr.SQL.Add(       'E_LETTRAGE,E_JOURNAL,J_CENTRALISABLE, E_NATUREPIECE, E_QUALIFPIECE, ') ;  
// FIN GC - 03/02/2002

If CritEdt.SQLGA<>'' Then QECR.SQL.Add(CritEdt.SQLGA) ;
Case CritEdt.Monnaie of
  0 : BEGIN QEcr.SQL.Add('E_DEBIT DEBIT,E_CREDIT CREDIT') ; END ;
  1 : BEGIN QEcr.SQL.Add('E_DEBITDEV DEBIT,E_CREDITDEV CREDIT') ; END ;
end ;
QEcr.SQL.Add(' From ECRITURE') ;
QEcr.SQL.Add(' Left Join JOURNAL On J_JOURNAL=E_JOURNAL ') ;
{ Construction de la clause Where de la SQL }
QEcr.SQL.Add('Where E_GENERAL=:G_GENERAL ') ;
{ Tables explorées par la SQL }
QEcr.SQL.Add('And E_DATECOMPTABLE>="'+usdatetime(CritEdt.GL.Date21)+'" And E_DATECOMPTABLE<="'+usdatetime(CritEdt.Date2)+'"') ;
if FExercice.ItemIndex>0 then QEcr.SQL.Add(' And E_EXERCICE="'+CritEdt.Exo.Code+'" ') ;
QEcr.SQL.Add(   ' And E_NUMEROPIECE>='+IntToStr(CritEdt.GL.NumPiece1)) ;
QEcr.SQL.Add(   ' And E_NUMEROPIECE<='+IntToStr(CritEdt.GL.NumPiece2)) ;
QEcr.SQL.Add(TraduitNatureEcr(CritEdt.Qualifpiece, 'E_QUALIFPIECE', true, CritEdt.ModeRevision)) ;
if CritEdt.RefInterne<>'' then QEcr.SQL.Add('And UPPER(E_REFINTERNE) like "'+TraduitJoker(CritEdt.RefInterne)+'" ' );
if CritEdt.Etab<>'' then QEcr.SQL.Add(' And E_ETABLISSEMENT="'+CritEdt.Etab+'"') ;
if CritEdt.Valide<>'g' then QEcr.SQL.Add(' And E_VALIDE="'+CritEdt.Valide+'" ') ;
if CritEdt.DeviseSelect<>'' then QEcr.SQL.Add(' And E_DEVISE="'+CritEdt.DeviseSelect+'"') ;
QEcr.SQL.Add(' And E_ECRANOUVEAU="N" and E_QUALIFPIECE<>"C" ') ;
Case CritEdt.GL.Lettrable of
  2 : QEcr.SQL.Add(' And E_LETTRAGE=""  ') ;
  1 : QEcr.SQL.Add(' And E_LETTRAGE<>"" ') ;
 End ;
(* *)
Case CritEdt.GL.Pointable of
  2 : QEcr.SQL.Add(' And E_REFPOINTAGE=""  ') ;
  1 : QEcr.SQL.Add(' And E_REFPOINTAGE<>"" ') ;
 End ;
(**)
{ Construction de la clause Order By de la SQL }
If CritEDT.SQLPlus<>'' Then QEcr.SQL.Add(CritEDT.SQLPlus) ;
QEcr.SQL.Add(' Order By E_GENERAL,E_DATECOMPTABLE,E_JOURNAL,E_EXERCICE,E_NUMEROPIECE,E_NUMLIGNE') ;
ChangeSql(QEcr) ; //QEcr.Prepare ;
PrepareSQLODBC(QEcr) ;
QEcr.Open ;
END ;

procedure TGdLivGen.RenseigneCritere ;
{ Récupération des champs du multicritère dans l'entête d'état }
BEGIN
Inherited ;
If EstSerie(S5) Then RLegende.Caption:=MsgRien.Mess[7] ;
TLE_VALIDE.Caption:='' ;
RRefInterne.Caption:=FRefInterne.Text ;
RNumPiece1.Caption:=IntToStr(CritEdt.GL.NumPiece1)   ; RNumPiece2.Caption:=IntToStr(CritEdt.GL.NumPiece2) ;
CaseACocher(FValide,RValide) ; CaseACocher(FLettrable,RLettrable) ; CaseACocher(FPointable,RPointable) ;
(*if FSansRupt.Checked then BEGIN RSansRupt.Caption:='þ' ; RAvecRupt.Caption:='o' ; END ;
if FAvecRupt.Checked then BEGIN RSansRupt.Caption:='o' ; RAvecRupt.Caption:='þ' ; END ;
if FPlanRupt.Enabled then RPlanRupt.Caption:=FPlanRupt.Text ;
CaseACocher(FSoldeProg,RSoldeProg)       ; CaseACocher(FSautPage,RSautPage) ;
CaseACocher(FLigneCptPied,RLigneCptPied) ;
*)
DateCumulAuGL(TCumul,CritEdt,MsgBox.Mess[2]) ;
DateCumulAuGL(AnoRupt,CritEdt,MsgBox.Mess[2]) ;
(*
TTotBil.Caption:=MsgBox.Mess[3] ; TTotCha.Caption:=MsgBox.Mess[4] ;
TTotPro.Caption:=MsgBox.Mess[5] ; TTotGes.Caption:=MsgBox.Mess[6] ;
TTotExt.Caption:=MsgBox.Mess[7] ;
TTotBilGen.Caption:=MsgBox.Mess[8] ; TTotChaGen.Caption:=MsgBox.Mess[9] ;
TTotProGen.Caption:=MsgBox.Mess[10] ; TTotGesGen.Caption:=MsgBox.Mess[11] ;
TTotExtGen.Caption:=MsgBox.Mess[12] ;
*)
END ;

procedure TGdLivGen.ChoixEdition ;
{ Initialisation des options d'édition }
BEGIN
Inherited ;
QRDLRupt.PrintBefore:=TRUE ;
Case CritEdt.Rupture of
  rLibres  :  BEGIN
              QRDLRupt.PrintBefore:=FALSE ;
              ChargeGroup(LRupt,['G00','G01','G02','G03','G04','G05','G06','G07','G08','G09']) ;
              END ;
  rRuptures : BEGIN
              ChargeRupt(LRupt, 'RUG', CritEdt.GL.PlanRupt, '', '') ;
              NiveauRupt(LRupt) ;
              END ;
  rCorresp  : BEGIN
              ChargeRuptCorresp(LRupt, CritEdt.GL.PlanRupt, '', '', False) ;
              NiveauRupt(LRupt) ;
              END ;
  End ;

ChgMaskChamp(QR2DEBIT , CritEdt.Decimale, CritEdt.AfficheSymbole, CritEdt.Symbole, False) ;
ChgMaskChamp(QR2CREDIT, CritEdt.Decimale, CritEdt.AfficheSymbole, CritEdt.Symbole, False) ;
ChargeGroup(L,['TOTJAL']) ;
ChargeGroup(LMulti,['Total mensuel : ']) ;
ChargeRecap(LCentr) ;
//ChargeBilan(LBil) ;
if CritEdt.GL.BilGes then
   BEGIN ChargeRecap(LBilan) ; ChargeRecap(LBil) ; ChargeRecap(LCha) ; ChargeRecap(LPro) ; ChargeRecap(LExt) ; END ;

// Rony 9/07/97
//If CritEdt.GA then BSDMVT.Height:=Sup1.Top+Sup1.Height else BSDMVT.Height:=Sup1.Top ;
END ;

procedure TGdLivGen.FinirPrint;
Var Solde : Double ;
BEGIN
Inherited ;
QEcr.Close ;
VideGroup(L) ; VideGroup(LMulti) ;  VideRecap(LCentr) ;
If GCalc<>NIL Then BEGIN GCalc.Free ; GCalc:=NIL ; END ;
if CritEdt.GL.BilGes then
   BEGIN VideRecap(LBilan) ; VideRecap(LBil) ; VideRecap(LCha) ; VideRecap(LPro) ; VideRecap(LExt) ; END ;

//VideBilan(LBil) ;
if CritEdt.Rupture<>rRien then VideRupt(LRupt) ;
//QBalC.Free ;
If QBalC<>NIL Then BEGIN QBalC.Free ; QBalC:=NIL ; END ;
if OkMajEdt Then
   BEGIN
   Solde:=TotEdt[1].TotDebit-TotEdt[1].TotCredit ;
   if Solde<0 then
      BEGIN
      If IsLegal then MajEditionLegal('GLG', CritEdt.Exo.Code, DateToStr(CritEdt.DateDeb), DateToStr(CritEdt.DateFin),'', TotEdt[1].TotDebit, TotEdt[1].TotCredit, Solde, 0)
                 Else MajEdition('GLG', CritEdt.Exo.Code, DateToStr(CritEdt.DateDeb), DateToStr(CritEdt.DateFin),'', TotEdt[1].TotDebit, TotEdt[1].TotCredit, Solde, 0) ;
      END Else
      BEGIN
      If IsLegal then MajEditionLegal('GLG', CritEdt.Exo.Code, DateToStr(CritEdt.DateDeb), DateToStr(CritEdt.DateFin),'', TotEdt[1].TotDebit, TotEdt[1].TotCredit, 0, Solde)
                 Else MajEdition('GLG', CritEdt.Exo.Code, DateToStr(CritEdt.DateDeb), DateToStr(CritEdt.DateFin),'', TotEdt[1].TotDebit, TotEdt[1].TotCredit, 0, Solde) ;
      END ;
  END ;
END ;

{***************** EVENEMENTS LIES A L'AFFICHAGE DES BANDES ********************}
{**************************** GESTION DU FILTRE *******************************}
procedure TGdLivGen.InitDivers ;
BEGIN
Inherited ;
OkGes:=False ; 
{ Les tableaux de montants }
Fillchar(TotGes,SizeOf(TotGes),#0) ;
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
Fillchar(TotAncRupt,SizeOf(TotAncRupt),#0) ;
{ Labels sur les bandes }
BFCpteGen.Frame.DrawTop:=CritEdt.GL.FormatPrint.PrSepCompte[2] ;
//BDetail.Frame.DrawTop:=CritEdt.GL.FormatPrint.PrSepCompte[2] ;
BRupt.Frame.DrawTop:=CritEdt.GL.FormatPrint.PrSepCompte[3] ;
BRupt.Frame.DrawBottom:=CritEdt.GL.FormatPrint.PrSepCompte[3] ;
If CritEdt.GA Then CritEdt.SQLGA:=GetParamListe(NomListeGA,Self) ;
OkEnteteRupt:=(CritEdt.Rupture=rCorresp) ;
FirstPassage:=True ;
END ;

procedure TGdLivGen.RecupCritEdt ;
Var NonLibres : Boolean ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  GL.CodeRupt1:='' ; GL.CodeRupt2:='' ;
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

//--
  if FLettrable.State=cbUnChecked then GL.Lettrable:=2 ;
  if FLettrable.State=cbChecked then GL.Lettrable:=1 ;
  if FLettrable.State=cbGrayed then GL.Lettrable:=0 ;
//--
  if FPointable.State=cbUnChecked then GL.Pointable:=2 ;
  if FPointable.State=cbChecked then GL.Pointable:=1 ;
  if FPointable.State=cbGrayed then GL.Pointable:=0 ;
//--
  GL.QuelAN:=FQuelAN(FAvecAN.Checked) ;
  GL.SansDetailCentralisation:=Not FCentralisation.Checked ;
  GL.BilGes:=FBilGes.Checked ;
//--
  GL.AvecDetailCentralise := FDetailCentralise.Checked;
  GL.AvecMemoEcr := FAvecMemo.Checked;
//--

  NonLibres:=((Rupture=rRuptures) or (Rupture=rCorresp)) ;
  if NonLibres then GL.PlanRupt:=FPlanRuptures.Value ;
  GL.RuptOnly:=QuelleTypeRupt(1,FSansRupt.Checked,FAvecRupt.Checked,FALSE) ;
  If (CritEdt.Rupture=rCorresp) Then GL.PlansCorresp:=FPlanRuptures.ItemIndex+1 ;
  GL.OnlyCptAssocie:=(Rupture<>rRien) and FOnlyCptAssocie.Checked ;
  With GL.FormatPrint Do
     BEGIN
     PrColl:=FCollectif.Checked ;
     PrSepCompte[2]:=FLigneCptPied.Checked ;
     PrSepCompte[3]:=FLigneRupt.Checked ;
     END ;
  END ;
END ;

function  TGdLivGen.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ;
If Result Then
   BEGIN
   (*
   QBal:=PrepareTotCpt(fbGene,QuelTypeEcr,Dev,Etab,Exo) ;
   *)
   QBalC:=PrepareTotCptSolde(QuelTypeEcr,Dev,Etab,Exo,DevEnP) ;
   Gcalc:=TGCalculCum.create(Un,fbGene,fbGene,QuelTypeEcr,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE) ;
   GCalc.initCalcul('','','','',CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                    CritEdt.Date1,CritEdt.GL.Date11,TRUE) ;
   END ;
END ;

procedure TGdLivGen.FormShow(Sender: TObject);
begin
HelpContext:=7415000 ;
//Standards.HelpContext:=7415010 ;
//Avances.HelpContext:=7415020 ;
//Mise.HelpContext:=7415030 ;
//Option.HelpContext:=7415040 ;
//TabRuptures.HelpContext:=7415050 ;
  inherited;
Floading:=FALSE ;
TabSup.TabVisible:=False;
FCollectif.Checked:=FALSE ; If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
FRefInterne.Text:='' ; FLigneRupt.Enabled:=False ;
FValide.Checked:=True ; FValide.State:=cbGrayed ;
FJoker.MaxLength:=VH^.Cpta[fbGene].lg ;
FSurRupt.Visible:=False ;
FCodeRupt1.Visible:=False  ; FCodeRupt2.Visible:=False ;
TFCodeRupt1.Visible:=False ; TFCodeRupt2.Visible:=False ;
FOnlyCptAssocie.Enabled:=False ;

  // GC - GP - 20/12/2001
  FDetailCentralise.Visible := TRUE Or (CtxPcl in V_PGI.PgiContexte);
  FAvecMemo.Visible := TRUE Or (CtxPcl in V_PGI.PgiContexte);
  if TRUE Or (CtxPcl in V_Pgi.PgiContexte) then
    FCentralisation.Caption := 'Edition avec les centralisations'
  else
    FCentralisation.Caption := 'Détail de la centralisation';
  FCentralisation.OnClick ( nil );
  if CritEdtChaine.Utiliser then
    InitEtatChaine('QRGLGENE');
  // GC - FIN

end;

procedure TGdLivGen.BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
var CumulAu : TabTot ;
    D, C  : Double ;
    TotCpt1M,TotCpt1S : TabTot ;
    TotCpt1MEURO,TotCpt1SEURO : TabTot ;
    OkPourColl : Boolean ; CptRupt : String ;
begin
  inherited;
RepotCent:=False ;
G_GENERAL.Caption:=MsgBox.Mess[20]+' '+Qr1G_GENERAL.AsString+' '+Qr1G_LIBELLE.AsString ;
Fillchar(TotCptGen,SizeOf(TotCptGen),#0) ;
Fillchar(TotJalCent,SizeOf(TotJalCent),#0) ;
Fillchar(TotCptGenMvt,SizeOf(TotCptGenMvt),#0) ;
Fillchar(TotCpt1M,SizeOf(TotCpt1M),#0) ;
Fillchar(TotCpt1S,SizeOf(TotCpt1S),#0) ;
Fillchar(TotCpt1MEURO,SizeOf(TotCpt1M),#0) ;
Fillchar(TotCpt1SEURO,SizeOf(TotCpt1S),#0) ;
Fillchar(CumulAu,SizeOf(CumulAu),#0) ;
VideRecap(LCentr) ; ChargeRecap(LCentr) ; TitreCentr.Visible:=TRUE ;
Fillchar(CritEDT.GL.FormatPrint.Report,SizeOf(CritEDT.GL.FormatPrint.Report),#0) ;
InitReport([2],CritEDT.GL.FormatPrint.Report) ;
Case CritEdt.Rupture of
  rLibres    : if CritEdt.GL.OnlyCptAssocie then PrintBand:=DansRuptLibre(Q,fbGene,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
  rRuptures  : if CritEdt.GL.OnlyCptAssocie then PrintBand:=DansRupt(LRupt,Qr1G_GENERAL.AsString) ;
  rCorresp   : if CritEdt.GL.OnlyCptAssocie then
                 if CritEDT.GL.PlansCorresp=1 then PrintBand:=(Qr1G_CORRESP1.AsString<>'') Else
                 if CritEDT.GL.PlansCorresp=2 then PrintBand:=(Qr1G_CORRESP2.AsString<>'') ;
  End;
Affiche:=PrintBand ;
if PrintBand then
   BEGIN
   OkPourColl:=(QR1G_COLLECTIF.AsString='X') And (CritEDT.GL.FormatPrint.PrColl) ;
   StReportGen:=Qr1G_GENERAL.AsString ;
   Case CritEdt.SautPage of
     0 : BDetail.forceNewPage:=(Qr1G_SAUTPAGE.AsString='X') ;
     1 : BDetail.forceNewPage:=True ;
     2 : BDetail.forceNewPage:=False ;
    end ;
   (*
   ExecuteTotCpt(QBal,QG_GENERAL.AsString, CritEdt.Date1,CritEdt.GL.Date11,
                 CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,Total,TRUE) ;
   *)
   if OkPourColl And (CritEdt.GL.QuelAN<>SansAN) Then
      BEGIN
      ExecuteTotCptSolde(QBalC,QR1G_GENERAL.AsString, CritEDT.Date1,CritEDT.Date1,
                         CritEDT.DeviseSelect,CritEDT.Etab,CritEDT.Exo.Code,TotCpt1M,TotCpt1S,TotCpt1MEURO,TotCpt1SEURO,FALSE,TRUE,
                         CritEdt.Decimale,V_PGI.OkDecE,CritEdt.Monnaie=2,DevEnP,QuelTypeEcr) ;
      END ;
   If (CritEdt.GL.QuelAN<>SansAN) Then
      BEGIN
      GCAlc.ReInitCalcul(Qr1G_GENERAL.AsString,'',0,0) ;
      GCalc.Calcul ; CumulAu:=GCalc.ExecCalc.TotCpt ;
      END ;
   If OkPourColl Then CumulAu[0]:=TotCpt1S[0] Else CumulVersSolde(CumulAu[0]) ;
   if CritEDT.Date1=CritEDT.GL.Date11 Then Fillchar(CumulAu[1],SizeOf(CumulAu[1]),#0) ;
   If CritEdt.GL.QuelAN=SansAN Then BEGIN CumulAu[0].TotDebit:=0 ; CumulAu[0].TotCredit:=0 ; END ;
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
      TotCptGen[1].TotDebit:= Arrondi(TotCptGen[1].TotDebit+D,CritEDT.Decimale) ;
      TotCptGen[1].TotCredit:=Arrondi(TotCptGen[1].TotCredit+C,CritEDT.Decimale) ;
      TotAnCRupt[1].TotDebit:= Arrondi(TotAnCRupt[1].TotDebit+D,CritEDT.Decimale) ;
      TotAnCRupt[1].TotCredit:=Arrondi(TotAnCRupt[1].TotCredit+C,CritEDT.Decimale) ;
      AddReport([1,2],CritEdt.GL.FormatPrint.Report,D,C,CritEdt.Decimale) ;
      END Else
      BEGIN
      AnvDebit.Caption:=MsgBox.Mess[17] ; AnvCredit.Caption:='' ; AnvSolde.Caption:='' ;
      CumDebit.Caption:=MsgBox.Mess[17] ; CumCredit.Caption:='' ;  CumSolde.Caption:='' ;
      END ;

   (*
   { A Nouveau }
   TotCptGen[1].TotDebit:=CumulAu[0].TotDebit ;
   TotCptGen[1].TotCredit:=CumulAu[0].TotCredit ;
   AnvDebit.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotCptGen[1].TotDebit, CritEdt.AfficheSymbole) ;
   AnvCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotCptGen[1].TotCredit, CritEdt.AfficheSymbole) ;
   AnvSolde.Caption:= PrintSolde(TotCptGen[1].TotDebit,TotCptGen[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;

   { A Nouveau + Cumul }
   TotCptGen[1].TotDebit:= Arrondi(TotCptGen[1].TotDebit+CumulAu[1].TotDebit,CritEdt.Decimale) ;
   TotCptGen[1].TotCredit:=Arrondi(TotCptGen[1].TotCredit+CumulAu[1].TotCredit,CritEdt.Decimale) ;
   CumDebit.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotCptGen[1].TotDebit, CritEdt.AfficheSymbole) ;
   CumCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotCptGen[1].TotCredit,CritEdt.AfficheSymbole) ;
   CumSolde.Caption:= PrintSolde(TotCptGen[1].TotDebit,TotCptGen[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;

   TotAnCRupt[1].TotDebit:= Arrondi(TotAnCRupt[1].TotDebit+TotCptGen[1].TotDebit,CritEdt.Decimale) ;
   TotAnCRupt[1].TotCredit:=Arrondi(TotAnCRupt[1].TotCredit+TotCptGen[1].TotCredit,CritEdt.Decimale) ;

   { Ajout (ANouveau + Cuml) dans le solde progr. }
   D:=CumulAu[0].TotDebit+CumulAu[1].TotDebit ;
   C:=CumulAu[0].TotCredit+CumulAu[1].TotCredit ;
   *)

   { Init du Cumul Progressif }
   Case CritEdt.SoldeProg of
     1 : BEGIN
                 Progressif(False,0,0) ; { Si Progressif  obligatoire alors initialise Solde Progressif }
                 Progressif(True,D,C) ;
                 END ;
     0 : if Qr1G_SOLDEPROGRESSIF.AsString='X' then
                    BEGIN
                    Progressif(False,0,0) ; { Si Progressif ou non alors seulement si demandé par le compte }
                    Progressif(True,D,C) ;
                    END ;
    end ;
   //AddReport([1,2],CritEdt.GL.FormatPrint.Report,D,C,CritEdt.Decimale) ;

   Case CritEdt.Rupture of
     rLibres   : AddGroupLibre(LRupt,Q,fbGene,CritEdt.LibreTrie,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
     rRuptures : AddRupt(Lrupt,Qr1G_GENERAL.AsString,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
     rCorresp  : BEGIN
                 Case CritEDT.GL.PlansCorresp Of
                   1 : If Qr1G_CORRESP1.AsString<>'' Then CptRupt:=Qr1G_CORRESP1.AsString+Qr2E_GENERAL.AsString
                                                     Else CptRupt:='.'+Qr2E_GENERAL.AsString ;
                   2 : If Qr1G_CORRESP2.AsString<>'' Then CptRupt:=Qr1G_CORRESP2.AsString+Qr2E_GENERAL.AsString
                                                     Else CptRupt:='.'+Qr2E_GENERAL.AsString ;
                  End ;
                 AddRuptCorres(LRupt,CptRupt,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
                 END ;
     End ;
   Quoi:=QuoiGen(1) ;
   END ;
end;

procedure TGdLivGen.BSDMvtBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var i : integer ;
    StJal,SDate1,SDate2 : String ;
    LigneCentralisable  : Boolean ;
    CptRupt             : String ;
begin
  inherited;
PrintBand:=Affiche And (Not QR2E_GENERAL.IsNull) ;
if Not PrintBand then Exit ;

// GC 02/02/2002
if CritEdt.Gl.AvecDetailCentralise then
  LigneCentralisable := False
else
  LigneCentralisable := (QR2J_CENTRALISABLE.AsString='X') and (QR1G_CENTRALISABLE.AsString='X');

QRRichEdit1.Visible := False;
QRRichEdit1.AutoStretch := False;
BSDMvt.Height := 13 ;

if CritEdt.GL.AvecMemoEcr then
begin
  FHRichEdit1.Clear;
  FHRichEdit1.Lines.Assign(QR2E_BLOCNOTE);

  QRRichEdit1.Visible     := Trim(FHRichEdit1.Text) <> '';
  QRRichEdit1.AutoStretch := Trim(FHRichEdit1.Text) <> '';
end;
// FIN - GC

If IsECC then
   BEGIN
   LE_DEBIT.Caption:='0' ;
   LE_CREDIT.Caption:='0' ;
   END ;
   
If CritEdt.GA Then For i:=1 To 7 Do AlimListe(QEcr,i,Self) ;
E_PIECELIGECH.Caption:=QR2E_NUMEROPIECE.AsString+' / '+QR2E_NUMLIGNE.AsString ;
LE_VALIDE.Caption:=ValiQuali(QR2E_VALIDE.AsString,QR2E_QUALIFPIECE.AsString) ;
Case CritEdt.SoldeProg of             { Affectation Du Calcul du Solde Progressif O/N Sinon, d'aprés l'info sur le compte }
  0 : if Qr1G_SOLDEPROGRESSIF.AsString='X' then LE_SOLDE.Caption:=PrintSolde(ProgressDebit+QR2DEBIT.AsFloat,ProgressCredit+QR2CREDIT.AsFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole)
                                           Else LE_SOLDE.Caption:=PrintSolde(QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
  1 : LE_SOLDE.Caption:=PrintSolde(ProgressDebit+QR2DEBIT.AsFloat,ProgressCredit+QR2CREDIT.AsFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
  2 : LE_SOLDE.Caption:=PrintSolde(QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
 end ;
TotCptGen[1].TotDebit :=Arrondi(TotCptGen[1].TotDebit + QR2DEBIT.AsFloat,CritEdt.Decimale) ;
TotCptGen[1].TotCredit:=Arrondi(TotCptGen[1].TotCredit+QR2CREDIT.AsFloat,CritEdt.Decimale) ;
TotCptGenMvt[1].TotDebit :=Arrondi(TotCptGenMvt[1].TotDebit + QR2DEBIT.AsFloat,CritEdt.Decimale) ;
TotCptGenMvt[1].TotCredit:=Arrondi(TotCptGenMvt[1].TotCredit+QR2CREDIT.AsFloat,CritEdt.Decimale) ;
If Not CritEdt.GL.ForceNonCentralisable Then PrintBand:= Not LigneCentralisable;

AddGroup(L,[QR2E_JOURNAL],[QR2DEBIT.AsFloat, QR2CREDIT.AsFloat]) ;

Case CritEdt.Rupture of
  rLibres   : AddGroupLibre(LRupt,Q,fbGene,CritEdt.LibreTrie,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0]) ;
  rRuptures : AddRupt(LRupt,Qr2E_GENERAL.AsString,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0]) ;
  rCorresp  : BEGIN
              Case CritEDT.GL.PlansCorresp Of
                1 : If Qr1G_CORRESP1.AsString<>'' Then CptRupt:=Qr1G_CORRESP1.AsString+Qr2E_GENERAL.AsString
                                                  Else CptRupt:='.'+Qr2E_GENERAL.AsString ;
                2 : If Qr1G_CORRESP2.AsString<>'' Then CptRupt:=Qr1G_CORRESP2.AsString+Qr2E_GENERAL.AsString
                                                  Else CptRupt:='.'+Qr2E_GENERAL.AsString ;
                Else CptRupt:=Qr2E_GENERAL.AsString ;
               End ;
              AddRuptCorres(LRupt,CptRupt,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0]) ;
              END ;
  End ;
if Not PrintBand then
   BEGIN
   TotJalCent[1].TotDebit :=Arrondi(TotJalCent[1].TotDebit + QR2DEBIT.AsFloat, CritEdt.Decimale) ;
   TotJalCent[1].TotCredit:=Arrondi(TotJalCent[1].TotCredit+ QR2CREDIT.AsFloat,CritEdt.Decimale) ;
   StJal:=Format_String(Qr2E_JOURNAL.AsString,3) ;
   SDate1:=FormatDatetime('mmyyyy',Qr2E_DATECOMPTABLE.AsDateTime) ;
   SDate2:=FormatDatetime('mmmm yyyy',Qr2E_DATECOMPTABLE.AsDateTime) ;
   // GC - GP 09/01/2002
   if TRUE Or (CtxPcl in V_PGI.PgiContexte) then
   begin
     LCentr.Sorted := False;
     AddRecap(LCentr,[Sdate1,StJal],[Qr2J_LIBELLE.AsString,SDAte2],[QR2DEBIT.AsFloat,QR2CREDIT.AsFloat])
   end
   else
     AddRecap(LCentr,[StJal,SDate1],[Qr2J_LIBELLE.AsString,SDAte2],[QR2DEBIT.AsFloat,QR2CREDIT.AsFloat]) ;
   // FIN - GC
   END Else
   BEGIN
   AddGroup(LMulti, [QR2E_DATECOMPTABLE], [QR2DEBIT.AsFloat, QR2CREDIT.AsFloat]) ;
   AddReport([1,2],CritEdt.GL.FormatPrint.Report,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,CritEdt.Decimale) ;
   END ;
//AddReport([1,2],CritEdt.GL.FormatPrint.Report,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,CritEdt.Decimale) ;

// GC - 21/01/2002
(*if (CtxPcl in V_Pgi.PgiContexte) and PrintBand then
begin
  if (QR2E_GENERAL.AsString = VH^.EccEuroDebit) or (QR2E_GENERAL.AsString = VH^.EccEuroCredit) then
  begin
    Printband := not ((QR2DEBIT.AsFloat = 0) and (QR2CREDIT.AsFloat = 0));
  end;
end;*)
// FIN GC

Quoi:=QuoiMvt ;
end;

procedure TGdLivGen.BSDMvtAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Case CritEdt.SoldeProg of
  1 :  Progressif(True,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat) ;
  0 :  if Qr1G_SOLDEPROGRESSIF.AsString='X' then Progressif(True,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat) ;
  end ;
end;

procedure TGdLivGen.BFCpteGenBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var Top1, Top2, Top3 : Byte ;
    PourBilGes       : Array[0..12] of double ;
begin
  inherited;
PrintBand:=Affiche ;
if PrintBand then
   BEGIN
   G_GENERAL_.Caption:=MsgBox.Mess[21]+' '+Qr1G_GENERAL.AsString+' '+Qr1G_LIBELLE.AsString ;
   TotJalNonCentral.Visible:=(Qr1G_CENTRALISABLE.asString='X') And (Not CritEdt.GL.ForceNonCentralisable) ;
   TotJCentDebit.Visible:=(Qr1G_CENTRALISABLE.asString='X') And (Not CritEdt.GL.ForceNonCentralisable) ;
   TotJCentCredit.Visible:=(Qr1G_CENTRALISABLE.asString='X') And (Not CritEdt.GL.ForceNonCentralisable) ;
   TotJCentSolde.Visible:=(Qr1G_CENTRALISABLE.asString='X') And (Not CritEdt.GL.ForceNonCentralisable) ;
   TotJCentDebit.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotJalCent[1].TotDebit, CritEdt.AfficheSymbole) ;
   TotJCentCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotJalCent[1].TotCredit, CritEdt.AfficheSymbole) ;
   TotJCentSolde.Caption:=PrintSolde(TotJalCent[1].TotDebit, TotJalCent[1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
   TotGenDebit.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotCptGen[1].TotDebit, CritEdt.AfficheSymbole) ;
   TotGenCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotCptGen[1].TotCredit, CritEdt.AfficheSymbole) ;
   TotGenSolde.Caption:=PrintSolde(TotCptGen[1].TotDebit,TotCptGen[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
   TotGenMvtDebit.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotCptGenMvt[1].TotDebit, CritEdt.AfficheSymbole) ;
   TotGenMvtCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotCptGenMvt[1].TotCredit, CritEdt.AfficheSymbole) ;
   TotGenMvtSolde.Caption:=PrintSolde(TotCptGenMvt[1].TotDebit,TotCptGenMvt[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
   { Alimentation des valeurs pour le Total général }
   TotEdt[1].TotDebit :=Arrondi(TotEdt[1].TotDebit+TotCptGen[1].TotDebit,CritEdt.Decimale) ;
   TotEdt[1].TotCredit:=Arrondi(TotEdt[1].TotCredit+TotCptGen[1].TotCredit,CritEdt.Decimale) ;

   if CritEdt.GL.BilGes then
      BEGIN
      PourBilGes[0]:=TotCptGenMvt[1].TotDebit ;
      PourBilGes[1]:=TotCptGenMvt[1].TotCredit ;
      PourBilGes[2]:=TotCptGen[1].TotDebit ;
      PourBilGes[3]:=TotCptGen[1].TotCredit ;
      RecapPourQui(Qr1G_GENERAL.AsString,LBilan,PourBilGes) ;
//      RecapPourQui(Qr1G_GENERAL.AsString,LBil,LCha,LPro,LExt,PourBilGes) ;
      END ;
//   AddBilan(LBil, Qr1G_GENERAL.AsString, [0,0,TotCptGenMvt[1].TotDebit, TotCptGenMvt[1].TotCredit, TotCptGen[1].TotDebit, TotCptGen[1].TotCredit]) ;
   if TotJalNonCentral.Visible then
      BEGIN
      Top1:=0 ;
      TotJalNonCentral.Top:=Top1 ; TotJCentDebit.Top:=Top1 ;
      TotJCentCredit.Top:=Top1 ; TotJCentSolde.Top:=Top1 ;
      Top2:=Top1+TotJalNonCentral.Height+1 ;
      TitreTotGenMvt.Top:=Top2 ; TotGenMvtDebit.Top:=Top2 ;
      TotGenMvtCredit.Top:=Top2 ; TotGenMvtSolde.Top:=Top2 ;
      Top3:=Top2+TitreTotGenMvt.Height+1 ;
      G_GENERAL_.Top:=Top3 ; TotGenDebit.Top:=Top3 ;
      TotGenCredit.Top:=Top3 ; TotGenSolde.Top:=Top3 ;
      BFCpteGen.Height:=Top3+G_GENERAL_.Height+2 ;
      END Else
      BEGIN
      Top1:=0 ;
      TitreTotGenMvt.Top:=Top1 ; TotGenMvtDebit.Top:=Top1 ;
      TotGenMvtCredit.Top:=Top1 ; TotGenMvtSolde.Top:=Top1 ;
      Top2:=Top1+TitreTotGenMvt.Height+1 ;
      G_GENERAL_.Top:=Top2 ; TotGenDebit.Top:=Top2 ;
      TotGenCredit.Top:=Top2 ; TotGenSolde.Top:=Top2 ;
      BFCpteGen.Height:=Top2+G_GENERAL_.Height+2 ;
      END ;
   END ;
Quoi:=QuoiGen(2) ;
end;

procedure TGdLivGen.BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
TotalDebit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotEdt[1].TotDebit, CritEdt.AfficheSymbole) ;
TotalCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotEdt[1].TotCredit, CritEdt.AfficheSymbole) ;
TotalSoldes.Caption:=PrintSolde(TotEdt[1].TotDebit,TotEdt[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
//ShowMessage('Après édit GL : '+DateToStr(CritEdt.Date1)+'   '+DateToStr(CritEdt.Date2)) ;
end;

procedure TGdLivGen.QRDLRuptNeedData(var MoreData: Boolean);
var SumD, SumC       : Double ;
    CodRupt, LibRupt, Lib1, CptRupt, StCode : String ;
    TotRupt          : Array[0..12] Of Double ;
    OkOk             : Boolean ;
    Quellerupt       : Integer ;
    Col : TColor ;
    LibRuptInf : Array[1..10] Of TRuptInf ;
begin
  inherited;
MoreData:=false ;
Case CritEdt.Rupture of
  rLibres   : BEGIN
              MoreData:=PrintGroupLibre(LRupt,Q,fbGene,CritEdt.LibreTrie,CodRupt,LibRupt,Lib1,TotRupt,Quellerupt,Col,LibRuptInf) ;
              If MoreData then
                 BEGIN
                 StCode:=CodRupt ;
                 Delete(StCode,1,Quellerupt+2) ;
                 MoreData:=DansChoixCodeLibre(StCode,Q,fbGene,CritEdt.LibreCodes1,CritEdt.LibreCodes2, CritEdt.LibreTrie) ;
                 BRupt.Font.Color:=Col ;
                 END ;
              END ;
  rRuptures : MoreData:=PrintRupt(LRupt,Qr1G_GENERAL.AsString,CodRupt,LibRupt,DansTotGen,QRP.EnRupture,TotRupt) ;
  rCorresp  : BEGIN
              OkOk:=TRUE ;
              Case CritEDT.GL.PlansCorresp  Of
                1 : If Qr1G_CORRESP1.AsString<>'' Then CptRupt:=Qr1G_CORRESP1.AsString+Qr2E_GENERAL.AsString
                                                  Else CptRupt:='.'+Qr2E_GENERAL.AsString ;
                2 : If Qr1G_CORRESP2.AsString<>'' Then CptRupt:=Qr1G_CORRESP2.AsString+Qr2E_GENERAL.AsString
                                                  Else CptRupt:='.'+Qr2E_GENERAL.AsString ;
                Else OkOk:=FALSE ;
                END ;
              If OkOk Then MoreData:=PrintRupt(LRupt,CptRupt,CodRupt,LibRupt,DansTotGen,QRP.EnRupture,TotRupt) Else MoreData:=FALSE ;
              END ;
  End ;

if MoreData then
   BEGIN
   TCodRupt.Width:=TLE_DATECOMPTABLE.Width+TE_JOURNAL.Width+TE_PIECE.Width+TE_REFINTERNE.Width+1 ;
   TCodRupt.Caption:='' ;
   BRupt.Height:=HauteurBandeRuptIni ;
   Case CritEdt.Rupture of
     rLibres   : BEGIN
                 insert(MsgBox.Mess[26]+' ',CodRupt,Quellerupt+2) ;
                 TCodRupt.Caption:=CodRupt+' '+Lib1 ;
                 AlimRuptSup(HauteurBandeRuptIni,QuelleRupt,TCodRupt.Width,BRupt,LibRuptInf,Self) ;
                 END ;
     rRuptures, rCorresp : TCodRupt.Caption:=CodRupt+'  '+LibRupt ;
     End ;
   SumD:=Arrondi(TotAnCRupt[1].TotDebit+ TotRupt[1],CritEdt.Decimale) ;
   SumC:=Arrondi(TotAnCRupt[1].TotCredit+TotRupt[2],CritEdt.Decimale) ;

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
      OkEnteteRupt:=True ;
      //Copy(CodRupt,)
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
      TCodRupt.Top:=2 ; DebitRupt.Top:=2 ; CreditRupt.Top:=2 ; SoldeRupt.Top:=2 ;
      BRupt.Height:=20 ;
      END;
   DebitRupt.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, SumD, CritEdt.AfficheSymbole) ;
   CreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, SumC, CritEdt.AfficheSymbole) ;
   SoldeRupt.Caption:=PrintSolde(SumD, SumC, CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;
  (**)
   END ;
end;

procedure TGdLivGen.BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.GL.RuptOnly<>Sans) ;
end;

procedure TGdLivGen.BRuptAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Fillchar(TotAncRupt,SizeOf(TotAncRupt),#0) ;

end;

procedure TGdLivGen.BSDMultiBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
Case CritEdt.TotMens of             { Total Mensuel O/N Sinon, d'aprés l'info sur le compte }
 0 : PrintBand:=(Qr1G_TOTAUXMENSUELS.AsString='X') ;
 1 : PrintBand:=True  ;
 2 : PrintBand:=False  ;
   End;

end;

procedure TGdLivGen.QRDLMultiNeedData(var MoreData: Boolean);
Var Cod, Lib : String ;
    Tot      : Array[0..12] of Double ;
    QuelleRupt: Integer ;
begin
  inherited;
MoreData:=PrintGroup(LMulti, QEcr, [QR2E_DATECOMPTABLE], Cod, Lib, Tot,Quellerupt) ;
LibTotalMulti.Caption:=MsgBox.Mess[16]+' '+Cod ;
TotMultiDebit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,Tot[0], CritEdt.AfficheSymbole) ;
TotMultiCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,Tot[1],CritEdt.AfficheSymbole) ;
TotSoldeMulti.Caption:=PrintSolde(Tot[0], Tot[1], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
(* Rony 10/04/97 -- Pas de Progerssif  pour le Mens
Case CritEdt.SoldeProg of             { Affectation Du Calcul du Solde Progressif O/N Sinon, d'aprés l'info sur le compte }
  0 : if Qr1G_SOLDEPROGRESSIF.AsString='X' then TotSoldeMulti.Caption:=PrintSolde(ProgressDebit,ProgressCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole)
                                           Else TotSoldeMulti.Caption:=PrintSolde(Tot[0], Tot[1], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
  1 : TotSoldeMulti.Caption:=PrintSolde(ProgressDebit,ProgressCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
  2 : TotSoldeMulti.Caption:=PrintSolde(Tot[0], Tot[1], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
 End ;
*)
end;

procedure TGdLivGen.TOPREPORTBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
Titre1Rep.Caption:=Titre2Rep.Caption ;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
Report1Solde.Caption:=Report2Solde.Caption ;
end;

procedure TGdLivGen.BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var MDebit,MCredit : Double ;
begin
  inherited;
MDebit:=0 ; MCredit:=0 ;
Case QuelReport(CritEdt.GL.FormatPrint.Report,MDebit,MCredit) Of
  1 : BEGIN TITREREPORTB.Caption:=MsgBox.Mess[14] ; Titre2Rep.Caption:='' ; END ;
  2 : BEGIN TITREREPORTB.Caption:=MsgBox.Mess[15] ; Titre2Rep.Caption:=StReportGen ; END ;
  3 : BEGIN TITREREPORTB.Caption:=MsgBox.Mess[22] ; Titre2Rep.Caption:='' ; END ;
  End ;
Report2Debit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MDebit, CritEdt.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MCredit, CritEdt.AfficheSymbole ) ;
Report2Solde.Caption:=PrintSolde(MDebit,MCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
end;

procedure TGdLivGen.QEcrAfterOpen(DataSet: TDataSet);
begin
  inherited;
QR2E_GENERAL       :=TStringField(QEcr.FindField('E_GENERAL')) ;
QR2E_EXERCICE      :=TStringField(QEcr.FindField('E_EXERCICE')) ;
QR2E_DATECOMPTABLE :=TDateTimeField(QEcr.FindField('E_DATECOMPTABLE')) ;
QR2E_DATECOMPTABLE.Tag:=1 ;
QR2E_NUMEROPIECE   :=TIntegerField(QEcr.FindField('E_NUMEROPIECE')) ;
QR2E_NUMLIGNE      :=TIntegerField(QEcr.FindField('E_NUMLIGNE')) ;
QR2J_LIBELLE       :=TStringField(QEcr.FindField('J_LIBELLE')) ;
QR2E_REFINTERNE    :=TStringField(QEcr.FindField('E_REFINTERNE')) ;
QR2E_ETABLISSEMENT :=TStringField(QEcr.FindField('E_ETABLISSEMENT')) ;
QR2E_LIBELLE       :=TStringField(QEcr.FindField('E_LIBELLE')) ;
QR2E_DEVISE        :=TStringField(QEcr.FindField('E_DEVISE')) ;
QR2E_VALIDE        :=TStringField(QEcr.FindField('E_VALIDE')) ;
QR2E_QUALIFPIECE   :=TStringField(QEcr.FindField('E_QUALIFPIECE')) ;
QR2E_LETTRAGE      :=TStringField(QEcr.FindField('E_LETTRAGE')) ;
QR2E_JOURNAL       :=TStringField(QEcr.FindField('E_JOURNAL')) ;
QR2J_CENTRALISABLE :=TStringField(QEcr.FindField('J_CENTRALISABLE')) ;
QR2DEBIT           :=TFloatField(QEcr.FindField('DEBIT')) ;
QR2CREDIT          :=TFloatField(QEcr.FindField('CREDIT')) ;
QR2E_NATUREPIECE   :=TStringField(QEcr.FindField('E_NATUREPIECE')) ;
// GC - 03/02/2002
if CritEdt.GL.AvecMemoEcr then
  QR2E_BLOCNOTE    :=TBlobField(QEcr.FindField('E_BLOCNOTE')) ;
// FIN GC - 03/02/2002

ChgMaskChamp(Qr2DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
ChgMaskChamp(Qr2CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
IsECC:=(FDevises.Value<>V_PGI.DevisePivot)and(FMontant.ITemIndex=1)and(QEcr.FindField('E_NATUREPIECE').AsString='ECC') ;
end;

procedure TGdLivGen.BTitreAfterPrint(BandPrinted: Boolean);
begin
  inherited;
//ShowMessage('Avant édit GL : '+DateToStr(CritEdt.Date1)+'   '+DateToStr(CritEdt.Date2)) ;
end;

procedure TGdLivGen.QRDLCENTRNeedData(var MoreData: Boolean);
Var Cod,Lib : String ;
//    TotCentr : Array[0..12] Of Double ;
begin
  inherited;
MoreData:=PrintRecap(LCentr, Cod, Lib, TotCentr) ;
CodCentr.Caption:='RIEN' ;
If MoreData Then
   BEGIN
   Invisible.Caption:=Cod ; CodCentr.Caption:=Lib ;
   TotGenDCentr.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotCentr[0], CritEdt.AfficheSymbole) ;
   TotGenCCentr.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotCentr[1],CritEdt.AfficheSymbole) ;
   TotGenSCentr.Caption:=PrintSolde(TotCentr[0], TotCentr[1], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
   // XX AddReport([3],CritEdt.GL.FormatPrint.Report,TotCentr[0],TotCentr[1],CritEdt.Decimale) ;
   END ;
end;

procedure TGdLivGen.BCentrBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var SJal,SDate1,SDate2,St,SAn,SMois : String ;
    DD1 : TDateTime ;
begin
  inherited;
If CritEdt.GL.SansDetailCentralisation Then BEGIN PrintBand:=FALSE ; Exit ; END ;
// XX
If RepotCent then InitReport([3],CritEDT.GL.FormatPrint.Report) ;
AddReport([3],CritEdt.GL.FormatPrint.Report,TotCentr[0],TotCentr[1],CritEdt.Decimale) ;
St:=Invisible.Caption ;

// GC-GP - 09/01/2002
if TRUE Or (CtxPcl in V_PGI.PgiContexte) then
begin
  SDate1 := Copy(St,1,6);
  SJal  := Copy(St,7,3);
end
else
begin
  SJal:=Copy(St,1,3) ; SDate1:=Copy(St,4,Length(St)-3) ;
end;
// FIN - GC
SMois:=Copy(SDate1,1,2) ; SAn:=Copy(SDate1,3,4) ;
SDate1:=DateToStr(EncodeDate(StrToInt(San),StrToInt(SMois),01)) ;
DD1:=StrTodate(SDate1) ; SDAte2:=DateToStr(FinDeMois(DD1)) ;

if CtxPcl in V_Pgi.PgiContexte then
  Quoi:=Qr1G_GENERAL.AsString+' '+CodCentr.Caption+'#'+TotGenSCentr.Caption+
        '@'+'!;'+Qr1G_General.ASString+';'+SDate1+';'+SDate2+';'+SJal+';'+CritEdt.Exo.Code+';'
else
  Quoi:=Qr1G_GENERAL.AsString+' '+CodCentr.Caption+'#'+TotGenSCentr.Caption+
        '@'+'!;'+Qr1G_General.ASString+';'+SJal+';'+SDate1+';'+SDate2+';'+CritEdt.Exo.Code+';' ;

//Quoi:=Qr1G_GENERAL.AsString+' '+CodCentr.Caption+'#'+TotGenSCentr.Caption+
//      '@'+'!;'+Qr1G_General.ASString+';'+SJal+';'+SDate1+';'+SDate2+';'+CritEdt.Exo.Code+';' ;
end;

procedure TGdLivGen.BCentrAfterPrint(BandPrinted: Boolean);
begin
  inherited;
TitreCentr.Visible:=FALSE ;
RepotCent:=False ;
end;

procedure TGdLivGen.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
//QR2E_GENERAL       :=TStringField(QEcr.FindField('E_GENERAL')) ;
Qr1G_GENERAL        :=TStringField(Q.FindField('G_GENERAL'));
Qr1G_LIBELLE        :=TStringField(Q.FindField('G_LIBELLE'));
Qr1G_COLLECTIF      :=TStringField(Q.FindField('G_COLLECTIF'));
Qr1G_SAUTPAGE       :=TStringField(Q.FindField('G_SAUTPAGE')) ;
Qr1G_SOLDEPROGRESSIF:=TStringField(Q.FindField('G_SOLDEPROGRESSIF')) ;
Qr1G_TOTAUXMENSUELS :=TStringField(Q.FindField('G_TOTAUXMENSUELS')) ;
Qr1G_CENTRALISABLE  :=TStringField(Q.FindField('G_CENTRALISABLE')) ;
If (CritEDT.Rupture=rCorresp) then
   BEGIN
   Qr1G_CORRESP1         :=TStringField(Q.FindField('G_CORRESP1'));
   Qr1G_CORRESP2         :=TStringField(Q.FindField('G_CORRESP2'));
   END ;
end;

procedure TGdLivGen.FSelectCpteChange(Sender: TObject);
begin
  inherited;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
end;

procedure TGdLivGen.FExerciceChange(Sender: TObject);
begin
FLoading:=TRUE ;
  inherited;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
FLoading:=FALSE ;

end;

procedure TGdLivGen.FDateCompta1Change(Sender: TObject);
begin
  inherited;
If FLoading Then Exit ;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
end;

Procedure TGdLivGen.CalculGestion(MO : Array of double ) ;
BEGIN
TotGes[1].TotDebit:=Arrondi(TotGes[1].TotDebit+MO[0],CritEDT.Decimale) ;
TotGes[1].TotCredit:=Arrondi(TotGes[1].TotCredit+MO[1],CritEDT.Decimale) ;
TotGes[2].TotDebit:=Arrondi(TotGes[2].TotDebit+MO[2],CritEDT.Decimale) ;
TotGes[2].TotCredit:=Arrondi(TotGes[2].TotCredit+MO[3],CritEDT.Decimale) ;
END ;

procedure TGdLivGen.RecapPourQui(Cpt : String ; LQui : TStringList ; MO : Array of double) ;
Var I      : Integer ;
BEGIN
For i:=1 to high(VH^.FBIL) do
    if (Cpt>=VH^.FBIL[i].Deb) And (Cpt<=VH^.FBIL[i].Fin) then BEGIN AddRecap(LQui,['LB'],Cpt,MO) ; Break ; Exit END ;
For i:=1 to high(VH^.FCHA) do
    if (Cpt>=VH^.FCHA[i].Deb) And (Cpt<=VH^.FCHA[i].Fin) then BEGIN AddRecap(LQui,['LC'],Cpt,MO) ; Break ; Exit END ;
For i:=1 to high(VH^.FPRO) do
    if (Cpt>=VH^.FPRO[i].Deb) And (Cpt<=VH^.FPRO[i].Fin) then BEGIN AddRecap(LQui,['LP'],Cpt,MO) ; Break ; Exit END ;
For i:=1 to high(VH^.FExt) do
    if (Cpt>=VH^.FExt[i].Deb) And (Cpt<=VH^.FExt[i].Fin) then BEGIN AddRecap(LQui,['LE'],Cpt,MO) ; Break ; Exit END ;
END ;

procedure TGdLivGen.BFCpteGenAfterPrint(BandPrinted: Boolean);
begin
  inherited;
// XX
//If Not CritEdt.GL.SansDetailCentralisation Then InitReport([3],CritEDT.GL.FormatPrint.Report) ;
If Not CritEdt.GL.SansDetailCentralisation Then RepotCent:=True else
   BEGIN
   RepotCent:=False ;
   END ;
InitReport([2],CritEDT.GL.FormatPrint.Report) ;
end;

procedure TGdLivGen.DLRuptHNeedData(var MoreData: Boolean);
begin
  inherited;
(* Gestion En-tête de rupture/corresp
MoreData:=False ;
If OkEnteteRupt then
   BEGIN
   Case CritEdt.GL.PlansCorresp of
     1 : BEGIN MoreData:=Qr1G_CORRESP1.AsString<>'' ;  END ;
     2 : BEGIN MoreData:=Qr1G_CORRESP2.AsString<>'' ;  END ;
     End ;
   END ;
OkEnteteRupt:=False ;
*)
end;

procedure TGdLivGen.BRupTeteBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
(* En-tête de rupture/corresp
PrintBand:=False ;
Case CritEdt.GL.PlansCorresp of
  1 : BEGIN PrintBand:=Qr1G_CORRESP1.AsString<>'' ; CodeRuptAu.Caption:=Qr1G_CORRESP1.AsString+'x' ; END ;
  2 : BEGIN PrintBand:=Qr1G_CORRESP2.AsString<>'' ; CodeRuptAu.Caption:=Qr1G_CORRESP2.AsString+'x' ; END ;
  End ;
*)
end;

procedure TGdLivGen.FSansRuptClick(Sender: TObject);
begin
  inherited;
FLigneRupt.Enabled:=Not FSansRupt.Checked ;
FLigneRupt.checked:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
FRupturesClick(Nil) ;
end;

procedure TGdLivGen.FRupturesClick(Sender: TObject);
begin
  inherited;
if FPlansCo.Checked then FGroupRuptures.Caption:=' '+MsgBox.Mess[24] ;
if FRuptures.Checked then If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
end;

procedure TGdLivGen.DLBilNeedData(var MoreData: Boolean);
Var Tot : Array[0..12] of Double ;
    Cod,Lib, LibMvt, LibGen : String ;
begin
  inherited;
MoreData:=False ;
if (Not CritEdt.GL.BilGes) Or OkZoomEdt then Exit ;
MoreData:=PrintRecap(LBilan,Cod,Lib,Tot) ;
if Cod='LB' then BEGIN LibMvt:=MsgBox.Mess[3] ; LibGen:=MsgBox.Mess[8] ; END Else
if Cod='LC' then BEGIN LibMvt:=MsgBox.Mess[4] ; LibGen:=MsgBox.Mess[9] ; END Else
if Cod='LP' then BEGIN LibMvt:=MsgBox.Mess[5] ; LibGen:=MsgBox.Mess[10] ; END Else
if Cod='LE' then BEGIN LibMvt:=MsgBox.Mess[7] ; LibGen:=MsgBox.Mess[12] ; END ;
If MoreData Then
   BEGIN
   TTotBil.Caption:=LibMvt ; TTotBilGen.Caption:=LibGen ;
   if (Cod='LC') or (Cod='LP') then
      BEGIN OkGes:=True ; CalculGestion(Tot) ; END ;
   TotBilDebit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[0] , CritEDT.AfficheSymbole ) ;
   TotBilCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[1] , CritEDT.AfficheSymbole ) ;
   TotBilSolde.Caption:=PrintSolde(Tot[0],Tot[1],CritEdt.Decimale,CritEdt.Symbole, CritEdt.AfficheSymbole);

   TotGenBilDebit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[2] , CritEDT.AfficheSymbole ) ;
   TotGenBilCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[3] , CritEDT.AfficheSymbole ) ;
   TotGenBilSolde.Caption:=PrintSolde(Tot[2],Tot[3],CritEdt.Decimale,CritEdt.Symbole, CritEdt.AfficheSymbole);
  END ;
end;

procedure TGdLivGen.BGesBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
TotGesDebit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGes[1].TotDebit , CritEDT.AfficheSymbole ) ;
TotGesCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGes[1].TotCredit , CritEDT.AfficheSymbole ) ;
TotGesSolde.Caption:=PrintSolde(TotGes[1].TotDebit,TotGes[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole, CritEdt.AfficheSymbole);
TotGenGesDebit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGes[2].TotDebit , CritEDT.AfficheSymbole ) ;
TotGenGesCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGes[2].TotCredit , CritEDT.AfficheSymbole ) ;
TotGenGesSolde.Caption:=PrintSolde(TotGes[2].TotDebit,TotGes[2].TotCredit,CritEdt.Decimale,CritEdt.Symbole, CritEdt.AfficheSymbole);
end;

procedure TGdLivGen.DLGesNeedData(var MoreData: Boolean);
begin
  inherited;
MoreData:=OkGes ;
OkGes:=False ;
end;

procedure TGdLivGen.Timer1Timer(Sender: TObject);
begin
// GC - 20/12/2001
  inherited;
// GC - 20/12/2001 - FIN
end;

procedure TGdLivGen.FCentralisationClick(Sender: TObject);
begin
  inherited;
  // GC-GP
  if TRUE Or (ctxPcl in V_Pgi.PgiContexte) then
  begin
    FCentralisation.Enabled := not FDetailCentralise.Checked;
    FDetailCentralise.Enabled := not FCentralisation.Checked;
  end;
end;

end.
