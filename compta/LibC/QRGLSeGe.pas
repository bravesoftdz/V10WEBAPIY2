unit QRGLSEGE;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, Hent1, Ent1, UtilEdt, CpteUtil, HQry, QRRupt, SaisUtil, CritEDT,
  Menus, HSysMenu, Calcole, HTB97, HPanel, UiUtil, tCalcCum ;

procedure GLSEGE ;
procedure GLSEGEZoom(Crit : TCritEdt) ;

type
  TFQRGLSEGE = class(TFQR)
    TFJokerGene: THLabel;
    FJokerGene: TEdit;
    TFa2: TLabel;
    FGeneral2: THCpteEdit;
    FGeneral1: THCpteEdit;
    TFGeneral: THLabel;
    HLabel3: THLabel;
    FNumPiece1: TMaskEdit;
    Label2: TLabel;
    FNumPiece2: TMaskEdit;
    HLabel9: THLabel;
    FRefInterne: TEdit;
    HLabel10: THLabel;
    FExceptionG: TEdit;
    FSautPage: TCheckBox;
    FLigneGenEntete: TCheckBox;
    FLigneGenPied: TCheckBox;
    FLigneSecPied: TCheckBox;
    TRGeneral: TQRLabel;
    RGeneral1: TQRLabel;
    TRaG: TQRLabel;
    RGeneral2: TQRLabel;
    QRLabel10: TQRLabel;
    RExceptionG: TQRLabel;
    QRLabel13: TQRLabel;
    RNumPiece1: TQRLabel;
    QRLabel14: TQRLabel;
    RNumPiece2: TQRLabel;
    TRRefInterne: TQRLabel;
    RRefInterne: TQRLabel;
    TValide: TQRLabel;
    TTYPEANALYTIQUE: TQRLabel;
    T_NUMLIGNE: TQRLabel;
    T_GENERAL: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel21: TQRLabel;
    QRLabel16: TQRLabel;
    T_DEBIT: TQRLabel;
    T_CREDIT: TQRLabel;
    QRLabel17: TQRLabel;
    TITRE1REP: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    REPORT1SOLDE: TQRLabel;
    TS_SECTION: TQRLabel;
    BGENEH: TQRBand;
    QRLabel25: TQRLabel;
    TCumul: TQRLabel;
    ANvDEBIT: TQRLabel;
    CumDEBIT: TQRLabel;
    ANvCREDIT: TQRLabel;
    CumCREDIT: TQRLabel;
    CumSolde: TQRLabel;
    ANvSolde: TQRLabel;
    TG_GENERAL: TQRLabel;
    BMVT: TQRBand;
    Y_DATECOMPTABLE: TQRDBText;
    CREDIT: TQRDBText;
    DEBIT: TQRDBText;
    Y_REFINTERNE: TQRDBText;
    Y_VALIDE: TQRDBText;
    Y_LIBELLE: TQRDBText;
    Y_TYPEANALYTIQUE: TQRDBText;
    SOLDE: TQRLabel;
    Y_JOURNAL: TQRDBText;
    E_PIECELIGNE: TQRLabel;
    BSECF: TQRBand;
    TOT1DEBIT: TQRLabel;
    TOT1CREDIT: TQRLabel;
    TOT1SOLDE: TQRLabel;
    TS_SECTION_: TQRLabel;
    BMULTI: TQRBand;
    TDEBMULTI: TQRLabel;
    TCREMULTI: TQRLabel;
    TSOLMULTI: TQRLabel;
    TLibmulti: TQRLabel;
    QRLabel33: TQRLabel;
    TOT2DEBIT: TQRLabel;
    TOT2CREDIT: TQRLabel;
    TOT2SOLDE: TQRLabel;
    QRBand1: TQRBand;
    Trait0: TQRLigne;
    Trait1: TQRLigne;
    Trait2: TQRLigne;
    Trait3: TQRLigne;
    Trait4: TQRLigne;
    Ligne1: TQRLigne;
    TITRE2REP: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    REPORT2SOLDE: TQRLabel;
    MsgLibel: THMsgBox;
    QAnal: TQuery;
    SAna: TDataSource;
    DLMVT: TQRDetailLink;
    DLTETE: TQRDetailLink;
    DLMULTI: TQRDetailLink;
    QGene: TQuery;
    SGene: TDataSource;
    BGenF: TQRBand;
    TG_GENERAL_: TQRLabel;
    TOTSDEBIT: TQRLabel;
    TOTSCREDIT: TQRLabel;
    TOTSSOLDE: TQRLabel;
    QRLabel30: TQRLabel;
    TotDesMvtsDebit: TQRLabel;
    TotDesMvtsCredit: TQRLabel;
    TotDesMvtsSolde: TQRLabel;
    FSoldeProg: TCheckBox;
    FTotMens: TCheckBox;
    FValide: TCheckBox;
    FTypeAnal: TCheckBox;
    TRLegende: TQRLabel;
    RLegende: TQRLabel;
    RValide: TQRLabel;
    QRLabel6: TQRLabel;
    BRupt: TQRBand;
    CodeRupt: TQRLabel;
    DebitRupt: TQRLabel;
    CreditRupt: TQRLabel;
    SoldeRupt: TQRLabel;
    DLRupt: TQRDetailLink;
    AnoRupt: TQRLabel;
    CumRupt: TQRLabel;
    CumDebRupt: TQRLabel;
    AnoDebRupt: TQRLabel;
    AnoCreRupt: TQRLabel;
    CumCreRupt: TQRLabel;
    AnoSolRupt: TQRLabel;
    CumSolRupt: TQRLabel;
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BGENEHBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BMVTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BMVTAfterPrint(BandPrinted: Boolean);
    procedure DLMULTINeedData(var MoreData: Boolean);
    procedure BSECFBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FGeneral1Change(Sender: TObject);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QAnalAfterOpen(DataSet: TDataSet);
    procedure BMULTIBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure BGenFBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QAnalBeforeOpen(DataSet: TDataSet);
    procedure QGeneAfterOpen(DataSet: TDataSet);
    procedure FormShow(Sender: TObject);
    procedure FNatureCptChange(Sender: TObject);
    procedure BSECFAfterPrint(BandPrinted: Boolean);
    procedure BGenFAfterPrint(BandPrinted: Boolean);
    procedure BNouvRechClick(Sender: TObject);
    procedure DLRuptNeedData(var MoreData: Boolean);
    procedure FSansRuptClick(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private    { Déclarations privées }
    Qr11S_SECTION, Qr11S_AXE, Qr11S_SECTIONTRIE,
    Qr11S_LIBELLE, Qr11S_SOLDEPROGRESSIF,
    Qr11S_SAUTPAGE, Qr11S_TOTAUXMENSUELS,
    Qr11S_CORRESP1, Qr11S_CORRESP2            : TStringField;
    Qr11S_TOTALDEBIT                          : TFloatField;
    Qr11S_TOTALCREDIT                         : TFloatField;
    Qr12G_GENERAL                             : TStringField;
    Qr12G_LIBELLE                             : TStringField;
    {-----}
    QR2Y_AXE,QR2Y_SECTION,QR2Y_TYPEANALYTIQUE,
    QR2Y_EXERCICE, QR2Y_GENERAL                : TStringField ;
    QR2Y_DATECOMPTABLE                         : TDateTimeField;
    QR2Y_NUMEROPIECE,QR2Y_NUMLIGNE             : TIntegerField;
    QR2Y_ETABLISSEMENT,QR2Y_VALIDE             : TStringField ;
    QR2Y_QTE1, QR2Y_QTE2                       : TFloatField ;
    QR2Y_LIBELLE, QR2Y_QUALIFPIECE             : TStringField ;
    QR2Y_POURCENTQTE1, QR2Y_POURCENTQTE2       : TFloatField ;
    QR2Y_REFINTERNE,QR2Y_QUALIFQTE1            : TStringField ;
    QR2Y_QUALIFQTE2, QR2Y_JOURNAL              : TStringField ;
    QR2Debit                                   : TFloatField ;
    QR2Credit                                  : TFloatField ;
    TotCptSect, TotCptGene, TotEdt, TotDesMvts : TabTot ;
    L, LRupt                                   : TStringList ;
    ListeCodesRupture                          : TStringList ;
    StReportGen, StReportSec                   : String ;
    OkBand1,OkBand2, Affiche,
    forceaffichepiedcpt,IsECC                  : Boolean ;
    Function  QuoiSec(i : Integer) : String ;
    Function  QuoiGen(i : Integer) : String ;
    Function  QuoiMvt : String ;
    Procedure GLSeGeZoom(Quoi : String) ;
    Procedure GenereSQLSub ;
    Procedure GenereSQLSub2 ;
    Function OkPourRupt : Boolean ;
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

Uses CPSection_TOM ;

procedure GLSEGE ;
var QR : TFQRGLSEGE ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQRGLSEGE.Create(Application) ;
Edition.Etat:=etGlAnaGen ;
QR.QRP.QRPrinter.OnSynZoom:=QR.GLSeGeZoom ;
QR.InitType (nbSec,neGL,msSecAna,'QRGLSEGE','',TRUE,FALSE,TRUE,Edition) ;
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

procedure GLSEGEZoom(Crit : TCritEdt) ;
var QR : TFQRGLSEGE ;
    Edition : TEdition ;
BEGIN
QR:=TFQRGLSEGE.Create(Application) ;
Edition.Etat:=etGlAnaGen ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.GLSeGeZoom ;
 QR.CritEdt:=Crit ;
 QR.InitType (nbSec,neGL,msSecAna,'QRGLSEGE','',FALSE,TRUE,TRUE,Edition) ;
 QR.ShowModal ;
 finally
 QR.Free ;
 end ;
END ;

Function TFQRGLSEGE.QuoiSec(i : Integer) : String ;
BEGIN
Case i Of
  1 : Result:=Qr11S_SECTION.AsString+'#'+Qr11S_LIBELLE.AsString+'@'+'3'+Qr11S_AXE.AsString ;
  2 : Result:=Qr11S_SECTION.AsString+'#'+Qr11S_LIBELLE.AsString+' '+
              TOT1SOLDE.Caption+'@'+'3' ;
 end ;
END ;

Function TFQRGLSEGE.QuoiGen(i : Integer) : String ;
BEGIN
Case i Of
  1 : Result:=Qr12G_GENERAL.AsString+' '+Qr11S_SECTION.AsString+'#'+Qr12G_LIBELLE.AsString+'@'+'1' ;
  2 : Result:=Qr12G_GENERAL.AsString+' '+Qr11S_SECTION.AsString+'#'+Qr12G_LIBELLE.AsString+' '+
              TSOLMULTI.Caption+'@'+'1' ;
 end ;
END ;

Function TFQRGLSEGE.QuoiMvt : String ;
BEGIN
Result:=Qr2Y_GENERAL.AsString+' '+Qr11S_SECTION.AsString+' '+SOLDE.Caption+
        '#'+Qr2Y_JOURNAL.AsString+' N° '+IntToStr(Qr2Y_NUMEROPIECE.AsInteger)+' '+DateToStr(Qr2Y_DateComptable.AsDAteTime)+'-'+
         PrintSolde(Qr2DEBIT.AsFloat,Qr2Credit.AsFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole)+
         '@'+'7;'+Qr2Y_JOURNAL.AsString+';'+UsDateTime(Qr2Y_DateComptable.AsDateTime)+';'+Qr2Y_NUMEROPIECE.AsString+';'+Qr2Y_EXERCICE.asString+';'+
         IntToStr(Qr2Y_NumLigne.AsInteger)+';'+Qr11S_Axe.AsString+';' ;
END ;

Procedure TFQRGLSEGE.GLSeGeZoom(Quoi : String) ;
Var Lp,i: Integer ;
BEGIN
LP:=Pos('@',Quoi) ; If LP=0 Then Exit ;
i:=StrToInt(Copy(Quoi,LP+1,1)) ;
If (i=5) Then  { Band Mvt }
   BEGIN
   Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
   If QRP.QRPRinter.FSynShiftDblClick Then i:=6 ;
   END ;
If (i=7) Then Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
ZoomEdt(i,Quoi) ;
END ;

procedure TFQRGLSEGE.FinirPrint;
BEGIN
Inherited ;
VideGroup(L) ;
QAnal.Close ;
If GCalc<>NIL Then BEGIN GCalc.Free ; GCalc:=NIL ; END ;
if CritEdt.Rupture<>rRien then VideRupt(LRupt) ;
END ;

Function TFQRGLSEGE.OkPourRupt : Boolean ;
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

procedure TFQRGLSEGE.InitDivers ;
BEGIN
Inherited ;
{ Booléen pour l'aafichage de l'entête Du Cpt Géné. }
{ Init du Total Général }
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
TValide.Caption:='' ; TTYPEANALYTIQUE.Caption:='' ;
BGENEH.Frame.DrawTop:=CritEdt.GL.FormatPrint.PrSepCompte[2] ;
BSECF.Frame.DrawTop:=CritEdt.GL.FormatPrint.PrSepCompte[4] ;
BSECF.Frame.DrawBottom:=CritEdt.GL.FormatPrint.PrSepCompte[4] ;
OkBand1:=TRUE ; OkBand2:=TRUE ; forceaffichepiedcpt:=FALSE ;
END ;

procedure TFQRGLSEGE.GenereSQL ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
AlimSQLMul(Q,0) ; Q.Close ; ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
Q.Open ;
GenereSQLSub ; GenereSQLSub2 ;
END;

Procedure TFQRGLSEGE.GenereSQLSub ;
BEGIN
Inherited ;
AlimSQLMul(QGene,1) ; QGene.Close ; ChangeSQL(QGene) ; //QGene.Prepare ;
PrepareSQLODBC(QGene) ;
QGene.Open ;
END ;

Procedure TFQRGLSEGE.GenereSQLSub2 ;
BEGIN
QAnal.Close ;
QAnal.SQL.Clear ;
QAnal.SQL.Add(' Select Y_AXE, Y_SECTION, Y_TYPEANALYTIQUE, Y_GENERAL, ') ;
QAnal.SQL.Add(' Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_REFINTERNE, ') ;
QAnal.SQL.Add(' Y_ETABLISSEMENT, Y_VALIDE, Y_QTE1, Y_QTE2, Y_LIBELLE, ') ;
QAnal.SQL.Add(' Y_POURCENTQTE1, Y_POURCENTQTE2, Y_QUALIFQTE1, Y_QUALIFQTE2, Y_NATUREPIECE, ') ;
QAnal.SQL.Add(' Y_JOURNAL, Y_QUALIFPIECE, ') ;
Case Fmontant.ItemIndex of
  0 : BEGIN QAnal.SQL.Add('Y_DEBIT as DEBIT, Y_CREDIT as CREDIT') ; END ;
  1 : BEGIN QAnal.SQL.Add('Y_DEBITDEV as DEBIT, Y_CREDITDEV as CREDIT') ; END ;
  end ;
QAnal.SQL.Add(' From ANALYTIQ ') ;
{ Construction de la clause Where de la SQL }
QAnal.SQL.Add(' Where ') ;
QAnal.SQL.Add(' Y_AXE="'+CritEDT.GL.Axe+'" ') ;
QAnal.SQL.Add(' And Y_SECTION=:S_SECTION AND Y_GENERAL=:G_GENERAL ') ;
QAnal.SQL.Add(' And Y_EXERCICE="'+CritEdt.Exo.Code+'" ') ;
QAnal.SQL.Add(TraduitNatureEcr(CritEDT.QualifPiece, 'Y_QUALIFPIECE', true,CritEdt.ModeRevision)) ;
QAnal.SQL.Add(' And Y_DATECOMPTABLE>="'+usdatetime(CritEdt.GL.Date21)+'" ');
QAnal.SQL.Add(' And Y_DATECOMPTABLE<="'+usdatetime(CritEdt.Date2)+'" ');
QAnal.SQL.Add(' And Y_NUMEROPIECE>='+IntToStr(CritEdt.GL.NumPiece1)+' and Y_NUMEROPIECE<='+IntToStr(CritEdt.GL.NumPiece2)+' ') ;
if CritEdt.RefInterne<>'' then QAnal.SQL.Add('and UPPER(Y_REFINTERNE) like "'+TraduitJoker(CritEdt.RefInterne)+'" ' );
if CritEdt.Etab<>'' then QAnal.SQL.Add(' And Y_ETABLISSEMENT="'+CritEdt.Etab+'" ') ;
if CritEdt.Valide<>'g' then QAnal.SQL.Add(' And Y_VALIDE="'+CritEdt.Valide+'" ') ;
if CritEdt.GL.AnalPur<>'g' then QAnal.SQL.Add(' And Y_TYPEANALYTIQUE="'+CritEdt.GL.AnalPur+'" ') ;
if CritEdt.DeviseSelect<>'' then QAnal.SQL.Add(' And Y_DEVISE="'+CritEdt.DeviseSelect+'" ') ;
QAnal.SQL.Add(' And Y_ECRANOUVEAU="N" and Y_QUALIFPIECE<>"C" ') ;
{ Construction de la clause Order By de la SQL }
// GG Tri ????
If CritEDT.SQLPlus<>'' Then QAnal.SQL.Add(CritEDT.SQLPlus) ;
QAnal.SQL.Add(' Order By Y_SECTION, Y_GENERAL,Y_AXE, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE ') ;
ChangeSQL(QAnal) ; //QAnal.Prepare ;
PrepareSQLODBC(QAnal) ;
QAnal.Open ;
END ;

procedure TFQRGLSEGE.RenseigneCritere ;
{ Récupération des champs du multicritère dans l'entête d'état }
BEGIN
Inherited ;
If EstSerie(S5) Then RLegende.Caption:=MsgRien.Mess[9] ;
if CritEdt.SJoker then
   BEGIN
   TRGeneral.Caption:=MsgLibel.Mess[5] ;
   RGeneral1.Caption:=FJokerGene.Text ;
   END Else
   BEGIN
   TRGeneral.Caption:=MsgLibel.Mess[4] ;
   RGeneral1.Caption:=CritEdt.LSCpt1 ; RGeneral2.Caption:=CritEdt.LSCpt2 ;
   END ;
RGeneral2.Visible:=Not CritEdt.SJoker ; TRaG.Visible:=Not CritEdt.SJoker ;
RNumPiece1.Caption:=IntToStr(CritEdt.GL.NumPiece1) ; RNumPiece2.Caption:=IntToStr(CritEdt.GL.NumPiece2) ;
RRefInterne.Caption:=FRefInterne.Text ;
RExceptionG.Caption:=FExceptionG.Text ;
CaseACocher(FValide, RValide) ;
(*CaseACocher(FSautPage, RSautPage) ;
CaseACocher(FSoldeProg, RSoldeProg) ;
CaseACocher(FTotMens, RTotMens) ;*)
DateCumulAuGL(TCumul,CritEdt,MsgLibel.Mess[0]) ;
//TCumul.Caption:=MsgLibel.Mess[0]+' '+DateTimeToStr(StrToDateTime(FDateCompta1.Text)-1) ;
END ;

procedure TFQRGLSEGE.ChoixEdition ;
{ Initialisation des options d'édition }
BEGIN
Inherited ;
ChargeGroup(L,['MOIS']) ;
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

END ;

procedure TFQRGLSEGE.RecupCritEdt ;
var st        : String ;
    NonLibres : Boolean ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  SJoker:=FJokerGene.Visible ; Composite:=True ;
  If SJoker Then BEGIN SCpt1:=FJokerGene.Text ; SCpt2:=FJokerGene.Text ; END
            Else BEGIN
                 SCpt1:=FGeneral1.Text  ; SCpt2:=FGeneral2.Text ;
                 PositionneFourchetteSt(FGeneral1,FGeneral2,CritEdt.LSCpt1,CritEdt.LSCpt2) ;
                 END ;
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
  If FNumPiece1.Text<>'' then GL.NumPiece1:=StrToInt(FNumPiece1.Text) else GL.NumPiece1:=0 ;
  If FNumPiece2.Text<>'' then GL.NumPiece2:=StrToInt(FNumPiece2.Text) else GL.NumPiece2:=999999999 ;
  RefInterne:=FRefInterne.text ;
  GL.Axe:=FNatureCpt.Value ;
  GL.SSauf:=Trim(FExceptionG.Text) ;
  If GL.SSauf<>'' then St:=' And '+AnalyseCompte(FExceptionG.Text,fbGene,True,False) ;
  Gl.SSqlSauf:=St ;
  if FTypeAnal.State=cbGrayed then GL.AnalPur:='g' ;
  if FTypeAnal.State=cbChecked then GL.AnalPur:='X' ;
  if FTypeAnal.State=cbUnChecked then GL.AnalPur:='-' ;
  GL.RuptOnly:=QuelleTypeRupt(1,FSAnsRupt.Checked,FAvecRupt.Checked,FALSE) ;
  NonLibres:=((Rupture=rRuptures) or (Rupture=rCorresp)) ;
  if NonLibres then GL.PlanRupt:=FPlanRuptures.Value ;
  If (CritEdt.Rupture=rCorresp) Then GL.PlansCorresp:=FPlanRuptures.ItemIndex+1 ;
  GL.OnlyCptAssocie:=(Rupture<>rRien) and FOnlyCptAssocie.Checked ;
  With GL.FormatPrint Do
     BEGIN
     PrSepCompte[2]:=FLigneGenEntete.Checked ;
     PrSepCompte[3]:=FLigneGenPied.Checked ;
     PrSepCompte[4]:=FLigneSecPied.Checked ;
     END ;
  END ;
END ;

Function TFQRGLSEGE.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK and OkPourRupt ;
If Result Then
   BEGIN
   Gcalc:=TGCalculCum.create(Deux,AxeToFb(CritEdt.GL.Axe),fbGene,QuelTypeEcr,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE) ;
   GCalc.initCalcul('','','',CritEdt.GL.Axe,CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                    CritEdt.Date1,CritEdt.GL.Date11,TRUE) ;
   END ;
END ;

procedure TFQRGLSEGE.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
TITRE1REP.Caption:=TITRE2REP.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
Report1Solde.Caption:=Report2Solde.Caption ;
end;

procedure TFQRGLSEGE.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
Fillchar(TotDesMvts,SizeOf(TotDesMvts),#0) ;
TS_SECTION.Caption:=MsgLibel.Mess[6]+'   '+Qr11S_SECTION.AsString+'   '+Qr11S_LIBELLE.AsString ;
StReportSec:=Qr11S_SECTION.AsString ;
InitReport([2],CritEdt.GL.FormatPrint.Report) ;
Fillchar(TotCptSect,SizeOf(TotCptSect),#0) ;{ Init des variables pour le calcul du Total de chaque Section }
OkBand1:=TRUE ; OkBand2:=TRUE ; forceaffichepiedcpt:=FALSE ;
If (CritEdt.GL.TypCpt=1) Then
   BEGIN
   If Qr12G_GENERAL.IsNull Then BEGIN OkBand1:=FALSE ; OkBand2:=FALSE ; forceaffichepiedcpt:=TRUE ; END ;
   END Else
   BEGIN
   If Qr12G_GENERAL.IsNull Then BEGIN PrintBand:=FALSE ; OkBand1:=FALSE ; OkBand2:=FALSE ; Exit ; END ;
   END ;
Case CritEdt.SautPage of             { Saut de Page O/N Sinon, d'aprés l'info sur le compte}
    0 : BDetail.forceNewPage:=Qr11S_SAUTPAGE.AsString='X' ;
    1 : BDetail.forceNewPage:=True  ;
    2 : BDetail.forceNewPage:=False ;
  End ;
If PrintBand then  // Rony 03/11/97, Pour Cpts associés aux Ruptures
   Case CritEdt.Rupture of
     rLibres    : if CritEdt.GL.OnlyCptAssocie then PrintBand:=DansRuptLibre(Q,AxeToFb(CritEdt.Gl.Axe),CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
     rRuptures  : if CritEdt.GL.OnlyCptAssocie then PrintBand:=(Qr11S_SECTIONTRIE.AsString<>'') ;
     rCorresp   : if CritEdt.GL.OnlyCptAssocie then
                    if CritEDT.GL.PlansCorresp=1 then PrintBand:=(Qr11S_CORRESP1.AsString<>'') Else
                    if CritEDT.GL.PlansCorresp=2 then PrintBand:=(Qr11S_CORRESP2.AsString<>'') ;
     End;
Affiche:=PrintBand ;
If Not PrintBand then Exit ;
Quoi:=QuoiSec(1) ;
(*Case CritEdt.TypCpt of  { 0 : Mvt sur L'exo,    3 : MVt sur la Période }
  0 :  PrintBand:=True ;
  3 :  PrintBand:= TRUE ;
 end ;*)
end;

procedure TFQRGLSEGE.BGENEHBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var D,C : Double ;
    CumulAu : TabTot ;
    CptRupt : String ;
begin
  inherited;
Fillchar(TotCptGene,SizeOf(TotCptGene),#0) ;{ Init des variables pour le calcul du Total de chaque Cpte Géné }
InitReport([3],CritEdt.GL.FormatPrint.Report) ;
Fillchar(CumulAu,SizeOf(CumulAu),#0) ;
PrintBand:=Affiche ; If Not PrintBand then Exit ; // Rony 03/11/97, Pout Cpts associés aux ruptures
If Not OkBand1 Then BEGIN PrintBand:=FALSE ; Exit ; END ;
TG_GENERAL.Caption:=MsgLibel.Mess[7]+'   '+Qr12G_GENERAL.AsString+'   '+Qr12G_LIBELLE.AsString ;
StReportGen:=Qr11S_SECTION.AsString+' / '+Qr12G_GENERAL.AsString ;
GCAlc.ReInitCalcul(Qr11S_SECTION.AsString, Qr12G_GENERAL.AsString,0,0) ;
GCalc.Calcul ; CumulAu:=GCalc.ExecCalc.TotCpt ;
CumulVersSolde(CumulAu[0]) ;
If CritEdt.Date1=CritEdt.GL.Date11 Then Fillchar(CumulAu[1],SizeOf(CumulAu[1]),#0) ;
D:=CumulAu[0].TotDebit+CumulAu[1].TotDebit ;
C:=CumulAu[0].TotCredit+CumulAu[1].TotCredit ;
TotCptSect[1].TotDebit:= Arrondi(TotCptSect[1].TotDebit + D,CritEdt.Decimale) ;
TotCptSect[1].TotCredit:=Arrondi(TotCptSect[1].TotCredit + C,CritEdt.Decimale) ;
TotCptGene[1].TotDebit:= Arrondi(TotCptGene[1].TotDebit + D,CritEdt.Decimale) ;
TotCptGene[1].TotCredit:= Arrondi(TotCptGene[1].TotCredit + C,CritEdt.Decimale) ;
TotEdt[1].TotDebit:=     Arrondi(TotEdt[1].TotDebit + D,CritEdt.Decimale) ;
TotEdt[1].TotCredit:=    Arrondi(TotEdt[1].TotCredit + C,CritEdt.Decimale) ;
Case CritEdt.SoldeProg of             { SI Solde Progressif, donc Init du Solde Progressif O/N Sinon, d'aprés l'info sur le compte }
    0 : If Qr11S_SOLDEPROGRESSIF.AsString='X' then Progressif(False,0,0) ; { Init du Solde Progressif car Solde Prog demandé par le compte }
    1 : Progressif(False,0,0)  ;
   end ;
AnvDebit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,CumulAu[0].TotDebit, CritEdt.AfficheSymbole ) ;
AnvCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,CumulAu[0].TotCredit, CritEdt.AfficheSymbole) ;
AnvSolde.Caption:=PrintSolde(CumulAu[0].TotDebit,CumulAu[0].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;

CumDebit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,D, CritEdt.AfficheSymbole) ;
CumCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,C, CritEdt.AfficheSymbole) ;
CumSolde.Caption:=PrintSolde(D,C,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;

Case CritEdt.SoldeProg of
  0 :  if Qr11S_SOLDEPROGRESSIF.AsString='X' then Progressif(True,D,C) ;
  1 :  Progressif(True,D,C) ;
  end ;
AddReport([1,2,3],CritEdt.GL.FormatPrint.Report,D,C,CritEdt.Decimale) ;
{ Les montants d'Anouveau ne sont pas inclus dans les ruptures }
Case CritEdt.Rupture of
  rLibres  : AddGroupLibre(LRupt,Q,AxeToFb(CritEdt.Gl.Axe),CritEdt.LibreTrie,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
  rRuptures : AddRupt(Lrupt,Qr11S_SECTIONTRIE.AsString,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
  rCorresp : BEGIN
             Case CritEDT.GL.PlansCorresp Of
               1 : If Qr11S_CORRESP1.AsString<>'' Then CptRupt:=Qr11S_CORRESP1.AsString+QR2Y_SECTION.AsString
                                                  Else CptRupt:='.'+QR2Y_SECTION.AsString ;
               2 : If Qr11S_CORRESP2.AsString<>'' Then CptRupt:=Qr11S_CORRESP2.AsString+QR2Y_SECTION.AsString
                                                  Else CptRupt:='.'+QR2Y_SECTION.AsString ;
              End ;
             AddRuptCorres(Lrupt,CptRupt,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
             END ;
  End ;
Quoi:=QuoiGen(1) ;
end;

procedure TFQRGLSEGE.BMVTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var CptRupt : String ;
begin
  inherited;
PrintBand:=Affiche And (Not QR2Y_GENERAL.IsNull)  ;
If Not PrintBand Then Exit ;
If IsECC then
   BEGIN
   DEBIT.Caption:='0' ;
   CREDIT.Caption:='0' ;
   END ;
{ Totalisation Cpte General }
TotCptGene[1].TotDebit:= Arrondi(TotCptGene[1].TotDebit + Qr2DEBIT.AsFloat,CritEDT.Decimale) ;
TotCptGene[1].TotCredit:=Arrondi(TotCptGene[1].TotCredit + Qr2Credit.AsFloat,CritEDT.Decimale) ;
{ Totalisation du Cpte Section }
TotCptSect[1].TotDebit:= Arrondi(TotCptSect[1].TotDebit + QR2DEBIT.AsFloat,CritEdt.Decimale) ;
TotCptSect[1].TotCredit:=Arrondi(TotCptSect[1].TotCredit + QR2Credit.AsFloat,CritEdt.Decimale) ;
{ Total Général (Fin Etat) }
TotEdt[1].TotDebit:=     Arrondi(TotEdt[1].TotDebit + QR2DEBIT.AsFloat,CritEdt.Decimale) ;
TotEdt[1].TotCredit:=    Arrondi(TotEdt[1].TotCredit + QR2Credit.AsFloat,CritEdt.Decimale) ;
{ Total des mvts}
TotDesMvts[1].TotDebit :=Arrondi(TotDesMvts[1].TotDebit + QR2DEBIT.AsFloat,CritEdt.Decimale) ;
TotDesMvts[1].TotCredit:=Arrondi(TotDesMvts[1].TotCredit+QR2CREDIT.AsFloat,CritEdt.Decimale) ;
Y_VALIDE.Caption:=ValiQuali(Qr2Y_VALIDE.AsString,QR2Y_QUALIFPIECE.AsString) ;
If QR2Y_TYPEANALYTIQUE.AsString='X' then  Y_TYPEANALYTIQUE.Caption:='A' else  Y_TYPEANALYTIQUE.Caption:='' ;
E_PIECELIGNE.Caption:=IntToStr(QR2Y_NUMEROPIECE.AsInteger)+'/'+IntToStr(QR2Y_NUMLIGNE.AsInteger) ;
Case CritEdt.SoldeProg of             { Affectation Du Calcul du Solde Progressif O/N Sinon, d'aprés l'info sur le compte }
    0 : If Qr11S_SOLDEPROGRESSIF.AsString='X' then SOLDE.Caption:=PrintSolde(QR2DEBIT.AsFloat+ProgressDebit, QR2CREDIT.AsFloat+ProgressCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole)
                                              Else SOLDE.Caption:=PrintSolde(QR2DEBIT.AsFloat, QR2CREDIT.AsFloat, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
    1 : SOLDE.Caption:=PrintSolde(QR2DEBIT.AsFloat+ProgressDebit, QR2CREDIT.AsFloat+ProgressCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
    2 : SOLDE.Caption:=PrintSolde(QR2DEBIT.AsFloat, QR2CREDIT.AsFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
   end ;
AddGroup(L,[Qr2Y_DATECOMPTABLE],[QR2DEBIT.AsFloat,QR2CREDIT.AsFloat]) ;
AddReport([1,2,3],CritEdt.GL.FormatPrint.Report,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,CritEdt.Decimale) ;
Case CritEdt.Rupture of
  rLibres   : AddGroupLibre(LRupt,Q,AxeToFb(CritEdt.Gl.Axe),CritEdt.LibreTrie,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0]) ;
  rRuptures : AddRupt(Lrupt,Qr11S_SECTIONTRIE.AsString,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0]) ;
  rCorresp  : BEGIN
              Case CritEDT.GL.PlansCorresp Of
                1 : If Qr11S_CORRESP1.AsString<>'' Then CptRupt:=Qr11S_CORRESP1.AsString+QR2Y_SECTION.AsString
                                                   Else CptRupt:='.'+QR2Y_SECTION.AsString ;
                2 : If Qr11S_CORRESP2.AsString<>'' Then CptRupt:=Qr11S_CORRESP2.AsString+QR2Y_SECTION.AsString
                                                   Else CptRupt:='.'+QR2Y_SECTION.AsString ;
               End ;
              AddRuptCorres(Lrupt,CptRupt,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0]) ;
              END ;
  End ;

Quoi:=QuoiMvt ;
end;

procedure TFQRGLSEGE.BMVTAfterPrint(BandPrinted: Boolean);
begin
  inherited;
{ Incrément pour le solde progressif si OK }
Case CritEdt.SoldeProg of             { Calcul du Solde Progressif O/N Sinon, d'aprés l'info sur le compte }
    0 : If Qr11S_SOLDEPROGRESSIF.AsString='X' then Progressif(True,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat) ;
    1 : Progressif(True,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat)  ;
   end ;
end;

procedure TFQRGLSEGE.DLMULTINeedData(var MoreData: Boolean);
Var Cod,Lib : String ;
    Tot : Array[0..12] of Double ;
    Quellerupt : integer ;
begin
  inherited;
MoreData:=PrintGroup(L,QAnal,[Qr2Y_DATECOMPTABLE],Cod,Lib,Tot,Quellerupt);
TLibmulti.Caption:=MsgLibel.Mess[3]+' '+Cod ;
TDEBMULTI.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, Tot[0], CritEdt.AfficheSymbole) ;
TCREMULTI.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, Tot[1], CritEdt.AfficheSymbole) ;
TSOLMULTI.Caption:=PrintSolde(Tot[0],Tot[1],CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
end;

procedure TFQRGLSEGE.BSECFBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=Affiche ; If Not PrintBand then Exit ; // Rony 03/11/97, Pour Cpts associés aux Ruptures
If CritEdt.GL.TypCpt=1 Then
   BEGIN
   If (Not OkBand1) Then
      BEGIN
      If Not forceaffichepiedcpt Then BEGIN PrintBand:=FALSE ; Exit ; END ;
      END ;
   END Else
   BEGIN
   If OkBand1=FALSE Then BEGIN PrintBand:=FALSE ; Exit ; END ;
   END ;
TS_SECTION_.Caption:=MsgLibel.Mess[9]+'   '+Qr11S_SECTION.AsString+'   '+Qr11S_LIBELLE.AsString ;
   { Total et Solde par Section }
TOT1DEBIT.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptSect[1].TotDebit, CritEdt.AfficheSymbole ) ;
TOT1CREDIT.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCptSect[1].TotCredit, CritEdt.AfficheSymbole ) ;
TOT1SOLDE.Caption:=PrintSolde(TotCptSect[1].TotDebit, TotCptSect[1].TotCredit ,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
TotDesMvtsDebit.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotDesMvts[1].TotDebit, CritEdt.AfficheSymbole) ;
TotDesMvtsCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotDesMvts[1].TotCredit, CritEdt.AfficheSymbole) ;
TotDesMvtsSolde.Caption:=PrintSolde(TotDesMvts[1].TotDebit,TotDesMvts[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
Quoi:=QuoiSec(2) ;
end;

procedure TFQRGLSEGE.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TOT2DEBIT.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotEdt[1].TotDebit, CritEdt.AfficheSymbole ) ;
TOT2CREDIT.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotEdt[1].TotCredit, CritEdt.AfficheSymbole ) ;
TOT2SOLDE.Caption:=PrintSolde(TotEdt[1].TotDebit, TotEdt[1].TotCredit ,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFQRGLSEGE.FGeneral1Change(Sender: TObject);
Var AvecJokerG : Boolean ;
begin
  inherited;
AvecJokerG:=Joker(FGeneral1, FGeneral2, FJokerGene ) ;
TFa2.Visible:=Not AvecJokerG ;
TFGeneral.Visible:=Not AvecJokerG ;
TFJokerGene.Visible:=AvecJokerG ;
end;

procedure TFQRGLSEGE.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var D,C : Double ;
begin
  inherited;
Case QuelReport(CritEdt.GL.FormatPrint.Report,D,C) Of
  1 : BEGIN TITREREPORTB.Caption:=MsgLibel.Mess[1] ; Titre2Rep.Caption:='' ; END ;
  2 : BEGIN TITREREPORTB.Caption:=MsgLibel.Mess[2] ; Titre2Rep.Caption:=StReportSec ; END ;
  3 : BEGIN TITREREPORTB.Caption:=MsgLibel.Mess[2] ; Titre2Rep.Caption:=StReportGen ; END ;
  END ;
Report2Debit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, D, CritEdt.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, C, CritEdt.AfficheSymbole ) ;
Report2Solde.Caption:=PrintSolde(D,C,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
end;

procedure TFQRGLSEGE.QAnalAfterOpen(DataSet: TDataSet);
begin
  inherited;
QR2Y_AXE:=TStringField(Qanal.FindField('Y_AXE')) ;
QR2Y_SECTION:=TStringField(Qanal.FindField('Y_SECTION')) ;
QR2Y_GENERAL:=TStringField(Qanal.FindField('Y_GENERAL')) ;
QR2Y_TYPEANALYTIQUE:=TStringField(Qanal.FindField('Y_TYPEANALYTIQUE')) ;
QR2Y_EXERCICE:=TStringField(Qanal.FindField('Y_EXERCICE')) ;
QR2Y_DATECOMPTABLE:=TDateTimeField(Qanal.FindField('Y_DATECOMPTABLE')) ;
QR2Y_DATECOMPTABLE.Tag:=1 ;
QR2Y_NUMEROPIECE:=TIntegerField(Qanal.FindField('Y_NUMEROPIECE')) ;
QR2Y_NUMLIGNE:=TIntegerField(Qanal.FindField('Y_NUMLIGNE')) ;
QR2Y_REFINTERNE:=TStringField(Qanal.FindField('Y_REFINTERNE')) ;
QR2Y_ETABLISSEMENT:=TStringField(Qanal.FindField('Y_ETABLISSEMENT')) ;
QR2Y_VALIDE:=TStringField(Qanal.FindField('Y_VALIDE')) ;
QR2Y_QTE1:=TFloatField(Qanal.FindField('Y_QTE1')) ;
QR2Y_QTE2:=TFloatField(Qanal.FindField('Y_QTE2')) ;
QR2Y_LIBELLE:=TStringField(Qanal.FindField('Y_LIBELLE')) ;
QR2Y_POURCENTQTE1:=TFloatField(Qanal.FindField('Y_POURCENTQTE1')) ;
QR2Y_POURCENTQTE2:=TFloatField(Qanal.FindField('Y_POURCENTQTE2')) ;
QR2Y_REFINTERNE:=TStringField(Qanal.FindField('Y_REFINTERNE')) ;
QR2Y_QUALIFQTE1:=TStringField(Qanal.FindField('Y_QUALIFQTE1')) ;
QR2Y_QUALIFQTE2:=TStringField(Qanal.FindField('Y_QUALIFQTE2')) ;
QR2Y_JOURNAL:=TStringField(Qanal.FindField('Y_JOURNAL')) ;
QR2Y_QUALIFPIECE:=TStringField(Qanal.FindField('Y_QUALIFPIECE')) ;
QR2Debit:=TFloatField(Qanal.FindField('DEBIT')) ;
QR2Credit:=TFloatField(Qanal.FindField('CREDIT')) ;
ChgMaskChamp(Qr2DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
ChgMaskChamp(Qr2CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
IsECC:=(FDevises.Value<>V_PGI.DevisePivot)and(FMontant.ITemIndex=1)and(QAnal.FindField('Y_NATUREPIECE').AsString='ECC') ;
end;

procedure TFQRGLSEGE.BMULTIBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
If OkZoomEdt Then PrintBand:=True Else
   Case CritEdt.TotMens of             { Total Mensuel O/N Sinon, d'aprés l'info sur le compte }
       0 : PrintBand:=(Qr11S_TOTAUXMENSUELS.AsString='X') ;
       1 : PrintBand:=True  ;
       2 : PrintBand:=False  ;
      end ;
end;

procedure TFQRGLSEGE.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr11S_AXE             :=TStringField(Q.FindField('S_AXE'));
Qr11S_SECTION         :=TStringField(Q.FindField('S_SECTION'));
Qr11S_LIBELLE         :=TStringField(Q.FindField('S_LIBELLE'));
Qr11S_SOLDEPROGRESSIF :=TStringField(Q.FindField('S_SOLDEPROGRESSIF'));
Qr11S_SAUTPAGE        :=TStringField(Q.FindField('S_SAUTPAGE'));
Qr11S_TOTAUXMENSUELS  :=TStringField(Q.FindField('S_TOTAUXMENSUELS'));
Qr11S_TOTALDEBIT      :=TFloatField(Q.FindField('S_TOTALDEBIT'));
Qr11S_TOTALCREDIT     :=TFloatField(Q.FindField('S_TOTALCREDIT'));
Qr11S_SECTIONTRIE    :=TStringField(Q.FindField('S_SECTIONTRIE'));
If CritEDT.Rupture=rCorresp then
   BEGIN
   Qr11S_CORRESP1         :=TStringField(Q.FindField('S_CORRESP1'));
   Qr11S_CORRESP2         :=TStringField(Q.FindField('S_CORRESP2'));
   END ;
end;

procedure TFQRGLSEGE.BGenFBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
PrintBand:=Affiche ; If Not PrintBand then Exit ; // Rony 03/11/97, Pour Cpts associés aux Ruptures
If Not OkBand1 Then BEGIN PrintBand:=FALSE ; Exit ; END ;
TG_GENERAL_.Caption:=MsgLibel.Mess[8]+'   '+Qr12G_GENERAL.AsString+'   '+Qr12G_LIBELLE.AsString ;
Quoi:=QuoiGen(2) ;
{ Total et Solde par General }
TOTSDEBIT.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotCptGene[1].TotDebit, CritEDT.AfficheSymbole ) ;
TOTSCREDIT.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotCptGene[1].TotCredit, CritEDT.AfficheSymbole ) ;
TOTSSOLDE.Caption:=PrintSolde(TotCptGene[1].TotDebit, TotCptGene[1].TotCredit ,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
end;

procedure TFQRGLSEGE.QAnalBeforeOpen(DataSet: TDataSet);
begin
  inherited;
QAnal.PArams[0].AsString:=Q.FindField('S_SECTION').AsString ;
end;

procedure TFQRGLSEGE.QGeneAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr12G_GENERAL:=TStringField(QGene.FindField('G_GENERAL'));
Qr12G_LIBELLE:=TStringField(QGene.FindField('G_LIBELLE'));
end;

procedure TFQRGLSEGE.FormShow(Sender: TObject);
begin
FGeneral1.Text:='' ; FGeneral2.Text:='' ;
HelpContext:=7439000 ;
//Standards.HelpContext:=7439010 ;
//Avances.HelpContext:=7439020 ;
//Mise.HelpContext:=7439030 ;
//Option.HelpContext:=7439040 ;
//TabRuptures.HelpContext:=7439050 ;
  inherited;
{$IFDEF CCS3}
FNatureCpt.Visible:=FALSE ; HLabel2.Visible:=FALSE ;
{$ENDIF}
TabSup.TabVisible:=False;
If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
FSurRupt.Visible:=False ;
FCodeRupt1.Visible:=False  ; FCodeRupt2.Visible:=False ;
TFCodeRupt1.Visible:=False ; TFCodeRupt2.Visible:=False ;
FOnlyCptAssocie.Enabled:=False ;
end;

procedure TFQRGLSEGE.FormCreate(Sender: TObject);
begin
  inherited;
ListeCodesRupture:=TStringList.Create ;
end;

procedure TFQRGLSEGE.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
ListeCodesRupture.Free ;
end;

procedure TFQRGLSEGE.FNatureCptChange(Sender: TObject);
begin
  inherited;
If QRLoading then Exit ;
FGeneral1.Clear ; FGeneral2.Clear ; FJokerGene.Clear ;
if FPlansCo.checked then
   BEGIN CorrespToCombo(FPlanRuptures,AxeToFb(FNatureCpt.Value)) ; Exit ; END ;
case FNatureCpt.ItemIndex of
  0 : BEGIN FGeneral1.ZoomTable:=tzGVentil1 ; FGeneral2.ZoomTable:=tzGVentil1 ; END ;
  1 : BEGIN FGeneral1.ZoomTable:=tzGVentil2 ; FGeneral2.ZoomTable:=tzGVentil2 ; END ;
  2 : BEGIN FGeneral1.ZoomTable:=tzGVentil3 ; FGeneral2.ZoomTable:=tzGVentil3 ; END ;
  3 : BEGIN FGeneral1.ZoomTable:=tzGVentil4 ; FGeneral2.ZoomTable:=tzGVentil4 ; END ;
  4 : BEGIN FGeneral1.ZoomTable:=tzGVentil5 ; FGeneral2.ZoomTable:=tzGVentil5 ; END ;
  End ;
Case FNatureCpt.ItemIndex of
   0 : FPlanRuptures.Datatype:='ttRuptSect1' ;
   1 : FPlanRuptures.Datatype:='ttRuptSect2' ;
   2 : FPlanRuptures.Datatype:='ttRuptSect3' ;
   3 : FPlanRuptures.Datatype:='ttRuptSect4' ;
   4 : FPlanRuptures.Datatype:='ttRuptSect5' ;
  end ;
If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
end;

procedure TFQRGLSEGE.BSECFAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([2],CritEdt.GL.FormatPrint.Report) ;
end;

procedure TFQRGLSEGE.BGenFAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([3],CritEdt.GL.FormatPrint.Report) ;
end;

procedure TFQRGLSEGE.BNouvRechClick(Sender: TObject);
begin
  inherited;
If FJokerGene.Visible then FJokerGene.Text:='' ;
end;

procedure TFQRGLSEGE.DLRuptNeedData(var MoreData: Boolean);
Var TotRupt : Array[0..12] of Double ;
    SumD, SumC  : Double ;
    Quellerupt  : Integer ;
    CodRupt, LibRupt, Lib1, CptRupt, stCode     : String ;
    Col                  : TColor ;
    DansTotal, OkOk      : Boolean ;
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
  rRuptures : MoreData:=PrintRupt(Lrupt,Qr11S_SECTIONTRIE.AsString,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt) ;
  rCorresp  : BEGIN
              OkOk:=TRUE ;
              Case CritEDT.GL.PlansCorresp  Of
                1 : If Qr11S_CORRESP1.AsString<>'' Then CptRupt:=Qr11S_CORRESP1.AsString+QR2Y_SECTION.AsString
                                                  Else CptRupt:='.'+QR2Y_SECTION.AsString ;
                2 : If Qr11S_CORRESP2.AsString<>'' Then CptRupt:=Qr11S_CORRESP2.AsString+QR2Y_SECTION.AsString
                                                  Else CptRupt:='.'+QR2Y_SECTION.AsString ;
                Else OkOk:=FALSE ;
                END ;
              If OkOk Then MoreData:=PrintRupt(Lrupt,CptRupt,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt) Else MoreData:=FALSE ;
              END ;
  End ;

If MoreData Then
   BEGIN
   SumD:=Arrondi(TotRupt[5]+ TotRupt[1],CritEdt.Decimale) ;
   SumC:=Arrondi(TotRupt[6]+TotRupt[2],CritEdt.Decimale) ;
   CodeRupt.Caption:='' ;
   Case CritEdt.Rupture of
     rLibres   : BEGIN
                 insert(MsgLibel.Mess[11]+' ',CodRupt,Quellerupt+2) ;
                 CodeRupt.Caption:=CodRupt+' '+Lib1 ;
                 END ;
     rRuptures, rCorresp : CodeRupt.Caption:=CodRupt+' '+LibRupt ;
     End ;
   AnoRupt.Visible:=(CritEdt.Rupture=rCorresp)    ; AnoDebRupt.Visible:=(CritEdt.Rupture=rCorresp) ;
   AnoCreRupt.Visible:=(CritEdt.Rupture=rCorresp) ; AnoSolRupt.Visible:=(CritEdt.Rupture=rCorresp) ;
   CumRupt.Visible:=(CritEdt.Rupture=rCorresp)    ; CumDebRupt.Visible:=(CritEdt.Rupture=rCorresp) ;
   CumCreRupt.Visible:=(CritEdt.Rupture=rCorresp) ; CumSolRupt.Visible:=(CritEdt.Rupture=rCorresp) ;
   If (CritEdt.Rupture=rCorresp) then
      BEGIN
      AnoDebRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[3], CritEdt.AfficheSymbole) ;
      AnoCreRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[4], CritEdt.AfficheSymbole) ;
      AnoSolRupt.Caption:=PrintSolde(TotRupt[3], TotRupt[4], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

      CumDebRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[5], CritEdt.AfficheSymbole) ;
      CumCreRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[6], CritEdt.AfficheSymbole) ;
      CumSolRupt.Caption:=PrintSolde(TotRupt[5], TotRupt[6], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;
      CodeRupt.Top:=37 ; DebitRupt.Top:=37 ; CreditRupt.Top:=37 ; SoldeRupt.Top:=37 ;
      BRupt.Height:=54 ;
      END Else
      BEGIN
      CodeRupt.Top:=3 ; DebitRupt.Top:=3 ; CreditRupt.Top:=3 ; SoldeRupt.Top:=3;
      BRupt.Height:=20 ;
      END ;
   DebitRupt.Caption:= AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumD, CritEDT.AfficheSymbole) ;
   CreditRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumC, CritEDT.AfficheSymbole) ;
   SoldeRupt.Caption:=PrintSolde(SumD, SumC, CritEDT.Decimale, CritEDT.Symbole, CritEDT.AfficheSymbole) ;
   END ;
end;

procedure TFQRGLSEGE.FSansRuptClick(Sender: TObject);
begin
  inherited;
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
FRupturesClick(Nil) ;
end;

procedure TFQRGLSEGE.FRupturesClick(Sender: TObject);
begin
  inherited;
if FPlansCo.Checked then FGroupRuptures.Caption:=' '+MsgLibel.Mess[14] ;
if FRuptures.Checked then FGroupRuptures.Caption:=' '+MsgLibel.Mess[13] ;

end;

end.
