unit QRGLGESE;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, Hent1, Ent1, UtilEdt, CpteUtil, HQry, QRRupt,
  SaisUtil, CritEDT, Menus, HSysMenu, Calcole, HTB97, HPanel, UiUtil, tCalcCum ;

procedure GLGESE ;
procedure GLGESEZoom(Crit : TCritEdt) ;

type
  TFQRGLGESE = class(TFQR)
    HLabel3: THLabel;
    FAxe: THValComboBox;
    HLabel9: THLabel;
    FNumPiece1: TMaskEdit;
    Label2: TLabel;
    FNumPiece2: TMaskEdit;
    HLabel10: THLabel;
    FRefInterne: TEdit;
    FValide: TCheckBox;
    FTypeAnal: TCheckBox;
    HLabel11: THLabel;
    FExceptionS: TEdit;
    FSautPage: TCheckBox;
    FLigneSecEntete: TCheckBox;
    FLigneSecPied: TCheckBox;
    FLigneGenPied: TCheckBox;
    TFSection: THLabel;
    FSection1: THCpteEdit;
    TFaS: TLabel;
    FSection2: THCpteEdit;
    TFJokerSect: THLabel;
    FJokerSect: TEdit;
    TRAxe: TQRLabel;
    RAxe: TQRLabel;
    TRSection: TQRLabel;
    RSection1: TQRLabel;
    TRaS: TQRLabel;
    RSection2: TQRLabel;
    QRLabel13: TQRLabel;
    RNumPiece1: TQRLabel;
    QRLabel14: TQRLabel;
    RNumPiece2: TQRLabel;
    TRRefInterne: TQRLabel;
    RRefInterne: TQRLabel;
    QRLabel3: TQRLabel;
    RExceptionS: TQRLabel;
    TRLegende: TQRLabel;
    RLegende: TQRLabel;
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
    BMULTI: TQRBand;
    TOTMENSD: TQRLabel;
    TOTMENSC: TQRLabel;
    TOTMENSS: TQRLabel;
    BGENEF: TQRBand;
    TG_GENERAL_: TQRLabel;
    TOT1DEBIT: TQRLabel;
    TOT1CREDIT: TQRLabel;
    TOT1SOLDE: TQRLabel;
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
    QAnal: TQuery;
    SAna: TDataSource;
    DLMULTI: TQRDetailLink;
    MsgLibel: THMsgBox;
    Libmulti: TQRLabel;
    BSSec: TQRBand;
    BSectF: TQRBand;
    DLMVT: TQRDetailLink;
    TS_SECTION_: TQRLabel;
    TOTSDEBIT: TQRLabel;
    TOTSCREDIT: TQRLabel;
    TOTSSOLDE: TQRLabel;
    QSec: TQuery;
    SSec: TDataSource;
    TG_GENERAL: TQRLabel;
    QRLabel25: TQRLabel;
    TCumul: TQRLabel;
    ANvDEBIT: TQRLabel;
    CumDEBIT: TQRLabel;
    CumCREDIT: TQRLabel;
    ANvCREDIT: TQRLabel;
    ANvSolde: TQRLabel;
    CumSolde: TQRLabel;
    DLSECT: TQRDetailLink;
    QRLabel36: TQRLabel;
    TotDesMvtsDebit: TQRLabel;
    TotDesMvtsCredit: TQRLabel;
    TotDesMvtsSolde: TQRLabel;
    FSoldeProg: TCheckBox;
    FTotMens: TCheckBox;
    TS_SECTION: TQRLabel;
    QRLabel6: TQRLabel;
    RValide: TQRLabel;
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
    procedure FormShow(Sender: TObject);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BMULTIBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BGENEFBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FSection1Change(Sender: TObject);
    procedure DLMULTINeedData(var MoreData: Boolean);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BMVTAfterPrint(BandPrinted: Boolean);
    procedure BMVTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BSectFBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QAnalAfterOpen(DataSet: TDataSet);
    procedure BSSecBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QAnalBeforeOpen(DataSet: TDataSet);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure QSecAfterOpen(DataSet: TDataSet);
    procedure FAxeChange(Sender: TObject);
    procedure BGENEFAfterPrint(BandPrinted: Boolean);
    procedure BSectFAfterPrint(BandPrinted: Boolean);
    procedure BNouvRechClick(Sender: TObject);
    procedure DLRuptNeedData(var MoreData: Boolean);
    procedure BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FSansRuptClick(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure FJokerSectKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GenereSQL ; Override ;
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
  private    { Déclarations privées }
//    NumR                                     : Integer ;
    L, LRupt                                : TStringList ;
//    QBal                                     : TQuery ;
    StReportGen, StReportSec : String ;
    TotSection,TotGene, TotEdt, TotDesMvts   : TabTot ;
    Qr11G_GENERAL                            : TStringField;
    Qr11G_LIBELLE                            : TStringField;
    Qr11G_SOLDEPROGRESSIF                    : TStringField;
    Qr11G_SAUTPAGE                           : TStringField;
    Qr11G_TOTAUXMENSUELS,
    Qr11G_CORRESP1, Qr11G_CORRESP2           : TStringField;
    Qr11G_TOTALDEBIT                         : TFloatField;
    Qr11G_TOTALCREDIT                        : TFloatField;
    Qr12S_AXE, Qr12S_SECTION                 : TStringField;
    Qr12S_LIBELLE                            : TStringField;
    Qr2Debit,Qr2Credit                       : TFloatField ;
    Qr2Y_Section, Qr2Y_JOURNAL,Qr2Y_EXERCICE : TStringField ;
    Qr2Y_Valide, Qr2Y_TypeAnalytique,
    Qr2Y_QUALIFPIECE, Qr2Y_GENERAL           : TStringField ;
    Qr2Y_NUMEROPIECE,Qr2Y_NUMLIGNE           : TIntegerField ;
    Qr2Y_DATECOMPTABLE                       : TDateTimeField ;
    OkBand1,OkBand2,forceaffichepiedcpt,
    IsECC, Affiche                           : Boolean ;
    Procedure GLGeSeZoom(Quoi : String) ;
    Function QuoiSec(i : Integer) : String ;
    Function QuoiGen(i : Integer) : String ;
    Function QuoiMvt : String ;
    Procedure GenereSQLSub ;
    Procedure GenereSQLSub2 ;
  public     { Déclarations publiques }
  end;

implementation

{$R *.DFM}

procedure GLGESE ;
var QR : TFQRGLGESE ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQRGLGESE.Create(Application) ;
Edition.Etat:=etGlGenAna ;
QR.QRP.QRPrinter.OnSynZoom:=QR.GLGeSeZoom ;
QR.InitType (nbGen,neGL,msGenAna,'QRGLGESE','',TRUE,FALSE,TRUE,Edition) ;
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

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/01/2003
Modifié le ... :   /  /    
Description .. : - 09/01/2003 - On n'affiche pas le fenetre pares l'impression 
Suite ........ : de l'etat
Mots clefs ... : 
*****************************************************************}
procedure GLGESEZoom(Crit : TCritEdt) ;
var QR : TFQRGLGESE ;
    Edition : TEdition ;
BEGIN
QR:=TFQRGLGESE.Create(Application) ;
Edition.Etat:=etGlGenAna ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.GLGeSeZoom ;
 QR.CritEdt:=Crit ;
 QR.InitType (nbGen,neGL,msGenAna,'QRGLGESE','',FALSE,TRUE,TRUE,Edition) ;
 //QR.ShowModal ;
 finally
 QR.Free ;
 end ;
END ;


Function TFQRGLGESE.QuoiSec(i : Integer) : String ;
BEGIN
Inherited ;
Case i Of
  1 : Result:=Qr12S_SECTION.AsString+' '+Qr11G_GENERAL.AsString+'#'+Qr12S_LIBELLE.AsString+'@'+'3'+Qr12S_AXE.AsString ;
  2 : Result:=Qr12S_SECTION.AsString+' '+Qr11G_GENERAL.AsString+'#'+Qr12S_LIBELLE.AsString+' '+
              TOT1SOLDE.Caption+'@'+'3'+Qr12S_AXE.AsString ;
 end ;
END ;

Function TFQRGLGESE.QuoiGen(i : Integer) : String ;
BEGIN
Inherited ;
Case i Of
  1 : Result:=Qr11G_GENERAL.AsString+'#'+Qr11G_LIBELLE.AsString+'@'+'1' ;
  2 : Result:=Qr11G_GENERAL.AsString+'#'+Qr11G_LIBELLE.AsString+' '+
              TOT1SOLDE.Caption+'@'+'1' ;
 end ;
END ;

Function TFQRGLGESE.QuoiMvt : String ;
BEGIN
Inherited ;
Result:=Qr12S_SECTION.AsString+' '+Qr11G_GENERAL.AsString+' '+SOLDE.Caption+
        '#'+Qr2Y_JOURNAL.AsString+' N° '+IntToStr(Qr2Y_NUMEROPIECE.AsInteger)+' '+DateToStr(Qr2Y_DateComptable.AsDAteTime)+'-'+
         PrintSolde(Qr2DEBIT.AsFloat,Qr2Credit.AsFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)+
         '@'+'7;'+Qr2Y_JOURNAL.AsString+';'+UsDateTime(Qr2Y_DateComptable.AsDateTime)+';'+Qr2Y_NUMEROPIECE.AsString+';'+Qr2Y_EXERCICE.asString+';'+
         IntToStr(Qr2Y_NumLigne.AsInteger)+';'+Qr12S_Axe.AsString+';' ;
END ;

Procedure TFQRGLGESE.GLGeSeZoom(Quoi : String) ;
Var Lp,i: Integer ;
BEGIN
Inherited ;
LP:=Pos('@',Quoi) ; If LP=0 Then Exit ;
i:=StrToInt(Copy(Quoi,LP+1,1)) ;
If (i=5) Then  { Band Mvt }
   BEGIN
   Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
   If QRP.QRPrinter.FSynShiftDblClick Then i:=6 ;
   END ;
If (i=7) Then Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
ZoomEdt(i,Quoi) ;
END ;

procedure TFQRGLGESE.FinirPrint;
BEGIN
Inherited ;
VideGroup(L) ;
//QBal.Free ;
if CritEdt.Rupture<>rRien then VideRupt(LRupt) ;
QAnal.Close ;
If GCalc<>NIL Then BEGIN GCalc.Free ; GCalc:=NIL ; END ;
END ;

procedure TFQRGLGESE.InitDivers ;
BEGIN
Inherited ;
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
TValide.Caption:='' ; TTYPEANALYTIQUE.Caption:='' ;
BSSec.Frame.DrawTop:=CritEDT.GL.FormatPrint.PrSepCompte[2] ;
BSectF.Frame.DrawTop:=CritEDT.GL.FormatPrint.PrSepCompte[3] ;
BGENEF.Frame.DrawTop:=CritEDT.GL.FormatPrint.PrSepCompte[4] ;
BGENEF.Frame.DrawBottom:=CritEDT.GL.FormatPrint.PrSepCompte[4] ;
OkBand1:=TRUE ; OkBand2:=TRUE ; forceaffichepiedcpt:=FALSE ;
END ;


Procedure TFQRGLGESE.GenereSQLSub2 ;
BEGIN
QAnal.Close ;
QAnal.SQL.Clear ;
QAnal.SQL.Add(' Select Y_GENERAL, Y_AXE, Y_SECTION, Y_TYPEANALYTIQUE, ') ;
QAnal.SQL.Add(' Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_REFINTERNE, Y_EXERCICE, ') ;
QAnal.SQL.Add(' Y_ETABLISSEMENT, Y_VALIDE, Y_LIBELLE, Y_NATUREPIECE, Y_QUALIFPIECE, ') ;
QAnal.SQL.Add(' Y_JOURNAL, ') ;

Case Fmontant.ItemIndex of
  0 : BEGIN QAnal.SQL.Add('Y_DEBIT DEBIT, Y_CREDIT CREDIT') ; END ;
  1 : BEGIN QAnal.SQL.Add('Y_DEBITDEV DEBIT, Y_CREDITDEV CREDIT') ; END ;
  end ;
QAnal.SQL.Add(' From ANALYTIQ ') ;
{ Construction de la clause Where de la SQL }
QAnal.SQL.Add(' Where ') ;
QAnal.SQL.Add(' Y_AXE="'+CritEDT.GL.Axe+'" ') ;
QAnal.SQL.Add(' And Y_GENERAL=:G_GENERAL AND Y_SECTION=:S_SECTION') ;
QAnal.SQL.Add(' And Y_EXERCICE="'+CritEDT.Exo.Code+'" ') ;
QAnal.SQL.Add(TraduitNatureEcr(CritEDT.QualifPiece, 'Y_QUALIFPIECE', true,CritEdt.ModeRevision)) ;
QAnal.SQL.Add(' And Y_DATECOMPTABLE>="'+usdatetime(CritEdt.GL.Date21)+'" ');
QAnal.SQL.Add(' And Y_DATECOMPTABLE<="'+usdatetime(CritEDT.Date2)+'" ');
QAnal.SQL.Add(' And Y_NUMEROPIECE>='+IntToStr(CritEDT.GL.NumPiece1)+' and Y_NUMEROPIECE<='+IntToStr(CritEDT.GL.NumPiece2)+' ') ;
if CritEDT.RefInterne<>'' then QAnal.SQL.Add('and UPPER(Y_REFINTERNE) like "'+TraduitJoker(CritEDT.RefInterne)+'" ' );
if CritEDT.Etab<>'' then QAnal.SQL.Add(' And Y_ETABLISSEMENT="'+CritEDT.Etab+'" ') ;
if CritEDT.Valide<>'g' then QAnal.SQL.Add(' And Y_VALIDE="'+CritEDT.Valide+'" ') ;
if CritEDT.GL.AnalPur<>'g' then QAnal.SQL.Add(' And Y_TYPEANALYTIQUE="'+CritEDT.GL.AnalPur+'" ') ;
if CritEDT.DeviseSelect<>'' then QAnal.SQL.Add(' And Y_DEVISE="'+CritEDT.DeviseSelect+'" ') ;
QAnal.SQL.Add(' And Y_ECRANOUVEAU="N" and Y_QUALIFPIECE<>"C" ') ;
If CritEDT.SQLPlus<>'' Then QAnal.SQL.Add(CritEDT.SQLPlus) ;
{ Construction de la clause Order By de la SQL }
QAnal.SQL.Add(' Order By Y_GENERAL, Y_SECTION, Y_AXE, Y_DATECOMPTABLE, Y_EXERCICE, Y_NUMEROPIECE, Y_NUMLIGNE ') ;
QAnal.Close ; ChangeSQL(QAnal) ;
//QAnal.Prepare ;
PrepareSQLODBC(QAnal) ;
QAnal.Open ;
END ;                

procedure TFQRGLGESE.GenereSQLSub ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
AlimSQLMul(QSec,1) ; QSec.Close ; ChangeSQL(QSec) ; //QSec.Prepare ;
PrepareSQLODBC(QSec) ;
QSec.Open ;
END;

procedure TFQRGLGESE.GenereSQL ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
AlimSQLMul(Q,0) ; Q.Close ; ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
Q.Open ;
GenereSQLSub ; GenereSQLSub2 ;
END;

procedure TFQRGLGESE.RenseigneCritere ;
{ Récupération des champs du multicritère dans l'entête d'état }
BEGIN
Inherited ;
If EstSerie(S5) Then RLegende.Caption:=MsgRien.Mess[7] ;
RAxe.Caption:=FAxe.Text ;
if CritEdt.SJoker then
   BEGIN
   TRSection.Caption:=MsgLibel.Mess[6] ;
   RSection1.Caption:=FJokerSect.Text ;
   END Else
   BEGIN
   TRSection.Caption:=MsgLibel.Mess[5] ;
   RSection1.Caption:=CritEdt.LSCpt1 ; RSection2.Caption:=CritEdt.LScpt2 ;
   END ;
RSection2.Visible:=Not CritEdt.SJoker ; TRaS.Visible:=Not CritEdt.SJoker ;
RNumPiece1.Caption:=IntToStr(CritEdt.GL.NumPiece1)   ; RNumPiece2.Caption:=IntToStr(CritEdt.GL.NumPiece2) ;
RRefInterne.Caption:=FRefInterne.Text ;
CaseACocher(FValide, RValide) ;
// Rony 10 /04 / 97 CaseACocher(FTypeAnal, RTypeAnal) ;
RExceptionS.Caption:=FExceptionS.Text ;
(*CaseACocher(FSautPage, RSautPage) ; CaseACocher(FSoldeProg, RSoldeProg) ;
CaseACocher(FTotMens, RTotMens)   ; *)
DateCumulAuGL(TCumul,CritEdt,MsgLibel.Mess[0]) ;
//TCumul.Caption:=MsgLibel.Mess[0]+' '+DateTimeToStr(StrToDateTime(FDateCompta1.Text)-1) ;
END ;

procedure TFQRGLGESE.ChoixEdition ;
{ Initialisation des options d'édition }
//Var AvecJokerS  : Boolean ;
//    RDev : RDevise ;
BEGIN
Inherited ;
ChargeGroup(L,['MOIS']) ;
DLRupt.PrintBefore:=TRUE ;
Case CritEdt.Rupture of
  rLibres  :  BEGIN
              DLRupt.PrintBefore:=FALSE ;
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
END ;

procedure TFQRGLGESE.RecupCRITEdt ;
Var ST       : String ;
    NonLibres : Boolean ;
BEGIN
Inherited ;
With CritEDT Do
  BEGIN
  Composite:=TRUE ;
  SJoker:=FJokerSect.Visible ;
  If SJoker Then BEGIN SCpt1:=FJokerSect.Text ; SCpt2:=FJokerSect.Text ; END
            Else BEGIN
                 SCpt1:=FSection1.Text  ; SCpt2:=FSection2.Text ;
                 PositionneFourchetteST(FSection1,FSection2,LSCpt1,LSCpt2) ;
                 END ;
  if FValide.State=cbGrayed then Valide:='g' ;
  if FValide.State=cbChecked then Valide:='X' ;
  if FValide.State=cbUnChecked then Valide:='-' ;
  if FTypeAnal.State=cbGrayed then GL.AnalPur:='g' ;
  if FTypeAnal.State=cbChecked then GL.AnalPur:='X' ;
  if FTypeAnal.State=cbUnChecked then GL.AnalPur:='-' ;
  If FNumPiece1.Text<>'' then GL.NumPiece1:=StrToInt(FNumPiece1.Text) else GL.NumPiece1:=0 ;
  If FNumPiece2.Text<>'' then GL.NumPiece2:=StrToInt(FNumPiece2.Text) else GL.NumPiece2:=999999999 ;
  RefInterne:=FRefInterne.text ;
  GL.Axe:=FAxe.Value ;
  GL.SSauf:=Trim(FExceptionS.Text) ;

  If GL.SSauf<>'' then St:=' And '+AnalyseCompte(FExceptionS.Text,AxeToFb(FAxe.Value),True,False) ;
  Gl.SSqlSauf:=St ;
  GL.RuptOnly:=QuelleTypeRupt(1,FSAnsRupt.Checked,FAvecRupt.Checked,FALSE) ;
  NonLibres:=((Rupture=rRuptures) or (Rupture=rCorresp)) ;
  if NonLibres then GL.PlanRupt:=FPlanRuptures.Value ;
  If (CritEdt.Rupture=rCorresp) Then GL.PlansCorresp:=FPlanRuptures.ItemIndex+1 ;
  With GL.FormatPrint Do
     BEGIN
     PrSepCompte[2]:=FLigneSecEntete.Checked ;
     PrSepCompte[3]:=FLigneSecPied.Checked ;
     PrSepCompte[4]:=FLigneGenPied.Checked ;
     END ;
  END ;
END ;

Function TFQRGLGESE.CRITOk : Boolean ;
BEGIN
Result:=Inherited CritOK ;
If Result Then
   BEGIN
   (*
   QBal:=PrepareTotCptM(fbGene,AxeToFb(FAxe.Value),[],Dev,Etab,Exo) ;
   *)
   Gcalc:=TGCalculCum.create(Deux,fbGene,AxeToFb(CritEdt.GL.Axe),QuelTypeEcr,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE) ;
   GCalc.initCalcul('','','',CritEdt.GL.Axe,CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                    CritEdt.Date1,CritEdt.GL.Date11,TRUE) ;
   END ;
END ;

procedure TFQRGLGESE.FormShow(Sender: TObject);
begin
FAxe.ItemIndex:=0 ;
HelpContext:=7436000 ;
//Standards.HelpContext:=7436010 ;
//Avances.HelpContext:=7436020 ;
//Mise.HelpContext:=7436030 ;
//Option.HelpContext:=7436040 ;
//TabRuptures.HelpContext:=7436050 ;
  inherited;
{$IFDEF CCS3}
FAxe.Visible:=FALSE ; HLabel3.Visible:=FALSE ;
{$ENDIF}
TabSup.TabVisible:=False;
FSection1.Text:='' ; FSection2.Text:='' ;
If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
FSurRupt.Visible:=False ;
FCodeRupt1.Visible:=False  ; FCodeRupt2.Visible:=False ;
TFCodeRupt1.Visible:=False ; TFCodeRupt2.Visible:=False ;
FOnlyCptAssocie.Enabled:=False ;
FSection1.SynJoker:=TRUE ; FSection2.SynJoker:=TRUE ;
end;

procedure TFQRGLGESE.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
Titre1Rep.Caption:=Titre2Rep.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
Report1Solde.Caption:=Report2Solde.Caption ;

end;

procedure TFQRGLGESE.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
Fillchar(TotDesMvts,SizeOf(TotDesMvts),#0) ;
TG_GENERAL.Caption:=MsgLibel.Mess[7]+'   '+Qr11G_GENERAL.AsString+'   '+Qr11G_LIBELLE.AsString ;
StReportGen:=Qr11G_GENERAL.AsString ;
InitReport([2],CritEDT.GL.FormatPrint.Report) ;
Fillchar(TotGene,SizeOf(TotGene),#0) ;{ Init des variables pour le calcul du Total de chaque Cpte Géné }
Case CritEdt.Rupture of
  rLibres    : if CritEdt.GL.OnlyCptAssocie then PrintBand:=DansRuptLibre(Q,fbGene,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
  rRuptures  : if CritEdt.GL.OnlyCptAssocie then PrintBand:=DansRupt(LRupt,Qr11G_GENERAL.AsString) ;
  rCorresp   : if CritEdt.GL.OnlyCptAssocie then
                 if CritEDT.GL.PlansCorresp=1 then PrintBand:=(Qr11G_CORRESP1.AsString<>'') Else
                 if CritEDT.GL.PlansCorresp=2 then PrintBand:=(Qr11G_CORRESP2.AsString<>'') ;
  End;
Affiche:=PrintBand ;
if PrintBand then
   BEGIN
   OkBand1:=TRUE ; OkBand2:=TRUE ; forceaffichepiedcpt:=FALSE ;
   If (CritEdt.GL.TypCpt=1) Then
      BEGIN
      If Qr12S_Section.IsNull Then BEGIN OkBand1:=FALSE ; OkBand2:=FALSE ; forceaffichepiedcpt:=TRUE ; END ;
      END Else
      BEGIN
      If Qr12S_Section.IsNull Then BEGIN PrintBand:=FALSE ; OkBand1:=FALSE ; OkBand2:=FALSE ; Exit ; END ;
      END ;
   Case FSautPage.State of             { Saut de Page O/N Sinon, d'aprés l'info sur le compte}
       cbChecked   : BDetail.forceNewPage:=True  ;
       cbUnChecked : BDetail.forceNewPage:=False ;
       cbGrayed    : BDetail.forceNewPage:=Qr11G_SAUTPAGE.AsString='X' ;
      end ;
   Quoi:=QuoiGen(1) ;
   END ;
end;

procedure TFQRGLGESE.BMULTIBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
(*
Case NumR of
 0 : BEGIN
     Quoi:=QuoiSec(2) ;
     END ;
 1 : BEGIN
     Case FTOTMENS.State of             { Total Mensuel O/N Sinon, d'aprés l'info sur le compte }
         cbChecked   : PrintBand:=True  ;
         cbUnChecked : PrintBand:=False  ;
         cbGrayed    : PrintBand:=(Qr11G_TOTAUXMENSUELS.AsString='X') ;
        end ;
     END ;
 end ;
*)
If OkZoomEdt Then PrintBand:=True Else
   Case FTOTMENS.State of             { Total Mensuel O/N Sinon, d'aprés l'info sur le compte }
       cbChecked   : PrintBand:=True  ;
       cbUnChecked : PrintBand:=False  ;
       cbGrayed    : PrintBand:=(Qr11G_TOTAUXMENSUELS.AsString='X') ;
      end ;
end;

procedure TFQRGLGESE.BGENEFBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=Affiche ; If Not PrintBand then Exit ;// Rony 03/11/97, pour Cpts associés aux ruptures
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
TG_GENERAL_.Caption:=MsgLibel.Mess[10]+'   '+Qr11G_GENERAL.AsString+'   '+Qr11G_LIBELLE.AsString ;
{ Total et Solde par General }
TOT1DEBIT.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGene[1].TotDebit, CritEDT.AfficheSymbole ) ;
TOT1CREDIT.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGene[1].TotCredit, CritEDT.AfficheSymbole ) ;
TOT1SOLDE.Caption:=PrintSolde(TotGene[1].TotDebit, TotGene[1].TotCredit ,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
TotDesMvtsDebit.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotDesMvts[1].TotDebit, CritEdt.AfficheSymbole) ;
TotDesMvtsCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotDesMvts[1].TotCredit, CritEdt.AfficheSymbole) ;
TotDesMvtsSolde.Caption:=PrintSolde(TotDesMvts[1].TotDebit,TotDesMvts[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
Quoi:=QuoiGen(2) ;
end;

procedure TFQRGLGESE.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TOT2DEBIT.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,TotEdt[1].TotDebit, CritEDT.AfficheSymbole ) ;
TOT2CREDIT.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,TotEdt[1].TotCredit, CritEDT.AfficheSymbole ) ;
TOT2SOLDE.Caption:=PrintSolde(TotEdt[1].TotDebit, TotEdt[1].TotCredit ,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFQRGLGESE.FSection1Change(Sender: TObject);
Var AvecJokerS : Boolean ;
begin
  inherited;
AvecJokerS:=Joker(FSection1, FSection2, FJokerSect ) ;
TFaS.Visible:=Not AvecJokerS ;
TFSection.Visible:=Not AvecJokerS ;
TFJokerSect.Visible:=AvecJokerS ;
FExceptionS.Enabled:=Not AvecJokerS ;
end;

procedure TFQRGLGESE.DLMULTINeedData(var MoreData: Boolean);
Var Cod,Lib : String ;
    Tot : Array[0..12] of Double ;
    QuelleRupt : integer ;
begin
  inherited;
MoreData:=PrintGroup(L,QAnal,[Qr2Y_DATECOMPTABLE],Cod,Lib,Tot,Quellerupt);
LibMulti.Caption:=MsgLibel.Mess[4]+' '+Cod ;
TOTMENSD.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[0], CritEDT.AfficheSymbole) ;
TOTMENSC.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[1], CritEDT.AfficheSymbole) ;
TOTMENSS.Caption:=PrintSolde(Tot[0],Tot[1],CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
end;

procedure TFQRGLGESE.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var D,C : Double ;
begin
  inherited;
Case QuelReport(CritEDT.GL.FormatPrint.Report,D,C) Of
  1 : BEGIN TITREREPORTB.Caption:=MsgLibel.Mess[1] ; Titre2Rep.Caption:='' ; END ;
  2 : BEGIN TITREREPORTB.Caption:=MsgLibel.Mess[2] ; Titre2Rep.Caption:=StReportGen ;END ;
  3 : BEGIN TITREREPORTB.Caption:=MsgLibel.Mess[3] ; Titre2Rep.Caption:=StReportSec ;END ;
  END ;
Report2Debit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, D, CritEDT.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, C, CritEDT.AfficheSymbole ) ;
Report2Solde.Caption:=PrintSolde(D,C,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
end;

procedure TFQRGLGESE.BMVTAfterPrint(BandPrinted: Boolean);
begin
  inherited;
{ Incrément pour le solde progressif si OK }
Case FSoldeProg.State of             { Calcul du Solde Progressif O/N Sinon, d'aprés l'info sur le compte }
    cbChecked   : Progressif(True,Qr2DEBIT.AsFloat,Qr2CREDIT.AsFloat)  ;
    cbGrayed    : If Qr11G_SOLDEPROGRESSIF.AsString='X' then Progressif(True,Qr2DEBIT.AsFloat,Qr2CREDIT.AsFloat) ;
   end ;
end;

procedure TFQRGLGESE.BMVTBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var CptRupt : String ;
begin
  inherited;
PrintBand:=Affiche And (Not Qr2Y_SECTION.IsNull) ;
If Not PrintBand Then Exit ;
If IsECC then
   BEGIN
   DEBIT.Caption:='0' ;
   CREDIT.Caption:='0' ;
   END ;
{ Totalisation Section }
TotSection[1].TotDebit:= Arrondi(TotSection[1].TotDebit + Qr2DEBIT.AsFloat,CritEDT.Decimale) ;
TotSection[1].TotCredit:=Arrondi(TotSection[1].TotCredit + Qr2Credit.AsFloat,CritEDT.Decimale) ;
{ Totalisation du Cpte Général }
TotGene[1].TotDebit:= Arrondi(TotGene[1].TotDebit + Qr2DEBIT.AsFloat,CritEDT.Decimale) ;
TotGene[1].TotCredit:=Arrondi(TotGene[1].TotCredit + Qr2Credit.AsFloat,CritEDT.Decimale) ;
{ Total Général (Fin etat) }
TotEdt[1].TotDebit:=     Arrondi(TotEdt[1].TotDebit + Qr2DEBIT.AsFloat,CritEDT.Decimale) ;
TotEdt[1].TotCredit:=    Arrondi(TotEdt[1].TotCredit + Qr2Credit.AsFloat,CritEDT.Decimale) ;
{ Total des mvts}
TotDesMvts[1].TotDebit :=Arrondi(TotDesMvts[1].TotDebit + QR2DEBIT.AsFloat,CritEdt.Decimale) ;
TotDesMvts[1].TotCredit:=Arrondi(TotDesMvts[1].TotCredit+QR2CREDIT.AsFloat,CritEdt.Decimale) ;
Y_VALIDE.Caption:=ValiQuali(Qr2Y_VALIDE.AsString,QR2Y_QUALIFPIECE.AsString) ;
If Qr2Y_TYPEANALYTIQUE.AsString='X' then  Y_TYPEANALYTIQUE.Caption:='A' Else Y_TYPEANALYTIQUE.Caption:='' ;
E_PIECELIGNE.Caption:=IntToStr(Qr2Y_NUMEROPIECE.AsInteger)+'/'+IntToStr(Qr2Y_NUMLIGNE.AsInteger) ;
Case FSoldeProg.State of             { Affectation Du Calcul du Solde Progressif O/N Sinon, d'aprés l'info sur le compte }
    cbChecked   : SOLDE.Caption:=PrintSolde(Qr2DEBIT.AsFloat+ProgressDebit, Qr2CREDIT.AsFloat+ProgressCredit, CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
    cbUnChecked : SOLDE.Caption:=PrintSolde(Qr2DEBIT.AsFloat, Qr2CREDIT.AsFloat,CritEDT.Decimale,CritEDT.FormatMontant,CritEDT.AfficheSymbole) ;
    cbGrayed    : If Qr11G_SOLDEPROGRESSIF.AsString='X' then SOLDE.Caption:=PrintSolde(Qr2DEBIT.AsFloat+ProgressDebit, Qr2CREDIT.AsFloat+ProgressCredit, CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)
                                                     Else SOLDE.Caption:=PrintSolde(Qr2DEBIT.AsFloat, Qr2CREDIT.AsFloat, CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
   end ;
AddGroup(L,[Qr2Y_DATECOMPTABLE],[Qr2DEBIT.AsFloat,Qr2CREDIT.AsFloat]) ;
AddReport([1,2,3],CritEDT.GL.FormatPrint.Report,Qr2DEBIT.AsFloat,Qr2CREDIT.AsFloat,CritEDT.Decimale) ;
Case CritEdt.Rupture of
  rLibres   : AddGroupLibre(LRupt,Q,fbGene,CritEdt.LibreTrie,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0]) ;
  rRuptures : AddRupt(LRupt,Qr2Y_GENERAL.AsString,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0]) ;
  rCorresp  : BEGIN
              Case CritEDT.GL.PlansCorresp Of
                1 : If Qr11G_CORRESP1.AsString<>'' Then CptRupt:=Qr11G_CORRESP1.AsString+Qr2Y_GENERAL.AsString
                                                   Else CptRupt:='.'+Qr2Y_GENERAL.AsString ;
                2 : If Qr11G_CORRESP2.AsString<>'' Then CptRupt:=Qr11G_CORRESP2.AsString+Qr2Y_GENERAL.AsString
                                                   Else CptRupt:='.'+Qr2Y_GENERAL.AsString ;
                Else CptRupt:=Qr2Y_GENERAL.AsString ;
               End ;
              AddRuptCorres(LRupt,CptRupt,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0]) ;
              END ;
  End ;
Quoi:=QuoiMvt ;
end;

procedure TFQRGLGESE.BSectFBeforePrint(var PrintBand: Boolean ; var Quoi: string);
begin
  inherited;
PrintBand:=Affiche ; If Not PrintBand then Exit ;// Rony 03/11/97, Pour cpts associés aux Ruptures
If Not OkBand1 Then BEGIN PrintBand:=FALSE ; Exit ; END ;
Quoi:=QuoiSec(2) ;
TS_SECTION_.Caption:=MsgLibel.Mess[9]+'   '+Qr12S_SECTION.AsString+'   '+Qr12S_LIBELLE.AsString ;
{ Total et Solde par General }
TOTSDEBIT.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotSection[1].TotDebit, CritEDT.AfficheSymbole ) ;
TOTSCREDIT.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotSection[1].TotCredit, CritEDT.AfficheSymbole ) ;
TOTSSOLDE.Caption:=PrintSolde(TotSection[1].TotDebit, TotSection[1].TotCredit ,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
end;

procedure TFQRGLGESE.QAnalAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr2Debit:=TFloatField(Qanal.FindField('DEBIT')) ;
Qr2Credit:=TFloatField(Qanal.FindField('CREDIT')) ;
Qr2Y_Section:=TStringField(QAnal.FindField('Y_SECTION')) ;
Qr2Y_GENERAL:=TStringField(QAnal.FindField('Y_GENERAL')) ;
Qr2Y_JOURNAL:=TStringField(QAnal.FindField('Y_JOURNAL')) ;
Qr2Y_EXERCICE:=TStringField(QAnal.FindField('Y_EXERCICE')) ;
Qr2Y_Valide:=TStringField(QAnal.FindField('Y_VALIDE')) ;
Qr2Y_QUALIFPIECE:=TStringField(QAnal.FindField('Y_QUALIFPIECE')) ;
Qr2Y_TypeAnalytique:=TStringField(QAnal.FindField('Y_TYPEANALYTIQUE')) ;
Qr2Y_NUMEROPIECE:=TIntegerField(QAnal.FindField('Y_NUMEROPIECE')) ;
Qr2Y_NUMLIGNE:=TIntegerField(QAnal.FindField('Y_NUMLIGNE')) ;
Qr2Y_DATECOMPTABLE:=TDateTimeField(QAnal.FindField('Y_DATECOMPTABLE')) ; Qr2Y_DATECOMPTABLE.Tag:=1 ;
ChgMaskChamp(Qr2DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
ChgMaskChamp(Qr2CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
IsECC:=(FDevises.Value<>V_PGI.DevisePivot)and(FMontant.ITemIndex=1)and(QAnal.FindField('Y_NATUREPIECE').AsString='ECC') ;
end;

procedure TFQRGLGESE.BSSecBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var D,C : Double ;
    CumulAu : TabTot ;
    CptRupt : String ;
begin
  inherited;
Fillchar(TotSection,SizeOf(TotSection),#0) ;{ Init des variables pour le calcul du Total de chaque Cpte Géné }
InitReport([3],CritEDT.GL.FormatPrint.Report) ;
Fillchar(CumulAu,SizeOf(CumulAu),#0) ;
PrintBand:=Affiche ; If Not PrintBand then Exit ; //rony 03/11/97, pour cpts associé aux Ruptures
If Not OkBand1 Then BEGIN PrintBand:=FALSE ; Exit ; END ;
Quoi:=QuoiSec(1) ;
TS_SECTION.Caption:=MsgLibel.Mess[8]+'   '+Qr12S_SECTION.AsString+'   '+Qr12S_LIBELLE.AsString ;
StReportSec:=Qr11G_GENERAL.AsString+' / '+Qr12S_SECTION.AsString ;
(*
ExecuteTotCptM(QBal,Qr11G_GENERAL.AsString, Qr12S_SECTION.AsString, CritEDT.Date1,CritEDT.GL.Date11,
             CritEDT.DeviseSelect,CritEDT.Etab,CritEDT.Exo.Code,CumulAu,FALSE) ;
*)
GCAlc.ReInitCalcul(Qr11G_GENERAL.AsString, Qr12S_SECTION.AsString,0,0) ;
GCalc.Calcul ; CumulAu:=GCalc.ExecCalc.TotCpt ;
CumulVersSolde(CumulAu[0]) ;
If CritEDT.Date1=CritEDT.GL.Date11 Then Fillchar(CumulAu[1],SizeOf(CumulAu[1]),#0) ;
D:=CumulAu[0].TotDebit+CumulAu[1].TotDebit ;
C:=CumulAu[0].TotCredit+CumulAu[1].TotCredit ;
TotGene[1].TotDebit :=Arrondi(TotGene[1].TotDebit + D,CritEDT.Decimale) ;
TotGene[1].TotCredit:=Arrondi(TotGene[1].TotCredit + C,CritEDT.Decimale) ;
TotSection[1].TotDebit :=Arrondi(TotSection[1].TotDebit + D,CritEDT.Decimale) ;
TotSection[1].TotCredit:=Arrondi(TotSection[1].TotCredit + C,CritEDT.Decimale) ;
TotEdt[1].TotDebit  :=Arrondi(TotEdt[1].TotDebit + D,CritEDT.Decimale) ;
TotEdt[1].TotCredit :=Arrondi(TotEdt[1].TotCredit + C,CritEDT.Decimale) ;
Case FSoldeProg.State of             { SI Solde Progressif, donc Init du Solde Progressif O/N Sinon, d'aprés l'info sur le compte }
    cbChecked   : Progressif(False,0,0)  ;
    cbGrayed    : If Qr11G_SOLDEPROGRESSIF.AsString='X' then Progressif(False,0,0) ; { Init du Solde Progressif car Solde Prog demandé par le compte }
   end ;
AnvDebit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,CumulAu[0].TotDebit, CritEDT.AfficheSymbole ) ;
AnvCredit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,CumulAu[0].TotCredit, CritEDT.AfficheSymbole) ;
AnvSolde.Caption:=PrintSolde(CumulAu[0].TotDebit,CumulAu[0].TotCredit,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;

CumDebit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,D, CritEDT.AfficheSymbole) ;
CumCredit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,C, CritEDT.AfficheSymbole) ;
CumSolde.Caption:=PrintSolde(D,C,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;

Case FSoldeProg.State of
  cbChecked :  Progressif(True,D,C) ;
  cbGrayed  :  if Qr11G_SOLDEPROGRESSIF.AsString='X' then Progressif(True,D,C) ;
  end ;
AddReport([1,2,3],CritEDT.GL.FormatPrint.Report,D,C,CritEDT.Decimale) ;
{ Il n'y a pas de montants d'Anouveau affectés aux Ruptures }
Case CritEdt.Rupture of
  rLibres   : AddGroupLibre(LRupt,Q,fbGene,CritEdt.LibreTrie,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
  rRuptures : AddRupt(Lrupt,Qr11G_GENERAL.AsString,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
  rCorresp  : BEGIN
              Case CritEDT.GL.PlansCorresp Of
                1 : If Qr11G_CORRESP1.AsString<>'' Then CptRupt:=Qr11G_CORRESP1.AsString+Qr2Y_GENERAL.AsString
                                                   Else CptRupt:='.'+Qr2Y_GENERAL.AsString ;
                2 : If Qr11G_CORRESP2.AsString<>'' Then CptRupt:=Qr11G_CORRESP2.AsString+Qr2Y_GENERAL.AsString
                                                   Else CptRupt:='.'+Qr2Y_GENERAL.AsString ;
               End ;
              AddRuptCorres(LRupt,CptRupt,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
              END ;
  End ;
end;

procedure TFQRGLGESE.QAnalBeforeOpen(DataSet: TDataSet);
begin
  inherited;
QAnal.PArams[0].AsString:=Q.FindField('G_general').AsString ;
end;

procedure TFQRGLGESE.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr11G_GENERAL         :=TStringField(Q.FindField('G_GENERAL'));
Qr11G_LIBELLE         :=TStringField(Q.FindField('G_LIBELLE'));
Qr11G_SOLDEPROGRESSIF :=TStringField(Q.FindField('G_SOLDEPROGRESSIF'));
Qr11G_SAUTPAGE        :=TStringField(Q.FindField('G_SAUTPAGE'));
Qr11G_TOTAUXMENSUELS  :=TStringField(Q.FindField('G_TOTAUXMENSUELS'));
Qr11G_TOTALDEBIT      :=TFloatField(Q.FindField('G_TOTALDEBIT'));
Qr11G_TOTALCREDIT     :=TFloatField(Q.FindField('G_TOTALCREDIT'));
If (CritEDT.Rupture=rCorresp) then
   BEGIN
   Qr11G_CORRESP1         :=TStringField(Q.FindField('G_CORRESP1'));
   Qr11G_CORRESP2         :=TStringField(Q.FindField('G_CORRESP2'));
   END ;
end;

procedure TFQRGLGESE.QSecAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr12S_AXE    :=TStringField(QSec.FindField('S_AXE'));
Qr12S_SECTION:=TStringField(QSec.FindField('S_SECTION'));
Qr12S_LIBELLE:=TStringField(QSec.FindField('S_LIBELLE'));
end;

procedure TFQRGLGESE.FAxeChange(Sender: TObject);
Var St : String ;
     A : Char ;
begin
  inherited;
If QRLoading then Exit ;
FCpte1.Clear ; FCpte2.Clear ; FJoker.Clear ;
FSection1.Clear ; FSection2.Clear ; FJokerSect.Clear ;

ST:=FAxe.Value ; A:=St[2] ;
Case A of
   '1' : BEGIN
         FSection1.ZoomTable:=tzSection ; FSection2.ZoomTable:=tzSection ;
         FCpte1.ZoomTable:=tzGVentil1 ; FCpte2.ZoomTable:=tzGVentil1 ;
         END ;
   '2' : BEGIN
         FSection1.ZoomTable:=tzSection2 ; FSection2.ZoomTable:=tzSection2 ;
         FCpte1.ZoomTable:=tzGVentil2 ; FCpte2.ZoomTable:=tzGVentil2 ;
         END ;
   '3' : BEGIN
         FSection1.ZoomTable:=tzSection3 ; FSection2.ZoomTable:=tzSection3 ;
         FCpte1.ZoomTable:=tzGVentil3 ; FCpte2.ZoomTable:=tzGVentil3 ;
         END ;
   '4' : BEGIN
         FSection1.ZoomTable:=tzSection4 ; FSection2.ZoomTable:=tzSection4 ;
         FCpte1.ZoomTable:=tzGVentil4 ; FCpte2.ZoomTable:=tzGVentil4 ;
         END ;
   '5' : BEGIN
         FSection1.ZoomTable:=tzSection5 ; FSection2.ZoomTable:=tzSection5 ;
         FCpte1.ZoomTable:=tzGVentil5 ; FCpte2.ZoomTable:=tzGVentil5 ;
         END ;
  end ;
end;

procedure TFQRGLGESE.BGENEFAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([2],CritEDT.GL.FormatPrint.Report) ;
end;

procedure TFQRGLGESE.BSectFAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([3],CritEDT.GL.FormatPrint.Report) ;
end;

procedure TFQRGLGESE.BNouvRechClick(Sender: TObject);
begin
  inherited;
If FJokerSect.Visible then FJokerSect.Text:='' ;

end;

procedure TFQRGLGESE.DLRuptNeedData(var MoreData: Boolean);
Var TotRupt : Array[0..12] of Double ;
    SumD, SumC  : Double ;
    Quellerupt  : Integer ;
    CodRupt, LibRupt, Lib1, CptRupt, Stcode     : String ;
    Col         : TColor ;
    OkOk, DansGen    : Boolean ;
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
  rRuptures : MoreData:=PrintRupt(LRupt,Qr11G_GENERAL.AsString,CodRupt,LibRupt,DansGen,QRP.EnRupture,TotRupt) ;
  rCorresp  : BEGIN
              OkOk:=TRUE ;
              Case CritEDT.GL.PlansCorresp  Of
                1 : If Qr11G_CORRESP1.AsString<>'' Then CptRupt:=Qr11G_CORRESP1.AsString+Qr2Y_GENERAL.AsString
                                                   Else CptRupt:='.'+Qr2Y_GENERAL.AsString ;
                2 : If Qr11G_CORRESP2.AsString<>'' Then CptRupt:=Qr11G_CORRESP2.AsString+Qr2Y_GENERAL.AsString
                                                   Else CptRupt:='.'+Qr2Y_GENERAL.AsString ;
                Else OkOk:=FALSE ;
                END ;
              If OkOk Then MoreData:=PrintRupt(LRupt,CptRupt,CodRupt,LibRupt,DansGen,QRP.EnRupture,TotRupt) Else MoreData:=FALSE ;
              END ;
  End ;
(******************************)
If MoreData Then
   BEGIN
   SumD:=Arrondi(TotRupt[5]+ TotRupt[1],CritEdt.Decimale) ;
   SumC:=Arrondi(TotRupt[6]+TotRupt[2],CritEdt.Decimale) ;
   CodeRupt.Caption:='' ;
   Case CritEdt.Rupture of
     rLibres   : BEGIN
                 insert(MsgLibel.Mess[12]+' ',CodRupt,Quellerupt+2) ;
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
      Case CritEDT.GL.PlansCorresp  Of
        1 : CodeRupt.Caption:=Qr11G_CORRESP1.AsString+'x' ;
        2 : CodeRupt.Caption:=Qr11G_CORRESP2.AsString+'x' ;
        End ;

      //CodeRupt.Caption:= ;
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

procedure TFQRGLGESE.BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.Rupture<>rRien) ;
end;

procedure TFQRGLGESE.FSansRuptClick(Sender: TObject);
begin
  inherited;
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
FRupturesClick(Nil) ;
end;

procedure TFQRGLGESE.FRupturesClick(Sender: TObject);
begin
  inherited;
if FPlansCo.Checked then FGroupRuptures.Caption:=' '+MsgLibel.Mess[15] ;
if FRuptures.Checked then FGroupRuptures.Caption:=' '+MsgLibel.Mess[14] ;
end;

procedure TFQRGLGESE.FJokerSectKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var CtrlF5 : Boolean ;
BEGIN
CtrlF5:=(Shift=[ssCtrl]) And (Key=VK_F5) ;
If CtrlF5 Then GereZoneJoker(Sender,'FJokerSect',FSection1) ;
  inherited;
end;

end.
