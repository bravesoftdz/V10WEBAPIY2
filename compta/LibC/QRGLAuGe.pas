unit QRGLAuGe;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, UtilEDT, CritEDT, QRRupt, CpteUtil, EdtLegal,
  Ent1, HEnt1, HQry, HSysMenu, Menus, Calcole, HTB97, HPanel, UiUtil, tCalcCum ;

procedure GLAuxiGen ;
procedure GLAuxiGenZoom(Crit : TCritEdt) ;

type
  TGLAuxGen = class(TFQR)
    HLabel3: THLabel;
    FGen1: THCpteEdit;
    Label1: TLabel;
    FGen2: THCpteEdit;
    FGenJoker: TEdit;
    HLabel9: THLabel;
    FValide: TCheckBox;
    HLabel10: THLabel;
    FNumPiece1: TMaskEdit;
    Label12: TLabel;
    FNumPiece2: TMaskEdit;
    FRefInterne: TEdit;
    HLabel11: THLabel;
    HLabel12: THLabel;
    FExcepS: TEdit;
    FSautPage: TCheckBox;
    FLigneGenEntete: TCheckBox;
    FLigneGenPied: TCheckBox;
    FLigneTiersPied: TCheckBox;
    RGen: TQRLabel;
    RGen1: TQRLabel;
    TRaG: TQRLabel;
    RGen2: TQRLabel;
    QRLabel13: TQRLabel;
    RNumPiece1: TQRLabel;
    QRLabel14: TQRLabel;
    RNumPiece2: TQRLabel;
    QRLabel10: TQRLabel;
    RRefInterne: TQRLabel;
    QRLabel6: TQRLabel;
    RExceptS: TQRLabel;
    TRLegende: TQRLabel;
    RLegende: TQRLabel;
    QRLabel17: TQRLabel;
    RValide: TQRLabel;
    TValide: TQRLabel;
    TE_DATECOMPTABLE: TQRLabel;
    TE_PIECE: TQRLabel;
    TE_REFINTERNE: TQRLabel;
    TE_LIBELLE: TQRLabel;
    TE_DEBIT: TQRLabel;
    TE_CREDIT: TQRLabel;
    TE_SOLDE: TQRLabel;
    TE_LETTRAGE: TQRLabel;
    TITRE1REP: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    REPORT1SOLDE: TQRLabel;
    TT_AUXILIAIRE: TQRLabel;
    BSDHCpteGen: TQRBand;
    TG_GENERAL: TQRLabel;
    AnvDebit: TQRLabel;
    AnvCredit: TQRLabel;
    AnvSolde: TQRLabel;
    CumSolde: TQRLabel;
    CumCredit: TQRLabel;
    CumDebit: TQRLabel;
    QRLabel9: TQRLabel;
    TCumul: TQRLabel;
    BSDMvt: TQRBand;
    LE_DATECOMPTABLE: TQRDBText;
    LE_CREDIT: TQRDBText;
    LE_DEBIT: TQRDBText;
    E_LIBELLE: TQRDBText;
    E_REFINTERNE: TQRDBText;
    LE_LETTRAGE: TQRDBText;
    E_QTE2: TQRDBText;
    LE_SOLDE: TQRLabel;
    LE_VALIDE: TQRLabel;
    E_PIECELIGECH: TQRLabel;
    BSDMulti: TQRBand;
    TLibMulti: TQRLabel;
    TDebMulti: TQRLabel;
    TCreMulti: TQRLabel;
    TSolMulti: TQRLabel;
    BFCpteAux: TQRBand;
    TT_AUXILIAIRE_: TQRLabel;
    TotAuxSolde: TQRLabel;
    TotAuxCredit: TQRLabel;
    TotAuxDebit: TQRLabel;
    QRLabel33: TQRLabel;
    TotalDebit: TQRLabel;
    TotalCredit: TQRLabel;
    TotalSoldes: TQRLabel;
    QROVERLAY: TQRBand;
    Trait0: TQRLigne;
    Trait5: TQRLigne;
    Trait4: TQRLigne;
    Trait3: TQRLigne;
    Trait2: TQRLigne;
    Trait1: TQRLigne;
    Ligne1: TQRLigne;
    TITRE2REP: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    REPORT2SOLDE: TQRLabel;
    SEcr: TDataSource;
    QEcr: TQuery;
    MsgLibel: THMsgBox;
    QRDLCptGenH: TQRDetailLink;
    QRDLMvt: TQRDetailLink;
    QRDLMulti: TQRDetailLink;
    QGene: TQuery;
    SGene: TDataSource;
    BGenF: TQRBand;
    TG_GENERAL_: TQRLabel;
    TOTSDEBIT: TQRLabel;
    TOTSCREDIT: TQRLabel;
    TOTSSOLDE: TQRLabel;
    QRLabel21: TQRLabel;
    TotDesMvtsDebit: TQRLabel;
    TotDesMvtsCredit: TQRLabel;
    TotDesMvtsSolde: TQRLabel;
    TFaA: TLabel;
    FSoldeProg: TCheckBox;
    FTotMens: TCheckBox;
    FTotEche: TCheckBox;
    DLRupt: TQRDetailLink;
    BRupt: TQRBand;
    CodeRupt: TQRLabel;
    DebitRupt: TQRLabel;
    CreditRupt: TQRLabel;
    SoldeRupt: TQRLabel;
    AnoRupt: TQRLabel;
    CumRupt: TQRLabel;
    CumDebRupt: TQRLabel;
    AnoDebRupt: TQRLabel;
    AnoCreRupt: TQRLabel;
    CumCreRupt: TQRLabel;
    AnoSolRupt: TQRLabel;
    CumSolRupt: TQRLabel;
    procedure FormShow(Sender: TObject);
    procedure QRDLMultiNeedData(var MoreData: Boolean);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BSDHCpteGenBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BSDMvtBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BSDMvtAfterPrint(BandPrinted: Boolean);
    procedure BSDMultiBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BFCpteAuxBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure QEcrAfterOpen(DataSet: TDataSet);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure QGeneAfterOpen(DataSet: TDataSet);
    procedure QEcrBeforeOpen(DataSet: TDataSet);
    procedure BGenFBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FNatureCptChange(Sender: TObject);
    procedure BGenFAfterPrint(BandPrinted: Boolean);
    procedure BFCpteAuxAfterPrint(BandPrinted: Boolean);
    procedure BNouvRechClick(Sender: TObject);
    procedure FGen1Change(Sender: TObject);
    procedure DLRuptNeedData(var MoreData: Boolean);
    procedure BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FRupturesClick(Sender: TObject);
    procedure FSansRuptClick(Sender: TObject);
  private    { Déclarations privées }
    Qr11T_AUXILIAIRE, Qr11T_LIBELLE,
    Qr11T_SOLDEPROGRESSIF,Qr11T_SAUTPAGE,
    Qr11T_TOTAUXMENSUELS,
    Qr11T_CORRESP1, Qr11T_CORRESP2               : TStringField;
    Qr11T_TOTALDEBIT                             : TFloatField;
    Qr11T_TOTALCREDIT                            : TFloatField;
    Qr12G_GENERAL, Qr12G_LIBELLE                 : TStringField;
    StReportAux, StReportGen                     : string ;
    NumR, AncienEche                             : Integer ;
    TotGen, TotEdt, TotAux, TotDesMvts           : TabTot ;
    LMulti, LRupt                                : TStringList ;
    QR2E_GENERAL, QR2E_EXERCICE,
    QR2E_REFINTERNE, QR2E_ETABLISSEMENT,
    QR2E_LIBELLE, QR2E_DEVISE, QR2E_AUXILIAIRE,
    QR2E_VALIDE, QR2E_QUALIFQTE1, QR2E_QUALIFPIECE,
    QR2E_QUALIFQTE2, QR2E_LETTRAGE, QR2E_JOURNAL : TStringField ;
    QR2E_DATECOMPTABLE                           : TDateTimeField ;
    QR2E_NUMEROPIECE, QR2E_NUMLIGNE,
    QR2E_NUMECHE, QR2MULTIECHE                   : TIntegerField ;
    QR2E_QTE1, QR2E_QTE2, QR2DEBIT, QR2CREDIT    : TFloatField ;
    OkBand1,OkBand2,forceaffichepiedcpt,IsECC,
    OkEnteteRupt,Affiche                                 : Boolean ;
    Function  QuoiGen(i : Integer) : String ;
    Function  QuoiAux(i : Integer) : String ;
    Function  QuoiMvt : String ;
    Procedure GLAuGeZoom(Quoi : String) ;
    Procedure GenereSQLSub2 ;
    procedure GenereSQLSub ;
  public     { Déclarations publiques }
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

procedure GLAuxiGen ;
var QR : TGLAuxGen ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TGLAuxGen.Create(Application) ;
Edition.Etat:=etGlAuxGen ;
QR.QRP.QRPrinter.OnSynZoom:=QR.GLAuGeZoom ;
QR.InitType(nbAux,neGl,msAuxEcr,'QRGLAUGE','',TRUE,FALSE,TRUE,Edition) ;
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

procedure GLAuxiGenZoom(Crit : TCritEdt) ;
var QR : TGLAuxGen ;
    Edition : TEdition ;
BEGIN
QR:=TGLAuxGen.Create(Application) ;
Edition.Etat:=etGlAuxGen ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.GLAuGeZoom ;
 QR.CritEdt:=Crit ;
 QR.InitType(nbAux,neGl,msAuxEcr,'QRGLAUGE','',FALSE,TRUE,TRUE,Edition) ;
 QR.ShowModal ;
 finally
 QR.Free ;
 end ;
END ;

Function TGLAuxGen.QuoiAux(i : Integer) : String ;
BEGIN
Case i Of
  1 : Result:=Qr11T_AUXILIAIRE.AsString+'#'+Qr11T_LIBELLE.AsString+'@'+'2' ;
  2 : Result:=Qr11T_AUXILIAIRE.AsString+'#'+Qr11T_LIBELLE.AsString+' '+
              TotAuxSolde.Caption+'@'+'2' ;
 end ;
END ;

Function TGLAuxGen.QuoiGen(i : Integer) : String ;
BEGIN
Case i Of
  1 : Result:=QR2E_GENERAL.AsString+' '+Qr11T_AUXILIAIRE.AsString+'#'+Qr12G_LIBELLE.AsString+'@'+'1' ;
  2 : Result:=QR2E_GENERAL.AsString+' '+Qr11T_AUXILIAIRE.AsString+'#'+Qr12G_LIBELLE.AsString+' '+
              TotAuxSolde.Caption+'@'+'1' ;
 end ;
END ;

Function TGLAuxGen.QuoiMvt : String ;
BEGIN
Result:=QR2E_GENERAL.AsString+' '+Qr11T_LIBELLE.AsString+' '+LE_SOLDE.Caption+
        '#'+QR2E_JOURNAL.AsString+' N° '+IntToStr(QR2E_NUMEROPIECE.AsInteger)+' '+DateToStr(QR2E_DateComptable.AsDAteTime)+'-'+
         PrintSolde(QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole)+
         '@'+'5;'+QR2E_JOURNAL.AsString+';'+UsDateTime(QR2E_DATECOMPTABLE.AsDateTime)+';'+QR2E_NUMEROPIECE.AsString+';'+QR2E_EXERCICE.asString+';'+
         IntToStr(QR2E_NumLigne.AsInteger)+';' ;
END ;

Procedure TGLAuxGen.GLAuGeZoom(Quoi : String) ;
Var Lp,i: Integer ;
BEGIN
LP:=Pos('@',Quoi) ; If LP=0 Then Exit ;
i:=StrToInt(Copy(Quoi,LP+1,1)) ;
If (i=5) Then
   BEGIN
   Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
   If QRP.QRPrinter.FSynShiftDblClick Then i:=6 ;
   If QRP.QRPrinter.FSynCtrlDblClick Then i:=11 ;
   END ;
ZoomEdt(i,Quoi) ;
END ;

procedure TGLAuxGen.FinirPrint;
BEGIN
Inherited ;
QEcr.Close ; VideGroup(LMulti) ;
If GCalc<>NIL Then BEGIN GCalc.Free ; GCalc:=NIL ; END ;
if CritEdt.Rupture<>rRien then VideRupt(LRupt) ;
END ;

procedure TGLAuxGen.GenereSQL ;
BEGIN
Inherited ;
AlimSQLMul(Q,0) ; Q.Close ; ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ; Q.Open ;
GenereSQLSub ; GenereSQLSub2 ;
END ;

procedure TGLAuxGen.GenereSQLSub ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
AlimSQLMul(QGene,1) ; QGene.Close ; ChangeSQL(QGene) ; //QGene.Prepare ;
PrepareSQLODBC(QGene) ;
QGene.Open ;
END ;

Procedure TGLAuxGen.GenereSQLSub2 ;
BEGIN
QEcr.Close ;
QEcr.SQL.Clear ;
QEcr.SQL.Add('Select E_AUXILIAIRE,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_JOURNAL, ') ;
QEcr.SQL.Add(      ' E_REFINTERNE,E_ETABLISSEMENT,E_LIBELLE,E_NUMECHE,E_DEVISE,E_VALIDE, ') ;
QEcr.SQL.Add(      ' E_QTE1,E_QTE2,E_QUALIFQTE1,E_QUALIFQTE2,E_LETTRAGE,E_GENERAL, E_NATUREPIECE, ') ;
// GP : bug edition avec filtre sur N° pièce trop grand (à cause du *10)
//QEcr.SQL.Add(      ' E_NUMEROPIECE*10+E_NUMLIGNE as MULTIECHE, E_QUALIFPIECE,') ;
QEcr.SQL.Add(      ' E_NUMLIGNE as MULTIECHE, E_QUALIFPIECE,') ;

Case Fmontant.ItemIndex of
  0 : BEGIN QEcr.SQL.Add(' E_DEBIT DEBIT,E_CREDIT CREDIT ') ; END ;
  1 : BEGIN QEcr.SQL.Add(' E_DEBITDEV DEBIT,E_CREDITDEV CREDIT ') ; END ;
end ;

{ Tables explorées par la SQL }
QEcr.SQL.Add(' From ECRITURE ') ;
{ Construction de la clause Where de la SQL }
QEcr.SQL.Add(' Where E_AUXILIAIRE=:T_AUXILIAIRE AND E_GENERAL=:G_GENERAL ') ;
QEcr.SQL.Add('And E_DATECOMPTABLE>="'+usdatetime(CritEdt.GL.Date21)+'" And E_DATECOMPTABLE<="'+usdatetime(CritEdt.Date2)+'" ') ;
if FExercice.ItemIndex<>0 then QEcr.SQL.Add(' And E_EXERCICE="'+CritEdt.Exo.Code+'" ') ;
QEcr.SQL.Add(   ' And E_NUMEROPIECE>='+IntToStr(CritEdt.GL.NumPiece1)+' ') ;
QEcr.SQL.Add(   ' And E_NUMEROPIECE<='+IntToStr(CritEdt.GL.NumPiece2)+' ') ;
QEcr.SQL.Add(TraduitNatureEcr(CritEDT.QualifPiece, 'E_QUALIFPIECE', true,CritEdt.ModeRevision)) ;
if FRefInterne.Text<>'' then QEcr.SQL.Add(' And UPPER(E_REFINTERNE) like "'+TraduitJoker(CritEdt.RefInterne)+'" ' );
if FEtab.ItemIndex<>0 then QEcr.SQL.Add(' And E_ETABLISSEMENT="'+CritEdt.Etab+'"') ;
if FValide.State<>cbGrayed then QEcr.SQL.Add(' And E_VALIDE="'+CritEdt.Valide+'" ') ;
if FDevises.ItemIndex<>0 then QEcr.SQL.Add(' And E_DEVISE="'+CritEdt.DeviseSelect+'"') ;
QEcr.SQL.Add(' And E_ECRANOUVEAU="N" and E_QUALIFPIECE<>"C" ') ;
{ Construction de la clause Order By de la SQL }
If CritEDT.SQLPlus<>'' Then QEcr.SQL.Add(CritEDT.SQLPlus) ;
QEcr.SQL.Add(' Order By E_AUXILIAIRE,E_GENERAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE ') ;
ChangeSQL(QEcr) ; QEcr.Open ;
END ;

procedure TGLAuxGen.RenseigneCritere ;
{ Récupération des champs du multicritère dans l'entête d'état }
BEGIN
Inherited ;
If OkZoomEdt Then Exit ;
If EstSerie(S5) Then RLegende.Caption:=MsgRien.Mess[7] ;
if CritEdt.SJoker then
   BEGIN
   RGen.Caption:=MsgLibel.Mess[9] ;
   RGen1.Caption:=FGenJoker.Text ;
   END else
   BEGIN
   RGen.Caption:=MsgLibel.Mess[8] ;
   Rgen1.Caption:=CritEdt.LSCpt1 ; RGen2.Caption:=CritEdt.LSCpt2 ;
   END ;
RGen.Visible:=Not CritEdt.SJoker ; RGen2.Visible:=Not CritEdt.SJoker ;
RRefInterne.Caption:=FRefInterne.Text ;
RNumPiece1.Caption:=IntToStr(CritEdt.GL.NumPiece1) ; RNumPiece2.Caption:=IntToStr(CritEdt.GL.NumPiece2) ;
RExceptS.Caption:=FExcepS.Text ;
CaseACocher(FValide,RValide)            ;
(*
CaseACocher(FSoldeProg,RSoldeProg) ;
CaseACocher(FSautPage,RSautPage)        ; CaseACocher(FTotMens,RTotMens) ;
CaseACocher(FTotEche,RTotEche)          ; CaseACocher(FLigneGenEntete,RLigneGenEntete) ;
CaseACocher(FLigneGenPied ,RLigneGenPied) ; CaseACocher(FLigneTiersPied,RLigneTiersPied) ;
*)
DateCumulAuGL(TCumul,CritEdt,MsgLibel.Mess[5]) ;
END ;

procedure TGLAuxGen.ChoixEdition ;
BEGIN
Inherited ;
ChargeGroup(LMulti,['MOIS','ECHE']) ;
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
END ;

procedure TGLAuxGen.InitDivers ;
BEGIN
Inherited ;
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
TValide.Caption:='' ;
{ Gestion multiéchéances }
AncienEche:=1 ;
{ Séparateurs de bandes }
BSDHCpteGen.Frame.DrawTop:=CritEdt.GL.FormatPrint.PrSepCompte[2] ;
BFCpteAux.Frame.DrawTop:=CritEdt.GL.FormatPrint.PrSepCompte[4] ;
BFCpteAux.Frame.DrawBottom:=CritEdt.GL.FormatPrint.PrSepCompte[4] ;
OkBand1:=TRUE ; OkBand2:=TRUE ; forceaffichepiedcpt:=FALSE ;
END ;

procedure TGLAuxGen.RecupCritEdt ;
var st : String ;
    NonLibres : Boolean ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  SJoker:=FGenJoker.Visible ; Composite:=True ;
  if SJoker Then BEGIN SCpt1:=FGenJoker.Text ; SCpt2:=FGenJoker.Text ; END
            Else BEGIN
                 SCpt1:=FGen1.Text ;  SCpt2:=FGen2.Text ;
                 PositionneFourchetteSt(FGen1,FGen2,CritEdt.LSCpt1,CritEdt.LSCpt2) ;
                 END ;
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
  GL.TotEche:=FTotEche.Checked ;
  GL.SSauf:=Trim(FExcepS.Text) ; St:=GL.SSauf ;
  if GL.SSauf<>'' Then St:=' And '+AnalyseCompte(FExcepS.Text,fbGene,True,False) ;
  Gl.SSqlSauf:=St ;
  GL.RuptOnly:=QuelleTypeRupt(1,FSAnsRupt.Checked,FAvecRupt.Checked,FALSE) ;
  NonLibres:=((Rupture=rRuptures) or (Rupture=rCorresp)) ;
  if NonLibres then GL.PlanRupt:=FPlanRuptures.Value ;
  If (CritEdt.Rupture=rCorresp) Then GL.PlansCorresp:=FPlanRuptures.ItemIndex+1 ;
  GL.OnlyCptAssocie:=(Rupture<>rRien) and FOnlyCptAssocie.Checked ;
   With GL.FormatPrint Do
    BEGIN
    PrSepCompte[2]:=FLigneGenEntete.Checked ;
    PrSepCompte[3]:=FLigneGenPied.Checked ;
    PrSepCompte[4]:=FLigneTiersPied.Checked ;
    END ;
  END ;
END ;

function  TGLAuxGen.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ;
If Result Then
   BEGIN
   Gcalc:=TGCalculCum.create(Deux,fbAux,fbGene,QuelTypeEcr,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE) ;
   GCalc.initCalcul('','','','',CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                    CritEdt.Date1,CritEdt.GL.Date11,TRUE) ;
   END ;
END ;

procedure TGLAuxGen.FormShow(Sender: TObject);
begin
HelpContext:=7430000 ;
//Standards.HelpContext:=7430010 ;
//Avances.HelpContext:=7430020 ;
//Mise.HelpContext:=7430030 ;
//Option.HelpContext:=7430040 ;
//TabRuptures.HelpContext:=7430050 ;
  inherited;
TabSup.TabVisible:=False;
FGen1.Text:='' ; FGen2.Text:='' ; FRefInterne.Text:='' ;
FValide.Checked:=True ; FValide.State:=cbGrayed ;
If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
FSurRupt.Visible:=False ;
FCodeRupt1.Visible:=False  ; FCodeRupt2.Visible:=False ;
TFCodeRupt1.Visible:=False ; TFCodeRupt2.Visible:=False ;
FOnlyCptAssocie.Enabled:=False ;
end;

procedure TGLAuxGen.QRDLMultiNeedData(var MoreData: Boolean);
var LibMulti, CodMulti : string ;
    TotMulti           : Array[0..12] of Double ;
begin
  inherited;
MoreData:=PrintGroup(LMulti, QEcr, [QR2E_DATECOMPTABLE , QR2MULTIECHE], CodMulti, LibMulti, TotMulti, NumR) ;
if NumR=0 then
   BEGIN
   BSDMulti.Frame.DrawTop:=False ;
   BSDMulti.Font.Color:=clRed ;
   TLibmulti.Caption:=MsgLibel.Mess[3]+' '+CodMulti ;
   TDebMulti.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotMulti[0], CritEdt.AfficheSymbole) ;
   TCreMulti.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotMulti[1], CritEdt.AfficheSymbole) ;
   TSolMulti.Caption:=PrintSolde(TotMulti[0],TotMulti[1],CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
   END Else
if NumR=1 then
   BEGIN
   BSDMulti.Frame.DrawTop:=False ;
   BSDMulti.Font.Color:=clPurple ;
   TLibmulti.Caption:=MsgLibel.Mess[4]+' '+IntToStr(QR2E_NUMEROPIECE.AsInteger)+' / '+QR2E_NUMLIGNE.AsString ;
   TDebMulti.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotMulti[0], CritEdt.AfficheSymbole) ;
   TCreMulti.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotMulti[1], CritEdt.AfficheSymbole) ;
   TSolMulti.Caption:='' ;
   END ;
end;

procedure TGLAuxGen.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
Fillchar(TotDesMvts,SizeOf(TotDesMvts),#0) ;
TT_AUXILIAIRE.Caption:=MsgLibel.Mess[10]+'   '+Qr11T_AUXILIAIRE.AsString+'   '+Qr11T_LIBELLE.AsString ;
StReportAux:=Qr11T_AUXILIAIRE.AsString ;
InitReport([2],CritEdt.GL.FormatPrint.Report) ;
Fillchar(TotAux,SizeOf(TotAux),#0) ;
Case CritEdt.Rupture of
  rLibres    : if CritEdt.GL.OnlyCptAssocie then PrintBand:=DansRuptLibre(Q,fbAux,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
  rRuptures  : if CritEdt.GL.OnlyCptAssocie then PrintBand:=DansRupt(LRupt,Qr11T_AUXILIAIRE.AsString) ;
  rCorresp   : if CritEdt.GL.OnlyCptAssocie then
                 if CritEDT.GL.PlansCorresp=1 then PrintBand:=(Qr11T_CORRESP1.AsString<>'') Else
                 if CritEDT.GL.PlansCorresp=2 then PrintBand:=(Qr11T_CORRESP2.AsString<>'') ;
  End;
If PrintBand then
   BEGIN
   OkBand1:=TRUE ; OkBand2:=TRUE ; forceaffichepiedcpt:=FALSE ;
   If (CritEdt.GL.TypCpt=1) Then
      BEGIN
      If Qr12G_GENERAL.IsNull Then BEGIN OkBand1:=FALSE ; OkBand2:=FALSE ; forceaffichepiedcpt:=TRUE ; END ;
      END Else
      BEGIN
      If Qr12G_GENERAL.IsNull Then BEGIN PrintBand:=FALSE ; OkBand1:=FALSE ; OkBand2:=FALSE ; Exit ; END ;
      END ;
   END ;
Affiche:=PrintBand ; { Rony 03/11/97, Pour ne pas afficher les sous comptes }
If Not PrintBand then Exit ;
Case CritEdt.SautPage of
  0 : BDetail.forceNewPage:=(Qr11T_SAUTPAGE.AsString='X') ;
  1 : BDetail.forceNewPage:=True ;
  2 : BDetail.forceNewPage:=False ;
 end ;
Quoi:=QuoiAux(1) ;
end;

procedure TGLAuxGen.BSDHCpteGenBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var D, C : Double ;
    CumulAu : TabTot ;
    CptRupt : String ;
begin
  inherited;
PrintBand:=Affiche ; If Not PrintBand then Exit ; // Rony 03/11/97, Pour cpts associés aux Ruptures
If not PrintBand Then Exit ;
TG_GENERAL.Caption:=MsgLibel.Mess[11]+'   '+Qr12G_GENERAL.AsString+'   '+Qr12G_LIBELLE.AsString ;
Fillchar(TotGen,SizeOf(TotGen),#0) ;
Quoi:=QuoiGen(1) ;
InitReport([3],CritEdt.GL.FormatPrint.Report) ; StReportGen:=Qr11T_AUXILIAIRE.AsString+' / '+Qr12G_GENERAL.AsString ;
Fillchar(CumulAu,SizeOf(CumulAu),#0) ;
If Not OkBand1 Then BEGIN PrintBand:=FALSE ; Exit ; END ;

GCAlc.ReInitCalcul(Qr11T_AUXILIAIRE.AsString, Qr12G_GENERAL.AsString,0,0) ;
GCalc.Calcul ; CumulAu:=GCalc.ExecCalc.TotCpt ;
CumulVersSolde(CumulAu[0]) ;
If CritEdt.Date1=CritEdt.GL.Date11 Then Fillchar(CumulAu[1],SizeOf(CumulAu[1]),#0) ;
D:=CumulAu[0].TotDebit+CumulAu[1].TotDebit ;
C:=CumulAu[0].TotCredit+CumulAu[1].TotCredit ;
// Rony 07/04/1997  ---> TotGen
TotGen[1].TotDebit:= Arrondi(TotGen[1].TotDebit + D,CritEdt.Decimale) ;
TotGen[1].TotCredit:= Arrondi(TotGen[1].TotCredit + C,CritEdt.Decimale) ;
TotAux[1].TotDebit:= Arrondi(TotAux[1].TotDebit + D,CritEdt.Decimale) ;
TotAux[1].TotCredit:=Arrondi(TotAux[1].TotCredit + C,CritEdt.Decimale) ;
TotEdt[1].TotDebit:=     Arrondi(TotEdt[1].TotDebit + D,CritEdt.Decimale) ;
TotEdt[1].TotCredit:=    Arrondi(TotEdt[1].TotCredit + C,CritEdt.Decimale) ;
Case CritEdt.SoldeProg of
  0 :  if Qr11T_SOLDEPROGRESSIF.AsString='X' then Progressif(False,0,0) ;
  1 :  Progressif(False,0,0) ;
 end ;
{ A Nouveau }
AnvDebit.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,CumulAu[0].TotDebit, CritEdt.AfficheSymbole) ;
AnvCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,CumulAu[0].TotCredit, CritEdt.AfficheSymbole) ;
AnvSolde.Caption:= PrintSolde(CumulAu[0].TotDebit,CumulAu[0].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
{ A Nouveau + Cumul }
CumDebit.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,D , CritEdt.AfficheSymbole) ;
CumCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,C ,CritEdt.AfficheSymbole) ;
CumSolde.Caption:= PrintSolde(D, C,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
Case CritEdt.SoldeProg of
  0 :  if Qr11T_SOLDEPROGRESSIF.AsString='X' then Progressif(True,D,C) ;
  1 :  Progressif(True,D,C) ;
 end ;
AddReport([1,2,3],CritEdt.GL.FormatPrint.Report,D,C,CritEdt.Decimale) ;
Case CritEdt.Rupture of
  rLibres   : AddGroupLibre(LRupt,Q,fbAux,CritEdt.LibreTrie,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
  rRuptures : AddRupt(Lrupt,Qr2E_AUXILIAIRE.AsString,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
  rCorresp  : BEGIN
              Case CritEDT.GL.PlansCorresp Of
                1 : If Qr11T_CORRESP1.AsString<>'' Then CptRupt:=Qr11T_CORRESP1.AsString+Qr2E_AUXILIAIRE.AsString
                                                   Else CptRupt:='.'+Qr2E_AUXILIAIRE.AsString ;
                2 : If Qr11T_CORRESP2.AsString<>'' Then CptRupt:=Qr11T_CORRESP2.AsString+Qr2E_AUXILIAIRE.AsString
                                                   Else CptRupt:='.'+Qr2E_AUXILIAIRE.AsString ;
               End ;
              AddRuptCorres(LRupt,CptRupt,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
              END ;
  End ;

end;

procedure TGLAuxGen.BSDMvtBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var CptRupt : String ;
begin
  inherited;
AncienEche:=QR2E_NUMECHE.AsInteger ;
PrintBand:=Affiche And (Not QR2E_GENERAL.IsNull) ;
if Not PrintBand then Exit ;
If IsECC then
   BEGIN
   LE_DEBIT.Caption:='0' ;
   LE_CREDIT.Caption:='0' ;
   END ;
TotAux[1].TotDebit:= Arrondi(TotAux[1].TotDebit+QR2DEBIT.AsFloat,CritEdt.Decimale) ;
TotAux[1].TotCredit:=Arrondi(TotAux[1].TotCredit+QR2CREDIT.AsFloat,CritEdt.Decimale) ;
TotGen[1].TotDebit:= Arrondi(TotGen[1].TotDebit + QR2DEBIT.AsFloat,CritEdt.Decimale) ;
TotGen[1].TotCredit:=Arrondi(TotGen[1].TotCredit + QR2CREDIT.AsFloat,CritEdt.Decimale) ;
TotEdt[1].TotDebit:=     Arrondi(TotEdt[1].TotDebit + Qr2DEBIT.AsFloat,CritEDT.Decimale) ;
TotEdt[1].TotCredit:=    Arrondi(TotEdt[1].TotCredit + Qr2Credit.AsFloat,CritEDT.Decimale) ;
TotDesMvts[1].TotDebit :=Arrondi(TotDesMvts[1].TotDebit + QR2DEBIT.AsFloat,CritEdt.Decimale) ;
TotDesMvts[1].TotCredit:=Arrondi(TotDesMvts[1].TotCredit+QR2CREDIT.AsFloat,CritEdt.Decimale) ;
E_PIECELIGECH.Caption:=QR2E_NUMEROPIECE.AsString+' / '+QR2E_NUMLIGNE.AsString+'   '+QR2E_NUMECHE.AsString ;
LE_VALIDE.Caption:=ValiQuali(QR2E_VALIDE.AsString,QR2E_QUALIFPIECE.AsString) ;
Case CritEdt.SoldeProg of             { Affectation Du Calcul du Solde Progressif O/N Sinon, d'aprés l'info sur le compte }
  0 : if Qr11T_SOLDEPROGRESSIF.AsString='X'
         then LE_SOLDE.Caption:=PrintSolde(ProgressDebit+QR2DEBIT.asFloat,ProgressCredit+QR2CREDIT.asFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole)
         else LE_SOLDE.Caption:=PrintSolde(QR2DEBIT.asFloat,QR2CREDIT.asFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
  1 : LE_SOLDE.Caption:=PrintSolde(ProgressDebit+QR2DEBIT.asFloat,ProgressCredit+QR2CREDIT.asFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
  2 : LE_SOLDE.Caption:=PrintSolde(QR2DEBIT.asFloat,QR2CREDIT.asFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
 end ;
AddGroup(LMulti, [QR2E_DATECOMPTABLE, QR2MULTIECHE],[QR2DEBIT.AsFloat,QR2CREDIT.AsFloat]) ;
AddReport([1,2,3],CritEdt.GL.FormatPrint.Report,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,CritEdt.Decimale) ;
Case CritEdt.Rupture of
  rLibres   : AddGroupLibre(LRupt,Q,fbAux,CritEdt.LibreTrie,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0]) ;
  rRuptures : AddRupt(LRupt,Qr2E_AUXILIAIRE.AsString,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0]) ;
  rCorresp  : BEGIN
              Case CritEDT.GL.PlansCorresp Of
                1 : If Qr11T_CORRESP1.AsString<>'' Then CptRupt:=Qr11T_CORRESP1.AsString+Qr2E_AUXILIAIRE.AsString
                                                   Else CptRupt:='.'+Qr2E_AUXILIAIRE.AsString ;
                2 : If Qr11T_CORRESP2.AsString<>'' Then CptRupt:=Qr11T_CORRESP2.AsString+Qr2E_AUXILIAIRE.AsString
                                                   Else CptRupt:='.'+Qr2E_AUXILIAIRE.AsString ;
                Else CptRupt:=Qr2E_AUXILIAIRE.AsString ;
               End ;
              AddRuptCorres(LRupt,CptRupt,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0])
              END ;
  End ;

Quoi:=QuoiMvt ;
end;

procedure TGLAuxGen.BSDMvtAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Case CritEdt.SoldeProg of
  0 :  if Qr11T_SOLDEPROGRESSIF.AsString='X' then Progressif(True,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat) ;
  1 :  Progressif(True,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat) ;
  end ;
end;

procedure TGLAuxGen.BSDMultiBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
If NumR=0 then
   BEGIN
     If OkZoomEdt Then PrintBand:=True Else
     Case CritEdt.TotMens of         { Total Mensuel O/N Sinon, d'aprés l'info sur le compte }
       0 : PrintBand:=(Qr11T_TOTAUXMENSUELS.AsString='X') ;
       1 : PrintBand:=True  ;
       2 : PrintBand:=False  ;
      end;
   END Else
If NumR=1 then                       { Affiche les échéances }
   BEGIN
   if AncienEche<=1 then PrintBand:=False else
      begin If OkZoomEdt Then PrintBand:=True Else PrintBand:=CritEDT.GL.TotEche ; end ;
   END ;
end;

procedure TGLAuxGen.BFCpteAuxBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=Affiche ; If Not PrintBand then Exit ;//Rony 03/11/97, pour cpts associés aux Ruptures
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
TT_AUXILIAIRE_.Caption:=MsgLibel.Mess[14]+'   '+Qr11T_AUXILIAIRE.AsString+'   '+Qr11T_LIBELLE.AsString ;
{ Total et Solde par Auxiliaire }
TotAuxDebit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotAux[1].TotDebit, CritEdt.AfficheSymbole ) ;
TotAuxCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotAux[1].TotCredit, CritEdt.AfficheSymbole ) ;
TotAuxSolde.Caption:=PrintSolde(TotAux[1].TotDebit, TotAux[1].TotCredit ,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
TotDesMvtsDebit.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotDesMvts[1].TotDebit, CritEdt.AfficheSymbole) ;
TotDesMvtsCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotDesMvts[1].TotCredit, CritEdt.AfficheSymbole) ;
TotDesMvtsSolde.Caption:=PrintSolde(TotDesMvts[1].TotDebit,TotDesMvts[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
Quoi:=QuoiAux(2) ;
end;

procedure TGLAuxGen.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TotalDebit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotEdt[1].TotDebit, CritEdt.AfficheSymbole ) ;
TotalCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotEdt[1].TotCredit, CritEdt.AfficheSymbole ) ;
TotalSoldes.Caption:=PrintSolde(TotEdt[1].TotDebit,TotEdt[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TGLAuxGen.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
Titre1Rep.Caption:=Titre2Rep.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
Report1Solde.Caption:=Report2Solde.Caption ;
end;

procedure TGLAuxGen.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var D,C : Double ;
begin
  inherited;
Case QuelReport(CritEdt.GL.FormatPrint.Report,D,C) Of
  1 : BEGIN TITREREPORTB.Caption:=MsgLibel.Mess[6] ; Titre2Rep.Caption:='' ; END ;
  2 : BEGIN TITREREPORTB.Caption:=MsgLibel.Mess[7] ; Titre2Rep.Caption:=StReportAux ; END ;
  3 : BEGIN TITREREPORTB.Caption:=MsgLibel.Mess[7] ; Titre2Rep.Caption:=StReportGen ; END ;
  END ;
Report2Debit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, D, CritEdt.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, C, CritEdt.AfficheSymbole ) ;
Report2Solde.Caption:=PrintSolde(D,C,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
end;

procedure TGLAuxGen.QEcrAfterOpen(DataSet: TDataSet);
begin
  inherited;
QR2E_GENERAL       :=TStringField(QEcr.FindField('E_GENERAL')) ;
QR2E_AUXILIAIRE    :=TStringField(QEcr.FindField('E_AUXILIAIRE')) ;
QR2E_EXERCICE      :=TStringField(QEcr.FindField('E_EXERCICE')) ;
QR2E_DATECOMPTABLE :=TDateTimeField(QEcr.FindField('E_DATECOMPTABLE')) ;
QR2E_DATECOMPTABLE.Tag:=1 ;
QR2E_NUMEROPIECE   :=TIntegerField(QEcr.FindField('E_NUMEROPIECE')) ;
QR2E_NUMLIGNE      :=TIntegerField(QEcr.FindField('E_NUMLIGNE')) ;
QR2E_NUMECHE       :=TIntegerField(QEcr.FindField('E_NUMECHE')) ;
QR2E_REFINTERNE    :=TStringField(QEcr.FindField('E_REFINTERNE')) ;
QR2E_ETABLISSEMENT :=TStringField(QEcr.FindField('E_ETABLISSEMENT')) ;
QR2E_LIBELLE       :=TStringField(QEcr.FindField('E_LIBELLE')) ;
QR2E_DEVISE        :=TStringField(QEcr.FindField('E_DEVISE')) ;
QR2E_VALIDE        :=TStringField(QEcr.FindField('E_VALIDE')) ;
QR2E_QTE1          :=TFloatField(QEcr.FindField('E_QTE1')) ;
QR2E_QTE2          :=TFloatField(QEcr.FindField('E_QTE2')) ;
QR2E_QUALIFQTE1    :=TStringField(QEcr.FindField('E_QUALIFQTE1')) ;
QR2E_QUALIFQTE2    :=TStringField(QEcr.FindField('E_QUALIFQTE2')) ;
QR2E_LETTRAGE      :=TStringField(QEcr.FindField('E_LETTRAGE')) ;
QR2E_JOURNAL       :=TStringField(QEcr.FindField('E_JOURNAL')) ;
QR2E_QUALIFPIECE   :=TStringField(QEcr.FindField('E_QUALIFPIECE')) ;
QR2DEBIT           :=TFloatField(QEcr.FindField('DEBIT')) ;
QR2CREDIT          :=TFloatField(QEcr.FindField('CREDIT')) ;
QR2MULTIECHE       :=TIntegerField(QEcr.FindField('MULTIECHE')) ;
IsECC:=(FDevises.Value<>V_PGI.DevisePivot)and(FMontant.ITemIndex=1)and(QEcr.FindField('E_NATUREPIECE').AsString='ECC') ;
ChgMaskChamp(Qr2DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
ChgMaskChamp(Qr2CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
end;

procedure TGLAuxGen.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr11T_AUXILIAIRE      :=TStringField(Q.FindField('T_AUXILIAIRE'));
Qr11T_LIBELLE         :=TStringField(Q.FindField('T_LIBELLE'));
Qr11T_SOLDEPROGRESSIF :=TStringField(Q.FindField('T_SOLDEPROGRESSIF'));
Qr11T_SAUTPAGE        :=TStringField(Q.FindField('T_SAUTPAGE'));
Qr11T_TOTAUXMENSUELS  :=TStringField(Q.FindField('T_TOTAUXMENSUELS'));
Qr11T_TOTALDEBIT      :=TFloatField(Q.FindField('T_TOTALDEBIT'));
Qr11T_TOTALCREDIT     :=TFloatField(Q.FindField('T_TOTALCREDIT'));
If CritEDT.Rupture=rCorresp then
   BEGIN
   Qr11T_CORRESP1         :=TStringField(Q.FindField('T_CORRESP1'));
   Qr11T_CORRESP2         :=TStringField(Q.FindField('T_CORRESP2'));
   END ;
end;

procedure TGLAuxGen.QGeneAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr12G_GENERAL:=TStringField(QGene.FindField('G_GENERAL'));
Qr12G_LIBELLE:=TStringField(QGene.FindField('G_LIBELLE'));
end;

procedure TGLAuxGen.QEcrBeforeOpen(DataSet: TDataSet);
begin
  inherited;
QEcr.PArams[0].AsString:=Q.FindField('T_AUXILIAIRE').AsString ;
end;

procedure TGLAuxGen.BGenFBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
PrintBand:=Affiche ; If Not PrintBand then Exit ; // Rony 03/11/97, pour Cpts associés aux Ruptures
If Not OkBand1 Then BEGIN PrintBand:=FALSE ; Exit ; END ;
TG_GENERAL_.Caption:=MsgLibel.Mess[13]+'   '+Qr12G_GENERAL.AsString+'   '+Qr12G_LIBELLE.AsString ;
Quoi:=QuoiGen(2) ;
{ Total et Solde par Auxiliaire }
TOTSDEBIT.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGen[1].TotDebit, CritEDT.AfficheSymbole ) ;
TOTSCREDIT.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGen[1].TotCredit, CritEDT.AfficheSymbole ) ;
TOTSSOLDE.Caption:=PrintSolde(TotGen[1].TotDebit, TotGen[1].TotCredit ,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
end;

procedure TGLAuxGen.FNatureCptChange(Sender: TObject);
begin
  inherited;
If Not QRLoading Then BEGIN FGen1.Clear ; FGen2.Clear ; FGenJoker.Clear ; END ;
Case FNatureCpt.ItemIndex of
  0     : BEGIN FGen1.ZoomTable:=tzGCollectif ; FGen2.ZoomTable:=tzGCollectif ; END ;
  1,2,4 : BEGIN FGen1.ZoomTable:=tzGCollDivers ; FGen2.ZoomTable:=tzGCollDivers ; END ;
  3     : BEGIN FGen1.ZoomTable:=tzGCollClient ; FGen2.ZoomTable:=tzGCollClient ; END ;
  5     : BEGIN FGen1.ZoomTable:=tzGCollFourn ; FGen2.ZoomTable:=tzGCollFourn ; END ;
  6     : BEGIN FGen1.ZoomTable:=tzGCollSalarie ; FGen2.ZoomTable:=tzGCollSalarie ; END ;
  End ;
end;

procedure TGLAuxGen.BGenFAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([3],CritEdt.GL.FormatPrint.Report) ;
end;

procedure TGLAuxGen.BFCpteAuxAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([2],CritEdt.GL.FormatPrint.Report) ;
end;

procedure TGLAuxGen.FGen1Change(Sender: TObject);
Var AvecJokerS : Boolean ;
begin
  inherited;
AvecJokerS:=Joker(FGen1, FGen2, FGenJoker) ;
TFaA.Visible:=Not AvecJokerS ;
TFGen.Visible:=Not AvecJokerS ;
TFGenJoker.Visible:=AvecJokerS ;
end;


procedure TGLAuxGen.BNouvRechClick(Sender: TObject);
begin
  inherited;
If FGenJoker.Visible then FGenJoker.Text:='' ;

end;


procedure TGLAuxGen.DLRuptNeedData(var MoreData: Boolean);
Var TotRupt : Array[0..12] of Double ;
    SumD, SumC  : Double ;
    Quellerupt  : Integer ;
    CodRupt, LibRupt, Lib1, CptRupt, Stcode     : String ;
    Col         : TColor ;
    DansGen, OkOk        : Boolean ;
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
  rRuptures : MoreData:=PrintRupt(LRupt,Qr11T_AUXILIAIRE.AsString,CodRupt,LibRupt,DansGen,QRP.EnRupture,TotRupt) ;
  rCorresp  : BEGIN
              OkOk:=TRUE ;
              Case CritEDT.GL.PlansCorresp  Of
                1 : If Qr11T_CORRESP1.AsString<>'' Then CptRupt:=Qr11T_CORRESP1.AsString+Qr2E_AUXILIAIRE.AsString
                                                   Else CptRupt:='.'+Qr2E_AUXILIAIRE.AsString ;
                2 : If Qr11T_CORRESP2.AsString<>'' Then CptRupt:=Qr11T_CORRESP2.AsString+Qr2E_AUXILIAIRE.AsString
                                                   Else CptRupt:='.'+Qr2E_AUXILIAIRE.AsString ;
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
                 insert(MsgLibel.Mess[15]+' ',CodRupt,Quellerupt+2) ;
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
OkEnteteRupt:=MoreData and (CritEdt.Rupture=rCorresp);
end;

procedure TGLAuxGen.BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.Rupture<>rRien) ;
end;

procedure TGLAuxGen.FSansRuptClick(Sender: TObject);
begin
  inherited;
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
FRupturesClick(Nil) ;
end;

procedure TGLAuxGen.FRupturesClick(Sender: TObject);
begin
  inherited;
if FPlansCo.Checked then CorrespToCombo(FPlanRuptures,FbAux) ;
if FRuptures.Checked then FGroupRuptures.Caption:=' '+MsgLibel.Mess[17] ;
end;


end.
