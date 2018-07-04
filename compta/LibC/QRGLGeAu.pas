unit QRGLGeAu;

interface
           
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, UtilEDT, CritEDT, QRRupt, CpteUtil, EdtLegal,
  Ent1, HEnt1, HQry, HSysMenu, Menus, Calcole, HTB97, HPanel, UiUtil, tCalcCum ;

 procedure GLGenAuxi ;
 procedure GLGenAuxiZoom(Crit : TCritEdt) ;

type
  TGLGenAux = class(TFQR)
    TFAux: THLabel;
    TFaA: TLabel;
    FAux2: THCpteEdit;
    FAuxJoker: TEdit;
    TFAuxJoker: THLabel;
    FValide: TCheckBox;
    HLabel9: THLabel;
    HLabel11: THLabel;
    FRefInterne: TEdit;
    FNumPiece1: TMaskEdit;
    Label12: TLabel;
    FNumPiece2: TMaskEdit;
    HLabel3: THLabel;
    FExcepAux: TEdit;
    FSautPage: TCheckBox;
    FLigneTiersentete: TCheckBox;
    FLigneTiersPied: TCheckBox;
    FLigneGenPied: TCheckBox;
    RAux: TQRLabel;
    RAux1: TQRLabel;
    TRaT: TQRLabel;
    RAux2: TQRLabel;
    QRLabel36: TQRLabel;
    RExcepAux: TQRLabel;
    QRLabel13: TQRLabel;
    RNumPiece1: TQRLabel;
    QRLabel14: TQRLabel;
    RNumPiece2: TQRLabel;
    QRLabel10: TQRLabel;
    RRefInterne: TQRLabel;
    QRLabel17: TQRLabel;
    RValide: TQRLabel;
    TRLegende: TQRLabel;
    RLegende: TQRLabel;
    TValide: TQRLabel;
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
    TG_GENERAL: TQRLabel;
    BSDHCptAux: TQRBand;
    TCumul: TQRLabel;
    CumCredit: TQRLabel;
    CumSolde: TQRLabel;
    AnvSolde: TQRLabel;
    AnvCredit: TQRLabel;
    CumDebit: TQRLabel;
    AnvDebit: TQRLabel;
    QRLabel25: TQRLabel;
    BSDMvt: TQRBand;
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
    TLibMulti: TQRLabel;
    TDebMulti: TQRLabel;
    TCreMulti: TQRLabel;
    TSolMulti: TQRLabel;
    BFCpteGen: TQRBand;
    TG_GENERAL_: TQRLabel;
    TotGenSolde: TQRLabel;
    TotGenCredit: TQRLabel;
    TotGenDebit: TQRLabel;
    QRLabel33: TQRLabel;
    TotalDebit: TQRLabel;
    TotalCredit: TQRLabel;
    TotalSoldes: TQRLabel;
    QROVERLAY: TQRBand;
    Trait1: TQRLigne;
    Trait2: TQRLigne;
    Trait3: TQRLigne;
    Trait4: TQRLigne;
    Trait0: TQRLigne;
    Trait5: TQRLigne;
    Ligne1: TQRLigne;
    TITRE2REP: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    REPORT2SOLDE: TQRLabel;
    QEcr: TQuery;
    SEcr: TDataSource;
    MsgLibel: THMsgBox;
    QRDLAuxH: TQRDetailLink;
    QRDLMulti: TQRDetailLink;
    QRDLMvt: TQRDetailLink;
    QAux: TQuery;
    SAux: TDataSource;
    BAuxF: TQRBand;
    TT_AUXILIAIRE_: TQRLabel;
    TOTSDEBIT: TQRLabel;
    TOTSCREDIT: TQRLabel;
    TOTSSOLDE: TQRLabel;
    QRLabel21: TQRLabel;
    TotDesMvtsDebit: TQRLabel;
    TotDesMvtsCredit: TQRLabel;
    TotDesMvtsSolde: TQRLabel;
    FSoldeProg: TCheckBox;
    FTotMens: TCheckBox;
    FTotEche: TCheckBox;
    TT_AUXILIAIRE: TQRLabel;
    FAux1: THCpteEdit;
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
    procedure QRDLMultiNeedData(var MoreData: Boolean);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BSDHCptAuxBeforePrint(var PrintBand: Boolean;      var Quoi: string);
    procedure BSDMvtBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BSDMvtAfterPrint(BandPrinted: Boolean);
    procedure BSDMultiBeforePrint(var PrintBand: Boolean;        var Quoi: string);
    procedure BFCpteGenBeforePrint(var PrintBand: Boolean;       var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean;        var Quoi: string);
    procedure FAuxJokerChange(Sender: TObject);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QEcrAfterOpen(DataSet: TDataSet);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure QAuxAfterOpen(DataSet: TDataSet);
    procedure BAuxFBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QEcrBeforeOpen(DataSet: TDataSet);
    procedure FNatureCptChange(Sender: TObject);
    procedure BFCpteGenAfterPrint(BandPrinted: Boolean);
    procedure BAuxFAfterPrint(BandPrinted: Boolean);
    procedure BNouvRechClick(Sender: TObject);
    procedure DLRuptNeedData(var MoreData: Boolean);
    procedure BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FRupturesClick(Sender: TObject);
    procedure FSansRuptClick(Sender: TObject);
    procedure GenereSQL ; Override ;
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
  private    { Déclarations privées }
    Qr11G_GENERAL, Qr11G_LIBELLE,
    Qr11G_COLLECTIF, Qr11G_SOLDEPROGRESSIF,
    Qr11G_SAUTPAGE, Qr11G_TOTAUXMENSUELS,
    Qr11G_CORRESP1, Qr11G_CORRESP2              : TStringField ;
    Qr12T_AUXILIAIRE,Qr12T_LIBELLE              : TStringField;
    StReportGen, StReportAux                    : string ;
    NumR, AncienEche                            : Integer ;
    UneSeuleEche, OkBand1, OkBand2,
    forceaffichepiedcpt, IsECC, Affiche         : Boolean ;
    TotGen, TotEdt, TotAux, TotDesMvts          : TabTot ;
    LMulti, LRupt                               : TStringList ;
    QR2E_GENERAL, QR2E_EXERCICE,
    QR2E_REFINTERNE, QR2E_ETABLISSEMENT,
    QR2E_LIBELLE, QR2E_DEVISE, QR2E_AUXILIAIRE,
    QR2E_VALIDE, QR2E_QUALIFQTE1,
    QR2E_QUALIFQTE2, QR2E_LETTRAGE,
    QR2E_JOURNAL, QR2E_QUALIFPIECE              : TStringField ;
    QR2E_DATECOMPTABLE                          : TDateTimeField ;
    QR2E_NUMEROPIECE, QR2E_NUMLIGNE,
    QR2E_NUMECHE, QR2MULTIECHE                  : TIntegerField ;
    QR2E_QTE1 ,QR2E_QTE2,
    QR2DEBIT, QR2CREDIT                         : TFloatField ;
    Procedure GenereSQLSub2 ;
    procedure GenereSQLSub ;
    Function  QuoiGen(i : Integer) : String ;
    Function  QuoiAux(i : Integer) : String ;
    Function  QuoiMvt : String ;
    Procedure GLGeAuZoom(Quoi : String) ;
  public     { Déclarations publiques }
  end;

implementation

{$R *.DFM}

procedure GLGenAuxi ;
var QR : TGLGenAux ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TGLGenAux.Create(Application) ;
Edition.Etat:=etGlGenAux ;
QR.QRP.QRPrinter.OnSynZoom:=QR.GLGeAuZoom ;
QR.InitType(nbGen,neGl,msGenEcr,'QRGLGEAU','',TRUE,FALSE,TRUE,Edition) ;
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

procedure GLGenAuxiZoom(Crit : TCritEdt) ;
var QR : TGLGenAux ;
    Edition : TEdition ;
BEGIN
QR:=TGLGenAux.Create(Application) ;
Edition.Etat:=etGlGenAux ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.GLGeAuZoom ;
 QR.CritEdt:=Crit ;
 QR.InitType(nbGen,neGl,msGenEcr,'QRGLGEAU','',FALSE,TRUE,TRUE,Edition) ;
 QR.ShowModal ;
 finally
 QR.Free ;
 end ;
END ;

Function TGLGenAux.QuoiGen(i : Integer) : String ;
BEGIN
Case i Of
  1 : Result:=Qr11G_General.AsString+'#'+Qr11G_LIBELLE.AsString+'@'+'1' ;
  2 : Result:=Qr11G_General.AsString+'#'+Qr11G_LIBELLE.AsString+' '+
              TotGenSolde.Caption+'@'+'1' ;
  END ;
END ;

Function TGLGenAux.QuoiAux(i : Integer) : String ;
BEGIN
Case i Of
  1 : Result:=QR12T_AUXILIAIRE.AsString+' '+Qr11G_GENERAL.AsString+'#'+QR12T_LIBELLE.AsString+'@'+'2' ;
  2 : Result:=QR12T_AUXILIAIRE.AsString+' '+Qr11G_GENERAL.AsString+'#'+QR12T_LIBELLE.AsString+' '+
              TotSSolde.Caption+'@'+'2' ;
  END ;
END ;

Function TGLGenAux.QuoiMvt : String ;
BEGIN
Result:=QR2E_AUXILIAIRE.AsString+' '+QR12T_LIBELLE.AsString+' '+Le_Solde.Caption+
        '#'+QR2E_JOURNAL.AsString+' N° '+IntToStr(QR2E_NUMEROPIECE.AsInteger)+' '+DateToStr(QR2E_DateComptable.AsDAteTime)+'-'+
         PrintSolde(QR2Debit.AsFloat,QR2Credit.AsFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole)+
         '@'+'5;'+QR2E_JOURNAL.AsString+';'+UsDateTime(QR2E_DATECOMPTABLE.AsDateTime)+';'+QR2E_NUMEROPIECE.AsString+';'+QR2E_EXERCICE.asString+';'+
         IntToStr(QR2E_NumLigne.AsInteger)+';' ;
END ;

Procedure TGLGenAux.GLGeAuZoom(Quoi : String) ;
Var Lp,i: Integer ;
BEGIN
Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
i:=StrToInt(Copy(Quoi,Lp+1,1)) ;
If (i=5) Then
   BEGIN
   Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
   If QRP.QRPrinter.FSynShiftDblClick Then i:=6 ;
   If QRP.QRPrinter.FSynCtrlDblClick Then i:=11 ;
   END ;
ZoomEdt(i,Quoi) ;
END ;

procedure TGLGenAux.GenereSQL ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
AlimSQLMul(Q,0) ; Q.Close ; ChangeSQL(Q) ;
//Q.Prepare ;
PrepareSQLODBC(Q) ;
Q.Open ;
GenereSQLSub ; GenereSQLSub2 ;
END ;

procedure TGLGenAux.GenereSQLSub ;
BEGIN
Inherited ;
AlimSQLMul(QAux,1) ; QAux.Close ; ChangeSQL(QAux) ;
//QAux.Prepare ;
PrepareSQLODBC(QAux) ;
QAux.Open ;
END ;

Procedure TGLGenAux.GenereSQLSub2 ;
BEGIN
QEcr.Close ;
QEcr.SQL.Clear ;
QEcr.SQL.Add('Select E_AUXILIAIRE,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE, ') ;
QEcr.SQL.Add(      ' E_REFINTERNE,E_ETABLISSEMENT,E_LIBELLE,E_NUMECHE,E_DEVISE,E_VALIDE, ') ;
QEcr.SQL.Add(      ' E_QTE1,E_QTE2,E_QUALIFQTE1,E_QUALIFQTE2,E_LETTRAGE,E_GENERAL, ') ;
// GP : bug edition avec filtre sur N° pièce trop grand (à cause du *10)
//QEcr.SQL.Add(      'E_JOURNAL, E_NUMEROPIECE*10+E_NUMLIGNE as MULTIECHE, E_NATUREPIECE, E_QUALIFPIECE, ') ;
QEcr.SQL.Add(      'E_JOURNAL, E_NUMLIGNE as MULTIECHE, E_NATUREPIECE, E_QUALIFPIECE, ') ;

Case Fmontant.ItemIndex of
  0 : BEGIN QEcr.SQL.Add(' E_DEBIT DEBIT,E_CREDIT CREDIT ') ; END ;
  1 : BEGIN QEcr.SQL.Add(' E_DEBITDEV DEBIT,E_CREDITDEV CREDIT ') ; END ;
end ;
{ Tables explorées par la SQL }
QEcr.SQL.Add(' From ECRITURE ') ;
{ Construction de la clause Where de la SQL }
QEcr.SQL.Add(' Where E_GENERAL=:G_GENERAL and E_AUXILIAIRE=:T_AUXILIAIRE ') ;
QEcr.SQL.Add('And E_DATECOMPTABLE>="'+usdatetime(CritEdt.GL.Date21)+'" And E_DATECOMPTABLE<="'+usdatetime(CritEdt.Date2)+'"') ;
if FExercice.ItemIndex<>0 then QEcr.SQL.Add(' And E_EXERCICE="'+CritEdt.Exo.Code+'" ') ;
if FNumPiece1.Text<>'' then QEcr.SQL.Add(   ' And E_NUMEROPIECE>='+IntToStr(CritEdt.GL.NumPiece1)) ;
if FNumPiece2.Text<>'' then QEcr.SQL.Add(   ' And E_NUMEROPIECE<='+IntToStr(CritEdt.GL.NumPiece2)) ;
QEcr.SQL.Add(TraduitNatureEcr(CritEdt.Qualifpiece, 'E_QUALIFPIECE', true, CritEdt.ModeRevision)) ;
if FRefInterne.Text<>'' then QEcr.SQL.Add('And UPPER(E_REFINTERNE) like "'+TraduitJoker(CritEdt.RefInterne)+'" ' );
if FEtab.ItemIndex<>0 then QEcr.SQL.Add(' And E_ETABLISSEMENT="'+CritEdt.Etab+'"') ;
if FValide.State<>cbGrayed then QEcr.SQL.Add(' And E_VALIDE="'+CritEdt.Valide+'" ') ;
if FDevises.ItemIndex<>0 then QEcr.SQL.Add(' And E_DEVISE="'+CritEdt.DeviseSelect+'"') ;
QEcr.SQL.Add(' And E_ECRANOUVEAU="N" and E_QUALIFPIECE<>"C" ') ;
If CritEDT.SQLPlus<>'' Then QEcr.SQL.Add(CritEDT.SQLPlus) ;
{ Construction de la clause Order By de la SQL }
QEcr.SQL.Add(' Order By E_GENERAL,E_AUXILIAIRE,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE ') ;
ChangeSQL(QEcr) ; QEcr.Open ;
END ;

procedure TGLGenAux.RenseigneCritere ;
{ Récupération des champs du multicritère dans l'entête d'état }
BEGIN
Inherited ;
If EstSerie(S5) Then RLegende.Caption:=MsgRien.Mess[7] ;
PositionneFourchetteSt(Faux1,FAux2,CritEdt.LSCpt1,CritEdt.LSCpt2) ;
if CritEdt.SJoker then
   BEGIN
   RAux.Caption:=MsgLibel.Mess[9] ;
   RAux1.Caption:=FAuxJoker.Text
   END Else
   BEGIN
   RAux.Caption:=MsgLibel.Mess[8] ;
   RAux1.Caption:=CritEdt.LSCpt1 ; RAux2.Caption:=CritEdt.LSCpt2 ;
   END ;
RAux2.Visible:=Not CritEdt.SJoker ; TRaT.Visible:=Not CritEdt.SJoker ;
RRefInterne.Caption:=FRefInterne.Text ;
RNumPiece1.Caption:=IntToStr(CritEdt.GL.NumPiece1)   ; RNumPiece2.Caption:=IntToStr(CritEdt.GL.NumPiece2) ;
RExcepAux.Caption:=FExcepAux.Text ;
CaseACocher(FValide,RValide)            ;
(*
CaseACocher(FSoldeProg,RSoldeProg) ;
CaseACocher(FSautPage,RSautPage)        ; CaseACocher(FTotMens,RTotMens) ;
CaseACocher(FTotEche,RTotEche)          ; CaseACocher(FLigneTiersEntete ,RLigneTiersEntete) ;
CaseACocher(FLigneGenPied ,RLigneGenPied) ;
CaseACocher(FLigneTiersPied,RLigneTiersPied) ;
*)
DateCumulAuGL(TCumul,CritEdt,MsgLibel.Mess[5]) ;
END ;

procedure TGLGenAux.ChoixEdition ;
{ Initialisation des options d'édition }
BEGIN
Inherited ;
ChargeGroup(LMulti,['MOIS','ECHE']) ;
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

procedure TGLGenAux.FinirPrint;
BEGIN
Inherited ;
VideGroup(LMulti) ; QEcr.Close ;
If GCalc<>NIL Then BEGIN GCalc.Free ; GCalc:=NIL ; END ;
if CritEdt.Rupture<>rRien then VideRupt(LRupt) ;
END ;

procedure TGLGenAux.InitDivers ;
BEGIN
Inherited ;
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
TValide.Caption:='' ;
{ Gestion multiéchéances }
AncienEche:=1 ;
{ Labels sur les bandes }
BSDHCptAux.Frame.DrawTop:=CritEdt.GL.FormatPrint.PrSepCompte[2] ;
BFCpteGen.Frame.DrawTop:=CritEdt.GL.FormatPrint.PrSepCompte[4] ;
BFCpteGen.Frame.DrawBottom:=CritEdt.GL.FormatPrint.PrSepCompte[4] ;
OkBand1:=TRUE ; OkBand2:=TRUE ; forceaffichepiedcpt:=FALSE ;
END ;

procedure TGLGenAux.RecupCritEdt ;
var st        : string ;
    NonLibres : Boolean ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  SJoker:=FAuxJoker.Visible ; Composite:=True ;
  if SJoker Then BEGIN SCpt1:=FAuxJoker.Text ; SCpt2:=FAuxJoker.Text ; END
            Else BEGIN
                 SCpt1:=FAux1.Text ; SCpt2:=FAux2.Text ;
                 PositionneFourchetteSt(FAux1,FAux2,CritEdt.LSCpt1,CritEdt.LSCpt2) ;
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
  GL.SSauf:=Trim(FExcepAux.Text) ; St:=GL.SSauf ;
  if GL.SSauf<>'' then St:=' And '+AnalyseCompte(FExcepAux.Text,fbAux,True,False) ;
  Gl.SSqlSauf:=St ;
  GL.RuptOnly:=QuelleTypeRupt(1,FSAnsRupt.Checked,FAvecRupt.Checked,FALSE) ;
  NonLibres:=((Rupture=rRuptures) or (Rupture=rCorresp)) ;
  if NonLibres then GL.PlanRupt:=FPlanRuptures.Value ;
  If (CritEdt.Rupture=rCorresp) Then GL.PlansCorresp:=FPlanRuptures.ItemIndex+1 ;
  GL.OnlyCptAssocie:=(Rupture<>rRien) and FOnlyCptAssocie.Checked ;
  With GL.FormatPrint Do
    BEGIN
    PrSepCompte[2]:=FLigneTiersEntete.Checked ;
    PrSepCompte[3]:=FLigneTiersPied.Checked ;
    PrSepCompte[4]:=FLigneGenPied.Checked ;
    END ;
  END ;
END ;

function  TGLGenAux.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ;
If Result Then
   BEGIN
(*
   QBal:=PrepareTotCptM(fbGene, fbAux, [], Dev, Etab, Exo) ;
*)
   Gcalc:=TGCalculCum.create(Deux,fbGene,fbAux,QuelTypeEcr,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE) ;
   GCalc.initCalcul('','','','',CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                    CritEdt.Date1,CritEdt.GL.Date11,TRUE) ;
   END ;
END ;

procedure TGLGenAux.FormShow(Sender: TObject);
begin
HelpContext:=7427000 ;
//Standards.HelpContext:=7427010 ;
//Avances.HelpContext:=7427020 ;
//Mise.HelpContext:=7427030 ;
//Option.HelpContext:=7427040 ;
//TabRuptures.HelpContext:=7427050 ;
  inherited;
TabSup.TabVisible:=False;
FCpte1.ZoomTable:=tzGCollectif ; FCpte2.ZoomTable:=tzGCollectif ;
FValide.Checked:=True ; FValide.State:=cbGrayed ;
FAux1.Text:='' ; FAux2.Text:='' ;
FRefInterne.Text:='' ;
If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
FSurRupt.Visible:=False ;
FCodeRupt1.Visible:=False  ; FCodeRupt2.Visible:=False ;
TFCodeRupt1.Visible:=False ; TFCodeRupt2.Visible:=False ;
FOnlyCptAssocie.Enabled:=False ;
end;

procedure TGLGenAux.QRDLMultiNeedData(var MoreData: Boolean);
var LibMulti, CodMulti : string ;
    TotMulti           : Array[0..12] of Double ;
begin
  inherited;
MoreData:=PrintGroup(LMulti, QEcr, [QR2E_DATECOMPTABLE, QR2MULTIECHE], CodMulti, LibMulti, TotMulti, NumR) ;
Case NumR of  { Mensuel }
  0 : BEGIN
      BSDMulti.Frame.DrawTop:=False ;
      BSDMulti.Font.Color:=clRed ;
      TLibmulti.Caption:=MsgLibel.Mess[3]+' '+CodMulti;
      TDebMulti.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotMulti[0], CritEdt.AfficheSymbole) ;
      TCreMulti.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotMulti[1], CritEdt.AfficheSymbole) ;
      TSolMulti.Caption:=PrintSolde(TotMulti[0],TotMulti[1],CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
      END ;
  1 : BEGIN { Eché. }
      BSDMulti.Frame.DrawTop:=False ;
      BSDMulti.Font.Color:=clPurple ;
      TLibmulti.Caption:=MsgLibel.Mess[4]+' '+IntToStr(QR2E_NUMEROPIECE.AsInteger)+' / '+QR2E_NUMLIGNE.AsString ;
      TDebMulti.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotMulti[0], CritEdt.AfficheSymbole) ;
      TCreMulti.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotMulti[1], CritEdt.AfficheSymbole) ;
      TSolMulti.Caption:='' ;
      END ;
 end;
end;

procedure TGLGenAux.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
Fillchar(TotDesMvts,SizeOf(TotDesMvts),#0) ;
TG_GENERAL.Caption:=MsgLibel.Mess[10]+'   '+Qr11G_GENERAL.AsString+'   '+Qr11G_LIBELLE.AsString ;
StReportGen:=Qr11G_GENERAL.AsString ;
InitReport([2],CritEdt.GL.FormatPrint.Report) ;
Fillchar(TotGen,SizeOf(TotGen),#0) ;
Case CritEdt.Rupture of
  rLibres    : if CritEdt.GL.OnlyCptAssocie then PrintBand:=DansRuptLibre(Q,fbGene,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
  rRuptures  : if CritEdt.GL.OnlyCptAssocie then PrintBand:=DansRupt(LRupt,Qr11G_GENERAL.AsString) ;
  rCorresp   : if CritEdt.GL.OnlyCptAssocie then
                 if CritEDT.GL.PlansCorresp=1 then PrintBand:=(Qr11G_CORRESP1.AsString<>'') Else
                 if CritEDT.GL.PlansCorresp=2 then PrintBand:=(Qr11G_CORRESP2.AsString<>'') ;
  End;
Affiche:=PrintBand ;
If PrintBand then
   BEGIN
   OkBand1:=TRUE ; OkBand2:=TRUE ; forceaffichepiedcpt:=FALSE ;
   If (CritEdt.GL.TypCpt=1) Then
      BEGIN
      If Qr12T_AUXILIAIRE.IsNull Then BEGIN OkBand1:=FALSE ; OkBand2:=FALSE ; forceaffichepiedcpt:=TRUE ; END ;
      END Else
      BEGIN
      If Qr12T_AUXILIAIRE.IsNull Then BEGIN PrintBand:=FALSE ; OkBand1:=FALSE ; OkBand2:=FALSE ; Exit ; END ;
      END ;
   Case CritEdt.SautPage of
     0 : BDetail.forceNewPage:=(Qr11G_SAUTPAGE.AsString='X') ;
     1 : BDetail.forceNewPage:=True ;
     2 : BDetail.forceNewPage:=False ;
    end ;
   Quoi:=QuoiGen(1) ;
   END ;
end;

procedure TGLGenAux.BSDHCptAuxBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var D,C : Double ;
    CumulAu : TabTot ;
    CptRupt : String ;
begin
  inherited;
Fillchar(TotAux,SizeOf(TotAux),#0) ;
InitReport([3],CritEdt.GL.FormatPrint.Report) ;
TT_AUXILIAIRE.Caption:=MsgLibel.Mess[11]+'   '+Qr12T_AUXILIAIRE.AsString+'   '+QR12T_LIBELLE.AsString ;
StReportAux:=Qr11G_GENERAL.AsString+' / '+Qr12T_AUXILIAIRE.AsString ;
Fillchar(CumulAu,SizeOf(CumulAu),#0) ;
PrintBand:=Affiche ; If Not PrintBand then Exit ; // Rony 03/11/97, Pour cpt Associés aux Ruptures
If Not OkBand1 Then BEGIN PrintBand:=FALSE ; Exit ; END ;
(*
ExecuteTotCptM(QBal,QG_GENERAL.AsString,QR2E_AUXILIAIRE.AsString, CritEdt.Date1,CritEdt.GL.Date11,
               CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,TotAnCRupt,FALSE) ;
*)

Quoi:=QuoiAux(1) ;
GCAlc.ReInitCalcul(Qr11G_GENERAL.AsString, Qr12T_AUXILIAIRE.AsString,0,0) ;
GCalc.Calcul ; CumulAu:=GCalc.ExecCalc.TotCpt ;
CumulVersSolde(CumulAu[0]) ;
If CritEDT.Date1=CritEDT.GL.Date11 Then Fillchar(CumulAu[1],SizeOf(CumulAu[1]),#0) ;
D:=CumulAu[0].TotDebit+CumulAu[1].TotDebit ;
C:=CumulAu[0].TotCredit+CumulAu[1].TotCredit ;
// Rony 07/04/1997  ---> TotAux
TotAux[1].TotDebit:= Arrondi(TotAux[1].TotDebit + D,CritEdt.Decimale) ;
TotAux[1].TotCredit:= Arrondi(TotAux[1].TotCredit + C,CritEdt.Decimale) ;
TotGen[1].TotDebit:= Arrondi(TotGen[1].TotDebit + D,CritEdt.Decimale) ;
TotGen[1].TotCredit:=Arrondi(TotGen[1].TotCredit + C,CritEdt.Decimale) ;
TotEdt[1].TotDebit  :=Arrondi(TotEdt[1].TotDebit + D,CritEDT.Decimale) ;
TotEdt[1].TotCredit :=Arrondi(TotEdt[1].TotCredit + C,CritEDT.Decimale) ;
Case CritEdt.SoldeProg of
  0 :  if Qr11G_SOLDEPROGRESSIF.AsString='X' then Progressif(False,0,0) ;
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
  0 :  if Qr11G_SOLDEPROGRESSIF.AsString='X' then Progressif(True,D,C) ;
  1 :  Progressif(True,D,C) ;
 end ;
AddReport([1,2,3],CritEdt.GL.FormatPrint.Report,D,C,CritEdt.Decimale) ;
{ Il n'y a pas de montants d'Anouveau affectés aux Ruptures }
   Case CritEdt.Rupture of
     rLibres   : AddGroupLibre(LRupt,Q,fbGene,CritEdt.LibreTrie,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
     rRuptures : AddRupt(Lrupt,Qr11G_GENERAL.AsString,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
     rCorresp  : BEGIN
                 Case CritEDT.GL.PlansCorresp Of
                   1 : If Qr11G_CORRESP1.AsString<>'' Then CptRupt:=Qr11G_CORRESP1.AsString+Qr2E_GENERAL.AsString
                                                      Else CptRupt:='.'+Qr2E_GENERAL.AsString ;
                   2 : If Qr11G_CORRESP2.AsString<>'' Then CptRupt:=Qr11G_CORRESP2.AsString+Qr2E_GENERAL.AsString
                                                      Else CptRupt:='.'+Qr2E_GENERAL.AsString ;
                  End ;
                 AddRuptCorres(LRupt,CptRupt,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
                 END ;
     End ;

(*
if CritEdt.Rupture=rLibres then
   BEGIN
   AddGroupLibre(LRupt,Q,fbGene,CritEdt.LibreTrie,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
   END ;
*)
end;

procedure TGLGenAux.BSDMvtBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var CptRupt : String ;
begin
  inherited;
AncienEche:=QR2E_NUMECHE.AsInteger ;
PrintBand:=Affiche And (Not QR2E_AUXILIAIRE.IsNull) ;
if Not PrintBand then exit ;
If IsECC then
   BEGIN
   LE_DEBIT.Caption:='0' ;
   LE_CREDIT.Caption:='0' ;
   END ;
{ Totalisation Compte}
TotAux[1].TotDebit:= Arrondi(TotAux[1].TotDebit + Qr2DEBIT.AsFloat,CritEDT.Decimale) ;
TotAux[1].TotCredit:=Arrondi(TotAux[1].TotCredit + Qr2Credit.AsFloat,CritEDT.Decimale) ;
TotGen[1].TotDebit:= Arrondi(TotGen[1].TotDebit + QR2DEBIT.AsFloat,CritEdt.Decimale) ;
TotGen[1].TotCredit:=Arrondi(TotGen[1].TotCredit + QR2CREDIT.AsFloat,CritEdt.Decimale) ;
{ Total Général }
TotEdt[1].TotDebit:=     Arrondi(TotEdt[1].TotDebit + Qr2DEBIT.AsFloat,CritEDT.Decimale) ;
TotEdt[1].TotCredit:=    Arrondi(TotEdt[1].TotCredit + Qr2Credit.AsFloat,CritEDT.Decimale) ;
TotDesMvts[1].TotDebit :=Arrondi(TotDesMvts[1].TotDebit + QR2DEBIT.AsFloat,CritEdt.Decimale) ;
TotDesMvts[1].TotCredit:=Arrondi(TotDesMvts[1].TotCredit+QR2CREDIT.AsFloat,CritEdt.Decimale) ;
LE_VALIDE.Caption:=ValiQuali(QR2E_VALIDE.AsString,QR2E_QUALIFPIECE.AsString) ;
if QR2E_NUMECHE.AsInteger>1 then UneSeuleEche:=False ;
E_PIECELIGECH.Caption:=QR2E_NUMEROPIECE.AsString+' / '+QR2E_NUMLIGNE.AsString+'   '+QR2E_NUMECHE.AsString ;
Case CritEdt.SoldeProg of             { Affectation Du Calcul du Solde Progressif O/N Sinon, d'aprés l'info sur le compte }
  0 : if Qr11G_SOLDEPROGRESSIF.AsString='X'
         then LE_SOLDE.Caption:=PrintSolde(ProgressDebit+QR2DEBIT.asFloat,ProgressCredit+QR2CREDIT.asFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole)
         else LE_SOLDE.Caption:=PrintSolde(QR2DEBIT.asFloat,QR2CREDIT.asFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
  1 : LE_SOLDE.Caption:=PrintSolde(ProgressDebit+QR2DEBIT.asFloat,ProgressCredit+QR2Credit.asFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
  2 : LE_SOLDE.Caption:=PrintSolde(QR2DEBIT.asFloat,QR2CREDIT.asFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
 end ;
AddGroup(LMulti, [QR2E_DATECOMPTABLE,QR2MULTIECHE],[QR2DEBIT.AsFloat,QR2CREDIT.AsFloat]) ;
AddReport([1,2,3],CritEdt.GL.FormatPrint.Report,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,CritEdt.Decimale) ;

Case CritEdt.Rupture of
  rLibres   : AddGroupLibre(LRupt,Q,fbGene,CritEdt.LibreTrie,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0]) ;
  rRuptures : AddRupt(LRupt,Qr2E_GENERAL.AsString,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0]) ;
  rCorresp  : BEGIN
              Case CritEDT.GL.PlansCorresp Of
                1 : If Qr11G_CORRESP1.AsString<>'' Then CptRupt:=Qr11G_CORRESP1.AsString+Qr2E_GENERAL.AsString
                                                   Else CptRupt:='.'+Qr2E_GENERAL.AsString ;
                2 : If Qr11G_CORRESP2.AsString<>'' Then CptRupt:=Qr11G_CORRESP2.AsString+Qr2E_GENERAL.AsString
                                                   Else CptRupt:='.'+Qr2E_GENERAL.AsString ;
                Else CptRupt:=Qr2E_GENERAL.AsString ;
               End ;
              AddRuptCorres(LRupt,CptRupt,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0]) ;
              END ;
  End ;
(*
if CritEdt.Rupture=rLibres then
   BEGIN
   AddGroupLibre(LRupt,Q,fbGene,CritEdt.LibreTrie,[1,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,0,0,0,0]) ;
   END ;
*)
Quoi:=QuoiMvt ;
end;

procedure TGLGenAux.BSDMvtAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Case CritEdt.SoldeProg of
  0 :  if Qr11G_SOLDEPROGRESSIF.AsString='X' then Progressif(True,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat) ;
  1 :  Progressif(True,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat) ;
  end ;
end;

procedure TGLGenAux.BSDMultiBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
Case NumR of
 0 : BEGIN
     If OkZoomEdt Then PrintBand:=True Else
     Case CritEdt.TotMens of        { Total Mensuel O/N Sinon, d'aprés l'info sur le compte }
       0 : PrintBand:=(Qr11G_TOTAUXMENSUELS.AsString='X') ;
       1 : PrintBand:=True  ;
       2 : PrintBand:=False  ;
      end;
     END ;
 1 : BEGIN
     if AncienEche<=1 then PrintBand:=False else
        begin If OkZoomEdt Then PrintBand:=True Else PrintBand:=CritEDT.GL.TotEche ; end ;
     END ;
 end ;
end;

procedure TGLGenAux.BFCpteGenBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=Affiche ; If Not PrintBand then Exit ; // rony 03/11/97, Pour cpts associés aux ruptures
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
TG_GENERAL_.Caption:=MsgLibel.Mess[14]+'   '+Qr11G_GENERAL.AsString+'   '+Qr11G_LIBELLE.AsString ;
{ Total et Solde par Section, sans oublier le 'Cumul au' }
TotGenDebit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotGen[1].TotDebit, CritEdt.AfficheSymbole ) ;
TotGenCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotGen[1].TotCredit, CritEdt.AfficheSymbole ) ;
TotGenSolde.Caption:=PrintSolde(TotGen[1].TotDebit, TotGen[1].TotCredit ,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
TotDesMvtsDebit.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotDesMvts[1].TotDebit, CritEdt.AfficheSymbole) ;
TotDesMvtsCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotDesMvts[1].TotCredit, CritEdt.AfficheSymbole) ;
TotDesMvtsSolde.Caption:=PrintSolde(TotDesMvts[1].TotDebit,TotDesMvts[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
Quoi:=QuoiGen(2) ;
end;

procedure TGLGenAux.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TotalDebit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotEdt[1].TotDebit, CritEdt.AfficheSymbole ) ;
TotalCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotEdt[1].TotCredit, CritEdt.AfficheSymbole ) ;
TotalSoldes.Caption:=PrintSolde(TotEdt[1].TotDebit,TotEdt[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TGLGenAux.FAuxJokerChange(Sender: TObject);
Var AvecJokerX : Boolean ;
begin
inherited;
AvecJokerX:=Joker(FAux1, FAux2, FAuxJoker) ;
TFaA.Visible:=Not AvecJokerX ;
TFAux.Visible:=Not AvecJokerX ;
TFAuxJoker.Visible:=AvecJokerX ;
end;

procedure TGLGenAux.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var D,C : Double ;
begin
  inherited;
Case QuelReport(CritEdt.GL.FormatPrint.Report,D,C) Of
  1 : BEGIN TITREREPORTB.Caption:=MsgLibel.Mess[6] ; Titre2Rep.Caption:='' ; END ;
  2 : BEGIN TITREREPORTB.Caption:=MsgLibel.Mess[7] ; Titre2Rep.Caption:=StReportGen ; END ;
  3 : BEGIN TITREREPORTB.Caption:=MsgLibel.Mess[7] ; Titre2Rep.Caption:=StReportAux ; END ;
  END ;
Report2Debit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, D, CritEdt.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, C, CritEdt.AfficheSymbole ) ;
Report2Solde.Caption:=PrintSolde(D,C,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
end;

procedure TGLGenAux.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
{
Report1Debit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotEdt[1].TotDebit, CritEdt.AfficheSymbole ) ;
Report1Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotEdt[1].TotCredit, CritEdt.AfficheSymbole ) ;
Report1Solde.Caption:=PrintSolde(TotEdt[1].TotDebit,TotEdt[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
}
Titre1Rep.Caption:=Titre2Rep.Caption ;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
Report1Solde.Caption:=Report2Solde.Caption ;
end;

procedure TGLGenAux.QEcrAfterOpen(DataSet: TDataSet);
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
QR2E_QUALIFPIECE   :=TStringField(QEcr.FindField('E_QUALIFPIECE')) ;
QR2E_QTE1          :=TFloatField(QEcr.FindField('E_QTE1')) ;
QR2E_QTE2          :=TFloatField(QEcr.FindField('E_QTE2')) ;
QR2E_QUALIFQTE1    :=TStringField(QEcr.FindField('E_QUALIFQTE1')) ;
QR2E_QUALIFQTE2    :=TStringField(QEcr.FindField('E_QUALIFQTE2')) ;
QR2E_LETTRAGE      :=TStringField(QEcr.FindField('E_LETTRAGE')) ;
QR2E_JOURNAL       :=TStringField(QEcr.FindField('E_JOURNAL')) ;
QR2DEBIT           :=TFloatField(QEcr.FindField('DEBIT')) ;
QR2CREDIT          :=TFloatField(QEcr.FindField('CREDIT')) ;
QR2MULTIECHE       :=TIntegerField(QEcr.FindField('MULTIECHE')) ;
ChgMaskChamp(Qr2DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
ChgMaskChamp(Qr2CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
IsECC:=(FDevises.Value<>V_PGI.DevisePivot)and(FMontant.ITemIndex=1)and(QEcr.FindField('E_NATUREPIECE').AsString='ECC') ;
end;

procedure TGLGenAux.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr11G_GENERAL         :=TStringField(Q.FindField('G_GENERAL'));
Qr11G_LIBELLE         :=TStringField(Q.FindField('G_LIBELLE'));
Qr11G_COLLECTIF       :=TStringField(Q.FindField('G_COLLECTIF'));
Qr11G_SOLDEPROGRESSIF :=TStringField(Q.FindField('G_SOLDEPROGRESSIF'));
Qr11G_SAUTPAGE        :=TStringField(Q.FindField('G_SAUTPAGE'));
Qr11G_TOTAUXMENSUELS  :=TStringField(Q.FindField('G_TOTAUXMENSUELS'));
If (CritEDT.Rupture=rCorresp) then
   BEGIN
   Qr11G_CORRESP1         :=TStringField(Q.FindField('G_CORRESP1'));
   Qr11G_CORRESP2         :=TStringField(Q.FindField('G_CORRESP2'));
   END ;
end;

procedure TGLGenAux.QAuxAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr12T_AUXILIAIRE:=TStringField(QAux.FindField('T_AUXILIAIRE'));
Qr12T_LIBELLE:=TStringField(QAux.FindField('T_LIBELLE'));
end;

procedure TGLGenAux.BAuxFBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
PrintBand:=Affiche ; If Not PrintBand then Exit ; // rony 03/11/97, Pour cpts associés aux ruptures
If Not OkBand1 Then BEGIN PrintBand:=FALSE ; Exit ; END ;
TT_AUXILIAIRE_.Caption:=MsgLibel.Mess[13]+'   '+Qr12T_AUXILIAIRE.AsString+'   '+Qr12T_LIBELLE.AsString ;
Quoi:=QuoiAux(2) ;
{ Total et Solde par General }
TOTSDEBIT.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotAux[1].TotDebit, CritEDT.AfficheSymbole ) ;
TOTSCREDIT.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotAux[1].TotCredit, CritEDT.AfficheSymbole ) ;
TOTSSOLDE.Caption:=PrintSolde(TotAux[1].TotDebit, TotAux[1].TotCredit ,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
end;

procedure TGLGenAux.QEcrBeforeOpen(DataSet: TDataSet);
begin
  inherited;
QEcr.PArams[0].AsString:=Q.FindField('G_GENERAL').AsString ;
end;

procedure TGLGenAux.FNatureCptChange(Sender: TObject);
begin
  inherited;
If Not QRLoading Then BEGIN FAux1.Clear ; FAux2.Clear ; FAuxJoker.Clear ; END ;
Case FNatureCpt.ItemIndex of
  0     : BEGIN FAux1.ZoomTable:=tzTiers ; FAux2.ZoomTable:=tzTiers ; END ;
  1     : BEGIN FAux1.ZoomTable:=tzTClient ; FAux2.ZoomTable:=tzTClient ; END ;
  2     : BEGIN FAux1.ZoomTable:=tzTDivers ; FAux2.ZoomTable:=tzTDivers ; END ;
  3     : BEGIN FAux1.ZoomTable:=tzTFourn ; FAux2.ZoomTable:=tzTFourn ; END ;
  4     : BEGIN FAux1.ZoomTable:=tzTSalarie ; FAux2.ZoomTable:=tzTSalarie ; END ;
  End ;
(* En CollDivers Zoom sur tzTCrediteur ET tzTDebiteur ET tzTDivers *)

end;

procedure TGLGenAux.BFCpteGenAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([2],CritEdt.GL.FormatPrint.Report) ;
end;

procedure TGLGenAux.BAuxFAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([3],CritEdt.GL.FormatPrint.Report) ;
end;

procedure TGLGenAux.BNouvRechClick(Sender: TObject);
begin
  inherited;
If FAuxJoker.Visible then FAuxJoker.Text:='' ;

end;

procedure TGLGenAux.DLRuptNeedData(var MoreData: Boolean);
Var TotRupt     : Array[0..12] of Double ;
    SumD, SumC  : Double ;
    Quellerupt  : Integer ;
    CodRupt, LibRupt, Lib1, CptRupt,Stcode     : String ;
    Col         : TColor ;
    DansGen, OkOk        : Boolean ;
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
                1 : If Qr11G_CORRESP1.AsString<>'' Then CptRupt:=Qr11G_CORRESP1.AsString+Qr2E_GENERAL.AsString
                                                   Else CptRupt:='.'+Qr2E_GENERAL.AsString ;
                2 : If Qr11G_CORRESP2.AsString<>'' Then CptRupt:=Qr11G_CORRESP2.AsString+Qr2E_GENERAL.AsString
                                                   Else CptRupt:='.'+Qr2E_GENERAL.AsString ;
                Else OkOk:=FALSE ;
                END ;
              If OkOk Then MoreData:=PrintRupt(LRupt,CptRupt,CodRupt,LibRupt,DansGen,QRP.EnRupture,TotRupt) Else MoreData:=FALSE ;
              END ;
  End ;
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
      Case CritEDT.GL.PlansCorresp  Of
        1 : CodeRupt.Caption:=Qr11G_CORRESP1.AsString+'x' ;
        2 : CodeRupt.Caption:=Qr11G_CORRESP2.AsString+'x' ;
        End ;

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

procedure TGLGenAux.BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.Rupture<>rRien) ;
end;


procedure TGLGenAux.FSansRuptClick(Sender: TObject);
begin
  inherited;
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
FRupturesClick(Nil) ;
end;

procedure TGLGenAux.FRupturesClick(Sender: TObject);
begin
  inherited;
if FPlansCo.Checked then FGroupRuptures.Caption:=' '+MsgLibel.Mess[18] ;
if FRuptures.Checked then FGroupRuptures.Caption:=' '+MsgLibel.Mess[17] ;
end;

end.
