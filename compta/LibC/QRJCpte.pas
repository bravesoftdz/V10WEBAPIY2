unit QRJCpte;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, HSysMenu, Menus, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF} StdCtrls, Buttons,
  Hctrls, ExtCtrls, Mask, Hcompte, ComCtrls, CritEdt, UtilEDT, QRRupt,
  CpteUtil, Ent1, HEnt1, HQry, Calcole, HTB97, HPanel, UiUtil, tCalcCum ;

procedure JalCpteGe ;
procedure JalCpteGeZoom(Crit : TCritEdt) ;

type
  TQRJalCpt = class(TFQR)
    HLabel3: THLabel;
    FRefInterne: TEdit;
    HLabel9: THLabel;
    FNumPiece1: TMaskEdit;
    Label12: TLabel;
    FNumPiece2: TMaskEdit;
    FValide: TCheckBox;
    FAvecJalAN: TCheckBox;
    FSautDePage: TCheckBox;
    FLignePerEntete: TCheckBox;
    FLigneCptGene: TCheckBox;
    FLigneJalPied: TCheckBox;
    FLignePerPied: TCheckBox;
    QRLabel10: TQRLabel;
    RRefInterne: TQRLabel;
    QRLabel13: TQRLabel;
    RNumPiece1: TQRLabel;
    QRLabel14: TQRLabel;
    RNumPiece2: TQRLabel;
    QRLabel25: TQRLabel;
    RValide: TQRLabel;
    QRLabel9: TQRLabel;
    RAN: TQRLabel;
    TL_COMPTE: TQRLabel;
    TLE_LIBELLE: TQRLabel;
    TLE_DEBIT: TQRLabel;
    TLE_CREDIT: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    E_JOURNAL: TQRDBText;
    J_NATUREJAL: TQRDBText;
    BHPeriode: TQRBand;
    TLaPeriode: TQRLabel;
    LaPeriodeH: TQRLabel;
    BBDetail: TQRBand;
    BMULTI: TQRBand;
    TDEBMULTI: TQRLabel;
    TCREMULTI: TQRLabel;
    LG_LIBELLE: TQRLabel;
    TCodMulti: TQRLabel;
    BFJournal: TQRBand;
    TTotJal: TQRLabel;
    DEBITJal: TQRLabel;
    CREDITJal: TQRLabel;
    QRLabel33: TQRLabel;
    TOTDEBIT: TQRLabel;
    TOTCREDIT: TQRLabel;
    QRBand1: TQRBand;
    Trait4: TQRLigne;
    Trait3: TQRLigne;
    Trait0: TQRLigne;
    Trait2: TQRLigne;
    Ligne1: TQRLigne;
    REPORT2CREDIT: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    QEcr: TQuery;
    SEcr: TDataSource;
    MsgLibel: THMsgBox;
    DLMULTI: TQRDetailLink;
    DLHMOIS: TQRDetailLink;
    DLMVT: TQRDetailLink;
    QRLigne1: TQRLigne;
    DLJalPied: TQRDetailLink;
    procedure FormShow(Sender: TObject);
    procedure BBDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFJournalBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BMULTIAfterPrint(BandPrinted: Boolean);
    procedure BFJournalAfterPrint(BandPrinted: Boolean);
    procedure BHPeriodeBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure DLMULTINeedData(var MoreData: Boolean);
    procedure QEcrAfterOpen(DataSet: TDataSet);
    procedure DLHMOISNeedData(var MoreData: Boolean);
    procedure DLJalPiedNeedData(var MoreData: Boolean);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure GenereSQL ; Override ;
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
  private
    { Déclarations privées }
//    NumR                     : Integer ;
    LMulti                   : TStringList ;
    TotGen, TotJour          : TabTot ;
    StReportJal{, StReportJalMois }: String ;
    QR1J_JOURNAL, QR1J_NATUREJAL,
    QR1J_LIBELLE             : TStringField ;
    QR2E_JOURNAL             : TStringField ;
    QR2E_GENERAL             : TStringField ;
    QR2G_LIBELLE             : TStringField ;
    QR2E_DATECOMPTABLE       : TDateTimeField ;
    QR2MOIS                  : TDateTimeField ;
    QR2DEBIT                 : TFloatField ;
    QR2CREDIT                : TFloatField ;
    TotCentr                 : Array[0..12] Of Double ;
    PeriodeAEditer,PeriodeEditee,OkpourEditPeriode,
    OkPiedJal, OkMens : Boolean ;
    LabelPeriodeAEditer      : String ;
//    TotCentr                 : Boolean ;
    procedure GenereSQLSub ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

procedure JalCpteGe ;
var QR : TQRJalCpt ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TQRJalCpt.Create(Application) ;
Edition.Etat:=etJalDivGen ;
// QR.QRP.QRPrinter.OnSynZoom:=QR.JournalZoom ;
QR.InitType(nbJal,neJalR,msRien,'QRJCPTE','',TRUE,FALSE,FALSE,Edition) ;
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

procedure JalCpteGeZoom(Crit : TCritEdt) ;
var QR : TQRJalCpt ;
    Edition : TEdition ;
BEGIN
QR:=TQRJalCpt.Create(Application) ;
Edition.Etat:=etJalDivGen ;
try
// QR.QRP.QRPrinter.OnSynZoom:=QR.JournalZoom ;
 QR.CritEDT:=Crit ;
 QR.InitType(nbJal,neJalR,msRien,'QRJDIVIS','',FALSE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;


procedure TQRJalCpt.FormShow(Sender: TObject);
begin
FValide.State:=cbGrayed ;
HelpContext:=7409000 ;
//Standards.HelpContext:=7409010 ;
//Avances.HelpContext:=7409020 ;
//Mise.HelpContext:=7409030 ;
//Option.HelpContext:=7409040 ;
  inherited;
end;

procedure TQRJalCpt.FinirPrint;
BEGIN
Inherited ;
QEcr.Close ;
//VideGroup(LMulti) ;
VideRecap(LMulti) ;
END ;

function  TQRJalCpt.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ;
If Result Then
   BEGIN
   END ;
END ;

procedure TQRJalCpt.RecupCritEdt ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  RefInterne:=FRefInterne.Text ;
  If FNumPiece1.Text<>'' then JalR.NumPiece1:=StrToInt(FNumPiece1.Text) else JalR.NumPiece1:=0 ;
  If FNumPiece2.Text<>'' then JalR.NumPiece2:=StrToInt(FNumPiece2.Text) else JalR.NumPiece2:=999999999 ;
  if FValide.State=cbGrayed then Valide:='g' ;
  if FValide.State=cbChecked then Valide:='X' ;
  if FValide.State=cbUnChecked then Valide:='-' ;
  If FSautDePage.Checked Then SautPage:=1 ;
  JalR.AvecJalAN:=FAvecJalAN.Checked ;
  With JalR.FormatPrint Do
       BEGIN
       PrSepMontant:=FTrait.Checked ;
       PrSepCompte[2]:=FLignePerEntete.Checked ;
       PrSepCompte[3]:=FLigneJalPied.Checked ;
       PrSepCompte[4]:=FLignePerPied.Checked ;
       PrSepCompte[5]:=FLigneCptGene.Checked ;
       END ;
  END ;
END ;

procedure TQRJalCpt.InitDivers ;
//Var ll : Integer ;
BEGIN
Inherited ;
LaPeriodeH.Caption:='' ;
Fillchar(TotGen,SizeOf(TotGen),#0) ;
BDetail.ForceNewPage:=CritEdt.SautPage=1 ;
BHPeriode.Frame.DrawTop:=CritEdt.JalR.FormatPrint.PrSepCompte[2] ;
BFJournal.Frame.DrawTop:=CritEdt.JalR.FormatPrint.PrSepCompte[3] ;
PeriodeAEditer:=True ; PeriodeEditee:=FALSE ; OkPourEditPeriode:=TRUE ;
LabelPeriodeAEditer:='' ;
ChangeFormatCompte(fbGene,fbGene,Self,TCodMulti.Width,1,2,TCodMulti.Name) ;
END ;

procedure TQRJalCpt.ChoixEdition ;
BEGIN
Inherited ;
//ChargeGroup(Lmulti,['MOIS','GENE']) ;
ChargeRecap(LMulti) ;

END ;

procedure TQRJalCpt.GenereSQL ;
Var St, StJalAuto : String ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
Q.Sql.Clear ;
St:='Select J_JOURNAL, J_LIBELLE, J_NATUREJAL' ;
Q.SQL.Add(St) ;
{ Tables explorées par la SQL }
St:=' From JOURNAL ' ;
Q.SQL.Add(St) ;
Q.SQL.Add(' Where J_JOURNAL<>"'+w_w+'" ') ;
{ Construction de la clause Where de la SQL }
St:='' ;
If CritEdt.Joker then
   BEGIN Q.SQL.Add(' And J_JOURNAL like "'+TraduitJoker(CritEdt.Cpt1)+'"') ;END
   Else
   BEGIN
   if CritEdt.Cpt1<>'' then Q.SQL.Add(' And J_JOURNAL>="'+CritEdt.Cpt1+'"') ;
   if CritEdt.Cpt2<>'' then Q.SQL.Add(' And J_JOURNAL<="'+CritEdt.Cpt2+'"') ;
   END ;
If St<>'' Then Q.SQL.Add(St) ;
St:=' AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ODA" AND J_NATUREJAL<>"ANA" ' ;
If Not CritEdt.JalR.AvecJalAN Then St:=St+' AND J_NATUREJAL<>"ANO" ' ;
If St<>'' Then Q.SQL.Add(St) ;
if (OkV2 and (V_PGI.Confidentiel<>'1')) then
   BEGIN
   StJalAuto:=VH^.JalAutorises ;
   if StJalAuto<>'' then
      BEGIN St:=St+' And '+AnalyseCompte(StJalAuto,fbJal,False,False) ; Q.SQL.Add(St) ; END ;
   END ;
St:=' Order By J_JOURNAL' ;
Q.SQL.Add(St) ;
ChangeSql(Q) ; Q.Open ;
GenereSQLSub ;
END;

procedure TQRJalCpt.GenereSQLSUB ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
QEcr.Close ;
QEcr.SQL.Clear ;
{ Construction de la clause Select de la SQL }
QEcr.SQL.Add(' Select E_JOURNAL, E_GENERAL, E_DATECOMPTABLE, G_LIBELLE, E_DATECOMPTABLE as MOIS, ') ;
QEcr.SQL.Add(' E_EXERCICE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE, ') ;
Case CritEdt.Monnaie of
  0 : BEGIN QEcr.SQL.Add(' E_DEBIT DEBIT, E_CREDIT CREDIT ')         ; END ;
  1 : BEGIN QEcr.SQL.Add(' E_DEBITDEV DEBIT, E_CREDITDEV CREDIT ')   ; END ;
end ;
QEcr.SQL.Add(' From ECRITURE') ;
QEcr.SQL.Add('LEFT JOIN GENERAUX On G_GENERAL=E_GENERAL ') ;
{ Construction de la clause Where de la SQL }
QEcr.SQL.Add(' Where ') ;
QEcr.SQL.Add(' E_JOURNAL=:J_JOURNAL ') ;
QEcr.SQL.Add(' and E_EXERCICE="'+CritEdt.Exo.Code+'"') ;
QEcr.SQL.Add(' And E_DATECOMPTABLE>="'+USDATETIME(CritEdt.Date1)+'"') ;
QEcr.SQL.Add(' And E_DATECOMPTABLE<="'+USDATETIME(CritEdt.Date2)+'"') ;
QEcr.SQL.Add(TraduitNatureEcr(CritEdt.QualifPiece, 'E_QUALIFPIECE', true, CritEdt.ModeRevision)) ;
QEcr.SQL.Add(' and E_QUALIFPIECE<>"C" ') ;
//QCpt.SQL.Add(' And J_NATUREJAL<>"ODA"  And J_NATUREJAL<>"ANA"') ;
QEcr.SQL.Add(' And E_NUMEROPIECE>='+IntToStr(CritEdt.JalR.NumPiece1)+' and E_NUMEROPIECE<='+IntToStr(CritEdt.JalR.NumPiece2)+' ') ;
if CritEdt.RefInterne<>'' then QEcr.SQL.Add('and UPPER(E_REFINTERNE) like "'+TraduitJoker(CritEdt.RefInterne)+'" ' );
if CritEdt.Etab<>'' then QEcr.SQL.Add(' And E_ETABLISSEMENT="'+CritEdt.Etab+'"') ;
if CritEdt.Valide<>'g' then QEcr.SQL.Add(' And E_VALIDE="'+CritEdt.Valide+'"') ;
if CritEdt.DeviseSelect<>'' then QEcr.SQL.Add(  ' And E_DEVISE="'+CritEdt.DeviseSelect+'"') ;
{ Construction de la clause Order By de la SQL }
QEcr.SQL.Add(' Order By E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_GENERAL, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE ') ;
ChangeSql(QEcr) ;
QEcr.Open ;
END;

procedure TQRJalCpt.RenseigneCritere ;
{ Récupération des champs du multicritère dans l'entête d'état : OK }
BEGIN
Inherited ;
RRefInterne.Caption:=FRefInterne.Text ;
RNumPiece1.Caption:=IntToStr(CritEdt.JalR.NumPiece1) ; RNumPiece2.Caption:=IntToStr(CritEdt.JalR.NumPiece2) ;
CaseACocher(FValide,RValide)   ; CaseACocher(FAvecJalAN,RAN)   ;
CaseACocher(FAvecJalAN,RAN)   ;
END ;

procedure TQRJalCpt.BBDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var SDate1,SDate2,StCompte1,StCompte2 : String ;
begin
  inherited;
If QR2E_JOURNAL.IsNull then
   BEGIN
   PrintBand:=False ;
   Exit ;
   END ;
StCompte1:=Format_String(QR2E_GENERAL.AsString,VH^.Cpta[fbGene].Lg) ;
StCompte2:=Format_String('zzzzzzzzzzzzzzzz',VH^.Cpta[fbGene].Lg) ;
SDate1:=FormatDatetime('mmyyyy',Qr2E_DATECOMPTABLE.AsDateTime) ;
SDate2:=FormatDatetime('mmmm yyyy',Qr2E_DATECOMPTABLE.AsDateTime) ;
AddRecap(LMulti,[Sdate1,StCompte1+QR2G_LIBELLE.AsString],[QR2E_GENERAL.AsString,SDate2],[QR2DEBIT.AsFloat,QR2CREDIT.AsFloat]) ;
AddRecap(LMulti,[Sdate1,StCompte2],[StCompte2,SDate2],[QR2DEBIT.AsFloat,QR2CREDIT.AsFloat]) ;
//AddGroup(Lmulti,[QR2MOIS,QR2E_GENERAL],[QR2DEBIT.AsFloat,QR2Credit.AsFloat]) ;
//LabelPeriodeAEditer:=Sdate2 ;
// Rony 21/05/97 If LabelPeriodeAEditer='' Then LabelPeriodeAEditer:=Sdate2 ;
AddReport([1,2],CritEdt.JalR.FormatPrint.Report,QR2DEBIT.AsFloat,QR2Credit.AsFloat,CritEdt.Decimale) ;
{ Incrémentation des montants pour ... }
{ Chaque Journal }
TotJour[1].TotDebit:=  Arrondi(TotJour[1].TotDebit+QR2DEBIT.AsFloat, CritEdt.Decimale) ;
TotJour[1].TotCredit:= Arrondi(TotJour[1].TotCredit+QR2Credit.AsFloat, CritEdt.Decimale) ;
{ Le Total Gébéral }
TotGen[1].TotDebit:=  Arrondi(TotGen[1].TotDebit+QR2DEBIT.AsFloat, CritEdt.Decimale) ;
TotGen[1].TotCredit:= Arrondi(TotGen[1].TotCredit+QR2Credit.AsFloat, CritEdt.Decimale) ;

PrintBand:=False ;
end;

procedure TQRJalCpt.BFJournalBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
//  PrintBand:=OkPiedJal ;
TTotJal.Caption:=MsgLibel.Mess[4]+'    '+QR2E_JOURNAL.AsString ;
{ Affiche les Totaux Par journal }
DEBITJal.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotJour[1].TotDebit,CritEdt.AfficheSymbole ) ;
CREDITJal.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotJour[1].TotCredit,CritEdt.AfficheSymbole ) ;
LabelPeriodeAEditer:='' ;
end;


procedure TQRJalCpt.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TOTDEBIT.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotGEN[1].TotDebit,CritEdt.AfficheSymbole ) ;
TOTCREDIT.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotGEN[1].TotCredit,CritEdt.AfficheSymbole ) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TQRJalCpt.BMULTIAfterPrint(BandPrinted: Boolean);
begin
  inherited;
(*If OkMens Then
   BEGIN
   InitReport([3],CritEdt.JalR.FormatPrint.Report) ;
   END ;
OkMens:=False ;
(*If OkPourEditperiode Then
   BEGIN
   PeriodeEditee:=TRUE ; PeriodeAEditer:=FALSE ; OkPourEditPeriode:=FALSE ;
   END ; *)
end;

procedure TQRJalCpt.BFJournalAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Fillchar(TotJour,SizeOf(TotJour),#0) ;
// Roro
InitReport([2],CritEdt.JalR.FormatPrint.Report) ;
end;

procedure TQRJalCpt.BHPeriodeBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=not QR2E_DATECOMPTABLE.IsNull ;
IF PrintBand then
   BEGIN
   //StReportJalMois:=Qr1J_Journal.AsString+' / '+LabelPeriodeAEditer ;
   //InitReport([3],CritEdt.JalR.FormatPrint.Report) ;
   TLaPeriode.Caption:=MsgLibel.Mess[2]+' '+LabelPeriodeAEditer ;
   END ;
end;

procedure TQRJalCpt.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;

end;

procedure TQRJalCpt.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PeriodeAEditer:=True ; PeriodeEditee:=FALSE ; OkPourEditPeriode:=TRUE ;
LabelPeriodeAEditer:='' ;
E_JOURNAL.Caption:='  '+QR1J_JOURNAL.AsString+'  '+QR1J_LIBELLE.AsString ;
StReportJal:=QR1J_JOURNAL.AsString ;
VideRecap(LMulti) ; ChargeRecap(LMulti) ;
InitReport([2],CritEdt.JalR.FormatPrint.Report) ;
J_NATUREJAL.Caption:=MsgLibel.Mess[9]+'  '+RechDom('ttNatJal',QR1J_NATUREJAL.AsString,False) ;
OkPiedJal:=PrintBand ;
end;

procedure TQRJalCpt.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var D,C : Double ;
begin
  inherited;
D:=0 ; C:=0 ;
Case QuelReport(CritEdt.JalR.FormatPrint.Report,D,C) Of
  1 : TitreReportB.Caption:=MsgLibel.Mess[5] ;
  2 : TitreReportB.Caption:=MsgLibel.Mess[6]+' '+StReportJal ;
  //3 : TitreReportB.Caption:=MsgLibel.Mess[10]+' '+StReportJalMois ;
  END ;
Report2Debit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, D, CritEdt.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, C, CritEdt.AfficheSymbole ) ;
end;

procedure TQRJalCpt.DLMULTINeedData(var MoreData: Boolean);
{
Var CodMulti,LibMulti : String ;
    TotMulti : Array[0..12] of Double ;
}
Var Cod,Lib, StGene : String ;
    RuptPeriode : Boolean ;
begin
  inherited;
//MoreData:=PrintGroup2(Lmulti,QEcr,[QR2E_DATECOMPTABLE,QR2E_GENERAL],CodMulti,LibMulti,TotMulti,NumR);
MoreData:=PrintRecap(LMulti, Cod, Lib, TotCentr) ;
If Not MoreData Then Exit ;
RuptPeriode:=Copy(Cod,7,3)='zzz' ;
if MoreData Then
   BEGIN
   If Ruptperiode Then
      BEGIN PeriodeAEditer:=TRUE ; PeriodeEditee:=FALSE ; LabelPeriodeAEditer:=Copy(Lib,VH^.Cpta[fbGene].Lg+2,Length(Lib)-VH^.Cpta[fbGene].Lg) ;
      END Else
      BEGIN
      If PeriodeAEditer Then
         BEGIN
         If (Not PeriodeEditee) Then
            BEGIN
            OkPourEditPeriode:=TRUE ; LabelPeriodeAEditer:=Copy(Lib,VH^.Cpta[fbGene].Lg+2,Length(Lib)-VH^.Cpta[fbGene].Lg) ;
            END Else OkPourEditPeriode:=FALSE ;
         END Else OkPourEditPeriode:=FALSE ;
      END ;
   END Else OkPourEditPeriode:=FALSE ;
OkMens:=RuptPeriode ;
If RuptPeriode Then { Rupture/Mois}
   BEGIN
   //AddReport([3],CritEdt.JalR.FormatPrint.Report,TotCentr[0],TotCentr[1],CritEdt.Decimale) ;
   BMulti.Frame.DrawTop:=CritEdt.JalR.FormatPrint.PrSepCompte[4] ;
   BMulti.Font.color:=clPurple  ; BMulti.Font.Style:=[fsBold] ;
//   LaPeriodeH.Caption:=Copy(Cod,1,6) ;
   LaPeriodeH.Caption:='' ; TCodMulti.Caption:='' ;
   LG_LIBELLE.Alignment:=taRightJustify ;
   LG_LIBELLE.Caption:=MsgLibel.Mess[3]+' '+LabelPeriodeAEditer;
   // Rony 20/05/97 LG_LIBELLE.Caption:=MsgLibel.Mess[3]+' '+Copy(Cod,1,6);
   TDEBMULTI.Caption:=AfficheMontant(CRITEdt.FormatMontant, CRITEdt.Symbole, TotCentr[0] , CRITEdt.AfficheSymbole) ;
   TCREMULTI.Caption:=AfficheMontant(CRITEdt.FormatMontant, CRITEdt.Symbole, TotCentr[1] , CRITEdt.AfficheSymbole) ;
   END Else
   BEGIN  { Rupture/General }
   BMulti.Frame.DrawTop:=CritEdt.JalR.FormatPrint.PrSepCompte[5] ;
   BMulti.Font.color:=clblack  ; BMulti.Font.Style:=[] ;
   StGene:=Copy(Cod,7,VH^.Cpta[fbGene].Lg) ;
   TCodMulti.Caption:=StGene ; LG_LIBELLE.Alignment:=taLeftJustify ;
   LG_LIBELLE.Caption:=Copy(Cod,7+VH^.Cpta[fbGene].Lg,(Length(Cod)+1)-(VH^.Cpta[fbGene].Lg+7)) ;
   //TCodMulti.Caption:=Cod ; LG_LIBELLE.Caption:=Lib;
   TDEBMULTI.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCentr[0], CritEdt.AfficheSymbole) ;
   TCREMULTI.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotCentr[1], CritEdt.AfficheSymbole) ;
   END ;
(*
if ((MoreData) and (Not QEcr.EOF) And (NumR=0)) then OKEnteteM:=True ;
Case NumR of  { Rupture/Mois }
   0 : BEGIN
       BMulti.Frame.DrawTop:=CritEdt.JalR.FormatPrint.PrSepCompte[4] ;
       BMulti.Font.color:=clPurple  ; BMulti.Font.Style:=[fsBold] ;
       LaPeriodeH.Caption:=CodMulti ;
       TCodMulti.Caption:=''        ;  LG_LIBELLE.Caption:=MsgLibel.Mess[7]+' '+CodMulti;
       TDEBMULTI.Caption:=AfficheMontant(CRITEdt.FormatMontant, CRITEdt.Symbole, TotMulti[0] , CRITEdt.AfficheSymbole) ; { + 'ANouveau' + 'Cumul au' en Débit  }
       TCREMULTI.Caption:=AfficheMontant(CRITEdt.FormatMontant, CRITEdt.Symbole, TotMulti[1] , CRITEdt.AfficheSymbole) ; { + 'ANouveau' + 'Cumul au' en Crédit }
       END ;
   1 : BEGIN  { Rupture/General }
       BMulti.Frame.DrawTop:=CritEdt.JalR.FormatPrint.PrSepCompte[5] ;
       BMulti.Font.color:=clblack  ; BMulti.Font.Style:=[] ;
       TCodMulti.Caption:=CodMulti ; LG_LIBELLE.Caption:=QR2G_LIBELLE.AsString ;
       TDEBMULTI.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotMulti[0], CritEdt.AfficheSymbole) ;
       TCREMULTI.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotMulti[1], CritEdt.AfficheSymbole) ;
       END ;
   END ;
*)
end;

procedure TQRJalCpt.QEcrAfterOpen(DataSet: TDataSet);
begin
  inherited;
QR2E_JOURNAL:=TStringField(QEcr.FindField('E_JOURNAL')) ;
QR2E_GENERAL:=TStringField(QEcr.FindField('E_GENERAL')) ;
QR2G_LIBELLE:=TStringField(QEcr.FindField('G_LIBELLE')) ;
QR2E_DATECOMPTABLE:=TDateTimeField(QEcr.FindField('E_DATECOMPTABLE')) ;
QR2E_DATECOMPTABLE.Tag:=1 ;
QR2MOIS:=TDateTimeField(QEcr.FindField('E_DATECOMPTABLE')) ;
QR2MOIS.Tag:=1 ;
QR2DEBIT:=TFloatField(QEcr.FindField('DEBIT')) ;
QR2CREDIT:=TFloatField(QEcr.FindField('CREDIT')) ;
end;

procedure TQRJalCpt.DLHMOISNeedData(var MoreData: Boolean);
begin
  inherited;
MoreData:=OkPourEditperiode;
If OkPourEditperiode Then
   BEGIN
   PeriodeEditee:=TRUE ; PeriodeAEditer:=FALSE ; OkPourEditPeriode:=FALSE ;
   END ;
end;

procedure TQRJalCpt.DLJalPiedNeedData(var MoreData: Boolean);
begin
  inherited;
MoreData:=OkPiedJal ;
OkPiedJal:=False ;
end;

procedure TQRJalCpt.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
QR1J_JOURNAL:=TStringField(Q.FindField('J_JOURNAL')) ;
QR1J_NATUREJAL:=TStringField(Q.FindField('J_NATUREJAL')) ;
QR1J_LIBELLE:=TStringField(Q.FindField('J_LIBELLE')) ;
end;

end.
