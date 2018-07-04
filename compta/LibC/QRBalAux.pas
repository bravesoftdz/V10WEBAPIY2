unit QRBalAux;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  StdCtrls, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, UtilEdt, CritEdt,
  EdtLegal,
  QrRupt, Ent1, HEnt1,
  CpteUtil, HSysMenu, Menus, HTB97, HPanel, UiUtil, tCalcCum ;

procedure BalanceAuxi ;
// CA - 21/12/2001
procedure BalanceAuxiChaine( vCritEdtChaine : TCritEdtChaine );
// CA - 21/12/2001
procedure BalanceAuxiZoom(Crit : TCritEdt) ;
procedure BalanceAuxiLegal(Crit : TCritEdt) ;

type
  TFBalAux = class(TFQR)
    FLigneRupt: TCheckBox;
    TitreColCpt: TQRLabel;
    TitreColLibelle: TQRLabel;
    TTitreColAvant: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel40: TQRLabel;
    QRLabel13: TQRLabel;
    TTitreColSelection: TQRLabel;
    QRLabel14: TQRLabel;
    TTitreColApres: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel21: TQRLabel;
    TTitreColSolde: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel25: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    REPORT3DEBIT: TQRLabel;
    REPORT3CREDIT: TQRLabel;
    REPORT4DEBIT: TQRLabel;
    REPORT4CREDIT: TQRLabel;
    T_AUXILIAIRE: TQRDBText;
    T_LIBELLE: TQRDBText;
    CumDebit: TQRLabel;
    CumCredit: TQRLabel;
    DetDebit: TQRLabel;
    DetCredit: TQRLabel;
    SumDebit: TQRLabel;
    SumCredit: TQRLabel;
    SolDebit: TQRLabel;
    SolCredit: TQRLabel;
    BRupt: TQRBand;
    DetDebRupt: TQRLabel;
    CumCreRupt: TQRLabel;
    CumDebRupt: TQRLabel;
    TCodRupt: TQRLabel;
    SumDebRupt: TQRLabel;
    SumCreRupt: TQRLabel;
    DetCreRupt: TQRLabel;
    SolCreRupt: TQRLabel;
    SolDebRupt: TQRLabel;
    QRLabel33: TQRLabel;
    TotCumDebit: TQRLabel;
    TotCumCredit: TQRLabel;
    TotDetDebit: TQRLabel;
    TotDetCredit: TQRLabel;
    TotSumDebit: TQRLabel;
    TotSumCredit: TQRLabel;
    TotSolDebit: TQRLabel;
    TotSolCredit: TQRLabel;
    QRBand1: TQRBand;
    Trait6: TQRLigne;
    Trait0: TQRLigne;
    Trait5: TQRLigne;
    Trait4: TQRLigne;
    Trait3: TQRLigne;
    Trait2: TQRLigne;
    Ligne1: TQRLigne;
    REPORT5DEBIT: TQRLabel;
    REPORT5CREDIT: TQRLabel;
    REPORT6DEBIT: TQRLabel;
    REPORT6CREDIT: TQRLabel;
    REPORT7DEBIT: TQRLabel;
    REPORT7CREDIT: TQRLabel;
    REPORT8DEBIT: TQRLabel;
    REPORT8CREDIT: TQRLabel;
    DLRupt: TQRDetailLink;
    MsgBox: THMsgBox;
    FCol1: TCheckBox;
    FCol2: TCheckBox;
    FCol3: TCheckBox;
    FCol4: TCheckBox;
    FAvecAN: TCheckBox;
    BLibreHaut: TQRBand;
    TCodRupt_: TQRLabel;
    DLHLibre: TQRDetailLink;
    FSpeed: TCheckBox;
    BRuptTRI: TQRBand;
    DetDebRuptTRI: TQRLabel;
    CumCreRuptTRI: TQRLabel;
    CumDebRuptTRI: TQRLabel;
    TCodRuptTRI: TQRLabel;
    SumDebRuptTRI: TQRLabel;
    SumCreRuptTRI: TQRLabel;
    DetCreRuptTRI: TQRLabel;
    SolCreRuptTRI: TQRLabel;
    SolDebRuptTRI: TQRLabel;
    DLRUPTTRI: TQRDetailLink;
    procedure FormShow(Sender: TObject);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean;      var Quoi: string);
    procedure DLRuptNeedData(var MoreData: Boolean);
    procedure BRuptBeforePrint(var PrintBand: Boolean;      var Quoi: string);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean;      var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean;      var Quoi: string);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure FExerciceChange(Sender: TObject);
    procedure FDateCompta1Change(Sender: TObject);
    procedure FSelectCpteChange(Sender: TObject);
    procedure FAvecANClick(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure FSansRuptClick(Sender: TObject);
    procedure DLHLibreNeedData(var MoreData: Boolean);
    procedure BDetailAfterPrint(BandPrinted: Boolean);
    procedure BDetailCheckForSpace;
    procedure Timer1Timer(Sender: TObject);
    procedure DLRUPTTRINeedData(var MoreData: Boolean);
    procedure BRuptTRIBeforePrint(var PrintBand: Boolean;
      var Quoi: String);
    procedure BRuptTRIAfterPrint(BandPrinted: Boolean);
  private    { Déclarations privées }
    DansTotGen       : Boolean ;
    Librupt, CodRupt : string ;
    // Rony 17/03/97 QBal,QBalCum     : TQuery ;
    TotEdt           : TabTot ;
    LRupt,LRuptTri   : TStringList ;
    Qr11T_AUXILIAIRE : TStringField ;
    Qr11T_LIBELLE    : TStringField ;
    FLoading               : Boolean ;
    IsLegal, OkEnteteLibre,SautPageRupture : Boolean ;
    Qr11T_CORRESP1, Qr11T_CORRESP2 : TStringField ;
    Qr11STTS : tVariantField ;
    Procedure BalAuxiZoom(Quoi : String) ;
    Procedure CalcTotalEdt(AnVD, AnVC, D, C : Double );
    Function  ColonnesOK : Boolean ;
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

Uses Filtre ;

Procedure TFBalAux.BalAuxiZoom(Quoi : String) ;
BEGIN
Inherited ;
ZoomEdt(2,Quoi) ;
END ;

procedure BalanceAuxi ;
var QR : TFBalAux ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:= TFBalAux.Create(Application) ;
Edition.Etat:=etBalAux ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalAuxiZoom ;
QR.IsLegal:=FALSE ;
QR.InitType (nbAux,neBal,msAuxEcr,'QRBALAUX','',TRUE,FALSE,FALSE,Edition) ;
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
END;

// GC - 20/12/2001
procedure BalanceAuxiChaine( vCritEdtChaine : TCritEdtChaine );
var QR : TFBalAux ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:= TFBalAux.Create(Application) ;
Edition.Etat:=etBalAux ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalAuxiZoom ;
QR.IsLegal:=FALSE ;
QR.CRitEdtChaine := vCritEdtChaine;
QR.InitType (nbAux,neBal,msAuxEcr,'QRBALAUX','',TRUE,FALSE,FALSE,Edition) ;
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
END;
// GC - 20/12/2001 - FIN

procedure BalanceAuxiZoom(Crit : TCritEdt) ;
var QR : TFBalAux ;
    Edition : TEdition ;
BEGIN
QR:=TFBalAux.Create(Application) ;
Edition.Etat:=etBalAux ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.BalAuxiZoom ;
 QR.IsLegal:=FALSE ;
 QR.CritEdt:=Crit ;
 QR.InitType (nbAux,neBal,msAuxEcr,'QRBALAUX','',FALSE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure BalanceAuxiLegal(Crit : TCritEdt) ;
var QR : TFBalAux ;
    Edition : TEdition ;
BEGIN
QR:=TFBalAux.Create(Application) ;
Edition.Etat:=etBalAux ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.BalAuxiZoom ;
 QR.IsLegal:=True ;
 QR.CritEdt:=Crit ;
 QR.InitType (nbAux,neBal,msAuxEcr,'QRBALAUX','',TRUE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFBalAux.GenereSQL ;
BEGIN
Inherited ;
If CritEdt.Bal.STTS<>'' Then GenereSQLBaseTS
                        Else GenereSQLBase ;
END ;

procedure TFBalAux.RenseigneCritere;
BEGIN
Inherited ;
(* Rony 17/03/97
RPlanRupt.Caption:=FPlanRupt.Text ;
if FSansRupt.Checked then BEGIN RSansRupt.Caption:='þ' ; RAvecRupt.Caption:='o' ; RSurRupt.Caption:='o' ; END ;
if FAvecRupt.Checked then BEGIN RSansRupt.Caption:='o' ; RAvecRupt.Caption:='þ' ; RSurRupt.Caption:='o' ; END ;
if FSurRupt.Checked  then BEGIN RSansRupt.Caption:='o' ; RAvecRupt.Caption:='o' ; RSurRupt.Caption:='þ' ; END ;
if FPlanRupt.Enabled then RPlanRupt.Caption:=FPlanRupt.Text ;
CaseACocher(FLigneRupt ,RLigneRupt) ;
CaseACocher(FCol1,RCol1)            ; CaseACocher(FCol2,RCol2) ;
CaseACocher(FCol3 ,RCol3)           ; CaseACocher(FCol4,RCol4) ;
*)
END;

procedure TFBalAux.ChoixEdition ;
{ Initialisation des options d'édition : Format des montant vis à vis des devises}
BEGIN
Inherited ;
If CritEdt.Bal.STTS<>'' Then
  BEGIN
  DLRuptTRI.PrintBefore:=FALSE ;
  ChargeGroup(LRuptTRI,['TRI']) ;
  END  ;

DLRupt.PrintBefore:=TRUE ;
Case CritEdt.Rupture of
  rLibres   : BEGIN
              DLRupt.PrintBefore:=FALSE ;
              ChargeGroup(LRupt,['T00','T01','T02','T03','T04','T05','T06','T07','T08','T09']) ;
              END ;
  rRuptures : BEGIN
              ChargeRupt(LRupt, 'RUT', CritEdt.Bal.PlanRupt,CritEdt.Bal.CodeRupt1,CritEdt.Bal.CodeRupt2) ;
              if CritEdt.Bal.RuptOnly=Sur then NiveauRupt(LRupt);
              END ;
  rCorresp  : BEGIN
              ChargeRuptCorresp(LRupt, CritEdt.Bal.PlanRupt,CritEdt.Bal.CodeRupt1,CritEdt.Bal.CodeRupt2, False) ;
              if CritEdt.Bal.RuptOnly=Sur then NiveauRupt(LRupt);
              END ;
  End ;
END ;

procedure TFBalAux.InitDivers ;
BEGIN
Inherited ;
{ Initialise les tableaux de montants }
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
{ Labels sur les bandes }
DateBandeauBalance(TTitreColAvant,TTitreColSelection,TTitreColApres,TTitreColSolde,CritEdt,
                   MsgBox.Mess[0],MsgBox.Mess[1],MsgBox.Mess[2]) ;

BRupt.Frame.DrawTop:=CritEDT.BAL.FormatPrint.PrSepRupt  ;
BRupt.Frame.DrawBottom:=BRupt.Frame.DrawTop ;
ChangeFormatCompte(fbAux,fbAux,Self,T_AUXILIAIRE.Width,1,2,T_AUXILIAIRE.Name) ;
OkEnteteLibre:=True ; SautPageRupture:=FALSE ;
Brupt.ForceNewPage:=FALSE ; BDetail.ForceNewPage:=FALSE ;
END ;

Procedure TFBalAux.FinirPrint ;
BEGIN
Inherited ;
if OkMajEdt Then
   BEGIN
   If IsLegal then MajEditionLegal('BLT',CritEdt.Exo.Code,DateToStr(CritEdt.DateDeb),DateToStr(CritEdt.DateFin),
                   '',TotEdt[3].TotDebit,TotEdt[3].TotCredit,TotEdt[4].TotDebit,TotEdt[4].TotCredit)
              Else MajEdition('BLT',CritEdt.Exo.Code,DateToStr(CritEdt.DateDeb),DateToStr(CritEdt.DateFin),
                   '',TotEdt[3].TotDebit,TotEdt[3].TotCredit,TotEdt[4].TotDebit,TotEdt[4].TotCredit) ;
   END ;
(*  Rony 17/03/97
QBal.Free ;
QBalCum.Free ; *)
If GCalc<>NIL Then BEGIN GCalc.Free ; GCalc:=NIL ; END ;
if CritEdt.Rupture<>rRien then VideRupt(LRupt) ;
If CritEdt.Bal.STTS<>'' Then VideRupt(LRuptTRI) ;
END ;

procedure TFBalAux.RecupCritEdt ;
Var NonLibres : Boolean ;
BEGIN
InHerited ;
With CritEdt Do
  BEGIN
  Bal.CodeRupt1:='' ; Bal.CodeRupt2:='' ;
  NonLibres:=((Rupture=rRuptures) or (Rupture=rCorresp)) ;
  If NonLibres Then Bal.PlanRupt:=FPlanRuptures.Value ;
  BAL.RuptOnly:=QuelleTypeRupt(0,FSAnsRupt.Checked,FAvecRupt.Checked,FSurRupt.Checked) ;
  If NonLibres Then BEGIN Bal.CodeRupt1:=FCodeRupt1.Text ; Bal.CodeRupt2:=FCodeRupt2.Text ; END ;
  Bal.QuelAN:=FQuelAN(FAvecAN.Checked) ;
  If Bal.QuelAN=SansAN Then FCol1.Checked:=FALSE ;
  If (CritEdt.Rupture=rCorresp) Then Bal.PlansCorresp:=FPlanRuptures.ItemIndex+1 ;
  Bal.OnlyCptAssocie:=(Rupture<>rRien) and FOnlyCptAssocie.Checked ;
  Bal.SpeedOk:=FSpeed.checked ;
  With Bal.FormatPrint Do
    BEGIN
    PrSepRupt:=FLigneRupt.Checked ;
    PutTabCol(1, TabColl[1], TitreColCpt.Tag, TRUE) ;
    PutTabCol(2, TabColl[2], TitreColLibelle.Tag, TRUE) ;
    PutTabCol(3, TabColl[3], TTitreColAvant.Tag, FCol1.Checked) ;
    PutTabCol(4, TabColl[4], TTitreColSelection.Tag, FCol2.Checked) ;
    PutTabCol(5, TabColl[5], TTitreColApres.Tag, FCol3.Checked) ;
    PutTabCol(6, TabColl[6], TTitreColSolde.Tag, FCol4.Checked) ;
    END ;
  END ;
END ;

Function TFBalAux.CritOK : Boolean ;
BEGIN
Result:=Inherited CritOK and ColonnesOK;
If Result Then
   BEGIN
//   QBal:=PrepareTotCpt(fbAux,QuelTypeEcr,Dev,Etab,Exo) ;
// Rony 17/03/97 QBal:=PrepareTotCptCum(QBalCum,fbAux,QuelTypeEcr,Dev,Etab,Exo,DevEnP) ;
   Gcalc:=TGCalculCum.create(Un,fbAux,fbAux,QuelTypeEcr,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE,CritEdt.Bal.SpeedOk,'',CritEdt.FiltreSup,CritEdt.Bal.STTS) ;
   GCalc.initCalcul('','','','',CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                    CritEdt.Date1,CritEdt.Date2,TRUE) ;
   //GCalc.InitAnalyseCompte(CritEdt.Cpt1,CritEdt.Cpt2,'','') ; LG 07/02/2007
   END ;
InitReport([1],CritEDT.BAL.FormatPrint.Report) ;
END ;

Function TFBalAux.ColonnesOK : Boolean ;
BEGIN
If OkZoomEdt then begin Result:=True ; exit ; end ;
Result:=(FCol1.Checked or FCol2.Checked  or FCol3.Checked or FCol4.Checked) ;
If Not Result then NumErreurCrit:=7 ;
END ;

procedure TFBalAux.FormShow(Sender: TObject);
begin
HelpContext:=7448000 ;
//Standards.HelpContext:=7448010 ;
//Avances.HelpContext:=7448020 ;
//Mise.HelpContext:=7448030 ;
//Option.HelpContext:=7448040 ;
//TabRuptures.HelpContext:=7448050 ;
  inherited;
Floading:=FALSE ;
TabSup.TabVisible:=False;
TFGenJoker.Visible:=False ;
If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
FCol1.Checked:=TRUE ; FCol2.Checked:=TRUE ; FCol3.Checked:=TRUE ; FCol4.Checked:=TRUE ;
// RL 17/10/97 FSansRuptClick(Nil);
(*
PutTabCol(1,PremTabCol[1],TitreColCpt.Tag,TRUE) ;
PutTabCol(2,PremTabCol[2],TitreColLibelle.Tag,TRUE) ;
PutTabCol(3,PremTabCol[3],TTitreColAvant.Tag,TRUE) ;
PutTabCol(4,PremTabCol[4],TTitreColSelection.Tag,TRUE) ;
PutTabCol(5,PremTabCol[5],TTitreColApres.Tag,TRUE) ;
PutTabCol(6,PremTabCol[6],TTitreColSolde.Tag,TRUE) ;
*)
FOnlyCptAssocie.Enabled:=False ;
{ JP 23/08/04 : FQ 14242 : Nouvelle gestion des filtrs
// GC - 20/12/2001
  if CritEdtChaine.Utiliser then InitEtatchaine('QRBALAUX');
// GC - 20/12/2001 - FIN
ChargeFiltrePourRupture ;}
end;

procedure TFBalAux.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var TotCpt1, TotCpt2, tot : TabTot ;
    Solde                 : Double ;
    MReport               : TabTRep ;
    CptRupt               : String ;
    OkRupt                : Boolean ;
    VCS                   : Variant ;
begin
  inherited;
VCS:=RecupVCS(Qr11STTS) ;
Fillchar(TotCpt1,SizeOf(TotCpt1),#0) ;
Fillchar(TotCpt2,SizeOf(TotCpt2),#0) ;
Fillchar(MReport,SizeOf(MReport),#0) ;
Fillchar(Tot,SizeOf(Tot),#0) ;
Case CritEdt.Bal.TypCpt of
  0 :  PrintBand:=True ;
  3 :  PrintBand:=True ;
 end ;
if PrintBand then
   BEGIN
   Case CritEdt.Rupture of
//     rLibres    : if CritEdt.Bal.OnlyCptAssocie then PrintBand:=DansRuptLibre(Q,fbAux,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
     rLibres    : BEGIN
                  OkRupt:=DansRuptLibre(Q,fbAux,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
                  if CritEdt.Bal.OnlyCptAssocie then PrintBand:=OkRupt ;
                  END ;
     rRuptures  : if CritEdt.Bal.OnlyCptAssocie then PrintBand:=DansRupt(LRupt,Qr11T_AUXILIAIRE.AsString) ;
     rCorresp   : if CritEdt.Bal.OnlyCptAssocie then
                    if CritEDT.Bal.PlansCorresp=1 then PrintBand:=DansRuptCorresp(LRupt,Qr11T_CORRESP1.AsString) Else
                    if CritEDT.Bal.PlansCorresp=2 then PrintBand:=DansRuptCorresp(LRupt,Qr11T_CORRESP2.AsString) ;
//                    if CritEDT.Bal.PlansCorresp=1 then PrintBand:=(Qr11T_CORRESP1.AsString<>'') Else
//                    if CritEDT.Bal.PlansCorresp=2 then PrintBand:=(Qr11T_CORRESP2.AsString<>'') ;

     End;
   If Not PrintBand then Exit ;
   Quoi:=Qr11T_AUXILIAIRE.AsString+' '+Qr11T_LIBELLE.AsString ;
   (* Rony 17/03/97
   ExecuteTotCptCum(QBal,QBalCum,QT_AUXILIAIRE.AsString,'',CritEdt.Date1,CritEdt.Bal.Date11,
                 CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,TotCpt1,FALSE) ;
   ExecuteTotCptCum(QBal,QBalCum,QT_AUXILIAIRE.AsString,'',CritEdt.Bal.Date21,CritEdt.Date2,
                 CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,TotCpt2,FALSE) ;
                 *)
   If CritEdt.Bal.QuelAN<>SansAN Then
      BEGIN
      If VH^.UseTC Then GCAlc.ReInitCalculVCS(Qr11T_AUXILIAIRE.AsString,'',CritEDT.BAL.Date21,CritEDT.Date2,VCS)
                   Else GCAlc.ReInitCalculVCS(Qr11T_AUXILIAIRE.AsString,'',CritEDT.Date1,CritEDT.BAL.Date11,VCS) ;
//      GCAlc.ReInitCalcul(Qr11T_AUXILIAIRE.AsString, '',CritEdt.Date1,CritEdt.Bal.Date11) ;
      GCalc.Calcul ; TotCpt1:=GCalc.ExecCalc.TotCpt ;
      END ;

//   GCAlc.ReInitCalcul(Qr11T_AUXILIAIRE.AsString, '',CritEdt.Bal.Date21,CritEdt.Date2) ;
   GCAlc.ReInitCalculVCS(Qr11T_AUXILIAIRE.AsString,'',CritEDT.BAL.Date21,CritEDT.Date2,VCS) ;
   GCalc.Calcul ; TotCpt2:=GCalc.ExecCalc.TotCpt ;
   CumulVersSolde(TotCpt1[0]) ;

   if CritEdt.Date1=CritEdt.Bal.Date11 Then Fillchar(TotCpt1[1],SizeOf(TotCpt1[1]),#0) ;
   If CritEdt.Bal.QuelAN=SansAN Then
      BEGIN
      TotCpt1[0].TotDebit:=0 ; TotCpt1[0].TotCredit:=0 ;
      TotCpt1[1].TotDebit:=0 ; TotCpt1[1].TotCredit:=0 ;
      END ;
   Tot[1].TotDebit:= Arrondi(TotCpt1[0].TotDebit+TotCpt1[1].TotDebit,CritEdt.Decimale) ;
   Tot[1].TotCredit:=Arrondi(TotCpt1[0].TotCredit+TotCpt1[1].TotCredit,CritEdt.Decimale) ;
   Tot[2].TotDebit:= TotCpt2[1].TotDebit ;
   Tot[2].TotCredit:=TotCpt2[1].TotCredit ;
   Tot[3].TotDebit:= Arrondi(Tot[2].TotDebit+Tot[1].TotDebit,CritEdt.Decimale) ;
   Tot[3].TotCredit:=Arrondi(Tot[2].TotCredit+Tot[1].TotCredit,CritEdt.Decimale) ;

   MReport[1].TotDebit:= Arrondi(TotCpt1[0].TotDebit+TotCpt1[1].TotDebit,CritEdt.Decimale) ;
   MReport[1].TotCredit:=Arrondi(TotCpt1[0].TotCredit+TotCpt1[1].TotCredit,CritEdt.Decimale) ;
   MReport[2].TotDebit:= TotCpt2[1].TotDebit ;
   MReport[2].TotCredit:=TotCpt2[1].TotCredit ;
   MReport[3].TotDebit:= Arrondi(Tot[2].TotDebit+Tot[1].TotDebit,CritEdt.Decimale) ;
   MReport[3].TotCredit:=Arrondi(Tot[2].TotCredit+Tot[1].TotCredit,CritEdt.Decimale) ;

   Tot[4]:=Tot[3] ;
   CumulVersSolde(Tot[4]) ;
   if Tot[4].TotDebit=0 then
      BEGIN
      MReport[4].TotCredit:=Tot[4].TotCredit ;
      SolDebit.Caption:='' ;
      SolCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[4].TotCredit,CritEdt.AfficheSymbole) ;
      END else
      BEGIN
      MReport[4].TotDebit:=Tot[4].TotDebit ;
      SolDebit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[4].TotDebit,CritEdt.AfficheSymbole ) ;
      SolCredit.Caption:='' ;
      END ;
   If CritEdt.Bal.STTS<>'' Then AddGroup(LRuptTri,[Qr11STTS],[Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit]) ;

{ succeptible d'être modifié en fonction de ce que l'on veut afficher }
   Solde:=Tot[4].TotDebit-Tot[4].TotCredit ;
   if Solde>=0 then Quoi:=Quoi+'#'+MsgBox.Mess[5]+' '+
                    AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Solde,CritEdt.AfficheSymbole )
               else Quoi:=Quoi+'#'+MsgBox.Mess[6]+' '+
                    AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Abs(Solde),CritEdt.AfficheSymbole ) ;

   CumDebit.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[1].TotDebit,CritEdt.AfficheSymbole ) ;
   CumCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[1].TotCredit,CritEdt.AfficheSymbole ) ;
   DetDebit.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[2].TotDebit,CritEdt.AfficheSymbole ) ;
   DetCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[2].TotCredit,CritEdt.AfficheSymbole ) ;
   SumDebit.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[3].TotDebit,CritEdt.AfficheSymbole ) ;
   SumCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[3].TotCredit,CritEdt.AfficheSymbole ) ;
   AddReportBAL([1],CritEdt.Bal.FormatPrint.Report,MReport,CritEdt.Decimale) ;

   Case CritEdt.Rupture of
     rLibres   : AddGroupLibre(LRupt,Q,fbAux,CritEdt.LibreTrie,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,0]) ;
     rRuptures : AddRupt(LRupt,Qr11T_AUXILIAIRE.AsString,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,0]) ;
     rCorresp  : BEGIN
                 Case CritEDT.Bal.PlansCorresp Of
                   1 : If Qr11T_CORRESP1.AsString<>'' Then CptRupt:=Qr11T_CORRESP1.AsString+Qr11T_AUXILIAIRE.AsString
                                                      Else CptRupt:='.'+Qr11T_AUXILIAIRE.AsString ;
                   2 : If Qr11T_CORRESP2.AsString<>'' Then CptRupt:=Qr11T_CORRESP2.AsString+Qr11T_AUXILIAIRE.AsString
                                                      Else CptRupt:='.'+Qr11T_AUXILIAIRE.AsString ;
                   Else CptRupt:=Qr11T_AUXILIAIRE.AsString ;
                   End ;
                 AddRuptCorres(LRupt,CptRupt,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,0]) ;
                 END ;
     End ;

   If CritEdt.Bal.RuptOnly=Sur then
      BEGIN
      PrintBand:=False ;
      IF CritEdt.Rupture=rLibres then CalcTotalEdt(Tot[1].TotDebit, Tot[1].TotCredit, Tot[2].TotDebit, Tot[2].TotCredit) ;
      END Else CalcTotalEdt(Tot[1].TotDebit, Tot[1].TotCredit, Tot[2].TotDebit, Tot[2].TotCredit) ;
   END ;
end;

procedure TFBalAux.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TotCumDebit.Caption :=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[1].TotDebit,CritEdt.AfficheSymbole ) ;
TotCumCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[1].TotCredit,CritEdt.AfficheSymbole ) ;

TotDetDebit.Caption :=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[2].TotDebit,CritEdt.AfficheSymbole ) ;
TotDetCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[2].TotCredit,CritEdt.AfficheSymbole ) ;

TotSumDebit.Caption :=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[3].TotDebit,CritEdt.AfficheSymbole ) ;
TotSumCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[3].TotCredit,CritEdt.AfficheSymbole ) ;

TotSolDebit.Caption :=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[4].TotDebit,CritEdt.AfficheSymbole ) ;
TotSolCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[4].TotCredit,CritEdt.AfficheSymbole ) ;

BOTTOMREPORT.enabled:=FALSE ;
//ShowMessage('Après édit Bal : '+DateToStr(CritEdt.Date1)+'   '+DateToStr(CritEdt.Bal.Date11)+'   '+DateToStr(CritEdt.Date2)) ;
end;

procedure TFBalAux.DLRuptNeedData(var MoreData: Boolean);
var SumD, SumC, TotSoldeD, TotSoldeC : Double ;
    TotRupt    : Array[0..12] of Double ;
    CptRupt,Lib1, StCode    : String ;
    Quellerupt : Integer ;
    OkOk,AddTotEdt : Boolean ;
    Col            : TColor ;
    LibRuptInf : Array[1..10] Of TRuptInf ;
begin
  inherited;
MoreData:=false ;
Case CritEdt.Rupture of
  rLibres   : BEGIN
              MoreData:=PrintGroupLibre(LRupt,Q,fbAux,CritEdt.LibreTrie,CodRupt,LibRupt,Lib1,TotRupt,Quellerupt,Col,LibRuptInf,CritEdt.Bal.STTS) ;
              If MoreData then
                 BEGIN
                 StCode:=CodRupt ;
                 Delete(StCode,1,Quellerupt+2) ;
                 MoreData:=DansChoixCodeLibre(StCode,Q,fbAux,CritEdt.LibreCodes1,CritEdt.LibreCodes2, CritEdt.LibreTrie) ;
                 END ;
              END ;
//  rRuptures : MoreData:=PrintRupt(LRupt,Qr11T_AUXILIAIRE.AsString,CodRupt,LibRupt,DansTotGen,QRP.EnRupture,TotRupt) ;
  rRuptures : MoreData:=PrintRupt2(LRupt,Qr11T_AUXILIAIRE.AsString,CodRupt,LibRupt,DansTotGen,QRP.EnRupture,TotRupt,[Qr11T_AUXILIAIRE],Q,CritEdt.Bal.STTS) ;
  rCorresp  : BEGIN
              OkOk:=TRUE ;
              Case CritEDT.Bal.PlansCorresp  Of
                1 : If Qr11T_CORRESP1.AsString<>'' Then CptRupt:=Qr11T_CORRESP1.AsString+Qr11T_AUXILIAIRE.AsString
                                                   Else CptRupt:='.'+Qr11T_AUXILIAIRE.AsString ;
                2 : If Qr11T_CORRESP2.AsString<>'' Then CptRupt:=Qr11T_CORRESP2.AsString+Qr11T_AUXILIAIRE.AsString
                                                   Else CptRupt:='.'+Qr11T_AUXILIAIRE.AsString ;
                Else OkOk:=FALSE ;
                END ;
              If OkOk Then MoreData:=PrintRupt(LRupt,CptRupt,CodRupt,LibRupt,DansTotGen,QRP.EnRupture,TotRupt) Else MoreData:=FALSE ;
              END ;
  End ;
if MoreData then
   BEGIN
   TCodRupt.Width:=TitreColCpt.Width+TitreColLibelle.Width+1 ;
   TCodRupt.Caption:='' ;
   BRupt.Height:=HauteurBandeRuptIni ;
   if CritEdt.Rupture=rLibres then
      BEGIN
      insert(MsgBox.Mess[11]+' ',CodRupt,Quellerupt+2) ;
      TCodRupt.Caption:=CodRupt+' '+Lib1 ;
      //TCodRupt.Caption:=MsgBox.Mess[11]+' '+CodRupt+' '+Lib1 ;
      AlimRuptSup(HauteurBandeRuptIni,QuelleRupt,TCodRupt.Width,BRupt,LibRuptInf,Self) ;
      END Else TCodRupt.Caption:=CodRupt+'   '+LibRupt ;
   SumD:=TotRupt[1]+TotRupt[3] ;
   SumC:=TotRupt[2]+TotRupt[4] ;
   CalCulSolde(SumD,SumC,TotSoldeD,TotSoldeC) ;
   CumDebRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[1], CritEdt.AfficheSymbole) ;
   CumCreRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[2], CritEdt.AfficheSymbole) ;
   DetDebRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[3], CritEdt.AfficheSymbole) ;
   DetCreRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[4], CritEdt.AfficheSymbole) ;
   SumDebRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, SumD  , CritEdt.AfficheSymbole) ;
   SumCreRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, SumC  , CritEdt.AfficheSymbole) ;
   SolDebRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotSoldeD, CritEdt.AfficheSymbole) ;
   SolCreRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotSoldeC, CritEdt.AfficheSymbole) ;
//   SautPageRuptAFaire(CritEdt,BDetail,QuelleRupt) ;
   SautPageRupture:=CritEdt.SautPageRupt And (CritEdt.Bal.RuptOnly=Avec) And (CritEdt.SautPage<>1) And SautPageRuptAFaire(CritEdt,BDetail,QuelleRupt) ;
   END ;

OkEnteteLibre:=MoreData and (CritEdt.Rupture=rLibres) ;

if CritEdt.Bal.RuptOnly=Sur then
   BEGIN
   AddTotEdt:=False ;
   if (CritEdt.Rupture=rLibres) then AddTotEdt:=False else
   if (CritEdt.Rupture=rRuptures) or (CritEdt.Rupture=rCorresp) then AddTotEdt:=DansTotGen ;
   if AddTotEdt then CalcTotalEdt(TotRupt[1], TotRupt[2], TotRupt[3], TotRupt[4]) ;
   END ;
(*  Rony 21/10/97 Avant Grosse MAJ
if (CritEdt.Bal.RuptOnly=Sur) And (DansTotGen) then
   BEGIN
   CalcTotalEdt(TotRupt[1], TotRupt[2], TotRupt[3], TotRupt[4]) ;
   END ;
*)
end;

procedure TFBalAux.BRuptBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.Rupture<>rRien) ;
end;

Procedure TFBalAux.CalcTotalEdt(AnVD, AnVC, D, C : Double );
Var SumD, SumC, TotSoldeD, TotSoldeC : Double ;
BEGIN
SumD:=AnVD+D ;
SumC:=AnVC+C ;
CalCulSolde(SumD,SumC,TotSoldeD,TotSoldeC) ;
TotEdt[1].TotDebit:= Arrondi(TotEdt[1].TotDebit  + AnVD, CritEdt.Decimale) ;
TotEdt[1].TotCredit:=Arrondi(TotEdt[1].TotCredit + AnVC, CritEdt.Decimale) ;
TotEdt[2].TotDebit:= Arrondi(TotEdt[2].TotDebit  + D, CritEdt.Decimale) ;
TotEdt[2].TotCredit:=Arrondi(TotEdt[2].TotCredit + C, CritEdt.Decimale) ;
TotEdt[3].TotDebit:= Arrondi(TotEdt[3].TotDebit  + SumD, CritEdt.Decimale) ;
TotEdt[3].TotCredit:=Arrondi(TotEdt[3].TotCredit + SumC, CritEdt.Decimale) ;
TotEdt[4].TotDebit:= Arrondi(TotEdt[4].TotDebit + TotSoldeD, CritEdt.Decimale) ;
TotEdt[4].TotCredit:=Arrondi(TotEdt[4].TotCredit + TotSoldeC, CritEdt.Decimale) ;
END ;

procedure TFBalAux.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
Report1Debit.Caption :=Report5Debit.Caption ;
Report1Credit.Caption:=Report5Credit.Caption ;
Report2Debit.Caption :=Report6Debit.Caption ;
Report2Credit.Caption:=Report6Credit.Caption ;
Report3Debit.Caption :=Report7Debit.Caption ;
Report3Credit.Caption:=Report7Credit.Caption ;
Report4Debit.Caption :=Report8Debit.Caption ;
Report4Credit.Caption:=Report8Credit.Caption ;
end;

procedure TFBalAux.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var MReport : TabTRep ;
begin
  inherited;
Case QuelReportBAL(CritEdt.Bal.FormatPrint.Report,MReport) of
  1 : TITREREPORTB.Caption:=MsgBox.Mess[7] ;
 end ;
Report5Debit.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[1].TotDebit,  CritEdt.AfficheSymbole ) ;
Report5Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[1].TotCredit, CritEdt.AfficheSymbole ) ;
Report6Debit.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[2].TotDebit,  CritEdt.AfficheSymbole ) ;
Report6Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[2].TotCredit, CritEdt.AfficheSymbole ) ;
Report7Debit.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[3].TotDebit,  CritEdt.AfficheSymbole ) ;
Report7Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[3].TotCredit, CritEdt.AfficheSymbole ) ;
Report8Debit.Caption :=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[4].TotDebit,  CritEdt.AfficheSymbole ) ;
Report8Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[4].TotCredit, CritEdt.AfficheSymbole ) ;
end;

procedure TFBalAux.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr11T_AUXILIAIRE  :=TStringField(Q.FindField('T_AUXILIAIRE'));
Qr11T_LIBELLE     :=TStringField(Q.FindField('T_LIBELLE'));
If CritEDT.Rupture=rCorresp then
   BEGIN
   Qr11T_CORRESP1         :=TStringField(Q.FindField('T_CORRESP1'));
   Qr11T_CORRESP2         :=TStringField(Q.FindField('T_CORRESP2'));
   END ;
If CritEdt.Bal.STTS<>'' Then
  BEGIN
  Qr11STTS:=TVariantField(Q.FindField(CritEdt.Bal.STTS)) ;
  END ;
end;

procedure TFBalAux.FExerciceChange(Sender: TObject);
begin
FLoading:=TRUE ;
  inherited;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
FLoading:=FALSE ;
end;

procedure TFBalAux.FDateCompta1Change(Sender: TObject);
begin
  inherited;
If FLoading Then Exit ;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
end;

procedure TFBalAux.FSelectCpteChange(Sender: TObject);
begin
  inherited;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
end;

procedure TFBalAux.FAvecANClick(Sender: TObject);
begin
  inherited;
FCol1.Checked:=FAvecAN.Checked ;
end;

procedure TFBalAux.FRupturesClick(Sender: TObject);
begin
  inherited;
If FPlansCo.Checked then FGroupRuptures.Caption:=' '+MsgBox.Mess[9] ;
If FRuptures.Checked then FGroupRuptures.Caption:=' '+MsgBox.Mess[8] ;
end;

procedure TFBalAux.FSansRuptClick(Sender: TObject);
begin
  inherited;
FLigneRupt.Enabled:=Not FSansRupt.Checked ;
FLigneRupt.checked:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
FRupturesClick(Nil) ;
end;


procedure TFBalAux.DLHLibreNeedData(var MoreData: Boolean);
begin
  inherited;
MoreData:=FALSE ;
//If BDetail.ForceNewPage Then BEGIN MoreData:=TRUE ; TCodRupt_.Caption:='Rupture Tiers sans tables associées' END ;
//MoreData:=DansRuptLibre2(Q,fbAux,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
//MoreData:=OkEnteteLibre ;
//OkEnteteLibre:=False ;
end;

procedure TFBalAux.BDetailAfterPrint(BandPrinted: Boolean);
begin
  inherited;
If CritEdt.SautPageRupt And (CritEdt.Bal.RuptOnly=Avec) And (CritEdt.SautPage<>1) Then
  BEGIN
  BDetail.ForceNewPage:=FALSE ; SautPageRupture:=FALSE ;
  END Else BDetail.ForceNewPage:=FALSE ;
end;

procedure TFBalAux.BDetailCheckForSpace;
begin
  inherited;
If SautPageRupture Then BDetail.ForceNewPage:=TRUE ;
end;

procedure TFBalAux.Timer1Timer(Sender: TObject);
begin
// GC - 20/12/2001
  inherited;
// GC - 20/12/2001 - FIN
end;

procedure TFBalAux.DLRUPTTRINeedData(var MoreData: Boolean);
Var CodTri,LibTri : String ;
    TotTri : Array[0..12] of Double ;
    NumR : Integer ;
    SumD, SumC, TotSoldeD, TotSoldeC : Double ;
begin
  inherited;
If CritEdt.Bal.STTS='' Then  BEGIN MoreData:=FALSE ; Exit ; END ;
MoreData:=PrintGroup(LRuptTri,Q,[Qr11STTS],CodTri,LibTri,TotTri,NumR);

if MoreData then
   BEGIN
   TCodRuptTRI.Width:=TitreColCpt.Width+TitreColLibelle.Width+1 ;
   TCodRuptTRI.Caption:='TOTAL '+GetNomChampSTTS+' '+CodTri ;
   BRuptTRI.Height:=HauteurBandeRuptIni ;
   SumD:=TotTRI[0]+TotTRI[2] ;
   SumC:=TotTRI[1]+TotTRI[3] ;

   CalCulSolde(SumD,SumC,TotSoldeD,TotSoldeC) ;
   CumDebRuptTRI.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotTRI[0], CritEDT.AfficheSymbole) ;
   CumCreRuptTRI.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotTRI[1], CritEDT.AfficheSymbole) ;
   DetDebRuptTRI.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotTRI[2], CritEDT.AfficheSymbole) ;
   DetCreRuptTRI.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotTRI[3], CritEDT.AfficheSymbole) ;
   SumDebRuptTRI.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumD  , CritEDT.AfficheSymbole) ;
   SumCreRuptTRI.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumC  , CritEDT.AfficheSymbole) ;
   SolDebRuptTRI.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotSoldeD, CritEDT.AfficheSymbole) ;
   SolCreRuptTRI.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotSoldeC, CritEDT.AfficheSymbole) ;
   If CritEdt.SautPageTRI Then
     BEGIN
     If CritEdt.BAL.RuptOnly=Sur Then Brupt.ForceNewPage:=TRUE Else BDetail.ForceNewPage:=TRUE ;
     END ;
   END ;

end;

procedure TFBalAux.BRuptTRIBeforePrint(var PrintBand: Boolean;
  var Quoi: String);
begin
  inherited;
PrintBand:=CritEdt.Bal.STTS<>'' ;
If PrintBand Then
  BEGIN
  if CritEdt.Rupture<>rRien then
    BEGIN
    ReinitRupt(LRupt) ;
    (*
    Case CritEdt.Rupture of
      rLibres   : BEGIN
                  DLRupt.PrintBefore:=FALSE ;
                  ChargeGroup(LRupt,['G00','G01','G02','G03','G04','G05','G06','G07','G08','G09']) ;
                  END ;
      rRuptures : BEGIN
                  ChargeRupt(LRupt, 'RUG',CritEdt.Bal.PlanRupt,CritEdt.Bal.CodeRupt1,CritEdt.Bal.CodeRupt2) ;
                  NiveauRupt(LRupt);
                  END ;
       rCorresp : BEGIN
                  ChargeRuptCorresp(LRupt, CritEdt.Bal.PlanRupt,CritEdt.Bal.CodeRupt1,CritEdt.Bal.CodeRupt2, False) ;
                  NiveauRupt(LRupt);
                  END  ;
      End ;
      *)
    END ;
  END ;
end;

procedure TFBalAux.BRuptTRIAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Brupt.ForceNewPage:=FALSE ;
end;

end.
