unit QRBalGen1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, StdCtrls, hmsgbox, HQuickrp, DB, DBTables, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, UtilEdt, HEnt1, Ent1, QRRupt, CpteUtil,
{$IFNDEF IMP}
{$IFNDEF CCMP}
  EdtLegal,
{$ENDIF}
  CALCOLE,
{$ENDIF}
  CritEDT, Menus, HSysMenu, HTB97, HPanel, UiUtil, tCalcCum
// CA - 20/11/2001
{$IFNDEF IMP}
  ,BalSit
{$ENDIF}  
;

procedure BalanceAno ;
procedure BalanceClo ;
procedure BalanceGene(VoirZoneEcart : Boolean = FALSE) ;
procedure BalanceGeneZoom(Crit : TCritEdt) ;
procedure BalanceGeneCERG ;
procedure BalanceGeneLegal(Crit : TCritEdt) ;
// GC - 20/12/2001
procedure BalanceGeneComp;
procedure BalanceGeneChaine( vCritEdtChaine : TCritEdtChaine );
// GC - 20/12/2001 - FIN
Function  Get_Balance(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif : String) : Variant ;
Function  GetBalanceSimple(Typ,Cod1 : String ; AvecAno,Etab,Devi,Exo : String ; var Dat11,Dat22 : TDateTime ; Collectif : Boolean) : String ;

Type TNatEtat = (neRien,neAno,neClo) ;

type
  TFBalGen1 = class(TFQR)
    FLigneRupt: TCheckBox;
    TitreColCpt: TQRLabel;
    TitreColLibelle: TQRLabel;
    TTitreColAvant: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel30: TQRLabel;
    TTitreColSelection: TQRLabel;
    QRLabel38: TQRLabel;
    QRLabel39: TQRLabel;
    TTitreColApres: TQRLabel;
    QRLabel41: TQRLabel;
    QRLabel42: TQRLabel;
    TTitreColSolde: TQRLabel;
    QRLabel44: TQRLabel;
    QRLabel45: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    REPORT3DEBIT: TQRLabel;
    REPORT3CREDIT: TQRLabel;
    REPORT4DEBIT: TQRLabel;
    REPORT4CREDIT: TQRLabel;
    G_GENERAL: TQRDBText;
    G_LIBELLE: TQRDBText;
    LibCollectif: TQRLabel;
    CumDebit: TQRLabel;
    CumCredit: TQRLabel;
    DetDebit: TQRLabel;
    DetCredit: TQRLabel;
    SumDebit: TQRLabel;
    SumCredit: TQRLabel;
    SolDebit: TQRLabel;
    SolCredit: TQRLabel;
    solsolcredit: TQRLabel;
    solsoldebit: TQRLabel;
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
    DLRupt: TQRDetailLink;
    MsgBox: THMsgBox;
    REPORT5DEBIT: TQRLabel;
    REPORT5CREDIT: TQRLabel;
    REPORT6DEBIT: TQRLabel;
    REPORT6CREDIT: TQRLabel;
    REPORT7DEBIT: TQRLabel;
    REPORT7CREDIT: TQRLabel;
    REPORT8DEBIT: TQRLabel;
    REPORT8CREDIT: TQRLabel;
    MsgLabel: THMsgBox;
    FCol1: TCheckBox;
    FCol2: TCheckBox;
    FCol3: TCheckBox;
    FCol4: TCheckBox;
    FCollectif: TCheckBox;
    FAvecAN: TCheckBox;
    FBilGes: TCheckBox;
    FOnlyClo: TCheckBox;
    BBil: TQRBand;
    BGes: TQRBand;
    TTotBil: TQRLabel;
    TotBilCumDebit: TQRLabel;
    TotBilCumCredit: TQRLabel;
    TotBilDetDebit: TQRLabel;
    TotBilDetCredit: TQRLabel;
    TotBilSumDebit: TQRLabel;
    TotBilSumCredit: TQRLabel;
    TotBilSolDebit: TQRLabel;
    TotBilSolCredit: TQRLabel;
    TotGesSumDebit: TQRLabel;
    TotGesSumCredit: TQRLabel;
    TTotGes: TQRLabel;
    TotGesCumDebit: TQRLabel;
    TotGesCumCredit: TQRLabel;
    TotGesDetDebit: TQRLabel;
    TotGesDetCredit: TQRLabel;
    TotGesSolDebit: TQRLabel;
    TotGesSolCredit: TQRLabel;
    DLBil: TQRDetailLink;
    DLGes: TQRDetailLink;
    FSpeed: TCheckBox;
    FSansEcart: TCheckBox;
    GBComparatif: TGroupBox;
    LDateComparatif1: TLabel;
    Label3: TLabel;
    FDateComparatif1: THCritMaskEdit;
    FDateComparatif2: THCritMaskEdit;
    FExercice2: THValComboBox;
    FAvecComparatif: TCheckBox;
    RBCOMPEXO: TRadioButton;
    GBComparatif2: TGroupBox;
    TFBALSIT: THLabel;
    TFLIBELLEBALSIT: THLabel;
    FBALSIT: THCritMaskEdit;
    RBCOMPBALSIT: TRadioButton;
    SolCumDebit: TQRLabel;
    SolCumCredit: TQRLabel;
    SolDetDebit: TQRLabel;
    SolDetCredit: TQRLabel;
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
    FOkRegroup: TCheckBox;
    FRegroup: TEdit;
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure DLRuptNeedData(var MoreData: Boolean);
    procedure BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FormShow(Sender: TObject);
    procedure BBilanGenBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure FExerciceChange(Sender: TObject);
    procedure FDateCompta1Change(Sender: TObject);
    procedure FSelectCpteChange(Sender: TObject);
    procedure FAvecANClick(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure FSansRuptClick(Sender: TObject);
    procedure DLBilNeedData(var MoreData: Boolean);
    procedure BGesBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure DLGesNeedData(var MoreData: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FAvecComparatifClick(Sender: TObject);
    procedure FExercice2Change(Sender: TObject);
    procedure OnTypeComparatifClick(Sender: TObject);
    procedure FBALSITChange(Sender: TObject);
    procedure FBALSITElipsisClick(Sender: TObject);
    procedure FMontantClick(Sender: TObject);
    procedure DLRUPTTRINeedData(var MoreData: Boolean);
    procedure BRuptTRIBeforePrint(var PrintBand: Boolean;
      var Quoi: String);
    procedure BDetailAfterPrint(BandPrinted: Boolean);
    procedure BRuptAfterPrint(BandPrinted: Boolean);
    procedure FOkRegroupClick(Sender: TObject);
    procedure FRegroupDblClick(Sender: TObject);
  private { Déclarations privées }
    LibRupt, CodRupt : string ;
    DansTotal        : Boolean ;
    NET              : TNatEtat ;
    // Rony 10/03/97 * QBal, QBalC,QBalCum : TQuery ;
    QBalC : TQuery ;
    LRupt, LBilGes, LRuptTri,LRegroup      : TStringList ;
    TotEdt      : TabTot ;
    TotGes              : TabTot ;
    Qr11G_GENERAL       : TStringField ;
    Qr11G_LIBELLE       : TStringField ;
    Qr11G_COLLECTIF     : TStringField ;
    Qr11G_CORRESP1, Qr11G_CORRESP2 : TStringField ;
    Qr11STTS : tVariantField ;
    FLoading, IsLegal, OkTotGestion : Boolean ;
    IsLoading : Boolean ;
    VoirZoneEcart : Boolean ;
    procedure GenereSQL ; Override ;
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
    Procedure BalGeneZoom(Quoi : String) ;
    Function  ColonnesOK : Boolean ;
    Procedure CalcTotalEdt(AnVD, AnVC, D, C : Double );
    procedure RecapPourQui(Cpt : String ; LBG : TStringList ; MO : Array of double) ;
    Procedure CalculGestion(MO : Array of double ) ;
    procedure EnableOngletComparatif;
  public  { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Uses Filtre,
     // GC - 20/12/2001
     SaisUtil,Choix,
{$IFNDEF IMP}
     Exercice,
{$ENDIF}
     uLibWindows;
     //

Procedure TFBalGen1.BalGeneZoom(Quoi : String) ;
BEGIN
ZoomEdt(1,Quoi) ;
END ;

procedure BalanceGene(VoirZoneEcart : Boolean = FALSE) ;
var QR : TFBalGen1 ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:= TFBalGen1.Create(Application) ;
Edition.Etat:=etBalGen ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalGeneZoom ;
QR.IsLegal:=FALSE ;
QR.NET:=neRien ;
QR.VoirZoneEcart:=VoirZoneEcart ;
If VoirZoneEcart Then QR.InitType (nbGen,neBal,msGenEcr,'QRBALECCGEN','',TRUE,FALSE,FALSE,Edition)
                 Else QR.InitType (nbGen,neBal,msGenEcr,'QRBALGEN','',TRUE,FALSE,FALSE,Edition) ;
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
procedure BalanceGeneComp;
var
  QR: TFBalGen1;
  Edition: TEdition;
  PP: THPanel;
begin
  PP := FindInsidePanel;
  QR := TFBalGen1.Create(Application);
  Edition.Etat := etBalGeneComp;

  QR.QRP.QRPrinter.OnSynZoom := QR.BalGeneZoom;
  QR.IsLegal := FALSE;
  QR.NET := neRien;
  QR.InitType(nbGen, neBal, msGenEcr, 'QRBALGEN', '', TRUE, FALSE, FALSE, Edition);

  if PP = nil then
  begin
    try
      QR.ShowModal;
    finally
      QR.Free;
    end;
    Screen.Cursor := SyncrDefault;
  end
  else
  begin
    InitInside(QR, PP);
    QR.Show;
  end;

end;

//******************************************************************************
// GC
procedure BalanceGeneChaine( vCritEdtChaine : TCritEdtChaine );
var
  QR: TFBalGen1;
  Edition: TEdition;
  PP: THPanel;
begin
  PP := FindInsidePanel;
  QR := TFBalGen1.Create(Application);
  Edition.Etat := etBalGeneComp;

  QR.QRP.QRPrinter.OnSynZoom := QR.BalGeneZoom;
  QR.IsLegal := FALSE;
  QR.NET := neRien;
  QR.CritEdtChaine := vCritEdtChaine;
  QR.InitType(nbGen, neBal, msGenEcr, 'QRBALGEN', '', TRUE, FALSE, FALSE, Edition);

  if PP = nil then
  begin
    try
      QR.Timer1.Enabled := True;
      QR.QRP.QRPrinter.Copies := QR.CritEdtChaine.NombreExemplaire;
      QR.ShowModal;
    finally
      QR.Free;
    end;
    Screen.Cursor := SyncrDefault;
  end
  else
  begin
    InitInside(QR, PP);
    QR.Show;
  end;

end;
// GC - 20/12/2001 - FIN

procedure BalanceGeneCERG ;
var QR : TFBalGen1 ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:= TFBalGen1.Create(Application) ;
Edition.Etat:=etBalGenCerg ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalGeneZoom ;
V_PGI.CergEF:=TRUE ;
QR.IsLegal:=FALSE ;
QR.NET:=neRien ;
QR.InitType (nbGen,neBal,msGenEcr,'QRBALCER','',TRUE,FALSE,FALSE,Edition) ;
if PP=Nil then
   BEGIN
   try
    QR.ShowModal ;
    finally
    QR.Free ;
    V_PGI.CergEF:=FALSE ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   V_PGI.CergEF:=FALSE ;
   END Else
   BEGIN
   InitInside(QR,PP) ;
   QR.Show ;
   END ;
END ;


procedure BalanceGeneZoom(Crit : TCritEdt) ;
var QR : TFBalGen1 ;
    Edition : TEdition ;
BEGIN
QR:=TFBalGen1.Create(Application) ;
Edition.Etat:=etBalGen ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.BalGeneZoom ;
 QR.CritEdt:=Crit ;
 QR.IsLegal:=FALSE ;
 QR.NET:=neRien ;
 QR.InitType (nbGen,neBal,msGenEcr,'QRBALGEN','',FALSE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure BalanceGeneLegal(Crit : TCritEdt) ;
var QR : TFBalGen1 ;
    Edition : TEdition ;
BEGIN
QR:=TFBalGen1.Create(Application) ;
Edition.Etat:=etBalGen ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.BalGeneZoom ;
 QR.CritEdt:=Crit ;
 QR.IsLegal:=TRUE ;
 QR.NET:=neRien ;
 QR.InitType (nbGen,neBal,msGenEcr,'QRBALGEN','',TRUE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure BalanceAno ;
var QR : TFBalGen1 ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:= TFBalGen1.Create(Application) ;
Edition.Etat:=etBalAno ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalGeneZoom ;
QR.IsLegal:=FALSE ;
QR.NET:=neAno ;
QR.InitType (nbGen,neBal,msGenEcr,'QRBALANO','',TRUE,FALSE,FALSE,Edition) ;
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

procedure BalanceClo ;
var QR : TFBalGen1 ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:= TFBalGen1.Create(Application) ;
Edition.Etat:=etBalClo ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalGeneZoom ;
QR.IsLegal:=FALSE ;
QR.NET:=neClo ;
QR.InitType (nbGen,neBal,msGenEcr,'QRBALCLO','',TRUE,FALSE,FALSE,Edition) ;
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

procedure TFBalGen1.GenereSQL ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
InHerited ;
If CritEdt.Bal.STTS<>'' Then GenereSQLBaseTS
                        Else GenereSQLBase ;
END ;

procedure TFBalGen1.RenseigneCritere;
BEGIN
InHerited ;
TTotBil.Caption:=MsgLabel.Mess[0] ; TTotGes.Caption:=MsgLabel.Mess[3] ;
if NET=neRien then Exit ;
QRLabel7.Visible:=False ; RNatureEcr.Visible:=False ;
QRLabel24.Visible:=False ; RExcepGen.Visible:=False ;
END;

procedure TFBalGen1.ChoixEdition ;
BEGIN
InHerited ;
if CritEdt.Bal.BilGes then ChargeRecap(LBilGes) ;
If CritEdt.Bal.STTS<>'' Then
  BEGIN
  DLRuptTRI.PrintBefore:=FALSE ;
  ChargeGroup(LRuptTRI,['TRI']) ;
  END  ;

DLRupt.PrintBefore:=FALSE ;
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
If CritEdt.Bal.OkRegroup Then ChargeRegroup(LRegroup,CritEdt.Bal.StRegroup) ;

END ;

procedure TFBalGen1.RecupCritEdt ;
Var NonLibres : Boolean ;
BEGIN
Inherited ;
With CritEDT Do
  BEGIN
  BAL.CodeRupt1:='' ; BAL.CodeRupt2:='' ;
  NonLibres:=((Rupture=rRuptures) or (Rupture=rCorresp)) ;
  if NonLibres then BAL.PlanRupt:=FPlanRuptures.Value ;
  BAL.RuptOnly:=QuelleTypeRupt(0,FSAnsRupt.Checked,FAvecRupt.Checked,FSurRupt.Checked) ;
  if NonLibres then BEGIN BAL.CodeRupt1:=FCodeRupt1.Text ; BAL.CodeRupt2:=FCodeRupt2.Text ; END ;
  If (CritEdt.Rupture=rCorresp) Then Bal.PlansCorresp:=StrToInt(Copy(FPlanRuptures.Value,3,1)) ;
  Bal.OnlyCptAssocie:=(Rupture<>rRien) and FOnlyCptAssocie.Checked ;
  Bal.QuelAN:=FQuelAN(FAvecAN.Checked) ;
  If Bal.QuelAN=SansAN Then FCol1.Checked:=FALSE ;
  Bal.BilGes:=FBilGes.Checked ;
  if NET=neClo then
     BEGIN
     Bal.PourCloture:=TRUE ;
     if Not FBilGes.Checked then
        BEGIN
        Bal.Sauf:='...' ; { Pour forcer dans QR l'utilisation de Bal.SQLSauf }
        Bal.SQLSauf:=Bal.SQLSauf+' and ((G_NATUREGENE<>"CHA") and (G_NATUREGENE<>"PRO")) ' ;
        END ;
     END ;
  Bal.SpeedOk:=FSpeed.checked ;
  Bal.VoirEcart:=0 ;
  If VoirZoneEcart Then
    BEGIN
    Case FSansEcart.State Of cbUnchecked : Bal.VoirEcart:=1 ; cbChecked : Bal.VoirEcart:=2 ; END ;
    END ;
  With BAL.FormatPrint Do
    BEGIN
    PrColl:=FCollectif.Checked ;
    PrSepRupt:=FLigneRupt.Checked ;
    PutTabCol(1,TabColl[1],TitreColCpt.Tag,TRUE) ;
    PutTabCol(2,TabColl[2],TitreColLibelle.Tag,TRUE) ;
    PutTabCol(3,TabColl[3],TTitreColAvant.Tag,FCol1.Checked) ;
    PutTabCol(4,TabColl[4],TTitreColSelection.Tag,FCol2.Checked) ;
    PutTabCol(5,TabColl[5],TTitreColApres.Tag,FCol3.Checked) ;
    PutTabCol(6,TabColl[6],TTitreColSolde.Tag,FCol4.Checked) ;
    END ;
  Bal.OkRegroup:=FOkRegroup.Checked ;
  If Bal.OkRegroup Then Bal.StRegroup:=FRegroup.Text ;
  END ;
  CritEdt.AvecComparatif := FAvecComparatif.Checked;
  CritEdt.BalSit:=FBALSIT.Text;
  CritEdt.CompareBalSit := RBCOMPBALSIT.Checked;
  CritEdt.ExoComparatif.Code := FExercice2.Value;
  if IsValidDate(FDateComparatif1.Text) then
    CritEdt.ExoComparatif.Deb := StrToDate(FDateComparatif1.Text);
  if IsValidDate(FDateComparatif2.Text) then
    CritEdt.ExoComparatif.Fin := StrToDate(FDateComparatif2.Text);
END ;

Function TFBalGen1.CritOk : Boolean ;
Var StEcart : String ;
// CA - 20/12/2001
    BalSit : string;
// CA - 20/12/2001 - FIN
BEGIN
Result:=Inherited CritOK ;
if (CritEdt.AvecComparatif) then
begin
  if CritEdt.CompareBalSit then
  begin
    Result := Presence('CBALSIT','BSI_CODEBAL',CritEdt.BalSit);
    if not Result then MessageAlerte('Balance de situation incorrecte.');
  end
  else
  begin
    Result := IsValidDate(FDateComparatif1.Text) and IsValidDate(FDateComparatif2.Text);
    if not Result then MessageAlerte('Dates du comparatif incorrectes.');
  end;
  if not Result then NumErreurCrit := 0;
end;
If Result And (Not OkZoomEdt) Then Result:=ColonnesOK;
If Result Then
   BEGIN
   StEcart:='' ;
   Case CritEdt.Bal.VoirEcart Of
     1 : StEcart:='((E_QUALIFORIGINE<>"ECC" AND E_QUALIFORIGINE<>"CON") OR (E_QUALIFORIGINE IS NULL))' ;
     2 : StEcart:='(E_QUALIFORIGINE="ECC" OR E_QUALIFORIGINE="CON")' ;
     END ;
   If StEcart='' Then StEcart:=CritEdt.FiltreSup ;
    // CA - 20/12/2001
    // Récupération du code de la balance de situation le cas échéant :
    if TabSup.TabVisible and RBCOMPBALSIT.Checked and (FBALSIT.Text<>'') then
      BalSit := FBALSIT.Text else BalSit :='';
    // CA - 20/12/2001 - FIN
//   QBal:=PrepareTotCpt(fbGene,QuelTypeEcr,Dev,Etab,Exo) ;
// rony 10/03/97 * QBal:=PrepareTotCptCum(QBalCum,fbGene,QuelTypeEcr,Dev,Etab,Exo) ;
   QBalC:=PrepareTotCptSolde(QuelTypeEcr,Dev,Etab,Exo,DevEnP) ;
   // Rony 10/07/97 Gcalc:=TGCalculCum.create(Un,fbGene,fbGene,QuelTypeEcr,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE) ;
   if NET=neClo then
      BEGIN
      // CA - 20/12/2001 - Ajout paramètre false et balsit
      if FOnlyClo.Checked then Gcalc:=TGCalculCum.create(Un,fbGene,fbGene,[tpCloture],Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE,False,BalSit)
                          Else Gcalc:=TGCalculCum.create(Un,fbGene,fbGene,[tpReel,tpCloture],Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE,False,BalSit)
      END Else Gcalc:=TGCalculCum.create(Un,fbGene,fbGene,QuelTypeEcr,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE,CritEdt.Bal.SpeedOk,BalSit,StEcart,CritEdt.Bal.STTS) ;
      // CA - 20/12/2001 - FIN
   GCalc.initCalcul('','','','',CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                    CritEdt.Date1,CritEdt.Date2,TRUE) ;
   GCalc.InitAnalyseCompte(CritEdt.Cpt1,CritEdt.Cpt2,'','') ;
   END ;
InitReport([1],CritEDT.BAL.FormatPrint.Report) ;
END ;

Function TFBalGen1.ColonnesOK : Boolean ;
BEGIN
If OkZoomEdt then begin Result:=True ; exit ; end ;
Result:=(FCol1.Checked or FCol2.Checked  or FCol3.Checked or FCol4.Checked) ;
If Not Result then NumErreurCrit:=7 ;
END ;

Procedure TFBalGen1.FinirPrint ;
BEGIN
Inherited ;
if CritEdt.Bal.BilGes then VideRecap(LBilGes) ;
if CritEdt.Rupture<>rRien then VideRupt(LRupt) ;
If CritEdt.Bal.STTS<>'' Then VideRupt(LRuptTRI) ;
If CritEdt.Bal.OkRegroup Then VideRupt(LRegroup) ;
{$IFNDEF CCMP}
If OkMajEdt And (Not V_PGI.CergEF) Then
   BEGIN
{$IFNDEF IMP}
   If IsLegal Then MajEditionLegal('BLG',CritEDT.Exo.Code,DateToStr(CritEDT.DateDeb),DateToStr(CritEDT.DateFin),
                              '',TotEdt[3].TotDebit,TotEdt[3].TotCredit,TotEdt[4].TotDebit,TotEdt[4].TotCredit)
              Else MajEdition('BLG',CritEDT.Exo.Code,DateToStr(CritEDT.DateDeb),DateToStr(CritEDT.DateFin),
                              '',TotEdt[3].TotDebit,TotEdt[3].TotCredit,TotEdt[4].TotDebit,TotEdt[4].TotCredit) ;
{$ENDIF}
   END ;
{$ENDIF}
If GCalc<>NIL Then BEGIN GCalc.Free ; GCalc:=NIL ; END ;
If QBalC<>NIL Then BEGIN QBalC.Free ; QBalC:=NIL ; END ;
END ;

procedure TFBalGen1.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var TotCpt1,TotCpt2,tot,
    TotCpt1M,TotCpt1S,
    TotCpt2M,TotCpt2S   : TabTot ;
    TotCpt1MEURO,TotCpt1SEURO,
    TotCpt2MEURO,TotCpt2SEURO   : TabTot ;
    Solde               : Double ;
    MReport             : TabTRep ;
    Sommation           : Boolean ;
    PourBil             : Array[0..12] of double ;
    CptRupt     : String ;
    VCS : Variant ;
  // GC - 20/12/2001
  procedure CalculpourComparatif;
  var
    SurBalSit: boolean;
  begin
    SurBalSit := RBCOMPBALSIT.Checked;
    if (Qr11G_COLLECTIF.AsString = 'X') and (CritEDT.BAL.FormatPrint.PrColl) then
    begin
      ExecuteTotCptSolde(QBalC, Qr11G_GENERAL.AsString, CritEDT.Date1, CritEDT.Date2, CritEDT.DeviseSelect, CritEDT.Etab, CritEDT.Exo.Code, TotCpt2M, Totcpt1, TotCpt2MEURO, TotCpt2SEURO, TRUE, Sommation, CritEdt.Decimale, V_PGI.OkDecE, CritEdt.Monnaie = 2,DevEnP,QuelTypeEcr);
      if SurBalSit then
      begin
        GCAlc.ReInitCalcul(Qr11G_GENERAL.AsString, '', CritEDT.ExoComparatif.Deb, CritEDT.ExoComparatif.Fin, CritEDT.ExoComparatif.Code, SurBalSit);
        GCalc.Calcul;
        TotCpt2 := GCalc.ExecCalc.TotCpt;
      end
      else
        ExecuteTotCptSolde(QBalC, Qr11G_GENERAL.AsString, CritEDT.ExoComparatif.Deb, CritEDT.ExoComparatif.Fin, CritEDT.DeviseSelect, CritEDT.Etab, CritEDT.ExoComparatif.Code, TotCpt2M, TotCpt2, TotCpt2MEURO, TotCpt2SEURO, TRUE, Sommation, CritEdt.Decimale, V_PGI.OkDecE, CritEdt.Monnaie = 2,DevEnP,QuelTypeEcr)
    end
    else
    begin
      GCAlc.ReInitCalcul(Qr11G_GENERAL.AsString, '', CritEDT.Date1, CritEDT.Date2, CritEdt.Exo.Code);
      GCalc.Calcul;
      TotCpt1 := GCalc.ExecCalc.TotCpt;

      GCAlc.ReInitCalcul(Qr11G_GENERAL.AsString, '', CritEDT.ExoComparatif.Deb, CritEDT.ExoComparatif.Fin, CritEDT.ExoComparatif.Code, SurBalSit);
      GCalc.Calcul;
      TotCpt2 := GCalc.ExecCalc.TotCpt;
    end;
  end;
  // GC - 20/12/2001 - FIN

begin
  inherited;
VCS:=RecupVCS(Qr11STTS) ;
Sommation:=TRUE ;
Fillchar(Tot,SizeOf(Tot),#0) ; Fillchar(MReport,SizeOf(MReport),#0) ;
Fillchar(PourBil,SizeOf(PourBil),#0) ;
Fillchar(TotCpt1,SizeOf(TotCpt1),#0) ;
Fillchar(TotCpt2,SizeOf(TotCpt2),#0) ;
Fillchar(TotCpt1M,SizeOf(TotCpt1M),#0) ; Fillchar(TotCpt1S,SizeOf(TotCpt1S),#0) ;
Fillchar(TotCpt2M,SizeOf(TotCpt2M),#0) ; Fillchar(TotCpt2S,SizeOf(TotCpt2S),#0) ;
Fillchar(TotCpt1M,SizeOf(TotCpt1MEURO),#0) ; Fillchar(TotCpt1S,SizeOf(TotCpt1SEURO),#0) ;
Fillchar(TotCpt2M,SizeOf(TotCpt2MEURO),#0) ; Fillchar(TotCpt2S,SizeOf(TotCpt2SEURO),#0) ;

Case CritEDT.BAL.TypCpt of
  0 :  PrintBand:=True ;
  3 :  PrintBand:= TRUE ;
 end ;

if PrintBand then
   BEGIN
   Case CritEdt.Rupture of
     rLibres    : if CritEdt.Bal.OnlyCptAssocie then PrintBand:=DansRuptLibre(Q,fbGene,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
     rRuptures  : if CritEdt.Bal.OnlyCptAssocie then PrintBand:=DansRupt(LRupt,Qr11G_GENERAL.AsString) ;
     rCorresp   : if CritEdt.Bal.OnlyCptAssocie then
                    if CritEDT.Bal.PlansCorresp=1 then PrintBand:=DansRuptCorresp(LRupt,Qr11G_CORRESP1.AsString) Else
                    if CritEDT.Bal.PlansCorresp=2 then PrintBand:=DansRuptCorresp(LRupt,Qr11G_CORRESP2.AsString) ;
//                    if CritEDT.Bal.PlansCorresp=1 then PrintBand:=(Qr11G_CORRESP1.AsString<>'') Else
//                    if CritEDT.Bal.PlansCorresp=2 then PrintBand:=(Qr11G_CORRESP2.AsString<>'') ;
     End;
   If Not PrintBand then Exit ;
   Quoi:=Qr11G_GENERAL.AsString+' '+Qr11G_LIBELLE.AsString ;
   if (Qr11G_COLLECTIF.AsString='X') And (CritEDT.BAL.FormatPrint.PrColl) Then
      BEGIN
      // GC - 20/12/2001
      if CritEdt.AvecComparatif then
        CalculpourComparatif
      else
      begin
      If CritEdt.Bal.QuelAN<>SansAN Then
         BEGIN
         ExecuteTotCptSolde(QBalC,Qr11G_GENERAL.AsString, CritEDT.Date1,CritEDT.Date1,
                             CritEDT.DeviseSelect,CritEDT.Etab,CritEDT.Exo.Code,TotCpt1M,TotCpt1S,TotCpt1MEURO,TotCpt1SEURO,FALSE,Sommation,CritEdt.Decimale,V_PGI.OkDecE,CritEdt.Monnaie=2,DevEnP,QuelTypeEcr) ;
         If VH^.UseTC Then GCAlc.ReInitCalculVCS(Qr11G_GENERAL.AsString,'',CritEDT.BAL.Date21,CritEDT.Date2,VCS)
                      Else GCAlc.ReInitCalculVCS(Qr11G_GENERAL.AsString,'',CritEDT.Date1,CritEDT.BAL.Date11,VCS) ;
         GCalc.Calcul ; TotCpt1:=GCalc.ExecCalc.TotCpt ;
         END ;

      GCAlc.ReInitCalculVCS(Qr11G_GENERAL.AsString,'',CritEDT.BAL.Date21,CritEDT.Date2,VCS) ;
      GCalc.Calcul ; TotCpt2:=GCalc.ExecCalc.TotCpt ;
      ExecuteTotCptSolde(QBalC,Qr11G_GENERAL.AsString, CritEDT.Date1,CritEDT.Date2, CritEDT.DeviseSelect,CritEDT.Etab,CritEDT.Exo.Code,TotCpt2M,TotCpt2S,TotCpt2MEURO,TotCpt2SEURO,TRUE,Sommation,CritEdt.Decimale,V_PGI.OkDecE,CritEdt.Monnaie=2,DevEnP,QuelTypeEcr) ;
      TotCpt1[0]:=TotCpt1S[0] ;
      end;
      // GC - FIN

      END Else
      BEGIN
      (*rony 10/03/97*
      (*ExecuteTotCptCum(QBal,QBalCum,QG_GENERAL.AsString,'',CritEDT.Date1,CritEDT.BAL.Date11,
                    CritEDT.DeviseSelect,CritEDT.Etab,CritEDT.Exo.Code,TotCpt1,Sommation) ;
      ExecuteTotCptCum(QBal,QBalCum,QG_GENERAL.AsString,'',CritEDT.BAL.Date21,CritEDT.Date2,
                    CritEDT.DeviseSelect,CritEDT.Etab,CritEDT.Exo.Code,TotCpt2,Sommation) ;
      CumulVersSolde(TotCpt1[0]) ;
      *)
      // GC - 20/12/2001
      if CritEdt.AvecComparatif then CalculpourComparatif
      else
      begin
        If CritEdt.Bal.QuelAN<>SansAN Then
           BEGIN
           If VH^.UseTC Then GCAlc.ReInitCalculVCS(Qr11G_GENERAL.AsString,'',CritEDT.BAL.Date21,CritEDT.Date2,VCS)
                        Else GCAlc.ReInitCalculVCS(Qr11G_GENERAL.AsString,'',CritEDT.Date1,CritEDT.BAL.Date11,VCS) ;
//         GCAlc.ReInitCalcul(Qr11G_GENERAL.AsString,'',CritEDT.Date1,CritEDT.BAL.Date11) ;
           GCalc.Calcul ; TotCpt1:=GCalc.ExecCalc.TotCpt ;
           END ;

        GCAlc.ReInitCalculVCS(Qr11G_GENERAL.AsString,'',CritEDT.BAL.Date21,CritEDT.Date2,VCS) ;
        GCalc.Calcul ; TotCpt2:=GCalc.ExecCalc.TotCpt ;
        CumulVersSolde(TotCpt1[0]) ;

      end;
      // GC - FIN

      END ;
      (*
      if CritEDT.Date1 = CritEDT.BAL.Date11 then Fillchar(TotCpt1[1], SizeOf(TotCpt1[1]), #0);
      if CritEdt.Bal.QuelAN = SansAN then
      begin
        TotCpt1[0].TotDebit := 0;
        TotCpt1[0].TotCredit := 0;
        TotCpt1[1].TotDebit := 0;
        TotCpt1[1].TotCredit := 0;
      end;
      *)
   // GC - 06/03/2002
    if not CritEdt.AvecComparatif then
    begin
      if CritEDT.Date1 = CritEDT.BAL.Date11 then
        Fillchar(TotCpt1[1], SizeOf(TotCpt1[1]), #0);
      if CritEdt.Bal.QuelAN = SansAN then
      begin
        TotCpt1[0].TotDebit  := 0;
        TotCpt1[0].TotCredit := 0;
        TotCpt1[1].TotDebit  := 0;
        TotCpt1[1].TotCredit := 0;
      end;
    end;

    // GC - 09/02/2002
    if CritEdt.AvecComparatif then
    begin
      if (Qr11G_COLLECTIF.AsString = 'X') and (CritEDT.BAL.FormatPrint.PrColl) then
      begin
        Tot[1].TotDebit  := Arrondi(TotCpt1[1].TotDebit,  CritEDT.Decimale);
        Tot[1].TotCredit := Arrondi(Totcpt1[1].TotCredit, CritEDT.Decimale);
        Tot[2].TotDebit  := Arrondi(TotCpt2[1].TotDebit,  CritEDT.Decimale);
        Tot[2].TotCredit := Arrondi(TotCpt2[1].TotCredit, CritEDT.Decimale);
      end
      else
      begin
        // Total du compte pour l' exercice 1
        CalculSolde(TotCpt1[0].TotDebit + TotCpt1[1].TotDebit, TotCpt1[0].TotCredit + TotCpt1[1].TotCredit, Tot[1].TotDebit, Tot[1].TotCredit);
        // Total du compte pour l' exercice 2
        CalculSolde(TotCpt2[0].TotDebit + TotCpt2[1].TotDebit, TotCpt2[0].TotCredit + TotCpt2[1].TotCredit, Tot[2].TotDebit, Tot[2].TotCredit);
      end;
    end
    else
    begin
      Tot[1].TotDebit  := Arrondi(TotCpt1[0].TotDebit + TotCpt1[1].TotDebit, CritEDT.Decimale);
      Tot[1].TotCredit := Arrondi(TotCpt1[0].TotCredit + TotCpt1[1].TotCredit, CritEDT.Decimale);
      Tot[2].TotDebit  := TotCpt2[1].TotDebit;
      Tot[2].TotCredit := TotCpt2[1].TotCredit;
    end;
    // GC - FIN

   If CritEdt.Bal.okRegroup Then
     BEGIN
     Case TypeRegroup(LRegroup,Qr11G_GENERAL.AsString) Of
       1 : BEGIN
           AddRegroup(LRegroup,Qr11G_GENERAL.AsString,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit]) ;
           PrintBand:=FALSE ;
           Exit ;
           END ;
       2 : BEGIN
           PrintRegroup(LRegroup,Qr11G_GENERAL.AsString,Tot) ;
           END ;
       END ;
     END ;

   Tot[3].TotDebit:= Arrondi(Tot[2].TotDebit+Tot[1].TotDebit,CritEDT.Decimale) ;
   Tot[3].TotCredit:=Arrondi(Tot[2].TotCredit+Tot[1].TotCredit,CritEDT.Decimale) ;


   if CritEdt.AvecComparatif then
   begin
     MReport[1].TotDebit  := Tot[1].TotDebit;
     MReport[1].TotCredit := Tot[1].TotCredit;
     MReport[2].TotDebit  := Tot[2].TotDebit;
     MReport[2].TotCredit := Tot[2].TotCredit;
   end
   else
   begin
     MReport[1].TotDebit  := Arrondi(TotCpt1[0].TotDebit+TotCpt1[1].TotDebit,CritEDT.Decimale) ;
     MReport[1].TotCredit := Arrondi(TotCpt1[0].TotCredit+TotCpt1[1].TotCredit,CritEDT.Decimale) ;
     MReport[2].TotDebit  := TotCpt2[1].TotDebit ;
     MReport[2].TotCredit := TotCpt2[1].TotCredit ;
   end;

   MReport[3].TotDebit:= Arrondi(Tot[2].TotDebit+Tot[1].TotDebit,CritEDT.Decimale) ;
   MReport[3].TotCredit:=Arrondi(Tot[2].TotCredit+Tot[1].TotCredit,CritEDT.Decimale) ;

   Case CritEdt.Rupture of
     rLibres   : AddGroupLibre(LRupt,Q,fbGene,CritEdt.LibreTrie,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,0]) ;
     rRuptures : AddRupt(LRupt,Qr11G_GENERAL.AsString,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,0]) ;
     rCorresp  : BEGIN
                 Case CritEDT.Bal.PlansCorresp Of
                   1 : If Qr11G_CORRESP1.AsString<>'' Then CptRupt:=Qr11G_CORRESP1.AsString+Qr11G_GENERAL.AsString
                                                      Else CptRupt:='.'+Qr11G_GENERAL.AsString ;
                   2 : If Qr11G_CORRESP2.AsString<>'' Then CptRupt:=Qr11G_CORRESP2.AsString+Qr11G_GENERAL.AsString
                                                      Else CptRupt:='.'+Qr11G_GENERAL.AsString ;
                   Else CptRupt:=Qr11G_GENERAL.AsString ;
                  End ;
                 AddRuptCorres(LRupt,CptRupt,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,0]) ;
                 END ;
     End ;
   If CritEdt.Bal.STTS<>'' Then AddGroup(LRuptTri,[Qr11STTS],[Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit]) ;

   if Qr11G_COLLECTIF.AsString='X' Then
      BEGIN
      If CritEDT.BAL.FormatPrint.PrColl Then Tot[4]:=TotCpt2S[1] Else
         BEGIN
         Tot[4]:=Tot[3] ;
         CumulVersSolde(Tot[4]) ;
         END ;
      MReport[4].TotCredit:=Tot[4].TotCredit ;
      MReport[4].TotDebit :=Tot[4].TotDebit ;
      PourBil[6]:=Tot[4].TotDebit ;
      PourBil[7]:=Tot[4].TotCredit ;
      SolDebit.Caption :=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[4].TotDebit , CritEDT.AfficheSymbole ) ;
      SolCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[4].TotCredit, CritEDT.AfficheSymbole) ;
      END Else
      BEGIN
      Tot[4]:=Tot[3] ;
      CumulVersSolde(Tot[4]) ;
      if Tot[4].TotDebit=0 then
         BEGIN
         MReport[4].TotCredit:=Tot[4].TotCredit ;
         PourBil[6]:=0 ;
         PourBil[7]:=Tot[4].TotCredit ;
         SolDebit.Caption:='' ;
         SolCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[4].TotCredit, CritEDT.AfficheSymbole) ;
         END else
         BEGIN
         MReport[4].TotDebit:=Tot[4].TotDebit ;
         PourBil[6]:=Tot[4].TotDebit ;
         PourBil[7]:=0 ;
         SolDebit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,Tot[4].TotDebit,CritEDT.AfficheSymbole ) ;
         SolCredit.Caption:='' ;
         END ;
      END ;
   Solde:=Tot[4].TotDebit-Tot[4].TotCredit ;
   if Solde>=0 then Quoi:=Quoi+'#'+MsgBox.Mess[5]+' '+
                    AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,Solde,CritEDT.AfficheSymbole )
               else Quoi:=Quoi+'#'+MsgBox.Mess[6]+' '+
                    AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,Abs(Solde),CritEDT.AfficheSymbole ) ;
   CumDebit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,Tot[1].TotDebit,CritEDT.AfficheSymbole ) ;
   CumCredit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,Tot[1].TotCredit,CritEDT.AfficheSymbole ) ;
   DetDebit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,Tot[2].TotDebit,CritEDT.AfficheSymbole ) ;
   DetCredit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,Tot[2].TotCredit,CritEDT.AfficheSymbole ) ;
   SumDebit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,Tot[3].TotDebit,CritEDT.AfficheSymbole ) ;
   SumCredit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,Tot[3].TotCredit,CritEDT.AfficheSymbole ) ;

    // GC - 09/01/2002
    SolCumDebit.Visible   := (CritEdt.AvecComparatif) and (Qr11G_COLLECTIF.AsString='X') And (CritEDT.BAL.FormatPrint.PrColl);
    SolCumCredit.Visible  := (CritEdt.AvecComparatif) and (Qr11G_COLLECTIF.AsString='X') And (CritEDT.BAL.FormatPrint.PrColl);
    SolDetCredit.Visible  := (CritEdt.AvecComparatif) and (Qr11G_COLLECTIF.AsString='X') And (CritEDT.BAL.FormatPrint.PrColl);
    SolDetDebit.Visible   := (CritEdt.AvecComparatif) and (Qr11G_COLLECTIF.AsString='X') And (CritEDT.BAL.FormatPrint.PrColl);
    // GC - FIN

   if (Qr11G_COLLECTIF.AsString='X') And CritEDT.BAL.FormatPrint.PrColl And (Not V_PGI.CergEF) then
      BEGIN
      CumulVersSolde(TotCpt2S[1]) ;
      BDetail.Height:=LibCollectif.Top+LibCollectif.Height ;
      LibCollectif.Visible:=True ; SolSolDebit.Visible:=True ; SolSolCredit.Visible:=True ;
      SolSolDebit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotCpt2S[1].TotDebit,CritEDT.AfficheSymbole ) ;
      SolSolCredit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotCpt2S[1].TotCredit,CritEDT.AfficheSymbole ) ;
        // GC - 09/01/2002
        if CritEdt.AvecComparatif then
        begin
          CumulVersSolde(Tot[1]);
          CumulVersSolde(Tot[2]);
          SolCumDebit.Caption   := AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[1].TotDebit, CritEDT.AfficheSymbole);
          SolCumCredit.Caption  := AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[1].TotCredit,CritEDT.AfficheSymbole);
          SolDetDebit.Caption   := AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[2].TotDebit, CritEDT.AfficheSymbole);
          SolDetCredit.Caption  := AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[2].TotCredit,CritEDT.AfficheSymbole);
        end;
        // GC - FIN

      END else
      BEGIN
      BDetail.Height:=LibCollectif.Top ;
      LibCollectif.Visible:=False ; SolSolDebit.Visible:=False ; SolSolCredit.Visible:=False ;
      END ;
   AddReportBAL([1],CritEDT.BAL.FormatPrint.Report,MReport,CritEDT.Decimale) ;
   PourBil[0]:=Tot[1].TotDebit ;
   PourBil[1]:=Tot[1].TotCredit ;
   PourBil[2]:=Tot[2].TotDebit ;
   PourBil[3]:=Tot[2].TotCredit ;
   PourBil[4]:=Tot[3].TotDebit ;
   PourBil[5]:=Tot[3].TotCredit ;
   if CritEdt.Bal.BilGes then RecapPourQui(Qr11G_GENERAL.AsString,LBilGes,PourBil) ;
//   PrintBand:=CritEdt.Bal.RuptOnly<>Sur ;
   If CritEdt.Bal.RuptOnly=Sur then
      BEGIN
      PrintBand:=False ;
      IF CritEdt.Rupture=rLibres then CalcTotalEdt(Tot[1].TotDebit, Tot[1].TotCredit, Tot[2].TotDebit, Tot[2].TotCredit) ;
      END Else CalcTotalEdt(Tot[1].TotDebit, Tot[1].TotCredit, Tot[2].TotDebit, Tot[2].TotCredit) ;

   END ;
end;

procedure TFBalGen1.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TotCumDebit.Caption :=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotEdt[1].TotDebit,CritEDT.AfficheSymbole ) ;
TotCumCredit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotEdt[1].TotCredit,CritEDT.AfficheSymbole ) ;
TotDetDebit.Caption :=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotEdt[2].TotDebit,CritEDT.AfficheSymbole ) ;
TotDetCredit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotEdt[2].TotCredit,CritEDT.AfficheSymbole ) ;
TotSumDebit.Caption :=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotEdt[3].TotDebit,CritEDT.AfficheSymbole ) ;
TotSumCredit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotEdt[3].TotCredit,CritEDT.AfficheSymbole ) ;
TotSolDebit.Caption :=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotEdt[4].TotDebit,CritEDT.AfficheSymbole ) ;
TotSolCredit.Caption:=AfficheMontant(CritEDT.FormatMontant,CritEDT.Symbole,TotEdt[4].TotCredit,CritEDT.AfficheSymbole ) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFBalGen1.DLRuptNeedData(var MoreData: Boolean);
Var TotRupt : Array[0..12] of Double ;
    SumD, SumC, TotSoldeD, TotSoldeC : Double ;
    CptRupt,Lib1, stcode    : String ;
    OkOk, AddTotEdt : Boolean ;
    QuelleRupt : Integer ;
    Col        : TColor ;
    LibRuptInf : Array[1..10] Of TRuptInf ;
begin
  inherited;
MoreData:=false ;
case CritEdt.Rupture of
  rLibres   : BEGIN
              MoreData:=PrintGroupLibre(LRupt,Q,fbGene,CritEdt.LibreTrie,CodRupt,LibRupt,Lib1,TotRupt,Quellerupt,Col,LibRuptInf,CritEdt.Bal.STTS) ;
              If MoreData then
                 BEGIN
                 StCode:=CodRupt ;
                 Delete(StCode,1,Quellerupt+2) ;
                 MoreData:=DansChoixCodeLibre(StCode,Q,fbGene,CritEdt.LibreCodes1,CritEdt.LibreCodes2, CritEdt.LibreTrie) ;
                 BRupt.Font.Color:=Col ;
                 END ;
              END ;
  rRuptures : MoreData:=PrintRupt2(LRupt,Qr11G_GENERAL.AsString,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt,[Qr11G_GENERAL],Q,CritEdt.Bal.STTS) ;
//  rRuptures : MoreData:=PrintRupt2(LRupt,Qr11G_GENERAL.AsString,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt,Q,[Qr11G_GENERAL]) ;
  rCorresp  : BEGIN
              OkOk:=TRUE ;
              Case CritEDT.Bal.PlansCorresp  Of
                1 : If Qr11G_CORRESP1.AsString<>'' Then CptRupt:=Qr11G_CORRESP1.AsString+Qr11G_GENERAL.AsString
                                                   Else CptRupt:='.'+Qr11G_GENERAL.AsString ;
                2 : If Qr11G_CORRESP2.AsString<>'' Then CptRupt:=Qr11G_CORRESP2.AsString+Qr11G_GENERAL.AsString
                                                   Else CptRupt:='.'+Qr11G_GENERAL.AsString ;
                Else OkOk:=FALSE ;
                End ;
              If OkOk Then MoreData:=PrintRupt(LRupt,CptRupt,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt) Else MoreData:=FALSE ;
              END ;
  End ;
if MoreData then
   BEGIN
   TCodRupt.Width:=TitreColCpt.Width+TitreColLibelle.Width+1 ;
   TCodRupt.Caption:='' ;
   BRupt.Height:=HauteurBandeRuptIni ;
   if CritEdt.Rupture=rLibres then
      BEGIN
      insert(MsgBox.Mess[14]+' ',CodRupt,Quellerupt+2) ;
      TCodRupt.Caption:=CodRupt+' '+Lib1 ;
      AlimRuptSup(HauteurBandeRuptIni,QuelleRupt,TCodRupt.Width,BRupt,LibRuptInf,Self) ;
      END Else TCodRupt.Caption:=CodRupt+'   '+LibRupt ;
   SumD:=TotRupt[1]+TotRupt[3] ;
   SumC:=TotRupt[2]+TotRupt[4] ;

   CalCulSolde(SumD,SumC,TotSoldeD,TotSoldeC) ;
   CumDebRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotRupt[1], CritEDT.AfficheSymbole) ;
   CumCreRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotRupt[2], CritEDT.AfficheSymbole) ;
   DetDebRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotRupt[3], CritEDT.AfficheSymbole) ;
   DetCreRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotRupt[4], CritEDT.AfficheSymbole) ;
   SumDebRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumD  , CritEDT.AfficheSymbole) ;
   SumCreRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumC  , CritEDT.AfficheSymbole) ;
   SolDebRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotSoldeD, CritEDT.AfficheSymbole) ;
   SolCreRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotSoldeC, CritEDT.AfficheSymbole) ;
   END ;
if CritEdt.Bal.RuptOnly=Sur then
   BEGIN
   AddTotEdt:=False ;
   if (CritEdt.Rupture=rLibres) then AddTotEdt:=False else
   if (CritEdt.Rupture=rRuptures) or (CritEdt.Rupture=rCorresp) then AddTotEdt:=DansTotal ;
   if AddTotEdt then CalcTotalEdt(TotRupt[1], TotRupt[2], TotRupt[3], TotRupt[4]) ;
   END ;

end;

Procedure TFBalGen1.CalcTotalEdt(AnVD, AnVC, D, C : Double );
Var SumD, SumC, TotSoldeD, TotSoldeC : Double ;
BEGIN
SumD:=Arrondi(AnVD+D, CritEdt.Decimale) ;
SumC:=Arrondi(AnVC+C, CritEdt.Decimale) ;
CalCulSolde(SumD,SumC,TotSoldeD,TotSoldeC) ;
TotEdt[1].TotDebit:= Arrondi(TotEdt[1].TotDebit  + AnVD, CritEdt.Decimale) ;
TotEdt[1].TotCredit:=Arrondi(TotEdt[1].TotCredit + AnVC, CritEdt.Decimale) ;
TotEdt[2].TotDebit:= Arrondi(TotEdt[2].TotDebit  + D, CritEdt.Decimale) ;
TotEdt[2].TotCredit:=Arrondi(TotEdt[2].TotCredit + C, CritEdt.Decimale) ;
TotEdt[3].TotDebit:= Arrondi(TotEdt[3].TotDebit  + SumD, CritEdt.Decimale) ;
TotEdt[3].TotCredit:=Arrondi(TotEdt[3].TotCredit + SumC, CritEdt.Decimale) ;
TotEdt[4].TotDebit:= Arrondi(TotEdt[4].TotDebit + TotSoldeD, CritEdt.Decimale) ;
TotEdt[4].TotCredit:=Arrondi(TotEdt[4].TotCredit + TotSoldeC, CritEdt.Decimale) ;
(*
If CritEdt.Bal.RuptOnly=Sur then
   BEGIN
   TotEdt[4].TotDebit:= Arrondi(TotEdt[4].TotDebit + TotSoldeD, CritEdt.Decimale) ;
   TotEdt[4].TotCredit:=Arrondi(TotEdt[4].TotCredit + TotSoldeC, CritEdt.Decimale) ;
   END Else
   BEGIN
   TotEdt[4].TotDebit:= Arrondi(TotEdt[4].TotDebit + TotSoldeD, CritEdt.Decimale) ;
   TotEdt[4].TotCredit:=Arrondi(TotEdt[4].TotCredit + TotSoldeC, CritEdt.Decimale) ;
   END ;
*)
END ;

procedure TFBalGen1.BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.Bal.RuptOnly<>Sans) ;
end;

procedure TFBalGen1.InitDivers ;
BEGIN
InHerited ;
Fillchar(TotGes,SizeOf(TotGes),#0) ;
OkTotGestion:=false ;
{ Initialise les tableaux de montants }
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
{ Labels sur les bandes }
  if CritEdt.CompareBalSit and  (CritEdt.BalSit<>'') then
     DateBandeauBalance(TTitreColAvant, TTitreColSelection, TTitreColApres, TTitreColSolde, CritEdt, GetColonneSQL('CBALSIT', 'BSI_ABREGE','BSI_CODEBAL="' + FBALSIT.Text + '"'), MsgBox.Mess[1], MsgBox.Mess[2])
  else
     // CA - 20/11/2001 - FIN
     DateBandeauBalance(TTitreColAvant,TTitreColSelection,TTitreColApres,TTitreColSolde,CritEdt, MsgBox.Mess[0],MsgBox.Mess[1],MsgBox.Mess[2]) ;
BRupt.Frame.DrawTop:=CritEDT.BAL.FormatPrint.PrSepRupt  ;
BRupt.Frame.DrawBottom:=BRupt.Frame.DrawTop ;
ChangeFormatCompte(fbGene,fbGene,Self,G_GENERAL.Width,1,2,G_GENERAL.Name) ;
Brupt.ForceNewPage:=FALSE ; BDetail.ForceNewPage:=FALSE ;
END ;

procedure TFBalGen1.TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
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

procedure TFBalGen1.BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var MReport : TabTRep ;
begin
  inherited;
Case QuelReportBAL(CritEDT.BAL.FormatPrint.Report,MReport) of
  1 : TITREREPORTB.Caption:=MsgBox.Mess[7] ;
 end ;
Report5Debit.Caption :=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, MReport[1].TotDebit,  CritEDT.AfficheSymbole ) ;
Report5Credit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, MReport[1].TotCredit, CritEDT.AfficheSymbole ) ;
Report6Debit.Caption :=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, MReport[2].TotDebit,  CritEDT.AfficheSymbole ) ;
Report6Credit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, MReport[2].TotCredit, CritEDT.AfficheSymbole ) ;
Report7Debit.Caption :=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, MReport[3].TotDebit,  CritEDT.AfficheSymbole ) ;
Report7Credit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, MReport[3].TotCredit, CritEDT.AfficheSymbole ) ;
Report8Debit.Caption :=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, MReport[4].TotDebit,  CritEDT.AfficheSymbole ) ;
Report8Credit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, MReport[4].TotCredit, CritEDT.AfficheSymbole ) ;
end;

procedure TFBalGen1.FormShow(Sender: TObject);
begin
HelpContext:=7445000 ;
//Standards.HelpContext:=7445010 ;
//Avances.HelpContext:=7445020 ;
//Mise.HelpContext:=7445030 ;
//Option.HelpContext:=7445040 ;
//TabRuptures.HelpContext:=7445050 ;
Case NET Of
  neAno : BEGIN
          HelpContext:=7760000 ;
          //Standards.HelpContext:=7761000 ;
          //Avances.HelpContext:=7762000 ;
          //Mise.HelpContext:=7763000;
          //Option.HelpContext:=7764000 ;
          END ;
  neClo : BEGIN
          HelpContext:=7769000 ;
          //Standards.HelpContext:=7769100 ;
          //Avances.HelpContext:=7769200 ;
          //Mise.HelpContext:=7769300 ;
          //Option.HelpContext:=7769400 ;
          END ;
  END ;
  IsLoading := True;
  inherited;
Floading:=FALSE ;
  // CA-GP - 20/12/2001
  // Initialisation comparatif
  if TRUE Or (CtxPcl in V_PGI.PGIContexte) then
  begin
    TabSup.TabVisible := FAvecComparatif.Checked;
    if FAvecComparatif.Checked then
      RBCompExo.OnClick( nil );
    EnableOngletComparatif;
  end
  else
    TabSup.TabVisible := False;


  TFLIBELLEBALSIT.Caption := '';
  CritEdt.CompareBalSit := RBCOMPBALSIT.Checked;
  if CritEdt.CompareBalSit then CritEdt.BalSit:= FBALSIT.Text;
  FBALSITChange(FBALSIT);
  // CA - 20/12/2001 - FIN

If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
FCollectif.Checked:=FALSE ; FLigneRupt.Checked:=FALSE ;
// RL 17/10/97 If V_PGI.CergEF Then FGroupRuptures.Enabled:=FALSE ;
FGroupChoixRupt.Visible:=Not V_PGI.CergEF ;
If Not V_PGI.CergEF Then FSansRuptClick(Nil); TFGenJoker.Visible:=False ;
FCol1.Checked:=TRUE ; FCol2.Checked:=TRUE ; FCol3.Checked:=TRUE ; FCol4.Checked:=TRUE ;
FJoker.MaxLength:=VH^.Cpta[fbGene].lg ;
If Not V_PGI.CergEF Then ChargeFiltrePourRupture ;
If V_PGI.CergEF Then
   BEGIN
   Apercu.Checked:=TRUE ; FListe.Checked:=TRUE ; FMonetaire.Enabled:=FALSE ; FDevises.Enabled:=FALSE ;
   GroupBox3.Enabled:=FALSE ;
   Fliste.Enabled:=FALSE ; Apercu.Enabled:=FALSE ; FCol3.Checked:=FALSE ;
   FCollectif.Checked:=TRUE ;
   Self.Caption:=MsgBox.Mess[8] ;
   END ;
FOnlyCptAssocie.Checked:=False ;
FOnlyClo.Visible:=False ; FBilGes.Enabled:=true ;

// GC - 12/12/2001
if CritEdtChaine.Utiliser then InitEtatchaine('QRBALGEN');

If VoirZoneEcart Then
  BEGIN
  FSansEcart.Visible:=TRUE ;
  If VH^.TenueEuro Then FMontant.ItemIndex:=0 Else FMontant.ItemIndex:=2 ;
  END ;

IsLoading := False;
if NET=neRien then Exit ;
TFGen.Enabled:=False ; TFGenJoker.Enabled:=False ; FJoker.Enabled:=False ;
FCpte1.Enabled:=False ; FCpte2.Enabled:=False ;
HLabel2.Enabled:=False ; FNatureCpt.Enabled:=False ;
FSelectCpte.Value:='EXO' ;
TSelectCpte.Enabled:=False  ; FSelectCpte.Enabled:=False ;
{ Zone Dates (de, à) }
HLabel6.Enabled:=False ; Label7.Enabled:=False ;
FDateCompta1.Enabled:=False ; FDateCompta2.Enabled:=False ;
HLabel5.Visible:=False ; FNatureEcr.Visible:=False ;
HLabel1.Visible:=False ; FExcep.Visible:=False ;
{ Zone Ruptures }
FGroupChoixRupt.enabled:=False ;
FSansRupt.Enabled:=False ; FAvecRupt.Enabled:=False ; FSurRupt.Enabled:=False ;

AvecRevision.Visible:=False ; //OnCum.Visible:=False ;
Case NET of
 neAno : BEGIN
         FCol1.Checked:=True ; FCol1.Enabled:=False ;
         FCol2.Checked:=False ; FCol2.Enabled:=False ; FCol3.Checked:=False ; FCol3.Enabled:=False ;
         FCol4.Checked:=False ; FCol4.Enabled:=False ;
         Caption:=MsgBox.Mess[12] ;
         END ;
 neClo : BEGIN
         Caption:=MsgBox.Mess[13] ;
         FOnlyClo.Visible:=True ; FBilGes.Visible:=True ;
         FExercice.Value:=VH^.Precedent.Code ;
         END ;
   End ;
FCollectif.Checked:=False ; FCollectif.Enabled:=False ;
Avance.Checked:=False ; Avance.Visible:=False ;
FOnlyCptAssocie.Checked:=False ; FOnlyCptAssocie.Visible:=False ;
FGroupRuptures.Enabled:=False ;
QRP.ReportTitle:=Caption ;
UpdateCaption(Self) ;
end;

procedure TFBalGen1.BBilanGenBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
PrintBand:=False ;
end;

procedure TFBalGen1.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr11G_GENERAL         :=TStringField(Q.FindField('G_GENERAL'));
Qr11G_LIBELLE         :=TStringField(Q.FindField('G_LIBELLE'));
Qr11G_COLLECTIF       :=TStringField(Q.FindField('G_COLLECTIF'));
If CritEDT.Rupture=rCorresp then
   BEGIN
   Qr11G_CORRESP1         :=TStringField(Q.FindField('G_CORRESP1'));
   Qr11G_CORRESP2         :=TStringField(Q.FindField('G_CORRESP2'));
   END ;
If CritEdt.Bal.STTS<>'' Then
  BEGIN
  Qr11STTS:=TVariantField(Q.FindField(CritEdt.Bal.STTS)) ;
  END ;
end;

procedure TFBalGen1.FExerciceChange(Sender: TObject);
begin
FLoading:=TRUE ;
  inherited;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
FLoading:=FALSE ;
end;

procedure TFBalGen1.FDateCompta1Change(Sender: TObject);
begin
  inherited;
If FLoading Then Exit ;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
end;

procedure TFBalGen1.FSelectCpteChange(Sender: TObject);
begin
  inherited;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
end;

procedure TFBalGen1.FAvecANClick(Sender: TObject);
begin
  inherited;
FCol1.Checked:=FAvecAN.Checked ;
end;

procedure TFBalGen1.FRupturesClick(Sender: TObject);
begin
  inherited;
If V_PGI.CergEF Then Exit ;
if not FRuptures.Checked then FBilGes.Checked:=False ;
FBilGes.Enabled:=FRuptures.Checked ;
if FPlansCo.Checked then FGroupRuptures.Caption:=' '+MsgBox.Mess[10] ;
if FRuptures.Checked Then FGroupRuptures.Caption:=' '+MsgBox.Mess[9] ;
end;

procedure TFBalGen1.FSansRuptClick(Sender: TObject);
begin
  inherited;
FLigneRupt.Enabled:=Not FSansRupt.Checked ;
FLigneRupt.checked:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
FRupturesClick(Nil) ;
if FSansRupt.Checked then FBilGes.Enabled:=True ;
end;

Procedure TFBalGen1.CalculGestion(MO : Array of double ) ;
BEGIN
TotGes[1].TotDebit:=Arrondi(TotGes[1].TotDebit+MO[0],CritEDT.Decimale) ;
TotGes[1].TotCredit:=Arrondi(TotGes[1].TotCredit+MO[1],CritEDT.Decimale) ;
TotGes[2].TotDebit:=Arrondi(TotGes[2].TotDebit+MO[2],CritEDT.Decimale) ;
TotGes[2].TotCredit:=Arrondi(TotGes[2].TotCredit+MO[3],CritEDT.Decimale) ;
TotGes[3].TotDebit:=Arrondi(TotGes[3].TotDebit+MO[4],CritEDT.Decimale) ;
TotGes[3].TotCredit:=Arrondi(TotGes[3].TotCredit+MO[5],CritEDT.Decimale) ;
TotGes[4].TotDebit:=Arrondi(TotGes[4].TotDebit+MO[6],CritEDT.Decimale) ;
TotGes[4].TotCredit:=Arrondi(TotGes[4].TotCredit+MO[7],CritEDT.Decimale) ;
END ;

procedure TFBalGen1.RecapPourQui(Cpt : String ; LBG : TStringList ; MO : Array of double) ;
Var I      : Integer ;
BEGIN
For i:=1 to high(VH^.FBIL) do
    if (Cpt>=VH^.FBIL[i].Deb) And (Cpt<=VH^.FBIL[i].Fin) then BEGIN AddRecap(LBG,['LB'],Cpt,MO) ; Break ; Exit END ;
For i:=1 to high(VH^.FCHA) do
    if (Cpt>=VH^.FCHA[i].Deb) And (Cpt<=VH^.FCHA[i].Fin) then BEGIN AddRecap(LBG,['LC'],Cpt,MO) ; Break ; Exit END ;
For i:=1 to high(VH^.FPRO) do
    if (Cpt>=VH^.FPRO[i].Deb) And (Cpt<=VH^.FPRO[i].Fin) then BEGIN AddRecap(LBG,['LP'],Cpt,MO) ; Break ; Exit END ;
For i:=1 to high(VH^.FExt) do
    if (Cpt>=VH^.FExt[i].Deb) And (Cpt<=VH^.FExt[i].Fin) then BEGIN AddRecap(LBG,['LE'],Cpt,MO) ; Break ; Exit END ;
END ;

procedure TFBalGen1.DLBilNeedData(var MoreData: Boolean);
Var Tot : Array[0..12] of Double ;
    Cod,Lib, Libelle : String ;
begin
  inherited;
MoreData:=False ;
if (Not CritEdt.Bal.BilGes) Or OkZoomEdt then Exit ;
MoreData:=PrintRecap(LBilGes,Cod,Lib,Tot) ;
if Cod='LB' then Libelle:=MsgBox.Mess[17] Else
if Cod='LC' then Libelle:=MsgBox.Mess[18] Else
if Cod='LP' then Libelle:=MsgBox.Mess[19] Else
if Cod='LE' then Libelle:=MsgBox.Mess[20]  ;
If MoreData Then
   BEGIN
   TTotBil.Caption:=Libelle ;
   if (Cod='LC') or (Cod='LP') then
      BEGIN OkTotGestion:=True ; CalculGestion(Tot) ; END ;

   TotBilCumDebit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[0] , CritEDT.AfficheSymbole ) ;
   TotBilCumCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[1] , CritEDT.AfficheSymbole ) ;

   TotBilDetDebit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[2] , CritEDT.AfficheSymbole ) ;
   TotBilDetCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[3] , CritEDT.AfficheSymbole ) ;

   TotBilSumDebit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[4] , CritEDT.AfficheSymbole ) ;
   TotBilSumCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[5] , CritEDT.AfficheSymbole ) ;

   TotBilSolDebit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[6] , CritEDT.AfficheSymbole ) ;
   TotBilSolCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, Tot[7] , CritEDT.AfficheSymbole ) ;
   END ;
BBil.Enabled:=MoreData ;
end;

procedure TFBalGen1.BGesBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
  // GC - 20/12/2001
  if CritEdt.AvecComparatif then
  begin
    CumulVersSolde( TotGes[1] );
    CumulVersSolde( TotGes[2] );
  end
  else
  begin
  // GC - 20/12/2001 - FIN
TotGes[4]:=TotGes[3] ;
CumulVersSolde(TotGes[4]) ;
  // GC - 20/12/2001
  end;
  // GC - 20/12/2001 - FIN
  // GC - 20/12/2001
  if CritEdt.AvecComparatif then
  begin
    if TotGes[1].TotDebit = 0 then
    begin
      TotGesCumDebit.Caption  := '' ;
      TotGesCumCredit.Caption := AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGes[1].TotCredit, CritEDT.AfficheSymbole);
    end
    else
    begin
      TotGesCumDebit.Caption  := AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGes[1].TotDebit, CritEDT.AfficheSymbole);
      TotGesCumCredit.Caption := '';
    end;

    if TotGes[2].TotDebit = 0 then
    begin
      TotGesDetDebit.Caption  := '' ;
      TotGesDetCredit.Caption := AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGes[2].TotCredit, CritEDT.AfficheSymbole);
    end
    else
    begin
      TotGesDetDebit.Caption  := AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGes[2].TotDebit, CritEDT.AfficheSymbole);
      TotGesDetCredit.Caption := '';
    end;

    TotGesSumDebit.Caption  := '' ;
    TotGesSumCredit.Caption := '' ;
    TotGesSolDebit.Caption  := '' ;
    TotGesSolCredit.Caption := '' ;
  end
  else
  begin
  // GC - 20/12/2001 - FIN
if TotGes[4].TotDebit=0 then
   BEGIN
   TotGesSolDebit.Caption:='' ;
   TotGesSolCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGes[4].TotCredit , CritEDT.AfficheSymbole ) ;
   END Else
   BEGIN
   TotGesSolDebit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGes[4].TotDebit, CritEDT.AfficheSymbole ) ;
   TotGesSolCredit.Caption:='' ;
   END ;
TotGesCumDebit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGes[1].TotDebit, CritEDT.AfficheSymbole ) ;
TotGesCumCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGes[1].TotCredit , CritEDT.AfficheSymbole ) ;
TotGesDetDebit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGes[2].TotDebit, CritEDT.AfficheSymbole ) ;
TotGesDetCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGes[2].TotCredit , CritEDT.AfficheSymbole ) ;
TotGesSumDebit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGes[3].TotDebit, CritEDT.AfficheSymbole ) ;
TotGesSumCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotGes[3].TotCredit , CritEDT.AfficheSymbole ) ;
// GC - 20/12/2001
end;
// GC - 20/12/2001 - FIN
end;

procedure TFBalGen1.DLGesNeedData(var MoreData: Boolean);
begin
  inherited;
MoreData:=OkTotGestion ;
OkTotGestion:=False ;
end;

Function GetBalanceSimple(Typ,Cod1 : String ; AvecAno,Etab,Devi,Exo : String ;
                          var Dat11,Dat22 : TDateTime ; Collectif : Boolean) : String ;
Var St1,StSup,TypRub,LAxe,LAxe1,LAxe2,Cpt1,Cpt1Ex : String ;
    PreE,PreJ,PreB : String ;
    Crit : TCritEdt ;
    QRub : TQuery ;
    Lefb1,Lefb2 : TFichierBase ;
    Multiple : TTypeCalc ;
BEGIN
{$IFNDEF IMP}
{V_PGI.OnCumEdt:=TRUE ;} Result:='' ; StSup:='' ;
Fillchar(Crit,SizeOf(Crit),#0) ;
If Typ='RUB' Then
   BEGIN
   QRub:=OpenSql('Select * From Rubrique Where RB_RUBRIQUE="'+Cod1+'"',True) ;
   if QRub.Eof then BEGIN Ferme(QRub) ; Exit ; END ;
   TypRub:=QRub.FindField('RB_TYPERUB').AsString ;
   Laxe:=QRub.FindField('RB_AXE').AsString ;
   Cpt1:=QRub.FindField('RB_COMPTE1').AsString ;
   Cpt1Ex:=QRub.FindField('RB_EXCLUSION1').AsString ;
   Ferme(QRub) ;
   if(TypRub='A/B') or (TypRub='A/G') or (TypRub='B/A') or
     (TypRub='G/A') or (TypRub='G/T') or (TypRub='T/G') then Multiple:=Deux
                                                        else Multiple:=Un ;
   If Multiple=Deux Then Exit ;
   QuelFBRub(Lefb1,Lefb2,TypRub,Laxe,Laxe1,Laxe2,Multiple,TRUE,FALSE) ;
//   St:=AvecAno ; QuelTyp(St,SetTyp,AvecAno) ;
   St1:=AnalyseCompte(Cpt1,Lefb1,False,FALSE) ;
   if St1<>'' then StSup:='( '+St1+' )' ;
   St1:=AnalyseCompte(Cpt1Ex,Lefb1,True,FALSE) ;
   if St1<>'' then StSup:=StSup+' And ( '+St1+' )' ;
   if StSup<>'' Then
      BEGIN
      PREEJB(Lefb1,PreE,PreJ,PreB) ;
      StSup:=StSup+' And '+WhatExiste(PreE,PreJ,PreB,AvecANO,False,cbUnchecked,Dat11,Dat22,Exo,1,'') ;
      END ;
   END Else
   BEGIN
   If Typ='TIE' Then StSup:=Cod1  Else If Typ='ANA' then Else Exit ;
   END ;
Crit.Date1:=Dat11 ; Crit.Date2:=Dat22 ;
Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
Crit.NatureEtat:=neBal ;
InitCritEdt(Crit) ;
Crit.Bal.ZoomBalRub:=TRUE ;
Crit.SQLPLUSBase:=StSup ; Crit.Etab:=Etab ; Crit.DeviseSelect:=Devi ;
Crit.BAL.FormatPrint.PrColl:=Collectif ;
BalanceGeneZoom(Crit) ;
{$ENDIF}
END ;

Function  Get_Balance(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif : String) : Variant ;
Var Err : Integer ;
    DD1,DD2 : TDateTime ;
    Exo : TExoDate ;
    V : Variant ;
BEGIN
{$IFNDEF IMP}
if (Table='RUBRIQUE') Then Table:='RUB' Else Exit ;
V:=varEmpty ;
If (AvecAno='') Or (AvecAno='X') Then AvecAno:='N' ;
If WhatDate(SDate,DD1,DD2,Err,Exo)
   Then V:=GetBalanceSimple(Table,Qui,AvecAno,Etab,Devi,Exo.Code,DD1,DD2,Collectif='X')
   Else V:=Err ;
Result:=V ;
{$ENDIF}
END ;

procedure TFBalGen1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then V_PGI.CergEF:=FALSE ;
  inherited;
end;

// GC - 20/12/2001
procedure TFBalGen1.FAvecComparatifClick(Sender: TObject);
begin
  inherited;
  FCol1.Enabled := not FAvecComparatif.Checked;
  FCol2.Enabled := not FAvecComparatif.Checked;
  FCol3.Enabled := not FAvecComparatif.Checked;
  FCol4.Enabled := not FAvecComparatif.Checked;

  if IsLoading then Exit;
  { Autres critères de la balance comparative }
  FCol1.Checked := FAvecComparatif.Checked;
  FCol2.Checked := FAvecComparatif.Checked;
  FCol3.Checked := not FAvecComparatif.Checked;
  FCol4.Checked := not FAvecComparatif.Checked;
  // FIN - GC

  { Composant de l'onglet Comparatif }
  TabSup.TabVisible := FAvecComparatif.Checked ;

  FExercice2.Enabled       := FAvecComparatif.Checked;
  FDateComparatif1.Enabled := FAvecComparatif.Checked;
  FDateComparatif2.Enabled := FAvecComparatif.Checked;

  { Recherche de l' existence de l'exercice N-1 }
  if VH^.Precedent.Code <> '' then
    FExercice2.Value := VH^.Precedent.Code
  else
    FExercice2.Value := VH^.Encours.Code;
  EnableOngletComparatif;
  PAvances.tabVisible:=FAvecComparatif.Checked ;
  If Not PAvances.tabVisible Then BEffaceAvanceClick(Nil) ;
end;

// GC - 20/12/2001
procedure TFBalGen1.FExercice2Change(Sender: TObject);
var D1, D2 : TDateTime;
begin
  inherited;
  DExoToDates(FExercice2.Value, D1, D2);
  FDateComparatif1.Text := DateToStr(D1);
  FDateComparatif2.Text := DateToStr(D2);
end;

procedure TFBalGen1.OnTypeComparatifClick(Sender: TObject);
begin
  inherited;

  FExercice2.Color       := IIF(not RBCompBalsit.Checked, clWhite, clBtnFace);
  FDateComparatif1.Color := IIF(not RBCompBalsit.Checked, clWhite, clBtnFace);
  FDateComparatif2.Color := IIF(not RBCompBalsit.Checked, clWhite, clBtnFace);
  FBalSit.Color          := IIF(RBCompBalSit.Checked, clWhite, clBtnFace);

  GBComparatif.Enabled := RBCOMPEXO.Checked;
  GBComparatif2.Enabled := RBCOMPBALSIT.Checked and VH^.TenueEuro and (not (FMontant.ItemIndex=2));
  CritEdt.CompareBalSit := RBCOMPBALSIT.Checked;
  if CritEdt.CompareBalSit then CritEdt.BalSit := FBALSIT.Text;
end;

procedure TFBalGen1.FBALSITChange(Sender: TObject);
begin
  inherited;
  TFLIBELLEBALSIT.Caption := '';
  if FBALSIT.Text = '' then exit;
  TFLIBELLEBALSIT.Caption := GetColonneSQL ('CBALSIT','BSI_LIBELLE','BSI_CODEBAL="'+FBALSIT.Text+'"');
  if CritEdt.CompareBalSit then CritEdt.BalSit := FBALSIT.Text;
end;

procedure TFBalGen1.FBALSITElipsisClick(Sender: TObject);
begin
  inherited;
{$IFNDEF IMP}
  LookUpBalSit(THCritMaskEdit(Sender),'GEN', '', '' ) ;
  if CritEdt.CompareBalSit then CritEdt.BalSit := FBALSIT.Text;
{$ENDIF}
end;

procedure TFBalGen1.FMontantClick(Sender: TObject);
begin
  inherited;
  EnableOngletComparatif;
end;

procedure TFBalGen1.EnableOngletComparatif;
begin
  GBComparatif2.Enabled := VH^.TenueEuro and (FMontant.ItemIndex<>2);
  FBalSit.Enabled := VH^.TenueEuro and (FMontant.ItemIndex<>2);
  RBCOMPBALSIT.Enabled := VH^.TenueEuro and (FMontant.ItemIndex<>2);
  RBCompBalSit.Checked := VH^.TenueEuro and (FMontant.ItemIndex<>2);
  if not RBCompBalSit.Checked then FBalsit.Text := '';
end;

procedure TFBalGen1.DLRUPTTRINeedData(var MoreData: Boolean);
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

procedure TFBalGen1.BRuptTRIBeforePrint(var PrintBand: Boolean;
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

procedure TFBalGen1.BDetailAfterPrint(BandPrinted: Boolean);
begin
  inherited;
BDetail.ForceNewPage:=FALSE ;
end;

procedure TFBalGen1.BRuptAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Brupt.ForceNewPage:=FALSE ;
end;

procedure TFBalGen1.FOkRegroupClick(Sender: TObject);
Var QQ : TQuery ;
begin
  inherited;
FRegroup.Enabled:=FOkRegroup.Checked ;
If Not FRegroup.Enabled Then FRegroup.Text:='' Else
  BEGIN
  If Trim(FRegroup.Text)='' Then
    BEGIN
    QQ:=OpenSQL('SELECT CR_CORRESP FROM CORRESP WHERE CR_TYPE="ZG1" ',TRUE) ;
    While Not QQ.Eof Do BEGIN FRegroup.Text:=FRegroup.Text+QQ.FindField('CR_CORRESP').AsString+';' ; QQ.Next ; END ;
    Ferme(QQ) ;
    END ;
  END ;

end;

procedure TFBalGen1.FRegroupDblClick(Sender: TObject);
Var ST : String ;
begin
  inherited;
St:=FRegroup.Text ;
FRegroup.Text:=ChoisirMulti('Compte de regroupement','CORRESP','CR_CORRESP','',' CR_TYPE="ZG1" ','CR_CORRESP',St) ;
end;

end.
