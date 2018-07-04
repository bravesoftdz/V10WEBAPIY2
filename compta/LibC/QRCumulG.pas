unit QRCumulG;
{JP 22/06/05 : Remplacé en CWAS par CPCUMULPERIODIQUE_TOF}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, CritEdt,
  QR, HSysMenu, Menus, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF} StdCtrls, Buttons,
  Hctrls, ExtCtrls, Mask, Hcompte, ComCtrls,UtilEdt, HEnt1, Ent1,CpteUtil,
  EdtLegal,
  QRRupt, HTB97, HPanel, UiUtil, tCalcCum ;

procedure CumulPeriodique(Lefb : TFichierBase ; NET : TNatureEtatJal) ;
// GC - 20/12/2001
procedure CumulPeriodiqueChaine ( vCritEdtChaine : TCritEdtChaine ) ;
// GC - 20/12/2001 - FIN
procedure CumulPeriodiqueZoom(Crit : TCritEdt ; Lefb : TFichierBase ) ;

type
  TFQRCumulG = class(TFQR)
    FAfficheDevise: TCheckBox;
    FSautDePage: TCheckBox;
    TitreColCpt: TQRLabel;
    TitreColMonFE: TQRLabel;
    TCREDIT: TQRLabel;
    TSOLDE: TQRLabel;
    TDEBITDEV: TQRLabel;
    TitreColMonDev: TQRLabel;
    TCREDITDEV: TQRLabel;
    TSOLDEDEV: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    REPORT1SOLDE: TQRLabel;
    TitreAN: TQRLabel;
    ANDEBIT: TQRLabel;
    ANCREDIT: TQRLabel;
    ANSOLDE: TQRLabel;
    ANDEBITDEV: TQRLabel;
    ANCREDITDEV: TQRLabel;
    ANSOLDEDEV: TQRLabel;
    BCumul: TQRBand;
    SOLDE: TQRLabel;
    SOLDEDEV: TQRLabel;
    DEBIT: TQRLabel;
    CREDIT: TQRLabel;
    DEBITDEV: TQRLabel;
    CREDITDEV: TQRLabel;
    BGeneFooter: TQRBand;
    TitreTotal: TQRLabel;
    TOT1SOLDE: TQRLabel;
    TOT1SOLDEDEV: TQRLabel;
    TOT1DEBIT: TQRLabel;
    TOT1CREDIT: TQRLabel;
    TOT1DEBITDEV: TQRLabel;
    TOT1CREDITDEV: TQRLabel;
    TTotalGen: TQRLabel;
    TOT2DEBIT: TQRLabel;
    TOT2CREDIT: TQRLabel;
    TOT2SOLDE: TQRLabel;
    TOT2DEBITDEV: TQRLabel;
    TOT2CREDITDEV: TQRLabel;
    TOT2SOLDEDEV: TQRLabel;
    QRBand1: TQRBand;
    Trait4: TQRLigne;
    Trait3: TQRLigne;
    Trait2: TQRLigne;
    Trait0: TQRLigne;
    Ligne1: TQRLigne;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    REPORT2SOLDE: TQRLabel;
    MsgLibel: THMsgBox;
    FormateurPivot: THNumEdit;
    TitreColLibelle: TQRLabel;
    QRDLMois: TQRDetailLink;
    LaDate: TQRLabel;
    FDate1: THValComboBox;
    FDate2: THValComboBox;
    TCompte: TQRLabel;
    FSoldeProg: TCheckBox;
    FTrait5: TQRLigne;
    FTrait6: TQRLigne;
    REPORT4DEBIT: TQRLabel;
    REPORT4CREDIT: TQRLabel;
    REPORT4SOLDE: TQRLabel;
    REPORT3DEBIT: TQRLabel;
    REPORT3CREDIT: TQRLabel;
    REPORT3SOLDE: TQRLabel;
    HLabel3: THLabel;
    TFDateCpta2: TLabel;
    TSOLDE2: TQRLabel;
    TSOLDEDEV2: TQRLabel;
    FAvecJalAN: TCheckBox;
    QRLabel9: TQRLabel;
    RAN: TQRLabel;
    FDateCpta2: THValComboBox;
    FDateCpta1: THValComboBox;
    BTotPeriode: TQRBand;
    LibRecap: TQRLabel;
    MDEBIT: TQRLabel;
    MCREDIT: TQRLabel;
    TDEBIT: TQRLabel;
    QRDLRecap: TQRDetailLink;
    MDEBITDEV: TQRLabel;
    MCREDITDEV: TQRLabel;
    MSOLDE: TQRLabel;
    MSOLDEDEV: TQRLabel;
    RecapMens: TCheckBox;
    QRSUBTOT: TQRBand;
    QRLabel2: TQRLabel;
    TOT22DEBIT: TQRLabel;
    TOT22CREDIT: TQRLabel;
    TOT22SOLDE: TQRLabel;
    TOT22DEBITDEV: TQRLabel;
    TOT22CREDITDEV: TQRLabel;
    TOT22SOLDEDEV: TQRLabel;
    procedure FExerciceChange(Sender: TObject);
    procedure BGeneFooterBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BCumulBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure QRDLMoisNeedData(var MoreData: Boolean);
    procedure FDateCpta1Change(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BGeneFooterAfterPrint(BandPrinted: Boolean);
    procedure QRDLRecapNeedData(var MoreData: Boolean);
    procedure QRSUBTOTBeforePrint(var PrintBand: Boolean;
      var Quoi: string);
    procedure FMontantClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
  private
    { Déclarations privées }
    StReportGen : string ;
    TotEdt, TotProg, TotCptGene  : TabTot ;
    Qr1COMPTE                    : TStringField ;
    Qr1LIBELLE                   : TStringField ;
    Qr1SAUTPAGE                  : TStringField ;
    Qr1NATURE                    : TStringField ;
    Qr1SOLDEPROGRESSIF           : TStringField ;
    GCalcDev                     : TGCalculCum ;
    LaDate1,LaDate2              : TDateTime ;
    fbEdt                        : tfichierBase ;
    NETJ                         : TNatureEtatJal ;
    LRecap                       : TStringList ;
    Procedure TiersZoom(Quoi : String) ;
    procedure GenPerExo ;
    procedure ChangeDATECOMPTA ;
    Function  PrintMois : Boolean ;
  public
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

Uses SaisUtil ;

procedure CumulPeriodique(Lefb : TFichierBase ; NET : TNatureEtatJal) ;
var QR: TFQRCumulG ;
    NB : TNatureBase ;
    MS : TModeSelect ;
    NomFiltre : String ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:= TFQRCumulG.Create(Application) ;
QR.QRP.QRPrinter.OnSynZoom:=QR.TiersZoom ;
QR.fbEdt:=Lefb ;
QR.NETJ:=NET ;
Case QR.fbEdt Of
  fbGene : BEGIN NB:=nbGen ; MS:=msGenEcr ; NomFiltre:='QRCUMGEN' ; Edition.Etat:=etCumulGen ; END ;
  fbAux  : BEGIN NB:=nbAux ; MS:=msAuxEcr ; NomFiltre:='QRCUMAUX' ; Edition.Etat:=etCumulAux ; END ;
  fbSect : BEGIN NB:=nbSec ; MS:=msSecAna ; NomFiltre:='QRCUMSEC' ; Edition.Etat:=etCumulAna ; END ;
  fbJal  : BEGIN
           NB:=nbJal ; MS:=msRien   ; NomFiltre:='' ;
           Case NET Of
             neJalPer : BEGIN NB:=nbJal ; MS:=msRien   ; NomFiltre:='QRJALPER' ; Edition.Etat:=etJalPer ; END ;
             neJalCentr : BEGIN NB:=nbJal ; MS:=msRien   ; NomFiltre:='QRJALCEN' ; Edition.Etat:=etJalCentr ; END ;
             neJaG : BEGIN NB:=nbJal ; MS:=msRien   ; NomFiltre:='JALGENERAL' ; Edition.Etat:=etJalGen ; END ;
             END ;
           END ;
  Else Exit ;
  END ;
QR.InitType (NB,neCum,MS,NomFiltre,'',TRUE,FALSE,FALSE,Edition) ;
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
procedure CumulPeriodiqueChaine ( vCritEdtChaine : TCritEdtChaine );
//(Lefb : TFichierBase ; NET : TNatureEtatJal) ;
var QR: TFQRCumulG ;
    NB : TNatureBase ;
    MS : TModeSelect ;
    NomFiltre : String ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:= TFQRCumulG.Create(Application) ;
QR.QRP.QRPrinter.OnSynZoom:=QR.TiersZoom ;
QR.fbEdt:=fbJal;     // En dur, je c c nul mais c pour test
QR.NETJ:=neJalCentr; // Idem
Case QR.fbEdt Of
  fbGene : BEGIN NB:=nbGen ; MS:=msGenEcr ; NomFiltre:='QRCUMGEN' ; Edition.Etat:=etCumulGen ; END ;
  fbAux  : BEGIN NB:=nbAux ; MS:=msAuxEcr ; NomFiltre:='QRCUMAUX' ; Edition.Etat:=etCumulAux ; END ;
  fbSect : BEGIN NB:=nbSec ; MS:=msSecAna ; NomFiltre:='QRCUMSEC' ; Edition.Etat:=etCumulAna ; END ;
  fbJal  : BEGIN
           NB:=nbJal ; MS:=msRien   ; NomFiltre:='' ;
           Case QR.NETJ Of
             neJalPer : BEGIN NB:=nbJal ; MS:=msRien   ; NomFiltre:='QRJALPER' ; Edition.Etat:=etJalPer ; END ;
             neJalCentr : BEGIN NB:=nbJal ; MS:=msRien   ; NomFiltre:='QRJALCEN' ; Edition.Etat:=etJalCentr ; END ;
             neJaG : BEGIN NB:=nbJal ; MS:=msRien   ; NomFiltre:='JALGENERAL' ; Edition.Etat:=etJalGen ; END ;
             END ;
           END ;
  Else Exit ;
  END ;
QR.CritEdtChaine := vCritEdtChaine;
QR.InitType (NB,neCum,MS,NomFiltre,'',TRUE,FALSE,FALSE,Edition) ;
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

procedure CumulPeriodiqueZoom(Crit : TCritEdt ; Lefb : TFichierBase) ;
var QR : TFQRCumulG ;
    NB : TNatureBase ;
    MS : TModeSelect ;
    NomFiltre : String ;
    Edition : TEdition ;
BEGIN
QR:=TFQRCumulG.Create(Application) ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.TiersZoom ;
 QR.fbEdt:=Lefb ;
 QR.CritEdt:=Crit ;
 Case QR.fbEdt Of
   fbGene : BEGIN NB:=nbGen ; MS:=msGenEcr ; NomFiltre:='QRCUMGEN' ; Edition.Etat:=etCumulGen ; END ;
   fbAux  : BEGIN NB:=nbAux ; MS:=msAuxEcr ; NomFiltre:='QRCUMAUX' ; Edition.Etat:=etCumulAux ;END ;
   fbSect : BEGIN NB:=nbSec ; MS:=msSecAna ; NomFiltre:='QRCUMSEC' ; Edition.Etat:=etCumulAna ; END ;
   fbJal  : BEGIN
            NB:=nbJal ; MS:=msRien   ; NomFiltre:='QRJALPER' ;  Edition.Etat:=etJalPer ;
            END ;
   Else Exit ;
   END ;
 QR.InitType (NB,neCum,MS,NomFiltre,'',FALSE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

Procedure TFQRCumulG.TiersZoom(Quoi : String) ;
var i : Integer ;
BEGIN
i:=0 ;
Case fbEdt Of
  fbGene : i:=1 ;
  fbAux  : i:=2 ;
  fbSect : i:=3 ;
  fbJal  : i:=4 ;
  END ;
if i<>0 then ZoomEdt(i,Quoi) ;
END ;

procedure TFQRCumulG.InitDivers ;
BEGIN
Inherited ;
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
Case CritEdt.Monnaie of
  0 : TitreColMonFE.Caption:=Copy(FMontant.Items[0],2,Length(Fmontant.Items[0])-1) ;
  1 : TitreColMonFE.Caption:=MsgLibel.Mess[34] ;
End ;
BDetail.ForceNewPage:=FSautDePage.Checked ;
BGeneFooter.Frame.DrawBottom:=CritEdt.Cum.FormatPrint.PrSepCompte ;
AnDebit.Caption:='' ; ANCredit.Caption:='' ; ANSolde.Caption:='' ;
AnDebitDev.Caption:='' ; ANCreditDev.Caption:='' ; ANSoldeDev.Caption:='' ;
TitreAN.Caption:='' ; BDetail.Height:=17 ;
If fbEdt<>fbJal Then
   BEGIN
   If CritEdt.Cum.CalculAN Then
      BEGIN
      TitreAN.Caption:=MsgLibel.Mess[7] ;
      BDetail.Height:=35 ;
      END ;
   END Else
   BEGIN
   If NETJ=neJalCentr Then TitreColLibelle.Caption:=' ' ;
   If NETJ=neJaG Then
      BEGIN
      BTotPeriode.Frame.DrawTop:=CritEdt.Cum.FormatPrint.PrSepCompte ;
//      TitreColCpt.Caption:=' ' ; TitreColLibelle.Alignment:=taLeftJustify ;
      TitreColCpt.Caption:=TitreColLibelle.Caption ; TitreColLibelle.Caption:=' ' ;
      //TitreColCpt.Caption:=TitreColLibelle.Caption ; TitreColLibelle.Caption:='' ;
      END ;
   END ;
END ;

procedure TFQRCumulG.GenereSQL ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
GenereSQLBase ;
END;

procedure TFQRCumulG.RenseigneCritere ;
{ Récupération des champs du multicritère dans l'entête d'état }
BEGIN
Inherited ;
{ Nature }
QRLabel12.Visible:=(fbEdt<>fbJal) ;
RNatureCpt.Visible:=(fbEdt<>fbJal) ;
{ Mode Select }
TRSelectCpte.Visible:=(fbEdt<>fbJal) ;
RSelectCpte.Visible:=(fbEdt<>fbJal) ;
{ Sauf Cpt }
QRLabel24.Visible:=(fbEdt<>fbJal) ;
RExcepGen.Visible:=(fbEdt<>fbJal) ;
CaseACocher(FAvecJalAN,RAN) ;
if (NETJ=neJalPer) or (NETJ=neJalCentr) then
   BEGIN
   If CritEdt.Cum.DeuxMontant Then BTitre.Height:=92 else BTitre.Height:=155 ;
   END ;
if NETJ=neJaG then
   BEGIN
   {AnV}
   QRLabel9.Visible:=False ; RAN.Visible:=False ;
   {DateComptable}
   RDateCompta1.Visible:=False ; QRLabel4.Visible:=False ;
   RDateCompta2.Visible:=False ; QRLabel11.Visible:=False ;
   {Fourchette}
   RCpte.Visible:=False ; RCpte1.Visible:=False ;
   TRa.Visible:=False ; RCpte2.Visible:=False ; RJoker.Visible:=False ;
   {Positionnement de label}
//      BTitre.Height:=109 else BTitre.Height:=150 ;
   BTitre.Height:=73 ;
   If CritEdt.Cum.DeuxMontant Then
      BEGIN
      //BTitre.Height:=73 ;
        {type Ecr}  // --> Niveau d' Etab
      QRLabel7.Left:=QRLabel15.Left ; RNatureEcr.Left:=REtab.Left ;
      QRLabel7.Top:=QRLabel15.Top ; RNatureEcr.Top:=REtab.Top ;
        {Etab}      // --> Niveau de L'exo
      QRLabel15.Left:=QRLabel5.Left ; REtab.Left:=RExercice.Left ;
      QRLabel15.Top:=RCpte.Top ; REtab.Top:=RCpte.Top ;
        {Exo}       // --> Niveau de du Compte 1
      QRLabel5.Left:=RCpte.Left ; RExercice.Left:=RCpte1.Left ;
        {Devise}    // --> Niveau de Nature de Compte
      QRLabel18.Left:=QRLabel5.Left ; RDevises.Left:=RNatureCpt.Left;
      QRLabel18.Top:=QRLabel12.Top ; RDevises.Top:=RNatureCpt.Top ;
      END Else
      BEGIN
      //BTitre.Height:=72 ;
        {Etab}      // --> Niveau de L'exo
      QRLabel15.Left:=QRLabel5.Left ; REtab.Left:=RExercice.Left ;
      QRLabel15.Top:=RCpte.Top ; REtab.Top:=RCpte.Top ;
        {Exo}       // --> Niveau de du Compte 1
      QRLabel5.Left:=RCpte.Left ; RExercice.Left:=RCpte1.Left ;
        {type Ecr}  // --> Niveau de Nature de Cpt
      QRLabel7.Left:=QRLabel12.Left ; RNatureEcr.Left:=RNatureCpt.Left ;
      QRLabel7.Top:=RNatureCpt.Top ; RNatureEcr.Top:=RNatureCpt.Top ;
        {Devise}    // --> Niveau de Date Comptable
      QRLabel18.Left:=QRLabel4.Left ; RDevises.Left:=RDateCompta1.Left ;
      QRLabel18.Top:=QRLabel12.Top ; RDevises.Top:=QRLabel12.Top ;
      END ;
   END ;
END ;

procedure TFQRCumulG.ChoixEdition ;
{ Initialisation des options d'édition }
Var fb : TFichierBase ;
BEGIN
Inherited ;
if ((NETJ=neJaG)or((NETJ=neJalPer) and CritEdt.Cum.RecapMens)) Then ChargeRecap(LRecap) ;

(*Rony 21/05/97*)
fb:=fbEdt ;
If fbEdt=fbSect Then fb:=AxeToFb(CritEdt.Cum.Axe) ;
{Rony 19/09/97. Une édition en devise ne marche pas}
//Gcalc:=TGCalculCum.create(Un,fb,fb,QuelTypeEcr,Dev,Etab,Exo,TRUE,CritEdt.Monnaie=2,CritEdt.DecimalePivot,V_PGI.OkDecE) ;
Gcalc:=TGCalculCum.create(Un,fb,fb,QuelTypeEcr,Dev,Etab,Exo,(CritEdt.Monnaie=0)or(CritEdt.Monnaie=2),CritEdt.Monnaie=2,CritEdt.DecimalePivot,V_PGI.OkDecE) ;
GCalc.initCalcul('','','','',CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                 CritEdt.Date1,CritEdt.Cum.Date11,TRUE) ;
If CritEdt.Cum.DeuxMontant Then
   BEGIN
   GcalcDev:=TGCalculCum.create(Un,fb,fb,QuelTypeEcr,Dev,Etab,Exo,FALSE,FALSE,CritEdt.Decimale,V_PGI.OkDecE) ;
   GCalcDev.initCalcul('','','','',CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                       CritEdt.Date1,CritEdt.Cum.Date11,TRUE) ;
   END ;

{ Masque pour la Devise }
If CritEdt.Cum.DeuxMontant Then
   BEGIN
   ChangeMask(Formateur,CritEDT.Decimale,CritEDT.Symbole);
   CritEDT.FormatMontant:=Formateur.Masks.PositiveMask;
   ChangeMask(FormateurPivot,CritEdt.DecimalePivot,CritEdt.SymbolePivot);
   CritEdt.FormatMontantPivot:=FormateurPivot.Masks.PositiveMask;
   END Else
   BEGIN
   If (CritEdt.DevPourExist<>'') And (CritEdt.Monnaie<>2) Then
     BEGIN
     CritEdt.SymbolePivot:=CritEdt.Symbole ;
     CritEdt.DecimalePivot:=CritEdt.Decimale ;
     ChangeMask(FormateurPivot,CritEdt.Decimale,CritEdt.Symbole);
     CritEdt.FormatMontantPivot:=FormateurPivot.Masks.PositiveMask;
     END ;
   END ;
If FMonetaire.Checked And (CritEdt.Monnaie=2) And CritEdt.Cum.DeuxMontant Then CritEdt.AfficheSymbole:=TRUE ; 
END ;

Procedure TFQRCumulG.FinirPrint ;
BEGIN
Inherited ;
If GCalc<>NIL Then BEGIN GCalc.Free ; GCalc:=NIL ; END ;
If CritEdt.Cum.DeuxMontant Then If GCalcDev<>NIL Then BEGIN GCalcDev.Free ; GCalcDev:=NIL ; END ;
if ((NETJ=neJaG)or((NETJ=neJalPer)and CritEdt.Cum.RecapMens)) Then VideRecap(LRecap) ;

END ;

function  TFQRCumulG.CritOk : Boolean ;
//Var fb : TFichierBase ;
BEGIN
Result:=Inherited CritOK ;
(* Rony 21/05/97
fb:=fbEdt ;
If fbEdt=fbSect Then fb:=AxeToFb(CritEdt.Cum.Axe) ;
If Result Then
   BEGIN
   Gcalc:=TGCalculCum.create(False,fb,fb,QuelTypeEcr,Dev,Etab,Exo,TRUE,CritEdt.Monnaie=2,CritEdt.DecimalePivot,V_PGI.OkDecE) ;
   GCalc.initCalcul('','','','',CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                    CritEdt.Date1,CritEdt.Cum.Date11,TRUE) ;
   If CritEdt.Cum.DeuxMontant Then
      BEGIN
      GcalcDev:=TGCalculCum.create(False,fb,fb,QuelTypeEcr,Dev,Etab,Exo,FALSE,FALSE,CritEdt.Decimale,V_PGI.OkDecE) ;
      GCalcDev.initCalcul('','','','',CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                          CritEdt.Date1,CritEdt.Cum.Date11,TRUE) ;
      END ;
   END ;
*)
//InitReport([1],CritEDT.BAL.FormatPrint.Report) ;
END ;

procedure TFQRCumulG.RecupCritEdt ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  if FSoldeProg.State=cbGrayed then SoldeProg:=0 ;
  if FSoldeProg.State=cbChecked then SoldeProg:=1 ;
  if FSoldeProg.State=cbUnChecked then SoldeProg:=2 ;
  if FSautDePage.State=cbGrayed then SautPage:=0 ;
  if FSautDePage.State=cbChecked then SautPage:=1 ;
  if FSautDePage.State=cbUnChecked then SautPage:=2 ;
  If fbEdt=fbSect Then Cum.Axe:=FNatureCpt.Value ;
  Cum.AvecJalAN:=FAvecJalAN.Checked ; Cum.RecapMens:=False ;
  If (NETJ=neJalPer) then Cum.RecapMens:=RecapMens.Checked ;
  Cum.DeuxMontant:=False ;
  if (Monnaie>0) And (DeviseSelect<>'') And (DeviseSelect<>V_PGI.DevisePivot) then
     BEGIN
     DEVPourExist:=DeviseSelect ;
     END ;
  if ((Monnaie=0) or (Monnaie=2)) and (DeviseSelect<>'') And (DeviseSelect<>V_PGI.DevisePivot) then if FAfficheDevise.Checked then Cum.DeuxMontant:=TRUE ;
  With Cum.FormatPrint Do
       BEGIN
       PutTabCol(1,TabColl[1],TitreColCpt.Tag,True) ;
       PutTabCol(2,TabColl[2],TitreColLibelle.Tag,True) ;
(**)
       PutTabCol(3,TabColl[3],TitreColMonFE.Tag,True) ;
       PutTabCol(4,TabColl[4],TSolde.Tag,True) ;
(**)
       if (Monnaie=1) And (DeviseSelect<>'') And (DeviseSelect<>V_PGI.DevisePivot) then
          BEGIN
          (* Rony 19/09/97 Montants pas affiche en Devise uniquement
          PutTabCol(3,TabColl[3],TitreColMonFE.Tag,False) ;
          PutTabCol(4,TabColl[4],TSolde.Tag,False) ;
          PutTabCol(5,TabColl[5],TitreColMonDev.Tag,True) ;
          PutTabCol(6,TabColl[6],TSoldeDev.Tag,True) ;
          *)
          PutTabCol(3,TabColl[3],TitreColMonFE.Tag,True) ;
          PutTabCol(4,TabColl[4],TSolde.Tag,True) ;
          PutTabCol(5,TabColl[5],TitreColMonDev.Tag,False) ;
          PutTabCol(6,TabColl[6],TSoldeDev.Tag,False) ;

          END Else
          BEGIN
          PutTabCol(5,TabColl[5],TitreColMonDev.Tag,CritEdt.Cum.DeuxMontant) ;
          PutTabCol(6,TabColl[6],TSoldeDev.Tag,CritEdt.Cum.DeuxMontant) ;
          END ;
       END ;
  END ;
END ;

procedure TFQRCumulG.FExerciceChange(Sender: TObject);
begin
GenPerExo ;
  inherited;
end;

procedure TFQRCumulG.GenPerExo ;
{ Création des périodes comptables pour un exercice donné }
var DateExo            : TExoDate ;
    i                  : integer ;
    Annee,pMois,NbMois : Word ;
    DD : TdateTime ;
    //DebMois, LAnnee : Word ;
    D1, D2 : String ;
BEGIN
NbMois:=0 ; //DebMois:=0 ; LAnnee:=0 ;
QuelDateDeExo(FExercice.Value,DateExo) ; NOMBREPEREXO(DateExo,pMois,Annee,NbMois) ;
//DebMois:=pMois ; LAnnee:=Annee ;
{If Not QRloading then  // FQ 10248
   BEGIN}
   FDateCpta1.Items.Clear ; FDateCpta1.Values.Clear ; FDateCpta2.Items.clear ; FDateCpta2.Values.Clear ;
   FDate1.Items.Clear ; FDate1.Values.Clear ; FDate2.Items.Clear ; FDate2.Values.Clear ;
   for i:=0 to NbMois-1 do
       BEGIN
       DD:=PlusMois(DateExo.Deb,i) ;
       D1:=FormatDateTime('mmmm yyyy',DD) ;
       D2:=FormatDateTime('mmmm yyyy',DD) ;
       FDateCpta1.Items.Add(FirstMajuscule(D1));
       FDateCpta2.Items.Add(FirstMajuscule(D2));
       //FDateCpta1.Items.Add(FormatDateTime('mmmm yyyy',DD));
       //FDateCpta2.Items.Add(FormatDateTime('mmmm yyyy',DD));
       FDateCpta1.Values.Add(IntTostr(i)) ; // Rony 14/05/97 -- A cause du filtre
       FDateCpta2.Values.Add(IntTostr(i)) ; //
       FDate1.Items.Add(DateToStr(DebutdeMois(dd))) ;
       FDate2.Items.Add(DateToStr(FindeMois(dd))) ;
       END ;
  FDateCpta1.ItemIndex:=0 ; FDateCpta2.ItemIndex:=NbMois-1 ;
//  END;
FDate1.ItemIndex:=FDateCpta1.ItemIndex; FDate2.ItemIndex:=FDateCpta2.ItemIndex ;
ChangeDateCompta ;
END ;

procedure TFQRCumulG.ChangeDATECOMPTA ;
BEGIN
FDateCompta1.Text:=FDate1.Items[FDate1.ItemIndex] ;
FDateCompta2.Text:=FDate2.Items[FDate2.ItemIndex] ;
END ;

procedure TFQRCumulG.FDateCpta1Change(Sender: TObject);
begin
  inherited;
FDate1.ItemIndex:=FDateCpta1.ItemIndex ; FDate2.ItemIndex:=FDateCpta2.ItemIndex ;
ChangeDateCompta ;
end;

procedure TFQRCumulG.BGeneFooterBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
{ Affiche les Totaux Par Cpt. Géné. }
Case fbEdt Of
  fbGene,fbAux,fbSect : TitreTotal.Caption:=MsgLibel.Mess[2]+' '+QR1COMPTE.ASString ;
  fbJal : Case NETJ Of
            neJalPer : TitreTotal.Caption:=MsgLibel.Mess[24]+' '+QR1COMPTE.ASString ;
            neJalCentr : TitreTotal.Caption:=QR1COMPTE.ASString+' '+QR1Libelle.AsString ;
            END ;
  END ;
//rony 15/04/97 TitreTotal.Caption:=MsgLibel.Mess[2]+' '+QR1COMPTE.ASString+' '+QR1LIBELLE.AsString ;
Quoi:=Qr1COMPTE.AsString+' '+Qr1LIBELLE.AsString ;

TOT1DEBIT.Caption:= AfficheMontant(CritEdt.FormatMontantPivot,CritEdt.SymbolePivot,TotCptGene[1].TotDebit,CritEdt.AfficheSymbole) ;
TOT1CREDIT.Caption:=AfficheMontant(CritEdt.FormatMontantPivot,CritEdt.SymbolePivot,TotCptGene[1].TotCredit,CritEdt.AfficheSymbole) ;
TOT1SOLDE.Caption:=PrintSolde(TotCptGene[1].TotDebit,TotCptGene[1].TotCredit,CritEdt.Decimale,CritEdt.SymbolePivot,CritEdt.AfficheSymbole) ;

TOT1DEBITDEV.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCptGene[2].TotDebit,CritEdt.AfficheSymbole ) ;
TOT1CREDITDEV.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCptGene[2].TotCredit,CritEdt.AfficheSymbole ) ;
TOT1SOLDEDEV.Caption:=PrintSolde(TotCptGene[2].TotDebit,TotCptGene[2].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
if NETJ=neJaG then PrintBand:=False ;
end;

procedure TFQRCumulG.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TOT2DEBIT.Caption:= AfficheMontant(CritEdt.FormatMontantPivot,CritEdt.SymbolePivot,TotEdt[1].TotDebit,CritEdt.AfficheSymbole ) ;
TOT2CREDIT.Caption:=AfficheMontant(CritEdt.FormatMontantPivot,CritEdt.SymbolePivot,TotEdt[1].TotCredit,CritEdt.AfficheSymbole ) ;
TOT2SOLDE.Caption:=PrintSolde(TotEdt[1].TotDebit,TotEdt[1].TotCredit,CritEdt.DecimalePivot,CritEdt.SymbolePivot,CritEdt.AfficheSymbole) ;

TOT2DEBITDEV.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[2].TotDebit,CritEdt.AfficheSymbole ) ;
TOT2CREDITDEV.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[2].TotCredit,CritEdt.AfficheSymbole ) ;
TOT2SOLDEDEV.Caption:=PrintSolde(TotEdt[2].TotDebit,TotEdt[2].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFQRCumulG.BCumulBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
Quoi:=Qr1COMPTE.AsString+' '+Qr1LIBELLE.AsString ;
{ Mvt Mensuel + Solde Progressif de chaque Cpt Géné. }
If NETJ=neJalCentr Then PrintBand:=False ;
if NETJ=neJaG then PrintBand:=False ;
end;

procedure TFQRCumulG.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var TotCpt1                  : TabTot ;
    TotCpt1Dev               : TabTot ;
    MReport                  : TabTRep ;
    StAno                    : String ;
begin
  inherited;
LaDate1:=0 ; LaDate2:=0 ;
TCompte.Left:=TitreColCpt.Left+3 ;
TCompte.Width:=(TitreColCpt.Width+TitreColLibelle.Width)+1 ;
Quoi:=Qr1COMPTE.AsString+' '+Qr1LIBELLE.AsString ;
Case CritEdt.SautPage of
  0 : BDetail.forceNewPage:=(Qr1SAUTPAGE.AsString='X') ;
  1 : BDetail.forceNewPage:=True ;
  2 : BDetail.forceNewPage:=False ;
 end ;
StReportGen:=Qr1COMPTE.AsString ;
InitReport([2],CritEdt.Cum.FormatPrint.Report) ;
Fillchar(TotCptGene,SizeOf(TotCptGene),#0) ; { Init du total Cpt. Géné. }
Fillchar(TotProg,SizeOf(TotProg),#0) ;       { Init du solde Progressif }
Fillchar(TotCpt1,SizeOf(TotCpt1),#0) ;
Fillchar(TotCpt1Dev,SizeOf(TotCpt1Dev),#0) ;
Fillchar(MReport,SizeOf(MReport),#0) ;
TCompte.Caption:=Qr1COMPTE.AsString+' '+Qr1Libelle.AsString ;
If fbEdt=fbJal Then
   BEGIN
   If QR1NATURE.AsString='ANO' Then BDetail.Height:=35 Else BDetail.Height:=17 ;
   END ;
//If Not CritEdt.Cum.CalculAN Then Exit ; // 14388
GCAlc.ReInitCalcul(Qr1COMPTE.AsString,'',CritEDT.Date1,CritEDT.BAL.Date11) ;
GCalc.Calcul ; TotCpt1:=GCalc.ExecCalc.TotCpt ;
If fbEdt<>fbJal Then CumulVersSolde(TotCpt1[0]) ;
If CritEdt.Cum.DeuxMontant Then
   BEGIN
   GCAlcDev.ReInitCalcul(Qr1COMPTE.AsString,'',CritEDT.Date1,CritEDT.BAL.Date11) ;
   GCalcDev.Calcul ; TotCpt1Dev:=GCalcDev.ExecCalc.TotCpt ;
   If fbEdt<>fbJal Then CumulVersSolde(TotCpt1Dev[0]) ;
   END ;
{ Chaque Cpt. Géné. , }
TotCptGene[1].TotDebit:=  Arrondi(TotCptGene[1].TotDebit+TotCpt1[0].TotDebit, CritEdt.DecimalePivot) ;
TotCptGene[1].TotCredit:= Arrondi(TotCptGene[1].TotCredit+TotCpt1[0].TotCredit, CritEdt.DecimalePivot) ;
TotCptGene[2].TotDebit:=  Arrondi(TotCptGene[2].TotDebit+TotCpt1Dev[0].TotDebit, CritEdt.Decimale) ;
TotCptGene[2].TotCredit:= Arrondi(TotCptGene[2].TotCredit+TotCpt1Dev[0].TotCredit, CritEdt.Decimale) ;

{ Le Solde Progressif, }
TotProg[1].TotDebit:=  Arrondi(TotProg[1].TotDebit +TotCpt1[0].TotDebit, CritEdt.DecimalePivot) ;
TotProg[1].TotCredit:= Arrondi(TotProg[1].TotCredit+TotCpt1[0].TotCredit, CritEdt.DecimalePivot) ;
TotProg[2].TotDebit:=  Arrondi(TotProg[2].TotDebit +TotCpt1Dev[0].TotDebit, CritEdt.Decimale) ;
TotProg[2].TotCredit:= Arrondi(TotProg[2].TotCredit+TotCpt1Dev[0].TotCredit, CritEdt.Decimale) ;

{ Le Total Gébéral }
TotEdt[1].TotDebit:=  Arrondi(TotEdt[1].TotDebit+TotCpt1[0].TotDebit, CritEdt.DecimalePivot) ;
TotEdt[1].TotCredit:= Arrondi(TotEdt[1].TotCredit+TotCpt1[0].TotCredit, CritEdt.DecimalePivot) ;
TotEdt[2].TotDebit:=  Arrondi(TotEdt[2].TotDebit+TotCpt1Dev[0].TotDebit, CritEdt.Decimale) ;
TotEdt[2].TotCredit:= Arrondi(TotEdt[2].TotCredit+TotCpt1Dev[0].TotCredit, CritEdt.Decimale) ;

//Arrondi(TotEdt[1].TotDebit+TotCpt1[0].TotDebit, CritEdt.DecimalePivot) ;
AnDebit.Caption:=AfficheMontant(CritEdt.FormatMontantPivot,CritEdt.SymbolePivot,TotCpt1[0].TotDebit, CritEdt.AfficheSymbole ) ;
AnCredit.Caption:=AfficheMontant(CritEdt.FormatMontantPivot,CritEdt.SymbolePivot,TotCpt1[0].TotCredit, CritEdt.AfficheSymbole) ;
AnSolde.Caption:=PrintSolde(TotCpt1[0].TotDebit,TotCpt1[0].TotCredit,CritEdt.DecimalePivot,CritEdt.SymbolePivot,CritEdt.AfficheSymbole) ;

AnDebitDev.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt1Dev[0].TotDebit, CritEdt.AfficheSymbole ) ;
AnCreditDev.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt1Dev[0].TotCredit, CritEdt.AfficheSymbole) ;
AnSoldeDev.Caption:=PrintSolde(TotCpt1Dev[0].TotDebit,TotCpt1Dev[0].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
{ Affectation des ANouv. pour le REPORT }
MReport[1].TotDebit:=TotCpt1[0].TotDebit ;
MReport[1].TotCredit:=TotCpt1[0].TotCredit ;
MReport[2].TotDebit:=TotCpt1Dev[0].TotDebit ;
MReport[2].TotCredit:=TotCpt1Dev[0].TotCredit ;
AddReportBAL([1,2],CritEdt.Cum.FormatPrint.Report,MReport,CritEdt.DecimalePivot) ;
if ((NETJ=neJaG)or((NETJ=neJalPer)and CritEdt.Cum.RecapMens)) Then
   BEGIN
   StAno:='0000000000' ;
   AddRecap(LRecap, [StAno], [StAno],[1,TotCpt1[0].TotDebit, TotCpt1[0].TotCredit,TotCpt1Dev[0].TotDebit, TotCpt1Dev[0].TotCredit]) ;
   PrintBand:=False ;
   END ;
If NETJ=neJalCentr Then PrintBand:=False ;
end;

procedure TFQRCumulG.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
Report1Solde.Caption:=Report2Solde.Caption ;
Report3Debit.Caption:=Report4Debit.Caption ;
Report3Credit.Caption:=Report4Credit.Caption ;
Report3Solde.Caption:=Report4Solde.Caption ;
end;

procedure TFQRCumulG.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var MReport        : TabTRep ;
begin
  inherited;
Fillchar(MReport,SizeOf(MReport),#0) ;
Case QuelReportBAL(CritEdt.Cum.FormatPrint.Report,MReport) Of
  1 : TitreReportB.Caption:=MsgLibel.Mess[3] ;
  2 : TitreReportB.Caption:=MsgLibel.Mess[4]+' '+StReportGen ;
  END ;

Report2Debit.Caption:=AfficheMontant(CritEdt.FormatMontantPivot, CritEdt.SymbolePivot, MReport[1].TotDebit, CritEdt.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEdt.FormatMontantPivot, CritEdt.SymbolePivot, MReport[1].TotCredit, CritEdt.AfficheSymbole ) ;
REPORT2SOLDE.Caption:=PrintSolde(MReport[1].TotDebit,MReport[1].TotCredit,CritEdt.DecimalePivot,CritEdt.SymbolePivot,CritEdt.AfficheSymbole) ;
Report4Debit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[2].TotDebit, CritEdt.AfficheSymbole ) ;
Report4Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[2].TotCredit, CritEdt.AfficheSymbole ) ;
REPORT4SOLDE.Caption:=PrintSolde(MReport[2].TotDebit,MReport[2].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
end;


procedure TFQRCumulG.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Case fbEdt Of
  fbGene : BEGIN
           Qr1COMPTE:=TStringField(Q.FindField('G_GENERAL'));
           Qr1LIBELLE:=TStringField(Q.FindField('G_LIBELLE'));
           Qr1SAUTPAGE:=TStringField(Q.FindField('G_SAUTPAGE')) ;
           Qr1SOLDEPROGRESSIF:=TStringField(Q.FindField('G_SOLDEPROGRESSIF')) ;
           END ;
  fbAux   : BEGIN
           Qr1COMPTE:=TStringField(Q.FindField('T_AUXILIAIRE'));
           Qr1LIBELLE:=TStringField(Q.FindField('T_LIBELLE'));
           Qr1SAUTPAGE:=TStringField(Q.FindField('T_SAUTPAGE')) ;
           Qr1SOLDEPROGRESSIF:=TStringField(Q.FindField('T_SOLDEPROGRESSIF')) ;
           END ;
  fbSect  : BEGIN
           Qr1COMPTE:=TStringField(Q.FindField('S_SECTION'));
           Qr1LIBELLE:=TStringField(Q.FindField('S_LIBELLE'));
           Qr1SAUTPAGE:=TStringField(Q.FindField('S_SAUTPAGE')) ;
           Qr1SOLDEPROGRESSIF:=TStringField(Q.FindField('S_SOLDEPROGRESSIF')) ;
           END ;
  fbJal   : BEGIN
           Qr1COMPTE:=TStringField(Q.FindField('J_JOURNAL'));
           Qr1LIBELLE:=TStringField(Q.FindField('J_LIBELLE'));
           Qr1NATURE:=TStringField(Q.FindField('J_NATUREJAL'));
           END ;
  END ;
end;

Function TFQRCumulG.PrintMois : Boolean ;
BEGIN
If (LaDate1=0) And (LaDate2=0) Then LaDate1:=DebutDeMois(CritEdt.Cum.Date21)
                               Else LaDate1:=PlusMois(LaDate1,1) ;
LaDate2:=FinDeMois(LaDate1) ;
PrintMois:=LaDate1<CritEdt.DateFin ;
END ;


procedure TFQRCumulG.QRDLMoisNeedData(var MoreData: Boolean);
Var MReport : TabTRep ;
    TotMois,TotMoisDev : TabTot ;
    StAno, StDate : String ;
begin
  inherited;
MoreData:=PrintMois ;
If MoreData Then
   BEGIN
   StDate:=FormatDatetime('mmmm yyyy',LaDate1) ;
   LaDate.Caption:=FirstMajuscule(StDate) ;
   Fillchar(MReport,SizeOf(MReport),#0) ;
   Fillchar(TotMois,SizeOf(TotMois),#0) ;
   Fillchar(TotMoisDev,SizeOf(TotMoisDev),#0) ;
   GCAlc.ReInitCalcul(Qr1COMPTE.AsString,'',LaDate1,LaDate2) ;
   GCalc.Calcul ; TotMois:=GCalc.ExecCalc.TotCpt ;
   If CritEdt.Cum.DeuxMontant Then
      BEGIN
      GCAlcDev.ReInitCalcul(Qr1COMPTE.AsString,'',LaDate1,LaDate2) ;
      GCalcDev.Calcul ; TotMoisDev:=GCalcDev.ExecCalc.TotCpt ;
      END ;

   if ((NETJ=neJaG)or((NETJ=neJalPer)and CritEdt.Cum.RecapMens)) Then
      BEGIN
      If QR1NATURE.AsString='ANO' then StAno:='0000000000' Else StAno:='' ;
      AddRecap(LRecap, [FormatDatetime('mmyyyy',LaDate1)], [FormatDatetime('mmmm yyyy',LaDate1)],[1,TotMois[1].TotDebit, TotMois[1].TotCredit,TotMoisDev[1].TotDebit, TotMoisDev[1].TotCredit]) ;
      AddRecap(LRecap, [StAno], [StAno],[1,TotMois[1].TotDebit, TotMois[1].TotCredit,TotMoisDev[1].TotDebit, TotMoisDev[1].TotCredit]) ;
      END ;
   { Incrémentation des montants pour ... }
   { Chaque Cpt General }
   TotCptGene[1].TotDebit:=  Arrondi(TotCptGene[1].TotDebit+TotMois[1].TotDebit, CritEdt.DecimalePivot) ;
   TotCptGene[1].TotCredit:= Arrondi(TotCptGene[1].TotCredit+TotMois[1].TotCredit, CritEdt.DecimalePivot) ;
   TotCptGene[2].TotDebit:=  Arrondi(TotCptGene[2].TotDebit+TotMoisDev[1].TotDebit, CritEdt.Decimale) ;
   TotCptGene[2].TotCredit:= Arrondi(TotCptGene[2].TotCredit+TotMoisDev[1].TotCredit, CritEdt.Decimale) ;
   { Le Total Gébéral }
   TotEdt[1].TotDebit:=  Arrondi(TotEdt[1].TotDebit+TotMois[1].TotDebit, CritEdt.DecimalePivot) ;
   TotEdt[1].TotCredit:= Arrondi(TotEdt[1].TotCredit+TotMois[1].TotCredit, CritEdt.DecimalePivot) ;
   TotEdt[2].TotDebit:=  Arrondi(TotEdt[2].TotDebit+TotMoisDev[1].TotDebit, CritEdt.Decimale) ;
   TotEdt[2].TotCredit:= Arrondi(TotEdt[2].TotCredit+TotMoisDev[1].TotCredit, CritEdt.Decimale) ;

   DEBIT.Caption:=AfficheMontant(CritEdt.FormatMontantPivot, CritEdt.SymbolePivot, TotMois[1].TotDebit, CritEdt.AfficheSymbole) ;
   CREDIT.Caption:=AfficheMontant(CritEdt.FormatMontantPivot, CritEdt.SymbolePivot, TotMois[1].TotCredit, CritEdt.AfficheSymbole) ;

   DEBITDEV.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotMoisDev[1].totDebit, CritEdt.AfficheSymbole) ;
   CREDITDEV.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotMoisDev[1].totCredit, CritEdt.AfficheSymbole) ;

   Case CritEdt.SoldeProg of             { Affectation Du Calcul du Solde Progressif O/N Sinon, d'aprés l'info sur le compte }
     0 : if Qr1SOLDEPROGRESSIF.AsString='X' then
            BEGIN
            SOLDE.Caption:=PrintSolde(TotMois[1].totDebit+TotProg[1].TotDebit,TotMois[1].TotCredit+TotProg[1].TotCredit,CritEdt.DecimalePivot,CritEdt.SymbolePivot,CritEdt.AfficheSymbole) ;
            SOLDEDEV.Caption:=PrintSolde(TotMoisDev[1].TotDebit+TotProg[2].TotDebit,TotMoisDev[1].TotCredit+TotProg[2].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
            END Else
            BEGIN
            SOLDE.Caption:=PrintSolde(TotMois[1].totDebit,TotMois[1].TotCredit,CritEdt.DecimalePivot,CritEdt.SymbolePivot,CritEdt.AfficheSymbole) ;
            SOLDEDEV.Caption:=PrintSolde(TotMoisDev[1].totDebit,TotMoisDev[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
            END ;
     1 : BEGIN
         SOLDE.Caption:=PrintSolde(TotMois[1].totDebit+TotProg[1].TotDebit,TotMois[1].TotCredit+TotProg[1].TotCredit,CritEdt.DecimalePivot,CritEdt.SymbolePivot,CritEdt.AfficheSymbole) ;
         SOLDEDEV.Caption:=PrintSolde(TotMoisDev[1].TotDebit+TotProg[2].TotDebit,TotMoisDev[1].TotCredit+TotProg[2].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
         END ;
     2 : BEGIN
         SOLDE.Caption:=PrintSolde(TotMois[1].totDebit,TotMois[1].TotCredit,CritEdt.DecimalePivot,CritEdt.SymbolePivot,CritEdt.AfficheSymbole) ;
         SOLDEDEV.Caption:=PrintSolde(TotMoisDev[1].totDebit,TotMoisDev[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
         END ;
    end ;

   {Pour le Calcul du Report uniquement en Devise Pivot }
   MReport[1].TotDebit:= Arrondi(MReport[1].TotDebit+TotMois[1].TotDebit, CritEdt.DecimalePivot) ;
   MReport[1].TotCredit:=Arrondi(MReport[1].TotCredit+TotMois[1].TotCredit, CritEdt.DecimalePivot) ;

   MReport[2].TotDebit:= Arrondi(MReport[2].TotDebit+TotMoisDev[1].TotDebit, CritEdt.DecimalePivot) ;
   MReport[2].TotCredit:=Arrondi(MReport[2].TotCredit+TotMoisDev[1].TotCredit, CritEdt.DecimalePivot) ;
   AddReportBAL([1,2],CritEdt.Cum.FormatPrint.Report,MReport,CritEdt.DecimalePivot) ;

   TotProg[1].TotDebit:=  Arrondi(TotProg[1].TotDebit +TotMois[1].TotDebit, CritEdt.DecimalePivot) ;
   TotProg[1].TotCredit:= Arrondi(TotProg[1].TotCredit+TotMois[1].TotCredit, CritEdt.DecimalePivot) ;
   TotProg[2].TotDebit:=  Arrondi(TotProg[2].TotDebit +TotMoisDev[1].TotDebit, CritEdt.Decimale) ;
   TotProg[2].TotCredit:= Arrondi(TotProg[2].TotCredit+TotMoisDev[1].TotCredit, CritEdt.Decimale) ;
   END ;
end;

procedure TFQRCumulG.FormShow(Sender: TObject);
Var i,ii,ll : Integer ;
    st : String ;
begin
If fbEdt=fbJal then
   BEGIN
   FSautDePage.AllowGrayed:=False;
   FSautDePage.State:=cbUnChecked ;
   case NETJ of
     neJalCentr : BEGIN {Centralisateur}
                  HelpContext:=7397000 ;
                  //Standards.HelpContext:=7397010 ;
                  //Avances.HelpContext:=7397020 ;
                  //Mise.HelpContext:=7397030 ;
                  //Option.HelpContext:=7397040 ;
                  END ;
     neJalPer   : BEGIN {Périodique}
                  HelpContext:=7403000 ;
                  //Standards.HelpContext:=7403010 ;
                  //Avances.HelpContext:=7403020 ;
                  //Mise.HelpContext:=7403030 ;
                  //Option.HelpContext:=7403040 ;
                  END ;
     neJaG      : BEGIN {Etablissement}
                  HelpContext:=7406000 ;
                  //Standards.HelpContext:=7406010 ;
                  //Avances.HelpContext:=7406020 ;
                  //Option.HelpContext:=7406030 ;
                  //Mise.HelpContext:=7406020 ;
                  HLabel7.Parent:=Standards ; FEtab.Parent:=Standards ;
                  {Devise}
                  HLabel8.Parent:=Standards ; FDevises.Parent:=Standards ; Avances.TabVisible:=False;
                  END ;

     END ;
   END ;
  inherited;
{$IFDEF CCS3}
If fbEdt=fbSect Then BEGIN FNatureCpt.Visible:=FALSE ; HLabel2.Visible:=FALSE ; END ;
{$ENDIF}
i:=8 ; ii:=16 ; FSautDePage.AllowGrayed:=True ; RecapMens.Visible:=False ;
Case fbEdt Of
   fbGene : BEGIN
            HelpContext:=7478000 ;
            //Standards.HelpContext:=7478010 ;
            //Avances.HelpContext:=7478020 ;
            //Mise.HelpContext:=7478030 ;
            //Option.HelpContext:=7478040 ;

            i:=8 ; HLabel2.Caption:=MsgLibel.Mess[14] ;
            FNatureCpt.Vide:=TRUE ; ii:=16 ;
            TitreColCpt.Caption:=MsgLibel.Mess[25] ;
            HLabel1.Caption:=MsgLibel.Mess[27] ;
            END ;
   fbAux :  BEGIN
            HelpContext:=7479000 ;
            //Standards.HelpContext:=7479010 ;
            //Avances.HelpContext:=7479020 ;
            //Mise.HelpContext:=7479030 ;
            //Option.HelpContext:=7479040 ;

            i:=10 ; HLabel2.Caption:=MsgLibel.Mess[14] ;
            FNatureCpt.Vide:=TRUE ; ii:=17 ;
            TitreColCpt.Caption:=MsgLibel.Mess[25] ;
            HLabel1.Caption:=MsgLibel.Mess[27] ;
            END ;
   fbSect : BEGIN
            HelpContext:=7480000 ;
            //Standards.HelpContext:=7480010 ;
            //Avances.HelpContext:=7480020 ;
            //Mise.HelpContext:=7480030 ;
            //Option.HelpContext:=7480040 ;

            i:=12 ; HLabel2.Caption:=MsgLibel.Mess[15] ;
            FNatureCpt.Vide:=FALSE ; ii:=18 ;
            TitreColCpt.Caption:=MsgLibel.Mess[25] ;
            HLabel1.Caption:=MsgLibel.Mess[27] ;
            END ;
   fbJal :  BEGIN
            FAvecJalAN.Visible:=TRUE ; i:=20 ; OnCum.Visible:=False       ;
            QRLabel9.Visible:=TRUE ; RAN.Visible:=TRUE ;
            Case NETJ Of neJalPer : ii:=22 ; neJalCentr: ii:=29 ; neJaG: ii:=31 ; END ;
            FSoldeProg.State:=cbUnChecked ; FSoldeProg.Enabled:=False ;
            If NETJ=neJalCentr Then FLigneCpt.Checked:=FALSE ;
            If NETJ=neJalPer Then RecapMens.Visible:=V_PGI.SAV ;
            TSOLDE2.Caption:=MsgLibel.Mess[23] ;
            TSOLDEDEV2.Caption:=MsgLibel.Mess[23] ;
            TitreColCpt.Caption:=MsgLibel.Mess[26] ;
            HLabel1.Caption:=MsgLibel.Mess[28] ;
            { Combos : Nature de Cpt & Mvtés sur ...  adieu ... en jal  }
            FNatureCpt.Vide:=TRUE ; HLabel2.Enabled:=False ; FNatureCpt.Visible:=False ;
            FSelectCpte.Visible:=False ; FSelectCpte.Visible:=False ;
            TSelectCpte.Visible:=False ; TSelectCpte.Visible:=False ;
            { Top des autres combos }
            If (NETJ=neJalPer)or(NETJ=neJalCentr) Then
               BEGIN
              (**) { Exo}
               HLabel4.Top:=HLabel2.Top ;  FExercice.Top:=FNatureCpt.Top ;
               { Periode }
               HLabel3.Top:=TSelectCpte.Top ;  FDateCpta1.Top:=FSelectCpte.Top ;
               TFDateCpta2.Top:=TSelectCpte.Top ; FDateCpta2.Top:=FSelectCpte.Top ;
               { Type Ecr }
               HLabel5.Top:=HLabel1.Top ;  FNatureEcr.Top:=FExcep.Top ;
               (**)
               (*
                 {Date Compta}  // --> Niveau de Nature de Cpt
               QRLabel4.Left:=QRLabel12.Left ; RDateCompta1.Left:=RNatureCpt.Left ; QRLabel11.Left:=TRa.Left ; RDateCompta2.Left:=RCpte2.Left ;
               QRLabel4.SynCritere:=QRLabel12.SynCritere ; RDateCompta1.SynCritere:=RNatureCpt.SynCritere ; QRLabel11.SynCritere:=TRa.SynCritere ; RDateCompta2.SynCritere:=RCpte2.SynCritere ;
               QRLabel12.SynCritere:=0 ; RNatureCpt.SynCritere:=0 ; TRa.SynCritere:=0 ; RCpte2.SynCritere:=0 ;

                 {type Ecr}  // --> Niveau de Nature du Mode Seclect
               QRLabel7.Left:=TRSelectCpte.Left ; RNatureEcr.Left:=RNatureCpt.Left ;
               QRLabel7.SynCritere:=TRSelectCpte.SynCritere ; RNatureEcr.SynCritere:=RNatureCpt.SynCritere ;
               TRSelectCpte.SynCritere:=0 ; RNatureCpt.SynCritere:=0 ;
               *)
               END ;
            If NETJ=neJaG Then
               BEGIN
               TFGen.Visible:=False       ; TFGenJoker.Visible:=False ;
               FCpte1.Visible:=False      ; FCpte2.Visible:=False ; FJoker.Visible:=False ;
               HLabel2.Visible:=False     ; FNatureCpt.Visible:=False ;
               TSelectCpte.Visible:=False ; FSelectCpte.Visible:=False ;
               HLabel3.Visible:=False     ; FDateCpta1.Visible:=False ;
               TFDateCpta2.Visible:=False ; FDateCpta2.Visible:=False ;
               HLabel1.Visible:=False     ; FExcep.Visible:=False ;
               FAvecJalAN.Visible:=False  ; FAvecJalAN.Checked:=True ;
               OnCum.Visible:=False       ; TabSup.TabVisible:=False;
               {Positionnement des labels }
               HLabel4.top:=TFGen.Top ; FExercice.top:=FCpte1.Top ; {Exo}
               HLabel5.top:=HLabel2.Top ; FNatureEcr.top:=FNatureCpt.Top ;{Type Ecr}
               HLabel7.top:=TSelectCpte.Top ; FEtab.top:=FSelectCpte.Top ; {Etab}
               HLabel8.top:=FExcep.Top ; FDevises.top:=FExcep.Top ;{Devise}
               FLigneCpt.Caption:=MsgLibel.Mess[33] ;
               FSautDePage.Enabled:=False ;
               END ;
            FSautDePage.AllowGrayed:=False;
            FSautDePage.State:=cbUnChecked ;
            FSautDePage.Enabled:=FALSE ;
            END ;
  End ;
TFGen.Caption:=MsgLibel.Mess[i] ; TFGenJoker.Caption:=MsgLibel.Mess[i+1] ;
St:=MsgLibel.Mess[i] ;
ll:=Pos('&',St) ; If ll>0 Then St[ll]:=' ' ;
RCpte.Caption:=trim(St) ;
Self.Caption:=MsgLibel.Mess[ii] ;
QRP.ReportTitle:=Self.Caption ;
UpdateCaption(Self) ;
// GC - 20/12/2001
  if CritEdtChaine.Utiliser then InitEtatchaine('QRJALCEN');
// GC - 20/12/2001 - FIN
end;

procedure TFQRCumulG.BGeneFooterAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([2],CritEdt.Cum.FormatPrint.Report) ;
end;

procedure TFQRCumulG.QRDLRecapNeedData(var MoreData: Boolean);
Var Tot : Array[0..12] of Double ;
    Cod,Lib : String ; Ano : Boolean ;
begin
  inherited;
if Not ((NETJ=neJaG)or((NETJ=neJalPer)and CritEdt.Cum.RecapMens)) Then BEGIN MoreData:=FALSE ; exit ; END ;

//if NETJ<>neJaG then BEGIN MoreData:=FALSE ; exit ; END ;
MoreData:=PrintRecap(LRecap,Cod,Lib,Tot) ;
Ano:=Copy(Cod,1,3)='000' ;
If MoreData Then
   BEGIN
   If Ano then
      BEGIN
      LibRecap.Caption:=MsgLibel.Mess[32] ;
      LibRecap.Left:=TitreColLibelle.Left ;
      //LibRecap.Tag:=2 ;
      END Else
      BEGIN
      LibRecap.Caption:='   '+FirstMajuscule(Lib) ;
      //LibRecap.Tag:=-1 ;
      LibRecap.Left:=TitreColCpt.Left ;
      END ;
   MDEBIT.Caption:=AfficheMontant(CritEdt.FormatMontantPivot, CritEdt.SymbolePivot, Tot[1],  CritEdt.AfficheSymbole ) ;
   MCREDIT.Caption:=AfficheMontant(CritEdt.FormatMontantPivot, CritEdt.SymbolePivot, Tot[2],  CritEdt.AfficheSymbole ) ;
   MSOLDE.Caption:=PrintSolde(Tot[1],Tot[2],CritEdt.DecimalePivot,CritEdt.SymbolePivot,CritEdt.AfficheSymbole) ;
   MDEBITDEV.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, Tot[3],  CritEdt.AfficheSymbole ) ;
   MCREDITDEV.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, Tot[4],  CritEdt.AfficheSymbole ) ;
   MSOLDEDEV.Caption:=PrintSolde(Tot[3],Tot[4],CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
   END ;
end;

procedure TFQRCumulG.QRSUBTOTBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
PrintBand:=(NETJ=neJalPer)and(RecapMens.checked) ;
If Not PrintBand then Exit ;
TOT22DEBIT.Caption:= AfficheMontant(CritEdt.FormatMontantPivot,CritEdt.SymbolePivot,TotEdt[1].TotDebit,CritEdt.AfficheSymbole ) ;
TOT22CREDIT.Caption:=AfficheMontant(CritEdt.FormatMontantPivot,CritEdt.SymbolePivot,TotEdt[1].TotCredit,CritEdt.AfficheSymbole ) ;
TOT22SOLDE.Caption:=PrintSolde(TotEdt[1].TotDebit,TotEdt[1].TotCredit,CritEdt.DecimalePivot,CritEdt.SymbolePivot,CritEdt.AfficheSymbole) ;

TOT22DEBITDEV.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[2].TotDebit,CritEdt.AfficheSymbole ) ;
TOT22CREDITDEV.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotEdt[2].TotCredit,CritEdt.AfficheSymbole ) ;
TOT22SOLDEDEV.Caption:=PrintSolde(TotEdt[2].TotDebit,TotEdt[2].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
TOPREPORT.enabled:=FALSE ;
end;


procedure TFQRCumulG.FMontantClick(Sender: TObject);
begin
  inherited;
if FMontant.ItemIndex=1 then
   BEGIN
   FAfficheDevise.Checked:=False ;
   FAfficheDevise.Enabled:=False ;
   END Else FAfficheDevise.Enabled:=True ;
end;
procedure TFQRCumulG.Timer1Timer(Sender: TObject);
begin
// GC - 20/12/2001
  inherited;
// GC - 20/12/2001 - FIN
end;

end.
