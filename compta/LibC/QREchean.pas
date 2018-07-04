unit QREchean;


interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, Ent1, HEnt1, QrRupt, CritEdt, UtilEdt, HQry,SaisUtil,
{$IFNDEF CCMP}
  EdtLegal,
{$ENDIF}
  CpteUtil, ParamDBG, HSysMenu, Menus, DBCtrls, HTB97, HPanel, UiUtil, tCalcCum ;

procedure Echeancier(Origine : tSuiviMP = smpAucun) ;
procedure EcheancierZoom(Crit : TCritEdt) ;

type  TOR = Class
       Filled    : Boolean ;
       Printed   : Boolean ;
       Principal : Boolean ;
       Libelle   : String ;
       Tot       : Array[0..12] of Double ;
       END ;

type
  TFQREchean = class(TFQR)
    TFPaie: THLabel;
    E_MODEPAIE: THMultiValComboBox;
    TFEcheancier: THLabel;
    TFCollJoker: THLabel;
    FCollJoker: TEdit;
    TFaC: TLabel;
    FColl2: THCpteEdit;
    FColl: THLabel;
    FColl1: THCpteEdit;
    FSens: THValComboBox;
    TFSens: THLabel;
    FEdition: THValComboBox;
    HLabel10: THLabel;
    FLigneEchPied: TCheckBox;
    FSautPage: TCheckBox;
    TRPaie: TQRLabel;
    RPaie: TQRLabel;
    TREcheancier: TQRLabel;
    REcheancier: TQRLabel;
    RGen: TQRLabel;
    RGen1: TQRLabel;
    TRaG: TQRLabel;
    RGen2: TQRLabel;
    QRLabel28: TQRLabel;
    REdition: TQRLabel;
    TRSens: TQRLabel;
    RSens: TQRLabel;
    TRLegende: TQRLabel;
    RLegende: TQRLabel;
    QRLabel10: TQRLabel;
    TE_JOURNAL: TQRLabel;
    TE_PIECELIGNE: TQRLabel;
    TE_GENERAL: TQRLabel;
    TE_AUXILIAIRE: TQRLabel;
    TT_LIBELLE: TQRLabel;
    TLE_DEBIT: TQRLabel;
    TLE_CREDIT: TQRLabel;
    TE_REFINTERNE: TQRLabel;
    TE_DATECOMPTABLE: TQRLabel;
    TE_MODEPAIE: TQRLabel;
    TITRE1REP: TQRLabel;
    REPORT1MONTANT: TQRLabel;
    BEcheH: TQRBand;
    TEcheH: TQRLabel;
    LEVALIDE: TQRLabel;
    LE_JOURNAL: TQRDBText;
    LE_PIECELIG: TQRLabel;
    LE_MONTANT: TQRLabel;
    LE_CUMUL: TQRLabel;
    E_REFINTERNE: TQRDBText;
    E_DATECOMPTABLE: TQRDBText;
    LE_MODEPAIE: TQRDBText;
    BSDMulti: TQRBand;
    TCodMulti: TQRLabel;
    TSolMulti: TQRLabel;
    TNbPaie: TQRLabel;
    NbPaieMulti: TQRLabel;
    TCumMulti: TQRLabel;
    TLibMulti: TQRLabel;
    QRLabel36: TQRLabel;
    TotMontantG: TQRLabel;
    TotCumulG: TQRLabel;
    TotDetCredit: TQRLabel;
    NbGene: TQRLabel;
    QRBand1: TQRBand;
    Trait0: TQRLigne;
    Trait1: TQRLigne;
    Trait2: TQRLigne;
    Trait4: TQRLigne;
    Ligne1: TQRLigne;
    TITRE2REP: TQRLabel;
    REPORT2MONTANT: TQRLabel;
    MsgLibel: THMsgBox;
    DLEcheH: TQRDetailLink;
    QRDLMulti: TQRDetailLink;
    QModP: TQuery;
    QModPMAX: TStringField;
    QModPDEBIT: TFloatField;
    QModPCREDIT: TFloatField;
    QModPCOUNT: TIntegerField;
    FRecapPaie: TCheckBox;
    FRuptPaie: TCheckBox;
    TFGenTCTD: THLabel;
    TFGenTCTDJoker: THLabel;
    TFaTCTD: TLabel;
    TFNatureCptTCTD: THLabel;
    FGenTCTD1: THCpteEdit;
    FGenTCTDJoker: TEdit;
    FNatureCptTCTD: THValComboBox;
    FGenTCTD2: THCpteEdit;
    RColl: TQRLabel;
    RColl1: TQRLabel;
    TRaC: TQRLabel;
    RColl2: TQRLabel;
    QRLabel21: TQRLabel;
    RNatureCptTCTD: TQRLabel;
    QRBandGG: TQRBand;
    QRDLGG: TQRDetailLink;
    CodMP: TQRLabel;
    QRSUBTOT: TQRBand;
    QRLabel6: TQRLabel;
    QRLabel20: TQRLabel;
    NbGene2: TQRLabel;
    TotMontantG2: TQRLabel;
    TotCumulG2: TQRLabel;
    LibMP: TQRLabel;
    TotRecapMP: TQRLabel;
    NbMP: TQRLabel;
    FSautRecapMP: TCheckBox;
    QRLabel1: TQRLabel;
    FLigneEchTete: TCheckBox;
    E_GENERAL: TQRLabel;
    E_AUXILIAIRE: TQRLabel;
    T_LIBELLE: TQRLabel;
    FEcheancier: THValComboBox;
    HLabel3: THLabel;
    FJal1: THCpteEdit;
    Label1: TLabel;
    FJal2: THCpteEdit;
    FAvecLibEcr: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure FColl1Change(Sender: TObject);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BEcheHBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BSDMultiBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure QRDLMultiNeedData(var MoreData: Boolean);
    procedure DLEcheHNeedData(var MoreData: Boolean);
    procedure FEcheancierChange(Sender: TObject);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure FGenTCTD1Change(Sender: TObject);
    procedure QRDLGGNeedData(var MoreData: Boolean);
    procedure QRBandGGAfterPrint(BandPrinted: Boolean);
    procedure EntetePageAfterPrint(BandPrinted: Boolean);
    procedure FRecapPaieClick(Sender: TObject);
    procedure EntetePageBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FNatureCptChange(Sender: TObject);
    procedure QRSUBTOTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FNatureCptTCTDChange(Sender: TObject);
    procedure BSDMultiAfterPrint(BandPrinted: Boolean);
    procedure FEditionChange(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure GenereSQL ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
  private    { Déclarations privées }
    TotMulti                 : Array[0..12] of Double ;
    CumulD, CumulC, NbGen    : Double ;
    StReportEche             : String ;
    NbEch, NumR              : integer ;
    LMulti, LModePaie        : TStringList ;
    TotGeneral, TotEche      : TabTot ;
    OKEnteteEche, OkEche     : Boolean ;
    ChampCpt                 : TStringField ;
    Origine                  : TSuiviMP ;
    QR2G_LIBELLE,
    QR2G_NATUREGENE,
    QR2E_REFINTERNE,
    QR2E_AUXILIAIRE,
    QR2E_GENERAL,
    QR2E_ETABLISSEMENT,
    QR2E_LIBELLE,
    QR2T_LIBELLE,
    QR2T_NATUREAUXI,
    QR2E_EXERCICE,
    QR2E_VALIDE,
    QR2E_JOURNAL,
    QR2E_MODEPAIE            : TStringField ;
    QR2E_DATECOMPTABLE,
    QR2E_DATEECHEANCE        : TDateTimeField ;
    QR2E_NUMEROPIECE,
    QR2E_NUMLIGNE            : TIntegerField ;
    QR2E_COUVERTURE,
    QR2DEBIT,
    QR2CREDIT, QR2COUVERTURE : TFloatField ;
    Function  QuoiMvt : String ;
    Procedure EcheanZoom(Quoi : String) ;
    function Calculs : Boolean ;
  public    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

procedure Echeancier(Origine : tSuiviMP = smpAucun) ;
var QR : TFQREchean ;
    Edition : TEdition ;
    PP : THPanel ;
    NomFiltre : String ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQREchean.Create(Application) ;
Edition.Etat:=etEche ;
QR.QRP.QRPrinter.OnSynZoom:=QR.EcheanZoom ;
QR.Origine:=Origine ;
{$IFDEF CCMP}
If Origine=smpEncTous Then NomFiltre:='QRECHEANCEEC' Else NomFiltre:='QRECHEANCEDE' ;
{$ELSE}
NomFiltre:='QRECHEANCE' ;
{$ENDIF}
QR.InitType (nbAux,neEch,msRien,NomFiltre,'',TRUE,FALSE,FALSE,Edition) ;
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

procedure EcheancierZoom(Crit : TCritEdt) ;
var QR : TFQREchean ;
    Edition : TEdition ;
BEGIN
QR:=TFQREchean.Create(Application) ;
Edition.Etat:=etEche ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.EcheanZoom ;
 QR.CritEdt:=Crit ;
 QR.InitType (nbAux,neEch,msRien,'QRECHEANCE','',FALSE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

Function TFQREchean.QuoiMvt : String ;
BEGIN
Result:=QR2E_AUXILIAIRE.AsString+' '+QR2T_LIBELLE.AsString+' '+
        '#'+QR2E_JOURNAL.AsString+' N° '+IntToStr(QR2E_NUMEROPIECE.AsInteger)+' '+DateToStr(QR2E_DateComptable.AsDAteTime)+'-'+
         //PrintSolde(QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,CritJustif.Decimale,CritJustif.Symbole,CritJustif.AfficheSymbole)+' '+
         //QR2E_LETTRAGE.AsString+
         '@'+'5;'+QR2E_JOURNAL.AsString+';'+UsDateTime(QR2E_DATECOMPTABLE.AsDateTime)+';'+QR2E_NUMEROPIECE.AsString+';'+QR2E_EXERCICE.asString+';'+
         IntToStr(QR2E_NumLigne.AsInteger)+';' ;
END ;

Procedure TFQREchean.EcheanZoom(Quoi : String) ;
Var Lp,i: Integer ;
BEGIN
LP:=Pos('@',Quoi) ; If LP=0 Then Exit ;
i:=StrToInt(Copy(Quoi,LP+1,1)) ;
If (i=5) Then
   BEGIN
   Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
   If QRP.QRPRinter.FSynShiftDblClick Then i:=6 ;
   If QRP.QRPrinter.FSynCtrlDblClick Then i:=11 ;
   END ;
ZoomEdt(i,Quoi) ;
END ;

procedure TFQREchean.FinirPrint;
{$IFNDEF CCMP}
Var Solde : double ;
{$ENDIF}
BEGIN
Inherited ;
VideGroup(LMulti) ;
VideRecap(LModePaie) ;
{$IFNDEF CCMP}
if OkMajEdt Then
   BEGIN
   Solde:=TotGeneral[1].TotDebit-TotGeneral[1].TotCredit ;
   if Solde<0
      then MajEdition('ECH', CritEdt.Exo.Code, DateToStr(CritEdt.Date1), DateToStr(CritEdt.Date2),'', TotGeneral[1].TotDebit, TotGeneral[1].TotCredit, Solde, 0)
      else MajEdition('ECH', CritEdt.Exo.Code, DateToStr(CritEdt.Date1), DateToStr(CritEdt.Date2),'', TotGeneral[1].TotDebit, TotGeneral[1].TotCredit, 0, Solde) ;
   END ;
{$ENDIF}
END ;

procedure TFQREchean.InitDivers ;
BEGIN
Inherited ;
//BEcheF.Frame.DrawBottom:=Not (CritEdt.Edition=3) and Not FSautPage.checked ;
Case CritEdt.Ech.Edition of
  0,1 : BEGIN  { Mvts, Mvts+cumuls}
        TE_JOURNAL.Caption:=MsgLibel.Mess[29] ; TE_PIECELIGNE.Caption:=MsgLibel.Mess[30] ;
        TE_GENERAL.Caption:=MsgLibel.Mess[31] ; TE_AUXILIAIRE.Caption:=MsgLibel.Mess[32] ;
        TT_LIBELLE.Caption:=MsgLibel.Mess[33] ; TE_REFINTERNE.Caption:=MsgLibel.Mess[17] ;
        TE_DATECOMPTABLE.Caption:=MsgLibel.Mess[18] ; TE_MODEPAIE.Caption:=MsgLibel.Mess[19] ;
        END ;
  2,3 : BEGIN  {Cumul/tiers, Cumul/Date}
        TE_JOURNAL.Caption:='' ; TE_PIECELIGNE.Caption:='' ;
        TE_GENERAL.Caption:='' ; TE_AUXILIAIRE.Caption:='' ;
        TT_LIBELLE.Caption:='' ; TE_REFINTERNE.Caption:='' ;
        TE_DATECOMPTABLE.Caption:='' ; TE_MODEPAIE.Caption:='' ;
        END ;
   End ;
TRLegende.Visible:=(CritEdt.Ech.Edition=0)or(CritEdt.Ech.Edition=1) ;
RLegende.Visible:=TRLegende.Visible ;
BEcheH.ForceNewPage:=(CritEdt.SautPage=1) ;
BEcheH.Frame.DrawBottom:=CritEDT.Jal.FormatPrint.PrSepCompte[2] ;
NbGen:=0 ;
{ Initialisation des variables de cumul }
CumulD:=0 ; cumulC:=0 ;
{ Gestion de l'apparition de l'entête de compte général }
OKEnteteEche:=True ; OKEche:=True ; QRBandGG.ForceNewPage:=CritEdt.Ech.NewPageRecapMP ;
END ;

procedure TFQREchean.GenereSQL ;
Var WhereGen,WhereAux,WhereColl,St,StSupCouv : String ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
{ Construction de la clause Select de la SQL }
Q.Close ;
Q.SQL.Clear ;
Q.SQL.Add('Select E_AUXILIAIRE, E_GENERAL, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_VALIDE, ') ;
Q.SQL.Add(       'E_REFINTERNE, E_ETABLISSEMENT, E_LIBELLE, E_JOURNAL, T_LIBELLE, ') ;
Q.SQL.Add(       'E_DATEECHEANCE, E_MODEPAIE, E_EXERCICE, E_COUVERTURE, T_NATUREAUXI, G_NATUREGENE, G_LIBELLE, G_LETTRABLE, G_COLLECTIF, T_LETTRABLE, ') ;
StSupCouv:='' ;
Case CritEdt.Monnaie of
  0 : BEGIN
      Q.SQL.Add(' E_DEBIT as DEBIT, E_CREDIT as CREDIT, E_COUVERTURE as COUVERTURE ') ;
      StSupCouv:='AND (NOT (E_ETATLETTRAGE="PL" AND ((E_DEBIT+E_CREDIT)=E_COUVERTURE)))'
      END ;
  1 : BEGIN
      Q.SQL.Add(' E_DEBITDEV as DEBIT, E_CREDITDEV as CREDIT, E_COUVERTUREDEV as COUVERTURE ')   ;
      StSupCouv:='AND (NOT (E_ETATLETTRAGE="PL" AND ((E_DEBITDEV+E_CREDITDEV)=E_COUVERTUREDEV)))'
      END ;
end ;
{ Tables explorées par la SQL }
Q.SQL.Add(' From ECRITURE') ;
{ Construction de la clause Left Join de la SQL }
Q.SQL.Add(' Left Join TIERS on E_AUXILIAIRE=T_AUXILIAIRE ') ;
Q.SQL.Add(' Left Join GENERAUX on E_GENERAL=G_GENERAL ') ;
{ Construction de la clause Where de la SQL }
Q.SQL.Add(' Where E_DATEECHEANCE>="'+USDateTime(CritEdt.Date1)+'" And E_DATEECHEANCE<="'+USDateTime(CritEdt.Date2)+'" ') ;
if VH^.ExoV8.Code<>'' then Q.SQL.Add(' And E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'" ') ;
WhereAux:='' ;

if CritEdt.Joker then
   BEGIN
   WhereAux:=' And E_AUXILIAIRE like "'+TraduitJoker(CritEdt.Cpt1)+'" ';
   END Else
   BEGIN
   if CritEdt.Cpt1<>'' then WhereAux:=WhereAux+' And E_AUXILIAIRE>="'+CritEdt.Cpt1+'" ' ;
   if CritEdt.Cpt2<>'' then WhereAux:=WhereAux+' And E_AUXILIAIRE<="'+CritEdt.Cpt2+'" ' ;
   END ;
//if CritEdt.NatureCpt<>'' then WhereAux:=WhereAux+' And T_NATUREAUXI="'+CritEdt.NatureCpt+'" ' ;

{ Rony 25/03/97 }
Case CritEdt.Ech.Echeancier of
  0 : BEGIN
      if CritEdt.NatureCpt<>'' then WhereAux:=WhereAux+' And T_NATUREAUXI="'+CritEdt.NatureCpt+'" ' ;
      END ;
  1 : BEGIN
      if CritEdt.NatureCpt<>'' then WhereAux:=WhereAux+' And T_NATUREAUXI="'+CritEdt.NatureCpt+'" '
                               Else WhereAux:=WhereAux+' And (T_NATUREAUXI="FOU" or T_NATUREAUXI="AUC" or T_NATUREAUXI="DIV" )   ' ;
      END ;
  2 : BEGIN
      if CritEdt.NatureCpt<>'' then WhereAux:=WhereAux+' And T_NATUREAUXI="'+CritEdt.NatureCpt+'" '
                               Else WhereAux:=WhereAux+' And (T_NATUREAUXI="CLI" or T_NATUREAUXI="AUD" or T_NATUREAUXI="DIV" )   ' ;
      END ;
  End ;
{ Fin 25/03/97 }

WhereColl:='' ;
if CritEdt.Ech.SSJoker then
   BEGIN
   WhereColl:=' And G_GENERAL like "'+TraduitJoker(CritEdt.ECH.Cpt3)+'" ';
   END Else
   BEGIN
   if CritEdt.Ech.Cpt3<>'' then WhereColl:=WhereColl+' And G_GENERAL>="'+CritEdt.Ech.Cpt3+'" ' ;
   if CritEdt.Ech.Cpt4<>'' then WhereColl:=WhereColl+' And G_GENERAL<="'+CritEdt.Ech.Cpt4+'" ' ;
   END ;
WhereColl:=WhereColl+' And G_COLLECTIF="X" ' ;

WhereAux:=Trim(WhereAux+WhereColl) ;
WhereAux:='('+Copy(WhereAux,4,Length(WhereAux)-3)+')' ;

WhereGen:='' ;
if CritEdt.SJoker then
   BEGIN
   WhereGen:=' And E_GENERAL like "'+TraduitJoker(CritEdt.SCpt1)+'" ';
   END Else
   BEGIN
   if CritEdt.SCpt1<>'' then WhereGen:=WhereGen+' And E_GENERAL>="'+CritEdt.SCpt1+'" ' ;
   if CritEdt.SCpt2<>'' then WhereGen:=WhereGen+' And E_GENERAL<="'+CritEdt.SCpt2+'" ' ;
   END ;
if CritEdt.Ech.NatureSCpt<>'' then WhereGen:=WhereGen+' And G_NATUREGENE="'+CritEdt.Ech.NatureSCpt+'" '
                               Else WhereGen:=WhereGen+' And (G_NATUREGENE="TID" Or G_NATUREGENE="TIC") ' ;

WhereGen:=Trim(WhereGen) ;
WhereGen:='('+Copy(WhereGen,4,Length(WhereGen)-3)+')' ;

St:=' And (('+WhereAux+') Or ('+WhereGen+'))' ;

Q.SQL.Add(St) ;

St:='' ;
if CritEdt.Ech.Jal1<>'' then St:=St+' And E_JOURNAL>="'+CritEdt.Ech.Jal1+'" ' ;
if CritEdt.Ech.Jal2<>'' then St:=St+' And E_JOURNAL<="'+CritEdt.Ech.Jal2+'" ' ;
If St<>'' Then Q.SQL.Add(St) ;

Q.SQL.Add(TraduitNatureEcr(CritEdt.Qualifpiece, 'E_QUALIFPIECE', true, cbUnchecked(*AvecRevision.State*))) ;
if CritEdt.Etab<>'' then Q.SQL.Add(' And E_ETABLISSEMENT="'+CritEdt.Etab+'"') ;
if CritEdt.DeviseSelect<>'' then Q.SQL.Add(' And E_DEVISE="'+CritEdt.DeviseSelect+'"') ;
Q.SQL.Add(' And E_ETATLETTRAGE<>"RI" And E_ETATLETTRAGE<>"TL"') ;
Q.SQL.Add(' And E_ECRANOUVEAU<>"CLO" And E_ECRANOUVEAU<>"OAN" and E_QUALIFPIECE<>"C" ') ;
if (OkV2 and (V_PGI.Confidentiel<>'1')) then Q.SQL.Add('AND E_CONFIDENTIEL<>"1" ') ;
Case CritEdt.Ech.Sens of
  0 : BEGIN
      Case CritEdt.Monnaie of
        0 : BEGIN Q.SQL.Add(' And E_CREDIT<>0 ' ) ;    END ;
        1 : BEGIN Q.SQL.Add(' And E_CREDITDEV<>0 ') ;  END ;
       end ;
      END ;
  1 : BEGIN
      Case CritEdt.Monnaie of
        0 : BEGIN Q.SQL.Add(' And E_DEBIT<>0 ' ) ;    END ;
        1 : BEGIN Q.SQL.Add(' And E_DEBITDEV<>0 ') ;  END ;
       end ;
      END ;
 end ;
If StSupCouv<>'' Then Q.SQL.add(StSupCouv) ;
If CritEDT.SQLPlus<>'' Then Q.SQL.Add(CritEDT.SQLPlus) ;
Q.SQL.Add(E_MODEPAIE.GetSQLValue) ;
{ Construction de la clause Order By de la SQL }
if CritEdt.Ech.RuptPaie then Q.SQL.Add(' Order By E_DATEECHEANCE, E_MODEPAIE, E_AUXILIAIRE, E_GENERAL ')
                        Else Q.SQL.Add(' Order By E_DATEECHEANCE, E_AUXILIAIRE, E_GENERAL') ;
ChangeSQL(Q) ; Q.Open ;
END ;

procedure TFQREchean.RenseigneCritere ;
Var St11,St22, St3, St4 : String ;
BEGIN
Inherited ;
if CritEdt.SJoker then
   BEGIN
   RGen.Caption:=MsgLibel.Mess[25] ;
   RGen1.Caption:=FGenTCTDJoker.Text
   END Else
   BEGIN
   RGen.Caption:=MsgLibel.Mess[24] ;
   //RGen1.Caption:=FGenTCTD1.Text ; RGen2.Caption:=FGenTCTD2.Text ;
   PositionneFourchetteST(FGenTCTD1,FGenTCTD2,St11,St22) ;
   RGen1.Caption:=St11 ;
   RGen2.Caption:=St22 ;
   END ;
TRaG.Visible:=Not CritEdt.SJoker ; RGen2.Visible:=Not CritEdt.SJoker ;

if CritEdt.Ech.SSjoker then
   BEGIN
   RColl.Caption:=MsgLibel.Mess[27] ;
   RColl1.Caption:=FCollJoker.Text
   END Else
   BEGIN
   RColl.Caption:=MsgLibel.Mess[26] ;
   //RColl1.Caption:=FColl1.Text ; RColl2.Caption:=FColl2.Text ;
   PositionneFourchetteST(FColl1,FColl2,St3,St4) ;
   RColl1.Caption:=St3 ;
   RColl2.Caption:=St4 ;
   END ;
TRaC.Visible:=Not CritEdt.Ech.SSjoker ; RColl2.Visible:=Not CritEdt.Ech.SSjoker ;

RNatureCptTCTD.Caption:=FNatureCptTCTD.Text ;
REcheancier.Caption:=FEcheancier.Text   ; RSens.Caption:=FSens.Text ;
RPaie.Caption:=E_MODEPAIE.Text          ; REdition.Caption:=FEdition.Text ;
END ;

procedure TFQREchean.ChoixEdition ;
BEGIN
Inherited ;
ChargeGroup(LMulti,['ECHE','PAIE','AUXI']) ;
ChargeRecap(LModePaie) ;
END ;

procedure TFQREchean.RecupCritEdt ;
Var St3, St4 : String ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  Ech.ModePaie:='' ;
  SJoker:=FGenTCTDJoker.Visible ; Ech.SSjoker:=FCollJoker.Visible ;
  if SJoker Then
     BEGIN
     SCpt1:=FGenTCTDJoker.Text ; SCpt2:=FGenTCTDJoker.Text ;
     LSCpt1:=SCpt1 ; LSCpt2:=SCpt2 ;
     END Else
     BEGIN
     SCpt1:=FGenTCTD1.Text ;  SCpt2:=FGenTCTD2.Text ;
     PositionneFourchetteSt(FGenTCTD1,FGenTCTD2,CritEdt.LSCpt1,CritEdt.LSCpt2) ;
     END ;
  if Ech.SSjoker  Then
     BEGIN
     Ech.Cpt3:= FCollJoker.Text ; Ech.Cpt4:= FCollJoker.Text ;
     Ech.LCpt3:=Ech.Cpt3 ; Ech.LCpt4:=Ech.Cpt4 ;
     END Else
     BEGIN
     Ech.Cpt3:=FColl1.Text ; Ech.Cpt4:=FColl2.Text ;
     St3:='' ; St4:='' ;
     PositionneFourchetteSt(FColl1,FColl2,St3,St4) ;
     CritEdt.Ech.LCpt3:=St3 ; CritEdt.Ech.LCpt4:=St4 ;
     END ;
 (* if FEcheancier.ItemIndex<>0 then
     BEGIN
     NatureCpt:=FNatureCpt.Value ;
     END else if FNatureCpt.ItemIndex<>0 then NatureCpt:=FNatureCpt.Value ;
  *)

  { Rony 25/03/97 }
  if FNatureCpt.ItemIndex<>0 then NatureCpt:=FNatureCpt.Value ;
  { Fin 25/03/97 }

  Ech.RecapMP:=FRecapPaie.Checked ;
  If FNatureCptTCTD.ItemIndex<>0 Then Ech.NatureSCpt:=FNatureCptTCTD.Value ;
  //if FModePaie.ItemIndex<>0 then ModePaie:=FModePaie.Value ;
  Ech.RuptPaie:=FRuptPaie.Checked ; Ech.NewPageRecapMP:=FSautRecapMP.Checked ;
  Ech.Echeancier:=FEcheancier.ItemIndex ;
  Ech.Edition:=FEdition.ItemIndex ;
  Ech.Sens:=FSens.ItemIndex ;
  if FSautPage.Checked then SautPage:=1 else  SautPage:=0 ;
  Ech.Jal1:=Trim(FJal1.Text) ; Ech.Jal2:=Trim(FJal2.Text) ;
  Ech.AvecLibEcr:=FAvecLibEcr.Checked ;
  With Ech.FormatPrint Do
       BEGIN
       PrSepCompte[2]:=FLigneEchTete.Checked ;
       PrSepCompte[3]:=FLigneEchPied.Checked ;
       END ;
  END ;
END ;

function  TFQREchean.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ;
If Result Then
   BEGIN
   Fillchar(TotGeneral,SizeOf(TotGeneral),#0) ;
   END ;
END ;

function TFQREchean.Calculs : Boolean ;
Var MReport : TabTRep ;
    Reste   : Double ;
BEGIN
FillChar(MReport,Sizeof(MReport),#0);

if QR2DEBIT.AsFloat<>0 then
   BEGIN
   Reste:=QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat ;
   AddGroup(LMulti, [QR2E_DATEECHEANCE, QR2E_MODEPAIE, ChampCpt], [1, Reste, QR2CREDIT.AsFloat]) ;
//   AddGroup(LMulti, [QR2E_DATEECHEANCE, QR2E_AUXILIAIRE, QR2E_MODEPAIE ], [1, Reste, QR2CREDIT.AsFloat]) ;
   LE_MONTANT.Caption:=PrintSolde(Reste, QR2CREDIT.AsFloat, CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;
   TotEche[1].TotDebit:=   Arrondi(TotEche[1].TotDebit+Reste, CritEdt.Decimale) ;
   TotGeneral[1].TotDebit:=Arrondi(TotGeneral[1].TotDebit+Reste, CritEdt.Decimale) ;
   MReport[1].TotDebit:=   Arrondi(MReport[1].TotDebit+Reste,  CritEdt.Decimale) ;
   MReport[2].TotDebit:=   Arrondi(MReport[2].TotDebit+Reste,  CritEdt.Decimale) ;
   MReport[3].TotDebit:=   Arrondi(MReport[3].TotDebit+CumulD, CritEdt.Decimale) ;
   END else
   BEGIN
   Reste:=QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat ;
   AddGroup(LMulti, [QR2E_DATEECHEANCE, QR2E_MODEPAIE, ChampCpt], [1, QR2DEBIT.AsFloat, Reste]) ;
//   AddGroup(LMulti, [QR2E_DATEECHEANCE, QR2E_AUXILIAIRE, QR2E_MODEPAIE], [1, QR2DEBIT.AsFloat, Reste]) ;
   LE_MONTANT.Caption:=PrintSolde(QR2DEBIT.AsFloat, Reste, CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;
   TotEche[1].TotCredit:=   Arrondi(TotEche[1].TotCredit+Reste, CritEdt.Decimale) ;
   TotGeneral[1].TotCredit:=Arrondi(TotGeneral[1].TotCredit+Reste, CritEdt.Decimale) ;
   MReport[1].TotCredit:=   Arrondi(MReport[1].TotCredit+Reste,  CritEdt.Decimale) ;
   MReport[2].TotCredit:=   Arrondi(MReport[2].TotCredit+Reste,  CritEdt.Decimale) ;
   MReport[3].TotCredit:=   Arrondi(MReport[3].TotCredit+CumulC, CritEdt.Decimale) ;
   END ;
if (CritEdt.Ech.Edition=3) then
   BEGIN                    {type d'édition : Cumul/Dates }
   AddReportBAL([1],CritEdt.Ech.FormatPrint.Report,MReport,CritEdt.Decimale) ;
   END Else
   BEGIN                    {type d'édition : Mvts ou Mvts + Tiers ou Cumul/Tiers }
   AddReportBAL([1,2],CritEdt.Ech.FormatPrint.Report,MReport,CritEdt.Decimale) ;
   END ;
if Reste=0 then Inc(NbEch) ;
Result:=Not(Reste=0) ;
END ;

procedure TFQREchean.FormShow(Sender: TObject);
begin
HelpContext:=7541000 ;
//Standards.HelpContext:=7541010 ;
//Avances.HelpContext:=7541020 ;
//Mise.HelpContext:=7541030 ;
//Option.HelpContext:=7541040 ;

FEcheancier.ItemIndex:=0 ; FSens.ItemIndex:=2 ;
FNatureCptTCTD.ItemIndex:=0 ;
FEdition.ItemIndex:=0 ; E_MODEPAIE.Text:=Traduirememoire('<<Tous>>') ;
{$IFDEF CCMP}
FEcheancier.Enabled:=FALSE ;
If Origine=smpEncTous Then FEcheancier.Value:='ENC' Else FEcheancier.Value:='DEC' ;
FNatureCpt.Vide := False;
if (VH^.CCMP.LotCli) then begin FNatureCpt.Plus := ' AND (CO_CODE="AUD" OR CO_CODE="CLI" OR CO_CODE="DIV")'; FNatureCpt.Value:='CLI'; end
                     else begin FNatureCpt.Plus := ' AND (CO_CODE="AUC" OR CO_CODE="DIV" OR CO_CODE="FOU" OR CO_CODE="SAL")'; FNatureCpt.Value:='FOU'; end;
{$ENDIF}
  inherited;
TabSup.TabVisible:=False;
//FColl1.Text:='' ; FColl2.Text:='' ;
end;

procedure TFQREchean.FColl1Change(Sender: TObject);
Var AvecJokerS : Boolean ;
begin
  inherited;
AvecJokerS:=Joker(FColl1, FColl2, FCollJoker) ;
TFaC.Visible:=Not AvecJokerS ;
FColl.Visible:=Not AvecJokerS ;
TFCollJoker.Visible:=AvecJokerS ;
end;

procedure TFQREchean.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TITRE1REP.Tag:=TE_DATECOMPTABLE.Tag ;
TITRE1REP.Caption:=TITRE2REP.Caption ;
REPORT1MONTANT.Caption:=REPORT2MONTANT.Caption ;
TITREREPORTH.Caption:='' ;
//REPORT1CUMULE.Caption:=REPORT2CUMULE.Caption ;
end;

procedure TFQREchean.BEcheHBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TEcheH.Caption:=MsgLibel.Mess[21]+' '+DateToStr(QR2E_DATEECHEANCE.AsDateTime) ;
Fillchar(TotEche,SizeOf(TotEche),#0) ;
StReportEche:=DateToStr(QR2E_DATEECHEANCE.AsDateTime) ;
InitReport([2],CritEdt.Ech.FormatPrint.Report) ;
NbEch:=0 ;
PrintBand:=Not (CritEdt.Ech.Edition=3)
           or (Not (CritEdt.Ech.Edition=3) or CritEdt.Ech.RuptPaie );
end;

procedure TFQREchean.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
// XX OkEche:=True ;

If ChampCpt<>Nil then ChampCpt:=Nil ;
IF QR2E_AUXILIAIRE.AsString<>'' then ChampCpt:=TStringField(Q.FindField('E_AUXILIAIRE'))
                                Else ChampCpt:=TStringField(Q.FindField('E_GENERAL')) ;

PrintBand:=(Calculs) And (Not((CritEdt.Ech.Edition=3) Or (CritEdt.Ech.Edition=2))) ;
// XX if ((CritEdt.Ech.Edition=0)or(CritEdt.Ech.Edition=1)) then OKEche:=PrintBand ;
If Not PrintBand then Exit Else Quoi:=QuoiMvt ;
// RR
if (QR2E_AUXILIAIRE.AsString='') then
   BEGIN
   E_AUXILIAIRE.Caption:=QR2E_GENERAL.AsString ;
   E_GENERAL.Caption:='' ;
   If CritEdt.Ech.AvecLibEcr And (QR2E_LIBELLE.AsString<>'') Then T_LIBELLE.Caption:=QR2E_LIBELLE.AsString
                                                             Else T_LIBELLE.Caption:=QR2G_LIBELLE.AsString ;
   END Else
   BEGIN
   E_AUXILIAIRE.Caption:=QR2E_AUXILIAIRE.AsString ;
   E_GENERAL.Caption:=QR2E_GENERAL.AsString ;
   If CritEdt.Ech.AvecLibEcr And (QR2E_LIBELLE.AsString<>'') Then T_LIBELLE.Caption:=QR2E_LIBELLE.AsString
                                                             Else T_LIBELLE.Caption:=QR2T_LIBELLE.AsString ;
   END ;
LE_PIECELIG.Caption:=QR2E_NUMEROPIECE.AsString+'/'+QR2E_NUMLIGNE.AsString ;
if QR2E_VALIDE.AsString='X' then LEVALIDE.Caption:='V' else LEVALIDE.Caption:=' ' ;
LE_MODEPAIE.Caption:=RechDom('ttModepaie',QR2E_MODEPAIE.Value,False) ;
{$B+}
//if ((CritEdt.Ech.Edition=0)or(CritEdt.Ech.Edition=1)) then OKEche:=PrintBand ;
{$B-}
LE_CUMUL.Caption:=PrintSolde(TotEche[1].TotDebit,TotEche[1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;

end;

procedure TFQREchean.BSDMultiBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
Case NumR of
 0 : BEGIN       { Total Echeance }
     PrintBand:=OkEche ;
     CumulD:=CumulD+TotEche[1].TotDebit ;
     CumulC:=CumulC+TotEche[1].TotCredit ;
     TCumMulti.Caption:=  PrintSolde(CumulD, CumulC, CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;
     NbGen:=NbGen+(TotMulti[0]-NbEch)  ;
     END ;
 1 : BEGIN       { Total Paiement }
     PrintBand:=CritEdt.Ech.RuptPaie ;
     END ;
 2 : BEGIN      { Total Auxiliaire }
     PrintBand:=(CritEdt.Ech.Edition=1) or (CritEdt.Ech.Edition=2) ;
     END ;
 end;
end;

procedure TFQREchean.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
{ Affichage montants finaux }
NbGene.Caption:=floatToStr(NbGen) ;
TotCumulG.Caption:=PrintSolde(CumulD, CumulC, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
TotMontantG.Caption:=PrintSolde(TotGeneral[1].TotDebit,TotGeneral[1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFQREchean.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var MReport        : TabTRep ;
begin
  inherited;
TITREREPORTB.Caption:='' ;
TITRE2REP.Tag:=TE_DATECOMPTABLE.Tag ;
Fillchar(MReport,SizeOf(MReport),#0) ;
Case QuelReportBAL(CritEdt.Ech.FormatPrint.Report,MReport) Of
  1 : BEGIN
      Titre2Rep.Caption:=MsgLibel.Mess[5] ;
      { Cas particulier pour le Montant Cumulé en Report Géné. }
      //Report2CUMULE.Caption:=PrintSolde(CumulD,CumulC, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
      //Report2CUMULE.Caption:=PrintSolde(MReport[3].TotDebit, MReport[3].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
      END ;
  2 : BEGIN
      Titre2Rep.Caption:=MsgLibel.Mess[6]+' '+StReportEche ;
      //Report2CUMULE.Caption:=PrintSolde(MReport[2].TotDebit,MReport[2].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
      END ;
  END ;
Report2MONTANT.Caption:=PrintSolde(MReport[1].TotDebit, MReport[1].TotCredit, CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;
//Report2CUMULE.Caption:=PrintSolde(MReport[3].TotDebit, MReport[3].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
end;

procedure TFQREchean.QRDLMultiNeedData(var MoreData: Boolean);
var LibMulti, CodMulti : string ;
begin
  inherited;
MoreData:=PrintGroup(LMulti, Q, [QR2E_DATEECHEANCE,QR2E_MODEPAIE, ChampCpt], CodMulti, LibMulti, TotMulti, NumR) ;
//MoreData:=PrintGroup2(LMulti, Q, [QR2E_DATEECHEANCE,QR2E_AUXILIAIRE, QR2E_MODEPAIE], CodMulti, LibMulti, TotMulti, NumR) ;
if ((MoreData) and (Not Q.EOF) and (NumR=0)) then OKEnteteEche:=True ;
If Not MoreData then exit ;
TCodMulti.Tag:=5 ; //TLibMulti.AutoSize:=False ;
BSDMulti.Frame.DrawTop:=False ;
BSDMulti.Frame.DrawBottom:=False ;
Case NumR of
  0 : BEGIN { Eche }
      BSDMulti.Frame.DrawTop:=CritEdt.GL.FormatPrint.PrSepCompte[3] ;
      BSDMulti.Frame.DrawBottom:=CritEdt.GL.FormatPrint.PrSepCompte[1] ;
      BSDMulti.Font.Color:=clNavy ;
      BSDMulti.Font.Style:=[fsBold] ; 
      TLibMulti.Caption:=MsgLibel.Mess[7]+' '+CodMulti ; TLibMulti.Alignment:=taRightJustify ;
      if(CritEdt.Ech.Edition=3)and Not CritEdt.Ech.RuptPaie then
         BEGIN
         TLibMulti.Left:=TE_JOURNAL.Left ; //TLibMulti.AutoSize:=True ;
         TLibMulti.Tag:=TE_JOURNAL.Tag ; TLibMulti.Alignment:=taLeftJustify ;
         END Else
         BEGIN
         TLibMulti.Left:=TE_DATECOMPTABLE.Left ; TLibMulti.Width:=TE_DATECOMPTABLE.Width+TE_MODEPAIE.Width+1 ;
         TLibMulti.Tag:=8 ;
         END ;
      TCodMulti.Caption:='' ;
      END ;
  1 : BEGIN { Mode Paie }
      BSDMulti.Font.Color:=clRed ;
      BSDMulti.Font.Style:=[fsBold] ;
      TLibMulti.Caption:=MsgLibel.Mess[8] ; TLibMulti.Left:=18 ;  TLibMulti.Tag:=2 ;
      TCodMulti.Caption:=CodMulti ; TCodMulti.Alignment:=taLeftJustify ;
      TCodMulti.Left:=106 ; TCodMulti.Tag:=4 ;
      TCumMulti.Caption:='' ;
      AddRecap(LModePaie, [QR2E_MODEPAIE.AsString], [QR2E_MODEPAIE.AsString],[TotMulti[0]-NbEch,TotMulti[1], TotMulti[2]]) ;
      END ;
  2 : BEGIN { Auxi }
      BSDMulti.Font.Color:=clGreen ;
      BSDMulti.Font.Style:=[fsBold] ;
      TLibMulti.Caption:=MsgLibel.Mess[9] ; TLibMulti.Left:=18 ; TLibMulti.Tag:=2 ;
      TCodMulti.Caption:=CodMulti ; TCodMulti.Alignment:=taLeftJustify ;
      TCodMulti.Left:=106 ; TCodMulti.Tag:=4 ;
      TCumMulti.Caption:='' ;
      END ;
  end ;
TSolMulti.Caption:=PrintSolde(TotMulti[1], TotMulti[2], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;
NbPaieMulti.Caption:=FloatToStr(TotMulti[0]-NbEch) ;

end;

procedure TFQREchean.DLEcheHNeedData(var MoreData: Boolean);
begin
  inherited;
MoreData:=OKEnteteEche ;
OKEnteteEche:=False ;
end;

procedure TFQREchean.FEcheancierChange(Sender: TObject);
begin
//If QRLoading then Exit ;
  inherited;
If Not QRLoading then
   BEGIN
   FCpte1.Clear ; FCpte2.Clear ; FJoker.Clear ;
   FGenTCTD1.Clear ; FGenTCTD2.Clear ; FGenTCTDJoker.Clear ;
   FColl1.Clear ; FColl2.Clear ; FCollJoker.Clear ; // Rony 06/06/97
   END ;
Case FEcheancier.ItemIndex of
  0 : BEGIN { Tous }
      FCpte1.ZoomTable:=tzTiers ;   FCpte2.ZoomTable:=tzTiers ;
      FGenTCTD1.ZoomTable:=tzGTIDTIC ; FGenTCTD2.ZoomTable:=tzGTIDTIC ;
      FColl1.ZoomTable:=tzGCollectif ; FColl2.ZoomTable:=tzGCollectif ;
      FNatureCptTCTD.ItemIndex:=0 ;
      FNatureCpt.Vide:=True ; FNatureCpt.Reload ;
      FNatureCptTCTD.Enabled:=True ; FNatureCpt.ItemIndex:=0 ;
      END ;
  1 : BEGIN { Décaissement }
      FCpte1.ZoomTable:=tzTfourn ;  FCpte2.ZoomTable:=tzTfourn ;
      FGenTCTD1.ZoomTable:=tzGTIC ; FGenTCTD2.ZoomTable:=tzGTIC ;
      FColl1.ZoomTable:=tzGCollToutCredit ; FColl2.ZoomTable:=tzGCollToutCredit ;
      FNatureCpt.Vide:=False ; FNatureCpt.Items.Clear ; FNatureCpt.Values.Clear ;
      FNatureCpt.Items.Add(MsgLibel.Mess[16]) ; FNatureCpt.Values.Add('') ;
      FNatureCpt.Items.Add(MsgLibel.Mess[11]) ; FNatureCpt.Values.Add('FOU') ;
      FNatureCpt.Items.Add(MsgLibel.Mess[12]) ; FNatureCpt.Values.Add('AUC') ;
      FNatureCpt.Items.Add(MsgLibel.Mess[15]) ; FNatureCpt.Values.Add('DIV') ;
      FNatureCpt.ItemIndex:=0 ;
      FNatureCptTCTD.ItemIndex:=1 ; FNatureCptTCTD.Enabled:=False ;
      END ;
  2 : BEGIN { Encaissement }
      FCpte1.ZoomTable:=tzTclient ; FCpte2.ZoomTable:=tzTclient ;
      FGenTCTD1.ZoomTable:=tzGTID ; FGenTCTD2.ZoomTable:=tzGTID ;
      FColl1.ZoomTable:=tzGCollToutDebit ; FColl2.ZoomTable:=tzGCollToutDebit ;
      FNatureCpt.Vide:=False ; FNatureCpt.Items.clear ; FNatureCpt.Values.Clear ;
      FNatureCpt.Items.Add(MsgLibel.Mess[16]) ; FNatureCpt.Values.Add('') ;
      FNatureCpt.Items.Add(MsgLibel.Mess[13]) ; FNatureCpt.Values.Add('CLI') ;
      FNatureCpt.Items.Add(MsgLibel.Mess[14]) ; FNatureCpt.Values.Add('AUD') ;
      FNatureCpt.Items.Add(MsgLibel.Mess[15]) ; FNatureCpt.Values.Add('DIV') ;
      FNatureCpt.ItemIndex:=0 ;
      FNatureCptTCTD.ItemIndex:=2 ; FNatureCptTCTD.Enabled:=False ;
      END ;
 end ;
//  inherited;
end;

procedure TFQREchean.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
QR2G_LIBELLE        :=TStringField(Q.FindField('G_LIBELLE')) ;
QR2G_NATUREGENE     :=TStringField(Q.FindField('E_NATUREGENE')) ;
QR2E_GENERAL        :=TStringField(Q.FindField('E_GENERAL')) ;
QR2E_ETABLISSEMENT  :=TStringField(Q.FindField('E_ETABLISSEMENT')) ;
QR2E_LIBELLE        :=TStringField(Q.FindField('E_LIBELLE')) ;
QR2T_LIBELLE        :=TStringField(Q.FindField('T_LIBELLE')) ;
QR2T_NATUREAUXI     :=TStringField(Q.FindField('T_NATUREAUXI')) ;
QR2E_EXERCICE       :=TStringField(Q.FindField('E_EXERCICE')) ;
QR2E_AUXILIAIRE     :=TStringField(Q.FindField('E_AUXILIAIRE')) ;
QR2E_MODEPAIE       :=TStringField(Q.FindField('E_MODEPAIE')) ;
QR2E_DATECOMPTABLE  :=TDateTimeField(Q.FindField('E_DATECOMPTABLE')) ;
QR2E_DATEECHEANCE   :=TDateTimeField(Q.FindField('E_DATEECHEANCE')) ;
QR2E_DATECOMPTABLE.Tag:=1 ;
QR2E_NUMEROPIECE    :=TIntegerField(Q.FindField('E_NUMEROPIECE')) ;
QR2E_NUMLIGNE       :=TIntegerField(Q.FindField('E_NUMLIGNE')) ;
QR2E_REFINTERNE     :=TStringField(Q.FindField('E_REFINTERNE')) ;
QR2E_VALIDE         :=TStringField(Q.FindField('E_VALIDE')) ;
QR2E_JOURNAL        :=TStringField(Q.FindField('E_JOURNAL')) ;
QR2E_COUVERTURE     :=TFloatField(Q.FindField('E_COUVERTURE')) ;
QR2COUVERTURE       :=TFloatField(Q.FindField('COUVERTURE')) ;
QR2DEBIT            :=TFloatField(Q.FindField('DEBIT')) ;
QR2CREDIT           :=TFloatField(Q.FindField('CREDIT')) ;
ChgMaskChamp(Qr2DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
ChgMaskChamp(Qr2CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
end;

procedure TFQREchean.FGenTCTD1Change(Sender: TObject);
Var AvecJokerS : Boolean ;
begin
  inherited;
AvecJokerS:=Joker(FGenTCTD1, FGenTCTD2, FGenTCTDJoker) ;
TFaTCTD.Visible:=Not AvecJokerS ;
TFGenTCTD.Visible:=Not AvecJokerS ;
TFGenTCTDJoker.Visible:=AvecJokerS ;
end;

procedure TFQREchean.QRDLGGNeedData(var MoreData: Boolean);
Var Tot : Array[0..12] of Double ;
    Cod,Lib : String ;
begin
  inherited;
if Not CritEdt.Ech.RecapMP then BEGIN MoreData:=FALSE ; exit ; END ;
MoreData:=PrintRecap(LModePaie,Cod,Lib,Tot) ;
If MoreData Then
   BEGIN
   CodMP.Caption:=Cod ; LibMP.Caption:=RechDom('ttModepaie',Cod,False) ;
   NbMP.Caption:=FloatToStr(Tot[0]) ;
   TotRecapMP.Caption:=PrintSolde(Tot[1], Tot[2], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;
   END ;
end;

procedure TFQREchean.QRBandGGAfterPrint(BandPrinted: Boolean);
begin
  inherited;
QRBandGG.ForceNewPage:=FALSE ;
end;

procedure TFQREchean.EntetePageAfterPrint(BandPrinted: Boolean);
begin
  inherited;
if CritEdt.Ech.RecapMp then if Q.Eof then TOPREPORT.Enabled:=False ;
end;

procedure TFQREchean.FRecapPaieClick(Sender: TObject);
begin
  inherited;
FSautRecapMP.Enabled:=FRecapPaie.Checked ;
end;

procedure TFQREchean.EntetePageBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
if (CritEdt.Ech.RecapMp)and(Q.Eof) then
   BEGIN
   TE_JOURNAL.Caption:='' ;   TE_PIECELIGNE.Caption:=MsgLibel.Mess[34] ;
   TE_GENERAL.Caption:= '' ; TE_AUXILIAIRE.Caption:='' ;
   TT_LIBELLE.Caption:= MsgLibel.Mess[35] ; TE_REFINTERNE.Caption:='' ;
   TE_DATECOMPTABLE.Caption:='' ; TE_MODEPAIE.Caption:='' ;
   TE_PIECELIGNE.Alignment:=taLeftJustify ; TT_LIBELLE.Alignment:=taLeftJustify ;
   END Else BEGIN TE_PIECELIGNE.Alignment:=taCenter ; TT_LIBELLE.Alignment:=taCenter ; END ;
end;

procedure TFQREchean.FNatureCptChange(Sender: TObject);
begin
  inherited;
if FEcheancier.Value='' then Exit ;
if FEcheancier.Value='DEC' then
   BEGIN
   FCpte1.ZoomTable:=tzTfourn ;  FCpte2.ZoomTable:=tzTfourn ;
   END Else
if FEcheancier.Value='ENC' then
   BEGIN
   FCpte1.ZoomTable:=tzTclient ; FCpte2.ZoomTable:=tzTclient ;
   END ;
if FNatureCpt.Value='FOU' then BEGIN FCpte1.ZoomTable:=tzTfourn ;  FCpte2.ZoomTable:=FCpte1.ZoomTable ; END ELse
if FNatureCpt.Value='AUC' then BEGIN FCpte1.ZoomTable:=tzTCrediteur ;  FCpte2.ZoomTable:=FCpte1.ZoomTable ;END ELse
if FNatureCpt.Value='CLI' then BEGIN FCpte1.ZoomTable:=tzTclient ;  FCpte2.ZoomTable:=FCpte1.ZoomTable ;END ELse
if FNatureCpt.Value='AUD' then BEGIN FCpte1.ZoomTable:=tzTDebiteur ;  FCpte2.ZoomTable:=FCpte1.ZoomTable ;END ELse
if FNatureCpt.Value='DIV' then BEGIN FCpte1.ZoomTable:=tzTDivers ;  FCpte2.ZoomTable:=FCpte1.ZoomTable ;END ;

end;

procedure TFQREchean.QRSUBTOTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
if Not CritEdt.Ech.RecapMP then BEGIN PrintBand:=FALSE ; exit ; END ;
NbGene2.Caption:=floatToStr(NbGen) ;
TotCumulG2.Caption:=PrintSolde(CumulD, CumulC, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
TotMontantG2.Caption:=PrintSolde(TotGeneral[1].TotDebit,TotGeneral[1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
TOPREPORT.enabled:=FALSE ;
end;

procedure TFQREchean.FNatureCptTCTDChange(Sender: TObject);
begin
  inherited;
If Not QRLoading then
   BEGIN
   FGenTCTDJoker.Clear ; FGenTCTD1.Clear ; FGenTCTD2.Clear ;
   END ;
If FNatureCptTCTD.Value='TIC' then
   BEGIN
   FGenTCTD1.ZoomTable:=tzGTIC ; FGenTCTD2.ZoomTable:=tzGTIC ;
   END Else
If FNatureCptTCTD.Value='TID' then
   BEGIN
   FGenTCTD1.ZoomTable:=tzGTID ; FGenTCTD2.ZoomTable:=tzGTID ;
   END Else
   BEGIN
   FGenTCTD1.ZoomTable:=tzGTIDTIC ; FGenTCTD2.ZoomTable:=tzGTIDTIC ;
   END ;
end;

procedure TFQREchean.BSDMultiAfterPrint(BandPrinted: Boolean);
begin
  inherited;
if NumR=0 then InitReport([2],CritEdt.Ech.FormatPrint.Report) ;
end;

procedure TFQREchean.FEditionChange(Sender: TObject);
begin
  inherited;
FLigneEchTete.Checked:=Not(FEdition.ItemIndex=3)or(Not(FEdition.ItemIndex=3) or FRuptPaie.Checked) ;
FLigneEchTete.Enabled:=Not(FEdition.ItemIndex=3)or(Not(FEdition.ItemIndex=3) or FRuptPaie.Checked) ;
FLigneEchPied.Checked:=Not(FEdition.ItemIndex=3)or(Not(FEdition.ItemIndex=3) or FRuptPaie.Checked) ;
FLigneEchPied.Enabled:=Not(FEdition.ItemIndex=3)or(Not(FEdition.ItemIndex=3) or FRuptPaie.Checked) ;
If (FEdition.ItemIndex=3) or(Not(FEdition.ItemIndex=3) or FRuptPaie.Checked) then FSautPage.Checked:=False ;
FSautPage.Enabled:=Not(FEdition.ItemIndex=3)or(Not(FEdition.ItemIndex=3) or FRuptPaie.Checked) ;
end;

procedure TFQREchean.BNouvRechClick(Sender: TObject);
begin
  inherited;
If FGenTCTDJoker.Visible then FGenTCTDJoker.Text:='' ;
If FCollJoker.Visible then FCollJoker.Text:='' ;
FEcheancierChange(Nil) ; // Tout est géré par Cette evenement
                         // or avec un vide à true du combo il n'y a pas d'evenement...
end;

procedure TFQREchean.FFiltresChange(Sender: TObject);
begin
  inherited;
//If FEcheancier.ItemIndex=0 then ;
end;

end.
