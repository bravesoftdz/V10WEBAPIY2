unit QRGLVen;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, HSysMenu, Menus, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Buttons,
  Hctrls, ExtCtrls, Mask, Hcompte, ComCtrls, Spin,UtilEdt, HEnt1, Ent1,
  CpteUtil,
{$IFNDEF CCMP}
  EdtLegal,
{$ENDIF}
  CritEDT, HQry,UtilEdt1, QRRupt, HTB97, HPanel, UiUtil, tCalcCum ;


procedure GLVentile ;
procedure GLVentileZoom(Crit : TCritEdt) ;

type
  TFQRGLVen = class(TFQR)
    HLabel3: THLabel;
    FChoixEcart: TRadioGroup;
    FEcart: TSpinEdit;
    HLabel9: THLabel;
    HLabel10: THLabel;
    FTypeGL: TRadioGroup;
    TFCollJoker: THLabel;
    FCollJoker: TEdit;
    TFaC: TLabel;
    FColl2: THCpteEdit;
    TFPaie: THLabel;
    E_MODEPAIE: THMultiValComboBox;
    TFSens: THLabel;
    FSens: THValComboBox;
    FTypePar: TRadioGroup;
    HLabel12: THLabel;
    TFColl: THLabel;
    FColl1: THCpteEdit;
    FTriPar: TRadioGroup;
    FEntete: TCheckBox;
    FPied: TCheckBox;
    FSautPage: TCheckBox;
    TitreDates: TLabel;
    HLabel14: THLabel;
    FP1: TMaskEdit;
    TFDateCpta2: TLabel;
    FP5: TMaskEdit;
    FP6: TMaskEdit;
    Label2: TLabel;
    FP2: TMaskEdit;
    HLabel15: THLabel;
    HLabel16: THLabel;
    FP3: TMaskEdit;
    Label3: TLabel;
    FP7: TMaskEdit;
    FP8: TMaskEdit;
    Label4: TLabel;
    FP4: TMaskEdit;
    HLabel17: THLabel;
    TFPeriodicite: TLabel;
    FPeriodicite: THValComboBox;
    RColl: TQRLabel;
    RColl1: TQRLabel;
    TRaC: TQRLabel;
    RColl2: TQRLabel;
    QRLabel30: TQRLabel;
    RPrevision: TQRLabel;
    QRLabel1: TQRLabel;
    RRetard: TQRLabel;
    QRLabel26: TQRLabel;
    TRSens: TQRLabel;
    RSens: TQRLabel;
    TRSituation: TQRLabel;
    RParAux: TQRLabel;
    QRLabel9: TQRLabel;
    RParPaie: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel10: TQRLabel;
    RE_MODEPAIE: TQRLabel;
    TE_VALIDE: TQRLabel;
    TLE_JOURNAL: TQRLabel;
    TE_PIECELIGNE: TQRLabel;
    TLE_DATECOMPTABLE: TQRLabel;
    TE_REFINTERNE: TQRLabel;
    Col11: TQRLabel;
    Col1: TQRLabel;
    Col22: TQRLabel;
    Col2: TQRLabel;
    col33: TQRLabel;
    col3: TQRLabel;
    Col44: TQRLabel;
    Col4: TQRLabel;
    col5: TQRLabel;
    Col55: TQRLabel;
    Col66: TQRLabel;
    Col6: TQRLabel;
    Col7: TQRLabel;
    REPORTCOL1: TQRLabel;
    REPORTCOL2: TQRLabel;
    REPORTCOL3: TQRLabel;
    REPORTCOL4: TQRLabel;
    REPORTCOL5: TQRLabel;
    REPORTCOL6: TQRLabel;
    REPORTTOTAL1: TQRLabel;
    TLCLEF: TQRLabel;
    BDetail2: TQRBand;
    Solde4: TQRLabel;
    Solde3: TQRLabel;
    Solde2: TQRLabel;
    Solde1: TQRLabel;
    SoldeTotal: TQRLabel;
    Solde0: TQRLabel;
    LE_VALIDE: TQRLabel;
    LE_DATEECHEANCE: TQRDBText;
    E_REFINTERNE: TQRDBText;
    LE_JOURNAL: TQRDBText;
    Solde5: TQRLabel;
    E_PIECELIGECH: TQRLabel;
    BMdp: TQRBand;
    QRLabel3: TQRLabel;
    MdpSolde5: TQRLabel;
    MdpSolde4: TQRLabel;
    MdpSolde3: TQRLabel;
    MdpSolde2: TQRLabel;
    MdpSolde1: TQRLabel;
    MdpSolde0: TQRLabel;
    MdpTotal: TQRLabel;
    BFCompteAux: TQRBand;
    TLCLEF_: TQRLabel;
    AuxCum1: TQRLabel;
    AuxCum2: TQRLabel;
    AuxCum4: TQRLabel;
    AuxCum3: TQRLabel;
    AuxCum0: TQRLabel;
    AuxTotal: TQRLabel;
    AuxCum5: TQRLabel;
    QRLabel33: TQRLabel;
    TotCum5: TQRLabel;
    TotCum4: TQRLabel;
    TotCum3: TQRLabel;
    TotCum2: TQRLabel;
    TotCum1: TQRLabel;
    TotCum0: TQRLabel;
    TotTotal: TQRLabel;
    QROVERLAY: TQRBand;
    Trait8: TQRLigne;
    Trait0: TQRLigne;
    Trait4: TQRLigne;
    Trait3: TQRLigne;
    Trait2: TQRLigne;
    Trait1: TQRLigne;
    Trait5: TQRLigne;
    Trait6: TQRLigne;
    Trait7: TQRLigne;
    Ligne1: TQRLigne;
    REPORTCOL7: TQRLabel;
    REPORTCOL8: TQRLabel;
    REPORTCOL9: TQRLabel;
    REPORTCOL10: TQRLabel;
    REPORTCOL11: TQRLabel;
    REPORTCOL12: TQRLabel;
    REPORTTOTAL2: TQRLabel;
    SEcr: TDataSource;
    QEcr: TQuery;
    QRDLAuxF: TQRDetailLink;
    GMdp: TQRGroup;
    MsgBox: THMsgBox;
    TRLegende: TQRLabel;
    RLegende: TQRLabel;
    BRupt: TQRBand;
    Libre2: TQRLabel;
    TCodRupt: TQRLabel;
    Libre1: TQRLabel;
    Libre0: TQRLabel;
    Libre5: TQRLabel;
    Libre4: TQRLabel;
    Libre3: TQRLabel;
    LibreTotal: TQRLabel;
    DLRupt: TQRDetailLink;
    FLigneRupt: TCheckBox;
    FSaufCptSolde: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure FColl2Change(Sender: TObject);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFCompteAuxBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BMdpBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FChoixEcartClick(Sender: TObject);
    procedure FTypeGLClick(Sender: TObject);
    procedure FP5Exit(Sender: TObject);
    procedure FDateCompta1Exit(Sender: TObject);
    procedure tabSupEnter(Sender: TObject);
    procedure FTypeParClick(Sender: TObject);
    procedure FPeriodiciteChange(Sender: TObject);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QEcrAfterOpen(DataSet: TDataSet);
    procedure BDetail2BeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FNatureCptChange(Sender: TObject);
    procedure BFCompteAuxAfterPrint(BandPrinted: Boolean);
    procedure BNouvRechClick(Sender: TObject);
    procedure DLRuptNeedData(var MoreData: Boolean);
    procedure BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FSansRuptClick(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure BDetailAfterPrint(BandPrinted: Boolean);
    procedure BDetailCheckForSpace;
    procedure GenereSQL ; Override ;
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
  private
    { Déclarations privées }
    StReportAux               : string ;
    LMdp, LRupt               : TStringList ;
    TotEdt, TotAux, TotMdp    : TabTot ;
    Affiche,SautPageRupture   : Boolean ;
    TabDate                   : TTabDate4 ;
    Qr1T_AUXILIAIRE, Qr1T_LIBELLE,
    Qr1T_CORRESP1, Qr1T_CORRESP2          : TStringField ;
    QR2E_REFINTERNE, QR2E_GENERAL,
    QR2E_AUXILIAIRE, QR2E_VALIDE,
    QR2E_JOURNAL, QR2E_EXERCICE,
    QR2E_MODEPAIE                         : TStringField ;
    QR2E_DATECOMPTABLE, QR2E_DATEECHEANCE : TDateTimeField ;
    QR2E_NUMEROPIECE, QR2E_NUMLIGNE,
    QR2E_NUMECHE                          : TIntegerField ;
    QR2COUVERTURE,
    QR2DEBIT, QR2CREDIT                   : TFloatField ;
    procedure GenereSQLSub ;
    Function  QuoiCpt(i : Integer) : String ;
    Procedure GLVentZoom(Quoi : String) ;
    Procedure Calculs(Solde : TQRLabel; var Report1,Report2 : TReport; i : integer; Bool1,Bool2 : Boolean ) ;
    procedure SaisieDates ;
    //Procedure ChangeColl(Nature : String ; C1, C2 : THCpteEdit) ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Function TFQRGLVen.QuoiCpt(i : Integer) : String ;
BEGIN
Case CritEdt.GlV.TypePar of
  0 : BEGIN
      Case i Of
        0 : Result:=Q.FindField('T_AUXILIAIRE').AsString+' '+Q.FindField('T_LIBELLE').AsString+'#'+' '+'@'+'2' ;
        1 : Result:=QR2E_AUXILIAIRE.AsString+' '+Q.FindField('T_LIBELLE').AsString+' '+
             '#'+QR2E_JOURNAL.AsString+' N° '+IntToStr(QR2E_NUMEROPIECE.AsInteger)+' '+DateToStr(QR2E_DateComptable.AsDAteTime)+'-'+
              PrintSolde(Qr2DEBIT.AsFloat,Qr2Credit.AsFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)+' '+
              '@'+'5;'+QR2E_JOURNAL.AsString+';'+UsDateTime(QR2E_DATECOMPTABLE.AsDateTime)+';'+QR2E_NUMEROPIECE.AsString+';'+QR2E_EXERCICE.asString+';'+
              IntToStr(QR2E_NumLigne.AsInteger)+';'
        end ;
      END ;
  1 : BEGIN
      Case i Of
        0 : Result:=Q.FindField('MP_MODEPAIE').AsString+' '+Q.FindField('MP_LIBELLE').AsString+'#'+' '+'@'+'99' ;
        // Rony 14/04/97 -- 1 : Result:=QR2E_AUXILIAIRE.AsString+' '+Q.FindField('T_LIBELLE').AsString+' '+
        1 : Result:=QR2E_MODEPAIE.AsString+' '+Q.FindField('MP_LIBELLE').AsString+' '+QR2E_GENERAL.AsString+' '+QR2E_AUXILIAIRE.AsString+' '+
             '#'+QR2E_JOURNAL.AsString+' N° '+IntToStr(QR2E_NUMEROPIECE.AsInteger)+' '+DateToStr(QR2E_DateComptable.AsDAteTime)+'-'+
              PrintSolde(Qr2DEBIT.AsFloat,Qr2Credit.AsFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)+' '+
              '@'+'5;'+QR2E_JOURNAL.AsString+';'+UsDateTime(QR2E_DATECOMPTABLE.AsDateTime)+';'+QR2E_NUMEROPIECE.AsString+';'+QR2E_EXERCICE.asString+';'+
              IntToStr(QR2E_NumLigne.AsInteger)+';'
        end ;
      END ;
  END ;
END ;

Procedure TFQRGLVen.GLVentZoom(Quoi : String) ;
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

procedure GLVentile ;
var QR : TFQRGLVen ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFQRGLVen.Create(Application) ;
Edition.Etat:=etGlVen ;
QR.QRP.QRPrinter.OnSynZoom:=QR.GLVentZoom ;
QR.InitType (nbAux,neGlV,msRien,'QRGLVEN','',TRUE,FALSE,FALSE,Edition) ;
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

procedure GLVentileZoom(Crit : TCritEdt) ;
var QR : TFQRGLVen ;
    Edition : TEdition ;
BEGIN
QR:=TFQRGLVen.Create(Application) ;
Edition.Etat:=etGlVen ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.GLVentZoom ;
 QR.CritEdt:=Crit ;
 QR.InitType (nbAux,neGlV,msRien,'QRGLVEN','',FALSE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFQRGLVen.InitDivers ;
BEGIN
Inherited ;
{ Calcul des différentes fourchettes de dates }
CalculDateTiers(neGL,Ventile,CritEdt,TabDate,FP1,FP2,FP3,FP4,FP5,FP6,FP7,FP8) ;
{ Titres des colonnes de montants }
Case CritEdt.GlV.TypeGL of
  0 : BEGIN
      Col1.Caption:=MsgBox.Mess[2] ;                              Col11.Caption:=Format_String(' ',1) ;
      Col2.Caption:=MsgBox.Mess[4]+DateToStr(CritEdt.Date1+1) ; Col22.Caption:=MsgBox.Mess[5]+DateToStr(TabDate[1]) ;
      Col3.Caption:=MsgBox.Mess[4]+DateToStr(TabDate[1]+1) ;      Col33.Caption:=MsgBox.Mess[5]+DateToStr(TabDate[2]) ;
      Col4.Caption:=MsgBox.Mess[4]+DateToStr(TabDate[2]+1) ;      Col44.Caption:=MsgBox.Mess[5]+DateToStr(TabDate[3]) ;
      Col5.Caption:=MsgBox.Mess[4]+DateToStr(TabDate[3]+1) ;      Col55.Caption:=MsgBox.Mess[5]+DateToStr(TabDate[4]) ;
      Col6.Caption:=MsgBox.Mess[6] ;                              Col66.Caption:=DateToStr(TabDate[4]+1) ;
      END ;
  1 : BEGIN
      Col1.Caption:=MsgBox.Mess[7] ;                              Col11.Caption:=DateToStr(TabDate[4]-1) ;
      Col2.Caption:=MsgBox.Mess[4]+DateToStr(TabDate[4]) ;        Col22.Caption:=MsgBox.Mess[5]+DateToStr(TabDate[3]-1) ;
      Col3.Caption:=MsgBox.Mess[4]+DateToStr(TabDate[3]) ;        Col33.Caption:=MsgBox.Mess[5]+DateToStr(TabDate[2]-1) ;
      Col4.Caption:=MsgBox.Mess[4]+DateToStr(TabDate[2]) ;        Col44.Caption:=MsgBox.Mess[5]+DateToStr(TabDate[1]-1) ;
      Col5.Caption:=MsgBox.Mess[4]+DateToStr(TabDate[1]) ;        Col55.Caption:=MsgBox.Mess[5]+DateToStr(CritEdt.Date1-1) ;
      Col6.Caption:=MsgBox.Mess[3] ;                              Col66.Caption:=Format_String(' ',1) ;
      END ;
 end;
{ Labels sur la bande pied d'état }
BDetail.ForceNewPage:=FSautPage.checked ;
{ Séparateurs de bandes }
//BDetail.Frame.DrawTop:=CritEdt.GlV.FormatPrint.PrSepCompte[1] ; // compte
BDetail.Frame.DrawBottom:=CritEdt.GlV.FormatPrint.PrSepCompte[3] ; // EnTete

BFCompteAux.Frame.DrawTop:=CritEdt.GlV.FormatPrint.PrSepCompte[2] ; // Pied
BFCompteAux.Frame.DrawBottom:=CritEdt.GlV.FormatPrint.PrSepCompte[1] ; //Compte

BRupt.Frame.DrawTop:=CritEdt.GlV.FormatPrint.PrSepCompte[4] ;
BRupt.Frame.DrawBottom:=BRupt.Frame.DrawTop ;
//BMdp.Frame.DrawTop:=CritEdt.FormatPrint.GlV.PrSepCompte[3] ; en attente
{ Edition par auxiliaire ou par mode de paiement }
(* Rony 14/04/97
Case CritEdt.GlV.TypePar of
  0 : BEGIN
      TLCLEF.Caption:=MsgBox.Mess[8] ; TTOTCLEF.Caption:=MsgBox.Mess[13] ;
      LCLEF.DataField:='T_AUXILIAIRE' ; LLIB.DataField:='T_LIBELLE' ; TOTCLEF.DataField:='T_AUXILIAIRE' ;
      END ;
  1 : BEGIN
      TLCLEF.Caption:=MsgBox.Mess[9] ; TTOTCLEF.Caption:=MsgBox.Mess[14] ;
      LCLEF.DataField:='MP_MODEPAIE' ; LLIB.DataField:='MP_LIBELLE' ; TOTCLEF.DataField:='MP_MODEPAIE' ;
      END ;
 end ;
*)
if Not E_MODEPAIE.Tous then ChargeRecapMPTiers(LMdp, E_MODEPAIE);
(* Rony 14/04/97  car Titre + Num + Libelle == Label
If CritEdt.GLV.TypePAR=0 Then
   BEGIN
   ChangeFormatCompte(fbAux,fbAux,Self,LCLEF.Width,4,5,LCLEF.Name) ;
   LCleF.Alignment:=taLeftJustify ;
   END Else LCleF.Alignment:=taCenter ;
*)
TE_VALIDE.Caption:='' ;
If CritEdt.SautPageRupt And (CritEdt.GlV.RuptOnly=Avec) And (CritEdt.SautPage<>1) Then BDetail.ForceNewPage:=FALSE ;
SautPageRupture:=FALSE ;
END ;

procedure TFQRGLVen.GENERESQL ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
GenereSQLTiers(neGL,Ventile,CritEdt,Q,FALSE,AvecRevision.State) ;
GenereSQLSUB ;
END ;

procedure TFQRGLVen.GenereSQLSUB ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
GenereSQLSUBTiers(neGL,Ventile,CritEdt,QEcr,FALSE) ;
END ;

procedure TFQRGLVen.RenseigneCritere ;
Var St11,St22 : String ;
BEGIN
Inherited ;
if CritEdt.SJoker then
   BEGIN
   RColl.Caption:=MsgBox.Mess[15] ;
   RColl1.Caption:=FCollJoker.Text ;
   END else
   BEGIN
   RColl.Caption:=MsgBox.Mess[14] ;
   //RColl1.Caption:=FColl1.Text ; RColl2.Caption:=FColl2.Text ;
   PositionneFourchetteST(FColl1,FColl2,St11,St22) ;
   RColl1.Caption:=St11 ;
   RColl2.Caption:=St22 ;
   END ;
TRaC.Visible:=Not CritEdt.SJoker ; RColl2.Visible:=Not CritEdt.SJoker ;
RSens.Caption:=FSens.Text ;
RE_MODEPAIE.Caption:=E_MODEPAIE.Text ;
Case FTypePar.ItemIndex of
  0 : BEGIN RParAux.Caption:='þ'    ; RParPaie.Caption:='o'  ; END ;
  1 : BEGIN RParAux.Caption:='o'    ; RParPaie.Caption:='þ' ; END ;
end ;
Case FTypeGL.ItemIndex of
  0 : BEGIN RPrevision.Caption:='þ' ; RRetard.Caption:='o' ;  END ;
  1 : BEGIN RPrevision.Caption:='o' ; RRetard.Caption:='þ' ; END ;
end ;
END ;

procedure TFQRGLVen.ChoixEdition ;
{ Initialisation des options d'édition }
BEGIN
Inherited ;
DLRupt.PrintBefore:=TRUE ;
Case CritEdt.Rupture of
  rLibres   : BEGIN
              DLRupt.PrintBefore:=FALSE ;
              ChargeGroup(LRupt,['T00','T01','T02','T03','T04','T05','T06','T07','T08','T09']) ;
              END ;
  rRuptures : BEGIN
              ChargeRupt(LRupt, 'RUT', CritEdt.GlV.PlanRupt, '', '') ;
              NiveauRupt(LRupt);
              END ;
  rCorresp  : BEGIN
              ChargeRuptCorresp(LRupt, CritEdt.GlV.PlanRupt, '', '', False) ;
              NiveauRupt(LRupt);
              END ;
  End ;
END ;

procedure TFQRGLVen.FinirPrint ;
{$IFNDEF CCMP}
Var SOlde : Double ;
{$ENDIF}
BEGIN
Inherited ;
QEcr.Close ;
if (CritEdt.Rupture<>rRien) then VideRupt(LRupt) ;
{$IFNDEF CCMP}
if OkMajEdt Then
   BEGIN
   Solde:=TotEdt[6].TotDebit-TotEdt[6].TotCredit ;
   if Solde<0
      then MajEdition('GLV', '', DateToStr(CritEdt.Date1), DateToStr(CritEdt.Date1),'', TotEdt[6].TotDebit, TotEdt[6].TotCredit, Solde, 0)
      else MajEdition('GLV', '', DateToStr(CritEdt.Date1), DateToStr(CritEdt.Date1),'', TotEdt[6].TotDebit, TotEdt[6].TotCredit, 0, Solde) ;
   END ;
{$ENDIF}
END ;

function  TFQRGLVen.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ;
If Result Then
   BEGIN
   if CritEdt.GLV.ChoixEcart=1 then Result:=ValidPeriodeTiers(FP1,FP2,FP3,FP4,FP5,FP6,FP7,FP8) ;
   Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
   END ;
END ;

procedure TFQRGLVen.RecupCritEdt ;
Var NonLibres : Boolean ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  SJoker:=FCollJoker.Visible ;
  if SJoker Then
     BEGIN
     SCpt1:=FCollJoker.Text ; SCpt2:=FCollJoker.Text ;
     LSCpt1:=SCpt1 ; LSCpt2:=SCpt2 ;
     END Else
     BEGIN
     SCpt1:=FColl1.Text     ; SCpt2:=FColl2.Text ;
     PositionneFourchetteSt(FColl1,FColl2,CritEdt.LSCpt1,CritEdt.LSCpt2) ;
     END ;
  GlV.TriePar:=FTriPar.ItemIndex ;
  GlV.TypeGL:=FTypeGL.ItemIndex ;
  GlV.TypePar:=FTypePar.ItemIndex ;
  GlV.Sens:=FSens.ItemIndex ;
  GlV.Ecart:=FEcart.Value ;
  GlV.ChoixEcart:=FChoixEcart.ItemIndex ;
  GlV.RuptOnly:=QuelleTypeRupt(1,FSAnsRupt.Checked,FAvecRupt.Checked,FSurRupt.Checked) ;
  NonLibres:=((Rupture=rRuptures) or (Rupture=rCorresp)) ;
  If NonLibres Then GlV.PlanRupt:=FPlanRuptures.Value ;
  GlV.OnlyCptAssocie:=(CritEdt.Rupture<>rRien) and FOnlyCptAssocie.Checked ;
  If (CritEdt.Rupture=rCorresp) Then GLV.PlansCorresp:=FPlanRuptures.ItemIndex+1 ;
  GlV.SaufCptSolde:=FSaufCptSolde.Checked ;
  With GlV.FormatPrint Do
    BEGIN
    PrSepCompte[2]:=FPied.Checked ;
    PrSepCompte[3]:=FEntete.Checked ;
    PrSepCompte[4]:=FLigneRupt.Checked ;
    END ;
  END ;
END ;

Procedure TFQRGLVen.Calculs(Solde : TQRLabel; var Report1,Report2 : TReport; i : integer; Bool1,Bool2 : Boolean ) ;
{ Incrémentations de toutes les variables de l'état }
Var TotRupt : Array[0..13] of double ;
    TotTemp : TabTot ;
    CptRupt : String ;
BEGIN
Fillchar(TotTemp,SizeOf(TotTemp),#0) ;
if QR2DEBIT.AsFloat<>0 then
   BEGIN
   if Bool1 then
      BEGIN
      Solde.Caption:=PrintSolde(QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat,QR2CREDIT.AsFloat, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
      Report1.TotDebit:= Arrondi(QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat,CritEdt.Decimale) ;
      Report2.TotDebit:=Report1.TotDebit ;
      TotAux[i].TotDebit:= Arrondi(TotAux[i].TotDebit+QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat,CritEdt.Decimale) ;
      TotEdt[i].TotDebit:= Arrondi(TotEdt[i].TotDebit+QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat,CritEdt.Decimale) ;
      TotTemp[i].TotDebit:= Arrondi(TotTemp[i].TotDebit+QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      if Bool2 then
         TotAux[6].TotDebit:= Arrondi(TotAux[6].TotDebit+QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat,CritEdt.Decimale) ;
      TotEdt[6].TotDebit:= Arrondi(TotEdt[6].TotDebit+QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat,CritEdt.Decimale) ;
      TotTemp[6].TotDebit:= Arrondi(TotTemp[6].TotDebit+QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      END else
      BEGIN
      TotMdp[i].TotDebit:= Arrondi(TotMdp[i].TotDebit+QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat,CritEdt.Decimale) ;
      TotMdp[6].TotDebit:= Arrondi(TotMdp[6].TotDebit+QR2DEBIT.AsFloat-QR2COUVERTURE.AsFloat,CritEdt.Decimale) ;
      END ;
   END else
   BEGIN
   if Bool1 then
      BEGIN
      Solde.Caption:=PrintSolde(QR2DEBIT.AsFloat,QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
      Report1.TotCredit:= Arrondi(QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat,CritEdt.Decimale) ;
      Report2.TotCredit:=Report1.TotCredit ;
      TotAux[i].TotCredit:= Arrondi(TotAux[i].TotCredit+QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat,CritEdt.Decimale) ;
      TotEdt[i].TotCredit:= Arrondi(TotEdt[i].TotCredit+QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat,CritEdt.Decimale) ;
      TotTemp[i].TotCredit:= Arrondi(TotTemp[i].TotCredit+QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      if Bool2 then
         TotAux[6].TotCredit:= Arrondi(TotAux[6].TotCredit+QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat,CritEdt.Decimale) ;
      TotEdt[6].TotCredit:= Arrondi(TotEdt[6].TotCredit+QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat,CritEdt.Decimale) ;
      TotTemp[6].TotCredit:= Arrondi(TotTemp[6].TotCredit+QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat, CritEdt.Decimale) ;
      END else
      BEGIN
      TotMdp[i].TotCredit:= Arrondi(TotMdp[i].TotCredit+QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat,CritEdt.Decimale) ;
      TotMdp[6].TotCredit:= Arrondi(TotMdp[6].TotCredit+QR2CREDIT.AsFloat-QR2COUVERTURE.AsFloat,CritEdt.Decimale) ;
      END ;
   END ;
Soldetotal.Caption:=PrintSolde(TotAux[6].TotDebit, TotAux[6].TotCredit, CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;
TotRupt[0]:=Arrondi(TotTemp[0].TotDebit,CritEdt.Decimale) ;
TotRupt[1]:=Arrondi(TotTemp[0].TotCredit,CritEdt.Decimale) ;

TotRupt[2]:=Arrondi(TotTemp[1].TotDebit,CritEdt.Decimale) ;
TotRupt[3]:=Arrondi(TotTemp[1].TotCredit,CritEdt.Decimale) ;

TotRupt[4]:=Arrondi(TotTemp[2].TotDebit,CritEdt.Decimale) ;
TotRupt[5]:=Arrondi(TotTemp[2].TotCredit,CritEdt.Decimale) ;

TotRupt[6]:=Arrondi(TotTemp[3].TotDebit,CritEdt.Decimale) ;
TotRupt[7]:=Arrondi(TotTemp[3].TotCredit,CritEdt.Decimale) ;

TotRupt[8]:=Arrondi(TotTemp[4].TotDebit,CritEdt.Decimale) ;
TotRupt[9]:=Arrondi(TotTemp[4].TotCredit,CritEdt.Decimale) ;

TotRupt[10]:=Arrondi(TotTemp[5].TotDebit,CritEdt.Decimale) ;
TotRupt[11]:=Arrondi(TotTemp[5].TotCredit,CritEdt.Decimale) ;

TotRupt[12]:=Arrondi(TotTemp[6].TotDebit,CritEdt.Decimale) ;
TotRupt[13]:=Arrondi(TotTemp[6].TotCredit,CritEdt.Decimale) ;
(**)
Case CritEdt.Rupture of
  rLibres   : AddGroupLibre(LRupt,Q,fbAux,CritEdt.LibreTrie,TotRupt) ;
  rRuptures : AddRupt(LRupt,Qr1T_AUXILIAIRE.AsString,TotRupt) ;
  rCorresp  : BEGIN
              Case CritEDT.GlV.PlansCorresp Of
                1 : If Qr1T_CORRESP1.AsString<>'' Then CptRupt:=Qr1T_CORRESP1.AsString+Qr1T_AUXILIAIRE.AsString
                                                  Else CptRupt:='.'+Qr1T_AUXILIAIRE.AsString ;
                2 : If Qr1T_CORRESP2.AsString<>'' Then CptRupt:=Qr1T_CORRESP2.AsString+Qr1T_AUXILIAIRE.AsString
                                                  Else CptRupt:='.'+Qr1T_AUXILIAIRE.AsString ;
                Else CptRupt:=Qr1T_AUXILIAIRE.AsString ;
                End ;
              AddRuptCorres(LRupt,CptRupt,TotRupt) ;
              END ;
  End ;
END ;

procedure TFQRGLVen.SaisieDates ;
Var i : byte ;
BEGIN { Selon le type de GL (Prévision ou Retard), acces aux zones saisissables et le report de la Date d'arrêtée}
If FTypeGL.ItemIndex=0 then
   BEGIN
   TabSup.Caption:=MsgBox.Mess[10] ; FP1.Text:=DateToStr(StrToDate(FDateCompta1.Text)+1) ;
   For i:=1 to 4 do TMaskEdit(FindComponent('FP'+InttoStr(i))).Enabled:=False ;
   For i:=5 to 8 do TMaskEdit(FindComponent('FP'+InttoStr(i))).Enabled:=True ;
   END else
   BEGIN
   TabSup.Caption:=MsgBox.Mess[11] ; FP8.Text:=DateToStr(StrToDate(FDateCompta1.Text)-1) ; ;
   FP1.Enabled:=True ; FP2.Enabled:=False ; FP3.Enabled:=False ; FP4.Enabled:=False ;
   FP5.Enabled:=True ; FP6.Enabled:=True  ; FP7.Enabled:=True  ; FP8.Enabled:=True ;
   END ;
If FPeriodicite.Values.Count>0 Then FPeriodicite.Value:=FPeriodicite.Values[0] ;
END ;

procedure TFQRGLVen.FormShow(Sender: TObject);
begin
HelpContext:=7559000 ;
//Standards.HelpContext:=7559010 ;
//TabSup.HelpContext:=7559025 ;
//Avances.HelpContext:=7559020 ;
//Mise.HelpContext:=7559030 ;
//Option.HelpContext:=7559040 ;
//TabRuptures.HelpContext:=7559050 ;

FSens.ItemIndex:=2 ;
E_MODEPAIE.Text:=Traduirememoire('<<Tous>>') ;
FDateCompta1.Text:=DateToStr(V_PGI.DateEntree) ;
  inherited;
If FFiltres.Text='' then FDateCompta1.Text:=DateToStr(V_PGI.DateEntree) ;
TabSup.TabVisible:=(FChoixEcart.ItemIndex=1) ;
FSurRupt.Visible:=False ;
FCodeRupt1.Visible:=False  ; FCodeRupt2.Visible:=False ;
TFCodeRupt1.Visible:=False ; TFCodeRupt2.Visible:=False ;
FOnlyCptAssocie.Enabled:=False ; FLigneRupt.Enabled:=False ;
{$IFDEF CCMP}
FNatureCpt.Vide := False;
if (VH^.CCMP.LotCli) then begin FNatureCpt.Plus := ' AND (CO_CODE="AUD" OR CO_CODE="CLI" OR CO_CODE="DIV")'; FNatureCpt.Value:='CLI'; end
                     else begin FNatureCpt.Plus := ' AND (CO_CODE="AUC" OR CO_CODE="DIV" OR CO_CODE="FOU" OR CO_CODE="SAL")'; FNatureCpt.Value:='FOU'; end;
{$ENDIF}
end;

procedure TFQRGLVen.FColl2Change(Sender: TObject);
Var AvecJokerS : Boolean ;
begin
  inherited;
AvecJokerS:=Joker(FColl1, FColl2, FCollJoker) ;
TFaC.Visible:=Not AvecJokerS ;
TFColl.Visible:=Not AvecJokerS ;
TFCollJoker.Visible:=AvecJokerS ;
end;

procedure TFQRGLVen.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
{ Initialisations des variables sur la rupture "par auxiliaire" ou "par modes de paiement"
{ !!!! -> 'TotAux' concerne les ruptures "par auxiliaire" ou "par mode de paiment"
       -> 'TotMdp' concerne les totaux pour les modes de paiement non retenus s'ils existent }
begin
  inherited;
Fillchar(TotAux,SizeOf(TotAux),#0) ;
Fillchar(TotMdp,SizeOf(TotMdp),#0) ;
TLCLEF.Left:=TLE_JOURNAL.Left ;
//TLCLEF.Width:=TE_REFINTERNE.Left-TE_VALIDE.Width ;
TLCLEF.Width:=TLE_JOURNAL.Width+TE_PIECELIGNE.Width+TLE_DATECOMPTABLE.Width+TE_REFINTERNE.Left+3 ;
Case CritEdt.GlV.TypePar of
  0 : BEGIN
      StReportAux:=Q.Findfield('T_AUXILIAIRE').AsString ;
      TLCLEF.Caption:=MsgBox.Mess[16]+' '+Q.Findfield('T_AUXILIAIRE').AsString+' '+Q.Findfield('T_LIBELLE').AsString ;
      END ;
  1 : BEGIN
      StReportAux:=Q.Findfield('MP_MODEPAIE').AsString ;
      TLCLEF.Caption:=MsgBox.Mess[17]+' '+Q.Findfield('MP_MODEPAIE').AsString+' '+Q.Findfield('MP_LIBELLE').AsString ;
      END ;
 end ;
InitReport([2],CritEdt.GlV.FormatPrint.Report) ;
Case CritEdt.Rupture of
  rLibres    : if CritEdt.GLV.OnlyCptAssocie then PrintBand:=DansRuptLibre(Q,fbAux,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
  rRuptures  : if CritEdt.GlV.OnlyCptAssocie then PrintBand:=DansRupt(LRupt,Qr1T_AUXILIAIRE.AsString) ;
  rCorresp   : if CritEdt.GlV.OnlyCptAssocie then
                 if CritEDT.GlV.PlansCorresp=1 then PrintBand:=(Qr1T_CORRESP1.AsString<>'') Else
                 if CritEDT.GlV.PlansCorresp=2 then PrintBand:=(Qr1T_CORRESP2.AsString<>'') ;
  End;
Affiche:=PrintBand ;
if PrintBand then Quoi:=QuoiCpt(0) ;
end;

procedure TFQRGLVen.BFCompteAuxBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=Affiche ;
If Not PrintBand then Exit ;
TLCLEF_.Left:=TLCLEF.Left ;
TLCLEF_.Width:=TLCLEF.Width ;
Case CritEdt.GlV.TypePar of
  0 : BEGIN
      TLCLEF_.Caption:=MsgBox.Mess[18]+' '+Q.Findfield('T_AUXILIAIRE').AsString+' '+Q.Findfield('T_LIBELLE').AsString ;
      END ;
  1 : BEGIN
      TLCLEF_.Caption:=MsgBox.Mess[19]+' '+Q.Findfield('MP_MODEPAIE').AsString+' '+Q.Findfield('MP_LIBELLE').AsString ;
      END ;
 end ;
Quoi:=QuoiCpt(0) ;
AuxCum5.Caption:= PrintSolde(TotAux[5].TotDebit, TotAux[5].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
AuxCum4.Caption:= PrintSolde(TotAux[4].TotDebit, TotAux[4].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
AuxCum3.Caption:= PrintSolde(TotAux[3].TotDebit, TotAux[3].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
AuxCum2.Caption:= PrintSolde(TotAux[2].TotDebit, TotAux[2].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
AuxCum1.Caption:= PrintSolde(TotAux[1].TotDebit, TotAux[1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
AuxCum0.Caption:= PrintSolde(TotAux[0].TotDebit, TotAux[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
AuxTotal.Caption:=PrintSolde(TotAux[6].TotDebit,TotAux[6].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
end;

procedure TFQRGLVen.BMdpBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=((TotMdp[6].TotDebit<>0) Or (TotMdp[6].TotCredit<>0)) And (FtypePar.ItemIndex=0) ;
MdpSolde5.Caption:=PrintSolde(TotMdp[5].TotDebit, TotMdp[5].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
MdpSolde4.Caption:=PrintSolde(TotMdp[4].TotDebit, TotMdp[4].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
MdpSolde3.Caption:=PrintSolde(TotMdp[3].TotDebit, TotMdp[3].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
MdpSolde2.Caption:=PrintSolde(TotMdp[2].TotDebit, TotMdp[2].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
MdpSolde1.Caption:=PrintSolde(TotMdp[1].TotDebit, TotMdp[1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
MdpSolde0.Caption:=PrintSolde(TotMdp[0].TotDebit, TotMdp[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
MdpTotal.Caption:= PrintSolde(TotMdp[6].TotDebit, TotMdp[6].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
end;

procedure TFQRGLVen.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TotCum5.Caption:= PrintSolde(TotEdt[5].TotDebit, TotEdt[5].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
TotCum4.Caption:= PrintSolde(TotEdt[4].TotDebit, TotEdt[4].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
TotCum3.Caption:= PrintSolde(TotEdt[3].TotDebit, TotEdt[3].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
TotCum2.Caption:= PrintSolde(TotEdt[2].TotDebit, TotEdt[2].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
TotCum1.Caption:= PrintSolde(TotEdt[1].TotDebit, TotEdt[1].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
TotCum0.Caption:= PrintSolde(TotEdt[0].TotDebit, TotEdt[0].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
TotTotal.Caption:=PrintSolde(TotEdt[6].TotDebit, TotEdt[6].TotCredit, CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFQRGLVen.FChoixEcartClick(Sender: TObject);
begin
  inherited;
FEcart.Enabled:=(FChoixEcart.ItemIndex=0) ;
TabSup.TabVisible:=(FChoixEcart.ItemIndex=1) ;
If (FChoixEcart.ItemIndex=1) then SaisieDates ;
end;

procedure TFQRGLVen.FTypeGLClick(Sender: TObject);
begin
  inherited;
If FChoixEcart.ItemIndex=1 then SaisieDates ;
end;

procedure TFQRGLVen.FP5Exit(Sender: TObject);
Var St : String ;
    A  : Char   ;
begin
  inherited;
St:=TMaskEdit(Sender).Name ; A:=St[3] ;
if (TMaskEdit(Sender).Text<>'  /  /    ') and IsValidDate(TMaskEdit(Sender).Text)  then
   BEGIN
   Case FTypeGL.ItemIndex of
     0 : BEGIN
         FP2.Text:=DateToStr(StrToDate(FP5.Text)+1) ;
         FP3.Text:=DateToStr(StrToDate(FP6.Text)+1) ;
         FP4.Text:=DateToStr(StrToDate(FP7.Text)+1) ;
        END ;
     1 : BEGIN
         FP2.Text:=DateToStr(StrToDate(FP5.Text)+1)  ;
         FP3.Text:=DateToStr(StrToDate(FP6.Text)+1) ;
         FP4.Text:=DateToStr(StrToDate(FP7.Text)+1) ;
         FP7.Text:=DateToStr(StrToDate(FP4.Text)-1) ;
         If A='8' then FDateCompta1.Text:=DateToStr(StrToDate(FP8.Text)+1) ;
         END ;
   end ;
   END else
   BEGIN
   MsgRien.Execute(1,'',' '+TMaskEdit(Sender).Text) ;
   If FTypeGL.ItemIndex=0 then TMaskEdit(Sender).Text:=DateToStr(VH^.ExoV8.Deb)
                          Else TMaskEdit(Sender).Text:=DateToStr(VH^.ExoV8.Fin);
   TMaskEdit(Sender).SetFocus ;
   END ;
end;

procedure TFQRGLVen.FDateCompta1Exit(Sender: TObject);
begin
  inherited;
If FChoixEcart.ItemIndex=1 then SaisieDates ;
end;

procedure TFQRGLVen.tabSupEnter(Sender: TObject);
begin
  inherited;
FPeriodicite.SetFocus ;
end;

procedure TFQRGLVen.FTypeParClick(Sender: TObject);
begin
  inherited;
if FTypePar.ItemIndex=1 then
   BEGIN
   E_MODEPAIE.Text:=Traduirememoire('<<Tous>>') ;
   FTriPar.ItemIndex:=0 ;
   FSansRupt.Checked:=True ;
   FSaufCptSolde.Checked:=FALSE ; FSaufCptSolde.Enabled:=FALSE ;
   END Else BEGIN FSaufCptSolde.Enabled:=TRUE ; END ;
E_MODEPAIE.Enabled:=(FTypePar.ItemIndex=0) ;
FTriPar.Enabled:=(FTypePar.ItemIndex=0) ;
FGroupChoixRupt.Enabled:=(FTypePar.ItemIndex=0) ;
end;

procedure TFQRGLVen.FPeriodiciteChange(Sender: TObject);
Var i : Integer ;
    TabD : TTabDate8 ;
begin
  inherited;
If QRLoading then Exit ;
PeriodiciteChangeTiers(neGL,Ventile,FPeriodicite.ItemIndex,FP8,FP1,FTypeGL.ItemIndex=1,TabD) ;
For i:=1 to 8 do TMaskEdit(FindComponent('FP'+InttoStr(i))).text:=DateToStr(TabD[i] ) ;
end;


procedure TFQRGLVen.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TitreReportH.Caption:=TitreReportB.Caption ;
ReportCol1.Caption:=ReportCol7.Caption ;
ReportCol2.Caption:=ReportCol8.Caption ;
ReportCol3.Caption:=ReportCol9.Caption ;
ReportCol4.Caption:=ReportCol10.Caption ;
ReportCol5.Caption:=ReportCol11.Caption ;
ReportCol6.Caption:=ReportCol12.Caption ;
ReportTotal1.Caption:=ReportTotal2.Caption ;
end;

procedure TFQRGLVen.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var MReport : TabTRep ;
begin
  inherited;
Case QuelReportBAL(CritEdt.GlV.FormatPrint.Report,MReport) of
  1 : TitreReportB.Caption:=MsgBox.Mess[20] ;
  2 : TitreReportB.Caption:=MsgBox.Mess[20]+' '+StReportAux ;
 end ;
ReportCol7.Caption  :=PrintSolde(MReport[1].TotDebit,MReport[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
ReportCol8.Caption  :=PrintSolde(MReport[2].TotDebit,MReport[2].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
ReportCol9.Caption  :=PrintSolde(MReport[3].TotDebit,MReport[3].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
ReportCol10.Caption :=PrintSolde(MReport[4].TotDebit,MReport[4].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
ReportCol11.Caption :=PrintSolde(MReport[5].TotDebit,MReport[5].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
ReportCol12.Caption :=PrintSolde(MReport[6].TotDebit,MReport[6].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
ReportTotal2.Caption:=PrintSolde(MReport[7].TotDebit,MReport[7].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
end;

procedure TFQRGLVen.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
If CritEDT.Rupture<>rRien then
   BEGIN
   Qr1T_AUXILIAIRE  :=TStringField(Q.FindField('T_AUXILIAIRE'));
   Qr1T_LIBELLE     :=TStringField(Q.FindField('T_LIBELLE'));
   If CritEDT.Rupture=rCorresp then
      BEGIN
      Qr1T_CORRESP1         :=TStringField(Q.FindField('T_CORRESP1'));
      Qr1T_CORRESP2         :=TStringField(Q.FindField('T_CORRESP2'));
      END ;
   END ;
end;

procedure TFQRGLVen.QEcrAfterOpen(DataSet: TDataSet);
begin
  inherited;
QR2E_AUXILIAIRE    :=TStringField(QEcr.FindField('E_AUXILIAIRE')) ;
QR2E_MODEPAIE      :=TStringField(QEcr.FindField('E_MODEPAIE')) ;
QR2E_GENERAL       :=TStringField(QEcr.FindField('E_GENERAL')) ;
QR2E_DATECOMPTABLE :=TDateTimeField(QEcr.FindField('E_DATECOMPTABLE')) ;
QR2E_DATEECHEANCE  :=TDateTimeField(QEcr.FindField('E_DATEECHEANCE')) ;
QR2E_NUMEROPIECE   :=TIntegerField(QEcr.FindField('E_NUMEROPIECE')) ;
QR2E_NUMLIGNE      :=TIntegerField(QEcr.FindField('E_NUMLIGNE')) ;
QR2E_NUMECHE       :=TIntegerField(QEcr.FindField('E_NUMECHE')) ;
QR2E_REFINTERNE    :=TStringField(QEcr.FindField('E_REFINTERNE')) ;
QR2E_EXERCICE      :=TStringField(QEcr.FindField('E_EXERCICE')) ;
QR2E_VALIDE        :=TStringField(QEcr.FindField('E_VALIDE')) ;
QR2E_JOURNAL       :=TStringField(QEcr.FindField('E_JOURNAL')) ;
QR2COUVERTURE    :=TFloatField(QEcr.FindField('COUVERTURE')) ;
QR2DEBIT           :=TFloatField(QEcr.FindField('DEBIT')) ;
QR2CREDIT          :=TFloatField(QEcr.FindField('CREDIT')) ;
ChgMaskChamp(Qr2DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
ChgMaskChamp(Qr2CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
end;

procedure TFQRGLVen.BDetail2BeforePrint(var PrintBand: Boolean;  var Quoi: string);
var MReport  : TabTRep ;
    OkCalCul : Boolean ;
begin
  inherited;
Fillchar(MReport,SizeOf(MReport),#0) ;
if (Not E_MODEPAIE.Tous) And (FTypePar.ItemIndex=0) then PrintBand:=MDPRetenuTiers(LMdp, QR2E_MODEPAIE.AsString) ;
OkCalCul:=Affiche ;
If OkCalCul then
   BEGIN
   E_PIECELIGECH.Caption:=QR2E_NUMEROPIECE.AsString+' / '+QR2E_NUMLIGNE.AsString;
   solde0.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, 0, CritEdt.AfficheSymbole) ;
   solde1.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, 0, CritEdt.AfficheSymbole) ;
   solde2.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, 0, CritEdt.AfficheSymbole) ;
   solde3.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, 0, CritEdt.AfficheSymbole) ;
   solde4.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, 0, CritEdt.AfficheSymbole) ;
   solde5.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, 0, CritEdt.AfficheSymbole) ;
   soldeTotal.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, 0, CritEdt.AfficheSymbole) ;
   if QR2E_VALIDE.AsString='X' then LE_VALIDE.Caption:='V' else LE_VALIDE.Caption:=' ' ;
   Case CritEdt.GlV.TypeGL of
     0 : BEGIN
         if QR2E_DateECHEANCE.AsDateTime<=CritEdt.Date1
            then case CritEdt.GlV.TypePar of
                   0 : Calculs(Solde5, MReport[1], MReport[7], 5, PrintBand, True) ;
                   1 : Calculs(Solde5, MReport[1], MReport[7], 5, PrintBand, False) ;
                 end ;
         if (QR2E_DateECHEANCE.AsDateTime>CritEdt.Date1) And (QR2E_DateECHEANCE.AsDateTime<=TabDate[1])
            then Calculs(Solde4, MReport[2], MReport[7], 4, PrintBand, True) ;
         if (QR2E_DateECHEANCE.AsDateTime>TabDate[1]) And (QR2E_DateECHEANCE.AsDateTime<=TabDate[2])
            then Calculs(Solde3, MReport[3], MReport[7], 3, PrintBand, True) ;
         if (QR2E_DateECHEANCE.AsDateTime>TabDate[2]) And (QR2E_DateECHEANCE.AsDateTime<=TabDate[3])
            then Calculs(Solde2, MReport[4], MReport[7], 2, PrintBand, True) ;
         if (QR2E_DateECHEANCE.AsDateTime>TabDate[3]) And (QR2E_DateECHEANCE.AsDateTime<=TabDate[4])
            then Calculs(Solde1, MReport[5], MReport[7], 1, PrintBand, True) ;
         if QR2E_DateECHEANCE.AsDateTime>TabDate[4]
            then Calculs(Solde0, MReport[6], MReport[7], 0, PrintBand, True) ;
         END ;
     1 : BEGIN
         if QR2E_DateECHEANCE.AsDateTime<TabDate[4]
            then Calculs(Solde5, MReport[1], MReport[7], 5, PrintBand, True) ;
         if (QR2E_DateECHEANCE.AsDateTime>=TabDate[4]) And (QR2E_DateECHEANCE.AsDateTime<TabDate[3])
            then Calculs(Solde4, MReport[2], MReport[7], 4, PrintBand, True) ;
         if (QR2E_DateECHEANCE.AsDateTime>=TabDate[3]) And (QR2E_DateECHEANCE.AsDateTime<TabDate[2])
            then Calculs(Solde3, MReport[3], MReport[7], 3, PrintBand, True) ;
         if (QR2E_DateECHEANCE.AsDateTime>=TabDate[2]) And (QR2E_DateECHEANCE.AsDateTime<TabDate[1])
            then Calculs(Solde2, MReport[4], MReport[7], 2, PrintBand, True) ;
         if (QR2E_DateECHEANCE.AsDateTime>=TabDate[1]) And (QR2E_DateECHEANCE.AsDateTime<CritEdt.Date1)
            then Calculs(Solde1, MReport[5], MReport[7], 1, PrintBand, True) ;
         if QR2E_DateECHEANCE.AsDateTime>=CritEdt.Date1
            then case CritEdt.GlV.TypePar of
                   0 : Calculs(Solde0, MReport[6], MReport[7], 0, PrintBand, True) ;
                   1 : Calculs(Solde0, MReport[6], MReport[7], 0, PrintBand, False) ;
                 end ;
         END ;
    end;
   END ;
Case CritEdt.GlV.TypePar of
  0 : PrintBand:=(Not QR2E_AUXILIAIRE.IsNull) And Affiche ;
  1 : PrintBand:=(Not QR2E_MODEPAIE.IsNull) and Affiche ;
 end ;
AddReportBAL([1,2], CritEdt.GlV.FormatPrint.Report, MReport, CritEdt.Decimale) ;
if PrintBand then Quoi:=QuoiCpt(1) ;
end;

procedure TFQRGLVen.FNatureCptChange(Sender: TObject);
begin
  inherited;
If QRLoading then Exit ; // Rony 05/06/97
FColl1.clear ; FColl2.clear ; FCollJoker.clear ;
ChangeColl(FNatureCpt.Value,FColl1, FColl2) ;
end;

procedure TFQRGLVen.BFCompteAuxAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([2],CritEdt.GlV.FormatPrint.Report) ;
end;

procedure TFQRGLVen.BNouvRechClick(Sender: TObject);
begin
  inherited;
If FCollJoker.Visible then FCollJoker.Text:='' ;
end;

procedure TFQRGLVen.DLRuptNeedData(var MoreData: Boolean);
var TotRupt          : Array[0..13] of Double ;
    Librupt, CodRupt, Lib1, Cptrupt, stcode          : String ;
    Quellerupt       : Integer ;
    Col              : TColor ;
    OkOk, DansTotal  : Boolean ;
    LibRuptInf : Array[1..10] Of TRuptInf ;
begin
  inherited;
MoreData:=false ;
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
  rRuptures : MoreData:=PrintRupt(LRupt,Qr1T_AUXILIAIRE.AsString,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt) ;
  rCorresp  : BEGIN
              OkOk:=TRUE ;
              Case CritEDT.GlV.PlansCorresp  Of
                1 : If Qr1T_CORRESP1.AsString<>'' Then CptRupt:=Qr1T_CORRESP1.AsString+Qr1T_AUXILIAIRE.AsString
                                                   Else CptRupt:='.'+Qr1T_AUXILIAIRE.AsString ;
                2 : If Qr1T_CORRESP2.AsString<>'' Then CptRupt:=Qr1T_CORRESP2.AsString+Qr1T_AUXILIAIRE.AsString
                                                   Else CptRupt:='.'+Qr1T_AUXILIAIRE.AsString ;
                Else OkOk:=FALSE ;
                END ;
              If OkOk Then MoreData:=PrintRupt(LRupt,CptRupt,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt) Else MoreData:=FALSE ;
              END ;
  End ;
if MoreData then
   BEGIN
   TCodRupt.Caption:='' ;
   if CritEdt.Rupture=rLibres then
      BEGIN
      insert(MsgBox.Mess[22]+' ',CodRupt,Quellerupt+2) ;
      TCodRupt.Caption:=CodRupt+' '+Lib1 ;
      END Else TCodRupt.Caption:=CodRupt+'   '+LibRupt ;
   Libre5.Caption:=PrintSolde(TotRupt[10], TotRupt[11], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
   Libre4.Caption:=PrintSolde(TotRupt[8], TotRupt[9], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
   Libre3.Caption:=PrintSolde(TotRupt[6], TotRupt[7], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
   Libre2.Caption:=PrintSolde(TotRupt[4], TotRupt[5], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
   Libre1.Caption:=PrintSolde(TotRupt[2], TotRupt[3], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
   Libre0.Caption:=PrintSolde(TotRupt[0], TotRupt[1], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
   LibreTotal.Caption:=PrintSolde(TotRupt[12], TotRupt[13], CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
   SautPageRupture:=CritEdt.SautPageRupt And (CritEdt.glV.RuptOnly=Avec) And (CritEdt.SautPage<>1) And SautPageRuptAFaire(CritEdt,BDetail,QuelleRupt) ;
   END ;
end;

procedure TFQRGLVen.BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.Rupture<>rRien) ;
end;

procedure TFQRGLVen.FSansRuptClick(Sender: TObject);
begin
  inherited;
FLigneRupt.Enabled:=Not FSansRupt.Checked ;
FLigneRupt.checked:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
FRupturesClick(Nil) ;
end;

procedure TFQRGLVen.FRupturesClick(Sender: TObject);
begin
  inherited;
If FPlansCo.Checked then FGroupRuptures.Caption:=' '+MsgBox.Mess[25] ;
If FRuptures.Checked then FGroupRuptures.Caption:=' '+MsgBox.Mess[24] ;
end;

procedure TFQRGLVen.BDetailAfterPrint(BandPrinted: Boolean);
begin
  inherited;
If CritEdt.SautPageRupt And (CritEdt.GlV.RuptOnly=Avec) And (CritEdt.SautPage<>1) Then
  BEGIN
  BDetail.ForceNewPage:=FALSE ; SautPageRupture:=FALSE ;
  END ;
end;

procedure TFQRGLVen.BDetailCheckForSpace;
begin
  inherited;
If SautPageRupture Then BDetail.ForceNewPage:=TRUE ;
end;

end.
