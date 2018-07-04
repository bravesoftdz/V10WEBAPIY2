unit QRAFF;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  StdCtrls, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, UtilEdt, CritEdt, QrRupt, Ent1, HEnt1,ParamQR,
  CpteUtil, CPSection_TOM, HSysMenu, Menus, HTB97, HPanel, UiUtil, HStatus, Formule,
  Spin,tcalccum, ADODB;

procedure BLAffaire ;

Type tQRMethode = (tQRNormal,tQRcompte,tQRrequete) ;

type
  TFQRAFF = class(TFQR)
    TitreColCpt: TQRLabel;
    TitreColLibelle: TQRLabel;
    TTitreColAvant: TQRLabel;
    QrC11: TQRLabel;
    QRC12: TQRLabel;
    TTitreColSelection: TQRLabel;
    QRC21: TQRLabel;
    QRC22: TQRLabel;
    TTitreColApres: TQRLabel;
    QRC31: TQRLabel;
    QRC32: TQRLabel;
    QRC41: TQRLabel;
    QRC42: TQRLabel;
    TTitreColSolde: TQRLabel;
    REPORT1DEBITanv: TQRLabel;
    REPORT1CREDITanv: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    REPORT1DEBITsum: TQRLabel;
    REPORT1CREDITsum: TQRLabel;
    REPORT1DEBITsol: TQRLabel;
    REPORT1CREDITsol: TQRLabel;
    S_SECTION: TQRDBText;
    S_LIBELLE: TQRDBText;
    ANvDEBIT: TQRLabel;
    ANvCREDIT: TQRLabel;
    DEBIT: TQRLabel;
    CREDIT: TQRLabel;
    DEBITsum: TQRLabel;
    CREDITsum: TQRLabel;
    SOLDEdeb: TQRLabel;
    SOLDEcre: TQRLabel;
    BRUPT: TQRBand;
    TOTT3: TQRLabel;
    TOTT4: TQRLabel;
    TCodRupt: TQRLabel;
    TOTT7: TQRLabel;
    TOTT8: TQRLabel;
    TOTT5: TQRLabel;
    TOTT6: TQRLabel;
    TOTT1: TQRLabel;
    TOTT2: TQRLabel;
    QRLabel33: TQRLabel;
    TOT2DEBITanv: TQRLabel;
    TOT2CREDITanv: TQRLabel;
    TOT2DEBIT: TQRLabel;
    TOT2CREDIT: TQRLabel;
    TOT2DEBITsum: TQRLabel;
    TOT2CREDITsum: TQRLabel;
    TOT2SOLdeb: TQRLabel;
    TOT2SOLcre: TQRLabel;
    QRBand1: TQRBand;
    Trait2: TQRLigne;
    Trait3: TQRLigne;
    Trait4: TQRLigne;
    Trait5: TQRLigne;
    Trait0: TQRLigne;
    Trait6: TQRLigne;
    Ligne1: TQRLigne;
    REPORT2DEBITanv: TQRLabel;
    REPORT2CREDITanv: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    REPORT2CREDITsum: TQRLabel;
    REPORT2DEBITsum: TQRLabel;
    REPORT2DEBITsol: TQRLabel;
    REPORT2CREDITsol: TQRLabel;
    MsgLibel: THMsgBox;
    DLrupt: TQRDetailLink;
    FCol1: TCheckBox;
    FCol2: TCheckBox;
    FCol3: TCheckBox;
    FCol4: TCheckBox;
    FLigneRupt: TCheckBox;
    FAvecAN: TCheckBox;
    QSup: TQuery;
    TNatLib: TLabel;
    NatLib: THValComboBox;
    OKC1: TCheckBox;
    OkC2: TCheckBox;
    OkC3: TCheckBox;
    OkC4: TCheckBox;
    SPC1: TEdit;
    SPC2: TEdit;
    SPC3: TEdit;
    SPC4: TEdit;
    BDSMulti: TQRBand;
    TOTST3: TQRLabel;
    TOTST4: TQRLabel;
    TCodeST: TQRLabel;
    TOTST7: TQRLabel;
    TOTST8: TQRLabel;
    TOTST5: TQRLabel;
    TOTST6: TQRLabel;
    TOTST1: TQRLabel;
    TOTST2: TQRLabel;
    QRDLMulti: TQRDetailLink;
    BRecapG: TQRBand;
    TOTR3: TQRLabel;
    TOTR4: TQRLabel;
    TCodeRecap: TQRLabel;
    TOTR7: TQRLabel;
    TOTR8: TQRLabel;
    TOTR5: TQRLabel;
    TOTR6: TQRLabel;
    TOTR1: TQRLabel;
    TOTR2: TQRLabel;
    QRDLRecapG: TQRDetailLink;
    SautDePage: TCheckBox;
    NbCharRupt: TSpinEdit;
    TNbCharRupt: TLabel;
    procedure FormShow(Sender: TObject);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean;      var Quoi: string);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BRUPTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean;      var Quoi: string);
    procedure DLruptNeedData(var MoreData: Boolean);
    procedure FNatureCptChange(Sender: TObject);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean;      var Quoi: string);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure FExerciceChange(Sender: TObject);
    procedure FDateCompta1Change(Sender: TObject);
    procedure FSelectCpteChange(Sender: TObject);
    procedure FAvecANClick(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure FSansRuptClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure OKC1Click(Sender: TObject);
    procedure OkC2Click(Sender: TObject);
    procedure OkC3Click(Sender: TObject);
    procedure OkC4Click(Sender: TObject);
    procedure SPC1DblClick(Sender: TObject);
    procedure SPC1Change(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
    procedure QRDLMultiNeedData(var MoreData: Boolean);
    procedure BRecapGBeforePrint(var PrintBand: Boolean; var Quoi: String);
    procedure QRDLRecapGNeedData(var MoreData: Boolean);
    procedure BRecapGAfterPrint(BandPrinted: Boolean);
    procedure BDSMultiBeforePrint(var PrintBand: Boolean;
      var Quoi: String);
    procedure BDetailAfterPrint(BandPrinted: Boolean);
    procedure BDetailCheckForSpace;
  private    { Déclarations privées }
    LRupt,LMulti,LRecapG      : TStringList ;
    ListeCodesRupture : TStringList ;
    ListeCalc         : TStringList ;
    // Rony 17/03/97 QBal,QBalCum  : TQuery ;
    TotEdt     : TabTot ;
    DansTotal         : Boolean ;
    Qr11S_AXE         : TStringField ;
    Qr11S_SECTION     : TStringField ;
    Qr11S_LIBELLE     : TStringField ;
    Qr11S_SECTIONTRIE : TStringField ;
    Qr11S_CORRESP1, Qr11S_CORRESP2: TStringField ;
    FLoading          : Boolean ;
    CritEdtSup        : TCritEdtSup ;
    QRMethode         : tQRMethode ;
    DC4Formule        : TabDC4 ;
    OkImprimeSt       : Boolean ;
    AfterPrintDetail  : Boolean ;
    procedure GenereSQL ; Override ;
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
    Procedure BalSectionZoom(Quoi : String) ;
    Procedure CalcTotalEdt(D1,C1,D2,C2,D3,C3,D4,C4 : Double) ;
    Function  OkPourRupt : Boolean ;
    Function  ColonnesOK : Boolean ;
    function  SQLModeSelect1 : String ;
    procedure GenereSQLSup ;
    procedure ChargeNatLib ;
    procedure RecupCritEdtSup ;
    Procedure RecupCritCol(Var PC : tParamCol ; OkC : Boolean ; St : String) ;
    Procedure ChangeTitreCB(FCol : TCheckBox ; SPC : TEdit ; OkC : Boolean ; i : Integer) ;
    Procedure INITCHAMPTABLE(LibreOrder : String) ;
    Function  ChercheMontantFormule(i,j : Integer ; Var CritEdtSup : TCritEdtSup ; Var TDC : TabDC4) : Boolean ;
    function  Load_QR ( St : String ) : Variant ;
    Procedure EstLigneSeule ;
    Procedure ReinitFormatBal ;
  public    { Déclarations publiques }
  end;

implementation

{$R *.DFM}
Uses Filtre ;

procedure BLAffaire ;
var QR : TFQRAFF ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
QR:=TFQRAFF.Create(Application) ;
Edition.Etat:=etBalAna ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BalSectionZoom ;
QR.InitType (nbSec,neBal,msSecAna,'QRBALANASPECIF','',TRUE,FALSE,FALSE,Edition) ;
PP:=FindInsidePanel ;
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

Procedure TFQRAFF.BalSectionZoom(Quoi : String) ;
Var LP : Integer ;
    St,St1 : String ;
    Lefb : TFichierBase ;
BEGIN
Inherited ;
If QRP.QrPrinter.FSynShiftDblClick Then BEGIN ZoomEdt(3,Quoi) ; Exit ; END ;
Lp:=Pos('@',Quoi) ;  If Lp=0 Then Exit ;
St:=Quoi ; St1:=Copy(Quoi,Lp+2,2) ; LeFb:=AxeToFb(St1) ;
St:=Copy(Quoi,1,VH^.Cpta[Lefb].Lg) ;
CritEdt.Cpt1:=St ;
//MultiCritereAnaZoom(taModif,CritEdt);
END ;

Procedure TFQRAFF.FinirPrint ;
BEGIN
Inherited ;
VideGroup(LMulti) ; VideRecap(LRecapG) ;
if CritEdt.Rupture<>rRien then VideRupt(LRupt) ;
if QRMethode=tQRNormal Then GCalc.Free Else BEGIN ListeCalc.Clear ; ListeCalc.Free ; END ;
//QSup.Close ;
(* QBal.Free ; Rony 17/03/97
QBalCum.Free ; *)
END ;

Procedure TFQRAFF.ReinitFormatBal ;
Var T1,T2 : String ;
    LeTag,LeTag2 : Integer ;
    LeLeft,LeWidth : Integer ;
    T   : TComponent ;
    C   : TQRCustomControl ;
    i,j : Integer ;
BEGIN
For i:=1 To 4 Do
  BEGIN
  T1:=CritEdtSup.PC[i].Titre1 ; T2:=CritEdtSup.PC[i].Titre2 ;
  If CritEdtSup.PC[i].OkOk And ((T1='') Or (T2='')) Then
    BEGIN
    Case i Of
      1 : BEGIN LeLeft:=TTitreColAvant.Left ; LeWidth:=TTitreColAvant.Width ; END ;
      2 : BEGIN LeLeft:=TTitreColSelection.Left ; LeWidth:=TTitreColSelection.Width ; END ;
      3 : BEGIN LeLeft:=TTitreColApres.Left ; LeWidth:=TTitreColApres.Width ; END ;
      4 : BEGIN LeLeft:=TTitreColSolde.Left ; LeWidth:=TTitreColSolde.Width ; END ;
      Else BEGIN LeLeft:=0 ; LeWidth:=0 ; END ;
      END ;
    LeTag:=0 ;
    If T1='' Then BEGIN LeTag2:=(i*2)-1+2 ; LeTag:=(i*2)+2 ; END Else If T2='' Then BEGIN LeTag2:=(i*2)+2 ; LeTag:=(i*2)-1+2 ; END ;
    If (LeTag<>0) And (LeLeft<>0) And (LeWidth<>0) Then
      BEGIN
      Case i Of
        1 : If LeTag=3 Then BEGIN QRC11.Left:=LeLeft ; QRC11.Width:=LeWidth ; QRC12.visible:=FALSE ; END
                       Else BEGIN QRC12.Left:=LeLeft ; QRC12.Width:=LeWidth ; QRC11.visible:=FALSE ; END ;
        2 : If LeTag=5 Then BEGIN QRC21.Left:=LeLeft ; QRC21.Width:=LeWidth ; QRC22.visible:=FALSE ; END
                       Else BEGIN QRC22.Left:=LeLeft ; QRC22.Width:=LeWidth ; QRC21.visible:=FALSE ; END ;
        3 : If LeTag=7 Then BEGIN QRC31.Left:=LeLeft ; QRC31.Width:=LeWidth ; QRC32.visible:=FALSE ; END
                       Else BEGIN QRC32.Left:=LeLeft ; QRC32.Width:=LeWidth ; QRC31.visible:=FALSE ; END ;
        4 : If LeTag=9 Then BEGIN QRC41.Left:=LeLeft ; QRC41.Width:=LeWidth ; QRC42.visible:=FALSE ; END
                       Else BEGIN QRC42.Left:=LeLeft ; QRC42.Width:=LeWidth ; QRC41.visible:=FALSE ; END ;
        END ;
      for j:=0 to Self.ComponentCount-1 do
        BEGIN
        T:=Self.Components[j] ;
        if (T is TQRLabel) Or (T Is TQRDBText) then
           BEGIN
           C:=TQRCustomControl(T) ;
           If C.Tag=LeTag Then
             BEGIN
             C.Left:=LeLeft ; //C.Width:=LeWidth ;
//             If (T is TQRLabel) And (TQRLabel(C).SynColGroup=0) Then TQRLabel(C).Alignment:=taCenter ;
             END ;
           If C.Tag=LeTag2 Then BEGIN C.Visible:=FALSE ; END ;
          END ;
        END ;
      END ;
    END ;
  END ;
END ;

procedure TFQRAFF.InitDivers ;
BEGIN
Inherited ;
{ Init du Total Général }
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
{ Labels sur les bandes }
(*
DateBandeauBalance(TTitreColAvant,TTitreColSelection,TTitreColApres,TTitreColSolde,CritEdt,
                   MsgLibel.Mess[0],MsgLibel.Mess[1],MsgLibel.Mess[2]) ;
*)
If CritEdtSup.PC[1].OkOk Then
  BEGIN
  TTitreColAvant.Caption:=CritEdtSup.PC[1].Titre ; QRC11.Caption:=CritEdtSup.PC[1].Titre1 ; QRC12.Caption:=CritEdtSup.PC[1].Titre2 ;
  END ;
If CritEdtSup.PC[2].OkOk Then
  BEGIN
  TTitreColSelection.Caption:=CritEdtSup.PC[2].Titre ; QRC21.Caption:=CritEdtSup.PC[2].Titre1 ; QRC22.Caption:=CritEdtSup.PC[2].Titre2 ;
  END ;
If CritEdtSup.PC[3].OkOk Then
  BEGIN
  TTitreColApres.Caption:=CritEdtSup.PC[3].Titre ; QRC31.Caption:=CritEdtSup.PC[3].Titre1 ; QRC32.Caption:=CritEdtSup.PC[3].Titre2 ;
  END ;
If CritEdtSup.PC[4].OkOk Then
  BEGIN
  TTitreColSolde.Caption:=CritEdtSup.PC[4].Titre ; QRC41.Caption:=CritEdtSup.PC[4].Titre1 ; QRC42.Caption:=CritEdtSup.PC[4].Titre2 ;
  END ;
BRupt.Frame.DrawTop:=CritEDT.BAL.FormatPrint.PrSepRupt  ;
BRupt.Frame.DrawBottom:=BRupt.Frame.DrawTop ;
ChangeFormatCompte(AxeToFb(CritEdt.Bal.Axe),AxeToFb(CritEdt.Bal.Axe),Self,S_SECTION.Width,1,2,S_SECTION.Name) ;
ReinitFormatBal ;
BRecapG.ForceNewPage:=TRUE ;
END ;

function  TFQRAFF.SQLModeSelect1 : String ;
{ Contruction d'une partie de SQLAux }
Var TOTDEBIT,TOTCREDIT,TOTDEBITANO,TOTCREDITANO,St,St1 : String ;
    PreB,PreE,PreJ,ResC,BilC,PerC,BenC,BilO,BenO,PerO : String ;
    Exo : TExoDate ;
    ts : Integer ;
    EdtSimple : Boolean ;
    Valeur : String ;
    DD1,DD2 : TDateTime ;
    WhatAN : TQuelAN ;
    QClo : TQuery ;
BEGIN
//Simon
// GP : à revoir dans le cas : Que les révisions
Result:='' ; EdtSimple:=TRUE ; WhatAN:=AvecAN ;
WhatAN:=CritEdt.Bal.QuelAN ;
PreJ:='SECTION' ; PreB:='S_' ; PreE:='Y_' ;
Ts:=CritEDT.BAL.TypCpt ; Exo:=CritEDT.Exo ; Valeur:=CritEdt.QualifPiece ;
TOTDEBIT:='' ; TOTCREDIT:='' ; TOTDEBITANO:='' ; TOTCREDITANO:='' ;
//Type ttTypePiece  = (tpReel,tpSim,tpPrev,tpSitu,tpRevi) ;
//TS = 0 : Exo, 1 = Tous, 2 = Cpt non soldés, 3 = Période
Case TS of
  0 : BEGIN
      if (Valeur<>'TOU') Then
         BEGIN
         St:=SQLExiste(PreE,PreJ,PreB,TOTDEBIT,TOTCREDIT,TOTDEBITANO,TOTCREDITANO,Valeur,V_PGI.Controleur,AvecRevision.State,CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,CritEdt.DEVPourExist) ;
         END Else
         BEGIN
         St:=WhatExiste(PreE,PreJ,PreB,Valeur,V_PGI.Controleur,AvecRevision.State,CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,0,CritEdt.DEVPourExist) ;
         END ;
      Result:=St ;
      END ;
  1 : BEGIN Result:=PreB+PreJ+'='+PreB+PreJ+' ' ; END ;
  2 : BEGIN
      if (Valeur<>'TOU') Then
         BEGIN
         if Valeur='NOR' then
            BEGIN
            // GP : Et les révisions ???
            Result:=PreB+'TOTALDEBIT<>'+PreB+'TOTALCREDIT ' ;
            Result:=Result+' Or (('+PreB+'TOTALDEBIT='+PreB+'TOTALCREDIT '+') And (' ;
            Result:=Result+WhatExiste(PreE,PreJ,PreB,Valeur,V_PGI.Controleur,AvecRevision.State,CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,0,CritEdt.DEVPourExist)+'))' ;
            END ;
         if Valeur='NSS' then
            BEGIN
            Result:=PreB+'TOTALDEBIT<>'+PreB+'TOTALCREDIT ' ;
            Result:=Result+' Or (('+PreB+'TOTALDEBIT='+PreB+'TOTALCREDIT '+') And (' ;
            Result:=Result+WhatExiste(PreE,PreJ,PreB,Valeur,V_PGI.Controleur,AvecRevision.State,CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,0,CritEdt.DEVPourExist)+'))' ;
            END ;
         END Else
         BEGIN
         Result:=PreB+'TOTALDEBIT<>'+PreB+'TOTALCREDIT ' ;
         Result:=Result+' Or (('+PreB+'TOTALDEBIT='+PreB+'TOTALCREDIT '+') And (' ;
         Result:=Result+WhatExiste(PreE,PreJ,PreB,Valeur,V_PGI.Controleur,AvecRevision.State,CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,0,CritEdt.DEVPourExist)+'))' ;
         END ;
      If Result<>'' Then Result:=' ('+Result+') ' ;
      END ;
  3 : BEGIN
      DD1:=CritEdt.DateDeb ; DD2:=CritEdt.DateFin ;
      If EdtSimple And (WhatAN=SansAN) Then BEGIN TOTDEBITANO:='' ; TOTCREDITANO:='' ; END ;
      If EdtSimple And (WhatAN=AvecAN) And (FNatureBase=nbGen) Then BEGIN DD1:=CritEdt.Date1 ; END ;
      if (Valeur<>'TOU') Then
         BEGIN
         St:=SQLExiste(PreE,PreJ,PreB,'','',TOTDEBITANO,TOTCREDITANO,Valeur,V_PGI.Controleur,AvecRevision.State,DD1,DD2,CritEdt.Exo.Code,CritEdt.DEVPourExist) ;
         END Else
         BEGIN
         St:=WhatExiste(PreE,PreJ,PreB,Valeur,V_PGI.Controleur,AvecRevision.State,DD1,DD2,CritEdt.Exo.Code,0,CritEdt.DEVPourExist) ;
         END ;
      Result:=St ;
      END ;
 end ;
END ;

Procedure MajListeCalc(QSup : TQuery  ; ListeCalc : TStringList ; Var CritEdt : TCritEdt ; Var CritEdtSup : TCritEdtSup) ;
Var DD,CD,DE,CE : Double ;
    i : Integer ;
    St,St1 : String ;
BEGIN
St:='' ;
For i:=0 To 9 Do If CritEdtSup.ChampTable[i]<>'' Then
  BEGIN
  St:=St+QSup.FindField(CritEdtSup.ChampTable[i]).AsString+';' ;
  END ;
St:=St+QSup.FindField(CritEdt.CptLibre1).AsString+';' ;
St:=St+QSup.FindField('Y_SECTION').AsString+';' ;
DD:=QSup.FindField('DD').AsFloat ; CD:=QSup.FindField('CD').AsFloat ;
//DE:=QSup.FindField('DE').AsFloat ; CE:=QSup.FindField('CE').AsFloat ;
St1:=FloattoStr(DD) ;  St:=St+St1+';' ;
St1:=FloattoStr(CD) ;St:=St+St1+';' ;
//St1:=FloattoStr(DE) ;St:=St+St1+';' ;
//St1:=FloattoStr(CE) ;St:=St+St1+';' ;
ListeCalc.Add(St) ;
//showmessage('ALIM '+St+' / ') ;
//showmessage('ALIM '+FloatToStr(QSup.FindField('DD').AsFloat)+' / '+FloatToStr(QSup.FindField('CD').AsFloat)) ;
END ;

Function FindDansListeCalc(Var CritEdtSup : TCritEdtSup ; St : String ;Var ValChamp : tChampTable11 ; Var TDC : TabTot) : Boolean ;
Var ValChampSt : tChampTable11 ;
    St1 : String ;
    i : Integer ;
    OkOk : Boolean ;
    St2 : String ;
    XX : Double ;
BEGIN
Result:=FALSE ; St1:=St ; Fillchar(ValChampSt,SizeOf(ValChampSt),#0) ;
For i:=0 To 9 Do If CritEdtSup.ChampTable[i]<>'' Then ValChampSt[i]:=ReadTokenSt(St1) ;
ValChampSt[10]:=ReadTokenSt(St1) ; ValChampSt[11]:=ReadTokenSt(St1) ;
Okok:=TRUE ;
For i:=0 To 11 Do
  BEGIN
  If ValChampSt[i]<>ValChamp[i] Then BEGIN OkOk:=FALSE ; Break ; END ;
  END ;
Result:=OkOk ;
If Result Then
  BEGIN
  St2:=ReadTokenSt(St1) ; XX:=StrToFloat(St2) ; TDC[0].TotDebit:=XX ;  // Debit dev
  St2:=ReadTokenSt(St1) ; XX:=StrToFloat(St2) ;TDC[0].TotCredit:=XX ; // Crédit dev
  St2:=ReadTokenSt(St1) ; XX:=StrToFloat(St2) ; TDC[1].TotDebit:=XX ;  // Debit opposé
  St2:=ReadTokenSt(St1) ; XX:=StrToFloat(St2) ; TDC[1].TotCredit:=XX ; // Crédit opposé
//showmessage('FIND '+St+' / ') ;
//showmessage('FIND '+FloatToStr(TDC[0].TotDebit)+' / '+FloatToStr(TDC[0].TotCredit)) ;
  END ;
END ;

Function ChercheMontantListeCalc(Q : TQuery  ; Value2 : String ; ListeCalc : TStringList ;
                                  Var CritEdt : TCritEdt ; Var CritEdtSup : TCritEdtSup ;
                                  Var TDC : TabTot) : Boolean ;
Var DD,CD,DE,CE : Double ;
    i : Integer ;
    St,St1 : String ;
    ValChamp : tChampTable11 ;
BEGIN
Result:=FALSE ;
Fillchar(TDC,SizeOf(TDC),#0) ; Fillchar(ValChamp,SizeOf(ValChamp),#0) ;
For i:=0 To 9 Do If CritEdtSup.ChampTable[i]<>'' Then ValChamp[i]:=Q.FindField(CritEdtSup.ChampTable[i]).AsString ;
ValChamp[10]:=Value2 ;
ValChamp[11]:=St+Q.FindField('S_SECTION').AsString ;
For i:=0 To ListeCalc.Count-1 Do
  BEGIN
  If FindDansListeCalc(CritEdtSup,ListeCalc[i],ValChamp,TDC) Then BEGIN Result:=TRUE ; Break ; END ;
  END ;
END ;

function TFQRAFF.Load_QR ( St : String ) : Variant ;
Var V    : Variant ;
BEGIN
V:=#0 ; Result:=V ;
St:=uppercase(Trim(St)) ; if St='' then Exit ;
If St='11' Then V:=DC4Formule[1].TotDebit Else
If St='12' Then V:=DC4Formule[1].TotCredit Else
If St='21' Then V:=DC4Formule[2].TotDebit Else
If St='22' Then V:=DC4Formule[2].TotCredit Else
If St='31' Then V:=DC4Formule[3].TotDebit Else
If St='32' Then V:=DC4Formule[3].TotCredit Else
If St='41' Then V:=DC4Formule[4].TotDebit Else
If St='42' Then V:=DC4Formule[4].TotCredit Else Exit ;
Load_QR:=V ;
END ;

Function TFQRAFF.ChercheMontantFormule(i,j : Integer ; Var CritEdtSup : TCritEdtSup ; Var TDC : TabDC4) : Boolean ;
Var Formule : String ;
    XX : Double ;
    St : String ;
BEGIN
Result:=TRUE ;
Fillchar(DC4Formule,SizeOf(DC4Formule),#0) ;
If i=1 Then Formule:=CritEdtSup.PC[j].Formule1 Else Formule:=CritEdtSup.PC[j].Formule2 ;
DC4Formule:=TDC ;
St:=GFormule(Formule,Load_QR,Nil,1) ;
If V_PGI.ErreurFormule Then BEGIN Result:=FALSE ; XX:=999.99 ; END Else XX:=Valeur(St) ;
If i=1 Then TDC[j].TotDebit:=XX Else TDC[j].TotCredit:=XX ;
END ;

procedure TFQRAFF.GenereSQLSup ;
Var Q   : TQuery ;
    StCalc : String ;
    SUM_Montant : String ;
    P : String[2] ;
    STABLE : String ;
    i : Integer ;
    Pref : String ;
    GTABLE : String ;
    fb : TFichierBase ;

BEGIN
If QRMethode=tQRNormal Then Exit ;
STABLE:=''; GTABLE:=CritEdt.CptLibre1 ; fb:=AxeTofb(CritEdt.Bal.Axe) ; Pref:='Y_' ;
(*
DecodeOkNatLibre(CritEdt.LibreCodes1,OkNatLibre) ;
For i:=0 To 9 Do If OkNatLibre[i+1] Then STABLE:=STABLE+' '+Pref+'TABLE'+IntToStr(i)+', ' ;
*)
For i:=0 To 9 Do If CritEdtSup.ChampTable[i]<>'' Then STABLE:=STABLE+CritEdtSup.ChampTable[i]+', ';
QSup.SQL.Clear ;
If QRMethode=tQRCompte Then
  BEGIN
  Case Fb Of
    FbGene : StCalc:='SELECT E_GENERAL, E_EXERCICE, ' ;
    FbAux  : StCalc:='SELECT E_AUXILIAIRE, E_EXERCICE, ';
    fbAxe1..fbAxe5 : StCalc:='SELECT Y_SECTION, Y_EXERCICE, ' ;
    fbJal : StCalc:='SELECT E_JOURNAL, E_EXERCICE, ' ;
    END ;
  END Else
  BEGIN
  Case Fb Of
    fbAxe1..fbAxe5 : StCalc:='SELECT '+STABLE+' '+GTABLE+', Y_SECTION, ' ;
    END ;
  END ;
QSup.SQL.Add(StCalc) ;
Case fb Of
  fbGene,fbAux,fbJal : If EnDevise(Dev,DevEnP) Then SUM_Montant:='sum(E_DEBITDEV) AS DD, sum(E_CREDITDEV) AS CD'
                                                            Else SUM_Montant:='sum(E_DEBIT) AS DD, sum(E_CREDIT) AS CD' ;
  fbAxe1..fbAxe5 : If EnDevise(Dev,DevEnP) Then SUM_Montant:='sum(Y_DEBITDEV) AS DD, sum(Y_CREDITDEV) AS CD'
                                                                     Else SUM_Montant:='sum(Y_DEBIT) AS DD, sum(Y_CREDIT) AS CD ' ;
  END ;
Case fb Of
  fbGene,fbAux,fbJal : StCalc:=SUM_MONTANT+' FROM ECRITURE ' ;
  fbAxe1..fbAxe5 : StCalc:=SUM_MONTANT+' FROM ANALYTIQ ' ;
  END ;
QSup.SQL.Add(StCalc) ;
StCalc:='' ;
Case fb Of
  fbAxe1..fbAxe5 : StCalc:=' LEFT JOIN SECTION ON S_SECTION=Y_SECTION '+' LEFT JOIN GENERAUX ON G_GENERAL=Y_GENERAL ' ;
  END ;
If StCalc<>'' Then QSup.SQL.Add(StCalc) ;
If QRMethode=tQRCompte Then
  BEGIN
  Case Fb Of
    fbAxe1..fbAxe5 :
      BEGIN
      StCalc:='WHERE Y_SECTION=:CPTE AND Y_AXE="'+CritEdt.Bal.Axe+'" AND Y_DATECOMPTABLE>="'+USDATETIME(CritEdt.Date1)+'" AND Y_DATECOMPTABLE<="'+USDATETIME(CritEdt.Date2)+'" ' ;
      END ;
    END ;
  END Else
  BEGIN
  Case Fb Of
    fbAxe1..fbAxe5 :
      BEGIN
      StCalc:='WHERE Y_AXE="'+CritEdt.Bal.Axe+'" AND Y_DATECOMPTABLE>="'+USDATETIME(CritEdt.Date1)+'" AND Y_DATECOMPTABLE<="'+USDATETIME(CritEdt.Date2)+'" ' ;
      StCalc:=StCalc+SQLQuelCpt ;
      END ;
    END ;
  END ;
QSup.SQL.Add(StCalc) ;
If Dev Then
   BEGIN
   Case Fb Of
     FbGene,FbAux,fbJal : StCalc:='AND E_DEVISE="'+CritEdt.DeviseSelect+'" ' ;
     fbAxe1..fbAxe5 : StCalc:='AND Y_DEVISE="'+CritEdt.DeviseSelect+'" ' ;
     END ;
   QSup.SQL.Add(StCalc) ;
   END ;
If Etab Then
   BEGIN
   Case Fb Of
     FbGene,FbAux,fbJal : StCalc:='AND E_ETABLISSEMENT="'+CritEdt.Etab+'" ' ;
     fbAxe1..fbAxe5 : StCalc:='AND Y_ETABLISSEMENT="'+CritEdt.Etab+'" ' ;
     END ;
   QSup.SQL.Add(StCalc) ;
   END ;
If Exo Then
   BEGIN
   Case Fb Of
     FbGene,FbAux,fbJal : StCalc:='AND E_EXERCICE="'+CritEdt.Exo.Code+'" ' ;
     fbAxe1..fbAxe5 : StCalc:='AND Y_EXERCICE="'+CritEdt.Exo.Code+'" '  ;
     END ;
   QSup.SQL.Add(StCalc) ;
   END ;
P:='E_' ; If fb In [fbAxe1..fbAxe5] Then P:='Y_' ;
StCalc:=WhereSupp(P,QuelTypeEcr) ; If StCalc<>'' Then QSup.SQL.Add(StCalc) ;
IF CritEdt.Bal.OnlyCptAssocie then
  BEGIN
  StCalc:=WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,AxeToFb(CritEdt.Bal.Axe),CritEdt.Bal.OnlyCptAssocie) ;
  If StCalc<>'' Then QSup.SQL.Add(StCalc) ;
  END ;
StCalc:='' ;
If CritEdt.CptLibre1<>'' Then For i:=1 To 4 Do If CritEdtSup.PC[i].OkOk Then
  BEGIN
  If CritEdtSup.PC[i].Value1<>'' Then StCalc:=StCalc+' OR '+CritEdt.CptLibre1+'="'+CritEdtSup.PC[i].VAlue1+'" ' ;
  If CritEdtSup.PC[i].Value2<>'' Then StCalc:=StCalc+' OR '+CritEdt.CptLibre1+'="'+CritEdtSup.PC[i].VAlue2+'" ' ;
  END ;
If StCalc<>'' Then
  BEGIN
  delete(StCalc,1,3) ;
  StCalc:=' AND ('+StCalc+') ' ; QSup.SQL.Add(StCalc) ;
  END ;
Case Fb Of
  fbAxe1..fbAxe5 : StCalc:='GROUP BY '+STABLE+' '+GTABLE+',Y_SECTION ' ;
  END ;
QSup.SQL.Add(StCalc) ;
ChangeSQL(QSup) ;
If QRMethode=tQRCompte Then QSup.Prepare Else
  BEGIN
  QSup.Open ;
  InitMove(RecordsCount(QSup),'') ;
  While Not QSup.Eof Do
    BEGIN
    MoveCur(FALSE) ; MajListeCalc(QSup,ListeCalc,CritEdt,CritEdtSup) ;
    QSup.Next ;
    END ;
  FiniMove ;
  QSup.Close ;
  END ;
END ;

procedure TFQRAFF.GenereSQL ;
{ Construction de la requête SQL en fonction du multicritère }
Var LeSauf,LeSQLSauf,Axe,NatCpt,St, StJalAuto : String ;
    StSelect1,stSelect2,stSelect3,StFrom,StCpt,StOrder,StWhere1 : String ;
    ITypCpt : Integer ;
    ERuptOnly : TRuptEtat ;
    NatureRupt : String3 ;
    P          : String ;
    EOnlyCompteAssocie : Boolean ;
    EPlansCorresp : Byte ;
BEGIN
Inherited ;
Q.Close ; Q.SQL.Clear ; ERuptOnly:=Sans ;
StSelect1:='' ; StSelect2:='' ; StFrom:='' ; StCpt:='' ; StOrder:='' ;
StWhere1:='' ; stSelect3:='' ; EOnlyCompteAssocie:=FALSE ;
EPlansCorresp:=0 ;
Axe:=CritEdt.Bal.Axe ; ITypCpt:=CritEDT.BAL.TypCpt ; NatCpt:=CritEDT.NatureCpt ;
LeSauf:=CritEDT.BAL.Sauf ; LeSQLSauf:=CritEdt.Bal.SQLSauf ;
ERuptOnly:=CritEdt.BAl.RuptOnly ;
EOnlyCompteAssocie:=CritEdt.Bal.OnlyCptAssocie ;
EPlansCorresp:=CritEDT.Bal.PlansCorresp ;
StSelect1:='Select S_AXE, S_SECTION, S_LIBELLE, S_TOTDEBANO, S_TOTCREANO, S_TOTDEBE, S_TOTCREE,S_SECTIONTRIE, S_TOTDEBANON1, S_TOTCREANON1 ' ;
Case CritEdt.Rupture of
  rLibres  : BEGIN
             StSelect1:='Select S_AXE, S_SECTION, S_LIBELLE, S_TOTDEBANO, S_TOTCREANO, S_TOTDEBE, S_TOTCREE,S_SECTIONTRIE, S_TOTDEBANON1, S_TOTCREANON1 ' ;
             StSelect3:=', S_TABLE0, S_TABLE1, S_TABLE2, S_TABLE3, S_TABLE4, S_TABLE5, S_TABLE6, S_TABLE7, S_TABLE8, S_TABLE9 ' ;
             END ;
  rCorresp : StSelect1:='Select S_AXE, S_SECTION, S_LIBELLE, S_TOTDEBANO, S_TOTCREANO, S_TOTDEBE, S_TOTCREE,S_SECTIONTRIE, S_TOTDEBANON1, S_TOTCREANON1, S_CORRESP1, S_CORRESP2 ' ;
  End ;
StFrom:=' From SECTION Q Where ' ;
StWhere1:=' AND S_AXE="'+Axe+'" ' ;
if (OkV2 and (V_PGI.Confidentiel<>'1')) then StWhere1:=StWhere1+'AND S_CONFIDENTIEL<>"1" ' ;
StOrder:=' Order By S_AXE, S_SECTION ' ;
Case CritEdt.Rupture of
  rLibres   : BEGIN
              IF EOnlyCompteAssocie then StWhere1:=StWhere1+WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,AxeToFb(Axe),EOnlyCompteAssocie) ;
              StOrder:='Order By '+OrderLibre(CritEdt.LibreTrie)+'S_SECTION' ;
              END ;
  rRuptures : StOrder:=' Order By S_AXE, S_SECTIONTRIE ' ;
  rCorresp  : Case EPlansCorresp Of
                1 : StOrder:=' Order By S_AXE, S_CORRESP1, S_SECTION' ;
                2 : StOrder:=' Order By S_AXE, S_CORRESP2, S_SECTION' ;
                Else StOrder:=' Order By S_AXE, S_SECTION ' ;
                End ;
  End ;
If StSelect1<>'' Then Q.SQL.Add(StSelect1) ;
If StSelect2<>'' Then Q.SQL.Add(StSelect2) ;
If StSelect3<>'' Then Q.SQL.Add(StSelect3) ;
If StFrom<>'' Then Q.SQL.Add(StFrom) ;
St:=SQLModeSelect1 ; If (St<>'') Then Q.SQL.Add(St) ;
St:=SQLQuelCpt ; If St<>'' Then Q.SQL.Add(St) ;
If StWhere1<>'' Then Q.SQL.Add(StWhere1) ;
if LeSauf<>'' then Q.SQL.Add(LeSQLSauf) ;
Case CritEdt.Rupture of
  rCorresp : BEGIN
//             if FOnlyCptAssocie.Checked then
             If EOnlyCompteAssocie Then
                BEGIN
                Case FNatureBase Of
                  nbGen : BEGIN P:='G' ;NatureRupt:='GE'+IntToStr(EPlansCorresp) ; END ;
                  nbAux : BEGIN P:='T' ;NatureRupt:='AU'+IntToStr(EPlansCorresp) ; END ;
                  nbSec : BEGIN P:='S' ;NatureRupt:=Axe+IntToStr(EPlansCorresp) ; END ;
                  END ;
                Q.SQL.Add(' AND Exists(SELECT CR_CORRESP FROM CORRESP WHERE CR_TYPE="'+NatureRupt+'"') ;
                Q.SQL.Add(' AND CR_CORRESP=Q.'+P+'_CORRESP'+IntToStr(EPlansCorresp)+')') ;
                END ;
              END ;
  End ;
If StOrder<>'' Then Q.SQL.Add(StOrder) ;
ChangeSQL(Q) ; Q.Open ;
END;

procedure TFQRAFF.RenseigneCritere ;
{ Récupération des champs du multicritère dans l'entête d'état }
BEGIN
Inherited ;
(*if FSansRupt.Checked=True then BEGIN RSansRupt.Caption:='þ' ; RAvecRupt.Caption:='o' ; RSurRupt.Caption:='o' ; END Else
if FAvecRupt.Checked=True then BEGIN RSansRupt.Caption:='o' ; RAvecRupt.Caption:='þ' ; RSurRupt.Caption:='o' ; END Else
if FSurRupt.Checked=True  then BEGIN RSansRupt.Caption:='o' ; RAvecRupt.Caption:='o' ; RSurRupt.Caption:='þ' ; END ;
If FPlanRupt.Enabled=True then BEGIN RPlanRupt.Caption:=FPlanRupt.Text ; RTriRupt1.Caption:=FTriRupt1.Text ; RTriRupt2.Caption:=FTriRupt2.Text ; END ;
if not CritEdt.Bal.RuptOnly Then begin RTriRupt1.Caption:='' ; RTriRupt2.Caption:='' ; End ;
CaseACocher(FCol1, RCol1) ; CaseACocher(FCol2, RCol2) ;
CaseACocher(FCol3, RCol3) ; CaseACocher(FCol4, RCol4) ;
CaseACocher(FLigneRupt, RLigneRupt) ;*)
END ;

procedure TFQRAFF.ChoixEdition ;
BEGIN
Inherited ;
Case CritEdt.Rupture of
  rLibres   : BEGIN
              DLrupt.PrintBefore:=FALSE ;
              ChargeGroup(LRupt,['S00','S01','S02','S03','S04','S05','S06','S07','S08','S09']) ;
              END ;
  rCorresp : BEGIN
             DLrupt.PrintBefore:=TRUE ;
             ChargeRuptCorresp(LRupt,CritEdt.Bal.PlanRupt,CritEdt.Bal.CodeRupt1 ,CritEdt.Bal.CodeRupt2, False) ;
             NiveauRupt(LRupt) ;
             END ;
  End ;
ChargeGroup(LMulti,[MsgLibel.Mess[16]]) ; ChargeRecap(LRecapG) ;
END ;

Procedure TFQRAFF.RecupCritCol(Var PC : tParamCol ; OkC : Boolean ; St : String) ;
Var St1 : String ;
BEGIN
Fillchar(PC,SizeOf(PC),#0) ;
PC.OkOk:=OkC ;
If OkC Then
  BEGIN
  St:=Trim(St) ; If St='' Then Exit ;
  If St<>'' Then PC.Titre:=ReadTokenPipe(St,Pipe) ;
  If St<>'' Then PC.Value1:=ReadTokenPipe(St,Pipe) ;
  If St<>'' Then PC.Titre1:=ReadTokenPipe(St,Pipe) ;
  If St<>'' Then PC.Value2:=ReadTokenPipe(St,Pipe) ;
  If St<>'' Then PC.Titre2:=ReadTokenPipe(St,Pipe) ;
  If St<>'' Then PC.DebitPos1:=(ReadTokenPipe(St,Pipe)='X') ;
  If St<>'' Then PC.DebitPos2:=(ReadTokenPipe(St,Pipe)='X') ;
  If St<>'' Then PC.Formule1:=Trim(ReadTokenPipe(St,Pipe)) ;
  If St<>'' Then PC.Formule2:=Trim(ReadTokenPipe(St,Pipe)) ;
  END ;
END ;

Procedure TFQRAFF.INITCHAMPTABLE(LibreOrder : String) ;
Var St, StCode, StOrder : String ;
    i : integer ;
BEGIN
St:=LibreOrder ; StCode:='' ; i:=0 ;
While St<>'' do
      BEGIN
      StCode:=ReadTokenSt(St) ;
      if (StCode='') then Continue ;
      if Stcode[1]='B' then CritEdtSup.CHAMPTABLE[i]:=Stcode[1]+'G_TABLE'+Stcode[3] Else
      if Stcode[1]='D' then CritEdtSup.CHAMPTABLE[i]:='BS_TABLE'+Stcode[3] Else
      CritEdtSup.CHAMPTABLE[i]:=Stcode[1]+'_TABLE'+Stcode[3] ;
      Inc(i) ;
      END ;
END ;

procedure TFQRAFF.RecupCritEdtSup ;
Var OkNatLibre : TOkNAtLibre ;
BEGIN
Fillchar(CritEdtSup,SizeOf(CritEdtSup),#0) ;
With CritEdtSup Do
  BEGIN
  RecupCritCol(CritEdtSup.PC[1],OkC1.Checked,SPC1.text) ; RecupCritCol(CritEdtSup.PC[2],OkC2.Checked,SPC2.text) ;
  RecupCritCol(CritEdtSup.PC[3],OkC3.Checked,SPC3.text) ; RecupCritCol(CritEdtSup.PC[4],OkC4.Checked,SPC4.text) ;
  END ;
(*
DecodeOkNatLibre(CritEdt.LibreCodes1,OkNatLibre) ;
For i:=0 To 9 Do If OkNatLibre[i] Then CritEdt.Y_TABLE[i]:=Pref+'TABLE'+IntToStr(i)+', ' ;
*)
InitChampTable(CritEdt.LibreTrie) ;
END ;

procedure TFQRAFF.RecupCritEdt ;
Var NonLibres : Boolean ;
    i : Integer ;
BEGIN
InHerited ;
With CritEdt Do
  BEGIN
  Bal.CodeRupt1:='' ; Bal.CodeRupt2:='' ; Bal.PlanRupt:='' ;
  NonLibres:=((Rupture=rRuptures) or (Rupture=rCorresp)) ;
  if NonLibres then BAL.PlanRupt:=FPlanRuptures.Value ;

  BAL.RuptOnly:=QuelleTypeRupt(0,FSAnsRupt.Checked,FAvecRupt.Checked,FSurRupt.Checked) ;
  if NonLibres then BEGIN Bal.CodeRupt1:=FCodeRupt1.Text ; Bal.CodeRupt2:=FCodeRupt2.Text ; END ;
  Bal.Axe:=FNatureCpt.Value ;
  Bal.QuelAN:=FQuelAN(FAvecAN.Checked) ;
  If Bal.QuelAN=SansAN Then FCol1.Checked:=FALSE ;
  If (CritEdt.Rupture=rCorresp) Then Bal.PlansCorresp:=FPlanRuptures.ItemIndex+1 ;
  Bal.OnlyCptAssocie:=(Rupture<>rRien) and FOnlyCptAssocie.Checked ;
  RecupCritEdtSup ;
  i:=StrToInt(Copy(NatLib.Value,2,2)) ;
  CritEdt.CptLibre1:='G_TABLE'+FormatFloat('0',i) ;
  If SautDePage.Checked Then CritEdt.SautPage:=2 ;
  With Bal.FormatPrint Do
       BEGIN
       PrSepRupt:=FLigneRupt.Checked ;
       PutTabCol(1,TabColl[1],TitreColCpt.Tag,TRUE) ;
       PutTabCol(2,TabColl[2],TitreColLibelle.Tag,TRUE) ;
       PutTabCol(3,TabColl[3],TTitreColAvant.Tag,FCol1.Checked) ;
       PutTabCol(4,TabColl[4],TTitreColSelection.Tag,FCol2.Checked) ;
       PutTabCol(5,TabColl[5],TTitreColApres.Tag,FCol3.Checked) ;
       PutTabCol(6,TabColl[6],TTitreColSolde.Tag,FCol4.Checked) ;
       END ;
  END ;
  LgSpecif1:=NbCharRupt.Value ;
END ;

Function TFQRAFF.CritOK : Boolean ;
BEGIN
Result:=Inherited CritOK and OkPourRupt and ColonnesOK ;
If Result Then
If Result Then
   BEGIN
//   QBal:=PrepareTotCpt(AxeToFb(CritEdt.Bal.Axe),QuelTypeEcr,Dev,Etab,Exo) ;
// Rony 17/03/97 QBal:=PrepareTotCptCum(QBalCum,AxeToFb(CritEdt.Bal.Axe),QuelTypeEcr,Dev,Etab,Exo,DevEnP) ;
   Gcalc:=TGCalculCum.create(Un,AxeToFb(CritEdt.Bal.Axe),AxeToFb(CritEdt.Bal.Axe),QuelTypeEcr,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE) ;
   GCalc.initCalcul('','','',CritEdt.Bal.Axe,CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                    CritEdt.Date1,CritEdt.Bal.Date11,TRUE) ;
   ListeCalc:=TStringList.Create ; ListeCalc.Sorted:=TRUE ; ListeCalc.Duplicates:=DupIgnore ;
   GenereSQLSup ;
   END ;
InitReport([1],CritEDT.BAL.FormatPrint.Report) ;
END ;

Function TFQRAFF.ColonnesOK : Boolean ;
BEGIN
If OkZoomEdt then begin Result:=True ; exit ; end ;
Result:=(FCol1.Checked or FCol2.Checked  or FCol3.Checked or FCol4.Checked) ;
If Not Result then NumErreurCrit:=9 ;
END ;

Procedure TFQRAFF.CalcTotalEdt(D1,C1,D2,C2,D3,C3,D4,C4 : Double );
BEGIN
  inherited;
TotEdt[1].TotDebit:= Arrondi(TotEdt[1].TotDebit  + D1, CritEdt.Decimale) ;
TotEdt[1].TotCredit:=Arrondi(TotEdt[1].TotCredit + C1, CritEdt.Decimale) ;
TotEdt[2].TotDebit:= Arrondi(TotEdt[2].TotDebit  + D2, CritEdt.Decimale) ;
TotEdt[2].TotCredit:=Arrondi(TotEdt[2].TotCredit + C2, CritEdt.Decimale) ;
TotEdt[3].TotDebit:= Arrondi(TotEdt[3].TotDebit  + D3, CritEdt.Decimale) ;
TotEdt[3].TotCredit:=Arrondi(TotEdt[3].TotCredit + C3, CritEdt.Decimale) ;
TotEdt[4].TotDebit:= Arrondi(TotEdt[4].TotDebit + D4, CritEdt.Decimale) ;
TotEdt[4].TotCredit:=Arrondi(TotEdt[4].TotCredit + C4, CritEdt.Decimale) ;
END ;

procedure TFQRAFF.FormShow(Sender: TObject);
begin
ChargeNatLib ;
FSansRupt.Enabled:=FALSE ; FAvecRupt.Checked:=TRUE ; FTablesLibres.Checked:=TRUE ;
HelpContext:=7451010 ;
Standards.HelpContext:=7451010 ;
Avances.HelpContext:=7451020 ;
Mise.HelpContext:=7451030 ;
Option.HelpContext:=7451040 ;
TabRuptures.HelpContext:=7451050 ;
  inherited;
Floading:=FALSE ;
//TabSup.TabVisible:=False;
TFGenJoker.Visible:=False ; TFGen.Visible:=True ;
FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
FSansRuptClick(Nil);
FCol1.Checked:=TRUE ; 
FOnlyCptAssocie.Enabled:=False ;
{ChargeFiltrePourRupture ;}
QRMethode:=tQRRequete ;
FReport.Checked:=FALSE ;
end;

procedure TFQRAFF.FormCreate(Sender: TObject);
begin
  inherited;
ListeCodesRupture:=TStringList.Create ;
end;

procedure TFQRAFF.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
ListeCodesRupture.Free ;
end;

procedure TFQRAFF.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
Report1DebitAnv.Caption:=Report2DebitAnv.Caption ;
Report1CreditAnv.Caption:=Report2CreditAnv.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
Report1DebitSum.Caption:=Report2DebitSum.Caption ;
Report1CreditSum.Caption:=Report2CreditSum.Caption ;
Report1DebitSol.Caption:=Report2DebitSol.Caption ;
Report1CreditSol.Caption:=Report2CreditSol.Caption ;
end;

Procedure CalcMontant(i : Integer ; Var PC : tParamCol ; Var TDC : TabTot ; Var Tot : TabDC ; EnEuro : Boolean) ;
Var XX : Double ;
    DebP : Boolean ;
BEGIN
If EnEuro Then XX:=TDC[1].TotDebit-TDC[1].TotCredit Else XX:=TDC[0].TotDebit-TDC[0].TotCredit ;
If i=1 Then DebP:=PC.DebitPos1 Else DebP:=PC.DebitPos2 ;
If Not DebP Then XX:=(-1)*XX ;
If i=1 Then Tot.TotDebit:=XX Else Tot.TotCredit:=XX ;
END ;

Procedure TFQRAFF.EstLigneSeule ;
Var CodRupt,LibRupt,St : String ;
    Quellerupt : Integer ;
    Tot    : Array[0..12] of Double ;
    nb : Double ;
BEGIN
//Exit ;
S_SECTION.Font.Style:=S_SECTION.Font.Style-[fsItalic] ;
If SimulePrintGroup(LMulti, Q, [Qr11S_SECTION], CodRupt, LibRupt, Tot,Quellerupt) Then
  BEGIN
  Nb:=Tot[0] ;
  If Nb<2 Then S_SECTION.Font.Style:=S_SECTION.Font.Style+[fsItalic] ;
  END ;
S_LIBELLE.Font.Style:=S_SECTION.Font.Style ;
ANvDEBIT.Font.Style:=S_SECTION.Font.Style ;
ANvCREDIT.Font.Style:=S_SECTION.Font.Style ;
DEBIT.Font.Style:=S_SECTION.Font.Style ;
CREDIT.Font.Style:=S_SECTION.Font.Style ;
DEBITsum.Font.Style:=S_SECTION.Font.Style ;
CREDITsum.Font.Style:=S_SECTION.Font.Style ;
SOLDEdeb.Font.Style:=S_SECTION.Font.Style ;
SOLDEcre.Font.Style:=S_SECTION.Font.Style ;
END ;

procedure TFQRAFF.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var TotCpt : TabDC4 ;
    Tot : TabTot ;
    TDC : TabTot ;
    MReport : TabTRep ;
    CptRupt     : String ;
    i : Integer ;
    SautLigne : Boolean ;
    LeTop,LeHeight : Integer ;
begin
  inherited;
PrintBand:=True ; AfterPrintDetail:=FALSE ; SautLigne:=FALSE ;
Fillchar(Tot,SizeOf(Tot),#0) ; Fillchar(TotCpt,SizeOf(TotCpt),#0) ;
Fillchar(MReport,SizeOf(MReport),#0) ;
Case CritEdt.Bal.TypCpt of  { 0 : Mvt sur L'exo,    3 : MVt sur la Période }
  0 :  PrintBand:=True ;
  3 :  PrintBand:= TRUE ;
 end ;
if PrintBand then
   BEGIN
   Case CritEdt.Rupture of
     rLibres    : if CritEdt.Bal.OnlyCptAssocie then PrintBand:=DansRuptLibre(Q,AxeToFb(CritEdt.Bal.Axe),CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
     rRuptures  : if CritEdt.Bal.OnlyCptAssocie then PrintBand:=(Qr11S_SECTIONTRIE.AsString<>'') ;
     rCorresp   : if CritEdt.Bal.OnlyCptAssocie then
                    if CritEDT.Bal.PlansCorresp=1 then PrintBand:=DansRuptCorresp(LRupt,Qr11S_CORRESP1.AsString) Else
                    if CritEDT.Bal.PlansCorresp=2 then PrintBand:=DansRuptCorresp(LRupt,Qr11S_CORRESP2.AsString) ;
//                    if CritEDT.Bal.PlansCorresp=1 then PrintBand:=(Qr11S_CORRESP1.AsString<>'') Else
//                    if CritEDT.Bal.PlansCorresp=2 then PrintBand:=(Qr11S_CORRESP2.AsString<>'') ;
     End;
   END ;
if PrintBand then
   BEGIN
   Quoi:=Qr11S_SECTION.AsString+' '+'#'+Qr11S_LIBELLE.AsString+'@'+'3'+Qr11S_AXE.AsString ;
   If QRMethode=tQRRequete Then
     BEGIN
     For i:=1 To 4 Do // Calcul Cumuls
       BEGIN
       If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value1<>'') And (CritEdtSup.PC[i].Formule1='') Then
         If ChercheMontantListeCalc(Q,CritEdtSup.PC[i].Value1,ListeCalc,CritEdt,CritEdtSup,TDC) Then CalcMontant(1,CritEdtSup.PC[i],TDC,TotCpt[i],CritEdt.Monnaie=2) ;
       If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value2<>'') And (CritEdtSup.PC[i].Formule2='') Then
         If ChercheMontantListeCalc(Q,CritEdtSup.PC[i].Value2,ListeCalc,CritEdt,CritEdtSup,TDC) Then CalcMontant(2,CritEdtSup.PC[i],TDC,TotCpt[i],CritEdt.Monnaie=2) ;
       END ;
     For i:=1 To 4 Do // Calcul Formules
       BEGIN
       If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value1='') And (CritEdtSup.PC[i].Formule1<>'') Then ChercheMontantFormule(1,i,CritEdtSup,TotCpt) ;
       If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value2='') And (CritEdtSup.PC[i].Formule2<>'') Then ChercheMontantFormule(2,i,CritEdtSup,TotCpt) ;
       END ;

     END ;

   Tot[1].TotDebit:= Arrondi(TotCpt[1].TotDebit,CritEdt.Decimale) ; Tot[1].TotCredit:=Arrondi(TotCpt[1].TotCredit,CritEdt.Decimale) ;
   Tot[2].TotDebit:= Arrondi(TotCpt[2].TotDebit,CritEdt.Decimale) ; Tot[2].TotCredit:=Arrondi(TotCpt[2].TotCredit,CritEdt.Decimale) ;
   Tot[3].TotDebit:= Arrondi(TotCpt[3].TotDebit,CritEdt.Decimale) ; Tot[3].TotCredit:=Arrondi(TotCpt[3].TotCredit,CritEdt.Decimale) ;
   Tot[4].TotDebit:= Arrondi(TotCpt[4].TotDebit,CritEdt.Decimale) ; Tot[4].TotCredit:=Arrondi(TotCpt[4].TotCredit,CritEdt.Decimale) ;
   AnvDebit.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[1].TotDebit,CritEdt.AfficheSymbole ) ;
   AnvCredit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[1].TotCredit,CritEdt.AfficheSymbole ) ;

   debit.caption:=    AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[2].TotDebit,CritEdt.AfficheSymbole ) ;
   credit.caption:=   AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[2].TotCredit,CritEdt.AfficheSymbole ) ;

   DEBITsum.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[3].TotDebit,CritEdt.AfficheSymbole ) ;
   CREDITsum.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[3].TotCredit,CritEdt.AfficheSymbole ) ;

   SoldeDeb.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[4].TotDebit,CritEdt.AfficheSymbole ) ;
   SoldeCre.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,Tot[4].TotCredit,CritEdt.AfficheSymbole ) ;

   AddGroup(LMulti, [Qr11S_SECTION], [1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,Tot[3].TotDebit,Tot[3].TotCredit,Tot[4].TotDebit,Tot[4].TotCredit]) ;

   Case CritEdt.Rupture of
     rLibres   : AddGroupLibre(LRupt,Q,AxeToFb(CritEdt.Bal.Axe),CritEdt.LibreTrie,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,Tot[3].TotDebit,Tot[3].TotCredit,Tot[4].TotDebit,Tot[4].TotCredit,0]) ;
     rRuptures : AddRupt(LRupt,Qr11S_SECTIONTRIE.AsString,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,Tot[3].TotDebit,Tot[3].TotCredit,Tot[4].TotDebit,Tot[4].TotCredit,0]) ;
     rCorresp  : BEGIN
                 Case CritEDT.Bal.PlansCorresp Of
                   1 : If Qr11S_CORRESP1.AsString<>'' Then CptRupt:=Qr11S_CORRESP1.AsString+Qr11S_SECTION.AsString
                                                      Else CptRupt:='.'+Qr11S_SECTION.AsString ;
                   2 : If Qr11S_CORRESP2.AsString<>'' Then CptRupt:=Qr11S_CORRESP2.AsString+Qr11S_SECTION.AsString
                                                      Else CptRupt:='.'+Qr11S_SECTION.AsString ;
                   Else CptRupt:=Qr11S_SECTION.AsString ;
                  End ;
                 AddRuptCorres(LRupt,CptRupt,[1,Tot[1].TotDebit,Tot[1].TotCredit,Tot[2].TotDebit,Tot[2].TotCredit,Tot[3].TotDebit,Tot[3].TotCredit,Tot[4].TotDebit,Tot[4].TotCredit,0]) ;
                 END ;
     End ;

   If (CritEdt.Bal.RuptOnly=Sur) then
      BEGIN
      PrintBand:=False ;
      IF CritEdt.Rupture=rLibres then CalcTotalEdt(Tot[1].TotDebit, Tot[1].TotCredit, Tot[2].TotDebit, Tot[2].TotCredit,Tot[3].TotDebit,Tot[3].TotCredit,Tot[4].TotDebit,Tot[4].TotCredit) ;
      END Else CalcTotalEdt(Tot[1].TotDebit, Tot[1].TotCredit, Tot[2].TotDebit, Tot[2].TotCredit,Tot[3].TotDebit,Tot[3].TotCredit,Tot[4].TotDebit,Tot[4].TotCredit) ;

   EstLigneSeule ;
   If Pos('*',Qr11S_LIBELLE.AsString)>0 Then SautLigne:=TRUE ;
   If SautLigne Then BEGIN LeTop:=20 ; LeHeight:=36 ; END Else BEGIN LeTop:=2 ; LeHeight:=18 ; END ;
   BDetail.Height:=LeHeight ;
   For i:=0 To BDetail.ControlCount-1 Do BDetail.Controls[i].Top:=LeTop ;
//   AddReportBAL([1],CritEdt.Bal.FormatPrint.Report,MReport,CritEdt.Decimale) ;
   END ;
end;

procedure TFQRAFF.BRUPTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.Bal.RuptOnly<>Sans) ;
end;

procedure TFQRAFF.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var i : Integer ;
    TotCpt : TabDC4 ;
begin
  inherited;
For i:=1 To 4 Do
  BEGIN
  TotCpt[i]:=TotEdt[i] ;
  If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value1='') And (CritEdtSup.PC[i].Formule1<>'') Then TotCpt[i].TotDebit:=0 ;
  If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value2='') And (CritEdtSup.PC[i].Formule2<>'') Then TotCpt[i].TotCredit:=0 ;
  END ;
For i:=1 To 4 Do // Calcul Formules
  BEGIN
  If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value1='') And (CritEdtSup.PC[i].Formule1<>'') Then
    BEGIN
    ChercheMontantFormule(1,i,CritEdtSup,TotCpt) ;
    END ;
  If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value2='') And (CritEdtSup.PC[i].Formule2<>'') Then
    BEGIN
    ChercheMontantFormule(2,i,CritEdtSup,TotCpt) ;
    END ;
  END ;


TOT2DEBITanv.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[1].TotDebit,CritEdt.AfficheSymbole ) ;
TOT2CREDITanv.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[1].TotCredit,CritEdt.AfficheSymbole ) ;

TOT2DEBIT.Caption:=    AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[2].TotDebit,CritEdt.AfficheSymbole ) ;
TOT2CREDIT.Caption:=   AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[2].TotCredit,CritEdt.AfficheSymbole ) ;

TOT2DEBITsum.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[3].TotDebit,CritEdt.AfficheSymbole ) ;
TOT2CREDITsum.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[3].TotCredit,CritEdt.AfficheSymbole ) ;

TOT2SOLdeb.Caption:=   AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[4].TotDebit,CritEdt.AfficheSymbole ) ;
TOT2SOLcre.Caption:=   AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[4].TotCredit,CritEdt.AfficheSymbole ) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFQRAFF.DLruptNeedData(var MoreData: Boolean);
Var CodRupt,LibRupt,Lib1  : String ;
    SumD, SumC, TotSoldeD, TotSoldeC : Double ;
    Tot    : Array[0..12] of Double ;
    CptRupt, StCode     : String ;
    Quellerupt : Integer ;
    OkOk,AddTotEdt : Boolean ;
    Col        : TColor ;
    LibRuptInf : Array[1..10] Of TRuptInf ;
    TotCpt : TabDC4 ;
    i : Integer ;
    OkAddRecap : Boolean ;
begin
  inherited;
MoreData:=False ;
Case CritEdt.Rupture of
  rLibres   : BEGIN
              MoreData:=PrintGroupLibre(LRupt,Q,AxeToFb(CritEdt.Bal.Axe),CritEdt.LibreTrie,CodRupt,LibRupt,Lib1,Tot,Quellerupt,Col,LibRuptInf) ;
              If MoreData then
                 BEGIN
                 StCode:=CodRupt ;
                 Delete(StCode,1,Quellerupt+2) ;
                 MoreData:=DansChoixCodeLibre(StCode,Q,AxeToFb(CritEdt.Bal.Axe),CritEdt.LibreCodes1,CritEdt.LibreCodes2, CritEdt.LibreTrie) ;
                 BRupt.Font.Color:=Col ;
                 END ;
              END ;
  rRuptures : MoreData:=PrintRupt(LRupt,Qr11S_SECTIONTRIE.AsString,CodRupt,LibRupt,DansTotal,QRP.EnRupture,Tot) ;
  rCorresp  : BEGIN
              OkOk:=TRUE ;
              Case CritEDT.Bal.PlansCorresp  Of
                1 : If Qr11S_CORRESP1.AsString<>'' Then CptRupt:=Qr11S_CORRESP1.AsString+Qr11S_SECTION.AsString
                                                   Else CptRupt:='.'+Qr11S_SECTION.AsString ;
                2 : If Qr11S_CORRESP2.AsString<>'' Then CptRupt:=Qr11S_CORRESP2.AsString+Qr11S_SECTION.AsString
                                                   Else CptRupt:='.'+Qr11S_SECTION.AsString ;
                Else OkOk:=FALSE ;
                End ;
              If OkOk Then MoreData:=PrintRupt(LRupt,CptRupt,CodRupt,LibRupt,DansTotal,QRP.EnRupture,Tot) Else MoreData:=FALSE ;
              END ;
  End ;

If MoreData then
   BEGIN
   TCodRupt.Width:=TitreColCpt.Width+TitreColLibelle.Width+1 ;
   TCodRupt.Caption:='' ;
   BRupt.Height:=HauteurBandeRuptIni ;
   OkAddRecap:=TRUE ;
   if CritEdt.Rupture=rLibres then
      BEGIN
      insert(MsgLibel.Mess[13]+' ',CodRupt,Quellerupt+2) ;
      TCodRupt.Caption:=CodRupt+' '+Lib1 ;
      AlimRuptSup(HauteurBandeRuptIni,QuelleRupt,TCodRupt.Width,BRupt,LibRuptInf,Self) ;
      OkAddRecap:=FALSE ; If QuelleRupt=0 Then OkAddRecap:=TRUE ;
      END Else TCodRupt.Caption:=CodRupt+'   '+LibRupt ;
   TCodRupt.Caption:=FindEtReplace(TCodRupt.Caption,'*',' ',TRUE) ;
   (*
   SumD:=TotRupt[1]+TotRupt[3] ;
   SumC:=TotRupt[2]+TotRupt[4] ;
   CalCulSolde(SumD,SumC,TotSoldeD,TotSoldeC) ;
   *)
   For i:=1 To 4 Do
     BEGIN
     TotCpt[i].TotDebit:=Tot[(2*i)-1] ; TotCpt[i].TotCredit:=Tot[(2*i)] ;
     If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value1='') And (CritEdtSup.PC[i].Formule1<>'') Then TotCpt[i].TotDebit:=0 ;
     If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value2='') And (CritEdtSup.PC[i].Formule2<>'') Then TotCpt[i].TotCredit:=0 ;
     END ;
   For i:=1 To 4 Do // Calcul Formules
     BEGIN
     If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value1='') And (CritEdtSup.PC[i].Formule1<>'') Then
       BEGIN
       ChercheMontantFormule(1,i,CritEdtSup,TotCpt) ; Tot[2*(i-1)+1]:=TotCpt[i].TotDebit ;
       END ;
     If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value2='') And (CritEdtSup.PC[i].Formule2<>'') Then
       BEGIN
       ChercheMontantFormule(2,i,CritEdtSup,TotCpt) ; Tot[2*i]:=TotCpt[i].TotCredit ;
       END ;
     END ;
   TotCpt[1].TotDebit:= Arrondi(TotCpt[1].TotDebit,CritEdt.Decimale) ; TotCpt[1].TotCredit:=Arrondi(TotCpt[1].TotCredit,CritEdt.Decimale) ;
   TotCpt[2].TotDebit:= Arrondi(TotCpt[2].TotDebit,CritEdt.Decimale) ; TotCpt[2].TotCredit:=Arrondi(TotCpt[2].TotCredit,CritEdt.Decimale) ;
   TotCpt[3].TotDebit:= Arrondi(TotCpt[3].TotDebit,CritEdt.Decimale) ; TotCpt[3].TotCredit:=Arrondi(TotCpt[3].TotCredit,CritEdt.Decimale) ;
   TotCpt[4].TotDebit:= Arrondi(TotCpt[4].TotDebit,CritEdt.Decimale) ; TotCpt[4].TotCredit:=Arrondi(TotCpt[4].TotCredit,CritEdt.Decimale) ;

   TotT1.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[1].TotDebit,CritEdt.AfficheSymbole ) ;
   TotT2.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[1].TotCredit,CritEdt.AfficheSymbole ) ;

   TotT3.caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[2].TotDebit,CritEdt.AfficheSymbole ) ;
   TOTT4.caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[2].TotCredit,CritEdt.AfficheSymbole ) ;

   TOTT5.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[3].TotDebit,CritEdt.AfficheSymbole ) ;
   TOTT6.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[3].TotCredit,CritEdt.AfficheSymbole ) ;

   TOTT7.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[4].TotDebit,CritEdt.AfficheSymbole ) ;
   TOTT8.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[4].TotCredit,CritEdt.AfficheSymbole ) ;
//   If OkAddRecap Then  AddRecap(LRecapG,[TCodRupt.Caption],[''],[TotCpt[1].TotDebit, TotCpt[1].TotCredit, TotCpt[2].TotDebit, TotCpt[2].TotCredit, TotCpt[3].TotDebit, TotCpt[3].TotCredit, TotCpt[4].TotDebit, TotCpt[4].TotCredit]) ;
   If OkAddRecap Then  AddRecap(LRecapG,[Lib1],[''],[TotCpt[1].TotDebit, TotCpt[1].TotCredit, TotCpt[2].TotDebit, TotCpt[2].TotCredit, TotCpt[3].TotDebit, TotCpt[3].TotCredit, TotCpt[4].TotDebit, TotCpt[4].TotCredit]) ;
   if CritEdt.Bal.RuptOnly=Sur then
      BEGIN
      AddTotEdt:=False ;
      if (CritEdt.Rupture=rLibres) then AddTotEdt:=False else
      if (CritEdt.Rupture=rRuptures) or (CritEdt.Rupture=rCorresp) then AddTotEdt:=DansTotal ;
      if AddTotEdt then CalcTotalEdt(TotCpt[1].TotDebit, TotCpt[1].TotCredit, TotCpt[2].TotDebit, TotCpt[2].TotCredit, TotCpt[3].TotDebit, TotCpt[3].TotCredit, TotCpt[4].TotDebit, TotCpt[4].TotCredit) ;
      END ;
   If CritEdt.SautPage=2 Then BDetail.ForceNewPage:=TRUE ;
   END ;
end;

procedure TFQRAFF.FNatureCptChange(Sender: TObject);
begin
FCpte1.Clear ; FCpte2.Clear ;
  inherited;
if FPlansCo.checked then
   BEGIN CorrespToCombo(FPlanRuptures,AxeToFb(FNatureCpt.Value)) ; Exit ; END ;
Case FNatureCpt.ItemIndex of
   0 : FPlanRuptures.Datatype:='ttRuptSect1' ;
   1 : FPlanRuptures.Datatype:='ttRuptSect2' ;
   2 : FPlanRuptures.Datatype:='ttRuptSect3' ;
   3 : FPlanRuptures.Datatype:='ttRuptSect4' ;
   4 : FPlanRuptures.Datatype:='ttRuptSect5' ;
  end ;
FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
end;

procedure TFQRAFF.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var MReport : TabTRep ;
begin
  inherited;
Case QuelReportBAL(CritEdt.Bal.FormatPrint.Report,MReport) Of
  1 : TITREREPORTB.Caption:=MsgLibel.Mess[7] ;
  END ;
Report2DebitAnv.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[1].TotDebit, CritEdt.AfficheSymbole ) ;
Report2CreditAnv.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[1].TotCredit, CritEdt.AfficheSymbole ) ;
Report2Debit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[2].TotDebit, CritEdt.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[2].TotCredit, CritEdt.AfficheSymbole ) ;
Report2DebitSum.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[3].TotDebit, CritEdt.AfficheSymbole ) ;
Report2CreditSum.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[3].TotCredit, CritEdt.AfficheSymbole ) ;
Report2DebitSol.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[4].TotDebit, CritEdt.AfficheSymbole ) ;
Report2CreditSol.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[4].TotCredit, CritEdt.AfficheSymbole ) ;
end;

procedure TFQRAFF.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr11S_AXE         :=TStringField(Q.FindField('S_AXE'));
Qr11S_SECTION     :=TStringField(Q.FindField('S_SECTION'));
Qr11S_SECTION.Tag :=100 ; // CERIC
Qr11S_LIBELLE     :=TStringField(Q.FindField('S_LIBELLE'));
Qr11S_SECTIONTRIE :=TStringField(Q.FindField('S_SECTIONTRIE'));
if CritEdt.Rupture=rCorresp then
   BEGIN
   Qr11S_CORRESP1         :=TStringField(Q.FindField('S_CORRESP1'));
   Qr11S_CORRESP2         :=TStringField(Q.FindField('S_CORRESP2'));
   END ;
end;

Function TFQRAFF.OkPourRupt : Boolean ;
Var i : Integer ;
BEGIN
Result:=True ;
if (CritEdt.Rupture=rRuptures) then
   BEGIN
   Result:=False ; DLrupt.PrintBefore:=TRUE ;
   if ListeCodesRupture<>Nil then ListeCodesRupture.Clear ;
   For i:=0 to FCodeRupt1.Items.Count-1 do ListeCodesRupture.Add(FCodeRupt1.Items[i]) ;
   ChargeRupt(LRupt,'RU'+Char(CritEdt.Bal.Axe[2]),CritEdt.Bal.PlanRupt,CritEdt.Bal.CodeRupt1 ,CritEdt.Bal.CodeRupt2) ;
   NiveauRupt(LRupt) ;
   Case SectionRetrie(CritEdt.Bal.PlanRupt,CritEdt.Bal.Axe,ListeCodesRupture) of
     srOk              : Result:=True ;
     srNonStruct       : NumErreurCrit:=7 ; //MsgRien.Execute(7,'','') ;
     srPasEnchainement : NumErreurCrit:=8 ; // MsgRien.Execute(8,'','') ;
     End ;

   END ;
END ;

procedure TFQRAFF.FExerciceChange(Sender: TObject);
begin
FLoading:=TRUE ;
  inherited;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
FLoading:=FALSE ;
end;

procedure TFQRAFF.FDateCompta1Change(Sender: TObject);
begin
  inherited;
If FLoading Then Exit ;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
end;

procedure TFQRAFF.FSelectCpteChange(Sender: TObject);
begin
  inherited;
VoirQuelAN(FSelectCpte.Value,FExercice.Value,FDateCompta1,FAvecAN) ;
end;

procedure TFQRAFF.FAvecANClick(Sender: TObject);
begin
  inherited;
FCol1.Checked:=FAvecAN.Checked ;
end;

procedure TFQRAFF.FRupturesClick(Sender: TObject);
begin
  inherited;
If FPlansCo.Checked then FGroupRuptures.Caption:=' '+MsgLibel.Mess[11] ;
If FRuptures.Checked then FGroupRuptures.Caption:=' '+MsgLibel.Mess[10] ;
end;

procedure TFQRAFF.FSansRuptClick(Sender: TObject);
begin
  inherited;
FLigneRupt.Enabled:=Not FSansRupt.Checked ;
FLigneRupt.checked:=Not FSansRupt.Checked ;
(*
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
*)
FOnlyCptAssocie.Enabled:=FALSE ;
FOnlyCptAssocie.Checked:=TRUE ;
FRupturesClick(Nil) ;
end;

procedure TFQRAFF.ChargeNatLib ;
Var QLoc : TQuery ;
begin
QLoc:=OpenSql('Select CO_CODE,CO_LIBELLE From COMMUN Where CO_TYPE="NAT" And CO_CODE Like"G%" ORDER BY CO_CODE',True) ;
NatLib.Values.Clear ; NatLib.Items.Clear ;
While Not QLoc.Eof do
   BEGIN
   NatLib.Values.Add(QLoc.Fields[0].AsString) ; NatLib.Items.Add(QLoc.Fields[1].AsString) ;
   QLoc.Next ;
   END ;
Ferme(QLoc) ;
if NatLib.Values.Count>0 then NatLib.Value:=NatLib.Values[0] ;
end;

procedure TFQRAFF.OKC1Click(Sender: TObject);
begin
  inherited;
SPC1.Visible:=OkC1.Checked ; If Not OKC1.Checked Then SPC1.Text:='' ;
FCol1.Checked:=OKC1.Checked ; FCol1.Enabled:=OKC1.Checked ;
If OkC1.Checked Then FCol1.Caption:=MsgLibel.Mess[15]+' 1' ;
end;

procedure TFQRAFF.OkC2Click(Sender: TObject);
begin
  inherited;
SPC2.Visible:=OkC2.Checked ; If Not OKC2.Checked Then SPC2.Text:='' ;
FCol2.Checked:=OKC2.Checked ; FCol2.Enabled:=OKC2.Checked ;
If OkC2.Checked Then FCol2.Caption:=MsgLibel.Mess[15]+' 2' ;
end;

procedure TFQRAFF.OkC3Click(Sender: TObject);
begin
  inherited;
SPC3.Visible:=OkC3.Checked ; If Not OKC3.Checked Then SPC3.Text:='' ;
FCol3.Checked:=OKC3.Checked ; FCol3.Enabled:=OKC3.Checked ;
If OkC3.Checked Then FCol3.Caption:=MsgLibel.Mess[15]+' 3' ;
end;

procedure TFQRAFF.OkC4Click(Sender: TObject);
begin
  inherited;
SPC4.Visible:=OkC4.Checked ; If Not OKC4.Checked Then SPC4.Text:='' ;
FCol4.Checked:=OKC4.Checked ; FCol4.Enabled:=OKC4.Checked ;
If OkC4.Checked Then FCol4.Caption:=MsgLibel.Mess[15]+' 4' ;
end;

procedure TFQRAFF.SPC1DblClick(Sender: TObject);
Var St : String ;
    i : Integer ;
begin
  inherited;
If NatLib.ItemIndex<0 Then Exit ;
If TEdit(Sender).Name='SPC1' Then i:=1 Else If TEdit(Sender).Name='SPC2' Then i:=2 Else
  If TEdit(Sender).Name='SPC3' Then i:=3 Else If TEdit(Sender).Name='SPC4' Then i:=4 Else Exit ;
St:=TEdit(Sender).Text ;
If ParamCol(St,i,NatLib.Values[NatLib.ItemIndex],NatLib.Items[NatLib.ItemIndex]) Then TEdit(Sender).Text:=St ;
end;

Procedure TFQRAFF.ChangeTitreCB(FCol : TCheckBox ; SPC : TEdit ; OkC : Boolean ; i : Integer) ;
Var St,St1 : String ;
BEGIN
FCol.Checked:=OkC ; FCol.Enabled:=OkC ;
If OkC Then
  BEGIN
  St:=SPC.Text ; If St<>'' Then St1:=' "'+ReadTokenPipe(St,Pipe)+'"' Else St1:=' '+IntToStr(i) ;
  FCol.Caption:=MsgLibel.Mess[15]+St1 ;
  END Else
  BEGIN
  FCol.Caption:=MsgLibel.Mess[15]+' '+IntToStr(i) ;
  END ;
END ;

procedure TFQRAFF.SPC1Change(Sender: TObject);
Var OkC : Boolean ;
    FCol : TCheckBox ;
    SPC : TEdit ;
    i : Integer ;
begin
  inherited;
If TEdit(Sender).Name='SPC1' Then BEGIN OkC:=OkC1.Checked ; FCol:=FCol1 ; i:=1 ; SPC:=SPC1 ; END Else
  If TEdit(Sender).Name='SPC2' Then BEGIN OkC:=OkC2.Checked ; FCol:=FCol2 ; i:=2 ; SPC:=SPC2 ;  END Else
    If TEdit(Sender).Name='SPC3' Then BEGIN OkC:=OkC3.Checked ; FCol:=FCol3 ; i:=3 ; SPC:=SPC3 ; END Else
      If TEdit(Sender).Name='SPC4' Then BEGIN OkC:=OkC4.Checked ; FCol:=FCol4 ; i:=4 ; SPC:=SPC4 ; END Else Exit ;
ChangeTitreCB(FCol,SPC,OkC,i) ;
end;

procedure TFQRAFF.FFiltresChange(Sender: TObject);
begin
  inherited;
OkC1Click(Nil) ; OkC2Click(Nil) ; OkC3Click(Nil) ; OkC4Click(Nil) ;
ChangeTitreCB(FCol1,SPC1,OkC1.Checked,1) ; ChangeTitreCB(FCol2,SPC2,OkC2.Checked,2) ;
ChangeTitreCB(FCol3,SPC3,OkC3.Checked,3) ; ChangeTitreCB(FCol4,SPC4,OkC4.Checked,4) ;
end;

procedure TFQRAFF.QRDLMultiNeedData(var MoreData: Boolean);
Var CodRupt,LibRupt,St : String ;
    Quellerupt : Integer ;
    Tot    : Array[0..12] of Double ;
    TotCpt : TabDC4 ;
    i : Integer ;
    Q1 : TQuery ;
    Nb : Double ;
begin
  inherited;
//If BDetail.forceNewPage Then BDetail.forceNewPage:=FALSE ;
MoreData:=PrintGroup(LMulti, Q, [Qr11S_SECTION], CodRupt, LibRupt, Tot,Quellerupt) ;
If MoreData Then
  BEGIN
  OkImprimeSt:=TRUE ;
  Nb:=Tot[0] ;
  If Nb<2 Then BEGIN OkImprimeSt:=FALSE ; END ;
  For i:=1 To 4 Do
    BEGIN
    TotCpt[i].TotDebit:=Tot[(2*i)-1] ; TotCpt[i].TotCredit:=Tot[(2*i)] ;
    If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value1='') And (CritEdtSup.PC[i].Formule1<>'') Then TotCpt[i].TotDebit:=0 ;
    If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value2='') And (CritEdtSup.PC[i].Formule2<>'') Then TotCpt[i].TotCredit:=0 ;
    END ;
  For i:=1 To 4 Do // Calcul Formules
    BEGIN
    If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value1='') And (CritEdtSup.PC[i].Formule1<>'') Then
      BEGIN
      ChercheMontantFormule(1,i,CritEdtSup,TotCpt) ; Tot[2*(i-1)+1]:=TotCpt[i].TotDebit ;
      END ;
    If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value2='') And (CritEdtSup.PC[i].Formule2<>'') Then
      BEGIN
      ChercheMontantFormule(2,i,CritEdtSup,TotCpt) ; Tot[2*i]:=TotCpt[i].TotCredit ;
      END ;
    END ;
  TotCpt[1].TotDebit:= Arrondi(TotCpt[1].TotDebit,CritEdt.Decimale) ; TotCpt[1].TotCredit:=Arrondi(TotCpt[1].TotCredit,CritEdt.Decimale) ;
  TotCpt[2].TotDebit:= Arrondi(TotCpt[2].TotDebit,CritEdt.Decimale) ; TotCpt[2].TotCredit:=Arrondi(TotCpt[2].TotCredit,CritEdt.Decimale) ;
  TotCpt[3].TotDebit:= Arrondi(TotCpt[3].TotDebit,CritEdt.Decimale) ; TotCpt[3].TotCredit:=Arrondi(TotCpt[3].TotCredit,CritEdt.Decimale) ;
  TotCpt[4].TotDebit:= Arrondi(TotCpt[4].TotDebit,CritEdt.Decimale) ; TotCpt[4].TotCredit:=Arrondi(TotCpt[4].TotCredit,CritEdt.Decimale) ;

  TotST1.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[1].TotDebit,CritEdt.AfficheSymbole ) ;
  TotST2.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[1].TotCredit,CritEdt.AfficheSymbole ) ;

  TotST3.caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[2].TotDebit,CritEdt.AfficheSymbole ) ;
  TOTST4.caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[2].TotCredit,CritEdt.AfficheSymbole ) ;

  TOTST5.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[3].TotDebit,CritEdt.AfficheSymbole ) ;
  TOTST6.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[3].TotCredit,CritEdt.AfficheSymbole ) ;

  TOTST7.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[4].TotDebit,CritEdt.AfficheSymbole ) ;
  TOTST8.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[4].TotCredit,CritEdt.AfficheSymbole ) ;
  St:=CodRupt ;
  Q1:=OpenSQL('SELECT S_SECTION,S_LIBELLE FROM SECTION WHERE S_SECTION LIKE "'+CodRupt+'%" AND S_AXE="'+CritEdt.Bal.Axe+'" ORDER BY S_AXE,S_SECTION',TRUE) ;
  If Not Q1.Eof Then St:=Q1.Fields[1].AsString ;
  Ferme(Q1) ;
  TCodeST.Width:=TitreColCpt.Width+TitreColLibelle.Width+1 ;
  TCodeST.Caption:='     '+' S/T '+St ;
  END ;
end;

procedure TFQRAFF.BRecapGBeforePrint(var PrintBand: Boolean;
  var Quoi: String);
begin
  inherited;
//PrintBand:=FALSE ;
end;

procedure TFQRAFF.QRDLRecapGNeedData(var MoreData: Boolean);
Var Cod,Lib : String ;
    TotRecapG : Array[0..12] Of Double ;
    TotCpt : TabDC4 ;
    i : Integer ;
begin
  inherited;
MoreData:=PrintRecap(LRecapG, Cod, Lib, TotRecapG) ;
If MoreData Then
   BEGIN
   For i:=1 To 4 Do
     BEGIN
     TotCpt[i].TotDebit:=TotRecapG[2*(i-1)] ; TotCpt[i].TotCredit:=TotRecapG[(2*i)-1] ;
     If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value1='') And (CritEdtSup.PC[i].Formule1<>'') Then TotCpt[i].TotDebit:=0 ;
     If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value2='') And (CritEdtSup.PC[i].Formule2<>'') Then TotCpt[i].TotCredit:=0 ;
     END ;
   For i:=1 To 4 Do // Calcul Formules
     BEGIN
     If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value1='') And (CritEdtSup.PC[i].Formule1<>'') Then
       BEGIN
       ChercheMontantFormule(1,i,CritEdtSup,TotCpt) ; TotRecapG[2*(i-1)+1]:=TotCpt[i].TotDebit ;
       END ;
     If (CritEdtSup.PC[i].OkOk) And (CritEdtSup.PC[i].Value2='') And (CritEdtSup.PC[i].Formule2<>'') Then
       BEGIN
       ChercheMontantFormule(2,i,CritEdtSup,TotCpt) ; TotRecapG[2*i]:=TotCpt[i].TotCredit ;
       END ;
     END ;
   TotCpt[1].TotDebit:= Arrondi(TotCpt[1].TotDebit,CritEdt.Decimale) ; TotCpt[1].TotCredit:=Arrondi(TotCpt[1].TotCredit,CritEdt.Decimale) ;
   TotCpt[2].TotDebit:= Arrondi(TotCpt[2].TotDebit,CritEdt.Decimale) ; TotCpt[2].TotCredit:=Arrondi(TotCpt[2].TotCredit,CritEdt.Decimale) ;
   TotCpt[3].TotDebit:= Arrondi(TotCpt[3].TotDebit,CritEdt.Decimale) ; TotCpt[3].TotCredit:=Arrondi(TotCpt[3].TotCredit,CritEdt.Decimale) ;
   TotCpt[4].TotDebit:= Arrondi(TotCpt[4].TotDebit,CritEdt.Decimale) ; TotCpt[4].TotCredit:=Arrondi(TotCpt[4].TotCredit,CritEdt.Decimale) ;

   TotR1.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[1].TotDebit,CritEdt.AfficheSymbole ) ;
   TotR2.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[1].TotCredit,CritEdt.AfficheSymbole ) ;

   TotR3.caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[2].TotDebit,CritEdt.AfficheSymbole ) ;
   TOTR4.caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[2].TotCredit,CritEdt.AfficheSymbole ) ;

   TOTR5.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[3].TotDebit,CritEdt.AfficheSymbole ) ;
   TOTR6.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[3].TotCredit,CritEdt.AfficheSymbole ) ;

   TOTR7.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[4].TotDebit,CritEdt.AfficheSymbole ) ;
   TOTR8.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,TotCpt[4].TotCredit,CritEdt.AfficheSymbole ) ;
   TCodeRecap.Caption:=Cod+' '+Lib ;
   END ;

end;

procedure TFQRAFF.BRecapGAfterPrint(BandPrinted: Boolean);
begin
  inherited;
BRecapG.ForceNewPage:=FALSE ;
end;

procedure TFQRAFF.BDSMultiBeforePrint(var PrintBand: Boolean;
  var Quoi: String);
begin
  inherited;
If Not OkImprimeSt then
  BEGIN
  TotST1.Caption:='' ; TotST2.Caption:='' ;
  TotST3.Caption:='' ; TotST4.Caption:='' ;
  TotST4.Caption:='' ; TotST6.Caption:='' ;
  TotST7.Caption:='' ; TotST8.Caption:='' ;
  BDSMulti.Height:=0 ;
  TCodeST.Caption:='' ;
  END Else BDSMulti.Height:=36 ;
If (CritEdt.Bal.RuptOnly=Sur) then
  BEGIN
  BDSMulti.Height:=0 ;
  TotST1.Caption:='' ; TotST2.Caption:='' ;
  TotST3.Caption:='' ; TotST4.Caption:='' ;
  TotST4.Caption:='' ; TotST6.Caption:='' ;
  TotST7.Caption:='' ; TotST8.Caption:='' ;
  BDSMulti.Height:=0 ;
  TCodeST.Caption:='' ;
  END ;
end;

procedure TFQRAFF.BDetailAfterPrint(BandPrinted: Boolean);
begin
  inherited;
If AfterPrintDetail Then ;
If BDetail.forceNewPage Then BDetail.forceNewPage:=FALSE ;
AfterPrintDetail:=TRUE ;
end;

procedure TFQRAFF.BDetailCheckForSpace;
Var OldForce : Boolean ;
begin
OldForce:=BDetail.ForceNewPage ;
  inherited;
If OldForce Then BDetail.ForceNewPage:=TRUE ;
end;

end.
