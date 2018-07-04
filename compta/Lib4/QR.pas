unit QR;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HQuickRP, StdCtrls, ExtCtrls, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Buttons, Mask, Hctrls,
  ULibExercice,  // ajout me
  Cpteutil,
  ComCtrls, HQry, HMsgBox, Ent1, HCompte, HEnt1, Filtre, UtilEdt, SaisUTil,
  ParamDat, Grids, Printers, CritEdt, Menus, HSysMenu,SaisComm,
{$IFNDEF IMP}
{$IFNDEF CCMP}
{$IFNDEF CCS3}
  SaisBud,
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$IFDEF VER150}
  variants,
{$ENDIF}
  utilPGI,     // TSQLAnaCroise
  UListByUser, UTob, UTXml,{FQ 14242}
  HTB97, HPanel, UiUtil, ParamSoc  ,UentCommun, ADODB, TntStdCtrls,
  TntExtCtrls;



Const OnImpr : Boolean = FALSE;
      OkV2 : Boolean = TRUE;
type
  TFQR = class(TForm)
    Pages: TPageControl;
    Standards: TTabSheet;
    TFaG: TLabel;
    Label7: TLabel;
    FNatureEcr: THValComboBox;
    FDateCompta2: TMaskEdit;
    FDateCompta1: TMaskEdit;
    FExercice: THValComboBox;
    FNatureCpt: THValComboBox;
    FCpte1: THCpteEdit;
    FCpte2: THCpteEdit;
    Avances: TTabSheet;
    FDevises: THValComboBox;
    FEtab: THValComboBox;
    Option: TTabSheet;
    Apercu: TCheckBox;
    Panel1: TPanel;
    BTitre: TQRBand;
    TTitre: TQRSysData;
    EnteteAutrePage: TQRBand;
    TitreEntete: TQRSysData;
    TitreBarre: TQRShape;
    BDetail: TQRBand;
    TFGen: THLabel;
    Hlabel2: THLabel;
    HLabel4: THLabel;
    HLabel5: THLabel;
    HLabel6: THLabel;
    HLabel7: THLabel;
    HLabel8: THLabel;
    Formateur: THNumEdit;
    BFinEtat: TQRBand;
    Q: TQuery;
    S: TDataSource;
    QRP: TQuickReport;
    MsgRien: THMsgBox;
    TFGenJoker: THLabel;
    FJoker: TEdit;
    HLabel1: THLabel;
    FExcep: TEdit;
    EntetePage: TQRBand;
    FListe: TCheckBox;
    RCpte: TQRLabel;
    TRa: TQRLabel;
    RCpte1: TQRLabel;
    RCpte2: TQRLabel;
    RExcepGen: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel15: TQRLabel;
    RDateCompta1: TQRLabel;
    RDateCompta2: TQRLabel;
    RExercice: TQRLabel;
    RNatureEcr: TQRLabel;
    REtab: TQRLabel;
    QRLabel18: TQRLabel;
    RDevises: TQRLabel;
    QRLabel12: TQRLabel;
    RNatureCpt: TQRLabel;
    GroupBox3: TGroupBox;
    Bevel1: TBevel;
    TOPREPORT: TQRBand;
    BOTTOMREPORT: TQRBand;
    piedpage: TQRBand;
    RCopyright: TQRLabel;
    QRSysData1: TQRSysData;
    QRSysData2: TQRSysData;
    RSociete: TQRLabel;
    RUtilisateur: TQRLabel;
    QRLabel24: TQRLabel;
    TSelectCpte: THLabel;
    FSelectCpte: THValComboBox;
    RSelectCpte: TQRLabel;
    TRSelectCpte: TQRLabel;
    AvecRevision: TCheckBox;
    OnCum: TCheckBox;
    FCouleur: TCheckBox;
    BtitreZoom: TQRBand;
    QRSysData3: TQRSysData;
    QRLabel8Zoom: TQRLabel;
    RCptZoom: TQRLabel;
    QRLabel16Zoom: TQRLabel;
    QRLabel21Zoom: TQRLabel;
    QRLabel25Zoom: TQRLabel;
    QRLabel26Zoom: TQRLabel;
    RDateCompta1Zoom: TQRLabel;
    RDateCompta2Zoom: TQRLabel;
    RNatureEcrZoom: TQRLabel;
    REtabZoom: TQRLabel;
    QRLabel38Zoom: TQRLabel;
    RDevisesZoom: TQRLabel;
    HMCrit: THMsgBox;
    tabSup: TTabSheet;
    Mise: TTabSheet;
    FLigneCpt: TCheckBox;
    Reduire: TCheckBox;
    FTrait: TCheckBox;
    GroupBox1: TGroupBox;
    Bevel2: TBevel;
    FMonetaire: TCheckBox;
    FReport: TCheckBox;
    Avance: TCheckBox;
    FMontant: TRadioGroup;
    TITREREPORTB: TQRLabel;
    TITREREPORTH: TQRLabel;
    MsgQR: THMsgBox;
    RJoker: TQRLabel;
    RNumversion: TQRLabel;
    FRappelCrit: TCheckBox;
    BtitreCrit: TQRBand;
    QRSysData4: TQRSysData;
    RappelCrit: TQRLabel;
    CadreCrit: TQRShape;
    TabRuptures: TTabSheet;
    FGroupLibres: TGroupBox;
    TFLibre1: THLabel;
    FLibre1: TEdit;
    TFLibre2: THLabel;
    FLibre2: TEdit;
    FLibTriPar: TEdit;
    TFLibTriPar: THLabel;
    FGroupChoixRupt: TGroupBox;
    FSansRupt: TRadioButton;
    FAvecRupt: TRadioButton;
    FSurRupt: TRadioButton;
    FGroupQuelleRupture: TGroupBox;
    FRuptures: TRadioButton;
    FTablesLibres: TRadioButton;
    FPlansCo: TRadioButton;
    FGroupRuptures: TGroupBox;
    TFPlanRuptures: TLabel;
    TFCodeRupt1: TLabel;
    TFCodeRupt2: TLabel;
    FPlanRuptures: THValComboBox;
    FCodeRupt1: TComboBox;
    FCodeRupt2: TComboBox;
    TabComplement: TTabSheet;
    TFNat0: THLabel;
    FNat01: THCpteEdit;
    FNat02: THCpteEdit;
    TFNat1: THLabel;
    FNat11: THCpteEdit;
    FNat12: THCpteEdit;
    TFNat2: THLabel;
    FNat21: THCpteEdit;
    FNat22: THCpteEdit;
    TFNat3: THLabel;
    FNat31: THCpteEdit;
    FNat32: THCpteEdit;
    FOnlyCptAssocie: TCheckBox;
    HMTrad: THSystemMenu;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    BDupFiltre: TMenuItem;
    Dock971: TDock97;
    HPB: TToolWindow97;
    BParamListe: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    PanelFiltre: TToolWindow97;
    BFiltre: TToolbarButton97;
    FFiltres: THValComboBox;
    FSautPageRupt: TCheckBox;
    Timer1: TTimer;
    PAvances: TTabSheet;
    bEffaceAvance: TToolbarButton97;
    Z_C3: THValComboBox;
    Z_C2: THValComboBox;
    Z_C1: THValComboBox;
    ZO1: THValComboBox;
    ZO2: THValComboBox;
    ZO3: THValComboBox;
    ZV3: TEdit;
    ZV2: TEdit;
    ZV1: TEdit;
    ZG1: TComboBox;
    ZG2: TComboBox;
    Z_C4: THValComboBox;
    Z_C5: THValComboBox;
    Z_C6: THValComboBox;
    ZO6: THValComboBox;
    ZO5: THValComboBox;
    ZO4: THValComboBox;
    ZV4: TEdit;
    ZV5: TEdit;
    ZV6: TEdit;
    ZG4: TComboBox;
    ZG5: TComboBox;
    ZG3: TComboBox;
    CBLib: TCheckBox;
    CTri1: TCheckBox;
    CTri2: TCheckBox;
    CTri3: TCheckBox;
    CTri4: TCheckBox;
    CTri5: TCheckBox;
    CTri6: TCheckBox;
    HTri: TLabel;
    FSautPageTRI: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure ParamLaDate(Sender: TObject; var Key: Char);
    procedure BSaveFiltreClick(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
    procedure FExerciceChange(Sender: TObject);
    procedure FMontantClick(Sender: TObject);
    procedure FDevisesChange(Sender: TObject);
    procedure piedpageAfterPrint(BandPrinted: Boolean);
    procedure BTitreAfterPrint(BandPrinted: Boolean);
    procedure FCpte1Change(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FNatureCptChange(Sender: TObject);
    procedure EnteteAutrePageAfterPrint(BandPrinted: Boolean);
    procedure RemplirCritEdtList ; virtual ;
    function  SQLModeSelect : String ;
    function  SQLModeSelectMul(Where1,Where2,PreBTot : String) : String ;
    function  QuelTypeEcr : SetttTypePiece ;
    function  AfficheMontant( Formatage, LeSymbole : String ; LeMontant : Double ; OkSymbole : Boolean ) : String ;
    Procedure Progressif(CalculMoi : Boolean ; D,C : Double);
    procedure CalculSolde(MontantD,MontantC : Double ; var SoldeD,SoldeC : Double) ; //GC - 20/12/2001
    procedure TOPREPORTAfterPrint(BandPrinted: Boolean);
    procedure BParamListeClick(Sender: TObject);
    procedure AvanceClick(Sender: TObject);
    Function  AlimSQLMul(Q : TQuery ; Niveau : integer) : Boolean ;
    Function  AlimSQLMulNat(Q : TQuery ; Niveau : integer ; tz1,tz2 : TZoomTable) : Boolean ;
    procedure BCreerFiltreClick(Sender: TObject);
    procedure BDelFiltreClick(Sender: TObject);
    procedure BRenFiltreClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    function  FquelAN(Inclure : Boolean) : TQuelAN ;
    procedure BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BOTTOMREPORTAfterPrint(BandPrinted: Boolean);
    procedure EnteteAutrePageBeforePrint(var PrintBand: Boolean; var Quoi: string);
    Procedure GenereSQLBase  ;
    Procedure GenereSQLBaseTS ;
    Function  SQLQuelCpt : String ;
    procedure BtitreCritAfterPrint(BandPrinted: Boolean);
    procedure BtitreCritBeforePrint(var PrintBand: Boolean; var Quoi: string);
    Function  AvecReport : Boolean ;
    (*
    procedure TrouveOverlay(Ena : Boolean) ;
    *)
    procedure FormActivate(Sender: TObject);
    procedure FLibre1DblClick(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure FSansRuptClick(Sender: TObject);
    procedure ApercuClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure RajouteLabelRupture(NomBandRupt : String);
    Function  OkRajouteLabelRupture : Boolean ;
    procedure BDetailCheckForSpace;
    procedure BAideClick(Sender: TObject);
    procedure FJokerKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure POPFPopup(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    Procedure InitVarInterne ;
    procedure FPlanRupturesChange(Sender: TObject);
    Function  QuelfbPourRupture : TFichierBase ;
    procedure FCpte1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FCpte1KeyPress(Sender: TObject; var Key: Char);
    procedure Timer1Timer(Sender: TObject);
    procedure OnCumClick(Sender: TObject);
    procedure CBLibClick(Sender: TObject);
    procedure CTri1Click(Sender: TObject);
    procedure Z_C1Change(Sender: TObject);
    procedure bEffaceAvanceClick(Sender: TObject);
    procedure BDupFiltreClick(Sender: TObject);
  private    { Déclarations privées }
    RuptAvantRappelFiltre : TQuelleRupt ;
    procedure ChargeAvances(OnLib : Boolean) ;
    Function MultiTri : String ;
  public    { Déclarations publiques }
    AGLDesign : boolean ;  
    NomListeGA       : String ;
    ProgressDebit, ProgressCredit            : Double ;
    Exo,Etab,Dev,DevEnP : Boolean ;
    AvecJoker        : Boolean ;
    BTitreOK         : Boolean ;
    PremTabColBal    : TTabCollAff ;
    PremTabColEdt    : TTabCollAffEdt ;
    CritEDT          : TCritEdt ;
    // GC - le 13/12/2001
    CritEdtChaine    : TCritEdtChaine;
    //
    FNatureBase      : TNatureBase ;
    FNatureEtat      : TNatureEtat ;
    FModeSelect      : TModeSelect ;
    FComposite       : Boolean ;
    FNomFiltre       : String ;
    OkMajEdt         : Boolean ;
    OkZoomEdt        : Boolean ;
    NumErreurCrit    : Integer ;
    GCAlc            : TGCalculCum ;
    Etat             : TEtat ;
    HauteurBandeRuptIni : Integer ;
    QRSautPage       : Boolean ;
    QRLoading,QrloadingExo,QRChangeDEV,QRLoadingCptAssocie : Boolean ;
    QRCritMemoire,QRCritMemoireDefaut    : HTStrings ;
    OkChargeCritMemoire, OKFiltreDefaut : Boolean ;


    {JP 23/08/04 : FQ 14242 : nouvelle gestion des filtres pour corrigés une violation d'accès}
    FListeByUser    : TListByUser;

    procedure GenereSQL ; virtual ;
    procedure FinirPrint ; virtual ;
    procedure InitDivers ; virtual ;
    procedure RenseigneCritere ; virtual ;
    procedure ChoixEdition ; virtual ;
    procedure RecupCritEdt ; virtual ;
    function  CritOk : Boolean ; virtual ;


    procedure InitAddFiltre    (T : TOB);
    procedure InitGetFiltre    (T : TOB);
    procedure InitSelectFiltre (T : TOB);
    procedure UpgradeFiltre    (T : TOB);
    procedure ParseParamsFiltre(Params : HTStrings);
    {FIN FQ 14242}

    (*    STTS : String ; // Champ tri sup pour edition cumuls*)
    procedure AfficheEuroSurFiltreDevise ;
    Procedure InitType (NB : TNatureBase ; NE : TNatureEtat ; MS : TModeSelect ; NF,NomGa : String ; MajEdt,ZoomEdt,Composite  : Boolean ; ED : TEdition) ;
    procedure GeneOuTiers ;
    Procedure QuiVisible ;
    Procedure GereZoneJoker(Sender: TObject ; NomControle : String ; HCpt : THCpteEdit) ;
    procedure InitEtatChaine( vCodeEtat : String );  // GC - 20/12/2001
    Function  RecupVCS(QV : TVariantField) : Variant ;
    Function  FaitSQLCompQR : String ;
    Function  GetPref : String ;
    Function  GetNomChampSTTS : String ;
    {JP 28/06/06 : FQ 16149 : gestion des réstrictions Etablissements et à défaut des ParamSoc}
    procedure GereEtablissement;
    {JP 28/06/06 : FQ 16149 : on s'assure que le filtre coincide avec les restrictions utilisateurs sur l'établissement}
    procedure ControlEtab;
  end;

Procedure ZoomEdt(Qui : Integer ; Quoi : String) ;
Procedure ZoomEdtEtat(Quoi : String) ;
function SQLExiste(PreE,PreJ,PreB,TOTDEBIT,TOTCREDIT,TOTDEBITANO,TOTCREDITANO,Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState ;
                   LDate1,LDate2 : TDateTime ; CodeExo : String ; DEV : String ; TheAxe : String = '' ; FiltreSup : String = '') : String ;

// GC - 20/12/2001
function SQLExisteComp(PreE,PreJ,PreB,TOTDEBIT,TOTCREDIT,TOTDEBITANO,TOTCREDITANO,Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState ;
                       LDate1,LDate2 : TDateTime ; CodeExo : String ; DEV : String ; VCritEdt : TCritEdt ) : String ;

Function WhatExisteComp(PreE,PreJ,PreB,Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState ;
                        LDate1,LDate2 : TDateTime ; CodeExo : String ; WhatQualif : Byte ; DEV : String ; vCritEdt : TCritEdt ) : String ;
// GC - 20/12/2001

{$IFNDEF IMP}
{$IFDEF AMORTISSEMENT}
Procedure ZoomEdtEtatImmo(Quoi : String) ;
{$ENDIF}
{$ENDIF}

implementation

{$R *.DFM}

uses ParamDBG,DBCtrls,QRE, TabLiEdt, TriTabLi, QRRupt,
     CPSection_TOM,
{$IFDEF AMORTISSEMENT}
  IMMO_TOM,
{$ENDIF}
{$IFNDEF IMP}
  Saisie, Lettrage, SaisBor,
{$IFNDEF CCMP}
     SaisODA,
     {$IFNDEF NOQR}
     QRGLGen,
     {$ENDIF}
   {$IFNDEF CCS3}
     BudGene, BudSect,
   {$ENDIF}
{$ENDIF}
{$ENDIF}
{$IFNDEF CCS3}
     HZoomSP,
{$ENDIF}
     CPJournal_TOM,
     CPGeneraux_TOM,
     CPTiers_TOM
    ,CbpMCD
    ,CbpEnumerator
(*
{$IFNDEF IMP}
     ,Exercice   // GC - 20/12/2001
{$ENDIF}
*)
  ;
Procedure NumPourSpeed(ED : TEdition) ;
Var Num : Integer ;
BEGIN
Num:=0 ;
Case ED.Etat Of
  etBalGen,etBalAux : Num:=-1 ;
  END ;
MajBeforeDispatch(Num) ;
END ;

Procedure TFQR.InitType (NB : TNatureBase ; NE : TNatureEtat ; MS : TModeSelect ; NF,NomGa : String ; MajEdt,ZoomEdt,Composite : Boolean ; ED : TEdition) ;
BEGIN
NumPourSpeed(ED) ;
TabComplement.TabVisible:=False ; HauteurBandeRuptIni:=0 ; QRSautPage:=TRUE ;
//TabComplement.TabVisible:=(NE=neBal) or (NE=neGl) or (NE=neJU) or (NE=neGLV) and ( Not Composite) ;
If ed.Etat In [etTvaHT,etTvaFac] Then TabRuptures.TabVisible:=FALSE Else
   BEGIN
   TabRuptures.TabVisible:=(NE=neBal) or (NE=neGl) or (NE=neJU) or (NE=neGLV) or (NE=neBalBud) or (NE=neGlBud);
   END ;
FGroupChoixRupt.Visible:=TabRuptures.TabVisible ;
If Not FGroupChoixRupt.Visible Then
   BEGIN
   FSansRupt.Visible:=FALSE ; FAvecRupt.Visible:=FALSE ; FSurRupt.Visible:=FALSE ;
   END ;
BTitreZoom.Visible:=FALSE ;
FNomFiltre:=NF ; FNatureBase:=NB ; FNatureEtat:=NE ; FModeSelect:=MS ; FComposite:=Composite ;
OkMajEdt:=MajEdt ;
Etat:=ED.Etat ;
//OkMajEdt:=FALSE ;
// GP Pour voir l'impact sur les temps en sortie de preview pour les testeurs
OkZoomEdt:=ZoomEdt ; NomListeGA:=NomGA ;
Avance.Visible:=NomListeGA<>'' ;
Avance.Visible:=FALSE ; // GP Fiche n° 1413
OnCum.Visible:=FALSE ; OnCum.Checked:=FALSE ;
(* GP N° 10092 le 07/06/2002
If V_PGI.SAV Then if (NE=neBal) or (NE=neBalLib) or (NE=neCum) Then OnCum.Visible:=TRUE ;
*)
PAvances.TabVisible:=FALSE ;
If EstSpecif('51187') Then
  BEGIN
  if (NE=neBal) Then OnCum.Visible:=TRUE ;
  END ;
If EstSpecif('51186') Then
  BEGIN
  if (NE=neBal) Then PAvances.TabVisible:=TRUE ;
  END ;
BParamListe.Visible:=Avance.Visible And Avance.Checked ;
if FModeSelect=msRien then
   BEGIN
   TSelectCpte.Visible:=FALSE ; FSelectCpte.Visible:=FALSE ;
   TRSelectCpte.Visible:=FALSE ; RSelectCpte.Visible:=FALSE ;
   END ;
(* Rony 19/11/97
FNatureCpt.Visible:=Not (NE=neBalLib) ;
HLabel2.Visible:=FNatureCpt.Visible ; *)
if OkZoomEdt then
   BEGIN
   FListe.Checked:=FALSE ;
   Reduire.Checked:=FALSE ;
   Apercu.Checked:=TRUE ;
   FCouleur.Checked:=V_PGI.QRCouleur ;
   // Simon
   Case NE of
      neBal: InitFormatBal(Self,PremTabColBal) ;
      neGL : InitFormatEdt(self,PremTabColEdt,CritEDT.GL.FormatPrint.TabColl) ;
      neJU : InitFormatEdt(Self,PremTabColEdt,CritEDT.JU.FormatPrint.TabColl) ;
      neJal: InitFormatEdt(Self,PremTabColEdt,CritEdt.Jal.FormatPrint.TabColl) ;
      neJalR: InitFormatEdt(Self,PremTabColEdt,CritEdt.JalR.FormatPrint.TabColl) ;
      neEch: InitFormatEdt(Self,PremTabColEdt,CritEdt.Ech.FormatPrint.TabColl) ;
      neGLV: InitFormatEdt(Self,PremTabColEdt,CritEdt.GlV.FormatPrint.TabColl) ;
      neCum: InitFormatBal(Self,PremTabColBal) ;
      neBrouAna : InitFormatEdt(self,PremTabColEdt,CritEDT.BrouAna.FormatPrint.TabColl) ;
      neJbq: InitFormatEdt(Self,PremTabColEdt,CritEdt.Jbq.FormatPrint.TabColl) ;
      nePoi: InitFormatEdt(Self,PremTabColEdt,CritEdt.Poi.FormatPrint.TabColl) ;
      neRap: InitFormatEdt(Self,PremTabColEdt,CritEdt.Rap.FormatPrint.TabColl) ;

      neBalBud : InitFormatEdt(Self,PremTabColEdt,CritEdt.BalBud.FormatPrint.TabColl) ;
      neBroBud : InitFormatEdt(self,PremTabColEdt,CritEDT.BroBud.FormatPrint.TabColl) ;
      neBroBudAna : InitFormatEdt(self,PremTabColEdt,CritEDT.BroBudAna.FormatPrint.TabColl) ;
      neBalLib : InitFormatBal(Self,PremTabColBal) ;
      nePlanEdt: InitFormatEdt(Self,PremTabColEdt,CritEdt.PlanEdt.FormatPrint.TabColl) ;
      END ;
   InitEdit(EntetePage,QRP) ;
   If OnImpr Then Apercu.Checked:=FALSE ;
   RemplirCritEdtList ;
   BValiderClick(Nil) ;
   END ;
END ;
(*
function WhatTypeEcr(Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState) : SetttTypePiece ;
//Type ttTypePiece  = (tpReel,tpSim,tpPrev,tpSitu,tpRevi) ;
BEGIN
Result:=[] ;
if (Valeur<>'TOU') Then
   BEGIN
   if Valeur='NOR' then Result:=[tpReel]  ;
   if Valeur='NSS' then Result:=[tpReel,tpSim,tpSitu] ;
   if Valeur='SSI' then Result:=[tpSim,tpSitu] ;
   if Valeur='PRE' then Result:=[tpPrev] ;
   END Else Result:=[tpReel,tpSim,tpSitu,tpPrev] ;
If Controleur Then
   BEGIN
   Case EtatRevision Of
     cbchecked : Result:=[tpRevi] ;
     cbGrayed  : Result:=Result+[tpRevi] ;
     END ;
   END ;
END ;

Function WhereSupp(P : String ; SetTyp : SetttTypePiece) : String ;
BEGIN
Result:='' ;
If tpReel In SetTyp Then Result:=Result+'OR '+P+'QUALIFPIECE="N" ' ;
If tpSim  In SetTyp Then Result:=Result+'OR '+P+'QUALIFPIECE="S" ' ;
If tpPrev In SetTyp Then Result:=Result+'OR '+P+'QUALIFPIECE="P" ' ;
If tpSitu In SetTyp Then Result:=Result+'OR '+P+'QUALIFPIECE="U" ' ;
If tpRevi In SetTyp Then Result:=Result+'OR '+P+'QUALIFPIECE="R" ' ;
If Result<>'' Then
   BEGIN
   Result:=Copy(Result,4,Length(Result)-3) ;
   Result:=' And ('+Result+')' ;
   END ;
END ;


Function WhatExiste(PreE,PreJ,PreB,Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState ;
                    LDate1,LDate2 : TDateTime ; CodeExo : String) : String ;
Var StTable,St,St2 : String ;
    SetType : SetttTypePiece ;
begin
if PreE='Y_' then StTable:='ANALYTIQ' else StTable:='ECRITURE' ;
St:='(exists (Select '+PreE+PreJ+','+PreE+'EXERCICE,'+PreE+'DATECOMPTABLE From '+StTable+' ' ;
St:=St+' Where ('+PreE+PreJ+'=Q.'+PreB+PreJ+' ' ;
St:=St+' And '+PreE+'EXERCICE="'+CodeExo+'" ';
St:=St+' And '+PreE+'DATECOMPTABLE>="'+UsDateTime(LDate1)+'" And '+PreE+'DATECOMPTABLE<="'+UsDateTime(LDate2)+'" ' ;
SetType:=WhatTypeEcr(Valeur,Controleur,EtatRevision) ;
St2:=WhereSupp(PreE,SetType) ;
If St2<>'' Then St:=St+St2 ;
St:=St+')))' ;
Result:=St ;
end ;
*)
function SQLExiste(PreE,PreJ,PreB,TOTDEBIT,TOTCREDIT,TOTDEBITANO,TOTCREDITANO,Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState ;
                   LDate1,LDate2 : TDateTime ; CodeExo : String ; DEV : String ; TheAxe : String ='' ; FiltreSup : String = '') : String ;
Var St,St1,St2 : String ;
BEGIN
St:='' ; St1:='' ; St2:='' ;
If (TOTDEBIT<>'') And (TOTCREDIT<>'') Then St1:=TOTDEBIT+'<>0 OR '+TOTCREDIT+'<>0 OR ' ;
If (TOTDEBITANO<>'') And (TOTCREDITANO<>'') Then St2:=TOTDEBITANO+'<>0 OR '+TOTCREDITANO+'<>0 ' ;
if Valeur='NOR' then
   BEGIN
   If (St1<>'') And (St2<>'') Then St:='('+St1+St2+') ';
   If st<>'' Then
      BEGIN
      (* GP le 05/11/99 pour bug sur non apparition des comptes d'écarts de conversionsur les monnaies opposées
      If V_PGI.Controleur Then
      *)
         St:='('+St+'Or '+WhatExiste(PreE,PreJ,PreB,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,0,Dev,TheAxe,FiltreSup)+')' ;
      END Else
         St:=WhatExiste(PreE,PreJ,PreB,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,0,Dev,TheAxe,FiltreSup) ;
   END ;
if Valeur='NSS' then
   BEGIN
   If (st1<>'') And (St2<>'') Then
      BEGIN
      St:='(('+St1+St2+') ';
      St:=St+'Or '+WhatExiste(PreE,PreJ,PreB,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,0,Dev,TheAxe,FiltreSup)+')' ;
      END Else
      BEGIN
      St:=WhatExiste(PreE,PreJ,PreB,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,0,Dev,TheAxe,FiltreSup) ;
      END ;
   END ;
if Valeur='SSI' then
   BEGIN
   St:=WhatExiste(PreE,PreJ,PreB,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,0,Dev,TheAxe,FiltreSup) ;
   END ;
if Valeur='PRE' then
   BEGIN
   St:=WhatExiste(PreE,PreJ,PreB,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,0,Dev,TheAxe,FiltreSup) ;
   END ;
Result:=St ;
END ;

// GC - 20/12/2001
function SQLExisteComp(PreE,PreJ,PreB,TOTDEBIT,TOTCREDIT,TOTDEBITANO,TOTCREDITANO,Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState ;
                       LDate1,LDate2 : TDateTime ; CodeExo : String ; DEV : String ; VCritEdt : TCritEdt ) : String ;
Var St,St1,St2 : String ;
begin
  St:='' ; St1:='' ; St2:='' ;

  if (TOTDEBIT <> '') and (TOTCREDIT <> '' ) then      St1 := TOTDEBIT + '<>0 OR ' + TOTCREDIT + '<>0 OR ' ;
  if (TOTDEBITANO <> '') and (TOTCREDITANO <> '') then St2 := TOTDEBITANO + '<>0 OR ' + TOTCREDITANO + '<>0 ' ;

  // Ecritures de type normale
  if Valeur='NOR' then
  begin
    if (St1<>'') and (St2<>'') then St:='('+St1+St2+') ';
    if st <> '' then
    begin
      (* GP le 05/11/99 pour bug sur non apparition des comptes d'écarts de conversionsur les monnaies opposées
      If V_PGI.Controleur Then
      *)
        // GC
      St := '('+St+'Or '+WhatExisteComp(PreE,PreJ,PreB,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,0,Dev, vCritEdt)+')'
    end
    else
      St:= WhatExisteComp(PreE,PreJ,PreB,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,0,Dev, vCritEdt) ;
  end ;

  // Ecritures de type Normal, Simulation, Situtation
  if Valeur = 'NSS' then
  begin
    if (st1 <> '') and ( St2 <> '') then
    begin
      St := '(('+St1+St2+') ';
      St:=St+'Or '+WhatExisteComp(PreE,PreJ,PreB,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,0,Dev,vCritEdt) + ')' ;
    end
    else
    begin
      St:=WhatExisteComp(PreE,PreJ,PreB,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,0,Dev, vCritEdt) ;
    end ;
  end ;

  // Ecritures de type Simulation et Situation
  if Valeur = 'SSI' then
  begin
    St := WhatExisteComp(PreE,PreJ,PreB,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,0,Dev, vCritEdt) ;
  end ;

  // Ecritures de type Prevision
  if Valeur='PRE' then
  begin
    St:=WhatExisteComp(PreE,PreJ,PreB,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,0,Dev, vCritEdt) ;
  end ;

Result:=St ;

END ;

// GC
Function WhatExisteComp(PreE,PreJ,PreB,Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState ;
                        LDate1,LDate2 : TDateTime ; CodeExo : String ; WhatQualif : Byte ; DEV : String ; vCritEdt : TCritEdt ) : String ;
Var StTable,St,St2 : String ;
    SetType : SetttTypePiece ;
    i : integer;
begin

  if PreE='Y_' then
    StTable:='ANALYTIQ'
  else
    if PreE='BE_' then
      StTable:='BUDECR'
    else StTable:='ECRITURE' ;

  St := '(';
  for i:=1 to 2 do
  begin
    if (i = 2) and (vCritEdt.CompareBalSit) then // Cas comparatif avec balance de situation
    begin
      St := St + ' ' + '(exists (Select BSE_COMPTE1 From CBALSITECR ' ;
      St := St + ' Where (BSE_CODEBAL="'+vCritEdt.BalSit+'" AND BSE_COMPTE1=Q.'+PreB+PreJ+'))) ' ;
      break;
    end
    else
    begin
      St := St + ' ' + '(exists (Select '+PreE+PreJ+','+PreE+'EXERCICE,'+PreE+'DATECOMPTABLE From '+StTable+' ' ;
      St := St + ' Where ('+PreE+PreJ+'=Q.'+PreB+PreJ+' ' ;
    end;

    if CodeExo <> '' then
    begin
      if (i = 1) then
        St := St + ' And ' + PreE + 'EXERCICE="' + CodeExo + '" '
      else
        St := St + ' And ' + PreE + 'EXERCICE="' + vCritEdt.ExoComparatif.Code + '" ' ;
    end;

    if (i = 1) then
      St :=St + ' And '+ PreE + 'DATECOMPTABLE>="' + UsDateTime(LDate1) + '" And ' + PreE + 'DATECOMPTABLE<="' + UsDateTime(LDate2) + '" '
    else
      St :=St + ' And '+ PreE + 'DATECOMPTABLE>="' + UsDateTime(vCritEdt.ExoComparatif.Deb) + '" And ' + PreE + 'DATECOMPTABLE<="' + UsDateTime(vCritEdt.ExoComparatif.Fin) + '" ';

    if PreE='BE_' then
    begin
      if DEV <> '' then
        St:=St+' And '+PreE+'BUDJAL="'+DEV+'" ' ;
    end
    else
    begin
      if DEV<>'' Then St:=St+' And '+PreE+'DEVISE="'+DEV+'" ' ;
    end ;

    if WhatQualif=0 then
      SetType:=WhatTypeEcr(Valeur,Controleur,EtatRevision)
    else
      if WhatQualif=1 then
        SetType:=WhatTypeEcr2(Valeur) ;

    St2:=WhereSupp(PreE,SetType) ;

    if St2 <> '' then
      St:=St+St2 ;

    // St:=St+'))' ;
    St:=St+')))' ;

    (*
    if (i = 1) then
      St := St + ') OR ';
    else
      St := St + ')';*)

    if (i = 1) then
      St := St + ' OR ';
  end;

  Result:=St + ')';
end ;
// GC - 20/12/2001 - FIN

function TFQR.FquelAN(Inclure : Boolean) : TQuelAN ;
Var LTypCpt : Integer ;
     OkExo : Boolean ;
BEGIN
Result:=AvecAN ;
If OkZoomEdt Then Exit ;
OkExo:=(CritEdt.MonoExo) ;
Case FNatureEtat Of
  neBal : BEGIN LTypCpt:=CritEdt.Bal.TypCpt ; OkExo:=OkExo And (CritEdt.Date1=CritEdt.Bal.Date11) ; END ;
  neGL  : BEGIN LTypCpt:=CritEdt.GL.TypCpt ; OkExo:=OkExo And (CritEdt.Date1=CritEdt.GL.Date11) ; END ;
  else Exit ;
  END ;
If (LTypCpt=3) And (Not OkExo) And (Not Inclure) Then Result:=SansAN ;
END ;

Function CptBalClo(Cpt : String) : String ;
BEGIN
Result:='' ; If Cpt<>'' Then Result:=' OR (G_GENERAL="'+Cpt+'" )' ;
END ;
(*
Function CptBalEcc(Cpt : String ; FNatureBase : TNatureBase ; fModeSelect : tModeSelect ; Var CritEdt : tCritEdt) : String ;
Var OkOk : Boolean ;
    i,ll : Integer ;
BEGIN
// Forcer une recherche sur les comptes d'écart de convertion quipeuvent être mouvements en opposé et jamais en pivot
Result:='' ; If Cpt='' Then Exit ;
If (FNatureBase=nbGen) And (FModeSelect=msGenEcr) And (Not CritEdt.Composite) Then Else Exit ;
ll:=Length(Cpt) ; OkOk:=TRUE ;
if CritEDT.Joker then
  BEGIN
  For i:=0 To Length(CritEDT.Cpt1) Do
    BEGIN
    If (CritEdt.Cpt1[i]<>'?') And (CritEdt.Cpt1[i]<>'*') And (i<=ll) Then If CritEdt.Cpt1[i]<>Cpt[i] Then OkOk:=FALSE ;
    If Not OkOk Then Break ;
    END
  END Else
  BEGIN
  if CritEDT.Cpt1<>'' then If Cpt<CritEDT.Cpt1 Then OkOk:=FALSE ;
  if CritEDT.Cpt2<>'' then If Cpt>CritEDT.Cpt2 Then OkOk:=FALSE ;
  END ;
If Not Okok Then Exit ;
Result:=CptBalClo(Cpt) ;
END ;
*)
function  TFQR.SQLModeSelect : String ;
{ Contruction d'une partie de SQLAux }
Var TOTDEBIT,TOTCREDIT,TOTDEBITANO,TOTCREDITANO,St,St1,StEcc : String ;
    PreB,PreE,PreJ,ResC,BilC,PerC,BenC,BilO,BenO,PerO : String ;
    TheAxe : String ;
    Exo : TExoDate ;
    ts : Integer ;
    me : Boolean ;
    EdtSimple : Boolean ;
    Valeur : String ;
    DD1,DD2 : TDateTime ;
    WhatAN : TQuelAN ;
{$IFDEF SPEC302}
    QClo : TQuery ;
{$ENDIF}
BEGIN
//Simon
// GP : à revoir dans le cas : Que les révisions
Result:='' ; EdtSimple:=FALSE ; WhatAN:=AvecAN ; StEcc:='' ; TheAxe:='' ;
If (Not CritEdt.Composite) And
   ((FModeSelect=msGenEcr) Or (FModeSelect=msGenAna) or (FModeSelect=msAuxEcr)) Then
   BEGIN
   If (FNatureBase=nbGen) Or (FNatureBase=nbAux) Or (FNatureBase=nbSec) Then EdtSimple:=TRUE ;
   If EdtSimple Then
      BEGIN
      Case FNatureEtat Of
        neBal : WhatAN:=CritEdt.Bal.QuelAN ;
        neGL : WhatAN:=CritEdt.GL.QuelAN ;
        neCum : WhatAN:=SansAN ;
        Else BEGIN WhatAN:=AvecAN ; EdtSimple:=FALSE ; END ;
        END ;
      END Else EdtSimple:=FALSE ;
   END ;
Case FModeSelect of
   msGenEcr,msGenAna : PreJ:='GENERAL' ;
   msAuxEcr : PreJ:='AUXILIAIRE' ;
   msSecAna : PreJ:='SECTION' ;
   else Exit;
   END ;

Case FNatureBase of
   nbGen,nbGenT : PreB:='G_' ;
   nbAux : PreB:='T_' ;
   nbSec : PreB:='S_' ;
   nbJal,nbJalAna : PreB:='J_' ;
   nbBudGen : PreB:='BG_' ;
   nbBudSec : PreB:='BS_' ;
   nbBudJal : PreB:='BJ_' ;
   else Exit ;
   END ;

Case FModeSelect of
   msGenEcr,msAuxEcr : PreE:='E_' ;
   msGenAna,msSecAna : BEGIN
                       PreE:='Y_' ;
                       Case FNatureEtat Of
                         neBal : TheAxe:=CritEdt.Bal.Axe ;
                         neGL : TheAxe:=CritEdt.GL.Axe ;
                         END ;
                       END ;
   msGenBEcr, msSecBAna : PreE:='BE_' ;
   else Exit ;
   END ;

TOTDEBIT:=PreB+'TOTDEBE' ; TOTCREDIT:=PreB+'TOTCREE' ;
Case FNatureEtat of
   neBal : BEGIN me:=CritEDT.MonoExo ; Ts:=CritEDT.BAL.TypCpt ; Exo:=CritEDT.Exo ; END ;
   neGL  : BEGIN me:=CritEDT.MonoExo ; Ts:=CritEDT.GL.TypCpt ; Exo:=CritEDT.Exo ; END ;
   neCum : BEGIN me:=CritEDT.MonoExo ; Ts:=CritEDT.Cum.TypCpt ; Exo:=CritEDT.Exo ; END ;
   Else exit ;
   END ;
If ME Then
   BEGIN
   If Exo.Code=VH^.Precedent.Code Then BEGIN TOTDEBIT:=PreB+'TOTDEBP' ; TOTCREDIT:=PreB+'TOTCREP' ; END ;
   If Exo.Code=VH^.EnCours.Code   Then BEGIN TOTDEBIT:=PreB+'TOTDEBE' ; TOTCREDIT:=PreB+'TOTCREE' ; END ;
   If Exo.Code=VH^.Suivant.Code   Then BEGIN TOTDEBIT:=PreB+'TOTDEBS' ; TOTCREDIT:=PreB+'TOTCRES' ; END ;
   END ;
Valeur:=CritEdt.QualifPiece ;
TOTDEBITANO:=PreB+'TOTDEBANO' ; TOTCREDITANO:=PreB+'TOTCREANO' ;
If ME And (Exo.Code=VH^.Suivant.Code) Then
   BEGIN
   TOTDEBITANO:=PreB+'TOTDEBANON1' ; TOTCREDITANO:=PreB+'TOTCREANON1' ;
   END ;
//Type ttTypePiece  = (tpReel,tpSim,tpPrev,tpSitu,tpRevi) ;
//TS = 0 : Exo, 1 = Tous, 2 = Cpt non soldés, 3 = Période
Case TS of
  0 : BEGIN
      if (Valeur<>'TOU') Then
         BEGIN
         // GC - 20/12/2001
         if (Etat = etBalGeneComp) and (CritEdt.AvecComparatif) then
           St:=SQLExisteComp(PreE,PreJ,PreB,TOTDEBIT,TOTCREDIT,TOTDEBITANO,TOTCREDITANO,Valeur,V_PGI.Controleur,AvecRevision.State,CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,CritEdt.DEVPourExist,CritEdt)
         else
         // GC - 20/12/2001 - FIN
         St:=SQLExiste(PreE,PreJ,PreB,TOTDEBIT,TOTCREDIT,TOTDEBITANO,TOTCREDITANO,Valeur,V_PGI.Controleur,AvecRevision.State,CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,CritEdt.DEVPourExist,TheAxe,CritEdt.FiltreSup) ;
         If CritEdt.Bal.PourCloture Then
            BEGIN
            ResC:='' ; BilC:='' ; PerC:='' ; BenC:='' ; BilO:='' ; PerO:='' ; BenO:='' ;
{$IFDEF SPEC302}
            QClo:=OpenSQL('SELECT SO_FERMEBIL, SO_OUVREBIL, SO_RESULTAT, SO_FERMEPERTE, SO_OUVREPERTE, SO_FERMEBEN, SO_OUVREBEN FROM SOCIETE WHERE SO_SOCIETE="'+V_PGI.CodeSociete+'" ',TRUE) ;
            If Not QCLO.EOF Then
               BEGIN
               BilC:=QCLO.FindField('SO_FERMEBIL').AsString ;
               ResC:=QCLO.FindField('SO_RESULTAT').AsString ;
               PerC:=QCLO.FindField('SO_FERMEPERTE').AsString ;
               BenC:=QCLO.FindField('SO_FERMEBEN').AsString ;
               BilO:=QCLO.FindField('SO_OUVREBIL').AsString ;
               PerO:=QCLO.FindField('SO_OUVREPERTE').AsString ;
               BenO:=QCLO.FindField('SO_OUVREBEN').AsString ;
               END ;
            Ferme(QClo) ;
{$ELSE}
            BilC:=GetParamSocSecur('SO_FERMEBIL','') ;
            ResC:=GetParamSocSecur('SO_RESULTAT','') ;
            PerC:=GetParamSocSecur('SO_FERMEPERTE','') ;
            BenC:=GetParamSocSecur('SO_FERMEBEN','') ;
            BilO:=GetParamSocSecur('SO_OUVREBIL','') ;
            PerO:=GetParamSocSecur('SO_OUVREPERTE','') ;
            BenO:=GetParamSocSecur('SO_OUVREBEN','') ;
{$ENDIF}
            St1:=CptBalClo(BilC)+CptBalClo(ResC)+CptBalClo(PerC)+CptBalClo(BenC)+
                 CptBalClo(BilO)+CptBalClo(PerO)+CptBalClo(BenO) ;
            If St1<>'' Then St:='( '+St+' '+St1+' )' ;
            END ;
         END Else
         BEGIN
         St:=WhatExiste(PreE,PreJ,PreB,Valeur,V_PGI.Controleur,AvecRevision.State,CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,0,CritEdt.DEVPourExist,TheAxe,CritEdt.FiltreSup) ;
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
            // GC - 20/12/2001
            if (Etat = etBalGeneComp) and (CritEdt.AvecComparatif)then
              Result:=Result+WhatExisteComp(PreE,PreJ,PreB,'SSI',V_PGI.Controleur,AvecRevision.State,CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,0,CritEdt.DEVPourExist,CritEdt)+'))'
            else
            // GC - 20/12/2001 - FIN
            Result:=Result+WhatExiste(PreE,PreJ,PreB,'SSI',V_PGI.Controleur,AvecRevision.State,CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,0,CritEdt.DEVPourExist,TheAxe,CritEdt.FiltreSup)+'))' ;
            END ;
         if Valeur='NSS' then
            BEGIN
            Result:=PreB+'TOTALDEBIT<>'+PreB+'TOTALCREDIT ' ;
            Result:=Result+' Or (('+PreB+'TOTALDEBIT='+PreB+'TOTALCREDIT '+') And (' ;
            // GC - 20/12/2001
            if (Etat = etBalGeneComp) and (CritEdt.AvecComparatif)then
              Result:=Result+WhatExisteComp(PreE,PreJ,PreB,'SSI',V_PGI.Controleur,AvecRevision.State,CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,0,CritEdt.DEVPourExist,CritEdt)+'))'
            else
            // GC - 20/12/2001 - FIN
            Result:=Result+WhatExiste(PreE,PreJ,PreB,'SSI',V_PGI.Controleur,AvecRevision.State,CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,0,CritEdt.DEVPourExist,TheAxe,CritEdt.FiltreSup)+'))' ;
            END ;
         END Else
         BEGIN
         Result:=PreB+'TOTALDEBIT<>'+PreB+'TOTALCREDIT ' ;
         Result:=Result+' Or (('+PreB+'TOTALDEBIT='+PreB+'TOTALCREDIT '+') And (' ;
         Result:=Result+WhatExiste(PreE,PreJ,PreB,Valeur,V_PGI.Controleur,AvecRevision.State,CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,0,CritEdt.DEVPourExist,TheAxe,CritEdt.FiltreSup)+'))' ;
         END ;
      If Result<>'' Then Result:=' ('+Result+') ' ;
      END ;
  3 : BEGIN
      DD1:=CritEdt.DateDeb ; DD2:=CritEdt.DateFin ;
      If EdtSimple And (WhatAN=SansAN) Then BEGIN TOTDEBITANO:='' ; TOTCREDITANO:='' ; END ;
      If EdtSimple And (WhatAN=AvecAN) And (FNatureBase=nbGen) Then BEGIN DD1:=CritEdt.Date1 ; END ;
      if (Valeur<>'TOU') Then
         BEGIN
         // GC - 20/12/2001
         if (Etat = etBalGeneComp) and (CritEdt.AvecComparatif)then
           St:=SQLExisteComp(PreE,PreJ,PreB,'','',TOTDEBITANO,TOTCREDITANO,Valeur,V_PGI.Controleur,AvecRevision.State,DD1,DD2,CritEdt.Exo.Code,CritEdt.DEVPourExist,CritEdt)
         else
         // GC - 20/12/2001 - FIN
         St:=SQLExiste(PreE,PreJ,PreB,'','',TOTDEBITANO,TOTCREDITANO,Valeur,V_PGI.Controleur,AvecRevision.State,DD1,DD2,CritEdt.Exo.Code,CritEdt.DEVPourExist,TheAxe,CritEdt.FiltreSup) ;
         END Else
         BEGIN
         St:=WhatExiste(PreE,PreJ,PreB,Valeur,V_PGI.Controleur,AvecRevision.State,DD1,DD2,CritEdt.Exo.Code,0,CritEdt.DEVPourExist,TheAxe,CritEdt.FiltreSup) ;
         END ;
      Result:=St ;
      END ;
 end ;
END ;
(*
function SQLExisteMul(PreE,PreJ1,PreJ2,PreB1,PreB2,TOTDEBIT,TOTCREDIT,Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState ;
                      LDate1,LDate2 : TDateTime ; CodeExo : String ; Where1,Where2 : String) : String ;
Var St : String ;
BEGIN
St:='' ;
if Valeur='NOR' then
   BEGIN
   St:='('+TOTDEBIT+'<>0 OR '+TOTCREDIT+'<>0 OR '+PreB1+'TOTDEBANO<>0 OR '+PreB1+'TOTCREANO<>0) ';
   St:='('+St+'And '+
       WhatExisteMul(PreE,PreJ1,PreJ2,PreB2,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,Where1,Where2,0)
        +')' ;
   END ;
if Valeur='NSS' then
   BEGIN
   St:='(('+TOTDEBIT+'<>0 OR '+TOTCREDIT+'<>0 OR '+PreB1+'TOTDEBANO<>0 OR '+PreB1+'TOTCREANO<>0) ';
   St:='('+St+'And '+
       WhatExisteMul(PreE,PreJ1,PreJ2,PreB2,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,Where1,Where2,0)
       +')' ;
   END ;
if Valeur='SSI' then
   BEGIN
   St:=WhatExisteMul(PreE,PreJ1,PreJ2,PreB2,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,Where1,Where2,0) ;
   END ;
if Valeur='PRE' then
   BEGIN
   St:=WhatExisteMul(PreE,PreJ1,PreJ2,PreB2,Valeur,Controleur,EtatRevision,LDate1,LDate2,CodeExo,Where1,Where2,0) ;
   END ;
Result:=St ;
END ;
*)
function WhatPREBPREJ(FM : TModeSelect ; Var PreJ1,PreJ2,PreB1,PreB2,PreE,Table1,Table2 : String) : Boolean ;
BEGIN
result:=FALSE ;
Case FM of
   msGenEcr : BEGIN
              PreJ1:='GENERAL' ; PreJ2:='AUXILIAIRE' ; PreB1:='G_' ; PreB2:='T_' ; PreE:='E_' ;
              Table1:='GENERAUX' ; Table2:='TIERS' ;
              END ;
   msGenAna : BEGIN
              PreJ1:='GENERAL' ; PreJ2:='SECTION' ; PreB1:='G_' ; PreB2:='S_' ; PreE:='Y_' ;
              Table1:='GENERAUX' ; Table2:='SECTION' ;
              END ;
   msAuxEcr : BEGIN
              PreJ1:='AUXILIAIRE' ; PreJ2:='GENERAL' ; PreB1:='T_' ; PreB2:='G_' ; PreE:='E_' ;
              Table1:='TIERS' ; Table2:='GENERAUX' ;
              END ;
   msSecAna : BEGIN
              PreJ1:='SECTION' ; PreJ2:='GENERAL' ; PreB1:='S_' ; PreB2:='G_' ; PreE:='Y_' ;
              Table1:='SECTION' ; Table2:='GENERAUX' ;
              END ;
   else Exit;
   END ;
Result:=TRUE ;
END ;

function TFQR.QuelTypeEcr : SetttTypePiece ;
Var Valeur : String ;
BEGIN
Result:=[] ; Valeur:=CritEdt.QualifPiece ; Result:=WhatTypeEcr(valeur,V_PGI.Controleur,AvecRevision.State) ;
END ;

procedure TFQR.RenseigneCritere;
Var St1,St2 : String ;
BEGIN
If OkZoomEdt Then
   BEGIN
   QRP.ReportTitle:=Caption ;
   InitQrPrinter(QRP,FListe.Checked,Reduire.Checked,FCouleur.Checked) ;
   BTitre.Enabled:=FALSE ; BTitreZoom.Enabled:=TRUE ;
   RCptZoom.Caption:=CritEdt.Cpt1 ;
   RDateCompta1Zoom.caption:=DateToStr(CritEdt.DateDeb) ;
   RDateCompta2Zoom.caption:=DateToStr(CritEdt.DateFin) ;
   If CritEdt.QualifPiece='' Then RNatureEcrZoom.Caption:=Traduirememoire('<<Tous>>') Else RNatureEcrZoom.Caption:=RechDom('ttQualPieceCrit',CritEdt.QualifPiece,FALSE) ;
   If CritEdt.DeviseSelect='' Then RDevisesZoom.Caption:=Traduirememoire('<<Tous>>') Else RDevisesZoom.Caption:=RechDom('ttDevise',CritEdt.DeviseSelect,FALSE) ;
   If CritEdt.Etab='' Then REtabZoom.Caption:=Traduirememoire('<<Tous>>') Else REtabZoom.Caption:=RechDom('ttEtablissement',CritEdt.Etab,FALSE) ;
   Exit ;
   END Else BTitre.Enabled:=True ;
ReInitEdit(EntetePage,FCouleur.Checked) ;
if FJoker.Visible then
   BEGIN
   RCpte.Visible:=False ; TRa.visible:=False ;
   RCpte2.Visible:=False ; RJoker.Visible:=True;
   RCpte1.Caption:=FJoker.Text ; RCpte2.Caption:=FJoker.Text ;
   END else
   BEGIN
   RCpte.Visible:=True ; TRa.visible:=True ;
   RCpte2.Visible:=True ; RJoker.Visible:=False;
   PositionneFourchetteST(FCpte1,FCpte2,St1,St2) ;
   RCpte1.Caption:=St1 ;
   RCpte2.Caption:=St2 ;
   END ;
RDateCompta1.Caption:=FDateCompta1.Text ; RDateCompta2.Caption:=FDateCompta2.Text ;
RNatureCpt.Caption:=FNatureCpt.Text     ; RNatureEcr.Caption:=FNatureEcr.Text ;
REtab.Caption:=FEtab.Text               ; RDevises.Caption:=FDevises.Text ;
RExcepGen.Caption:=FExcep.Text          ; RExercice.Caption:=FExercice.Text ;
RSelectCpte.Caption:=FSelectCpte.Text   ;
Case CritEdt.Monnaie Of
  0 : RDevises.Caption:=RDevises.Caption+' / Aff. '+VH^.LibDevisePivot ;
  1 : RDevises.Caption:=RDevises.Caption+TraduireMemoire(' / Aff. Devise') ;
END ;
InitQrPrinter(QRP,FListe.Checked,Reduire.Checked,FCouleur.Checked) ;
END;

procedure TFQR.AfficheEuroSurFiltreDevise ;
BEGIN
If CritEdt.Monnaie=2 Then
  BEGIN
  CritEdt.Decimale:=V_PGI.OkDecE ;
  If Not VH^.TenueEuro Then CritEdt.Symbole:='' ;
  END ;
END ;

procedure TFQR.ChoixEdition ;
{ Initialisation des options d'édition : Format des montant vis à vis des devises }
Var RDev : RDevise ;
BEGIN
//Simon
Case FNatureEtat of
   neBal,neBalLib : BEGIN
         if (CritEDT.DeviseSelect='') or (CritEDT.Monnaie=0) Or (CritEDT.DeviseSelect=V_PGI.DevisePivot) then
            BEGIN
            CritEDT.DeviseAffichee:=V_PGI.DevisePivot ;
            CritEDT.Decimale:=V_PGI.OkDecV ;
            CritEDT.Symbole:=V_PGI.SymbolePivot ;
            END Else
            BEGIN
            CritEDT.DeviseAffichee:=CritEDT.DeviseSelect ;
            RDev.Code:=CritEDT.DeviseAffichee ;
            GetInfosDevise(RDev) ;
            CritEDT.Decimale:=RDev.Decimale ;
            CritEDT.Symbole:=RDev.Symbole ;
            AfficheEuroSurFiltreDevise ;
            END ;
         ChangeMask(Formateur,CritEDT.Decimale,CritEDT.Symbole);
         CritEDT.FormatMontant:=Formateur.Masks.PositiveMask;
         END ;
   neGL : BEGIN
         if (CritEDT.DeviseSelect='') or (CritEDT.Monnaie=0) Or (CritEDT.DeviseSelect=V_PGI.DevisePivot) then
            BEGIN
            CritEDT.DeviseAffichee:=V_PGI.DevisePivot ;
            CritEDT.Decimale:=V_PGI.OkDecV ;
            CritEDT.Symbole:=V_PGI.SymbolePivot ;
            END Else
            BEGIN
            CritEDT.DeviseAffichee:=CritEDT.DeviseSelect ;
            RDev.Code:=CritEDT.DeviseAffichee ;
            GetInfosDevise(RDev) ;
            CritEDT.Decimale:=RDev.Decimale ;
            CritEDT.Symbole:=RDev.Symbole ;
            AfficheEuroSurFiltreDevise ;
            END ;
         ChangeMask(Formateur,CritEDT.Decimale,CritEDT.Symbole) ;
         CritEDT.FormatMontant:=Formateur.Masks.PositiveMask ;
         END ;
   neJU : BEGIN
          if (CritEDT.DeviseSelect='') or (CritEDT.Monnaie=0) Or (CritEDT.DeviseSelect=V_PGI.DevisePivot) then
             BEGIN
             CritEDT.DeviseAffichee:=V_PGI.DevisePivot ;
             CritEDT.Decimale:=V_PGI.OkDecV ;
             CritEDT.Symbole:=V_PGI.SymbolePivot ;
             END Else
             BEGIN
             CritEDT.DeviseAffichee:=CritEDT.DeviseSelect ;
             RDev.Code:=CritEDT.DeviseAffichee ;
             GetInfosDevise(RDev) ;
             CritEDT.Decimale:=RDev.Decimale ;
             CritEDT.Symbole:=RDev.Symbole ;
             AfficheEuroSurFiltreDevise ;
             END ;
          ChangeMask(Formateur,CritEDT.Decimale,CritEDT.Symbole);
          CritEDT.FormatMontant:=Formateur.Masks.PositiveMask;
          END ;
   neJal,
   neJalR,
   neEch,
   neGLV: BEGIN
          if (CritEdt.DeviseSelect='') or (CritEdt.Monnaie=0) Or (CritEdt.DeviseSelect=V_PGI.DevisePivot) then
             BEGIN
             CritEdt.DeviseAffichee:=V_PGI.DevisePivot ;
             CritEdt.Decimale:=V_PGI.OkDecV ;
             CritEdt.Symbole:=V_PGI.SymbolePivot ;
             END Else
             BEGIN
             CritEdt.DeviseAffichee:=CritEdt.DeviseSelect ;
             RDev.Code:=CritEdt.DeviseAffichee ;
             GetInfosDevise(RDev) ;
             CritEdt.Decimale:=RDev.Decimale ;
             CritEdt.Symbole:=RDev.Symbole ;
             AfficheEuroSurFiltreDevise ;
            END ;
          ChangeMask(Formateur,CritEdt.Decimale,CritEdt.Symbole);
          CritEdt.FormatMontant:=Formateur.Masks.PositiveMask;
          END ;
   neCum :BEGIN
          if (CritEDT.DeviseSelect='') (*or (CritEDT.Monnaie=0)*) Or (CritEDT.DeviseSelect=V_PGI.DevisePivot) then
             BEGIN
             CritEDT.DeviseAffichee:=V_PGI.DevisePivot ;
             CritEDT.Decimale:=V_PGI.OkDecV ;
             CritEDT.Symbole:=V_PGI.SymbolePivot ;
             CritEDT.DecimalePivot:=V_PGI.OkDecV ;
             CritEDT.SymbolePivot:=V_PGI.SymbolePivot ;
             If CritEdt.Monnaie=2 Then
               BEGIN
               CritEdt.DecimalePivot:=V_PGI.OkDecE ;
               If Not VH^.TenueEuro Then CritEdt.SymbolePivot:='' ;
               END ;
             END Else
             BEGIN
             CritEDT.DeviseAffichee:=CritEDT.DeviseSelect ;
             RDev.Code:=CritEDT.DeviseAffichee ;
             GetInfosDevise(RDev) ;
             CritEDT.Decimale:=RDev.Decimale ;
             CritEDT.Symbole:=RDev.Symbole ;
             CritEDT.DecimalePivot:=V_PGI.OkDecV ;
             CritEDT.SymbolePivot:=V_PGI.SymbolePivot ;
             If CritEdt.Monnaie=2 Then
               BEGIN
               CritEdt.DecimalePivot:=V_PGI.OkDecE ;
               If Not VH^.TenueEuro Then CritEdt.SymbolePivot:='' ;
               END ;
             END ;
          ChangeMask(Formateur,CritEDT.DecimalePivot,CritEDT.SymbolePivot);
          CritEDT.FormatMontantPivot:=Formateur.Masks.PositiveMask;
          CritEDT.FormatMontant:=Formateur.Masks.PositiveMask;
          END ;
   neBrouAna : BEGIN
               CritEDT.DeviseAffichee:=V_PGI.DevisePivot ;
               CritEDT.Decimale:=V_PGI.OkDecV ;
               CritEDT.Symbole:=V_PGI.SymbolePivot ;
               ChangeMask(Formateur,CritEDT.Decimale,CritEDT.Symbole) ;
               CritEDT.FormatMontant:=Formateur.Masks.PositiveMask ;
               AfficheEuroSurFiltreDevise ;
               END ;
   neJbq,nePoi,neRap : BEGIN
               if (CritEDT.DeviseSelect='') or (CritEDT.Monnaie=0) Or (CritEDT.DeviseSelect=V_PGI.DevisePivot) then
                  BEGIN
                  CritEDT.DeviseAffichee:=V_PGI.DevisePivot ;
                  CritEDT.Decimale:=V_PGI.OkDecV ;
                  CritEDT.Symbole:=V_PGI.SymbolePivot ;
                  END Else
                  BEGIN
                  CritEDT.DeviseAffichee:=CritEDT.DeviseSelect ;
                  RDev.Code:=CritEDT.DeviseAffichee ;
                  GetInfosDevise(RDev) ;
                  CritEDT.Decimale:=RDev.Decimale ;
                  CritEDT.Symbole:=RDev.Symbole ;
                  AfficheEuroSurFiltreDevise ;
                  END ;
               ChangeMask(Formateur,CritEDT.Decimale,CritEDT.Symbole);
               CritEDT.FormatMontant:=Formateur.Masks.PositiveMask;
               If FNatureEtat=neRap Then Tra.Visible:=FALSE ;
               END ;

   neBalBud,
   neGLBud,
   neJalBud,nePlanEdt : BEGIN
              if (CritEDT.Monnaie=0) then
                 BEGIN
                 CritEDT.DeviseAffichee:=V_PGI.DevisePivot ;
                 CritEDT.Decimale:=V_PGI.OkDecV ;
                 CritEDT.Symbole:=V_PGI.SymbolePivot ;
                 END Else
                 BEGIN
                 CritEDT.DeviseAffichee:=CritEDT.DeviseSelect ;
                 RDev.Code:=CritEDT.DeviseAffichee ;
                 GetInfosDevise(RDev) ;
                 CritEDT.Decimale:=RDev.Decimale ;
                 CritEDT.Symbole:=RDev.Symbole ;
                 AfficheEuroSurFiltreDevise ;
                 END ;
              ChangeMask(Formateur,CritEDT.Decimale,CritEDT.Symbole) ;
              CritEDT.FormatMontant:=Formateur.Masks.PositiveMask ;
              END ;
   neBroBud,
   neBroBudAna : BEGIN
               CritEDT.DeviseAffichee:=V_PGI.DevisePivot ;
               CritEDT.Decimale:=V_PGI.OkDecV ;
               CritEDT.Symbole:=V_PGI.SymbolePivot ;
               AfficheEuroSurFiltreDevise ;
               ChangeMask(Formateur,CritEDT.Decimale,CritEDT.Symbole) ;
               CritEDT.FormatMontant:=Formateur.Masks.PositiveMask ;
               END ;
   END ;
If CritEdt.Monnaie=2 Then CritEdt.AfficheSymbole:=FALSE ;
END ;

Procedure TFQR.Progressif(CalculMoi : Boolean ; D,C : Double);
BEGIN      { Gestion de la progression du Cumul du Solde Progressif }
if CalculMoi then BEGIN ProgressDebit:=ProgressDebit+D ; ProgressCredit:=ProgressCredit+C ; END
             Else BEGIN ProgressDebit:=0 ; ProgressCredit:=0 ; END ;
END ;


Function TFQR.AfficheMontant( Formatage, LeSymbole : String ; LeMontant : Double ; OkSymbole : Boolean ) : String ;
BEGIN
if OkSymbole then
   BEGIN
   if LeMontant=0 then AfficheMontant:=''
                  else AfficheMontant:=FormatFloat(Formatage+' '+LeSymbole,LeMontant) ;
   END else
   BEGIN
   if LeMontant=0 then AfficheMontant:=''
                  else AfficheMontant:=FormatFloat(Formatage,LeMontant) ;
   END ;
END ;


procedure addcritedtlist(LL : TStringList ; St : String) ;
begin
If St='' Then St:='-' ; LL.Add(St) ;
end ;

procedure TFQR.RemplirCritEdtList ;
(*
0 Généraux         1 Tiers      2 Sections         3 Journaux
4 Dates            5 Devise     6 Etablissement    7 Nature
8 Type d'écriture  9 Exercice
20 de              21 à         22 Tous            23 Toutes
*)
Var G:Record
      _Cpt1,_Cpt2,_SCpt1,_SCpt2,_Date1,_Date2,_Etab,_Dev,_NatureCpt,_Qualif : String ;
      End ;
    fnb : string ;
BEGIN
Fillchar(G,SizeOf(G),#0) ;
QRP.QRPrinter.CritEdtList.Clear ;

If Not OkZoomEdt Then
   BEGIN
   END ;
Case FNatureBase Of
  nbGen,NbGenT: G._Cpt1:=HMCrit.Mess[0] ; nbAux : G._Cpt1:=HMCrit.Mess[1] ;
  nbSec : G._Cpt1:=HMCrit.Mess[2] ; nbJal,nbJalAna : G._Cpt1:=HMCrit.Mess[3] ;
  END ;
G._Cpt1:=G._Cpt1+' '+HMCrit.Mess[20]+CritEdt.LCpt1 ;
If Not CritEdt.Joker Then G._Cpt2:=HMCrit.Mess[21]+' '+CritEdt.LCpt2 ;

G._Date1:=HMCrit.Mess[4]+' '+HMCrit.Mess[20]+DateToStr(CritEdt.DateDeb) ;
G._Date2:=HMCrit.Mess[21]+' '+DateToStr(CritEdt.DateFin) ;
If CritEdt.Composite Then
   BEGIN
   Case FModeSelect of
      msGenEcr : G._SCpt1:=HMCrit.Mess[1] ; msGenAna : G._Cpt2:=HMCrit.Mess[2] ;
      msAuxEcr : G._SCpt1:=HMCrit.Mess[0] ; msSecAna : G._Cpt2:=HMCrit.Mess[0] ;
      END ;
   G._SCpt1:=G._SCpt1+' '+HMCrit.Mess[20]+CritEdt.LSCpt1 ;
   If Not CritEdt.SJoker Then G._SCpt2:=HMCrit.Mess[21]+' '+CritEdt.LSCpt2 ;
   END ;
G._Dev:=HMCrit.Mess[5] ;
If CritEdt.DeviseSelect='' Then G._Dev:=G._Dev+' '+HMCrit.Mess[23]
                           Else G._Dev:=G._Dev+' '+RechDom('ttDevise',CritEdt.DeviseSelect,FALSE) ; ;
G._Etab:=HMCrit.Mess[6] ;
If CritEdt.Etab='' Then G._Etab:=G._Etab+' '+HMCrit.Mess[22]
                   Else G._Etab:=G._Etab+' '+RechDom('ttEtablissement',CritEdt.Etab,FALSE) ;

G._Qualif:=HMCrit.Mess[8] ;
If CritEdt.QualifPiece='' Then G._Qualif:=G._Qualif+' '+HMCrit.Mess[23]
                     Else G._Qualif:=G._Qualif+' '+RechDom('ttQualPieceCrit',CritEdt.QualifPiece,FALSE) ;

G._NatureCpt:=HMCrit.Mess[7]+' ' ;
Case FNatureBase Of
  nbGen,NbGenT: G._NatureCpt:=G._NatureCpt+HMCrit.Mess[0] ;
  nbAux : G._NatureCpt:=G._NatureCpt+HMCrit.Mess[1] ;
  nbSec : G._NatureCpt:=G._NatureCpt+HMCrit.Mess[2] ;
  nbJal,nbJalAna : G._NatureCpt:=G._NatureCpt+HMCrit.Mess[3] ;
  END ;
fnb:='' ;
Case FNatureBase of
   nbGen  :  fnb:='ttNatGene' ;
   nbGenT :  fnb:='ttNatGeneTIDTIC' ;
   nbAux  :  fnb:='ttNatTiersCpta';
   nbSec  :  fnb:='ttAxe' ;
   nbJal,nbJalAna  :  fnb:='ttNatJal' ;
   END ;
If CritEdt.NatureCpt='' Then G._NatureCpt:=G._NatureCpt+' '+HMCrit.Mess[22]
                        Else G._NatureCpt:=G._NatureCpt+' '+RechDom(fnb,CritEdt.NatureCpt,FALSE) ;
AddCritEdtList(QRP.QRPrinter.CritEdtList,G._Cpt1) ;
AddCritEdtList(QRP.QRPrinter.CritEdtList,G._Cpt2) ;
AddCritEdtList(QRP.QRPrinter.CritEdtList,G._SCpt1) ;
AddCritEdtList(QRP.QRPrinter.CritEdtList,G._SCpt2) ;
AddCritEdtList(QRP.QRPrinter.CritEdtList,G._Date1) ;
AddCritEdtList(QRP.QRPrinter.CritEdtList,G._Date2) ;
AddCritEdtList(QRP.QRPrinter.CritEdtList,G._Etab) ;
AddCritEdtList(QRP.QRPrinter.CritEdtList,G._Dev) ;
AddCritEdtList(QRP.QRPrinter.CritEdtList,G._NatureCpt) ;
AddCritEdtList(QRP.QRPrinter.CritEdtList,G._Qualif) ;
END ;

Function QuelTypeCpt(Valeur : String) : Integer ;
BEGIN
Result:=0 ;
If Valeur='EXO' Then Result:=0 Else If Valeur='FOU' Then Result:=1 Else
If Valeur='NSL' Then Result:=2 Else If Valeur='PER' Then Result:=3 ;
END ;

Function TFQR.MultiTri : String ;
Var k : Integer ;
//    SST : Array[1..6] Of String ;
    TC : tControl ;
    Tr : TCheckBox ;
    ZC : THValComboBox ;
BEGIN
Result:='' ;
For k:=1 To 6 Do
  BEGIN
  TC:=TControl(FindComponent('CTri'+IntToStr(k))) ;
  If TC<>NIL Then
    BEGIN
    tr:=TCheckBox(TC) ;
    If TR.Checked Then
      BEGIN
      TC:=TControl(FindComponent('Z_C'+IntToStr(k))) ;
      If TC<>NIL Then
        BEGIN
        ZC:=THvalComboBox(TC) ; If ZC.ItemIndex>=0 Then BEGIN Result:=ZC.Value ; Break ; END ;
        END ;
      END ;
    END ;
  END ;
END ;

procedure TFQR.RecupCritEdt ;
Var ST : String ;
    i : Integer ;
BEGIN
Fillchar(CritEDT,SizeOf(CritEDT),#0) ;
CritEDT.NatureEtat:=FNatureEtat ;
With CritEDT Do
  BEGIN
  OnCumul:=FALSE ; Composite:=FALSE ;
  If OnCum.Visible Then OnCumul:=OnCum.Checked ;
  If OnCumul Then VH^.OnCumEdt:=TRUE Else VH^.OnCumEdt:=FALSE ;
  Decimale:=V_PGI.OkDecV ; Symbole:=V_PGI.SymbolePivot ;
  DeviseSelect:='' ;
  If FDevises.Vide Then
     BEGIN
     if FDevises.ItemIndex<>0 then DeviseSelect:=FDevises.Value ;
     END Else
     BEGIN
     DeviseSelect:=FDevises.Value ;
     END ;
  DeviseAffichee:=V_PGI.DevisePivot ;
  Monnaie:=FMontant.ItemIndex ;
  if Monnaie=1 then DeviseAffichee:=DeviseSelect ;
  DevEnP:=False ;
  if (Monnaie=0) And (DeviseSelect<>'') And (DeviseSelect<>V_PGI.DevisePivot) then DevEnP:=TRUE ;
  Joker:=FJoker.Visible ; RappelCrit:=SansRappel ;
  Case FRappelCrit.State Of cbChecked : RappelCrit:=QueRappel ; cbGrayed : RappelCrit:=AvecRappel ; Else RappelCrit:=SansRappel ; END ;
  if Joker then
     BEGIN
     Cpt1:=FJoker.Text ; Cpt2:=FJoker.Text ; LCpt1:=Cpt1 ; LCpt2:=Cpt2 ;
     END else
     BEGIN
     Cpt1:=FCpte1.Text ; Cpt2:=FCpte2.Text ;
     PositionneFourchetteSt(FCpte1,FCpte2,CritEdt.LCpt1,CritEdt.LCpt2) ;
     END ;
  If CritEdt.NatureEtat=NeRap Then
     BEGIN
     Date2:=StrToDate(FDateCompta2.Text) ;
     QuelExoDate(Date2,Date2,MonoExo,Exo) ;
     DateDeb:=Exo.Deb ; Date1:=Exo.Deb ;
     DateFin:=DAte2 ;
     END Else
     BEGIN
     Date1:=StrToDate(FDateCompta1.Text)    ; Date2:=StrToDate(FDateCompta2.Text) ;
     DateDeb:=Date1 ; DateFin:=Date2 ;
     MonoExo:=FExercice.ItemIndex>0 ;
     if MonoExo then Exo.Code:=FExercice.Value ;
     QuelExoDate(Date1,Date2,MonoExo,Exo) ;
     END ;
  if FEtab.ItemIndex<>0 then Etab:=FEtab.Value ;
  If FNatureCpt.Vide Then
     BEGIN
     if FNatureCpt.ItemIndex<>0 then NatureCpt:=FNatureCpt.Value ;
     END Else NatureCpt:=FNatureCpt.Value ;
  AfficheSymbole:=FMonetaire.Checked ;
  St:='' ;
  if Trim(FExcep.Text)<>'' then
    Case FNatureBase of
       nbGen,nbGenT : St:=' And '+AnalyseCompte(FExcep.Text,fbGene,True,False) ;
       nbAux : St:=' And '+AnalyseCompte(FExcep.Text,fbAux,True,False) ;
       nbSec : St:=' And '+AnalyseCompte(FExcep.Text,AxeToFb(FNatureCpt.Value),True,False) ;
       nbJal : St:=' And '+AnalyseCompte(FExcep.Text,fbJal,True,False) ;
       nbBudGen : St:=' And '+AnalyseCompte(FExcep.Text,fbBudGen,True,False) ;
       nbBudSec : St:=' And '+AnalyseCompte(FExcep.Text,AxeToFbBud(FNatureCpt.Value),True,False) ;
       nbBudJal : St:=' And '+AnalyseCompte(FExcep.Text,fbBudJal,True,False) ;
       END ;
  Rupture:=rRien ;
  if Not FSansRupt.Checked and TabRuptures.TabVisible then
     BEGIN
     if FRuptures.Checked then Rupture:=rRuptures else
     if FTablesLibres.Checked then Rupture:=rLibres else
     if FPlansCo.Checked then Rupture:=rCorresp ;
     END ;
  RuptTabLibre:=(Rupture=rLibres);
   if RuptTabLibre then
      BEGIN
      if FLibTriPar.Text<>'' then
         BEGIN
         LibreTrie:=FLibTriPar.Text ; For i:=1 To Length(LibreTrie) Do If LibreTrie[i]=';' Then Inc(NbLibreTrie) ;
         END ;
      if FLibre1.text<>'' then LibreCodes1:=FLibre1.text ;
      if FLibre2.text<>'' then LibreCodes2:=FLibre2.text ;
      END ;
   If FNat01.Text<>'' then Cptlibre1:=FNat01.Text ;
   If FNat02.Text<>'' then Cptlibre11:=FNat02.Text ;
   If FNat11.Text<>'' then Cptlibre2:=FNat11.Text ;
   If FNat12.Text<>'' then Cptlibre21:=FNat12.Text ;
   If FNat21.Text<>'' then Cptlibre3:=FNat21.Text ;
   If FNat22.Text<>'' then Cptlibre31:=FNat22.Text ;
   If FNat31.Text<>'' then Cptlibre4:=FNat31.Text ;
   If FNat32.Text<>'' then Cptlibre41:=FNat32.Text ;
  (*
    Case FNatureBase of
       nbGen,nbGenT : if SqlCptInterdit('G_GENERAL', ST, FExcep) then ;
       nbAux : if SqlCptInterdit('T_AUXILIAIRE', ST, FExcep) then ;
       nbSec : if SqlCptInterdit('S_SECTION', ST, FExcep) then ;
       nbBud : if SqlCptInterdit('B_BUDGET', ST, FExcep) then ;
       nbJal : if SqlCptInterdit('J_JOURNAL', ST, FExcep) then ;
       END ;
  *)
  CritEdt.Qualifpiece:=FNatureEcr.Value ;
  CritEdt.ModeRevision:=cbUnchecked ;
  If Not OKZoomEdt Then If V_PGI.Controleur Then
     BEGIN
     If Etat In [etTvaHT,etTvaFac] Then CritEdt.ModeRevision:=cbUnchecked Else CritEdt.ModeRevision:=AvecRevision.State ;
     END ;
  SautPageRupt:=FSautPageRupt.Checked ;
  SautPageTRI:=FSautPageTri.Checked ; 
  Case CritEDT.NatureEtat Of
    neBal : BEGIN
            if Not MonoExo then Bal.TypCpt:=3 ; // Périodes
            IntervalleDateBALGL1(MonoExo,Exo,Date1,Bal.Date11,Bal.Date21) ;
            (*
            if V_PGI.EnCours.Deb=Date1 then
               BEGIN
               Bal.Date11:=StrToDate(FDateCompta1.Text) ; Bal.Date21:=Bal.Date11 ;
               END else
               BEGIN
               Date1:=V_PGI.EnCours.Deb ; Bal.Date11:=StrToDate(FDateCompta1.Text)-1 ; Bal.Date21:=BAL.Date11+1 ;
               END ;
            *)
            Bal.TypCpt:=QuelTypeCpt(FSelectCpte.Value) ;
            Bal.Sauf:=Trim(FExcep.Text) ;
          //Simon
            if St<>'' then Bal.SQLSauf:=St ;
            With Bal.FormatPrint Do
              BEGIN
              PrSepMontant:=FTrait.Checked ;
              PrSepCompte:=FLigneCpt.Checked ;
              Report.OkAff:=FReport.Checked ;
              END ;
            Bal.ZoomBalRub:=FALSE ; Bal.QuelAN:=AvecAN ;
            Bal.STTS:=MultiTri ;
            END ;
    neGL  : BEGIN
            if Not MonoExo then GL.TypCpt:=3 ; // Périodes
            IntervalleDateBALGL1(MonoExo,Exo,Date1,GL.Date11,GL.Date21) ;
            GL.CodeRupt1:='' ; GL.CodeRupt2:='' ;
            (*
            if V_PGI.EnCours.Deb=Date1 then
               BEGIN
               GL.Date11:=StrToDate(FDateCompta1.Text) ; GL.Date21:=GL.Date11 ;
               END else
               BEGIN
               Date1:=V_PGI.EnCours.Deb ; GL.Date11:=DateDeb-1 ; GL.Date21:=GL.Date11+1 ;
               END ;
            *)
            GL.TypCpt:=QuelTypeCpt(FSelectCpte.Value) ;
            GL.Sauf:=Trim(FExcep.Text) ;
            if St<>'' then GL.SQLSauf:=St ;
            With GL.FormatPrint Do
               BEGIN
               PrSepMontant:=FTrait.Checked ;
               PrSepCompte[1]:=FLigneCpt.Checked ;
               Report.OkAff:=FReport.Checked ;
               END ;
            GL.QuelAN:=AvecAN ;
            END ;
    neJU  : BEGIN
            With JU.FormatPrint Do
                 BEGIN
                 PrSepMontant:=FTrait.Checked ;
                 PrSepCompte[1]:=FLigneCpt.Checked ;
                 Report.OkAff:=FReport.Checked ;
                 END ;
            END ;
    neJal : BEGIN
            With Jal.FormatPrint Do
                 BEGIN
                 PrSepMontant:=FTrait.Checked ;
                 PrSepCompte[1]:=FLigneCpt.Checked ;
                 Report.OkAff:=FReport.Checked ;
                 END ;
             END ;
    neJalR : BEGIN
             With JalR.FormatPrint Do
                  BEGIN
                  PrSepMontant:=FTrait.Checked ;
                  PrSepCompte[1]:=FLigneCpt.Checked ;
                  Report.OkAff:=FReport.Checked ;
                  END ;
             END ;
    neEch : BEGIN
//            DansExo ???
            With Ech.FormatPrint Do
                 BEGIN
                 PrSepMontant:=FTrait.Checked ;
                 PrSepCompte[1]:=FLigneCpt.Checked  ;
                 Report.OkAff:=FReport.Checked ;
                 END ;
            END ;
    neGLV : BEGIN
            With GlV.FormatPrint Do
              BEGIN
              PrSepMontant:=FTrait.Checked ;
              PrSepCompte[1]:=FLigneCpt.Checked ;
              Report.OkAff:=FReport.Checked ;
              END ;
            END ;
    neCum : BEGIN
            if Not MonoExo then Cum.TypCpt:=3 ; // Périodes
            IntervalleDateBALGL1(MonoExo,Exo,Date1,Cum.Date11,Cum.Date21) ;
            Cum.CalculAN:=DAte1=Cum.Date11 ;
            if (Monnaie>0) And (DeviseSelect<>'') And (DeviseSelect<>V_PGI.DevisePivot) then
               BEGIN
               Cum.DeuxMontant:=TRUE ;
               DEVPourExist:=DeviseSelect ;
               END ;
            (*
            if V_PGI.EnCours.Deb=Date1 then
               BEGIN
               Bal.Date11:=StrToDate(FDateCompta1.Text) ; Bal.Date21:=Bal.Date11 ;
               END else
               BEGIN
               Date1:=V_PGI.EnCours.Deb ; Bal.Date11:=StrToDate(FDateCompta1.Text)-1 ; Bal.Date21:=BAL.Date11+1 ;
               END ;
            *)
            Cum.TypCpt:=QuelTypeCpt(FSelectCpte.Value) ;
            Cum.Sauf:=Trim(FExcep.Text) ;
          //Simon
            if St<>'' then Cum.SQLSauf:=St ;
            With Cum.FormatPrint Do
              BEGIN
              PrSepMontant:=FTrait.Checked ;
              PrSepCompte:=FLigneCpt.Checked ;
              Report.OkAff:=FReport.Checked ;
              END ;
            END ;
    neBrouAna  : BEGIN
                 With BrouAna.FormatPrint Do
                    BEGIN
                    PrSepMontant:=FTrait.Checked ;
                    PrSepCompte[1]:=FLigneCpt.Checked ;
                    Report.OkAff:=FReport.Checked ;
                    END ;
                 END ;
    neJbq : BEGIN
            With JBq.FormatPrint Do
              BEGIN
              PrSepMontant:=FTrait.Checked ;
//              PrSepCompte[1]:=FLigneBanque.Checked ;
              Report.OkAff:=FReport.Checked ;
              END ;
            END ;
    nePoi : BEGIN
            With Poi.FormatPrint Do
              BEGIN
              PrSepMontant:=FTrait.Checked ;
//              PrSepCompte[1]:=FLigneBanque.Checked ;
              Report.OkAff:=FReport.Checked ;
              END ;
            END ;
    neRap : BEGIN
            With Rap.FormatPrint Do
              BEGIN
              PrSepMontant:=FTrait.Checked ;
//              PrSepCompte[1]:=FLigneBanque.Checked ;
              Report.OkAff:=FReport.Checked ;
              END ;
            END ;
    neBalBud : BEGIN
               BalBud.MvtSur:=FSelectCpte.Value ;
               With BalBud.FormatPrint Do
                 BEGIN
                 PrSepMontant:=FTrait.Checked ;
                 PrSepCompte[1]:=FLigneCpt.Checked ;
                 Report.OkAff:=FReport.Checked ;
                 END ;
            END ;
    neBroBud  : BEGIN
                 With BroBud.FormatPrint Do
                    BEGIN
                    PrSepMontant:=FTrait.Checked ;
                    PrSepCompte[1]:=FLigneCpt.Checked ;
                    Report.OkAff:=FReport.Checked ;
                    END ;
                 END ;
    neBroBudAna  : BEGIN
                 With BroBudAna.FormatPrint Do
                    BEGIN
                    PrSepMontant:=FTrait.Checked ;
                    PrSepCompte[1]:=FLigneCpt.Checked ;
                    Report.OkAff:=FReport.Checked ;
                    END ;
                 END ;
    neBalLib : BEGIN
               IntervalleDateBALGL1(MonoExo,Exo,Date1,BalLib.Date11,BalLib.Date21) ;
               Bal.Sauf:=Trim(FExcep.Text) ;
             // Simon
               if St<>'' then BalLib.SQLSauf:=St ;
                 With BalLib.FormatPrint Do
                   BEGIN
                   PrSepMontant:=FTrait.Checked ;
                   PrSepCompte:=FLigneCpt.Checked ;
                   Report.OkAff:=FReport.Checked ;
                   END ;
               END ;
    nePlanEdt : BEGIN
                With PlanEdt.FormatPrint Do
                  BEGIN
                  PrSepMontant:=FTrait.Checked ;
                  PrSepCompte[1]:=FLigneCpt.Checked ;
                  Report.OkAff:=FReport.Checked ;
                  END ;
            END ;
    END ;
 END ;
CritEdt.GA:=(NomListeGA<>'') And Avance.Checked ;
If PAvances.TabVisible Then CritEdt.FiltreSup:=FaitSQLCompQR ;
END ;

Function TFQR.CritOk : Boolean ;
BEGIN
Result:=TRUE ;
EntetePage.Color:=clWhite ;
EnteteAutrePage.Color:=clWhite ;
TopReport.Color:=clWhite ;
BDetail.Color:=clWhite ;
BFinEtat.Color:=clWhite ;
BottomReport.Color:=clWhite ;
NumErreurCrit:=1 ;
//Rony
//Simon
If Not OkZoomEdt Then BEGIN RecupCritEdt ; RemplirCritEdtList ; END ;
Case FNatureEtat of
 neBAl : BEGIN
         Result:=CtrlPerExo(CritEDT.DateDeb,CritEDT.DateFin) ;
         If Not OkZoomEdt Then
            BEGIN
            If Result Then
               BEGIN
               Result:=CritEdt.MonoExo ;
               If Not Result Then NumErreurCrit:=101 ;
               END ;
            If Result Then
               BEGIN
               Result:=CtrlMouvementeSur(CritEdt.Bal.TypCpt,CritEdt.QualifPiece) ;
               If Not Result Then NumErreurCrit:=6 ;
               END ;
            END ;
         If Result Then
            BEGIN
            Dev:=CritEDT.DeviseSelect<>'' ;
            Etab:=CritEDT.Etab<>'' ;
            Exo:=CritEDT.MonoExo And (CritEDT.Exo.Code<>'') ;
            END ;
         END ;
  neGL : BEGIN
         Result:=CtrlPerExo(CritEDT.DateDeb,CritEDT.DateFin) ;
         If Not OkZoomEdt Then
            BEGIN
            If Result Then
               BEGIN
               Result:=CritEdt.MonoExo ;
               If Not Result Then NumErreurCrit:=101 ;
               END ;
            If Result Then
               BEGIN
               Result:=CtrlMouvementeSur(CritEdt.GL.TypCpt,CritEdt.QualifPiece) ;
               If Not Result Then NumErreurCrit:=6 ;
               END ;
            END ;
         if Result then
            BEGIN
            Dev:=CritEDT.DeviseSelect<>'' ;
            Etab:=CritEDT.Etab<>'' ;
            Exo:=CritEDT.MonoExo And (CritEDT.Exo.Code<>'') ;
            END ;
         END ;
  neJU : BEGIN
         Result:=IsValidDate(DateToStr(CritEDT.DateDeb)) ;
         END ;
  neJal,
  NeJalR: BEGIN
          Result:=CtrlPerExo(CritEDT.DateDeb,CritEDT.DateFin) ;
          If Result Then
             BEGIN
             Result:=CritEdt.MonoExo ;
             If Not Result Then NumErreurCrit:=101 ;
             END ;
          END ;
  neEch: BEGIN
         Result:=CtrlPerExo(CritEDT.DateDeb,CritEDT.DateFin) ;
         END ;
  neGLV: BEGIN
         Result:=IsValidDate(DateToStr(CritEDT.DateDeb)) ;
         END ;
 neCum : BEGIN
         Result:=CtrlPerExo(CritEDT.DateDeb,CritEDT.DateFin) ;
         If Result Then
            BEGIN
            Result:=CritEdt.MonoExo ;
            If Not Result Then NumErreurCrit:=102 ;
            END ;
         If Result Then
            BEGIN
            Result:=CtrlMouvementeSur(CritEdt.Cum.TypCpt,CritEdt.QualifPiece) ;
            If Not Result Then NumErreurCrit:=6 ;
            END ;
         If Result Then
            BEGIN
            Dev:=CritEDT.DeviseSelect<>'' ;
            Etab:=CritEDT.Etab<>'' ;
            Exo:=CritEDT.MonoExo And (CritEDT.Exo.Code<>'') ;
            END ;
         END ;
  neBrouAna : BEGIN
              Result:=CtrlPerExo(CritEDT.DateDeb,CritEDT.DateFin) ;
              If Result Then
                 BEGIN
                 Result:=CritEdt.MonoExo ;
                 If Not Result Then NumErreurCrit:=101 ;
                 END ;
              if Result then
                 BEGIN
                 Dev:=FALSE ;
                 Etab:=CritEDT.Etab<>'' ;
                 Exo:=CritEDT.MonoExo And (CritEDT.Exo.Code<>'') ;
                 END ;
              END ;
  neJbq : BEGIN
          Result:=CtrlPerExo(CritEDT.DateDeb,CritEDT.DateFin) ;
          END ;
  neRap : BEGIN
          Dev:=CritEDT.DeviseSelect<>'' ;
          Etab:=FALSE ;
          Exo:=CritEDT.Exo.Code<>'' ;
          END ;
  neBalLib : BEGIN
          Dev:=CritEDT.DeviseSelect<>'' ;
          Etab:=CritEDT.Etab<>'' ;
          Exo:=CritEDT.Exo.Code<>'' ;
          END ;
   END ;
If Result And (CritEdt.Rupture=rLibres) And (Not FSansRupt.Checked) And (Trim(FLibTriPar.Text)='') Then BEGIN Result:=FALSE ; NumerreurCrit:=2005 ; END ;
END ;

procedure TFQR.GenereSQL  ;
BEGIN
//sdsds
END ;

procedure TFQR.FinirPrint ;
BEGIN
If OkMajEdt Then If QRP.QRPrinter.Canceled Then OkMajEdt:=FALSE ;
//sdsds
END ;

procedure TFQR.BValiderClick(Sender: TObject);
  Label 0 ;
begin
SourisSablier ;
If CritOK then
   BEGIN
   If Not OkZoomEdt Then If CritEdt.RappelCrit=QueRappel Then BEGIN FinirPrint ; Goto 0 ; END ;
   EnableControls(Self,False) ;
   InitDivers ;
   GenereSQL ;
   RenseigneCritere ;
   ChoixEdition ;
//   BTitre.Height:=TTitre.Height ;
   (*
   if Q.Eof then BEGIN OkMajEdt:=FALSE ; MsgRien.Execute(0,Caption,'') ; END
            else BEGIN if Apercu.Checked then QRP.Preview else QRP.Print ; END ;
   *)
   // GCO - 07/04/2002
   if Q.Eof then
   BEGIN
     if not CritEdtChaine.Utiliser then
     BEGIN
       OkMajEdt:=FALSE ;
       MsgRien.Execute(0,Caption,'') ;
     END;
   END
   else
   BEGIN
//     if Apercu.Checked then QRP.Preview else QRP.Print ;
     if CritEdtChaine.Utiliser then
     begin
       if CritEdtchaine.AuFormatPDF then
       begin
         if CritEdtChaine.MultiPdf then
           V_PGI.QRPDFQueue := ExtractFilePath(CritEdtChaine.NomPDF) + '\' + V_PGI.NoDossier + '-' +  ExtractFileName(CritEdtChaine.NomPDF)
         else
           V_PGI.QRPDFQueue := CritEdtChaine.NomPDF;

         V_PGI.QRPDFMerge := '' ;
         if (V_PGI.QRPDFQueue <> '') and FileExists(V_PGI.QRPDFQueue) then
           V_PGI.QRPDFMerge := V_PGI.QRPDFQueue ;
         QRP.Preview;
       end
       else
         QRP.Print;

        V_PGI.QRPDFQueue := '' ;
        V_PGI.QRPDFMerge := '' ;

       // Test pour savoir si on a abandonne l' état
       if QRP.QRPrinter.Canceled then
       begin
         if PgiAsk('Voulez vous annuler les impressions en cours ?', 'Etats chaînés') = MrYes then
           VH^.StopEdition := True;
       end;
     end
     else
     begin
       if Apercu.Checked then QRP.Preview else QRP.Print ;
     end;
   END ;
   // FIN GCO
   Q.Close ;
   FinirPrint ;
   EnableControls(Self,True) ;
   END else
   BEGIN
   If (NumErreurCrit>0) And (NumErreurCrit<10000) Then
      BEGIN
      If NumErreurCrit>2000 Then MsgQR.Execute(NumErreurCrit Mod 100,Caption,'') else
         If NumErreurCrit>100 Then MsgRien.Execute(NumErreurCrit Mod 100,Caption,'') else MsgRien.Execute(NumErreurCrit,Caption,'') ;
      END ;
   END ;
0:
If Not OkZoomEdt Then
   BEGIN
   If Apercu.Checked Then
      BEGIN
      If CritEdt.RappelCrit=QueRappel Then
         BEGIN
         PrintPageDeGarde(Pages,FCouleur.Checked,Apercu.Checked,TRUE,Caption,0) ;
         END ;
      END Else
      BEGIN
      If CritEdt.RappelCrit<>SansRappel Then
         BEGIN
         PrintPageDeGarde(Pages,FCouleur.Checked,Apercu.Checked,TRUE,Caption,0) ;
         END ;
      END ;
   END ;
SourisNormale ;
end;

procedure TFQR.GeneOuTiers ;
begin
//If FComposite And (FNatureBase=nbGen) And (FModeSelect=msGenEcr) Then
if FComposite and ((FNatureBase=nbGen) And (FModeSelect=msGenEcr)) then
   BEGIN
   FNatureCpt.DataType:='ttNatGeneColl' ; FCpte1.ZoomTable:=tzGCollectif ; FCpte2.ZoomTable:=FCpte1.ZoomTable ;
   FPlanRuptures.DataType:='ttRuptGene' ;
   END Else
if FComposite and ((FNatureBase=nbBudGen) And (FModeSelect=msGenBAna)) then
   BEGIN
   FNatureCpt.DataType:='ttAxe' ; FCpte1.ZoomTable:=tzBudgen ; FCpte2.ZoomTable:=FCpte1.ZoomTable ;
   END Else
if FComposite and ((FNatureBase=nbBudSec) And (FModeSelect=msSecBAna)) then
   BEGIN
   FNatureCpt.DataType:='ttAxe' ; FCpte1.ZoomTable:=tzBudSec1 ; FCpte2.ZoomTable:=FCpte1.ZoomTable ;
   END Else
   Case FNatureBase of
      nbGen  :  BEGIN
                If (FNatureEtat=neJbq) Or (FNatureEtat=neRap) Then
                   BEGIN
                   FCpte1.ZoomTable:=tzGBanque ; FCpte2.ZoomTable:=FCpte1.ZoomTable ;
                   END Else
                If (FNatureEtat=nePoi) Then
                   BEGIN
                   FCpte1.ZoomTable:=tzGPointable ; FCpte2.ZoomTable:=FCpte1.ZoomTable ;
                   END Else
                   BEGIN
                   FNatureCpt.DataType:='ttNatGene' ; FCpte1.ZoomTable:=tzGeneral ; FCpte2.ZoomTable:=FCpte1.ZoomTable ;
                   FPlanRuptures.DataType:='ttRuptGene' ;
                   FNat01.ZoomTable:=tzNatEcrE0 ; FNat02.ZoomTable:=FNat01.ZoomTable ;
                   FNat11.ZoomTable:=tzNatEcrE1 ; FNat12.ZoomTable:=FNat11.ZoomTable ;
                   FNat21.ZoomTable:=tzNatEcrE2 ; FNat22.ZoomTable:=FNat21.ZoomTable ;
                   FNat31.ZoomTable:=tzNatEcrE3 ; FNat32.ZoomTable:=FNat31.ZoomTable ;
                   END ;
                END ;
      nbGenT   : BEGIN FNatureCpt.DataType:='ttNatGeneTIDTIC' ; FCpte1.ZoomTable:=tzGTIDTIC ; FCpte2.ZoomTable:=FCpte1.ZoomTable ; END ;
      nbAux    : BEGIN
                 {JP 23/08/04 : est-ce bien utile ?
                 If (Etat=etJuSold) Then // Correction bug sur rappel de filtre dans justif : Zone nature compte à vide
                 BEGIN
                   If QRLoading Then
                   BEGIN
                     If (uppercase(FNatureCpt.DataType)<>'TTNATTIERSCPTA') Then FNatureCpt.DataType:='ttNatTiersCpta';
                   END
                   Else FNatureCpt.DataType:='ttNatTiersCpta';
                 END
                 Else}
                 FNatureCpt.DataType:='ttNatTiersCpta';
                 FCpte1.ZoomTable:=tzTiers   ; FCpte2.ZoomTable:=FCpte1.ZoomTable ; FPlanRuptures.DataType:='ttRuptTiers' ;
                 FNat01.ZoomTable:=tzNatEcrE0 ; FNat02.ZoomTable:=FNat01.ZoomTable ;
                 FNat11.ZoomTable:=tzNatEcrE1 ; FNat12.ZoomTable:=FNat11.ZoomTable ;
                 FNat21.ZoomTable:=tzNatEcrE2 ; FNat22.ZoomTable:=FNat21.ZoomTable ;
                 FNat31.ZoomTable:=tzNatEcrE3 ; FNat32.ZoomTable:=FNat31.ZoomTable ;
                 END ;
      nbSec    : BEGIN
                 FNatureCpt.DataType:='ttAxe' ; FCpte1.ZoomTable:=tzSection   ; FCpte2.ZoomTable:=FCpte1.ZoomTable ; FPlanRuptures.DataType:='ttRuptSect1' ;
                 Fcpte1.SynJoker:=TRUE ; Fcpte2.SynJoker:=TRUE ;
                 FNat01.ZoomTable:=tzNatEcrA0 ; FNat02.ZoomTable:=FNat01.ZoomTable ;
                 FNat11.ZoomTable:=tzNatEcrA1 ; FNat12.ZoomTable:=FNat11.ZoomTable ;
                 FNat21.ZoomTable:=tzNatEcrA2 ; FNat22.ZoomTable:=FNat21.ZoomTable ;
                 FNat31.ZoomTable:=tzNatEcrA3 ; FNat32.ZoomTable:=FNat31.ZoomTable ;
                 END ;
      nbJal    : BEGIN FNatureCpt.DataType:='ttAxe' ; FCpte1.ZoomTable:=tzJournal   ; FCpte2.ZoomTable:=FCpte1.ZoomTable ; END ;
      nbJalAna : BEGIN FNatureCpt.DataType:='ttAxe' ; FCpte1.ZoomTable:=tzJAna   ; FCpte2.ZoomTable:=FCpte1.ZoomTable ; END ;
      nbBudGen : BEGIN FCpte1.ZoomTable:=tzBudgen ; FCpte2.ZoomTable:=FCpte1.ZoomTable ; END ;
      nbBudSec : BEGIN FNatureCpt.DataType:='ttAxe' ; FCpte1.ZoomTable:=tzBudSec1   ; FCpte2.ZoomTable:=FCpte1.ZoomTable ; END ;
      nbBudJal : BEGIN FNatureCpt.DataType:='ttAxe' ; FCpte1.ZoomTable:=tzBudJal   ; FCpte2.ZoomTable:=FCpte1.ZoomTable ; END ;
      nbLibecr : BEGIN FNatureCpt.DataType:='' ; FCpte1.ZoomTable:=tzNatEcrE0   ; FCpte2.ZoomTable:=FCpte1.ZoomTable ; END ;
      END ;
end ;

Procedure TFQR.InitVarInterne ;
BEGIN
RuptAvantRappelFiltre:=rRien ;
If FRuptures.Checked Then RuptAvantRappelFiltre:=rRuptures ;
If FPlansCo.Checked Then RuptAvantRappelFiltre:=rCorresp ;
If FTablesLibres.Checked Then RuptAvantRappelFiltre:=rLibres ;
END ;

Procedure AddZC(ZC : THValCombobox ; TDE : IFieldCom ; OnLib : Boolean) ;
BEGIN
If OnLib Then ZC.Items.Add(TDE.Libelle) Else ZC.Items.Add(TDE.Name) ;
ZC.Values.Add(TDE.Name) ;
END ;

procedure TFQR.ChargeAvances(OnLib : Boolean) ;
Var i,j,k : Integer ;
    SST : Array[1..6] Of String ;
    TC : tControl ;
    ZC : THValComboBox ;
    St : String ;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;
    Field     : IFieldCOM ;
BEGIN
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  St:=Getpref ;
  If St='' Then Exit ;
  Table := Mcd.GetTable(Mcd.PrefixeToTable(St));
  FieldList := Table.Fields;
  FieldList.Reset();

  For k:=1 To 6 Do
  BEGIN
    SST[k]:='' ;
    TC:=TControl(FindComponent('Z_C'+IntToStr(k))) ;
    If TC<>NIL Then
    BEGIN
      ZC:=THValComboBox(TC) ; If ZC.ItemIndex>=0 Then SST[k]:=ZC.Value ; ZC.Clear ;
    END ;
  END ;
  While FieldList.MoveNext do
  begin
    if (FieldList.Current as IFieldCOM).name <>'' then
    BEGIN
      For k:=1 To 6 Do
      BEGIN
        TC:=TControl(FindComponent('Z_C'+IntToStr(k))) ;
        If TC<>NIL Then
        BEGIN
          ZC:=THValComboBox(TC) ;
          AddZC(ZC,FieldList.Current as IFieldCom,OnLib) ;
        END ;
      END ;
    END ;
  end;
  For k:=1 To 6 Do
  BEGIN
    TC:=TControl(FindComponent('Z_C'+IntToStr(k))) ;
    If TC<>NIL Then
    BEGIN
      ZC:=THValComboBox(TC) ;
      If SST[k]<>'' Then ZC.Value:=SST[k] ;
    END ;
  END ;
END ;

procedure TFQR.FormShow(Sender: TObject);
begin
ChargeAvances(FALSE) ;
OkChargeCritMemoire:=TRUE ;
QRLoading:=FALSE ; QRLoadingExo:=FALSE ; QRChangeDev:=FALSE ; BTitreOk:=TRUE ;
QRLoadingCptAssocie:=FALSE ;
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;

QRP.ReportTitle:=Caption ;
Pages.ActivePage:=Standards ;

FCouleur.Checked:=V_PGI.QRCouleur ;
FSansRuptClick(Nil) ;
//Simon
GeneOuTiers ;

FCpte1.Text:='' ; FCpte2.Text:='' ;
FSelectCpte.DataType:='ttSelectionCompte' ;
If FComposite Then FSelectCpte.DataType:='ttSelectionCompte2' ;
FNatureCpt.ItemIndex:=0 ;

{JP 28/06/06 : FQ 16149 : refonte de la gestion des établissements
FEtab.ItemIndex:=0 ;
PositionneEtabUser(FEtab) ;}
GereEtablissement;

FDevises.ItemIndex:=0 ;
FNatureEcr.ItemIndex:=0 ; FMontant.ItemIndex:=0 ; FSelectCpte.ItemIndex:=0 ;
if VH^.CPExoRef.Code<>'' then FExercice.Value:=VH^.CPExoRef.Code else FExercice.Value:=VH^.Entree.Code ;
FTrait.Checked:=TRUE ;
//Simon
Case FNatureEtat of
 neBal : BEGIN
         FLigneCpt.Checked:=FALSE ;
         InitFormatBal(Self,PremTabColBal) ;
         END ;
  neGL : BEGIN
         FLigneCpt.Checked:=TRUE ;
         InitFormatEdt(Self,PremTabColEdt,CritEDT.GL.FormatPrint.TabColl) ;
         END ;
  neJU : BEGIN
         InitFormatEdt(Self,PremTabColEdt,CritEDT.JU.FormatPrint.TabColl) ;
         FLigneCpt.Checked:=TRUE ;
         END ;
  neJal : BEGIN
          InitFormatEdt(Self,PremTabColEdt,CritEDT.Jal.FormatPrint.TabColl) ;
          FLigneCpt.Checked:=TRUE ;
          END ;
  neJalR: BEGIN
          TabSup.TabVisible:=FALSE ;
          InitFormatEdt(Self,PremTabColEdt,CritEDT.JalR.FormatPrint.TabColl) ;
          FLigneCpt.Checked:=TRUE ;
          END ;
  neEch: BEGIN
         InitFormatEdt(Self,PremTabColEdt,CritEDT.Ech.FormatPrint.TabColl) ;
         FLigneCpt.Checked:=TRUE ;
         END ;
  neGLV: BEGIN
         InitFormatEdt(Self,PremTabColEdt,CritEDT.GlV.FormatPrint.TabColl) ;
         FLigneCpt.Checked:=False ;
         END ;
  neCum: BEGIN
         FLigneCpt.Checked:=TRUE ;
         InitFormatBal(Self,PremTabColBal) ;
         TabSup.TabVisible:=FALSE ;
         If FNatureBase=nbSec Then FNAtureCpt.Vide:=FALSE Else FNAtureCpt.Vide:=TRUE ;
         FNatureCpt.Reload ; FNatureCpt.ItemIndex:=0 ;
         END ;
  neBrouAna : BEGIN
              FLigneCpt.Checked:=TRUE ;
              InitFormatEdt(Self,PremTabColEdt,CritEDT.BrouAna.FormatPrint.TabColl) ;
              END ;
  neJbq: BEGIN
         TabSup.TabVisible:=FALSE ;
         InitFormatEdt(Self,PremTabColEdt,CritEDT.Jbq.FormatPrint.TabColl) ;
         FLigneCpt.Checked:=TRUE ;
         FTrait.Checked:=TRUE ;
         FExercice.Vide:=TRUE ;
         FExercice.Reload ;
         FDevises.Vide:=FALSE ;
         FDevises.Reload ;
         FExercice.Value:=VH^.Entree.Code ; // rony 28/05/97
         FDevises.Value:=V_PGI.DevisePivot;
         END ;
  nePoi: BEGIN
         TabSup.TabVisible:=FALSE ;
         Avances.TabVisible:=FALSE ;
         InitFormatEdt(Self,PremTabColEdt,CritEDT.Poi.FormatPrint.TabColl) ;
         FLigneCpt.Checked:=TRUE ;
         FTrait.Checked:=TRUE ;
         FDevises.Vide:=FALSE ;
         FDevises.Reload  ;
         FDevises.Value:=V_PGI.DevisePivot;
         END ;
  neRap: BEGIN
         TabSup.TabVisible:=FALSE ;
         Avances.TabVisible:=FALSE ;
         InitFormatEdt(Self,PremTabColEdt,CritEDT.Rap.FormatPrint.TabColl) ;
         FLigneCpt.Checked:=TRUE ;
         FTrait.Checked:=TRUE ;
// GP le 15/06/99         FDevises.Vide:=FALSE ;
         FDevises.Reload  ;
         FDevises.Value:=V_PGI.DevisePivot;
         FDevises.Visible:=FALSE ; HLabel8.Visible:=FALSE ;
         END ;
 neBalBud : BEGIN
         InitFormatEdt(Self,PremTabColEdt,CritEDT.BalBud.FormatPrint.TabColl) ;
         FLigneCpt.Checked:=False ;
         FRuptures.Enabled:=FALSE ;
         FPlansCo.Enabled:=FALSE ;
         FTablesLibres.Checked:=TRUE ;
         FRuptures.Checked:=FALSE ;
         FSelectCpte.DataType:='ttSelectionCptBudget' ;
         FSelectCpte.Value:='TOO' ;
         FNatureEcr.Value:='N' ;
         FReport.Checked:=FALSE ; FReport.Enabled:=FALSE ;
         END ;
 neBroBud : BEGIN
         InitFormatEdt(Self,PremTabColEdt,CritEDT.BroBud.FormatPrint.TabColl) ;
         FLigneCpt.Checked:=False ;
         END ;
 neBroBudAna: BEGIN
         InitFormatEdt(Self,PremTabColEdt,CritEDT.BroBudAna.FormatPrint.TabColl) ;
         FLigneCpt.Checked:=False ;
         END ;
 neBalLib : BEGIN
         FLigneCpt.Checked:=FALSE ;
         InitFormatBal(Self,PremTabColBal) ;
         END ;
 nePlanEdt : BEGIN
         InitFormatEdt(Self,PremTabColEdt,CritEDT.PlanEdt.FormatPrint.TabColl) ;
         FLigneCpt.Checked:=False ;
         FRuptures.Enabled:=FALSE ;
         FPlansCo.Enabled:=FALSE ;
         FTablesLibres.Checked:=TRUE ;
         FRuptures.Checked:=FALSE ;
         FSelectCpte.DataType:='ttSelectionCptBudget' ;
         FSelectCpte.Value:='TOO' ;
         FNatureEcr.Value:='N' ;
         FReport.Checked:=FALSE ; FReport.Enabled:=FALSE ;
         END ;
   END ;
InitEdit(EntetePage,QRP) ;
AvecRevision.Visible:=V_PGI.Controleur And ((Etat in [etTvaHT,etTvaFac])=FALSE) And
                      (EstSerie(S3)=FALSE) And (EstSerie(S5)=FALSE) ;
If AvecRevision.Visible Then
   BEGIN
   AvecRevision.Checked:=TRUE ; AvecRevision.State:=cbUnchecked ;
   END Else
   BEGIN
   AvecRevision.Checked:=FALSE ; AvecRevision.State:=cbUnchecked ;
   END ;
QRLoading:=TRUE ; QRLoadingExo:=TRUE ; QRLoadingCptAssocie:=TRUE ;
//Rony 3/06/97
//SauveCritMemoire (QRCritMemoireDefaut,Pages);
SauveCritMemoire (QRCritMemoire,Pages);
InitVarInterne ;
//ChargeFiltre(FNomFiltre,FFiltres,Pages) ;
// GC*
(*
if CritEdtChaine.Utiliser then
  ChargeFiltre(FNomFiltre,FFiltres,Pages,True)
else
  ChargeFiltre(FNomFiltre,FFiltres,Pages,False);
// FIN GC*
*)
  {JP 23/08/04 : FQ 14242 : Nouvelle gestion des filtres}
  if FListeByUser <> nil then
    FListeByUser.LoadDB(FNomFiltre);
QRLoading:=FALSE ; QRLoadingExo:=FALSE ; QRLoadingCptAssocie:=FALSE ;
OKFiltreDefaut:=FFiltres.Text<>'' ;


// Merci Vincent :
//If (Not EstSerie(S3)) and (not EstSerie(S5)=FALSE) Then FTablesLibres.Visible:=FALSE ;
{$IFDEF CCS3}
FPlansCo.Visible:=FALSE ;
{$ENDIF}
If EstSpecif('51187') Then
  BEGIN
  If OnCum.Visible Then OnCum.Checked:=VH^.UseTC ;
  END ;

// BPY le 14/10/2004 => Fiche n° 14745 : repositionnement des bouton :)
    BFiltre.Left := 6;
    FFiltres.Left := BFiltre.Left + BFiltre.Width + 4;

    BAide.Left := HPB.Width - 6 - BAide.Width;
    BFerme.Left := BAide.Left - 4 - BFerme.Width;
    BValider.Left := BFerme.Left - 4 - BValider.Width;
// Fin BPY

// GCO - 06/12/2004 - FQ 13799
// Ne marche pas due à la modification du V_PGI.AGLDesigning ( ligne 3839 )
{$IFDEF CCSTD}
  FListeByUser.ForceAccessibilite := FaPublic;
{$ENDIF}
end;

procedure TFQR.BFermeClick(Sender: TObject);
begin
Close ;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

procedure TFQR.ParamLaDate(Sender: TObject; var Key: Char);
begin
ParamDate(Self,Sender,Key) ;
end;


procedure TFQR.FExerciceChange(Sender: TObject);
begin
//if FExercice.ItemIndex<>0 then ExoToDates(FExercice.Value,FDateCompta1,FDateCompta2) ;
// GP 13/03/97 N° 1138
{ AJOUT ME 18/04/2001 }
// Dans le cadre de la gestion des standards
// les flitres ramenés n'ont pas notion d'exercice
if (ctxPCL in V_PGI.PGIContexte) and (not (ctxStandard in V_PGI.PGIContexte)) and (FExercice.Value = '')  and QRLoadingExo then
begin
  if VH^.CPExoRef.Code <> '' then FExercice.Value := VH^.CPExoRef.Code
  else FExercice.Value := FExercice.values[0];
  ExoToDates(FExercice.Value,FDateCompta1,FDateCompta2) ;
end;
If QRLoadingExo then Exit ;
ExoToDates(FExercice.Value,FDateCompta1,FDateCompta2) ;
end;

procedure TFQR.FMontantClick(Sender: TObject);
{ Blocage accès devise si devise pivot demandée }
begin
If QRChangeDev Then Exit ;
//If Etat=etBalAgeDev Then Exit ;
if ((FDevises.ItemIndex=0)and(FDevises.vide=True)) or (FDevises.Value=V_PGI.DevisePivot) then
   BEGIN
   if FMontant.ItemIndex=1 then
      BEGIN
      // Rony 20/10/97 --> Bal Budget
      if (FNatureEtat<>neBalBud) And (FNatureEtat<>neBroBud) And (FNatureEtat<>neBroBudAna) then MsgQR.Execute(0,caption,'') ;
      FMontant.ItemIndex:=0 ;
      END ;
   (*
   If (FNatureEtat=nePoi)or(FNatureEtat=neJbq)or(FNatureEtat=neRap) then If FMontant.ItemIndex=2 Then FMontant.ItemIndex:=0 ;
   *)
   END Else
   BEGIN  // rony 29/05/97 -- > Pointage & JustifBanque
   If (FNatureEtat=nePoi)or(FNatureEtat=neJbq)or(FNatureEtat=neRap) then
      BEGIN
      FMontant.ItemIndex:=1 ;
      END ;
   END ;
If (FNatureEtat=nePoi)or(FNatureEtat=neJbq)or(FNatureEtat=neRap) then Exit ;

(*
if (FDevises.ItemIndex=0) or (FDevises.Value=V_PGI.DevisePivot) then
   BEGIN
   if FMontant.ItemIndex=1 then BEGIN MsgQR.Execute(0,caption,'') ; FMontant.ItemIndex:=0 ; END ;
   END ;
If Not V_PGI.Sav Then If FMontant.ItemIndex=2 Then FMontant.ItemIndex:=0 ;
*)
end;

procedure TFQR.FDevisesChange(Sender: TObject);
{ Blocage accès devise si devise pivot demandée }
begin
QRChangeDev:=TRUE ; // Rony 29/05/1997 -- > Pointage & JustifBanque
if ((FDevises.ItemIndex=0)and(FDevises.vide=True)) or (FDevises.Value=V_PGI.DevisePivot) then
   BEGIN
   if FMontant.ItemIndex=1 then FMontant.ItemIndex:=0 ;
   if (FNatureEtat=nePoi)or(FNatureEtat=neJbq)or(FNatureEtat=neRap) then if FMontant.ItemIndex=0 then FMontant.ItemIndex:=1 ;
   END Else if (FNatureEtat=nePoi)or(FNatureEtat=neJbq)or(FNatureEtat=neRap) then FMontant.ItemIndex:=1 ;
QRChangeDev:=FALSE ;

(*
if (FDevises.ItemIndex=0) or (FDevises.Value=V_PGI.DevisePivot) then
   BEGIN
   if FMontant.ItemIndex=1 then FMontant.ItemIndex:=0 ;
   END ;
*)
end;

procedure TFQR.piedpageAfterPrint(BandPrinted: Boolean);
begin
If BTitreOk Then
   BEGIN
   TitreEntete.Visible:=True ; TitreBarre.Visible:=True ;
   EnteteAutrePage.Enabled:=TRUE ;
   END Else BTitreOk:=TRUE ;
If AvecReport Then BottomReport.enabled:=TRUE ;
end;

procedure TFQR.BTitreAfterPrint(BandPrinted: Boolean);
begin
TitreEntete.Visible:=False ;
TitreBarre.Visible:=False ;
EntetePage.Enabled:=TRUE ;
//If CritEdt.RappelCrit Then TrouveOverlay(TRUE) ;
BTitre.ForceNewPage:=FALSE ;
end;

Procedure TFQR.GereZoneJoker(Sender: TObject ; NomControle : String ; HCpt : THCpteEdit) ;
{$IFNDEF CCS3}
var Te : TEdit ;
    fb : TFichierBase ;
    StS : String ;
{$ENDIF}
begin
{$IFNDEF CCS3}
  if TControl(Sender).Name<>NomControle Then
    Exit
  else
  begin
    fb := CaseFic(HCpt.Zoomtable) ;
    Te := TEdit(Sender) ;
    if (fb in [fbAxe1..fbAxe5]) and
       VH^.Cpta[fb].Structure and
       // GCO - 29/11/2006 - FQ 19175
       ExisteSQL('SELECT SS_AXE FROM STRUCRSE WHERE SS_AXE = "' + FBToAxe(fb) + '"') then
    begin
      StS := Te.text ;
      if ChoisirSousPlan(fb,StS,TRUE,taModif) then
        Te.text := StS;
    end;
  end;
{$ENDIF}
end;

procedure TFQR.FCpte1Change(Sender: TObject);
begin
If FNatureEtat=neRap then exit ;
AvecJoker:=Joker(FCpte1, FCpte2, FJoker) ;
TFaG.Visible:=Not AvecJoker ;
TFGen.Visible:=Not AvecJoker ;
TFGenJoker.Visible:=AvecJoker ;
end;

Function TFQR.AvecReport : Boolean ;
BEGIN
Result:=FALSE ;
Case FNatureEtat of
 neBal       : Result:=CritEDT.BAL.FormatPrint.Report.OkAff ;
 neGL        : Result:=CritEDT.GL.FormatPrint.Report.OkAff ;
 neJU        : Result:=CritEDT.JU.FormatPrint.Report.OkAff ;
 neJal       : Result:=CritEDT.Jal.FormatPrint.Report.OkAff ;
 neJalR      : Result:=CritEDT.JalR.FormatPrint.Report.OkAff ;
 neEch       : Result:=CritEDT.Ech.FormatPrint.Report.OkAff ;
 neGLV       : Result:=CritEDT.GlV.FormatPrint.Report.OkAff ;
 neCum       : Result:=CritEDT.BAL.FormatPrint.Report.OkAff ;
 neBrouAna   : Result:=CritEDT.BrouAna.FormatPrint.Report.OkAff ;
 neJbq       : Result:=CritEDT.Jbq.FormatPrint.Report.OkAff ;
 nePoi       : Result:=CritEDT.Poi.FormatPrint.Report.OkAff ;
 neRap       : Result:=CritEDT.Rap.FormatPrint.Report.OkAff ;
 neBalBud    : Result:=CritEDT.BalBud.FormatPrint.Report.OkAff ;
 neBroBud    : Result:=CritEDT.BroBud.FormatPrint.Report.OkAff ;
 neBroBudAna : Result:=CritEDT.BroBudAna.FormatPrint.Report.OkAff ;
 nePlanEdt   : Result:=CritEDT.PlanEdt.FormatPrint.Report.OkAff ;
 END ;
END ;

Function TFQR.OkRajouteLabelRupture : Boolean ;
BEGIN
Result:=FALSE ;
If (HauteurBandeRuptIni=0) Or (CritEdt.Rupture=rLibres) Then Result:=TRUE ;
END ;

procedure TFQR.InitDivers ;
Var ll : Integer ;
    okok : boolean ;
    i : Integer ;
BEGIN
{ Labels pied de page }
BTitreOk:=TRUE ; QRSautPage:=TRUE ;
  // GC - 27/11/2001 - Modif pour grisés plus clairs et bas de page à la mode PCL
  if CtxPCL in V_PGI.PGIContexte then
  begin
    if FCouleur.Checked then  // Uniquement en non couleur
    begin
      for i:=0 to EntetePage.ControlCount -1 do
      begin
        if EntetePage.Controls[i] is TQrLabel then
        begin
          TQrLabel(EntetePage.Controls[i]).Color := $E6E6E6;
          TQrLabel(EntetePage.Controls[i]).Font.Color := clBlack;
        end;
      end;
      QrSysData3.Color := $E6E6E6;
      QrsysData3.Font.Color := ClBlack;
      QrSysData4.Color := $E6E6E6;
      QrsysData4.Font.Color := ClBlack;
      TTitre.Color := $E6E6E6;
      TTitre.Font.Color := ClBlack;
    end;
    QrsysData1.Visible := False;
    RSociete.Caption := MsgRien.Mess[4] + ' ' + V_PGI.NoDossier + ' - ' + GetParamSocSecur('SO_LIBELLE', '');
    RSociete.Top := QrSysData2.Top;
    RCopyRight.Visible := False;
    RNumVersion.Visible := False;
    RUtilisateur.Visible:= False;
  end
  else
  begin
  RSociete.Caption:=MsgRien.Mess[4]+' '+V_PGI.CodeSociete+' '+V_PGI.NomSociete ;
  RUtilisateur.Caption:=MsgRien.Mess[5]+' '+V_PGI.User+' '+V_PGI.UserName ;
  RCopyright.Caption := Copyright + ' - ' + TitreHalley ;
  RNumversion.Caption:=MsgQR.Mess[3]+' '+V_PGI.NumVersion+' '+MsgQR.Mess[4]+' '+DateToStr(V_PGI.DateVersion);
  END ;
EnteteAutrePage.Enabled:=False ; //OkOk:=TRUE ;
BTitreCrit.Enabled:=FALSE ; //BTitreCrit.ForceNewPage:=CritEdt.RappelCrit ;
(*
If CritEdt.RappelCrit then
   BEGIN
   EntetePage.Enabled:=False ;
   piedpage.Enabled:=False ;  TitreBarre.Enabled:=False ;
   END ;
//BTitre.ForceNewPage:=FALSE ;
*)
//Simon
OkOk:=AvecReport ;
Case FNatureEtat of
 neBal       : BDetail.Frame.DrawTop:=CritEDT.BAL.FormatPrint.PrSepCompte ;
 neGL        : BDetail.Frame.DrawTop:=CritEDT.GL.FormatPrint.PrSepCompte[1] ;
 neJU        : BDetail.Frame.DrawTop:=CritEDT.JU.FormatPrint.PrSepCompte[1] ;
 neJal       : BDetail.Frame.DrawTop:=CritEDT.Jal.FormatPrint.PrSepCompte[1] ;
 neJalR      : BDetail.Frame.DrawTop:=CritEDT.JalR.FormatPrint.PrSepCompte[1] ;
 neEch       : ;
 neGLV       : ;
 neCum       : BDetail.Frame.DrawTop:=CritEDT.CUM.FormatPrint.PrSepCompte ;
 neBrouAna   : BDetail.Frame.DrawTop:=CritEDT.BrouAna.FormatPrint.PrSepCompte[1] ;
 neJbq       : ;
 nePoi       : ;
 neRap       : ;
 neBalBud    : ;
 neGLBud     : ;
 neBroBud    : BDetail.Frame.DrawTop:=CritEDT.BroBud.FormatPrint.PrSepCompte[1] ;
 neBroBudAna : BDetail.Frame.DrawTop:=CritEDT.BroBudAna.FormatPrint.PrSepCompte[1] ;
 neBalLib    : BDetail.Frame.DrawTop:=CritEDT.BalLib.FormatPrint.PrSepCompte ;
 nePlanEdt   : ;
 END ;
if okok then
   BEGIN
   TopReport.Enabled:=FALSE ; BottomReport.Enabled:=TRUE ;
   END else
   BEGIN
   TopReport.Enabled:=FALSE ; BottomReport.Enabled:=FALSE ;
   END ;
//Simon
Case FNatureEtat of
 neBal : BEGIN
         If OkRajouteLabelRupture Then RajouteLabelRupture('BRUPT') ;
         ll:=ChangeFormatBal(Self,PremTabColBal,CritEDT.BAL.FormatPrint.TabColl) ;
         CalculPourMiseEnPage(ll,QRP,Self,CritEDT.BAL.FormatPrint) ;
         END ;
  neGL : BEGIN
         If OkRajouteLabelRupture Then RajouteLabelRupture('BRUPT') ;
         ll:=ChangeFormatEdt(Self,PremTabColEdt,CritEDT.GL.FormatPrint.TabColl) ;
         CalculPourMiseEnPageEdt(ll,QRP,Self,CritEDT.GL.FormatPrint) ;
         END ;
  neJU : BEGIN
         If OkRajouteLabelRupture Then RajouteLabelRupture('BRUPT') ;
         ll:=ChangeFormatEdt(Self,PremTabColedt,CritEDT.JU.FormatPrint.TabColl) ;
         CalculPourMiseEnPageEdt(ll,QRP,Self,CritEDT.JU.FormatPrint) ;
         END ;
  neJal: BEGIN
//         CritEDT.Jal.FormatPrint.ModePrint:=PPaysage ;
         // GC - 20/12/2001 ( false à la fin )
         if (Etat = etJalDivPCL) Or (Etat = etJalDiv) then
             ll:=ChangeFormatEdt(Self,PremTabColedt,CritEDT.Jal.FormatPrint.TabColl,FALSE)
           else
         // GC - 20/12/2001 - FIN
         ll:=ChangeFormatEdt(Self,PremTabColedt,CritEDT.Jal.FormatPrint.TabColl) ;
         CalculPourMiseEnPageEdt(ll,QRP,Self,CritEDT.Jal.FormatPrint) ;
         END ;
  neJalR:BEGIN
//         CritEDT.Jal.FormatPrint.ModePrint:=PPaysage ;
         ll:=ChangeFormatEdt(Self,PremTabColedt,CritEDT.JalR.FormatPrint.TabColl) ;
         CalculPourMiseEnPageEdt(ll,QRP,Self,CritEDT.JalR.FormatPrint) ;
         END ;
  neEch: BEGIN
         ll:=ChangeFormatEdt(Self,PremTabColedt,CritEDT.Ech.FormatPrint.TabColl) ;
         CalculPourMiseEnPageEdt(ll,QRP,Self,CritEDT.Ech.FormatPrint) ;
         END ;
  neGlV: BEGIN
         If OkRajouteLabelRupture Then RajouteLabelRupture('BRUPT') ;
         ll:=ChangeFormatEdt(Self,PremTabColedt,CritEDT.GlV.FormatPrint.TabColl) ;
         CalculPourMiseEnPageEdt(ll,QRP,Self,CritEDT.GlV.FormatPrint) ;
         END ;
  neCum: BEGIN
         ll:=ChangeFormatBal(Self,PremTabColBal,CritEDT.Cum.FormatPrint.TabColl) ;
         CalculPourMiseEnPage(ll,QRP,Self,CritEDT.Cum.FormatPrint) ;
         END ;
  neBrouAna : BEGIN
         ll:=ChangeFormatEdt(Self,PremTabColEdt,CritEDT.BrouAna.FormatPrint.TabColl) ;
         CalculPourMiseEnPageEdt(ll,QRP,Self,CritEDT.BrouAna.FormatPrint) ;
         END ;
  neJbq: BEGIN
         ll:=ChangeFormatEdt(Self,PremTabColedt,CritEDT.Jbq.FormatPrint.TabColl) ;
         CalculPourMiseEnPageEdt(ll,QRP,Self,CritEDT.Jbq.FormatPrint) ;
         END ;
  nePoi: BEGIN
         ll:=ChangeFormatEdt(Self,PremTabColedt,CritEDT.Poi.FormatPrint.TabColl) ;
         CalculPourMiseEnPageEdt(ll,QRP,Self,CritEDT.Poi.FormatPrint) ;
         END ;
  neRap: BEGIN
         ll:=ChangeFormatEdt(Self,PremTabColedt,CritEDT.Rap.FormatPrint.TabColl) ;
         CalculPourMiseEnPageEdt(ll,QRP,Self,CritEDT.Rap.FormatPrint) ;
         END ;
 neBalbud : BEGIN
         If OkRajouteLabelRupture Then RajouteLabelRupture('BRappelLibreHaut') ;
         ll:=ChangeFormatEdt(Self,PremTabColEdt,CritEDT.BalBud.FormatPrint.TabColl) ;
         CalculPourMiseEnPageEdt(ll,QRP,Self,CritEDT.BalBud.FormatPrint) ;
         CritEdt.BalBud.SANbud:='N-' ;
         END ;
neBroBud : BEGIN
         ll:=ChangeFormatEdt(Self,PremTabColEdt,CritEDT.BroBud.FormatPrint.TabColl) ;
         CalculPourMiseEnPageEdt(ll,QRP,Self,CritEDT.BroBud.FormatPrint) ;
         END ;
neBroBudAna : BEGIN
         ll:=ChangeFormatEdt(Self,PremTabColEdt,CritEDT.BroBudAna.FormatPrint.TabColl) ;
         CalculPourMiseEnPageEdt(ll,QRP,Self,CritEDT.BroBudAna.FormatPrint) ;
         END ;
neBalLib : BEGIN
         ll:=ChangeFormatBal(Self,PremTabColBal,CritEDT.BalLib.FormatPrint.TabColl) ;
         CalculPourMiseEnPage(ll,QRP,Self,CritEDT.BalLib.FormatPrint) ;
         END ;
nePlanEdt : BEGIN
         If OkRajouteLabelRupture Then RajouteLabelRupture('BRappelLibreHaut') ;
         ll:=ChangeFormatEdt(Self,PremTabColEdt,CritEDT.PlanEdt.FormatPrint.TabColl) ;
         CalculPourMiseEnPageEdt(ll,QRP,Self,CritEDT.PlanEdt.FormatPrint) ;
         END ;
   END ;
END ;

// procedure TFQR.CalculSolde(Var MontantD,MontantC,SoldeD,SoldeC : Double) ;
procedure TFQR.CalculSolde( MontantD,MontantC : Double ; var SoldeD,SoldeC : Double) ; // GC - 20/12/2001
var Solde : Double ;
BEGIN
Solde:=MontantD-MontantC ;
if Solde<0
   then BEGIN SoldeD:=0 ; SoldeC:=Abs(Solde) ; END
   else BEGIN SoldeD:=Solde ; SoldeC:=0 ; END ;
END ;

procedure TFQR.FNatureCptChange(Sender: TObject);
begin
//Simon
If Not QRLoading Then
   BEGIN
   Fcpte1.Clear ; FCpte2.Clear ; FJoker.Clear ;
   END ;
If FComposite And (FNatureBase=nbGen) And (FModeSelect=msGenEcr) Then
   BEGIN
   Case FNatureCpt.ItemIndex of
     0 : BEGIN FCpte1.ZoomTable:=tzGCollectif   ; FCpte2.ZoomTable:=tzGCollectif    ; END ;
     1 : BEGIN FCpte1.ZoomTable:=tzGCollClient  ; FCpte2.ZoomTable:=tzGCollClient   ; END ;
     2 : BEGIN FCpte1.ZoomTable:=tzGCollDivers  ; FCpte2.ZoomTable:=tzGCollDivers   ; END ;
     3 : BEGIN FCpte1.ZoomTable:=tzGCollFourn   ; FCpte2.ZoomTable:=tzGCollFourn    ; END ;
     4 : BEGIN FCpte1.ZoomTable:=tzGCollSalarie ; FCpte2.ZoomTable:=tzGCollSalarie  ; END ;
     END ;
   END Else
   BEGIN
   Case FNatureBase Of
   nbGen : BEGIN
           Case FNatureCpt.ItemIndex of
             0 : BEGIN FCpte1.ZoomTable:=tzGeneral      ; FCpte2.ZoomTable:=tzGeneral      ; END ;
             1 : BEGIN FCpte1.ZoomTable:=tzGbanque      ; FCpte2.ZoomTable:=tzGbanque      ; END ;
             2 : BEGIN FCpte1.ZoomTable:=tzGcaisse      ; FCpte2.ZoomTable:=tzGcaisse      ; END ;
             3 : BEGIN FCpte1.ZoomTable:=tzGcharge      ; FCpte2.ZoomTable:=tzGcharge      ; END ;
             4 : BEGIN FCpte1.ZoomTable:=TzGCollClient  ; FCpte2.ZoomTable:=TzGCollClient  ; END ;
             5 : BEGIN FCpte1.ZoomTable:=tzGCollDivers  ; FCpte2.ZoomTable:=tzGCollDivers  ; END ;
             6 : BEGIN FCpte1.ZoomTable:=tzGCollFourn   ; FCpte2.ZoomTable:=tzGCollFourn   ; END ;
             7 : BEGIN FCpte1.ZoomTable:=tzGCollSalarie ; FCpte2.ZoomTable:=tzGCollSalarie ; END ;
             8 : BEGIN FCpte1.ZoomTable:=tzGDivers      ; FCpte2.ZoomTable:=tzGDivers      ; END ;
             9 : BEGIN FCpte1.ZoomTable:=tzGextra       ; FCpte2.ZoomTable:=tzGextra       ; END ;
            10 : BEGIN FCpte1.ZoomTable:=tzGcharge      ; FCpte2.ZoomTable:=tzGcharge      ; END ;
            11 : BEGIN FCpte1.ZoomTable:=tzGimmo        ; FCpte2.ZoomTable:=tzGimmo        ; END ;
            12 : BEGIN FCpte1.ZoomTable:=tzGproduit     ; FCpte2.ZoomTable:=tzGproduit     ; END ;
            13 : BEGIN FCpte1.ZoomTable:=tzGTIC         ; FCpte2.ZoomTable:=tzGTIC         ; END ;
            14 : BEGIN FCpte1.ZoomTable:=tzGTID         ; FCpte2.ZoomTable:=tzGTID         ; END ;
            end;
           END ;
   nbGenT : BEGIN
{           Case FNatureCpt.ItemIndex of
            0 : BEGIN FCpte1.ZoomTable:=tzGTIDTIC      ; FCpte2.ZoomTable:=tzGTIDTIC      ; END ;
            1 : BEGIN FCpte1.ZoomTable:=tzGTIC         ; FCpte2.ZoomTable:=tzGTIC         ; END ;
            2 : BEGIN FCpte1.ZoomTable:=tzGTID         ; FCpte2.ZoomTable:=tzGTID         ; END ;
            end;}
            if (FNatureCpt.Value = 'TIC') then begin FCpte1.ZoomTable := tzGTIC ; FCpte2.ZoomTable := tzGTIC; end
            else if (FNatureCpt.Value = 'TID') then begin FCpte1.ZoomTable := tzGTID; FCpte2.ZoomTable := tzGTID; end
            else begin FCpte1.ZoomTable := tzGTIDTIC; FCpte2.ZoomTable := tzGTIDTIC; end;
           END ;
   nbAux : BEGIN
{           Case FNatureCpt.ItemIndex of
             0 : BEGIN FCpte1.ZoomTable:=tzTiers      ; FCpte2.ZoomTable:=tzTiers      ; END ;  // ''
             1 : BEGIN FCpte1.ZoomTable:=tzTCrediteur ; FCpte2.ZoomTable:=tzTCrediteur ; END ;  // AUC
             2 : BEGIN FCpte1.ZoomTable:=tzTDebiteur  ; FCpte2.ZoomTable:=tzTDebiteur  ; END ;  // AUD
             3 : BEGIN FCpte1.ZoomTable:=tzTClient    ; FCpte2.ZoomTable:=tzTClient    ; END ;  // CLI
             4 : BEGIN FCpte1.ZoomTable:=tzTDivers    ; FCpte2.ZoomTable:=tzTDivers    ; END ;  // DIV
             5 : BEGIN FCpte1.ZoomTable:=tzTFourn     ; FCpte2.ZoomTable:=tzTFourn     ; END ;  // FOU
             6 : BEGIN FCpte1.ZoomTable:=tzTSalarie   ; FCpte2.ZoomTable:=tzTSalarie   ; END ;  // SAL
            end;}
            if FNatureCpt.Value='AUC' then begin FCpte1.ZoomTable:=tzTCrediteur ; FCpte2.ZoomTable:=tzTCrediteur; end else
            if FNatureCpt.Value='AUD' then begin FCpte1.ZoomTable:=tzTDebiteur  ; FCpte2.ZoomTable:=tzTDebiteur ; end else
            if FNatureCpt.Value='CLI' then begin FCpte1.ZoomTable:=tzTClient    ; FCpte2.ZoomTable:=tzTClient   ; end else
            if FNatureCpt.Value='DIV' then begin FCpte1.ZoomTable:=tzTDivers    ; FCpte2.ZoomTable:=tzTDivers   ; end else
            if FNatureCpt.Value='FOU' then begin FCpte1.ZoomTable:=tzTFourn     ; FCpte2.ZoomTable:=tzTFourn    ; end else
            if FNatureCpt.Value='SAL' then begin FCpte1.ZoomTable:=tzTSalarie   ; FCpte2.ZoomTable:=tzTSalarie  ; end
                                      else begin FCpte1.ZoomTable:=tzTiers      ; FCpte2.ZoomTable:=tzTiers     ; end;
           END ;
   nbSec : BEGIN
           Case FNatureCpt.ItemIndex of
            0 : BEGIN FCpte1.ZoomTable:=tzSection ; FCpte2.ZoomTable:=tzSection   ; END ;
            1 : BEGIN FCpte1.ZoomTable:=tzSection2 ; FCpte2.ZoomTable:=tzSection2 ; END ;
            2 : BEGIN FCpte1.ZoomTable:=tzSection3 ; FCpte2.ZoomTable:=tzSection3 ; END ;
            3 : BEGIN FCpte1.ZoomTable:=tzSection4 ; FCpte2.ZoomTable:=tzSection4 ; END ;
            4 : BEGIN FCpte1.ZoomTable:=tzSection5 ; FCpte2.ZoomTable:=tzSection5 ; END ;
           end ;
           END ;
   nbJal : BEGIN //Simon
           END ;
   nbJalAna : BEGIN // Rony --> Brouil Anal
              Case FNatureCpt.ItemIndex of
               0 : BEGIN FCpte1.ZoomTable:=tzJAna ; FCpte2.ZoomTable:=tzJAna   ; END ;
               1 : BEGIN FCpte1.ZoomTable:=tzJAna1 ; FCpte2.ZoomTable:=tzJAna1 ; END ;
               2 : BEGIN FCpte1.ZoomTable:=tzJAna2 ; FCpte2.ZoomTable:=tzJAna2 ; END ;
               3 : BEGIN FCpte1.ZoomTable:=tzJAna3 ; FCpte2.ZoomTable:=tzJAna3 ; END ;
               4 : BEGIN FCpte1.ZoomTable:=tzJAna4 ; FCpte2.ZoomTable:=tzJAna4 ; END ;
               5 : BEGIN FCpte1.ZoomTable:=tzJAna5 ; FCpte2.ZoomTable:=tzJAna5 ; END ;
              end ;
              END ;
   nbBudSec : BEGIN
              Case FNatureCpt.ItemIndex of
               0 : BEGIN FCpte1.ZoomTable:=tzBudSec1 ; FCpte2.ZoomTable:=tzBudSec1   ; END ;
               1 : BEGIN FCpte1.ZoomTable:=tzBudSec2 ; FCpte2.ZoomTable:=tzBudSec2 ; END ;
               2 : BEGIN FCpte1.ZoomTable:=tzBudSec3 ; FCpte2.ZoomTable:=tzBudSec3 ; END ;
               3 : BEGIN FCpte1.ZoomTable:=tzBudSec4 ; FCpte2.ZoomTable:=tzBudSec4 ; END ;
               4 : BEGIN FCpte1.ZoomTable:=tzBudSec5 ; FCpte2.ZoomTable:=tzBudSec5 ; END ;
              end ;
              END ;
   nbBudJal : BEGIN
              END ;
     END ;
   END ;
end;

procedure TFQR.EnteteAutrePageAfterPrint(BandPrinted: Boolean);
begin
// XX Rony 10/04/97 TOPREPORT.Enabled:=TRUE ;
//TitreBarre.Width:=EnteteAutrePage.Width ;
end;

procedure TFQR.TOPREPORTAfterPrint(BandPrinted: Boolean);
begin
// XX Rony 10/04/97 TOPREPORT.Enabled:=TRUE ;
end;

procedure TFQR.BParamListeClick(Sender: TObject);
begin
ParamListe(NomListeGA,NIL,NIL) ;
end;

procedure TFQR.AvanceClick(Sender: TObject);
begin
BParamListe.Visible:=Avance.Checked ;
end;

function  TFQR.SQLModeSelectMul(Where1,Where2,PreBTot : String) : String ;
{ Contruction d'une partie de SQLAux }
Var TOTDEBIT,TOTCREDIT,TOTDEBITANO,TOTCREDITANO,St : String ;
    TheAxe : String ;
    PreB1,PreB2,PreE,PreJ1,PreJ2,Table1,Table2 : String ;
    Exo : TExoDate ;
    ts : Integer ;
    me : Boolean ;
    Valeur : String ;
BEGIN
//Simon
Result:='' ; TheAxe:='' ;
If Not WhatPREBPREJ(FModeSelect,PreJ1,PreJ2,PreB1,PreB2,PreE,Table1,Table2) Then Exit ;
If PreBTot='' Then PreBtot:=PreB1 ;
TOTDEBIT:=PreBTot+'TOTDEBE' ; TOTCREDIT:=PreBTot+'TOTCREE' ;
TOTDEBITANO:=PreBTot+'TOTDEBANO' ; TOTCREDITANO:=PreBTot+'TOTCREANO' ;
Case FNatureEtat of
   neBal : BEGIN me:=CritEDT.MonoExo ; Ts:=CritEDT.BAL.TypCpt ; Exo:=CritEDT.Exo ; END ;
   neGL  : BEGIN me:=CritEDT.MonoExo ; Ts:=CritEDT.GL.TypCpt ; Exo:=CritEDT.Exo ; END ;
   Else exit ;
   END ;
If ME Then
   BEGIN
   If Exo.Code=VH^.Precedent.Code Then BEGIN TOTDEBIT:=PreBTot+'TOTDEBP' ; TOTCREDIT:=PreBTot+'TOTCREP' ; END ;
   If Exo.Code=VH^.EnCours.Code   Then BEGIN TOTDEBIT:=PreBTot+'TOTDEBE' ; TOTCREDIT:=PreBTot+'TOTCREE' ; END ;
   If Exo.Code=VH^.Suivant.Code   Then BEGIN TOTDEBIT:=PreBTot+'TOTDEBS' ; TOTCREDIT:=PreBTot+'TOTCRES' ; END ;
   END ;
If ME And (Exo.Code=Vh^.Suivant.Code) Then
   BEGIN
   TOTDEBITANO:=PreBTot+'TOTDEBANON1' ; TOTCREDITANO:=PreBTot+'TOTCREANON1' ;
   END ;
Valeur:=CritEdt.QualifPiece ;
//Type ttTypePiece  = (tpReel,tpSim,tpPrev,tpSitu,tpRevi) ;
//TS = 0 : Exo, 1 = Tous, 2 = Cpt non soldés, 3 = Période
If Where1='' Then Where1:='Q.'+PreB1+PreJ1 ;
Case FModeSelect of
   msGenAna,msSecAna : BEGIN
                       Case FNatureEtat Of
                         neBal : TheAxe:=CritEdt.Bal.Axe ;
                         neGL : TheAxe:=CritEdt.GL.Axe ;
                         END ;
                       END ;
   END ;
Case TS of
  0 : BEGIN
  (*
      if (Valeur<>'TOU') Then
         BEGIN
         St:=SQLExisteMul(PreE,PreJ1,PreJ2,PreB1,PreB2,TOTDEBIT,TOTCREDIT,Valeur,V_PGI.Controleur,AvecRevision.State,CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,Where1,Where2) ;
         END Else
         BEGIN
         St:=WhatExisteMul(PreE,PreJ1,PreJ2,PreB2,Valeur,V_PGI.Controleur,AvecRevision.State,
                           CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,Where1,Where2,0) ;
         END ;
      *)
      St:=WhatExisteMul(PreE,PreJ1,PreJ2,PreB2,Valeur,V_PGI.Controleur,AvecRevision.State,
                        CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,Where1,Where2,0,TheAxe) ;
      Result:=St ;
      END ;
  1 : BEGIN
      Result:='' ;
      (* GP le 04/11/99
      If OkZoomEdt Then Else
      *)
         Result:='('+
                 WhatExisteMul(PreE,PreJ1,PreJ2,PreB2,Valeur,V_PGI.Controleur,AvecRevision.State,
                              CritEdt.DateDeb,CritEdt.DateFin,CritEdt.Exo.Code,Where1,Where2,0,TheAxe)
                 +')' ;
      (*
          Result:='(('+PreB1+PreJ1+'='+PreB1+PreJ1+') AND ('+
                  WhatExisteMul(PreE,PreJ1,PreJ2,PreB2,Valeur,V_PGI.Controleur,AvecRevision.State,
                               CritEdt.DateDeb,CritEdt.DateFin,CritEdt.Exo.Code,Where1,Where2,0)
                 +'))' ;
      *)
      END ;
  2 : BEGIN
      (*
      if (Valeur<>'TOU') Then
         BEGIN
         if Valeur='NOR' then BEGIN Result:=PreB1+'TOTALDEBIT<>'+PreB1+'TOTALCREDIT ' ; END ;
         if Valeur='NSS' then
            BEGIN
            Result:=PreB1+'TOTALDEBIT<>'+PreB1+'TOTALCREDIT ' ;
            Result:=Result+' Or (('+PreB1+'TOTALDEBIT='+PreB1+'TOTALCREDIT '+') And (' ;
            Result:=Result+
                    WhatExisteMul(PreE,PreJ1,PreJ2,PreB2,Valeur,V_PGI.Controleur,AvecRevision.State,
                    CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,Where1,Where2,0)+'))' ;
            END ;
         END Else
         BEGIN
         Result:=PreB1+'TOTALDEBIT<>'+PreB1+'TOTALCREDIT ' ;
         Result:=Result+' Or (('+PreB1+'TOTALDEBIT='+PreB1+'TOTALCREDIT '+') And (' ;
         Result:=Result+
                 WhatExisteMul(PreE,PreJ1,PreJ2,PreB2,Valeur,V_PGI.Controleur,AvecRevision.State,
                 CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,Where1,Where2,0)+'))' ;
         END ;
      *)
      St:=WhatExisteMul(PreE,PreJ1,PreJ2,PreB2,Valeur,V_PGI.Controleur,AvecRevision.State,
                        CritEdt.Date1,CritEdt.Date2,CritEdt.Exo.Code,Where1,Where2,0,TheAxe) ;
      Result:=St ;
      END ;
  3 : BEGIN
      (*
      if (Valeur<>'TOU') Then
         BEGIN
         St:=SQLExisteMul(PreE,PreJ1,PreJ2,PreB1,PreB2,TOTDEBIT,TOTCREDIT,Valeur,V_PGI.Controleur,AvecRevision.State,CritEdt.DateDeb,CritEdt.DateFin,CritEdt.Exo.Code,Where1,Where2) ;
         END Else
         BEGIN
         St:=WhatExisteMul(PreE,PreJ1,PreJ2,PreB2,Valeur,V_PGI.Controleur,AvecRevision.State,
                           CritEdt.DateDeb,CritEdt.DateFin,CritEdt.Exo.Code,Where1,Where2,0) ;
         END ;
      *)
      St:=WhatExisteMul(PreE,PreJ1,PreJ2,PreB2,Valeur,V_PGI.Controleur,AvecRevision.State,
                        CritEdt.DateDeb,CritEdt.DateFin,CritEdt.Exo.Code,Where1,Where2,0,TheAxe) ;
      Result:=St ;
      END ;
 end ;
END ;

Function TFQR.AlimSQLMul(Q : TQuery ; Niveau : integer) : Boolean ;
Var St,Axe,PreB1,PreB2,PreJ1,PreJ2,PreE,Table1,Table2,LeSauf,LeSSauf,LeSQLSauf,LeSSQLSauf,NatCpt, NatureRupt : String ;
//    ItypCpt : Integer ;
    LeFb : TFichierBase ;
    EPlansCorresp : Byte;
    EOnlyCompteAssocie : Boolean ;
begin
Result:=FALSE ;
//EPlansCorresp:=0 ;
Case FNatureEtat Of
   neBal : BEGIN
           Axe:=CritEdt.Bal.Axe ; {ITypCpt:=CritEDT.BAL.TypCpt ;} NatCpt:=CritEDT.NatureCpt ;
           LeSauf:=CritEDT.BAL.Sauf ; LeSSauf:=CritEDT.BAL.SSqlSauf ;
           LeSQLSauf:=CritEdt.Bal.SQLSauf ; LeSSQLSauf:=CritEdt.Bal.SSQLSauf ;
           EPlansCorresp:=CritEdt.Bal.PlansCorresp ;
           EOnlyCompteAssocie:=CritEdt.Bal.OnlyCptAssocie ;
           END ;
   neGl : BEGIN
          Axe:=CritEdt.GL.Axe ; {ITypCpt:=CritEDT.GL.TypCpt ;} NatCpt:=CritEDT.NatureCpt ;
          LeSauf:=CritEDT.GL.Sauf ; LeSSauf:=CritEDT.GL.SSauf ;
          LeSQLSauf:=CritEdt.GL.SQLSauf ; LeSSQLSauf:=CritEdt.GL.SSQLSauf ;
          EPlansCorresp:=CritEdt.GL.PlansCorresp ;
          EOnlyCompteAssocie:=CritEdt.GL.OnlyCptAssocie ;
          END ;
   Else Exit ;
   END ;
If Not WhatPREBPREJ(FModeSelect,PreJ1,PreJ2,PreB1,PreB2,PreE,Table1,Table2) Then Exit ;
Q.CLose ;

LeFb:=fbGene ;
if PreB1='G_' then LeFb:=fbGene else
if PreB1='T_' then LeFb:=fbAux else
if PreB1='BG_' then LeFb:=fbBudgen else
if PreB1='S_' then LeFb:=AxeToFb(Axe) else
if PreB1='BS_' then LeFb:=AxeToFbBud(Axe) ;
If Niveau=0 Then
   BEGIN
   Q.Sql.Clear ;
   St:='Select '+PreB1+PreJ1+', '+PreB1+'LIBELLE, '+PreB1+'TOTDEBANO, '+PreB1+'TOTCREANO, '+
       PreB1+'TOTDEBANON1, '+PreB1+'TOTCREANON1, '+
       PreB1+'TOTDEBE, '+PreB1+'TOTCREE, '+PreB1+'SOLDEPROGRESSIF, '+PreB1+'SAUTPAGE, '+
       PreB1+'TOTAUXMENSUELS, '+PreB1+'TOTALDEBIT, '+PreB1+'TOTALCREDIT' ;
   If FModeSelect=msSecAna then St:=St+' , S_AXE, S_SECTIONTRIE ' ;
//CP : 09/02/98 : Ajout SQL : Etats TVA
   if Etat in [etTvaAcc..etTvaHT] then St:=St+', '+PreB1+'REGIMETVA,'+PreB1+'TVA, '+PreB1+'NATUREGENE ' ;
   Q.SQL.Add(St) ;
   Case CritEdt.Rupture of
     rLibres  : Q.Sql.Add(', '+PreB1+'TABLE0, '+PreB1+'TABLE1, '+PreB1+'TABLE2, '+PreB1+'TABLE3, '+PreB1+'TABLE4, '+PreB1+'TABLE5, '+PreB1+'TABLE6, '+PreB1+'TABLE7, '+PreB1+'TABLE8, '+PreB1+'TABLE9 ') ;
     rCorresp : Q.Sql.Add(', '+PreB1+'CORRESP1, '+PreB1+'CORRESP2 ') ;
     End ;
//   if (CritEdt.Rupture=rLibres) then Q.Sql.Add(', '+PreB1+'TABLE0, '+PreB1+'TABLE1, '+PreB1+'TABLE2, '+PreB1+'TABLE3, '+PreB1+'TABLE4, '+PreB1+'TABLE5, '+PreB1+'TABLE6, '+PreB1+'TABLE7, '+PreB1+'TABLE8, '+PreB1+'TABLE9 ') ;
   { Tables explorées par la SQL }
   St:=' From '+Table1+' Q Where ' ;
   if (OkV2 and (V_PGI.Confidentiel<>'1')) then St:=St+' '+PreB1+'CONFIDENTIEL<>"1" AND ' ;
   Case FModeSelect of
      msGenEcr : BEGIN
                 if (Etat<>etBALHTGA) and (Etat<>etTvaHT) then St:=St+'G_COLLECTIF="X" AND ' ;
//CP : 09/02/98 : Ajout SQL : Etat tva / Encaiss.
//                 if Etat in [etTvaHT] then St:=St+'G_TVASURENCAISS="X" AND ' ;
                 if (Etat=etBALHTGA) or (Etat=etTvaHT) then
                   BEGIN
                   if NatCpt='' then St:=St+'(G_NATUREGENE="CHA" OR G_NATUREGENE="PRO") AND '
                                else St:=St+'G_NATUREGENE="'+NatCpt+'" AND ' ;
                   If etat=etTvaHT Then St:=St+'G_TVA<>"" AND ' ;
                   END Else If NatCpt<>'' Then St:=St+'G_NATUREGENE="'+NatCpt+'" AND ' ;
                 END ;
      msGenAna : BEGIN
                 St:=St+'G_VENTILABLE="X" AND G_VENTILABLE'+Axe[2]+'="X" AND ' ;
                 if NatCpt<>'' then St:=St+'G_NATUREGENE="'+NatCpt+'" AND ' ;
                 END ;
      msAuxEcr : BEGIN
                 St:=St+'T_NATUREAUXI<>"NPC" AND ' ;
                 if NatCpt<>'' then St:=St+'T_NATUREAUXI="'+NatCpt+'" AND ' ;
                 END ;
      msSecAna : BEGIN
                 St:=St+'S_AXE="'+Axe+'" AND ' ;
                 END ;
      END ;
   Q.SQL.Add(St) ;
   St:=SQLModeSelect ;
   If St<>'' Then BEGIN Q.SQL.Add(St) ; St:='' ; END ;
   case Etat of
     etBalHTAG : St:=' AND (EXISTS(SELECT '+PreE+PreJ2+' FROM ECRITURE WHERE E_CONTREPARTIEAUX=Q.T_AUXILIAIRE ' ;
     etBalHTGA : St:=' AND (EXISTS(SELECT '+PreE+PreJ2+' FROM ECRITURE WHERE E_CONTREPARTIEGEN=Q.G_GENERAL ' ;
     etTvaAcc : St:=' AND (EXISTS(SELECT '+PreE+PreJ2+' FROM ECRITURE WHERE '
                   +'E_GENERAL=Q.G_GENERAL AND (E_ECHEENC1<>0 OR E_ECHEENC2<>0 '
                   +'OR E_ECHEENC3<>0 OR E_ECHEENC4<>0) AND E_EDITEETATTVA="-" ' ;
     etTvaHT : BEGIN
               Case CritEdt.GL.Encaissement Of
                  0 : St:=' AND (EXISTS(SELECT '+PreE+PreJ2+' FROM ECRITURE WHERE '
                            +'E_GENERAL=Q.G_GENERAL And E_TVAENCAISSEMENT="-" AND E_EDITEETATTVA="-"' ;
                  1 : St:=' AND (EXISTS(SELECT '+PreE+PreJ2+' FROM ECRITURE WHERE '
                            +'E_GENERAL=Q.G_GENERAL And E_TVAENCAISSEMENT="X" AND E_EDITEETATTVA="-"' ;
                  2 : St:=' AND (EXISTS(SELECT '+PreE+PreJ2+' FROM ECRITURE WHERE '
                            +'E_GENERAL=Q.G_GENERAL And (E_TVAENCAISSEMENT="X" OR E_TVAENCAISSEMENT="-") AND E_EDITEETATTVA="-"' ;
                  END ;
               END ;
     END ;
   if Etat in [etBalHTAG,etBalHTGA] then
     BEGIN
     if CritEDT.SJoker then St:=St+'AND '+PreE+PreJ2+' like "'+TraduitJoker(CritEdt.SCpt1)+'" ' Else
        BEGIN
        if CritEDT.SCpt1<>'' then St:=St+'AND '+PreE+PreJ2+'>="'+CritEdt.SCpt1+'" ' ;
        if CritEDT.SCpt2<>'' then St:=St+'AND '+PreE+PreJ2+'<="'+CritEdt.SCpt2+'" ' ;
        END ;
     END ;
   if St<>'' then BEGIN Q.SQL.Add(St+'))') ; St:='' ; END ;
   if CritEdt.Joker then St:='AND '+PreB1+PreJ1+' like "'+TraduitJoker(CritEdt.Cpt1)+'" ' Else
      BEGIN
      if CritEdt.Cpt1<>'' then St:='AND  '+PreB1+PreJ1+'>="'+CritEdt.Cpt1+'" ' ;
      if CritEdt.Cpt2<>'' then St:=St+'AND  '+PreB1+PreJ1+'<="'+CritEdt.Cpt2+'" ' ;
      END ;
   If St<>'' Then Q.SQL.Add(St) ;
   st:='' ; if LeSauf<>'' then St:=LeSqlSauf ; if st<>'' Then Q.SQL.Add(St) ;
   { Construction de la clause Order By de la SQL }
   Case CritEdt.Rupture of
     rRuptures : If leFb in [fbAxe1..fbAxe5] Then Q.Sql.Add(' Order By S_AXE, S_SECTIONTRIE ')
                                             Else Q.Sql.Add('Order By '+PreB1+PreJ1) ;
     rLibres  : BEGIN
                St:=WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,LeFb,EOnlyCompteAssocie) ;
                If St<>'' Then Q.Sql.Add(St) ;
                Q.Sql.Add('Order By '+OrderLibre(CritEdt.LibreTrie)+PreB1+PreJ1) ;
                END ;
     rCorresp : BEGIN
                if FOnlyCptAssocie.Checked then
                   BEGIN

                   Case FNatureBase Of
                     nbGen : NatureRupt:='GE'+IntToStr(EPlansCorresp) ;
                     nbAux : NatureRupt:='AU'+IntToStr(EPlansCorresp) ;
                     nbSec : NatureRupt:=CritEdt.Bal.Axe+IntToStr(EPlansCorresp) ;
                     END ;
                   Q.SQL.Add(' AND Exists(SELECT CR_CORRESP FROM CORRESP WHERE CR_TYPE="'+NatureRupt+'"') ;
                   Q.SQL.Add(' AND CR_CORRESP=Q.'+PreB1+'CORRESP'+IntToStr(EPlansCorresp)+')') ;
                   END ;
                Case EPlansCorresp Of
                    1 : Q.Sql.Add('Order By '+PreB1+'CORRESP1, '+PreB1+PreJ1) ;
                    2 : Q.Sql.Add('Order By '+PreB1+'CORRESP2, '+PreB1+PreJ1) ;
                    Else Q.Sql.Add('Order By '+PreB1+PreJ1) ;
                    END ;
                 END ;
     Else Q.Sql.Add('Order By '+PreB1+PreJ1) ;
     End ;
   (*
   If (CritEdt.Rupture=rLibres) then
      BEGIN
      Q.Sql.Add(WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,LeFb)) ;
      Q.Sql.Add('Order By '+OrderLibre(CritEdt.LibreTrie)+PreB1+PreJ1) ;
      END Else Q.SQL.Add(' Order By  '+PreB1+PreJ1) ;
   *)
   Result:=TRUE ;
   END ;
If Niveau=1 Then
   BEGIN
   Q.Sql.Clear ;
   St:='Select '+PreB2+PreJ2+', '+PreB2+'LIBELLE ' ;
   if FModeSelect=msGenAna then St:=St+' , S_AXE' ;
   if Etat in [EtTvaFac] then St:=St+', '+PreB2+'COLLECTIF,'+PreB2+'NATUREAUXI,'+PreB2+'REGIMETVA' ;
   Q.SQL.Add(St) ;
   St:=' From '+Table2+' Q1 Where ' ;
   if (OkV2 and (V_PGI.Confidentiel<>'1')) then St:=St+' '+PreB2+'CONFIDENTIEL<>"1" AND ' ;
   Case FModeSelect of
      msGenEcr : BEGIN
                 if Etat in [etTvaAcc..etTvaFac] then
                   BEGIN
                   St:=St+'(EXISTS(SELECT '+PreE+PreJ2+' FROM ECRITURE WHERE '
                         +'E_AUXILIAIRE=Q1.T_AUXILIAIRE AND (E_ECHEENC1<>0 OR E_ECHEENC2<>0 '
                         +'OR E_ECHEENC3<>0 OR E_ECHEENC4<>0) AND E_EDITEETATTVA="-")) AND ' ;
                   END ;
                 St:=St+'T_NATUREAUXI<>"NPC" ' ;
                 END ;
      msGenAna : BEGIN
                 St:=St+'S_AXE="'+Axe+'"' ;
                 END ;
      msAuxEcr : BEGIN
                 St:=St+'G_COLLECTIF="X" ' ;
                 END ;
      msSecAna : BEGIN
                 St:=St+'G_VENTILABLE="X" AND G_VENTILABLE'+Axe[2]+'="X" ';
                 END ;
      END ;
   Q.SQL.Add(St) ;
   St:=SQLModeSelectMul(':'+PreB1+PreJ1,'',PreB2) ;
   If St<>'' Then
      BEGIN
//      If ITypCpt=1 Then Q.SQL.Add('OR '+St) Else Q.SQL.Add('AND '+St) ;
      Q.SQL.Add('AND '+St) ;
      END ;
   St:='' ;
   if CritEDT.SJoker then St:='AND '+PreB2+PreJ2+' like "'+TraduitJoker(CritEdt.SCpt1)+'" ' Else
      BEGIN
      if CritEDT.SCpt1<>'' then St:='AND  '+PreB2+PreJ2+'>="'+CritEdt.SCpt1+'" ' ;
      if CritEDT.SCpt2<>'' then St:=St+'AND  '+PreB2+PreJ2+'<="'+CritEdt.SCpt2+'" ' ;
      END ;
   If St<>'' Then Q.SQL.Add(St) ;
   st:='' ; if LeSSauf<>'' then St:=LeSSqlSauf ; if st<>'' Then Q.SQL.Add(St) ;
   Q.SQL.Add(' Order By '+PreB2+PreJ2) ;
   END ;
end ;

Function TFQR.AlimSQLMulNat(Q : TQuery ; Niveau : integer ; tz1,tz2 : TZoomTable) : Boolean ;
Var St,PreJ1,PreJ2,PreE,LeSauf,LeSSauf,LeSQLSauf,LeSSQLSauf,Where1,Where2 : String ;
    leTz1,leTz2 : TZoomTable ;
    DD1,DD2 : TDateTime ;
    Valeur : String ;
begin
Result:=FALSE ;

Case FNatureEtat Of
   neBal : BEGIN
           LeSauf:=CritEDT.BAL.Sauf ; LeSSauf:=CritEDT.BAL.SSqlSauf ;
           LeSQLSauf:=CritEdt.Bal.SQLSauf ; LeSSQLSauf:=CritEdt.Bal.SSQLSauf ;
           END ;
   neGl : BEGIN
          LeSauf:=CritEDT.GL.Sauf ; LeSSauf:=CritEDT.GL.SSauf ;
          LeSQLSauf:=CritEdt.GL.SQLSauf ; LeSSQLSauf:=CritEdt.GL.SSQLSauf ;
          END ;
   neBalLib : BEGIN
          LeSauf:=CritEDT.BalLib.Sauf ; LeSSauf:=CritEDT.BalLib.SSauf ;
          LeSQLSauf:=CritEdt.BalLib.SQLSauf ; LeSSQLSauf:=CritEdt.BalLib.SSQLSauf ;
          END ;
   Else Exit ;
   END ;
PreE:='E_' ;
Case FModeSelect of
   msGenEcr : PreE:='E_' ;
   msGenAna : PreE:='Y_' ;
   END ;
leTz1:=tz1 ; leTz2:=tz2 ;
Q.CLose ;

If Niveau=0 Then
   BEGIN
   Q.Sql.Clear ;
   St:='Select NT_NATURE,NT_LIBELLE From NATCPTE Q Where NT_TYPECPTE="'+tzToNature(leTz1)+'" ' ;
   Q.SQL.Add(St) ;
   DD1:=CritEdt.DateDeb ; DD2:=CritEdt.DateFin ;
   preJ1:=tzToChampNature(letz1,FALSE) ;
   Valeur:=CritEdt.QualifPiece ;
   St:=WhatExisteNat(PreE,PreJ1,'NT_','NATURE',Valeur,V_PGI.Controleur,AvecRevision.State,DD1,DD2,CritEdt.Exo.Code,0,CritEdt.DEVPourExist) ;
   If St<>'' Then Q.SQL.Add(' AND '+St) ;
   St:='' ;
   if CritEdt.Joker then St:='AND NT_NATURE like "'+TraduitJoker(CritEdt.Cpt1)+'" ' Else
      BEGIN
      if CritEdt.Cpt1<>'' then St:='AND NT_NATURE >="'+CritEdt.Cpt1+'" ' ;
      if CritEdt.Cpt2<>'' then St:=St+'AND NT_NATURE <="'+CritEdt.Cpt2+'" ' ;
      END ;
   If St<>'' Then Q.SQL.Add(St) ;
   st:='' ; if LeSauf<>'' then St:=LeSqlSauf ; if st<>'' Then Q.SQL.Add(St) ;
   { Construction de la clause Order By de la SQL }
   Q.Sql.Add('Order By NT_TYPECPTE, NT_NATURE') ;
   (*
   If (CritEdt.Rupture=rLibres) then
      BEGIN
      Q.Sql.Add(WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,LeFb)) ;
      Q.Sql.Add('Order By '+OrderLibre(CritEdt.LibreTrie)+PreB1+PreJ1) ;
      END Else Q.SQL.Add(' Order By  '+PreB1+PreJ1) ;
   *)
   Result:=TRUE ;
   END ;
If Niveau=1 Then
   BEGIN
   Q.Sql.Clear ;
   St:='Select NT_NATURE,NT_LIBELLE From NATCPTE Q1 Where NT_TYPECPTE="'+tzToNature(leTz2)+'" ' ;
   Q.SQL.Add(St) ;
   DD1:=CritEdt.DateDeb ; DD2:=CritEdt.DateFin ;
   preJ1:=tzToChampNature(letz1,FALSE) ;
   preJ2:=tzToChampNature(letz2,FALSE) ;
   Valeur:=CritEdt.QualifPiece ;
   Where1:=':NT_NATURE' ; Where2:='NT_NATURE' ;
   St:=WhatExisteMulNat(PreE,PreJ1,PreJ2,'NT_',Valeur,V_PGI.Controleur,AvecRevision.State,
                        DD1,DD2,CritEdt.Exo.Code,Where1,Where2,0) ;
   If St<>'' Then Q.SQL.Add(' AND '+St) ;
   St:='' ;
   if CritEDT.SJoker then St:='AND NT_NATURE like "'+TraduitJoker(CritEdt.SCpt1)+'" ' Else
      BEGIN
      if CritEDT.SCpt1<>'' then St:='AND NT_NATURE >="'+CritEdt.SCpt1+'" ' ;
      if CritEDT.SCpt2<>'' then St:=St+'AND NT_NATURE<="'+CritEdt.SCpt2+'" ' ;
      END ;
   If St<>'' Then Q.SQL.Add(St) ;

   st:='' ; if LeSSauf<>'' then St:=LeSSqlSauf ; if st<>'' Then Q.SQL.Add(St) ;
   Q.SQL.Add(' Order By NT_NATURE') ;
   END ;
end ;

procedure TFQR.BFinEtatBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
TopReport.Enabled:=FALSE ;
BOTTOMREPORT.enabled:=FALSE ;
EnteteAutrePage.Enabled:=False ;
end;

procedure TFQR.BOTTOMREPORTAfterPrint(BandPrinted: Boolean);
begin
// XX Rony 10/04/97
TOPREPORT.Enabled:=BOTTOMREPORT.Enabled ;
end;

procedure TFQR.EnteteAutrePageBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
//TitreBarre.Width:=EnteteAutrePage.Width ;
end;

Function TFQR.SQLQuelCpt : String ;
Var COMPTE,St : String ;
BEGIN
COMPTE:='' ;
Case FNatureBase Of
  nbGen : COMPTE:='G_GENERAL' ;
  nbAux : COMPTE:='T_AUXILIAIRE' ;
  nbSec : COMPTE:='S_SECTION' ;
  nbJal : COMPTE:='J_JOURNAL' ;
  nbBudGen : COMPTE:='BG_GENERAL' ;
  nbBudSec : COMPTE:='BS_SECTION' ;
  nbBudJal : COMPTE:='BJ_JOURNAL' ;
  nbLibEcr : COMPTE:='NT_NATURE' ;
  END ;
St:='' ;
if CritEDT.Joker then St:='AND '+COMPTE+' like "'+TraduitJoker(CritEDT.Cpt1)+'" ' Else
   BEGIN
   if CritEDT.Cpt1<>'' then St:='AND '+COMPTE+'>="'+CritEDT.Cpt1+'" ' ;
   if CritEDT.Cpt2<>'' then St:=St+'AND '+COMPTE+'<="'+CritEDT.Cpt2+'" ' ;
   END ;
Result:=St ;
END ;

Procedure TFQR.GenereSQLBaseTS ;
Var ZoomRub : Boolean ;
    LeSauf,LeSQLSauf,Axe,NatCpt,St, StJalAuto : String ;
    StSelect1,stSelect2,stSelect3,StFrom,StCpt,StOrder,StWhere1,StWhere2,StWhere3 : String ;
//    ITypCpt : Integer ;
//    ERuptOnly : TRuptEtat ;
    NatureRupt : String3 ;
    P          : String ;
    EOnlyCompteAssocie : Boolean ;
    EPlansCorresp : Byte ;
    STTS : String ;
    STREGROUP : String ;
BEGIN
Q.Close ; Q.SQL.Clear ; //ERuptOnly:=Sans ;
ZoomRub:=FALSE ; StSelect1:='' ; StSelect2:='' ; StFrom:='' ; StCpt:='' ; StOrder:='' ;
StWhere1:='' ; StWhere2:='' ; stSelect3:='' ; StWhere3:='' ; EOnlyCompteAssocie:=FALSE ;
EPlansCorresp:=0 ; STTS:='' ; StRegroup:='' ;
Case FnatureEtat Of
   neBal : BEGIN
           Axe:=CritEdt.Bal.Axe ; {ITypCpt:=CritEDT.BAL.TypCpt ;} NatCpt:=CritEDT.NatureCpt ;
           LeSauf:=CritEDT.BAL.Sauf ; LeSQLSauf:=CritEdt.Bal.SQLSauf ;
           ZoomRub:=CritEdt.Bal.ZoomBalRub ; //ERuptOnly:=CritEdt.BAl.RuptOnly ;
           EOnlyCompteAssocie:=CritEdt.Bal.OnlyCptAssocie ;
           EPlansCorresp:=CritEDT.Bal.PlansCorresp ;
           STTS:=CritEdt.Bal.STTS ;
           StRegroup:=CritEdt.Bal.StRegroup ;
           END ;
   neGl : BEGIN
          Axe:=CritEdt.GL.Axe ; {ITypCpt:=CritEDT.GL.TypCpt ;} NatCpt:=CritEDT.NatureCpt ;
          LeSauf:=CritEDT.GL.Sauf ; LeSQLSauf:=CritEdt.GL.SQLSauf ; //ERuptOnly:=CritEdt.GL.RuptOnly ;
          EOnlyCompteAssocie:=CritEdt.GL.OnlyCptAssocie ;
          EPlansCorresp:=CritEDT.GL.PlansCorresp ;
          Case FNatureBase Of
            nbGen : BEGIN
                    stSelect2:=' ,G_SAUTPAGE, G_SOLDEPROGRESSIF, G_TOTAUXMENSUELS, G_CENTRALISABLE ' ;
                    END ;
            nbAux : BEGIN
                    stSelect2:=' ,T_SAUTPAGE, T_SOLDEPROGRESSIF, T_TOTAUXMENSUELS ' ;
                    END ;
            nbSec : BEGIN
                    stSelect2:=' ,S_SAUTPAGE, S_SOLDEPROGRESSIF, S_TOTAUXMENSUELS ' ;
                    END ;
            END ;
          END ;
   neCum : BEGIN
          Axe:=CritEdt.Cum.Axe ; {ITypCpt:=CritEDT.Cum.TypCpt ;} NatCpt:=CritEDT.NatureCpt ;
          LeSauf:=CritEDT.Cum.Sauf ; LeSQLSauf:=CritEdt.Cum.SQLSauf ;
          Case FNatureBase Of
            nbGen : BEGIN
                    stSelect2:=' ,G_SAUTPAGE, G_SOLDEPROGRESSIF ' ;
                    END ;
            nbAux : BEGIN
                    stSelect2:=' ,T_SAUTPAGE, T_SOLDEPROGRESSIF ' ;
                    END ;
            nbSec : BEGIN
                    stSelect2:=' ,S_SAUTPAGE, S_SOLDEPROGRESSIF ' ;
                    END ;
            nbJal : BEGIN
                    stSelect2:='' ;
                    END ;
            END ;
          END ;
   Else Exit ;
   END ;
Case FNatureBase Of
  nbGen : BEGIN
          Case CritEdt.Rupture of
            rLibres  : BEGIN
                       StSelect3:=','+OrderLibre(CritEdt.LibreTrie,TRUE) ;
                       END ;
            rCorresp : Case EPlansCorresp Of
                         1 : StSelect3:=', G_CORRESP1 ' ;
                         2 : StSelect3:=', G_CORRESP2 ' ;
                        END ;
            End ;
          StSelect1:='Select '+STTS+StSelect3+',G_GENERAL,G_LIBELLE,G_COLLECTIF ' ;
          StFrom:=' From ECRITURE Q LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL Where ' ;
          if CritEDT.NatureCpt<>'' then StWhere1:='AND G_NATUREGENE="'+CritEDT.NatureCpt+'" ' ;
          if (OkV2 and (V_PGI.Confidentiel<>'1')) then StWhere1:=StWhere1+'AND G_CONFIDENTIEL<>"1" ' ;
          Case CritEdt.Rupture of
            rLibres  : BEGIN
                       IF EOnlyCompteAssocie then StWhere1:=StWhere1+WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,fbGene,EOnlyCompteAssocie) ;
//                       StOrder:='Order By '+OrderLibre(CritEdt.LibreTrie)+'G_GENERAL' ;
                       StOrder:=' GROUP BY '+STTS+StSelect3+',G_GENERAL,G_LIBELLE,G_COLLECTIF' ;
                       StOrder:=StOrder+' ORDER BY '+STTS+StSelect3+',G_GENERAL,G_LIBELLE,G_COLLECTIF' ;
                       END ;
            rCorresp : BEGIN
                       StOrder:=' GROUP BY '+STTS+StSelect3+',G_GENERAL,G_LIBELLE,G_COLLECTIF' ;
                       StOrder:=StOrder+' ORDER BY '+STTS+StSelect3+',G_GENERAL,G_LIBELLE,G_COLLECTIF' ;
                       END ;
            Else BEGIN
                 StOrder:=' GROUP BY '+STTS+',G_GENERAL,G_LIBELLE,G_COLLECTIF' ;
                 StOrder:=StOrder+' ORDER BY '+STTS+',G_GENERAL,G_LIBELLE,G_COLLECTIF' ;
                 END ;
            END ;
            Case FNatureEtat Of
              neBal : BEGIN
                      StWhere3:='And E_DATECOMPTABLE>="'+usdatetime(CritEdt.Bal.Date21)+
                                '" And E_DATECOMPTABLE<="'+usdatetime(CritEdt.Date2)+'" ' ;
                      if FExercice.ItemIndex>0 then StWhere3:=StWhere3+' And E_EXERCICE="'+CritEdt.Exo.Code+'" ' ;
                      StWhere3:=StWhere3+TraduitNatureEcr(CritEdt.Qualifpiece, 'E_QUALIFPIECE', true, CritEdt.ModeRevision) ;
                      if CritEdt.Etab<>'' then StWhere3:=StWhere3+' And E_ETABLISSEMENT="'+CritEdt.Etab+'"' ;
                      if CritEdt.DeviseSelect<>'' then StWhere3:=StWhere3+' And E_DEVISE="'+CritEdt.DeviseSelect+'"' ;
                      END ;
              END ;
          END ;
  nbSec : BEGIN
          Case CritEdt.Rupture of
            rLibres  : BEGIN
                       StSelect3:=','+OrderLibre(CritEdt.LibreTrie,TRUE)+', S_SECTION,S_AXE ' ;
                       END ;
            rCorresp : Case EPlansCorresp Of
                         1 : StSelect3:=', S_CORRESP1, S_SECTION,S_AXE ' ;
                         2 : StSelect3:=', S_CORRESP2,S_SECTION,S_AXE ' ;
                        END ;
            rRuptures : StSelect3:=', S_SECTIONTRIE,S_AXE,S_SECTION ' ;
            Else StSelect3:=', S_SECTION, S_AXE ' ;
            End ;
          StSelect1:='Select '+STTS+StSelect3+', S_LIBELLE ' ;
          StFrom:=' From SECTION Q Where ' ;
          {b FP FQ16854}
          //StFrom:=' From ANALYTIQ Q LEFT JOIN SECTION ON Y_SECTION=S_SECTION Where ' ;
          StFrom:=' From ANALYTIQ Q LEFT JOIN SECTION ON '+TSQLAnaCroise.ChampSection(Axe)+'=S_SECTION Where ' ;
          {e FP FQ16854}
          StWhere1:=' AND S_AXE="'+Axe+'" ' ;
          if (OkV2 and (V_PGI.Confidentiel<>'1')) then StWhere1:=StWhere1+'AND S_CONFIDENTIEL<>"1" ' ;
          Case CritEdt.Rupture of
            rLibres  : BEGIN
                       IF EOnlyCompteAssocie then StWhere1:=StWhere1+WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,AxeTofb(Axe),EOnlyCompteAssocie) ;
//                       StOrder:='Order By '+OrderLibre(CritEdt.LibreTrie)+'G_GENERAL' ;
                       StOrder:=' GROUP BY '+STTS+StSelect3+',S_LIBELLE ' ;
                       StOrder:=StOrder+' ORDER BY '+STTS+StSelect3+',S_LIBELLE ' ;
                       END ;
            rCorresp : BEGIN
                       StOrder:=' GROUP BY '+STTS+StSelect3+',S_LIBELLE ' ;
                       StOrder:=StOrder+' ORDER BY '+STTS+StSelect3+',S_LIBELLE ' ;
                       END ;
            rRuptures : BEGIN
                       StOrder:=' GROUP BY '+STTS+StSelect3+',S_LIBELLE ' ;
                       StOrder:=StOrder+' ORDER BY '+STTS+StSelect3+',S_LIBELLE ' ;
                       END ;
            Else BEGIN
                 StOrder:=' GROUP BY '+STTS+StSelect3+',S_LIBELLE ' ;
                 StOrder:=StOrder+' ORDER BY '+STTS+StSelect3+',S_LIBELLE ' ;
                 END ;
            END ;
            Case FNatureEtat Of
              neBal : BEGIN
                      StWhere3:='And Y_DATECOMPTABLE>="'+usdatetime(CritEdt.Bal.Date21)+
                                '" And Y_DATECOMPTABLE<="'+usdatetime(CritEdt.Date2)+'" ' ;
                      if FExercice.ItemIndex>0 then StWhere3:=StWhere3+' And Y_EXERCICE="'+CritEdt.Exo.Code+'" ' ;
                      StWhere3:=StWhere3+TraduitNatureEcr(CritEdt.Qualifpiece, 'Y_QUALIFPIECE', true, CritEdt.ModeRevision) ;
                      if CritEdt.Etab<>'' then StWhere3:=StWhere3+' And Y_ETABLISSEMENT="'+CritEdt.Etab+'"' ;
                      if CritEdt.DeviseSelect<>'' then StWhere3:=StWhere3+' And Y_DEVISE="'+CritEdt.DeviseSelect+'"' ;
                      END ;
              END ;
          END ;
  nbAux : BEGIN
          Case CritEdt.Rupture of
            rLibres  : BEGIN
                       StSelect3:=','+OrderLibre(CritEdt.LibreTrie,TRUE) ;
                       END ;
            rCorresp : Case EPlansCorresp Of
                         1 : StSelect3:=', T_CORRESP1 ' ;
                         2 : StSelect3:=', T_CORRESP2 ' ;
                        END ;
            End ;
          StSelect1:='Select '+STTS+StSelect3+',T_AUXILIAIRE,T_LIBELLE ' ;
          StFrom:=' From ECRITURE Q LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE Where ' ;
          if CritEDT.NatureCpt<>'' then StWhere1:='AND T_NATUREAUXI="'+CritEDT.NatureCpt+'" ' ;
          if (OkV2 and (V_PGI.Confidentiel<>'1')) then StWhere1:=StWhere1+'AND T_CONFIDENTIEL<>"1" ' ;
          Case CritEdt.Rupture of
            rLibres  : BEGIN
                       IF EOnlyCompteAssocie then StWhere1:=StWhere1+WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,fbAux,EOnlyCompteAssocie) ;
//                       StOrder:='Order By '+OrderLibre(CritEdt.LibreTrie)+'G_GENERAL' ;
                       StOrder:=' GROUP BY '+STTS+StSelect3+',T_AUXILIAIRE, T_LIBELLE ' ;
                       StOrder:=StOrder+' ORDER BY '+STTS+StSelect3+',T_AUXILIAIRE, T_LIBELLE ' ;
                       END ;
            rCorresp : BEGIN
                       StOrder:=' GROUP BY '+STTS+StSelect3+',T_AUXILIAIRE, T_LIBELLE ' ;
                       StOrder:=StOrder+' ORDER BY '+STTS+StSelect3+',T_AUXILIAIRE, T_LIBELLE ' ;
                       END ;
            Else BEGIN
                 StOrder:=' GROUP BY '+STTS+',T_AUXILIAIRE, T_LIBELLE ' ;
                 StOrder:=StOrder+' ORDER BY '+STTS+',T_AUXILIAIRE, T_LIBELLE ' ;
                 END ;
            END ;
            Case FNatureEtat Of
              neBal : BEGIN
                      StWhere3:='And E_DATECOMPTABLE>="'+usdatetime(CritEdt.Bal.Date21)+
                                '" And E_DATECOMPTABLE<="'+usdatetime(CritEdt.Date2)+'" ' ;
                      if FExercice.ItemIndex>0 then StWhere3:=StWhere3+' And E_EXERCICE="'+CritEdt.Exo.Code+'" ' ;
                      StWhere3:=StWhere3+TraduitNatureEcr(CritEdt.Qualifpiece, 'E_QUALIFPIECE', true, CritEdt.ModeRevision) ;
                      if CritEdt.Etab<>'' then StWhere3:=StWhere3+' And E_ETABLISSEMENT="'+CritEdt.Etab+'"' ;
                      if CritEdt.DeviseSelect<>'' then StWhere3:=StWhere3+' And E_DEVISE="'+CritEdt.DeviseSelect+'"' ;
                      END ;
              END ;
          (*
          StSelect1:='Select T_AUXILIAIRE,T_LIBELLE,T_COLLECTIF, T_TOTDEBANO, T_TOTCREANO, T_TOTDEBE, T_TOTCREE, T_TOTDEBANON1, T_TOTCREANON1 ' ;
          Case CritEdt.Rupture of
            rLibres  : BEGIN
                       StSelect1:='Select T_AUXILIAIRE,T_LIBELLE,T_COLLECTIF, T_TOTDEBANO, T_TOTCREANO, T_TOTDEBE, T_TOTCREE, T_TOTDEBANON1, T_TOTCREANON1 ' ;
                       StSelect3:=', T_TABLE0, T_TABLE1, T_TABLE2, T_TABLE3, T_TABLE4, T_TABLE5, T_TABLE6, T_TABLE7, T_TABLE8, T_TABLE9 ' ;
                       END ;
            rCorresp : StSelect1:='Select T_AUXILIAIRE,T_LIBELLE,T_COLLECTIF, T_TOTDEBANO, T_TOTCREANO, T_TOTDEBE, T_TOTCREE, T_TOTDEBANON1, T_TOTCREANON1, T_CORRESP1, T_CORRESP2 ' ;
            End ;
          StFrom:=' From TIERS Q Where ' ;
          if CritEdt.NatureCpt<>'' then StWhere1:='AND T_NATUREAUXI="'+CritEdt.NatureCpt+'" ' ;
          if (OkV2 and (V_PGI.Confidentiel<>'1')) then StWhere1:=StWhere1+'AND T_CONFIDENTIEL<>"1" ' ;
          StOrder:=' Order By T_AUXILIAIRE' ;
          Case CritEdt.Rupture of
            rLibres  : BEGIN
                       IF EOnlyCompteAssocie then StWhere1:=StWhere1+WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,fbAux,EOnlyCompteAssocie) ;
                       StOrder:='Order By '+OrderLibre(CritEdt.LibreTrie)+'T_AUXILIAIRE' ;
                       END ;
            rCorresp : Case EPlansCorresp Of
                         1 : StOrder:=' Order By T_CORRESP1, T_AUXILIAIRE' ;
                         2 : StOrder:=' Order By T_CORRESP2, T_AUXILIAIRE' ;
                         Else StOrder:=' Order By T_AUXILIAIRE' ;
                         END ;
            End ;
          *)
          END ;
  nbJal : BEGIN
          StSelect1:='Select J_JOURNAL,J_LIBELLE,J_NATUREJAL ' ;
          StFrom:=' From JOURNAL Q  Where ' ;
          StOrder:=' Order By J_JOURNAL' ;
          If FNatureEtat=neCum Then
             BEGIN
             StWhere1:=' AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ODA" AND J_NATUREJAL<>"ANA" ' ;
             If Not CritEdt.Cum.AvecJalAN Then StWhere1:=StWhere1+' AND J_NATUREJAL<>"ANO" ' ;
             END ;
          if (OkV2 and (V_PGI.Confidentiel<>'1')) then
             BEGIN
             StJalAuto:=VH^.JalAutorises ;
             if StJalAuto<>'' then StWhere1:=StWhere1+' And '+AnalyseCompte(StJalAuto,fbJal,False,False) ;
             END ;
          END ;
  END ;
If StSelect1<>'' Then Q.SQL.Add(StSelect1) ;
If StSelect2<>'' Then Q.SQL.Add(StSelect2) ;
//If StSelect3<>'' Then Q.SQL.Add(StSelect3) ;
If StFrom<>'' Then Q.SQL.Add(StFrom) ;
Case FNatureBase Of
  nbGen : BEGIN St:=' G_GENERAL<>"'+W_W+'" ' ; Q.SQL.Add(St) ; END ;
  nbAux : BEGIN St:=' T_AUXILIAIRE<>"'+W_W+'" ' ; Q.SQL.Add(St) ; END ;
  nbSec : BEGIN St:=' S_SECTION<>"'+W_W+'" ' ; Q.SQL.Add(St) ; END ;
  nbJal : BEGIN St:=' J_JOURNAL<>"'+W_W+'" ' ; Q.SQL.Add(St) ; END ;
  END ;
St:=SQLQuelCpt ; If St<>'' Then Q.SQL.Add(St) ;
St:=' AND '+STTS+'<>"" ' ; If St<>'' Then Q.SQL.Add(St) ;
If StWhere1<>'' Then Q.SQL.Add(StWhere1) ;
If StWhere2<>'' Then Q.SQL.Add(StWhere2) ;
If StWhere3<>'' Then Q.SQL.Add(StWhere3) ;
if LeSauf<>'' then Q.SQL.Add(LeSQLSauf) ;
If CritEdt.FiltreSup<>'' Then Q.SQL.Add(' AND '+CritEdt.FiltreSup) ;
if StRegroup<>'' then
  Case FNatureBase of
     nbGen,nbGenT : BEGIN
                    St:=' Or '+AnalyseCompte(StRegroup,fbGene,FALSE,False) ; Q.SQL.Add(St) ;
                    END ;
     (*
     nbAux : St:=' And '+AnalyseCompte(FExcep.Text,fbAux,True,False) ;
     nbSec : St:=' And '+AnalyseCompte(FExcep.Text,AxeToFb(FNatureCpt.Value),True,False) ;
     nbJal : St:=' And '+AnalyseCompte(FExcep.Text,fbJal,True,False) ;
     nbBudGen : St:=' And '+AnalyseCompte(FExcep.Text,fbBudGen,True,False) ;
     nbBudSec : St:=' And '+AnalyseCompte(FExcep.Text,AxeToFbBud(FNatureCpt.Value),True,False) ;
     nbBudJal : St:=' And '+AnalyseCompte(FExcep.Text,fbBudJal,True,False) ;
     *)
     END ;

{Rony 23/09/97 : zoom à partir des Rubrique n'a pas de champ aprés le where ...}
if Not ZoomRub then St:=CritEDT.SQLPlusBase ; If St<>'' Then Q.SQL.Add(St) ;
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
                Q.SQL.Add(' AND CR_CORRESP='+P+'_CORRESP'+IntToStr(EPlansCorresp)+')') ;
                END ;
              END ;
  End ;
If StOrder<>'' Then Q.SQL.Add(StOrder) ;
ChangeSQL(Q);
Q.Open ;
END ;


Procedure TFQR.GenereSQLBase ;
Var ZoomRub : Boolean ;
    LeSauf,LeSQLSauf,Axe,NatCpt,St, StJalAuto : String ;
    StSelect1,stSelect2,stSelect3,StFrom,StCpt,StOrder,StWhere1,StWhere2 : String ;
//    ITypCpt : Integer ;
//    ERuptOnly : TRuptEtat ;
    NatureRupt : String3 ;
    P          : String ;
    EOnlyCompteAssocie : Boolean ;
    EPlansCorresp : Byte ;
    StRegroup : String ;
BEGIN
Q.Close ; Q.SQL.Clear ; //ERuptOnly:=Sans ;
ZoomRub:=FALSE ; StSelect1:='' ; StSelect2:='' ; StFrom:='' ; StCpt:='' ; StOrder:='' ;
StWhere1:='' ; StWhere2:='' ; stSelect3:='' ; EOnlyCompteAssocie:=FALSE ;
EPlansCorresp:=0 ; StRegroup:='' ;
Case FnatureEtat Of
   neBal : BEGIN
           Axe:=CritEdt.Bal.Axe ; {ITypCpt:=CritEDT.BAL.TypCpt ;} NatCpt:=CritEDT.NatureCpt ;
           LeSauf:=CritEDT.BAL.Sauf ; LeSQLSauf:=CritEdt.Bal.SQLSauf ;
           ZoomRub:=CritEdt.Bal.ZoomBalRub ; //ERuptOnly:=CritEdt.BAl.RuptOnly ;
           EOnlyCompteAssocie:=CritEdt.Bal.OnlyCptAssocie ;
           EPlansCorresp:=CritEDT.Bal.PlansCorresp ;
           StRegroup:=CritEdt.Bal.StRegroup ;
           END ;
   neGl : BEGIN
          Axe:=CritEdt.GL.Axe ; {ITypCpt:=CritEDT.GL.TypCpt ;} NatCpt:=CritEDT.NatureCpt ;
          LeSauf:=CritEDT.GL.Sauf ; LeSQLSauf:=CritEdt.GL.SQLSauf ; //ERuptOnly:=CritEdt.GL.RuptOnly ;
          EOnlyCompteAssocie:=CritEdt.GL.OnlyCptAssocie ;
          EPlansCorresp:=CritEDT.GL.PlansCorresp ;
          Case FNatureBase Of
            nbGen : BEGIN
                    stSelect2:=' ,G_SAUTPAGE, G_SOLDEPROGRESSIF, G_TOTAUXMENSUELS, G_CENTRALISABLE ' ;
                    END ;
            nbAux : BEGIN
                    stSelect2:=' ,T_SAUTPAGE, T_SOLDEPROGRESSIF, T_TOTAUXMENSUELS ' ;
                    END ;
            nbSec : BEGIN
                    stSelect2:=' ,S_SAUTPAGE, S_SOLDEPROGRESSIF, S_TOTAUXMENSUELS ' ;
                    END ;
            END ;
          END ;
   neCum : BEGIN
          Axe:=CritEdt.Cum.Axe ; {ITypCpt:=CritEDT.Cum.TypCpt ;} NatCpt:=CritEDT.NatureCpt ;
          LeSauf:=CritEDT.Cum.Sauf ; LeSQLSauf:=CritEdt.Cum.SQLSauf ;
          Case FNatureBase Of
            nbGen : BEGIN
                    stSelect2:=' ,G_SAUTPAGE, G_SOLDEPROGRESSIF ' ;
                    END ;
            nbAux : BEGIN
                    stSelect2:=' ,T_SAUTPAGE, T_SOLDEPROGRESSIF ' ;
                    END ;
            nbSec : BEGIN
                    stSelect2:=' ,S_SAUTPAGE, S_SOLDEPROGRESSIF ' ;
                    END ;
            nbJal : BEGIN
                    stSelect2:='' ;
                    END ;
            END ;
          END ;
   Else Exit ;
   END ;
Case FNatureBase Of
  nbGen : BEGIN
          StSelect1:='Select G_GENERAL,G_LIBELLE,G_COLLECTIF, G_TOTDEBANO, G_TOTCREANO, G_TOTDEBE, G_TOTCREE, G_TOTDEBANON1, G_TOTCREANON1 ' ;
          Case CritEdt.Rupture of
            rLibres  : BEGIN
                       StSelect1:='Select G_GENERAL,G_LIBELLE,G_COLLECTIF, G_TOTDEBANO, G_TOTCREANO, G_TOTDEBE, G_TOTCREE, G_TOTDEBANON1, G_TOTCREANON1 ' ;
                       StSelect3:=', G_TABLE0, G_TABLE1, G_TABLE2, G_TABLE3, G_TABLE4, G_TABLE5, G_TABLE6, G_TABLE7, G_TABLE8, G_TABLE9 ' ;
                       END ;
            rCorresp : StSelect1:='Select G_GENERAL,G_LIBELLE,G_COLLECTIF, G_TOTDEBANO, G_TOTCREANO, G_TOTDEBE, G_TOTCREE, G_TOTDEBANON1, G_TOTCREANON1, G_CORRESP1, G_CORRESP2 ' ;
            End ;
          StFrom:=' From GENERAUX Q Where ' ;
          if CritEDT.NatureCpt<>'' then StWhere1:='AND G_NATUREGENE="'+CritEDT.NatureCpt+'" ' ;
          if (OkV2 and (V_PGI.Confidentiel<>'1')) then StWhere1:=StWhere1+'AND G_CONFIDENTIEL<>"1" ' ;
          StOrder:=' Order By G_GENERAL' ;
          (*
          Case FNatureEtat Of
            neBal : BEGIN
            StWhere2('And E_DATECOMPTABLE>="'+usdatetime(CritEdt.GL.Date21)+'" And E_DATECOMPTABLE<="'+usdatetime(CritEdt.Date2)+'"') ;
            if FExercice.ItemIndex>0 then QEcr.SQL.Add(' And E_EXERCICE="'+CritEdt.Exo.Code+'" ') ;
                    END ;
            *)
          Case CritEdt.Rupture of
            rLibres  : BEGIN
                       IF EOnlyCompteAssocie then StWhere1:=StWhere1+WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,fbGene,EOnlyCompteAssocie) ;
                       StOrder:='Order By '+OrderLibre(CritEdt.LibreTrie)+'G_GENERAL' ;
                       END ;
            rCorresp : Case EPlansCorresp Of
                         1 : StOrder:=' Order By G_CORRESP1, G_GENERAL' ;
                         2 : StOrder:=' Order By G_CORRESP2, G_GENERAL' ;
                         Else StOrder:=' Order By G_GENERAL' ;
                         END ;
            End ;
          END ;
  nbAux : BEGIN
          StSelect1:='Select T_AUXILIAIRE,T_LIBELLE,T_COLLECTIF, T_TOTDEBANO, T_TOTCREANO, T_TOTDEBE, T_TOTCREE, T_TOTDEBANON1, T_TOTCREANON1 ' ;
          Case CritEdt.Rupture of
            rLibres  : BEGIN
                       StSelect1:='Select T_AUXILIAIRE,T_LIBELLE,T_COLLECTIF, T_TOTDEBANO, T_TOTCREANO, T_TOTDEBE, T_TOTCREE, T_TOTDEBANON1, T_TOTCREANON1 ' ;
                       StSelect3:=', T_TABLE0, T_TABLE1, T_TABLE2, T_TABLE3, T_TABLE4, T_TABLE5, T_TABLE6, T_TABLE7, T_TABLE8, T_TABLE9 ' ;
                       END ;
            rCorresp : StSelect1:='Select T_AUXILIAIRE,T_LIBELLE,T_COLLECTIF, T_TOTDEBANO, T_TOTCREANO, T_TOTDEBE, T_TOTCREE, T_TOTDEBANON1, T_TOTCREANON1, T_CORRESP1, T_CORRESP2 ' ;
            End ;
          StFrom:=' From TIERS Q Where ' ;
          if CritEdt.NatureCpt<>'' then StWhere1:='AND T_NATUREAUXI="'+CritEdt.NatureCpt+'" ' ;
          if (OkV2 and (V_PGI.Confidentiel<>'1')) then StWhere1:=StWhere1+'AND T_CONFIDENTIEL<>"1" ' ;
          StOrder:=' Order By T_AUXILIAIRE' ;
          Case CritEdt.Rupture of
            rLibres  : BEGIN
                       IF EOnlyCompteAssocie then StWhere1:=StWhere1+WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,fbAux,EOnlyCompteAssocie) ;
                       StOrder:='Order By '+OrderLibre(CritEdt.LibreTrie)+'T_AUXILIAIRE' ;
                       END ;
            rCorresp : Case EPlansCorresp Of
                         1 : StOrder:=' Order By T_CORRESP1, T_AUXILIAIRE' ;
                         2 : StOrder:=' Order By T_CORRESP2, T_AUXILIAIRE' ;
                         Else StOrder:=' Order By T_AUXILIAIRE' ;
                         END ;
            End ;
          END ;
  nbSec : BEGIN
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
          END ;
  nbJal : BEGIN
          StSelect1:='Select J_JOURNAL,J_LIBELLE,J_NATUREJAL ' ;
          StFrom:=' From JOURNAL Q  Where ' ;
          StOrder:=' Order By J_JOURNAL' ;
          If FNatureEtat=neCum Then
             BEGIN
             StWhere1:=' AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ODA" AND J_NATUREJAL<>"ANA" ' ;
             If Not CritEdt.Cum.AvecJalAN Then StWhere1:=StWhere1+' AND J_NATUREJAL<>"ANO" ' ;
             END ;
          if (OkV2 and (V_PGI.Confidentiel<>'1')) then
             BEGIN
             StJalAuto:=VH^.JalAutorises ;
             if StJalAuto<>'' then StWhere1:=StWhere1+' And '+AnalyseCompte(StJalAuto,fbJal,False,False) ;
             END ;
          END ;
  END ;
If StSelect1<>'' Then Q.SQL.Add(StSelect1) ;
If StSelect2<>'' Then Q.SQL.Add(StSelect2) ;
If StSelect3<>'' Then Q.SQL.Add(StSelect3) ;
If StFrom<>'' Then Q.SQL.Add(StFrom) ;
If FNatureBase=nbJal Then
   BEGIN
   St:=' J_JOURNAL<>"'+W_W+'" ' ; Q.SQL.Add(St) ;
   END Else
   BEGIN
   St:=SQLModeSelect ; If (St<>'') And (Not ZoomRub) Then Q.SQL.Add(St) ;
   {Rony 23/09/97 : zoom à partir des Rubrique n'a pas de champ aprés le where ...}
   if ZoomRub then Q.SQL.Add(CritEDT.SQLPlusBase) ;
   END ;
St:=SQLQuelCpt ; If St<>'' Then Q.SQL.Add(St) ;
If StWhere1<>'' Then Q.SQL.Add(StWhere1) ;
If StWhere2<>'' Then Q.SQL.Add(StWhere2) ;
if LeSauf<>'' then Q.SQL.Add(LeSQLSauf) ;
{Rony 23/09/97 : zoom à partir des Rubrique n'a pas de champ aprés le where ...}
if StRegroup<>'' then
  Case FNatureBase of
     nbGen,nbGenT : BEGIN
                    St:=' Or '+AnalyseCompte(StRegroup,fbGene,FALSE,False) ; Q.SQL.Add(St) ;
                    END ;
     (*
     nbAux : St:=' And '+AnalyseCompte(FExcep.Text,fbAux,True,False) ;
     nbSec : St:=' And '+AnalyseCompte(FExcep.Text,AxeToFb(FNatureCpt.Value),True,False) ;
     nbJal : St:=' And '+AnalyseCompte(FExcep.Text,fbJal,True,False) ;
     nbBudGen : St:=' And '+AnalyseCompte(FExcep.Text,fbBudGen,True,False) ;
     nbBudSec : St:=' And '+AnalyseCompte(FExcep.Text,AxeToFbBud(FNatureCpt.Value),True,False) ;
     nbBudJal : St:=' And '+AnalyseCompte(FExcep.Text,fbBudJal,True,False) ;
     *)
     END ;

if Not ZoomRub then St:=CritEDT.SQLPlusBase ; If St<>'' Then Q.SQL.Add(St) ;
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
ChangeSQL(Q);
Q.Open ;
END ;

procedure TFQR.BtitreCritAfterPrint(BandPrinted: Boolean);
begin
(*
BTitre.ForceNewPage:=TRUE ;
// Rony 4/06/97
EnteteAutrePage.Enabled:=False ; TOPREPORT.Enabled:=False ;
EntetePage.Enabled:=False      ;
piedpage.Enabled:=True ;
// Fin rony
*)
end;

(*
procedure TFQR.TrouveOverlay(Ena : Boolean) ;
Var i : integer ;
    T   : TComponent ;
    BO  : TQRBAnd ;
BEGIN
for i:=0 to Self.ComponentCount-1 do
    BEGIN
    T:=Components[i] ;
    if (T is TQRBAnd) then
       BEGIN
       BO:=TQRBand(T) ;
       If BO.BandType=rbOverlay Then BEGIN BO.Enabled:=Ena ; Exit ; END ;
       END ;
    END ;
END ;
*)
procedure TFQR.BtitreCritBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
(*
BTitre.ForceNewPage:=TRUE ;
BtitreOk:=FALSE ; EntetePage.Enabled:=FALSE ; BottomReport.enabled:=FALSE ;
TrouveOverlay(FALSE) ;
CadreCrit.Width:=BtitreCrit.Width ;
*)
end;

procedure TFQR.FormActivate(Sender: TObject);
begin
//If OkChargeCritMemoire Then SauveCritMemoire (QRCritMemoire,Pages);
If OkChargeCritMemoire Then SauveCritMemoire (QRCritMemoireDefaut,Pages);
OkChargeCritMemoire:=FALSE ;
end;


procedure TFQR.FLibre1DblClick(Sender: TObject);
Var StTri, StCod1, StCod2, St, StListe, Pref : String ;
    LeFb : TFichierBase ;
    I    : Byte ;
begin
if OkZoomEdt then Exit ;
if (FNatureEtat=neBal) or (FNatureEtat=neGl) or (FNatureEtat=neJU)
   or (FNatureEtat=neGLV) or (FNatureEtat=neBalBud) or (FNatureEtat=neGlBud) then
   BEGIN
   StCod1:=FLibre1.text ; StCod2:=FLibre2.text ; LeFb:=fbGene ;
   if FNatureBase=nbGen then BEGIN LeFb:=fbGene ; Pref:='G0' ; END Else
   if FNatureBase=nbAux then BEGIN LeFb:=fbAux ; Pref:='T0' ; END Else
   if FNatureBase=nbSec then BEGIN LeFb:=AxeToFb(FNatureCpt.Value) ; Pref:='S0' ; END else
   if FNatureBase=nbBudGen then BEGIN LeFb:=fbBudgen ; Pref:='B0' ; END else
   if FNatureBase=nbBudSec then BEGIN LeFb:=AxeToFbBud(FNatureCpt.Value) ; Pref:='D0' ; END ;
   If TEdit(Sender).Name='FLibTriPar' then
      BEGIN
      StTri:=TEdit(Sender).Text ;
      ChoixTriTableLibre(LeFb,StTri,True,FLibre1.Text) ;
      END Else
      BEGIN
      StTri:=FLibTriPar.Text ;
      IF FOnlyCptAssocie.Checked Then ChoixTableLibreSur(LeFb,StTri, StCod1, StCod2) Else ChoixTableLibre(LeFb,StTri, StCod1, StCod2) ;
      FLibre1.text:=StCod1 ; FLibre2.text:=StCod2 ;
      StListe:=StCod1 ; I:=0 ; StTri:='' ;
      While StListe<>'' do
            BEGIN
            St:=ReadTokenSt(StListe) ;
            if St='' then begin inc(i) ; continue ; end ;
            if (St[1]<>'#') and (St[1]<>'-') then StTri:=StTri+Pref+IntToStr(i)+';' ;
            Inc(i) ;
            END ;
      END ;
      FLibTriPar.Text:=StTri ;
   END ;
end;

Function TFQR.QuelfbPourRupture : TFichierBase ;
BEGIN
Result:=fbJal ;
Case Etat Of
  etBalAna,etGlAna,etGlAnaGen : Result:=AxeToFb(FNatureCpt.Value) ;
  // GC - 28/12/2001 - Ajout etBalGeneComp
  etBalGen,etBalGenAux,etGlGenAux,etGlGen,etGlGenAna,etBalGeneComp : Result:=fbGene ;
  etBalAux,etBalAuxGen,etBalAgeDev,etBalVen,etGlAge,etGlAuxGen,etGlAux,etGlVen : Result:=fbAux ;
  END ;
END ;

procedure TFQR.FRupturesClick(Sender: TObject);
Var Sauve : String ;
    leFb : TFichierBase ;
begin
leFb:=QuelfbPourRupture ;
QuiVisible ;
if leFb=fbJal then exit ;
Sauve:=Trim(FPlanRuptures.Value) ;
If FPlansCo.Checked then
   BEGIN
   CorrespToCombo(FPlanRuptures,leFb) ;
   If QrLoading And (Sauve<>'') Then FPlanRuptures.Value:=Sauve
                                Else If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
   END ;
If FRuptures.Checked then
   BEGIN
   FPlanRuptures.Reload ;
   If QrLoading And (Sauve<>'') Then FPlanRuptures.Value:=Sauve
                                Else If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
   END ;
end;

procedure TFQR.FSansRuptClick(Sender: TObject);
begin
TabRuptures.TabVisible:=Not FSansRupt.Checked ;
if TabRuptures.TabVisible then QuiVisible ;
end;

Procedure TFQR.QuiVisible ;
BEGIN
FGroupRuptures.Visible:=FRuptures.Checked or FPlansCo.Checked;
FGroupLibres.Visible:=FTablesLibres.Checked ;
FOnlyCptAssocie.Visible:=FTablesLibres.Checked Or FPlansCo.Checked ;
If Not QRLoadingCptAssocie Then
   BEGIN
   If Not FOnlyCptAssocie.Visible Then FOnlyCptAssocie.Checked:=FALSE Else FOnlyCptAssocie.Checked:=TRUE ;
   END ;
if Not FTablesLibres.Checked then
   BEGIN
   FLibre1.Text:='' ; FLibre2.Text:='' ; FLibTriPar.Text:='' ;
   END Else
   BEGIN
   If FSansRupt.Checked Then If Not FOnlyCptAssocie.Checked Then
      BEGIN
      FLibre1.Text:='' ; FLibre2.Text:='' ;
      END ;
   END ;
FSautPageRupt.Visible:=FALSE ;
If FAvecRupt.Checked Then
  Case Etat Of
    etBalAge,etBalAgeDev,etBalVen,etJuSold,etGlAge,etGlVen,etGlAux,etBalAux : FSautPageRupt.Visible:=TRUE ;
    END ;
If Not FSautPageRupt.Visible Then FSautPageRupt.Checked:=FALSE ;
END ;

procedure TFQR.ApercuClick(Sender: TObject);
begin
Reduire.Enabled := Apercu.Checked ;
FListe.Enabled := Apercu.Checked ;
if not Reduire.Enabled then Reduire.Checked:=False ;
if not FListe.Enabled then FListe.Checked:=False ;
end;

procedure TFQR.FormCreate(Sender: TObject);
begin
QRCritMemoire:=HTStringList.Create ;
QRCritMemoireDefaut:=HTStringList.Create ;
VertScrollBar.Position:=0 ; HorzScrollBar.Position:=0 ;
HorzScrollBar.Range:=0 ;HorzScrollBar.Visible:=FALSE ;
VertScrollBar.Range:=0 ;VertScrollBar.Visible:=FALSE ;
Pages.Top:=0 ;
ClientHeight:=Pages.Top+Pages.Height+HPB.Height+PanelFiltre.Height-1 ;
ClientWidth:=Pages.Left+Pages.Width ;
if V_PGI.OutLook then Pages.Align:=alClient ;

//RR 27/10/2004
//Changement du mode de l'AGLDesigning afin de
// ne pas perdre les titres et les noms de colonnes en visu, en mode Look2003
//RR en attente de solution AGLDesign := V_PGI.AGLDesigning ;
//RR V_PGI.AGLDesigning := true ;

  {JP 23/08/04 : FQ 14242 : nouvelle gestion des filtres}
  FListeByUser := TListByUser.Create(FFiltres, BFiltre, toFiltre);
  with FListeByUser do begin
    OnSelect  := InitSelectFiltre;
    OnInitGet := InitGetFiltre;
    OnInitAdd := InitAddFiltre;
    OnUpgrade := UpgradeFiltre;
    OnParams  := ParseParamsFiltre;
  end;
end;

procedure TFQR.FormDestroy(Sender: TObject);
begin
If QRCritMemoire<>NIL Then QRCritMemoire.Free ;
If QRCritMemoireDefaut<>NIL Then QRCritMemoireDefaut.Free ;
  {JP 23/08/04 : FQ 14242 : nouvelle gestion des filtres}
  if Assigned(FListeByUser) then FreeAndNil(FListeByUser);
end;

procedure TFQR.RajouteLabelRupture(NomBandRupt : String) ;
Var TC : TComponent ;
    T,LabelPrincipal : TQRLabel ;
    LeParent : TQRBAND ;
    i,LastTop : Integer ;
BEGIN
//If Not V_PGI.SAV Then Exit ; If Not V_PGI.VersionDev Then Exit ;
TC:=Self.FindComponent(NomBandRupt) ; If TC=NIL Then Exit ;
If Not (TC is TQRBAND) Then Exit ; LeParent:=TQRBAND(TC) ;
If HauteurBandeRuptIni=0 Then HauteurBandeRuptIni:=LeParent.Height ;
If (CritEdt.Rupture<>rLibres) Then Exit ;
TC:=Self.FindComponent('TCODRUPT') ; If TC=NIL Then Exit ;
If Not (TC is TQRLABEL) Then Exit ; LabelPrincipal:=TQRLABEL(TC) ;
TC:=Self.FindComponent('TCODRUPT9') ; If TC<>NIL Then Exit ;
LastTop:=LabelPrincipal.Top+LabelPrincipal.Height+2 ;
If HauteurBandeRuptIni=0 Then HauteurBandeRuptIni:=LeParent.Height ;
For i:=1 To 10 Do
  BEGIN
  T:=TQRLabel.Create(Self) ; T.Parent:=LeParent ; T.Name:='TCODRUPT'+IntToStr(i) ;
  T.Font.Style:=[] ; T.Font.Style:=[fsitalic] ; T.Font.Name:=LabelPrincipal.Font.Name ;
  T.Color:=LabelPrincipal.Color ; T.Font.Color:=LabelPrincipal.Font.Color ;
  T.Tag:=LabelPrincipal.Tag ; T.Alignment:=taRightJustify ; T.AutoSize:=FALSE ;
  T.SynColGroup:=LabelPrincipal.SynColGroup ; T.SynData:=LabelPrincipal.SynData ;
  T.SynCritere:=LabelPrincipal.SynCritere ;  T.SynTitreCol:=LabelPrincipal.SynTitreCol ;
  T.Font.Height:=LabelPrincipal.Font.Height ;
  T.Caption:='' ; T.Width:=LabelPrincipal.Width ; T.Left:=LabelPrincipal.Left ;
  T.Top:=LastTop ; T.SynLigne:=i ;
  LastTop:=T.Top+T.Height+2 ;
  END ;
END ;

procedure TFQR.BDetailCheckForSpace;
Var OkRupt : Boolean ;
    fb : tFichierBase ;
begin
If CritEdt.Rupture<>rLibres Then Exit ;
Case FNatureBase of
   nbGen,nbGenT : fb:=fbGene ;
   nbAux : fb:=fbAux ;
   nbSec : fb:=fbAxe1 ;
   nbBudGen : fb:=fbBudgen ;
   nbBudSec : fb:=fbBudSec1 ;
   Else Exit ;
   END ;
Case FNatureEtat Of
  neBal : BEGIN
          If (CritEdt.Bal.STTS<>'') And (BDetail.ForceNEwPage) Then Else BDetail.ForceNEwPage:=FALSE ;
          if Not CritEdt.Bal.OnlyCptAssocie then
             BEGIN
             OkRupt:=DansRuptLibre(Q,fb,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
             If OkRupt And QRSautPage Then BEGIN QRSautPage:=FALSE ; BDetail.ForceNewPage:=TRUE ; END ;
             END ;
          END ;
  neGL  : BEGIN
          BDetail.ForceNEwPage:=FALSE ;
          if Not CritEdt.GL.OnlyCptAssocie then
             BEGIN
             OkRupt:=DansRuptLibre(Q,fb,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
             If OkRupt And QRSautPage Then BEGIN QRSautPage:=FALSE ; BDetail.ForceNewPage:=TRUE ; END ;
             END ;
          END ;
  neJu  : BEGIN
          BDetail.ForceNEwPage:=FALSE ;
          if Not CritEdt.Bal.OnlyCptAssocie then
             BEGIN
             OkRupt:=DansRuptLibre(Q,fb,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
             If OkRupt And QRSautPage Then BEGIN QRSautPage:=FALSE ; BDetail.ForceNewPage:=TRUE ; END ;
             END ;
          END ;
  neGlV  : BEGIN
          BDetail.ForceNEwPage:=FALSE ;
          if Not CritEdt.GlV.OnlyCptAssocie then
             BEGIN
             OkRupt:=DansRuptLibre(Q,fb,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
             If OkRupt And QRSautPage Then BEGIN QRSautPage:=FALSE ; BDetail.ForceNewPage:=TRUE ; END ;
             END ;
          END ;
  neBalBud : BEGIN
             BDetail.ForceNEwPage:=FALSE ;
             if Not CritEdt.BalBud.OnlyCptAssocie then
                BEGIN
                OkRupt:=DansRuptLibre(Q,fb,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
                If OkRupt And QRSautPage Then BEGIN QRSautPage:=FALSE ; BDetail.ForceNewPage:=TRUE ; END ;
                END ;
             END ;
  Else BDetail.ForceNewPage:=FALSE ;
  END ;
end;

procedure TFQR.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;


Type Tab10Str = Array[1..10] Of String ;

(*======================================================================*)
procedure QuellePiece ( StAuto : String ; Var Tab : Tab10Str ) ;
Var StA,StC : String ;
    i    : integer ;
BEGIN
StA:=StAuto ; i:=0 ; Fillchar(Tab,SizeOf(Tab),#0) ;
Repeat
 StC:=ReadTokenSt(StA) ;
 if StC<>'' then
    BEGIN
    Inc(i) ; Tab[i]:=StC ;
    END ;
Until (StC='') ;
END ;

(*======================================================================*)
{$IFDEF AMORTISSEMENT}
procedure ZoomEdtImmo(Qui : Integer ; Quoi : String) ;
begin
  case Qui of
    91 : AMLanceFiche_FicheImmobilisation(Quoi,taConsult,'');
    92 : FicheGene(nil,'',Quoi,taConsult,0) ;
  end;
  Screen.Cursor:=crDefault ;
end;

Procedure ZoomEdtEtatImmo(Quoi : String) ;
var i,Qui: Integer ;
begin
  i := Pos(';',Quoi);
  Qui := StrToInt (Copy (Quoi,1,i-1));
  Quoi:=Copy(Quoi,i+1,Length(Quoi)-i) ;
  ZoomEdtImmo(Qui,Quoi) ;
end;
{$ENDIF}

Procedure ZoomEdtEtat(Quoi : String) ;
var i: Integer ;
BEGIN
i:=StrToInt(Copy(Quoi,1,1)) ;
Quoi:=Copy(Quoi,3,Length(Quoi)-2) ;
ZoomEdt(i,Quoi) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... : 18/08/2004
Suite ........ : - LG - 18/08/2004 - Suppression de la fct debutdemois pour 
Suite ........ : l'appel de la saisie bor, ne fct pas avec les exercices 
Suite ........ : decalees
Mots clefs ... : 
*****************************************************************}
Procedure ZoomEdt(Qui : Integer ; Quoi : String) ;
{$IFNDEF IMP}
Var St,St1 : String ;
    Lefb : TFichierBase ;
    ll,Lp : Integer ;
    Crit : TCritEdt ;
    D1,D2 : TDateTime ;
{$IFNDEF CCS3}
    Lp2 : Integer ;
{$ENDIF}
    Q : TQuery ;
    Trouv : Boolean ;
    Tab : Tab10Str ;
    M : RMvt ;
    R : RLettR ;
    DD : TDateTime ;
    JalP : String ;
    StNumP,StNumL : String ;
//{$IFNDEF IMP}
    P  : RParFolio ;
{$ENDIF}
BEGIN
{$IFNDEF IMP}
Case Qui Of
  1 : BEGIN
      St:=Copy(Quoi,1,VH^.Cpta[fbGene].Lg);
      FicheGene(Nil,'',St,taConsult,0);
      END ;
  2 : BEGIN
      St:=Copy(Quoi,1,VH^.Cpta[fbAux].Lg) ;
      FicheTiers(Nil,'',St,taConsult,1);
      END ;
  3 : BEGIN
        Lp:=Pos('@',Quoi) ;
        // GCO - 28/11/2006 - FQ 18927
        if Lp <> 0 Then
        begin
          St   := Quoi ;
          St1  := Copy(Quoi,Lp+2,2);
          LeFb := AxeToFb(St1);
          St:=Copy(Quoi,1,VH^.Cpta[Lefb].Lg) ;
        end
        else
        begin
          St  := ReadTokenSt(Quoi); // Section
          St1 := ReadTokenSt(Quoi); // Axe
        end;
        // FIN GCO
        FicheSection(Nil,St1,st,taConsult,0) ;
      END ;
  4 : BEGIN
      St:=Quoi ; ll:=Pos(' ',St) ; If ll>0 Then St:=Copy(St,1,ll-1) ;
      FicheJournal(Nil,'',st,taConsult,0) ;
      END ;
  5 : BEGIN { Zoom Pièce }
      QuellePiece(Quoi,tab) ;
      If PieceSurFolio(Tab[1]) Then
        BEGIN
        Q:=OpenSQL('Select E_DATECOMPTABLE from Ecriture where E_JOURNAL="'+Tab[1]+'"'
                  +' AND E_EXERCICE="'+Tab[4]+'"'
                  +' AND E_DATECOMPTABLE="'+Tab[2]+'"'
                  +' AND E_NUMEROPIECE='+Tab[3],True) ;
        If Not Q.Eof Then
          BEGIN
          DD:=Q.Fields[0].AsDateTime ;
          Ferme(Q) ;
          FillChar(P, Sizeof(P), #0) ;
          P.ParPeriode:=DateToStr(DD) ;
          P.ParCodeJal:=Tab[1] ;
          P.ParNumFolio:=Tab[3] ;
          P.ParNumLigne:=StrToInt(Tab[5]) ;
          ChargeSaisieFolio(P, taConsult) ;
          END Else Ferme(Q) ;
        END Else
        BEGIN
        Q:=OpenSQL('Select * from Ecriture where E_JOURNAL="'+Tab[1]+'"'
                  +' AND E_EXERCICE="'+Tab[4]+'"'
                  +' AND E_DATECOMPTABLE="'+Tab[2]+'"'
                  +' AND E_NUMEROPIECE='+Tab[3],True) ;
        Trouv:=Not Q.EOF ; if Trouv then M:=MvtToIdent(Q,fbGene,False) ; Ferme(Q) ;
        M.NumLigVisu:=StrToInt(Tab[5]) ;
  //      if Trouv then ZoomSaisie(M) ; ZoomSaisie n'est plus maintenue ...
        if Trouv then LanceSaisie(Nil,taConsult,M) ;
        END ;
      END ;
  6 : BEGIN { ZoomLettrage }
      QuellePiece(Quoi,tab) ;
      Q:=OpenSQL('Select E_GENERAL, E_AUXILIAIRE, E_DEVISE, E_LETTRAGE, E_LETTRAGEDEV from Ecriture where E_JOURNAL="'+Tab[1]+'"'
                +' AND E_EXERCICE="'+Tab[4]+'"'
                +' AND E_DATECOMPTABLE="'+Tab[2]+'"'
                +' AND E_NUMEROPIECE='+Tab[3]
                +' AND E_NUMLIGNE='+Tab[5],True) ;
      Trouv:=Not Q.EOF ;
      if Trouv then
         BEGIN
         FillChar(R,Sizeof(R),#0) ;
         R.General:=Q.Fields[0].AsString ; R.Auxiliaire:=Q.Fields[1].AsString ;
         R.CritDev:=Q.Fields[2].AsString ; R.GL:=NIL ; R.CritMvt:='' ; R.Appel:=tlMenu ;
         R.CodeLettre:=Q.Fields[3].AsString ; R.DeviseMvt:=R.CritDev ;
         R.LettrageDevise:=(Q.Fields[4].AsString='X') ;
         END ;
      Ferme(Q) ;
      If Trouv Then LettrageManuel(R,False,taConsult) ;
      END ;
  7 : BEGIN { Zoom Pièce analytique }
{$IFNDEF CCMP}
      QuellePiece(Quoi,tab) ;
      Q:=OpenSQL('Select * from ANALYTIQ where Y_JOURNAL="'+Tab[1]+'"'
                +' AND Y_EXERCICE="'+Tab[4]+'"'
                +' AND Y_DATECOMPTABLE="'+Tab[2]+'"'
                +' AND Y_NUMEROPIECE='+Tab[3]
                +' AND Y_NUMLIGNE='+Tab[5]
                {b FP FQ16854}
                //+' AND Y_AXE="'+Tab[6]+'"',True) ;
                +' AND '+TSQLAnaCroise.ConditionAxe(Tab[6]),True) ;
                {e FP FQ16854}
      Trouv:=Not Q.EOF ; if Trouv then M:=MvtToIdent(Q,fbSect,False) ; Ferme(Q) ;
      //M.NumLigVisu:=StrToInt(Tab[5]) ;
      if Trouv then LanceSaisieODA(Nil,taConsult,M) ;
{$ENDIF}
      END ;
  8 : BEGIN
{$IFNDEF CCMP}
{$IFNDEF CCS3}
      Lp:=Pos(' ',Quoi) ;
      St:=Copy(Quoi,1,Lp-1) ;
      FicheBudgene(Nil,'',St,taConsult,0) ;
{$ENDIF}
{$ENDIF}
      END ;
  9 : BEGIN
{$IFNDEF CCMP}
{$IFNDEF CCS3}
      Lp:=Pos('@',Quoi) ; Lp2:=Pos(' ',Quoi) ;
      If Lp=0 Then Exit ;
      St:=Quoi ; St1:=Copy(Quoi,Lp+2,2) ; //LeFb:=AxeToFbBud(St1) ;
      St:=Copy(Quoi,1,Lp2-1) ;
      FicheBudsect(Nil,St1,st,taConsult,0) ;
{$ENDIF}
{$ENDIF}
      END ;
  10 : BEGIN { Zoom Budecr }
{$IFNDEF CCMP}
{$IFNDEF CCS3}
      QuellePiece(Quoi,tab) ;
      Q:=OpenSQL('Select * from BUDECR where BE_BUDJAL="'+Tab[1]+'"'
                +' AND BE_EXERCICE="'+Tab[4]+'"'
                +' AND BE_DATECOMPTABLE="'+Tab[2]+'"'
                +' AND BE_NUMEROPIECE='+Tab[3],True) ;
      Trouv:=Not Q.EOF ;
      if Trouv then TrouveEtLanceSaisBud(Q,taconsult) ;
      Ferme(Q) ;
{$ENDIF}
{$ENDIF}
      END ;
  11 : BEGIN { Zoom Facture avec TP }
      QuellePiece(Quoi,tab) ;
      If VH^.OuiTP And ((Tab[1]=VH^.JalVTP) Or (Tab[1]=VH^.JalATP)) Then
        BEGIN
        Q:=OpenSQL('Select E_PIECETP from Ecriture where E_JOURNAL="'+Tab[1]+'"'
                  +' AND E_EXERCICE="'+Tab[4]+'"'
                  +' AND E_DATECOMPTABLE="'+Tab[2]+'"'
                  +' AND E_NUMEROPIECE='+Tab[3],True) ;
        If Not Q.Eof Then
          BEGIN
          St:=Q.Fields[0].AsString ;
          St1:=ReadTokenSt(St) ;
          St1:=ReadTokenSt(St) ; StNumP:=St1 ;
          St1:=ReadTokenSt(St) ; StNumL:=St1 ;
          St1:=ReadTokenSt(St) ;
          St1:=ReadTokenSt(St) ; JalP:=St1 ;
          Tab[1]:=JalP ; Tab[3]:=StNumP ; Tab[5]:=StNumL ;
          END ;
        Ferme(Q) ;
        END ;
      Q:=OpenSQL('Select * from Ecriture where E_JOURNAL="'+Tab[1]+'"'
                +' AND E_EXERCICE="'+Tab[4]+'"'
                +' AND E_DATECOMPTABLE="'+Tab[2]+'"'
                +' AND E_NUMEROPIECE='+Tab[3],True) ;
      Trouv:=Not Q.EOF ; if Trouv then M:=MvtToIdent(Q,fbGene,False) ; Ferme(Q) ;
      M.NumLigVisu:=StrToInt(Tab[5]) ;
      if Trouv then LanceSaisie(Nil,taConsult,M) ;
      END ;
  100 : BEGIN { GL pour justif centralisation }
{$IFNDEF CCMP}
{$IFNDEF NOQR}
        QuellePiece(Quoi,tab) ;
        Fillchar(Crit,SizeOf(Crit),#0) ;
        D1:=StrToDate(Tab[3]) ; D2:=StrToDate(Tab[4]) ;
        Crit.Date1:=D1 ; Crit.Date2:=D2 ;
        Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
        Crit.NatureEtat:=neGL ;
        InitCritEdt(Crit) ;
        Crit.GL.ForceNonCentralisable:=TRUE ;
        Crit.Cpt1:=Tab[1] ; Crit.Cpt2:=Crit.Cpt1 ;
        Crit.SQLPLUS:='AND E_JOURNAL="'+TAB[2]+'"' ;
        GLGeneralZoom(Crit) ;
{$ENDIF}
{$ENDIF}
        END ;
  END ;
Screen.Cursor:=crDefault ;
{$ENDIF}
END ;

procedure TFQR.FJokerKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var CtrlF5 : Boolean ;
BEGIN
CtrlF5:=(Shift=[ssCtrl]) And (Key=VK_F5) ;
If CtrlF5 Then GereZoneJoker(Sender,'FJoker',FCpte1) ;
inherited KeyDown(Key,Shift) ;
end;

procedure TFQR.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//RR V_PGI.AGLDesigning := AGLDesign ;
if Parent is THPanel then Action:=caFree ;
end;

procedure TFQR.FPlanRupturesChange(Sender: TObject);
Var Sauve1,Sauve2 : Integer ;
    leFb : TFichierBase ;
begin
If V_PGI.CergEF and (Etat=etBalGen) Then Exit ;
leFb:=QuelfbPourRupture ; if leFb=fbJal then exit ;
Sauve1:=FcodeRupt1.ItemIndex ; Sauve2:=FcodeRupt2.ItemIndex ;
if FPlansCo.Checked then CorrespToCodes(FPlanRuptures,FCodeRupt1,FCodeRupt2) Else
if FRuptures.Checked then RuptureToCodes(FPlanRuptures,FCodeRupt1,FCodeRupt2,leFb) ;
If QrLoading And (Sauve1<>-1) Then FCodeRupt1.ItemIndex:=Sauve1
                              Else FCodeRupt1.ItemIndex:=0 ;
If QrLoading And (Sauve2<>-1) Then FCodeRupt2.ItemIndex:=Sauve2
                              Else FCodeRupt2.ItemIndex:=FCodeRupt2.Items.Count-1 ;
end;

procedure TFQR.FCpte1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
If (Shift=[]) And (Key=187) Then
  BEGIN
  Key:=0 ; FCpte2.Text:=FCpte1.Text ;
  END ;
inherited KeyDown(Key,Shift) ;
end;

procedure TFQR.FCpte1KeyPress(Sender: TObject; var Key: Char);
begin
If Key='=' Then Key:=#0 ;
end;

// GC - 11/12/2001
procedure TFQR.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled := False;
//  Apercu.Checked := False;
  BValider.click;
  Close;
end;

// GC - 11/12/2001
procedure TFQR.InitEtatChaine(vCodeEtat: String);
begin
  // Positionnement du itemindex de la combo avec le nom du filtre
  FFiltres.ItemIndex := FFiltres.Items.Indexof(CritEdtChaine.FiltreUtilise);
  LoadFiltre(vCodeEtat, FFiltres, Pages);

  if CritEdtChaine.UtiliseCritStd then
  begin
    // Chargement des critères standards des états chainés
    FExercice.Value   := CritEdtChaine.Exercice.Code;
    FDateCompta1.Text := DateToStr(CritEdtChaine.Exercice.Deb);
    FDateCompta2.Text := DateToStr(CritEdtChaine.Exercice.Fin);
    FSelectCpte.Value := CritEdtChaine.ModeSelection;

    if CritEdtChaine.NatureCompte <> FNatureCpt.Value then
      FNatureCpt.Value  := CritEdtChaine.NatureCompte;
  end;

end;

procedure TFQR.OnCumClick(Sender: TObject);
begin
If EstSpecif('51187') Then
  BEGIN
  VH^.UseTC:=OnCum.Checked ;
  END ;
end;

procedure TFQR.CBLibClick(Sender: TObject);
begin
ChargeAvances(CBLib.Checked) ;
end;

procedure TFQR.CTri1Click(Sender: TObject);
Var TCB : TCheckBox ;
    i,j : Integer ;
begin
  inherited;
If Sender<>NIL Then
  BEGIN
  TCB:=TCheckBox(Sender) ; If Not TCB.Checked Then BEGIN FSautPageTRI.Enabled:=FALSE ; FSautPageTRI.Checked:=FALSE ; Exit ; END ;
  FSautPageTRI.Enabled:=TRUE ;
  j:=StrToInt(Copy(TCB.Name,5,1)) ;
  For i:=1 To 6 Do
    BEGIN
    If i<>j Then
      BEGIN
      TCB:=TCheckBox(FindComponent('CTri'+IntToStr(i))) ;
      If TCB<>NIL Then TCB.Checked:=FALSE ;
      END ;
    END ;
  END ;
end;

procedure TFQR.Z_C1Change(Sender: TObject);
Var i : Integer ;
    TCB : TCheckBox ;
begin
If Sender<>NIL Then
  BEGIN
  i:=StrToInt(Copy(THValComboBox(Sender).Name,4,1)) ;
  TCB:=TCheckBox(FindComponent('CTri'+IntToStr(i))) ;
  If TCB=Nil Then Exit ;
  If THValComboBox(Sender).Itemindex<0 Then
    BEGIN
    TCB.Enabled:=FALSE ; TCB.Checked:=FALSE ;
    END Else TCB.Enabled:=TRUE ;
  END ;
end;

Function TFQR.RecupVCS(QV : TVariantField) : Variant ;
BEGIN
Result:=null ;
If CritEdt.bal.STTS<>'' Then Result:=QV.AsVariant ;
END ;

Function TransCompQR(Ch,Comp,Val : String) : String ;
Var St,typ,Val1,Val2 : String ;
BEGIN
if (Ch='') or (Comp='') or (Val='') then BEGIN result:='' ; Exit ; END ;
if (Comp='I') or  (Comp='J') then St:='' else
if (Comp='E') or (Comp='G') then St:=' BETWEEN ' else
if ((Comp='C') or (Comp='L')) then St:=' LIKE ' else
if ((Comp='D') or (Comp='M')) then St:=' NOT LIKE ' else St:=Comp ;

typ:=ChampToType(Ch) ;
if (Comp='E') or (Comp='G') then
   BEGIN
   Val:=FindetReplace(Val,':',';',TRUE) ;
   Val:=FindetReplace(Val,' et ',';',TRUE) ;
   Val:=FindetReplace(Val,' ET ',';',TRUE) ;
   Val:=FindetReplace(Val,' Et ',';',TRUE) ;
   Val1:=ReadTokenSt(Val) ; Val2:=ReadTokenSt(Val) ;
   if Val2='' then Val2:=Val1 ;
   if (Typ='INTEGER') or (Typ='SMALLINT') then St:=St+IntToStr(ValeurI(Val1))+' AND '+IntToStr(ValeurI(Val2)) else
   if (Typ='DOUBLE') or (Typ='RATE')  then St:=St+StrfPoint(Valeur(Val1))+' AND '+StrfPoint(Valeur(Val2)) else
   if Typ='DATE'    then St:=St+'"'+USDATETIME(StrToDate(Val1))+'" AND "'+USDATETIME(StrToDate(Val2))+'"' else
      BEGIN
      St:=St+'"'+Val1+'" AND "'+Val2+'"' ;
      END ;
   if Comp='G' then St:=' NOT '+St ;
   END else
if (Comp='I') or (Comp='J') then
   BEGIN
   Val:=FindetReplace(Val,' ou ',';',TRUE) ;
   Val:=FindetReplace(Val,' OU ',';',TRUE) ;
   Val:=FindetReplace(Val,' Ou ',';',TRUE) ;
   Val2:='' ;
   While Val<>'' do
      BEGIN
      Val1:=ReadTokenSt(Val) ;
      if Val1<>'' then
         BEGIN
         if (Typ='INTEGER') or (Typ='SMALLINT') then St:=St+IntToStr(ValeurI(Val1))+', ' else
         if (Typ='DOUBLE') or (Typ='RATE')  then St:=St+StrfPoint(Valeur(Val1))+' ,' else
         if Typ='DATE'    then St:=St+'"'+USDATETIME(StrToDate(Val1))+'", ' else
                    St:=St+'"'+Val1+'", ' ;
         Val2:='Ok' ;
         END ;
      END ;
   if Val2<>'' then
      BEGIN
      system.Delete(St,Length(St)-1,2) ;
      St:=' IN ('+St+')' ;
      if Comp='J' then St:=' NOT '+St ;
      END
   END else
   BEGIN
   if (Typ='INTEGER') or (Typ='SMALLINT') then St:=St+IntToStr(ValeurI(Val)) else
   if (Typ='DOUBLE') or (Typ='RATE')  then St:=St+StrfPoint(Valeur(Val)) else
   if Typ='DATE'    then St:=St+'"'+USDATETIME(StrToDate(Val))+'"' else
      BEGIN
      St:=St+'"' ;
      if ((Comp='L') or (Comp='M')) then St:=St+'%' ;
      St:=St+CheckdblQuote(Val) ;
      if ((Comp='C') or (Comp='L') or (Comp='D') or (Comp='M')) then St:=St+'%' ;
      St:=St+'"' ;
      END ;
   END ;
Result:=Ch+St ;
END ;

Function TFQR.FaitSQLCompQR : String ;
Var StComp : Array[1..6] Of String ;
    G : Array[1..6] Of Boolean ;
    i : Integer ;
    Z_C,ZO : THvalComboBox ;
    ZV : tEdit ;
    ZG : TComboBox ;
    OkAvance : Boolean ;
    St : String ;
BEGIN
Result:='' ;
Fillchar(StComp,SizeOf(StComp),#0) ; Fillchar(G,SizeOf(G),#0) ;
OkAvance:=FALSE ;
For i:=1 To 6 Do
  BEGIN
  Z_C:=THvalComboBox(FindComponent('Z_C'+IntToStr(i))) ;
  ZO:=THvalComboBox(FindComponent('ZO'+IntToStr(i))) ;
  ZV:=TEdit(FindComponent('ZV'+IntToStr(i))) ;
  If i<6 Then ZG:=THvalComboBox(FindComponent('ZG'+IntToStr(i))) Else ZG:=Nil ;
  If (Z_C<>NIL) And (ZO<>NIL) And (ZV<>NIL) Then
    BEGIN
    StComp[i]:=TransCompQR(Z_C.Value,ZO.Value,ZV.Text) ;
    If Trim(StComp[i])<>'' Then OkAvance:=TRUE ;
    END ;
  If ZG<>NIL Then G[i]:=(ZG.ItemIndex=1) ;
  END ;
If Not OkAvance Then Exit ;
St:='' ;
If ((StComp[1]<>'') and (StComp[2]='') and (StComp[3]='') and (StComp[4]='') and (StComp[5]='')) then
   BEGIN
   if G[1] then
      BEGIN
      if St<>'' then St:='('+St+')' ;
      St:=St+' OR  ('+StComp[1]+')' ;
      END else St:=St+' AND ('+StComp[1]+')' ;
   END else
If ((StComp[1]<>'') and (StComp[2]<>'') and (StComp[3]='') and (StComp[4]='') and (StComp[5]='')) then
   BEGIN
   if G[1] then St:=St+' AND ('+StComp[1]+' OR '+StComp[2]+')'
           else St:=St+' AND ('+StComp[1]+' AND '+StComp[2]+')' ;
   END else
If ((StComp[1]<>'') and (StComp[2]<>'') and (StComp[3]<>'') and (StComp[4]='') and (StComp[5]='')) then
   BEGIN
   if G[1] then
      BEGIN
      if G[2] then St:=St+' AND ('+StComp[1]+' OR '+StComp[2]+' OR '+StComp[3]+')'
              else St:=St+' AND ('+StComp[1]+' OR ('+StComp[2]+' AND '+StComp[3]+'))' ;
      END else
      BEGIN
      if G[2] then St:=St+' AND ('+StComp[1]+' AND ('+StComp[2]+' OR '+StComp[3]+'))'
              else St:=St+' AND ('+StComp[1]+' AND '+StComp[2]+' AND '+StComp[3]+')' ;
      END ;
   END Else
If ((StComp[1]<>'') and (StComp[2]<>'') and (StComp[3]<>'') and (StComp[4]<>'') and (StComp[5]='')) then
   BEGIN
   if G[1] then
      BEGIN
      if G[2] then
        BEGIN
        If G[3] Then St:=St+' AND ('+StComp[1]+' OR '+StComp[2]+' OR '+StComp[3]+' OR '+StComp[4]+')'
                Else St:=St+' AND ('+StComp[1]+' OR '+StComp[2]+' OR ('+StComp[3]+' AND '+StComp[4]+'))' ;
        END Else
        BEGIN
        If G[3] Then St:=St+' AND ('+StComp[1]+' OR '+StComp[2]+' AND ('+StComp[3]+' OR '+StComp[4]+'))'
                Else St:=St+' AND ('+StComp[1]+' OR ('+StComp[2]+' AND '+StComp[3]+' AND '+StComp[4]+'))' ;
        END ;
      END else
      BEGIN
      if G[2] then
        BEGIN
        If G[3] Then St:=St+' AND ('+StComp[1]+' AND ('+StComp[2]+' OR '+StComp[3]+' OR '+StComp[4]+'))'
                Else St:=St+' AND ('+StComp[1]+' AND '+StComp[2]+' OR ('+StComp[3]+' AND '+StComp[4]+'))' ;
        END Else
        BEGIN
        If G[3] Then St:=St+' AND ('+StComp[1]+' AND '+StComp[2]+' AND ('+StComp[3]+' OR '+StComp[4]+'))'
                Else St:=St+' AND ('+StComp[1]+' AND '+StComp[2]+' AND '+StComp[3]+' AND '+StComp[4]+')' ;
        END ;
      END ;
   END Else
If ((StComp[1]<>'') and (StComp[2]<>'') and (StComp[3]<>'') and (StComp[4]<>'') and (StComp[5]<>'')) then
   BEGIN
   if G[1] then
      BEGIN
      if G[2] then
        BEGIN
        If G[3] Then
          BEGIN
          If G[4] Then St:=St+' AND ('+StComp[1]+' OR '+StComp[2]+' OR '+StComp[3]+' OR '+StComp[4]+' OR '+StComp[5]+')'
                  Else St:=St+' AND ('+StComp[1]+' OR '+StComp[2]+' OR '+StComp[3]+' OR ('+StComp[4]+' AND '+StComp[5]+'))' ;
          END Else
          BEGIN
          If G[4] Then St:=St+' AND ('+StComp[1]+' OR '+StComp[2]+' OR '+StComp[3]+' AND ('+StComp[4]+' OR '+StComp[5]+'))'
                  Else St:=St+' AND ('+StComp[1]+' OR '+StComp[2]+' OR ('+StComp[3]+' AND '+StComp[4]+' AND '+StComp[5]+'))' ;
          END ;
        END Else
        BEGIN
        If G[3] Then
          BEGIN
          If G[4] Then St:=St+' AND ('+StComp[1]+' OR ('+StComp[2]+' AND ('+StComp[3]+' OR '+StComp[4]+' OR '+StComp[5]+')))'
                  Else St:=St+' AND ('+StComp[1]+' OR ('+StComp[2]+' AND '+StComp[3]+' OR ('+StComp[4]+' AND '+StComp[5]+')))' ;
          END Else
          BEGIN
          If G[4] Then St:=St+' AND ('+StComp[1]+' OR ('+StComp[2]+' AND '+StComp[3]+' AND ('+StComp[4]+' OR '+StComp[5]+')))'
                  Else St:=St+' AND ('+StComp[1]+' OR ('+StComp[2]+' AND ('+StComp[3]+' AND '+StComp[4]+' AND '+StComp[5]+'))' ;
          END ;
        END ;
      END else
      BEGIN
      if G[2] then
        BEGIN
        If G[3] Then
          BEGIN
          If G[4] Then St:=St+' AND ('+StComp[1]+' AND ('+StComp[2]+' OR '+StComp[3]+' OR '+StComp[4]+' OR '+StComp[5]+'))'
                  Else St:=St+' AND ('+StComp[1]+' AND ('+StComp[2]+' OR '+StComp[3]+' OR ('+StComp[4]+' AND '+StComp[5]+')))' ;
          END Else
          BEGIN
          If G[4] Then St:=St+' AND ('+StComp[1]+' AND ('+StComp[2]+' OR '+StComp[3]+' AND ('+StComp[4]+' OR '+StComp[5]+')))'
                  Else St:=St+' AND ('+StComp[1]+' AND ('+StComp[2]+' OR ('+StComp[3]+' AND '+StComp[4]+' AND '+StComp[5]+')))' ;
          END ;
        END Else
        BEGIN
        If G[3] Then
          BEGIN
          If G[4] Then St:=St+' AND ('+StComp[1]+' AND ('+StComp[2]+' AND ('+StComp[3]+' OR '+StComp[4]+' OR '+StComp[5]+')))'
                  Else St:=St+' AND ('+StComp[1]+' AND '+StComp[2]+' AND '+StComp[3]+' OR ('+StComp[4]+' AND '+StComp[5]+'))' ;
          END Else
          BEGIN
          If G[4] Then St:=St+' AND ('+StComp[1]+' AND '+StComp[2]+' AND '+StComp[3]+' AND ('+StComp[4]+' OR '+StComp[5]+'))'
                  Else St:=St+' AND ('+StComp[1]+' AND '+StComp[2]+' AND ('+StComp[3]+' AND '+StComp[4]+' AND '+StComp[5]+')' ;
          END ;
        END ;
      END ;
   END ;

If Trim(St)<>'' Then Delete(St,1,5) ;
Result:=St ;
END ;

Function TFQR.GetPref : String ;
BEGIN
Result:='' ;
Case FModeSelect of
   msGenEcr,msAuxEcr : Result:='E' ;
   msGenAna,msSecAna : Result:='Y' ;
   End ;
END ;

Function TFQR.GetNomChampSTTS : String ;
Var i,NC : integer ;
    St{,St1} : String ;
		Mcd : IMCDServiceCOM;
		Field     : IFieldCOM ;
BEGIN
MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();
//
Result:='' ; St:=Getpref ;
If St='' Then Exit ; IF CritEdt.Bal.STTS='' Then Exit ;
Field := Mcd.GetField(CritEdt.Bal.STTS);
Result:=Field.libelle ;
END ;


procedure TFQR.bEffaceAvanceClick(Sender: TObject);
begin
Z_C1.ItemIndex:=-1 ;ZO1.ItemIndex:=-1 ; ZV1.text:='' ; ZG1.ItemIndex:=-1 ;
Z_C2.ItemIndex:=-1 ;ZO2.ItemIndex:=-1 ; ZV2.text:='' ; ZG2.ItemIndex:=-1 ;
Z_C3.ItemIndex:=-1 ;ZO3.ItemIndex:=-1 ; ZV3.text:='' ; ZG3.ItemIndex:=-1 ;
Z_C4.ItemIndex:=-1 ;ZO4.ItemIndex:=-1 ; ZV4.text:='' ; ZG4.ItemIndex:=-1 ;
Z_C5.ItemIndex:=-1 ;ZO5.ItemIndex:=-1 ; ZV5.text:='' ; ZG5.ItemIndex:=-1 ;
Z_C6.ItemIndex:=-1 ;ZO6.ItemIndex:=-1 ; ZV6.text:='' ;
end;

{---------------------------------------------------------------------------------------}
procedure TFQR.POPFPopup(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  UpdatePopFiltre( BSaveFiltre, BDelFiltre, BRenFiltre, FFiltres) ;
end;

{---------------------------------------------------------------------------------------}
procedure TFQR.BCreerFiltreClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  FListeByUser.Creer;
  FListeByUser.Save;
end;

{---------------------------------------------------------------------------------------}
procedure TFQR.BSaveFiltreClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  FListeByUser.Save;
end;

{---------------------------------------------------------------------------------------}
procedure TFQR.BDupFiltreClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  FListeByUser.Duplicate;
  FListeByUser.Save;
end;

{---------------------------------------------------------------------------------------}
procedure TFQR.BDelFiltreClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  VideFiltre(FFiltres, Pages) ;
  FListeByUser.Delete;
  FListeByUser.Save;
end;

{---------------------------------------------------------------------------------------}
procedure TFQR.BRenFiltreClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  FListeByUser.Rename;
  FListeByUser.Save;
end;

{---------------------------------------------------------------------------------------}
procedure TFQR.InitAddFiltre(T: TOB);
{---------------------------------------------------------------------------------------}
var
  Lines : HTStrings ;
begin
  Lines := HTStringList.Create ;
  SauveCritMemoire(Lines, Pages) ;
  FListeByUser.AffecteTOBFiltreMemoire(T, Lines);
  Lines.Free ;
end;

{---------------------------------------------------------------------------------------}
procedure TFQR.InitGetFiltre(T : TOB);
{---------------------------------------------------------------------------------------}
var
  Lines : HTStrings ;
begin
  Lines := HTStringList.Create ;
  SauveCritMemoire(Lines, Pages) ;
  FListeByUser.AffecteTOBFiltreMemoire(T, Lines);
  Lines.Free ;
end;

{---------------------------------------------------------------------------------------}
procedure TFQR.InitSelectFiltre(T : TOB);
{---------------------------------------------------------------------------------------}
var
  Lines : HTStrings ;
  i : integer;
  stChamp, stVal : string;
begin
  if T = nil then Exit;
  Lines := HTStringList.Create ;
  for i := 0 to T.Detail.Count - 1 do begin
    stChamp := T.Detail[i].GetValue('N');
    stVal   := T.Detail[i].GetValue('V');
    Lines.Add(stChamp + ';' + stVal);
  end;
  VideFiltre (FFiltres, Pages, False);

  ChargeCritMemoire(Lines, Pages);
  Lines.Free ;

  {JP 28/06/06 : FQ 16149 : après le chargement du filtre, on s'assure que l'établissement
                 reste en cohérence avec les restrictions utilisateurs}
  ControlEtab;
end;

{---------------------------------------------------------------------------------------}
procedure TFQR.ParseParamsFiltre(Params : HTStrings);                                 
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  FListeByUser.AddVersion;
  T := FListeByUser.Add;
  //en position 0 de Params se trouve le nom du filtre
  T.PutValue('NAME', XMLDecodeSt(Params[0])) ;
  T.PutValue('USER','---') ;
  Params.Delete(0);
  FListeByUser.AffecteTOBFiltreMemoire(T, Params);
end;

{---------------------------------------------------------------------------------------}
procedure TFQR.UpgradeFiltre( T: TOB);
{---------------------------------------------------------------------------------------}
begin

end;

{---------------------------------------------------------------------------------------}
procedure TFQR.FFiltresChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  QRloading:=TRUE ; QRLoadingExo:=TRUE ; QRLoadingCptAssocie:=TRUE ;
  InitVarInterne ;
  {JP 23/08/04 : FQ 14242 : Nouvelle gestion des filtres}
  FListeByUser.Select(FFiltres.Value);
  QRloading:=FALSE ; QRLoadingExo:=FALSE ; QRLoadingCptAssocie:=FALSE ;
end;

{---------------------------------------------------------------------------------------}
procedure TFQR.BNouvRechClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  QRloading:=TRUE ; QRLoadingExo:=TRUE ; QRLoadingCptAssocie:=TRUE ;
  If FJoker.Visible then FJoker.Text:='' ;
  InitVarInterne ;

  {JP 23/08/04 : FQ 14242 : Nouvelle gestion des filtres}
  VideFiltre(FFiltres, Pages);
  FListeByUser.New;

  QRloading:=FALSE ; QRLoadingExo:=FALSE ; QRLoadingCptAssocie:=FALSE ;
end;

{---------------------------------------------------------------------------------------}
procedure TFQR.GereEtablissement;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(FEtab) then begin
    {Si l'on ne gère pas les établissement ...}
    if not VH^.EtablisCpta  then begin
      {... on affiche l'établissement par défaut}
      FEtab.Value := VH^.EtablisDefaut;
      {... on désactive la zone}
      FEtab.Enabled := False;
    end

    {On gère l'établisement, donc ...}
    else begin
      {... On commence par regarder les restrictions utilisateur}
      PositionneEtabUser(FEtab);
      {... s'il n'y a pas de restrictions, on reprend le paramSoc}
      if FEtab.Value = '' then begin
        {... on affiche l'établissement par défaut}
        FEtab.Value := VH^.EtablisDefaut;
        {... on active la zone}
        FEtab.Enabled := True;
      end;
    end;
  end;

end;

{---------------------------------------------------------------------------------------}
procedure TFQR.ControlEtab;
{---------------------------------------------------------------------------------------}
var
  Eta : string;
begin
  if not Assigned(FEtab) then Exit;
  {S'il n'y a pas de gestion des établissement, logiquement, on ne force pas l'établissement !!!}
  if not VH^.EtablisCpta then Exit;
  
  Eta := EtabForce;
  {S'il y a une restriction utilisateur et qu'elle ne correspond pas au contenu de la combo ...}
  if (Eta <> '') and (Eta <> FEtab.Value) then begin
    {... on affiche l'établissement des restrictions}
    FEtab.Value := Eta;
    {... on désactive la zone}
    FEtab.Enabled := False;
  end;
end;

end.

