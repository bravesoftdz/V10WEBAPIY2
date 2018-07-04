unit SuivMP ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, Grids, DBGrids, Hcompte, Hctrls, ComCtrls,
  DBCtrls, Mask, DB, DBTables, Ent1, HEnt1, Hqry, hmsgbox, ParamDBG, PrintDBG,
  SaisUtil, SaisComm, Saisie, Guide, LettUtil, Menus, EncUtil, RIBModf, BanqueCP,FichComm,
  EcheMono, HStatus, EtapeReg, EcheUnit, Filtre, HDB, HSysMenu, ParamDat, EcheMPA,
{$IFDEF V530}
     EdtDoc,
{$ELSE}
     EdtRDoc,
{$ENDIF}
  Hspliter, EncTiers, Tiers, DecaBqe, Choix,
{$IFDEF V530}
      EdtEtat,
{$ELSE}
      EdtREtat,
{$ENDIF}
  HTB97, ed_tools, HPanel, UiUtil ;

procedure EncaisseDecaisse(Enc : boolean ; Circuit1,Circuit2 : String ; Generation,AffectBanque : boolean ; SorteLettre : TSorteLettre ) ;
procedure DecaisseCircuit ( n : Integer ) ;
procedure EncDecChequeTraiteBOR ( Chq : TSorteLettre ) ;

type
  TFSuiviMP = class(TForm)
    DockTop: TDock97;
    DockRight: TDock97;
    DockLeft: TDock97;
    DockBottom: TDock97;
    FindMvt: TFindDialog;
    HMEnc: THMsgBox;
    QEnc: THQuery;
    HDiv: THMsgBox;
    POPS: TPopupMenu;
    PFiltres: TToolWindow97;
    FFiltres: THValComboBox;
    BChercher: TToolbarButton97;
    Pages: TPageControl;
    PComptes: TTabSheet;
    HLabel4: THLabel;
    Bevel1: TBevel;
    E_GENERAL: THCpteEdit;
    PParam: TTabSheet;
    PEcritures: TTabSheet;
    TE_JOURNAL: THLabel;
    E_JOURNAL: THValComboBox;
    TE_NATUREPIECE: THLabel;
    E_NATUREPIECE: THValComboBox;
    TE_DATECOMPTABLE: THLabel;
    E_DATECOMPTABLE: THCritMaskEdit;
    TE_DATECOMPTABLE2: THLabel;
    E_DATECOMPTABLE_: THCritMaskEdit;
    E_NUMECHE: THCritMaskEdit;
    XX_WHERE: TEdit;
    MP_ENCAISSEMENT: THCritMaskEdit;
    E_QUALIFPIECE: THCritMaskEdit;
    E_ECHE: THCritMaskEdit;
    E_ETATLETTRAGE: THCritMaskEdit;
    HLabel7: THLabel;
    MP_CATEGORIE: THValComboBox;
    Label14: TLabel;
    E_MODEPAIE: THValComboBox;
    HLabel1: THLabel;
    GUIDE: THValComboBox;
    ETAPE: THValComboBox;
    HEtape: THLabel;
    HLabel17: THLabel;
    E_DEVISE: THValComboBox;
    HCpteGen: THLabel;
    GENEGENERATION: THCpteEdit;
    H_GENERAL: TLabel;
    H_GENEGENERATION: TLabel;
    Bevel2: TBevel;
    Bevel3: TBevel;
    HLabel9: THLabel;
    TYPEETAPE: THValComboBox;
    Label1: TLabel;
    Label2: TLabel;
    BZOOMETAPE: TToolbarButton97;
    E_TRESOLETTRE: THCritMaskEdit;
    HMTrad: THSystemMenu;
    PComplements: TTabSheet;
    Bevel4: TBevel;
    TE_NUMEROPIECE: THLabel;
    E_NUMEROPIECE: THCritMaskEdit;
    HLabel2: THLabel;
    E_NUMEROPIECE_: THCritMaskEdit;
    HLabel3: THLabel;
    E_AUXILIAIRE: THCpteEdit;
    TE_DEBIT: TLabel;
    E_DEBIT: THCritMaskEdit;
    TE_DEBIT_: TLabel;
    E_DEBIT_: THCritMaskEdit;
    TE_CREDIT: TLabel;
    E_CREDIT: THCritMaskEdit;
    TE_CREDIT_: TLabel;
    E_CREDIT_: THCritMaskEdit;
    CodeAFB: THMultiValComboBox;
    G: THGrid;
    Options97: TToolbar97;
    BAgrandir: TToolbarButton97;
    BRecherche: TToolbarButton97;
    BZoomPiece: TToolbarButton97;
    BModifEche: TToolbarButton97;
    Valide97: TToolbar97;
    BValide: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BAide: TToolbarButton97;
    XX_WHEREENC: TEdit;
    BImprimer: TToolbarButton97;
    FSelectAll: TCheckBox;
    iTiers: TToolbarButton97;
    iPiece: TToolbarButton97;
    PCircuit: TTabSheet;
    g1: TGroupBox;
    HLabel6: THLabel;
    E_SUIVDEC: THValComboBox;
    E_NOMLOT: THCritMaskEdit;
    HLabel12: THLabel;
    g2: TGroupBox;
    HLabel13: THLabel;
    HLotDest: THLabel;
    CircuitDest: THValComboBox;
    LotDest: TEdit;
    XX_WHERE1: TEdit;
    Panel1: TPanel;
    rDetail: TRadioButton;
    rMasse: TRadioButton;
    bAffectBqe: TToolbarButton97;
    bListeLots: TToolbarButton97;
    cVoirCircuit: TCheckBox;
    XX_WHERE3: TEdit;
    PGenere: TTabSheet;
    Bevel5: TBevel;
    H_MODEGENERE: TLabel;
    MODEGENERE: THValComboBox;
    HDateReg: THLabel;
    DATEGENERATION: THCritMaskEdit;
    CAutoEnreg: TCheckBox;
    GLOBAL: TCheckBox;
    DATEECHEANCE: THCritMaskEdit;
    cBAP: TCheckBox;
    cModeleBAP: THValComboBox;
    CChoixDate: TCheckBox;
    TE_EXERCICE: THLabel;
    E_EXERCICE: THValComboBox;
    H_OrdreTri: THLabel;
    OrdreTri: THValComboBox;
    TE_DATEECHEANCE: THLabel;
    E_DATEECHEANCE: THCritMaskEdit;
    TE_DATEECHEANCE2: THLabel;
    E_DATEECHEANCE_: THCritMaskEdit;
    TE_ETABLISSEMENT: THLabel;
    E_ETABLISSEMENT: THValComboBox;
    PLibres: TTabSheet;
    Bevel6: TBevel;
    TT_TABLE0: THLabel;
    TT_TABLE1: THLabel;
    TT_TABLE2: THLabel;
    TT_TABLE3: THLabel;
    TT_TABLE4: THLabel;
    TT_TABLE5: THLabel;
    TT_TABLE6: THLabel;
    TT_TABLE7: THLabel;
    TT_TABLE8: THLabel;
    TT_TABLE9: THLabel;
    T_TABLE4: THCpteEdit;
    T_TABLE3: THCpteEdit;
    T_TABLE2: THCpteEdit;
    T_TABLE1: THCpteEdit;
    T_TABLE0: THCpteEdit;
    T_TABLE5: THCpteEdit;
    T_TABLE6: THCpteEdit;
    T_TABLE7: THCpteEdit;
    T_TABLE8: THCpteEdit;
    T_TABLE9: THCpteEdit;
    Label3: TLabel;
    BSwapA: TBitBtn;
    BSwapR: TBitBtn;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    BFiltre: TToolbarButton97;
    H_ModeSaisie: TLabel;
    CModeSaisie: THValComboBox;
    Dock971: TDock97;
    H_TOTDEVISE: TLabel;
    ER_SOLDED: THNumEdit;
    BModifRIB: TToolbarButton97;
    GereAccept: TCheckBox;
    BZoomModele: TToolbarButton97;
    cFactCredit: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GDblClick(Sender: TObject);
    procedure GMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BSaveFiltreClick(Sender: TObject);
    procedure BZoomPieceClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure BAgrandirClick(Sender: TObject);
    procedure BRechercheClick(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure DATEGENERATIONExit(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
    procedure FindMvtFind(Sender: TObject);
    procedure BModePaieClick(Sender: TObject);
    procedure POPSPopup(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure E_DEVISEChange(Sender: TObject);
    procedure MP_CATEGORIEChange(Sender: TObject);
    procedure ETAPEChange(Sender: TObject);
    procedure E_GENERALExit(Sender: TObject);
    procedure GENEGENERATIONExit(Sender: TObject);
    procedure E_GENERALChange(Sender: TObject);
    procedure GENEGENERATIONChange(Sender: TObject);
    procedure BZOOMETAPEClick(Sender: TObject);
    procedure BModifEcheClick(Sender: TObject);
    procedure BDelFiltreClick(Sender: TObject);
    procedure BRenFiltreClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure BCreerFiltreClick(Sender: TObject);
    procedure E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
    procedure BAideClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure cTypeVisuChange(Sender: TObject);
    procedure GRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure bAffectBqeClick(Sender: TObject);
    procedure bListeLotsClick(Sender: TObject);
    procedure cBAPClick(Sender: TObject);
    procedure cVoirCircuitClick(Sender: TObject);
    procedure CChoixDateClick(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure POPFPopup(Sender: TObject);
    procedure BModifRIBClick(Sender: TObject);
    procedure BZoomModeleClick(Sender: TObject);

  private
    DateGene,NowFutur : TDateTime ;
    DEV        : RDEVISE ;
    CollecContreP,CollecTreso,NoBanque,BanqueVide,JalFerme : Boolean ;
    TGuide,LesM : TList ;
    LaSelection    : TStringList ;
    EntGuide       : RMVT ;
    TauxRef        : Double ;
    GX,GY          : integer ;
    PreVirMP,FormatCFONB,Document,EnvoiTrans : String3 ;
    PreVirEche     : TDateTime ;
    PreVirEgal,OkCFONB,OkBordereau : boolean ;
    XX  : Array[1..100] of String ;
    WMinX,WMinY        : Integer ;
    ParTiers,EnDev : Boolean ;
    TitreF         : String ;
    MinNum,MaxNum  : Longint ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
// Entete, Chargements
    procedure MajCaption ;
    procedure EtudierGuide ;
    procedure PositionneFerme( Jal : String ) ;
    procedure GUIDEChange(Sender: TObject);
    procedure SwapCat ( Vide : Boolean ) ;
    procedure InitEcran ;
    procedure InitCircuit ;
    procedure InitLettreCheque ;
    procedure InitLettreTraite ;
    procedure InitGeneration ;
    procedure InitLeReste ;
    procedure InitEncOuDec ;
    procedure PositionneBB ;
// Encaissement
    procedure CreerGuideTemp ;
    procedure CreerLignesTemp ( k : integer ) ;
    procedure LGContreCommun ( Q : TQuery ; OGuiContre : TOBM ; Auxi,GeneMP,OldRef,OldLib,OldPrevi,OldRib,OldReftire : String ; k : integer ) ;
    procedure LGContreLigne ( Q : TQuery ; DCur,CCur : double ; OldEche : TDateTime ; OldPaie : String3 ) ;
    procedure LGPremLigne ( Q : TQuery ; O : TOBM ) ;
    function  MontantGuideSt ( O : TOBM ; Debit : boolean ) : String ;
    function  MontantARegler ( O : TOBM ; Debit,AvecCouv : boolean ; Quoi : TSorteMontant ) : Double ;
    Function  JalBanque ( k : integer ) : String ;
    procedure FabricLesM ;
// Contrôles
    function  JeValidePas : boolean ;
    function  CoherGroupe : byte ;
    function  AcqModeDate ( Var MP : String3 ; Var DD : TDateTime ) : boolean ;
    function  CoherVirPre ( Var MP : String3 ; Var DD : TDateTime ) : boolean ;
    function  LesMeme : boolean ;
    function  VerifLesEche : boolean ;
    function  VerifLesDevise : boolean ;
    function  VerifLesRIB : boolean ;
    function  VerifLesPrevi : boolean ;
// Recherches, SQL
    procedure RempliGrid ;
    procedure RempliTiers(TEche : TList ; Lig : integer) ;
    procedure ChangeTypeGrid ;
// Sélections
    procedure CocheDecoche ( Lig : integer ; Next : boolean ) ;
    procedure CalculDebitCredit ;
    procedure AjouteOBM(O : TOBM ; Lig1,Lig2 : Integer ; Plus : Boolean) ;
    procedure ConstitueLaSelection ;
    Function  GetLeDebitCur ( O : TOBM ; AvecCouv : boolean ) : double ;
    Function  GetLeCreditCur ( O : TOBM ; AvecCouv : boolean ) : double ;
    Function  GetLaCouvCur ( O : TOBM ) : double ;
// Circuit
    procedure UpdateCircuit ;
    procedure UpdateOBMCircuit(O : TOBM) ;
    procedure GereOngletCircuit ( Montrer : boolean ) ;
// Affectation prévisionnelle des banques
    procedure AddGridAffBqe(G : THGrid ; O : TOBM) ;
    procedure DoAffectBqe ;
    Function  CountBqe : integer ;
  public
    Enc,FindFirst,Generation  : Boolean ;
    Circuit1,Circuit2         : String ;
    AffectBanque : Boolean ;
    SorteLettre  : TSorteLettre ;
  end;

implementation

{$R *.DFM}

uses SaisBor ;

procedure EncaisseDecaisse(Enc : boolean ; Circuit1,Circuit2 : String ; Generation,AffectBanque : boolean ; SorteLettre : TSorteLettre ) ;
Var X  : TFEncaDeca ;
    PP : THPanel ;
BEGIN
if PasCreerDate(V_PGI.DateEntree) then Exit ;
if Blocage(['nrBatch','nrCloture','nrLettrage'],True,'nrSaisieModif') then Exit ;
X:=TFEncaDeca.Create(Application) ;
X.ENC:=Enc ; X.Circuit1:=Circuit1 ; X.Circuit2:=Circuit2 ;
X.Generation:=Generation ; X.AffectBanque:=AffectBanque ;
X.SorteLettre:=SorteLettre ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     X.ShowModal ;
    Finally
     X.Free ;
     Bloqueur('nrSaisieModif',False) ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

procedure EncDecChequeTraiteBOR ( Chq : TSorteLettre ) ;
var Q  : TQuery ;
    s1,s2,stSQL : string ;
begin
s1:='' ; s2:='' ;
if Chq<>tslTraite then
   BEGIN
   stSQL:='SELECT * FROM CHOIXCOD WHERE CC_TYPE="CID" AND CC_LIBELLE<>"" ORDER BY CC_CODE DESC' ;
   Q:=OpenSQL(stSQL,True) ;
   if not Q.EOF then
      begin
      s2:=Q.FindField('CC_CODE').AsString ;
      Q.Next ;
      if not Q.EOF then s1:=Q.FindField('CC_CODE').AsString ;
      end ;
   Ferme(Q) ;
   EncaisseDecaisse(False,s1,s2,True,True,Chq) ;
   END else
   BEGIN
   EncaisseDecaisse(True,'','',True,True,Chq) ;
   END ;
end ;

procedure DecaisseCircuit(n : Integer);
var Q  : TQuery ;
    s1,s2,s3,stSQL : string ;
    Generation  : Boolean ;
    AffectBanque: Boolean ;
begin
if n=1 then s1:='' else s1:='CD'+IntToStr(n-1) ;
s2:='CD'+IntToStr(n) ;
//Charger l'étape du circuit
stSQL:='SELECT * FROM CHOIXCOD WHERE CC_TYPE="CID" AND CC_LIBELLE<>"" AND CC_CODE="'+s2+'"' ;
Q:=OpenSQL(stSQL,True) ;
AffectBanque:=(Q.FindField('CC_ABREGE').AsString='X') ;
Ferme(Q) ;
//Il faut vérifier si c'est est la dernière étape du circuit
s3:='CD'+Inttostr(n+1) ;
stSQL:='SELECT * FROM CHOIXCOD WHERE CC_TYPE="CID" AND CC_LIBELLE<>"" AND CC_CODE="'+s3+'"' ;
Q:=OpenSQL(stSQL,True) ;
Generation:=Q.EOF ;
Ferme(Q) ;
EncaisseDecaisse(False,s1,s2,Generation,AffectBanque,tslAucun) ;
end ;

{====================================== FORM, GRID ==========================================}
procedure TFSuiviMP.FormClose(Sender: TObject; var Action: TCloseAction);
begin
DetruitGuideEnc(True) ;
VideListe(TGuide) ; TGuide.Free ;
VideListe(LesM) ; LesM.Free ;
LaSelection.Clear ; LaSelection.Free ;
PurgePopup(POPS) ; G.VidePile(True) ;
RegSaveToolbarPos(Self,'EncaDeca') ;
if Parent is THPanel then
   BEGIN
   Bloqueur('nrSaisieModif',False) ;
   Action:=caFree ;
   END ;
end;

procedure TFSuiviMP.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
TGuide:=TList.Create ; LesM:=TList.Create ;
LaSelection:=TStringList.Create ; LaSelection.Sorted:=True ; LaSelection.Clear ;
OkCFONB:=False ; FormatCFONB:='' ; OkBordereau:=False ; Document:='' ; EnvoiTrans:='' ; 
WMinX:=Width ; WMinY:=258 ; MinNum:=999999999 ; MaxNum:=0 ;
RempliSelectEuro(CModeSaisie,HDiv.Mess[8]) ; JalFerme:=False ;
RegLoadToolbarPos(Self,'EncaDeca') ;
end;

procedure TFSuiviMP.InitEcran ;
BEGIN
{Généralités, Grid}
G.ListeParam:='ENCEURODET' ;
HMTrad.ResizeGridColumns(G) ;
if SorteLettre=tslCheque then TitreF:='CHEQUE' else
 if SorteLettre=tslTraite then TitreF:='ENCTRAITE' else
  if SorteLettre=tslBOR then TitreF:='DECBOR' else
     TitreF:='ENCDEC'+Chr(48+Ord(ENC)) ;
G.GetCellCanvas:=GetCellCanvas ;
G.RowCount:=2 ; OrdreTri.ItemIndex:=0 ; NoBanque:=True ; BanqueVide:=False ;
{Type de liste}
rDetail.Checked:=True ; cTypeVisuChange(nil) ;
{Initialisations zones}
DateGene:=V_PGI.DateEntree ;
E_EXERCICE.Value:=EXRF(VH^.Entree.Code) ;
DATEGENERATION.Text:=DateToStr(DateGene) ; DATEECHEANCE.Text:=StDate1900 ;
E_DATECOMPTABLE.Text:=StDate1900 ; E_DATECOMPTABLE_.Text:=StDate2099 ;
E_DATEECHEANCE.Text:=StDate1900  ; E_DATEECHEANCE_.Text:=StDate2099 ;
E_DEVISE.Value:=V_PGI.DevisePivot ;
ChangeMask(ER_SOLDED,V_PGI.OkDecV,V_PGI.SymbolePivot) ;
{Affectations diverses}
if AffectBanque then bAffectBqe.Hint:=HMEnc.Mess[39]
                else bAffectBqe.Hint:=HMEnc.Mess[40] ;
CircuitDest.Value:=Circuit2 ; CircuitDest.Enabled:=False ;
{Positionnement}
Pages.ActivePage:=PComptes ;
END ;

procedure TFSuiviMP.InitCircuit ;
Var Q : TQuery ;
BEGIN
PCircuit.TabVisible:=(Circuit2<>'') ;
if PCircuit.TabVisible then
   BEGIN
   if Circuit1<>'' then E_SUIVDEC.Value:=Circuit1 ;
   g1.Visible:=(Circuit1<>'') ;
   E_SUIVDEC.Value:=Circuit1 ; E_SUIVDEC.Enabled:=False ;
   cVoirCircuit.Visible:=SorteLettre in [tslCheque,tslBOR] ;
   cBAP.Enabled:=True ; cBAP.Checked:=True ; cModeleBAP.Enabled:=True ;
   BAffectBqe.Enabled:=AffectBanque ;
   END else
   BEGIN
   cBAP.Enabled:=False ; cModeleBAP.Enabled:=False ;
   if Not ENC then
      BEGIN
      Q:=OpenSQL('Select count(*) from CHOIXCOD Where CC_TYPE="CID" and CC_LIBELLE<>""',True) ;
      if Q.Fields[0].AsInteger<=0 then
         BEGIN
         cVoirCircuit.Visible:=False ;
         cVoirCircuit.Checked:=False ; GereOngletCircuit(False) ;
         END ;
      Ferme(Q) ;
      END ;
   END ;
END ;

procedure TFSuiviMP.InitLettreCheque ;
BEGIN
if SorteLettre<>tslCheque then Exit ;
CircuitDest.DataType:='' ; CircuitDest.Items.Add(HMEnc.Mess[45]) ; CircuitDest.ItemIndex:=0 ;
XX_WHERE3.Text:='MP_LETTRECHEQUE="X"' ;
GLOBAL.Checked:=False ; GLOBAL.Enabled:=False ;
Caption:=HMEnc.Mess[44] ;
E_MODEPAIE.Vide:=False ;
BModifEche.Enabled:=False ;
cVoirCircuit.Checked:=False ; GereOngletCircuit(False) ;
END ;

procedure TFSuiviMP.InitLettreTraite ;
BEGIN
if ((SorteLettre<>tslTraite) and (SorteLettre<>tslBOR)) then Exit ;
CircuitDest.DataType:='' ; CircuitDest.Items.Add(HMEnc.Mess[47]) ; CircuitDest.ItemIndex:=0 ;
XX_WHERE3.Text:='MP_LETTRETRAITE="X"' ;
if SorteLettre=tslBOR then Caption:=HMEnc.Mess[48] else Caption:=HMEnc.Mess[51] ;
E_MODEPAIE.Vide:=False ;
BModifEche.Enabled:=False ;
cVoirCircuit.Checked:=False ; GereOngletCircuit(False) ;
END ;

procedure TFSuiviMP.InitGeneration ;
BEGIN
if Generation then
   begin
   HEtape.Caption:=HMEnc.Mess[42] ;
   LotDest.Text:='' ; LotDest.Visible:=False ; HLotDest.Visible:=False ;
   cBAP.Visible:=False ; CBAP.Checked:=False ; cModeleBAP.Visible:=False ;
   end else
   begin
   HEtape.Caption:=HMEnc.Mess[43] ; H_GENEGENERATION.Visible:=False ;
   HCpteGen.Visible:=False ; GeneGeneration.Visible:=False ; Global.Visible:=False ;
   HDateReg.Visible:=False ; DateGeneration.Visible:=False ;
   end ;
PGenere.TabVisible:=Generation ;
BModifRIB.Enabled:=Generation ;
END ;

procedure TFSuiviMP.InitEncOuDec ;
BEGIN
if ENC then
   BEGIN
   HelpContext:=7496000 ;
   if SorteLettre=tslTraite then ETAPE.DataType:='ttEtapeEncTraite'
                            else ETAPE.DataType:='ttEtapeEncais' ;
   BAffectBqe.Enabled:=False ;
   E_GENERAL.ZoomTable:=tzGEncais ; Caption:=HMEnc.Mess[1] ;
   XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC"))' ;
   MP_ENCAISSEMENT.Text:='DEC' ;
   cVoirCircuit.Visible:=False ;
   END else
   BEGIN
   HelpContext:=7500000 ;
   if SorteLettre=tslCheque then ETAPE.DataType:='ttEtapeDecCheque' else
    if SorteLettre=tslBOR then ETAPE.DataType:='ttEtapeDecTraite' else
     if ((Generation) and (Circuit2<>'')) then ETAPE.DataType:='ttEtapeDecRegle' else
        ETAPE.DataType:='ttEtapeDecais' ;
   E_GENERAL.ZoomTable:=tzGDecais ; Caption:=HMEnc.Mess[2] ;
   XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF"))' ;
   MP_ENCAISSEMENT.Text:='ENC' ;
   PLibres.TabVisible:=False ;
   cFactCredit.Caption:=HDiv.Mess[9] ; 
   END ;
InitTablesLibresTiers(PLibres) ;
END ;

procedure TFSuiviMP.InitLeReste ;
BEGIN
if ETAPE.Values.Count=1 then ETAPE.Value:=ETAPE.Values[0] ;
if SorteLettre<>tslAucun then
   BEGIN
   if E_MODEPAIE.Values.Count>0 then E_MODEPAIE.Value:=E_MODEPAIE.Values[0] ;
   END ;
if ((SorteLettre=tslAucun) and (Not ENC) and (Generation)) then CAutoEnreg.Visible:=True ;
// Helpcontext
BZoomModele.Visible:=False ;
if SorteLettre=tslCheque then
   BEGIN
   HelpContext:=7505000 ;
   PComptes.HelpContext:=7505100 ; PParam.HelpContext:=7505200 ; PGenere.HelpContext:=7505300 ;
   PEcritures.HelpContext:=7505400 ; PComplements.HelpContext:=7505500 ;
   BZoomModele.Enabled:=True ; BZoomModele.Visible:=True ; 
   END else if ((Circuit1='') and (Circuit2<>'')) then
   BEGIN
   HelpContext:=7496200 ;
   PComptes.HelpContext:=7502100 ; PParam.HelpContext:=7502200 ; PGenere.HelpContext:=0 ;
   PCircuit.HelpContext:=7502300 ; PEcritures.HelpContext:=7502400 ; PComplements.HelpContext:=7502500 ;
   END else if ((Circuit1<>'') and (Circuit2<>'')) then
   BEGIN
   HelpContext:=7496200 ;
   PComptes.HelpContext:=7503100 ; PParam.HelpContext:=7503200 ; PGenere.HelpContext:=7503300 ;
   PCircuit.HelpContext:=7503400 ; PEcritures.HelpContext:=7503500 ; PComplements.HelpContext:=7503600 ;
   END else if Not ENC then
   BEGIN
   HelpContext:=7500000 ;
   PComptes.HelpContext:=7501100 ; PParam.HelpContext:=7501200 ; PGenere.HelpContext:=7501300 ;
   PCircuit.HelpContext:=0 ; PEcritures.HelpContext:=7501400 ; PComplements.HelpContext:=7501500 ;
   END ;
if SorteLettre=tslTraite then HelpContext:=7496100 ;
if SorteLettre=tslBOR then HelpContext:=7505760 ;
END ;

procedure TFSuiviMP.FormShow(Sender: TObject);
begin
InitEcran ;
InitCircuit ;
InitGeneration ;
InitEncOuDec ;
InitLettreCheque ;
InitLettreTraite ;
InitLeReste ;
ChargeFiltre(TitreF,FFiltres,Pages) ;
UpdateCaption(Self) ;
end;

procedure TFSuiviMP.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var OkG,Vide : boolean ;
begin
OkG:=G.Focused ; Vide:=(Shift=[]) ;
Case Key of
   VK_F9     : if Vide then BChercherClick(Nil) ;
   VK_F10    : if Vide then BValideClick(Nil) ;
   VK_SPACE  : if ((OkG) and (Vide)) then CocheDecoche(G.Row,False) ;
   VK_F5     : BEGIN Key:=0 ; BZoomPieceClick(Nil) ; END ;
          80 : if Shift=[ssCtrl] then BEGIN Key:=0 ; BImprimerClick(Nil) ; END ;
  END ;
end;

procedure TFSuiviMP.GMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
Var C,R : Longint ;
begin
GX:=X ; GY:=Y ;
if ((ssCtrl in Shift) and (Button=mbLeft)) then
   BEGIN
   G.MouseToCell(X,Y,C,R) ;
   if (R>0) then CocheDecoche(G.Row,False) ; //and R<G.RowCount-1
   END ;
end;

{======================================= SELEC =============================================}
Function TFSuiviMP.GetLaCouvCur ( O : TOBM ) : double ;
BEGIN
if E_DEVISE.Value<>V_PGI.DevisePivot then Result:=O.GetMvt('E_COUVERTUREDEV') else
   if CModeSaisie.Value='ECU' then Result:=O.GetMvt('E_COUVERTUREEURO') else
      Result:=O.GetMvt('E_COUVERTURE') ;
END ;

Function TFSuiviMP.GetLeDebitCur ( O : TOBM ; AvecCouv : boolean ) : double ;
BEGIN
if E_DEVISE.Value<>V_PGI.DevisePivot then Result:=MontantARegler(O,True,AvecCouv,tsmDevise) else
   if CModeSaisie.Value='ECU' then Result:=MontantARegler(O,True,AvecCouv,tsmEuro) else
      Result:=MontantARegler(O,True,AvecCouv,tsmPivot) ;
END ;

Function TFSuiviMP.GetLeCreditCur ( O : TOBM ; AvecCouv : boolean ) : double ;
BEGIN
if E_DEVISE.Value<>V_PGI.DevisePivot then Result:=MontantARegler(O,False,AvecCouv,tsmDevise) else
   if CModeSaisie.Value='ECU' then Result:=MontantARegler(O,False,AvecCouv,tsmEuro) else
      Result:=MontantARegler(O,False,AvecCouv,tsmPivot) ;
END ;

procedure TFSuiviMP.CalculDebitCredit ;
Var TotDeb,TotCred : double ;
    i,j   : integer;
    O     : TOBM ;
    TEche : TList ;
BEGIN
TotDeb:=0 ; TotCred:=0 ;
if ParTiers then
   begin
   {Calcul des totaux par tiers}
   for i:=1 to G.RowCount-1 do
      begin
      if EstSelect(G,i) then
         begin
         TEche:=TList(G.Objects[0,i]) ; if TEche=Nil then Continue ;
         for j:=0 to TEche.Count-1 do
             BEGIN
             O:=TOBM(TEche[j]) ; if O=Nil then Break ;
             TotDeb:=TotDeb+GetLeDebitCur(O,True) ;
             TotCred:=TotCred+GetLeCreditCur(O,True) ;
             END ;
         end ;
      end ;
   end else
   begin
   {Calcul des totaux en détail}
   for i:=1 to G.RowCount-1 do if EstSelect(G,i) then
       begin
       O:=TOBM(G.Objects[0,i]) ; if O=Nil then Break ;
       TotDeb:=TotDeb+GetLeDebitCur(O,True) ;
       TotCred:=TotCred+GetLeCreditCur(O,True) ;
       end ;
    end ;
AfficheLeSolde(ER_SOLDED,TotDeb,TotCred) ;
END ;

procedure TFSuiviMP.AjouteOBM(O : TOBM ; Lig1,Lig2 : Integer ; Plus : Boolean) ;
var St,MP : String ;
    ii    : integer ;
begin
MP:=O.GetMvt('E_MODEPAIE') ;
St:=Format_String(O.GetMvt('E_AUXILIAIRE'),17)+';'
   +Format_String(O.GetMvt('E_CONTREPARTIEGEN'),17)+';'
   +InttoStr(Lig1)+';'+Format_String(MP,3)+';'
   +DateToStr(O.GetMvt('E_DATEECHEANCE'))+';'
   +InttoStr(Lig2)+';' ;
if Plus then LaSelection.Add(St) else
   begin
   ii:=LaSelection.IndexOf(St) ;
   if ii>=0 then LaSelection.Delete(ii) ;
   end ;
end ;

procedure TFSuiviMP.CocheDecoche ( Lig : integer ; Next : boolean ) ;
Var Plus,Cancel : boolean ;
BEGIN
//if Lig=G.RowCount-1 then Exit ;
if EstSelect(G,Lig) then BEGIN G.Cells[G.ColCount-1,Lig]:=' ' ; Plus:=False ; END
                    else BEGIN G.Cells[G.ColCount-1,Lig]:='+' ; Plus:=True ; END ;
Cancel:=False ; GRowEnter(Nil,Lig,Cancel,False);
if ((Lig=G.RowCount-1) and (Next)) then Next:=False ;
if Next then G.Row:=Lig+1 ;
G.Invalidate ;
CalculDebitCredit ;
END ;

procedure TFSuiviMP.GDblClick(Sender: TObject);
Var C,R : longint ;
begin
G.MouseToCell(GX,GY,C,R) ;
if (R>0) then CocheDecoche(G.Row,True) ;
end;

{======================================= SQL ===============================================}
procedure TFSuiviMP.ChangeTypeGrid ;
begin
ParTiers:=(rMasse.Checked) ;
if (G.RowCount<=2) then //liste vide
   begin
   if (ParTiers) then G.ListeParam:='ENCEUROT' else G.ListeParam:='ENCEURODET' ;
   HMTrad.ResizeGridColumns(G) ;
   end ;
G.SortedCol:=-1 ;
if ParTiers then
   begin
   bModifEche.Hint:=iTiers.Hint ;
   bModifEche.Glyph:=iTiers.Glyph ;
   bZoomPiece.Hint:=HMEnc.Mess[33] ;
   BModifEche.Enabled:=True ;
   end else
   begin
   bModifEche.Hint:=iPiece.Hint ;
   bModifEche.Glyph:=iPiece.Glyph ;
   bZoomPiece.Hint:=HMEnc.Mess[34] ;
   BModifEche.Enabled:=(SorteLettre=tslAucun) ;
   end ;
end ;

procedure TFSuiviMP.RempliGrid ;
Var O,O1 : TOBM ;
    i,k : integer ;
    OldGeneAuxi,NewGeneAuxi,StS : String ;
    Premier,Okok,OkAccept : Boolean ;
    TEche,TPiece : TList ;
    XD,XC,XV,X : Double ;
begin
LaSelection.Clear ;
TEche:=TList.Create ; TPiece:=TList.Create ;
Premier:=True ; Okok:=False ;
while not QEnc.EOF do
   begin
   O:=TOBM.Create(EcrGen,'',True) ; FillChar(O.T,Sizeof(O.T),#0) ; k:=0 ; Okok:=True ;
   while Copy(VH^.DescriEcr[k],1,2)='E_' do begin O.T[k].V:=QEnc.Fields[k].AsVariant ; Inc(k) ; END ;
   if FSelectAll.Checked then O.PutMvt('E_ETAT','*') else O.PutMvt('E_ETAT',' ') ;
   {Sauvegarde mémoire des infos}
   O.PutMvt('E_MULTIPAIEMENT',O.GetMvt('E_LIBRETEXTE0')) ;
   StS:=O.GetMvt('E_CONTREPARTIEGEN')+';'+O.GetMvt('E_CONTREPARTIEAUX') ;
   O.PutMvt('E_TRACE',StS) ;
   {Maj mémoire des infos}
   O.PutMvt('E_LIBRETEXTE0',QEnc.FindField('T_LIBELLE').AsString) ; {ruse fichier TIERS}
   O.PutMvt('E_CONTREPARTIEGEN',QEnc.FindField('MP_GENERAL').AsString) ; {ruse fichier MP}
   O.PutMvt('E_CONTREPARTIEAUX',QEnc.FindField('MP_ENCAISSEMENT').AsString) ; {ruse fichier MP}
   NewGeneAuxi:=O.GetMvt('E_GENERAL')+O.GetMvt('E_AUXILIAIRE') ;
   if Premier then OldGeneAuxi:=NewGeneAuxi ; Premier:=False ;
   if (ParTiers) and (NewGeneAuxi<>OldGeneAuxi) then
      begin
      RempliTiers(TEche,G.RowCount-1) ; TEche:=TList.Create ; OldGeneAuxi:=NewGeneAuxi ;
      G.RowCount:=G.RowCount+1 ;
      end ;
   if ParTiers then
      begin
      O.PutMvt('E_ETAT','*') ;
      TEche.Add(O) ;
      end else
      begin
      TPiece.Add(O) ;
      end ;
   QEnc.Next ;
   end ;
{Cas du dernier tiers}
if ParTiers then
   begin
   if Okok then //Au moins 1 tiers
      begin
      RempliTiers(TEche,G.RowCount-1) ;
      G.RowCount:=G.RowCount+1 ;
      end ;
   end ;
{Cas du traitement en détail : on traite TPieces}
if not ParTiers then
   {Lister tous les mouvements}
   begin
   for i:=0 to TPiece.Count-1 do
       begin
       O1:=TOBM(TPiece[i]) ; X:=0 ;
       O:=Nil ; EgaliseOBM(O1,O) ;
       {Sauvegarde des zones OBM}
       XD:=O.GetMvt('E_DEBIT') ; XC:=O.GetMvt('E_CREDIT') ; XV:=O.GetMvt('E_COUVERTURE') ;
       {Manipulation des zones pour affichage contextuel}
       O.PutMvt('E_DEBIT',GetLeDebitCur(O,False)) ;
       O.PutMvt('E_CREDIT',GetLeCreditCur(O,False)) ;
       O.PutMvt('E_COUVERTURE',GetLaCouvCur(O)) ;
       {Affichage}
       ComCom1(G,O) ;
       G.Objects[0,G.RowCount-1]:=O ;
       {Restauration des zones OBM}
       O.PutMvt('E_DEBIT',XD) ; O.PutMvt('E_CREDIT',XC) ; O.PutMvt('E_COUVERTURE',XV) ;
       if FSelectAll.Checked then G.Cells[G.ColCount-1,G.RowCount-1]:='*' ;
       G.RowCount:=G.RowCount+1 ;
       end ;
   VideListe(TPiece) ;
   end ;
if G.RowCount>2 then G.RowCount:=G.RowCount-1 ;
TPiece.Free ;
end ;

procedure TFSuiviMP.RempliTiers(TEche : TList ; Lig : integer) ;
Var i : integer ;
    O : TOBM ;
    Gene,Aux,Libelle : String ;
    DebitCur,CreditCur,TotDebCur,TotCredCur : Double ;
    CouvCur,TotCouvCur : Double ;
BEGIN
DebitCur:=0 ; CreditCur:=0 ; TotDebCur:=0 ; TotCredCur:=0 ; CouvCur:=0 ; TotCouvCur:=0 ;
for i:=0 to TEche.Count-1 do
    BEGIN
    O:=TOBM(TEche[i]) ;
    Gene:=O.GetMvt('E_GENERAL') ; Aux:=O.GetMvt('E_AUXILIAIRE') ;
    Libelle:=O.GetMvt('E_LIBRETEXTE0') ;
    if O.GetMvt('E_ETAT')='*' then
       BEGIN
       DebitCur:=GetLeDebitCur(O,False) ;
       CreditCur:=GetLeCreditCur(O,False) ;
       CouvCur:=GetLaCouvCur(O) ;
       TotDebCur:=TotDebCur+DebitCur ; TotCredCur:=TotCredCur+CreditCur ;
       if DebitCur<>0 then TotCouvCur:=TotCouvCur+CouvCur else TotCouvCur:=TotCouvCur-CouvCur ;
       END ;
    END ;
if Not Enc then TotCouvCur:=-TotCouvCur ;
G.Cells[0,Lig]:=Gene ; G.Cells[1,Lig]:=Aux ; G.Cells[2,Lig]:=Libelle ;
G.Cells[3,Lig]:=FormatFloat(G.ColFormats[3],TotDebCur) ;
G.Cells[4,Lig]:=FormatFloat(G.ColFormats[4],TotCredCur) ;
G.Cells[5,Lig]:=FormatFloat(G.ColFormats[5],TotCouvCur) ;
G.Objects[0,Lig]:=TEche ;
if FSelectAll.Checked then G.Cells[G.ColCount-1,Lig]:='*' else G.Cells[G.ColCount-1,Lig]:='' ;
END ;

{--------------------------------------------------------------------------------------------}
Procedure TFSuiviMP.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
if (ARow<>0) And (EstSelect(G,ARow)) then
   G.Canvas.Font.Style:=G.Canvas.Font.Style+[fsItalic] else
   G.Canvas.Font.Style:=G.Canvas.Font.Style-[fsItalic] ;
END ;

procedure TFSuiviMP.BChercherClick(Sender: TObject);
Var StV8,StXP,StXN,StXP2,StXN2 : String ;
begin
if ENC then
   BEGIN
   if cFactCredit.Checked
      then XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC" OR E_NATUREPIECE="FC"))'
      else XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC"))' ;
   END else
   BEGIN
   if cFactCredit.Checked
      then XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF" OR E_NATUREPIECE="FF"))'
      else XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF"))' ;
   END ;
if PCircuit.TabVisible then
   BEGIN
   if ((G1.Visible) and (E_SUIVDEC.Value='')) then begin HMEnc.Execute(26,caption,'') ; Exit ; end ;
   if ((E_SUIVDEC.Value<>'') and (Trim(E_NOMLOT.Text)='')) then begin HMEnc.Execute(37,caption,'') ; Exit ; end ;
   END ;
if SorteLettre in [tslCheque,tslBOR] then
   BEGIN
   if cVoirCircuit.Checked then XX_WHERE1.Text:='E_SUIVDEC<>""' else
      BEGIN
      XX_WHERE1.Text:='E_SUIVDEC=""' ; E_SUIVDEC.Value:='' ; E_NOMLOT.Text:='' ;
      END ;
   END else
   BEGIN
   if cVoirCircuit.Visible then
      BEGIN
      if cVoirCircuit.Checked then XX_WHERE1.Text:='' else XX_WHERE1.Text:='E_SUIVDEC=""' ;
      END else
      BEGIN
      if ((E_SUIVDEC.Value='') and (Not Generation) and (Not ENC)) then XX_WHERE1.Text:='E_SUIVDEC=""' ;
      END ;
   END ;
if ETAPE.Value='' then BEGIN HMEnc.Execute(27,caption,'') ; if ETAPE.CanFocus then ETAPE.SetFocus ; Exit ; END ;
if E_GENERAL.Text='' then BEGIN HMEnc.Execute(0,caption,'') ; if E_GENERAL.CanFocus then E_GENERAL.SetFocus ; Exit ; END ;
if Not BonCompteEnc(ENC,E_GENERAL.Text) then BEGIN HMEnc.Execute(8,caption,'') ; if E_GENERAL.CanFocus then E_GENERAL.SetFocus ; Exit ; END ;
if GUIDE.Value='' then BEGIN HMEnc.Execute(7,caption,'') ; if GUIDE.CanFocus then GUIDE.SetFocus ; Exit ; END ;
if ((MP_CATEGORIE.Value='') and (Not MP_CATEGORIE.Vide)) then BEGIN HMEnc.Execute(16,caption,'') ; if MP_CATEGORIE.CanFocus then MP_CATEGORIE.SetFocus ; Exit ; END ;
EnDev:=(E_DEVISE.Value<>V_PGI.DevisePivot) ;
G.VidePile(True) ; ChangeTypeGrid ;
QEnc.Close ; QEnc.SQL.Clear ;
QEnc.SQL.Add('Select Ecriture.*,Tiers.T_LIBELLE,ModePaie.* from Ecriture ') ;
QEnc.SQL.Add(' Left outer join Tiers on E_AUXILIAIRE=T_AUXILIAIRE') ;
QEnc.SQL.Add(' Left outer join Modepaie on E_MODEPAIE=MP_MODEPAIE') ;
QEnc.SQL.Add(' Where '+WhereGeneCritEcr(Pages,QEnc,False)) ;
StV8:=LWhereV8 ; if StV8<>'' then QEnc.SQL.Add(' AND '+StV8) ;
if E_DEVISE.Value<>V_PGI.DevisePivot then
   BEGIN
   QEnc.SQL.Add(' AND (E_ETATLETTRAGE<>"PL" OR E_LETTRAGEDEV="X")') ;
   StXP:=StrFPoint(9*Resolution(DEV.Decimale+1)) ; StXN:=StrFPoint(-9*Resolution(DEV.Decimale+1)) ;
   QEnc.SQL.Add(' and (E_DEBITDEV+E_CREDITDEV-E_COUVERTUREDEV not between '+StXN+' AND '+StXP+')') ;
   END else
   begin
   if CModeSaisie.Value='ECU' then
      BEGIN
      QEnc.SQL.Add(' AND ((E_ETATLETTRAGE="AL" AND E_SAISIEEURO="X") OR (E_ETATLETTRAGE="PL" AND E_LETTRAGEEURO="X"))') ;
      END else if CModeSaisie.Value=V_PGI.DevisePivot then
      BEGIN
      QEnc.SQL.Add(' AND ((E_ETATLETTRAGE="AL" AND E_SAISIEEURO="-") OR (E_ETATLETTRAGE="PL" AND E_LETTRAGEEURO="-"))') ;
      END else
      BEGIN
      QEnc.SQL.Add(' AND ((E_ETATLETTRAGE="AL") OR (E_ETATLETTRAGE="PL" AND E_LETTRAGEEURO="-"))') ;
      END ;
   StXP:=StrFPoint(9*Resolution(V_PGI.OkDecV+1))  ; StXN:=StrFPoint(-9*Resolution(V_PGI.OkDecV+1)) ;
   StXP2:=StrFPoint(9*Resolution(V_PGI.OkDecE+1)) ; StXN2:=StrFPoint(-9*Resolution(V_PGI.OkDecE+1)) ;
   QEnc.SQL.Add(' And (E_DEBIT+E_CREDIT-E_COUVERTURE not between '+StXN+' AND '+StXP+')') ;
   QEnc.SQL.Add(' And (E_DEBITEURO+E_CREDITEURO-E_COUVERTUREEURO not between '+StXN2+' AND '+StXP2+')') ;
   end ;
if TypeEtape.Value='POR' then
   begin
   QEnc.SQL.Add(' AND ((E_NATUREPIECE<>"FC" AND E_NATUREPIECE<>"AC" AND E_NATUREPIECE<>"FF" AND E_NATUREPIECE<>"AF") OR (MP_GENERAL<>""))') ;
   end ;
if ((GereAccept.Visible) and (TypeEtape.Value='BQE') and (ENC) and (GereAccept.Checked)) then
   BEGIN
   QEnc.SQL.Add(' AND (MP_CATEGORIE<>"LCR" OR MP_CODEACCEPT="ACC" OR MP_CODEACCEPT="NON" OR MP_CODEACCEPT="BOR")') ;
   END ;
if ((Not ENC) and (TypeEtape.Value='BQE') and (SorteLettre=tslAucun) and (Circuit1='') and (Circuit2='')) then
   BEGIN
   QEnc.SQL.Add(' AND (E_BANQUEPREVI="" OR E_BANQUEPREVI="'+GeneGeneration.Text+'")') ;
   END ;
if ParTiers then
   begin
   QEnc.SQL.Add('order by E_GENERAL, E_AUXILIAIRE, E_EXERCICE, E_DATECOMPTABLE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ') ;
   end else
   begin
   case OrdreTri.ItemIndex of
      0 : QEnc.SQL.Add('order by E_GENERAL, E_EXERCICE, E_DATECOMPTABLE, E_AUXILIAIRE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ') ;
      1 : QEnc.SQL.Add('order by E_AUXILIAIRE, E_GENERAL, E_EXERCICE, E_DATECOMPTABLE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ') ;
      2 : QEnc.SQL.Add('order by E_EXERCICE, E_DATECOMPTABLE, E_GENERAL, E_AUXILIAIRE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ') ;
      end ;
   end ;
ChangeSQL(QEnc) ; EnableControls(Self,False) ; BZOOMETAPE.Enabled:=False ;
QEnc.Open ; EnableControls(Self,True) ; BZOOMETAPE.Enabled:=True ;
RempliGrid ; QEnc.Close ; G.Enabled:=True ;
CalculDebitCredit ; PreVirEgal:=False ; PreVirMP:='' ; PreVirEche:=0 ;
end;

{======================================== ENTETE ===========================================}
procedure TFSuiviMP.E_DEVISEChange(Sender: TObject);
begin
if E_DEVISE.Value='' then BEGIN FillChar(DEV,Sizeof(DEV),#0) ; Exit ; END ;
DEV.Code:=E_DEVISE.Value ; GetInfosDevise(DEV) ;
ChangeMask(ER_SOLDED,DEV.Decimale,DEV.Symbole) ;
if E_DEVISE.Value=V_PGI.DevisePivot then CModeSaisie.Visible:=True else
   BEGIN
   CModeSaisie.Visible:=False ; CModeSaisie.Value:='MIX' ;
   END ;
end;

procedure TFSuiviMP.MP_CATEGORIEChange(Sender: TObject);
begin
CatToMP(MP_CATEGORIE.Value,E_MODEPAIE.Items,E_MODEPAIE.Values,SorteLettre) ;
if ((E_MODEPAIE.Items.Count>0) and (E_MODEPAIE.Vide) and (SorteLettre<>tslAucun)) then E_MODEPAIE.ItemIndex:=0 ;
G.Enabled:=False ;
end;

procedure TFSuiviMP.PositionneFerme( Jal : String ) ;
Var QQ : TQuery ;
BEGIN
if Jal='' then Exit ;
QQ:=OpenSQL('Select J_FERME From JOURNAL Where J_JOURNAL="'+Jal+'"',True) ;
if Not QQ.EOF then
   BEGIN
   if QQ.Fields[0].AsString='X' then JalFerme:=True ;
   END ;
Ferme(QQ) ;
END ;

procedure TFSuiviMP.EtudierGuide ;
Var Code : String3 ;
    Q     : TQuery ;
    O     : TOBM ;
    Trouv : boolean ;
    Auxi1,Auxi2,Coll1,Coll2 : String17 ;
BEGIN
VideListe(TGuide) ; FillChar(EntGuide,Sizeof(EntGuide),#0) ;
Code:=GUIDE.Value ; JalFerme:=False ; if Code='' then Exit ;
{Entête du guide}
if ENC then Q:=OpenSQL('Select * from GUIDE where GU_TYPE="ENC" AND GU_GUIDE="'+Code+'"',True)
       else Q:=OpenSQL('Select * from GUIDE where GU_TYPE="DEC" AND GU_GUIDE="'+Code+'"',True) ;
Trouv:=Not Q.EOF ;
if Trouv then
   BEGIN
   EntGuide.Etabl:=Q.FindField('GU_ETABLISSEMENT').AsString ;
   EntGuide.Jal:=Q.FindField('GU_JOURNAL').AsString ;
   EntGuide.Nature:=Q.FindField('GU_NATUREPIECE').AsString ;
   EntGuide.CodeD:=Q.FindField('GU_DEVISE').AsString ;
   EntGuide.ModeOppose:=(Q.FindField('GU_SAISIEEURO').AsString='X') ;
   END ;
Ferme(Q) ; if Not Trouv then Exit ;
PositionneFerme(EntGuide.Jal) ;
EntGuide.DateC:=DateGene  ; EntGuide.Exo:=QuelExo(DATEGENERATION.Text) ;
EntGuide.Simul:='N'       ; EntGuide.General:=E_GENERAL.Text ;
EntGuide.LeGuide:=Code    ; EntGuide.Valide:=False ;
if ENC then EntGuide.TypeGuide:='ENC' else EntGuide.TypeGuide:='DEC' ;
if ((EntGuide.CodeD<>'') and (EntGuide.CodeD<>DEV.Code)) then BEGIN HMEnc.Execute(20,caption,'') ; Exit ; END ;
{Lignes du guide}
if ENC then Q:=OpenSQL('Select * from ECRGUI where EG_TYPE="ENC" AND EG_GUIDE="'+Code+'" Order By EG_TYPE, EG_GUIDE, EG_NUMLIGNE',True)
       else Q:=OpenSQL('Select * from ECRGUI where EG_TYPE="DEC" AND EG_GUIDE="'+Code+'" Order By EG_TYPE, EG_GUIDE, EG_NUMLIGNE',True) ;
Trouv:=Not Q.EOF ;
if Trouv then
   BEGIN
   While Not Q.EOF do
      BEGIN
      O:=TOBM.Create(EcrGui,'',True) ; O.ChargeMvt(Q) ; TGuide.Add(O) ;
      Q.Next ;
      END ;
   END ;
Ferme(Q) ; if Not Trouv then Exit ;
{cohérence du guide}
if TGuide.Count<>2 then BEGIN HMEnc.Execute(17,caption,'') ; Exit ; END ;
Auxi1:=TOBM(TGuide[0]).GetMvt('EG_AUXILIAIRE') ; Coll1:=TOBM(TGuide[0]).GetMvt('EG_GENERAL') ;
Auxi2:=TOBM(TGuide[1]).GetMvt('EG_AUXILIAIRE') ; Coll2:=TOBM(TGuide[1]).GetMvt('EG_GENERAL') ;
CollecTreso:=EstCollectif(Coll1) ; CollecContreP:=EstCollectif(Coll2) ;
if ((TypeEtape.Value='BQE') and (CollecContreP)) then BEGIN HMEnc.Execute(21,caption,'') ; Exit ; END ;
if EntGuide.ModeOppose then
   BEGIN
   CModeSaisie.Value:='ECU' ; CModeSaisie.Enabled:=False ;
   END else
   BEGIN
   CModeSaisie.Value:=V_PGI.DevisePivot ; CModeSaisie.Enabled:=True ;
   END ;
END ;

procedure TFSuiviMP.MajCaption ;
Var St : String ;
BEGIN
if Generation then
   begin
   if ENC then St:=HMEnc.Mess[1] else St:=HMEnc.Mess[2] ;
   St:=St+' - '+Lowercase(TypeEtape.Text) ;
   Caption:=St ;
   end else
   begin
   Caption:=CircuitDest.Items[CircuitDest.ItemIndex] ;
   end ;
END ;

procedure TFSuiviMP.GUIDEChange(Sender: TObject);
begin
EtudierGuide ;
end;

procedure TFSuiviMP.DATEGENERATIONExit(Sender: TObject);
Var DD  : TDateTime ;
    Err : integer ;
begin
if csDestroying in ComponentState then Exit ;
if Not IsValidDate(DATEGENERATION.Text) then
   BEGIN
   Pages.ActivePage:=PComptes ;
   HMEnc.Execute(9,caption,'') ;
   DATEGENERATION.Text:=DateToStr(V_PGI.DateEntree) ; DateGene:=V_PGI.DateEntree ;
   END else
   BEGIN
   DD:=StrToDate(DATEGENERATION.Text) ; Err:=DateCorrecte(DD) ;
   if Err>0 then
      BEGIN
      Pages.ActivePage:=PComptes ;
      HMEnc.Execute(9+Err,caption,'') ;
      DATEGENERATION.Text:=DateToStr(V_PGI.DateEntree) ;
      DateGene:=V_PGI.DateEntree ;
      END else
      BEGIN
      if RevisionActive(DD) then
         BEGIN
         DATEGENERATION.Text:=DateToStr(V_PGI.DateEntree) ;
         DateGene:=V_PGI.DateEntree ;
         END else
         BEGIN
         DateGene:=DD ;
         END ;
      END ;
   END ;
end;

{======================================= OUTILS ============================================}
procedure TFSuiviMP.BCreerFiltreClick(Sender: TObject);
begin
NewFiltre(TitreF,FFiltres,Pages) ;
end;

procedure TFSuiviMP.BSaveFiltreClick(Sender: TObject);
begin
SaveFiltre(TitreF,FFiltres,Pages) ;
end;

procedure TFSuiviMP.BDelFiltreClick(Sender: TObject);
begin
DeleteFiltre(TitreF,FFiltres) ;
end;

procedure TFSuiviMP.BRenFiltreClick(Sender: TObject);
begin
RenameFiltre(TitreF,FFiltres) ;
end;

procedure TFSuiviMP.BNouvRechClick(Sender: TObject);
begin VideFiltre(FFiltres,Pages) ; end;

procedure TFSuiviMP.FFiltresChange(Sender: TObject);
begin
LoadFiltre(TitreF,FFiltres,Pages) ;
end;

procedure TFSuiviMP.BAgrandirClick(Sender: TObject);
begin
Pages.Visible:=not Pages.Visible ;
if Pages.Visible then
   BEGIN
   BAgrandir.Glyph:=BSwapA.Glyph ;
   BAgrandir.Hint:=BSwapA.Hint ;                       
   BAgrandir.Caption:=BSwapA.Caption ;
   END else
   BEGIN
   BAgrandir.Glyph:=BSwapR.Glyph ;
   BAgrandir.Hint:=BSwapR.Hint ;
   BAgrandir.Caption:=BSwapR.Caption ;
   END ;
end ;

procedure TFSuiviMP.FindMvtFind(Sender: TObject);
begin Rechercher(G,FindMvt,FindFirst) ; end;

procedure TFSuiviMP.BRechercheClick(Sender: TObject);
begin FindFirst:=True ; FindMvt.Execute ; end;

procedure TFSuiviMP.BModePaieClick(Sender: TObject);
begin FicheModePaie_AGL('') ; end;

procedure TFSuiviMP.POPSPopup(Sender: TObject);
begin InitPopUp(Self) ; end;

procedure TFSuiviMP.BImprimerClick(Sender: TObject);
begin PrintDBGrid(G,Pages,Caption,'') ; End;

procedure TFSuiviMP.E_GENERALChange(Sender: TObject);
begin G.Enabled:=False ; end;

procedure TFSuiviMP.GENEGENERATIONChange(Sender: TObject);
begin G.Enabled:=False ; end;

procedure TFSuiviMP.BZoomPieceClick(Sender: TObject);
Var Lig : integer ;
    O   : TOBM ;
    M   : RMVT ;
begin
Lig:=G.Row ; if ((Lig<=0) or (Lig>G.RowCount-1)) then Exit ;
if ParTiers then
   begin
   {Zoom tiers}
   if G.Cells[1,Lig]<>'' then FicheTiers(Nil,'',G.Cells[1,Lig],taConsult,1) ;
   end else
   begin
   {Zoom pièce}
   if G.Cells[1,Lig]<>'' then
      begin
      O:=TOBM(G.Objects[0,Lig]) ; if O=Nil then Exit ;
      M:=OBMToIdent(O,False) ;
      if ((M.ModeSaisieJal<>'') And (M.ModeSaisieJal<>'-')) Then LanceSaisieFolioOBM(O,taConsult)
                                                            Else LanceSaisie(Nil,taConsult,M)
      end ;
   end ;
end;

{================================= CONTROLES ==========================================}
function TFSuiviMP.CoherGroupe : byte ;
Var k : integer ;
    OldMP,NewMP  : String3 ;
    St,Junk,Auxi,OldAux : String ;
BEGIN
Result:=0 ;
if LaSelection.Count<=0 then Exit ;
{Cohérence des modes de paiement}
if ((TypeEtape.Value='BQE') and (GLOBAL.Checked)) then
   BEGIN
   for k:=0 to LaSelection.Count-1 do
       BEGIN
       St:=LaSelection[k] ; Auxi:=Trim(ReadtokenSt(St)) ; Junk:=Trim(ReadTokenSt(St)) ;
       Junk:=ReadTokenSt(St) ; NewMP:=ReadTokenSt(St) ;
       if MODEGENERE.Value<>'' then NewMP:=MODEGENERE.Value ;
       if k=0 then BEGIN OldMP:=NewMP ; OldAux:=Auxi ; END ;
       if ((OldAux=Auxi) and (OldMP<>NewMP)) then BEGIN Result:=2 ; Exit ; END ;
       OldMP:=NewMP ; OldAux:=Auxi ;
       END ;
   END ;
END ;

function TFSuiviMP.AcqModeDate ( Var MP : String3 ; Var DD : TDateTime ) : boolean ;
Var X : T_MONOECH ;
BEGIN
Result:=False ;
if HMEnc.Execute(24,caption,'')<>mrYes then Exit ;
X.DateEche:=DD ; X.ModePaie:=MP ; X.Cat:=MP_CATEGORIE.Value ; X.Treso:=False ;
X.OkInit:=False ; X.OkVal:=False ; X.Action:=taModif ; X.DateValeur:=DD ;
if SaisirMonoEcheance(X) then
   BEGIN
   DD:=X.DateEche ; MP:=X.ModePaie ;
   PreVirEche:=DD ; PreVirMP:=MP ;
   Result:=True ;
   END ;
END ;

function TFSuiviMP.CoherVirPre ( Var MP : String3 ; Var DD : TDateTime ) : boolean ;
Var St,NewMP,OldMP,Junk : String ;
    NewDate,OldDate     : TDateTime ;
    k,i   : integer ;
BEGIN
Result:=False ; OldDate:=0 ;
for k:=0 to LaSelection.Count-1 do
    BEGIN
    St:=LaSelection[k] ; for i:=1 to 3 do Junk:=Trim(ReadtokenSt(St)) ;
    NewMP:=ReadTokenSt(St) ; if MODEGENERE.Value<>'' then NewMP:=MODEGENERE.Value ;
    NewDate:=StrToDate(ReadTokenSt(St)) ; if DateEcheance.Enabled then NewDate:=StrToDate(DateEcheance.Text) ;
    if k=0 then BEGIN OldMP:=NewMP ; OldDate:=NewDate ; END ;
    if ((OldMP<>NewMP) or (OldDate<>NewDate)) then BEGIN MP:=NewMP ; DD:=NewDate ; Exit ; END ;
    OldMP:=NewMP ; OldDate:=NewDate ;
    END ;
Result:=True ;
END ;

function TFSuiviMP.LesMeme : boolean ;
Var OldDate,NewDate : TDateTime ;
    St,Junk         : String ;
    OldMP,NewMP     : String3 ;
    Idem            : boolean ;
    k,i             : integer ;
    Auxi,OldAux     : String ;
BEGIN
Idem:=True ; Result:=True ; OldDate:=0 ;
if Not GLOBAL.Checked then Exit ;
for k:=0 to LaSelection.Count-1 do
    BEGIN
    St:=LaSelection[k] ; Auxi:=Trim(ReadtokenSt(St)) ; for i:=1 to 2 do Junk:=Trim(ReadtokenSt(St)) ;
    NewMP:=ReadTokenSt(St) ; if MODEGENERE.Value<>'' then NewMP:=MODEGENERE.Value ;
    NewDate:=StrToDate(ReadTokenSt(St)) ; if DateEcheance.Enabled then NewDate:=StrToDate(DateEcheance.Text) ;
    if k=0 then BEGIN OldMP:=NewMP ; OldDate:=NewDate ; END ;
    if (OldAux=Auxi) and ((OldMP<>NewMP) or (OldDate<>NewDate)) then BEGIN Idem:=False ; Break ; END ;
    OldMP:=NewMP ; OldDate:=NewDate ; OldAux:=Auxi ;
    END ;
if Not Idem then BEGIN if HMEnc.Execute(30,caption,'')=mrYes then Idem:=True ; END ;
Result:=Idem ;
END ;

function TFSuiviMP.VerifLesPrevi : boolean ;
Var Okok  : boolean ;
    Lig,j : integer ;
    O     : TOBM ;
    TEche : TList ;
    Previ : String ;
BEGIN
Result:=True ; Okok:=True ;
if Not ENC then Exit ;
if SorteLettre<>tslAucun then Exit ;
for Lig:=1 to G.RowCount-1 do if EstSelect(G,Lig) then
    BEGIN
    if ParTiers then
       BEGIN
       TEche:=TList(G.Objects[0,Lig]) ;
       for j:=0 to TEche.Count-1 do
           BEGIN
           O:=TOBM(TEche[j]) ; if O.GetMvt('E_ETAT')<>'*' then Continue ;
           Previ:=O.GetMvt('E_BANQUEPREVI') ;
           if ((Previ<>'') and (Previ<>GeneGeneration.Text)) then BEGIN Okok:=False ; Break ; END ;
           END ;
       if Not Okok then Break ;
       END else
       BEGIN
       O:=TOBM(G.Objects[0,Lig]) ; if O=Nil then Continue ;
       Previ:=O.GetMvt('E_BANQUEPREVI') ;
       if ((Previ<>'') and (Previ<>GeneGeneration.Text)) then BEGIN Okok:=False ; Break ; END ;
       END ;
    END ;
if Not Okok then
   BEGIN
   if HMEnc.Execute(53,caption,'')=mrYes then Okok:=True ;
   END ;
Result:=Okok ;
END ;

function TFSuiviMP.VerifLesEche : boolean ;
Var NewEche : TDateTime ;
    Junk,St : String ;
    k,i     : integer ;
    okok    : boolean ;
BEGIN
Okok:=True ;
for k:=0 to LaSelection.Count-1 do
    BEGIN
    St:=LaSelection[k] ; for i:=1 to 4 do Junk:=Trim(ReadtokenSt(St)) ;
    NewEche:=StrToDate(ReadTokenSt(St)) ; if DateEcheance.Enabled then NewEche:=StrToDate(DateEcheance.Text) ;
    if Not NbJoursOk(DATEGENE,NewEche) then BEGIN Okok:=False ; Break ; END ;
    END ;
Result:=Okok ;
if Not Okok then HMEnc.Execute(31,caption,'') ;
END ;

function TFSuiviMP.VerifLesRIB : boolean ;
Var Okok  : boolean ;
    Lig,j : integer ;
    O     : TOBM ;
    TEche : TList ;
BEGIN
Okok:=True ; Result:=True ;
if SorteLettre<>tslTraite then Exit ;
for Lig:=1 to G.RowCount-1 do if EstSelect(G,Lig) then
    BEGIN
    if ParTiers then
       BEGIN
       TEche:=TList(G.Objects[0,Lig]) ;
       for j:=0 to TEche.Count-1 do
           BEGIN
           O:=TOBM(TEche[j]) ; if O.GetMvt('E_ETAT')<>'*' then Continue ;
           if Trim(O.GetMvt('E_RIB'))='' then BEGIN Okok:=False ; Break ; END ;
           END ;
       if Not Okok then Break ;
       END else
       BEGIN
       O:=TOBM(G.Objects[0,Lig]) ; if O=Nil then Continue ;
       if Trim(O.GetMvt('E_RIB'))='' then BEGIN Okok:=False ; Break ; END ;
       END ;
    END ;
Result:=Okok ;
if Not Okok then HMEnc.Execute(52,caption,'') ;
END ;

function TFSuiviMP.VerifLesDevise : boolean ;
Var Okok,Premier : boolean ;
    Taux  : Double ;
    Lig,j : integer ;
    O     : TOBM ;
    TEche : TList ;
BEGIN
Premier:=True ; Okok:=True ; Result:=True ;
TauxRef:=1 ; Taux:=1 ;
if DEV.Code=V_PGI.DevisePivot then Exit ;
for Lig:=1 to G.RowCount-1 do if EstSelect(G,Lig) then
    BEGIN
    if ParTiers then
       BEGIN
       TEche:=TList(G.Objects[0,Lig]) ;
       for j:=0 to TEche.Count-1 do
           BEGIN
           O:=TOBM(TEche[j]) ; if O.GetMvt('E_ETAT')<>'*' then Continue ;
           if Premier then
              BEGIN
              Taux:=O.GetMvt('E_TAUXDEV') ; Premier:=False ;
              END else
              BEGIN
              if Arrondi(O.GetMvt('E_TAUXDEV')-Taux,8)<>0 then BEGIN Okok:=False ; Break ; END ;
              END ;
           END ;
       if Not Okok then Break ;
       END else
       BEGIN
       O:=TOBM(G.Objects[0,Lig]) ; if O=Nil then Continue ;
       if Premier then
          BEGIN
          Taux:=O.GetMvt('E_TAUXDEV') ; Premier:=False ;
          END else
          BEGIN
          if Arrondi(O.GetMvt('E_TAUXDEV')-Taux,8)<>0 then BEGIN Okok:=False ; Break ; END ;
          END ;
       END ;
    END ;
Result:=Okok ; TauxRef:=Taux ;
if Not Okok then HMEnc.Execute(32,caption,'') ;
END ;

function TFSuiviMP.JeValidePas : boolean ;
Var ii : integer ;
    MP   : String3 ;
    DD   : TDateTime ;
BEGIN
Result:=True ; PreVirEgal:=False ;
if ETAPE.Value='' then BEGIN HMEnc.Execute(27,caption,'') ; Exit ; END ;
if Not G.Enabled then BEGIN HMEnc.Execute(14,caption,'') ; Exit ; END ;
if E_General.Text='' then BEGIN HMEnc.Execute(8,caption,'') ; Exit ; END ;
if ((TypeEtape.Value<>'POR') and (GENEGENERATION.Text='')) then BEGIN HMEnc.Execute(25,caption,'') ; Exit ; END ;
if GUIDE.Value='' then BEGIN HMEnc.Execute(7,caption,'') ; Exit ; END ;
if LaSelection.Count<=0 then BEGIN HMEnc.Execute(6,caption,'') ; Exit ; END ;
if TypeEtape.Value='' then BEGIN HMEnc.Execute(19,caption,'') ; Exit ; END ;
if ((Generation) and (JalFerme)) then BEGIN HMEnc.Execute(55,caption,'') ; Exit ; END ;
if RevisionActive(DateGene) then Exit ;
if Not LesMeme then Exit ;
if Not VerifLesEche then Exit ;
if Not VerifLesDevise then Exit ;
if Not VerifLesRIB then Exit ;
if Not VerifLesPrevi then Exit ;
if TypeEtape.Value='BQE' then
   BEGIN
   ii:=CoherGroupe ;  if ii>0 then BEGIN HMEnc.Execute(21+ii,caption,'') ; Exit ; END ;
   if (GLOBAL.Checked) and ((MP_CATEGORIE.Value='PRE') or (MP_CATEGORIE.Value='VIR')) then
      BEGIN
      if Not CoherVirPre(MP,DD) then
         BEGIN
         if Not AcqModeDate(MP,DD) then Exit else PreVirEgal:=True ;
         END ;
      END ;
   END ;
Case SorteLettre of
   tslCheque : if HMEnc.Execute(46,caption,'')<>mrYes then Exit ;
   tslBOR    : if HMEnc.Execute(49,caption,'')<>mrYes then Exit ;
   tslTraite : if HMEnc.Execute(50,caption,'')<>mrYes then Exit ;
    else if HMEnc.Execute(5-Ord(ENC),caption,'')<>mrYes then Exit ;
   END ;
Result:=False ;
END ;

{================================ GENERATION ==========================================}
function TFSuiviMP.MontantARegler ( O : TOBM ; Debit,AvecCouv : boolean ; Quoi : TSorteMontant ) : Double ;
var X : double ;
begin
Result:=0 ; X:=0 ;
Case Quoi of
    tsmPivot : BEGIN
               if Debit then X:=O.GetMvt('E_DEBIT') else X:=O.GetMvt('E_CREDIT') ;
               if ((X<>0) and (AvecCouv)) then X:=X-O.GetMvt('E_COUVERTURE') ;
               END ;
   tsmDevise : BEGIN
               if Debit then X:=O.GetMvt('E_DEBITDEV') else X:=O.GetMvt('E_CREDITDEV') ;
               if ((X<>0) and (AvecCouv)) then X:=X-O.GetMvt('E_COUVERTUREDEV') ;
               END ;
     tsmEuro : BEGIN
               if Debit then X:=O.GetMvt('E_DEBITEURO') else X:=O.GetMvt('E_CREDITEURO') ;
               if ((X<>0) and (AvecCouv)) then X:=X-O.GetMvt('E_COUVERTUREEURO') ;
               END ;
   END ;
Result:=X ;
end ;

function TFSuiviMP.MontantGuideSt (O : TOBM ; Debit : boolean ) : String ;
begin
if Debit then Result:=FloatToStr(GetLeDebitCur(O,True))
         else Result:=FloatToStr(GetLeCreditCur(O,True)) ;
end ;

procedure TFSuiviMP.LGPremLigne ( Q : TQuery ; O : TOBM ) ;
Var R  : RMVT ;
    T  : TStrings ;
    St : String ;
    VV : Variant ;
BEGIN
Q.FindField('EG_GENERAL').AsString:=O.GetMvt('E_GENERAL') ;
Q.FindField('EG_AUXILIAIRE').AsString:=O.GetMvt('E_AUXILIAIRE') ;
Q.FindField('EG_RIB').AsString:=O.GetMvt('E_RIB') ;
Q.FindField('EG_REFEXTERNE').AsString:=O.GetMvt('E_REFEXTERNE') ;
Q.FindField('EG_BANQUEPREVI').AsString:=O.GetMvt('E_BANQUEPREVI') ;
Q.FindField('EG_DEBITDEV').AsString:=MontantGuideSt(O,False) ;
Q.FindField('EG_CREDITDEV').AsString:=MontantGuideSt(O,True) ;
if MODEGENERE.Value<>'' then Q.FindField('EG_MODEPAIE').AsString:=MODEGENERE.Value
                        else Q.FindField('EG_MODEPAIE').AsString:=O.GetMvt('E_MODEPAIE') ;
if DateEcheance.Enabled then VV:=StrToDate(DateEcheance.Text) else VV:=O.GetMvt('E_DATEECHEANCE') ;
Q.FindField('EG_DATEECHEANCE').AsString:=VarAsType(VV,VarString) ;
if Q.FindField('EG_REFINTERNE').AsString='' then Q.FindField('EG_REFINTERNE').AsString:=O.GetMvt('E_REFINTERNE') ;
if Q.FindField('EG_LIBELLE').AsString='' then Q.FindField('EG_LIBELLE').AsString:=O.GetMvt('E_LIBELLE') ;
R:=OBMToIdent(O,True) ; St:=EncodeLC(R) ;
T:=TStringList.Create ; T.Add(St) ;
TMemoField(Q.FindField('EG_ECHEANCES')).Assign(T) ; T.Free ;
END ;

procedure TFSuiviMP.LGContreCommun ( Q : TQuery ; OGuiContre : TOBM ; Auxi,GeneMP,OldRef,OldLib,OldPrevi,OldRib,OldRefTire : String ; k : integer ) ;
BEGIN
if TypeEtape.Value='POR' then Q.FindField('EG_GENERAL').AsString:=GeneMP else
   BEGIN
   if ((NoBanque) or (Not Generation) or (TypeEtape.Value<>'BQE'))
      then Q.FindField('EG_GENERAL').AsString:=OGuiContre.GetMvt('EG_GENERAL')
      else Q.FindField('EG_GENERAL').AsString:=XX[k] ;
   END ;
if ((CollecContreP) and (NoBanque) and (TypeEtape.Value<>'BQE'))
   then Q.FindField('EG_AUXILIAIRE').AsString:=Auxi
   else Q.FindField('EG_AUXILIAIRE').AsString:='' ;
if ((OldRef<>'') and (OGuiContre.GetMvt('EG_REFINTERNE')='')) then Q.FindField('EG_REFINTERNE').AsString:=OldRef ;
if ((OldLib<>'') and (OGuiContre.GetMvt('EG_LIBELLE')='')) then Q.FindField('EG_LIBELLE').AsString:=OldLib ;
if ((OldRib<>'') and (OGuiContre.GetMvt('EG_RIB')='')) then Q.FindField('EG_RIB').AsString:=OldRib ;
if ((OldrefTire<>'') and (OGuiContre.GetMvt('EG_REFEXTERNE')='')) then Q.FindField('EG_REFEXTERNE').AsString:=OldRefTire ;
if OldPrevi<>'' then Q.FindField('EG_BANQUEPREVI').AsString:=OldPrevi ;
END ;

procedure TFSuiviMP.LGContreLigne ( Q : TQuery ; DCur,CCur : double ; OldEche : TDateTime ; OldPaie : String3 ) ;
Var VV : Variant ;
    Deci : integer ;
BEGIN
if E_DEVISE.Value<>V_PGI.DevisePivot then Deci:=DEV.Decimale else
   if CModeSaisie.Value='ECU' then Deci:=V_PGI.OkDecE else Deci:=V_PGI.OkDecV ;
if Abs(DCur)>Abs(CCur) then
   BEGIN
   DCur:=DCur-CCur ; CCur:=0 ;
   if ((TypeEtape.Value='BQE') and (DCur<0)) then BEGIN CCur:=-DCur ; DCur:=0 ; END ;
   END else
   BEGIN
   CCur:=CCur-DCur ; DCur:=0 ;
   if ((TypeEtape.Value='BQE') and (CCur<0)) then BEGIN DCur:=-CCur ; CCur:=0 ; END ;
   END ;
if DCur<>0 then Q.FindField('EG_DEBITDEV').AsString:=StrS(DCur,Deci) ;
if CCur<>0 then Q.FindField('EG_CREDITDEV').AsString:=StrS(CCur,Deci) ;
Q.FindField('EG_MODEPAIE').AsString:=OldPaie ;
VV:=OldEche ; Q.FindField('EG_DATEECHEANCE').AsString:=VarAsType(VV,VarString) ;
END ;

procedure TFSuiviMP.CreerLignesTemp ( k : integer ) ;
Var i,Lig,Lig2,NumL : integer ;
    OEcr,OGuiAux,OGuiContre : TOBM ;
    Q : TQuery ;
    St,Auxi,GeneMP,OldAux,OldGeneMP,OldRef,OldLib,OldPrevi,OldRefTire,OldRib : String ;
    DCur,CCur             : double ;
    OldPaie               : String3 ;
    OldEche               : TDateTime ;
    Rupt,Premier          : Boolean ;
    Junk,StB              : String ;
    TEche                 : TList ;
BEGIN
OGuiAux:=TOBM(TGuide[0]) ; if OGuiAux=Nil then Exit ;
OGuiContre:=TOBM(TGuide[1]) ; if OGuiContre=Nil then Exit ;
DCur:=0 ; CCur:=0 ; OEcr:=Nil ; OldEche:=0 ;
OldRef:='' ; OldLib:='' ; OldPrevi:='' ; OldRefTire:='' ; OldRib:='' ;
Q:=OpenSQL('Select * from ECRGUI WHERE EG_TYPE="RrR"',False) ; NumL:=0 ;
InitMove(LaSelection.Count,'') ; OldAux:='' ; OldGeneMP:='' ; Premier:=True ;
for i:=0 to LaSelection.Count-1 do
    BEGIN
    MoveCur(False) ;
    St:=LaSelection[i] ;
    Auxi:=Trim(ReadtokenSt(St)) ;
    GeneMP:=Trim(ReadTokenSt(St)) ;
    St:=Trim(St) ; Lig:=ReadTokenI(St) ;
    Junk:=ReadTokenSt(St) ; Junk:=ReadTokenSt(St) ;
    Lig2:=ReadTokenI(St) ;
//    if ((Lig<=0) or (Lig>=G.RowCount)) then Continue ;
    if TypeEtape.Value='POR' then CollecContreP:=EstCollectif(GeneMP) ;
    if ParTiers then
       begin
       TEche:=TList(G.Objects[0,Lig]) ;
       OEcr:=TOBM(TEche[Lig2-1]) ;
       end else OEcr:=TOBM(G.Objects[0,Lig]) ;
    StB:=OEcr.GetMvt('E_BANQUEPREVI') ; if StB='' then StB:=GENEGENERATION.Text ;
    if ((Not NoBanque) and (StB<>XX[k])) then Continue ;
    { Rupture sur Tiers ou contrepartie systématique}
    if Not Premier then
       BEGIN
       Rupt:=False ;
       if Not GLOBAL.Checked then Rupt:=True else
        if ((TypeEtape.Value='POR') and (OldGeneMP<>GeneMP)) then Rupt:=True else
         if ((TypeEtape.Value='BQE') or (TypeEtape.Value='EFF')) and (Not CollecContreP) then Rupt:=False else
          if Auxi<>OldAux then Rupt:=True ;
       if ((SorteLettre in [tslBOR,tslTraite]) and (TypeEtape.Value<>'BQE') and (Not CollecContreP)) then Rupt:=True ;
       if Rupt then
          BEGIN
          Q.Insert ; OGuiContre.EgalChamps(Q) ;
          Inc(NumL) ; LGCommun(Q,NumL) ;
          LGContreCommun(Q,OGuiContre,OldAux,OldGeneMP,OldRef,OldLib,OldPrevi,OldRib,OldRefTire,k) ;
          if PreVirEgal then LGContreLigne(Q,DCur,CCur,PreVirEche,PreVirMP)
                        else LGContreLigne(Q,DCur,CCur,OldEche,OldPaie) ;
          if NoBanque then Q.FindField('EG_GUIDE').AsString:=V_PGI.User
                      else Q.FindField('EG_GUIDE').AsString:=FormatFloat('000',k) ;
          Q.Post ;
          DCur:=0 ; CCur:=0 ;
          END ;
       END ;
    { Report 1 pour 1 de la première ligne }
    Q.Insert ; OGuiAux.EgalChamps(Q) ;
    Inc(NumL) ; LGCommun(Q,NumL) ; LGPremLigne(Q,OEcr) ;
    if NoBanque then Q.FindField('EG_GUIDE').AsString:=V_PGI.User
                else Q.FindField('EG_GUIDE').AsString:=FormatFloat('000',k) ;
    Q.Post ;
    DCur:=DCur+GetLeDebitCur(OEcr,True) ;
    CCur:=CCur+GetLeCreditCur(OEcr,True) ;
    OldAux:=Auxi ;
    if DateEcheance.Enabled then OldEche:=StrToDate(DateEcheance.Text) else OldEche:=OEcr.GetMvt('E_DATEECHEANCE') ;
    if MODEGENERE.Value<>'' then OldPaie:=MODEGENERE.Value else OldPaie:=OEcr.GetMvt('E_MODEPAIE') ;
    OldGeneMP:=GeneMP ; OldRefTire:=OEcr.GetMvt('E_REFEXTERNE') ; OldRib:=OEcr.GetMvt('E_RIB') ;
    if Not GLOBAL.Checked then
       BEGIN
       OldRef:=OEcr.GetMvt('E_REFINTERNE') ; OldLib:=OEcr.GetMvt('E_LIBELLE') ;
       OldPrevi:=OEcr.GetMvt('E_BANQUEPREVI') ;
       END ;
    Premier:=False ;
    END ;
FiniMove ;
if V_PGI.IOError<>oeOK then BEGIN Ferme(Q) ; Exit ; END ;
{ Dernière rupture }
Q.Insert ; OGuiContre.EgalChamps(Q) ;
Inc(NumL) ; LGCommun(Q,NumL) ;
LGContreCommun(Q,OGuiContre,OldAux,GeneMP,OldRef,OldLib,OldPrevi,OldRib,OldRefTire,k) ;
if PreVirEgal then LGContreLigne(Q,DCur,CCur,PreVirEche,PreVirMP)
              else LGContreLigne(Q,DCur,CCur,OldEche,OldPaie) ;
if NoBanque then Q.FindField('EG_GUIDE').AsString:=V_PGI.User
            else Q.FindField('EG_GUIDE').AsString:=FormatFloat('000',k) ;
Q.Post ;
Ferme(Q) ;
END ;

Function TFSuiviMP.CountBqe : integer ;
Var i,j,k : integer ;
    O   : TOBM ;
    Nb  : integer ;
    Trouv : boolean ;
    StB : String ;
    TEche : TList ;
BEGIN
Nb:=0 ; FillChar(XX,Sizeof(XX),#0) ; NoBanque:=False ; BanqueVide:=False ;
if ((ENC) or (SorteLettre in [tslBOR,tslTraite])) then BEGIN Result:=1 ; NoBanque:=True ; Exit ; END ;
if ((Not ENC) and (Generation) and (TypeEtape.Value<>'BQE')) then BEGIN Result:=1 ; NoBanque:=True ; Exit ; END ;
if ParTiers then
   begin
   for i:=1 to G.RowCount-1 do if EstSelect(G,i) then
      begin
      TEche:=TList(G.Objects[0,i]) ;
      for j:=0 to TEche.Count-1 do
         begin
         O:=TOBM(TEche[j]) ;
         if O.GetMvt('E_ETAT')='*' then
            BEGIN
            StB:=O.GetMvt('E_BANQUEPREVI') ;
            if StB='' then BEGIN StB:=GENEGENERATION.Text ; BanqueVide:=True ; END ;
            Trouv:=False ; for k:=1 to Nb do if StB=XX[k] then BEGIN Trouv:=True ; Break ; END ;
            if Not Trouv then BEGIN Inc(Nb) ; XX[Nb]:=StB ; END ;
            END ;
         end ;
      end ;
   end else
   begin
   for i:=1 to G.RowCount-1 do
      begin
      if EstSelect(G,i) then
         begin
         O:=TOBM(G.Objects[0,i]) ;
         StB:=O.GetMvt('E_BANQUEPREVI') ;
         if StB='' then BEGIN StB:=GENEGENERATION.Text ; BanqueVide:=True ; END ;
         Trouv:=False ; for k:=1 to Nb do if StB=XX[k] then BEGIN Trouv:=True ; Break ; END ;
         if Not Trouv then BEGIN Inc(Nb) ; XX[Nb]:=StB ; END ;
         END ;
      end ;
   end ;
if ((Nb=1) and ((XX[1]='') or (XX[1]=GENEGENERATION.Text))) then NoBanque:=True ;
Result:=Nb ;
END ;

Function TFSuiviMP.JalBanque ( k : integer ) : String ;
Var QJ : TQuery ;
BEGIN
Result:='' ; if XX[k]='' then Exit ;
QJ:=OpenSQL('Select J_JOURNAL from JOURNAL where J_CONTREPARTIE="'+XX[k]+'" AND J_NATUREJAL="BQE"',True) ;
if Not QJ.EOF then Result:=QJ.Fields[0].AsString ;
Ferme(QJ) ;
END ;

procedure TFSuiviMP.CreerGuideTemp ;
Var QS,QD  : TQuery ;
    i,Nb,k : integer ;
    JalB   : String ;
BEGIN
BanqueVide:=False ;
DetruitGuideEnc(True) ;
Nb:=CountBqe ;
if ENC then QS:=OpenSQL('Select * from GUIDE where GU_TYPE="ENC" AND GU_GUIDE="'+GUIDE.Value+'"',True)
       else QS:=OpenSQL('Select * from GUIDE where GU_TYPE="DEC" AND GU_GUIDE="'+GUIDE.Value+'"',True) ;
if Not QS.EOF then
   BEGIN
   for k:=1 to Nb do
       BEGIN
       QD:=OpenSQL('Select * from GUIDE WHERE GU_TYPE="Rer"',False) ;
       QD.Insert ; InitNew(QD) ;
       for i:=0 to QS.FieldCount-1 do QD.Fields[i].Assign(QS.Fields[i]) ;
       QD.FindField('GU_TYPE').AsString:=V_PGI.User ;
       if NoBanque then QD.FindField('GU_GUIDE').AsString:=V_PGI.User else
          BEGIN
          QD.FindField('GU_GUIDE').AsString:=FormatFloat('000',k) ;
          JalB:=JalBanque(k) ; if JalB<>'' then QD.FindField('GU_JOURNAL').AsString:=JalB ;
          END ;
       if EntGuide.CodeD='' then EntGuide.CodeD:=DEV.Code ;
       QD.FindField('GU_DEVISE').AsString:=EntGuide.CodeD ;
       if CModeSaisie.Value='ECU' then QD.FindField('GU_SAISIEEURO').AsString:='X' ;
       QD.Post ;
       Ferme(QD) ;
       if V_PGI.IOError=oeOK then CreerLignesTemp(k) ;
       END ;
   END else
   BEGIN
   V_PGI.IOError:=oeUnknown ;
   END ;
Ferme(QS) ;
END ;

procedure TFSuiviMP.ConstitueLaSelection ;
{Constitution de la liste de sélection avant validation}
var i,j   : integer ;
    TEche : TList ;
    O     : TOBM ;
begin
LaSelection.Clear ;
if ParTiers then
   begin
   for i:=1 to G.RowCount-1 do if EstSelect(G,i) then
      begin
      TEche:=TList(G.Objects[0,i]) ;
      for j:=0 to TEche.Count-1 do
         begin
         O:=TOBM(TEche[j]) ;
         if O.GetMvt('E_ETAT')='*' then AjouteOBM(O,i,j+1,True) ;
         end ;
      end ;
   end else
   begin
   for i:=1 to G.RowCount-1 do
      begin
      if EstSelect(G,i) then
         begin
         O:=TOBM(G.Objects[0,i]) ;
         AjouteOBM(O,i,0,True) ;
         end ;
      end ;
   end ;
end ;

procedure TFSuiviMP.UpdateCircuit ;
var O      : TOBM ;
    StS,CPG,CPA : String ;
    i,j    : Integer ;
    TEche  : TList ;
begin
if not PCircuit.TabVisible then exit ;
if ParTiers then {masse}
   begin
   for i:=1 to G.RowCount-1 do if EstSelect(G,i) then
      begin
      TEche:=TList(G.Objects[0,i]) ;
      for j:=0 to TEche.Count-1 do
         begin
         O:=TOBM(TEche[j]) ;
         if O.GetMvt('E_ETAT')='*' then
            begin
            {Restore infos "bricolées" en mémoire}
            O.PutMvt('E_LIBRETEXTE0',O.GetMvt('E_MULTIPAIEMENT')) ;
            StS:=O.GetMvt('E_TRACE') ; CPG:=ReadTokenSt(StS) ; CPA:=ReadTokenSt(StS) ;
            O.PutMvt('E_CONTREPARTIEGEN',CPG) ; O.PutMvt('E_CONTREPARTIEAUX',CPA) ;
            O.PutMvt('E_MULTIPAIEMENT','') ; O.PutMvt('E_TRACE','') ;
            {Maj circuit}
            O.PutMvt('E_SUIVDEC',CircuitDest.Value) ;
            O.PutMvt('E_NOMLOT',LotDest.Text) ;
            O.PutMvt('E_FLAGECR',V_PGI.User) ;
            UpdateOBMCircuit(O) ;
            if V_PGI.IOError<>oeOK then break ;
            end ;
         end ;
      if V_PGI.IOError<>oeOK then break ;
      end ;
   end else {détail}
   begin
   for i:=1 to G.RowCount-1 do
      begin
      if EstSelect(G,i) then
         begin
         O:=TOBM(G.Objects[0,i]) ;
         {Restore infos "bricolées" en mémoire}
         O.PutMvt('E_LIBRETEXTE0',O.GetMvt('E_MULTIPAIEMENT')) ;
         StS:=O.GetMvt('E_TRACE') ; CPG:=ReadTokenSt(StS) ; CPA:=ReadTokenSt(StS) ;
         O.PutMvt('E_CONTREPARTIEGEN',CPG) ; O.PutMvt('E_CONTREPARTIEAUX',CPA) ;
         O.PutMvt('E_MULTIPAIEMENT','') ; O.PutMvt('E_TRACE','') ;
         {Maj circuit}
         O.PutMvt('E_SUIVDEC',CircuitDest.Value) ;
         O.PutMvt('E_NOMLOT',LotDest.Text) ;
         O.PutMvt('E_FLAGECR',V_PGI.User) ;
         UpdateOBMCircuit(O) ;
         if V_PGI.IOError<>oeOK then break ;
         end ;
      end ;
   end ;
end ;

procedure TFSuiviMP.UpdateOBMCircuit(O : TOBM) ;
var St,SQL : string ;
    M      : RMVT ;
    NumP   : Longint ;
begin
St:=O.StPourUpdate ; if St='' then Exit ;
M:=OBMToIdent(O,True) ; NumP:=O.GetMvt('E_NUMEROPIECE') ;
if NumP>MaxNum then MaxNum:=NumP ;
if NumP<MinNum then MinNum:=NumP ;
SQL:='UPDATE ECRITURE SET '+St+', E_DATEMODIF="'+USTime(NowFutur)+'"' ;
SQL:=SQL+' Where  '+WhereEcriture(tsGene,M,True)+' AND E_DATEMODIF="'+USTime(O.GetMvt('E_DATEMODIF'))+'"' ;
if ExecuteSQL(SQL)<>1 then V_PGI.IOError:=oeUnknown ;
end ;

procedure TFSuiviMP.FabricLesM ;
Var Q : TQuery ;
    Code : String ;
    P   : P_MV ;
    M   : RMVT ;
    k   : integer ;
BEGIN
{Init communs}
FillChar(M,Sizeof(M),#0) ; VideListe(LesM) ;
M.DateC:=DateGene ; M.Exo:=QuelExo(DATEGENERATION.Text) ;
M.Simul:='N' ; M.Valide:=False ;
M.MajDirecte:=((CAutoEnreg.Checked) and (Generation) and (Not ENC) and (SorteLettre=tslAucun)) ;
if ENC then M.TypeGuide:='ENC' else M.TypeGuide:='DEC' ;
M.ExportCFONB:=OKCFONB ;
M.FormatCFONB:=FormatCFONB ; M.EnvoiTrans:=EnvoiTrans ;
If Not VH^.OldTeleTrans Then M.EnvoiTrans:='' ;
M.Bordereau:=OKBordereau ; M.Document:=Document ;
M.SorteLettre:=SorteLettre ;
M.Globalise:=((GLOBAL.Checked) and (SorteLettre in [tslAucun,tslCheque])) ;
M.CodeD:=DEV.Code ; M.DateTaux:=DateGene ; M.TauxD:=TauxRef ;
M.ModeOppose:=(CModeSaisie.Value='ECU') ;
{Recherche du ou des guides créés}
Q:=OpenSQL('SELECT GU_GUIDE, GU_TYPE from GUIDE Where GU_TYPE="'+V_PGI.User+'" Order By GU_TYPE, GU_GUIDE',True) ;
While Not Q.EOF do
   BEGIN
   if NoBanque then
      BEGIN
      M.LeGuide:=V_PGI.User ; {ruse pour contreparties}
      if TypeEtape.Value='BQE' then M.General:=GENEGENERATION.Text else
         if TypeEtape.Value<>'POR' then M.Section:=GENEGENERATION.Text ;
      END else
      BEGIN
      Code:=Q.FindField('GU_GUIDE').AsString ; k:=StrToInt(Copy(Code,3,1)) ;
      M.LeGuide:=Code ; M.General:=XX[k] ;
      END ;
   P:=P_MV.Create   ; P.R:=M ; LesM.Add(P) ;
   Q.Next ;
   END ;
Ferme(Q) ;
END ;

procedure TFSuiviMP.BValideClick(Sender: TObject);
Var M  : RMVT ;
    io : TIOErr ;
    s,SQL : string ;
begin
NowFutur:=NowH ;
if (CircuitDest.Value<>'') and (not Generation) then
   begin
   if Trim(LotDest.Text)='' then begin HMEnc.Execute(35,caption,'') ; Exit ; end ;
   if HMEnc.Execute(54,caption,' '+Trim(LotDest.Text)+' ?')<>mrYes then Exit ;
   if Transactions(UpdateCircuit,3)=oeOk then
      BEGIN
      if (cBAP.Checked) and (cModeleBAP.Value<>'') then
         begin
         s:='E_FLAGECR="'+V_PGI.User+'" AND E_GENERAL="'+E_GENERAL.Text+'" '
           +'AND E_NUMEROPIECE>='+IntToStr(MinNum)+' AND E_NUMEROPIECE<='+IntToStr(MaxNum) ;
         LanceEtat('E','BAP',cModeleBAP.Value,True,False,False,nil,s,'',False) ;
         end ;
      // RAZ de E_FLAGECR
      SQL:='UPDATE ECRITURE SET E_FLAGECR="" WHERE E_GENERAL="'+E_GENERAL.Text+'" '
          +'AND E_QUALIFPIECE="N" AND E_ECHE="X" AND E_NUMECHE>0 AND E_ETATLETTRAGE<>"TL" '
          +'AND E_NUMEROPIECE>='+IntToStr(MinNum)+' AND E_NUMEROPIECE<='+IntToStr(MaxNum)+' '
          +'AND E_FLAGECR="'+V_PGI.User+'"' ;
      ExecuteSQL(SQL) ;
      if Not Generation then HMEnc.Execute(38,caption,'') ;
      END else
      BEGIN
      MessageAlerte(HDiv.Mess[2]) ;
      Exit ;
      END ;
   end ;
if Generation then
   begin
   { Validation normale }
   ConstitueLaSelection ;
   if JeValidePas then Exit ;
   Application.ProcessMessages ;
   io:=Transactions(CreerGuideTemp,5) ;
   Case io of
      oeOK      : ;
      oeUnknown : BEGIN MessageAlerte(HDiv.Mess[1]) ; Exit ; END ;
      oeSaisie  : BEGIN MessageAlerte(HDiv.Mess[2]) ; Exit ; END ;
      END ;
   FabricLesM ; G.VidePile(True) ;
   if ((Generation) and (Not ENC)) then LanceMultiSaisie(LesM) else
      BEGIN
      M:=P_MV(LesM[0]).R ; M.MajDirecte:=False ;
      LanceSaisie(Nil,taCreat,M) ;
      END ;
   end ;
Application.ProcessMessages ;
BChercherClick(Nil) ;
end;

procedure TFSuiviMP.SwapCat ( Vide : Boolean ) ;
Var ValCat,ValMode : String ;
BEGIN
ValCat:=MP_CATEGORIE.Value ; ValMode:=E_MODEPAIE.Value ;
MP_CATEGORIE.Vide:=Vide ; MP_Categorie.Enabled:=True ;
MP_CATEGORIE.Reload ;
MP_CATEGORIE.Value:=ValCat ; E_MODEPAIE.Value:=ValMode ;
END ;

procedure TFSuiviMP.PositionneBB ;
BEGIN
if ((Circuit2='') and (AffectBanque)) then
   BEGIN
   if ((Not ENC) and (SorteLettre=tslBOR)) then BAffectBqe.Enabled:=True else
   if ((ENC) and (SorteLettre=tslTraite)) then BAffectBqe.Enabled:=True else
      BAffectBqe.Enabled:=(TypeEtape.Value='BQE') ;
   END ;
END ;

procedure TFSuiviMP.ETAPEChange(Sender: TObject);
Var Q : TQuery ;
    StW,TypeET : String ;
    Okok,OkCpte : boolean ;
begin
G.Enabled:=False ; Okok:=True ; OkCFONB:=False ; FormatCFONB:='' ; EnvoiTrans:='' ; 
OkBordereau:=False ; Document:='' ; JalFerme:=False ;
if ETAPE.Value='' then Exit ;
if ENC then StW:='ER_ENCAISSEMENT="X"' else StW:='ER_ENCAISSEMENT="-"' ;
Q:=OpenSQL('Select * from ETAPEREG Where '+StW+' AND ER_ETAPE="'+ETAPE.Value+'"',True) ;
if Not Q.EOF then
   BEGIN
   E_GENERAL.Text:=Q.FindField('ER_CPTEDEPART').AsString ;
   TypeET:=Q.FindField('ER_TYPEETAPE').AsString ;
   if ((TYPEET='POR') or (TYPEET='EFF')) then
      BEGIN
      if ENC then GENEGENERATION.ZoomTable:=tzGBQCaissCli else GENEGENERATION.ZoomTable:=tzGBQCaissFou ;
      END else if TYPEET='BQE' then
      BEGIN
      GENEGENERATION.ZoomTable:=tzGBanque ;
      END ;
   Okok:=((Okok) and (E_GENERAL.ExisteH>0)) ;
   if Okok then
      BEGIN
      GENEGENERATION.Text:=Q.FindField('ER_CPTEARRIVEE').AsString ;
      if (GENEGENERATION.Text<>'') then OkCpte:=(GENEGENERATION.ExisteH>0) else OkCpte:=(TypeET='POR') ;
      Okok:=((Okok) and (OkCpte)) ;
      if Okok then
         BEGIN
         MP_CATEGORIE.Value:='' ;
         if ((ENC) and (SorteLettre=tslAucun)) then MP_CATEGORIE.Value:=Q.FindField('ER_CATEGORIEMP').AsString else
          if ((NOT ENC) and (Generation) and (SorteLettre=tslAucun)) then MP_CATEGORIE.Value:=Q.FindField('ER_CATEGORIEMP').AsString ;
         if Generation then E_MODEPAIE.Value:=Q.FindField('ER_MODEPAIE').AsString ;
         E_DEVISE.Value:=Q.FindField('ER_DEVISE').AsString ;
         GUIDE.Reload ; GUIDE.Value:=Q.FindField('ER_GUIDE').AsString ;
         GLOBAL.Checked:=(Q.FindField('ER_GLOBALISE').AsString='X') ;
         OkCFONB:=(Q.FindField('ER_EXPORTCFONB').AsString='X') ;
         FormatCFONB:=Q.FindField('ER_FORMATCFONB').AsString ;
{$IFNDEF SPEC302}
         EnvoiTrans:=Q.FindField('ER_ENVOITRANS').AsString ;
         If Not VH^.OldTeleTrans Then EnvoiTrans:='' ;
{$ENDIF}         
         OkBordereau:=(Q.FindField('ER_BORDEREAU').AsString='X') ;
         Document:=Q.FindField('ER_DOCUMENT').AsString ;
         CodeAFB.Text:=Q.FindField('ER_CODEAFB').AsString ;
         TYPEETAPE.Value:=TypeET ;
         GereAccept.Visible:=((ENC) and (TypeET='BQE')) ;
         PositionneBB ;
         END ;
      END ;
   END ;
Ferme(Q) ;
if Not Okok then
   BEGIN
   E_GENERAL.Text:='' ; GENEGENERATION.Text:='' ;
   HMEnc.Execute(29,caption,'') ;
   Etape.Value:='' ; Exit ;
   END ;
E_GENERAL.Enabled:=True ; DateGeneration.Enabled:=True ;
Global.Enabled:=(SorteLettre in [tslAucun,tslCheque]) ;
if TYPEETAPE.Value='POR' then
   BEGIN
   GENEGENERATION.Enabled:=False ; GENEGENERATION.Text:='**********' ; H_GENEGENERATION.Caption:=HDiv.Mess[0] ;
   END else if ETAPE.Value='EFF' then
   BEGIN
   GENEGENERATION.Enabled:=True ;
   END else if ETAPE.Value='BQE' then
   BEGIN
   GENEGENERATION.Enabled:=False ;
   END ;
if SorteLettre<>tslAucun then
   BEGIN
   MP_CATEGORIE.Vide:=True ; MP_CATEGORIE.Value:='' ; MP_Categorie.Enabled:=False ;
   MODEGENERE.Visible:=False ; MODEGENERE.Value:='' ; H_MODEGENERE.Visible:=False ;
   END else if OkCFONB then
   BEGIN
   SwapCat(False) ;
   MODEGENERE.Visible:=False ; MODEGENERE.Value:='' ; H_MODEGENERE.Visible:=False ;
   END else
   BEGIN
   SwapCat(True) ;
   MODEGENERE.Visible:=True ; H_MODEGENERE.Visible:=True ;
   END ;
if SorteLettre=tslAucun then MajCaption ;
GuideChange(Nil) ;
end;

procedure TFSuiviMP.E_GENERALExit(Sender: TObject);
Var St : String ;
begin
if csDestroying in ComponentState then Exit ;
St:=E_GENERAL.Text ; if St='' then BEGIN H_GENERAL.Caption:='' ; Exit ; END ;
if St=GENEGENERATION.Text then BEGIN HMEnc.Execute(28,caption,'') ; Exit ; END ;
E_GENERAL.ExisteH ;
if GUIDE.Value='' then Exit ;
if ENC then EXECUTESQL('UPDATE ECRGUI SET EG_GENERAL="'+St+'" WHERE EG_TYPE="ENC" AND EG_GUIDE="'+GUIDE.Value+'" AND EG_NUMLIGNE=1')
       else EXECUTESQL('UPDATE ECRGUI SET EG_GENERAL="'+St+'" WHERE EG_TYPE="DEC" AND EG_GUIDE="'+GUIDE.Value+'" AND EG_NUMLIGNE=1') ;
GuideChange(Nil) ;
end;

procedure TFSuiviMP.GENEGENERATIONExit(Sender: TObject);
Var St : String ;
begin
if csDestroying in ComponentState then Exit ;
St:=GENEGENERATION.Text ; if St='' then Exit ;
if St=E_GENERAL.Text then BEGIN HMEnc.Execute(28,caption,'') ; Exit ; END ;
if GUIDE.Value='' then Exit ;
if TypeEtape.Value='POR' then Exit ;
if ENC then EXECUTESQL('UPDATE ECRGUI SET EG_GENERAL="'+St+'" WHERE EG_TYPE="ENC" AND EG_GUIDE="'+GUIDE.Value+'" AND EG_NUMLIGNE=2')
       else EXECUTESQL('UPDATE ECRGUI SET EG_GENERAL="'+St+'" WHERE EG_TYPE="DEC" AND EG_GUIDE="'+GUIDE.Value+'" AND EG_NUMLIGNE=2') ;
GuideChange(Nil) ;
END ;

procedure TFSuiviMP.BZOOMETAPEClick(Sender: TObject);
Var VV : String3 ;
    OldG : Boolean ;
begin
VV:=Etape.Value ; ParamEtapeReg(ENC,VV,False) ;
OldG:=V_PGI.GestionMessage ; V_PGI.GestionMessage:=True ;
Etape.Reload ; Etape.Value:=VV ;
V_PGI.GestionMessage:=OldG ; 
end;

procedure TFSuiviMP.BModifEcheClick(Sender: TObject);
Var M : RMVT ;
    EU    : T_ECHEUNIT ;
    TAN   : String3 ;
    Lig,k : integer ;
    O     : TOBM ;
    Aux,Coll : string ;
    TotDebCur,TotCredCur,TotCouvCur : Double ;
    TEche : TList ;
begin
Lig:=G.Row ; if ((Lig<=0) or (Lig>G.RowCount-1)) then Exit ;
if ParTiers then
   begin {Liste des échéances du tiers}
   Aux:=G.Cells[1,Lig] ;
   TotDebCur:=0 ; TotCredCur:=0 ; TotCouvCur:=0 ;
   TEche:=TList(G.Objects[0,Lig]) ; if TEche=nil then exit ;
   if EncaDecaTiers(Aux,TEche,TotDebCur,TotCredCur,TotCouvCur,Enc,DEV,SorteLettre<>tslAucun,(CModeSaisie.Value='ECU')) then
      begin
      G.Cells[3,Lig]:=FormatFloat(G.ColFormats[3],TotDebCur) ;
      G.Cells[4,Lig]:=FormatFloat(G.ColFormats[4],TotCredCur) ;
      G.Cells[5,Lig]:=FormatFloat(G.ColFormats[5],TotCouvCur) ;
      // Sélection dynamique du tiers
      if TotDebCur+TotCredCur<>0 then G.Cells[G.ColCount-1,Lig]:='+' else G.Cells[G.ColCount-1,Lig]:=' ' ;
      CalculDebitCredit ;
      end ;
   end else
   begin
   O:=TOBM(G.Objects[0,Lig]) ; if O=Nil then Exit ;
   if O.GetMvt('E_ANA')='X' then Exit ; 
   M:=OBMToIdent(O,True) ;
   FillChar(EU,Sizeof(EU),#0) ;
   EU.DateEche:=O.GetMvt('E_DATEECHEANCE') ; EU.ModePaie:=O.GetMvt('E_MODEPAIE') ;
   EU.DebitDEV:=O.GetMvt('E_DEBITDEV') ; EU.CreditDEV:=O.GetMvt('E_CREDITDEV') ;
   EU.DebitEuro:=O.GetMvt('E_DEBITEURO') ; EU.CreditEuro:=O.GetMvt('E_CREDITEURO') ;
   EU.Debit:=O.GetMvt('E_DEBIT') ; EU.Credit:=O.GetMvt('E_CREDIT') ;
   EU.DEVISE:=O.GetMvt('E_DEVISE') ; EU.TauxDEV:=O.GetMvt('E_TAUXDEV') ;
   EU.DateComptable:=O.GetMvt('E_DATECOMPTABLE') ;
   EU.DateModif:=O.GetMvt('E_DATEMODIF') ;
   EU.SaisieEuro:=O.GetMvt('E_SAISIEEURO')='X' ;
   EU.ModeSaisie:=O.GetMvt('E_MODESAISIE') ;
   {#TVAENC}
   if VH^.OuiTvaEnc then
      BEGIN
      Coll:=O.GetMvt('E_GENERAL') ;
      if EstCollFact(Coll) then
         BEGIN
         for k:=1 to 4 do EU.TabTva[k]:=O.GetMvt('E_ECHEENC'+IntToStr(k)) ;
         EU.TabTva[5]:=O.GetMvt('E_ECHEDEBIT') ;
         END ;
      END ;
   TAN:=O.GetMvt('E_ECRANOUVEAU') ;
   if TAN='OAN' then
      BEGIN
      if M.CodeD<>V_PGI.DevisePivot then Exit ;
      if ((VH^.EXOV8.Code<>'') and (M.DateC<VH^.EXOV8.Deb)) then Exit ;
      END ;
   if ModifUneEcheance(M,EU) then BChercherClick(Nil) ;
   end ;
end;

procedure TFSuiviMP.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFSuiviMP.E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
begin
ParamDate(Self,Sender,Key) ;
end;

procedure TFSuiviMP.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFSuiviMP.BAnnulerClick(Sender: TObject);
begin Close ; end;

procedure TFSuiviMP.cTypeVisuChange(Sender: TObject);
begin
OrdreTri.Visible:=(rDetail.Checked) ;
H_OrdreTri.Visible:=OrdreTri.Visible ;
BModifRIB.Enabled:=((Generation) and (OrdreTri.Visible)) ;
end;

procedure TFSuiviMP.GRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin
end;

{------------------ Affectation prévisionnelle des banques --------------------}

procedure TFSuiviMP.AddGridAffBqe(G : THGrid ; O : TOBM) ;
begin
// Ajoute l'échéance dans le grid détail d'affectation
if AffectBanque then O.PutMvt('E_NUMPIECEINTERNE',O.GetMvt('E_BANQUEPREVI')) ;
if G.Cells[0,G.RowCount-1]<>'' then G.RowCount:=G.RowCount+1 ;
ComCom1(G,O) ;
G.Objects[0,G.RowCount-1]:=O ;
end ;

procedure TFSuiviMP.DoAffectBqe ;
var X     : TFAffBqe ;
    i,j,n : Integer ;
    TEche : TList ;
    O     : TOBM ;
begin
X:=TFAffBqe.Create(Application) ;
if not AffectBanque then X.cVisu.Checked:=True ;
X.Devise.Text:=E_DEVISE.Value ; n:=0 ;
if ParTiers then
   begin
   for i:=1 to G.RowCount-1 do if EstSelect(G,i) then
      begin
      TEche:=TList(G.Objects[0,i]) ;
      for j:=0 to TEche.Count-1 do
         begin
         O:=TOBM(TEche[j]) ;
         if O.GetMvt('E_ETAT')='*' then begin AddGridAffBqe(X.G,O) ; Inc(n) ; end ;
         end ;
      end ;
   end else
   begin
   for i:=1 to G.RowCount-1 do
      begin
      if EstSelect(G,i) then
         begin
         O:=TOBM(G.Objects[0,i]) ;
         AddGridAffBqe(X.G,O) ;
         Inc(n) ;
         end ;
      end ;
   end ;
   try
   if n>0 then X.ShowModal ;
   finally
   X.Free ;
   end ;
end ;

procedure TFSuiviMP.bAffectBqeClick(Sender: TObject);
begin
if Not BAffectBqe.Enabled then Exit ;
DoAffectBqe ;
end;

procedure TFSuiviMP.bListeLotsClick(Sender: TObject);
var sWhere : String ;
begin
if Trim(E_GENERAL.Text)='' then begin HMEnc.Execute(27,caption,'') ; Exit ; end ;
if E_SUIVDEC.Value='' then begin HMEnc.Execute(26,caption,'') ; Exit ; end ;
sWhere:='E_GENERAL="' + E_GENERAL.Text + '" AND E_SUIVDEC="' + E_SUIVDEC.Value + '"' ;
E_NOMLOT.Text:=Choisir(HMEnc.Mess[41],'ECRITURE','DISTINCT E_NOMLOT','',sWhere,'E_NOMLOT') ;
end;

procedure TFSuiviMP.cBAPClick(Sender: TObject);
begin
cModeleBAP.Enabled:=cBAP.Checked ;
end;

procedure TFSuiviMP.GereOngletCircuit ( Montrer : boolean ) ;
BEGIN
if SorteLettre in [tslCheque,tslBOR,tslAucun] then
   begin
   PCircuit.TabVisible:=cVoirCircuit.Checked ;
   if CVoirCircuit.Checked then
      BEGIN
      E_SUIVDEC.Value:=Circuit1 ;
      END else
      BEGIN
      E_SUIVDEC.Value:='' ; E_NOMLOT.Text:='' ;
      END ;
   end ;
if Montrer then Pages.ActivePage:=PEcritures else Pages.ActivePage:=PComptes ; 
END ;

procedure TFSuiviMP.cVoirCircuitClick(Sender: TObject);
begin
GereOngletCircuit(True) ;
end;

procedure TFSuiviMP.CChoixDateClick(Sender: TObject);
begin
if CChoixDate.Checked then
   BEGIN
   DateEcheance.Enabled:=True ;
   if DateEcheance.Text=StDate1900 then DateEcheance.Text:=DateToStr(V_PGI.DateEntree) ;
   END else
   BEGIN
   DateEcheance.Text:=StDate1900 ;
   DateEcheance.Enabled:=False ;
   END ;
end;

procedure TFSuiviMP.E_EXERCICEChange(Sender: TObject);
begin
ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

procedure TFSuiviMP.POPFPopup(Sender: TObject);
begin
UpdatePopFiltre(BSaveFiltre,BDelFiltre,BRenFiltre,FFiltres) ;
end;

procedure TFSuiviMP.BModifRIBClick(Sender: TObject);
Var Lig : integer ;
    O   : TOBM ;
    RIB,Aux : String ;
    MM      : RMVT ;
begin
if Not Generation then Exit ;
if ParTiers then Exit ;
Lig:=G.Row ; if ((Lig<=0) or (Lig>G.RowCount-1)) then Exit ;
O:=TOBM(G.Objects[0,Lig]) ; if O=Nil then Exit ;
Aux:=O.GetMvt('E_AUXILIAIRE') ; if Aux='' then Exit ;
RIB:=O.GetMvt('E_RIB') ;
if ModifLeRIB(RIB,Aux) then
   BEGIN
   O.PutMvt('E_RIB',RIB) ;
   MM:=OBMToIdent(O,True) ;
   ExecuteSQL('UPDATE ECRITURE SET E_RIB="'+RIB+'" Where '+WhereEcriture(tsGene,MM,True)) ;
   END ;
end;

procedure TFSuiviMP.BZoomModeleClick(Sender: TObject);
Var QQ : TQuery ;
    Modele : String ;
begin
if SorteLettre<>tslCheque then Exit ;
if ETAPE.Value='' then Exit ;
if GeneGeneration.Text='' then Exit ;
FicheBanqueCP(GeneGeneration.text,taModif,0) ;
end;

end.
