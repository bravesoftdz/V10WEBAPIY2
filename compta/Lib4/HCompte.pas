unit Hcompte;

interface

uses
  {$IFNDEF EAGLSERVER}
  Graphics,
  Controls,
  Forms,
  Dialogs,
  StdCtrls,
  Buttons,
  ExtCtrls,
  Grids,
  HSysMenu,
  hmsgbox,
  {$ENDIF}
  SysUtils, WinTypes, WinProcs, Messages, Classes,  Hctrls
{$IFDEF EAGLCLIENT}
  ,uTOB
{$ELSE}
  ,DB
  {$IFNDEF DBXPRESS} ,dbtables {$ELSE} ,uDbxDataSet {$ENDIF}
{$ENDIF}
{$IFNDEF ERADIO}
  {$IFNDEF EAGLSERVER}
    ,LookUp // Pour le Lookuplist
  {$ENDIF !EAGLSERVER}
{$ENDIF !ERADIO}
  ,HEnt1
  ,Ent1
  {$IFDEF MODENT1}
  , CPTypeCons
  , CPProcMetier
  {$ENDIF MODENT1}

  ,Paramsoc
  ,UtilPGI,uEntCommun, TntGrids, TntButtons
  ;


{$IFDEF EAGLCLIENT}
type TPROCZOOM = Procedure (Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
Var ProcZoomGene    : TProcZoom ;
    ProcZoomTiers   : TProcZoom ;
    ProcZoomSection : TProcZoom ;
    ProcZoomJal     : TProcZoom ;
    ProcZoomBudGen  : TProcZoom ;
    ProcZoomBudSec  : TProcZoom ;
    ProcZoomBudJal  : TProcZoom ;
    ProcZoomCorresp : TProcZoom ;
    ProcZoomNatCpt  : TProcZoom ;
    ProcZoomRub     : TProcZoom ;
{$ENDIF}


Type T5B = Array[1..5] Of Boolean ;

Type TGGeneral = Class
       General          : String ;
       Libelle          : String ;
       Abrege           : String ;
       NatureGene       : String3 ;
       RegimeTVA        : String ;
       Collectif        : Boolean ;
       Ventilable       : T5B ;
       Lettrable        : Boolean ;
       TotalDebit       : Double ;
       TotalCredit      : Double ;
       Budgene          : String ;
       Tva              : String ;
       Tpf              : String ;
       SoumisTPF        : Boolean ;
       ModeRegle        : String ;
       Tva_Encaissement : String3 ;
       CpteTVA,CpteTPF  : String ;
       TauxTVA,TauxTPF  : Double ;
       Pointable        : Boolean ;
       Ferme            : boolean ;
       Sens             : byte ;
       QQ1,QQ2          : String ;
       SuiviTreso,CodeConso : String3 ;
       TvaSurEncaissement : boolean ;
       Confidentiel       : String[1] ;
       Effet              : boolean ;
       constructor Create(StCompte : String) ;
       Procedure QueryToGeneral  ( G_Query : TQuery ) ;
       // Penser à SSAI
       end ;

Type TGTiers = Class
       Auxi          : String ;
       Libelle       : String ;
       Abrege        : String ;
       Collectif     : String ;
       NatureAux     : String3 ;
       Lettrable     : Boolean ;
       TotalDebit    : Double ;
       TotalCredit   : Double ;
       ModeRegle     : String ;
       JourPaiement1 : Integer ;
       JourPaiement2 : Integer ;
       RegimeTva     : String ;
       SoumisTPF,ReleveFacture,Ferme : Boolean ;
       Tva_Encaissement : String3 ;
       Confidentiel     : String[1] ;
       Devise,CodeConso : String3 ;
       TL0              : String17 ;
       MultiDevise      : boolean ;
       IsPayeur,DebrayePayeur,AvoirRbt : Boolean ;
       AuxiPayeur       : String ;
       constructor Create(StCompte : String) ;
       Procedure QueryToTiers  ( T_Query : TQuery ) ;
       // Penser à SSAI
       end ;

Type TGSection = Class
       Axe          : String ;
       Sect         : String ;
       Libelle      : String ;
       Abrege       : String ;
       BudSect      : String ;
       TotalDebit   : Double ;
       TotalCredit  : Double ;
       Ferme        : boolean ;
       Confidentiel : String[1] ;
       constructor create(StCompte : String ; StAxe : TZoomTable) ;
       Procedure QueryToSection ( S_Query : TQuery ) ;
       procedure QueryY2ToSection(S_Query: TQuery);
       end ;

Type TGBudSect = Class
       Axe          : String ;
       BudSect      : String ;
       Libelle      : String ;
       Abrege       : String ;
       Signe        : String ;
       Sens         : String ;
       Ferme        : Boolean ;
       Confidentiel : String[1] ;
       constructor create(StCompte : String ; StAxe : TZoomTable) ;
       Procedure QueryToBudSect ( BS_Query : TQuery ) ;
       end ;

Type TGBudGene = Class
       BudGene      : String ;
       Libelle      : String ;
       Abrege       : String ;
       Signe        : String ;
       Sens         : String ;
       Ferme        : Boolean ;
       Confidentiel : String[1] ;
       constructor Create(StCompte : String) ;
       Procedure QueryToBudGene ( BG_Query : TQuery ) ;
       end ;

{$IFNDEF EAGLSERVER}
type
  THCpteEdit = class(TEdit)
  private
    FVide         : Boolean ;
    FBourre       : Boolean ;
    FZoomTable    : TZoomTable ;
    FLibelle      : TLabel ;
    FOkLocate     : Boolean ;
    FSynPlus      : String ;
    FSynJoker     : Boolean ;
    Procedure SetFBourre ( Value : Boolean ) ;
    Procedure SetFSynJoker ( Value : Boolean ) ;
    Procedure SetZoomTable ( Value : TZoomTable ) ;
    Function  ZoomSousPlanAna : Boolean ;
  protected
    procedure DoExit ; override ;
  public
    QuelleRef     : String ;
    TypesArtAut   : String ; // Types d'articles autorisés
    TypesTiersAut : String ; // Types de tiers autorisés
    procedure DblClick; override ;
    procedure keydown(var Key: Word; Shift: TShiftState); Override ;
    function  ExisteH : Byte ;
  published
    Property ZoomTable : TZoomTable Read FZoomTable Write SetZoomTable ;
    Property Vide      : Boolean Read FVide Write FVide ;
    Property Bourre    : Boolean Read FBourre Write SetFBourre ;
    Property Libelle   : TLabel Read FLibelle Write FLibelle ;
    Property okLocate  : Boolean Read FOkLocate Write FOkLocate ;
    Property SynPlus   : String Read FSynPlus Write FSynPlus ;
    Property SynJoker  : Boolean Read FSynJoker Write SetFSynJoker ;
  end;

type
  TFZoom = class(TForm)
    HPanel1: TPanel;
    HPanel2: TPanel;
    OKBtn: THBitBtn;
    CancelBtn: THBitBtn;
    HelpBtn: THBitBtn;
    TCompte: TLabel;
    TLibelle: TLabel;
    SelectCompte: TEdit;
    Selectlib: TEdit;
    Appelf: THBitBtn;
    HM: THMsgBox;
    CreatBn: THBitBtn;
    HMTrad: THSystemMenu;
    FListe	: THGrid;
    procedure APPELFICHE(Sender: TObject);
    procedure CREATEFICHE(Sender: TObject);
    procedure SelectCompteChange(Sender: TObject);
    procedure TCompteMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure FListeDblClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    Constructor MyCreateNew( AOwner : TComponent ; Fic : TZoomTable ; StCompte : String ; PP : TProcZoom ; OkLib : boolean ; StRub : String ) ;
  private { Private declarations }
    FonctionZoom : TProcZoom ; // EAGL : EN ATTENTE
    FNomTable,FNomCode,FNomLibelle,FWhere,FWhereSup : String ;
    ZoomTable : TZoomTable ;
    FDebraye,Croissant : Boolean ;
    CptLocate : String ;
    SynPlus   : String ;
    procedure souligne(lab : Tlabel ; ok,croissant : boolean) ;
    PROCEDURE LesNoms ;
    Procedure AfficheTitre(Lefb : TFichierBase) ;
    Procedure RemplirGrille(stReq : String);
    Procedure Localiser(stCode : String);
  end;

Type TChildForm = Class(TForm)
  protected
    procedure Loaded ; override ;
    procedure CreateParams ( var Params : TCreateParams ) ; override ;
  end;

Function  Joker(Zone1, Zone2 : THCpteEdit ; ZoneJoker : TEdit) : Boolean ;
Function  GChercheCompte(FCompte : THCpteEdit ; pp : TPRocZoom ) : Boolean ;
Function  CRenseigneTitreLookup( vFB : TFichierBase ) : String ;
Function  CRenseigneTagLookup( vFB : TFichierBase ) : Integer ;
Function  CRenseigneSQLLookup( FCompte : THCpteEdit ; Var vStTable, vStColonne, vStSelect, vStWhere, vStOrder : String ; vBoSurLib : Boolean = False ) : String ;
Procedure InsereFenetre( FF : TForm ; CT : TWinControl ) ;
{$ENDIF}

function  ExisteCpte( Cpte : String ; fb : TFichierBase ) : Boolean ;
function  LongMax ( fb : TFichierBase ) : integer ;
function  SelectGene : String ;
function  SelectAuxi : String ;
function  SelectSection : String ;
function SelectSectionY2 : String ;
function  SelectBudSect : String ;
function  SelectBudGene : String ;
Function  RecupWhere( tt : TZoomTable) : string ;
Function  Casefic ( Fic : tZoomTable) : TFichierBase ;
{$IFNDEF EAGLSERVER}
Function  AfficheTableOk( tz : TZoomTable ) : Boolean  ;
{$ENDIF}

implementation

{$IFNDEF EAGLSERVER}
{$R *.DFM}
{$ENDIF}

{$IFDEF CCS3}
{$ELSE}
  {$IFNDEF ERADIO}
    {$IFNDEF EAGLSERVER}
      Uses HZoomSP;
    {$ENDIF !EAGLSERVER}
  {$ENDIF !ERADIO}
{$ENDIF}

Const COL_CODE : Integer = 1;
Const COL_LIBELLE : Integer = 2;

Function  tzToCorresp( tz : TZoomTable ) : String ;
BEGIN
Case tz of
   tzCorrespGene1 : Result:='GE1' ;
   tzCorrespGene2 : Result:='GE2' ;
   tzCorrespAuxi1 : Result:='AU1' ;
   tzCorrespAuxi2 : Result:='AU2' ;
   tzCorrespBud1  : Result:='BU1' ;
   tzCorrespBud2  : Result:='BU2' ;
   tzCorrespSec11 : Result:='A11' ;
   tzCorrespSec12 : Result:='A12' ;
   tzCorrespSec21 : Result:='A21' ;
   tzCorrespSec22 : Result:='A22' ;
   tzCorrespSec31 : Result:='A31' ;
   tzCorrespSec32 : Result:='A32' ;
   tzCorrespSec41 : Result:='A41' ;
   tzCorrespSec42 : Result:='A42' ;
   tzCorrespSec51 : Result:='A51' ;
   tzCorrespSec52 : Result:='A52' ;
   END ;
END ;

{$IFNDEF EAGLSERVER}
Function  AfficheTableOk( tz : TZoomTable ) : Boolean  ;
Var tl : TTableLibreCompta ;
    i : Integer ;

BEGIN
Result:=TRUE ;
i:=255 ; tl:=tlGene ;
Case tz of
  tzNatGene0,tzNatGene1,tzNatGene2,tzNatGene3,tzNatGene4,tzNatGene5,
  tzNatGene6,tzNatGene7,tzNatGene8,tzNatGene9 : BEGIN tl:=tlGene ; i:=Byte(tz)-Byte(tzNatGene0) ; END ;
  tzNatTiers0,tzNatTiers1,tzNatTiers2,tzNatTiers3,tzNatTiers4,tzNatTiers5,
  tzNatTiers6,tzNatTiers7,tzNatTiers8,tzNatTiers9 : BEGIN tl:=tlAux ; i:=Byte(tz)-Byte(tzNatTiers0) ; END ;
  tzNatSect0,tzNatSect1,tzNatSect2,tzNatSect3,tzNatSect4,tzNatSect5,
  tzNatSect6,tzNatSect7,tzNatSect8,tzNatSect9 : BEGIN tl:=tlSect ; i:=Byte(tz)-Byte(tzNatSect0) ; END ;
  tzNatBud0,tzNatBud1,tzNatBud2,tzNatBud3,tzNatBud4,tzNatBud5,
  tzNatBud6,tzNatBud7,tzNatBud8,tzNatBud9 : BEGIN tl:=tlCptBud ; i:=Byte(tz)-Byte(tzNatBud0) ; END ;
  tzNatBudS0,tzNatBudS1,tzNatBudS2,tzNatBudS3,tzNatBudS4,tzNatBudS5,
  tzNatBudS6,tzNatBudS7,tzNatBudS8,tzNatBudS9 : BEGIN tl:=tlSectBud ; i:=Byte(tz)-Byte(tzNatBudS0) ; END ;
  tzNatEcrE0,tzNatEcrE1,tzNatEcrE2,tzNatEcrE3 : BEGIN tl:=tlEcrGen ; i:=Byte(tz)-Byte(tzNatEcrE0) ; END ;
  tzNatEcrA0,tzNatEcrA1,tzNatEcrA2,tzNatEcrA3 : BEGIN tl:=tlEcrAna ; i:=Byte(tz)-Byte(tzNatEcrA0) ; END ;
  tzNatEcrU0,tzNatEcrU1,tzNatEcrU2,tzNatEcrU3 : BEGIN tl:=tlEcrBud ; i:=Byte(tz)-Byte(tzNatEcrU0) ; END ;
  tzNatImmo0,tzNatImmo1,tzNatImmo2,tzNatImmo3,tzNatImmo4,tzNatImmo5,
  tzNatImmo6,tzNatImmo7,tzNatImmo8,tzNatImmo9 : BEGIN tl:=tlimmo ; i:=Byte(tz)-Byte(tzNatImmo0) ; END ;
  END ;
If i<>255 Then Result:=VH^.OkTableLibre[tl,i] ;
END ;
{$ENDIF}


{=====================================================================}
function LongMax ( fb : TFichierBase ) : integer ;
Var i : Integer ;
BEGIN
if fb in [fbAxe1..fbAux] then Result:=GetInfoCpta(fb).lg else
Case fb of
  fbJal     : Result:=3 ;
  fbAxe1SP1..fbAxe5SP6 : BEGIN
                         { DONE -olaurent -cdll serveur : a reecrire }
                         i:=((Byte(fb)-Byte(fbAxe1SP1)) Mod 6)+1 ;
                         Result:=GetSousPlanAxe(fbSousPlanTofb(fb),i).Longueur ;
                         END ;
  else Result:=17 ;
  end ;
END ;

{$IFNDEF EAGLSERVER}
{====================== Child Form ==============}
procedure TChildForm.Loaded;
begin
Inherited Loaded ;
Visible:=False ; Position:=poDefault ;
BorderIcons:=[] ; BorderStyle:=bsNone ;
HandleNeeded ; SetBounds(0,0,Width,Height) ;
end;

procedure TChildForm.CreateParams(var Params:TCreateParams);
begin
Inherited CreateParams(Params) ;
with Params do
  BEGIN
  WndParent:=TForm(Owner).Handle ;
  Style:=ws_Child or ws_ClipSiblings ;
  x:=0 ; y:=0 ;
  END ;
end;

Procedure InsereFenetre( FF : TForm ; CT : TWinControl ) ;
BEGIN
FF.Parent:=CT ;
WinProcs.SetParent(FF.Handle,CT.Handle) ;
FF.SetBounds(0,0,CT.Width,CT.Height) ;
FF.Show ;
FF.BringToFront ;
END ;
{$ENDIF}

{=====================================================================}
Function RecupWhere( tt : TZoomTable) : string ;
Var nc : String ;
    i : Integer ;
    StConf : String ;
BEGIN
Result:='' ;
Case tt Of
  {Généraux}
  tzGCollectif  : Result:='AND G_COLLECTIF="X"';
  tzGCollClient : Result:='AND G_COLLECTIF="X" AND G_NATUREGENE="COC"';
  tzGCollFourn  : Result:='AND G_COLLECTIF="X" AND G_NATUREGENE="COF"';
  tzGCollDivers : Result:='AND G_COLLECTIF="X" AND G_NATUREGENE="COD"';
  tzGCollSalarie : Result:='AND G_COLLECTIF="X" AND G_NATUREGENE="COS"';
  tzGNonCollectif : Result:='AND G_COLLECTIF="-"';
  tzGCollToutDebit  : Result:='AND G_COLLECTIF="X" AND (G_NATUREGENE="COC" Or G_NATUREGENE="COD")';
  tzGCollToutCredit : Result:='AND G_COLLECTIF="X" AND (G_NATUREGENE="COF" Or G_NATUREGENE="COD")';
  tzGToutDebit  : Result:='AND ((G_COLLECTIF="X" AND (G_NATUREGENE="COC" Or G_NATUREGENE="COD")) Or (G_NATUREGENE="TID") OR ((G_NATUREGENE="DIV") AND (G_LETTRABLE="X")))'; // Ajout des comptes lettrables divers... FQ 16136 SBO 09/08/2005
  tzGToutCredit : Result:='AND ((G_COLLECTIF="X" AND (G_NATUREGENE="COF" Or G_NATUREGENE="COD")) Or (G_NATUREGENE="TIC") OR ((G_NATUREGENE="DIV") AND (G_LETTRABLE="X")))'; // Ajout des comptes lettrables divers... FQ 16136 SBO 09/08/2005
  tzGTID        : Result:='AND G_NATUREGENE="TID"' ;
  tzGTIC        : Result:='AND G_NATUREGENE="TIC"' ;
  tzGTIDTIC     : Result:='AND (G_NATUREGENE="TID" OR G_NATUREGENE="TIC")' ;
  tzGVentilable : Result:='AND G_VENTILABLE="X"' ;
  tzGVentil1    : Result:='AND G_VENTILABLE="X" AND G_VENTILABLE1="X"' ;
  tzGVentil2    : Result:='AND G_VENTILABLE="X" AND G_VENTILABLE2="X"' ;
  tzGVentil3    : Result:='AND G_VENTILABLE="X" AND G_VENTILABLE3="X"' ;
  tzGVentil4    : Result:='AND G_VENTILABLE="X" AND G_VENTILABLE4="X"' ;
  tzGVentil5    : Result:='AND G_VENTILABLE="X" AND G_VENTILABLE5="X"' ;
  tzGlettrable  : Result:='AND G_LETTRABLE="X"' ;
  tzGPointable  : Result:='AND G_POINTABLE="X"' ;
  tzGBanque     : Result:='AND G_NATUREGENE="BQE"' ;
  tzGCaisse     : Result:='AND G_NATUREGENE="CAI"' ;
  tzGCharge     : Result:='AND G_NATUREGENE="CHA"' ;
  tzGProduit    : Result:='AND G_NATUREGENE="PRO"' ;
  tzGImmo       : Result:='AND G_NATUREGENE="IMO"' ;
  tzGDivers     : Result:='AND G_NATUREGENE="DIV"' ;
  tzGExtra      : Result:='AND G_NATUREGENE="EXT"' ;
  tzGLettColl   : Result:='AND (G_LETTRABLE="X" OR G_COLLECTIF="X")' ;
  tzGBQCaiss    : Result:='AND (G_NATUREGENE="BQE" OR G_NATUREGENE="CAI" OR G_LETTRABLE="X" OR G_COLLECTIF="X")' ;
  tzGBQCaissCli : Result:='AND (((G_NATUREGENE<>"COF" AND G_NATUREGENE<>"COS" AND G_NATUREGENE<>"TIC") AND (G_LETTRABLE="X" OR G_COLLECTIF="X")) OR G_NATUREGENE="BQE" OR G_NATUREGENE="CAI")' ;
  tzGBQCaissFou : Result:='AND (((G_NATUREGENE<>"COC" AND G_NATUREGENE<>"TID") AND (G_LETTRABLE="X" OR G_COLLECTIF="X")) OR G_NATUREGENE="BQE" OR G_NATUREGENE="CAI")' ;
  tzGNLettColl  : Result:='AND (G_LETTRABLE="-" AND G_COLLECTIF="-")' ;
  tzGEncais     : Result:='AND (G_SUIVITRESO="ENC" OR G_SUIVITRESO="MIX") AND (G_COLLECTIF="X" or G_LETTRABLE="X")' ;
  tzGDecais     : Result:='AND (G_SUIVITRESO="DEC" OR G_SUIVITRESO="MIX") AND (G_COLLECTIF="X" or G_LETTRABLE="X")' ;
  tzGBilan      : Result:='AND G_NATUREGENE<>"CHA" AND G_NATUREGENE<>"PRO"' ;
  tzGEffet      : Result:='AND G_EFFET="X"';
  {Tiers}
  tzTclient     : Result:='AND T_NATUREAUXI="CLI"' ;
  tzTfourn      : Result:='AND T_NATUREAUXI="FOU"' ;
  tzTsalarie    : Result:='AND T_NATUREAUXI="SAL"' ;
  tzTdivers     : Result:='AND T_NATUREAUXI="DIV"' ;
  tzTDebiteur   : Result:='AND T_NATUREAUXI="AUD"' ;
  tzTCrediteur  : Result:='AND T_NATUREAUXI="AUC"' ;
  tzTPayeur     : Result:='AND T_ISPAYEUR="X"' ;
  tzTPayeurCLI  : Result:='AND T_ISPAYEUR="X" AND (T_NATUREAUXI="AUD" OR T_NATUREAUXI="CLI" OR T_NATUREAUXI="DIV")' ;
  tzTPayeurFOU  : Result:='AND T_ISPAYEUR="X" AND (T_NATUREAUXI="AUC" OR T_NATUREAUXI="FOU" OR T_NATUREAUXI="DIV")' ;
  tzTToutDebit  : Result:='AND (T_NATUREAUXI="AUD" OR T_NATUREAUXI="CLI" OR T_NATUREAUXI="DIV")' ;
  tzTToutCredit : Result:='AND (T_NATUREAUXI="AUC" OR T_NATUREAUXI="FOU" OR T_NATUREAUXI="DIV" OR T_NATUREAUXI="SAL")' ;
  tzTlettrable  : Result:='AND T_LETTRABLE="X"' ;
  tzTReleve     : Result:='AND T_LETTRABLE="X" AND T_RELEVEFACTURE="X" AND (T_NATUREAUXI="CLI" or T_NATUREAUXI="AUD")' ;
  // Modif GP 23/05/2002 N° 10026
  tzTiers       : Result:='AND T_NATUREAUXI<>"NCP" AND T_NATUREAUXI<>"CON" AND T_NATUREAUXI<>"PRO" AND T_NATUREAUXI<>"SUS"' ;
  {Journaux}
  tzJvente      : Result:='AND J_NATUREJAL="VEN"' ;
  tzJachat      : Result:='AND J_NATUREJAL="ACH"' ;
  tzJbanque     : Result:='AND J_NATUREJAL="BQE"' ;
  tzJcaisse     : Result:='AND J_NATUREJAL="CAI"' ;
  tzJOD         : Result:='AND J_NATUREJAL="OD"' ;
  tzJAN         : Result:='AND J_NATUREJAL="ANO"' ;
  tzJCloture    : Result:='AND J_NATUREJAL="CLO"' ;
  tzJEcartChange: Result:='AND J_NATUREJAL="ECC"' ;
  tzJAna        : Result:='AND J_NATUREJAL="ODA"' ;
  tzJAna1       : Result:='AND J_NATUREJAL="ODA" AND J_AXE="A1"' ;
  tzJAna2       : Result:='AND J_NATUREJAL="ODA" AND J_AXE="A2"' ;
  tzJAna3       : Result:='AND J_NATUREJAL="ODA" AND J_AXE="A3"' ;
  tzJAna4       : Result:='AND J_NATUREJAL="ODA" AND J_AXE="A4"' ;
  tzJAna5       : Result:='AND J_NATUREJAL="ODA" AND J_AXE="A5"' ;
  {Sections}
  tzSection     : begin
                    if VH^.LiaisonY2ViaShare then Result:='AND CSP_AXE="A1"'
                                             else Result:='AND S_AXE="A1"' ;
                  end;
  tzSection2    : begin
                    if VH^.LiaisonY2ViaShare then Result:='AND CSP_AXE="A2"'
                                             else Result:='AND S_AXE="A2"' ;
                  end;
  tzSection3    : begin
                    if VH^.LiaisonY2ViaShare then Result:='AND CSP_AXE="A3"'
                                             else Result:='AND S_AXE="A3"' ;
                  end;
  tzSection4    : begin
                    if VH^.LiaisonY2ViaShare then Result:='AND CSP_AXE="A4"'
                                             else Result:='AND S_AXE="A4"' ;
                  end;
  tzSection5    : begin
                    if VH^.LiaisonY2ViaShare then Result:='AND CSP_AXE="A5"'
                                             else Result:='AND S_AXE="A5"' ;
                  end;
  {Budget}
  tzBudGenAtt   : Result:='AND BG_ATTENTE="X"' ;
  tzBudSecAtt1  : Result:='AND BS_ATTENTE="X" AND BS_AXE="A1"' ;
  tzBudSecAtt2  : Result:='AND BS_ATTENTE="X" AND BS_AXE="A2"' ;
  tzBudSecAtt3  : Result:='AND BS_ATTENTE="X" AND BS_AXE="A3"' ;
  tzBudSecAtt4  : Result:='AND BS_ATTENTE="X" AND BS_AXE="A4"' ;
  tzBudSecAtt5  : Result:='AND BS_ATTENTE="X" AND BS_AXE="A5"' ;
  tzBudSec1     : Result:='AND BS_AXE="A1"' ;
  tzBudSec2     : Result:='AND BS_AXE="A2"' ;
  tzBudSec3     : Result:='AND BS_AXE="A3"' ;
  tzBudSec4     : Result:='AND BS_AXE="A4"' ;
  tzBudSec5     : Result:='AND BS_AXE="A5"' ;
  {Correspondance}
  tzCorrespGene1 : Result:='AND CR_TYPE="GE1"';
  tzCorrespGene2 : Result:='AND CR_TYPE="GE2"';
  tzCorrespAuxi1 : Result:='AND CR_TYPE="AU1"';
  tzCorrespAuxi2 : Result:='AND CR_TYPE="AU2"';
  tzCorrespBud1  : Result:='AND CR_TYPE="BU1"';
  tzCorrespBud2  : Result:='AND CR_TYPE="BU2"';
  tzCorrespSec11 : Result:='AND CR_TYPE="A11"';
  tzCorrespSec12 : Result:='AND CR_TYPE="A12"';
  tzCorrespSec21 : Result:='AND CR_TYPE="A21"';
  tzCorrespSec22 : Result:='AND CR_TYPE="A22"';
  tzCorrespSec31 : Result:='AND CR_TYPE="A31"';
  tzCorrespSec32 : Result:='AND CR_TYPE="A32"';
  tzCorrespSec41 : Result:='AND CR_TYPE="A41"';
  tzCorrespSec42 : Result:='AND CR_TYPE="A42"';
  tzCorrespSec51 : Result:='AND CR_TYPE="A51"';
  tzCorrespSec52 : Result:='AND CR_TYPE="A52"';
  {Nomenclature}
  tzNomenclature : Result:='AND N_NUMLIGNE=1';

  { Naure compte }
  tzNatGene0 : Result:='AND NT_TYPECPTE="G00"' ;
  tzNatGene1 : Result:='AND NT_TYPECPTE="G01"' ;
  tzNatGene2 : Result:='AND NT_TYPECPTE="G02"' ;
  tzNatGene3 : Result:='AND NT_TYPECPTE="G03"' ;
  tzNatGene4 : Result:='AND NT_TYPECPTE="G04"' ;
  tzNatGene5 : Result:='AND NT_TYPECPTE="G05"' ;
  tzNatGene6 : Result:='AND NT_TYPECPTE="G06"' ;
  tzNatGene7 : Result:='AND NT_TYPECPTE="G07"' ;
  tzNatGene8 : Result:='AND NT_TYPECPTE="G08"' ;
  tzNatGene9 : Result:='AND NT_TYPECPTE="G09"' ;
  tzNatTiers0 : Result:='AND NT_TYPECPTE="T00"' ;
  tzNatTiers1 : Result:='AND NT_TYPECPTE="T01"' ;
  tzNatTiers2 : Result:='AND NT_TYPECPTE="T02"' ;
  tzNatTiers3 : Result:='AND NT_TYPECPTE="T03"' ;
  tzNatTiers4 : Result:='AND NT_TYPECPTE="T04"' ;
  tzNatTiers5 : Result:='AND NT_TYPECPTE="T05"' ;
  tzNatTiers6 : Result:='AND NT_TYPECPTE="T06"' ;
  tzNatTiers7 : Result:='AND NT_TYPECPTE="T07"' ;
  tzNatTiers8 : Result:='AND NT_TYPECPTE="T08"' ;
  tzNatTiers9 : Result:='AND NT_TYPECPTE="T09"' ;
  tzNatSect0 : Result:='AND NT_TYPECPTE="S00"' ;
  tzNatSect1 : Result:='AND NT_TYPECPTE="S01"' ;
  tzNatSect2 : Result:='AND NT_TYPECPTE="S02"' ;
  tzNatSect3 : Result:='AND NT_TYPECPTE="S03"' ;
  tzNatSect4 : Result:='AND NT_TYPECPTE="S04"' ;
  tzNatSect5 : Result:='AND NT_TYPECPTE="S05"' ;
  tzNatSect6 : Result:='AND NT_TYPECPTE="S06"' ;
  tzNatSect7 : Result:='AND NT_TYPECPTE="S07"' ;
  tzNatSect8 : Result:='AND NT_TYPECPTE="S08"' ;
  tzNatSect9 : Result:='AND NT_TYPECPTE="S09"' ;
  tzNatBud0 : Result:='AND NT_TYPECPTE="B00"' ;
  tzNatBud1 : Result:='AND NT_TYPECPTE="B01"' ;
  tzNatBud2 : Result:='AND NT_TYPECPTE="B02"' ;
  tzNatBud3 : Result:='AND NT_TYPECPTE="B03"' ;
  tzNatBud4 : Result:='AND NT_TYPECPTE="B04"' ;
  tzNatBud5 : Result:='AND NT_TYPECPTE="B05"' ;
  tzNatBud6 : Result:='AND NT_TYPECPTE="B06"' ;
  tzNatBud7 : Result:='AND NT_TYPECPTE="B07"' ;
  tzNatBud8 : Result:='AND NT_TYPECPTE="B08"' ;
  tzNatBud9 : Result:='AND NT_TYPECPTE="B09"' ;
  tzNatBudS0 : Result:='AND NT_TYPECPTE="D00"' ;
  tzNatBudS1 : Result:='AND NT_TYPECPTE="D01"' ;
  tzNatBudS2 : Result:='AND NT_TYPECPTE="D02"' ;
  tzNatBudS3 : Result:='AND NT_TYPECPTE="D03"' ;
  tzNatBudS4 : Result:='AND NT_TYPECPTE="D04"' ;
  tzNatBudS5 : Result:='AND NT_TYPECPTE="D05"' ;
  tzNatBudS6 : Result:='AND NT_TYPECPTE="D06"' ;
  tzNatBudS7 : Result:='AND NT_TYPECPTE="D07"' ;
  tzNatBudS8 : Result:='AND NT_TYPECPTE="D08"' ;
  tzNatBudS9 : Result:='AND NT_TYPECPTE="D09"' ;
  tzNatEcrE0 : Result:='AND NT_TYPECPTE="E00"' ;
  tzNatEcrE1 : Result:='AND NT_TYPECPTE="E01"' ;
  tzNatEcrE2 : Result:='AND NT_TYPECPTE="E02"' ;
  tzNatEcrE3 : Result:='AND NT_TYPECPTE="E03"' ;
  tzNatEcrA0 : Result:='AND NT_TYPECPTE="A00"' ;
  tzNatEcrA1 : Result:='AND NT_TYPECPTE="A01"' ;
  tzNatEcrA2 : Result:='AND NT_TYPECPTE="A02"' ;
  tzNatEcrA3 : Result:='AND NT_TYPECPTE="A03"' ;
  tzNatEcrU0 : Result:='AND NT_TYPECPTE="U00"' ;
  tzNatEcrU1 : Result:='AND NT_TYPECPTE="U01"' ;
  tzNatEcrU2 : Result:='AND NT_TYPECPTE="U02"' ;
  tzNatEcrU3 : Result:='AND NT_TYPECPTE="U03"' ;
         { DONE -olaurent -cdll serveur :  }
  tzAxe1SP1..tzAxe1SP6 : BEGIN i:=Ord(tt)-Ord(tzAxe1SP1)+1 ; Result:='AND PS_AXE="A1" AND PS_SOUSSECTION="'+GetSousPlanAxe(fbAxe1,i).Code+'"' ; END ;
  tzAxe2SP1..tzAxe2SP6 : BEGIN i:=Ord(tt)-Ord(tzAxe2SP1)+1 ; Result:='AND PS_AXE="A2" AND PS_SOUSSECTION="'+GetSousPlanAxe(fbAxe2,i).Code+'"' ; END ;
  tzAxe3SP1..tzAxe3SP6 : BEGIN i:=Ord(tt)-Ord(tzAxe3SP1)+1 ; Result:='AND PS_AXE="A3" AND PS_SOUSSECTION="'+GetSousPlanAxe(fbAxe3,i).Code+'"' ; END ;
  tzAxe4SP1..tzAxe4SP6 : BEGIN i:=Ord(tt)-Ord(tzAxe4SP1)+1 ; Result:='AND PS_AXE="A4" AND PS_SOUSSECTION="'+GetSousPlanAxe(fbAxe4,i).Code+'"' ; END ;
  tzAxe5SP1..tzAxe5SP6 : BEGIN i:=Ord(tt)-Ord(tzAxe5SP1)+1 ; Result:='AND PS_AXE="A5" AND PS_SOUSSECTION="'+GetSousPlanAxe(fbAxe5,i).Code+'"' ; END ;
  tzRubCPTA  : Result:='AND (RB_NATRUB="CPT")' ;
  tzRubBUDG  : Result:='AND RB_NATRUB="BUD" AND RB_FAMILLES LIKE "CBG%"' ;
  tzRubBUDS  : Result:='AND RB_NATRUB="BUD" AND RB_FAMILLES LIKE "CBS%"' ;
  tzRubBUDSG : Result:='AND RB_NATRUB="BUD" AND RB_FAMILLES LIKE "S/G%"' ;
  tzRubBUDGS : Result:='AND RB_NATRUB="BUD" AND RB_FAMILLES LIKE "G/S%"' ;
  tzNatImmo0 : Result:='AND NT_TYPECPTE="I00"' ;
  tzNatImmo1 : Result:='AND NT_TYPECPTE="I01"' ;
  tzNatImmo2 : Result:='AND NT_TYPECPTE="I02"' ;
  tzNatImmo3 : Result:='AND NT_TYPECPTE="I03"' ;
  tzNatImmo4 : Result:='AND NT_TYPECPTE="I04"' ;
  tzNatImmo5 : Result:='AND NT_TYPECPTE="I05"' ;
  tzNatImmo6 : Result:='AND NT_TYPECPTE="I06"' ;
  tzNatImmo7 : Result:='AND NT_TYPECPTE="I07"' ;
  tzNatImmo8 : Result:='AND NT_TYPECPTE="I08"' ;
  tzNatImmo9 : Result:='AND NT_TYPECPTE="I09"' ;
  end ;

  // INDICATEUR DE FERMETURE DES COMPTES
  // On exclut des liste de recherche les comptes/sections/budget/... qui sont fermés... (SBO 30/08/2004 : FQ 10947)
  // Code réintégré - FQ 16037 : SBO 22/09/2005
  Case tt Of
    // Généraux
    tzGeneral,tzGCollectif,tzGNonCollectif,tzGventilable,tzGlettrable,TzGCollClient,
    tzGCollFourn,tzGCollDivers,tzGCollSalarie,tzGpointable,tzGbanque,tzGcaisse,
    tzGcharge,tzGproduit,tzGimmo,tzGDivers,tzGextra,tzGTID,tzGTIC,tzGTIDTIC,tzGLettColl,
    tzGCollToutDebit,tzGBQCaissCli,tzGBQCaissFou,tzGToutDebit,tzGToutCredit,
    tzGVentil1,tzGVentil2,tzGVentil3,tzGVentil4,tzGVentil5,tzGCollToutCredit,tzGBilan,
    tzGEncais,tzGDecais,tzGNLettColl,tzGEffet :
               begin
               Result := Result + ' AND G_FERME="-"';
               end ;
    // Auxiliaires
    tzTiers,tzTclient,tzTfourn,tzTsalarie,tzTdivers,tzTlettrable,
    tzTDebiteur,tzTCrediteur,tzTToutDebit,tzTToutCredit,tzTReleve,
    tzTPayeur,tzTPayeurCLI,tzTPayeurFOU :
               begin
               Result := Result + ' AND T_FERME="-"';
               end ;
    // Sections
    tzSection, tzSection2, tzSection3, tzSection4, tzSection5 :
               begin
                 if VH^.LiaisonY2ViaShare then
                 begin
                  Result := Result + ' AND CSP_FERME="-"';
                 end else
                 begin
                  Result := Result + ' AND S_FERME="-"';
                 end;
               end ;
    // Journaux
    tzJournal,tzJvente,tzJachat,tzJbanque,tzJcaisse,tzJOD,tzJAN,tzJCloture,
    tzJAna,tzJAna1,tzJAna2,tzJAna3,tzJAna4,tzJAna5 :
               begin
               Result := Result + ' AND J_FERME="-"';
               end ;
    // Comptes budgétaires
    tzBudGen,tzBudGenAtt :
               begin
               Result := Result + ' AND BG_FERME="-"';
               end ;
    // Sections budgétaires
    tzBudSecAtt1, tzBudSecAtt2, tzBudSecAtt3, tzBudSecAtt4, tzBudSecAtt5, tzBudSec1,
    tzBudSec2, tzBudSec3, tzBudSec4, tzBudSec5 :
               begin
               Result := Result + ' AND BS_FERME="-"';
               end ;
    // Budget
    tzBudJal : begin
               Result := Result + ' AND BJ_FERME="-"';
               end ;
  end ;

  // confidentialité
  Case CaseFic(tt) of
     fbGene   : nc:='GENERAUX' ;
     fbAux    : nc:='TIERS' ;
     fbAxe1..fbAxe5 : begin
                        if VH^.LiaisonY2ViaShare then
                        begin
                          nc:='CSECTION' ;
                        end else
                        begin
                          nc:='SECTION' ;
                        end;
                      end;
     fbBudGen : nc:='BUDGENE' ;
     fbBudSec1..fbBudSec5 : nc:='BUDSECT' ;
     else nc:='' ;
     END ;
    StConf := SQLConf(nc) ;
    if ( StConf <> '' )
      then Result := Result + ' AND ' + StConf ;

END ;

{=====================================================================}
Function Casefic ( Fic : tZoomTable) : TFichierBase ;
begin
Case fic Of
  tzGeneral,tzGCollectif,tzGNonCollectif,tzGventilable,tzGlettrable,TzGCollClient,
  tzGCollFourn,tzGCollDivers,tzGCollSalarie,tzGpointable,tzGbanque,tzGcaisse,
  tzGcharge,tzGproduit,tzGimmo,tzGDivers,tzGextra,tzGTID,tzGTIC,tzGTIDTIC,tzGLettColl,
  tzGCollToutDebit,tzGBQCaissCli,tzGBQCaissFou,tzGToutDebit,tzGToutCredit,
  tzGVentil1,tzGVentil2,tzGVentil3,tzGVentil4,tzGVentil5,tzGCollToutCredit,tzGBilan,
  tzGEncais,tzGDecais,tzGNLettColl,tzGEffet : casefic:=fbGene ;

  tzTiers,tzTclient,tzTfourn,tzTsalarie,tzTdivers,tzTlettrable,
  tzTDebiteur,tzTCrediteur,tzTToutDebit,tzTToutCredit,tzTReleve,
  tzTPayeur,tzTPayeurCLI,tzTPayeurFOU : casefic:=fbAux ;

  tzSection      : CaseFic:=fbAxe1 ;
  tzSection2     : CaseFic:=fbAxe2 ;
  tzSection3     : CaseFic:=fbAxe3 ;
  tzSection4     : CaseFic:=fbAxe4 ;
  tzSection5     : CaseFic:=fbAxe5 ;

  tzJournal,tzJvente,tzJachat,tzJbanque,tzJcaisse,tzJOD,tzJAN,tzJCloture,
  tzJAna,tzJAna1,tzJAna2,tzJAna3,tzJAna4,tzJAna5 : CaseFic:=fbJal ;

  tzBudGen,tzBudGenAtt : CaseFic:=fbBudGen ;
  tzBudSec1,tzBudSecAtt1 : CaseFic:=fbBudSec1 ;
  tzBudSec2,tzBudSecAtt2 : CaseFic:=fbBudSec2 ;
  tzBudSec3,tzBudSecAtt3 : CaseFic:=fbBudSec3 ;
  tzBudSec4,tzBudSecAtt4 : CaseFic:=fbBudSec4 ;
  tzBudSec5,tzBudSecAtt5 : CaseFic:=fbBudSec5 ;
  tzBudJal : CaseFic:=fbBudJal ;

  tzImmo   : CaseFic:=fbImmo ;

  tzCorrespGene1,tzCorrespGene2,tzCorrespAuxi1,tzCorrespAuxi2,tzCorrespBud1,tzCorrespBud2,
  tzCorrespSec11,tzCorrespSec12,tzCorrespSec21,tzCorrespSec22,tzCorrespSec31,tzCorrespSec32,
  tzCorrespSec41,tzCorrespSec42,tzCorrespSec51,tzCorrespSec52 : CaseFic:=fbCorresp ;

  tzNatGene0,tzNatGene1,tzNatGene2,tzNatGene3,tzNatGene4,
  tzNatGene5,tzNatGene6,tzNatGene7,tzNatGene8,tzNatGene9,
  tzNatTiers0,tzNatTiers1,tzNatTiers2,tzNatTiers3,tzNatTiers4,
  tzNatTiers5,tzNatTiers6,tzNatTiers7,tzNatTiers8,tzNatTiers9,
  tzNatSect0,tzNatSect1,tzNatSect2,tzNatSect3,tzNatSect4,
  tzNatSect5,tzNatSect6,tzNatSect7,tzNatSect8,tzNatSect9,
  tzNatBud0,tzNatBud1,tzNatBud2,tzNatBud3,tzNatBud4,
  tzNatBud5,tzNatBud6,tzNatBud7,tzNatBud8,tzNatBud9,
  tzNatBudS0,tzNatBudS1,tzNatBudS2,tzNatBudS3,tzNatBudS4,
  tzNatBudS5,tzNatBudS6,tzNatBudS7,tzNatBudS8,tzNatBudS9,
  tzNatEcrE0,tzNatEcrE1,tzNatEcrE2,tzNatEcrE3,
  tzNatEcrA0,tzNatEcrA1,tzNatEcrA2,tzNatEcrA3,
  tzNatEcrU0,tzNatEcrU1,tzNatEcrU2,tzNatEcrU3,
  tzNatImmo0,tzNatImmo1,tzNatImmo2,tzNatImmo3,tzNatImmo4,
  tzNatImmo5,tzNatImmo6,tzNatImmo7,tzNatImmo8,tzNatImmo9  : CaseFic:=fbNatCpt ;
  tzAxe1SP1..tzAxe5SP6 : CaseFic:=TFichierBase(Ord(fbAxe1SP1)+ord(Fic)-Ord(tzAxe1SP1)) ;
  tzRubCPTA,tzRubBUDG,tzRubBUDS,tzRubBUDSG,tzRubBUDGS : CaseFic:=fbRubrique ;
  else casefic:=fbGene ;
  END ;
end ;

{$IFNDEF EAGLSERVER}
{=====================================================================}
PROCEDURE TFzoom.LesNoms ;
begin
FWhere:=RecupWhere(ZoomTable) ;
FWhereSup:='' ;
If (ZoomTable in [tzRubCPTA,tzRubBUDG,tzRubBUDS,tzRubBUDSG,tzRubBUDGS]) And
   (SynPlus<>'') Then
  BEGIN
  If ZoomTable=tzRubCPTA Then FWhereSup:='AND RB_FAMILLES like "%'+SynPlus+'%"'
                         Else FWhereSup:='AND RB_BUDJAL like "%'+SynPlus+'%"' ;
  END
{b FP 29/12/2005: Ajoute le filtre pour les restrictions analytiques}
else if (ZoomTable in [tzSection, tzSection2, tzSection3, tzSection4, tzSection5]) and (SynPlus<>'') then
  begin
  if FWhere <> '' then FWhere := FWhere + ' AND ';
  FWhere := FWhere + SynPlus;
  END ;
{e FP 29/12/2005}
Case Casefic(ZoomTable) Of
  fbGene         : BEGIN FNomCode:='G_GENERAL'      ; FNomLibelle:='G_LIBELLE' ; FNomTable:='GENERAUX' ; END ;
  fbAux          : BEGIN FNomCode:='T_AUXILIAIRE'   ; FNomLibelle:='T_LIBELLE' ; FNomTable:='TIERS'    ; END ;
  fbAxe1..fbAxe5 : BEGIN
                    if VH^.LiaisonY2ViaShare then
                    begin
                      FNomCode:='CSP_SECTION'      ;
                      FNomLibelle:='CSP_LIBELLE' ;
                      FNomTable:='CSECTION'  ;
                    end else
                    begin
                      FNomCode:='S_SECTION'      ;
                      FNomLibelle:='S_LIBELLE' ;
                      FNomTable:='SECTION'  ;
                    end;
                  END ;
  fbJal          : BEGIN FNomCode:='J_JOURNAL'      ; FNomLibelle:='J_LIBELLE' ; FNomTable:='JOURNAL'  ; END ;
  fbBudGen       : BEGIN FNomCode:='BG_BUDGENE'     ; FNomLibelle:='BG_LIBELLE' ; FNomTable:='BUDGENE'   ; END ;
  fbBudSec1..fbBudSec5 : BEGIN FNomCode:='BS_BUDSECT'     ; FNomLibelle:='BS_LIBELLE' ; FNomTable:='BUDSECT'   ; END ;
  fbBudJal       : BEGIN FNomCode:='BJ_BUDJAL'      ; FNomLibelle:='BJ_LIBELLE' ; FNomTable:='BUDJAL'   ; END ;
  fbImmo         : BEGIN FNomCode:='I_IMMO'       ; FNomLibelle:='I_LIBELLE' ; FNomTable:='IMMO'     ; END ;
  fbCorresp      : BEGIN FNomCode:='CR_CORRESP'     ; FNomLibelle:='CR_LIBELLE'; FNomTable:='CORRESP'  ; END ;
  fbNatCpt       : BEGIN FNomCode:='NT_NATURE'      ; FNomLibelle:='NT_LIBELLE'; FNomTable:='NATCPTE'  ; END ;
  fbAxe1SP1..fbAxe5SP6 : BEGIN FNomCode:='PS_CODE'  ; FNomLibelle:='PS_LIBELLE'; FNomTable:='SSSTRUCR'  ; END ;
  fbRubrique     : BEGIN FNomCode:='RB_RUBRIQUE'  ; FNomLibelle:='RB_LIBELLE'; FNomTable:='RUBRIQUE'  ; END ;
  end ;
end ;

{=====================================================================}
procedure TFZoom.APPELFICHE(Sender: TObject);
Var Axe : String3 ;
    fb  : TFichierBase ;
    NoPage : Integer ;
begin
Axe:='' ; fb:=CaseFic(ZoomTable) ;
NoPage:=0 ;
if fb in [fbAxe1..fbAxe5] then Axe:=fbToAxe(fb) else
if fb in [fbBudSec1..fbBudSec5] then Axe:=fbToAxe(fb) ;
if fb=fbCorresp then Axe:=tzToCorresp(ZoomTable) ;
If fb=fbNatCpt Then Axe:=tzToNature(ZoomTable) ;
If fb=fbAux Then NoPage:=1 ;
If (Fliste.Cells[COL_CODE,Fliste.Row]<>'') Then
   BEGIN
   If Assigned(FonctionZoom) And ((ZoomTable in [tzRubCPTA,tzRubBUDG,tzRubBUDS,tzRubBUDSG,tzRubBUDGS])=FALSE) then
      BEGIN
      FonctionZoom(nil,Axe,Fliste.Cells[COL_CODE,Fliste.Row],TaModif,NoPage) ;
      END Else
      BEGIN
      Case fb Of
        fbGene : if Assigned(ProcZoomGene) then ProcZoomGene(Nil,Axe,Fliste.Cells[COL_CODE,Fliste.Row],taModif,0) ;
        fbAux  : if Assigned(ProcZoomTiers) then ProcZoomTiers(Nil,Axe,Fliste.Cells[COL_CODE,Fliste.Row],taModif,1) ;
        fbAxe1..fbAxe5 : if Assigned(ProcZoomSection) then ProcZoomSection(Nil,Axe,Fliste.Cells[COL_CODE,Fliste.Row],taModif,0) ;
        fbJal : if Assigned(ProcZoomJal) then ProcZoomJal(Nil,Axe,Fliste.Cells[COL_CODE,Fliste.Row],taModif,0) ;
        fbBudGen : if Assigned(ProcZoomBudGen) then ProcZoomBudGen(Nil,Axe,Fliste.Cells[COL_CODE,Fliste.Row],taModif,0) ;
        fbBudSec1..fbBudSec5 : if Assigned(ProcZoomBudSec) then ProcZoomBudSec(Nil,Axe,Fliste.Cells[COL_CODE,Fliste.Row],taModif,0) ;
        fbBudJal : if Assigned(ProcZoomBudJal) then ProcZoomBudJal(Nil,Axe,Fliste.Cells[COL_CODE,Fliste.Row],taModif,0) ;
        fbCorresp : if Assigned(ProcZoomCorresp) then ProcZoomCorresp(Nil,Axe,Fliste.Cells[COL_CODE,Fliste.Row],taModif,0) ;
        fbNatCpt : if Assigned(ProcZoomNatCpt) then ProcZoomNatCpt(Nil,Axe,Fliste.Cells[COL_CODE,Fliste.Row],taModif,0) ;
        fbRubrique : If Assigned(ProcZoomRub) Then
                        BEGIN
                        If ZoomTable=tzRubCPTA Then ProcZoomRub(Nil,'',Fliste.Cells[COL_CODE,Fliste.Row],taModif,0)
                                               Else ProcZoomRub(Nil,'',Fliste.Cells[COL_CODE,Fliste.Row],taModif,1) ;
                        END ;
        END ;
      END ;
   END ;
AfficheTitre(fb) ;
Screen.Cursor:=SyncrDefault ;
end;

{=====================================================================}
procedure TFZoom.CREATEFICHE(Sender: TObject);
Var Axe : String3 ;
    fb  : TFichierBase ;
    NoPage : Integer ;
    St : String ;
begin
Axe:='' ; fb:=CaseFic(ZoomTable) ;
NoPage:=0 ;
if fb in [fbAxe1..fbAxe5] then Axe:=fbToAxe(fb) ;
if fb in [fbBudSec1..fbBudSec5] then Axe:=fbToAxe(fb) ;
If fb=fbNatCpt Then Axe:=tzToNature(ZoomTable) ;
if fb=fbCorresp then Axe:=tzToCorresp(ZoomTable) ;
If fb=fbAux Then NoPage:=1 ;
If FListe.rowCount<2 Then St:=SelectCompte.text Else St:='' ;
If Assigned(FonctionZoom) And ((ZoomTable in [tzRubCPTA,tzRubBUDG,tzRubBUDS,tzRubBUDSG,tzRubBUDGS])=FALSE) then
    BEGIN
    FonctionZoom(Nil,Axe,SelectCompte.Text,TaCreatOne,NoPage) ;
    END Else
    BEGIN
    Case fb Of
      fbGene : If Assigned(ProcZoomGene) Then ProcZoomGene(Nil,Axe,St,taCreatOne,0) ;
      fbAux  : If Assigned(ProcZoomTiers) Then ProcZoomTiers(Nil,Axe,St,taCreatOne,1) ;
      fbAxe1..fbAxe5 : If Assigned(ProcZoomSection) Then ProcZoomSection(Nil,Axe,St,taCreatOne,0) ;
      fbJal : If Assigned(ProcZoomJal) Then ProcZoomJal(Nil,Axe,St,taCreatOne,0) ;
      fbBudGen : If Assigned(ProcZoomBudGen) then ProcZoomBudGen(Nil,Axe,St,taCreatOne,0) ;
      fbBudSec1..fbBudSec5 : If Assigned(ProcZoomBudSec) then ProcZoomBudSec(Nil,Axe,St,taCreatOne,0) ;
      fbBudJal : If Assigned(ProcZoomBudJal) then ProcZoomBudJal(Nil,Axe,St,taCreatOne,0) ;
      fbCorresp : if Assigned(ProcZoomCorresp) then ProcZoomCorresp(Nil,Axe,St,taModif,0) ;
      fbnatCpt : If Assigned(ProcZoomNatCpt) Then ProcZoomNatCpt(Nil,Axe,St,taCreatOne,0) ;
      fbRubrique : If Assigned(ProcZoomRub) Then
                      BEGIN
                      If ZoomTable=tzRubCPTA Then ProcZoomRub(Nil,'',St,taCreatOne,0)
                                             Else ProcZoomRub(Nil,'',St,taCreatOne,1) ;
                      END ;
      END ;
    END ;
AfficheTitre(fb) ;
Screen.Cursor:=SyncrDefault ;
end;

{=====================================================================}
Function EstSouligne(lab : TLabel) : Boolean ;
BEGIN
EstSouligne:=fsUnderline in Lab.Font.Style ;
END ;

{=====================================================================}
procedure TFZoom.Souligne(lab : Tlabel ; ok,Croissant : boolean) ;
begin
if ok then
   BEGIN
   Lab.Font.Style:=Lab.Font.Style+[fsUnderline] ;
   if Croissant then Lab.Font.Style:=Lab.Font.Style-[fsBold] else Lab.Font.Style:=Lab.Font.Style+[fsBold] ;
   END else Lab.Font.Style:=Lab.Font.Style-[fsUnderline,fsBold] ;
if ok then
   BEGIN
   If Lab.Name='TCompte' then Souligne(TLibelle,FALSE,False) else Souligne(TCompte,FALSE,False) ;
   END ;
end ;

{=====================================================================}
Procedure TFZoom.SelectCompteChange ( Sender: TObject ) ;
Var StOrder,StCode,StLibelle,stReq : String ;
    Lefb : TFichierBase ;
    LgMax : Integer ;
Begin
	if FDebraye then exit ;
	SelectCompte.ReadOnly:=TRUE ;
	LesNoms ;
	stReq := 'Select ' + FNomCode + ', ' + FNomLibelle + ' FROM ' + FNomTable ;

	Lefb:=Casefic(ZoomTable) ;
  Case Lefb of
    fbJal  : LgMax := 3 ;
    fbGene : LgMax := GetInfoCpta(fbGene).Lg ;
    fbAux  : LgMax := GetInfoCpta(fbAux).Lg ;
    fbAxe1..fbAxe5 : LgMax:=GetInfoCpta(LeFb).Lg ;
    else LgMax:=17 ;
    End ;
  //Simon
{$IFDEF DECLA}
  LgMax:=17 ;
{$ENDIF}

	SelectCompte.MaxLength := LgMax ;
	StCode := SelectCompte.Text ;
  if Length(StCode)<LgMax then StCode := StCode + '%' ;

  if V_PGI.Driver=dbINFORMIX
  	then StLibelle := SelectLib.Text
    else StLibelle := Uppercase(SelectLib.Text) ;
	if Length(StLibelle) < 35 then StLibelle := StLibelle + '%' ;
	if Pos('?',StCode) > 0 Then StCode := TraduitJoker(StCode) ;
	if StCode <> '%' then
    BEGIN
    stReq := stReq + ' Where ' + FNomCode + ' Like "' + StCode + '" ';
    if StLibelle<>'%' then
      if V_PGI.Driver<>dbINFORMIX
      	then stReq := stReq + ' AND UPPER('+FNomLibelle+') Like "' + StLibelle + '" '
        else stReq := stReq + ' AND ' + FNomLibelle + ' Like "' + StLibelle + '" ' ;
    END
  else
    if StLibelle<>'%' then
      if V_PGI.Driver<>dbINFORMIX
      	then stReq := stReq + ' Where UPPER(' + FNomLibelle + ') Like "' + StLibelle + '" '
        else stReq := stReq + ' Where ' + FNomLibelle + ' Like "' + StLibelle + '" '
    else
      if FWhere<>'' then stReq := stReq + ' Where ' + FNomCode + '<>"' + W_W + '" ' ;

  if FWhere<>'' Then stReq := stReq + FWhere ;
	if (ZoomTable in [tzRubCPTA,tzRubBUDG,tzRubBUDS,tzRubBUDSG,tzRubBUDGS])
  	 And (FWhereSup<>'') then stReq := stReq + FWhereSup ;
	stOrder:='' ;
	if Sender is TLabel then
    BEGIN
    Souligne(TLabel(Sender),true,Croissant) ;
    If EstSouligne(TCompte) then stOrder := ' Order by ' + FNomCode
         	                  else stOrder := ' Order by ' + FNomLibelle ;
    if not croissant then stOrder := stOrder + '  Desc' ;
    END
  else
    if (stCode='%') and (stLibelle<>'%')
    	then stOrder := ' Order by ' + FNomLibelle
      else stOrder := ' Order by ' + FNomCode ;

	stReq := stReq + stOrder;

  RemplirGrille(stReq);

  {$IFNDEF EAGLSERVER}
	AfficheTitre(Lefb) ;
  {$ENDIF}
  SelectCompte.ReadOnly:=FALSE ;
	OkBtn.Enabled:=(FListe.rowCount>1) ;
  AppelF.Enabled:=OkBtn.Enabled ;
End ;

Procedure TFZoom.AfficheTitre(Lefb : TFichierBase) ;
Var ThefbAxe : TFichierBase ;
    NumSousPlan : Integer ;
BEGIN
  // Colonne Code
  Case Lefb of
    fbGene         :	FListe.Cells[COL_CODE,0] := HM.Mess[8] ;
    fbAux          :	FListe.Cells[COL_CODE,0] := HM.Mess[9] ;
    fbAxe1..fbAxe5 :	Fliste.Cells[COL_CODE,0] := HM.Mess[10] ;
    fbJal          :	Fliste.Cells[COL_CODE,0] := HM.Mess[11] ;
    fbBudGen       :	Fliste.Cells[COL_CODE,0] := HM.Mess[13] ;
    fbBudSec1..fbBudSec5 : Fliste.Cells[COL_CODE,0] := HM.Mess[22] ;
    fbBudJal       :	Fliste.Cells[COL_CODE,0] := HM.Mess[23] ;
    fbImmo         :	Fliste.Cells[COL_CODE,0] := HM.Mess[14] ;
    fbCorresp      :	Fliste.Cells[COL_CODE,0] := HM.Mess[8] ;
    fbNatCpt       :	Fliste.Cells[COL_CODE,0] := HM.Mess[19] ;
    fbAxe1SP1..fbAxe5SP6 : BEGIN
                           ThefbAxe:=FbSousPlantoFb(Lefb) ;
                           Case ThefbAxe Of
                             fbAxe1 : NumSousPlan:=Ord(Lefb)-Ord(fbAxe1SP1)+1 ;
                             fbAxe2 : NumSousPlan:=Ord(Lefb)-Ord(fbAxe2SP1)+1 ;
                             fbAxe3 : NumSousPlan:=Ord(Lefb)-Ord(fbAxe3SP1)+1 ;
                             fbAxe4 : NumSousPlan:=Ord(Lefb)-Ord(fbAxe4SP1)+1 ;
                             fbAxe5 : NumSousPlan:=Ord(Lefb)-Ord(fbAxe5SP1)+1 ;
                             else NumSousPlan:=Ord(Lefb)-Ord(fbAxe1SP1)+1 ;
                             End ;
                           Fliste.Cells[COL_CODE,0] := HM.Mess[24] ;
                           If VH^.SousPlanAxe[ThefbAxe,NumSousPlan].Lib<>'' Then
                              Fliste.Cells[COL_CODE,0] := VH^.SousPlanAxe[ThefbAxe,NumSousPlan].Lib ;
                           END ;
    fbRubrique     : Fliste.Cells[COL_CODE,0] := HM.Mess[27] ;
    end ;
  // Colonne Libellé
	FListe.Cells[COL_LIBELLE,0] := HM.Mess[15] ;
END ;


{=====================================================================}
Constructor TFZoom.MyCreateNew ( AOwner : TComponent ; Fic : TZoomTable ; StCompte : String ; PP : TProcZoom ; OkLib : boolean ; StRub : String ) ;
Var Thefb,ThefbAxe : TFichierBase ;
    NumSousPlan : Integer ;
BEGIN
inherited Create(AOwner) ;
FonctionZoom:=pp ;
FDebraye:=TRUE ;
ZoomTable:=Fic ; Croissant:=TRUE ; SynPlus:=StRub ; Thefb:=Casefic(ZoomTable) ;
Case Thefb Of
  fbGene         : BEGIN Caption:=HM.Mess[0] ; TCompte.Caption:=HM.Mess[8] ; END ;
  fbAux          : BEGIN Caption:=HM.Mess[1] ; TCompte.Caption:=HM.Mess[9] ; END ;
  fbAxe1..fbAxe5 : BEGIN Caption:=HM.Mess[2] ; TCompte.Caption:=HM.Mess[10] ; END ;
  fbJal          : BEGIN Caption:=HM.Mess[3] ; TCompte.Caption:=HM.Mess[11] ; END ;
  fbBudGen       : BEGIN Caption:=HM.Mess[5] ; TCompte.Caption:=HM.Mess[13] ; END ;
  fbBudSec1..fbBudSec5 : BEGIN Caption:=HM.Mess[20] ; TCompte.Caption:=HM.Mess[22] ; END ;
  fbBudJal       : BEGIN Caption:=HM.Mess[21] ; TCompte.Caption:=HM.Mess[23] ; END ;
  fbImmo         : BEGIN Caption:=HM.Mess[6] ; TCompte.Caption:=HM.Mess[14] ; END ;
  fbCorresp      : BEGIN Caption:=HM.Mess[16]; TCompte.Caption:=HM.Mess[8] ; END ;
  fbNatCpt       : BEGIN Caption:=HM.Mess[7] ; TCompte.Caption:=HM.Mess[19] ; END ;
  fbAxe1SP1..fbAxe5SP6 : BEGIN
                         ThefbAxe:=FbSousPlantoFb(Thefb) ;
                         Case ThefbAxe Of
                           fbAxe1 : NumSousPlan:=Ord(Thefb)-Ord(fbAxe1SP1)+1 ;
                           fbAxe2 : NumSousPlan:=Ord(Thefb)-Ord(fbAxe2SP1)+1 ;
                           fbAxe3 : NumSousPlan:=Ord(Thefb)-Ord(fbAxe3SP1)+1 ;
                           fbAxe4 : NumSousPlan:=Ord(Thefb)-Ord(fbAxe4SP1)+1 ;
                           fbAxe5 : NumSousPlan:=Ord(Thefb)-Ord(fbAxe5SP1)+1 ;
                           else NumSousPlan:=Ord(Thefb)-Ord(fbAxe1SP1)+1 ;
                           End ;
                         Caption:=HM.Mess[25] ; TCompte.Caption:=HM.Mess[24] ;
                         If VH^.SousPlanAxe[ThefbAxe,NumSousPlan].Lib<>'' Then
                            Caption:=HM.Mess[29]+' '+VH^.SousPlanAxe[ThefbAxe,NumSousPlan].Lib ;
                         END ;
  fbRubrique     : BEGIN Caption:=HM.Mess[28] ; TCompte.Caption:=HM.Mess[27] ; END ;
  end ;
if Not OkLib then
   BEGIN
   SelectLib.text:='' ; FDebraye:=FALSE ; SelectCompte.text:=Copy(StCompte,1,17) ;
   END else
   BEGIN
   SelectCompte.text:='' ; FDebraye:=FALSE ;
   if V_PGI.Driver=dbINFORMIX then SelectLib.text:=Copy(StCompte,1,35) else SelectLib.text:=uppercase(Copy(StCompte,1,35)) ;
   END ;
Souligne(TCompte,TRUE,True) ;
END ;
{$ENDIF}

{=====================================================================}
Procedure RechercheExiste ( Cpte : String ; fb : TFichierBase ; Bourre,SurLib : boolean ;
                            Var St,SQL : String ) ;
var QuelChamp : String ;
begin
  // Jamais de bourrage sur recherche Tiers par le libellé
  if ((fb=fbAux) and (SurLib))
    then Bourre := False ;
  St := Trim(Cpte) ;
  // Si avec bourrage, on bourre...
  if Bourre
    then St := Bourreladonc(St,fb) ;
  // Création requête suivant type d'objet recherché
  Case fb of
    fbGene         : SQL:='SELECT G_GENERAL, G_LIBELLE From GENERAUX WHERE G_GENERAL' ;
    fbAux          : begin
                     if SurLib
                      then QuelChamp:='UPPER(T_LIBELLE)'
                      else QuelChamp := 'T_AUXILIAIRE' ;
                     SQL:='SELECT T_AUXILIAIRE, T_LIBELLE From TIERS WHERE ' + QuelChamp ;
                     end ;
    fbAxe1..fbAxe5 : BEGIN
                        if VH^.LiaisonY2ViaShare then
                        begin
                          SQL:='SELECT CSP_SECTION, CSP_LIBELLE From CSECTION WHERE CSP_SECTION' ;
                        end else
                        begin
                          SQL:='SELECT S_SECTION, S_LIBELLE From SECTION WHERE S_SECTION' ;
                        end;
                     END;
    fbJal          : SQL:='SELECT J_JOURNAL, J_LIBELLE From JOURNAL WHERE J_JOURNAL' ;
    fbBudGen       : SQL:='SELECT BG_BUDGENE, BG_LIBELLE From BUDGENE WHERE BG_BUDGENE' ;
    fbBudSec1..fbBudSec5 : SQL:='SELECT BS_BUDSECT, BS_LIBELLE From BUDSECT WHERE BS_BUDSECT' ;
    fbBudJal       : SQL:='SELECT BJ_BUDJAL, BJ_LIBELLE From BUDJAL WHERE BJ_BUDJAL' ;
    fbImmo         : SQL:='SELECT I_IMMO, I_LIBELLE From IMMO WHERE I_IMMO' ;
    fbCorresp      : SQL:='SELECT CR_CORRESP, CR_LIBELLE From CORRESP WHERE CR_CORRESP' ;
    fbNatCpt       : SQL:='SELECT NT_NATURE, NT_LIBELLE From NATCPTE WHERE NT_NATURE' ;
    fbAxe1SP1..fbAxe5SP6 : SQL:='SELECT PS_CODE, PS_LIBELLE From SSSTRUCR WHERE PS_CODE' ;
    fbRubrique     : SQL:='SELECT RB_RUBRIQUE, RB_LIBELLE From RUBRIQUE WHERE RB_RUBRIQUE' ;
  end ; // fin case
  // Gestion spéciale informix...
  if ((fb=fbAux) and (SurLib))
      then St := uppercase(St) ;
  // Gestin bourrage auto de compte recherché
  {JP 19/04/07 : FQ 19920 : pas de LIKE pour les sections budgétaires}
  if Bourre or (fb in [fbBudSec1..fbBudSec5])
    then SQL := SQL + '="' + St + '"'
    else SQL := SQL + ' Like "' + St + '%"' ;
  // touche finale
  SQL := SQL + ' ' ;
END ;

{=====================================================================}
function ExisteCpte( Cpte : String ; fb : TFichierBase ) : Boolean ;
Var St,SQL : String ;
    Q : TQuery ;
begin
RechercheExiste(Cpte,fb,True,FALSE,ST,SQL) ;
Case fb Of
  fbAxe1..fbAxe5 : SQL:=SQL+' AND S_AXE="'+fbToAxe(fb)+'"' ;
  fbBudSec1..fbBudSec5 : SQL:=SQL+' AND BS_AXE="'+fbToAxe(fb)+'"' ;
  fbAxe1SP1..fbAxe5SP6 : SQL:=SQL+' AND PS_AXE="'+fbToAxe(FbSousPlantoFb(fb))+'"' ;
  END ;
Q:=OpenSQL(SQl,TRUE,-1, '', True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
end ;

{=====================================================================}

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 18/02/2003
Modifié le ... :   /  /
Description .. : 18/02/2003 : Ajout de la variante eAGL avec Lookup
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
Function GChercheCompte(FCompte : THCpteEdit ; pp : TPRocZoom ) : Boolean ;
Var
    ili         : integer ;
    LocateCpt   : String ;
    St          : String ;
    St1         : String ;
    StRub       : String ;
    fb          : TFichierBase ;
    lStTitre    : String ;   // Titre
    lStTable    : String ;   // Table
    lStColonne  : String ;   // Colonne de retour
    lStSelect   : String ;   // Select
    lStWhere    : String ;   // Where
    lStOrder    : String ;   // Orber by
    {$IFNDEF ERADIO}
      {$IFNDEF EAGLSERVER}
        lInTag    : Integer ;  // N° de tag de la fonction de creation / modification
        lItLargeur: Integer ;  // Largeur de la colonne 0
      {$ENDIF !EAGLSERVER}
    {$ENDIF !ERADIO}
    lStSQL      : String ;   // Requête
    lStVide     : String ;   // Pour paramètre
    Qzi         : integer ;
    FZoom       : TFZoom ;
begin
  // init
  St1       := FCompte.text ;
  StRub     := '' ;
  Result    := False ;
  ili       := FCompte.ExisteH ;
  LocateCpt :='' ;
  fb        := CaseFic(FCompte.ZoomTable) ;

  // Recherche préliminaire et bourrage si besoin
  if ili = 1 then
    begin
    if not FCompte.OkLocate then
      begin
      Result := TRUE ;
      Exit ;
      end
    else
      begin
      if not FCompte.Bourre then
        begin
        if fb = fbNatCpt
          then St := BourreLaDoncSurLaTable(TzToNature(FCompte.ZoomTable),FCompte.Text)
          else St := Bourreladonc(FCompte.Text,fb) ;
        end
      else St := Fcompte.Text ;
      if length(St1) = length(St)
        then locateCpt := FCompte.text
        else Exit ;
      end ;
    end ;


  // Utilisation du lookup en eAgl ET 2Tiers pour les zooms sur :
  // - les sections
  // -
  if fb in [fbAxe1,fbAxe2,fbAxe3,fbAxe4,fbAxe5,fbGene,fbAux,fbJal] then
    begin
    try
      // ==> Préparation de la recherche :
      // Titre
      {$IFNDEF EAGLSERVER}
      lStTitre  := CRenseigneTitreLookup ( fb ) ;
      {$ENDIF}
      // Tag
      {$IFNDEF ERADIO}
        {$IFNDEF EAGLSERVER}
          lInTag    := CRenseigneTagLookup ( fb ) ;
        {$ENDIF !EAGLSERVER}
      {$ENDIF !ERADIO}
      // SQL
      CRenseigneSQLLookup( FCompte, lStTable, lStColonne, lStSelect, lStWhere, lStOrder, ili = 2 ) ;
      // On recherche si la requête restourne au moins 1 ligne sinon on considère que l'utilisateur n'a rien saisie.
      if (ili = 0) and (FCompte.Text <> '') then
        begin
        lStVide := '' ;
        RechercheExiste( FCompte.Text, fb, False, False, lStVide, lStSQL ) ;
        if not ExisteSQL( lStSQL + lStWhere )
          then FCompte.Text := '' ;
        end ;

      // Cas particulier pour les auxiliaires :
      if (Not GetParamSoc('SO_CPCODEAUXIONLY')) and ( fb = fbAux ) and ( ili = 2 ) and (FCompte.Text <> '') then
        begin
        lStSQL := 'SELECT T_AUXILIAIRE, T_LIBELLE FROM TIERS WHERE ( ( UPPER(T_LIBELLE) LIKE "' + FCompte.text + '%" ) '
                                                                   +    'OR ( T_AUXILIAIRE LIKE "' + FCompte.text + '%" ) ) '
                                                                   +    lStWhere ;
        lStWhere := '' ;
        FCompte.Text := '' ;
        end
      else
        lStSQL := '' ;

      // ==> Recherche sur FCompte
{$IFNDEF ERADIO}
{$IFNDEF EAGLSERVER}
      lItLargeur := 0;
      Result := LookupList( FCompte,          // Contrôle lookup à renseigné
                            lStTitre,         // Titre
                            lStTable,         // Table
                            lStColonne,       // Colonne de retour
                            lStSelect,        // Select
                            lStWhere,         // Where
                            lStOrder,         // Orber by
                            true,             // Affichage des colonnes de titres
                            lInTag,           // N° de tag de la fonction de creation / modification
                            lStSQL,           // Ordre SQL Complet
                            tlDefault,        // Par défaut
                            lItLargeur        // Agrandissement de la largeur pour la recherche sur libelle
                        ) ;
{$ENDIF !EAGLSERVER}
{$ENDIF !ERADIO}
      Except
        PGIInfo('Erreur lors de la mise en place de la recherche', 'Recherche') ;
      end ;
    end
  else
    // Pour les autres zooms, reprise du traitement classique
    begin
    FZoom:=TFZoom.MyCreateNew(Application,FCompte.ZoomTable,FCompte.Text,pp,(ili=2),FCompte.SynPlus) ;
    FZoom.CptLocate:=LocateCpt ;
    try
      Qzi:=FZoom.ShowModal ;
      if ((Qzi=mrOk) and (FZoom.Fliste.Cells[COL_CODE,FZoom.Fliste.Row]<>'')) then
        begin
        FCompte.Text := FZoom.Fliste.Cells[COL_CODE,FZoom.Fliste.Row] ;
        if FCompte.Libelle<>nil
          then FCompte.Libelle.Caption := FZoom.Fliste.Cells[COL_LIBELLE,FZoom.Fliste.Row] ;
        Result:=TRUE ;
        end ;
      finally
        FZoom.Free ;
      end ;
    end ;

END ;

{=====================================================================}
procedure TFZoom.TCompteMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
Croissant:=not (ssCtrl in Shift) ;
end;

{=====================================================================}
procedure TFZoom.FListeDblClick(Sender: TObject);
begin
Modalresult:=mrOK ;
end;

{=====================================================================}
procedure TFZoom.FormResize(Sender: TObject);
Var R,R0 : TRect ;
    i : integer ;
begin
R:=TDrawGrid(FListe).CellRect(1,1) ; R0:=TDrawGrid(FListe).CellRect(0,0) ;
i:=(FListe.Height-(R0.Bottom-R0.Top+2)) mod (R.Bottom-R.Top+1) ;
Height:=Height-i ; Width:=361 ;
end;


{=====================================================================}
function THCpteEdit.ExisteH  : Byte ;
Var St,SQL : String ;
    Q : TQuery ;
    SurLib : Boolean ;
  Label 0 ;
begin
	Result  := 0 ;
  SurLib  := FALSE ;
  if Libelle<>NIL
    then Libelle.Caption:='' ;
	if Text = '' then Exit ;
0:
	RechercheExiste(Text,CaseFic(ZoomTable),Bourre,SurLib,St,SQL) ;
	SQL:=SQL+RecupWhere(ZoomTable) ;
        {b FP 29/12/2005: Ajoute le filtre pour les restrictions analytiques}
        if (ZoomTable in [tzSection, tzSection2, tzSection3, tzSection4, tzSection5]) and (SynPlus<>'') then
          begin
          SQL := SQL + ' AND ' + SynPlus;
          end ;
        {e FP 29/12/2005}
	Q:=OpenSQL(SQL,TRUE,-1, '', True) ;
	If ((Bourre) and (Not SurLib)) Then
		BEGIN
    if Not Q.EOF then
      BEGIN
      Result:=1 ;
      Text:=St ;
      if Libelle<>NIL
        then Libelle.Caption:=Q.Fields[1].AsString ;
      END ;
    END
	else
    BEGIN
    If Not Q.EOF Then
      If Q.RecordCount = 1 then
        BEGIN
        Result  := 1 ;
        Text    := Q.Fields[0].AsString ;
        if Libelle<>NIL
        	then Libelle.Caption:=Q.Fields[1].AsString ;
        END
      else
        BEGIN
        if SurLib then Result:=2
        else If (Not Bourre) And ((CaseFic(ZoomTable) In [fbBudJal])=TRUE) And (St<>Q.Fields[0].AsString) Then
          BEGIN
          Result  := 1 ;
          Text    := Q.Fields[0].AsString ;
          if Libelle<>NIL then Libelle.Caption:=Q.Fields[1].AsString ;
          end
        else If (Not Bourre) And ((CaseFic(ZoomTable) In [fbBudGen])=TRUE) Then BEGIN // 15159
          while not Q.EOF do begin
            if St=Q.Fields[0].AsString then begin
              Result  := 1 ;
              Text    := Q.Fields[0].AsString ;
              if Libelle<>NIL then Libelle.Caption:=Q.Fields[1].AsString;
              break;
            end;
            Q.Next;
          end;
        END;
      END;
  	END ;

	If (Not SurLib) And (CaseFic(ZoomTable) In [fbAux]) And (Result<=0) and (Not GetParamSoc('SO_CPCODEAUXIONLY')) Then
    BEGIN
    SurLib:=TRUE ;
    Ferme(Q) ;
    Goto 0 ;
    END ;

	Ferme(Q) ;

end ;


{=====================================================================}
Procedure THCpteEdit.SetFBourre ( Value : Boolean ) ;
BEGIN
if (CaseFic(FZoomTable)>fbAux) And (CaseFic(FZoomTable)<>fbNatCpt) then FBourre:=FALSE else FBourre:=Value ;
END ;

{=====================================================================}
Procedure THCpteEdit.SetFSynJoker ( Value : Boolean ) ;
BEGIN
if (CaseFic(FZoomTable) in [fbAxe1..fbAxe5])=FALSE then FSynJoker:=FALSE else FSynJoker:=Value ;
END ;


Procedure THCpteEdit.SetZoomTable ( Value : TZoomTable ) ;
BEGIN
FZoomTable:=Value ;
if csDesigning in ComponentState then else
   BEGIN
   MaxLength:=LongMax(Casefic(FZoomTable)) ;
   END ;
END ;

{=====================================================================}
Function THCpteEdit.ZoomSousPlanAna : Boolean ;
{$IFNDEF ERADIO}
  {$IFNDEF EAGLSERVER}
  Var fb : TFichierBase ;
      StS : String ;
  {$ENDIF !EAGLSERVER}
{$ENDIF !ERADIO}
BEGIN
Result := False;
{$IFNDEF CCS3}
  {$IFNDEF ERADIO}
    {$IFNDEF EAGLSERVER}
      fb:=CaseFic(Zoomtable) ;
      If (fb in [fbAxe1..fbAxe5]) And
         VH^.Cpta[fb].Structure and
         // GCO - 29/11/2006 - FQ 19175
         ExisteSQL('SELECT SS_AXE FROM STRUCRSE WHERE SS_AXE = "' + FBToAxe(fb) + '"')
      Then
      BEGIN
      Result:=TRUE ;
      StS:=Self.text ;
      If ChoisirSousPlan(fb,StS,FSynJoker,taModif) Then Self.text:=StS ;
      END ;
    {$ENDIF !EAGLSERVER}
  {$ENDIF !ERADIO}
{$ENDIF}
END ;

{=====================================================================}
procedure THCpteEdit.DblClick ;
BEGIN
if GChercheCompte(Self,Nil) then ;
inherited DblClick ;
END ;

procedure THCpteEdit.keydown(var Key: Word; Shift: TShiftState) ;
Var Vide : boolean ;
    CtrlF5 : Boolean ;
BEGIN
Vide:=(Shift=[]) ; CtrlF5:=(Shift=[ssCtrl]) And (Key=VK_F5) ;
if Vide And (Key=VK_F5) Then
   BEGIN
   Key:=0 ;
   if GChercheCompte(Self,Nil) then ;
   END Else If CtrlF5 Then ZoomSousPlanAna ;
inherited KeyDown(Key,Shift) ;
END ;


{=====================================================================}
procedure THCpteEdit.DoExit ;
BEGIN
inherited DoExit ;
END ;
{$ENDIF}

{=====================================================================}
{***********A.G.L.***********************************************
Auteur  ...... : Piot
Créé le ...... : 21/02/2005
Modifié le ... :   /  /    
Description .. : - LG - 21/02/2005 - gestion du cutoff
Mots clefs ... : 
*****************************************************************}
Procedure TGGeneral.QueryToGeneral  ( G_Query : TQuery ) ;
Var C : Char ;
    St : String ;
BEGIN
// Modif eAGL (Fields remplacer par FindField)
  General         :=G_Query.FindField('G_GENERAL').AsString ;
  Libelle         :=G_Query.FindField('G_LIBELLE').AsString ;
  Collectif       :=G_Query.FindField('G_COLLECTIF').AsString='X' ;
  NatureGene      :=G_Query.FindField('G_NATUREGENE').AsString ;
  Ventilable[1]   :=G_Query.FindField('G_VENTILABLE1').AsString='X' ;
  Ventilable[2]   :=G_Query.FindField('G_VENTILABLE2').AsString='X' ;
  Ventilable[3]   :=G_Query.FindField('G_VENTILABLE3').AsString='X' ;
  Ventilable[4]   :=G_Query.FindField('G_VENTILABLE4').AsString='X' ;
  Ventilable[5]   :=G_Query.FindField('G_VENTILABLE5').AsString='X' ;
  Lettrable       :=G_Query.FindField('G_LETTRABLE').AsString='X' ;
  TotalDebit      :=G_Query.FindField('G_TOTALDEBIT').AsFloat ;
  TotalCredit     :=G_Query.FindField('G_TOTALCREDIT').AsFloat ;
  Budgene         :=G_Query.FindField('G_BUDGENE').AsString ;
  Tva             :=G_Query.FindField('G_TVA').AsString ;
  Tpf             :=G_Query.FindField('G_TPF').AsString ;
  Tva_Encaissement:=G_Query.FindField('G_TVAENCAISSEMENT').AsString ;
  SoumisTPF       :=G_Query.FindField('G_SOUMISTPF').AsString='X' ;
  RegimeTVA       :=G_Query.FindField('G_REGIMETVA').AsString ;
  Moderegle       :=G_Query.FindField('G_MODEREGLE').AsString ;
  Pointable       :=G_Query.FindField('G_POINTABLE').AsString='X' ;
  St              :=G_Query.FindField('G_SENS').AsString ;
  if St<>'' then C:=St[1] else C:='M' ;
  Case C of 'D' : Sens:=1 ; 'C' : Sens:=2 ; else Sens:=3 ; END ;
  Ferme           :=G_Query.FindField('G_FERME').AsString='X' ;
  QQ1             :=G_Query.FindField('G_QUALIFQTE1').AsString ;
  QQ2             :=G_Query.FindField('G_QUALIFQTE2').AsString ;
  TvaSurEncaissement:=G_Query.FindField('G_TVASURENCAISS').AsString='X' ;
  Abrege          :=G_Query.FindField('G_ABREGE').AsString ;
  SuiviTreso      :=G_Query.FindField('G_SUIVITRESO').AsString ;
  Confidentiel    :=G_Query.FindField('G_CONFIDENTIEL').AsString ;
  CodeConso       :=G_Query.FindField('G_CONSO').AsString ;
{$IFNDEF SPEC350}
  {$IFNDEF SPEC302}
  Effet           :=G_Query.FindField('G_EFFET').AsString='X' ;
  {$ENDIF}
{$ENDIF}
END ;

Constructor TGGeneral.Create(StCompte : String ) ;
Var i : Integer ;
    G_Query : TQuery ;
BEGIN
Inherited Create ;
General:='' ; Libelle:='' ; Collectif:=False ; NatureGene:='DIV' ; Budgene:='' ;
for i:=1 to 5 do Ventilable[i]:=FALSE ; Pointable:=FALSE ;
Lettrable:=FALSE ; TotalDebit:=0 ; TotalCredit:=0 ;
Tva:='' ; Tpf:='' ; Tva_Encaissement:='TD' ;
if StCompte='' then exit ;
// Modif eAGL
  if EstTablePartagee( 'GENERAUX' )
    then G_Query := OpenSQL(SelectGene+' FROM GENERAUXMS WHERE G_GENERAL="'+StCompte+'"', True,-1, '', True)
    else G_Query := OpenSQL(SelectGene+' FROM GENERAUX WHERE G_GENERAL="'+StCompte+'"', True,-1, '', True);
  if Not G_Query.Eof Then QueryToGeneral(G_Query);
  HCtrls.Ferme(G_Query);
END ;

{=====================================================================}
Procedure TGTiers.QueryToTiers ( T_Query : TQuery ) ;
BEGIN
// Modif eAGL (Fields remplacer par FindField)
  Auxi            :=T_Query.FindField('T_AUXILIAIRE').AsString ;
  Libelle         :=T_Query.FindField('T_LIBELLE').AsString ;
  Collectif       :=T_Query.FindField('T_COLLECTIF').AsString ;
  NatureAux       :=T_Query.FindField('T_NATUREAUXI').AsString ;
  Lettrable       :=T_Query.FindField('T_LETTRABLE').AsString='X' ;
  TotalDebit      :=T_Query.FindField('T_TOTALDEBIT').AsFloat ;
  TotalCredit     :=T_Query.FindField('T_TOTALCREDIT').AsFloat ;
  ModeRegle       :=T_Query.FindField('T_MODEREGLE').AsString ;
  JourPaiement1   :=T_Query.FindField('T_JOURPAIEMENT1').AsInteger ;
  JourPaiement2   :=T_Query.FindField('T_JOURPAIEMENT2').AsInteger ;
  RegimeTva       :=T_Query.FindField('T_REGIMETVA').AsString ;
  SoumisTPF       :=T_Query.FindField('T_SOUMISTPF').AsString='X' ;
  Tva_Encaissement:=T_Query.FindField('T_TVAENCAISSEMENT').AsString ;
  Abrege          :=T_Query.FindField('T_ABREGE').AsString ;
  ReleveFacture   :=T_Query.FindField('T_RELEVEFACTURE').AsString='X' ;
  Confidentiel    :=T_Query.FindField('T_CONFIDENTIEL').AsString ;
  Ferme           :=T_Query.FindField('T_FERME').AsString='X' ;
  Devise          :=T_Query.FindField('T_DEVISE').AsString ;
  MultiDevise     :=T_Query.FindField('T_MULTIDEVISE').AsString='X' ;
  CodeConso       :=T_Query.FindField('T_CONSO').AsString ;
  IsPayeur        :=T_Query.FindField('T_ISPAYEUR').AsString='X' ;
  AuxiPayeur      :=T_Query.FindField('T_PAYEUR').AsString ;
  DebrayePayeur   :=T_Query.FindField('T_DEBRAYEPAYEUR').AsString='X' ;
  AvoirRbt        :=T_Query.FindField('T_AVOIRRBT').AsString='X' ;
  TL0             :=T_Query.FindField('T_TABLE0').AsString ;
END ;

Constructor TGTiers.Create(StCompte : String) ;
Var T_Query : TQuery ;
BEGIN
Inherited Create ;
Auxi:='' ; Libelle:='' ; Collectif:='' ; Lettrable:=FALSE ; TotalDebit:=0 ; TotalCredit:=0 ;
NatureAux:='DIV' ; ModeRegle:='' ; JourPaiement1:=0 ; JourPaiement2:=0 ; RegimeTva:='FRA' ;
if StCompte='' then exit ;
// Modif eAGL
  if EstTablePartagee( 'TIERS' )
    then T_Query := OpenSQL(SelectAuxi + ' FROM TIERSMS WHERE T_AUXILIAIRE="'+StCompte+'"', True,-1, '', True)
    else T_Query := OpenSQL(SelectAuxi + ' FROM TIERS WHERE T_AUXILIAIRE="' + StCompte+'"', True,-1, '', True);
  If Not T_Query.Eof Then QueryToTiers(T_Query) ;
  HCtrls.Ferme(T_Query);
END ;

{=====================================================================}
Procedure TGSection.QueryToSection ( S_Query : TQuery ) ;
BEGIN
// Modif eAGL (Fields remplacer par FindField)
  Axe						:= S_Query.FindField('S_AXE').AsString ;
  Sect					:= S_Query.FindField('S_SECTION').AsString ;
  Libelle				:= S_Query.FindField('S_LIBELLE').AsString ;
  TotalDebit		:= S_Query.FindField('S_TOTALDEBIT').AsFloat ;
  TotalCredit		:= S_Query.FindField('S_TOTALCREDIT').AsFloat ;
  Ferme					:= S_Query.FindField('S_FERME').AsString='X' ;
  Abrege				:= S_Query.FindField('S_ABREGE').AsString ;
  Confidentiel	:= S_Query.FindField('S_CONFIDENTIEL').AsString ;
  BudSect				:= S_Query.FindField('S_BUDSECT').AsString ;
END ;

Procedure TGSection.QueryY2ToSection ( S_Query : TQuery ) ;
BEGIN
// Modif eAGL (Fields remplacer par FindField)
  Axe						:= S_Query.FindField('CSP_AXE').AsString ;
  Sect					:= S_Query.FindField('CSP_SECTION').AsString ;
  Libelle				:= S_Query.FindField('CSP_LIBELLE').AsString ;
  TotalDebit		:= S_Query.FindField('CSP_TOTALDEBIT').AsFloat ;
  TotalCredit		:= S_Query.FindField('CSP_TOTALCREDIT').AsFloat ;
  Ferme					:= S_Query.FindField('CSP_FERME').AsString='X' ;
  Abrege				:= S_Query.FindField('CSP_ABREGE').AsString ;
  Confidentiel	:= S_Query.FindField('CSP_CONFIDENTIEL').AsString ;
  BudSect				:= S_Query.FindField('CSP_BUDSECT').AsString ;
END ;

Constructor TGSection.Create(StCompte : String ; StAxe : TZoomTable) ;
Var S_Query : TQuery ;
BEGIN
Inherited Create ;
Case StAxe Of
  tzSection  : Axe:='A1' ; tzSection2 : Axe:='A2' ;
  tzSection3 : Axe:='A3' ; tzSection4 : Axe:='A4' ;
  tzSection5 : Axe:='A5' ;
  end ;
Sect:='' ; Libelle:='' ; TotalDebit:=0 ; TotalCredit:=0 ;
if StCompte='' then exit ;
// Modif eAGL
  if VH^.LiaisonY2ViaShare then
  begin
    S_Query := OpenSQL(SelectSectionY2 + ' FROM CSECTION WHERE CSP_SECTION="'+StCompte+'" AND CSP_AXE="'+AXE+'"', True,-1, '', True);
    If Not S_Query.Eof Then QueryY2ToSection(S_Query) ;
    HCtrls.Ferme(S_Query);
  end else
  begin
    if EstTablePartagee( 'SECTION' )
      then S_Query := OpenSQL(SelectSection + ' FROM SECTIONMS WHERE S_SECTION="'+StCompte+'" AND S_AXE="'+AXE+'"', True,-1, '', True)
      else S_Query := OpenSQL(SelectSection + ' FROM SECTION WHERE S_SECTION="'+StCompte+'" AND S_AXE="'+AXE+'"', True,-1, '', True);
    If Not S_Query.Eof Then QueryToSection(S_Query) ;
    HCtrls.Ferme(S_Query);
  end;
END ;

{=====================================================================}
procedure TGBudSect.QueryToBudSect ( BS_Query : TQuery ) ;
BEGIN
// Modif eAGL (Fields remplacer par FindField)
  Axe						:= BS_Query.FindField('BS_AXE').AsString ;
  BudSect				:= BS_Query.FindField('BS_BUDSECT').AsString ;
  Libelle				:= BS_Query.FindField('BS_LIBELLE').AsString ;
  Abrege				:= BS_Query.FindField('BS_ABREGE').AsString ;
  Signe					:= BS_Query.FindField('BS_SIGNE').AsString ;
  Sens					:= BS_Query.FindField('BS_SENS').AsString ;
  Ferme					:= BS_Query.FindField('BS_FERME').AsString='X' ;
  Confidentiel	:= BS_Query.FindField('BS_CONFIDENTIEL').AsString ;
END ;

Constructor TGBudSect.Create(StCompte : String ; StAxe : TZoomTable) ;
Var BS_Query : TQuery ;
BEGIN
Inherited Create ;
Case StAxe Of
  tzBudSec1 : Axe:='A1' ; tzBudSec2 : Axe:='A2' ;
  tzBudSec3 : Axe:='A3' ; tzBudSec4 : Axe:='A4' ;
  tzBudSec5 : Axe:='A5' ;
  end ;
BudSect:='' ; Libelle:='' ;
if StCompte='' then exit ;
// Modif eAGL
	BS_Query:=OpenSQL(SelectBudSect+' FROM BUDSECT WHERE BS_BUDSECT="'+StCompte+'" AND BS_AXE="'+AXE+'"', True,-1, '', True);
	If Not BS_Query.Eof Then QueryToBudSect(BS_Query) ;
	HCtrls.Ferme(BS_Query);
END ;

{=====================================================================}
procedure TGBudGene.QueryToBudGene ( BG_Query : TQuery ) ;
BEGIN
// Modif eAGL (Fields remplacer par FindField)
  BudGene				:= BG_Query.FindField('BG_BUDGENE').AsString ;
  Libelle				:= BG_Query.FindField('BG_LIBELLE').AsString ;
  Abrege				:= BG_Query.FindField('BG_ABREGE').AsString ;
  Signe					:= BG_Query.FindField('BG_SIGNE').AsString ;
  Sens					:= BG_Query.FindField('BG_SENS').AsString ;
  Ferme					:= BG_Query.FindField('BG_FERME').AsString='X' ;
  Confidentiel	:= BG_Query.FindField('BG_CONFIDENTIEL').AsString ;
END ;

Constructor TGBudGene.Create(StCompte : String) ;
Var BG_Query : TQuery ;
BEGIN
Inherited Create ;
BudGene:='' ; Libelle:='' ;
if StCompte='' then exit ;
// Modif eAGL
	BG_Query := OpenSQL(SelectBudGene+' FROM BUDGENE WHERE BG_BUDGENE="'+StCompte+'"', True,-1, '', True);
	If Not BG_Query.Eof Then QueryToBudGene(BG_Query) ;
	HCtrls.Ferme(BG_Query);
END ;

{=====================================================================}
{=====================================================================}
{***********A.G.L.***********************************************
Auteur  ...... : Piot
Créé le ...... : 21/02/2005
Modifié le ... : 25/08/2005
Description .. : - LG - 21/02/2005 - gestion du cutoff
Suite ........ : - LG - 25/08/2005 - plus geere ici
Mots clefs ... : 
*****************************************************************}
function SelectGene : String ;
BEGIN
{$IFDEF SPEC350}
SelectGene:='Select G_GENERAL, G_LIBELLE, G_COLLECTIF, G_NATUREGENE, G_VENTILABLE1, '
           +'G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5, G_LETTRABLE, '
           +'G_TOTALDEBIT, G_TOTALCREDIT, G_BUDGENE, G_TVA, G_TPF, G_TVAENCAISSEMENT , G_SOUMISTPF, G_REGIMETVA , '
           +'G_MODEREGLE, G_POINTABLE, G_SENS, G_FERME , G_QUALIFQTE1, G_QUALIFQTE2, '
           +'G_TVASURENCAISS, G_ABREGE, G_SUIVITRESO, G_CONFIDENTIEL, G_CONSO ' ;
{$ELSE}
 {$IFDEF SPEC302}
 SelectGene:='Select G_GENERAL, G_LIBELLE, G_COLLECTIF, G_NATUREGENE, G_VENTILABLE1, '
            +'G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5, G_LETTRABLE, '
            +'G_TOTALDEBIT, G_TOTALCREDIT, G_BUDGENE, G_TVA, G_TPF, G_TVAENCAISSEMENT , G_SOUMISTPF, G_REGIMETVA , '
            +'G_MODEREGLE, G_POINTABLE, G_SENS, G_FERME , G_QUALIFQTE1, G_QUALIFQTE2, '
            +'G_TVASURENCAISS, G_ABREGE, G_SUIVITRESO, G_CONFIDENTIEL, G_CONSO ' ;
 {$ELSE}
 SelectGene:='Select G_GENERAL, G_LIBELLE, G_COLLECTIF, G_NATUREGENE, G_VENTILABLE1, '
            +'G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5, G_LETTRABLE, '
            +'G_TOTALDEBIT, G_TOTALCREDIT, G_BUDGENE, G_TVA, G_TPF, G_TVAENCAISSEMENT , G_SOUMISTPF, G_REGIMETVA , '
            +'G_MODEREGLE, G_POINTABLE, G_SENS, G_FERME , G_QUALIFQTE1, G_QUALIFQTE2, '
            +'G_TVASURENCAISS, G_ABREGE, G_SUIVITRESO, G_CONFIDENTIEL, G_CONSO, G_EFFET ' ;
 {$ENDIF}
{$ENDIF}
END ;

function SelectAuxi : String ;
BEGIN
SelectAuxi:='Select T_AUXILIAIRE, T_LIBELLE, T_COLLECTIF, T_NATUREAUXI, T_LETTRABLE, '
           +'T_TOTALDEBIT, T_TOTALCREDIT, T_MODEREGLE, T_JOURPAIEMENT1,T_JOURPAIEMENT2, T_REGIMETVA, '
           +'T_SOUMISTPF, T_TVAENCAISSEMENT, T_ABREGE, T_RELEVEFACTURE, T_CONFIDENTIEL, T_FERME, '
           +'T_DEVISE, T_MULTIDEVISE, T_CONSO, T_ISPAYEUR, T_PAYEUR, T_DEBRAYEPAYEUR, T_AVOIRRBT, T_TABLE0 ' ;
END ;

function SelectSection : String ;
BEGIN
SelectSection:='Select S_AXE, S_SECTION, S_LIBELLE, S_TOTALDEBIT, S_TOTALCREDIT, S_FERME , S_ABREGE, S_CONFIDENTIEL, S_BUDSECT ' ;
END ;

function SelectSectionY2 : String ;
BEGIN
  Result :='Select CSP_AXE, CSP_SECTION, CSP_LIBELLE, CSP_TOTALDEBIT, CSP_TOTALCREDIT, CSP_FERME , CSP_ABREGE, CSP_CONFIDENTIEL, CSP_BUDSECT ' ;
END ;

function SelectBudSect : String ;
BEGIN
SelectBudSect:='Select BS_AXE, BS_BUDSECT, BS_LIBELLE, BS_ABREGE, BS_SIGNE, BS_SENS, BS_FERME, BS_CONFIDENTIEL ' ;
END ;

function SelectBudGene : String ;
BEGIN
SelectBudGene:='Select BG_BUDGENE, BG_LIBELLE, BG_ABREGE, BG_SIGNE, BG_SENS, BG_FERME, BG_CONFIDENTIEL ' ;
END ;

{$IFNDEF EAGLSERVER}
{=====================================================================}
Function Joker(Zone1, Zone2 : THCpteEdit ; ZoneJoker : TEdit) : Boolean ;
Var St, St2           : String ;
    i, OuSuisJe, OuSuisJe2 : byte ;
    Trouve      : Boolean ;
BEGIN
Trouve:=False ; OuSuisJe:=0 ; OuSuisJe2:=0 ; St2:='' ;
if (Zone1.Text<>'')and(Zone2.Text<>'') then
   BEGIN
   OuSuisJe:=Zone1.SelStart ; St:=Zone1.Text ;
   OuSuisJe2:=Zone2.SelStart ; St2:=Zone2.Text ;
   END Else
if Zone1.Text<>'' then BEGIN OuSuisJe:=Zone1.SelStart ; St:=Zone1.Text ; END else
if Zone2.Text<>'' then BEGIN OuSuisJe:=Zone2.SelStart ; St:=Zone2.Text ; END else
if ZoneJoker.Text<>'' then BEGIN OuSuisJe:=ZoneJoker.SelStart ; St:=ZoneJoker.Text ; END ;
for i:=1 to Length(St) do
    if St[i] in ['*','?'] then  BEGIN Trouve:=True ; Break ; END ;
If (St2<>'')and Not Trouve then
   BEGIN
   St:=St2 ;
   for i:=1 to Length(St) do
       if St[i] in ['*','?'] then  BEGIN OuSuisJe:=OuSuisJe2 ; Trouve:=True ; Break ; END ;
   END ;
if Trouve Then
   BEGIN
   Zone1.Text:=''           ; Zone2.Text:='' ;
   ZoneJoker.Visible:=True  ; ZoneJoker.Text:=St ;
   if ZoneJoker.CanFocus then ZoneJoker.SetFocus ; ZoneJoker.SelStart:=OuSuisJe ;
   Zone1.Visible:=False     ; Zone2.Visible:=False ;
   End Else
   BEGIN
   ZoneJoker.Visible:=False ; Zone1.Visible:=True ; Zone2.Visible:=True ;
   if ZoneJoker.Visible=False then BEGIN ZoneJoker.Text:='' ;  END ;
   END ;

Result:=Trouve ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 17/10/2002
Modifié le ... :   /  /    
Description .. : Si besoin, se place sur l'enregistrement identifié par 
Suite ........ : CptLocate
Mots clefs ... : 
*****************************************************************}
procedure TFZoom.FormShow(Sender: TObject);
begin
  If (CptLocate<>'') Then
     BEGIN
     SelectCompte.text:='' ;
		 Localiser(CptLocate);
     END ;
end;

procedure TFZoom.HelpBtnClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 17/10/2002
Modifié le ... :   /  /
Description .. : Remplissage de la liste à partir de la requête passée en
Suite ........ : paramètre
Mots clefs ... :
*****************************************************************}
procedure TFZoom.RemplirGrille(stReq : String);
var
	QZoom : TQuery;
  i 		: Integer;
begin
	QZoom := OpenSQL(stReq,True,-1, '', True);
	if QZoom.Eof then
  	begin // pas de résultat, on vide la liste
	  FListe.RowCount := 2;
    FListe.Cells[COL_CODE,1]:='';
    FListe.Cells[COL_LIBELLE,1]:='';
    end
  else
  	begin // Resize de la liste (+1 pour entête)
    FListe.rowCount := QZoom.RecordCount + 1;
    i := 1;
    while not QZoom.Eof do
      begin
      FListe.Cells[COL_CODE,i] := QZoom.FindField(FNomCode).asString;
      FListe.Cells[COL_LIBELLE,i] := QZoom.FindField(FNomLibelle).asString;
      QZoom.Next;
      Inc(i);
      end;
    end;
  HCtrls.Ferme(QZoom);
  FListe.refresh;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 17/10/2002
Modifié le ... :   /  /
Description .. : Trouve un enregistrement dans la liste
Mots clefs ... :
*****************************************************************}
procedure TFZoom.Localiser(stCode: String);
var i : Integer;
begin
	for i := 2 to FListe.rowCount do
  	if FListe.Cells[COL_CODE,i] = stCode then
    	begin
			FListe.row := i;
      Exit;
      end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 17/02/2003
Modifié le ... :   /  /
Description .. : Retourne le titre de la fenêtre de recherche en fonction du
Suite ........ : type de fichier recherché
Mots clefs ... :
*****************************************************************}
Function CRenseigneTitreLookup ( vFB : TFichierBase ) : String ;
var lFbAxe       : TFichierBase ;
    lNumSousPlan : Integer ;
begin
  Result := 'Recherche' ;
  Case vFB of
    fbGene                : Result := 'Recherche d''un compte général' ;
    fbAux                 : Result := 'Recherche d''un compte auxiliaire' ;
    fbAxe1..fbAxe5        : Result := 'Recherche d''une section analytique' ;
    fbJal                 : Result := 'Recherche d''un journal' ;
    fbBudGen              : Result := 'Recherche d''un compte budgétaire' ;
    fbBudSec1..fbBudSec5  : Result := 'Recherche d''une section budgétaire' ;
    fbBudJal              : Result := 'Recherche d''un journal budgétaire' ;
    fbImmo                : Result := 'Recherche d''une immobilisation' ;
    fbCorresp             : Result := 'Recherche d''un compte de correspondance' ;
    fbNatCpt              : Result := 'Recherche d''une valeur' ;
    fbAxe1SP1..fbAxe5SP6  : begin
                            Result    := 'Recherche d''une sous section' ;
                            lFbAxe := FbSousPlantoFb(vFB) ;
                            case lFbAxe Of
                              fbAxe1 : lNumSousPlan := Ord(vFB) - Ord(fbAxe1SP1) + 1 ;
                              fbAxe2 : lNumSousPlan := Ord(vFB) - Ord(fbAxe2SP1) + 1 ;
                              fbAxe3 : lNumSousPlan := Ord(vFB) - Ord(fbAxe3SP1) + 1 ;
                              fbAxe4 : lNumSousPlan := Ord(vFB) - Ord(fbAxe4SP1) + 1 ;
                              fbAxe5 : lNumSousPlan := Ord(vFB) - Ord(fbAxe5SP1) + 1 ;
                              else lNumSousPlan := Ord(vFB) - Ord(fbAxe1SP1) + 1 ;
                              end ;
                            if VH^.SousPlanAxe[lFbAxe, lNumSousPlan].Lib <> ''
                              then Result := 'Recherche de ' + VH^.SousPlanAxe[lFbAxe, lNumSousPlan].Lib ;
                           end ;
    fbRubrique            : Result := 'Recherche d''une rubrique' ;
  end ;
end ;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 17/02/2003
Modifié le ... :   /  /
Description .. : Retourne le n° de Tag utiliser par la compta PGE pour
Suite ........ : l'appel de la fiche en modification / création / consultation...
Mots clefs ... :
*****************************************************************}
Function CRenseigneTagLookup ( vFB : TFichierBase ) : Integer ;
begin
  Case vFB of
    fbGene                : Result := 7112 ;
    fbAux                 : Result := 7145 ;
    fbAxe1                : Result := 71781 ;
    fbAxe2                : Result := 71782 ;
    fbAxe3                : Result := 71783 ;
    fbAxe4                : Result := 71784 ;
    fbAxe5                : Result := 71785 ;
    fbJal                 : Result := 7211 ;
    fbBudGen              : Result := 1 ;
    fbBudSec1..fbBudSec5  : Result := 1 ;
    fbBudJal              : Result := 1 ;
    fbImmo                : Result := 1 ;
    fbCorresp             : Result := 1 ;
    fbNatCpt              : Result := 1 ;
    fbAxe1SP1..fbAxe5SP6  : Result := 1 ;
    fbRubrique            : Result := 1 ;
    else Result := 1 ;
  end ;
end ;


{$IFNDEF EAGLSERVER}
Function CRenseigneSQLLookup( FCompte : THCpteEdit ; Var vStTable, vStColonne, vStSelect, vStWhere, vStOrder : String ; vBoSurLib : Boolean ) : String ;
var
  lFB    : TFichierBase ;
  lZT    : TZoomTable ;
  lStSyn : String ;
begin

  // Conversion TZoomTable --> TFichierBase
  lZT       := FCompte.ZoomTable ;
  lStSyn    := FCompte.SynPlus ;
  lFB       := CaseFic( lZT ) ;

  // Condition
  vStWhere  := RecupWhere(lZT) ;

  // Condition supplémentaires pour les rubriques
  if (lZT in [tzRubCPTA,tzRubBUDG,tzRubBUDS,tzRubBUDSG,tzRubBUDGS]) and (lStSyn<>'') then
    begin
    if vStWhere <> '' then vStWhere := vStWhere + ' AND' ;
    if lZT = tzRubCPTA
      then vStWhere := vStWhere + ' RB_FAMILLES like "%'+lStSyn+'%"'
      else vStWhere := vStWhere + ' RB_BUDJAL like "%'+lStSyn+'%"'  ;
    end
  {b FP 29/12/2005: Ajoute le filtre pour les restrictions analytiques
   JP 03/12/07 : FQ 21963 : ajout des sections tzSection2, tzSection3, tzSection4, tzSection5}
  else if (lZT in [tzSection, tzSection2, tzSection3, tzSection4, tzSection5]) and (lStSyn<>'') then
    begin
    if vStWhere <> '' then vStWhere := vStWhere + ' AND ';
    vStWhere := vStWhere + lStSyn;
    end ;
  {e FP 29/12/2005}
  // Table, Champ retourné, champ affiché, Ordre de tri...
  Case lFB Of
    // Comptes généraux
    fbGene         : begin
                     vStTable   := 'GENERAUX' ;
                     vStColonne := 'G_GENERAL' ;
                     vStSelect  := 'G_LIBELLE' ;
                     vStOrder   := 'G_GENERAL' ;
                     end ;
    // Comptes auxiliaires
    fbAux          : begin
                     vStTable   := 'TIERS' ;
                     vStColonne := 'T_AUXILIAIRE' ;
                     vStSelect  := 'T_LIBELLE' ;
                     vStOrder   := 'T_AUXILIAIRE' ;
                     end ;
    // Axes analytiques
    fbAxe1..fbAxe5 : begin
                      if VH^.LiaisonY2ViaShare then
                      begin
                       vStTable   := 'CSECTION' ;
                       vStColonne := 'CSP_SECTION' ;
                       vStSelect  := 'CSP_LIBELLE' ;
                       vStOrder   := 'CSP_SECTION' ;
                      end else
                      begin
                       vStTable   := 'SECTION' ;
                       vStColonne := 'S_SECTION' ;
                       vStSelect  := 'S_LIBELLE' ;
                       vStOrder   := 'S_SECTION' ;
                      end;
                     end ;
    // Journaux
    fbJal          : begin
                     vStTable   := 'JOURNAL' ;
                     vStColonne := 'J_JOURNAL' ;
                     vStSelect  := 'J_LIBELLE' ;
                     vStOrder   := 'J_JOURNAL' ;
                     end ;
    // Budgets généraux
    fbBudGen       : begin
                     vStTable   := 'BUDGENE' ;
                     vStColonne := 'BG_BUDGENE' ;
                     vStSelect  := 'BG_LIBELLE' ;
                     vStOrder   := 'BG_BUDGENE' ;
                     end ;
    // Budgets analytiques
    fbBudSec1..fbBudSec5 : begin
                           vStTable   := 'BUDSECT' ;
                           vStColonne := 'BS_BUDSECT' ;
                           vStSelect  := 'BS_LIBELLE' ;
                           vStOrder   := 'BS_BUDSECT' ;
                           end ;
    // Budgets journaux
    fbBudJal       : begin
                     vStTable   := 'BUDJAL' ;
                     vStColonne := 'BJ_BUDJAL' ;
                     vStSelect  := 'BJ_LIBELLE' ;
                     vStOrder   := 'BJ_BUDJAL' ;
                     end ;
    // Immobilisations
    fbImmo         : begin
                     vStTable   := 'I_LIBELLE' ;
                     vStColonne := 'I_IMMO' ;
                     vStSelect  := 'I_LIBELLE' ;
                     vStOrder   := 'I_IMMO' ;
                     end ;
    // Comptes de correspondance
    fbCorresp      : begin
                     vStTable   := 'CORRESP' ;
                     vStColonne := 'CR_CORRESP' ;
                     vStSelect  := 'CR_LIBELLE' ;
                     vStOrder   := 'CR_CORRESP' ;
                     end ;
    // Nature des comptes de correspondances
    fbNatCpt       : begin
                     vStTable   := 'NATCPTE' ;
                     vStColonne := 'NT_NATURE' ;
                     vStSelect  := 'NT_LIBELLE' ;
                     vStOrder   := 'NT_NATURE' ;
                     end ;
    // Sous-sections
    fbAxe1SP1..fbAxe5SP6  : begin
                            vStTable   := 'SSSTRUCR' ;
                            vStColonne := 'PS_CODE' ;
                            vStSelect  := 'PS_LIBELLE' ;
                            vStOrder   := 'PS_CODE' ;
                            end ;
    // Rubriques
    fbRubrique     : begin
                     vStTable   := 'RUBRIQUE' ;
                     vStColonne := 'RB_RUBRIQUE' ;
                     vStSelect  := 'RB_LIBELLE' ;
                     vStOrder   := 'RB_RUBRIQUE' ;
                     end ;
    end ;
end ;
{$ENDIF}

end.



