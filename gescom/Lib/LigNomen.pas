unit LigNomen;

//===========================================================================
//  Gestion des nomenclatures dans les lignes de documents
//===========================================================================

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, Grids, HDB, ExtCtrls, ComCtrls, Hctrls, StdCtrls, UTOB,
  SaisUtil, HEnt1, hmsgbox, NomenUtil, CasEmplois, HSysMenu, tarifUtil,
  Menus, ImgList,EntGC,AGLInitGC,Choix,UtilPgi,Aglinit,DicoBTP,FactTvaMilliem,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      MPlayer, DBGrids, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
      {$IFDEF V530}
        EdtDOC,
      {$ELSE}
        EdtRdoc,
      {$ENDIF}
{$ENDIF}
  FactTOB,
  Clipbrd,
  Math, NomenErr, M3FP, UtilArticle, HDimension, HPanel,
  UTOFBTMEMORISATION,FactUtil,ParamSoc, HImgList, factdomaines, FactVariante, UCotraitanceOuv,
  xmldom, XMLIntf, msxmldom, XMLDoc, TntExtCtrls, TntStdCtrls, TntGrids, uEntCommun
  ;

procedure Entree_LigneNomen (TobNomen, TobArticles : TOB ; Emploi : boolean ;
                             Niveau,OrdreCompo : integer; TAction : TActionFiche=taModif);
{$IFDEF BTP}
function Entree_LigneOuv (XP : Tform; TOBaffaire,TobNomen, TobArticles,TOBLigPiece,TOBRepart,TobMetres : TOB ; Emploi : boolean ; Niveau,OrdreCompo : integer;DEV:Rdevise; UniteVente:String;var Valeurs:T_valeurs;TAction:TActionFiche;Enht:boolean;TypeARp:boolean;ModeSaisieAch : boolean; CotraitOk : boolean): boolean;
{$ENDIF}

type
    TFLigNomen = class(TForm)
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BAide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BValider: TToolbarButton97;
    G_NLIG: THGrid;
    PANNOMENENT: TPanel;
    COMPOSE: TEdit;
    TGNE_ARTICLE: TLabel;
    MsgBox: THMsgBox;
    ImageList1: THImageList;
    bNewligne: TToolbarButton97;
    bDelLigne: TToolbarButton97;
    HMTrad: THSystemMenu;
    bNomenc: TToolbarButton97;
    PopupTVNLIG: TPopupMenu;
    PopTVNLIG_T: TMenuItem;
    PopTVNLIG_F: TMenuItem;
    PopupG_NLIG: TPopupMenu;
    PopG_NLIG_A: TMenuItem;
    PopG_NLIG_C: TMenuItem;
    N1: TMenuItem;
    PopG_NLIG_I: TMenuItem;
    PopG_NLIG_S: TMenuItem;
    PopG_NLIG_N: TMenuItem;
    N2: TMenuItem;
    PopTV_NLIG_O: TMenuItem;
    BInfos: TToolbarButton97;
    PANDIM: TPanel;
    N3: TMenuItem;
    PopG_NLIG_D: TMenuItem;
    TCODEDIM: TLabel;
    TCOMPOSE: THLabel;
    Panel1: TPanel;
    TV_NLIG: TTreeView;
    POuvrage: THPanel;
    TQTEDETAIL: THLabel;
    QTEDETAIL: THNumEdit;
    TUV: THLabel;
    PopG_NLIG_L: TMenuItem;
    BTypeArticle: TToolbarButton97;
    LModeIns: THLabel;
    Popup_Selection: TPopupMenu;
    MSelCopier: TMenuItem;
    MSelEnregistrer: TMenuItem;
    MSelColler: TMenuItem;
    ToolWindow972: TToolWindow97;
    Mtransfert: TMemo;
    N4: TMenuItem;
    POPG_NLIG_Coller: TMenuItem;
    Rappeler1: TMenuItem;
    POPG_NLIG_Copier: TMenuItem;
    BArborescence: TToolbarButton97;
    ToolBarDesc: TToolWindow97;
    Pligne: THPanel;
    HLabel1: THLabel;
    HLabel3: THLabel;
    HLabel2: THLabel;
    HLabel4: THLabel;
    TLigPVBase: THLabel;
    TPUTTCLIG: THLabel;
    NEDEV: THLabel;
    TLigPvSaisie: THLabel;
    TLIB_PREST: TLabel;
    TGA_FOURNPRINC: TLabel;
    Label1: TLabel;
    TGA_PRESTATION: TLabel;
    NEPA: THNumEdit;
    NECOEFFG: THNumEdit;
    NEPR: THNumEdit;
    NECOEFPRPV: THNumEdit;
    NEPVHT: THNumEdit;
    NEPVTTC: THNumEdit;
    NEPVFORFAIT: THNumEdit;
    ImTypeArticle: THImageList;
    HLabel5: THLabel;
    NETPS: THNumEdit;
    BInfosLigne: TToolbarButton97;
    Appel_XL: TToolbarButton97;
    Pvalo: THPanel;
    HPanel3: THPanel;
    HLabel8: THLabel;
    MONTANTPA: THNumEdit;
    HLabel9: THLabel;
    MONTANTPR: THNumEdit;
    HLabel10: THLabel;
    MONTANTHT: THNumEdit;
    HLabel11: THLabel;
    MONTANTTTC: THNumEdit;
    HLabel14: THLabel;
    MONTANTFG: THNumEdit;
    HLabel15: THLabel;
    MONTANTMARG: THNumEdit;
    HLabel12: THLabel;
    COEFFG: THNumEdit;
    HLabel13: THLabel;
    COEFMARG: THNumEdit;
    HLabel6: THLabel;
    TOTALTPS: THNumEdit;
    HLabel7: THLabel;
    PValounitaire: THPanel;
    HLabel16: THLabel;
    HLabel17: THLabel;
    HLabel18      : THLabel;
    HLabel19      : THLabel;
    HLabel20      : THLabel;
    HLabel21      : THLabel;
    HLabel22      : THLabel;
    MONTANTPAU    : THNumEdit;
    MONTANTPRU    : THNumEdit;
    MONTANTHTU    : THNumEdit;
    MONTANTTTCU   : THNumEdit;
    MONTANTFGU    : THNumEdit;
    MONTANTMARGU  : THNumEdit;
    HLabel23      : THLabel;
    TPSUNITAIRE   : THNumEdit;
    TDEVHT        : THLabel;
    TDEVTTC       : THLabel;
    TNEMTACH      : THLabel;
    NEMTACH       : THNumEdit;
    TNEMTPR       : THLabel;
    NEMTPR        : THNumEdit;
    HLabel24      : THLabel;
    NECOEFFC      : THNumEdit;
    HLabel25      : THLabel;
    NECOEFFR      : THNumEdit;
    BintegreExcel : TToolbarButton97;
    ODExcelFile   : TOpenDialog;
    VariablesMtrs1: TMenuItem;
    Application1  : TMenuItem;
    Gnrales1      : TMenuItem;
    Document1     : TMenuItem;
    Ligne1        : TMenuItem;

//  evenements sur la form
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
//  boutons
    procedure BValiderClick(Sender: TObject);
    procedure bNewligneClick(Sender: TObject);
    procedure bDelLigneClick(Sender: TObject);
    procedure bNomencClick(Sender: TObject);
//  evenements sur la grid
    procedure G_NLIGRowEnter(Sender: TObject; Ou: Integer;
                             var Cancel: Boolean; Chg: Boolean);
    procedure G_NLIGRowExit(Sender: TObject; Ou: Integer;
                            var Cancel: Boolean; Chg: Boolean);
    procedure G_NLIGCellEnter(Sender: TObject; var ACol, ARow: Integer;
                              var Cancel: Boolean);
    procedure G_NLIGCellExit(Sender: TObject; var ACol, ARow: Integer;
                             var Cancel: Boolean);
    procedure G_NLIGElipsisClick(Sender: TObject);
    procedure G_NLIGDblClick(Sender: TObject);
    procedure G_NLIGRowMoved(Sender: TObject; FromIndex, ToIndex: Integer);
    procedure G_NLIGMouseDown(Sender: TObject; Button: TMouseButton;
                              Shift: TShiftState; X, Y: Integer);
    procedure PopG_NLIG_AClick(Sender: TObject);
    procedure PopG_NLIG_CClick(Sender: TObject);
    procedure PopG_NLIG_IClick(Sender: TObject);
    procedure PopG_NLIG_SClick(Sender: TObject);
    procedure PopG_NLIG_NClick(Sender: TObject);
    procedure PopG_NLIG_DClick(Sender: TObject);
    procedure PopTVNLIG_TClick(Sender: TObject);
    procedure PopTVNLIG_FClick(Sender: TObject);
    procedure PopTV_NLIG_OClick(Sender: TObject);

//  evenements sur le treeview
    procedure TV_NLIGMouseUp(Sender: TObject; Button: TMouseButton;
                             Shift: TShiftState; X, Y: Integer);
    procedure TV_NLIGCollapsing(Sender: TObject; Node: TTreeNode;
                                var AllowCollapse: Boolean);
    procedure TV_NLIGExpanding(Sender: TObject; Node: TTreeNode;
                               var AllowExpansion: Boolean);
    procedure TV_NLIGDblClick(Sender: TObject);
    procedure PopG_NLIG_LClick(Sender: TObject);
    procedure BTypeArticleClick(Sender: TObject);
    procedure MSelCopierClick(Sender: TObject);
    procedure MSelEnregistrerClick(Sender: TObject);
    procedure MSelCollerClick(Sender: TObject);
    procedure POPG_NLIG_CollerClick(Sender: TObject);
    procedure POPG_NLIG_CopierClick(Sender: TObject);
    procedure Rappeler1Click(Sender: TObject);
    procedure G_NLIGFlipSelection(Sender: TObject);
    procedure G_NLIGBeforeFlip(Sender: TObject; ARow: Integer; var Cancel: Boolean);
    procedure BInfosLigneClick(Sender: TObject);
    procedure ToolBarDescClose(Sender: TObject);
    procedure G_NLIGEnter(Sender: TObject);
    procedure Appel_XLClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure QTEDETAILExit(Sender: TObject);
    procedure BintegreExcelClick(Sender: TObject);
    procedure G_NLIGMouseMove(Sender: TObject; Shift: TShiftState; X,Y: Integer);
    procedure VariablesApplicationClick(Sender: TObject);
    procedure VariablesLigneClick(Sender: TObject);
    procedure VariablesGeneraleClick(Sender: TObject);
    procedure VariablesDocumentClick(Sender: TObject);
  private
    { Déclarations privées }
    fGestQteAvanc : boolean;
		RecupPrix : string;
    Xp : Tform;
    fPieceCoTrait : TPieceCotraitOuv;
    //
    MetreDoc : boolean;
    ModeSaisieAch : boolean;
    CoTraitOk : boolean;
    OnValidation : boolean;
    SavAction      : TActionFiche ;
    IsModeSelect  : boolean;
    Cell_Text : string;
    TOBLig        : TOB;
    TobArticles   : TOB;
    TobSauv : TOB;
    iTableLigne : integer;
    InfoNiv       : Integer;
    InfoOrdre     : integer;
    ColsInter : Array of boolean ;
    LesColNLIG : String ;
    ART : THCritMaskEdit;
    LastKey : Word;
    // MODIF Btp
{$IFDEF BTP}
    SavOuvrage  : T_Valeurs;
    IndPou      : Integer;
    ToolBartree : TToolWindow97;
    Reajuste    : boolean;
    TOBREpart   : TOB;
    TOBmetres   : TOB;
{$ENDIF}
    // VARIANTE
    TheVariante : TVariante;
    Ouvrage     : boolean;
    prefixe     : string;
    DEV         : RDevise;
    TOBLigPiece : TOB;
    TOBPiece    : TOB;
    TOBaffaire  : TOB;
    QteDUdetail : double;
    GestionHt   : boolean;
    // Rajour Btp pour eviter effet de bord si appel en cascade.
    SGN_NumLig  : integer;
    SGN_IdComp  : integer;
    SGN_Comp    : integer;
    SGN_Lib     : integer;
    SGN_Qte     : integer;
    SGN_Compose : integer;
    // MODIF BTP
    SGN_Pvht    : Integer;
    SGN_UV      : Integer;
    SGN_TYPA    : Integer;
//
    SGN_FOURN   : integer;
    SGN_TARIFF  : integer;
    SGN_UNAC    : integer;
    SGN_REMF    : integer;
    SGN_PAHT    : integer;
    SGN_PVHTP   : integer;
    SGN_PREST   : integer;
    SGN_TPS     : integer;
    SGN_MO      : integer;
    SGN_MONTANT : integer;
    SGN_MONTANTHTDEV : integer;
    SGN_TYPETRAVAIL : integer;
    SGN_LIBFOU : integer;
    SGN_QTESAIS : integer;
    SGN_PERTE : integer;
    SGN_RENDEMENT : integer;
    SGN_UNITESAIS : integer;
    SGN_CODEMARCHE : integer;
//
{$IFDEF BTP}
    TotalOuv, ValPourcent,VALUOUV,VALUPOURCENT : T_Valeurs;
{$ENDIF}
    TOBImp : TOB;
    // ----------
//  Gestion du grid
    procedure EtudieColsListe ;

//  Gestion des champs
    function  TraiterComposant(ARow : LongInt) : boolean;
    function  TraiterQuantite(ARow : LongInt) : boolean;
    procedure FormateZoneSaisie(Acol,ARow : LongInt) ;
    function  Recherche_Art(ARow : LongInt) : boolean;

//  Gestion des colonnes
//  Gestion des lignes du Grid
//  Attention : les propriétés Objects sont utilisées...
//              Objects[SGN_NumLig, x] contient la Tob ligne de nomenclature
//              Objects[SGN_Comp, x] contient le TreeNode associé à la ligne
//              Objects[SGN_IdComp, x] contient la tob Article

// Traitements sur les lignes
    procedure InsertLigne (ARow : Longint; Ajout : Boolean) ;
    procedure VideLaLigne (ARow : integer);
    procedure InitialiseLigne (ARow : Longint) ;
    procedure SupprimeLigne (ARow : Longint) ;
    procedure Renumerote  ;
    function  ChercheLigne (tn1 : TTreeNode) : integer;

//  Pied de la form
    procedure AffectePied(ARow : Longint);
    procedure AffiGrillCodedimension(TOBArt : TOB);

//  Gestion du Treeview
    procedure SetImages(TV : TTreeView; TN : TTreeNode);
    procedure AfficheTVNLIG;
    function  ChercheNode(TV : TTreeView; TN : TTreeNode; Valeur : integer;
                         Niveau : integer = 1) : TTreeNode;
    procedure Re_AffecteLibelle (TobLND : TOB ; stCol : string ; Qte : double);

//  Création et Chargement des tobs
    procedure ChargeTobArticles;
    procedure CompleteTobArticles(TobLN : TOB);
    procedure RechercheArticle(TobGen : TOB ; stArticle : string);
    procedure AjouteTobLigne(ARow : integer);
    procedure RenseigneTOBLN(TobLN, TobArticle,TOBInfo : TOB);
    procedure RecupInfoPiece ;
    procedure ChargeNomen(CodeNomen : string ; TobLN : TOB ; Niv,OrdreCompo: integer );

//  Traitements de vérif sur les tobs
    procedure GereSousNomen(TobLN : TOB);
    function  VerifNomen(TobArt : TOB ; Arow : integer; Var retour : string; var Libelle : string ) : boolean;
    function  Circularite(CodeCompose : string ; TobLigNomen : TOB; var TOBErreur : TOB) : Boolean;
    function  TobVide(TobLN : TOB) : boolean;
    procedure DepointeSousNomen(tobln: tob);
// Modif BTP
{$IFDEF BTP}
    procedure RenseigneTOBOuv(TOBOuv, TobArticle: TOB; libelle : string = '';AvecPv:boolean = true);
    function ChargeOuvrage(CodeNomen: string; TobOuvrage: TOB; Niv,OrdreCompo: integer; var Tresult: T_valeurs):boolean;
{$ENDIF}
    procedure PostDrawCell(ACol, ARow: Integer; Canvas: TCanvas;
      AState: TGridDrawState);
    procedure GetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas;
      AState: TGridDrawState);
    procedure AfficheSelectionLigne(ou: integer);
    function AutoriseEntreeSelection(ou: Integer;
      var cancel: boolean): boolean;
    procedure ModeNormal;
{$IFDEF BTP}
    procedure ModeSelection;
    procedure TraiterFournisseur(Arow: integer);
    procedure TraiterTemps(ARow : integer);
    procedure TraiterPrestation (Arow : integer);
    procedure TraiterPAFourniture(ARow: integer);
{$ENDIF}
    procedure ExportTOB(TOBLoc: TOB; Mcopy: TstringList);
    function RecupTob(Atob:TOB): integer;
    procedure AjouteLigne (TOBLoc :tob;Arow:integer);
    procedure CollageDonnee;
    procedure positionneValeurPV (TOBL : TOB);
    function IsTypeTOB: boolean;
    function ZoneAccessible(ACol, ARow: Integer): boolean;
    procedure ZoneSuivanteOuOk(var ACol, ARow: Integer;var Cancel: boolean);
    procedure CopieDonnee;
    procedure EnregistreDonne;
    procedure RappellerMemo;
    procedure ChargeDonnee;
    procedure enablePopG_NLIG (ou : integer);
    {$IFDEF BTP}
    procedure AssigneEvenementBtn;
    procedure BArborescenceClick(Sender: TObject);
    procedure definiBarTree;
    procedure ToolBarTreeClose(Sender: Tobject);
    {$ENDIF}
    procedure InitLesCols;
    procedure DefiniColsGrid;
    function  ValideLigne(Ou : Longint) : boolean;
    procedure ZoomArticle;
    function EnregistreOK: boolean;
    procedure AppliqueQteDuDetail(TOBLig: TOB);
    procedure SetControlsEnabled;
    procedure PositionneLig(Ligne: integer);
    procedure SetContraitance (value : boolean);
    function ElipsisMetreClick(TOBL: TOB): Double;
    function AppelMetre(TOBL: TOB; OkClick: boolean=False): Double;
    function TraiteQteSais(Acol, Arow: integer): boolean;

  public
    { Déclarations publiques }
    Action      : TActionFiche ;
    TVCols : string;
    Data_Modified : Boolean;
    TypeArp : boolean;
    TOBInfo : TOB;
{$IFDEF BTP}
		property GS_NUMLIG : integer read SGN_NUMLIG;
    property GS_QTE : integer read SGN_QTE;
    property NiveauDec : integer read InfoNiv;
    property TOBLpiece : TOB read TOBLigPiece;
    property Thefacture : Tform read XP;
    property ActiveCotrait : boolean read coTraitOk write SetContraitance;
    procedure AfficheTotalisation;
    procedure afficheMontantLigne(ou: Integer);
    procedure CalculeOuvrageLoc;
    procedure SetTOBLigne(TOBL, TOBA,TOBInfo: TOB);
    function GetQteduDetail: double;
    function ChargeSousDetail(CodeNomen: string; TobOuvrage: TOB; Niv, OrdreCompo: integer; var Tresult: T_valeurs): boolean;
    procedure RenumeroteLesFilles(TOBLN: TOB; Niveau : integer);
    procedure SetNumerotation(TOBL,TOBMere: TOB; Niveau : integer; Arow: integer);
    procedure AffecteLibelle (TobLN : TOB ; stCol : string);
    procedure Affichegrid;
    procedure recalculOuvrage;
    procedure AfficheLaligne(Arow: integer);
{$ENDIF}
    end;

var
//  FLigNomen : TFLigNomen;
  {Acces,} bAnnule,bValide : boolean;
  CasEmploi_Open : Boolean;

implementation

{$IFDEF BTP}
uses factOuvrage,UtilsMetresXLS,UtilXlsBTP,
//UMetreArticle,
UImportLigneNomen,UCotraitance,BTSAISIEPAGEMETRE_TOF,
UspecifPOC
{$ENDIF}

     ,CbpMCD
     ,CbpEnumerator
;

{$R *.DFM}

function formatDec ( chaine : string; Nb : integer) : string;
var Indice : integer;
begin
result := '';
for Indice := 1 to Nb do
    begin
    result := result + chaine;
    end;
end;

{$IFDEF BTP}
procedure TFLigNomen.CalculeOuvrageLoc;
var Indice : Integer;
    TOBL : TOB;
    Qte,PPQ,QteDuDetail : Double;
    MontPourcent : T_Valeurs;
    ArticleOk : String;
begin
  IndPou := -1;
  InitTableau (ValPourcent);
  InitTableau (ValUPourcent);
  InitTableau (MontPourcent);
  InitTableau (TotalOuv);
  InitTableau (VALUOUV);
  for Indice := 1 to G_NLIG.RowCount - 1 do
  begin
    TOBL := TOB(G_NLIG.objects[SGN_Numlig,Indice]);
    if TOBL = nil then break;
    if TOBL.getvalue('BLO_TYPEARTICLE') = 'POU' then
    begin
      ArticleOk := TOBL.GetValue('BLO_LIBCOMPL');
      break;
    end;
  end;
  for Indice := 1 to G_NLIG.RowCount - 1 do
  begin
    TOBL := TOB(G_NLIG.objects[SGN_Numlig,Indice]);
    if TOBL = nil then break;
    // VARIANTE
    if isVariante (TOBL) then continue;
    // --
    if TOBL.getvalue('BLO_TYPEARTICLE') = 'POU' then
    begin
      Indpou := Indice;
      continue;
    end;
    Qte := TOBL.getvalue('BLO_QTEFACT'); //BRL 21/01 if Qte=0 then Qte := 1;
    PPQ := TOBL.getValue('BLO_PRIXPOURQTE'); if PPQ = 0 then PPQ := 1;
    if Indice = 1 then
    begin
      QteDuDetail := TOBL.getValue('BLO_QTEDUDETAIL'); if QteDuDetail = 0 then QteDuDetail := 1;
    end;
    // --- CALCUL TOTAL OUVRAGE ----- //
    TotalOuv[0] := TotalOuv [0] + Arrondi(TOBL.getvalue('BLO_DPA') * (Qte/PPQ),V_PGI.okdecP);
    TotalOuv[1] := TotalOuv [1] + Arrondi(TOBL.getvalue('BLO_DPR') * (Qte/PPQ),V_PGI.okdecP) ;
    TotalOuv[2] := TotalOuv [2] + Arrondi(TOBL.getvalue('BLO_PUHTDEV') * (Qte/PPQ),V_PGI.okdecP) ;
    TotalOuv[3] := TotalOuv [3] + Arrondi(TOBL.getvalue('BLO_PUTTCDEV') * (Qte/PPQ),V_PGI.okdecP);
    TotalOuv[6] := TotalOuv [6] + Arrondi(TOBL.getvalue('BLO_PMAP') * (Qte / PPQ),V_PGI.OkdecP);
    TotalOuv[7] := TotalOuv [7] + Arrondi(TOBL.getvalue('BLO_PMRP')* (Qte/PPQ),V_PGI.okdecp) ;
    //    if TOBL.getValue('BNP_TYPERESSOURCE')='SAL' then
    if ((Pos(TOBL.GetValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne)>0) and (not IsLigneExternalise(TOBL))) or
    	 ((TOBL.getvalue('BLO_TYPEARTICLE')='OUV') or ((TOBL.getvalue('BLO_TYPEARTICLE')='ARP') and (TOBL.detail.count > 0)))then
    begin
      TotalOuv[9] := TotalOuv [9] + TOBL.getvalue('BLO_TPSUNITAIRE')* (Qte/PPQ) ;
    end;
    TotalOuv[10] := TotalOuv[10] + Arrondi((((PPQ) * TOBL.GetValue ('BLO_MONTANTPAFG'))),4);
    TotalOuv[11] := TotalOuv[11] + Arrondi((((PPQ) * TOBL.GetValue ('BLO_MONTANTPAFR'))),4);
    TotalOuv[12] := TotalOuv[12] + Arrondi((((PPQ) * TOBL.GetValue ('BLO_MONTANTPAFC'))),4);
    TotalOuv[13] := TotalOuv[13] + Arrondi((((PPQ) * TOBL.GetValue ('BLO_MONTANTFG'))),4);
    TotalOuv[14] := TotalOuv[14] + Arrondi((((PPQ) * TOBL.GetValue ('BLO_MONTANTFR'))),4);
    TotalOuv[15] := TotalOuv[15] + Arrondi((((PPQ) * TOBL.GetValue ('BLO_MONTANTFC'))),4);
    TotalOuv[16] := TotalOuv[16] + Arrondi((((PPQ) * TOBL.GetValue ('BLO_MONTANTPA'))),V_PGI.okdecP);
    TotalOuv[17] := TotalOuv[17] + Arrondi((((PPQ) * TOBL.GetValue ('BLO_MONTANTPR'))),V_PGI.okdecP);
    //FV1 : 16/11/2015 - Empêche la division par zéro
    if Qtedudetail = 0 then Qtedudetail := 1;
    //
    VALUOUV[0] := arrondi(TotalOuv [0] / Qtedudetail,V_PGI.okdecP);
    VALUOUV[1] := arrondi(TotalOuv [1] / Qtedudetail,V_PGI.okdecP);
    VALUOUV[2] := arrondi(TotalOuv [2] / Qtedudetail,V_PGI.okdecP);
    VALUOUV[3] := arrondi(TotalOuv [3] / Qtedudetail,V_PGI.okdecP);
    VALUOUV[6] := arrondi(TotalOuv [6] / Qtedudetail,V_PGI.okdecP);
    VALUOUV[7] := arrondi(TotalOuv [7] / Qtedudetail,V_PGI.okdecP);
    if (Pos(TOBL.GetValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne)>0) or
   		 ((TOBL.getvalue('BLO_TYPEARTICLE')='OUV') or ((TOBL.getvalue('BLO_TYPEARTICLE')='ARP') and (TOBL.detail.count > 0)))then
    begin
      VALUOUV[9] := arrondi(TotalOuv [9] / Qtedudetail,V_PGI.okdecP);
    end;
    VALUOUV[10] := arrondi(TotalOuv[10] / Qtedudetail,V_PGI.okdecP);
    VALUOUV[11] := arrondi(TotalOuv[11] / Qtedudetail,V_PGI.okdecP);
    VALUOUV[12] := arrondi(TotalOuv[12] / Qtedudetail,V_PGI.okdecP);
    VALUOUV[13] := arrondi(TotalOuv[13] / Qtedudetail,V_PGI.okdecP);
    VALUOUV[14] := arrondi(TotalOuv[14] / Qtedudetail,V_PGI.okdecP);
    VALUOUV[15] := arrondi(TotalOuv[15] / Qtedudetail,V_PGI.okdecP);
    VALUOUV[16] := arrondi(TotalOuv[16] / Qtedudetail,V_PGI.okdecP);
    VALUOUV[17] := arrondi(TotalOuv[17] / Qtedudetail,V_PGI.okdecP);
    //
    CumuleLesTypes (TOBL,VALUOUV);
    //  --------
    if  not ArticleOKInPOUR (TOBL.GetValue('BLO_TYPEARTICLE'),ArticleOk) then continue;
    MontPourcent[0] := MontPourcent [0] + Arrondi(TOBL.getvalue('BLO_DPA') * (Qte/PPQ),V_PGI.okdecp);
    MontPourcent[1] := MontPourcent [1] + Arrondi(TOBL.getvalue('BLO_DPR') * (Qte/PPQ),V_PGI.okdecp) ;
    MontPourcent[2] := MontPourcent [2] + Arrondi(TOBL.getvalue('BLO_PUHTDEV') * (Qte/PPQ) ,V_PGI.okdecp);
    MontPourcent[3] := MontPourcent [3] + Arrondi(TOBL.getvalue('BLO_PUTTCDEV') * (Qte/PPQ),V_PGI.okdecp);
    MontPourcent[6] := MontPourcent [6] + Arrondi(TOBL.getvalue('BLO_PMAP') * (Qte / PPQ),V_PGI.okdecp);
    MontPourcent[7] := MontPourcent [7] + Arrondi(TOBL.getvalue('BLO_PMRP')* (Qte/PPQ),V_PGI.okdecp);
    //
    //
  end;
  FormatageTableau (TotalOuv,V_PGI.okdecP );
  FormatageTableau (VALUOUV,V_PGI.okdecP );
  FormatageTableau (MontPourcent,V_PGI.okdecP);
  if IndPou >= 0 then
  begin
    TOBL := TOB(G_NLIG.objects[SGN_NumLig,Indpou]);
    if TOBL = nil then exit;
    Qte := TOBL.GetValue ('BLO_QTEFACT'); //BRL 21/01 if Qte=0 then Qte := 1;
    PPQ := TOBL.getValue('BLO_PRIXPOURQTE'); if PPQ = 0 then PPQ := 1;
    TOBL.Putvalue('BLO_DPA',MontPourcent[0]);
    TOBL.Putvalue('BLO_DPR',MontPourcent[1]);
    TOBL.Putvalue('BLO_PUHT',devisetopivotEx(MontPourcent[2],DEV.Taux,DEV.quotite,V_PGI.okdecP));
    TOBL.Putvalue('BLO_PUHTDEV',MontPourcent[2]);
    TOBL.Putvalue('BLO_PUHTBASE',TOBL.getvalue('BLO_PUHT'));
    TOBL.Putvalue('BLO_PUTTCDEV',MontPourcent[3]);
    TOBL.Putvalue('BLO_PUTTC',devisetopivotEx(MontPourcent[3],DEV.Taux,DEV.quotite,V_PGI.okdecP));
    TOBL.Putvalue('BLO_PUTTCBASE',TOBL.getvalue('BLO_PUTTC'));
    TOBL.Putvalue('BLO_PMAP',MontPourcent[6]);
    TOBL.Putvalue('BLO_PMRP',MontPourcent[7]);
    if GestionHt then G_NLIG.Cells [SGN_PvHT,indpou] := TOBL.GetValue('BLO_PUHTDEV')
                 else G_NLIG.Cells [SGN_PvHT,indpou] := TOBL.GetValue('BLO_PUTTCDEV');
    G_NLIG.InvalidateCell(SGN_PvHT,indpou);
    valPourcent[0] := MontPourcent[0] * (Qte/PPQ);
    valPourcent[1] := MontPourcent[1] * (Qte/PPQ) ;
    valPourcent[2] := arrondi (MontPourcent[2] * (Qte/PPQ),V_PGI.OkDecP ) ;
    valPourcent[3] := MontPourcent[3] * (Qte/PPQ);
    valPourcent[6] := MontPourcent[6] * (Qte/PPQ);
    valPourcent[7] := MontPourcent[7] * (Qte/PPQ) ;
    FormatageTableau (ValPourcent,V_PGI.okdecP);
    // -- CALCUL POURCENTAGE EN MODE UNITAIRE
    valUPourcent[0] := MontPourcent[0] * (Qte/PPQ/QteDuDetail);
    valUPourcent[1] := MontPourcent[1] * (Qte/PPQ/QteDuDetail) ;
    valUPourcent[2] := arrondi (MontPourcent[2] * (Qte/PPQ/QteDuDetail),V_PGI.OkDecP ) ;
    valUPourcent[3] := MontPourcent[3] * (Qte/PPQ/QteDuDetail);
    valUPourcent[6] := MontPourcent[6] * (Qte/PPQ/QteDuDetail);
    valUPourcent[7] := MontPourcent[7] * (Qte/PPQ/QteDuDetail) ;
    FormatageTableau (valUPourcent,V_PGI.okdecP);
   	CalculMontantHtDevLigOuv (TOBL,DEV);
  end;
end;
{$ENDIF}

Procedure Entree_LigneNomen (TobNomen, TobArticles : TOB ; Emploi : boolean ;
                             Niveau,OrdreCompo : integer; TAction : TActionFiche=taModif);
var
  FLigNomen : TFLigNomen;
begin
if TobNomen.Detail.Count <= 0 then exit;
CasEmploi_Open := True;
if not Emploi then CasEmploi_Open := False;
FLigNomen := TFLigNomen.Create(Application);
TRY
FLigNomen.Action := TAction;
FLigNomen.TobLig:=TobNomen;
FLigNomen.TobArticles:=TobArticles;
FLigNomen.TobSauv := TOB.Create('LIGNENOMEN',nil,-1); //Sauvegarde des données en cas d'annulation
FLigNomen.TobSauv.Dupliquer(TobNomen,true,true);
FLigNomen.ouvrage := false;
FLigNomen.prefixe := 'GLN';
FLigNomen.InfoNiv := Niveau;
FLigNomen.InfoOrdre := OrdreCompo;
FLigNomen.ShowModal;

if bAnnule then TobNomen.Dupliquer(FLigNomen.TobSauv,true,true);
FligNomen.TobSauv.Free;
finally
FligNomen.free;
end;
end;

{$IFDEF BTP}
function Entree_LigneOuv (XP : Tform; TOBaffaire,TobNomen, TobArticles,TOBLigPiece,TOBRepart,TobMetres: TOB ; Emploi : boolean ; Niveau,OrdreCompo : integer;DEV:Rdevise; UniteVente:String;var Valeurs:T_valeurs;TAction:TActionFiche;Enht:boolean;TypeARp:boolean;ModeSaisieAch : boolean; CotraitOk : boolean): boolean;
var
  FLigNomen : TFLigNomen;
  IItem : TMenuItem;
  Indice : Integer;
begin

  result := false;

  if TobNomen.Detail.Count <= 0 then exit;

  if not Emploi then
    CasEmploi_Open      := False
  else
    CasEmploi_Open      := True;

  FLigNomen             := TFLigNomen.Create(Application);
  FLigNomen.action      := Taction;
  FligNomen.Xp          := Xp; // la fiche facture
  FLigNomen.TobLig      :=TobNomen;
  FLigNomen.TobAffaire  :=TobAffaire;
  FLigNomen.TobArticles :=TobArticles;
  FLigNomen.TOBREpart   :=TobRepart;
  FLigNomen.TOBmetres   := TOBmetres;
  FLigNomen.TobSauv     := TOB.Create('LIGNEOUV',nil,-1); //Sauvegarde des données en cas d'annulation

  FLigNomen.TobSauv.Dupliquer(TobNomen,true,true);

  FLigNomen.ouvrage         := true; // Définition du contexte ouvrage
  FLigNomen.InfoNiv         := Niveau;
  FLigNomen.InfoOrdre       := OrdreCompo;
  FLigNomen.PLigne.Visible  := true;
  FLigNomen.prefixe         := 'BLO';
  FLigNomen.DEV             := DEV;
  FligNomen.TOBLigPiece     := TOBLigPiece;
  FligNomen.GestionHt       := Enht;
  FligNomen.TypeArp         := TypeArp;
  FLigNomen.ModeSaisieAch   := ModeSaisieAch;
  FLigNomen.CoTraitOk       := CoTraitOk;
  //
  for Indice := 0 to FligNomen.PopupG_NLIG.Items.Count-1 do
    if FligNomen.PopupG_NLIG.Items[indice].name = 'PopG_NLIG_L' then FligNomen.PopupG_NLIG.Items[Indice].visible := true;

  if TobNomen.detail.count > 0 then FligNomen.QteDuDetail := TOBNomen.detail[0].getValue('BLO_QTEDUDETAIL')
  else
    FLigNomen.qteduDetail := 1;

  if FligNomen.qteDudetail = 0 then Flignomen.qteDudetail := 1;

//if FligNomen.qtedudetail > 1 then
//begin
  FLigNomen.TUV.Visible  := true;
  FLigNomen.QTEDETAIL.Visible  := true;
  FLigNomen.TQTEDETAIL.Visible  := true;
  FLigNomen.TUV.Caption  := UniteVente;
  FLigNomen.QTEDETAIL.Value  := FligNomen.QTeDudetail;
//end;

  TRY
    FLigNomen.ShowModal;

    if bAnnule then
    begin
      TobNomen.Dupliquer(FLigNomen.TobSauv,true,true);
      Valeurs := FligNomen.savouvrage;
    end
    else
    begin
      (*
      Valeurs := CalculSurTableau('+',Flignomen.TotalOuv,FligNomen.ValPourcent);
      Valeurs := CalculSurTableau('/',Valeurs,FligNomen.QteDudetail);
      *)
      Valeurs := CalculSurTableau('+',Flignomen.VALUOUV,FligNomen.ValUPourcent);
      FormatageTableau (Valeurs,V_PGI.okdecP);

      result := true;
    end;
    FLigNomen.TobSauv.Free;
  FINALLY
    FligNomen.free;
  END;

end;
{$ENDIF}

procedure TFLigNomen.InitLesCols;
begin

  SGN_NumLig      := -1;
  SGN_IdComp      := -1;
  SGN_Comp        := -1;
  SGN_Lib         := -1;
  SGN_Qte         := -1;
  SGN_Compose     := -1;
  SGN_Pvht        := -1;
  SGN_UV          := -1;
  SGN_TYPA        := -1;
  SGN_FOURN       := -1;
  SGN_TARIFF      := -1;
  SGN_UNAC        := -1;
  SGN_REMF        := -1;
  SGN_PAHT        := -1;
  SGN_PVHTP       := -1;
  SGN_PREST       := -1;
  SGN_TPS         := -1;
  SGN_MO          := -1;
  SGN_MONTANT     := -1;
  SGN_MONTANTHTDEV := -1;
  SGN_TYPETRAVAIL := -1;
  SGN_LIBFOU      := -1;
  SGN_QTESAIS     := -1;
  SGN_PERTE       := -1;
  SGN_RENDEMENT   := -1;
  SGN_UNITESAIS   := -1;
  SGN_CODEMARCHE  := -1;
end;

////////////////////////////////////////////////////////////////////////////////
//*************************** Gestion de la form *****************************//
////////////////////////////////////////////////////////////////////////////////


procedure TFLigNomen.FormCreate(Sender: TObject);
begin
OnValidation := false;
MetreDoc := GetParamSoc('SO_BTMETREDOC');
InitLesCols;
IsModeSelect := false;
TypeArp := false;
LmodeIns.Caption := '';
ART := THCritMaskEdit.Create (Self);
G_NLIG.GetCellCanvas  := GetcellCanvas;
G_NLIG.PostDrawCell  := PostDrawCell;
{$IFDEF BTP}
ToolBartree := TToolWindow97.Create (self);
definiBarTree;
AssigneEvenementBtn;
BInfosLigne.visible := true;
BArborescence.visible := true;
Panel1.Visible := false;
{$ENDIF}
// VARIANTE
TheVariante := TVariante.create (self);
fPieceCoTrait := TPieceCotraitOuv.create (self);
end;

procedure TFLigNomen.PostDrawCell(ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState);
var TOBOuv : TOB;
    R:Trect;
    NumGraph : Integer;
{$IFDEF BTP}
    LastBrush,LastPen : TColor;
    TheChaine : string;
{$ENDIF}
begin
  if (Arow < G_NLIG.FixedRows) or (Acol < G_NLIG.FixedCols) then exit;
  TOBOUv := TOB(G_NLIG.Objects[SGN_Numlig, ARow]);
  if TOBOuv = nil then exit;
  R := G_NLIG.CellRect (Acol,Arow);
  if (Acol = SGN_LIBFOU) then
  begin
    if (TOBOUv <> nil) then
    begin
    	Canvas.FillRect(R);
      G_NLIG.Canvas.Brush.Style := bsSolid;
      if TOBOUV.getValue('BLO_NATURETRAVAIL')<>'' then
      	G_NLIG.Canvas.TextOut (R.left,R.Top +2,TOBOUv.GetValue('LIBELLEFOU'));
    end;
  end;
  if (ACOl = SGN_TypA) then
  begin
    Canvas.FillRect (R);
    NumGraph := RecupTypeGraph (TOBouv);
    if NumGraph >= 0 then
    begin
      ImTypeArticle.DrawingStyle := dsTransparent;
      ImTypeArticle.Draw (CanVas,R.left,R.top,NumGraph);
    end;
  end;
{$IFDEF BTP}
  if (Acol=SGN_Qte) then
  begin
    if (TheMetreShare <> nil) And (TheMetreShare.OkExcel) then
    begin
      if TheMetreShare.MetreAttached(TobOuv, TOBLigPiece.getValue('GL_NUMORDRE')) then //TOBLigPiece
      begin
        TheChaine := trim(strs( TOBOuv.GetValue('BLO_QTEFACT'),V_PGI.okDecQ))  ;
        LastBrush := G_NLIG.Canvas.Brush.Color;
        LastPen := G_NLIG.Canvas.Pen.Color;
        G_NLIG.Canvas.FillRect (R);
        G_NLIG.Canvas.Brush.Style:=bsSolid;
        G_NLIG.Canvas.Brush.Color := clRed;
        G_NLIG.Canvas.Pen.Color := clRed;
        G_NLIG.Canvas.Polygon([Point(R.left,R.top), point(R.left+4,R.top), Point(R.left, R.top+4)]);
        G_NLIG.Canvas.Brush.Color := LastBrush;
        G_NLIG.Canvas.Pen.Color := LastPen;
        G_NLIG.Canvas.TextOut(R.Right - canvas.TextWidth(TheChaine)- 3,R.Top+2,TheChaine);
      end;
    end;
  end;
{$ENDIF}
end;

Procedure TFLigNomen.GetCellCanvas ( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) ;
{$IFDEF BTP}
var IndPv : integer;
    Valeurpv : double;
    TOBL : TOB;
{$ENDIF}
BEGIN
{$IFDEF BTP}
  if (Arow < G_NLIG.FixedRows) or (Acol < G_NLIG.FixedCols) then exit;

  if (ACol = SGN_PREST) or (Acol = SGN_TPS) or (ACOL = SGN_MO) then
  begin
  	if Action<> taConsult then
    begin
    	canvas.brush.color := clInfoBk;
    end;
  end;

  if (G_NLIG.objects [SGN_NumLig,Arow]= nil)  then exit;

  TOBL := TOB(G_NLIG.objects [SGN_NumLig,Arow]);

  indpv:=0;
  if (GestionHt) then IndPv := TOB(G_NLIG.objects[SGN_NUMLIG,ARow]).GetNumChamp ('BLO_PUHTBASE')
  else IndPv := TOB(G_NLIG.objects[SGN_NUMLIG,ARow]).GetNumChamp ('BLO_PUTTCBASE');

  if DEV.Code <> V_PGI.DevisePivot then
  begin
    valeurpv := pivottodevise (TOB(G_NLIG.objects[SGN_NUMLIG,ARow]).getvaleur(indpv),DEV.taux,DEV.quotite,V_PGI.okdecP);
  end else
  begin
    valeurpv := arrondi(valeur(TOB(G_NLIG.objects[SGN_NUMLIG,ARow]).getvaleur(indpv)),V_PGI.OKdecP);
  end;

  if (ACol = SGN_PVHT) and
     (TOBL.getValue('BLO_TYPEARTICLE') = 'POU') then
  begin
  	if Action <> TaConsult then
    begin
    	Canvas.Font.Color:=clBlue;
    end;
    canvas.Font.Style := Canvas.Font.Style+[fsBold]
  end;

  if (ACol = SGN_PVHT) and
     (TOBL.getValue('BLO_TYPEARTICLE') <> 'POU') and
     (arrondi(valeur(G_NLIG.cells [SGN_PVht,Arow]),V_PGI.OkDecP) <> valeurpv) then
  begin
  	if Action <> TaConsult then
    begin
      Canvas.Font.Color:=clMaroon;
    end;
    canvas.Font.Style := Canvas.Font.Style+[fsBold]
  end;

  // VARIANTE
  if (isVariante(TOBL)) then
  begin
    Canvas.Font.Style := Canvas.Font.Style + [fsbold,fsItalic];
    if Action <> taConsult then
    begin
    	Canvas.Font.Color := clMaroon;
    end;
  end;

{$ENDIF}

END;

procedure TFLigNomen.DefiniColsGrid;
begin
{$IFDEF BTP}
  G_NLIG.ColFormats [SGN_QTe] := '# ### ##0.'+ formatDec ('0',V_PGI.okdecQ);
  if SGN_QTESAIS<> -1 then G_NLIG.ColFormats [SGN_QTeSAIS] := '# ### ##0.'+ formatDec ('0',V_PGI.okdecQ);

  if (SGN_PAHT    <> -1)  then G_NLIG.colformats [SGN_PAHT]     := '# ### ##0.'+ formatDec ('0',V_PGI.okdecP);
  if (SGN_PVHT    <> -1)  then G_NLIG.colformats [SGN_PVHT]     := '# ### ##0.'+ formatDec ('0',V_PGI.okdecP);
  if (SGN_TPS     <> -1)  then G_NLIG.colformats [SGN_TPS]      := '# ### ##0.'+ formatDec ('0',V_PGI.okdecQ);
  if (SGN_TARIFF  <> -1)  then G_NLIG.colformats [SGN_TARIFF]   := '# ### ##0.'+ formatDec ('0',V_PGI.okdecP);
  if (SGN_MO      <> -1)  then G_NLIG.colformats [SGN_MO]       := '# ### ##0.'+ formatDec ('0',V_PGI.OkDecP );
  if (SGN_MONTANT <> -1)  then G_NLIG.colformats [SGN_MONTANT]  := '# ### ##0.'+ formatDec ('0',V_PGI.OkDecP );
  if (SGN_MONTANTHTDEV <> -1)  then G_NLIG.colformats [SGN_MONTANTHTDEV]  := '# ### ##0.'+ formatDec ('0',V_PGI.OkDecP );
  if (SGN_PVHTP   <> -1)  then G_NLIG.colformats [SGN_PVHTP]    := '# ### ##0.'+ formatDec ('0',V_PGI.OkDecP );
{$ENDIF}
// --
(*
G_NLIG.ColAligns[SGN_NumLig] := taCenter;
G_NLIG.ColWidths[G_NLIG.ColCount - 1] := 0;
G_NLIG.ColLengths[G_NLIG.ColCount - 1] := -1;
G_NLIG.ColWidths[G_NLIG.ColCount - 2] := 0;
G_NLIG.ColLengths[G_NLIG.ColCount - 2] := -1;
*)
  G_NLIG.ColAligns[SGN_NumLig] := taCenter;

  
  if SGN_IDCOMP >= 0 then
  begin
    G_NLIG.ColWidths[SGN_IDCOMP] := 0;
    G_NLIG.ColLengths[SGN_IDCOMP] := -1;
  end;

  if SGN_COMPOSE >= 0 then
  begin
    G_NLIG.ColWidths[SGN_COMPOSE] := 0;
    G_NLIG.ColLengths[SGN_COMPOSE] := -1;
  end;

  if SGN_TYPETRAVAIL > 0 then
  begin
    G_NLIG.ColEditables[SGN_TYPETRAVAIL] := False;
    G_NLIG.Colwidths[SG_NATURETRAVAIL] := 20;
    G_NLIG.ColFormats[SG_NATURETRAVAIL] := 'CB=BTNATURETRAVAIL';
    G_NLIG.ColDrawingModes[SG_NATURETRAVAIL]:= 'IMAGE'
  end;

  if SGN_CODEMARCHE > 0 then
  begin
    G_NLIG.ColEditables[SGN_CODEMARCHE] := False;
    G_NLIG.Colwidths[SGN_CODEMARCHE] := 12;
  end;

  if SGN_LIBFOU > 0 then G_NLIG.ColEditables[SGN_LIBFOU] := False;

  // Modif BTP
  if SGN_TARIFF <> -1 then
  BEGIN
    if Action <> TaConsult then G_NLIG.ColLengths [SGN_TARIFF] := -1;
  END;
  //
  if SGN_UNAC <> -1 then
  begin
    if Action <> TaConsult then G_NLIG.ColLengths [SGN_UNAC] := -1;
  end;
  //
  if SGN_REMF <> -1 then
  begin
  if Action <> TaConsult then G_NLIG.ColLengths [SGN_REMF] := -1;
  end;
  //
  if SGN_PVHTP <> -1 then
  begin
    if Action <> TaConsult then G_NLIG.ColLengths [SGN_PVHTP] := -1;
  end;
  //
  if SGN_MO <> -1 then
  begin
    if Action <> TaConsult then G_NLIG.ColLengths [SGN_MO] := -1;
  end;
  //
  if SGN_TYPA <> -1 then
  BEGIN
    G_NLIG.ColWidths [SGN_TYPA] := 15;
    if Action <> TaConsult then G_NLIG.ColLengths [SGN_TYPA] := -1;
  END;

  if TypeArp then
  BEGIN
    G_NLIG.Collengths[SGN_QTE] := -1;
    G_NLIG.ColWidths [SGN_QTE] := 0;
    G_NLIG.Collengths[SGN_TYPA] := -1;
    G_NLIG.ColWidths [SGN_TYPA] := 0;
  END;
{$IFDEF BTP}

  NEPA.Decimals             := V_PGI.okdecP;
  NEPA.NumericType          := ntDecimal;
  NEPR.Decimals             := V_PGI.okdecP;
  NEPR.NumericType          := ntDecimal;
  NEPVHT.Decimals           := V_PGI.okdecP;
  NEPVHT.NumericType        := ntDecimal;
  NEPVFORFAIT.Decimals      := V_PGI.okdecP;
  NEPVFORFAIT.NumericType   := ntDecimal;
  NEMTACH.Decimals          := V_PGI.okdecP;
  NEMTACH.NumericType       := ntDecimal;
  NEMTPR.Decimals           := V_PGI.okdecP;
  NEMTPR.NumericType        := ntDecimal;

  NEPVTTC.Decimals          := V_PGI.okdecP;
  NEPVTTC.NumericType       := ntDecimal;
  NETPS.Decimals            := V_PGI.okdecQ;
  NETPS.NumericType         := ntDecimal;
  //
  MONTANTPA.Decimals        := V_PGI.okdecP;
  MONTANTPA.NumericType     := ntDecimal;
  MONTANTFG.Decimals        := V_PGI.okdecP;
  MONTANTFG.NumericType     := ntDecimal;
  MONTANTPR.Decimals        := V_PGI.okdecP;
  MONTANTPR.NumericType     := ntDecimal;
  MONTANTMARG.Decimals      := V_PGI.okdecP;
  MONTANTMARG.NumericType   := ntDecimal;
  MONTANTHT.Decimals        := V_PGI.okdecP;
  MONTANTHT.NumericType     := ntDecimal;
  MONTANTTTC.Decimals       := V_PGI.okdecP;
  MONTANTTTC.NumericType    := ntDecimal;
  TOTALTPS.Decimals         := V_PGI.okdecQ;
  TOTALTPS.NumericType      := ntDecimal;
  //
  MONTANTPAU.Decimals       := V_PGI.okdecP;
  MONTANTPAU.NumericType    := ntDecimal;
  MONTANTFGU.Decimals       := V_PGI.okdecP;
  MONTANTFGU.NumericType    := ntDecimal;
  MONTANTPRU.Decimals       := V_PGI.okdecP;
  MONTANTPRU.NumericType    := ntDecimal;
  MONTANTMARGU.Decimals     := V_PGI.okdecP;
  MONTANTMARGU.NumericType  := ntDecimal;
  MONTANTHTU.Decimals       := V_PGI.okdecP;
  MONTANTHTU.NumericType    := ntDecimal;
  MONTANTTTCU.Decimals      := V_PGI.okdecP;
  MONTANTTTCU.NumericType   := ntDecimal;
  TPSUNITAIRE.Decimals      := V_PGI.okdecQ;
  TPSUNITAIRE.NumericType   := ntDecimal;
  //
{$ENDIF}
end;

procedure TFLigNomen.FormShow(Sender: TObject);
var StLib     : String;
    StCompose : string;
begin
{$IFDEF BTP}
  TOBPiece := TOBLigPiece.Parent;
  fGestQteAvanc := GetParamSocSecur('SO_BTGESTQTEAVANC',false)  and (ISDocumentChiffrage(TOBPiece));

  RecupPrix := GetInfoParPiece(TOBPIece.GetValue('GP_NATUREPIECEG'), 'GPP_APPELPRIX');

  Reajuste := isExistsArticle (trim(GetParamsoc('SO_BTECARTPMA')));

  if TOBLigpiece.getValue('GL_BLOQUETARIF')='X' then
  begin
    Hlabel10.Visible    := false;
    Hlabel11.Visible    := false;
    Hlabel13.Visible    := false;
    Hlabel15.Visible    := false;
    MONTANTHT.Visible   := false;
    MONTANTTTC.Visible  := false;
    COEFMARG.Visible    := false;
    MONTANTMARG.Visible := false;
    TDEVHT.Visible      := false;
    TDEVTTC.Visible     := false;
  end;
{$ENDIF}

  if TypeArp then G_NLIG.RowCount := 2;

  if Ouvrage then
  begin
    PopG_NLIG_N.Caption := 'Détail ouvrage';
    PopG_NLIG_N.ShortCut := ShortCut( Word('O'), [ssCtrl]);
    Top     := 0;
    left    := 0;
    Width   := Screen.Monitors [0].Width;
    height  := Screen.Monitors [0].Height - 80;
  end;

  PANDIM.Visible := PopG_NLIG_D.Checked;
  bAnnule := false;
  bValide := false;
  Data_Modified := False;

  stCompose         := TobLig.Detail[0].GetValue(prefixe+'_COMPOSE');
  stLib             := RechDom('GCARTICLE',stCompose,false);
  COMPOSE.Text      := stCompose ;
  TCOMPOSE.Caption  := Copy(stCompose,0,19) + stLib;

  // Modif BTP
  // iTableLigne := PrefixeToNum('GLN') ;
  iTableLigne := PrefixeToNum(prefixe) ;

  // Modif BTP
  // AffecteLibelle(TOBLig, 'GLN_CODEARTICLE');
  AffecteLibelle(TOBLig, prefixe+'_CODEARTICLE');

  // Modif BTP
  if Ouvrage then
  begin
    //uniquement en line
    G_NLIG.ListeParam:='BTLIGOUV'
  end
  else
  begin
   G_NLIG.ListeParam:='GCLIGNOMEN';
  end;

  EtudieColsListe;
  AffecteGrid (G_NLIG,Action) ;

  // Modif BTP
  DefiniColsGrid;

  //
  fPieceCoTrait.Affaire := TOBaffaire;
  fPieceCoTrait.TOBpiece  := TOBpiece;
  fPieceCoTrait.TOBOuvrage := TobLig;
  fPieceCoTrait.Coltravail := SGN_TYPETRAVAIL;
  fPieceCoTrait.CollibFou := SGN_LIBFOU;

  {$IFDEF DANSSOUSDETAIL}
  fPieceCoTrait.InUse := (CoTraitOk) and ((TOBLigPiece.getValue('GLC_NATURETRAVAIL')='') or (valeur(TOBLigPiece.getValue('GLC_NATURETRAVAIL'))>= 10));
  {$ELSE}
  fPieceCoTrait.InUse := false;
  {$ENDIF}
//
  if TobVide(TobLig) then
  begin
    G_NLIG.Cells[0, 1]        := IntToStr(1);
    G_NLIG.Cells [SGN_Qte, 1] := StrS(1,V_PGI.OkDecQ);
  end
  else
    TOBLig.PutGridDetail(G_NLIG, False, False, LesColNLIG);

  // Modif BTP
  if (Ouvrage) and (not GestionHt) then G_NLIG.cells[SGN_PvHt,0] := TraduireMemoire ('Prix TTC');

  HMTrad.ResizeGridColumns (G_NLIG) ;

  AfficheTVNLIG;

  if (InfoNiv = 1) and (TOBLig.Detail.Count >= 0) then
  begin
    RecupInfoPiece;
    GereSousNomen(TobLig);
  end;

    ChargeTobArticles;

    TCODEDIM.Caption := '';
    //Cancel := False;
  // Suppression de la fille contenant les infos de la pièce en cas de nomenclature
  // sans composants
  if (TobVide(TobLig)) and (TobLig.Detail.Count=1) then TobLig.Detail[0].free;

  {$IFDEF BTP}
  if Ouvrage then
  begin
    CalculeOuvrageLoc;
    savouvrage := CalculSurTableau('+',TotalOuv,ValPourcent);
    savouvrage := CalculSurTableau('/',savouvrage,QteDudetail);
    afficheToTalisation;
  end;

  // Modif BTP
  if Ouvrage then
  begin
//   if (TypeArp) and (MetreDoc) then Appel_Xl.Visible := true;
   POuvrage.Visible := true;
   if TypeArp then
      BEGIN
      caption := TraduireMemoire ('Valorisation article')+' '+TOBligPiece.GetValue('GL_LIBELLE');
      TGNE_ARTICLE.caption := 'Article';
      END else
      BEGIN
      Caption := TraduireMemoire ('Décomposition de l''ouvrage ');
      TGNE_ARTICLE.caption := 'Ouvrage';
      END;
   if SGN_TYPA <> -1 then
      BEGIN
      if (Action <> TaConsult) then G_NLIG.ColLengths[SGN_TYPA] := -1;
      END;
   if not GestionHt then
      begin
      // dans le cas d'une saisie de document en TTC les elements de base sont en TTC
      TLigPvBase.Caption    := TraduireMemoire('PU TTC de base');
      TLigPvsaisie.Caption  := TraduireMemoire('Montant ligne TTC');
      TPUTTCLIG.Caption     := TraduireMemoire('Prix de vente HT');
      end;
  end
  else
    Caption := TraduireMemoire ('Décomposition de la nomenclature ');

  if action <> taconsult then BintegreExcel.Visible := true;
{$ELSE}
   Caption := TraduireMemoire ('Décomposition de la nomenclature ');
{$ENDIF}

  TheVariante.document  := TOBLigPiece;
  TheVariante.Ouvrage   := TOBLIG;
  SetControlsEnabled ;
  if not fGestQteAvanc then
  begin
    if SGN_QTESAIS >= 0 then
    begin
      G_NLIG.ColWidths[SGN_QTESAIS] := -1;
      G_NLIG.ColLengths[SGN_QTESAIS] := -1;
    end;
    if SGN_UNITESAIS >= 0 then
    begin
      G_NLIG.ColWidths[SGN_UNITESAIS] := -1;
      G_NLIG.ColLengths[SGN_UNITESAIS] := -1;
    end;
    if SGN_PERTE >= 0 then
    begin
      G_NLIG.ColWidths[SGN_PERTE] := -1;
      G_NLIG.ColLengths[SGN_PERTE] := -1;
    end;
    if SGN_RENDEMENT >= 0 then
    begin
      G_NLIG.ColWidths[SGN_RENDEMENT] := -1;
      G_NLIG.ColLengths[SGN_RENDEMENT] := -1;
    end;
  end;
  G_NLIGEnter (self);

end;

procedure TFLigNomen.SetControlsEnabled;
begin
	QTEDETAIL.Enabled := (Action<>taConsult);
	BVALIDER.Visible := (Action<>taConsult);
	BNewLigne.Visible := (Action<>taConsult);
	BDelLigne.Visible := (Action<>taConsult);
  PopG_NLIG_S.enabled := (Action<>taConsult);
  PopG_NLIG_I.enabled := (Action<>taConsult);
end;

procedure TFLigNomen.FormKeyDown(Sender: TObject; var Key: Word;
                                 Shift: TShiftState);
var FocusGrid : Boolean;
    ARow : Longint;
begin
BTypeArticle.visible := false;
FocusGrid := False;
if(Screen.ActiveControl = G_NLIG) then FocusGrid := True;
ARow := G_NLIG.Row;

LastKey := Key;
Case Key of
    VK_F10: BEGIN
    					Key := 0;
              if Action <> taCOnsult then BvaliderClick (Self);
    				END;
  	VK_ESCAPE : BEGIN
    							Key := 0;
                  close;
    					  END;
    VK_RETURN : Key:=VK_TAB ;
    VK_INSERT : BEGIN
    						if Action = Taconsult then exit;
{$IFDEF BTP}
                if (focusGrid) and (Shift=[ssCtrl]) then
                   begin
                   if IsModeSelect = true then
                      begin
                      ModeNormal;
                      end else
                      begin
                      ModeSelection
                      end;
                   key := 0;
                   end else
                   begin
                   if (FocusGrid) then
                      BEGIN
                      Key := 0;
                      InsertLigne (ARow, False);
                      END;
                   end;
{$ELSE}
                if (FocusGrid) then
                   BEGIN
                   Key := 0;
                   InsertLigne (ARow, False);
                   END;
{$ENDIF}
                END;
//
    VK_F5: if (FocusGrid) then
      begin
        Key := 0;
        ZoomArticle;
      end;
//
    VK_DELETE : BEGIN
    						if Action = Taconsult then exit;
                if ((FocusGrid) and (Shift=[ssCtrl])) then
                    BEGIN
                    Key := 0 ;
                    SupprimeLigne (ARow) ;
                    END ;
                END;
    END;
end;

procedure TFLigNomen.FormClose(Sender: TObject; var Action: TCloseAction);
var w_rep : word;
    i_ind : integer;
begin
  for i_ind := 1 to G_NLIG.RowCount - 1 do
    begin
    if TOB(G_NLIG.Objects[SGN_IdComp, i_ind]) <> nil then TOB(G_NLIG.Objects[SGN_IdComp, i_ind]).free;
    end;
  //
  if (InfoNiv = 1) and (TOBLig.Detail.Count >= 0) then
   begin
   DepointeSousNomen(Toblig);
   end;
  //
  if TobInfo <> nil then TOBInfo.free;
  //
  ART.free;
  // VARIANTE
  TheVariante.free;
  fPieceCoTrait.free;
end;


////////////////////////////////////////////////////////////////////////////////
//***************************** Gestion du grid ******************************//
////////////////////////////////////////////////////////////////////////////////

procedure TFLigNomen.EtudieColsListe ;
Var NomCol,LesCols : String ;
    icol,ichamp, i_ind : integer ;
  Mcd : IMCDServiceCOM;
  Field : IFieldCOM  ;
BEGIN
MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();

SetLength(ColsInter,G_NLIG.ColCount) ;
for i_ind:=Low(ColsInter)to High(ColsInter) do ColsInter[i_ind]:=False ;
LesCols:=G_NLIG.Titres[0] ;
{$IFDEF BTP}
   LesCols:=FindEtReplace(LesCols,'GCA_PRIXVENTE','MONTANTHTDEV',False) ;
{$ENDIF}

{$IFDEF BTP}
  if (not GestionHt) then
  BEGIN
    LesCols:=FindEtReplace(LesCols,prefixe+'_PUHTDEV',prefixe+'_PUTTCDEV',False) ;
  END;
{$ENDIF}
  LesColNLIG:=LesCols ; icol:=0 ;

  Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    //
    if NomCol<>'' then
    BEGIN
    	Field := (Mcd.GetField(Nomcol) as IFieldCOM);
      ichamp:=ChampToNum(NomCol) ;
      if ichamp>=0 then
      BEGIN
        ColsInter[icol]:=false;
        if (Pos(prefixe+'_',NomCol) > 0 ) and
        	 (Pos('X',Field.Control)>0) then ColsInter[icol]:=True
           														else ColsInter[icol]:=true;
        //
        if NomCol=prefixe+'_NUMORDRE'                               then SGN_NumLig     := icol  else
        if NomCol=prefixe+'_ARTICLE'                                then SGN_IdComp     := icol  else
        if NomCol=prefixe+'_CODEARTICLE'                            then SGN_Comp       := icol  else
        if NomCol=prefixe+'_COMPOSE'                                then SGN_Compose    := icol  else
        if NomCol=prefixe+'_LIBELLE'                                then SGN_Lib        := icol  else
        if NomCol = prefixe+'_MONTANTHTDEV'                         Then SGN_MONTANTHTDEV := icol;
        if (NomCol='GLN_QTE')            or (NomCol='BLO_QTEFACT')  then SGN_Qte        := icol  else
        if (NomCol=Prefixe+'_PUHTDEV')  and (GestionHt)             then SGN_Pvht       := icol  else
        if (NomCol=Prefixe+'_PUTTCDEV') and (not GestionHt)         then SGN_Pvht       := icol  else
        if NomCol=prefixe+'_TYPEARTICLE'                            then SGN_TypA       := icol  else
        // MODIF BTP
        if NomCol = 'GA_FOURNPRINC'                                 Then SGN_FOURN      := icol  else
        if NomCol = 'GCA_PRIXBASE'                                  Then SGN_TARIFF     := icol  else
        if NomCol = 'GCA_QUALIFUNITEACH'                            Then SGN_UNAC       := icol  else
        if NomCol = 'GF_CALCULREMISE'                               Then SGN_REMF       := icol  else
        if NomCol = 'GA_PAHT'                                       Then SGN_PAHT       := icol  else
        if NomCol = 'GA_PVHT'                                       Then SGN_PVHTP      := icol  else
        if NomCol = 'GA_NATUREPRES'                                 Then SGN_PREST      := icol  else
        if NomCol = 'GA_HEURE'                                      Then SGN_TPS        := icol  else
        if NomCol = 'GF_PRIXUNITAIRE'                               Then SGN_MO         := icol  else
        if NomCol = prefixe+'_NATURETRAVAIL'                        Then SGN_TYPETRAVAIL:= icol  else
        if NomCol = prefixe+'_FOURNISSEUR'                          Then SGN_LIBFOU     := icol  else
        if NomCol = 'BLO_QTESAIS'                                   Then SGN_QTESAIS    := icol  else
        if NomCol = 'BLO_QUALIFHEURE'                               Then SGN_UNITESAIS  := icol  else
        if NomCol = 'BLO_PERTE'                                     Then SGN_PERTE      := icol  else
        if NomCol = 'BLO_RENDEMENT'                                 Then SGN_RENDEMENT  := icol  else
        if NomCol = 'BLO_CODEMARCHE'                                Then SGN_CODEMARCHE  := icol  else
        // --- FIN MODIF
        if (NomCol=Prefixe+'_QUALIFQTEVTE')                         then SGN_UV         := icol;
      END
      else
        if NomCol = 'MONTANTHTDEV'                                  Then SGN_MONTANT    := icol;
    End ;
    Inc(icol) ;
  Until ((LesCols='') or (NomCol='')) ;

end;

////////////////////////////////////////////////////////////////////////////////
//*************************** Evenement sur le grid **************************//
////////////////////////////////////////////////////////////////////////////////
procedure TFLigNomen.enablePopG_NLIG(ou:integer);
var ARect : Trect;
    NumGraph : Integer;
begin
if TOB(G_NLIG.Objects[SGN_NumLig, Ou]).GetValue('SOUSNOMEN')='X' then
     PopG_NLIG_N.Enabled := True
else PopG_NLIG_N.Enabled := False;

if CasEmploi_Open then
     PopG_NLIG_C.Enabled := True
else PopG_NLIG_C.Enabled := False;

if (SGN_TypA <> -1) and (Action <> taconsult) then
   BEGIN
   if TOB(G_NLIG.Objects[SGN_NumLig,Ou]) <> nil then
      BEGIN
      ARect:=G_NLIG.CellRect(SGN_Typa,G_NLIG.Row) ;
      NumGraph := RecupTypeGraph (TOB(G_NLIG.Objects[SGN_NumLig,Ou]));
      BTypeArticle.ImageIndex := NumGraph;
      BTYpeArticle.Opaque := false;
      With BTypeArticle do
         Begin
         Top := Arect.top - G_NLIG.GridLineWidth ;
         Left := Arect.Left;
         Width := Arect.Right - Arect.Left;
         Height  := Arect.Bottom - Arect.Top - G_NLIG.GridLineWidth ;
         end;
      BTypeArticle.Parent := G_NLIG;
      BTypeArticle.Visible := true;
      END;
   END;
end;

procedure TFLigNomen.G_NLIGRowEnter(Sender: TObject; Ou: Integer;
                                    var Cancel: Boolean; Chg: Boolean);
var TOBOuv : TOB;
    ARect : Trect;
    NumGraph : Integer;
begin
//  Verification qu'on n'entre pas sur une ligne vide au dela de la premiere possible
if IsModeSelect then AutoriseEntreeSelection (ou,cancel);
if (action <> taconsult) and (Ou>=G_NLIG.RowCount-1) and (not TypeArp) then G_NLIG.RowCount:=G_NLIG.RowCount+1;
if (Ou > TOBLig.detail.count) and (not OnValidation) then
    begin
    // Modif BTP
    if TypeArp then BEGIN Cancel:=True; Exit; END;
    // --
    G_NLIG.Cells[SGN_NumLig, Ou] := IntToStr(TOBlig.Detail.Count+1);
    AjouteTobLigne(Ou);
    G_NLIG.Cells [SGN_Qte, Ou] := StrS(1,V_PGI.OkDecQ);
    // Modif BTP
{$IFDEF BTP}
    if Ouvrage then
       begin
       afficheMontantLigne (Ou);
       end;
{$ENDIF}
    AffectePied(Ou);
    G_NLIG.Col := SGN_Comp;
    Cell_Text := '';
    exit;
    end;
{$IFDEF BTP}
if Ouvrage then afficheMontantLigne(ou);
{$ENDIF}
if G_NLIG.Cells[SGN_Comp, Ou] = '' then exit;

//  affichage synchrone du treeview avec la grid
TV_NLIG.HideSelection := False;
TV_NLIG.Selected := TTreeNode(G_NLIG.Objects[SGN_Comp, Ou]);
{$IFDEF BTP}
if (TOB(G_NLIG.objects[SGN_numlig,ou]) <> nil) and
   (TOB(G_NLIG.objects[SGN_numlig,ou]).getValue('BLO_TYPEARTICLE')<>'ARP') then
{$ENDIF}
if (PopTV_NLIG_O.Checked) and (TV_NLIG.Selected <> nil) then TV_NLIG.Selected.Expand(False);
SetImages(TV_NLIG, nil);

//  Mise a jour du pied
AffectePied(Ou);
enablePopG_NLIG (ou);
{$IFDEF BTP}
if Ouvrage then
   begin
   afficheMontantLigne (Ou);
   end;
{$ENDIF}
if (SGN_TypA <> -1) and (G_NLIG.ColLengths[SGN_TYPA] > 0) and (Action <> taconsult) then
   BEGIN
   TOBOuv := TOB(G_NLIG.Objects[SGN_NumLig,Ou]);
   if TOBOuv <> nil then
      BEGIN
      ARect:=G_NLIG.CellRect(SGN_Typa,G_NLIG.Row) ;
      NumGraph := RecupTypeGraph (TOBOuv);
      BTypeArticle.ImageIndex := NumGraph;
      BTYpeArticle.Opaque := false;
      With BTypeArticle do
         Begin
         Top := Arect.top - G_NLIG.GridLineWidth ;
         Left := Arect.Left;
         Width := Arect.Right - Arect.Left;
         Height  := Arect.Bottom - Arect.Top - G_NLIG.GridLineWidth ;
         end;
      BTypeArticle.Parent := G_NLIG;
      BTypeArticle.Visible := true;
      END;
   END;
  // VARIANTE
TheVariante.Ligne := TOB(G_NLIG.Objects[SGN_NumLig,Ou]);

// Mode Sélection (Copier coller)
if IsModeSelect then
   begin
   AfficheSelectionLigne (ou);
   end;
end;

procedure TFLigNomen.G_NLIGRowExit(Sender: TObject; Ou: Integer;
                                   var Cancel: Boolean; Chg: Boolean);
begin
if Action = TaConsult then exit;
if BTypeArticle.Visible then BTypeArticle.visible := false;
//  On remonte et on sort d'une ligne vide au dela des lignes renseignées
//  ou on clique ailleurs et on est au dela des lignes renseignées
if (TypeARP) and (G_NLIG.Row >TOBLig.detail.count) then BEGIN Cancel := true; EXit; END;
// --
if (G_NLIG.row > Ou) and ((G_NLIG.cells[SGN_COMP,ou]='') or ((G_NLIG.cells[SGN_COMP,G_NLIG.row]='') and (G_NLIG.cells[SGN_COMP,G_NLIG.row -1]=''))) then
   begin
   cancel:=true;
   exit;
   end;
  if (G_NLIG.cells[SGN_COMP,OU]='') and (OU = TOBLig.detail.Count) then
    begin
    VideLaLigne(Ou);
    exit;
    end;

if Data_Modified then ValideLigne(Ou);
//--------------------------------------------------------------------------
//  affichage synchrone du treeview avec la grid
//--------------------------------------------------------------------------
TV_NLIG.HideSelection := False;
TV_NLIG.Selected := TTreeNode(G_NLIG.Objects[SGN_Comp, Ou]);
if (PopTV_NLIG_O.Checked) and (TV_NLIG.Selected <> nil) then TV_NLIG.Selected.Collapse(False);
end;

procedure TFLigNomen.G_NLIGCellEnter(Sender: TObject; var ACol,
                                     ARow: Integer; var Cancel: Boolean);
var LastRow : integer;
begin
if Action=taConsult then Exit ;
LastRow := Arow;
ZoneSuivanteOuOk(ACol,ARow,Cancel) ;
if Not Cancel then
   BEGIN
//
  if (Arow > TOBLig.detail.count) then
  begin
    // Modif BTP
    if TypeArp then BEGIN Cancel:=True; Exit; END;
    // --
    G_NLIG.Cells[SGN_NumLig, Arow] := IntToStr(TOBlig.Detail.Count+1);
    AjouteTobLigne(Arow);
    G_NLIG.Cells [SGN_Qte, Arow] := StrS(1,V_PGI.OkDecQ);
    // Modif BTP
    {$IFDEF BTP}
    if Ouvrage then
    begin
    	afficheMontantLigne (Arow);
    end;
    {$ENDIF}
    AffectePied(Arow);
    Cell_Text := '';
    exit;
  end;
//
   G_NLIG.ElipsisButton := (G_NLIG.col = SGN_Comp) or
                           (G_NLIG.col = SGN_FOURN ) or
                           (G_NLIG.col = SGN_PREST) or
                           ((G_NLIG.col= SGN_Qte) and (MetreDoc));
   Cell_Text := G_NLIG.Cells[G_NLIG.Col, G_NLIG.Row];
   END else
   begin
    if (Lastrow > TOBLig.detail.count) then
    begin
      // Modif BTP
      if TypeArp then BEGIN Cancel:=True; Exit; END;
      // --
      G_NLIG.Cells[SGN_NumLig, Lastrow] := IntToStr(TOBlig.Detail.Count+1);
      AjouteTobLigne(Lastrow);
      G_NLIG.Cells [SGN_Qte, Lastrow] := StrS(1,V_PGI.OkDecQ);
      // Modif BTP
      {$IFDEF BTP}
      if Ouvrage then
      begin
        afficheMontantLigne (Lastrow);
      end;
      {$ENDIF}
      AffectePied(Lastrow);
      Cell_Text := '';
      exit;
    end;
   end;
end;

//=============================================================================
//  a la sortie d'une cellule
//=============================================================================
procedure TFLigNomen.G_NLIGCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Var NewQte, Qte, OldQte : Double;
begin
  if csDestroying in ComponentState then Exit ;
  OldQte := 0;
  if (G_NLIG.Cells[SGN_Comp,ARow] = '') and (Arow <> TOBLig.detail.count) then begin cancel := true; exit; end;
  if (G_NLIG.ColLengths[Acol]= -1) then exit;

  FormateZoneSaisie(ACol, ARow);

  if Cell_Text = G_NLIG.Cells[ACol, ARow] then Exit;
  if (TypeARP) and (G_NLIG.Row >TOBLig.detail.count) then BEGIN Cancel := true; EXit; END;

{$IFDEF BTP}
  if (Acol = SGN_QTESAIS) or (Acol = SGN_PERTE) or (Acol = SGN_RENDEMENT) then
  begin
    if not TraiteQteSais (Acol,Arow)  then exit; 
  end;
  
  if (ACol = SGN_Qte) and (TheMetreShare <> nil) then
  begin
    NewQte := Valeur(G_NLIG.Cells[SGN_QTE,ARow]);  //StrToFloat(G_NLIG.Cells[SGN_QTE,ARow]);
    OldQte := TOB(G_NLIG.Objects[SGN_NumLig, ARow]).GetDouble('BLO_QTEFACT');
    if NewQte <> OldQte then
    begin
  //Si on modifie la quantité à la main (en principe) on demande si l'on désire supprimer le fichier XLS associé...
    if TheMetreShare.AutorisationMetre(TOBPIece.GetValue('GP_NATUREPIECEG')) then
    begin
      if (not ISFromExcel(TOB(G_NLIG.Objects[SGN_NumLig, ARow]))) then
      begin
        if not (TheMetreShare.QteModifie(TOB(G_NLIG.Objects[SGN_NumLig, ARow]),TOBLigPiece.getValue('GL_NUMORDRE'),OldQte,NewQte,True)) then
        begin
            G_NLIG.Cells[SGN_QTE,ARow] := Strf00(oldqte,V_PGI.OkdecQ);
          TOBLig.detail[Arow -1].PutValue('GL_QTEFACT', Valeur(G_NLIG.Cells[SGN_QTE,ARow]));
        end;
      end
      else
      begin
        G_NLIG.Cells[SGN_QTE,ARow] := FloatToStr(NewQte);
      end;
    end;
    end;
  //if not TheMetreShare.QteModifie (TOB(G_NLIG.Objects[SGN_NumLig, ARow])) then
  //BEGIN
  //  G_NLIG.Cells[SGN_QTE,ARow] := Cell_Text;
  //  exit;
  //END;
  end;

{$ENDIF}

if (ACol = SGN_Qte) and (not TraiterQuantite (ARow)) then Exit;  //pour eviter erreur si sortie de l'application
//  On remonte et on est sur une ligne vide ou on clique ailleurs et on est sur la
//  derniere ligne
(*
if ((LastKey = VK_UP) and (G_NLIG.Cells[SGN_Comp,ARow] = '')) or
   ((G_NLIG.Cells[SGN_Comp,ARow] = '') and (G_NLIG.Cells[SGN_Comp,ARow + 1] = '')) then Exit;
*)
if (G_NLIG.Cells[SGN_Comp,ARow] = '') and (Arow = TOBLig.detail.count) and (lastkey = VK_UP) then exit;
//  Simule le OnChange inexistant sur la cellule pour gestion de modif
if Cell_Text = G_NLIG.Cells[ACol, ARow] then Exit;

if Action <> taConsult then Data_Modified := True;
//  Controle du contenu de la cellule
if (ACol = SGN_Comp) and (not TraiterComposant(ARow)) then begin G_NLIG.Cells[SGN_Comp,ARow] := '';cancel:= false; Exit; end;
{$IFDEF BTP}
//  Modif du fournisseur
if (ACol = SGN_FOURN) Then TraiterFournisseur(ARow);
if (ACol = SGN_PREST) Then TraiterPrestation(ARow);
if (ACol = SGN_TPS)   Then TraiterTemps(ARow);
if (Acol = SGN_PAHT)  Then TraiterPAFourniture (ARow);
{$ENDIF}
//  Formattage de la cellule en cours
FormateZoneSaisie (ACol,ARow);
//  Pas de tob associée, on sort.
if G_NLIG.Objects[SGN_NumLig, ARow] = nil then Exit;

TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('LIBELLE', G_NLIG.Cells[SGN_Comp, ARow] +
                                       ' (' + StrS0(Valeur(G_NLIG.Cells[SGN_Qte, ARow])) + ')');
if Data_Modified then
   if ACol=SGN_Comp then AfficheTVNLIG;
{$IFDEF BTP}
if (Ouvrage) then
   begin
   if Data_Modified then ValideLigne(Arow);
   CalculeOuvrageLoc;
   AfficheTotalisation;
   AfficheMontantLigne (Arow);
   end;
TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutLigneGrid (G_NLIG,Arow,false,false,LesColNLIG);
{$ENDIF}
enablePopG_NLIG (Arow);
AffectePied(ARow);
//Cell_Text := G_NLIG.Cells[G_NLIG.Col, G_NLIG.Row];
end;

//=============================================================================
//  Lancement de la recherche dans la table article
//=============================================================================
procedure TFLigNomen.G_NLIGElipsisClick(Sender: TObject);
Var Coord : TRect;
{$IFDEF BTP}
    SWhere,STChamps : string;
    TOBPiece,TOBLig,TOBART : TOB;
    fournisseur : string;
{$ENDIF}
begin
if (G_NLIG.Col = SGN_Comp) then
    BEGIN
    ZoomArticle;
    END;
{$IFDEF BTP}
if (G_NLIG.col = SGN_Qte) then
begin
  Appel_XLClick(self);
end;

if (G_NLIG.Col = SGN_FOURN ) then
   begin
   TOBLIg := TOB(G_NLIG.objects[SGN_numlig,g_nlig.row]);
   TOBArt := TOB(G_NLIG.objects[SGN_idComp,g_nlig.row]);
   fournisseur := LookupFournisseur (TOBLIg.GetValue('BLO_FOURNISSEUR'),TOBLIG.GetValue('BLO_ARTICLE'),
                                     TOBLIG.GetValue('BLO_LIBELLE'));
   if (fournisseur <> '') and (fournisseur <> TobLIG.GetValue('BLO_FOURNISSEUR')) then
      begin
      Cell_Text := fournisseur;
      G_NLIG.Cells[SGN_FOURN,G_NLIG.row] := fournisseur;
      TOBLIG.PutValue('BLO_FOURNISSEUR',fournisseur);
      TraiterFournisseur (G_NLIG.row);
      end;
   end;

if (G_NLIG.Col = SGN_PREST ) then
   begin
   TOBLIg := TOB(G_NLIG.objects[SGN_numlig,g_nlig.row]);
   Coord := G_NLIG.CellRect (G_NLIG.Col, G_NLIG.Row);
   ART.Parent := G_NLIG;
   ART.Top := Coord.Top;
   ART.Left := Coord.Left;
   ART.Width := 3; ART.Visible := False;
   ART.DataType := 'GCARTICLE';
   ART.Text := G_NLIG.Cells[G_NLIG.Col,G_NLIG.Row];
   sWhere:='GA_TYPEARTICLE="PRE"';
   StChamps:='GA_CODEARTICLE='+ART.Text+';XX_WHERE=AND '+sWhere ;
   DispatchRecherche(ART, 1, '', stchamps, '');
   if (Art.Text  <> '') and (Art.text <> TobLIG.GetValue('GA_NATUREPRES')) then
      begin
      Cell_Text := ART.text;
      G_NLIG.Cells[SGN_PREST,G_NLIG.row] := copy(ART.text,1,18);
      TOBLIG.PutValue('GA_NATUREPRES',copy(ART.text,1,18));
      TOBLIG.PutValue('GA_ARTICLE',ART.text);
      TraiterPrestation (G_NLIG.row);
      end;
   end;
{$ENDIF}
end;

//=============================================================================
//  double clic sur une ligne on ouvre la fiche article en consultation
//=============================================================================
procedure TFLigNomen.G_NLIGDblClick(Sender: TObject);
var Cancel : boolean;
begin
(*
if (G_NLIG.Cells[SGN_Comp, G_NLIG.Row] = '') then
    begin
    if TobLig.Detail.Count>1 then G_NLIGRowExit(Sender, G_NLIG.Row, Cancel, False);
    G_NLIG.Row:=1;
    Exit;
    end;
*)
if (G_NLIG.Cells[SGN_Comp, G_NLIG.Row] = '') then Exit;
// modif Btp
if Ouvrage then V_PGI.DispatchTT (7,taconsult,G_NLIG.Cells[SGN_IdComp, G_NLIG.Row],'','ACTION=CONSULTATION;MONOFICHE')
else AGLLanceFiche('GC','GCARTICLE','',G_NLIG.Cells[SGN_IdComp, G_NLIG.Row],'ACTION=CONSULTATION')
end;


//=============================================================================
//  Deplacement d'une ligne dans la grid
//=============================================================================
procedure TFLigNomen.G_NLIGRowMoved(Sender: TObject; FromIndex, ToIndex: Integer);
var
    i_ind1, i_ind2 : integer;
begin
//if (LastKey = VK_INSERT) or (LastKey = VK_DELETE) or (LastKey = VK_DOWN) then Exit;
if LastKey <> 0 then Exit;
//  On verifie que la ligne n'est pas deplacée au dela de la derniere ligne renseignee
if FromIndex < ToIndex then
    begin
    for i_ind1 := ToIndex - 1 downto 1 do
        if G_NLIG.Cells[SGN_Comp, i_ind1] <> '' then break;
    //  si c'est le cas on ramene la ligne deplacé apres la derniere
    if (i_ind1 + 1) <> ToIndex then
        begin
        G_NLIG.Objects[SGN_NumLig, i_ind1 + 1] := G_NLIG.Objects[SGN_NumLig, ToIndex];
        G_NLIG.Objects[SGN_NumLig, ToIndex] := nil;
        for i_ind2 := 0 to G_NLIG.ColCount - 1 do
            begin
            G_NLIG.Cells[i_ind2, i_ind1 + 1] := G_NLIG.Cells[i_ind2, ToIndex];
            G_NLIG.Cells[i_ind2, ToIndex] := '';
            end;
        ToIndex := i_ind1 + 1;
        end;
    end;
//  on renumerote toutes les lignes
Renumerote;
//  on retrie les tobs associées
if FromIndex < ToIndex then
    for i_ind1 := FromIndex to ToIndex - 1 do
        TOBLig.Detail.Exchange(i_ind1 - 1, i_ind1)
else
    for i_ind1 := FromIndex - 1 downto ToIndex do
        if i_ind1 < TobLig.Detail.Count then
        TOBLig.Detail.Exchange(i_ind1 - 1, i_ind1);

//  on rafraichit le treeview
AfficheTVNLIG;
end;

//=============================================================================
//  Mousedown dans la grid pour remettre lastkey à 0
//=============================================================================
procedure TFLigNomen.G_NLIGMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
LastKey := 0;
end;

////////////////////////////////////////////////////////////////////////////////
//*************************** Gestion des champs *****************************//
////////////////////////////////////////////////////////////////////////////////
{$IFDEF BTP}
procedure TFLigNomen.TraiterFournisseur (Arow :integer);
var TOBL : TOB;
    TOBART : TOB;
    valeur : T_valeurs;
    CalculPv : boolean;
    Qte : double;
begin
TOBL:= TOB(G_NLIG.objects[SGN_numlig,Arow]);
if TOBL = nil then exit;
Qte := TOBL.GetValue('BLO_QTEFACT');
if (TOBL.GetValue('BLO_TYPEARTICLE') <> 'MAR') and
   (TOBL.GetValue('BLO_TYPEARTICLE') <> 'ARP') then exit;
//
TOBL.PutValue ('BLO_FOURNISSEUR',G_NLIG.cells[SGN_FOURN,Arow]);
if (TOBL.GetValue('BLO_TYPEARTICLE')= 'ARP') and (TOBL.detail.count > 0) then TOBL.Detail[0].putvalue ('BLO_FOURNISSEUR',TOBL.GetValue('BLO_FOURNISSEUR'));
//
CalculPv := true;
if  (Cell_Text <> '' ) then
begin
 if PgiAskAf ('Répercussion sur le prix de vente ?',caption) = mrNo then
    begin
    CalculPv := false;
    end;
end;
TOBArt:= TOB(G_NLIG.objects[SGN_IdComp,Arow]);
//TOBArt.putValue ('GA_FOURNPRINC',G_NLIG.cells[SGN_FOURN,Arow]);
if (TOBL.GetValue('BLO_TYPEARTICLE')= 'ARP') and (TOBL.detail.count >0) then
   begin
   InitValoDetail (TOBPiece,TOBL.detail[0],TOBArt,DEV,CalculPv);
   AppliqueCoefDomaineActOuv (TOBL.detail[0],'A',RecupPrix);
   if VH_GC.BTCODESPECIF = '001' then ApliqueCoeflignePOC (TOBL.detail[0],DEV);
	 CalculeLigneAcOuv (TOBL.detail[0],DEV,CalculPv,TOBART);
   TOBL.detail[0].putvalue ('GA_PVHT',TOBL.detail[0].GetValue ('BLO_PUHT'));
   GetValoDetail (TOBL);
   CalculOuvFromDetail (TOBL,DEV,Valeur);
   TOBL.Putvalue ('BLO_DPA',Valeur[0]);
   TOBL.Putvalue ('BLO_DPR',Valeur[1]);
//   TOBL.putvalue ('GA_PVHT',valeur[2]);
   TOBL.Putvalue ('BLO_PUHT',Valeur[2]);
   TOBL.Putvalue ('BLO_PUHTBASE',Valeur[2]);
   TOBL.Putvalue ('BLO_PUHTDEV',pivottodevise(Valeur[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
   TOBL.Putvalue ('BLO_PUTTC',Valeur[3]);
   TOBL.Putvalue ('BLO_PUTTCDEV',pivottodevise(Valeur[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
   TOBL.PutValue('ANCPV', TOBL.GetValue('BLO_PUHTDEV'));
   TOBL.PutValue('ANCPA', TOBL.GetValue('BLO_DPA'));
   TOBL.PutValue('ANCPR', TOBL.GetValue('BLO_DPR'));
   TOBL.Putvalue ('BLO_PMAP',Valeur[6]);
   TOBL.Putvalue ('BLO_PMRP',Valeur[7]);
   //
    TOBL.PutValue('BLO_TPSUNITAIRE',Valeur[9]);
    TOBL.PutValue('GA_HEURE',Valeur[9]);
    TOBL.PutValue('BLO_MONTANTFG',Arrondi(Valeur[13]*Qte,V_PGI.OkdecP));
    TOBL.PutValue('BLO_MONTANTFC',Arrondi(Valeur[15]*Qte,V_PGI.OkdecP));
    TOBL.PutValue('BLO_MONTANTFR',Arrondi(Valeur[14]*Qte,V_PGI.OkdecP));
    CalculeLigneAcOuvCumul (TOBL);
   //
   CalculMontantHtDevLigOuv (TOBL,DEV);
   end else
   begin
   InitValoDetail (TOBPiece,TOBL,TOBArt,DEV,CalculPv);
   AppliqueCoefDomaineActOuv (TOBL,'A',RecupPrix);
   if VH_GC.BTCODESPECIF = '001' then ApliqueCoeflignePOC (TOBL,DEV);
	 CalculeLigneAcOuv (TOBL,DEV,CalculPv,TOBART);
   end;
TOBL.PutLigneGrid (G_NLIG,Arow,false,false,LesColNLIG);
cell_text := G_NLIG.cells[SGN_FOURN,Arow];
ValideLigne(Arow);
CalculeOuvrageLoc;
AfficheTotalisation;
AfficheMontantLigne (Arow);
end;

procedure TFLigNomen.TraiterTemps(ARow : integer);
var TOBL,TOBART : TOB;
    valeurs : T_valeurs;
    Qte : double;
begin
TOBL:= TOB(G_NLIG.objects[SGN_numlig,Arow]); if TOBL = nil then exit;
TOBArt := TOB(G_NLIG.objects[SGN_idComp,Arow]);
Qte := TOBL.GetValue('BLO_QTEFACT');
if (TOBL.GetValue('BLO_TYPEARTICLE') <> 'ARP') or (TOBL.detail.count = 0 ) then exit;
if (TOBL.GetValue('BLO_TYPEARTICLE') = 'ARP') and (TOBL.detail.count < 2 ) then exit;
TOBL.detail[1].putValue('BLO_QTEFACT',valeur (G_NLIG.Cells[SGN_TPS,Arow]));
if (Pos(TOBL.detail[1].GetValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne)>0) and
	 (not IsLigneExternalise(TOBL.detail[1])) then
begin
	TOBL.detail[1].putValue('BLO_TPSUNITAIRE',1);
end else
begin
	TOBL.detail[1].putValue('BLO_TPSUNITAIRE',0);
end;
TOBL.Putvalue('GA_HEURE',valeur(G_NLIG.Cells[SGN_TPS,Arow]));
TOBL.Putvalue('BLO_TPSUNITAIRE',valeur(G_NLIG.Cells[SGN_TPS,Arow]));
CalculeLigneAcOuv (TOBL.detail[1],DEV,true,TOBART);
//RecupValoDetailHrs (TOBLIG);
CalculOuvFromDetail (TOBL,DEV,Valeurs);
TOBL.Putvalue ('BLO_DPA',Valeurs[0]);
TOBL.Putvalue ('BLO_DPR',Valeurs[1]);
TOBL.Putvalue ('BLO_PUHT',Valeurs[2]);
TOBL.Putvalue ('BLO_PUHTBASE',Valeurs[2]);
TOBL.Putvalue ('BLO_PUHTDEV',pivottodevise(Valeurs[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
TOBL.Putvalue ('BLO_PUTTC',Valeurs[3]);
TOBL.Putvalue ('BLO_PUTTCDEV',pivottodevise(Valeurs[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
TOBL.PutValue('ANCPV', TOBL.GetValue('BLO_PUHTDEV'));
TOBL.PutValue('ANCPA', TOBL.GetValue('BLO_DPA'));
TOBL.PutValue('ANCPR', TOBL.GetValue('BLO_DPR'));
TOBL.Putvalue ('BLO_PMAP',Valeurs[6]);
TOBL.Putvalue ('BLO_PMRP',Valeurs[7]);
//
TOBL.PutValue('BLO_TPSUNITAIRE',Valeurs[9]);
//TOBL.PutValue('GA_HEURE',Valeurs[9]);
TOBL.PutValue('BLO_MONTANTFG',Arrondi(Valeurs[13]*Qte,4));
TOBL.PutValue('BLO_MONTANTFC',Arrondi(Valeurs[15]*Qte,4));
TOBL.PutValue('BLO_MONTANTFR',Arrondi(Valeurs[14]*Qte,4));
//
CalculeLigneAcOuvCumul (TOBL);
CalculMontantHtDevLigOuv (TOBL,DEV);
TOBL.PutLigneGrid (G_NLIG,Arow,false,false,LesColNLIG);
cell_text := G_NLIG.Cells[SGN_TPS,Arow];
ValideLigne(Arow);
CalculeOuvrageLoc;
AfficheTotalisation;
AfficheMontantLigne (Arow);
end;

procedure TFLigNomen.TraiterPrestation(Arow: integer);
var TOBL,TOBLDET : TOB;
    TOBART : TOB;
    valeur : T_valeurs;
    CalculPv : boolean;
    Req : string;
    QQ : Tquery;
    TheArticle : string;
    Qte : double;
begin
TOBL:= TOB(G_NLIG.objects[SGN_numlig,Arow]); if TOBL = nil then exit;
Qte := TOBL.GetValue('BLO_QTEFACT');
if (TOBL.GetValue('BLO_TYPEARTICLE') <> 'ARP') or (TOBL.detail.count = 0 ) then exit;

CalculPv := true;
if  (Cell_Text <> '' ) then
   begin
   if PgiAskAf ('Répercussion sur le prix de vente ?',caption) = mrNo then
      begin
      CalculPv := false;
      end;
   end;

if TOBL.detail.count = 1 then
begin
  TOBLDet := TOB.create('LIGNEOUV',TOBL,-1);
  InsertionChampSupOuv (TOBLDET,false);
  //
  TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
  TOBLDet.putValue('BLO_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));
  //
end;

if trim(G_NLIG.Cells[SGN_PREST,AROW]) <> '' then
begin
	TheArticle := trim(G_NLIG.Cells[SGN_PREST,AROW]);
  QQ := OPenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
              		'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
              		'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE GA_CODEARTICLE ="'+TheArticle+'"',True,-1, '', True);
  TOBArt:= TOB.create ('ARTICLE',nil,-1);
  TOBArt.selectDB ('',QQ);
  ferme (QQ);
  InitChampsSupArticle (TOBART);
  // mise en place de la nouvelle prestation
  TOBL.detail[1].putValue('BLO_ARTICLE',TOBART.GetValue('GA_ARTICLE'));
  TOBL.detail[1].putValue('BLO_CODEARTICLE',TOBART.GetValue('GA_CODEARTICLE'));
  TOBL.detail[1].putValue('BLO_REFARTSAISIE',TOBART.GetValue('GA_CODEARTICLE'));
  TOBL.detail[1].putValue('BLO_LIBELLE',TOBART.GetValue('GA_LIBELLE'));
  RenseigneTOBOuv (TOBL.detail[1],TOBART,'',CalculPv);
	CalculeLigneAcOuv (TOBL.detail[1],DEV,CalculPv,TOBART);
  TOBART.free;
end else
begin
  TOBL.detail[1].putValue('BLO_ARTICLE','');
  TOBL.detail[1].putValue('BLO_CODEARTICLE','');
  TOBL.detail[1].putValue('BLO_REFARTSAISIE','');
  TOBL.detail[1].putValue('BLO_LIBELLE','');
  RenseigneTOBOuv (TOBL.detail[1],nil,'',CalculPv);
	CalculeLigneAcOuv (TOBL.detail[1],DEV,CalculPv);
end;

TOBL.detail[1].putValue('BLO_COMPOSE',TOBL.detail[0].GetValue('BLO_COMPOSE'));
TOBL.detail[1].putValue('BLO_NIVEAU',TOBL.detail[0].GetValue('BLO_NIVEAU'));
TOBL.detail[1].putValue('BLO_NUMORDRE',TOBL.detail[0].GetValue('BLO_NUMORDRE')+1);
// Correction FQ NEW/15041
if TOBL.detail[1].getValue('BLO_ARTICLE')= '' then
begin
	TOBL.detail[1].free;
  TOBL.putvalue('GA_HEURE',0) ;
  TOBL.PutValue('GF_PRIXUNITAIRE',0);
  TOBL.putvalue('TGA_PRESTATION','') ;
  TOBL.putvalue('GA_NATUREPRES','') ;
  TOBL.putvalue('GA_ARTICLE','') ;
end;
// ----------------------
//RecupValoDetailHrs (TOBL);
GetValoDetail (TOBL);
CalculOuvFromDetail (TOBL,DEV,Valeur);
TOBL.Putvalue ('BLO_DPA',Valeur[0]);
TOBL.Putvalue ('BLO_DPR',Valeur[1]);
if (TOBL.GetValue('BLO_TYPEARTICLE')='ARP') and (TOBL.Detail.count > 0) then
begin
	TOBL.putvalue ('GA_PVHT',TOBL.detail[0].GetValue('BLO_PUHT'));
end else
begin
	TOBL.putvalue ('GA_PVHT',valeur[2]);
end;
TOBL.Putvalue ('BLO_PUHT',Valeur[2]);
TOBL.Putvalue ('BLO_PUHTBASE',Valeur[2]);
TOBL.Putvalue ('BLO_PUHTDEV',pivottodevise(Valeur[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
TOBL.Putvalue ('BLO_PUTTC',Valeur[3]);
TOBL.Putvalue ('BLO_PUTTCDEV',pivottodevise(Valeur[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
TOBL.PutValue('ANCPV', TOBL.GetValue('BLO_PUHTDEV'));
TOBL.PutValue('ANCPA', TOBL.GetValue('BLO_DPA'));
TOBL.PutValue('ANCPR', TOBL.GetValue('BLO_DPR'));
TOBL.Putvalue ('BLO_PMAP',Valeur[6]);
TOBL.Putvalue ('BLO_PMRP',Valeur[7]);
//
TOBL.PutValue('BLO_TPSUNITAIRE',Valeur[9]);
TOBL.PutValue('GA_HEURE',Valeur[9]);
//
TOBL.PutValue('BLO_MONTANTFG',Arrondi(Valeur[13]*Qte,V_PGI.OkdecP));
TOBL.PutValue('BLO_MONTANTFC',Arrondi(Valeur[15]*Qte,V_PGI.OkdecP));
TOBL.PutValue('BLO_MONTANTFR',Arrondi(Valeur[14]*Qte,V_PGI.OkdecP));
CalculeLigneAcOuvCumul (TOBL);
CalculMontantHtDevLigOuv (TOBL,DEV);
TOBL.PutLigneGrid (G_NLIG,Arow,false,false,LesColNLIG);
cell_text := TOBL.GetValue('GA_NATUREPRES');
ValideLigne(Arow);
CalculeOuvrageLoc;
AfficheTotalisation;
AfficheMontantLigne (Arow);
RenumeroteLesFilles (TOBL,TOBL.detail[0].GetValue('BLO_NIVEAU')-1);
end;

procedure TFLigNomen.TraiterPAFourniture (ARow : integer);
var TOBL : TOB;
    TOBART : TOB;
    valeurs : T_valeurs;
    CalculPv : boolean;
    Qte : double;
begin

  TOBL      := TOB(G_NLIG.objects[SGN_numlig,Arow]); if TOBL = nil then exit;
  Qte       := TOBL.GetValue('BLO_QTEFACT');
  //
  CalculPv  := true;
  //
  if  (Cell_Text <> '' ) then
   begin
    if (TOBPIece.GetValue('GP_NATUREPIECEG')<>'FRC') then
    begin
   if PgiAskAf ('Répercussion sur le prix de vente ?',caption) = mrNo then
      begin
      CalculPv := false;
      end;
   end;
  end;

  TOBArt:= TOB(G_NLIG.objects[SGN_IdComp,Arow]);
  TOBArt.putValue ('GA_PAHT',valeur(G_NLIG.Cells[SGN_PAHT,Arow]));

  if (TOBL.GetValue('BLO_TYPEARTICLE')= 'ARP') and (TOBL.detail.count >0) then
  begin
    TOBL.detail[0].putvalue ('GA_PAHT',TOBART.GetValue('GA_PAHT'));
    TOBL.detail[0].putvalue ('BLO_DPA',TOBART.GetValue('GA_PAHT'));
    if ((RecupPrix='DPR') or (RecupPrix = 'PUH')) then
    begin
   if TOBL.detail[0].Getvalue ('BLO_COEFFG')=0 then TOBL.detail[0].Putvalue ('BLO_COEFFG',TOBArt.GetValue('GA_COEFFG')-1);
    end;
    if (RecupPrix = 'PUH') then
    begin
      if TOBL.detail[0].Getvalue ('BLO_COEFMARG')=0 then
      begin
        TOBL.detail[0].Putvalue ('BLO_COEFMARG',TOBArt.GetValue('GA_COEFCALCHT'));
        TOBL.detail[0].PutValue('POURCENTMARG',Arrondi((TOBL.detail[0].GetValue('BLO_COEFMARG')-1)*100,2));
      end;
    end;

    CalculeLigneAcOuv (TOBL.detail[0],DEV,CalculPv,TOBART);
    TOBL.detail[0].putvalue ('GA_PVHT',TOBL.detail[0].GetValue ('BLO_PUHT'));
    CalculMontantHtDevLigOuv (TOBL.detail[0],DEV);
    CalculOuvFromDetail (TOBL,DEV,Valeurs);
    TOBL.Putvalue ('BLO_DPA',Valeurs[0]);
    TOBL.Putvalue ('BLO_DPR',Valeurs[1]);
    TOBL.putvalue ('GA_PVHT',Valeurs[2]);
    TOBL.Putvalue ('BLO_PUHT',Valeurs[2]);
    TOBL.Putvalue ('BLO_PUHTBASE',Valeurs[2]);
    TOBL.Putvalue ('BLO_PUHTDEV',pivottodevise(Valeurs[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
    TOBL.Putvalue ('BLO_PUTTC',Valeurs[3]);
    TOBL.Putvalue ('BLO_PUTTCDEV',pivottodevise(Valeurs[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
    TOBL.Putvalue ('BLO_PMAP',Valeurs[6]);
    TOBL.Putvalue ('BLO_PMRP',Valeurs[7]);
    //
    TOBL.PutValue('BLO_TPSUNITAIRE',Valeurs[9]);
    TOBL.PutValue('GA_HEURE',Valeurs[9]);
    TOBL.PutValue('BLO_MONTANTFG',Arrondi(Valeurs[13]*Qte,V_PGI.OkdecP));
    TOBL.PutValue('BLO_MONTANTFC',Arrondi(Valeurs[15]*Qte,V_PGI.OkdecP));
    TOBL.PutValue('BLO_MONTANTFR',Arrondi(Valeurs[14]*Qte,V_PGI.OkdecP));
    CalculeLigneAcOuvCumul (TOBL);
    //
    GetValoDetail (TOBL);
    if TOBL.GetValue('BLO_PUHT') <> 0 then
    begin
      TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('BLO_PUHT')- TOBL.GetValue('BLO_DPR'))/TOBL.GetValue('BLO_PUHT'))*100,2));
    end else
    begin
      TOBL.PutValue('POURCENTMARQ',0);
    end;
  end
   else
   begin
   TOBL.putvalue ('GA_PAHT',TOBART.GetValue('GA_PAHT'));
   TOBL.putvalue ('BLO_DPA',TOBART.GetValue('GA_PAHT'));
	 CalculeLigneAcOuv (TOBL,DEV,CalculPv,TOBART);
   TOBL.putvalue ('GA_PVHT',TOBL.GetValue ('BLO_PUHT'));
   end ;
CalculMontantHtDevLigOuv (TOBL,DEV);
if (TOBL.GetValue('BLO_TYPEARTICLE')= 'ARP') and (TOBL.detail.count >0) then
  begin
  TOBL.PutValue('ANCPV', TOBL.GetValue('BLO_PUHTDEV'));
  TOBL.PutValue('ANCPA', TOBL.GetValue('BLO_DPA'));
  TOBL.PutValue('ANCPR', TOBL.GetValue('BLO_DPR'));
  end;
TOBL.PutLigneGrid (G_NLIG,Arow,false,false,LesColNLIG);
cell_text := G_NLIG.Cells[SGN_PAHT,Arow];
ValideLigne(Arow);
CalculeOuvrageLoc;
AfficheTotalisation;
AfficheMontantLigne (Arow);
end;

{$ENDIF}

function TFLigNomen.TraiterComposant(ARow: LongInt) : boolean;
var
    st1,st2 : string;
    TOBTemp : TOB;
    tn1 : TTreeNode;
    i_Rep : integer;
{$IFDEF BTP}
    Tresult : T_Valeurs;
    Val : Double;
{$ENDIF}
begin
Result := True;
//--------------------------------------------------------------------------
//  Code composant obligatoire
//--------------------------------------------------------------------------
if G_NLIG.Cells[SGN_Comp,ARow] = '' then
    BEGIN
//  Message code composant obligatoire
    MsgBox.Execute (5,Caption,'') ;
//    G_NLIG.Col := SGN_Comp; G_NLIG.Row := ARow;
    Result := False;
    Exit;
    END;

//--------------------------------------------------------------------------
//  Code composant unique dans un niveau donné de nomenclature
//--------------------------------------------------------------------------
if not Ouvrage then
   begin
   for i_Rep := 0 to G_NLIG.RowCount - 1 do
       begin
       if G_NLIG.Cells[SGN_Comp, i_Rep] = '' then Break;
       st1 := G_NLIG.Cells[SGN_idComp, i_Rep];
       st2 := G_NLIG.Cells[SGN_IdComp, ARow];
       if (i_Rep <> ARow) and (st1 = st2) then
          begin
          //  Message code composant unique dans la nomenclature
          MsgBox.Execute (6,Caption,'') ;
//          G_NLIG.Col := SGN_Comp; G_NLIG.Row := ARow;
          Result := False;
          Exit;
          end;
       end;
   end;
st1:='';
//--------------------------------------------------------------------------
//  Code composant renseigne. s'il est different du code identifiant article
//  ou que ce code identifiant est vide, c'est un changement de composant
//  ou une nouvelle ligne
//--------------------------------------------------------------------------
st1 := Trim(Copy(G_NLIG.Cells[SGN_IdComp,ARow], 0, Length(G_NLIG.Cells[SGN_Comp,ARow])));
if (G_NLIG.Cells[SGN_IdComp,ARow] = '') or (G_NLIG.Cells[SGN_Comp,ARow] <> st1) then
    begin
    // Modif BTP
    TOB(G_NLIG.Objects[SGN_NumLig, ARow]).Free;
    G_NLIG.Objects[SGN_NumLig, ARow] := nil;
    AjouteTobLigne(ARow);
    // Modif BTP
    if Ouvrage then TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('BLO_QTEFACT',Valeur(G_NLIG.Cells[SGN_Qte, ARow]))
    else TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('GLN_QTE',Valeur(G_NLIG.Cells[SGN_Qte, ARow]));
//--------------------------------------------------------------------------
//  on passe par la recherche d'article
//--------------------------------------------------------------------------
    if not(Recherche_Art(ARow)) then
        begin
        Result := False;
        exit ;
        end;

//--------------------------------------------------------------------------
//  Mise à jour du treeview
//--------------------------------------------------------------------------
    if (ARow = 1) then
//--------------------------------------------------------------------------
//  premiere ligne, on insere le premier noeud enfant
//--------------------------------------------------------------------------
        G_NLIG.Objects[SGN_Comp, ARow] := TV_NLIG.Items.AddChildFirst(TV_NLIG.Items[0], G_NLIG.Cells[SGN_Comp, ARow])
    else if (G_NLIG.Cells[SGN_Comp, ARow + 1] = '') or (ARow = G_NLIG.RowCount - 1) then
//--------------------------------------------------------------------------
//  derniere  on ajoute un noeud enfant en fin de liste
//--------------------------------------------------------------------------
        G_NLIG.Objects[SGN_Comp, ARow] := TV_NLIG.Items.AddChild(TV_NLIG.Items[0], G_NLIG.Cells[SGN_Comp, ARow])
    else
        begin
        tn1 := TTreeNode(G_NLIG.Objects[SGN_Comp, ARow - 1]);
        tn1 := tn1.GetNext;
        G_NLIG.Objects[SGN_Comp, ARow] := TV_NLIG.Items.Insert(tn1, G_NLIG.Cells[SGN_Comp, ARow]);
        TTreeNode(G_NLIG.Objects[SGN_Comp, ARow]).Text := G_NLIG.Cells[SGN_Comp, ARow];
        end;
//--------------------------------------------------------------------------
//  on va charger l'eventuelle sous nomenclature complete
//--------------------------------------------------------------------------
    TOBTemp := TOB(G_NLIG.Objects[SGN_NumLig, ARow]);
{$IFDEF BTP}
    // MODIF BTP
    TTreeNode(G_NLIG.Objects[SGN_Comp, ARow]).Data  := TOBTemp; // on fait pointer le treenode sur la ligne

{$ENDIF}
    if TobTemp.Detail.Count > 0 then exit;
    // Modif BTP
{$IFDEF BTP}
    if Ouvrage then
       begin
       TobTemp.putvalue('BLO_QTEDUDETAIL',QteduDetail);
       if TobTemp.GetValue(prefixe+'_COMPOSE') <> '' then
          begin
          if not ChargeOuvrage (TobTemp.GetValue(prefixe+'_COMPOSE'), TOBTemp , InfoNiv+1,TOBTemp.GetValue(prefixe+'_NUMORDRE'),Tresult) then
             begin
             Result := False;
             G_NLIG.Cells[SGN_Comp,ARow] := '';
             TOBTEMP.InitValeurs;
             exit;
             end;
          RenumeroteLesFilles (TOBtemp,InfoNiv);
          TOBTemp.Putvalue ('BLO_DPA',Tresult[0]);
          TOBTemp.Putvalue ('BLO_DPR',Tresult[1]);
          TOBTemp.Putvalue ('BLO_PUHT',Tresult[2]);
          TOBTemp.Putvalue ('BLO_PUHTBASE',Tresult[2]);
          TOBTemp.Putvalue ('BLO_PUHTDEV',pivottodevise(Tresult[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
          TOBTemp.Putvalue ('BLO_PUTTC',Tresult[3]);
          TOBTemp.Putvalue ('BLO_PUTTCDEV',pivottodevise(Tresult[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
          TOBTemp.Putvalue ('BLO_PMAP',Tresult[6]);
          TOBTemp.PutValue('ANCPV', TOBTemp.getValue('BLO_PUHTDEV'));
          TOBTemp.PutValue('ANCPA', TOBTemp.getValue('BLO_DPA'));
          TOBTemp.PutValue('ANCPR', TOBTemp.getValue('BLO_DPR'));
          TOBTemp.Putvalue ('BLO_PMRP',Tresult[7]);
          TOBTEMP.Putvalue ('BLO_QTEDUDETAIL',qtedudetail);
          TOBTemp.putValue('BLO_TPSUNITAIRE',TResult[9]);
          TOBTemp.putValue('GA_HEURE',Tresult[9]);
          TOBTemp.putValue('GA_PAHT',Tresult[0]);
          TOBTemp.putValue('GA_PVHT',Tresult[2]);
          TOBTemp.PutValue('BLO_TPSUNITAIRE',Tresult[9]);
          TOBTemp.PutValue('GA_HEURE',Tresult[9]);
          TOBTemp.PutValue('BLO_MONTANTFG',Arrondi(Tresult[13]*TOBTemp.getValue('BLO_QTEFACT'),V_PGI.OkdecP));
          TOBTemp.PutValue('BLO_MONTANTFC',Arrondi(Tresult[15]*TOBTemp.getValue('BLO_QTEFACT'),V_PGI.OkdecP));
          TOBTemp.PutValue('BLO_MONTANTFR',Arrondi(Tresult[14]*TOBTemp.getValue('BLO_QTEFACT'),V_PGI.OkdecP));
          CalculeLigneAcOuvCumul (TOBTemp);
          CalculMontantHtDevLigOuv (TOBTEMP,DEV);
          GetValoDetail (TOBTEMP);
          if TOBLigPiece.getValue('GLC_NATURETRAVAIL') <> '' then
          begin
            TOBTemp.PutValue(prefixe+'_NATURETRAVAIL',TOBLigPiece.getValue('GLC_NATURETRAVAIL'));
            TOBTemp.PutValue(prefixe+'_FOURNISSEUR',TOBLigPiece.getValue('GL_FOURNISSEUR'));
            TOBTemp.PutValue('GA_FOURNPRINC',TOBLigPiece.getValue('GL_FOURNISSEUR'));
          end;
          if GestionHt then
             G_NLIG.cells[SGN_PVHT,Arow] := StrF00(TOB(G_NLIG.Objects[SGN_NumLig, ARow]).getvalue('BLO_PUHTDEV'),V_PGI.OkdecP)
          else G_NLIG.cells[SGN_PVHT,Arow] := StrF00(TOB(G_NLIG.Objects[SGN_NumLig, ARow]).getvalue('BLO_PUTTCDEV'),V_PGI.OkdecP);
          G_NLIG.InvalidateCell (SGN_PVHT,ARow);
          end;
       end;
{$ENDIF}
    //ChargeNomen(TOBTemp.GetValue(prefixe+'_COMPOSE'), TOBTemp , InfoNiv+1,TOBTemp.GetValue(prefixe+'_NUMORDRE'));
{$IFDEF BTP}
    // Lancement automatique du fichier XLS et Recupération de la valeur
    //if TheMetreShare <> nil then
    //	Val := TheMetreShare.CalculeMetre(TOBTemp, self);
    //else
    //Val := 0;
    Val := AppelMetre(TOBTEMP); //ElipsisMetreClick(TOBTEMP);

    //if Val <> 0 then        //FV1 : 30/09/2016
    //begin
        TOBtemp.putValue('BLO_QTEFACT',Val);
        TOBtemp.PutLigneGrid (G_NLIG, ARow,False, False, LesColNLIG);
        if G_NLIG.col = SGN_Qte then
          begin
           Cell_Text := G_NLIG.Cells[SGN_QTE,ARow];
          end;
    //end;
{$ENDIF}
    AffecteLibelle(TOBTemp, prefixe+'_CODEARTICLE');

//--------------------------------------------------------------------------
//  on rafraichit l'ensemble du treeview
//--------------------------------------------------------------------------
    TV_NLIG.Selected := TTreeNode(G_NLIG.Objects[SGN_Comp, ARow]);
    end;
end;

function TFLigNomen.TraiteQteSais (Acol,Arow : integer) : boolean;
var LastQte : double;
    TOBL : TOB;
begin
  result := true;
  TOBL        := Tob(G_NLIG.objects[SGN_Numlig,Arow]);
  //
  if (Acol = SGN_QTESAIS) then
  begin
    LastQte := TOBL.GetDouble('BLO_QTEFACT');
    TOBL.SetDouble('BLO_QTESAIS', Valeur(G_NLIG.Cells[ACol, ARow]))
  end else if Acol = SGN_PERTE then
  begin
    TOBL.SetDouble('BLO_PERTE', Valeur(G_NLIG.Cells[ACol, ARow]))
  end else if Acol = SGN_RENDEMENT then
  begin
    TOBL.SetDouble('BLO_RENDEMENT', Valeur(G_NLIG.Cells[ACol, ARow]));
  end;
  if POS(TOBL.GetString('BNP_TYPERESSOURCE'),'SAL;INT;ST') > 0 then
  begin
    // prise en compte du rendement
    if TOBL.GetDouble('BLO_RENDEMENT') <> 0 then TOBL.SetDouble('BLO_QTEFACT',TOBL.GetDouble ('BLO_QTESAIS')/TOBL.GetDouble('BLO_RENDEMENT'))
                                            else TOBL.SetDouble('BLO_QTEFACT',TOBL.GetDouble ('BLO_QTESAIS'));
    //
    TOBL.SetDouble('GA_HEURE',TOBL.GetDouble('BLO_QTEFACT'));
  end else
  begin
    // prise en compte du coef de perte
    if TOBL.GetDouble('BLO_PERTE') <> 0 then TOBL.SetDouble('BLO_QTEFACT',TOBL.GetDouble ('BLO_QTESAIS')*TOBL.GetDouble('BLO_PERTE'))
                                        else TOBL.SetDouble('BLO_QTEFACT',TOBL.GetDouble ('BLO_QTESAIS'));
  end;
  TOBL.PutLigneGrid (G_NLIG,Arow,false,false,LesColNLIG);
  Data_Modified := true;
end;


function TFLigNomen.TraiterQuantite(ARow: LongInt) : boolean;
var TOBL        : TOB;
    TOBX        : TOB;
		Qte         : double;
    QteSit      : Double;
    PUHTDev     : Double;
    PrixPourQte : Double;
    MtSituation : Double;
    MtMarche    : Double;
    Pourcent    : Double;
    ZoneCell    : Variant;
begin

  Result    := True;

  ZoneCell  := G_NLIG.Cells[SGN_Comp, ARow];

  Qte       := VALEUR(G_NLIG.Cells[SGN_Qte, ARow]);

  if ZoneCell = '' then exit;

  TOBL        := Tob(G_NLIG.objects[SGN_Numlig,Arow]);
  PUHTDev     := TOBL.GetValue('BLO_PUHTDEV');
  PrixPourQte := TOBL.GEtValue('BLO_PRIXPOURQTE');
  //

  Re_AffecteLibelle(TOBL,prefixe+'_CODEARTICLE',Valeur(ZoneCell));
  TTreeNode(G_NLIG.Objects[SGN_Comp, ARow]).Text:=TOBL.GetValue('LIBELLE');

  TOBX := Tob(G_NLIG.objects[SGN_IdComp,Arow]);

  //{$IFDEF BTP}
  // Modif BTP
  //Re_AffecteLibelle(TOB(G_NLIG.Objects[SGN_NumLig, ARow]),prefixe+'_CODEARTICLE',Valeur(G_NLIG.Cells[SGN_Qte, ARow]));
  //TTreeNode(G_NLIG.Objects[SGN_Comp, ARow]).Text:=TOB(G_NLIG.Objects[SGN_NumLig, ARow]).GetValue('LIBELLE');

  {if Tob(G_NLIG.objects[SGN_IdComp,Arow]).GetValue('GA_TYPEARTICLE')='PRE' then
  begin
    Tob(G_NLIG.objects[SGN_Numlig,Arow]).PutValue('GA_HEURE',Valeur(G_NLIG.Cells[SGN_Qte, ARow]));
  end;
  {$ENDIF}
  {if (Pos(TOBPiece.GetValue('GP_NATUREPIECEG'),'FBT;FBP')>0) then
  begin
    Tob(G_NLIG.objects[SGN_Numlig,Arow]).PutValue('BLO_COEFMARG',0);
  	if (TOBL.FieldExists ('BLF_NATUREPIECEG')) and (TOBL.GetValue ('BLF_NATUREPIECEG')<>'') then
    begin
      Qte := Valeur(G_NLIG.Cells[SGN_Qte, ARow]);
    	TOBL.Setdouble('BLF_QTESITUATION',Qte);
    	TOBL.SetDouble('BLF_MTSITUATION',Arrondi(TOBL.GetValue('BLF_QTESITUATION')*TOBL.GetValue('BLO_PUHTDEV')/TOBL.GEtValue('BLO_PRIXPOURQTE'),DEV.decimale));
      if TOBL.GetValue('BLF_MTMARCHE') > 0 then
      begin
      	Pourcent := Arrondi(TOBL.GetValue('BLF_MTSITUATION')/TOBL.GetValue('BLF_MTMARCHE')*100,2);
      end else
      begin
      	Pourcent := Arrondi(100,2);
      end;
    	TOBL.PutValue('BLF_POURCENTAVANC',Pourcent);
    	CalculMontantHtDevLigOuv (TOBL,DEV);
    end;
  end;
  Tob(G_NLIG.objects[SGN_Numlig,Arow]).PutValue('GA_HEURE',Valeur(G_NLIG.Cells[SGN_Qte, ARow]));}

  if TOBX.GetValue('GA_TYPEARTICLE')='PRE' then TOBL.PutValue('GA_HEURE',Qte);

  if (Pos(TOBPiece.GetValue('GP_NATUREPIECEG'),'FBT;FBP')>0) then
  begin
    TOBL.PutValue('BLO_COEFMARG',0);
  	if (TOBL.FieldExists ('BLF_NATUREPIECEG')) and (TOBL.GetValue ('BLF_NATUREPIECEG')<>'') then
    begin
    	TOBL.Setdouble('BLF_QTESITUATION',Qte);
      //
    	TOBL.SetDouble('BLF_MTSITUATION',Arrondi(Qte*PuHTDev/PrixPourQte,DEV.decimale));
      //
      MtSituation := TOBL.GetValue('BLF_MTSITUATION');
      MtMarche    := TOBL.GetValue('BLF_MTMARCHE');
      //
      if MtMarche > 0 then
      	Pourcent := Arrondi(MtSituation/MtMarche*100,2)
      else
      	Pourcent := Arrondi(100,2);
      //
      TOBL.PutValue('BLF_POURCENTAVANC',Pourcent);
      //
    	CalculMontantHtDevLigOuv (TOBL,DEV);
    end;
  end;
  TOBL.SetDouble('BLO_QTEFACT',Qte);
  TOBL.SetDouble('BLO_QTESAIS',CalculeQteSais (TOBL));
  TOBL.PutValue('GA_HEURE',Qte);

end;

procedure TFLigNomen.FormateZoneSaisie(ACol,ARow: LongInt);
var St,StC : String ;
    i_sep : integer;
begin

  St:=G_NLIG.Cells[ACol,ARow] ;
  StC:=St ;

  if (ACol=SGN_Comp) then
    StC:=uppercase(Trim(St))
  else if (ACol=SGN_Qte)  or (Acol = SGN_QTESAIS) then
        BEGIN
        i_sep := pos('.',St);
        if i_sep > 0 then
           begin
           st := copy (st,1,i_sep-1)+','+copy(st,i_sep+1,255);
           end;
        StC:=StrS(Valeur(St),V_PGI.OkDecQ);
        END
{$IFDEF BTP}
        else if (Acol = SGN_PVht) then
        BEGIN
        i_sep := pos('.',St);
        if i_sep > 0 then
           begin
           st := copy (st,1,i_sep-1)+','+copy(st,i_sep+1,255);
           end;
    stc := strs (Valeur(st),V_PGI.OkDecP);
  END
  else if (Acol = SGN_UV) then
    StC:=uppercase(Trim(St))
{$ENDIF}
  else
    G_NLIG.Cells[ACol,ARow]:=StC ;
end;

function chargeArticle (TheArticle : string ; TOBA : TOB) : boolean;
var QQ : Tquery;
begin
	result := false;
  QQ := OPenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
              		'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
              		'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE GA_ARTICLE ="'+TheArticle+'"',True,-1, '', True);
//	QQ := OPenSql ('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+TheArticle+'"',true);
  if not QQ.eof then
  begin
  	TOBA.selectdb ('',QQ);
    InitChampsSupArticle (TOBA);
    result := true;
  end;
  ferme (QQ);
end;

function TFLigNomen.Recherche_Art(ARow : LongInt) : boolean;
var RechArt : T_RechArt ;
    OkArt,Trouvok   : Boolean ;
    st1 : string;
    TOBArt,TobTemp : TOB;
    TOBPiece : TOB;
    Libelle, sWhere, stChamps : string;
    CodeArt : string;
    TypeArticle : string;
begin
    Libelle := '';
    OkArt:=False ;
    TOBArt := TOB.Create('ARTICLE',nil,-1);
    CodeArt := G_NLIG.Cells[SGN_COMP,Arow];
//-----------------------------------------------------------------
//  Recherche du code article saisi
//-----------------------------------------------------------------
    RechArt := TrouverArticle (CodeArt, TOBArt);
    TypeArticle := TOBArt.GetValue('GA_TYPEARTICLE');
    case RechArt of
            traOk : begin
//-----------------------------------------------------------------
//  Code article trouve, on regarde s'il a des nomenclatures associées
//  si oui et plusieurs, ouverture d'une fenetre de selection
//  si oui et une seule, on recupere directement son code
//-----------------------------------------------------------------
                    if ((TypeArticle = 'NOM') and (not Ouvrage)) or
                       (((TypeArticle= 'OUV') or (TypeArticle= 'ARP')) and (Ouvrage)) then
                       begin
                       Trouvok := VerifNomen(TobArt,ARow,st1,Libelle);
                       if st1 <> '' then
                          begin
//-----------------------------------------------------------------
//  on enregistre le code nomenclature dans la colonne GLN_COMPOSE
//-----------------------------------------------------------------
                          G_NLIG.Cells[SGN_Compose, ARow] := st1;
                          TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue(prefixe+'_COMPOSE',st1);
                          OkArt:=True ;
                          end else OkArt:=false;
                       end else OkArt := true;
                    end;
         traAucun : begin
//-----------------------------------------------------------------
// Recherche sur code via LookUp ou Recherche avancée
//-----------------------------------------------------------------
                    ART.Text:=G_NLIG.Cells[SGN_Comp, ARow];
{$IFDEF BTP}
                    TOBPiece := TOBLigPiece.Parent ;
                    sWhere:=FabricWhereNatArt(TOBPiece.getValue('GP_NATUREPIECEG'),'','') ;
                    StChamps:='GA_CODEARTICLE='+ART.Text+';XX_WHERE=AND '+sWhere ;
                    DispatchRecherche(ART, 1, '', stChamps , '');
{$ELSE}
                    DispatchRecherche(ART, 1, '', 'GA_CODEARTICLE=' + ART.Text, '');
{$ENDIF}
                    if ART.Text <>'' then
                        begin
                        	if not chargeArticle(Art.text,TOBARt) then begin result := false; exit; end;
(*
                        TOBArt.PutValue('GA_ARTICLE',Art.Text);
                        TOBArt.LoadDB();
*)
                        trouvOk := VerifNomen(TobArt,ARow,st1,Libelle);
                        if st1 <> '' then
                            begin
//-----------------------------------------------------------------
//  on enregistre le code nomenclature dans la colonne GLN_COMPOSE
//-----------------------------------------------------------------
                            G_NLIG.Cells[SGN_Compose, ARow] := st1;
                            TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue(prefixe+'_COMPOSE',st1);
                            OkArt:=True ;
                            end else if TOBArt.GetValue('GA_STATUTART')='GEN' then
                            begin
                            if ChoisirDimension (TOBArt.GetValue ('GA_ARTICLE'), TOBArt) then
                            OkArt := True;
                            end else if TrouvOk then OkArt:=True ;
                        end;
                    end ;
        traGrille : begin
//-----------------------------------------------------------------
// Forcement objet dimension avec saisie obligatoire
//if GetArticleLookUp (GF_ARTICLE, 'GA_STATUTART = "DIM"') then
//-----------------------------------------------------------------
                    if ChoisirDimension (TOBArt.GetValue ('GA_ARTICLE'), TOBArt) then
                        OkArt := True;
{$IFDEF BTP}
                    if (TypeArticle= 'ARP') then
                       begin
                       Trouvok := VerifNomen(TobArt,ARow,st1,Libelle);
                       if st1 <> '' then
                          begin
//-----------------------------------------------------------------
//  on enregistre le code nomenclature dans la colonne GLN_COMPOSE
//-----------------------------------------------------------------
                          G_NLIG.Cells[SGN_Compose, ARow] := st1;
                          TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue(prefixe+'_COMPOSE',st1);
                          OkArt:=True ;
                          end else OkArt:=false;
                       end else OkArt := true;
{$ENDIF}
                    end;
         end ; // Case

    if (Tobart.GetValue('GA_FERME') = 'X') then
    begin
      MsgBox.Execute(12, Caption, '');
      OkArt := False;
    end;

    if (OkArt) then
        begin
        if (ARow = 1) and (TOBLig.Detail.Count = 0) then
            begin
            G_NLIG.Objects[SGN_NumLig, ARow] := TOB.Create ('LIGNOMEN', TOBLig, ARow-1) ;
            TOB(G_NLIG.Objects[SGN_NumLig, ARow]).AddChampSup('LIBELLE', True);
            TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue(prefixe+'_NUMORDRE',ARow-1);
            end;
//--------------------------------------------------------------------------
//  on va charger la fiche article dans TobArticles
//--------------------------------------------------------------------------
        G_NLIG.Objects[SGN_IdComp, ARow] := TOB.Create('ARTICLE',nil,ARow);
        Tob(G_NLIG.Objects[SGN_IdComp, ARow]).Dupliquer(TOBArt, True, True);

        TobTemp := TOB(G_NLIG.Objects[SGN_NumLig, ARow]);
        // Modif BTP
        TobTemp.PutValue(prefixe+'_CODEARTICLE', TOBArt.GetValue('GA_CODEARTICLE'));
        G_NLIG.Cells[SGN_Comp, ARow] := TOBArt.GetValue('GA_CODEARTICLE');
        TobTemp.PutValue(prefixe+'_ARTICLE', TOBArt.GetValue('GA_ARTICLE'));
        G_NLIG.Cells[SGN_IdComp, ARow] := TOBArt.GetValue('GA_ARTICLE');
        FormateZoneSaisie(SGN_Comp, ARow);
        if Libelle = '' then
           begin
           TobTemp.PutValue(prefixe+'_LIBELLE', TOBArt.GetValue('GA_LIBELLE'));
           G_NLIG.Cells[SGN_Lib, ARow] := TOBArt.GetValue('GA_LIBELLE');
           end else
           BEGIN
           TobTemp.PutValue(prefixe+'_LIBELLE', Libelle);
           G_NLIG.Cells[SGN_Lib, ARow] := Libelle;
           END;
        TobTemp.PutValue(prefixe+'_ORDRECOMPO',InfoOrdre);
{$IFDEF BTP}
        if Ouvrage then RenseigneTOBOuv(TOB(G_NLIG.Objects[SGN_NumLig, ARow]),TobArt,Libelle)
        else
{$ENDIF}
        RenseigneTOBLN(TOB(G_NLIG.Objects[SGN_NumLig, ARow]),TobArt,TOBInfo);
        TobTemp.PutValue(prefixe+'_NIVEAU',InfoNiv);
        if Ouvrage then
           begin
           if gestionHt then
              G_NLIG.cells[SGN_PVHT,Arow] := StrF00(TOB(G_NLIG.Objects[SGN_NumLig, ARow]).getvalue('BLO_PUHTDEV'),V_PGI.OkdecP)
           else G_NLIG.cells[SGN_PVHT,Arow] := StrF00(TOB(G_NLIG.Objects[SGN_NumLig, ARow]).getvalue('BLO_PUTTCDEV'),V_PGI.OkdecP);
           G_NLIG.cells[SGN_UV,Arow] := TOB(G_NLIG.Objects[SGN_NumLig, ARow]).getvalue('BLO_QUALIFQTEVTE');
           G_NLIG.InvalidateCell (SGN_PVHT,ARow);
           G_NLIG.InvalidateCell (SGN_UV,ARow);
           end;
        end else
        begin
        G_NLIG.Cells[SGN_Comp, ARow] := '';
        G_NLIG.Col := SGN_Comp;
        end;
TOBArt.free;
Result := OkArt;
end;

////////////////////////////////////////////////////////////////////////////////
//*************************** Gestion des lignes *****************************//
////////////////////////////////////////////////////////////////////////////////

function TFLigNomen.ValideLigne(Ou : LongInt) : boolean;
var i_ind1 : integer ;
{$IFDEF BTP}
    valeurpv,valeuranc,ValeurPr,Qte,MtLoc,QteSit,MtSit : double;
    TOBPIece,TOBL,TOBART : TOB;
    Zvaleur : T_valeurs;
{$ENDIF}
begin
  //  Validation des différents champs de la ligne
  Result := True;
  if TOB(G_NLIG.Objects[SGN_NumLig, Ou]).getValue('BLO_CODEARTICLE')='' then exit; // pas de traitement sur ligne vide
  if (G_NLIG.col=SGN_Comp) and ( not TraiterComposant(Ou)) then
    begin
    Result := False;
    Exit;
    end;
  if not TraiterQuantite(Ou) then
    begin
    Result := False;
    Exit;
    end;
  bNewLigne.enabled := true;
  TOBL := TOB(G_NLIG.Objects[SGN_NumLig, Ou]);
  TOBArt := TOB(G_NLIG.objects[SGN_idComp,Ou]);

  //  Formattage et mise en place dans la Tob associée
  for i_ind1 := 0 to G_NLIG.ColCount - 1 do
    begin
    FormateZoneSaisie (i_ind1, Ou);
    //Modif BTP
    if i_ind1   = SGN_Comp                    then TOBL.PutValue(prefixe+'_CODEARTICLE', G_NLIG.Cells[i_ind1, Ou]);
    if i_ind1   = SGN_Lib                     then TOBL.PutValue(prefixe+'_LIBELLE', G_NLIG.Cells[i_ind1, Ou]);
    if (i_ind1  = SGN_Qte) and (not Ouvrage)  then TOBL.PutValue('GLN_QTE', G_NLIG.Cells[i_ind1, Ou]);
    if (i_ind1  = SGN_Qte) and (Ouvrage)      then TOBL.PutValue('BLO_QTEFACT', G_NLIG.Cells[i_ind1, Ou]);

    {$IFDEF BTP}
    if (i_ind1  = SGN_PvHt) and (ouvrage) then
    begin
      if Reajuste then
      begin
        TOBPiece := TOBLigPiece.Parent ;
// ----------------        
//        Pk ??? .. je laisse en commentaire pour le cas ou...
//        Valeurpv  := TOBL.GetDouble('BLO_PUHT');//valeur(G_NLIG.Cells[i_ind1, Ou]);
// -----------
        Valeurpv  := valeur(G_NLIG.Cells[i_ind1, Ou]);
        Valeuranc := TOBL.GetValue('ANCPV');
        if (valeurpv <> valeuranc) then
        begin
          if (valeurpv <> valeuranc) and (TOBL.Detail.Count > 0) then
            //FV1 : 22/01/2015 - FS#1359 - EMC - message d'erreur à la validation d'un devis
            ReajusteMontantOuvrage (TOBART,TOBpiece,TOBLigPiece,TOBL,valeuranc,ValeurPr,valeurpv,DEV,GestionHt);
          if GestionHt then
          BEGIN
            TOBL.PutValue('BLO_PUHTDEV', valeurpv);
            TOBL.PutValue('BLO_PUHT', devisetopivotEx(valeurpv,DEV.taux,DEV.quotite,V_PGI.OkdecP));
            CalculeLigneHTOuv  (TOBL,TOBPiece,DEV);
          END ELSE
          BEGIN
            TOBL.PutValue('BLO_PUTTCDEV', valeurpv);
            CalculeLigneTTCOuv  (TOBL,TOBPiece,DEV);
          END;
        end;
        TOBL.PutValue('ANCPV', valeurpv);
      end;
    end;
    {$ENDIF}
    end;
  {$IFDEF BTP}
  //
  if TOBL.detail.count > 0 then
  begin
	Qte := TOBL.GetValue('BLO_QTEFACT');
  GetValoDetail (TOBL);
  CalculOuvFromDetail (TOBL,DEV,Zvaleur);
  TOBL.Putvalue ('BLO_DPA',Zvaleur[0]);
  TOBL.Putvalue ('BLO_DPR',Zvaleur[1]);
    //  if Not ((TOBL.GetValue('BLO_TYPEARTICLE')= 'ARP') and (TOBL.detail.count > 1)) then TOBL.putvalue ('GA_PVHT',Zvaleur[2]);
  TOBL.Putvalue ('BLO_PUHT',Zvaleur[2]);
  TOBL.Putvalue ('BLO_PUHTBASE',Zvaleur[2]);
  TOBL.Putvalue ('BLO_PUHTDEV',pivottodevise(Zvaleur[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
  TOBL.Putvalue ('BLO_PUTTC',Zvaleur[3]);
  TOBL.Putvalue ('BLO_PUTTCDEV',pivottodevise(Zvaleur[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
  if (TOBL.GetValue('BLO_TYPEARTICLE')='ARP') and (TOBL.Detail.count > 0) then
  begin
		TOBL.putvalue ('GA_PVHT',TOBL.detail[0].GetValue('BLO_PUHT'));
  end else
  begin
		TOBL.putvalue ('GA_PVHT',Zvaleur[2]);
  end;
  TOBL.PutValue('ANCPV', TOBL.GetValue('BLO_PUHTDEV'));
  TOBL.PutValue('ANCPA', TOBL.GetValue('BLO_DPA'));
  TOBL.PutValue('ANCPR', TOBL.GetValue('BLO_DPR'));
  TOBL.Putvalue ('BLO_PMAP',Zvaleur[6]);
  TOBL.Putvalue ('BLO_PMRP',Zvaleur[7]);
  //
  TOBL.PutValue('BLO_TPSUNITAIRE',Zvaleur[9]);
  TOBL.PutValue('GA_HEURE',Zvaleur[9]);
    StockeLesTypes (TOBL,zValeur);
  //
  TOBL.PutValue('BLO_MONTANTFG',Arrondi(Zvaleur[13]*Qte,V_PGI.OkdecP));
  TOBL.PutValue('BLO_MONTANTFC',Arrondi(Zvaleur[15]*Qte,V_PGI.OkdecP));
  TOBL.PutValue('BLO_MONTANTFR',Arrondi(Zvaleur[14]*Qte,V_PGI.OkdecP));
  CalculeLigneAcOuvCumul (TOBL);
  end else
  begin
  TOBArt := TOB(G_NLIG.objects[SGN_idComp,Ou]);
	CalculeLigneAcOuv (TOBL,DEV,TOBLigPiece.GetValue('GLC_NONAPPLICFRAIS')='X',TOBART);
    StockeMontantTypeSurLigne (TOBL);
  end;
  //
  CalculMontantHtDevLigOuv (TOBL,DEV);
  IF (TOBLigPiece.GetString('GL_PIECEPRECEDENTE')='') and
  (TOBLigPiece.FieldExists('BLF_QTESITUATION')) and
  (InfoNiv=1) then
  begin
    QteSit := arrondi(TOBL.getValue('BLO_QTEFACT')*TOBLigPiece.getvalue('GL_QTEFACT') /TOBL.getValue('BLO_QTEDUDETAIL'),V_PGI.okdecQ);
    MtLoc := Arrondi(TOBL.getValue('BLO_QTEFACT') / TOBL.getValue('BLO_QTEDUDETAIL') * TOBL.getValue('BLO_PUHTDEV'),V_PGI.okdecP);
    MtSit := arrondi(TOBLigPiece.getvalue('GL_QTEFACT') * MtLoc,DEV.Decimale);
    TOBL.PutValue('BLF_QTESITUATION', QteSit); { NEWPIECE }
    TOBL.PutValue('BLF_MTSITUATION',MtSit);
    TOBL.PutValue('BLF_POURCENTAVANC',100);
  end;
  {$ENDIF}
  TOBL.PutValue('LIBELLE', G_NLIG.Cells[SGN_Comp, Ou] +' (' + G_NLIG.Cells[SGN_Qte, Ou] + ')');
  TOBL.PutValue(prefixe+'_COMPOSE',COMPOSE.Text);
  {$IFDEF BTP}
  TOBL.PutLigneGrid (G_NLIG,Ou,false,false,LesColNLIG);
  {$ENDIF}
end;

procedure TFLigNomen.InsertLigne (ARow : Longint; Ajout : boolean) ;
var
    TobTemp : TOB;
begin
if Action=taConsult then Exit ;
// Modif BTP
if TypeArp then Exit;
// --
if ARow < 1 then Exit ;
if ARow = G_NLIG.RowCount - 1 then G_NLIG.RowCount := G_NLIG.RowCount + 50;
if (not Ajout) and (ARow > TOBLig.Detail.Count) then Exit;
if (ARow = TOBLig.Detail.Count) and (G_NLIG.Cells[SGN_Comp,ARow] = '') then exit;
// Modif BTP
if Btypearticle.Visible then BtypeArticle.Visible := false;
// -----
G_NLIG.CacheEdit; G_NLIG.SynEnabled := False;
if (ARow <= TOBLig.Detail.Count) then G_NLIG.InsertRow (ARow);
G_NLIG.Row := ARow;
if ARow > 1 then
    G_NLIG.Cells[0, ARow] := IntToStr(StrToInt(G_NLIG.Cells[0, ARow - 1]) + 1)
    else
    G_NLIG.Cells[0, ARow] := IntToStr(1);
InitialiseLigne (ARow) ;
//--------------------------------------------------------------------------
//  on ajoute une tob fille associée à cette ligne
//--------------------------------------------------------------------------
AjouteTobLigne(ARow);
TOBTemp := TOB(G_NLIG.Objects[SGN_NumLig, ARow]);
AffecteLibelle(TOBTemp, prefixe+'_CODEARTICLE');
G_NLIG.MontreEdit; G_NLIG.SynEnabled := True;
G_NLIG.ElipsisButton := true;
G_NLIG.col := SGN_COMP;
Renumerote;
AffectePied(ARow);
// Modif BTP
// ben ....du fait qu'on est sur une nouvelle ligne le Cell_Text doit etre reinitialise
Cell_Text := G_NLIG.Cells[G_NLIG.Col, G_NLIG.Row];
end;

procedure TFLigNomen.VideLaLigne (ARow : integer);
begin
TOB(G_NLIG.Objects[SGN_NumLig, ARow]).Free;
G_NLIG.Objects[SGN_NumLig, ARow] := nil;
G_NLIG.Rows[ARow].Clear;
AfficheTVNLIG;
end;

//=============================================================================
//  Initialisation des colonnes de la ligne
//=============================================================================
procedure TFLigNomen.InitialiseLigne (ARow : Longint) ;
var
    i_ind : integer ;
begin
for i_ind := 1 to G_NLIG.ColCount-1 do
    if (i_ind = SGN_Qte)   then G_NLIG.Cells [i_ind, ARow] := '1,00' else
        G_NLIG.Cells [i_ind, ARow]:='' ;
end;

//=============================================================================
//  Suppression d'une ligne et de la tob associée et renumérotation
//=============================================================================
procedure TFLigNomen.SupprimeLigne (ARow : Longint) ;
var Acol : integer;
    Cancel: Boolean;
{$IFDEF BTP}
    TOBLN : TOB;
    TobDelMetre : TOB;
{$ENDIF}
begin
if Action=taConsult then Exit ;
// Modif BTP
if TypeArp then Exit;
// --
G_NLIG.OnRowMoved := nil;
if ARow < 1 then Exit ;
if (ARow > TOBLig.Detail.Count) then Exit;
if (ARow = TOBLig.Detail.Count) and (G_NLIG.Cells[SGN_Comp,ARow] = '') then exit;
// Modif BTP
{$IFDEF BTP}
TOBLN := TOB(G_NLIG.Objects[SGN_NumLig, Arow]);
if (TOBLN.FieldExists ('INDICELIENDEVCHA')) and (TOBLN.GetValue('INDICELIENDEVCHA') > 0) then
begin
  if (TobLN.GetValue('BLO_NATUREPIECEG')=GetParamSoc ('SO_AFNATAFFAIRE')) then pgiBox ('Impossible la prévision de chantier a été générée',Caption)
                                                                          else pgiBox ('Impossible Il existe un lien avec un devis',Caption);
  exit;
end;

{$ENDIF}

if Btypearticle.Visible then BtypeArticle.Visible := false;
// -----
G_NLIG.CacheEdit; G_NLIG.SynEnabled := False;
if G_NLIG.Objects[SGN_NumLig, ARow] <> nil then TOB(G_NLIG.Objects[SGN_NumLig, ARow]).Free;
G_NLIG.DeleteRow (ARow);
if G_NLIG.RowCount < NbRowsInit then G_NLIG.RowCount := NbRowsInit;
G_NLIG.MontreEdit; G_NLIG.SynEnabled := True;
Renumerote;
ACol := G_NLIG.col; Cancel := false;
//if (G_NLIG.Row = 1) then //and (TOBLig.Detail.Count = 0) then
//     G_NLIGRowEnter(Sender, 1, Cancel, False)
//else begin
//     G_NLIG.Row := ARow-1;
//     end;
G_NLIGRowEnter(Self, ARow, Cancel, False);
G_NLIGCellEnter (Self,Acol,Arow,cancel);
G_NLIG.row := Arow;
G_NLIG.col := ACol;
AffectePied(ARow);
{$IFDEF BTP}
if Ouvrage then
   begin
   CalculeOuvrageLoc;
   AfficheTotalisation;
   end;
{$ENDIF}
AfficheTVNLIG;
G_NLIG.OnRowMoved := G_NLIGRowMoved;

end;

//=============================================================================
//  Renumérotation des lignes à partir de la ligne donnée
//=============================================================================
procedure TFLigNomen.Renumerote;
var  i_ind1,i_ind2 : integer ;
     TobLN,TobLND : TOB;
begin

for i_ind1 := 1 to G_NLIG.RowCount-1 do
    begin
    TOBLN := TOB(G_NLIG.Objects[SGN_NumLig, i_ind1]);
    if TOBLN = nil then exit;
    G_NLIG.Cells [0, i_ind1] := IntToStr(i_ind1);
    SetNumerotation (TOBLN,TOBLig,InfoNiv,i_ind1);
    if TOBLN.detail.count > 0 then
    begin
      RenumeroteLesFilles(TOBLN,InfoNiv);
    end;
    // Modif BTP
    TOBLN.PutValue(prefixe+'_NUMORDRE',i_ind1);
    if TOBLN.GetValue('SOUSNOMEN')='X' then
      begin
       for i_ind2 := 0 to TOBLN.Detail.Count-1 do
           begin
           TOBLND := TOBLN.Detail[i_ind2];
           // Modif BTP
           TOBLND.PutValue(prefixe+'_ORDRECOMPO',i_ind1);
          end;
           end;
    end;
end;

//=============================================================================
//  Recherche d'une ligne de la grille en fonction d'une valeur chaine de caracteres
//=============================================================================
function TFLigNomen.ChercheLigne (tn1 : TTreeNode) : integer;
var i_ind1 : integer ;
begin
Result := 0;
for i_ind1 := 0 to G_NLIG.RowCount - 1 do
    begin
    if tn1=G_NLIG.Objects[SGN_Comp, i_ind1] then
        begin
        Result := i_ind1;
        Exit;
        end;
    end;
end;

////////////////////////////////////////////////////////////////////////////////
//****************************** Gestion du pied *****************************//
////////////////////////////////////////////////////////////////////////////////

procedure TFLigNomen.AffectePied (ARow : Longint) ;
begin
TCODEDIM.Caption := '';
if G_NLIG.Cells[SGN_Comp, ARow] = '' then
   begin
   bNomenc.Enabled:=false;
   exit;
   end;
if TOB(G_NLIG.Objects[SGN_IdComp, ARow])=nil then exit;
if TOB(G_NLIG.Objects[SGN_IdComp, ARow]).GetValue('GA_STATUTART') = 'DIM' then
AffiGrillCodedimension(TOB(G_NLIG.Objects[SGN_IdComp, ARow]));
if TOB(G_NLIG.Objects[SGN_NumLig, ARow]).GetValue('SOUSNOMEN')='X' then bNomenc.Enabled:=true
else bNomenc.Enabled:=false ;
end;

Procedure TFLigNomen.AffiGrillCodedimension(TOBArt : TOB);
var i_indDim : integer;
    GrilleDim,CodeDim,LibDim,StDim : string;
begin
StDim:='';
for i_indDim := 1 to MaxDimension do
    begin
    GrilleDim := TOBArt.GetValue ('GA_GRILLEDIM' + IntToStr (i_indDim));
    CodeDim := TOBArt.GetValue ('GA_CODEDIM' + IntToStr (i_indDim));
    LibDim := GCGetCodeDim (GrilleDim, CodeDim, i_indDim);
    if LibDim <> '' then
       if StDim='' then StDim:=StDim + LibDim
       else StDim := StDim + ' - ' + LibDim;
    end;
TCODEDIM.Caption := 'Dimension(s) : ' + StDim;
end;

////////////////////////////////////////////////////////////////////////////////
//*************************** Chargement des tobs ****************************//
////////////////////////////////////////////////////////////////////////////////

procedure TFLigNomen.ChargeTobArticles;
var i_ind1 : integer;
    st1 : string;
    QQ : TQuery;
    TOBArt : TOB;
begin
  for i_ind1 := 1 to G_NLIG.RowCount - 1 do
  begin
    //    if G_NLIG.Cells[SGN_Comp, i_ind1] = '' then Exit;
    if G_NLIG.Cells[SGN_Comp, i_ind1] = '' then continue;
    G_NLIG.Objects[SGN_IdComp, i_ind1] := TOB.Create('ARTICLE', nil, -1);
    st1 := G_NLIG.Cells[SGN_IdComp, i_ind1];
    TOBArt := TOB(G_NLIG.Objects[SGN_IdComp, i_ind1]);
    QQ := OPenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                  'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                  'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE GA_ARTICLE="'+St1+'"',True,-1, '', True);
    TobArt.SelectDB('',QQ);
    Ferme (QQ);
    InitChampsSupArticle (TOBART);
  end;
end;

procedure TFLigNomen.CompleteTobArticles(TobLN : TOB);
var i_ind : integer;
    stArticle : string;
begin
for i_ind:=0 to TobLN.Detail.Count-1 do
    begin
    stArticle:=TobLN.Detail[i_ind].GetValue(prefixe+'_ARTICLE');
    RechercheArticle(TobArticles,stArticle);
    CompleteTobArticles(TobLN.Detail[i_ind]);
    end;
end;

procedure TFLigNomen.RechercheArticle(TobGen : TOB ; stArticle : string);
var i_ind : integer;
    stDepot : string;
    TobArtD, TobDispo, TobLot : TOB;
    QLot,QDispo : TQuery;
    QQ : TQuery;
begin
stDepot:='';
TobArtD:=TobGen.FindFirst(['GA_ARTICLE'],[stArticle],false);
if TobArtD = nil then
   begin
   TobArtD:=TOB.Create('ARTICLE',TobGen,-1);
   QQ := OPenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
              		'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
              		'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE GA_ARTICLE="'+stArticle+'"',True,-1, '', True);
   TobArtD.SelectDB('',QQ);
   InitChampsSupArticle (TOBARTD);

   Ferme (QQ);
   TobArtD.AddChampSup('UTILISE',false);
   QDispo:=OpenSQL('SELECT * FROM DISPO WHERE GQ_ARTICLE="'+stArticle
         +'" AND GQ_CLOTURE="-"',true,-1, '', True);
   if not QDispo.Eof then
      begin
      TobDispo := TobArtD;
      TobDispo.LoadDetailDB('DISPO','','',QDispo,false);
      if TobArtD.GetValue('GA_LOT')='X' then
          begin
          for i_ind:=0 to TobDispo.Detail.Count-1 do
              begin
              stDepot:=TobDispo.Detail[i_ind].GetValue('GQ_DEPOT');
                  QLot:=OpenSQL('SELECT * FROM DISPOLOT WHERE GQL_ARTICLE="'+stArticle
                  +'" AND GQL_DEPOT="'+stDepot+'"',true,-1, '', True);
              if not QLot.Eof then
                 begin
                 TobLot:=TobDispo.Detail[i_ind];
                 TobLot.LoadDetailDB('DISPOLOT','','',QLot,false);
                 end;
              Ferme(QLot);
              end;
          end;
      end;
      Ferme(QDispo);
   end;
end;


procedure TFLigNomen.AjouteTobLigne(ARow : integer);
begin
if Action = taConsult then exit;
if G_NLIG.Objects[SGN_NumLig, ARow] = nil then
   begin
{$IFDEF BTP}
   if Ouvrage then
      begin
      G_NLIG.Objects[SGN_NumLig, ARow] := TOB.Create ('LIGNEOUV', TOBLig, ARow-1);
      InsertionChampSupOuv (TOB(G_NLIG.Objects[SGN_NumLig, ARow]),false);
      TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue(prefixe+'_NUMORDRE',StrToInt(G_NLIG.Cells[0, ARow]));
      //
      SetNumerotation(TOB(G_NLIG.Objects[SGN_NumLig, ARow]),TOBLig,InfoNiv,Arow);
      TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
      //
      TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue(prefixe+'_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));
      if TOBLigPiece.getValue('GLC_NATURETRAVAIL') <> '' then
      begin
      	TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue(prefixe+'_NATURETRAVAIL',TOBLigPiece.getValue('GLC_NATURETRAVAIL'));
      	TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue(prefixe+'_FOURNISSEUR',TOBLigPiece.getValue('GL_FOURNISSEUR'));
      end;
      TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue(prefixe+'_NATUREPIECEG',TOBLigPiece.getValue('GL_NATUREPIECEG'));
      TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue(prefixe+'_SOUCHE',TOBLigPiece.getValue('GL_SOUCHE'));
      TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue(prefixe+'_NUMERO',TOBLigPiece.getValue('GL_NUMERO'));
      TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue(prefixe+'_INDICEG',TOBLigPiece.getValue('GL_INDICEG'));
      TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue(prefixe+'_DOMAINE',TOBLigPiece.getValue('GL_DOMAINE'));
      //
      end else
      begin
{$ENDIF}
      G_NLIG.Objects[SGN_NumLig, ARow] := TOB.Create ('LIGNENOMEN', TOBLig, ARow-1) ;
      TOB(G_NLIG.Objects[SGN_NumLig, ARow]).AddChampSup('SOUSNOMEN',false);
      TOB(G_NLIG.Objects[SGN_NumLig, ARow]).AddChampSup('LIBELLE', False);
{$IFDEF BTP}
      end;
{$ENDIF}
   end;
end;

procedure TFLigNomen.RenseigneTOBLN(TobLN, TobArticle,TOBInfo : TOB);
begin
  If TOBInfo <> nil then
  begin
    // Modif BTP
    TobLN.PutValue(prefixe+'_NATUREPIECEG', TOBInfo.GetValue(prefixe+'_NATUREPIECEG'));
    TobLN.PutValue(prefixe+'_SOUCHE',       TOBInfo.GetValue(prefixe+'_SOUCHE'));
    TobLN.PutValue(prefixe+'_NUMERO',       TOBInfo.GetValue(prefixe+'_NUMERO'));
    TobLN.PutValue(prefixe+'_INDICEG',      TOBInfo.GetValue(prefixe+'_INDICEG'));
    TobLN.PutValue(prefixe+'_NUMLIGNE',     TOBInfo.GetValue(prefixe+'_NUMLIGNE'));
  end;

  if TobArticle <> nil then
  begin
    // Modif BTP
    TobLN.PutValue(prefixe+'_TENUESTOCK',   TobArticle.GetValue('GA_TENUESTOCK'));
    //TobLN.PutValue('GLN_QUALIFQTEACH',TobArticle.GetValue('GA_QUALIFQTEACH'));
    TobLN.PutValue(prefixe+'_QUALIFQTEVTE', TobArticle.GetValue('GA_QUALIFUNITEVTE'));
    TobLN.PutValue(prefixe+'_QUALIFQTESTO', TobArticle.GetValue('GA_QUALIFUNITESTO'));
    TobLN.PutValue(prefixe+'_DPA',          TobArticle.GetValue('GA_DPA'));
    TobLN.PutValue(prefixe+'_DPR',          TobArticle.GetValue('GA_DPR'));
    TobLN.PutValue(prefixe+'_PMAP',         TobArticle.GetValue('GA_PMAP'));
    TobLN.PutValue(prefixe+'_PMRP',         TobArticle.GetValue('GA_PMRP'));
    {$IFDEF BTP}
    TobLN.PutValue('BLO_LIBCOMPL',          TobArticle.GetValue('GA_LIBCOMPL'));
    TobLN.PutValue('BNP_TYPERESSOURCE',     TobArticle.GetValue('BNP_TYPERESSOURCE'));
    {$ENDIF}
  end;

end;

{$IFDEF BTP}
procedure TFLigNomen.RenseigneTOBOuv(TOBOuv, TobArticle : TOB; libelle : string;AvecPv:boolean);
var TOBPiece : TOB;
begin
  TOBPiece := TOBLigPiece.parent;

  If TOBInfo <> nil then
  begin
    // Modif BTP
    TOBOuv.PutValue(prefixe+'_NATUREPIECEG',TOBInfo.GetValue(prefixe+'_NATUREPIECEG'));
    TOBOuv.PutValue(prefixe+'_SOUCHE',      TOBInfo.GetValue(prefixe+'_SOUCHE'));
    TOBOuv.PutValue(prefixe+'_NUMERO',      TOBInfo.GetValue(prefixe+'_NUMERO'));
    TOBOuv.PutValue(prefixe+'_INDICEG',     TOBInfo.GetValue(prefixe+'_INDICEG'));
    TOBOuv.PutValue(prefixe+'_NUMLIGNE',    TOBInfo.GetValue(prefixe+'_NUMLIGNE'));
  end;

  CopieOuvFromLigne (TOBOuv,TOBLigPiece);

  if TobArticle <> nil then
  begin
    // Modif BTP
    TOBOuv.PutValue(prefixe+'_TENUESTOCK',  TobArticle.GetValue('GA_TENUESTOCK'));
    //TobLN.PutValue('GLN_QUALIFQTEACH',TobArticle.GetValue('GA_QUALIFQTEACH'));
    TOBOuv.PutValue(prefixe+'_QUALIFQTEVTE',TobArticle.GetValue('GA_QUALIFUNITEVTE'));
    TOBOuv.PutValue(prefixe+'_QUALIFQTESTO',TobArticle.GetValue('GA_QUALIFUNITESTO'));
    TOBOuv.PutValue(prefixe+'_PRIXPOURQTE', TobArticle.GetValue('GA_PRIXPOURQTE'));
    TOBOuv.PutValue(prefixe+'_DPA',         TobArticle.GetValue('GA_PAHT'));
    TOBOuv.PutValue(prefixe+'_PMAP',        TobArticle.GetValue('GA_PMAP'));
    TOBOuv.PutValue('ANCPA',                TobOuv.GetValue(prefixe+'_DPA'));
    TOBOuv.PutValue('ANCPR',                TobOuv.GetValue(prefixe+'_DPR'));
    //
    CopieOuvFromArt (TOBPiece,TOBOuv,TobArticle,DEV,RecupPrix);
    //
    TOBOuv.SetDouble('BLO_QTESAIS',TOBOuv.GetDouble('BLO_QTEFACT'));
    if fGestQteAvanc then
    begin
      if (POS(TOBOuv.GetString('BNP_TYPERESSOURCE'),'SAL;INT;ST') > 0)  then
      begin
        TOBOuv.PutValue('BLO_QUALIFHEURE',     TobArticle.GetValue('GA_QUALIFHEURE'));
        TOBOuv.PutValue('BLO_RENDEMENT',       TobArticle.GetValue('GA_COEFPROD'));
        TOBOuv.PutValue('BLO_PERTE',0);
      end else
      begin
        TOBOuv.SetString('BLO_QUALIFHEURE','');
        TOBOuv.PutValue('BLO_PERTE',           TobArticle.GetValue('GA_PERTEPROP'));
        TOBOuv.PutValue('BLO_RENDEMENT',0);
      end;
    end;
    TOBOuv.SetDouble('BLO_QTEFACT',CalculeQteFact (TOBOuv));
    //
    if ((RecupPrix='PAS') or (RecupPrix = 'DPA')) then
    begin
      if AvecPv then
      begin
        TOBOuv.PutValue(prefixe+'_PUHT',    TOBOuv.GetValue(prefixe+'_DPA'));
        TOBOuv.PutValue(prefixe+'_PUHTBASE',TOBOuv.GetValue(prefixe+'_DPA'));
        TOBOuv.PutValue(prefixe+'_PUHTDEV', pivottodevise(TOBOuv.GetValue(prefixe+'_DPA'),DEV.Taux,DEV.quotite,V_PGI.OkdecP ));
        TOBOuv.PutValue('ANCPV',            TobOuv.GetValue(prefixe+'_PUHTDEV'));
        TOBOuv.PutValue(prefixe+'_PUTTC',   TOBOuv.GetValue(prefixe+'_DPA'));
        TOBOuv.PutValue(prefixe+'_PUTTCBASE',TOBOuv.GetValue(prefixe+'_DPA'));
        TOBOuv.PutValue(prefixe+'_PUTTCDEV',pivottodevise(TOBOuv.GetValue(prefixe+'_DPA'),DEV.Taux,DEV.quotite,V_PGI.OkdecP ));
      end;
    end
    else if RecupPrix = 'DPR' then
    begin
      if AvecPv then
      begin
        TOBOuv.PutValue(prefixe+'_PUHT',    TOBOuv.GetValue(prefixe+'_DPR'));
        TOBOuv.PutValue(prefixe+'_PUHTBASE',TOBOuv.GetValue(prefixe+'_DPR'));
        TOBOuv.PutValue(prefixe+'_PUHTDEV', pivottodevise(TOBOuv.GetValue(prefixe+'_DPR'),DEV.Taux,DEV.quotite,V_PGI.OkdecP ));
        TOBOuv.PutValue('ANCPV',            TobOuv.GetValue(prefixe+'_PUHTDEV'));
        TOBOuv.PutValue(prefixe+'_PUTTC',   TOBOuv.GetValue(prefixe+'_DPR'));
        TOBOuv.PutValue(prefixe+'_PUTTCBASE',TOBOuv.GetValue(prefixe+'_DPR'));
        TOBOuv.PutValue(prefixe+'_PUTTCDEV',pivottodevise(TOBOuv.GetValue(prefixe+'_DPR'),DEV.Taux,DEV.quotite,V_PGI.OkdecP ));
      end;
    end
    else if RecupPrix = 'PUH' then
    begin
      if AvecPv then
      begin
        TOBOuv.PutValue(prefixe+'_PUHT',    TobArticle.GetValue('GA_PVHT'));
        TOBOuv.PutValue(prefixe+'_PUHTBASE',TobArticle.GetValue('GA_PVHT'));
        TOBOuv.PutValue(prefixe+'_PUHTDEV', pivottodevise(TobOuv.GetValue(prefixe+'_PUHTBASE'),DEV.Taux,DEV.quotite,V_PGI.OkdecP ));
        TOBOuv.PutValue('ANCPV',            TobOuv.GetValue(prefixe+'_PUHTDEV'));
        TOBOuv.PutValue(prefixe+'_PUTTC',   TobArticle.GetValue('GA_PVTTC'));
        TOBOuv.PutValue(prefixe+'_PUTTCBASE',TobArticle.GetValue('GA_PVTTC'));
        TOBOuv.PutValue(prefixe+'_PUTTCDEV',pivottodevise(TOBOuv.GetValue(prefixe+'_PUTTCBASE'),DEV.Taux,DEV.quotite,V_PGI.OkdecP ));
      end;
    end;
    // Modif BTP
    if Libelle <> '' then TOBOuv.PutValue(prefixe+'_LIBELLE',Libelle);
   	TOBOUv.putValue('BNP_TYPERESSOURCE',TOBArticle.getValue('BNP_TYPERESSOURCE'));
    //
    if (Pos (TOBOUV.getValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne) > 0) and (not IsLigneExternalise(TOBOUV)) then
      TOBOUV.PutValue('BLO_TPSUNITAIRE',1)
    else
      TOBOUV.PutValue('BLO_TPSUNITAIRE',0);
    //
    if (IsPrestationSt(TOBOUV)) then
    begin
      if TOBPiece.getValue('GP_APPLICFGST')='X' then
        TOBOuv.putValue('BLO_NONAPPLICFRAIS','-')
      else
        TOBOuv.putValue('BLO_NONAPPLICFRAIS','X');
      //
      if TOBPiece.getValue('GP_APPLICFCST')='X' then
        TOBOuv.putValue('BLO_NONAPPLICFC','-')
      else
        TOBOuv.putValue('BLO_NONAPPLICFC','X');
    end
    else
    begin
      TOBOuv.putValue('BLO_NONAPPLICFRAIS','-');
      TOBOuv.putValue('BLO_NONAPPLICFC','-');
    end;
    //
    if TobArticle.getValue('GA_PRIXPASMODIF')<>'X' Then
    begin
      AppliqueCoefDomaineActOuv (TOBOuv,'A',RecupPrix);
      if VH_GC.BTCODESPECIF = '001' then ApliqueCoeflignePOC (TOBOuv,DEV);
    end;
    //
    CalculeLigneAcOuv (TOBOUV,DEV,true,TOBARTICLE);
    //
    if TOBLigPiece.getValue('GLC_NATURETRAVAIL') <> '' then
    begin
      TOBOuv.PutValue('BLO_NATURETRAVAIL',TOBLigPiece.getValue('GLC_NATURETRAVAIL'));
      TOBOuv.PutValue('BLO_FOURNISSEUR',  TOBLigPiece.getValue('GL_FOURNISSEUR'));
      TOBOuv.PutValue('GA_FOURNPRINC',    TOBLigPiece.getValue('GL_FOURNISSEUR'));
    end;
  end
  else
  begin
    if AvecPv then
    begin
      TOBOuv.PutValue(prefixe+'_PUHT',0);
      TOBOuv.PutValue(prefixe+'_PUHTBASE',0);
      TOBOuv.PutValue(prefixe+'_PUHTDEV',0);
      TOBOuv.PutValue('ANCPV',  0);
      TOBOuv.PutValue(prefixe+'_PUTTC',0);
      TOBOuv.PutValue(prefixe+'_PUTTCBASE',0);
      TOBOuv.PutValue(prefixe+'_PUTTCDEV',0);
      TOBOuv.PutValue(prefixe+'_DPA',0);
      TOBOuv.PutValue(prefixe+'_DPR',0);
      TOBOuv.PutValue(prefixe+'_PMAP',0);
      TOBOuv.PutValue(prefixe+'_PMRP',0);
      TOBOuv.PutValue(prefixe+'_QUALIFQTEVTE','');
      TOBOuv.PutValue(prefixe+'_QUALIFQTESTO','');
    end;
    TOBOuv.PutValue('ANCPA',0);
    TOBOuv.PutValue('ANCPR',0);
    //
    CalculeLigneAcOuv (TOBOUV,DEV,True,TOBARTICLE);
  end;
end;
{$ENDIF}

procedure TFLigNomen.RecupInfoPiece ;
begin
// Modif BTP
{$IFDEF BTP}
if Ouvrage then
begin
TOBInfo:=TOB.Create('LIGNEOUV',nil,-1);
InsertionChampSupOuv (TOBInfo,false);
end else
{$ENDIF}
TOBInfo:=TOB.Create('LIGNENOMEN',nil,-1);

TOBInfo.PutValue(prefixe+'_NATUREPIECEG',TobLig.Detail[0].GetValue(prefixe+'_NATUREPIECEG'));
TOBInfo.PutValue(prefixe+'_SOUCHE',TobLig.Detail[0].GetValue(prefixe+'_SOUCHE'));
TOBInfo.PutValue(prefixe+'_NUMERO',TobLig.Detail[0].GetValue(prefixe+'_NUMERO'));
TOBInfo.PutValue(prefixe+'_INDICEG',TobLig.Detail[0].GetValue(prefixe+'_INDICEG'));
TOBInfo.PutValue(prefixe+'_NUMLIGNE',TobLig.Detail[0].GetValue(prefixe+'_NUMLIGNE'));
end;

procedure TFLigNomen.ChargeNomen(CodeNomen : string ; TobLN : TOB ; Niv,OrdreCompo: integer );
var
    i_ind1, i_ind2 : integer;
    Select,TheArticle : string;
    Q,QQ : TQuery;
    TobLND, TobLigne, TobFille : TOB;
    TobArt : TOB;
    // Modif BTP
//    QteDudetail : Double;
begin
  {$IFDEF BTP}

  //QteDuDetail := 1;
  //  CodeNomen obligatoire

  if CodeNomen <> '' then
  begin
    TobLN.PutValue('SOUSNOMEN','X');   //En cas de récursivité
    TobLigne := TOB.Create('NOMENLIG', nil, -1);
    //  Chargement du niveau courant de nomenclature
    Select := 'SELECT GNL_SOUSNOMEN,GNL_ARTICLE,GNL_CODEARTICLE,GNL_LIBELLE,GNL_QTE ';
    Select := Select + 'FROM NOMENLIG WHERE GNL_NOMENCLATURE = "' + CodeNomen + '"';
    Q := OpenSQL(Select, true,-1, '', True);
    if not Q.EOF then
    begin
      TobLigne.LoadDetailDB('NOMENLIG', '', '', Q, False);
      for i_ind1:=0 to TobLigne.Detail.count -1 do
      begin
        if Ouvrage then
        begin
          TobLND:=TOB.Create('LIGNEOUV',TobLN,-1);
          InsertionChampSupOuv (TobLND,false);
        end
        else
          TobLND:=TOB.Create('LIGNENOMEN',TobLN,-1);

        TobLND.AddChampSup('LIBELLE', True);   //  Ajout du champ sup. LIBELLE = CODEARTICLE (QUANTITE)
        TobLND.AddChampSup('SOUSNOMEN',False);

        TobLND.PutValue(prefixe+'_COMPOSE',TobLN.GetValue(prefixe+'_ARTICLE'));
        TobLND.PutValue(prefixe+'_NIVEAU',Niv);
        TobLND.PutValue(prefixe+'_ORDRECOMPO',OrdreCompo);

        TobLND.PutValue(prefixe+'_NUMORDRE',i_ind1+1);
        TobLND.AddChampSup('CODENOMEN', True);
        TobLND.PutValue('CODENOMEN',TobLigne.Detail[i_ind1].GetValue('GNL_SOUSNOMEN'));

        TobLND.PutValue(prefixe+'_ARTICLE', TobLigne.Detail[i_ind1].GetValue('GNL_ARTICLE'));
        TobArt:=TOB.Create('ARTICLE',nil,-1);
        //
        TheArticle := TobLigne.Detail[i_ind1].GetValue('GNL_ARTICLE');
        //

        QQ := OPenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
            'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
            'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE GA_ARTICLE ="'+TheArticle+'"',True,-1, '', True);
        TOBArt:= TOB.create ('ARTICLE',nil,-1);
        TOBArt.selectDB ('',QQ);
        ferme (QQ);
        InitChampsSupArticle (TOBART);
        //
        //
        TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
        TobLND.putValue('BLO_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));
        //
        RenseigneTOBLN(TobLND,TobArt,TOBInfo);
        TobLND.PutValue(prefixe+'_CODEARTICLE',TobLigne.Detail[i_ind1].GetValue('GNL_CODEARTICLE'));
        TobLND.PutValue(prefixe+'_LIBELLE',TobLigne.Detail[i_ind1].GetValue('GNL_LIBELLE'));
        TobLND.PutValue('GLN_QTE', TobLigne.Detail[i_ind1].GetValue('GNL_QTE'));
        TobArt.free;
      end;
    end;
    // MODIF LS
    TOBLigne.free;
    Ferme(Q);
    //  on balaye les filles pour charger les eventuelles nomenclatures
    //  de sous niveau
    for i_ind2 := 0 to TobLN.Detail.Count - 1 do
    begin
      TobFille := TobLN.Detail[i_ind2];
      ChargeNomen(TobFille.GetValue('CODENOMEN'), TobFille , Niv+1,TobFille.GetValue(prefixe+'_NUMORDRE'));
    end;
  end;
  {$ENDIF}

end;

{$IFDEF BTP}
function TFLigNomen.ChargeOuvrage(CodeNomen : string ; TobOuvrage : TOB ; Niv,OrdreCompo : integer ;var  Tresult: T_valeurs):boolean;
var
    i_ind1, i_ind2,IndPou : integer;
    Select : string;
    Q,QQ : TQuery;
    TOBDetOuv, TobLigne, TobFille : TOB;
    TobArt : TOB;
    // Modif BTP
    TOBLigNom : TOB;
    QteDudetail,QTe,QTedupv,MontantInterm,MontantPr : Double;
    Valloc : T_Valeurs;
    TOBPiece : TOB;
begin
if niv > 5 then
  BEGIN
  PGIBox ('Nombre de niveaux trop important : 5 Maximum','Ouvrages');
  Result := false;
  Exit;
  END;
TOBPiece := TOBLigPiece.parent;
result := true;
QteDuDetail := 1;
indpou := -1;
//  CodeNomen obligatoire
if CodeNomen <> '' then
   begin
   InitTableau (Tresult);
   TobOuvrage.PutValue('SOUSNOMEN','X');   //En cas de récursivité
   TobLigne := TOB.Create('NOMENLIG', nil, -1);
   //  Chargement du niveau courant de nomenclature
   Select := 'Select NOMENLIG.*,GNE_QTEDUDETAIL,GNE_LIBELLE from NOMENLIG LEFT OUTER JOIN NOMENENT ON GNE_NOMENCLATURE = GNL_NOMENCLATURE where GNL_NOMENCLATURE = "' + CodeNomen + '"';
   Q := OpenSQL(Select, true,-1, '', True);
   if not Q.EOF then
      begin
      TobLigne.LoadDetailDB('DETAILNOMEN', '', '', Q, False);
      ferme (Q);
      for i_ind1:=0 to TobLigne.Detail.count -1 do
          begin
          TOBLIGNom := TOBLIgne.detail[i_ind1];
          QteDuDetail := TOBLigNom.getValue('GNE_QTEDUDETAIL');
          if QteDuDetail = 0 then QteDuDetail := 1;
          TOBDetOuv:=TOB.Create('LIGNEOUV',TobOuvrage,-1);
          InsertionChampSupOuv (TOBDetOuv,false);
          //
          TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
          TOBDetOuv.putValue('BLO_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));
          //
          // prise en compte document d'origine
          TOBDetOuv.PutValue(prefixe+'_NATUREPIECEG',TOBLigPiece.getValue('GL_NATUREPIECEG'));
          TOBDetOuv.PutValue(prefixe+'_SOUCHE',TOBLigPiece.getValue('GL_SOUCHE'));
          TOBDetOuv.PutValue(prefixe+'_NUMERO',TOBLigPiece.getValue('GL_NUMERO'));
          TOBDetOuv.PutValue(prefixe+'_INDICEG',TOBLigPiece.getValue('GL_INDICEG'));
          // ------
          TOBDetOuv.PutValue(prefixe+'_COMPOSE',TobOuvrage.GetValue(prefixe+'_ARTICLE'));
          TOBDetOuv.PutValue(prefixe+'_NIVEAU',Niv);
          TOBDetOuv.PutValue(prefixe+'_ORDRECOMPO',OrdreCompo);

          TOBDetOuv.PutValue(prefixe+'_NUMORDRE',i_ind1+1);
          TOBDetOuv.PutValue('CODENOMEN',TobLigNom.GetValue('GNL_SOUSNOMEN'));
          TOBDETOuv.putvalue('BLO_QTEDUDETAIL',TOBLigNom.getvalue('GNE_QTEDUDETAIL'));
          TOBDetOuv.PutValue(prefixe+'_ARTICLE', TOBLigNom.GetValue('GNL_ARTICLE'));

//
          QQ := OPenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                          'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                          'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE GA_ARTICLE ="'+
                          TobLigNom.GetValue('GNL_ARTICLE')+'"',True,-1, '', True);
          TOBArt:= TOB.create ('ARTICLE',nil,-1);
          TOBArt.selectDB ('',QQ);
          ferme (QQ);
          InitChampsSupArticle (TOBART);
//
          TOBDetOuv.PutValue('BLO_QTEFACT', TOBLigNom.GetValue('GNL_QTE'));
          RenseigneTOBOuv(TOBDetOuv,TobArt);
          TOBDetOuv.PutValue(prefixe+'_CODEARTICLE',TobLigNom.GetValue('GNL_CODEARTICLE'));
          // Modif BTP
          TOBDetOuv.PutValue(prefixe+'_REFARTSAISIE',TobLigNom.GetValue('GNL_CODEARTICLE'));
          // ---
          TOBDetOuv.PutValue(prefixe+'_LIBELLE',TOBLigNom.GetValue('GNL_LIBELLE'));
          if TobDetOuv.GetValue('CODENOMEN') <> '' then
             begin
             if not ChargeOuvrage(TobDetOuv.GetValue('CODENOMEN'), TobDetOuv , Niv+1,TobDetOuv.GetValue(prefixe+'_NUMORDRE'),valloc) then
                begin
                TOBArt.free;
                TOBDetOuv.free;
                TOBligne.free;
                Result := false;
                Exit;
                end;
             TOBDetOuv.Putvalue ('BLO_DPA',valloc[0]);
             TOBDetOuv.Putvalue ('BLO_DPR',valloc[1]);
             TOBDetOuv.Putvalue ('BLO_PUHT',Valloc[2]);
             TOBDetOuv.Putvalue ('BLO_PUHTBASE',Valloc[2]);
             TOBDetOuv.Putvalue ('BLO_PUHTDEV',pivottodevise(Valloc[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
             TOBDetOuv.Putvalue ('BLO_PUTTC',Valloc[3]);
             TOBDetOuv.Putvalue ('BLO_PUTTCDEV',pivottodevise(Valloc[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
             TOBDetOuv.Putvalue ('BLO_PMAP',Valloc[6]);
             TOBDetOuv.Putvalue ('BLO_PMRP',VAlloc[7]);
             TOBDetOuv.Putvalue ('BLO_TPSUNITAIRE',VAlloc[9]);
            TOBDetOuv.PutValue('BLO_MONTANTFG',Arrondi(ValLoc[13]*TOBDetOuv.getValue('BLO_QTEFACT'),V_PGI.OkdecP));
            TOBDetOuv.PutValue('BLO_MONTANTFC',Arrondi(ValLoc[15]*TOBDetOuv.getValue('BLO_QTEFACT'),V_PGI.OkdecP));
            TOBDetOuv.PutValue('BLO_MONTANTFR',Arrondi(ValLoc[14]*TOBDetOuv.getValue('BLO_QTEFACT'),V_PGI.OkdecP));
            CalculeLigneAcOuvCumul (TOBDETOUV);
             CalculMontantHtDevLigOuv (TOBDetOuv,DEV);
             end;
          TOBDetOuv.PutValue('ANCPV',TOBDetOuv.Getvalue('BLO_PUHTDEV') );
          TOBDetOuv.PutValue('ANCPA',TOBDetOuv.Getvalue('BLO_DPA') );
          TOBDetOuv.PutValue('ANCPR',TOBDetOuv.Getvalue('BLO_DPR') );
          if TOBLigNom.GetValue('GNL_PRIXFIXE') <> 0 then
             begin
             TOBDetOuv.PutValue('BLO_PUHT', TOBLigNom.GetValue('GNL_PRIXFIXE'));
             TOBDetOuv.PutValue('BLO_PUHTDEV', pivottodevise(TOBLigNom.GetValue('GNL_PRIXFIXE'),DEV.Taux,DEV.quotite,V_PGI.OkdecP));
             end;
          Qte := TOBDetOuv.getValue('BLO_QTEFACT');
          QteDuPv := TOBDetOuv.getValue('BLO_PRIXPOURQTE'); if QteDuPv = 0 then QteDuPv := 1;
          CalculeLigneAcOuv (TOBDETOUV,DEV,true,TOBART);
          CalculMontantHtDevLigOuv (TOBDetOuv,DEV);
          TobArt.free;
          if TobDetOuv.Getvalue('BLO_TYPEARTICLE') <> 'POU' then
             begin
             Tresult[0] := Tresult [0] + arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_DPA')),V_PGI.okdecP);
             Tresult[1] := Tresult [1] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_DPR')),V_PGI.okdecP);
             Tresult[2] := Tresult [2] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PUHT')),V_PGI.okdecP);
             Tresult[3] := Tresult [3] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PUTTC')),V_PGI.okdecP);
             Tresult[6] := Tresult [6] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PMAP')),V_PGI.okdecP);
             Tresult[7] := Tresult [7] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PMRP')),V_PGI.okdecP);
             Tresult[9] := Tresult [9] +(Qte * TOBDetOuv.GetValue ('BLO_TPSUNITAIRE'));
             Tresult[13] := Tresult [13] +TOBDetOuv.GetValue('BLO_MONTANTFG');
             Tresult[14] := Tresult [14] +TOBDetOuv.GetValue('BLO_MONTANTFR');
             Tresult[15] := Tresult [15] +TOBDetOuv.GetValue('BLO_MONTANTFC');
             end else
             begin
             indpou := 1;
             end;
          end;
      end else Ferme(Q);
   TOBLigne.free;
   FormatageTableau (Tresult,V_PGI.OkdecP);
   // calcule des eventuels pourcentages
   if indpou >= 0 then
      begin
      for I_ind1 := 0 to TOBOuvrage.detail.count -1 do
          begin
          TOBLIGNom := TOBOuvrage.detail[i_ind1];
          if TobLigNom.Getvalue('BLO_TYPEARTICLE') <> 'POU' then continue;
          TOBLIGNom.Putvalue ('BLO_DPA',Tresult[0]);
          TOBLIGNom.Putvalue ('BLO_DPR',Tresult[1]);
          TOBLIGNom.Putvalue ('BLO_PUHT',Tresult[2]);
          TOBLIGNom.Putvalue ('BLO_PUHTBASE',Tresult[2]);
          TOBLIGNom.Putvalue ('BLO_PUHTDEV',pivottodevise(Tresult[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
          TOBLIGNom.Putvalue ('BLO_PUTTC',Tresult[3]);
          TOBLIGNom.Putvalue ('BLO_PUTTCDEV',pivottodevise(Tresult[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
          TOBLIGNom.Putvalue ('BLO_PMAP',Tresult[6]);
          TOBLIGNom.Putvalue ('BLO_PMRP',Tresult[7]);
          Qte := TOBDetOuv.getValue('BLO_QTEFACT');
          CalculMontantHtDevLigOuv (TOBLigNom,DEV);
          QteDuPv := TOBDetOuv.getValue('BLO_PRIXPOURQTE'); if QteDuPv = 0 then QteDuPv := 1;
          Tresult[0] := Tresult [0] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_DPA')),V_PGI.OkdecP);
          Tresult[1] := Tresult [1] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_DPR')),V_PGI.OkdecP);
          Tresult[2] := Tresult [2] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PUHT')),V_PGI.OkdecP);
          Tresult[3] := Tresult [3] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PUTTC')),V_PGI.OkdecP);
          Tresult[6] := Tresult [6] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PMAP')),V_PGI.OkdecP);
          Tresult[7] := Tresult [7] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PMRP')),V_PGI.OkdecP);
          FormatageTableau (Tresult,V_PGI.OkdecP );
          end;
      end;
   for i_ind2 := 0 to TobOuvrage.Detail.Count - 1 do
       begin
       TobFille := TobOuvrage.Detail[i_ind2];
       if (TobFille.getvalue('ANCPV') <> TOBFille.getValue('BLO_PUHTDEV')) and
          (TobFille.detail.count > 0) then
          begin
          MontantInterm := TOBfille.getValue('BLO_PUHTDEV');
          if reajuste then
          begin
          	ReajusteMontantOuvrage (nil,TOBPiece,TOBLigPiece,TOBFille,TOBFille.getValue('ANCPV'),MontantPr,MontantInterm,DEV,GestionHt);
            TOBFille.putvalue('BLO_PUHTDEV',MontantInterm);
          end;
          TOBDetOuv.PutValue('ANCPV',TOBDetOuv.Getvalue('BLO_PUHTDEV') );
          CalculMontantHtDevLigOuv (TOBDetOuv,DEV);
          end;
       end;
   GestionDetailPrixPose (TobOuvrage);
   end;
   Tresult := CalculSurTableau ('/',Tresult,QteDudetail);
   FormatageTableau (Tresult,V_PGI.OkdecP);
end;
{$ENDIF}
////////////////////////////////////////////////////////////////////////////////
//************************* Traitement sur les tobs **************************//
////////////////////////////////////////////////////////////////////////////////

procedure TFLigNomen.GereSousNomen(tobln : tob);
var toblnd : tob;
    i_ind1 : integer;
begin
for i_ind1:=0 to TobLN.Detail.Count-1 do
    begin
    TobLND:=TobLN.Detail[i_ind1];
    if not TOBLND.FieldExists ('SOUSNOMEN') then TobLND.AddchampSup('SOUSNOMEN',true);
    if TobLND.Detail.Count > 0 then
       begin
       TobLND.PutValue('SOUSNOMEN','X');
       GereSousNomen(TobLND);
       end;
     end;
end;

procedure TFLigNomen.DepointeSousNomen(tobln : tob);
var toblnd : tob;
    i_ind1 : integer;
begin
for i_ind1:=0 to TobLN.Detail.Count-1 do
    begin
    TobLND:=TobLN.Detail[i_ind1];
    TobLND.DelChampSup ('SOUSNOMEN',true);
    end;
end;

function  TFLigNomen.VerifNomen(TobArt : TOB ; Arow : integer; Var retour : string; var Libelle : string ) : boolean;
var Q : TQuery;
    st1 : string;
    i_ind1 : integer;
    TypeArticle : string;
begin
result := true;
st1 := '';
i_ind1 := 0;
TYpeArticle := TOBArt.GetValue('GA_TYPEARTICLE');
if ((TypeArticle <> 'NOM') and (not Ouvrage)) or
   ((TypeArticle <> 'OUV') and (TypeArticle <> 'ARP') and (Ouvrage)) then
   BEGIN
   exit;
   END else TOB(G_NLIG.Objects[SGN_NumLig, Arow]).PutValue('SOUSNOMEN','X');

Q := OpenSQL('Select Count(GNE_NOMENCLATURE) from NOMENENT Where GNE_ARTICLE="' +
             TOBArt.GetValue('GA_ARTICLE') + '"', True,-1, '', True);
if not(Q.Eof) then i_ind1 := Q.Fields[0].AsInteger;
Ferme(Q);

if i_ind1 > 1 then
   BEGIN
   st1:=Choisir('Choix d''une nomenclature','NOMENENT','GNE_LIBELLE',
          'GNE_NOMENCLATURE','GNE_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'"','');
   if st1 = '' then result := false;
   END else if i_ind1 = 1 then
   BEGIN
   Q := OpenSQL('Select GNE_NOMENCLATURE from NOMENENT Where GNE_ARTICLE="' +
                TOBArt.GetValue('GA_ARTICLE') + '"', True,-1, '', True);
   st1 := Q.FindField('GNE_NOMENCLATURE').AsString;
   Ferme(Q);
   end;
if st1 <> '' then // recupération du libellé
   begin
   Q := OpenSQL('Select GNE_LIBELLE from NOMENENT Where GNE_NOMENCLATURE="' + st1 + '"', True,-1, '', True);
   Libelle := Q.FindField('GNE_Libelle').AsString;
   Ferme(Q);
   end;
Retour := st1;
end;

function TFLigNomen.TobVide(TobLN : TOB): boolean;
begin
Result:=false;
if TobLN.Detail.Count=0 then Result:=true;
if TobLN.Detail.Count=1 then
   if TobLN.Detail[0].GetValue(prefixe+'_ARTICLE')='' then Result:=true;
end;

//=============================================================================
//  Verification de non circularité
//=============================================================================
function TFLigNomen.Circularite(CodeCompose : string ; TobLigNomen : TOB; var TOBErreur : TOB) : Boolean;
var i_ind1 : integer;
    TOBLND, TOBTempErr : TOB;
begin
Result := True;
for i_ind1 := 0 to TobLigNomen.Detail.Count - 1 do
    begin
    TOBLND := TobLigNomen.Detail[i_ind1];
//  pas de code de sous nomenclature, on continue
    if TOBLND.GetValue('SOUSNOMEN') <> 'X' then Continue;
{$IFDEF BTP}
    if TOBLND.GetValue('BLO_TYPEARTICLE')='ARP' Then continue;
{$ENDIF}
//  on compare avec le code de sous nomenclature eventuel
    if CodeCompose = TOBLND.GetValue(prefixe+'_ARTICLE') then
        begin
        TOBTempErr := TOB.Create('', TOBErreur, -1);
        TOBTempErr.AddChampSup(prefixe+'_COMPOSE', True);
        TOBTempErr.PutValue(prefixe+'_COMPOSE', CodeCompose);
        Result := False;
//      Exit;
        end else
        begin
        TOBTempErr := TOB.Create('', TOBErreur, -1);
        TOBTempErr.AddChampSup(prefixe+'_COMPOSE', True);
        TOBTempErr.PutValue(prefixe+'_COMPOSE', TOBLND.GetValue(prefixe+'_ARTICLE'));
        end;
    if not Circularite(CodeCompose, TOBLND, TOBTempErr) then
       begin
       Result := False;
//     Exit;
       end;
    end;
end;

////////////////////////////////////////////////////////////////////////////////
//*************************** Gestion du TreeView ****************************//
////////////////////////////////////////////////////////////////////////////////

//-----------------------------------------------------------------------------
//  Positionnement de l'icone de ImageList1 associée à un element.
//          0 : Nomenclature, non étendue
//          1 : Nomenclature, etendue
//          2 : article
//-----------------------------------------------------------------------------
procedure TFLigNomen.SetImages(TV : TTreeView; TN : TTreeNode);
var tn1 : TTreeNode;
begin
if (TV = nil) and (TN = nil) then Exit;
if TV <> nil then tn1 := TV.TopItem else
    if TN.Count <> 0 then tn1 := TN.getFirstChild else
        tn1 := TN;
while tn1 <> nil do
    begin
    if tn1.Count = 0 then
        begin
        tn1.ImageIndex := 2;
        tn1.SelectedIndex := 2;
        end
    else
        begin
        tn1.ImageIndex := 0;
        tn1.SelectedIndex := 0;
        SetImages(nil, tn1);
        if tn1.Expanded then
            begin
            tn1.ImageIndex := 1;
            end;
        end;
    if (TV <> nil) or (TN.Count <> 0) then tn1 := tn1.GetNextSibling else tn1 := nil;
    end;
end;

//-----------------------------------------------------------------------------
//  Affichage du treeview TV_NLIG
//-----------------------------------------------------------------------------
procedure TFLigNomen.AfficheTVNLIG;
var  i_ind1 : integer;
begin
AfficheTreeView(TV_NLIG, TOBLig, 'LIBELLE');
TV_NLIG.TopItem.Text := Copy(COMPOSE.Text,0,19);
TV_NLIG.TopItem.Expand(False);
SetImages(TV_NLIG, nil);
for i_ind1 := 0 to G_NLIG.RowCount - 1 do
    begin
    if G_NLIG.Cells[SGN_Comp, i_ind1] = '' then break;
    G_NLIG.Objects[SGN_Comp, i_ind1] := nil;
    if G_NLIG.Cells[SGN_Compose, i_ind1] <> COMPOSE.Text then
       G_NLIG.Objects[SGN_Comp, i_ind1] := ChercheNode(TV_NLIG, nil, i_ind1+1);
    if G_NLIG.Objects[SGN_Comp, i_ind1] = nil then
       G_NLIG.Objects[SGN_Comp, i_ind1] := ChercheNode(TV_NLIG, nil, i_ind1+1);
    end;
end;

function TFLigNomen.ChercheNode(TV : TTreeView; TN : TTreeNode; Valeur : integer;
                     Niveau : integer = 1) : TTreeNode;
var
    tn1 : TTreeNode;
    i : integer ;
begin
Result := nil;
i:=0;
if TV <> nil then tn1 := TV.Items.GetFirstNode else tn1 := TN.getFirstChild;
while tn1 <> nil do
    begin
    if tn1.Level <= Niveau then
        begin
        i:=i+1;
        if i = valeur then
            begin
            Result := tn1;
            Exit;
            end;
        end;
    tn1 := tn1.GetNext;
    end;
end;

//=============================================================================
//  Ajout du champ LIBELLE à la tob ligne : contient le libellé des TTreeNodes
//=============================================================================
procedure TFLigNomen.AffecteLibelle (TobLN : TOB; stCol : string);
var TOBL : TOB ;
    i_ind1 : integer ;
    st1, st2, st3, st4 : string;
begin
for i_ind1 := 0 to TobLN.Detail.Count - 1 do
    begin
    if i_ind1 = 0 then
       begin
       if not TOBLN.FieldExists ('LIBELLE') then TobLN.Detail[i_ind1].AddChampSup('LIBELLE', True);
       end;
    st2 := VarAsType(TobLN.Detail[i_ind1].GetValue(stCol), varString);
    // Modif BTP
    { avant
    st4 := VarAsType(TobLN.Detail[i_ind1].GetValue('GLN_CODEARTICLE'), varString);
    st3 := VarAsType(TobLN.Detail[i_ind1].GetValue('GLN_QTE'), varString);
    }
    st4 := VarAsType(TobLN.Detail[i_ind1].GetValue(prefixe+'_CODEARTICLE'), varString);
    if Ouvrage then st3 := VarAsType(TobLN.Detail[i_ind1].GetValue('BLO_QTEFACT'), varString)
    else st3 := VarAsType(TobLN.Detail[i_ind1].GetValue('GLN_QTE'), varString);
    if st4 <> '' then
        st1 := st4 + ' (' + StrS0(Valeur(st3)) + ')'
    else
        st1 := st2 + ' (' + StrS0(Valeur(st3)) + ')';
    TobLN.Detail[i_ind1].PutValue('LIBELLE', st1);
    if TobLN.Detail[i_ind1].Detail.Count > 0 then
        begin
        TOBL := TobLN.Detail[i_ind1];
        AffecteLibelle(TOBL, stCol);
        end;
    end;
end;

procedure TFLigNomen.Re_AffecteLibelle (TobLND : TOB ; stCol : string ; Qte : double);
var st1, st2, st3, st4 : string;
begin
st2 := VarAsType(TobLND.GetValue(stCol), varString);
st4 := VarAsType(TobLND.GetValue(prefixe+'_CODEARTICLE'), varString);
st3 := VarAsType(Qte, varString);
if   st4 <> '' then st1 := st4 + ' (' + StrS0(Valeur(st3)) + ')'
else st1 := st2 + ' (' + StrS0(Valeur(st3)) + ')';
TobLND.PutValue('LIBELLE', st1);
end;

//=============================================================================
//  Evenements du treeview
//=============================================================================
procedure TFLigNomen.TV_NLIGMouseUp(Sender: TObject; Button: TMouseButton;
                                    Shift: TShiftState; X, Y: Integer);
var
    tn1 : TTreeNode;
    i_ind1 : integer;
    b_temp : boolean;
begin
case Button of
    mbLeft  : begin
              tn1 := TV_NLIG.GetNodeAt(X, Y);
              if tn1 = nil then Exit;
//--------------------------------------------------------------------------
//  le noeud selectionné n'est pas de niveau 1 donc pas accessible directement
//  on positionne sur le parent de niveau 1
//--------------------------------------------------------------------------
              if tn1.Level > 1 then
                  begin
                  tn1 := FindParentNiv1(tn1);
                  TV_NLIG.Selected := tn1;
                  end;
//--------------------------------------------------------------------------
//  affichage synchrone de la grid avec le treeview
//--------------------------------------------------------------------------
              i_ind1 := ChercheLigne(tn1);
              if i_ind1 <> 0 then
                  begin
                  b_temp := False;
                  if G_NLIG.Row <> i_ind1 then G_NLIGRowExit(Sender, G_NLIG.Row, b_temp, b_temp);
                  G_NLIG.Col := SGN_Comp;
                  G_NLIG.Row := i_ind1;
                  G_NLIGRowEnter(Sender, i_ind1, b_temp, b_temp);
                  G_NLIG.SetFocus;
                  TV_NLIG.Selected := TTreeNode(G_NLIG.Objects[SGN_Comp, i_ind1]);
                  AffectePied(G_NLIG.Row);
                  Exit;
                  end;
              end;
    mbRight : begin
              end;
    end;
end;

procedure TFLigNomen.TV_NLIGCollapsing(Sender: TObject; Node: TTreeNode;
                                       var AllowCollapse: Boolean);
begin
Node.ImageIndex := 0;
end;

procedure TFLigNomen.TV_NLIGExpanding(Sender: TObject; Node: TTreeNode;
                                      var AllowExpansion: Boolean);
begin
Node.ImageIndex := 1;
end;

//=============================================================================
//  Double clic, on ouvre la decomposition dans une nouvelle fenetre
//=============================================================================
procedure TFLigNomen.TV_NLIGDblClick(Sender: TObject);
var Active_Ligne, i_ind1 : integer;
    TOBTemp,TobLND,TobNew : TOB;
    tn1 : TTreeNode;
{$IFDEF BTP}
    valloc : T_valeurs;
    OkTrait : boolean;
    LastAction : TactionFiche;
{$ENDIF}
    // ajout BTP
    savcasemploi : boolean;
begin
//  sauvegarde la ligne en cours
Active_Ligne := G_NLIG.Row;
tn1 := TV_NLIG.Selected;
//  on verifie que la ligne n'est pas vide et que c'est une nomenclature
if tn1 <> nil then i_ind1 := ChercheLigne(TN1)
else i_ind1 := Active_Ligne;
if i_ind1 = 0 then exit;
if G_NLIG.Cells[SGN_Comp, i_ind1] = '' then Exit;
TobLND:=TOB(G_NLIG.Objects[SGN_NumLig, i_ind1]);
if (TTreeNode(G_NLIG.Objects[SGN_Comp, G_NLIG.Row]).count = 0)
   and (TobLND.GetValue('SOUSNOMEN')<>'X') then exit;
{$IFDEF BTP}
if TobLND.GetValue('BLO_TYPEARTICLE')='ARP' then exit;
{$ENDIF}
//  si c'est une nomenclature sans détail on ajoute une tob fille
if (TTreeNode(G_NLIG.Objects[SGN_Comp, G_NLIG.Row]).count = 0)
    and (TobLND.GetValue('SOUSNOMEN')='X') then
    if TobLND.Detail.Count<=0 then
       begin
{$IFDEF BTP}
       if Ouvrage then
          begin
          TobNew:=TOB.Create('LIGNEOUV',TobLND,-1);
          InsertionChampSupOuv (TOBNew,false);
          TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
          TobNew.putValue('BLO_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));
          end
       else
{$ENDIF}
       TobNew:=TOB.Create('LIGNENOMEN',TobLND,-1);
       TobNew.PutValue(prefixe+'_COMPOSE',TobLND.GetValue(prefixe+'_ARTICLE'));
       end;
// Ajout Btp
savcasemploi := CasEmploi_Open;
//  on cree et on ouvre la nouvelle fenetre
// FLigNomen := TFLigNomen.Create(Application);
{$IFDEF BTP}
if Ouvrage then
   begin
   if IsModeSelect then LastAction := SavAction else lastAction := Action;
   OkTrait := Entree_LigneOuv(XP,nil,TobLND,TobArticles,TOBLigPiece, TOBREpart,TobMetres,  True , InfoNiv+1,TobLND.GetValue(prefixe+'_NUMORDRE'),DEV,TOBLND.GetValue(prefixe+'_QUALIFQTEVTE'),valloc,Lastaction,GestionHt,false,ModeSaisieAch,false);
   if (not (action = taconsult)) and (OkTrait) then
      begin
      TOBLND.Putvalue ('BLO_DPA',valloc[0]);
      TOBLND.Putvalue ('BLO_DPR',valloc[1]);
      TOBLND.Putvalue ('BLO_PUHT',Valloc[2]);
      TOBLND.Putvalue ('BLO_PUHTDEV',pivottodevise(Valloc[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
      TOBLND.Putvalue ('BLO_PUTTC',Valloc[3]);
      TOBLND.Putvalue ('BLO_PUTTCDEV',pivottodevise(Valloc[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
      TOBLND.Putvalue ('BLO_PMAP',Valloc[6]);
      TOBLND.Putvalue ('BLO_PMRP',VAlloc[7]);
      TOBLND.Putvalue ('ANCPV',TOBLND.GetValue('BLO_PUHTDEV'));
      TOBLND.Putvalue ('ANCPA',TOBLND.GetValue('BLO_DPA'));
      TOBLND.Putvalue ('ANCPR',TOBLND.GetValue('BLO_DPR'));
            // Modif LS le 31/03/03
      TOBLND.Putvalue ('BLO_TPSUNITAIRE',VAlloc[9]);
      TOBLND.Putvalue ('GA_HEURE',VAlloc[9]);
      TOBLND.Putvalue ('GA_PAHT',VAlloc[0]);
      TOBLND.Putvalue ('GA_PVHT',VAlloc[2]);
      // --
      CalculMontantHtDevLigOuv (TOBLND,DEV);
      AfficheLaligne (i_ind1);
      G_NLIG.InvalidateCell(SGN_PvHT,i_ind1);
//      G_NLIG.Cells [SGN_PvHT,i_ind1] := TOBLND.GetValue('BLO_PUHTDEV');
//      G_NLIG.InvalidateCell(SGN_PvHT,i_ind1);
    	Data_modified := true;
      end;
   end else
{$ENDIF}
	begin
    Entree_LigneNomen(TobLND,TobArticles,True , InfoNiv+1,TobLND.GetValue(prefixe+'_NUMORDRE'), taModif);
    Data_modified := true;
	end;
// restitution contexte
CasEmploi_Open := savcasemploi;
if TobVide(TobLND) then TobLND.ClearDetail; //si pas de composants, on supprime la fille créée
bAnnule := false;
//  on fait le menage, on reinitialise les données de la nomenclature appellée
TOBTemp := TOB(G_NLIG.Objects[SGN_NumLig, i_ind1]);
AffecteLibelle(TOBTemp, prefixe+'_CODEARTICLE');
//  on reactualise les affichages
AfficheTVNLIG;
G_NLIG.Row := Active_Ligne;
TV_NLIG.Selected := TTreeNode(G_NLIG.Objects[SGN_Comp, G_NLIG.Row]);
// Modif BTP
{$IFDEF BTP}
if Ouvrage then
   begin
   CalculeOuvrageLoc;
   AfficheTotalisation;
   AfficheMontantLigne (i_ind1);
   end;
// ---
{$ENDIF}
end;

//=============================================================================
//  Clic Droit Option Tout Ouvrir
//=============================================================================
procedure TFLigNomen.PopTVNLIG_TClick(Sender: TObject);
begin
TV_NLIG.FullExpand;
end;

//=============================================================================
//  Clic Droit Option Tout Fermer
//=============================================================================
procedure TFLigNomen.PopTVNLIG_FClick(Sender: TObject);
begin
TV_NLIG.FullCollapse;
TV_NLIG.TopItem.Expand(False);
end;

//=============================================================================
//  Clic Droit Option Tout Fermer
//=============================================================================
procedure TFLigNomen.PopTV_NLIG_OClick(Sender: TObject);
begin
PopTV_NLIG_O.Checked := not (PopTV_NLIG_O.Checked);
end;



////////////////////////////////////////////////////////////////////////////////
//*************************** Gestion des boutons ****************************//
////////////////////////////////////////////////////////////////////////////////

Function TFLIgNomen.EnregistreOK : boolean ;
var TOBTempErr : TOB;
begin
	result := true;
  if (IsModeSelect) then Modenormal;
  if TobVide(TobLig) then
  begin
    MsgBox.Execute (10,Caption,'') ;
    Result := False;
    Exit;
  end;
//  Validation de la ligne en cours dans la Grid

  if not ValideLigne(G_NLIG.Row) then
  begin
    Result := false;
    Exit;
  end;

  if (G_NLIG.row = TobLig.detail.count) and (G_NLIG.cells[SGN_COMP,G_NLIG.row]='') then
  begin
    VideLaLigne(G_NLIG.row);
    G_NLIG.row:= G_NLIG.row -1;
  end;

  {$IFDEF BTP}
  CalculeOuvrageLoc;
  {$ENDIF}

  //  Vérification de la circularité
  TOBTempErr:=TOB.Create(TCOMPOSE.Caption,nil,-1);
  TOBTempErr.AddChampSup(prefixe+'_COMPOSE', True);
  TOBTempErr.PutValue(prefixe+'_COMPOSE', TCOMPOSE.Caption);
  if not Circularite(Compose.Text,TOBLig, TOBTempErr) then
  begin
    FNomenErr := TFNomenErr.Create(Application);
    //  Message nomenclature contient une reference circulaire
    AfficheTreeView(FNomenErr.TV_ERREUR, TOBTempErr, prefixe+'_COMPOSE');
    FNomenErr.ShowModal;
    FNomenERR.Free;
    G_NLIG.Col := SGN_Comp; G_NLIG.Row := 1;
    Result := false;
    TOBTempErr.Free;
    Exit;
  end;
  TOBTempErr.Free;
end;

procedure TFLigNomen.BValiderClick(Sender: TObject);
var indice : integer;
begin
// Modif BTP
//NextControl (Self,true);
  OnValidation := True;

  SendMessage(G_NLIG.Handle, WM_KeyDown, VK_TAB, 0);

  if not EnregistreOK then
  BEGIN
    Modalresult := 0 ;
    OnValidation := false;
    exit;
  end;
  CompleteTobArticles(TobLig);

  Indice := 0;
  repeat
    if TOBlig.Detail[Indice].getValue('BLO_CODEARTICLE')='' then
    Begin
      //Suppression des métrés associés à la ligne....
      TheMetreShare.SuppressionLigneMetreDoc(TOBlig);
      //
      TOBlig.Detail[Indice].free;
    end
    else
      inc(Indice);
  until indice >= TOBLig.detail.count;

  Data_Modified := false;

  bValide := true;

  Close;

end;

procedure TFLigNomen.bNewligneClick(Sender: TObject);
begin
//  LastKey positionné pour eviter le traitement de G_NLIGRowMoved
LastKey := VK_INSERT;
InsertLigne (G_NLIG.Row, False);
LastKey := 0;
bNewLigne.enabled := false;
end;

procedure TFLigNomen.bDelLigneClick(Sender: TObject);
begin
//  LastKey positionné pour eviter le traitement de G_NLIGRowMoved
//LastKey := VK_DELETE;
SupprimeLigne (G_NLIG.Row) ;
//LastKey := 0;
bNewLigne.enabled := true;
end;

procedure TFLigNomen.bNomencClick(Sender: TObject);
begin
TV_NLIGDblClick(Sender);
end;


procedure TFLigNomen.PopG_NLIG_AClick(Sender: TObject);
begin
G_NLIGDblClick(Sender);
end;

procedure TFLigNomen.PopG_NLIG_CClick(Sender: TObject);
var
    st_Article : string;
begin
st_Article := G_NLIG.Cells[SGN_Comp, G_NLIG.Row];
Entree_CasEmploi(['NOMENCLATURE=NON', st_Article], 2);
end;

procedure TFLigNomen.PopG_NLIG_IClick(Sender: TObject);
begin
bNewLigneClick(Sender);
end;

procedure TFLigNomen.PopG_NLIG_SClick(Sender: TObject);
begin
bDelLigneClick(Sender);
end;

procedure TFLigNomen.PopG_NLIG_NClick(Sender: TObject);
begin
TV_NLIGDblClick(Sender);
end;

procedure TFLigNomen.PopG_NLIG_DClick(Sender: TObject);
begin
PopG_NLIG_D.Checked := not (PopG_NLIG_D.Checked);
PANDIM.Visible := PopG_NLIG_D.Checked;
end;

{$IFDEF BTP}
procedure TFLigNomen.AfficheTotalisation;
var Total,TOTALU : T_valeurs;
begin
  Total := CalculSurTableau('+',TotalOuv ,ValPourcent);
  ToTalU := CalculSurTableau('+',VALUOUV ,ValUPourcent);
  TDEVHT.Caption := DEV.Symbole;
  TDEVTTC.Caption := DEV.Symbole;
  //
  Montantpa.Value := Total [0];
  Montantpr.value := Total [1];
  MontantHT.value := Total [2];
  MontantTTC.value := Total [3];
  if Total[0] <> 0 then Coeffg.value := Total[1]/Total[0]
								   else Coeffg.value := 0;
  MontantFG.value := Total[1]-Total[0];
  if Total[1] <> 0 then CoefMarg.value := arrondi(devisetopivotEx( Total[2],DEV.taux,DEV.Quotite,V_PGI.okdecP )/Total[1],4)
								   else CoefMarg.value := 0;
  MontantMarg.value := devisetopivot (Total[2],DEV.Taux,DEV.Quotite ) -Total[1];
  TOTALTPS.value := Total[9];
  //
  if QteDetail.value <> 0 then
  begin
    MontantpaU.Value := ToTalU [0];
    MontantprU.value := ToTalU [1];
    MontantHTU.value := ToTalU [2];
    MontantTTCU.value := ToTalU [3];
    MontantFGU.value := ToTalU[1]-ToTalU[0];
    MontantMargU.value := devisetopivot (ToTalU[2],DEV.Taux,DEV.Quotite ) -ToTalU[1];
    TPSUNITAIRE.value := ToTalU[9];
  end;
  if QteDudetail = 1 then
  begin
	  Pvalo.visible := false;
  end else
  begin
	  Pvalo.visible := true;
  end;
  //
end;

procedure TFLigNomen.afficheMontantLigne(ou:Integer);
var valuePv : double;
    TOBL : TOB;
    MontantPAFG,MontantPAFC,MOntantPAFR : double;
    MontantFG,MontantFC,MontantFR : double;
begin
  TOBL := TOB(G_NLIG.Objects[SGN_NumLig, Ou]);
	//
  if TOBL <> nil then
  begin
    MontantPAFG := TOBL.GetValue('BLO_MONTANTPAFG');
    MontantPAFC := TOBL.GetValue('BLO_MONTANTPAFC');
    MontantPAFR := TOBL.GetValue('BLO_MONTANTPAFR');
    MontantFG := TOBL.GetValue('BLO_MONTANTFG');
    MontantFC := TOBL.GetValue('BLO_MONTANTFC');
    MontantFR := TOBL.GetValue('BLO_MONTANTFR');
  end;
  //
  TGA_FOURNPRINC.Caption  := '';
  TGA_PRESTATION.Caption  := '';
  NEDEV.Caption := DEV.Symbole;
  NEPA.Value := 0;
  NEPR.Value := 0;
  NEPVHT.Value := 0;
  NEPVTTC.Value := 0;
  NEPVFORFAIT.Value := 0;
  NEMTACH.Value := 0;
  NEMTPR.Value := 0;
  NETPS.value := 0;
  if TOBL <> nil then
  begin
    TGA_FOURNPRINC.caption := TOBL.GetValue ('TGA_FOURNPRINC');
    TGA_PRESTATION.caption := TOBL.GetValue ('TGA_PRESTATION');
    NEPA.value := TOBL.getvalue('BLO_DPA');
    NEPR.value := TOBL.getvalue('BLO_DPR');
    NETPS.Value := TOB(G_NLIG.Objects[SGN_Numlig, Ou]).getValue('BLO_TPSUNITAIRE') *
    							 TOB(G_NLIG.Objects[SGN_Numlig, Ou]).getValue('BLO_QTEFACT');
    CalculMontantHtDevLigOuv (TOBL,DEV);
    if GestionHt then
    BEGIN
      NEPVHT.value := TOBL.getvalue('BLO_PUHT');
      NEPVTTC.value := TOBL.getvalue('BLO_PUTTC');
      ValuepV := TOBL.getvalue('BLO_PUHTDEV');
      NEPVFORFAIT.value := TOBL.getvalue('BLO_MONTANTHTDEV');
      TLigPvSaisie.Caption := TraduireMemoire('Montant H.T');
    END ELSE
    BEGIN
      NEPVHT.value := TOBL.getvalue('BLO_PUTTC');
      NEPVTTC.value := TOBL.getvalue('BLO_PUHT');
      ValuePV:= TOBL.getvalue('BLO_PUTTCDEV');
      NEPVFORFAIT.value := TOBL.getvalue('BLO_MONTANTTTCDEV');
      TLigPvSaisie.Caption := TraduireMemoire('Montant T.T.C');
    END;
    NEMTPR.value := TOBL.getvalue('BLO_MONTANTPA')+TOBL.getvalue('BLO_MONTANTFG')+TOBL.getvalue('BLO_MONTANTFC')+TOBL.getvalue('BLO_MONTANTFR');
    if TOBL.detail.count >0  then
    begin
    	(*
      if NEPA.value <> 0 then NECOEFFG.value := NEPR.value/ NEPA.value
                         else NECOEFFG.value := 0;
      if NEPR.value <> 0 then NECOEFFG.value := devisetopivot (ValuePv,DEV.taux,Dev.quotite)/ NEPR.value
                         else NECOEFFG.value := 0;
    	*)
      if (MontantPAFG <> 0 ) and (MontantFG<>0) then
      begin
        NECOEFFG.value := arrondi((MontantFG+MontantPAFG)/(MontantPAFG),4);
      end else
      begin
        NECOEFFG.value := 1;
      end;

      if (MontantPAFC<>0) and (MontantFC<>0) then
      begin
        NECOEFFC.value := arrondi((MontantPAFC+MontantFG+MontantFC)/ (MontantPAFC+MontantFG),4);
      end else
      begin
        NECOEFFC.value := 0;
      end;
      if (MontantPAFR<>0) and (MontantFR<>0) then
      begin
        NECOEFFR.value := arrondi((MontantPAFR+MontantFG+MontantFC+MontantFR)/ (MontantPAFR+MontantFG+montantFC),4)
      end else
      begin
        NECOEFFR.value := 0;
      end;
      if NEPR.value <> 0 then NECOEFPRPV.value := devisetopivot (ValuePv,DEV.taux,Dev.quotite)/ NEPR.value
                         else NECOEFPRPV.value := 0;
    end else
    begin
      NECOEFFG.value := (TOBL.GetValue('BLO_COEFFG')+1);
      NECOEFFC.value := (TOBL.GetValue('BLO_COEFFC')+1);
      NECOEFFR.value := (TOBL.GetValue('BLO_COEFFR')+1);
      NECOEFPRPV.value := TOBL.GetValue('BLO_COEFMARG');
    end;
  end else
  begin
    if NEPA.value <> 0 then NECOEFFG.value := NEPR.value/ NEPA.value
                       else NECOEFFG.value := 0;
    if NEPR.value <> 0 then NECOEFPRPV.value := devisetopivot (ValuePv,DEV.taux,Dev.quotite)/ NEPR.value
                       else NECOEFPRPV.value := 0;
  end;
  if TOBL <> nil then
  begin
    NEMTACH.value := TOBL.getvalue('BLO_MONTANTPA');
//    NEMTPR.value := TOBL.getvalue('BLO_MONTANTPR');
  end;
end;
{$ENDIF}

procedure TFLigNomen.PopG_NLIG_LClick(Sender: TObject);

  procedure AppliqueChangementDetOuv (TOBOL,TOBL : TOB);
  begin
    changeTaxeOuvInterne (TOBL,TOBL.GetValue('BLO_FAMILLETAXE1'),TOBL.GetValue('BLO_FAMILLETAXE2'),
                               TOBL.GetValue('BLO_FAMILLETAXE3'),TOBL.GetValue('BLO_FAMILLETAXE4'),
                               TOBL.GetValue('BLO_FAMILLETAXE5'));
  end;



{$IFDEF BTP}
var CodeAffaire,CodeTiers : string;
    TOBL,TOBPIece,TOBOL,TOBART : TOB;
{$ENDIF}
begin
{$IFDEF BTP}
TheTOB:=TOB(G_NLIG.objects[SGN_Numlig,G_NLIG.Row]);
TOBArt := TOB(G_NLIG.objects[SGN_idComp,g_nlig.row]);
// sauvegarde de la ligne pour lancer les changements ulterieurement
TOBOL := TOB.Create ('LIGNEOUV',nil,-1);
InsertionChampSupOuv (TOBOL,false);
TOBOL.dupliquer (TheTOB,false,true);
//
if TheTob = nil then exit;
TheTOB.Data := TOBLigPiece;
TOBLigPiece.Data := TOBrepart;
TOBPiece := TOBLigPiece.Parent ;
CodeAffaire := TOBLigPiece.GetValue('GL_AFFAIRE');
CodeTiers := TOBLigPiece.GetValue('GL_TIERS');
AGLLanceFiche ('BTP','BTCOMPLLIOUV','','PRIXBLOQUE='+TOBLigPiece.GetValue('GLC_NONAPPLICFRAIS')+';AFFAIRE='+CodeAffaire+';TIERS=',ActionToString(Action));
TOBL:=TOB(G_NLIG.objects[SGN_Numlig,G_NLIG.Row]);
//
if (TOBL.getValue('BLO_FAMILLETAXE1')<>TOBOL.getValue('BLO_FAMILLETAXE1')) or
   (TOBL.getValue('BLO_FAMILLETAXE2')<>TOBOL.getValue('BLO_FAMILLETAXE2')) or
   (TOBL.getValue('BLO_FAMILLETAXE3')<>TOBOL.getValue('BLO_FAMILLETAXE3')) or
   (TOBL.getValue('BLO_FAMILLETAXE4')<>TOBOL.getValue('BLO_FAMILLETAXE4')) or
   (TOBL.getValue('BLO_FAMILLETAXE5')<>TOBOL.getValue('BLO_FAMILLETAXE5')) then
begin
  AppliqueChangementDetOuv (TOBOL,TOBL);
end;
//
if (TOBLigPiece.getValue('GL_BLOQUETARIF')='-') and (TOBL.getValue('BLO_COEFMARG')<>TOBOL.getValue('BLO_COEFMARG')) then
begin
  AppliqueCoefMargDetail (TOBLigPiece,TOBL,TOBL.GetValue('BLO_COEFMARG'),dev);
end;
//
if (TOBL.GetValue('BLO_NONAPPLICFRAIS') <> TOBOL.GetValue('BLO_NONAPPLICFRAIS')) or
   (TOBL.GetValue('BLO_NONAPPLICFC') <> TOBOL.GetValue('BLO_NONAPPLICFC')) or
   (TOBL.GetValue('BLO_NONAPPLICFG') <> TOBOL.GetValue('BLO_NONAPPLICFG')) then
begin
  AppliqueChangeFraisDetOuv (TOBPiece,TOBL,TOBL.getValue('BLO_NONAPPLICFRAIS'),
                             TOBL.getValue('BLO_NONAPPLICFC'),TOBL.getValue('BLO_NONAPPLICFG'));
end;
//
CalculeLigneAcOuv (TOBL,DEV,TOBLigPiece.GetValue('GLC_NONAPPLICFRAIS')='X',TOBART);
G_NLIG.Cells[SGN_Qte,G_NLIG.Row] := TOBL.GetValue('BLO_QTEFACT');
if GestionHt then
   BEGIN
   G_NLIG.Cells[SGN_Pvht,G_NLIG.Row] := TOBL.GetValue('BLO_PUHTDEV');
   CalculeLigneHTOuv(TOB(G_NLIG.Objects[SGN_NumLig, G_NLIG.Row]),TOBPiece,DEV);
   END ELSE
   BEGIN
   G_NLIG.Cells[SGN_Pvht,G_NLIG.Row] := TOBL.GetValue('BLO_PUTTCDEV');
   CalculeLigneTTCOuv(TOB(G_NLIG.Objects[SGN_NumLig, G_NLIG.Row]),TOBPiece,DEV);
   END;
TheTOB := nil;
TOBLigPiece.data := nil;
G_NLIG.Cells[SGN_UV,G_NLIG.Row] := TOBL.GetValue('BLO_QUALIFQTEVTE');
CalculeOuvrageLoc ;
AfficheTotalisation;
G_NLIG.InvalidateRow (G_NLIG.Row);
FreeAndNil(TOBOL);
{$ENDIF}
end;

procedure TFLigNomen.BTypeArticleClick(Sender: TObject);
var TOBOuv : TOB;
begin
TOBOuv := TOB(G_NLIG.Objects[SGN_NumLig,G_NLIG.Row]);
if TOBOuv <> nil then
   BEGIN
{$IFDEF BTP}
   if TOBOuv.GetValue('BLO_TYPEARTICLE') = 'OUV' then TV_NLIGDblClick (Sender);
{$ENDIF}
   END;
end;

function TFLigNomen.AutoriseEntreeSelection (ou : Integer; var cancel : boolean): boolean;
begin
// Pour l'instant rien n'interdit la sélection d'une ligne dans le détail d'ouvrage
// pour plus tard
result := true;
end;

procedure TFLigNomen.AfficheSelectionLigne (ou: integer);
begin
// ou = numero de ligne à sélectionner
// plus tard
end;

procedure TFLigNomen.ModeNormal;
begin
IsModeSelect := false;
Action := Savaction;
SavAction := Action;
affecteGrid (G_NLIG,action);
LmodeIns.Caption := '';
G_NLIG.AllSelected := false;
G_NLIG.Multiselect := false;
G_NLIG.popupmenu := PopupG_NLIG;
if (SGN_TYPA <> -1) and (action <> taconsult) then
   BEGIN
   G_NLIG.ColLengths[SGN_TYPA] := -1;
   END;
end;

{$IFDEF BTP}
procedure TFligNomen.ModeSelection;
begin
IsModeSelect := true;
savAction := action;
action := TaConsult;
LmodeIns.Caption := 'Mode sélection';
affecteGrid (G_NLIG,action);
G_NLIG.Multiselect := true;
if SGN_TYPA <> -1 then
   BEGIN
   G_NLIG.ColLengths[SGN_TYPA] := 27;
   END;
G_NLIG.popupmenu := Popup_SELECTION;
end;
{$ENDIF}

procedure TFLigNomen.CopieDonnee;
var Indice : Integer;
    Mcopy : TstringList;
    TOBLoc,TOBSEL : TOB;
begin
Clipboard.clear;
TOBSEL := TOB.create ('DONNEAINS',nil,-1);
TOBSEL.AddChampSupValeur ('TYPE','LIGNEOUV');
Mcopy := TstringList.Create;
for Indice := 1 to G_NLIG.RowCount do
    begin
    if (G_NLIG.IsSelected (Indice) ) then
       begin
       TOBLOC := TOB.Create ('LIGNEOUV',TOBSEL,-1);
       TOBLOC.Dupliquer (TOB(G_NLIG.objects[SGN_NumLig,Indice]),true,true);
       end;
    end;
ExportTob(TOBSel,Mcopy);
Clipboard.clear;
MTransfert.Lines :=Mcopy;
MTransfert.SelectAll;
Mtransfert.CopyToClipboard;
Mcopy.Clear ;
Mcopy.Free;
TOBSel.free;
end;

procedure TFLigNomen.MSelCopierClick(Sender: TObject);
begin
CopieDonnee;
end;

procedure TFLigNomen.ExportTOB(TOBLoc : TOB; Mcopy:TstringList);
begin
MCopy.Add(TobLoc.SaveToBuffer (true,true,''));
end;

procedure TFLigNomen.EnregistreDonne;
var CodeEnreg,LibEnreg : string;
    Indice : Integer;
    Mcopy : TstringList;
    TOBLoc,TOBSEL,TOBENREG : TOB;
begin
if DemandeNomEnreg ('LIGNEOUV',CodeEnreg,LibEnreg) then
   begin
   TOBSEL := TOB.create ('DONNEAINS',nil,-1);
   TOBSEL.AddChampSupValeur ('TYPE','LIGNEOUV');
   Mcopy := TstringList.Create;
   for Indice := 1 to G_NLIG.RowCount do
       begin
       if (G_NLIG.IsSelected (Indice) ) then
          begin
          TOBLOC := TOB.Create ('LIGNEOUV',TOBSEL,-1);
          TOBLOC.Dupliquer (TOB(G_NLIG.objects[SGN_NumLig,Indice]),true,true);
          end;
       end;
   ExportTob(TOBSel,Mcopy);
   MTransfert.Lines :=Mcopy;
   MTransfert.SelectAll;
   //
   TOBENREG := TOB.create ('BMEMORISATION',nil,-1);
   TOBENreg.PutValue('BMO_TYPEMEMO','LIGNEOUV');
   TOBENreg.PutValue('BMO_CODEMEMO',CodeEnreg);
   TOBENreg.PutValue('BMO_LIBMEMO',LibEnreg);
   TOBENREG.PutValue('BMO_MEMO',Mtransfert.Text);
   TRY
      TOBENreg.InsertOrUpdateDB (true);
   finally
      TOBEnreg.free;
      Mcopy.Clear ;
      Mcopy.Free;
      TOBSel.free;
      end;
   end;
end;

procedure TFLigNomen.MSelEnregistrerClick(Sender: TObject);
begin
enregistreDonne;
end;

function TFLigNomen.RecupTob(Atob:TOB): integer;
begin
TOBImp := ATob;
end;

procedure TFligNomen.AfficheLaligne (Arow : integer);
Var TobTemp,TOBArt : TOB;
begin
TOBTemp :=TOB(G_NLIG.Objects[SGN_NumLig, ARow]);
TOBArt := Tob(G_NLIG.Objects[SGN_IdComp, ARow]);
if Ouvrage then G_NLIG.Cells[SGN_Qte, ARow] := TobTemp.GetValue('BLO_QTEFACT')
           else G_NLIG.Cells[SGN_Qte, ARow] := TobTemp.GetValue('GLN_QTE');
G_NLIG.Cells[SGN_Comp, ARow] := TOBArt.GetValue('GA_CODEARTICLE');
G_NLIG.Cells[SGN_IdComp, ARow] := TOBArt.GetValue('GA_ARTICLE');
G_NLIG.Cells[SGN_Lib, ARow] := TOBTemp.GetValue('BLO_LIBELLE');
if gestionHt then G_NLIG.cells[SGN_PVHT,Arow] := StrF00(TobTemp.getvalue('BLO_PUHTDEV'),V_PGI.OkdecP)
             else G_NLIG.cells[SGN_PVHT,Arow] := StrF00(TobTemp.getvalue('BLO_PUTTCDEV'),V_PGI.OkdecP);
G_NLIG.cells[SGN_UV,Arow] := TobTemp.getvalue('BLO_QUALIFQTEVTE');
G_NLIG.InvalidateRow (Arow);
TOBTemp.PutLigneGrid (G_NLIG,Arow,false,false,LesColNLIG);
end;

procedure TFLigNomen.positionneValeurPV (TOBL : TOB);
var Indice : Integer;
    TOBS : TOB;
begin
TOBL.PutValue('ANCPV', Valeur(TOBL.GetValue('ANCPV')));
if TOBL.detail.count > 0 then
   begin
   for indice := 0 to TOBL.detail.count -1 do
       begin
       TOBS := TOBL.detail[Indice];
       positionneValeurPV (TOBS);
       end;
   end;
end;


procedure TFLigNomen.AjouteLigne (TOBLoc :tob;Arow:integer);
var TOBL,TOBArt: TOB;
    tn1 : TTreenode;
    Q : TQuery;
    Larticle : string;
begin
G_NLIG.InsertRow (Arow);
//G_NLIG.Row := ARow;
if ARow > 1 then G_NLIG.Cells[0, ARow] := IntToStr(StrToInt(G_NLIG.Cells[0, ARow - 1]) + 1)
            else G_NLIG.Cells[0, ARow] := IntToStr(1);
InitialiseLigne (ARow) ;
AjouteTobLigne(ARow);
//
TOBL :=TOB(G_NLIG.Objects[SGN_NumLig, ARow]);
TOBL.dupliquer (TOBLoc,true,true);
positionneValeurPV (TOBL);
TobL.putvalue('BLO_QTEDUDETAIL',QteduDetail);
TOBL.PutValue(prefixe+'_NUMORDRE',StrToInt(G_NLIG.Cells[0, ARow])); // numéro de ligne
// gestion de la TOB Article
G_NLIG.Objects[SGN_IdComp, ARow] := TOB.Create('ARTICLE',nil,ARow);
Larticle := TOBL.GetValue('BLO_CODEARTICLE');
Q := OPenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
           		'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
           		'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE GA_CODEARTICLE="'+
              Larticle+'" AND GA_STATUTART <> "DIM" ',True,-1, '', True);

Tob(G_NLIG.Objects[SGN_IdComp, ARow]).SelectDB('',Q);
TOBArt := Tob(G_NLIG.Objects[SGN_IdComp, ARow]);
InitChampsSupArticle (TOBART);

//
ferme (Q);
RepriseDonneeArticle (TOBLOC,TOBArt);
ReajusteLigneParDoc (TobLigPiece,TOBLOc,DEV,InfoNiv);
AfficheLaligne (Arow);
if (ARow = 1) then
   G_NLIG.Objects[SGN_Comp, ARow] := TV_NLIG.Items.AddChildFirst(TV_NLIG.Items[0], G_NLIG.Cells[SGN_Comp, ARow])
else if (G_NLIG.Cells[SGN_Comp, ARow + 1] = '') or (ARow = G_NLIG.RowCount - 1) then
   G_NLIG.Objects[SGN_Comp, ARow] := TV_NLIG.Items.AddChild(TV_NLIG.Items[0], G_NLIG.Cells[SGN_Comp, ARow])
else
    begin
    tn1 := TTreeNode(G_NLIG.Objects[SGN_Comp, ARow - 1]);
    tn1 := tn1.GetNext;
    G_NLIG.Objects[SGN_Comp, ARow] := TV_NLIG.Items.Insert(tn1, G_NLIG.Cells[SGN_Comp, ARow]);
    TTreeNode(G_NLIG.Objects[SGN_Comp, ARow]).Text := G_NLIG.Cells[SGN_Comp, ARow];
    end;
end;

function TFLigNomen.IsTypeTOB : boolean;
begin
result := false;
MTransfert.SelectAll;
Mtransfert.Clear;
MTransfert.PasteFromClipboard;
if copy(Mtransfert.Text,1,4)='$|0|' then Result := true;
Mtransfert.clear;
end;

procedure TFLigNomen.CollageDonnee;
begin
MTransfert.SelectAll;
Mtransfert.Clear;
MTransfert.PasteFromClipboard;
ChargeDonnee;
MselCopier.Enabled := true;
end;

procedure TFLigNomen.MSelCollerClick(Sender: TObject);
begin
if (not IsTypeTOB) or  (not IsModeSelect) then
   begin
   SendMessage (G_NLIG.InplaceEditor.Handle,WM_PASTE,0,0);
   exit;
   end;
CollageDonnee;
end;

procedure TFLigNomen.POPG_NLIG_CollerClick(Sender: TObject);
begin
if (not IsTypeTOB)  then
   begin
   SendMessage (G_NLIG.InplaceEditor.Handle,WM_PASTE,0,0);
   exit;
   end;
CollageDonnee;
end;


Function TFLigNomen.ZoneAccessible ( ACol,ARow : Longint) : boolean ;
var TOBLIG1 : TOB;
    TypeArticle : string;
BEGIN
  result := true;
  if Arow > TOBLig.detail.count then begin result := true;exit; end;
  TOBLIg1 := TOB(G_NLIG.objects [SGN_NumLig,Arow]);
  if TOBLIG1 = nil then exit;
  {$IFDEF BTP}
  TypeArticle := TOBLIG1.GetValue ('BLO_TYPEARTICLE');
  {$ENDIF}
	if not G_NLIG.ColEditables [Acol] then
  begin
  	result := false;
    exit;
  end;

  if G_NLIG.ColLengths[ACol]<0 then
   BEGIN
   result:=false ;
   Exit ;
   END ;
  if (Acol = SGN_Comp) and (TOBLIG1.GetValue('BLO_ARTICLE') <> '') then BEGIN Result:=False; exit; END;
  if (TypeArp) and (Acol = SGN_Comp) then BEGIN Result:=False; exit; END;
  if (SGN_MONTANT >= 0) and (Acol = SGN_MONTANT) then
  begin
     result := false;
     exit;
  end;
  if (SGN_MONTANTHTDEV >= 0) and (Acol = SGN_MONTANTHTDEV) then
  begin
     result := false;
     exit;
  end;



  if (SGN_UV <> -1) and (Acol = SGN_UV) then
   begin
   result := false;
   exit;
   end;
  {$IFDEF BTP}
  if TypeArticle <> 'ARP' then
   begin
   if (Acol = SGN_PREST) or (Acol = SGN_TPS) then BEGIN Result := false; exit; END;
   if (TypeArticle = 'OUV') and ((Acol = SGN_FOURN) or (Acol = SGN_PAHT )) then BEGIN Result := false; exit; END;
   end;
  if (Acol = SGN_PVHT) and (TOBLIgPiece.getValue('GL_BLOQUETARIF')<> '-') then
  begin
	result := false;
  exit;
  end;
  {$ENDIF}
  if ((Ouvrage) and (G_NLIG.objects [SGN_NumLig,Arow]<> nil) and
   (SGN_PVHT <> -1) and (ACol = SGN_PVHT) and
   (TypeArticle = 'POU'))  then
   begin
   result := false;
    exit;
  end;
  if (SGN_PVHT <> -1) and (ACol = SGN_PVHT) and (TOBLigPiece.getBoolean('GL_GESTIONPTC')) then
  begin
    result := false;
    exit;
  end;
  if (Acol = SGN_UNITESAIS) then
  begin
    result := false;
    exit;
  end;
  if (Acol = SGN_PERTE) and ((POS(TOBLIG1.GetString('BNP_TYPERESSOURCE'),'SAL;INT;ST') >0) or (IsOuvrage(TOBLIG1))) then
  begin
    result := false;
    exit
  end;
  if (Acol = SGN_RENDEMENT) and ((POS(TOBLIG1.GetString('BNP_TYPERESSOURCE'),'SAL;INT;ST') <=0) or (IsOuvrage(TOBLIG1))) then
  begin
    result := false;
    exit
  end;
END ;

procedure TFLigNomen.ZoneSuivanteOuOk ( Var ACol,ARow : Longint ; Var Cancel : boolean ) ;
Var Sens,ii,Lim : integer ;
    OldEna,ChgLig,ChgSens  : boolean ;
BEGIN
OldEna:=G_NLIG.SynEnabled ; G_NLIG.SynEnabled:=False ;
if (G_NLIG.col > Acol) and (G_NLIG.row >= Arow) and (G_NLIG.cells[SGN_Comp,Arow]='') then
   begin
   Acol := SGN_COMP;
   G_NLIG.synenabled := Oldena;
   cancel := true;
   exit;
   end;
Sens:=-1 ; ChgLig:=(G_NLIG.Row<>ARow) ; ChgSens:=False ;
if G_NLIG.Row>ARow then Sens:=1 else if ((G_NLIG.Row=ARow) and (ACol<G_NLIG.Col)) then Sens:=1 ;
ACol:=G_NLIG.Col ; ARow:=G_NLIG.Row ; ii:=0 ;
While Not ZoneAccessible(ACol,ARow)  do
   BEGIN
   Cancel:=True ; inc(ii) ; if ii>500 then Break ;
   if Sens=1 then
      BEGIN
      // Modif BTP
      //Lim:=G_NLIG.RowCount ;
      Lim:=TobLig.detail.count +1;
      // ---
      if ((ACol=G_NLIG.ColCount-1) and (ARow>=Lim)) then
         BEGIN
         if ChgSens then Break else BEGIN Sens:=-1 ; Continue ; ChgSens:=True ; END ;
         END ;
//      if ChgLig then BEGIN ACol:=G_NLIG.FixedCols-1 ; ChgLig:=False ; END ;
//      if ACol<G_NLIG.ColCount-1 then Inc(ACol) else BEGIN Inc(ARow) ; ACol:=G_NLIG.FixedCols ; END ;
      if ChgLig then BEGIN ACol:=SGN_Comp ; ChgLig:=False ; END ;
      if ACol<G_NLIG.ColCount-1 then
      begin
        Inc(ACol);
      end else
      BEGIN
        Inc(ARow) ;
        if (Arow > G_NLIG.RowCount -1) and (TypeArp) then dec(Arow);
        ACol:=SGN_Comp ;
      END ;

      END else
      BEGIN
      if ((ACol=G_NLIG.FixedCols) and (ARow=1)) then
         BEGIN
         if ChgSens then Break else BEGIN Sens:=1 ; Continue ; END ;
         END ;
//      if ChgLig then BEGIN ACol:=G_NLIG.ColCount ; ChgLig:=False ; END ;
      if ChgLig then BEGIN ACol:=SGN_COMP ; Sens := 1; ChgLig:=False ; continue ; END ;
      if ACol>G_NLIG.FixedCols then Dec(ACol) else BEGIN Dec(ARow) ; ACol:=G_NLIG.ColCount-1 ; END ;
      END ;
   END ;
G_NLIG.SynEnabled:=OldEna ;
END ;

procedure TFLigNomen.POPG_NLIG_CopierClick(Sender: TObject);
begin
SendMessage (G_NLIG.InplaceEditor.Handle,WM_COPY,0,0);
end;

procedure TFLigNomen.Rappeler1Click(Sender: TObject);
begin
RappellerMemo;
end;

procedure TFLigNomen.RappellerMemo;
var Retour : string;
    TOBMemo : TOB;
    QQ :TQuery;
begin
Retour := AGLLanceFiche ('BTP','BTMEMORIS_MUL','BMO_TYPEMEMO=LIGNEOUV','','ACTION=CONSULTATION');
if Retour <> '' then
   begin
   TOBMemo := TOB.create ('BMEMORISATION',nil,-1);
   QQ := OpenSQL('Select BMO_MEMO from BMEMORISATION Where BMO_TYPEMEMO="LIGNEOUV" AND' +
             ' BMO_CODEMEMO="'+Retour+'"',True,-1, '', True) ;
   TOBMemo.SelectDB ('',QQ);
   ferme(QQ);
   MTransfert.SelectAll;
   MTransfert.Clear;
   MTransfert.text :=  TOBMemo.getValue('BMO_MEMO');
   TobMemo.Free;
   ChargeDonnee;
   end;
end;

procedure TFLigNomen.ChargeDonnee;
var TOBLOC: TOB;
    Indice : Integer;
begin
TOBLoadFromBuffer (Mtransfert.Text,RecupTob);
if TOBImp.getValue('TYPE') = 'LIGNEOUV' then
   BEGIN
   For Indice :=0 to TOBImp.detail.count -1 do
       begin
       TOBLoc := TOBImp.detail[Indice];
       AjouteLigne (TOBLoc,G_NLIG.row);
       end;
   end;
Renumerote;
G_NLIG.refresh;
{$IFDEF BTP}
if Ouvrage then
   begin
   CalculeOuvrageLoc;
   AfficheTotalisation;
   end;
// ---
{$ENDIF}
AfficheTVNLIG;
TOBIMP.free;
end;

procedure TFLigNomen.G_NLIGFlipSelection(Sender: TObject);
begin
//
end;

procedure TFLigNomen.G_NLIGBeforeFlip(Sender: TObject; ARow: Integer;
  var Cancel: Boolean);
var TOBL : TOB;
begin
if G_NLIG.Cells [SGN_Numlig,Arow] = '' then
   begin
   cancel := true;
   exit;
   end;
{$IFDEF BTP}
   TOBL := TOB(G_NLIG.objects[SGN_Numlig,Arow]);
   if TOBL.GetValue('BLO_TYPEARTICLE')='POU' then
   begin
   cancel := true;
   exit;
   end;
{$ELSE}
  cancel := true;
{$ENDIF}
end;

procedure TFLigNomen.BInfosLigneClick(Sender: TObject);
begin
ToolBarDesc.Visible := BinfosLigne.Down;
end;

procedure TFLigNomen.ToolBarDescClose(Sender: TObject);
begin
BInfosLigne.Down := false;
end;

procedure TFLigNomen.G_NLIGEnter(Sender: TObject);
var Arow,Acol: integer;
    cancel : boolean;
begin
Arow := 1; Acol := SGN_Comp;
G_NLIGRowEnter(Sender, 1, Cancel, False);
G_NLIGCellEnter (sender,Acol,Arow,cancel);
G_NLig.row := Arow;
G_NLIG.col := Acol;
Cell_Text := G_NLIG.Cells[G_NLIG.Col, G_NLIG.Row];
end;

{$IFDEF BTP}
procedure TFLigNomen.definiBarTree;
begin
ToolBarTree.parent := self;
ToolBartree.caption := 'Structure de l''ouvrage';
ToolBarTree.BorderStyle := bsSingle;
ToolBarTree.clientAreaHeight := TV_NLIG.Height ;
ToolBarTree.clientAreaWidth := TV_NLIG.Width;
ToolBarTree.ClientHeight := ToolBarTree.clientAreaHeight;
ToolBarTree.clientWidth := ToolBarTree.clientAreaWidth;
ToolBarTree.DragHandleStyle := dhDouble;
ToolBarTree.HideWhenInactive := True;
ToolBarTree.fullsize := false;
ToolBarTree.Resizable := true;
ToolBartree.Height := TV_NLIG.Height + 200;
ToolBartree.Width := TV_NLIG.Width + 10;
ToolBarTree.top := G_NLIG.Top + ((height - ToolBartree.height) div 2);
ToolBarTree.left := G_NLIG.left + ((width - ToolBartree.width) div 2);
TV_NLIG.Parent := ToolBarTree;
ToolBarTree.OnClose := ToolBarTreeClose;
ToolBarTree.visible := false;
end;

procedure TFLigNomen.ToolBarTreeClose (Sender : Tobject);
begin
BArborescence.Down := false;
end;

procedure TFLigNomen.AssigneEvenementBtn;
begin
BArborescence.OnClick := BArborescenceClick;
BInfosLigne.OnClick := BInfosLigneClick;
end;

procedure TFLigNomen.BArborescenceClick (Sender : TObject);
begin
ToolBarTree.visible := BArborescence.Down;
end;
{$ENDIF}


procedure TFLigNomen.ZoomArticle;
Var Coord : TRect;
{$IFDEF BTP}
    SWhere,STChamps : string;
    TOBPiece,TOBLig,TOBART : TOB;
    fournisseur : string;
{$ENDIF}
begin
  Coord := G_NLIG.CellRect (G_NLIG.Col, G_NLIG.Row);
  ART.Parent := G_NLIG;
  ART.Top := Coord.Top;
  ART.Left := Coord.Left;
  ART.Width := 3; ART.Visible := False;
  ART.DataType := 'GCARTICLE';
  ART.Text := G_NLIG.Cells[G_NLIG.Col,G_NLIG.Row];
{$IFDEF BTP}
  TOBPiece := TOBLigPiece.Parent ;
  sWhere:=FabricWhereNatArt(TOBPiece.getValue('GP_NATUREPIECEG'),'','') ;
  StChamps:='GA_CODEARTICLE='+ART.Text+';XX_WHERE=AND '+sWhere ;
  DispatchRecherche(ART, 1, '', stchamps, '');
{$ELSE}
  DispatchRecherche(ART, 1, '', 'GA_CODEARTICLE=' + ART.Text, '');
{$ENDIF}
  if ART.Text <> '' then G_NLIG.Cells[G_NLIG.Col,G_NLIG.Row]:= copy(ART.Text, 0, Length(ART.Text) - 1);
  FormateZoneSaisie (G_NLIG.Col, G_NLIG.Row);
end;


procedure TFLigNomen.Appel_XLClick(Sender: TObject);
var TOBO : TOB;
    result : double;
begin

  TOBO := TOB(G_NLIG.Objects[SGN_NumLig, G_NLIG.Row]);

  Result := 0;

  Result := ElipsisMetreClick(TOBO);

  //if result <> 0 then        //FV1 - 30/09/2016
  //begin
      TOBO.putValue('BLO_QTEFACT',result);
      TOBO.PutLigneGrid (G_NLIG, G_NLIG.row,False, False, LesColNLIG);
      if G_NLIG.col = SGN_Qte then
      begin
       Cell_Text := G_NLIG.Cells[SGN_QTE,G_NLIG.Row];
      end;
      ValideLigne(G_NLIG.row);
      CalculeOuvrageLoc;
      AfficheTotalisation;
      AfficheMontantLigne (G_NLIG.row);
      AffectePied(G_NLIG.row);
  //END;

end;


procedure TFLigNomen.FormCloseQuery(Sender: TObject;var CanClose: Boolean);
var w_rep : word;
begin
{$IFNDEF BTP}
  if not bValide then
  begin
    if TobVide(TobSauv) then
    begin
      MsgBox.Execute (11,Caption,'');
//      Action := caNone;
      exit;
    end else bAnnule := true;
  end;
{$ELSE}
	bannule := false;
	if not bvalide then Bannule := true;
{$ENDIF}
  if Data_Modified then
  begin
    //  Demande de validation
    w_rep := MsgBox.Execute (0,Caption,'');
    case w_rep of
        mrYes    : begin
                    // BValiderClick(Sender);
                     if not EnregistreOK then canclose := false;
                     bAnnule := false;
                   end;
        mrCancel : begin
                     canclose := false;
                     bAnnule := false;
                   end;
    end;
  end;
//
end;

procedure TFLigNomen.AppliqueQteDuDetail ( TOBLig : TOB);
var Indice : integer;
begin
	for indice := 0 to TOBLIG.detail.count -1 do
  begin
  	TOBLIG.detail[Indice].putvalue('BLO_QTEDUDETAIL',QTEDETAIL.Value);
  end;
end;

procedure TFLigNomen.QTEDETAILExit(Sender: TObject);
begin
	if QTEDETAIL.Value = 0 then QTEDETAIL.Value := 1;
  QteDudetail := QTEDETAIL.Value;
	AppliqueQteDuDetail ( TOBLig);
{$IFDEF BTP}
  CalculeOuvrageLoc;
  AfficheTotalisation;
  AfficheMontantLigne (G_NLIG.row);
{$ENDIF}
end;

procedure TFLigNomen.SetNumerotation(TOBL,TOBMere: TOb; Niveau : integer ;Arow: integer);
begin
  if Niveau=1 then
  begin
    // provenance de la ligne de document
    TOBL.putValue('BLO_N1',Arow);
  end else
  begin
    // provenance d'un detail d'ouvrage
    if Niveau = 2 then
    begin
      TOBL.PutValue('BLO_N1',TOBMere.getValue('BLO_N1'));
      TOBL.PutValue('BLO_N2',Arow);
    end else if Niveau = 3 then
    begin
      TOBL.PutValue('BLO_N1',TOBMere.getValue('BLO_N1'));
      TOBL.PutValue('BLO_N2',TOBMere.getValue('BLO_N2'));
      TOBL.PutValue('BLO_N3',Arow);
    end else if Niveau = 4 then
    begin
      TOBL.PutValue('BLO_N1',TOBMere.getValue('BLO_N1'));
      TOBL.PutValue('BLO_N2',TOBMere.getValue('BLO_N2'));
      TOBL.PutValue('BLO_N3',TOBMere.getValue('BLO_N3'));
      TOBL.PutValue('BLO_N4',Arow);
    end else if Niveau = 5 then
    begin
      TOBL.PutValue('BLO_N1',TOBMere.getValue('BLO_N1'));
      TOBL.PutValue('BLO_N2',TOBMere.getValue('BLO_N2'));
      TOBL.PutValue('BLO_N3',TOBMere.getValue('BLO_N3'));
      TOBL.PutValue('BLO_N4',TOBMere.getValue('BLO_N4'));
      TOBL.PutValue('BLO_N5',Arow);
    end;
  end;
end;

procedure TFLigNomen.RenumeroteLesFilles(TOBLN : TOB; Niveau : integer);
var LigPere1,LigPere2,LigPere3,LigPere4,LigPere5 : integer;
    Indice,CeluiLa : integer;
    NiveauANumeroter : integer;
    TOBLLN : TOB;
begin
  NiveauANumeroter := Niveau +1;
  if Niveau = 1 then
  begin
    LigPere1 := TOBLN.GetValue('BLO_N1');
    LigPere2 := 0;
    LigPere3 := 0;
    LigPere4 := 0;
    LigPere5 := 0;
  end else if Niveau = 2 then
  begin
    LigPere1 := TOBLN.GetValue('BLO_N1');
    LigPere2 := TOBLN.GetValue('BLO_N2');
    LigPere3 := 0;
    LigPere4 := 0;
    LigPere5 := 0;
  end else if Niveau = 3 then
  begin
    LigPere1 := TOBLN.GetValue('BLO_N1');
    LigPere2 := TOBLN.GetValue('BLO_N2');
    LigPere3 := TOBLN.GetValue('BLO_N3');
    LigPere4 := 0;
    LigPere5 := 0;
  end else if Niveau = 4 then
  begin
    LigPere1 := TOBLN.GetValue('BLO_N1');
    LigPere2 := TOBLN.GetValue('BLO_N2');
    LigPere3 := TOBLN.GetValue('BLO_N3');
    LigPere4 := TOBLN.GetValue('BLO_N4');
    LigPere5 := 0;
  end;
  for Indice := 0 to TOBLN.detail.count -1 do
  begin
    TOBLLN := TOBLN.detail[Indice];
    if NiveauANumeroter = 2 then BEGIN Inc(LigPere2); CeluiLa := LigPere2; END else
    if NiveauANumeroter = 3 then BEGIN Inc(LigPere3) ; CeluiLa := LigPere3; END else
    if NiveauANumeroter = 4 then BEGIN Inc(LigPere4) ; CeluiLa := LigPere4; END else
    if NiveauANumeroter = 5 then BEGIN Inc(LigPere5); ; CeluiLa := LigPere5; END;
    SetNumerotation (TOBLLN,TOBLN,NiveauANumeroter, CeluiLa);
    if TOBLLN.detail.count > 0 then
    begin
      RenumeroteLesFilles (TOBLLN,NiveauANumeroter);
    end;
  end;
end;

procedure TFLigNomen.BintegreExcelClick(Sender: TObject);
var ImportExcel : TimportLigneNomen;
		TOBL : TOB;
    i_ind : integer;
begin
  ODExcelFile.Filter := 'fichiers Excel (*xls;*xlsx)|*.xls;*.xlsx';
  if ODExcelFile.execute then
  begin
    ImportExcel := TimportLigneNomen.create;
    ImportExcel.documentPGI := TForm(TFLigNomen);
    ImportExcel.InfoNiv := infoNiv;
    ImportExcel.QteDuDetail := QTEDETAIL.Value;
    ImportExcel.documentExcel := ODexcelFile.FileName;
    ImportExcel.LignePGI := G_NLIG.row;
    ImportExcel.TOBpiece := TOBpiece;
    ImportExcel.TOBLIG := TOBLigPiece;
    ImportExcel.TOBOUV := TOBLIG;
    ImportExcel.TOBinfo := TOBInfo;
    ImportExcel.TOBArticles := TOBArticles;
    ImportExcel.DEV := DEV;
    ImportExcel.gestionHt := GestionHt;
    TRY
      ImportExcel.execute;
      for i_ind := 1 to G_NLIG.RowCount - 1 do
      begin
      	if TOB(G_NLIG.Objects[SGN_IdComp, i_ind]) <> nil then TOB(G_NLIG.Objects[SGN_IdComp, i_ind]).free;
      end;
      G_NLIG.VidePile (false);
      // Raffraichissement de la grille et recalcul de l'ouvrage
      TOBLig.PutGridDetail(G_NLIG, False, False, LesColNLIG);
      ChargeTobArticles;
//      HMTrad.ResizeGridColumns (G_NLIG) ;
      AfficheTVNLIG;
      CalculeOuvrageLoc;
      savouvrage := CalculSurTableau('+',TotalOuv,ValPourcent);
      savouvrage := CalculSurTableau('/',savouvrage,QteDudetail);
      afficheToTalisation;
      PositionneLig(1);
      if G_NLIG.rowcount < NbRowsInit then G_NLIG.rowCount := NbRowsInit;
    FINALLY
      ImportExcel.free;
    END;
  end;
end;

procedure TFLigNomen.SetTOBLigne (TOBL,TOBA,TOBInfo: TOB);
begin
  If TOBInfo <> nil then
  begin
    // Modif BTP
    TobL.PutValue('BLO_NATUREPIECEG',TOBInfo.GetValue('BLO_NATUREPIECEG'));
    TobL.PutValue('BLO_SOUCHE',TOBInfo.GetValue('BLO_SOUCHE'));
    TobL.PutValue('BLO_NUMERO',TOBInfo.GetValue('BLO_NUMERO'));
    TobL.PutValue('BLO_INDICEG',TOBInfo.GetValue('BLO_INDICEG'));
    TobL.PutValue('BLO_NUMLIGNE',TOBInfo.GetValue('BLO_NUMLIGNE'));
  end;
  if TOBA <> nil then
  begin
    // Modif BTP
    TobL.PutValue('BLO_TENUESTOCK',TOBA.GetValue('GA_TENUESTOCK'));
    TobL.PutValue('BLO_QUALIFQTEVTE',TOBA.GetValue('GA_QUALIFUNITEVTE'));
    TobL.PutValue('BLO_QUALIFQTESTO',TOBA.GetValue('GA_QUALIFUNITESTO'));
    TobL.PutValue('BLO_DPA',TOBA.GetValue('GA_DPA'));
    TobL.PutValue('BLO_DPR',TOBA.GetValue('GA_DPR'));
    TobL.PutValue('BLO_PMAP',TOBA.GetValue('GA_PMAP'));
    TobL.PutValue('BLO_PMRP',TOBA.GetValue('GA_PMRP'));
    {$IFDEF BTP}
    TobL.PutValue('BLO_LIBCOMPL',TOBA.GetValue('GA_LIBCOMPL'));
    TobL.PutValue('BNP_TYPERESSOURCE',TOBA.GetValue('BNP_TYPERESSOURCE'));
    {$ENDIF}
  end;
end;

function TFLigNomen.GetQteduDetail : double;
begin
	result := QTEDETAIL.Value;
end;

function TFligNomen.ChargeSousDetail (CodeNomen: string; TobOuvrage: TOB; Niv,OrdreCompo: integer; var Tresult: T_valeurs):boolean;
begin
	result := ChargeOuvrage (CodeNomen,TobOuvrage,Niv,OrdreCompo,Tresult);
end;

procedure TFligNomen.Affichegrid;
begin
	TOBLig.PutGridDetail(G_NLIG, False, False, LesColNLIG);
	HMTrad.ResizeGridColumns (G_NLIG) ;
  AfficheTVNLIG;
end;

procedure TFligNomen.recalculOuvrage;
begin
 CalculeOuvrageLoc;
 CalculSurTableau('+',TotalOuv,ValPourcent);
 CalculSurTableau('/',savouvrage,QteDudetail);
 afficheToTalisation;
end;

procedure TFligNomen.PositionneLig(Ligne : integer);
var Arow,Acol : integer;
		Cancel : boolean;
begin
	G_NLIG.OnRowEnter := nil;
	G_NLIG.OnRowExit := nil;
	G_NLIG.OnCellEnter := nil;
	G_NLIG.OnCellExit := nil;
  G_NLIG.OnEnter := nil;
  if Ligne > TOBLIG.detail.count then Ligne := TOBLIG.detail.count;
  G_NLIG.Col := SGN_Lib;
  G_NLIG.Row := Ligne;
	G_NLIG.OnRowEnter := G_NLIGRowEnter;
	G_NLIG.OnRowExit := G_NLIGRowExit;
	G_NLIG.OnCellEnter := G_NLIGCellEnter;
	G_NLIG.OnCellExit := G_NLIGCellExit;
  G_NLIG.OnEnter := G_NLIGEnter;
  if self.ActiveControl.Name <> 'GS' then G_NLIG.setfocus;
  Acol := G_NLIG.col; Arow := G_NLIG.row;
  G_NLIGEnter (Self);
  G_NLIGRowEnter(G_NLIG,G_NLIG.row,cancel,false);
  G_NLIGCellEnter (G_NLIG,Acol,Arow,cancel);
end;


procedure TFLigNomen.SetContraitance(value: boolean);
begin
	coTraitOk := value;
end;

procedure TFLigNomen.G_NLIGMouseMove(Sender: TObject; Shift: TShiftState;X, Y: Integer);
var Acol,Arow : integer;
	TOBl : TOB;
begin
  G_NLIG.MouseToCell(X,Y,Acol,Arow);
  if (Acol <= G_NLIG.FixedCols) or (Arow <= G_NLIG.FixedRows) then exit;
  G_NLIG.Hint := '';
  G_NLIG.Showhint := false;
  Application.CancelHint;
  if (Acol = SGN_LIBFOU) then
  begin
    TOBL := TOB(G_NLIG.Objects[SGN_NumLig,Arow]); if TOBL = nil then exit;
  	Application.CancelHint;
  	G_NLIG.Hint := TOBL.getvalue('LIBELLEFOU');
  	G_NLIG.Showhint := true;
  end;
end;

Function TFLigNomen.ElipsisMetreClick(TOBL : TOB) : Double;
begin

  result := AppelMetre(TOBL , True);

end;

Function TFLigNomen.AppelMetre(TOBL : TOB; OkClick : boolean = False) : Double;
Var TobLigne : TOB;
    SauveQte : Double;
Begin

  Result := Tobl.GetValue('BLO_QTEFACT');

  SauveQte := Tobl.GetValue('BLO_QTEFACT');

  if IsExcelLaunched then
  Begin
    PGIBox ('Vous devez fermer préalablement les instances d''EXCEL', 'Gestion des métrés');
    Result := SauveQte;
    Exit;
  end;

  If TheMetreShare <> nil then   // restitution du contexte en retour ---
  begin
    if not TheMetreShare.AutorisationMetre(TOBPIece.GetValue('GP_NATUREPIECEG')) then exit;
    //
    if TheMetreShare.OKMetreDoc then
    begin
      if TheMetreShare.OkExcel then
      begin
      //**** Nouvelle Version *****
        self.Enabled := False;
        Result := TheMetreShare.GereMetreXLS(TobPiece, TOBL, TOBLig, TOBLigPiece, OkClick);    //Piece/Lignes/Ouvrages     
        self.Enabled := True;
    end
    else
      begin
        if OkClick then
        begin
          self.Enabled := False;
    result := GereMetre (TOBL,TOBMetres,Action,'BLO');
          self.Enabled := True;
        end;
      end;
    end;
  end;

  if (SauveQte <> 0) and ((Result = 0) OR (Result = 1)) then Result := SauveQte;

end;

procedure TFLigNomen.VariablesDocumentClick(Sender: TObject);
Var TypeVarDoc : String;
    Cledoc     : R_Cledoc;
    TOBL       : TOB;
begin
{$IFDEF BTP}
  //
  CleDoc.NaturePiece  := TOBPiece.GetString('GP_NATUREPIECEG');
  CleDoc.DatePiece    := V_PGI.DateEntree;
  CleDoc.Souche       := TOBPiece.GetString('GP_SOUCHE');
  CleDoc.NumeroPiece  := TOBPiece.GetValue('GP_NUMERO');
  CleDoc.Indice       := TOBPiece.GetValue('GP_INDICEG');

  TheMetreShare.SaisieVarDoc (Cledoc, 'D');
{$ENDIF}
end;

procedure TFLigNomen.VariablesApplicationClick(Sender: TObject);
Var TypeVarDoc : String;
    Cledoc     : R_Cledoc;
begin
{$IFDEF BTP}
  TypeVarDoc := 'A';

  AGLLanceFiche('BTP', 'BTVARIABLE', TypeVarDoc , '','APPLICATION');
{$ENDIF}

end;

procedure TFLigNomen.VariablesLigneClick(Sender: TObject);
Var TypeVarDoc : String;
    CleVarDoc  : R_Cledoc;
    TOBL       : TOB;
begin

  TOBl := TOB(G_NLIG.Objects[SGN_NumLig, G_NLIG.Row]);

  if TOBL = nil then exit;

  CleVardoc.NaturePiece := TOBL.GetValue('BLO_NATUREPIECEG');
  CleVardoc.Souche      := TOBL.GetValue('BLO_SOUCHE');
  CleVardoc.NumeroPiece := StrToInt(TOBL.GetValue('BLO_NUMERO'));
  CleVardoc.Indice      := StrToInt(TOBL.GetValue('BLO_INDICEG'));
  CleVardoc.NumOrdre    := StrToInt(TOBL.GetValue('BLO_NUMORDRE'));
  CleVardoc.UniqueBlo   := StrToInt(TOBL.GetValue('BLO_UNIQUEBLO'));
  //
  TheMetreShare.SaisieVarDoc (CleVardoc, 'LS');

end;

procedure TFLigNomen.VariablesGeneraleClick(Sender: TObject);
Var TypeVarDoc : String;
    Cledoc     : R_Cledoc;
begin
{$IFDEF BTP}
  TypeVarDoc := 'G';

  AGLLanceFiche('BTP', 'BTVARIABLE', TypeVarDoc , '','GENERALE');
{$ENDIF}

end;

end.





