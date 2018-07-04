unit NomenLig;

//===========================================================================
//  Gestion des nomenclatures.
//
//  Cette fenetre est conçue pour gerer un niveau de nomenclature, le premier
//  niveau d'une nomenclature donnée.
//===========================================================================

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, Grids, HDB, ExtCtrls, ComCtrls, Hctrls, StdCtrls, UTOB,
  SaisUtil, HEnt1, hmsgbox, TarifUtil, NomenUtil,EntGC,
{$IFDEF BTP}
  UTofBTAnalDev,
{$ENDIF}
  CasEmplois, HSysMenu,
  ImgList,Choix,
{$IFDEF EAGLCLIENT}
      MaineAGL,UtileAGL,
{$ELSE}
      MPlayer, DBGrids, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Db, Fe_Main,
      EdtREtat,uPDFBatch,
{$IFDEF V530}
     EdtDOc,
{$ELSE}
     EdtRDoc,
{$ENDIF}
{$ENDIF}
  Math, NomenErr, Menus, M3FP, AglInitGC, UtilArticle, HRichEdt, HRichOLE,
{$IFDEF GPAOLIGHT}
  WJetons, WNOMETET, wCommuns,
{$ENDIF GPAOLIGHT}
  HPanel, HImgList, TntStdCtrls, TntExtCtrls, TntGrids,
  UtilsMetresXLS,UEntCommun,Paramsoc;

procedure MajDateModifEntete (CodeArt : string);
Procedure MajDateModifOuvrage(CodeArt : string);
procedure Entree_NomenLig(Parms: array of variant; nb: integer);
procedure EntreeOuvrageDetail(Parms: array of variant; nb: integer);
function VerifExistanceEnteteOuv(CodeArt,Libelle,Domaine : String; var QteduDetail : Integer; Creation : Boolean): boolean;
procedure EntreeParcDetail(Parms: array of variant; nb: integer);
function VerifExistanceEnteteParc(CodeArt,Libelle : String;Creation : Boolean): boolean;

type
	TTypeNomen = (TTNNomen,TTNOuv,TTNPARC);
  TFNomenLig = class(TForm)
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BAide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BValider: TToolbarButton97;
    G_NLIG: THGrid;
    PANNOMENENT: TPanel;
    GNE_ARTICLE: TEdit;
    GNE_NOMENCLATURE: TEdit;
    TGNE_ARTICLE: TLabel;
    TGNE_NOMENCLATURE: TLabel;
    GNE_LIBELLE: TEdit;
    TGNE_LIBELLE: TLabel;
    MsgBox: THMsgBox;
    ImageList1: THImageList;
    PANPIED: TPanel;
    GNL_SOUSNOMEN: TEdit;
    TGNL_SOUSNOMEN: TLabel;
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
    bPrint: TToolbarButton97;
    BInfos: TToolbarButton97;
    PANDIM: TPanel;
    TGA_CODEDIM5: TLabel;
    TGA_CODEDIM4: TLabel;
    TGA_CODEDIM3: TLabel;
    TGA_CODEDIM2: TLabel;
    TGA_CODEDIM1: TLabel;
    N3: TMenuItem;
    PopG_NLIG_D: TMenuItem;
    BTPpanel: THPanel;
    TV_NLIG: TTreeView;
    POuvrage: THPanel;
    PopG_NLIG_O: TMenuItem;
    PANDetailBtp: TPanel;
    TGNE_QTEDUDETAIL: TLabel;
    TUNITE: THLabel;
    BTypeArticle: TToolbarButton97;
    BArborescence: TToolbarButton97;
    BInfosLigne: TToolbarButton97;
    ToolBarDesc: TToolWindow97;
    Pligne: THPanel;
    HLabel1: THLabel;
    HLabel3: THLabel;
    HLabel2: THLabel;
    HLabel4: THLabel;
    HLabel5: THLabel;
    HLabel6: THLabel;
    HLabel7: THLabel;
    TGA_FOURNPRINC: TLabel;
    Label1: TLabel;
    TGA_PRESTATION: TLabel;
    TLIB_PREST: TLabel;
    NEPA: THNumEdit;
    NECOEFPAPR: THNumEdit;
    NEPR: THNumEdit;
    NECOEFPRPV: THNumEdit;
    NEPVHT: THNumEdit;
    NEPVFORFAIT: THNumEdit;
    NEPVTTC: THNumEdit;
    ImTypeArticle: THImageList;
    HLabel16: THLabel;
    NETPS: THNumEdit;
    GNE_QTEDUDETAIL: THNumEdit;
    HPanel3: THPanel;
    HLabel12: THLabel;
    COEFFG: THNumEdit;
    HLabel13: THLabel;
    COEFMARG: THNumEdit;
    HPanel1: THPanel;
    PValo: THPanel;
    HLabel8: THLabel;
    HLabel9: THLabel;
    HLabel10: THLabel;
    HLabel11: THLabel;
    HLabel14: THLabel;
    HLabel15: THLabel;
    HLabel28: THLabel;
    MONTANTPA: THNumEdit;
    MONTANTPR: THNumEdit;
    MONTANTHT: THNumEdit;
    MONTANTTTC: THNumEdit;
    MONTANTFG: THNumEdit;
    MONTANTMARG: THNumEdit;
    PValoUnitaire: THPanel;
    HLabel18: THLabel;
    HLabel19: THLabel;
    HLabel20: THLabel;
    HLabel23: THLabel;
    HLabel24: THLabel;
    HLabel25: THLabel;
    HLabel27: THLabel;
    MONTANTPAU: THNumEdit;
    MONTANTPRU: THNumEdit;
    MONTANTHTU: THNumEdit;
    MONTANTTTCU: THNumEdit;
    MONTANTFGU: THNumEdit;
    MONTANTMARGU: THNumEdit;
    HLabel17: THLabel;
    TOTALTPS: THNumEdit;
    HLabel21: THLabel;
    TPSUNITAIRE: THNumEdit;
    Variables_Lig: TToolbarButton97;
    AppelXLS: TToolbarButton97;
    PopG_NLIG_V: TMenuItem;
    PopG_NLIG_E: TMenuItem;
//  evenements sur la forme
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
//  boutons
    procedure BValiderClick(Sender: TObject);
    procedure bPrintClick(Sender: TObject);
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
//  evenements sur le treeview
    procedure TV_NLIGMouseUp(Sender: TObject; Button: TMouseButton;
                             Shift: TShiftState; X, Y: Integer);
    procedure TV_NLIGCollapsing(Sender: TObject; Node: TTreeNode;
                                var AllowCollapse: Boolean);
    procedure TV_NLIGExpanding(Sender: TObject; Node: TTreeNode;
                               var AllowExpansion: Boolean);
    procedure TV_NLIGDblClick(Sender: TObject);
    procedure PopTVNLIG_TClick(Sender: TObject);
    procedure PopTVNLIG_FClick(Sender: TObject);
    procedure PopTV_NLIG_OClick(Sender: TObject);
    procedure ExitPVForfait(Sender: TObject);
    procedure PopG_NLIG_OClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BTypeArticleClick(Sender: TObject);
    procedure ToolBarDescClose(Sender: TObject);
    procedure GNE_QTEDUDETAILExit(Sender: TObject);
    function  GetDomaineOuvrage(CodeOuvrage : string) : string;
    procedure constitueTOBEtat (TOBEtat : TOB);
    procedure AppelXLSClick(Sender: TObject);
    procedure Variables_LigClick(Sender: TObject);
  private
    { Déclarations privées }
  SGN_NumLig  : integer;
  SGN_IdComp  : integer;
  SGN_Comp    : integer;
  SGN_Lib     : integer;
  SGN_Qte     : integer;
  SGN_Joker   : integer;
  SGN_SSNom   : integer;
  // Modif BTP
  SGA_UVTE    : integer;
  SGN_Prix    : integer;
  SGN_TypA    : Integer;
  // NEWONE
  SGN_FOURN :integer;
  SGN_TARIFF :integer;
  SGN_UNAC :integer;
  SGN_REMF :integer;
  SGN_PAHT :integer;
  SGN_PVHT :integer;
  SGN_PREST :integer;
  SGN_TPS :integer;
  SGN_MO :integer;
  SGN_MTHT : integer;
  SGN_QTESAIS : integer;
  SGN_UNITESAIS : integer;
  SGN_PERTE : integer;
  SGN_RENDEMENT : integer;

  //
    Cell_Text : string;
    iTableLigne : integer;
{$IFDEF BTP}
    itableLigneArt : integer;
    ToolBartree : TToolWindow97;
{$ENDIF}
    ColsInter : Array of boolean ;
    TOBLig : TOB;
    LesColNLIG : String ;
    ART : THCritMaskEdit;
    // -- MODIF --
    Domaine : String;
    //
    MetreArticle    : TMetreArt;
    //
    procedure EtudieColsListe ;
    procedure PostDrawCell(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
//  Gestion du Treeview
    procedure SetImages(TV : TTreeView; TN : TTreeNode);
    procedure AfficheTVNLIG;
//  Gestion des champs
    function  TraiterComposant(ARow : LongInt) : boolean;
    function  TraiterQuantite(ARow : LongInt) : boolean;
    function  TraiterJoker(ARow : LongInt) : boolean;
    procedure FormateZoneSaisie(Acol,ARow : LongInt) ;
    function  Recherche_Art(ARow : LongInt) : boolean;
//  Gestion des colonnes
//  Gestion des lignes du Grid
//  Attention : les propriétés Objects sont utilisées...
//              Objects[SGN_NumLig, x] contient la Tob ligne de nomenclature
//              Objects[SGN_Comp, x] contient le TreeNode associé à la ligne
//              Objects[SGN_IdComp, x] contient la tob Article
//   Modif BTP
//              Objects[SGN_Prix,x] contient le prix HT des nomenclatures (calculé)
    procedure InsertLigne (ARow : Longint; Ajout : Boolean) ;
    procedure InitialiseLigne (ARow : Longint) ;
    procedure SupprimeLigne (ARow : Longint) ;
    procedure Renumerote (ARow : Longint) ;
    procedure AffectePied(ARow : Longint);
    procedure EffaceDimensions;
{$IFNDEF BTP}
//    function  ChercheLigne (Valeur : string) : integer; // DBR Fiche 10995
    function  ChercheLigne (tn1 : TTreeNode) : integer; // DBR Fiche 10995
{$ENDIF}
    function  ValideLigne(Ou : Longint) : boolean;
    procedure ChargeTobArticles;
{$IFDEF BTP}
    function TraiterPrix(ARow: Integer): boolean;
    function ChercheLigneBtp(tn1: TTreeNode): integer;
    procedure ChargePrixOuvrage (TOBLIG : TOB;CodeNomenc: String; ARow : Integer);
    function calculeMontantLig(Ou: Integer): T_Valeurs;
    procedure AffecteValArticle(Arow: integer);
    procedure AlimenteValeurOuv();
    procedure CalculeMontantOuv();
    function RecupPrixBase(Arow: integer): double;
    procedure calculecoefs;
    procedure definiBarTree;
    procedure AssigneEvenementBtn;
    procedure BArborescenceClick(Sender: TObject);
    procedure BInfosLigneClick(Sender: TObject);
    procedure ToolBarTreeClose(Sender: Tobject);
    procedure CalculeValoOuvrage(TOBLig : TOB; var valeurs : T_valeurs);
    //FV1 : 17/02/2015 - Création de nouvelles Procédures pour rendre le source plus lisible !!!!
    procedure ChargeEcranOuvrage;
    procedure GestionGrilleOuvrage;
    procedure GestionTableauGrille;
    function TraiteQteSais(Acol,Arow: integer): boolean;

{$ENDIF}
    procedure RecupInfoPrixPose (TOBmere : TOB);
    function ZoneAccessible(ACol, ARow: Integer): boolean;
    procedure ZoneSuivanteOuOk(var ACol, ARow: Integer; var Cancel: boolean);
    function CalculeQteSais(TOBL: TOB): double;
    function CalculeQteFact(TOBL: TOB): double;

  public
    { Déclarations publiques }
    Action      : TActionFiche ;
    TVCols : string;
    Data_Modified : Boolean;
    TValeurOuv,TValeurPOU : T_Valeurs;
    TMontantLig : T_Valeurs;
    TPercent : T_Valeurs;
    IPercent : Integer;
    TypeSaisie :  TTypeNomen;
  end;

var
  LastKey : Word;
  CasEmploi_Open : Boolean;
  OkAffichage : boolean;

implementation

{$R *.DFM}
  uses
{$IFDEF GPAOLIGHT}
    wNomeLig,
{$ENDIF}
	CbpMCD
;


procedure MajDateModifEntete (CodeArt : string);
var TOBEnt : TOB;
		Req : String;
    QQ : TQuery;
begin
	TOBEnt := TOB.create ('NOMENENT',nil,-1);
	Req := 'SELECT GNE_ARTICLE,GNE_NOMENCLATURE,GNE_DATEMODIF FROM NOMENENT WHERE GNE_ARTICLE="'+ CodeArt +'"';
	QQ := Opensql (Req,true,-1, '', True);
	if not QQ.EOF then
	begin
		TOBENT.SelectDB ('',QQ);
    TOBENT.PutValue('GNE_DATEMODIF',Now);
    TOBENT.UpdateDB (false);
  end;
  Ferme (QQ);
  TOBENT.free;
end;

procedure MajDateModifOuvrage (CodeArt : string);
var TOBOuv  : TOB;
		Req     : String;
    QQ      : TQuery;
begin
	TOBOuv := TOB.create ('ARTICLE',nil,-1);
	Req := 'SELECT GA_ARTICLE,GA_DATEMODIF, GA_UTILISATEUR FROM ARTICLE WHERE GA_CODEARTICLE="'+ CodeArt +'" AND GA_TYPEARTICLE="OUV"';
	QQ := Opensql (Req,true,-1, '', True);
	if not QQ.EOF then
	begin
		TOBOuv.SelectDB ('',QQ);
    TOBOuv.PutValue('GA_DATEMODIF',Now);
    TOBOuv.PutValue('GA_UTILISATEUR',V_PGI.User);
    TOBOuv.UpdateDB (false);
  end;
  Ferme (QQ);
  TOBOuv.free;
end;

function formatDec ( chaine : string; Nb : integer) : string;
var Indice : integer;
begin
result := '';
for Indice := 1 to Nb do
    begin
    result := result + chaine;
    end;
end;

procedure Entree_NomenLig(Parms: array of variant; nb: integer);
var FNomLig : TFNomenLig;
    st1, st2 : string;
    Var_Appel : string;
{$IFDEF CHR}
    i: integer;
{$ENDIF}
begin
OkAffichage := false;
CasEmploi_Open := True;
Var_Appel := string(Parms[0]);
st1 := Copy(Var_Appel, 0, Pos('=',Var_Appel) - 1);
if st1 = 'CASEMPLOIS' then
    begin
    st2 := Copy(Var_Appel, Pos('=',Var_Appel) + 1, 255);
    if st2 = 'NON' then CasEmploi_Open := False;
    end;
FNomLig := TFNomenLig.Create(Application);
Try
FNomLig.TypeSaisie := TTNOuv;
FNomLig.GNE_ARTICLE.Text := String(Parms[1]);
FNomLig.GNE_NOMENCLATURE.Text := String(Parms[2]);
{$IFDEF BTP}
if nb > 3 then
    begin
    FNomLig.GNE_QTEDUDETAIL.Value  := StrToInt(String(Parms[3]));
    if nb = 4 then FNomLig.Action := TactionFiche(parms[4]);
    end
else FNomLig.GNE_QTEDUDETAIL.value := 0;
{$ENDIF}
{$IFDEF CHR}
    for i := 0 to FNomLig.PopupG_NLIG.Items.Count - 1 do
    begin
      if (i > 7) then
        FNomLig.PopupG_NLIG.Items[i].Visible := false;
    end;
{$ENDIF}
FNomLig.ShowModal;
finally
FNomLig.free;
MajDateModifEntete (String(Parms[1]));
end;
end;

procedure EntreeParcDetail(Parms: array of variant; nb: integer);
{$IFDEF BTP}
var FNomLig : TFNomenLig;
QteDudetail,Indice : Integer;
CodeArt,libelle : string;
{$ENDIF}
begin
{$IFDEF BTP}
  OkAffichage := false;
  CasEmploi_Open := True;
  CodeArt := String(Parms[0]);
  Libelle := String(Parms[1]);
  VerifExistanceEnteteParc(CodeArt,Libelle, true);
  FNomLig := TFNomenLig.Create(Application);
  Try
		FNomLig.TypeSaisie := TTNPARC;
    FNomLig.GNE_ARTICLE.Text := CodeArt;
    FNomLig.GNE_LIBELLE.Text := Libelle;
    FNomLig.GNE_NOMENCLATURE.Text := copy(CodeArt,0,18);
    FNomLig.GNE_QTEDUDETAIL.value := 1;
    for indice := 0 To FNomLig.PopupG_NLIG.Items.Count - 1 do
    begin
      if FNomLig.PopupG_NLIG.Items[Indice].Name = 'PopG_NLIG_O' then
      	FNomLig.PopupG_NLIG.Items[Indice].Visible := false;
    end;
    FNomLig.Action := TactionFiche(Parms[2]);
    FNomLig.ShowModal;
  finally
  	FNomLig.Free;
  	MajDateModifEntete (String(Parms[0]));
  end;
{$ENDIF}
end;

//
// Modif BTP
//
procedure EntreeOuvrageDetail(Parms: array of variant; nb: integer);
{$IFDEF BTP}
var FNomLig : TFNomenLig;
QteDudetail,Indice : Integer;
{$ENDIF}
begin
{$IFDEF BTP}
OkAffichage := false;
CasEmploi_Open := True;
if VerifExistanceEnteteOuv (string(Parms[0]),String(Parms[1]),String(Parms[2]),Qtedudetail,true) then
    begin
    FNomLig := TFNomenLig.Create(Application);
    Try
    FNomLig.TypeSaisie := TTNOuv;
    FNomLig.GNE_ARTICLE.Text := String(Parms[0]);
    FNomLig.GNE_NOMENCLATURE.Text := copy(String(Parms[0]),0,18);
    FNomLig.GNE_QTEDUDETAIL.value := Qtedudetail;
    for indice := 0 To FNomLig.PopupG_NLIG.Items.Count - 1 do
        begin
        if FNomLig.PopupG_NLIG.Items[Indice].Name = 'PopG_NLIG_O' then
            FNomLig.PopupG_NLIG.Items[Indice].Visible := true;
        end;
    FNomLig.Action := TactionFiche(Parms[3]);
    FNomLig.ShowModal;
    finally
    FNomLig.Free;
  		MajDateModifEntete (String(Parms[0]));
    end;
    end
else
    begin
    // Message Utilisez les décompositions
    PGIBox('Vous devez utiliser les déclinaisons','Ouvrages : ') ;
    end;
{$ENDIF}
end;

function VerifExistanceEnteteParc(CodeArt,Libelle : String;Creation : Boolean): boolean;
var
   TOBEnt : Tob;
   QQ : Tquery;
   Req : String;
begin
  TOBEnt := TOB.create ('NOMENENT',nil,-1);
  Req := 'SELECT GNE_NOMENCLATURE FROM NOMENENT WHERE GNE_ARTICLE="'+ CodeArt +'"';
  QQ := Opensql (Req,true,-1, '', True);
  if QQ.EOF then
  begin
    TOBEnt.PutValue ('GNE_ARTICLE',CodeArt);
    TOBEnt.PutValue ('GNE_LIBELLE',Libelle);
    TOBEnt.PutValue ('GNE_DOMAINE','');
    TOBEnt.putValue ('GNE_NOMENCLATURE',copy(CodeArt,0,18));
    TOBENt.PutValue ('GNE_QTEDUDETAIL',1);
    TOBENt.PutValue ('GNE_DATEMODIF', V_PGI.DateEntree);
    TOBENt.PutValue ('GNE_DATECREATION', V_PGI.DateEntree);
    TOBEnt.InsertOrUpdateDB(true);
    result := true;
  end;
  TOBEnt.Free;
  Ferme(QQ);
end;

function VerifExistanceEnteteOuv(CodeArt,Libelle,Domaine : String;var Qtedudetail : integer;Creation : Boolean): boolean;
var
   TOBEnt : Tob;
   QQ : Tquery;
   Req : String;
begin
     QteDuDetail := 0;
     TOBEnt := TOB.create ('NOMENENT',nil,-1);
     Req := 'SELECT GNE_NOMENCLATURE,GNE_QTEDUDETAIL FROM NOMENENT WHERE GNE_ARTICLE="'+ CodeArt +'"';
     QQ := Opensql (Req,true,-1, '', True);
     if not QQ.EOF then
     begin
          TOBENT.LoadDetailDB ('NOMENENT','','',QQ,false);
          if TOBENT.detail.Count > 1 then
             result := false
          else
          begin
             IF trim(copy (tobent.detail[0].GetValue('GNE_NOMENCLATURE'),0,18)) = trim(copy (codeArt,0,18)) then
             begin
                  QteDudetail := TOBENT.detail[0].getvalue('GNE_QTEDUDETAIL');
                  result := true;
             end
             else
                 result := false;
          end;
     end
     else
     begin
          TOBEnt.PutValue ('GNE_ARTICLE',CodeArt);
          TOBEnt.PutValue ('GNE_LIBELLE',Libelle);
          TOBEnt.PutValue ('GNE_DOMAINE',Domaine);
          TOBEnt.putValue ('GNE_NOMENCLATURE',copy(CodeArt,0,18));
          TOBENt.PutValue ('GNE_QTEDUDETAIL',1);
          TOBENt.PutValue ('GNE_DATEMODIF', V_PGI.DateEntree);
          TOBENt.PutValue ('GNE_DATECREATION', V_PGI.DateEntree);
          TOBEnt.InsertOrUpdateDB(true);
          QteDudetail := strtoint ('1');
          result := true;
     end;
     TOBEnt.Free;
     Ferme(QQ);
end;

procedure TFNomenLig.FormCreate(Sender: TObject);
begin
{$IFDEF BTP}
Domaine := '';
ToolBartree := TToolWindow97.Create (self);
definiBarTree;
BTPPanel.visible := false;
AssigneEvenementBtn;
BInfosLigne.visible := true;
BArborescence.visible := true;
{$ENDIF}
ART := THCritMaskEdit.Create (Self);
SGN_NumLig  := -1;
SGN_IdComp  := -1;
SGN_Comp    := -1;
SGN_Lib     := -1;
SGN_Qte     := -1;
SGN_Joker   := -1;
SGN_SSNom   := -1;
// Modif BTP
SGA_UVTE    := -1;
SGN_Prix    := -1;
SGN_TypA    := -1;
// NEWONE
SGN_FOURN :=-1;
SGN_TARIFF :=-1;
SGN_UNAC :=-1;
SGN_REMF :=-1;
SGN_PAHT :=-1;
SGN_PVHT :=-1;
SGN_PREST :=-1;
SGN_TPS :=-1;
SGN_MO :=-1;
SGN_MTHT := -1;
SGN_QTESAIS := -1;
SGN_UNITESAIS := -1;
SGN_PERTE := -1;
SGN_RENDEMENT := -1;
end;

procedure TFNomenLig.FormShow(Sender: TObject);
var
    Cancel : boolean;
    // Modif BRP
    Indice : Integer;
    TValLoc : T_Valeurs;
    ArticleOk : string;
//uniquement en line
{*
    TL : TLabel;
    TN : THNumEdit;
*}
    ACol,Arow : integer;
    TobEnt : TOB;
begin

  PANDIM.Visible  := PopG_NLIG_D.Checked;

  Data_Modified   := False;

  if TypeSaisie = TTNOUV then
  begin
    ChargeEcranOuvrage;
    G_NLIG.ListeParam:='BTNOMENLIG';
  end
  else if TypeSaisie = TTNPARC then
  begin
    Caption := TraduireMemoire('Saisie de la nomenclature de parc ' + GNE_NOMENCLATURE.Text);
    GNE_ARTICLE.Visible := false;
    TGNE_ARTICLE.visible := false;
    G_NLIG.ListeParam:='BTNOMENPARC';
  end;

  TOBEnt := TOB.Create('NOMENENT', nil, -1);
  TOBEnt.PutValue('GNE_NOMENCLATURE', GNE_NOMENCLATURE.Text);
  TOBEnt.PutValue('GNE_ARTICLE', GNE_ARTICLE.Text);
  TOBEnt.ChargeCle1;

  if not TOBEnt.LoadDB then
  begin
    {$IFNDEF CHR}
    MsgBox.Execute(8, Caption, '');
    Close;
    {$ELSE}
    TOBEnt.InsertDB(nil);
    {$ENDIF}
  end;
  GNE_LIBELLE.Text := TOBEnt.GetValue('GNE_LIBELLE');

  iTableLigne := PrefixeToNum('GNL') ;
  itableLigneArt := prefixeTOnum ('GA');
  ChargeNomen(GNE_NOMENCLATURE.Text, True, TOBLig);
  AffecteLibelle(TOBLig, 'GNL_CODEARTICLE');

  G_NLIG.PostDrawCell := PostDrawCell;

  EtudieColsListe;

  AffecteGrid (G_NLIG,Action) ;

  G_NLIG.ColAligns[SGN_NumLig] := taCenter;

  if TypeSaisie = TTNOUV then GestionGrilleOuvrage;

  G_NLIG.ColWidths[G_NLIG.ColCount - 1]  := 0;
  G_NLIG.ColWidths[G_NLIG.ColCount - 2]  := 0;
  G_NLIG.ColLengths[G_NLIG.ColCount - 1] := -1;
  G_NLIG.ColLengths[G_NLIG.ColCount - 2] := -1;

  if (not GetParamSocSecur('SO_BTGESTQTEAVANC',false)) then
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

  G_NLIG.ColFormats [SGN_QTe] := '# ### ##0.'+ formatDec ('0',V_PGI.okdecQ);

  HMTrad.ResizeGridColumns (G_NLIG) ;

  if TOBLig.Detail.Count = 0 then
  begin
    G_NLIG.Cells[0, 1] := IntToStr(1);
    G_NLIG.Cells [SGN_Qte, 1] := StrS(1, 4);
		if TypeSaisie=TTNOuv then
    begin
      G_NLIG.Cells[SGN_Prix, 1] := strs (valeur ('0'), V_PGI.OkDecP);
    end;
    G_NLIG.Cells[SGN_Joker, 1] := 'N';
  end;

  TOBLig.PutGridDetail(G_NLIG, False, False, LesColNLIG);

  AfficheTVNLIG;

  ChargeTobArticles;

  {$IFDEF BTP}
  if TypeSaisie = TTNOUV then GestionTableauGrille;
  {$ENDIF}

  TOBEnt.free;

  EffaceDimensions;

  cancel  := False;
  Acol    := G_NLIG.FixedCols;
  ARow    := 1;

  TOBLig.PutGridDetail(G_NLIG, False, False, LesColNLIG);

  G_NLIGRowEnter(Sender, ARow, Cancel, False);
  G_NLIGCellEnter (Sender,ACol,Arow,cancel);

  G_NLIG.col  := ACol;
  G_NLIG.row  := Arow;

  Cell_Text   := G_NLIG.Cells[G_NLIG.Col, G_NLIG.Row];

  OkAffichage := true;

end;

procedure TFNomenLig.FormKeyDown(Sender: TObject; var Key: Word;
                                 Shift: TShiftState);
var FocusGrid : Boolean;
    ARow : Longint;
    stWhere : string;
begin
ARow:=0;
BTypeArticle.visible := false;
FocusGrid := False;
if(Screen.ActiveControl = G_NLIG) then
    begin
    FocusGrid := True;
    ARow := G_NLIG.Row;
    end;
LastKey := Key;
Case Key of
    VK_RETURN : Key:=VK_TAB ;
    VK_F5     : BEGIN
                  if FocusGrid then
                  BEGIN
                    Key := 0;
                    G_NLIGElipsisClick (self);
                  END;
                END;
    VK_INSERT : BEGIN
                if FocusGrid then
                    BEGIN
                    Key := 0;
                    InsertLigne (ARow, False);
                    END;
                END;
    VK_DELETE : BEGIN
                if ((FocusGrid) and (Shift=[ssCtrl])) then
                    BEGIN
                    Key := 0 ;
                    SupprimeLigne (ARow) ;
                    END ;
                END;
    END;
end;

procedure TFNomenLig.FormClose(Sender: TObject; var Action: TCloseAction);
var
    w_rep : word;
    // Modif BTP
    i_ind1,i_col : Integer;
begin
  if Data_Modified then
    begin
      //  Demande de validation
      w_rep := MsgBox.Execute (0,Caption,'');
      case w_rep of
      mrYes     : BEGIN
                    BValiderClick(Sender);
                    if ModalResult = 0 then exit;
                  END;

      mrCancel  : Begin
                    Action := caNone;
                  end;
      MrNo      : Begin
                  // Modif BTP
                    for i_Col := 0 To G_NLIG.ColCount - 1 do
                    begin
                      for i_ind1 := 0 to G_NLIG.RowCount - 1 do
                      begin
                        if Tob(G_NLIG.Objects[i_col, i_ind1]) <> nil then
                          G_NLIG.Objects[i_col, i_ind1].free;
                      end;
                    end;
                    TobLig.free;
                    ToolBartree.Free;
                    ART.free;
                  //
                  end;
      end;
    end;
end;

//=============================================================================
procedure TFNomenLig.EtudieColsListe ;
Var NomCol,LesCols : String ;
  icol,ichamp, i_ind : integer ;
  Mcd : IMCDServiceCOM;
  Field : IFieldCOM  ;
BEGIN
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
//
SetLength(ColsInter,G_NLIG.ColCount) ;
for i_ind:=Low(ColsInter)to High(ColsInter) do ColsInter[i_ind]:=False ;
LesCols:=G_NLIG.Titres[0] ; LesColNLIG:=LesCols ; icol:=0 ;
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol<>'' then
        BEGIN
        Field := (Mcd.GetField(Nomcol) as IFieldCOM);
        ichamp:=ChampToNum(NomCol) ;
        if ichamp>=0 then
            BEGIN
{$IFDEF BTP}
            if (Pos('GNL_',NomCol) > 0) and (Pos('X',Field.Control)>0) then ColsInter[icol]:=True
            else if (Pos('GNL_',NomCol) = 0) then ColsInter[icol] := True; (* Pour les autres tables *)
{$ELSE}
        if Pos('X',V_PGI.Dechamps[iTableLigne,ichamp].Control)>0 then ColsInter[icol] := True;
{$ENDIF}
        if NomCol='GNL_NUMLIGNE'        then SGN_NumLig := icol else
        if NomCol='GNL_ARTICLE'         then SGN_IdComp := icol else
        if NomCol='GNL_CODEARTICLE'     then SGN_Comp   := icol else
        if NomCol='GNL_SOUSNOMEN'       then SGN_SSNom  := icol else
        if NomCol='GNL_LIBELLE'         then SGN_Lib    := icol else
        if NomCol='GNL_QTE'             then SGN_Qte    := icol else
{$IFDEF BTP}
        if NomCol='GNL_PRIXFIXE'        then SGN_Prix   := icol else
        if NomCol='GA_TYPEARTICLE'      then SGN_TypA   := icol else
        if NomCol='GA_QUALIFUNITEVTE'   then SGA_UVTE   := icol else
        // NEWONE
        if NomCol='GA_FOURNPRINC'       Then SGN_FOURN  := icol else
        if NomCol='GCA_PRIXBASE'        Then SGN_TARIFF := icol else
        if NomCol='GCA_QUALIFUNITEACH'  Then SGN_UNAC   := icol else
        if NomCol='GF_CALCULREMISE'     Then SGN_REMF   := icol else
        if NomCol='GA_PAHT'             Then SGN_PAHT   := icol else
        if NomCol='GA_PVHT'             Then SGN_PVHT   := icol else
        if NomCol='GA_NATUREPRES'       Then SGN_PREST  := icol else
        if NomCol='GA_HEURE'            Then SGN_TPS    := icol else
        if NomCol='GF_PRIXUNITAIRE'     Then SGN_MO     := icol else
        if NomCol='GCA_PRIXVENTE'       Then SGN_MTHT   := icol else
        if NomCol='GNL_QTESAIS'         Then SGN_QTESAIS := icol else
        if NomCol='GNL_QUALIFUNITEACH'  Then SGN_UNITESAIS := icol else
        if NomCol='GNL_PERTE'           Then SGN_PERTE := icol else
        if NomCol='GNL_RENDEMENT'       Then SGN_RENDEMENT := icol else
{$ENDIF}
        if NomCol='GNL_JOKER'           Then
        begin
          SGN_Joker:=icol;
          G_NLIG.ColLengths[icol] :=1;
          //G_NLIG.ColTypes[icol] := 'B';
        end;
      END ;
    END ;
    Inc(icol) ;
  Until ((LesCols='') or (NomCol='')) ;

end;

function TFNomenLig.Recherche_Art(ARow : LongInt) : boolean;
var RechArt   : T_RechArt ;
    Q         : TQuery;
    TF1       : TField;
    OkArt     : Boolean ;
    st1       : String;
    libelle   : String;
    StWhere   : string;
    TOBArt    : TOB;
    TOBLoc    : TOB;
    TLocal    : T_Valeurs;
    Larticle  : string;
begin
    Libelle:='';
    OkArt:=False ;
    TOBArt := TOB.Create('ARTICLE',nil,-1);
    ART.Text := G_NLIG.Cells[SGN_Comp, ARow];
//-----------------------------------------------------------------
//  Recherche du code article saisi
//-----------------------------------------------------------------
    RechArt := TrouverArticle (ART, TOBArt);
    case RechArt of
            traOk : begin
                    if (TOBArt.GetValue('GA_TYPEARTICLE') = 'ARP') then
                        begin
                        st1 := '';
                        Q := OpenSQL('Select GNE_NOMENCLATURE from NOMENENT Where GNE_NOMENCLATURE="' +
                                     TOBArt.GetValue('GA_ARTICLE') + '"', True);
                        TF1 := Q.FindField('GNE_NOMENCLATURE');
{$IFDEF EAGLCLIENT}
                        st1 := TF1.AsString ;
{$ELSE}
                        st1 := TF1.Text;
{$ENDIF}
                        Ferme(Q);
                        if st1 <> '' then
                            begin
//-----------------------------------------------------------------
//  on enregistre le code nomenclature dans la colonne GNL_SOUSNOMEN
//-----------------------------------------------------------------
                            G_NLIG.Cells[SGN_SSNom, ARow] := st1;
                            TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('GNL_SOUSNOMEN',st1);
                            OkArt:=True ;
                            end;
                        end
                    else

//-----------------------------------------------------------------
//  Code article trouve, on regarde s'il a des nomenclatures associées
//  si oui et plusieurs, ouverture d'une fenetre de selection
//  si oui et une seule, on recupere directement son code
//-----------------------------------------------------------------
                    // Modif BTP le 15/05/2001
                    if (TOBArt.GetValue('GA_TYPEARTICLE') = 'NOM') or (TOBArt.GetValue('GA_TYPEARTICLE') = 'OUV') then
                        begin
                        st1 := '';
                        Q := OpenSQL('Select GNE_NOMENCLATURE from NOMENENT Where GNE_ARTICLE="' +
                                     TOBArt.GetValue('GA_ARTICLE') + '"', True);
                        if Q.RecordCount > 1 then
                            begin
                            libelle := traduireMemoire( 'Choix d''un ouvrage');
                            st1 := Choisir(libelle,'NOMENENT','GNE_LIBELLE',
                            'GNE_NOMENCLATURE','GNE_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'"','');
                            end
                        else
                            if Q.RecordCount = 1 then
                                begin
                                TF1 := Q.FindField('GNE_NOMENCLATURE');
{$IFDEF EAGLCLIENT}
                                st1 := TF1.AsString ;
{$ELSE}
                                st1 := TF1.Text;
{$ENDIF}
                                end;
                        Ferme(Q);
                        Libelle:='';
                        if st1 <> '' then
                            begin
                            // recupération du libellé
                            Q := OpenSQL('Select GNE_LIBELLE from NOMENENT Where GNE_NOMENCLATURE="' + st1 + '"', True,-1, '', True);
                            Libelle := Q.FindField('GNE_Libelle').AsString;
                            Ferme(Q);
//-----------------------------------------------------------------
//  on enregistre le code nomenclature dans la colonne GNL_SOUSNOMEN
//-----------------------------------------------------------------
                            G_NLIG.Cells[SGN_SSNom, ARow] := st1;
                            TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('GNL_SOUSNOMEN',st1);
                            OkArt:=True ;
                            end;
                        end
                        else
                        begin
                        OkArt:=True ;
                        end;
                    end;
         traAucun : begin
//-----------------------------------------------------------------
// Recherche sur code via LookUp ou Recherche avancée
//-----------------------------------------------------------------
										Larticle := Trim(G_NLIG.Cells[SGN_Comp, ARow]);
                    ART.text := '';
                    if TypeSaisie = TTNOUV then
                    begin
                    StWhere := GetTypeArticleBTP;
                    end else
                    begin
                    StWhere := GetTypeArticleParc;
                    end;
                    StWhere := 'GA_CODEARTICLE=' + Larticle+';XX_WHERE=AND '+stWhere;

                    DispatchRecherche (ART, 1, '',stWhere,'');
                    if ART.Text <> '' then
                        begin
                        TOBArt.PutValue('GA_ARTICLE',ART.Text);
                        TOBArt.LoadDB();

                        // traitement si ouvrage
                        if (TOBArt.GetValue('GA_TYPEARTICLE') = 'NOM') or (TOBArt.GetValue('GA_TYPEARTICLE') = 'OUV') then
                           begin
                           st1 := '';
                           Q := OpenSQL('Select GNE_NOMENCLATURE from NOMENENT Where GNE_ARTICLE="' +
                                       TOBArt.GetValue('GA_ARTICLE') + '"', True,-1, '', True);
                           if Q.RecordCount > 1 then
                              begin
                              libelle := traduireMemoire( 'Choix d''un ouvrage');
                              st1 := Choisir(libelle,'NOMENENT','GNE_LIBELLE',
                              'GNE_NOMENCLATURE','GNE_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'"','');
                              end
                           else
                              if Q.RecordCount = 1 then
                                 begin
                                 TF1 := Q.FindField('GNE_NOMENCLATURE');
{$IFDEF EAGLCLIENT}
                                 st1 := TF1.AsString ;
{$ELSE}
                                 st1 := TF1.Text;
{$ENDIF}
                                 end;
                           Ferme(Q);
                           Libelle:='';
                           if st1 <> '' then
                              begin
                              // recupération du libellé
                              Q := OpenSQL('Select GNE_LIBELLE from NOMENENT Where GNE_NOMENCLATURE="' + st1 + '"', True,-1, '', True);
                              Libelle := Q.FindField('GNE_Libelle').AsString;
                              Ferme(Q);
//-----------------------------------------------------------------
//  on enregistre le code nomenclature dans la colonne GNL_SOUSNOMEN
//-----------------------------------------------------------------
                              G_NLIG.Cells[SGN_SSNom, ARow] := st1;
                              TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('GNL_SOUSNOMEN',st1);
                              OkArt:=True ;
                              end;
                           end
                        else
                           begin
                           if TOBArt.GetValue('GA_STATUTART')='GEN' then
                              begin
                              if ChoisirDimension (TOBArt.GetValue ('GA_ARTICLE'), TOBArt) then
                                 OkArt := True;
                              end else OkArt := true;

                           if (TOBArt.GetValue('GA_TYPEARTICLE') = 'ARP') then
                              begin
                              st1 := '';
                              Q := OpenSQL('Select GNE_NOMENCLATURE from NOMENENT Where GNE_NOMENCLATURE="' +
                                     TOBArt.GetValue('GA_ARTICLE') + '"', True);
                              TF1 := Q.FindField('GNE_NOMENCLATURE');
{$IFDEF EAGLCLIENT}
                              st1 := TF1.AsString ;
{$ELSE}
                              st1 := TF1.Text;
{$ENDIF}
                              Ferme(Q);
                              if st1 <> '' then
                                begin
//-----------------------------------------------------------------
//  on enregistre le code nomenclature dans la colonne GNL_SOUSNOMEN
//-----------------------------------------------------------------
                                G_NLIG.Cells[SGN_SSNom, ARow] := st1;
                                TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('GNL_SOUSNOMEN',st1);
                                OkArt:=True ;
                                end;
                              end;
                           end;
                        end;
                    end ;
        traGrille : begin
//-----------------------------------------------------------------
// Forcement objet dimension avec saisie obligatoire
//if GetArticleLookUp (GF_ARTICLE, 'GA_STATUTART = "DIM"') then
//-----------------------------------------------------------------
                    if ChoisirDimension (TOBArt.GetValue ('GA_ARTICLE'), TOBArt) then
                        OkArt := True;
                    if (TOBArt.GetValue('GA_TYPEARTICLE') = 'ARP') then
                        begin
                        st1 := '';
                        Q := OpenSQL('Select GNE_NOMENCLATURE from NOMENENT Where GNE_NOMENCLATURE="' +
                                     TOBArt.GetValue('GA_ARTICLE') + '"', True,-1, '', True);
                        TF1 := Q.FindField('GNE_NOMENCLATURE');
{$IFDEF EAGLCLIENT}
                        st1 := TF1.AsString ;
{$ELSE}
                        st1 := TF1.Text;
{$ENDIF}
                        Ferme(Q);
                        if st1 <> '' then
                            begin
//-----------------------------------------------------------------
//  on enregistre le code nomenclature dans la colonne GNL_SOUSNOMEN
//-----------------------------------------------------------------
                            G_NLIG.Cells[SGN_SSNom, ARow] := st1;
                            TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('GNL_SOUSNOMEN',st1);
                            OkArt:=True ;
                            end;
                        end;
                    end;
         end ; // Case

    if (Tobart.GetValue('GA_FERME') = 'X') then
    begin
      MsgBox.Execute(9, Caption, '');
      OkArt := False;
    end;

    if (OkArt) then
    begin
         // Modif BTP
    if (Tobart.getvalue('GA_TYPEARTICLE') <> 'POU')  or
       ((Tobart.getvalue('GA_TYPEARTICLE') = 'POU') and  (Ipercent = 0)) then
         // -----
        begin
        if (ARow = 1) and (TOBLig.Detail.Count = 0) then
            begin
            G_NLIG.Objects[SGN_NumLig, ARow] := TOB.Create ('NOMENLIG', TOBLig, ARow-1) ;
            ChampSupNomen (TOB(G_NLIG.Objects[SGN_NumLig, ARow]));
            TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('GNL_NOMENCLATURE',GNE_NOMENCLATURE.Text);
            TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('GNL_NUMLIGNE',1);
            end;
//--------------------------------------------------------------------------
//  on va charger la fiche article dans TobArticles
//--------------------------------------------------------------------------
        G_NLIG.Objects[SGN_IdComp, ARow] := TOB.Create('ARTICLE',nil,ARow);
        Tob(G_NLIG.Objects[SGN_IdComp, ARow]).Dupliquer(TOBArt, True, True);
        TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('GNL_CODEARTICLE', TOBArt.GetValue('GA_CODEARTICLE'));
        G_NLIG.Cells[SGN_Comp, ARow] := TOBArt.GetValue('GA_CODEARTICLE');
        TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('GA_TYPEARTICLE', TOBArt.GetValue('GA_TYPEARTICLE'));
        TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('GNL_ARTICLE', TOBArt.GetValue('GA_ARTICLE'));
        TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('BNP_TYPERESSOURCE', TOBArt.GetValue('BNP_TYPERESSOURCE'));
        //
        G_NLIG.Cells[SGN_IdComp, ARow] := TOBArt.GetValue('GA_ARTICLE');
        if (Pos(TOBArt.GetValue('GA_TYPEARTICLE'),'ARP;OUV')>0) then
        begin
          TOBLoc :=TOB(G_NLIG.Objects[SGN_NumLig, ARow]);
          if TOBart.getValue('GA_TYPEARTICLE')='ARP' then
          begin
            ChargeNomen (TOBArt.GetValue('GA_ARTICLE'),true,TOBLoc);
          end else
          begin
            ChargeNomen (TOBArt.GetValue('GA_CODEARTICLE'),true,TOBLoc);
          end;
        end;
        FormateZoneSaisie(SGN_Comp, ARow);
        if Libelle = '' then
           begin
           TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('GNL_LIBELLE', TOBArt.GetValue('GA_LIBELLE'));
           G_NLIG.Cells[SGN_Lib, ARow] := TOBArt.GetValue('GA_LIBELLE');
           end else
           BEGIN
           TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('GNL_LIBELLE', Libelle);
           G_NLIG.Cells[SGN_Lib, ARow] := Libelle;
           END;
				if TypeSaisie = TTNOUV then
        begin
          if TOB(G_NLIG.Objects[SGN_NumLig, ARow]).getvalue('GNL_SOUSNOMEN') = '' then
          begin
            renseigneValoArt (TOB(G_NLIG.Objects[SGN_NumLig, ARow]),TOBArt,nil);
          end;
          G_NLIG.cells[SGA_UVTE,Arow] := Tob(G_NLIG.Objects[SGN_IdComp, Arow]).getvalue('GA_QUALIFUNITEVTE');
          TOB(G_NLIG.Objects[SGN_NumLig, ARow]).Putvalue('UNITENOMENC', G_NLIG.cells[SGA_UVTE,Arow]);
          //
          if TOB(G_NLIG.Objects[SGN_NumLig, ARow]).getvalue('GNL_SOUSNOMEN') <> '' then
              begin
              chargePrixOuvrage (TOB(G_NLIG.Objects[SGN_NumLig, ARow]),
                                TOB(G_NLIG.Objects[SGN_NumLig, ARow]).getvalue('GNL_SOUSNOMEN'),Arow);
              if Tob(G_NLIG.Objects[SGN_IdComp, ARow]).getvalue('GA_TYPEARTICLE') = 'ARP' then
                begin
                RecupInfoPrixPose (TOB(G_NLIG.Objects[SGN_NumLig, ARow]));
                end;
              end;
          if (Tobart.getvalue('GA_TYPEARTICLE') = 'POU') then
              begin
              Ipercent := Arow;
              end;
          Tlocal := calculeMontantLig (Arow);
          CalculeMontantOuv ();
          AlimenteValeurOuv ();
          TMontantLig := TLocal;
          AffecteValArticle (G_NLIG.Row);
          G_NLIG.InvalidateCell (SGN_Prix,G_NLIG.row);
        end;
       end else
       begin
          // On ne peut pas insérer 2 articles gérés en pourcentage
       G_NLIG.Cells[SGN_Comp, ARow] := '';
       G_NLIG.Col := SGN_Comp;
       end;
        // ----------------
    end
    else
    begin
        G_NLIG.Cells[SGN_Comp, ARow] := '';
        G_NLIG.Col := SGN_Comp;
    end;
TOBArt.Free;
Result := OkArt;
end;

//=============================================================================
//  Gestion du TreeView
//=============================================================================
//-----------------------------------------------------------------------------
//  Positionnement de l'icone de ImageList1 associée à un element.
//          0 : Nomenclature, non étendue
//          1 : Nomenclature, etendue
//          2 : article
//-----------------------------------------------------------------------------
procedure TFNomenLig.SetImages(TV : TTreeView; TN : TTreeNode);
var
    tn1 : TTreeNode;
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
procedure TFNomenLig.AfficheTVNLIG;
var i_ind1  : integer;
    tn1     : TTreeNode;
begin
AfficheTreeView(TV_NLIG, TOBLig, 'LIBELLE');
TV_NLIG.TopItem.Text := GNE_NOMENCLATURE.Text;
TV_NLIG.TopItem.Expand(False);
SetImages(TV_NLIG, nil);

  tn1 := TV_NLIG.TopItem ;

  for i_ind1:= 1 to G_NLIG.rowcount-1 do
  begin
    while tn1.level <> 1  do
    begin
      tn1 :=tn1.GetNext;
      if tn1 = nil  then break;
    end;
    if tn1 = nil    then break;
    G_NLIG.Objects[SGN_Comp, i_ind1] := tn1;
    tn1 := tn1.GetNext;
    if tn1 = nil    then break;
  end;

end;

//=============================================================================
//  Gestion des champs
//=============================================================================
function TFNomenLig.TraiterComposant(ARow: LongInt) : boolean;
var
    st1{, st2} : string;
    TOBTemp : TOB;
    tn1 : TTreeNode;
    // i_Rep : Integer;
{$IFDEF BTP}
    Tlocal : T_Valeurs;
{$ELSE}
		i_rep : integer;
    st2: string;
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
    G_NLIG.Col := SGN_Comp; G_NLIG.Row := ARow;
    Result := False;
    Exit;
    END;
//--------------------------------------------------------------------------
//  Code composant renseigne. s'il est different du code identifiant article
//  ou que ce code identifiant est vide, c'est un changement de composant
//  ou une nouvelle ligne
//--------------------------------------------------------------------------
st1 := Trim(Copy(G_NLIG.Cells[SGN_IdComp,ARow], 0, Length(G_NLIG.Cells[SGN_Comp,ARow])));
if (G_NLIG.Cells[SGN_IdComp,ARow] = '') or (G_NLIG.Cells[SGN_Comp,ARow] <> st1) then
    begin
{$IFDEF BTP}
//    TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('GNL_PRIXFIXE',valeur(G_NLIG.Cells[SGN_Prix, ARow]));
{$ENDIF}

    TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('GNL_QTE',valeur(G_NLIG.Cells[SGN_Qte, ARow]));
    TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('GNL_JOKER',G_NLIG.Cells[SGN_Joker, ARow]);
//--------------------------------------------------------------------------
//  on passe par la recherche d'article
//--------------------------------------------------------------------------
    If Not Recherche_Art(ARow) then
    begin
      Result := False;
      Exit;
    end;
//--------------------------------------------------------------------------
//  Code composant unique dans un niveau donné de nomenclature
//--------------------------------------------------------------------------
{$IFNDEF BTP}
    for i_Rep := 0 to G_NLIG.RowCount - 1 do
        begin
        if G_NLIG.Cells[SGN_Comp, i_Rep] = '' then Break;
        st1 := G_NLIG.Cells[SGN_Comp, i_Rep] + G_NLIG.Cells[SGN_IdComp, i_Rep] + G_NLIG.Cells[SGN_SSNom, i_Rep];
        st2 := G_NLIG.Cells[SGN_Comp, ARow] + G_NLIG.Cells[SGN_IdComp, ARow] + G_NLIG.Cells[SGN_SSNom, ARow];
        if (i_Rep <> ARow) and (st1 = st2) then
            BEGIN
            //  Message code composant unique dans la nomenclature
            MsgBox.Execute (6,Caption,'') ;
            G_NLIG.Col := SGN_Comp; G_NLIG.Row := ARow;
            TOB(G_NLIG.Objects[SGN_NumLig, ARow]).Free;
            G_NLIG.Objects[SGN_NumLig, ARow] := nil;
// BB, correction fiche 10352
            G_NLIG.Rows[ARow].Clear;
            InsertLigne(ARow, True);
// BB, fin correction fiche 10352
            Result := False;
            Exit;
            END;
        end;
{$ENDIF}

//--------------------------------------------------------------------------
//  Mise à jour du treeview
//--------------------------------------------------------------------------
    if (ARow = 1) then
    begin
//--------------------------------------------------------------------------
//  premiere ligne, on insere le premier noeud enfant
//--------------------------------------------------------------------------
        G_NLIG.Objects[SGN_Comp, ARow] := TV_NLIG.Items.AddChildFirst(TV_NLIG.Items[0], G_NLIG.Cells[SGN_Comp, ARow]);
    end
    else if (G_NLIG.Cells[SGN_Comp, ARow + 1] = '') or (ARow = G_NLIG.RowCount - 1) then
//--------------------------------------------------------------------------
//  derniere  on ajoute un noeud enfant en fin de liste
//--------------------------------------------------------------------------
        G_NLIG.Objects[SGN_Comp, ARow] := TV_NLIG.Items.AddChild(TV_NLIG.Items[0], G_NLIG.Cells[SGN_Comp, ARow])
    else
        begin
        tn1 := TTreeNode(G_NLIG.Objects[SGN_Comp, ARow - 1]);
        tn1 := tn1.GetNextSibling;
        G_NLIG.Objects[SGN_Comp, ARow] := TV_NLIG.Items.Insert(tn1, G_NLIG.Cells[SGN_Comp, ARow]);
        TTreeNode(G_NLIG.Objects[SGN_Comp, ARow]).Text := G_NLIG.Cells[SGN_Comp, ARow];
        end;
//--------------------------------------------------------------------------
//  on va charger l'eventuelle sous nomenclature complete
//--------------------------------------------------------------------------
    TOBTemp := TOB(G_NLIG.Objects[SGN_NumLig, ARow]);
//    ChargeNomen(TOBTemp.GetValue('GNL_SOUSNOMEN'), True, TOBTemp);
    AffecteLibelle(TOBTemp, 'GNL_CODEARTICLE');
    //
//    renseigneValoArt (TOB(G_NLIG.Objects[SGN_NumLig, Arow]),Tob(G_NLIG.Objects[SGN_IdComp, Arow]),nil);
    //
   ChampSupNomen (TOBTemp);
   if TypeSaisie = TTNOUV then
   begin
     TOBTEMP.SetDouble('GNL_QTESAIS',TOBTEMP.GetDouble('GNL_QTE'));
     if TOBTemp.GetValue('GNL_SOUSNOMEN') = '' then
        begin
        renseigneValoArt (TOBTemp,Tob(G_NLIG.Objects[SGN_IdComp, Arow]),nil);
        end;
     if TOBTemp.getvalue('GA_TYPEARTICLE') = 'POU' then
        begin
        IPercent := Arow;
        end;
     if TOBTemp.GetValue('GNL_SOUSNOMEN') <> '' then
        begin
        chargePrixOuvrage (TOBTemp,TOBTemp.getvalue('GNL_SOUSNOMEN'),Arow);
        if TOBTemp.getvalue('GA_TYPEARTICLE') = 'ARP' then
          begin
          RecupInfoPrixPose (TOBTemp);
          end;
        end;
     TOBTEMP.PutLigneGrid (G_NLIG,Arow,false,false,LesColNLIG);
   end;
//--------------------------------------------------------------------------
//  on rafraichit l'ensemble du treeview
//--------------------------------------------------------------------------
    TV_NLIG.Selected := TTreeNode(G_NLIG.Objects[SGN_Comp, ARow]);
{$IFDEF BTP}
      if TypeSaisie = TTNOUV then
      begin
        Tlocal := calculeMontantLig (Arow);
        CalculeMontantOuv;
        AlimenteValeurOuv ();
        TMontantLig := TLocal;
        AffecteValArticle (Arow);
      end;
{$ENDIF}
    end;
end;

function TFNomenLig.TraiteQteSais (Acol,Arow : integer ) : boolean;
var TOBL : TOB;
    LastQTe : double;
begin
  result := true;
  if (not GetParamSocSecur('SO_BTGESTQTEAVANC',false)) then exit;
  TOBL := Tob(G_NLIG.objects[SGN_NumLig,Arow]);
  LastQte := TOBL.GetDouble('GNL_QTE');

  if Acol = SGN_QTESAIS then
  begin
    TOBL.SetDouble('GNL_QTESAIS', Valeur(G_NLIG.Cells[ACol, ARow]))
  end else if Acol = SGN_PERTE then
  begin
    TOBL.SetDouble('GNL_PERTE', Valeur(G_NLIG.Cells[ACol, ARow]))
  end else if Acol = SGN_RENDEMENT then
  begin
    TOBL.SetDouble('GNL_RENDEMENT', Valeur(G_NLIG.Cells[ACol, ARow]));
  end;
  TOBL.SetDouble('GNL_QTE',CalculeQteFact(TOBL));
  if (Acol = SGN_QTESAIS) and (TOBL.GetDouble('GNL_QTE')<> LastQte) then
  begin
    if TOBL.GetValue('GA_TYPEARTICLE')='PRE' then
    begin
      TobL.PutValue('GA_HEURE',TobL.GetValue('GNL_QTE'));
    end;
  end;
  TOBL.PutLigneGrid (G_NLIG,Arow,false,false,LesColNLIG);
end;

function TFNomenLig.TraiterQuantite(ARow: LongInt) : boolean;
begin
  Result := True;
  {$IFNDEF BTP}
    if Valeur(G_NLIG.Cells[SGN_Qte, ARow]) <= 0 then
    BEGIN
      //  Message quantité >= 0
      MsgBox.Execute (2,Caption,'') ;
      G_NLIG.Cells [SGN_Qte, Arow] := StrS(1, 4);
      G_NLIG.Col := SGN_Qte; G_NLIG.Row := ARow;
      Result := False;
    END;
  {$ELSE}
    if Tob(G_NLIG.objects[SGN_IdComp,Arow]).GetValue('GA_TYPEARTICLE')='PRE' then
    begin
      Tob(G_NLIG.objects[SGN_Numlig,Arow]).PutValue('GA_HEURE',Tob(G_NLIG.objects[SGN_Numlig,Arow]).GetValue('GNL_QTE'));
    end;
  {$ENDIF}
end;
//
// Modif BTP 06/02/2001
//
{$IFDEF BTP}
function TFNomenLig.TraiterPrix(ARow: LongInt) : boolean;
begin
Result := True;
if Valeur(G_NLIG.Cells[SGN_Qte, ARow]) < 0 then
    G_NLIG.Cells [SGN_Prix, ARow] := strs (0,V_PGI.OkDecP);
Tob(G_NLIG.objects[SGN_Numlig,Arow]).PutValue('GNL_PRIXFIXE',Valeur(G_NLIG.Cells [SGN_Prix, ARow]));
end;
{$ENDIF}

function TFNomenLig.TraiterJoker(ARow: LongInt) : boolean;
begin
Result := True;
{$IFDEF BTP}
  G_NLIG.Cells[SGN_Joker, ARow]:='N';
{$ENDIF}

G_NLIG.Cells[SGN_Joker, ARow]:=Uppercase (G_NLIG.Cells[SGN_Joker, ARow]);
if (G_NLIG.Cells[SGN_Joker, ARow] <> 'O') and (G_NLIG.Cells[SGN_Joker, ARow] <> 'N') then
    BEGIN
//  Message code = à O ou N
    MsgBox.Execute (4,Caption,'') ;
    G_NLIG.Cells [SGN_Joker, ARow] := 'N';
    G_NLIG.Col := SGN_Joker; G_NLIG.Row := ARow;
    Result := False;
    END;
end;

procedure TFNomenLig.FormateZoneSaisie(ACol,ARow: LongInt);
var St,StC : String ;
{$IFDEF BTP} // DBR Fiche 10147
    i_sep : integer;
{$ENDIF} // BTP
begin
St:=G_NLIG.Cells[ACol,ARow] ; StC:=St ;
if (ACol=SGN_Comp) then StC:=uppercase(Trim(St)) else
    if (ACol=SGN_Joker) then StC:=uppercase(Trim(St)) else
        if (ACol=SGN_Qte) then
        begin
        StC := StrF00(Valeur(St), 4); // DBR Fiche 10147
{        i_sep := pos('.',St); // StrF00 fait tout le travail
        if i_sep > 0 then
           begin
           st := copy (st,1,i_sep-1)+','+copy(st,i_sep+1,255);
           end;
        StC:=StrS(Valeur(St),V_PGI.OkdecQ) }
{$IFDEF BTP}
        end else
        begin
        if ACol=SGN_Prix then
            begin
            i_sep := pos('.',St);
            if i_sep > 0 then
               begin
               st := copy (st,1,i_sep-1)+','+copy(st,i_sep+1,255);
               end;
            StC:=StrS(valeur(St),V_PGI.OkDecP) ;
            end;
{$ENDIF}
        end;
G_NLIG.Cells[ACol,ARow]:=StC ;
end;

//=============================================================================
//  Gestion des colonnes
//=============================================================================

//=============================================================================
//  Gestion des lignes de la grid
//=============================================================================
//  Insertion d'une ligne avec création de la tob associée et renumérotation
//=============================================================================
procedure TFNomenLig.InsertLigne (ARow : Longint; Ajout : boolean) ;
var TOBL    : TOB;
    TOBL_1  : TOB;
    TobTemp : TOB;
    Ztype   : Boolean ;
    Acol    : Integer;
begin
  if Action=taConsult then Exit ;
  if ARow < 1 then Exit ;

  // Modif BTP
  if Btypearticle.Visible then BtypeArticle.Visible := false;
  // -----
  TOBL   := TOB(G_NLIG.Objects[SGN_NumLig, ARow]);
  TOBL_1 := TOB(G_NLIG.Objects[SGN_NumLig, ARow-1]);
  //
  if (Arow >= 1) then
  begin
    if (TOBL  <>nil) and (TOBL.GetValue   ('GNL_CODEARTICLE')='') then exit;
    if (TOBL_1<>nil) and (TOBL_1.GetValue ('GNL_CODEARTICLE')='') then exit;
  end;

  if ARow = G_NLIG.RowCount - 1 then G_NLIG.RowCount := G_NLIG.RowCount + 50;

  if (not Ajout) and (ARow > TOBLig.Detail.Count) then Exit;

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
  G_NLIG.Objects[SGN_NumLig, ARow] := TOB.Create ('NOMENLIG', TOBLig, ARow-1) ;


  ChampSupNomen (TOB(G_NLIG.Objects[SGN_NumLig, ARow]), False); // DBR Fiche 10995

  TOBL := TOB(G_NLIG.Objects[SGN_NumLig, ARow]);

  TOBL.PutValue('GNL_NOMENCLATURE',GNE_NOMENCLATURE.Text);
  TOBL.PutValue('GNL_NUMLIGNE'    ,StrToInt(G_NLIG.Cells[0, ARow]));
  TOBL.PutValue('GNE_DOMAINE'     ,Domaine);

  AffecteLibelle(TOBL, 'GNL_CODEARTICLE');

  Renumerote(ARow + 1);

  Ztype       := false;
  Acol        := SGN_Comp;
  G_NLIG.col  := Acol;
  G_NLIGCellEnter (Self,Acol,Arow,Ztype);
  G_NLIG.MontreEdit;
  G_NLIG.SynEnabled := True;

  {$IFDEF BTP}
  if TypeSaisie = TTNOUV then
  begin
    CalculeMontantOuv;
    AlimenteValeurOuv ();
  end;
  {$ENDIF}

end;

//=============================================================================
//  Initialisation des colonnes de la ligne
//=============================================================================
procedure TFNomenLig.InitialiseLigne (ARow : Longint) ;
var i_ind : integer ;
begin

  for i_ind := 1 to G_NLIG.ColCount-1 do
    if (i_ind = SGN_Joker)  then G_NLIG.Cells [i_ind, ARow] := 'N'                  else
    if (i_ind = SGN_Qte)    then G_NLIG.Cells [i_ind, ARow] := strs (1, 4)          else
    if i_ind = SGN_Prix     then G_NLIG.Cells [i_ind, ARow] := strs (0,V_PGI.OkDecP) else G_NLIG.Cells [i_ind, ARow]:='' ;

end;

//=============================================================================
//  Suppression d'une ligne et de la tob associée et renumérotation
//=============================================================================
procedure TFNomenLig.SupprimeLigne (ARow : Longint) ;
var bLocal      : boolean;
    TobInd      : TOB;
    RepPoste    : string;
    RepServeur  : string;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if (ARow > TOBLig.Detail.Count) then Exit;
{$IFDEF BTP}
if Btypearticle.Visible then BtypeArticle.Visible := false;
{$ENDIF}
G_NLIG.CacheEdit; G_NLIG.SynEnabled := False;

if G_NLIG.Objects[SGN_NumLig, ARow] <> nil then TOB(G_NLIG.Objects[SGN_NumLig, ARow]).Free;

G_NLIG.DeleteRow (ARow);

if G_NLIG.RowCount < NbRowsInit then G_NLIG.RowCount := NbRowsInit;
G_NLIG.MontreEdit; G_NLIG.SynEnabled := True;
AfficheTVNLIG;

{$IFDEF BTP}
	Renumerote (Arow);
  if TypeSaisie = TTNOUV then
  begin
    CalculeMontantOuv;
    AlimenteValeurOuv ();
  end;
{$ELSE}
Renumerote(1);
{$ENDIF}

if (TOBLig.Detail.Count = 0) then InsertLigne (G_NLIG.Row, true);
if G_NLIG.row > TobLig.detail.count then
begin
   G_NLIG.row := G_NLIG.row -1;
   blocal := false;
   G_NLIGRowEnter(Self, G_NLIG.row, Blocal, blocal);
end;
// ----
AffectePied(G_NLIG.Row);
//
//--- Suppression du fichier metré associé à la ligne...
//

  if (G_NLIG.Cells[SGN_Comp, G_NLIG.Row] <> '') then
  begin
    TobInd        := Tob(G_NLIG.Objects[SGN_NumLig, G_NLIG.Row]);
    MetreArticle  := TMetreArt.CreateArt(TobInd.getvalue('GA_TYPEARTICLE'), TobInd.getvalue('GNL_ARTICLE'));
    if MetreArticle.ControleMetre then
    begin
      RepServeur  := MetreArticle.RepMetre + MetreArticle.fFileName;
      RepPoste    := MetreArticle.RepMetreLocal + MetreArticle.fFileName;
      MetreArticle.SupprimeFichierXLS(RepPoste);
      //Penser à mettre ici la suppression du fichier du répertoire serveur quand on aura géré les fixhier XLS différents par sous-détail pour un même article !!!!!
    end;
    FreeAndNil(MetreArticle);
  end;

END;


//=============================================================================
//  Renumérotation des lignes à partir de la ligne donnée
//=============================================================================
procedure TFNomenLig.Renumerote (ARow : Longint) ;
var
    i_ind1 : integer ;
begin
for i_ind1 := ARow to G_NLIG.RowCount-1 do
    begin
    if G_NLIG.Cells [SGN_Comp, i_ind1] = '' then Exit;
    G_NLIG.Cells [0, i_ind1] := IntToStr(i_ind1);
    TOB(G_NLIG.Objects[SGN_NumLig, i_ind1]).PutValue('GNL_NUMLIGNE',StrToInt(G_NLIG.Cells[0, i_ind1]));
    end;
end;

//=============================================================================
//  Mise à jour du pied d'écran
//=============================================================================
procedure TFNomenLig.AffectePied (ARow : Longint) ;
var i_ind1, i_ind2 : integer;
    LibelleDim, GrilleDim : string;

begin
EffaceDimensions;
if G_NLIG.Cells[SGN_Comp, ARow] = '' then
    begin
    GNL_SOUSNOMEN.Text := '';
    GNL_SOUSNOMEN.Visible := False;
    TGNL_SOUSNOMEN.Visible := False;
    Exit;
    end;
GNL_SOUSNOMEN.Text := TOB(G_NLIG.Objects[SGN_NumLig, ARow]).GetValue('GNL_SOUSNOMEN');
if GNL_SOUSNOMEN.Text <> '' then
    begin
    GNL_SOUSNOMEN.Visible := True;
    TGNL_SOUSNOMEN.Visible := True;
    Exit;
    end;

GNL_SOUSNOMEN.Text := '';
GNL_SOUSNOMEN.Visible := False;
TGNL_SOUSNOMEN.Visible := False;
if TOB(G_NLIG.Objects[SGN_IdComp, ARow]).GetValue('GA_STATUTART') = 'DIM' then
    begin
    LibelleDim := '';
    GrilleDim := '';
    LibelleDimensions('', TOB(G_NLIG.Objects[SGN_IdComp, ARow]), LibelleDim, GrilleDim);
    TGA_CODEDIM1.Caption := ReadTokenSt(LibelleDim);
    TGA_CODEDIM2.Caption := ReadTokenSt(LibelleDim);
    TGA_CODEDIM3.Caption := ReadTokenSt(LibelleDim);
    TGA_CODEDIM4.Caption := ReadTokenSt(LibelleDim);
    TGA_CODEDIM5.Caption := ReadTokenSt(LibelleDim);
    if TGA_CODEDIM1.Caption <> '' then
        begin
        TGA_CODEDIM1.Caption := TGA_CODEDIM1.Caption + ' : ' + ReadTokenSt(GrilleDim);
        TGA_CODEDIM1.Visible := True;
        end;
    if TGA_CODEDIM2.Caption <> '' then
        begin
        TGA_CODEDIM2.Caption := TGA_CODEDIM2.Caption + ' : ' + ReadTokenSt(GrilleDim);
        TGA_CODEDIM2.Visible := True;
        end;
    if TGA_CODEDIM3.Caption <> '' then
        begin
        TGA_CODEDIM3.Caption := TGA_CODEDIM3.Caption + ' : ' + ReadTokenSt(GrilleDim);
        TGA_CODEDIM3.Visible := True;
        end;
    if TGA_CODEDIM4.Caption <> '' then
        begin
        TGA_CODEDIM4.Caption := TGA_CODEDIM4.Caption + ' : ' + ReadTokenSt(GrilleDim);
        TGA_CODEDIM4.Visible := True;
        end;
    if TGA_CODEDIM5.Caption <> '' then
        begin
        TGA_CODEDIM5.Caption := TGA_CODEDIM5.Caption + ' : ' + ReadTokenSt(GrilleDim);
        TGA_CODEDIM5.Visible := True;
        end;
    i_ind1 := TGA_CODEDIM1.Width + TGA_CODEDIM2.Width + TGA_CODEDIM3.Width +
              TGA_CODEDIM4.Width + TGA_CODEDIM5.Width;
    i_ind2 := PANPIED.Width - i_ind1 - 20;
    i_ind1 := i_ind2 div 4;
    TGA_CODEDIM1.Left := 10;
    TGA_CODEDIM2.Left := TGA_CODEDIM1.Left + TGA_CODEDIM1.Width + i_ind1;
    TGA_CODEDIM3.Left := TGA_CODEDIM2.Left + TGA_CODEDIM2.Width + i_ind1;
    TGA_CODEDIM4.Left := TGA_CODEDIM3.Left + TGA_CODEDIM3.Width + i_ind1;
    TGA_CODEDIM5.Left := TGA_CODEDIM4.Left + TGA_CODEDIM4.Width + i_ind1;
    end;
end;

procedure TFNomenLig.EffaceDimensions;
begin
TGA_CODEDIM1.Caption := '';
TGA_CODEDIM2.Caption := '';
TGA_CODEDIM3.Caption := '';
TGA_CODEDIM4.Caption := '';
TGA_CODEDIM5.Caption := '';
TGA_CODEDIM1.Visible := False;
TGA_CODEDIM2.Visible := False;
TGA_CODEDIM3.Visible := False;
TGA_CODEDIM4.Visible := False;
TGA_CODEDIM5.Visible := False;
end;

procedure TFNomenLig.ChargeTobArticles;
var
    i_ind1 : integer;
    st1 : string;
    QQ : TQuery ;
    TOBL : TOB;
begin
  for i_ind1 := 1 to G_NLIG.RowCount - 1 do
  begin
    if G_NLIG.Cells[SGN_Comp, i_ind1] = '' then Exit;
    G_NLIG.Objects[SGN_IdComp, i_ind1] := TOB.Create('ARTICLE', nil, -1);
    st1 := G_NLIG.Cells[SGN_IdComp, i_ind1];
    QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                   'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                   'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+st1+'"',true,-1, '', True);
    // Tob(G_NLIG.Objects[SGN_IdComp, i_ind1]).SelectDB('"' + st1 + '"', nil);
    TOBL := Tob(G_NLIG.Objects[SGN_IdComp, i_ind1]);
    TobL.SelectDB('',QQ);
    InitChampsSupArticle (TOBL);
    ferme (QQ);
  end;
end;

//=============================================================================
//  Recherche d'une ligne de la grille en fonction d'une valeur chaine de caracteres
//=============================================================================
{$IFNDEF BTP}
//function TFNomenLig.ChercheLigne (Valeur : string) : integer; // DBR Fiche 10995
function TFNomenLig.ChercheLigne (tn1 : TTreeNode) : integer; // DBR Fiche 10995
var
    i_ind1 : integer ;
//    st1, st2 : string; // DBR Fiche 10995
begin
Result := 0;
for i_ind1 := 0 to G_NLIG.RowCount - 1 do
    begin
(*    st1 := G_NLIG.Cells[SGN_Comp, i_ind1] + G_NLIG.Cells[SGN_IdComp, i_ind1] + // DBR Fiche 10995
           G_NLIG.Cells[SGN_SSNom, i_ind1];
    st2 := Trim(Copy(Valeur, 0, Pos('(', Valeur) - 1));
    if Pos(st2, st1) <> 0 then *)
    if tn1 = TtreeNode (G_NLIG.Objects[SGN_Comp, i_ind1]) then  // DBR Fiche 10995
        begin
        Result := i_ind1;
        Exit;
        end;
    end;
end;
{$ENDIF}

{$IFDEF BTP}
Function TFNomenLig.ChercheLigneBtp (tn1 : TTreeNode) : integer;
var i_ind1 : integer ;
begin
Result := 0;
for i_ind1 := 0 to G_NLIG.RowCount - 1 do
    begin
    if tn1= TtreeNode (G_NLIG.Objects[SGN_Comp, i_ind1]) then
        begin
        Result := i_ind1;
        Exit;
        end;
    end;
end;
{$ENDIF}

//=============================================================================
//  Validation d'une ligne
//=============================================================================
function TFNomenLig.ValideLigne(Ou : LongInt) : boolean;
var
    i_ind1 : integer ;
    bfalse : boolean;
    TOBL : TOB;

begin
  //  Validation des différents champs de la ligne
  Result := True;
  if (Ou > TOBLig.Detail.Count - 1) and (G_NLIG.Cells[SGN_IdComp, Ou] = '') then
  begin
    bfalse := False;
    G_NLIGRowExit(TObject(G_NLIG), Ou, bfalse, False);
    Exit;
  end;
  TOBL := TOB(G_NLIG.Objects[SGN_NumLig, Ou]);
  if not TraiterComposant(Ou) then
  begin
    Result := False;
    Exit;
  end;
  if not TraiterQuantite(Ou) then
  begin
    Result := False;
    Exit;
  end;
  if not TraiterJoker(Ou) then
  begin
    Result := False;
    Exit;
  end;
  //  Formattage et mise en place dans la Tob associée
  for i_ind1 := 0 to G_NLIG.ColCount - 1 do
  begin
    FormateZoneSaisie (i_ind1, Ou);
    if i_ind1 = SGN_Comp  then
      TOB(G_NLIG.Objects[SGN_NumLig, Ou]).PutValue('GNL_CODEARTICLE', G_NLIG.Cells[i_ind1, Ou]);
    if i_ind1 = SGN_Lib   then
      TOB(G_NLIG.Objects[SGN_NumLig, Ou]).PutValue('GNL_LIBELLE', G_NLIG.Cells[i_ind1, Ou]);
    if i_ind1 = SGN_Qte   then
      TOB(G_NLIG.Objects[SGN_NumLig, Ou]).PutValue('GNL_QTE', Valeur(G_NLIG.Cells[i_ind1, Ou]));
    if i_ind1 = SGN_QTESAIS   then
      TOB(G_NLIG.Objects[SGN_NumLig, Ou]).PutValue('GNL_QTESAIS', Valeur(G_NLIG.Cells[i_ind1, Ou]));
    if i_ind1 = SGN_PERTE   then
      TOB(G_NLIG.Objects[SGN_NumLig, Ou]).PutValue('GNL_PERTE', Valeur(G_NLIG.Cells[i_ind1, Ou]));
    if i_ind1 = SGN_RENDEMENT   then
      TOB(G_NLIG.Objects[SGN_NumLig, Ou]).PutValue('GNL_RENDEMENT', Valeur(G_NLIG.Cells[i_ind1, Ou]));
    if i_ind1 = SGN_Joker then
      TOB(G_NLIG.Objects[SGN_NumLig, Ou]).PutValue('GNL_JOKER', G_NLIG.Cells[i_ind1, Ou]);
  end;

  {$IFDEF BTP}

  if G_NLIG.Cells[SGN_SSNom, Ou] <> '' then
  begin
    if TOB(G_NLIG.Objects[SGN_NumLig, Ou]).getvalue('QTEDUDETAIL') <> valeur('0') then
    begin
      TOB(G_NLIG.Objects[SGN_NumLig, Ou]).PutValue('LIBELLE',G_NLIG.Cells[SGN_SSNom, Ou] +
                                              ' (' + G_NLIG.Cells[SGN_Qte, Ou] + ' Détail pour '+
      Strs0(TOB(G_NLIG.Objects[SGN_NumLig, Ou]).getvalue('QTEDUDETAIL')) + ' ' +
      TOB(G_NLIG.Objects[SGN_NumLig, Ou]).getvalue('UNITENOMENC') +')');
    end else
      TOB(G_NLIG.Objects[SGN_NumLig, Ou]).PutValue('LIBELLE', G_NLIG.Cells[SGN_SSNom, Ou] +
                                                  ' (' + G_NLIG.Cells[SGN_Qte, Ou] + ')');
  end else
  begin
    TOB(G_NLIG.Objects[SGN_NumLig, Ou]).PutValue('LIBELLE', G_NLIG.Cells[SGN_Comp, Ou] +
                                                  ' (' + G_NLIG.Cells[SGN_Qte, Ou] + ')');
  end;
  {$ELSE}
  if G_NLIG.Cells[SGN_SSNom, Ou] = '' then
    TOB(G_NLIG.Objects[SGN_NumLig, Ou]).PutValue('LIBELLE', G_NLIG.Cells[SGN_Comp, Ou] +
                                                ' (' + G_NLIG.Cells[SGN_Qte, Ou] + ')')
  else
  TOB(G_NLIG.Objects[SGN_NumLig, Ou]).PutValue('LIBELLE', G_NLIG.Cells[SGN_SSNom, Ou] +
                                              ' (' + G_NLIG.Cells[SGN_Qte, Ou] + ')');
  {$ENDIF}
  
  TOB(G_NLIG.Objects[SGN_NumLig, Ou]).PutLigneGrid (G_NLIG,Ou,false,false,LesColNLIG);
  if SGN_Prix <> -1 then G_NLIG.InvalidateCell (SGN_Prix,Ou);

end;

//=============================================================================
//  Gestion des boutons
//=============================================================================
procedure TFNomenLig.BValiderClick(Sender: TObject);
var
    i_ind1 : integer;
    TOBTemp, TOBTempErr : TOB;
    Q : TQuery;
// BBI : GP Light
{$IFDEF GPAOLIGHT}
    LastIdent : integer;
    TobTempL, TOBA : TOB;
    stSQL : string;
{$ENDIF}
		TheRef : string;
    //
    RepPoste    : string;
    RepServeur  : string;
// BBI Fin
begin
{
//  Validation de la ligne en cours dans la Grid
if (not (G_NLIG.Cells[SGN_Comp,G_NLIG.row] = '')) and ( not ValideLigne(G_NLIG.Row)) then
    begin
    ModalResult := 0;
    Exit;
    end;
}
  SendMessage(G_NLIG.Handle, WM_KeyDown, VK_TAB, 0);
  TheRef := GNE_NOMENCLATURE.Text;
  //  Verification de la non circularite d'appels de nomenclatures
  TOBTempErr := TOB.Create(GNE_NOMENCLATURE.Text, nil, -1);
  TOBTempErr.AddChampSup('GNL_SOUSNOMEN', True);
  TOBTempErr.PutValue('GNL_SOUSNOMEN', GNE_NOMENCLATURE.Text);

  if not Circularite(GNE_NOMENCLATURE.Text, TOBLig, TOBTempErr) then
    begin
    FNomenErr := TFNomenErr.Create(Application);
//  Message nomenclature contient une reference circulaire
    MsgBox.Execute (7,Caption,'') ;
    AfficheTreeView(FNomenErr.TV_ERREUR, TOBTempErr, 'GNL_SOUSNOMEN');
    FNomenErr.ShowModal;
    FNomenERR.Free;
    G_NLIG.Col := SGN_Comp; G_NLIG.Row := 1;
    ModalResult := 0;
    TOBTempErr.free;
    Exit;
    end;
  //
  //  Suppression des records de l'ancienne version de la nomenclature
  {New}
  TOBTemp := TOB.Create ('NOMENENT',nil,-1);
  try
  Q := OpenSQL('Select * from NOMENENT where GNE_NOMENCLATURE="' + GNE_NOMENCLATURE.Text + '"', true,-1, '', True);
  TOBTemp.selectDb ('',Q);
  ferme (Q);
  TOBTemp.putValue('GNE_QTEDUDETAIL',GNE_QTEDUDETAIL.Value);
  TOBTEmp.UpdateDB;
  finally
	TOBTemp.free;
  end;
  {Fin New}

  try
    ExecuteSQL('DELETE FROM NOMENLIG WHERE GNL_NOMENCLATURE="' + GNE_NOMENCLATURE.Text + '"');
  //  Création des records de la nouvelle version de la nomenclature
  for i_ind1 := 0 to TOBLig.Detail.Count - 1 do
      begin
      if (TOBLIG.Detail[i_ind1].getvalue('GNL_ARTICLE') <> '') then
         begin
         TOBLig.Detail[i_ind1].SetAllModifie(true);
         TOBLig.Detail[i_ind1].InsertOrUpdateDB();
         //Gestion des métrés...
         MetreArticle := TMetreArt.CreateArt(TOBLig.Detail[i_Ind1].GetValue('GA_TYPEARTICLE'),TOBLIG.Detail[i_ind1].getvalue('GNL_ARTICLE'));
         if MetreArticle.ControleMetre then
         begin
            RepServeur  := MetreArticle.RepMetre + MetreArticle.fFileName;
            RepPoste    := MetreArticle.RepMetreLocal + MetreArticle.fFileName;
            //Copie du fichier local sur le serveur
            MetreArticle.CopieLocaltoServeur(RepPoste, RepServeur);
            //suppression de la copie locale
            MetreArticle.SupprimeFichierXLS(RepPoste);
         end;
         //
         FreeAndNil(MetreArticle);
      end;
    end;
  finally
  end;

  TOBTempErr.free;

  // Modified by f.vautrain 03/10/2017 12:10:39 - FS#2730 - TREUIL - Modifier un sous-détail dans la bibliothèque n'affecte pas la date de modif. de l'ouvrage
  if Data_Modified Then MajDateModifOuvrage(GNE_NOMENCLATURE.Text);

  Data_Modified := False;

  Close;
end;

procedure TFNomenLig.bPrintClick(Sender: TObject);
Var {TL : TList ;
    TT : TStrings ;}
    SQL : String ;
    TOBEtat : TOB;
begin
{$IFNDEF PGIMAJVER}
{$IFNDEF BTP}
TL:=TList.Create ;
TT:=TStringList.Create ;
SQL:='SELECT * FROM NOMENLIG ' +
     'LEFT JOIN NOMENENT ON GNL_NOMENCLATURE=GNE_NOMENCLATURE ' +
     'LEFT JOIN ARTICLE ON GNL_ARTICLE=GA_ARTICLE ' +
     'WHERE GNL_NOMENCLATURE="' + GNE_NOMENCLATURE.Text + '" ';
TT.Add(SQL) ; TL.Add(TT) ;
LanceDocument('E','ART','NLI',TL,Nil,True,False) ;
TT.Free ; TL.Free ;
{$ELSE}
(*
SQL:='SELECT LIG.*,ENT.GNE_QTEDUDETAIL,ART.GA_QUALIFUNITEVTE AS UNITEOUV,AR1.GA_PVHT,AR1.GA_PAHT,AR1.GA_DPR,'+
		 'AR1.GA_QUALIFUNITEVTE AS UNITELIG FROM NOMENLIG LIG '+
     'LEFT JOIN NOMENENT ENT ON GNE_NOMENCLATURE=GNL_NOMENCLATURE '+
		 'LEFT JOIN ARTICLE ART ON ART.GA_ARTICLE=ENT.GNE_ARTICLE '+
		 'LEFT JOIN ARTICLE AR1 ON AR1.GA_ARTICLE=LIG.GNL_ARTICLE '+
		 'WHERE GNL_NOMENCLATURE="' + trim(GNE_NOMENCLATURE.Text) + '" ORDER BY GNL_NOMENCLATURE';
LanceEtat('E','ART','BDO',True,False,False,Nil,SQL,'',False);
*)
TOBETAT := TOB.Create ('L EDITION',nil,-1);
constitueTOBEtat (TOBEtat);
LanceEtatTob('E','ART','BDD',TOBEtat,True,false,False,nil,'','Détail de l''ouvrage',False);
TOBEtat.free;
{$ENDIF}
{$ENDIF}
end;

procedure TFNomenLig.bNewligneClick(Sender: TObject);
begin
//  LastKey positionné pour eviter le traitement de G_NLIGRowMoved
LastKey := VK_INSERT;
// Modif BTP
if G_NLIG.cells [SGN_Comp,G_NLIG.row] <> '' then
   InsertLigne (G_NLIG.Row, False);
LastKey := 0;
end;

procedure TFNomenLig.bDelLigneClick(Sender: TObject);
{$IFDEF BTP}
var
   Acol,Arow: integer;
   Cancel : boolean;
{$ENDIF}
begin
//  LastKey positionné pour eviter le traitement de G_NLIGRowMoved
LastKey := VK_DELETE;
SupprimeLigne (G_NLIG.Row) ;

{$IFDEF BTP}
Arow := G_NLIG.row;
Acol := SGN_COMP;
Cancel := false;
G_NLIGCellEnter (Sender,ACol,Arow,cancel);
{$ENDIF}

LastKey := 0;
end;

procedure TFNomenLig.bNomencClick(Sender: TObject);
begin
TV_NLIGDblClick(Sender);
end;

//=============================================================================
//  Evenements de la Grid
//=============================================================================
//  a l'entrée dans une ligne
//=============================================================================
procedure TFNomenLig.G_NLIGRowEnter(Sender: TObject; Ou: Integer;
                                    var Cancel: Boolean; Chg: Boolean);
var
   TOBA : TOB;
   ARect : Trect;
   NumGraph : Integer;
begin
//--------------------------------------------------------------------------
//  Verification qu'on n'entre pas sur une ligne vide au dela de la premiere possible
//--------------------------------------------------------------------------
if Ou > TOBLig.detail.count then
    begin
    G_NLIG.Row := TOBLig.detail.count + 1;
    G_NLIG.Col := 1;
    Ou := Math.Min(Ou, G_NLIG.Row);
    G_NLIG.row := Ou;
    //Exit;
    end;
// {$IFDEF BTP}
if (G_NLIG.Cells[SGN_Comp, Ou] = '') and (TOB(G_NLIG.objects[SGN_Numlig,Ou]) = nil) then InsertLigne (Ou, True);
// {$ELSE}
//if G_NLIG.Cells[SGN_Comp, Ou] = '' then InsertLigne (Ou, True);
//{$ENDIF}

//--------------------------------------------------------------------------
//  affichage synchrone du treeview avec la grid
//--------------------------------------------------------------------------
TV_NLIG.HideSelection := False;
TV_NLIG.Selected := TTreeNode(G_NLIG.Objects[SGN_Comp, Ou]);
if (PopTV_NLIG_O.Checked) and (TV_NLIG.Selected <> nil) then TV_NLIG.Selected.Expand(False);
SetImages(TV_NLIG, nil);
//--------------------------------------------------------------------------
//  Mise a jour du pied
//--------------------------------------------------------------------------
AffectePied(Ou);

if GNL_SOUSNOMEN.Text = '' then
    begin
    bNomenc.Enabled := False;
    PopG_NLIG_N.Enabled := False;
    end
    else
    begin
    bNomenc.Enabled := True;
    PopG_NLIG_N.Enabled := True;
    end;
if CasEmploi_Open then
    PopG_NLIG_C.Enabled := True
    else
    PopG_NLIG_C.Enabled := False;

{$IFDEF BTP}
  if TypeSaisie = TTNOUV then
  begin
    // Pour affectation de TMontantLig
    InitTableau (TMontantLig);
    // Stockage du montant ligne
    if G_NLIG.Cells[SGN_Comp,Ou] <> '' then
        begin
        TMontantLig := calculeMontantLig (Ou);
        end;
    AffecteValArticle (ou);
  end;
{$ENDIF}

if (SGN_TypA <> -1) and (Action <> taconsult) then
   BEGIN
   TOBA := TOB(G_NLIG.Objects[SGN_idcomp,Ou]);
   if TOBA <> nil then
      BEGIN
      ARect:=G_NLIG.CellRect(SGN_Typa,G_NLIG.Row) ;
      NumGraph := RecupTypeGraph (TOBA);
      BTypeArticle.ImageIndex := NumGraph;
      BTYpeArticle.Opaque := false;
      With BTypeArticle do
         Begin
         Top := Arect.top - G_NLIG.GridLineWidth ;
         Left := Arect.Left;
         Width := Arect.Right - Arect.Left;
         Height  := Arect.Bottom - Arect.Top - G_NLIG.GridLineWidth;
         end;
      BTypeArticle.Parent := G_NLIG;
      BTypeArticle.Visible := true;
      END;
   END;
end;

//=============================================================================
//  a la sortie d'une ligne
//=============================================================================
procedure TFNomenLig.G_NLIGRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
//  On remonte et on sort d'une ligne vide au dela des lignes renseignées
//  ou on clique ailleurs et on est au dela des lignes renseignées
if BTypeArticle.Visible then BTypeArticle.visible := false;
if ((LastKey = VK_UP) and (G_NLIG.Cells[SGN_Comp,Ou] = '') and (Ou >= TOBLig.detail.Count)) then // or
//   ((G_NLIG.Cells[SGN_Comp,Ou] = '') and (G_NLIG.Cells[SGN_Comp,Ou + 1] = '')) then
    begin
    if G_NLIG.Objects[SGN_NumLig, Ou] = nil then Exit;
    TOB(G_NLIG.Objects[SGN_NumLig, Ou]).Free;
    G_NLIG.Objects[SGN_NumLig, Ou] := nil;
    G_NLIG.Rows[Ou].Clear;
    AfficheTVNLIG;
//    G_NLIG.Row := G_NLIG.Row - 1;
    Exit;
    end;
//
//
// Modif BTP
//
if  (G_NLIG.Cells[SGN_Comp,Ou] = '') and (Ou >= TOBLig.detail.Count) and (G_NLIG.row < Ou) then
begin
    if G_NLIG.Objects[SGN_NumLig, Ou] = nil then Exit;
    TOB(G_NLIG.Objects[SGN_NumLig, Ou]).Free;
    G_NLIG.Objects[SGN_NumLig, Ou] := nil;
    G_NLIG.Rows[Ou].Clear;
    AfficheTVNLIG;
    exit;
end;
if (G_NLIG.Cells[SGN_Comp,Ou] = '') and (Ou >= TOBLig.detail.Count) and (G_NLIG.row >= Ou) then
begin
cancel := true;
exit;
end;
// -- Fin Modif BTP

if Data_Modified then ValideLigne(Ou);
//--------------------------------------------------------------------------
//  affichage synchrone du treeview avec la grid
//--------------------------------------------------------------------------
TV_NLIG.HideSelection := False;
TV_NLIG.Selected := TTreeNode(G_NLIG.Objects[SGN_Comp, Ou]);
if (PopTV_NLIG_O.Checked) and (TV_NLIG.Selected <> nil) then TV_NLIG.Selected.Collapse(False);
end;

//=============================================================================
//  a l'entrée sur une cellule
//=============================================================================
procedure TFNomenLig.G_NLIGCellEnter(Sender: TObject; var ACol,
                                     ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
ZoneSuivanteOuOk(ACol,ARow,Cancel) ;

if Not Cancel then
   begin
{$IFDEF BTP}
    G_NLIG.ElipsisButton := ((G_NLIG.Col = SGN_Comp) and (G_NLIG.Cells[G_NLIG.col,G_NLIG.row] = ''));
{$ELSE}
    G_NLIG.ElipsisButton := (G_NLIG.Col = SGN_Comp);
{$ENDIF}

   Cell_Text := G_NLIG.Cells[G_NLIG.Col, G_NLIG.Row];

{$IFDEF BTP}
   if G_NLIG.col = SGN_PRIX then
      if (valeur(G_NLIG.cells [G_NLIG.col,G_NLIG.row]) = valeur ('0')) then
          G_NLIG.cells [G_NLIG.col,G_NLIG.row] := strs (RecupPrixBase (G_NLIG.row),V_PGI.OkDecP);
{$ENDIF}
   end;
end;

//=============================================================================
//  a la sortie d'une cellule
//=============================================================================
procedure TFNomenLig.G_NLIGCellExit(Sender: TObject; var ACol,
                                    ARow: Integer; var Cancel: Boolean);
var
    tn1 : TTreeNode;
{$IFDEF BTP}
    TLocal: T_Valeurs;
    dValLoc,dcmp : double;
{$ENDIF}
//    libelle : string;
  TOBL : TOB;
begin
  TOBL := TOB(G_NLIG.Objects[SGN_NumLig, Arow]);

  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application

  //  On remonte et on est sur une ligne vide ou on clique ailleurs et on est sur la
  //  derniere ligne
  if ((LastKey = VK_UP) and (G_NLIG.Cells[SGN_Comp,ARow] = '')) or ((G_NLIG.Cells[SGN_Comp,ARow] = '') and (G_NLIG.Cells[SGN_Comp,ARow + 1] = '')) then Exit;

  //  Simule le OnChange inexistant sur la cellule pour gestion de modif
  if ACol = SGN_Comp then G_NLIG.Cells[ACol, ARow] := Uppercase(G_NLIG.Cells[ACol, ARow]);

  if (Acol = SGN_QTESAIS) or (Acol = SGN_PERTE) or (Acol = SGN_RENDEMENT) then
  begin
    if Arow <> IPercent then
    begin
      TraiteQteSais (Acol,Arow);
    end;
  end;
  {$IFDEF BTP}
  if (SGN_PRIX <> -1) and (Acol = SGN_PRIX)  then
  begin
    if Arow <> IPercent then
    begin
      dValLoc := RecupPrixBase (Arow);
      dcmp := valeur(G_NLIG.cells [Acol,Arow]);
      if dcmp = dValLoc then G_NLIG.cells [Acol,Arow] := cell_text;
    end;
  end;
  {$ENDIF}

  if Cell_Text = G_NLIG.Cells[ACol, ARow] then Exit;

  Data_Modified := True;
  //  Controle du contenu de la cellule
  if (ACol = SGN_Comp) and (not TraiterComposant(ARow)) then Exit;
  if (ACol = SGN_Qte)  and (not TraiterQuantite (ARow)) then Exit;

  {$IFDEF BTP}
  if (ACol = SGN_Prix) and (not TraiterPrix (ARow)) then Exit;
  {$ENDIF}

  if (ACol = SGN_Joker) and (not TraiterJoker (ARow)) then Exit;
  //  Formattage de la cellule en cours
  FormateZoneSaisie (ACol,ARow);
  //  Pas de tob associée, on sort.  ( ???????????? )
  if G_NLIG.Objects[SGN_NumLig, ARow] = nil then Exit;

  {$IFDEF BTP}
  if TypeSaisie = TTNOUV then
  begin
    Tlocal := calculeMontantLig (Arow);
    CalculeMontantOuv ();
    AlimenteValeurOuv ();
    TMontantLig := TLocal;
    AffecteValArticle (ARow);
    ValideLigne (Arow);
  end;
  {$ENDIF}

//  on sort de la colonne composant ou de la colonne qte, on met a jour le libelle du treeview
  if (ACol = SGN_Comp) or (ACol = SGN_Qte) or (Acol = SGN_QTESAIS)  then
  begin
    if G_NLIG.Cells[SGN_SSNom, Arow] <> '' then
    begin
      if TOB(G_NLIG.Objects[SGN_NumLig, Arow]).getvalue('QTEDUDETAIL') <> valeur('0') then
      begin
        TOB(G_NLIG.Objects[SGN_NumLig, Arow]).PutValue('LIBELLE',G_NLIG.Cells[SGN_SSNom, Arow] +
                                      ' (' + G_NLIG.Cells[SGN_Qte, Arow] + ' Détail pour '+
                                       Strs0(TOB(G_NLIG.Objects[SGN_NumLig, Arow]).getvalue('QTEDUDETAIL')) + ' ' +
                                       TOB(G_NLIG.Objects[SGN_NumLig, Arow]).getvalue('UNITENOMENC') +')');
      end
      else
          TOB(G_NLIG.Objects[SGN_NumLig, Arow]).PutValue('LIBELLE', G_NLIG.Cells[SGN_SSNom, Arow] +
                                      ' (' + G_NLIG.Cells[SGN_Qte, Arow] + ')');
    end
    else
      TOB(G_NLIG.Objects[SGN_NumLig, ARow]).PutValue('LIBELLE', G_NLIG.Cells[SGN_Comp, ARow] +
                                     ' (' + StrS0(Valeur(G_NLIG.Cells[SGN_Qte, ARow])) + ')');
    if Acol = SGN_QTE then
    begin
      TOBL.SetDOuble('GNL_QTE',valeur(G_NLIG.Cells[Acol, ARow]));
      TOBL.SetDouble('GNL_QTESAIS',CalculeQteSais(TOBL));
    end;
  end;

  tn1 := TTreeNode(G_NLIG.Objects[SGN_Comp, ARow]);
  if tn1 <> nil then
  begin
    tn1.Text := TOB(G_NLIG.Objects[SGN_NumLig, ARow]).GetValue('LIBELLE');
    if G_NLIG.Cells[SGN_SSNom, ARow] = '' then
    begin
      tn1.ImageIndex := 2;
      tn1.SelectedIndex := 2;
    end
    else
    begin
      tn1.ImageIndex := 0;
      tn1.SelectedIndex := 0;
    end;
  end;

end;

//=============================================================================
//  Lancement de la recherche dans la table article
//=============================================================================
procedure TFNomenLig.G_NLIGElipsisClick(Sender: TObject);
Var Coord : TRect;
    StWhere,StFiche : string;
{$IFDEF BTP}
    TheArticle : string;
{$ENDIF}
begin
if (G_NLIG.Col = SGN_Comp) then
    BEGIN
    Coord         := G_NLIG.CellRect (G_NLIG.Col, G_NLIG.Row);
    ART.Parent    := G_NLIG;
    ART.Top       := Coord.Top;
    ART.Left      := Coord.Left;
    ART.Width     := 3;
    ART.Visible   := False;
    ART.DataType  := 'GCARTICLE';
  	ART.Text      := trim(G_NLIG.Cells [G_NLIG.Col,G_NLIG.row]);
		TheArticle    := trim(G_NLIG.Cells [G_NLIG.Col,G_NLIG.row]);
    StFiche       := '';
    if TypeSaisie = TTNOUV then
    begin
    	StWhere := GetTypeArticleBTP;
    end else if TypeSaisie = TTNPARC then
    begin
      StWhere := GetTypeArticlePARC;
      stFiche := 'BTARTPARC_RECH';
    end;
    StWhere   := 'GA_CODEARTICLE=' + TheArticle+';XX_WHERE=AND '+stWhere;
    DispatchRecherche (ART, 1, '',stWhere, stFiche);
    if ART.Text <> '' then G_NLIG.Cells[G_NLIG.Col,G_NLIG.Row]:= copy(ART.Text, 0, Length(ART.Text) - 1);
    FormateZoneSaisie (G_NLIG.Col, G_NLIG.Row);
    END ;
end;

//=============================================================================
//  Mousedown dans la grid pour remettre lastkey à 0
//=============================================================================
procedure TFNomenLig.G_NLIGMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
    R_temp1 : TRect;
    ACol, ARow : integer;
    btemp : boolean;
begin
LastKey := 0;
R_temp1 := G_NLIG.CellRect(G_NLIG.Col, G_NLIG.Row);
btemp := False;
ACol := G_NLIG.Col;
ARow := G_NLIG.Row;
if (X < R_temp1.Left) or (X > R_temp1.Right) then
    G_NLIGCellExit(Sender, ACol, ARow, btemp);
if (Y < R_temp1.Top) or (Y > R_temp1.Bottom) then
    G_NLIGRowExit(Sender, ARow, btemp, False);
end;

//=============================================================================
//  double clic sur une ligne on ouvre la fiche article en consultation
//=============================================================================
procedure TFNomenLig.G_NLIGDblClick(Sender: TObject);
begin
if (G_NLIG.Cells[SGN_Comp, G_NLIG.Row] = '') then Exit;

{$IFDEF BTP}
V_PGI.DispatchTT (7,taconsult,G_NLIG.Cells[SGN_IdComp, G_NLIG.Row],'','ACTION=CONSULTATION;MONOFICHE');
{$ELSE}
{$IFDEF CHR}
AGLLanceFiche('H', 'HRPRESTATIONS', '', G_NLIG.Cells[SGN_IdComp, G_NLIG.Row], 'ACTION=CONSULTATION');
{$ELSE}
AGLLanceFiche('GC','GCARTICLE','',G_NLIG.Cells[SGN_IdComp, G_NLIG.Row],'ACTION=CONSULTATION');
{$ENDIF}
{$ENDIF}

end;
//=============================================================================
//  Deplacement d'une ligne dans la grid
//=============================================================================
procedure TFNomenLig.G_NLIGRowMoved(Sender: TObject; FromIndex, ToIndex: Integer);
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
        G_NLIG.Objects[SGN_IDComp, i_ind1 + 1] := G_NLIG.Objects[SGN_IdComp, ToIndex];
        G_NLIG.Objects[SGN_IdComp, ToIndex] := nil;
        G_NLIG.Objects[SGN_Comp, i_ind1 + 1] := G_NLIG.Objects[SGN_Comp, ToIndex];
        G_NLIG.Objects[SGN_Comp, ToIndex] := nil;
{$IFDEF BTP}
        G_NLIG.Objects[SGN_pRIX, i_ind1 + 1] := G_NLIG.Objects[SGN_pRIX, ToIndex];
        G_NLIG.Objects[SGN_Prix, ToIndex] := nil;
{$ENDIF}
        for i_ind2 := 0 to G_NLIG.ColCount - 1 do
            begin
            G_NLIG.Cells[i_ind2, i_ind1 + 1] := G_NLIG.Cells[i_ind2, ToIndex];
            G_NLIG.Cells[i_ind2, ToIndex] := '';
            end;
        ToIndex := i_ind1 + 1;
        end;
    end;
//  on renumerote toutes les lignes
Renumerote(1);
//  on retrie les tobs associées
if FromIndex < ToIndex then
    for i_ind1 := FromIndex to ToIndex - 1 do
        TOBLig.Detail.Exchange(i_ind1 - 1, i_ind1)
else
    for i_ind1 := FromIndex - 1 downto ToIndex do
        TOBLig.Detail.Exchange(i_ind1 - 1, i_ind1);
//  on rafraichit le treeview
AfficheTVNLIG;
//
end;

procedure TFNomenLig.PopG_NLIG_AClick(Sender: TObject);
begin
G_NLIGDblClick(Sender);
end;

procedure TFNomenLig.PopG_NLIG_CClick(Sender: TObject);
var
    st_Article : string;
begin
st_Article := G_NLIG.Cells[SGN_Comp, G_NLIG.Row];
Entree_CasEmploi(['NOMENCLATURE=NON', st_Article], 2);
end;

procedure TFNomenLig.PopG_NLIG_IClick(Sender: TObject);
begin
bNewLigneClick(Sender);
end;

procedure TFNomenLig.PopG_NLIG_SClick(Sender: TObject);
begin
bDelLigneClick(Sender);
end;

procedure TFNomenLig.PopG_NLIG_NClick(Sender: TObject);
begin
TV_NLIGDblClick(Sender);
end;

procedure TFNomenLig.PopG_NLIG_DClick(Sender: TObject);
begin
PopG_NLIG_D.Checked := not (PopG_NLIG_D.Checked);
PANDIM.Visible := PopG_NLIG_D.Checked;
end;

//=============================================================================
//  Evenements du treeview
//=============================================================================
procedure TFNomenLig.TV_NLIGMouseUp(Sender: TObject; Button: TMouseButton;
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

{$IFDEF BTP}
              i_ind1 := ChercheLigneBtp(TV_NLIG.Selected);
{$ELSE}
//              i_ind1 := ChercheLigne(tn1.Text); // DBR Fiche 10995
              i_ind1 := ChercheLigne(tn1); // DBR Fiche 10995
{$ENDIF}

              if i_ind1 <> 0 then
                  begin
                  b_temp := False;
                  if G_NLIG.Row <> i_ind1 then G_NLIGRowExit(Sender, G_NLIG.Row, b_temp, b_temp);
                  G_NLIG.Col := SGN_Comp;
                  G_NLIG.Row := i_ind1;
                  G_NLIGRowEnter(Sender, i_ind1, b_temp, b_temp);
                  G_NLIG.SetFocus;
                  TV_NLIG.Selected := TTreeNode(G_NLIG.Objects[SGN_Comp, i_ind1]);
                  Exit;
                  end;
              end;
    mbRight : begin
              end;
    end;
end;

procedure TFNomenLig.TV_NLIGCollapsing(Sender: TObject; Node: TTreeNode;
                                       var AllowCollapse: Boolean);
begin
Node.ImageIndex := 0;
end;

procedure TFNomenLig.TV_NLIGExpanding(Sender: TObject; Node: TTreeNode;
                                      var AllowExpansion: Boolean);
begin
Node.ImageIndex := 1;
end;

//=============================================================================
//  Double clic, on ouvre la decomposition dans une nouvelle fenetre
//=============================================================================
procedure TFNomenLig.TV_NLIGDblClick(Sender: TObject);
var
    st_Article, st_Nomen : string;
    Active_Ligne : integer;
    TOBEnt, TOBTemp : TOB;
    i_ind1 : integer;
{$IFDEF BTP}
    st_qtedudetail : string;
{$ENDIF}
begin
//  sauvegarde la ligne en cours
Active_Ligne := G_NLIG.Row;
//  on verifie que la ligne n'est pas vide et que c'est une nomenclature

{$IFDEF BTP}
i_ind1 := ChercheLigneBtp(TV_NLIG.Selected);
{$ELSE}
//i_ind1 := ChercheLigne(TV_NLIG.Selected.Text); // DBR Fiche 10995
i_ind1 := ChercheLigne(TV_NLIG.Selected); // DBR Fiche 10995
{$ENDIF}

if G_NLIG.Cells[SGN_Comp, i_ind1] = '' then Exit;
if TOB(G_NLIG.Objects[SGN_NumLig, i_ind1]).GetValue('GNL_SOUSNOMEN') = '' then Exit;
if TOB(G_NLIG.Objects[SGN_idcomp, i_ind1]).GetValue('GA_TYPEARTICLE') = 'ARP' then Exit;
//  on cree et on ouvre la nouvelle fenetre
//FNomLig := TFNomenLig.Create(Application);
TOBEnt := TOB.Create('NOMENENT', nil, -1);
TOBEnt.PutValue('GNE_ARTICLE', TOB(G_NLIG.Objects[SGN_NumLig, i_ind1]).GetValue('GNL_ARTICLE'));
TOBEnt.PutValue('GNE_NOMENCLATURE', TOB(G_NLIG.Objects[SGN_NumLig, i_ind1]).GetValue('GNL_SOUSNOMEN'));
TOBEnt.LoadDB();

st_Article := TOBEnt.GetValue('GNE_ARTICLE');
st_Nomen := TOBEnt.GetValue('GNE_NOMENCLATURE');

{$IFDEF BTP}
st_qtedudetail := TOBEnt.GetValue('GNE_QTEDUDETAIL');
Entree_NomenLig(['CASEMPLOIS=OUI', st_Article, st_Nomen,st_qtedudetail,taconsult], 5);
{$ELSE}
Entree_NomenLig(['CASEMPLOIS=OUI', st_Article, st_Nomen], 3);
{$ENDIF}

//  on fait le menage, on reinitialise les données de la nomenclature appellée
TOBEnt.Free;
TOBTemp := TOB(G_NLIG.Objects[SGN_NumLig, i_ind1]);
ChargeNomen(TOBTemp.GetValue('GNL_SOUSNOMEN'), True, TOBTemp);
AffecteLibelle(TOBTemp, 'GNL_CODEARTICLE');
//  on reactualise les affichages
AfficheTVNLIG;
G_NLIG.Row := Active_Ligne;
TV_NLIG.Selected := TTreeNode(G_NLIG.Objects[SGN_Comp, G_NLIG.Row]);
end;

//=============================================================================
//  Clic Droit Option Tout Ouvrir
//=============================================================================
procedure TFNomenLig.PopTVNLIG_TClick(Sender: TObject);
begin
TV_NLIG.FullExpand;
end;

//=============================================================================
//  Clic Droit Option Tout Fermer
//=============================================================================
procedure TFNomenLig.PopTVNLIG_FClick(Sender: TObject);
begin
TV_NLIG.FullCollapse;
TV_NLIG.TopItem.Expand(False);
end;

//=============================================================================
//  Clic Droit Option Tout Fermer
//=============================================================================
procedure TFNomenLig.PopTV_NLIG_OClick(Sender: TObject);
begin
PopTV_NLIG_O.Checked := not (PopTV_NLIG_O.Checked);
end;

/////////////////////////////////////////////////////////////////////////////
procedure InitNomenLig();
begin
RegisterAglProc( 'Entree_NomenLig', False , 3, Entree_NomenLig);
{$IFDEF BTP}
RegisterAglProc( 'Entree_NomenLig_Btp', False , 5, Entree_NomenLig);
RegisterAglProc( 'EntreeOuvrageDetail', False , 4, EntreeOuvrageDetail);
RegisterAglProc( 'EntreeParcDetail', False , 3, EntreeParcDetail);

{$ENDIF}
end;


procedure TFNomenLig.PostDrawCell(ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState);
{$IFDEF BTP}
var
    R:Trect;
    TheText : String;
    ValLoc : double;
    TOBA : TOB;
    NumGraph : Integer;
{$ENDIF}
begin
{$IFDEF BTP}
if (OkAffichage) and
    (not (gdfocused in AState)) and
    (arow >= G_NLIG.FixedRows) then
    begin
    TheText := G_NLIG.Cells[Acol,ARow];
    R := G_NLIG.CellRect (Acol,Arow);
    if (Acol=SGN_Prix) then
        begin

        with Canvas do
        BEGIN
        if (TOB(G_NLIG.Objects[SGN_Numlig, ARow]) <> nil ) and
            (TOB(G_NLIG.Objects[SGN_idcomp, ARow]) <> nil) then
            begin
            // Valeur dans le grid = 0
            if valeur(G_NLIG.cells[SGN_Prix,Arow]) = valeur('0') then
               begin
               if valeur(TOB(G_NLIG.Objects[SGN_Numlig, ARow]).GetValue('GNL_PRIXFIXE')) = valeur('0') then
                  begin
                //
                // Affichage du Pv
                //
                  TheText:= Strs(TOB(G_NLIG.Objects[SGN_Numlig, ARow]).GetValue('PVHT'),V_PGI.OkDecP)
                  end;
               end else
               begin
               if Arow = IPercent then
                  begin
                  Valloc := TOB(G_NLIG.Objects[SGN_Numlig, Arow]).GetValue('PVHT');
                  TheText:= Strs(ValLoc,V_PGI.OkDecP);
                  font.style := font.style + [fsBold];
                  font.color := clBlue;
                  end else
                  begin
                  font.style := font.style + [fsBold];
                  font.color := clMaroon;
                  end;
               end;
            end;
            FillRect (R);
            // cadrage à droite
            Textout(R.left + R.right-R.left-TextWidth(TheText) - 2 , R.Top+2 , TheText);
         end;
      end else if (SGN_TYPA <> 0 ) and (ACOl = SGN_TypA) then
      begin
      Canvas.FillRect (R);
      if (TOB(G_NLIG.Objects[SGN_Numlig, ARow]) <> nil ) and
         (TOB(G_NLIG.Objects[SGN_idcomp, ARow]) <> nil) then
         begin
         TOBA := TOB(G_NLIG.Objects[SGN_idcomp, ARow]);
         NumGraph := RecupTypeGraph (TOBA);
         if NumGraph >= 0 then
            begin
            ImTypeArticle.DrawingStyle := dsTransparent;
            ImTypeArticle.Draw (CanVas,R.left,R.top,NumGraph);
            end;
         end;
      end;
    end;
{$ENDIF}
end;

{$IFDEF BTP}
Procedure  TFNomenLig.chargePrixOuvrage (TOBLIG : TOB; CodeNomenc: String; ARow : Integer);
var
   valeurs : T_valeurs;
begin
    InitTableau (valeurs);
    CalculeValoOuvrage(TOBLig,valeurs);
    TOBLig.Putvalue ('DPA',valeurs[0]);
    TOBLig.Putvalue ('DPR',valeurs[1]);
    TOBLig.Putvalue ('PVHT',valeurs[2]);
    TOBLig.Putvalue ('PVTTC',valeurs[3]);
    TOBLIG.Putvalue ('PMAP',valeurs[6]);
    TOBLIG.Putvalue ('PMRP',valeurs[7]);
    // New
    TOBLIG.Putvalue ('TPSUNITAIRE',valeurs[9]);
    TOBLIG.Putvalue ('GA_HEURE',valeurs[9]);
//
    if Tob(G_NLIG.Objects[SGN_IdComp, Arow]).getvalue('GA_TYPEARTICLE') <> 'ARP' then
    begin
      TOBLIG.Putvalue ('GA_PAHT',valeurs[0]);
      TOBLIG.Putvalue ('GA_PVHT',valeurs[2]);
    end else
    begin
      TOBLIG.Putvalue ('GA_PAHT',TOBLIG.detail[0].GetValue('DPA'));
      TOBLIG.Putvalue ('GA_PVHT',TOBLIG.detail[0].GetValue('PVHT'));
    end;
  TOBLIG.putValue('GA_PAHT',TOBLIG.getValue('DPA'));
  TOBLIG.putValue('GA_PVHT',TOBLIG.getValue('PVHT'));
  TOBLIG.putvalue ('GF_PRIXUNITAIRE',TOBLIG.GetValue('PVHT'));
end;
{$ENDIF}

{$IFDEF BTP}
function TFNomenlig.calculeMontantLig (Ou : Integer) : T_Valeurs;
var
   Dqte,DPrixQte: double;
   TResult : T_Valeurs;
begin
     InitTableau (Tresult);
     Dqte := Valeur (G_NLIG.cells[SGN_Qte,Ou]);
     if TOB(G_NLIG.Objects[SGN_IdComp, ou]) <> nil then
     begin
        DPrixQte := TOB(G_NLIG.objects[SGN_IdComp,ou]).getValue ('GA_PRIXPOURQTE');
        if DPrixQte = valeur('0') then DPrixQte := valeur('1');
        if TOB(G_NLIG.Objects[SGN_Idcomp, Ou]).GetValue('GA_TYPEARTICLE') = 'POU' then
        //
        begin
           Tresult := TvaleurOuv;
           TPercent := TValeurPou;
           TOB(G_NLIG.Objects[SGN_Numlig, Ou]).PutValue('DPA',TPercent[0]);
           TOB(G_NLIG.Objects[SGN_Numlig, Ou]).PutValue('DPR',TPercent[1]);
           TOB(G_NLIG.Objects[SGN_Numlig, Ou]).PutValue('PVHT',TPercent[2]);
           Tob(G_NLIG.objects[SGN_Numlig,Ou]).PutValue('GNL_PRIXFIXE',TPercent[2]);
           TOB(G_NLIG.Objects[SGN_Numlig, Ou]).PutValue('PVTTC',TPercent[3]);
           TOB(G_NLIG.Objects[SGN_Numlig, Ou]).PutValue('PMAP',TPercent[6]);
           TOB(G_NLIG.Objects[SGN_Numlig, Ou]).PutValue('PMRP',TPercent[7]);
           G_NLIG.Cells [SGN_Prix,ou] := strs (TPercent[2],V_PGI.OkDecP );
           G_NLIG.InvalidateCell (SGN_Prix,Ipercent);
        end;
        //
           // Prix de vente HT
        if valeur(G_NLIG.cells[SGN_Prix,Ou]) = 0 then
           begin
             if (TOB(G_NLIG.Objects[SGN_NumLig, Ou])<> nil ) then
             begin
              TResult[2] := dQte * (TOB(G_NLIG.Objects[SGN_Numlig, ou]).GetValue('PVHT')/ DprixQte);
              TOB(G_NLIG.Objects[SGN_Numlig, ou]).PutValue('GCA_PRIXVENTE',Arrondi(dQte * (TOB(G_NLIG.Objects[SGN_Numlig, ou]).GetValue('PVHT')),V_PGI.okDecV));
             end else
             begin
              TResult[2] := DQte * valeur(G_NLIG.cells[SGN_Prix,Ou])/DPrixQte;
              TOB(G_NLIG.Objects[SGN_Numlig, ou]).PutValue('GCA_PRIXVENTE',Arrondi(DQte * valeur(G_NLIG.cells[SGN_Prix,Ou])/DPrixQte,V_PGI.okDecV));
             end;
           end
           else
           begin
              TResult[2] := DQte * valeur(G_NLIG.cells[SGN_Prix,Ou])/DPrixQte;
              TOB(G_NLIG.Objects[SGN_Numlig, ou]).PutValue('GCA_PRIXVENTE',Arrondi(DQte * valeur(G_NLIG.cells[SGN_Prix,Ou])/DPrixQte,V_PGI.okdecV));
           end;

        Tresult[2] := arrondi (Tresult[2],V_PGI.OkDecP );
        Tresult[0] := dQte * TOB(G_NLIG.objects[SGN_Numlig,ou]).getValue ('DPA')/DprixQte;
        Tresult[1] := Dqte * TOB(G_NLIG.objects[SGN_Numlig,ou]).getValue ('DPR')/Dprixqte;
        Tresult[3] := Dqte * TOB(G_NLIG.objects[SGN_Numlig,ou]).getValue ('PVTTC')/DprixQte;
        Tresult[6] := Dqte * TOB(G_NLIG.objects[SGN_Numlig,ou]).getValue ('PMAP')/DprixQte;
        Tresult[7] := Dqte * TOB(G_NLIG.objects[SGN_Numlig,ou]).getValue ('PMRP')/DprixQte;
        if (Pos(TOB(G_NLIG.objects[SGN_IdComp,ou]).GetValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne)>0) or
           (TOB(G_NLIG.objects[SGN_IdComp,ou]).getvalue('GA_TYPEARTICLE')='OUV') or
           (TOB(G_NLIG.objects[SGN_IdComp,ou]).getvalue('GA_TYPEARTICLE')='ARP') then
        begin
        	Tresult[9] := Dqte * TOB(G_NLIG.objects[SGN_Numlig,ou]).getValue ('TPSUNITAIRE');
        end;
     end;
     Result := TResult;
end;
{$ENDIF}

{$IFDEF BTP}
Procedure TFNomenlig.AffecteValArticle  (Arow : integer);
var TOBL : TOB;
begin
   if TOB(G_NLIG.Objects[SGN_Numlig, Arow]) <> nil then
      begin
      TOBL := TOB(G_NLIG.Objects[SGN_Numlig, Arow]);
      NEPA.value := TOB(G_NLIG.Objects[SGN_Numlig, Arow]).getValue('DPA');
      NEPR.value := TOB(G_NLIG.Objects[SGN_Numlig, Arow]).getValue('DPR');
      NEPVTTC.value := TOB(G_NLIG.Objects[SGN_Numlig, Arow]).getValue('PVTTC');
      NEPVHT.value := valeur(strs(TOB(G_NLIG.Objects[SGN_Numlig, Arow]).GetValue('PVHT'),V_PGI.OkDecP));
      NEPVFORFAIT.value := Tob(G_NLIG.objects[SGN_Numlig,Arow]).getValue('GNL_PRIXFIXE');
      // NEWONE
      TGA_FOURNPRINC.caption := TOB(G_NLIG.Objects[SGN_Numlig, Arow]).getValue('TGA_FOURNPRINC');
      TGA_PRESTATION.caption := TOB(G_NLIG.Objects[SGN_Numlig, Arow]).getValue('TGA_PRESTATION');
      NETPS.Value := TOB(G_NLIG.Objects[SGN_Numlig, Arow]).getValue('TPSUNITAIRE') *
                     TOB(G_NLIG.Objects[SGN_Numlig, Arow]).getValue('GNL_QTE');
      if TOBL.detail.count <> 0 then
      begin
        calculecoefs ();
      end else
      begin
        NECOEFPAPR.value := TOBL.getValue('COEFFG');
        if NEPR.value <> valeur ('0') then
        begin
          if NEPVFORFAIT.value <> valeur ('0') then
             NECOEFPRPV.value := NEPVFORFAIT.value/NEPR.value
          else
             NECOEFPRPV.value := TOBL.GetValue('COEFMARG');
        end;
      end;
      end else
      begin
      NEPA.value := 0;
      NEPR.value := 0;
      NEPVTTC.value := 0;
      NEPVHT.value := 0;
      NEPVFORFAIT.value := 0;
      NETPS.value := 0;
      TGA_FOURNPRINC.caption := '';
      TGA_PRESTATION.caption := '';
      calculecoefs ();
      end;
   If Arow = IPercent then
   begin
        NEPVFORFAIT.Enabled := false;
        NEPVFORFAIT.Color  := clInactiveCaptionText;
   end
   else
   begin
        NEPVFORFAIT.Enabled := true;
        NEPVFORFAIT.Color  := clWindow;
   end;
end;
{$ENDIF}

{$IFDEF BTP}
procedure TFNomenLig.calculecoefs ();
begin
   if NEPA.value <> valeur ('0') then
      NECOEFPAPR.value := NEPR.value/NEPA.value
   else
      NECOEFPAPR.value := valeur ('0');
   if NEPR.value <> valeur ('0') then
   begin
      if NEPVFORFAIT.value <> valeur ('0') then
         NECOEFPRPV.value := NEPVFORFAIT.value/NEPR.value
      else
         NECOEFPRPV.value := NEPVHT.value/NEPR.value;
   end
   else NECOEFPRPV.value := valeur ('0');
end;
{$ENDIF}


procedure TFNomenLig.ExitPVForfait(Sender: TObject);
{$IFDEF BTP}
var
   TOBLoc : TOB;
   TLocal : T_Valeurs;
{$ENDIF}
begin
{$IFDEF BTP}
TOBLOC := Tob(G_NLIG.objects[SGN_Numlig,G_NLIG.Row]);
TOBLOC.putvalue ('GNL_PRIXFIXE', NEPVFORFAIT.value);
G_NLIG.cells[SGN_Prix,G_NLIG.row] := strs (NEPVFORFAIT.value,V_PGI.OkDecP);
G_NLIG.InvalidateCell (SGN_PRIX,G_NLIG.row);
calculecoefs ();
Tlocal := calculeMontantLig (G_NLIG.row);
CalculeMontantOuv ();
AlimenteValeurOuv ();
TMontantLig := TLocal;
AffecteValArticle (G_NLIG.Row);
{$ENDIF}
end;

{$IFDEF BTP}
Procedure TFNomenLig.CalculeMontantOuv ();
var
   indice : Integer;
   TLocal : T_valeurs;
   ArticleOk : string;
begin
   IPercent := 0;
   InitTableau (TValeurOuv);
   InitTableau (TValeurPou);
   InitTableau (TPercent);
   for indice := 1 to G_NLIG.RowCount - 1 do
       begin
       if TOB(G_NLIG.Cells[SGN_IdComp, Indice]) = nil then break;
       if Tob(G_NLIG.Objects[SGN_IdComp, Indice]).getvalue('GA_TYPEARTICLE') = 'POU' then
          begin
          ArticleOk := Tob(G_NLIG.Objects[SGN_IdComp, Indice]).getvalue('GA_LIBCOMPL');
          break;
          end;
       end;

   for indice := 1 to G_NLIG.ColCount - 1 do
   begin
     if Tob (G_NLIG.objects[SGN_Idcomp,Indice]) <> nil then
     begin
        if Tob(G_NLIG.objects[SGN_Idcomp,Indice]).Getvalue('GA_TYPEARTICLE') <> 'POU' then
        begin
           Tlocal := calculeMontantLig (Indice);
           TValeurOuv := CalculSurTableau ('+',TValeurOuv,Tlocal);
           if  ArticleOKInPOUR (Tob(G_NLIG.objects[SGN_Idcomp,Indice]).Getvalue('GA_TYPEARTICLE'),ArticleOk) then
               TValeurPou := CalculSurTableau ('+',TValeurPou,Tlocal);
        end
        else IPercent := Indice;
     end;
   end;
   if IPercent <> 0 then Tpercent := calculeMontantLig (Ipercent);
end;
{$ENDIF}

procedure TFNomenLig.AlimenteValeurOuv ();
var TValeur : T_Valeurs;
begin
  //
  TValeur := CalculSurTableau ('+',TValeurOuv,TPercent);
  //
  MONTANTPA.value   :=  TValeur [0];
  MONTANTPR.value   :=  TValeur [1];
  MONTANTHT.value   :=  TValeur [2];
  MONTANTTTC.value  :=  TValeur [3];
  //
  if MONTANTPA.value <>  Valeur ('0') then
    COEFFG.value    :=  TValeur[1] / TValeur[0]
  else
    COEFFG.value    :=   Valeur ('0');
  //
  if MONTANTPR.value <>  Valeur ('0') then
    COEFMARG.value  :=  TValeur[2] / TValeur [1]
  else
    COEFMARG.value  :=   Valeur ('0');

  MONTANTFG.value   :=  TValeur[1] - TVAleur[0];
  MONTANTMARG.value :=  TValeur[2] - TValeur[1];
  TOTALTPS.value    :=  TValeur[9];
  //
  if GNE_QTEDUDETAIL.value <> 0 then
  begin
    MONTANTPAU.value  :=  TValeur [0]/GNE_QTEDUDETAIL.value;
    MONTANTPRU.value  :=  TValeur [1]/GNE_QTEDUDETAIL.value;
    MONTANTHTU.value  :=  TValeur [2]/GNE_QTEDUDETAIL.value;
    MONTANTTTCU.value :=  TValeur [3]/GNE_QTEDUDETAIL.value;
    MONTANTFGU.value  := (TValeur[1] - TVAleur[0])/GNE_QTEDUDETAIL.value;
    MONTANTMARGU.value:= (TValeur[2] - TValeur[1])/GNE_QTEDUDETAIL.value;
    TPSUNITAIRE.value := (TValeur[9])/GNE_QTEDUDETAIL.value;
  end;
  //
  POuvrage.repaint;
  //
  Pvalo.refresh;
  //
  if GNE_QTEDUDETAIL.Value  = 1 then
  begin
    PValo.Visible := false;
  end
  else
    PValo.Visible := true;

  PValoUnitaire.refresh;
//
end;

{$IFDEF BTP}
function TFNomenLig.RecupPrixBase (Arow : integer) : double;
var //Dqte : double;
    Tresult : T_Valeurs;
begin
result := valeur ('0');
//Dqte := Valeur (G_NLIG.cells[SGN_Qte,Arow]);
if TOB(G_NLIG.Objects[SGN_IdComp,Arow]) = nil then exit;
if TOB(G_NLIG.Objects[SGN_Idcomp, Arow]).GetValue('GA_TYPEARTICLE') <> 'POU' then
   begin
     if not (TOB (G_NLIG.objects[SGN_Numlig,Arow]) = nil) then
        result := valeur(strs(Tob(G_NLIG.objects[SGN_Numlig,Arow]).getValue('GNL_PRIXFIXE'),V_PGI.OkDecP));
     if (result = 0) and (TOB(G_NLIG.Objects[SGN_NumLig, Arow])<> nil ) then
        result := TOB(G_NLIG.Objects[SGN_Numlig, ARow]).GetValue('PVHT');
   end else
   begin
   Tresult := TvaleurOuv;
   result := Tresult [2];
   end;
end;
{$ENDIF}

// Modif BTP
procedure TFNomenLig.PopG_NLIG_OClick(Sender: TObject);
begin
{$IFDEF BTP}
EntreeAnalyseOuvrageBib (TobLig,GNE_NOMENCLATURE.Text);
{$ENDIF}
end;

procedure TFNomenLig.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFNomenLig.BTypeArticleClick(Sender: TObject);
var TOBA : TOB;
begin
TOBA := TOB(G_NLIG.Objects[SGN_idcomp,G_NLIG.Row]);
if TOBA <> nil then
   BEGIN
   if TOBA.GetValue('GA_TYPEARTICLE') = 'OUV' then TV_NLIGDblClick (Sender);
   END;
end;


Function TFNomenLig.ZoneAccessible ( ACol,ARow : Longint) : boolean ;
var TOBL : TOB;
BEGIN
  TOBL := TOB(G_NLIG.Objects[SGN_NumLig, ARow]);
  //
  result := true;
  if G_NLIG.ColLengths[ACol]<0 then
  BEGIN
    result:=false ;
    Exit ;
  END ;
  if (SGA_UVTE<> -1) and (G_NLIG.col = SGA_UVTE) then
  begin
    result:=false ;
    Exit ;
  end;
  if (SGN_PRIX <> -1) and (Acol = SGN_PRIX)  then
  begin
    if Arow = IPercent then
    begin
      result := false;
      Exit ;
    end;
  end;
  if ((Acol = SGN_QTESAIS) or (Acol = SGN_PERTE) or (Acol = SGN_RENDEMENT) or (Acol = SGN_UNITESAIS) ) and ((not GetParamSocSecur('SO_BTGESTQTEAVANC',false)))  then
  begin
    result := false;
    exit
  end;
  if (Acol = SGN_UNITESAIS) then
  begin
    result := false;
    exit;
  end;
  if (Acol = SGN_PERTE) and (POS(TOBL.GetString('BNP_TYPERESSOURCE'),'SAL;INT;ST') >0) then
  begin
    result := false;
    exit
  end;
  if (Acol = SGN_RENDEMENT) and (POS(TOBL.GetString('BNP_TYPERESSOURCE'),'SAL;INT;ST') <=0) then
  begin
    result := false;
    exit
  end;
END ;

procedure TFNomenLig.ZoneSuivanteOuOk ( Var ACol,ARow : Longint ; Var Cancel : boolean ) ;
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
      if ChgLig then BEGIN ACol:=G_NLIG.FixedCols-1 ; ChgLig:=False ; END ;
      if ACol<G_NLIG.ColCount-1 then Inc(ACol) else BEGIN Inc(ARow) ; ACol:=G_NLIG.FixedCols ; END ;
      END else
      BEGIN
      if ((ACol=G_NLIG.FixedCols) and (ARow=1)) then
         BEGIN
         if ChgSens then Break else BEGIN Sens:=1 ; Continue ; END ;
         END ;
      if ChgLig then BEGIN ACol:=G_NLIG.ColCount ; ChgLig:=False ; END ;
      if ACol>G_NLIG.FixedCols then Dec(ACol) else BEGIN Dec(ARow) ; ACol:=G_NLIG.ColCount-1 ; END ;
      END ;
   END ;
G_NLIG.SynEnabled:=OldEna ;
END ;

{$IFDEF BTP}
procedure TFNomenLig.definiBarTree;
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

procedure TFNomenLig.ToolBarTreeClose (Sender : Tobject);
begin
BArborescence.Down := false;
end;

procedure TFNomenLig.AssigneEvenementBtn;
begin
BArborescence.OnClick := BArborescenceClick;
BInfosLigne.OnClick := BInfosLigneClick;
end;

procedure TFNomenLig.BArborescenceClick (Sender : TObject);
begin
ToolBarTree.visible := BArborescence.Down;
end;

procedure TFNomenLig.BInfosLigneClick (Sender : TObject);
begin
ToolBarDesc.visible := BInfosLigne.Down
end;

{$ENDIF}

procedure TFNomenLig.ToolBarDescClose(Sender: TObject);
begin
Binfosligne.Down := false;
end;

procedure TFNomenLig.GNE_QTEDUDETAILExit(Sender: TObject);
begin
{$IFNDEF PGIMAJVER}
	AlimenteValeurOuv ();
{$ENDIF}
end;

function TFNomenLig.GetDomaineOuvrage(CodeOuvrage: string): string;
var QQ : TQuery;
begin
  result := '';
  QQ := OPenSql('SELECT GNE_DOMAINE FROM NOMENENT WHERE GNE_NOMENCLATURE="'+CodeOuvrage+'"',true,-1, '', True);
  if not QQ.eof then
  begin
    Result := QQ.findField('GNE_DOMAINE').AsString;
  end;
  ferme(QQ);
end;

{$IFDEF BTP}
procedure TFNomenLig.CalculeValoOuvrage(TOBLig: TOB; var valeurs: T_valeurs);
var Indice,Ipercent : integer;
    Valloc,TSumPourcent,Tlocal : T_valeurs;
    TOBIndice,TOBPP : TOB;
    ArticleOk : string;
begin
  IPercent := -1;
  InitTableau (valeurs);
  InitTableau (TSumPourcent);
  //
  TOBPP := TOBLig.FindFirst (['GA_TYPEARTICLE'],['POU'],false);
  if TOBPP <> nil then ArticleOk := TOBPP.getValue('GA_LIBCOMPL');
  //
  if TOBLIg.fieldExists('TPSUNITAIRE') then TOBLIG.putValue('TPSUNITAIRE',0);
  //
  for indice:= 0 to TOBLig.Detail.Count -1 do
  begin
    TOBIndice := TOBLig.detail[indice];
    if TobIndice.GetValue('GNL_SOUSNOMEN') = '' then
    begin
      renseigneValoArt (TOBIndice,nil,nil);
      if TOBLIg.fieldExists('TPSUNITAIRE') then
      begin
        TOBLIG.putValue('TPSUNITAIRE',TOBLig.getValue('TPSUNITAIRE')+Arrondi(TOBIndice.getValue('TPSUNITAIRE')*TOBIndice.getValue('GNL_QTE'),3));
      end;
      TOBIndice.putValue('QTEDUDETAIL',TOBLig.getValue('QTEDUDETAIL'));
      CumuleElementTable (TOBIndice,ExAucun,valeurs);
    end else
    if TobIndice.GetValue('GNL_SOUSNOMEN') <> '' then
    begin
      initTableau (Tlocal);
//      CalculeValoOuvrage (TobLig,Tlocal); arggggglllll
      CalculeValoOuvrage (TobIndice,Tlocal);
      TOBIndice.putvalue ('DPA',Tlocal[0]);
      TOBIndice.putvalue ('DPR',Tlocal[1]);
      TOBIndice.putvalue ('PVHT',Tlocal[2]);
      TOBIndice.putvalue ('PVTTC',Tlocal[3]);
      TOBIndice.putvalue ('PMAP',Tlocal[6]);
      TOBIndice.putvalue ('PMRP',Tlocal[7]);
      if TOBLIg.fieldExists('TPSUNITAIRE') then
      begin
        TOBLIG.putValue('TPSUNITAIRE',TOBLig.getValue('TPSUNITAIRE')+Arrondi(TOBIndice.getValue('TPSUNITAIRE')*TOBIndice.getValue('GNL_QTE'),3));
      end;
      //TOBIndice.putvalue ('TPSUNITAIRE',Tlocal[9]);
      //
      TOBIndice.putValue('GA_PAHT',TOBIndice.getValue('DPA'));
      TOBIndice.putValue('GA_PVHT',TOBIndice.getValue('PVHT'));
      TOBIndice.putvalue ('GF_PRIXUNITAIRE',TOBIndice.GetValue('PVHT'));
      //
      TOBIndice.putValue('QTEDUDETAIL',TOBLig.getValue('QTEDUDETAIL'));
      CumuleElementTable (TOBIndice,ExAucun,valeurs);
    end;
    if TobIndice.GetValue('GNL_PRIXFIXE') <> 0 then
    TOBIndice.putvalue ('PVHT',TOBIndice.getValue('GNL_PRIXFIXE'));
    // On est sur une ligne
    if TOBIndice.GetValue ('GA_TYPEARTICLE') = 'POU' then
    begin
      Ipercent := Indice;
      continue;
    end;
    (*
    TOBIndice.putvalue ('DPA',Tlocal[0]);
    TOBIndice.putvalue ('DPR',Tlocal[1]);
    TOBIndice.putvalue ('PVHT',Tlocal[2]);
    TOBIndice.putvalue ('PVTTC',Tlocal[3]);
    TOBIndice.putvalue ('PMAP',Tlocal[6]);
    TOBIndice.putvalue ('PMRP',Tlocal[7]);
    Valeurs[0]:=Valeurs[0]+ TOBIndice.Getvalue ('DPA');
    Valeurs[1]:=Valeurs[1]+ TOBIndice.Getvalue ('DPR');
    Valeurs[2]:=Valeurs[2]+ TOBIndice.Getvalue ('PVHT');
    Valeurs[3]:=Valeurs[3]+ TOBIndice.Getvalue ('PVTTC');
    Valeurs[6]:=Valeurs[6]+ TOBIndice.Getvalue ('PMAP');
    Valeurs[7]:=Valeurs[7]+ TOBIndice.Getvalue ('PMRP');
    *)
//    CumuleElementTable (TOBIndice,ExAucun,valeurs);
    if ArticleOKInPOUR (TOBIndice.Getvalue('GA_TYPEARTICLE'),ArticleOk) then CumuleElementTable (TOBIndice,ExAucun,TSumPourcent);
  end;
  FormatageTableau (valeurs,V_PGI.okdecP);
  if (IPercent >= 0) then
  begin
    Tlocal := CalculSurTableau ('*',TSumPourcent,TOBLig.detail[IPercent].getvalue('GNL_QTE')/valeur('100'));
    Tlocal[2] := arrondi(TLocal[2],V_PGI.OkDecP );
    FormatageTableau (Tlocal,V_PGI.okdecP);
    valeurs := CalculSurTableau ('+',valeurs,Tlocal);
    FormatageTableau (valeurs,V_PGI.okdecP);
  end;
end;
{$ENDIF}

procedure TFNomenLig.constitueTOBEtat(TOBEtat : TOB);
var Indice : integer;
    TOBL,TOBLE : TOB;
begin
  for Indice := 0 to TOBLig.detail.count -1 do
  begin
    TOBL := TOBLig.detail[Indice];
    TOBLE := TOB.Create('UNE LIGNE',TOBEtat,-1);
    TOBLE.AddChampSupValeur ('GNE_LIBELLE',GNE_LIBELLE.text);
    TOBLE.AddChampSupValeur ('GNL_NOMENCLATURE',TOBL.getValue('GNL_NOMENCLATURE'));
    TOBLE.AddChampSupValeur ('GNL_LIBELLE',TOBL.getValue('GNL_LIBELLE'));
    TOBLE.AddChampSupValeur ('GNE_QTEfDUDETAIL',TOBL.getValue('QTEDUDETAIL'));
    TOBLE.AddChampSupValeur('UNITEOUV',Tunite.Caption);
    TOBLE.AddChampSupValeur('GNL_CODEARTICLE','');
    TOBLE.putValue('GNL_CODEARTICLE',TOBL.GetValue('GNL_CODEARTICLE'));
    TOBLE.AddChampSupValeur('GNL_LIBELLE',TOBL.GetValue('GNL_LIBELLE'));
    TOBLE.AddChampSupValeur('QTELIGNE',TOBL.GetValue('GNL_QTE'));
    TOBLE.AddChampSupValeur('UNITELIG',TOBL.GetValue('UNITENOMENC'));
    TOBLE.AddChampSupValeur('PALIGNE',TOBL.GetValue('DPA'));
    if TOBL.GetValue('GNL_PRIXFIXE') <> 0 then
    	TOBLE.AddChampSupValeur('PUHTLIGNE',TOBL.GetValue('GNL_PRIXFIXE'))
    else
    	TOBLE.AddChampSupValeur('PUHTLIGNE',TOBL.GetValue('PVHT'));
    if TOBL.detail.count = 0 then
    begin
      TOBLE.AddChampSupValeur('COEFPAPR',TOBL.GetValue('COEFFG'));
      TOBLE.AddChampSupValeur('COEFPRPV',TOBL.GetValue('COEFMARG'));
    end else
    begin
      if TOBL.GetValue('DPA')<>0 then TOBLE.AddChampSupValeur('COEFPAPR',TOBL.GetValue('DPR')/TOBL.GetValue('DPA'));
      if TOBL.GetValue('DPR')<>0 then TOBLE.AddChampSupValeur('COEFPRPV',TOBLE.GetValue('PUHTLIGNE')/TOBL.GetValue('DPR'));
    end;
    TOBLE.AddChampSupValeur('PRLIGNE',TOBL.GetValue('DPR'));
  end;
end;

procedure TFNomenLig.RecupInfoPrixPose(TOBmere: TOB);
begin
  //
  if TOBmere.detail.count = 0 then exit;
  TOBmere.PutValue('GCA_PRIXBASE',TOBmere.detail[0].GetValue('GCA_PRIXBASE'));
  TOBmere.PutValue('GCA_QUALIFUNITEACH',TOBmere.detail[0].GetValue('GCA_QUALIFUNITEACH'));
  TOBmere.PutValue('GF_CALCULREMISE',TOBmere.detail[0].GetValue('GF_CALCULREMISE'));
  TOBmere.PutValue('GA_PAHT',TOBmere.detail[0].GetValue('GA_PAHT'));
  TOBmere.PutValue('GA_PVHT',TOBmere.detail[0].GetValue('GA_PVHT'));
  TOBmere.PutValue('GA_FOURNPRINC',TOBmere.detail[0].GetValue('GA_FOURNPRINC'));
  TOBmere.PutValue('TGA_FOURNPRINC',TOBmere.detail[0].GetValue('TGA_FOURNPRINC'));
  if TOBmere.detail.count = 1 then exit;
  TOBmere.putvalue('GA_NATUREPRES',TOBmere.detail[1].GetValue('GNL_CODEARTICLE')) ;
  TOBmere.putvalue('TGA_PRESTATION',TOBmere.detail[1].GetValue('GNL_LIBELLE')) ;
  TOBmere.putvalue('GA_HEURE',TOBmere.detail[1].GetValue('GNL_QTE')) ;
  TOBmere.PutValue('GF_PRIXUNITAIRE',TOBmere.detail[1].GetValue('PVHT'));
end;

Procedure TFNomenLig.ChargeEcranOuvrage;
begin

  Top     := 0;
  left    := 0;
  Width   := Screen.Monitors [0].Width;
  height  := Screen.Monitors [0].Height - 80;

  Caption := TraduireMemoire('Saisie de l''ouvrage ' + GNE_NOMENCLATURE.Text);
  Domaine := GetDomaineOuvrage(GNE_NOMENCLATURE.Text);

  TGNE_NOMENCLATURE.Caption := TraduireMemoire('Déclinaison');
  TGNE_LIBELLE.caption      := TraduireMemoire('Désignation');
  TGNL_SOUSNOMEN.Caption    := TraduireMemoire('Ouvrage');
  TGNE_ARTICLE.Caption      := TraduireMemoire('Ouvrage');
  //
  Tunite.caption := TraduireMemoire(UniteOuvrage(GNE_ARTICLE.text));

  if Action = taconsult then
  begin
    // grisage des boutons
    bNewligne.Enabled := false;
    bdelligne.Enabled := false;
  end;

  PANDetailBtp.Visible  := true;
  PLigne.visible        := true;
  Pouvrage.visible      := true;

end;

Procedure TFNomenlig.GestionGrilleOuvrage;
begin

    G_NLIG.Colwidths[SGN_Joker] := 0;
    G_NLIG.Collengths[SGN_Joker] := -1;

    if Action <> taConsult then
    begin
      if SGN_TYPA   <> -1 then G_NLIG.ColLengths[SGN_TYPA]    := -1;
      if SGN_FOURN  <> -1 then G_NLIG.ColLengths[SGN_FOURN]   := -1;
      if SGN_TARIFF <> -1 then G_NLIG.ColLengths[SGN_TARIFF]  := -1;
      if SGN_UNAC   <> -1 then G_NLIG.ColLengths[SGN_UNAC]    := -1;
      if SGN_REMF   <> -1 then G_NLIG.ColLengths[SGN_REMF]    := -1;
      if SGN_PAHT   <> -1 then G_NLIG.ColLengths[SGN_PAHT]    := -1;
      if SGN_PVHT   <> -1 then G_NLIG.ColLengths[SGN_PVHT]    := -1;
      if SGN_PREST  <> -1 then G_NLIG.ColLengths[SGN_PREST]   := -1;
      if SGN_TPS    <> -1 then G_NLIG.ColLengths[SGN_TPS]     := -1;
      if SGN_MO     <> -1 then G_NLIG.ColLengths[SGN_MO]      := -1;
    end;

    if SGA_UVTE     <> -1 then G_NLIG.ColLengths[SGA_UVTE]    := -1;

    G_NLIG.Collengths[SGN_MTHT] := -1;

    if (SGN_Prix <> -1)   then G_NLIG.colformats [SGN_PRIX]   := '# ### ##0.'+ formatDec ('0',V_PGI.OkDecP );
    if (SGN_PAHT <> -1)   then G_NLIG.colformats [SGN_PAHT]   := '# ### ##0.'+ formatDec ('0',V_PGI.okdecP);
    if (SGN_PVHT <> -1)   then G_NLIG.colformats [SGN_PVHT]   := '# ### ##0.'+ formatDec ('0',V_PGI.okdecP);
    if (SGN_TPS <> -1)    Then G_NLIG.colformats [SGN_TPS]    := '# ### ##0.'+ formatDec ('0',V_PGI.okdecQ);
    if (SGN_TARIFF <> -1) then G_NLIG.colformats [SGN_TARIFF] := '# ### ##0.'+ formatDec ('0',V_PGI.okdecP);
    if (SGN_MO <> -1)     then G_NLIG.colformats [SGN_MO]     := '# ### ##0.'+ formatDec ('0',V_PGI.OkDecP );
    if (SGN_QTESAIS <> -1) then G_NLIG.colformats [SGN_QTESAIS] := '# ### ##0.'+ formatDec ('0',V_PGI.OkDecQ );

    G_NLIG.colformats[SGN_MTHT] :='# ### ##0.'+ formatDec ('0',V_PGI.okdecV);

    NEPA.Decimals             := V_PGI.okdecP;
    NEPA.NumericType          := ntDecimal;
    NEPR.Decimals             := V_PGI.okdecP;
    NEPR.NumericType          := ntDecimal;
    NEPVHT.Decimals           := V_PGI.okdecP;
    NEPVHT.NumericType        := ntDecimal;
    NEPVFORFAIT.Decimals      := V_PGI.okdecP;
    NEPVFORFAIT.NumericType   := ntDecimal;
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
    if SGN_TYPA <> -1 then
    begin
      G_NLIG.ColWidths [SGN_TYPA] := 15;
      if action <> taconsult then G_NLIG.Collengths [SGN_TYPA] := -1;
    end;

end;

Procedure TFNomenLig.GestionTableauGrille;
Var TobInd    : TOB;
    TOBL      : TOB;
    Indice    : Integer;
    ArticleOk : string;
    TValLoc   : T_Valeurs;
begin

  // Chargement de l'unite de vente de l'article dans la cellule
  // + calcul des différentes nomenclatures incluses dans le niveau courant
  IPercent := 0;

  InitTableau (TValeurOuv);
  InitTableau (TValeurPou);
  InitTableau (TPercent);
  InitTableau (TMontantLig);

  for indice := 1 to G_NLIG.RowCount - 1 do
  begin
    //
    TobInd := Tob(G_NLIG.Objects[SGN_IdComp, Indice]);
    //
    if G_NLIG.Cells[SGN_Comp, Indice] = '' then break;
    //
    if TobInd.getvalue('GA_TYPEARTICLE') = 'POU' then
    begin
      ArticleOk := TobInd.getvalue('GA_LIBCOMPL');
      break;
    end;

  //end;

  //for indice := 1 to G_NLIG.RowCount - 1 do
  //begin
     TOBL := TOB(G_NLIG.Objects[SGN_NumLig, Indice]);
     //
     ChampSupNomen (Tobl);

     if TOBL.GetValue('GNL_SOUSNOMEN') = '' then
     begin
        renseigneValoArt (TOBL,TobInd,nil,false);
     end;

     TobL.putvalue('UNITENOMENC',TobInd.getvalue('GA_QUALIFUNITEVTE'));
     G_NLIG.cells[SGA_UVTE,Indice] := TobInd.getvalue('GA_QUALIFUNITEVTE');

     if TobInd.getvalue('GA_TYPEARTICLE') = 'POU' then
     begin
       IPercent := Indice;
       continue;
     end;

     if TOBL.GetValue('GNL_SOUSNOMEN') <> '' then
     begin
       chargePrixOuvrage (TOBL, TOBL.getvalue('GNL_SOUSNOMEN'), Indice);
       if TobInd.getvalue('GA_TYPEARTICLE') = 'ARP' then RecupInfoPrixPose (TOB(G_NLIG.Objects[SGN_NumLig, Indice]));
     end;

     InitTableau (TValloc);

     TobL.putvalue('QTEDUDETAIL',1);

     CumuleElementTable (TobL,ExAucun, TValLoc);

     TValeurOuv := CalculSurTableau ('+',TValeurOuv,TValloc);

     if ArticleOKInPOUR (TobInd.getvalue('GA_TYPEARTICLE'),ArticleOk) then

     TValeurPou := CalculSurTableau ('+',TValeurPou,TValloc);
  end;

  FormatageTableau (TvaleurOuv);

  if IPercent <> 0  then
  begin
    // Article géré en pourcentage de l'ouvrage
    TPercent := calculeMontantLig (IPercent);
    FormatageTableau (TPercent);
  end;

  if (GNE_QTEDUDETAIL.value <> 0) then
  begin
    TGNE_QTEDUDETAIL.Visible := true;
    GNE_QTEDUDETAIL.Visible := true;
  end;

  PANDetailBtp.Visible := true;

  AlimenteValeurOuv ();

end;


procedure TFNomenLig.AppelXLSClick(Sender: TObject);
Var TobInd      : TOB;
    RepServeur  : string;
    RepPoste    : string;
begin

  if (G_NLIG.Cells[SGN_Comp, G_NLIG.Row] = '') then Exit;

  TobInd        := Tob(G_NLIG.Objects[SGN_NumLig, G_NLIG.Row]);

  MetreArticle  := TMetreArt.CreateArt(TobInd.getvalue('GA_TYPEARTICLE'), TobInd.getvalue('GNL_ARTICLE'));

  if MetreArticle.ControleMetre then
  begin
    RepServeur  := MetreArticle.RepMetre + MetreArticle.fFileName;
    RepPoste    := MetreArticle.RepMetreLocal + MetreArticle.fFileName;

    //On copie le fichier de l'article se trouvant sur le serveur sur le poste Local
    if not FileExists(RepPoste) then
    begin
      if FileExists(RepServeur) then
        MetreArticle.CopieLocaltoServeur(RepServeur, RepPoste)
      else
        MetreArticle.CopieLocaltoServeur(MetreArticle.fFichierVide, RepPoste);     //On charge un fichier vide...
    end;

    // on mettra ici l'appel du fichier XLS identique à celui de l'article
    MetreArticle.OuvrirMetreXLs;
  end;

  FreeAndNil(MetreArticle);

end;

procedure TFNomenLig.Variables_LigClick(Sender: TObject);
Var Tobind : TOB;
begin

  if (G_NLIG.Cells[SGN_Comp, G_NLIG.Row] = '') then Exit;

  TobInd := Tob(G_NLIG.Objects[SGN_NumLig, G_NLIG.Row]);

  //On mettra ici l'appel des variables identiques à celui de l'a rticle...
  AGLLanceFiche('BTP', 'BTVARIABLE', 'B;'+TobInd.getvalue('GNL_ARTICLE'),'', 'LIBELLEARTICLE=' + TobInd.getvalue('GNL_LIBELLE') + ';TYPEARTICLE=' + TobInd.getvalue('GA_TYPEARTICLE'));

end;

function TFNomenLig.CalculeQteSais (TOBL: TOB) : double ;
begin
  Result:= TOBL.GeTDouble('GNL_QTE');
  if not GetParamSocSecur('SO_BTGESTQTEAVANC',false) then exit;
  //
  if POS(TOBL.GetString('BNP_TYPERESSOURCE'),'SAL;INT;ST') > 0 then
  begin
    if TOBL.GetDouble('GNL_RENDEMENT')<> 0 then
    begin
      result := TOBL.GeTDouble('GNL_QTE')* TOBL.GetDouble('GNL_RENDEMENT');
    end;
  end else
  begin
    if TOBL.GetDouble('GNL_PERTE')<> 0 then
    begin
      result := Arrondi(TOBL.GeTDouble('GNL_QTE')/ TOBL.GetDouble('GNL_PERTE'),V_PGI.okdecQ);
    end;
  end;
end;

function TFNomenLig.CalculeQteFact (TOBL: TOB) : double ;
begin
  Result:= TOBL.GeTDouble('GNL_QTESAIS');
  if not GetParamSocSecur('SO_BTGESTQTEAVANC',false) then exit;
  //
  if POS(TOBL.GetString('BNP_TYPERESSOURCE'),'SAL;INT;ST') > 0 then
  begin
    if TOBL.GetDouble('GNL_RENDEMENT')<> 0 then
    begin
      result := TOBL.GeTDouble('GNL_QTESAIS')/ TOBL.GetDouble('GNL_RENDEMENT');
    end;
  end else
  begin
    if TOBL.GetDouble('GNL_PERTE')<> 0 then
    begin
      result := Arrondi(TOBL.GeTDouble('GNL_QTESAIS')* TOBL.GetDouble('GNL_PERTE'),V_PGI.okdecQ);
    end;
  end;
end;

Initialization
InitNomenLig();

end.
