unit SaisRbp ;

//=======================================================
//========== Questions sur la saisie bordereau ==========
//=======================================================
// 1. Solde précédent dans l'entête ?

//=======================================================
//======== +++++++ === +++ ??? +++ === +++++++ ==========
//=======================================================

//=======================================================
//======== +++++++ ===   LEGENDE    === +++++++ =========
//=======================================================
// ( ) Reste à faire
// (o) Implémenté mais non testé
// (x) Implémenté et testé
// (?) Non reproduit
//=======================================================
//======== +++++++ === +++ TODO +++ === +++++++ =========
//=======================================================

//=======================================================
//===== +++++++ === +++ CORRECTIONS +++ === +++++++ =====
//=======================================================

//=======================================================
//===== +++++++ === +++ MODIF. BASE +++ === +++++++ =====
//=======================================================

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Grids, Hctrls, Mask, ExtCtrls, ComCtrls,
  Buttons, Hspliter, Ent1, HCompte, DBTables, DB,HEnt1,
  General, Tiers, Journal, hmsgbox, HQry,Menus,MulGene,MulTiers, SaisLett ,
  SaisTVA, EcheMono, SaisAnal, SaisComp, SaisUtil, SaisComm, EncUtil, LettUtil,
  Choix, Echeance,Formule, HFLabel, Chancel, About, SaisVisu, HStatus,
  Devise, FichComm, Scenario, HDebug, HLines, ValPerio, Filtre,
  HSysMenu, SelGuide, DocRegl, SaisEnc, SaisBase, Lookup,
  HPop97, HTB97, ed_tools, SaisEcar, HPanel, UiUtil, Zecrimvt, Saisie,
  TZ, ZTypes, ZReleve, ZGuide, ZEch, ZDevise, ZParams, ZCompte, ZTiers, ZJournal ;

procedure SaisieRbp(Action : TActionFiche) ;

const MAXLIGNES            = 2 ;

const SR_ETABL   : Integer = 0 ;
      SR_NEW     : Integer = 1 ;
      SR_NUML    : Integer = 2 ;
      SR_GEN     : Integer = 3 ;
      SR_JOUR    : Integer = 4 ;
      SR_PB      : Integer = 5 ;
      SR_DEBIT   : Integer = 6 ;
      SR_CREDIT  : Integer = 7 ;
      SR_REGIME  : Integer = 8 ;
      SR_GENERAL : Integer = 9 ;
      SR_IMP     : Integer = 10;
      SR_TVA     : Integer = 11;
      SR_LIB     : Integer = 12;
      SR_PIECE   : Integer = 13;
      SR_FIRST   : Integer = 4 ;
      SR_LAST    : Integer = 13;

const RC_SELGUIDE          = 0 ;
      RC_BADWRITE          = 1 ;
      RC_BADGUIDE          = 2 ;
      RC_SELREGIME         = 3 ;
      RC_SELCODETVA        = 4 ;

const MAX_ROW              = 100 ;

type
  TFSaisieRbp = class(TForm)
    GS              : THGrid;
    PEntete: THPanel;
    E_JOURNAL       : THValComboBox;
    H_JOURNAL       : THLabel;
    H_DATECOMPTABLE : THLabel;
    E_DEVISE        : THValComboBox;
    H_DEVISE        : THLabel;
    PPied: THPanel;
    G_LIBELLE: THLabel;
    SA_SOLDEG       : THNumEdit;
    SA_TOTALDEBIT: THNumEdit;
    HReleve: THMsgBox;
    POPS            : TPopupMenu;
    FindSais        : TFindDialog;
    HMTrad: THSystemMenu;
    DockTop: TDock97;
    DockRight: TDock97;
    DockLeft: TDock97;
    DockBottom: TDock97;
    Valide97: TToolbar97;
    BValide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    Outils: TToolbar97;
    BChercher: TToolbarButton97;
    ToolbarSep971: TToolbarSep97;
    HConf: TToolbarButton97;
    ISigneEuro: TImage;
    LSA_SOLDEG: THLabel;
    LSA_TOTALCREDIT: THLabel;
    LSA_TOTALDEBIT: THLabel;
    E_DATECOMPTABLE: THValComboBox;
    Bevel1: TBevel;
    SA_TOTALCREDIT: THNumEdit;
    Bevel5: TBevel;
    HLabel1: THLabel;
    BEVELSOLDEPROG: TBevel;
    HLSOLDEPROG: THLabel;
    SA_SOLDEPROG: THNumEdit;
    LSA_SOLDEPROG: THLabel;
    SELGUIDE: TEdit;
    FlashGuide: TFlashingLabel;
    BCreateRow: TToolbarButton97;
    BDeleteRow: TToolbarButton97;
    BTools: TToolbarButton97;
    FLASHDEVISE: TFlashingLabel;
    BGuide: TToolbarButton97;
    CRB_REGIMETVA: THValComboBox;
    BEdit: TToolbarButton97;
    CRB_CODETVA: THValComboBox;
    BHideFocus: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure E_DEVISEChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure GereAffSolde(Sender: TObject);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure GSElipsisClick(Sender: TObject);
    procedure GSEnter(Sender: TObject);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean;
      Chg: Boolean);
    procedure GSKeyPress(Sender: TObject; var Key: Char);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean;
      Chg: Boolean);
    procedure BValideClick(Sender: TObject);
    procedure E_DATECOMPTABLEChange(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BDeleteRowClick(Sender: TObject);
    procedure BCreateRowClick(Sender: TObject);
    procedure BToolsClick(Sender: TObject);
    procedure POPSPopup(Sender: TObject);
    procedure GSMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure GSExit(Sender: TObject);
    procedure FindSaisFind(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure BGuideClick(Sender: TObject);
    procedure E_JOURNALExit(Sender: TObject);
    procedure BEditClick(Sender: TObject);
    procedure BHideFocusEnter(Sender: TObject);
    procedure PFENMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    Jal                 : TZJournal ;
    memSaisie, memPivot : RDevise ;
    memOptions          : ROpt ;
    memParams           : RPar ;
//    NumReleve           : LongInt ;
    bFirst, bExo        : Boolean ;
    Comptes             : TZCompte ;
    ObjReleve           : TZReleve ;
    GridX, GridY        : Integer ;
    bGuideRun           : Boolean ;
//    LastModePaie        : string ;
    bStopDevise         : Boolean ;
//    CodeGen             : string ;
    bClosing            : Boolean ;
    bModeRO             : Boolean ;
    FindFirst           : Boolean ;
    CurJal              : string ;
    CurPeriode          : string ;
    // Fonctions ressource
    function  GetMessageRC(MessageID : Integer) : string ;
    procedure PrintMessageRC(MessageID : Integer) ;
    // Fonctions utilitaires
    procedure InitReleve ;
    procedure ChargeJal(CodeJal : string) ;
    procedure GetLastJal ;
    procedure PosLesSoldes ;
    procedure InitPivot ;
    procedure InitSaisie(bReleve : Boolean) ;
    function  CanCloseSaisie : Boolean ;
    procedure CloseSaisie ;
    procedure InitEntete ;
    procedure InitPied ;
    procedure InitGrid ;
    procedure InitTotaux ;
    procedure InfosPied(Row : LongInt) ;
    procedure SetOrdre ;
    function  GetRowDate(Row : LongInt) : TDateTime ;
    procedure EnableButtons ;
    procedure NextRegime(Row : LongInt) ;
    procedure NextCodeTva(Row : LongInt) ;
    // Outils comptabilité
//    function  GetCompte(var NumCompte : string) : Integer ;
//    function  GetAux(var NumAux : string) : Integer ;
    // Fontions TZF
    function  SetRowTZF(Row : LongInt) : TZF ;
    function  GetRowTZF(Row : LongInt) : Boolean ;
    procedure SetRowBad(Row : LongInt) ;
    // Saisie complémentaires
    function  GetGuideCode : string ;
    function  GenGuide(Row : LongInt) : string ;
    function  DelGuide(Code : string) : Boolean ;
    function  GetGuideValue(Col, Row : LongInt) : Boolean ;
    // Calculs
    procedure CalculTotalReleve ;
    procedure CalculSoldeCompte(TotDebit, TotCredit : double) ;
    procedure SoldeProgressif(Row : LongInt) ;
    function  GetRowFormule(var Formule : string ) : Integer ;
    function  GetFormule(Formule : string ) : Variant ;
    // Fonctions Grid
    procedure UpdateScrollBars ;
    procedure SetGridRO ;
    procedure SetGridOn(bMove : Boolean=TRUE) ;
    procedure SetGridOptions(Row : LongInt) ;
    procedure CreateRow(Row : LongInt) ;
    function  DeleteRow(Row : LongInt) : Integer ;
    function  NextRow(Row : LongInt) : Boolean ;
    function  IsRowValid(Row : LongInt; var ACol : Integer; bTest: Boolean) : Boolean ;
    function  GetGridSens(ACol, ARow : Integer) : Integer ;
    procedure PostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    procedure SetGridGrise(ACol, ARow : Integer ; Canvas : TCanvas) ;
    function  IsRowCanEdit(Row : LongInt) : Boolean ;
    function  IsRowOk(Row : LongInt) : Boolean ;
    function  IsCellOK(Col, Row : LongInt) : Boolean ;
    // SQL
    procedure WriteReleve ;
    function  SetLock : Boolean ;
    function  ReadReleve : Boolean ;
    // Click boutons
    function  ValideReleve : Boolean ;
    procedure SearchClick ;
    procedure ValClick ;
    procedure DelClick ;
    procedure InsClick ;
    procedure ToolsClick ;
    procedure GuiClick ;
    procedure GuideRun ;
    procedure GuideStop ;
    procedure EditClick ;
    // Fonctions Zoom
  public
    SoldeProg : Double ;
    nbLignes  : LongInt ;
    Action    : TActionFiche ;
  end;

implementation

{$R *.DFM}

//=======================================================
//======== Point d'entrée dans la saisie relevé =========
//=======================================================
procedure SaisieRbp(Action : TActionFiche) ;
var Releve : TFSaisieRbp ; PP : THPanel ;
begin
Releve:=TFSaisieRbp.Create(Application) ;
Releve.Action:=Action ;
PP:=FindInsidePanel ;
if PP=nil then
   begin
   try
     Releve.ShowModal ;
   finally
     Releve.Free ;
   end ;
  Screen.Cursor:=SyncrDefault ;
  end else
  begin
  InitInside(Releve, PP) ;
  Releve.Show ;
  end ;
end ;

//=======================================================
//================= Fonctions Ressource =================
//=======================================================
function TFSaisieRbp.GetMessageRC(MessageID : Integer) : string ;
begin
Result:=HReleve.Mess[MessageID] ;
end ;

procedure TFSaisieRbp.PrintMessageRC(MessageID : Integer) ;
begin
HReleve.Execute(MessageID, Caption, '') ;
end ;

//=======================================================
//================ Evénements de la Form ================
//=======================================================
procedure TFSaisieRbp.FormCreate(Sender: TObject);
begin
ObjReleve:=nil ; Jal:=nil ; bGuideRun:=FALSE ; bClosing:=FALSE ;
// Paramètres courant
CurJal:='' ; CurPeriode:='' ;
// Paramètres de saisie
memParams.bAnalytique:=FALSE ; memParams.bEcheance:=FALSE ;
memParams.bLibre:=FALSE ;
// Paramètres DEBUG
memParams.bPiece:=FALSE ;     memParams.bSoldeProg:=FALSE ;
InitPivot ;
InitGrid ;
PosLesSoldes ;
if Action=taCreat then Action:=taModif ;
end;

procedure TFSaisieRbp.FormShow(Sender: TObject);
begin
bModeRO:=(Action=taConsult) ;
LookLesDocks(Self) ;
InitSaisie(FALSE) ;
end;

procedure TFSaisieRbp.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
//if not CanCloseSaisie then begin Action:=caNone ; Exit ; end ;
if     Jal<>nil then begin     Jal.Free ;     Jal:=nil ; end ;
if Comptes<>nil then begin Comptes.Free ; Comptes:=nil ; end ;
CloseSaisie ;
if Parent is THPanel then Action:=caFree ;
bClosing:=TRUE ;
end;

procedure TFSaisieRbp.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
CanClose:=CanCloseSaisie ; if CanClose then CloseSaisie ;
end;

procedure TFSaisieRbp.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var Vide : Boolean ;
begin
if not GS.SynEnabled then begin Key:=0 ; Exit ; end ;
(* PFU : Pour tester une fonction *)
if (Key=VK_F10) and (ssAlt in Shift) and (ssCtrl in Shift) and (V_PGI.SAV) then ToolsClick ;
if not (Screen.ActiveControl=GS) then Exit ;
Vide:=(Shift=[]) ;
if (Vide) and ((GS.Col=SR_REGIME) or (GS.Col=SR_TVA))
   and (Key in [VK_DELETE,VK_BACK]) then
   begin
   if (GS.Col=SR_TVA) and (Key=VK_DELETE) then GS.Cells[SR_TVA, GS.Row]:='' ;
   Key:=0 ;Exit ;
   end ;
case Key of
  VK_RETURN : if (Vide) then Key:=VK_TAB ;
{+}  VK_ADD : if (Vide) and ((GS.Col=SR_IMP) or (GS.Col=SR_REGIME) or (GS.Col=SR_TVA))then begin Key:=0 ; GSElipsisClick(nil); end ;
   VK_SPACE : if (Vide) then
                if (GS.Col=SR_REGIME) then begin Key:=0 ; NextRegime(GS.Row) ; end else
                  if (GS.Col=SR_TVA) then begin Key:=0 ; NextCodeTva(GS.Row) ; end ;
  VK_INSERT : if (Vide) then begin Key:=0; InsClick ; end ;
  VK_DELETE : if  Shift=[ssCtrl] then begin Key:=0 ; DelClick ; end ;
  VK_ESCAPE,
     VK_F10 : if (Vide) then begin Key:=0 ; ValClick ; end ;
{CTRL+G} 71 : if Shift=[ssCtrl] then begin Key:=0 ; GuiClick ; end ;
     VK_END : if Shift=[ssCtrl] then
                 begin
                 Key:=0 ; SetGridOptions(nbLignes) ;
                 GS.Row:=nbLignes ;
                 if  GS.Col<>SR_LAST then
                   begin
                   PostMessage(GS.Handle, WM_KEYDOWN, VK_RIGHT,  0) ;
                   PostMessage(GS.Handle, WM_KEYDOWN, VK_LEFT, 0) ;
                   end else
                   begin
                   PostMessage(GS.Handle, WM_KEYDOWN, VK_LEFT,  0) ;
                   PostMessage(GS.Handle, WM_KEYDOWN, VK_RIGHT, 0) ;
                   end ;
                 end ;
  end ;
end;

//=======================================================
//================ Evénements des contrôles =============
//=======================================================
procedure TFSaisieRbp.E_JOURNALExit(Sender: TObject);
var CodeJal : string ;
begin
if bClosing then Exit ;
if not E_JOURNAL.Enabled then Exit ;
if CurJal=E_JOURNAL.Text then Exit ;
if Screen.ActiveControl=E_DATECOMPTABLE then Exit ;
if ValideReleve then InitReleve ;
CodeJal:=E_JOURNAL.Value ;
if CodeJal='' then begin if Jal<>nil then Jal.Free ; Jal:=nil ; Exit ; end ;
// Alimenter le record RJal
ChargeJal(CodeJal) ;
if not ReadReleve then begin E_JOURNAL.SetFocus ; Exit ; end ;
CalculTotalReleve ;
// Activer le Grid
InfosPied(-1) ;
EnableButtons ;
SetGridOn ;
end;

procedure TFSaisieRbp.E_DATECOMPTABLEChange(Sender: TObject);
var Year, Month, Day : Word ;
begin
if bClosing then Exit ;
if not E_DATECOMPTABLE.Enabled then Exit ;
if ValideReleve then InitReleve ;
// Renseigner les options
DecodeDate(StrToDate(E_DATECOMPTABLE.Value), Year, Month, Day) ;
memOptions.MaxJour:=DaysPerMonth(Year, Month) ;
memOptions.Year:=Year ;
memOptions.Month:=Month ;
GS.Enabled:=FALSE ;
CurPeriode:=E_DATECOMPTABLE.Value ;
end;

procedure TFSaisieRbp.E_DEVISEChange(Sender: TObject);
begin
if bClosing then Exit ;
memSaisie.Code:=E_DEVISE.Value ;
if memSaisie.Code<>'' then GetInfosDevise(memSaisie) ;
memSaisie.Taux:=GetTaux(memSaisie.Code,memSaisie.DateTaux,EncodeDate(memOptions.Year, memOptions.Month, 1)) ;
ChangeMask(SA_TOTALDEBIT, memSaisie.Decimale, memSaisie.Symbole) ;
ChangeMask(SA_TOTALCREDIT, memSaisie.Decimale, memSaisie.Symbole) ;
ChangeMask(SA_SOLDEPROG, memPivot.Decimale, memPivot.Symbole) ;
ChangeMask(SA_SOLDEG, memPivot.Decimale, memPivot.Symbole) ;
end;

procedure TFSaisieRbp.BHideFocusEnter(Sender: TObject);
begin
SetGridOn(FALSE) ;
end;

procedure TFSaisieRbp.GereAffSolde(Sender: TObject);
var Nam : string ; c : THLabel ;
begin
if bClosing then Exit ;
Nam:=THNumEdit(Sender).Name ; Nam:='L'+Nam ;
c:=THLabel(FindComponent(Nam)) ;
if c<>nil then c.Caption:=THNumEdit(Sender).Text ;
end;

//=======================================================
//================ Fonctions utilitaires ================
//=======================================================
procedure TFSaisieRbp.InitReleve ;
begin
CloseSaisie ;
InitPivot ;
InitGrid ;
PosLesSoldes ;
InitSaisie(TRUE) ;
end ;

procedure TFSaisieRbp.InitSaisie(bReleve : Boolean) ;
begin
if bClosing then Exit ;
nbLignes:=0 ; SoldeProg:=0 ; bExo:=FALSE ;
if not bReleve then InitEntete ;
InitPied ;
if (bFirst) and not (bReleve) then E_JOURNAL.SetFocus ;
end ;

function TFSaisieRbp.CanCloseSaisie : Boolean ;
begin
Result:=TRUE ;
end ;

procedure TFSaisieRbp.CloseSaisie ;
begin
PurgePopup(POPS) ;
end ;

procedure TFSaisieRbp.InitEntete ;
begin
// Bouton spécial
BHideFocus.Width:=0 ;
// Les boutons fonctionnels
BChercher.Enabled:=FALSE ;  BEdit.Enabled:=FALSE ;
BCreateRow.Enabled:=FALSE ; BDeleteRow.Enabled:=FALSE ;
BGuide.Enabled:=FALSE ;     BValide.Enabled:=FALSE ;    BAbandon.Enabled:=FALSE ;
// Remplir la combo Période
memOptions.Exo:=VH^.Entree.Code ;
if memOptions.Exo=VH^.Encours.Code then memOptions.TypeExo:=teEncours else
   if memOptions.Exo=VH^.Suivant.Code then memOptions.TypeExo:=teSuivant else
      memOptions.TypeExo:=tePrecedent ;
ListePeriode(memOptions.Exo, E_DATECOMPTABLE.Items, E_DATECOMPTABLE.Values, True) ;
// Positionnement par défaut
E_DATECOMPTABLE.Value:=DateToStr(DebutDeMois(V_PGI.DateEntree)) ;
//E_DATECOMPTABLE.Value:=FormatDateTime('01/mm/yyyy', V_PGI.DateEntree) ;
// Positionnement par défaut sur le dernier journal utilisé
GetLastJal ;
// Positionnement par défaut sur la devise pivot
E_DEVISE.Value:=V_PGI.DevisePivot ;
end ;

procedure TFSaisieRbp.InitTotaux ;
begin
end ;

procedure TFSaisieRbp.InitPied ;
begin
InfosPied(-1) ;
InitTotaux ;
ZeroBlanc(PPied) ;
end ;

procedure TFSaisieRbp.ChargeJal(CodeJal : string) ;
var NumCompte : string ;
begin
// Journal
if Jal<>nil then begin Jal.Free ; Jal:=nil ; end ;
Jal:=TZJournal.Create(CodeJal) ;
Jal.Load ;
// Compte de banque
if Comptes<>nil then begin Comptes.Free ; Comptes:=nil ; end ;
Comptes:=TZCompte.Create ;
NumCompte:=Jal.GetValue('J_CONTREPARTIE') ;
if Comptes.GetCompte(NumCompte)<0 then begin Comptes.Free ; Comptes:=nil ; end ;
end ;

procedure TFSaisieRbp.GetLastJal ;
var Q : TQuery ; LastDate : TDateTime ; LastNum : Integer ; LastCode : string ;
begin
LastDate:=EncodeDate(1900, 1, 1) ; LastNum:=-1 ; LastCode:='' ;
Q:=OpenSQL('SELECT J_JOURNAL, J_DATEDERNMVT, J_NUMDERNMVT FROM JOURNAL', TRUE) ;
while not Q.EOF do
      begin
      if Q.FindField('J_DATEDERNMVT').AsDateTime>LastDate then
         begin
         LastDate:=Q.FindField('J_DATEDERNMVT').AsDateTime ;
         LastNum:=Q.FindField('J_NUMDERNMVT').AsInteger ;
         LastCode:=Q.FindField('J_JOURNAL').AsString ;
         end ;
      Q.Next ;
      end ;
Ferme(Q) ;      
if (LastNum<>-1) and (LastCode<>'') then E_JOURNAL.Value:=LastCode ;
end ;

procedure TFSaisieRbp.PosLesSoldes ;
begin
LSA_SOLDEG.SetBounds(SA_SOLDEG.Left,SA_SOLDEG.Top,SA_SOLDEG.Width,SA_SOLDEG.Height) ;
LSA_SOLDEPROG.SetBounds(SA_SOLDEPROG.Left,SA_SOLDEPROG.Top,SA_SOLDEPROG.Width,SA_SOLDEPROG.Height) ;
LSA_TOTALDEBIT.SetBounds(SA_TOTALDEBIT.Left,SA_TOTALDEBIT.Top,SA_TOTALDEBIT.Width,SA_TOTALDEBIT.Height) ;
LSA_TOTALCREDIT.SetBounds(SA_TOTALCREDIT.Left,SA_TOTALCREDIT.Top,SA_TOTALCREDIT.Width,SA_TOTALCREDIT.Height) ;
SA_SOLDEG.Visible:=FALSE ;     SA_SOLDEPROG.Visible:=FALSE ;
SA_TOTALDEBIT.Visible:=FALSE ; SA_TOTALCREDIT.Visible:=FALSE ;
end ;

procedure TFSaisieRbp.InitGrid ;
begin
bFirst:=TRUE ;
// Avant le VidePile pour bien placer le curseur
GS.Row:=GS.FixedRows ; GS.Col:=GS.FixedCols ;
GS.VidePile(FALSE)  ;
GS.RowCount:=MAX_ROW ; GS.TypeSais:=tsFolio ;
GS.ElipsisButton:=FALSE ;
GS.Enabled:=FALSE ; UpdateScrollBars ;
GS.ColWidths[SR_ETABL]:=0 ;
GS.ColWidths[SR_NEW]:=0 ;
GS.ColWidths[SR_NUML]:=0 ;
GS.ColWidths[SR_GEN]:=0 ;
GS.ColAligns[SR_NUML]:=taCenter ;
GS.ColAligns[SR_GEN]:=taCenter ;
GS.ColAligns[SR_JOUR]:=taCenter ;
GS.ColAligns[SR_DEBIT]:=taRightJustify ;
GS.ColAligns[SR_CREDIT]:=taRightJustify ;
GS.PostDrawCell:=PostDrawCell ;
GS.Refresh ;
end ;

procedure TFSaisieRbp.InitPivot ;
begin
memPivot.Code:=V_PGI.DevisePivot ;
if memPivot.Code<>'' then GetInfosDevise(memPivot) ;
end ;

procedure TFSaisieRbp.SetOrdre ;
var i : Integer ;
begin
for i:=GS.FixedRows to nbLignes do GS.Cells[SR_NUML,i]:=IntToStr(i) ;
end ;

function TFSaisieRbp.GetRowDate(Row : LongInt) : TDateTime ;
begin
Result:=EncodeDate(memOptions.Year, memOptions.Month, StrToInt(GS.Cells[SR_JOUR, Row])) ;
end ;

procedure TFSaisieRbp.EnableButtons ;
var bEnabled : Boolean ; ACol : Integer ;
begin
BTools.Visible:=V_PGI.SAV ;
// Le relevé peut être validé ?
bEnabled:=TRUE ;
// Entête
E_JOURNAL.Enabled:=bEnabled ;
E_DATECOMPTABLE.Enabled:=bEnabled ;
BValide.Enabled:=bEnabled ;
// Boutons
BGuide.Enabled:=(not bModeRO)
                and ((nbLignes=0) or ((nbLignes>0) and (GS.Row=nbLignes)))
                and (GS.Cells[SR_DEBIT,  GS.Row]='')
                and (GS.Cells[SR_CREDIT, GS.Row]='');
// Bouton rechercher
BChercher.Enabled:=(nbLignes<>0) ;
// Boutons Ins/Del ligne
BCreateRow.Enabled:=(not bModeRO) and (nbLignes<>0) and IsRowValid(GS.Row, ACol, TRUE) ;
BDeleteRow.Enabled:=(not bModeRO) and (nbLignes<>0) ;
// Bouton édition du mouvement
BEdit.Enabled:=GS.Cells[SR_GEN, GS.Row]<>'' ;
end ;

procedure TFSaisieRbp.NextRegime(Row : LongInt) ;
var i : Integer ;
begin
i:=CRB_REGIMETVA.Items.IndexOf(GS.Cells[SR_REGIME, Row]) ; i:=i+1 ;
if i>CRB_REGIMETVA.Items.Count-1 then i:=0 ;
GS.Cells[SR_REGIME, Row]:=CRB_REGIMETVA.Items[i] ;
end ;

procedure TFSaisieRbp.NextCodeTva(Row : LongInt) ;
var i : Integer ;
begin
i:=CRB_CODETVA.Items.IndexOf(GS.Cells[SR_TVA, Row]) ; i:=i+1 ;
if i>CRB_CODETVA.Items.Count-1 then i:=0 ;
GS.Cells[SR_TVA, Row]:=CRB_CODETVA.Items[i] ;
end ;

procedure TFSaisieRbp.InfosPied(Row : LongInt) ;
var k : Integer ; DIG, CIG : double ; NumCompte : string ; RefRow : LongInt ;
begin
if Row<0 then RefRow:=GS.Row else RefRow:=Row ;
G_LIBELLE.Font.Color:=clBlack ; G_LIBELLE.Caption:='' ;
NumCompte:=GS.Cells[SR_GENERAL, RefRow] ; k:=Comptes.GetCompte(NumCompte) ;
if k>=0 then
   begin
   G_LIBELLE.Caption:=Comptes.GetValue('G_LIBELLE', 0) ;
   DIG:=Comptes.GetValue('G_TOTALDEBIT', 0) ;
   CIG:=Comptes.GetValue('G_TOTALCREDIT', 0) ;
   LSA_SOLDEG.Visible:=TRUE ;
   CalculSoldeCompte(DIG, CIG) ;
   end else LSA_SOLDEG.Visible:=FALSE ;
FLASHDEVISE.Caption:='' ; FLASHDEVISE.Visible:=FALSE ;
end ;

//=======================================================
//=================== Gestion des TZF ===================
//=======================================================
procedure TFSaisieRbp.SetRowBad(Row : LongInt) ;
var Ecr : TZF ;
begin
Ecr:=ObjReleve.GetRow(Row-GS.FixedRows) ;
if Ecr=nil then Exit ;
Ecr.PutValue('BADROW', 'X') ;
end ;

function TFSaisieRbp.SetRowTZF(Row : LongInt) : TZF ;
var Ecr : TZF ; Gen : Integer ;
begin
if Jal=nil then begin Result:=nil ; Exit ; end ;
Ecr:=ObjReleve.GetRow(Row-GS.FixedRows) ;
if Ecr=nil then begin Result:=nil ; Exit ; end ;
// Entête / Commun
Ecr.PutValue('CRB_EXERCICE',  memOptions.Exo) ;
Ecr.PutValue('CRB_JOURNAL',   Jal.GetValue('J_JOURNAL')) ;
if GS.Cells[SR_GEN, Row]='' then Gen:=0 else Gen:=StrToInt(GS.Cells[SR_GEN, Row]) ;
Ecr.PutValue('CRB_NUMEROPIECE', Gen) ;
if Ecr.GetValue('CRB_DEVISE')='' then Ecr.PutValue('CRB_DEVISE', memPivot.Code) ;
Ecr.SetMontants(Valeur(GS.Cells[SR_DEBIT, Row]), Valeur(GS.Cells[SR_CREDIT, Row]),
                memSaisie, FALSE, FALSE) ;
// Corps
Ecr.PutValue('CRB_DATECOMPTABLE', GetRowDate(Row)) ;
Ecr.PutValue('CRB_NUMRELEVE',     GS.Cells[SR_PB, Row]) ;
Ecr.PutValue('CRB_NUMLIGNE',      StrToInt(GS.Cells[SR_NUML, Row])) ;
Ecr.PutValue('CRB_REFINTERNE',    GS.Cells[SR_PIECE, Row]) ;
Ecr.PutValue('CRB_LIBELLE',       GS.Cells[SR_LIB, Row]) ;
Ecr.PutValue('CRB_IMPUTATION',    GS.Cells[SR_IMP, Row]) ;
Ecr.PutValue('CRB_ETABLISSEMENT', GS.Cells[SR_ETABL,Row]) ;
Ecr.PutValue('CRB_REGIMETVA',     CRB_REGIMETVA.Values[CRB_REGIMETVA.Items.IndexOf(GS.Cells[SR_REGIME, Row])]) ;
Result:=Ecr ;
end ;

function TFSaisieRbp.GetRowTZF(Row : LongInt) : Boolean ;
var Ecr : TZF ; Gen : Integer ; Year, Month, Day : Word ;
    EcrDebit, EcrCredit : Double ; i : Integer ;
begin
Ecr:=ObjReleve.GetRow(Row) ; if Ecr=nil then begin Result:=FALSE ; Exit ; end ;
// Ajustement
Row:=Row+GS.FixedRows ;
// Entête / Commun
// Corps
DecodeDate(Ecr.GetValue('CRB_DATECOMPTABLE'), Year, Month, Day) ;
GS.Cells[SR_JOUR, Row]:=IntToStr(Day) ;
GS.Cells[SR_NUML, Row]:=IntToStr(Row) ;
GS.Cells[SR_PB, Row]:=Ecr.GetValue('CRB_NUMRELEVE') ;
Gen:=StrToInt(Ecr.GetValue('CRB_NUMEROPIECE')) ;
if Gen<>0 then GS.Cells[SR_GEN, Row]:=IntToStr(Gen) else GS.Cells[SR_GEN, Row]:='' ;
EcrDebit:=Arrondi(Ecr.GetValue('CRB_DEBIT'), memPivot.Decimale) ;
EcrCredit:=Arrondi(Ecr.GetValue('CRB_CREDIT'), memPivot.Decimale) ;
if EcrDebit<>0 then GS.Cells[SR_DEBIT, Row]:=StrFMontant(EcrDebit,15,memPivot.Decimale,'',TRUE)
               else GS.Cells[SR_DEBIT, Row]:='' ;
if EcrCredit<>0 then GS.Cells[SR_CREDIT, Row]:=StrFMontant(EcrCredit,15,memPivot.Decimale,'',TRUE)
                else GS.Cells[SR_CREDIT, Row]:='' ;
i:=CRB_REGIMETVA.Values.IndexOf(Ecr.GetValue('CRB_REGIMETVA')) ;
if i<0 then i:=0 ;
GS.Cells[SR_REGIME, Row]:=CRB_REGIMETVA.Items[i] ;
GS.Cells[SR_PIECE, Row]:=Ecr.GetValue('CRB_REFINTERNE') ;
GS.Cells[SR_LIB, Row]:=Ecr.GetValue('CRB_LIBELLE') ;
GS.Cells[SR_IMP, Row]:=Ecr.GetValue('CRB_IMPUTATION') ;
GS.Cells[SR_ETABL,Row]:=Ecr.GetValue('CRB_ETABLISSEMENT') ;
Result:=TRUE ;
end ;

//=======================================================
//==================== Fonctions SQL ====================
//=======================================================
function TFSaisieRbp.SetLock : Boolean ;
var THEBADGUY : string ; Ecr : TZF ;
begin
// Gestion de la concurrence
Result:=TRUE ;
(*
if memOptions.NewObj then
  begin
  THEBADGUY:=ObjFolio.GetLock ;
  if THEBADGUY=''
    then ObjFolio.LockFolio(TRUE, nil)
    else begin PrintMessageRC(RC_THEBADGUY, THEBADGUY) ; Result:=FALSE ; end ;
  end else
  begin
  bModeRO:=FALSE ; Action:=taModif ;
  Ecr:=ObjFolio.GetRow(0) ; if (Ecr<>nil) and (Ecr.GetValue('E_CONTROLEUR')=V_PGI.User) then Exit ;
  if not ObjFolio.LockFolio(TRUE, Ecr) then
    begin PrintMessageRC(RC_FOLIOLOCK) ; bModeRO:=TRUE ; Action:=taConsult ; end ;
  end ;
*)  
end ;

procedure TFSaisieRbp.WriteReleve ;
var i : Integer ; ACol : LongInt ;
begin
// En colsultation, rien à faire ici !
if bModeRO then Exit ;
// Je positionne les colonnes
if not IsRowValid(nbLignes, ACol, FALSE) then DeleteRow(nbLignes) ;
for i:=1 to nbLignes do if SetRowTZF(i)=nil then Continue ;
(* LOCK à gérer
if memOptions.NewObj then ObjFolio.LockReleve(FALSE, nil) ;
if not ObjReleve.NeedWrite then
  begin ObjReleve.LockReleve(FALSE, ObjReleve.GetRow(0)) ; Exit ; end ;
*)  
if not memOptions.NewObj then ObjReleve.Del ;
if V_PGI.IOError=oeOK then
  if (nbLignes>=1) then ObjReleve.Write ;
(* ANCIEN CODE
if not memOptions.NewObj then ObjReleve.Del ;
if V_PGI.IOError=oeOK then
  begin
  for i:=1 to nbLignes do
      begin
      if not IsRowOk(i) then begin SetRowBad(i) ; Continue ; end ;
      if SetRowTZF(i)=nil then continue ;
      end ;
  if (nbLignes>1) then ObjReleve.Write ;
  end ;
*)
end ;

function TFSaisieRbp.ReadReleve : Boolean ;
var k : Integer ;
begin
Result:=TRUE ;
if ObjReleve<>nil then ObjReleve.Free ;
InitReleve ;
ObjReleve:=TZReleve.Create(memSaisie.Decimale, Jal.GetValue('J_JOURNAL'), memOptions) ;
memOptions.NewObj:=not ObjReleve.Read ;
if not bModeRO then if not SetLock then begin Result:=FALSE ; Exit ; end ;
// Ok, les paramètres courants peuvent être mis à jour
CurPeriode:=E_DATECOMPTABLE.Value ;
CurJal:=E_JOURNAL.Value ;
// Alimenter le grid
k:=0 ; //memJal.TotDebDebit:=0; memJal.TotDebCredit:=0 ;
while GetRowTZF(k) do
      begin
      nbLignes:=nbLignes+1 ; k:=k+1 ;
      // (PFU : !) Valeurs non présentes dans le TZF
      GS.Cells[SR_NEW,k]:='N' ; // Ligne non créée en saisie
      // Relevé au début de la saisie
//      memJal.TotDebDebit:=memJal.TotDebDebit+Valeur(GS.Cells[SR_DEBIT, k]) ;
//      memJal.TotDebCredit:=memJal.TotDebCredit+Valeur(GS.Cells[SR_CREDIT, k]) ;
      if GS.RowCount-2<nbLignes then GS.RowCount:=GS.RowCount+1 ;
      end ;
GS.Invalidate ;
end ;

//=======================================================
//================ Saisie complémentaires ===============
//=======================================================
function TFSaisieRbp.GetGuideCode : string ;
var Q : TQuery ; sNum : string ; iNum : Integer ;
begin
Q:=OpenSQL('SELECT MAX(GU_GUIDE) FROM GUIDE WHERE GU_TYPE="NOR"', TRUE) ;
sNum:='001' ;
if not Q.EOF then
   begin
   sNum:=Q.Fields[0].AsString ;
   if sNum<>'' then begin iNum:=StrToInt(sNum); sNum:=Format('%.3d', [iNum+1]); end
               else sNum:='001';
   end ;
Ferme(Q) ;   
Result:=sNum ;
end ;

function TFSaisieRbp.GenGuide(Row : LongInt) : string ;
var Q, QSource, QCompte : TQuery ; i, j, kTva : Integer ; TTva : Real ;
    bAchat, bNoTva : Boolean ; RegimeTVA, CodeTVA, TCpt, Cols : string ;
    bAxe : array[1..MAXAXE] of Boolean ;
begin
Result:='' ;
// 'select * from guide left join ecrgui on (gu_type=eg_type) and (gu_guide=eg_guide) where gu_type="RLV" and gu_guide="002" order by EG_TYPE, EG_GUIDE, EG_NUMLIGNE'
QSource:=OpenSQL('SELECT * FROM GUIDE LEFT JOIN ECRGUI ON (GU_TYPE=EG_TYPE)'
                 +' AND (GU_GUIDE=EG_GUIDE) WHERE GU_TYPE="RLV" AND GU_GUIDE="'
                 +GS.Cells[SR_IMP, Row]
                 +'" ORDER BY EG_TYPE, EG_GUIDE, EG_NUMLIGNE', TRUE) ;
if QSource.EOF then begin Ferme(QSource) ; Exit ; end ;
// Génération de l'entête du Guide
Q:=OpenSQL('SELECT * FROM GUIDE WHERE GU_GUIDE="'+W_W+'"', FALSE) ;
Q.Insert ; InitNew(Q) ;
Q.FindField('GU_TYPE').AsString:='NOR' ;
Result:=GetGuideCode ;
Q.FindField('GU_GUIDE').AsString:=Result ;
Q.FindField('GU_LIBELLE').AsString:='Guide généré par la saisie relevé' ;
Q.FindField('GU_JOURNAL').AsString:=Jal.GetValue('J_JOURNAL') ;
Q.FindField('GU_NATUREPIECE').AsString:=QSource.FindField('GU_NATUREPIECE').AsString ;
Q.FindField('GU_DEVISE').AsString:=memPivot.Code ;
//Q.FindField('GU_TYPECTREPARTIE').AsString:= ;
Q.FindField('GU_ETABLISSEMENT').AsString:=QSource.FindField('GU_ETABLISSEMENT').AsString ;
Q.FindField('GU_UTILISATEUR').AsString:=V_PGI.User ;
Q.FindField('GU_DATECREATION').AsDateTime:=Date ;
Q.FindField('GU_DATEMODIF').AsDateTime:=Date ;
Q.FindField('GU_SOCIETE').AsString:=V_PGI.CodeSociete ;
Q.Post ;
Ferme(Q) ;
// Génération de la ligne de banque
Q:=OpenSQL('SELECT * FROM ECRGUI WHERE EG_TYPE="NOR" AND EG_GUIDE="'+W_W+'"', FALSE) ;
Q.Insert ; InitNew(Q) ;
Q.FindField('EG_TYPE').AsString:='NOR' ;
Q.FindField('EG_GUIDE').AsString:=Result ;
Q.FindField('EG_NUMLIGNE').AsInteger:=1 ;
Q.FindField('EG_GENERAL').AsString:=Jal.GetValue('J_CONTREPARTIE') ;
Q.FindField('EG_AUXILIAIRE').AsString:='' ;
Q.FindField('EG_REFINTERNE').AsString:='' ;
Q.FindField('EG_LIBELLE').AsString:=QSource.FindField('GU_LIBELLE').AsString ;
// Calcul débit/crédit
Q.FindField('EG_DEBITDEV').AsString:=GS.Cells[SR_CREDIT, Row] ;
Q.FindField('EG_CREDITDEV').AsString:=GS.Cells[SR_DEBIT, Row] ;
Q.FindField('EG_MODEPAIE').AsString:='' ;
Q.FindField('EG_DATEECHEANCE').AsString:='' ;
Q.FindField('EG_REFEXTERNE').AsString:='' ;
Q.FindField('EG_DATEREFEXTERNE').AsString:='' ;
Q.FindField('EG_REFLIBRE').AsString:='' ;
Q.FindField('EG_AFFAIRE').AsString:='' ;
Q.FindField('EG_QTE1').AsString:='' ;
Q.FindField('EG_QTE2').AsString:='' ;
Q.FindField('EG_QUALIFQTE1').AsString:='' ;
Q.FindField('EG_QUALIFQTE2').AsString:='' ;
Q.FindField('EG_MODEREGLE').AsString:='' ;
//  Q.FindField('EG_ECHEANCES').AsString:='' ;
Q.FindField('EG_ARRET').AsString:='------' ;
Q.FindField('EG_TVAENCAIS').AsString:='-' ;
Q.FindField('EG_TVA').AsString:='' ;
Q.FindField('EG_RIB').AsString:='-' ;
Q.FindField('EG_BANQUEPREVI').AsString:='-' ;
Q.Post ;
Ferme(Q) ;
// Génération des lignes du Guide
bNoTva:=FALSE ; TTva:=0 ; bAchat:=FALSE ;
for i:=1 to MAXLIGNES do begin
  if QSource.EOF then break ;
  Q:=OpenSQL('SELECT * FROM ECRGUI WHERE EG_TYPE="NOR" AND EG_GUIDE="'+W_W+'"', FALSE) ;
  Q.Insert ; InitNew(Q) ;
  Q.FindField('EG_TYPE').AsString:='NOR' ;
  Q.FindField('EG_GUIDE').AsString:=Result ;
  Q.FindField('EG_NUMLIGNE').AsInteger:=i+1 ;
  Q.FindField('EG_GENERAL').AsString:=QSource.FindField('EG_GENERAL').AsString ;
  Q.FindField('EG_AUXILIAIRE').AsString:='' ;
  Q.FindField('EG_REFINTERNE').AsString:='' ;
  Q.FindField('EG_LIBELLE').AsString:=QSource.FindField('EG_LIBELLE').AsString ;
  // Calcul débit/crédit
  if (i=1) then begin
    CodeTVA:=QSource.FindField('EG_TVA').AsString ;
    RegimeTVA:=CRB_REGIMETVA.Values[CRB_REGIMETVA.Items.IndexOf(GS.Cells[SR_REGIME, Row])] ;
    Cols:='G_GENERAL, G_NATUREGENE, G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5' ;
    QCompte:=OpenSQL('SELECT '+Cols+' FROM GENERAUX WHERE G_GENERAL="'+QSource.FindField('EG_GENERAL').AsString+'"', TRUE) ;
    if not QCompte.EOF then
       begin
       bAchat:=(QCompte.FindField('G_NATUREGENE').AsString='CHA') ;
       for j:=1 to MAXAXE do
         bAxe[j]:=(QCompte.FindField('G_VENTILABLE'+IntToStr(j)).AsString='X') ;
       end ;
    Ferme(QCompte) ;
    // Le code TVA de la saisie est prédominant !
    kTva:=CRB_CODETVA.Items.IndexOf(GS.Cells[SR_TVA, Row]) ;
    if kTva>=0 then CodeTVA:=CRB_CODETVA.Values[kTva] ;
    TTva:=TVA2TAUX(RegimeTVA, CodeTVA, bAchat) ;
    TCpt:=TVA2CPTE(RegimeTVA, CodeTVA, bAchat) ;
  end ;
  if TTva>0 then
     begin
     if i=1 then
       begin
       Q.FindField('EG_DEBITDEV').AsString:=StrFMontant(Arrondi(Valeur(GS.Cells[SR_DEBIT, Row])/(1+TTva),memPivot.Decimale),15,memPivot.Decimale,'',TRUE) ;
       Q.FindField('EG_CREDITDEV').AsString:=StrFMontant(Arrondi(Valeur(GS.Cells[SR_CREDIT, Row])/(1+TTva),memPivot.Decimale),15,memPivot.Decimale,'',TRUE) ;
       end else
       begin
       Q.FindField('EG_GENERAL').AsString:=TCpt ;
       Q.FindField('EG_DEBITDEV').AsString:='[SOLDE]' ;
       Q.FindField('EG_CREDITDEV').AsString:='[SOLDE]' ;
       end ;
     end else
     begin
     Q.FindField('EG_DEBITDEV').AsString:=GS.Cells[SR_DEBIT, Row] ;
     Q.FindField('EG_CREDITDEV').AsString:=GS.Cells[SR_CREDIT, Row] ;
     bNoTva:=TRUE ;
     end ;
  Q.FindField('EG_MODEPAIE').AsString:='' ;
  Q.FindField('EG_DATEECHEANCE').AsString:='' ;
  Q.FindField('EG_REFEXTERNE').AsString:='' ;
  Q.FindField('EG_DATEREFEXTERNE').AsString:='' ;
  Q.FindField('EG_REFLIBRE').AsString:='' ;
  Q.FindField('EG_AFFAIRE').AsString:='' ;
  Q.FindField('EG_QTE1').AsString:='' ;
  Q.FindField('EG_QTE2').AsString:='' ;
  Q.FindField('EG_QUALIFQTE1').AsString:='' ;
  Q.FindField('EG_QUALIFQTE2').AsString:='' ;
  Q.FindField('EG_MODEREGLE').AsString:='' ;
//  Q.FindField('EG_ECHEANCES').AsString:='' ;
  Q.FindField('EG_ARRET').AsString:='------' ;
  Q.FindField('EG_TVAENCAIS').AsString:='-' ;
  Q.FindField('EG_TVA').AsString:=QSource.FindField('EG_TVA').AsString ;
  Q.FindField('EG_RIB').AsString:='-' ;
  Q.FindField('EG_BANQUEPREVI').AsString:='-' ;
  Q.Post ;
  Ferme(Q) ;
  // Génération de l'analytique. Pour l'instant Section d'attente
(*
  for NumAxe:=1 to MAXAXE do
    begin
    if (i<>1) then Break ;
    if not bAxe[NumAxe] then Continue ;
    QCompte:=OpenSQL('SELECT X_AXE, X_SECTIONATTENTE FROM AXE WHERE X_AXE="A'+IntToStr(NumAxe)+'"', TRUE) ;
    if QCompte.EOF then begin Ferme(QCompte) ; Continue ; end ;
    Q:=OpenSQL('SELECT * FROM ANAGUI WHERE AG_TYPE="NOR" AND AG_GUIDE="'+W_W+'"', FALSE) ;
    Q.Insert ; InitNew(Q) ;
    Q.FindField('AG_TYPE').AsString:='NOR' ;
    Q.FindField('AG_GUIDE').AsString:=Result ;
    Q.FindField('AG_NUMLIGNE').AsInteger:=1 ;
    Q.FindField('AG_NUMVENTIL').AsInteger:=1 ;
    Q.FindField('AG_AXE').AsString:='A'+IntToStr(NumAxe) ;
    Q.FindField('AG_SECTION').AsString:=QCompte.FindField('X_SECTIONATTENTE').AsString ;
    Q.FindField('AG_POURCENTAGE').AsString:='100' ;
    Q.FindField('AG_POURCENTQTE1').AsString:='' ;
    Q.FindField('AG_POURCENTQTE2').AsString:='' ;
    Q.FindField('AG_ARRET').AsString:='---' ;
    Q.Post ;
    Ferme(Q) ;
    Ferme(QCompte) ;
    end ;
*)    
  if bNoTva then break ;
  QSource.Next ;
  end ;
Ferme(QSource) ;
end ;

function TFSaisieRbp.DelGuide(Code : string) : Boolean ;
begin
Result:=ExecuteSQL('DELETE FROM GUIDE WHERE GU_TYPE="NOR" AND GU_GUIDE="'+Code+'"')>0 ;
ExecuteSQL('DELETE FROM ECRGUI WHERE EG_TYPE="NOR" AND EG_GUIDE="'+Code+'"') ;
(*ExecuteSQL('DELETE FROM ANAGUI WHERE AG_TYPE="NOR" AND AG_GUIDE="'+Code+'"') ;*)
end ;

function TFSaisieRbp.GetGuideValue(Col, Row : LongInt) : Boolean ;
var Ecr : TZF ; bPost : Boolean ; Arret : string ;
begin
Ecr:=ObjReleve.GetRow(Row-GS.FixedRows) ;
if Ecr=nil then begin Result:=FALSE ; Exit ; end ;
Arret:=Ecr.GetValue('EG_ARRET') ; bPost:=FALSE ;
if Col=SR_JOUR then bPost:=TRUE ;
if Col=SR_PIECE then
   begin
   GS.Cells[SR_PIECE, GS.Row]:=GFormule(Ecr.GetValue('EG_REFINTERNE'), GetFormule, nil, 1) ;
   if Arret[3]='-' then bPost:=TRUE ;
   end ;
if Col=SR_LIB then
   begin
   GS.Cells[SR_LIB, GS.Row]:=GFormule(Ecr.GetValue('EG_LIBELLE'), GetFormule, nil, 1) ;
   if Arret[4]='-' then bPost:=TRUE ;
   end ;
if Col=SR_DEBIT then
   begin
   GS.Cells[SR_DEBIT, GS.Row]:=GFormule(Ecr.GetValue('EG_DEBITDEV'), GetFormule, nil, 1) ;
   if Arret[5]='-' then bPost:=TRUE ;
   end ;
if Col=SR_CREDIT then
   begin
   GS.Cells[SR_CREDIT, GS.Row]:=GFormule(Ecr.GetValue('EG_CREDITDEV'), GetFormule, nil, 1) ;
   if Arret[6]='-' then bPost:=TRUE ;
   end ;
Result:=bPost ;
end ;

//=======================================================
//================= Fonctions de calcul =================
//=======================================================
procedure TFSaisieRbp.CalculTotalReleve ;
var i : Integer ; TotalD, TotalC : Double ;
begin
TotalD:=0; TotalC:=0 ;
for i:=GS.FixedRows to nbLignes do
    begin
    TotalD:=TotalD+Valeur(GS.Cells[SR_DEBIT,i]) ;
    TotalC:=TotalC+Valeur(GS.Cells[SR_CREDIT,i]) ;
    end ;
SA_TOTALDEBIT.Value:=TotalD ;
SA_TOTALCREDIT.Value:=TotalC ;
end ;

procedure TFSaisieRbp.CalculSoldeCompte(TotDebit, TotCredit : Double) ;
var i : Integer ; TotCDebit, TotCCredit : Double ;
begin
TotCDebit:=TotDebit ; TotCCredit:=TotCredit ;
for i:=GS.FixedRows to nbLignes do
    begin
    // Attention colonnes inversées
    TotCDebit:=TotCDebit+Valeur(GS.Cells[SR_CREDIT, i]) ;
    TotCCredit:=TotCCredit+Valeur(GS.Cells[SR_DEBIT, i]) ;
    end ;
AfficheLeSolde(SA_SOLDEG, TotCDebit, TotCCredit) ;
end ;

procedure TFSaisieRbp.SoldeProgressif(Row : LongInt) ;
var i : LongInt ; SoldeD, SoldeC : double ;
begin
SoldeD:=0 ; SoldeC:=0 ;
for i:=GS.FixedRows to Row do
    begin
    SoldeD:=SoldeD+Valeur(GS.Cells[SR_DEBIT,i]) ;
    SoldeC:=SoldeC+Valeur(GS.Cells[SR_CREDIT,i]) ;
    end ;
SoldeProg:=Arrondi(SoldeD-SoldeC, memPivot.Decimale) ;
AfficheLeSolde(SA_SOLDEPROG, SoldeD, SoldeC) ;
end ;

function TFSaisieRbp.GetRowFormule(var Formule : string ) : Integer ;
var Pos1, Pos2 : Integer ;
begin
Result:=0 ;
Pos2:=Pos(':L', Formule) ; if Pos2>0 then System.Delete(Formule, Pos2+1, 1) ;
Pos1:=Pos(':',  Formule) ; if (Pos1<=0) or (Pos1=Length(Formule)) then Exit ;
Result:=Round(Valeur(Copy(Formule, Pos1+1, 5))) ;
System.Delete(Formule, Pos1, 5) ;
end ;

function TFSaisieRbp.GetFormule(Formule : string ) : Variant ;
var Ecr, RefEcr : TZF ; RefRow, CurRow, ForRow : Integer ;
begin
Result:=#0 ;
Ecr:=ObjReleve.GetRow(GS.Row-GS.FixedRows) ; if Ecr=nil then Exit ;
Formule:=UpperCase(Trim(Formule)) ; if Formule='' then Exit ;
ForRow:=GetRowFormule(Formule) ; CurRow:=Ecr.GetValue('EG_NUMLIGNE') ;
//RefRow:=0 ;
if ForRow=0 then RefRow:=0 else
  if ForRow<CurRow then RefRow:=ForRow-CurRow else
    if ForRow>CurRow then RefRow:=CurRow-ForRow else
      RefRow:=CurRow ;
RefEcr:=ObjReleve.GetRow(GS.Row-GS.FixedRows+RefRow) ; if RefEcr=nil then Exit ;
if Pos('E_', Formule)>0 then
   begin
   Result:=RefEcr.GetValue(Formule) ;
   Exit ;
   end ;
end ;

//=======================================================
//=================== Gestion du Grid ===================
//=======================================================
function TFSaisieRbp.GetGridSens(ACol, ARow : Integer) : Integer;
begin
// Sens de déplacement dans le Grid
if (GS.Row=ARow) then
   begin
   if (GS.Col>ACol) then Result:=1 else Result:=-1;
   end else if (GS.Row>ARow) then Result:=1 else Result:=-1;
end;

procedure TFSaisieRbp.SetGridGrise(ACol, ARow : Integer ; Canvas : TCanvas) ;
var R : TRect ;
begin
Canvas.Brush.Color := GS.FixedColor ;
Canvas.Brush.Style := bsBDiagonal ;
Canvas.Pen.Color   := GS.FixedColor ;
Canvas.Pen.Mode    := pmCopy ;
Canvas.Pen.Style   := psClear ;
Canvas.Pen.Width   := 1 ;
R:=GS.CellRect(ACol, ARow) ;
Canvas.Rectangle(R.Left, R.Top, R.Right+1, R.Bottom+1) ;
end ;

procedure TFSaisieRbp.PostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
var bGrise : Boolean ;
begin
if (ARow<GS.FixedRows) or (ARow>nbLignes+1) then Exit ;
bGrise:=FALSE ;
if (ACol=SR_DEBIT)  and (GS.Cells[SR_CREDIT, ARow]<>'') then bGrise:=TRUE ;
if (ACol=SR_CREDIT) and (GS.Cells[SR_DEBIT, ARow]<>'') then bGrise:=TRUE ;
if bGrise then SetGridGrise(ACol, ARow, Canvas);
end ;

procedure TFSaisieRbp.UpdateScrollBars ;
begin
//if nbLignes>GS.VisibleRowCount-1 then GS.ScrollBars:=ssBoth else GS.ScrollBars:=ssHorizontal ;
end ;

procedure TFSaisieRbp.SetGridOn(bMove : Boolean) ;
begin
if bMove then begin GS.Enabled:=TRUE ; GS.Row:=GS.FixedRows+nbLignes ; end ;
if (GS.Enabled) and (GS.Visible) and (not bClosing) then
  begin
  if GS.Row<=GS.FixedRows then GS.Row:=GS.FixedRows+1 ;
  if GS.Row>GS.FixedRows+nbLignes then GS.Row:=GS.FixedRows+nbLignes ;
  if GS.Col<SR_FIRST then GS.Col:=SR_FIRST ;
  if GS.Col>SR_LAST then GS.Col:=SR_LAST ;
  GS.SetFocus ;
  end ;
if bModeRO then begin if bMove then GS.Row:=GS.Row-1 ; SetGridRO ; end ;
end ;

procedure TFSaisieRbp.SetGridRO ;
begin
GS.CacheEdit ;
GS.Options:=GS.Options-[GoEditing, GoTabs, GoAlwaysShowEditor] ;
GS.Options:=GS.Options+[GoRowSelect] ;
end ;

procedure TFSaisieRbp.SetGridOptions(Row : LongInt) ;
begin
if Action=taModif then
  begin
  if IsRowCanEdit(Row) then
    begin
    GS.Options:=GS.Options-[GoRowSelect] ;
    GS.Options:=GS.Options+[GoEditing,GoTabs,GoAlwaysShowEditor] ;
    GS.MontreEdit ;
    GS.Invalidate ;
    end else
    begin
    GS.CacheEdit ;
    GS.Options:=GS.Options-[GoEditing,GoTabs,GoAlwaysShowEditor] ;
    GS.Options:=GS.Options+[GoRowSelect] ;
    end ;
  end ;
end ;

procedure TFSaisieRbp.CreateRow(Row : LongInt) ;
var bInsert : Boolean ; RowRef : LongInt ; k : Integer ;
begin
RowRef:=Row-1 ; bInsert:=FALSE ;
if Row<=nbLignes then
   begin bInsert:=TRUE ; GS.InsertRow(Row) ; if Row=1 then RowRef:=Row+1 ; end ;
Inc(nbLignes) ; UpdateScrollBars ;
if GS.RowCount-2<nbLignes then GS.RowCount:=GS.RowCount+1 ;
// Etablissement par défaut
GS.Cells[SR_ETABL,Row]:=VH^.EtablisDefaut ;
// Régime par défaut
k:=CRB_REGIMETVA.Values.IndexOf(VH^.RegimeDefaut) ;
if k>=0 then GS.Cells[SR_REGIME, Row]:=CRB_REGIMETVA.Items[k] ;
// Ligne créée en saisie
GS.Cells[SR_NEW,Row]:='O' ;
// Numérotation de la ligne
SetOrdre ;
// Dernier jour du mois par défaut
if nbLignes>1 then GS.Cells[SR_JOUR,Row]:=GS.Cells[SR_JOUR, RowRef]
              else GS.Cells[SR_JOUR,Row]:=IntToStr(memOptions.MaxJour) ;
// Création de l'objet correspondant
ObjReleve.CreateRow(nil, Row-GS.FixedRows) ;
end ;

function TFSaisieRbp.DeleteRow(Row : LongInt) : Integer ;
var ACol : Integer ; bValid : Boolean ;
begin
Result:=0 ;
bValid:=IsRowValid(Row, ACol, FALSE) ;
if (nbLignes=1) and (not bValid) then Exit ;
ObjReleve.DeleteRow(Row-GS.FixedRows) ;
GS.DeleteRow(Row) ; Dec(nbLignes) ; UpdateScrollBars ;
if GS.RowCount>MAX_ROW then GS.RowCount:=GS.RowCount-1 ;
SetOrdre ; CalculTotalReleve ;
if (Row=GS.FixedRows) and (nbLignes=0) then NextRow(Row) ;
if (bValid) and (nbLignes>1) then Result:=-1 ;
EnableButtons ;
end ;

function TFSaisieRbp.NextRow(Row : LongInt) : Boolean ;
begin
if Row>nbLignes+1 then begin Result:=FALSE ; Exit ; end ;
if Row>nbLignes then begin CreateRow(Row) ; SetGridOptions(Row) ; end ;
Result:=TRUE ;
end ;

// En saisie, on considère que la ligne est valide si le compte général est renseigné
function TFSaisieRbp.IsRowValid(Row : LongInt; var ACol : Integer; bTest: Boolean) : Boolean ;
var Ecr : TZF ;
begin
Result:=TRUE ;
if ObjReleve=nil then Exit ;
Ecr:=ObjReleve.GetRow(Row-GS.FixedRows) ;
if Ecr=nil then begin Result:=FALSE ; Exit ; end ;
if (GS.Cells[SR_DEBIT, Row]='') and (GS.Cells[SR_CREDIT, Row]='') then
   begin
   if (not bTest) and (bGuideRun) then
     begin PrintMessageRC(RC_BADGUIDE) ; GuideStop ; GS.Invalidate ; end ;
   Result:=FALSE ; ACol:=SR_DEBIT ; Exit ;
   end ;
end ;

function TFSaisieRbp.IsRowCanEdit(Row : LongInt) : Boolean ;
var Ecr : TZF ;
begin
Result:=TRUE ; if ObjReleve=nil then Exit ;
Ecr:=ObjReleve.GetRow(Row-GS.FixedRows) ; if Ecr=nil then Exit ;
if (Ecr.GetValue('CRB_NUMEROPIECE')>0) then Result:=FALSE
end ;

function TFSaisieRbp.IsRowOk(Row : LongInt) : Boolean ;
begin
Result:=TRUE ;
if (GS.Cells[SR_DEBIT, Row]='') and (GS.Cells[SR_CREDIT, Row]='') then Result:=FALSE ;
if GS.Cells[SR_JOUR, Row]='' then Result:=FALSE ;
if GS.Cells[SR_IMP, Row]='' then Result:=FALSE ;
end ;

// Permet de tester si une cellule peut passer en mode saisie
function TFSaisieRbp.IsCellOK(Col, Row : LongInt) : Boolean ;
begin
Result:=FALSE ;
//if (Col<>SR_PB) and (Col<>SR_PIECE) and (GS.Cells[SR_GEN, Row]<>'') then Exit ;
if (Col=SR_DEBIT)  and (GS.Cells[SR_CREDIT, Row]<>'') then Exit ;
if (Col=SR_CREDIT) and (GS.Cells[SR_DEBIT,  Row]<>'') then Exit ;
Result:=TRUE ;
end ;

//=======================================================
//================= Evénements du Grid ==================
//=======================================================
procedure TFSaisieRbp.GSElipsisClick(Sender: TObject);
var Table, Critere, Val, NumCompte, StTable, StCode, StWhere, StPrefixe, StLib : string ;
begin
if bClosing then Exit ;
if Screen.ActiveControl<>GS then GS.SetFocus;
Critere:='';
// Comptes généraux
if GS.Col=SR_GENERAL then
   LookupList(GS,TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE',Critere,'G_GENERAL',FALSE, 1) ;
// Guide d'imputation
if GS.Col=SR_IMP then
   begin
   Val:=GS.Cells[GS.Col, GS.Row] ; GS.Cells[GS.Col, GS.Row]:='' ;
   NumCompte:=GS.Cells[SR_GENERAL, GS.Row] ;
   if NumCompte='' then
      begin
      Critere:='GU_TYPE="RLV"' ; Table:='GUIDE' ;
      LookupList(GS,TraduireMemoire('Guide'),Table,'GU_GUIDE','GU_LIBELLE',Critere,'GU_GUIDE',FALSE,0) ;
      end else
      begin
      Critere:='GU_TYPE="RLV" AND EG_GENERAL="'+NumCompte+'"' ;
      Table:='GUIDE LEFT JOIN ECRGUI ON GU_GUIDE=EG_GUIDE' ;
      LookupList(GS,TraduireMemoire('Guide'),Table,'GU_GUIDE','GU_LIBELLE',Critere,'GU_GUIDE',FALSE,0) ;
      end ;
//SELECT EG_GUIDE, EG_GENERAL FROM ECRGUI WHERE EG_GENERAL="60700000" AND EG_TYPE="RLV"
   if GS.Cells[GS.Col, GS.Row]='' then GS.Cells[GS.Col, GS.Row]:=Val ;
   end ;
if GS.Col=SR_REGIME then
   begin
   Val:=GS.Cells[GS.Col, GS.Row] ; GS.Cells[GS.Col, GS.Row]:='' ;
   GetCorrespType(CRB_REGIMETVA.DataType, StTable, StCode, StWhere, StPrefixe, StLib) ;
   LookupList(GS,TraduireMemoire('Régime'),StTable,StLib,StCode,StWhere,'',FALSE,0) ;
   if GS.Cells[GS.Col, GS.Row]='' then GS.Cells[GS.Col, GS.Row]:=Val ;
   end ;
if GS.Col=SR_TVA then
   begin
   Val:=GS.Cells[GS.Col, GS.Row] ; GS.Cells[GS.Col, GS.Row]:='' ;
   GetCorrespType(CRB_CODETVA.DataType, StTable, StCode, StWhere, StPrefixe, StLib) ;
   LookupList(GS,TraduireMemoire('Régime'),StTable,StLib,StCode,StWhere,'',FALSE,0) ;
   if GS.Cells[GS.Col, GS.Row]='' then GS.Cells[GS.Col, GS.Row]:=Val ;
   end ;
end ;

procedure TFSaisieRbp.GSEnter(Sender: TObject);
begin
if bFirst then
   begin
   if not bModeRO then NextRow(GS.Row) ;
   SoldeProgressif(GS.Row) ;
   EnableButtons ;
   if not bModeRO then
     begin
     PostMessage(GS.Handle, WM_KEYDOWN, VK_TAB,  0) ;
     PostMessage(GS.Handle, WM_KEYDOWN, VK_LEFT, 0) ;
     end ;
   end ;
bFirst:=FALSE ;
end;

procedure TFSaisieRbp.GSCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var lc, lr, sens : Integer ;
begin
if ARow<>GS.Row then SetGridOptions(GS.Row) ;
// Bouton Elipsis
GS.ElipsisButton:=FALSE ;
lc:=GS.Col ; lr:=GS.Row ; sens:=GetGridSens(ACol, ARow) ;
// Vérifier que l'on est pas en dehors des zones de saisies autorisées
if lr>nbLignes then begin InfosPied(ARow) ; Cancel:=TRUE ; Exit ; end ;
// En consultation, c'est fini !
if bModeRO then Exit ;
// En saisie guidée ou saisie non libre, interdit de sortir de la pièce
if bGuideRun then  begin Cancel:=TRUE ; Exit ; end ;
// Si le mouvement est généré, pas de modification autorisée
(*
if (GS.Cells[SR_GEN, lr]<>'') and (lc<>SR_PB) and (lc<>SR_PIECE) then
   begin
   if sens=1 then begin if lc>SR_PB then ACol:=SR_PIECE else ACol:=SR_PB ; end
             else begin if lc>SR_PB then ACol:=SR_PB else ACol:=SR_PIECE ; end ;
   ARow:=lr ;
   Cancel:=TRUE ;
   Exit ;
   end ;
*)
// Interdir la saisie d'un montant au débit si montant au crédit
if (lc=SR_DEBIT) and (GS.Cells[SR_CREDIT,lr]<>'') then
  begin
  if sens=1 then begin ACol:=SR_CREDIT ; ARow:=lr ; end
            else begin ACol:=SR_PB ;     ARow:=lr ; end ;
  Cancel:=TRUE ; Exit ;
  end ;
// Interdir la saisie d'un montant au crédit si montant au débit
if (lc=SR_CREDIT) and (GS.Cells[SR_DEBIT,lr]<>'') then
  begin
  if sens=1 then begin ACol:=SR_REGIME ; ARow:=lr ; end
            else begin Acol:=SR_DEBIT ;  Arow:=lr ; end ;
  Cancel:=TRUE ; Exit ;
  end ;
// Bouton Elipsis
GS.ElipsisButton:=(lc=SR_IMP) or (lc=SR_REGIME) or (lc=SR_TVA) ;
if (lc=SR_IMP)     then GS.ElipsisHint:=GetMessageRC(RC_SELGUIDE) else
  if (lc=SR_REGIME)  then GS.ElipsisHint:=GetMessageRC(RC_SELREGIME) else
    if (lc=SR_TVA) then GS.ElipsisHint:=GetMessageRC(RC_SELCODETVA) ;
// Ouvrir le lookup automatiquement !
if (ACol=SR_GENERAL) and (lc=SR_IMP) and
   (GS.Cells[SR_GENERAL, lr]<>'') and (GS.Cells[SR_IMP, lr]='') then
  PostMessage(GS.Handle, WM_KEYDOWN, VK_ADD, 0);    
// Guide
if bGuideRun then
   if GetGuideValue(lc, lr) then PostMessage(GS.Handle, WM_KEYDOWN, VK_TAB, 0) ;
end;

procedure TFSaisieRbp.GSCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var NumCompte : string ; k : Integer ;
begin
// En consultation, rien à faire ici !
if bModeRO then Exit ;
// Validité du jour
if (ACol=SR_JOUR) then
   begin
   GS.Cells[ACol, ARow]:=StrFMontant(Abs(Valeur(GS.Cells[ACol, ARow])),15,0,'',TRUE) ;
   if (Valeur(GS.Cells[ACol, ARow])>memOptions.MaxJour) or (Valeur(GS.Cells[ACol, ARow])<=0) then
      GS.Cells[ACol, ARow]:=IntToStr(memOptions.MaxJour) ;
   end ;
// Validité du compte général
if (ACol=SR_GENERAL) then
   begin
   NumCompte:=GS.Cells[ACol, ARow] ;
   if (NumCompte<>'') then
     begin
     k:=Comptes.GetCompte(NumCompte) ;
     if k<0 then
       begin
       LookupList(GS,TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE','','G_GENERAL',FALSE, 1) ;
       if GS.Cells[ACol, ARow]='' then GS.Cells[ACol, ARow]:=NumCompte ;
       end else GS.Cells[ACol, ARow]:=Comptes.GetValue('G_GENERAL', k) ;
     end ;
   end ;
// Formatage du montant
if (ACol=SR_DEBIT) or (ACol=SR_CREDIT) then
   begin
   FormatMontant(GS, ACol, ARow, memSaisie.Decimale) ;
   CalculTotalReleve ;
   end ;
EnableButtons ;   
end;

procedure TFSaisieRbp.GSRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
var R : RMVT ; ACol : Integer ; CodeGen : string ;
begin
// En consultation, rien à faire ici !
if bModeRO then Exit ;
if Ou>nbLignes then Exit ;
bStopDevise:=FALSE ;
if not IsRowValid(Ou, ACol, FALSE) then
   begin
   if (Ou=nbLignes) and (GetGridSens(GS.Col, Ou)=-1) and (Screen.ActiveControl=GS)
     then begin if IsCellOK(GS.Col, Ou-1) then DeleteRow(Ou) else Cancel:=TRUE ; end
     else Cancel:=TRUE ;
//   if (Ou=nbLignes) and (GetGridSens(GS.Col, Ou)=-1) then DeleteRow(Ou)
//                                                     else Cancel:=TRUE ;
   end else
   begin
   SetRowTZF(Ou) ;
   if (GS.Cells[SR_IMP, Ou]<>'') and (GS.Cells[SR_GEN, Ou]='') then
      begin
      CodeGen:=GenGuide(Ou) ;
      FillChar(R,Sizeof(R),#0) ;  R.DateC:=GetRowDate(Ou) ;
      R.Exo:=QuelExoDT(R.DateC) ; R.Simul:='N' ;
      R.LeGuide:=CodeGen ;        R.TypeGuide:='NOR' ;
      R.ANouveau:=False ;         R.SaisieGuidee:=True ;
      LanceSaisie(nil, taCreat, R) ;
      if (R.LastNumCreat<>-1) then GS.Cells[SR_GEN, Ou]:=IntToStr(R.LastNumCreat) ;
      DelGuide(CodeGen) ;
      end ;
   if GS.Row>Ou then Cancel:=not NextRow(GS.Row) ;
   end ;
SetGridOptions(Ou) ;
end ;

procedure TFSaisieRbp.GSRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
var Ecr : TZF ;
begin
SoldeProgressif(Ou) ;
EnableButtons ;
InfosPied(-1) ;
if ObjReleve<>nil then Ecr:=ObjReleve.GetRow(GS.Row-GS.FixedRows) else Ecr:=nil ;
if (Ecr<>nil) and (Ecr.GetValue('GUIDE')<>'X') then GuideStop ;
SetGridOptions(Ou) ;
GS.Invalidate ;
end;

procedure TFSaisieRbp.GSExit(Sender: TObject) ;
var Cancel, Chg : Boolean ; ACol, ARow : LongInt ;
begin
// En consultation, rien à faire ici !
if bModeRO then Exit ;
// Sinon...
ACol:=GS.Col ; ARow:=GS.Row ; Cancel:=FALSE ; Chg:=FALSE ;
GSCellExit(GS, ACol, ARow, Cancel) ; GSRowExit(GS, ARow, Cancel, Chg) ;
GS.Col:=ACol ; GS.Row:=ARow ;
end;

procedure TFSaisieRbp.GSKeyPress(Sender: TObject; var Key: Char);
begin
if not GS.SynEnabled then Key:=#0 else
   begin
   if (Key='+') and ((GS.Col=SR_IMP) or (GS.Col=SR_REGIME) or (GS.Col=SR_TVA)) then Key:=#0 ;
   if (Key=' ') and ((GS.Col=SR_REGIME) or (GS.Col=SR_TVA)) then Key:=#0 ;
   if (Key<>' ') and ((GS.Col=SR_REGIME) or (GS.Col=SR_TVA)) then Key:=#0 ;
   end ;
end;

procedure TFSaisieRbp.GSMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin GridX:=X ; GridY:=Y ; end ;

procedure TFSaisieRbp.PFENMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var r : TRect ;
begin
if (not GS.Enabled) and (not bClosing) then
  begin
  r:=Rect(GS.Left, GS.Top, GS.Left+GS.Width, GS.Top+GS.Height) ;
  if PtInRect(r, Point(X, Y)) then
    begin
    if Jal=nil then
    begin
    if E_JOURNAL.Value='' then
      begin if Jal<>nil then Jal.Free ; Jal:=nil ; Exit ; end ;
    // Alimenter le record RJal
    ChargeJal(E_JOURNAL.Value) ;
    end ;
    if not ReadReleve then begin E_JOURNAL.SetFocus ; Exit ; end ;
    CalculTotalReleve ;
    // Activer le Grid
    SetGridOn ;
    InfosPied(-1) ;
    EnableButtons ;
    end ;
  if bModeRO then PostMessage(GS.Handle, WM_LBUTTONUP, 0, 0) ;
  end ;
end;

//=======================================================
//=============== Evénements des boutons ================
//=======================================================
procedure TFSaisieRbp.BChercherClick(Sender: TObject);
begin SearchClick ; end ;

procedure TFSaisieRbp.BValideClick(Sender: TObject);
begin ValClick ; end ;

procedure TFSaisieRbp.BDeleteRowClick(Sender: TObject);
begin DelClick ; end ;

procedure TFSaisieRbp.BCreateRowClick(Sender: TObject);
begin InsClick ; end ;

procedure TFSaisieRbp.BGuideClick(Sender: TObject);
begin GuiClick ; end ;

procedure TFSaisieRbp.SearchClick ;
begin FindFirst:=TRUE ; FindSais.Execute ;  end ;

procedure TFSaisieRbp.BEditClick(Sender: TObject);
begin EditClick ; end ;

procedure TFSaisieRbp.BToolsClick(Sender: TObject);
begin ToolsClick ; end ;

procedure TFSaisieRbp.FindSaisFind(Sender: TObject);
var bCancel : Boolean ;
begin
Rechercher(GS, FindSais, FindFirst) ;
GSRowEnter(nil, GS.Row, bCancel, FALSE);
end ;

procedure TFSaisieRbp.ValClick ;
begin
if Screen.ActiveControl<>GS then Exit ;
CalculTotalReleve ;
EnableButtons ;
if not BValide.Enabled then Exit ;
if ValideReleve then InitReleve ;
if not bClosing then E_JOURNAL.SetFocus ;
end ;

function TFSaisieRbp.ValideReleve : Boolean ;
//var Start : LongInt ;
begin
Result:=FALSE ;
if nbLignes<=0 then Exit ;
//start := GetTickCount;
if Transactions(WriteReleve, 5)<>oeOk then
   begin PrintMessageRC(RC_BADWRITE) ; FiniMove ; end ;
//InitReleve ;
//ShowMessage(Format('%f', [(GetTickCount-Start)/1000])) ;
Result:=TRUE ;
end ;

procedure TFSaisieRbp.DelClick ;
var lr : LongInt ; Decal : Integer ;
begin
if BDeleteRow.Enabled=FALSE then Exit ;
lr:=GS.Row ; Decal:=DeleteRow(lr) ;
if lr>nbLignes then GS.Row:=nbLignes else GS.Row:=lr+Decal ;
if GS.Row<>GS.FixedRows then GS.ElipsisButton:=FALSE ;
EnableButtons ;
SetGridOptions(GS.Row) ;
GS.Invalidate ;
(* ANCIEN CODE
lr:=GS.Row ; ltop:=GS.TopRow ; Decal:=DeleteRow(lr) ;
if lr>nbLignes then GS.Row:=nbLignes else GS.Row:=lr+Decal ;
GS.TopRow:=ltop ; GS.ElipsisButton:=FALSE ;
*)
end ;

procedure TFSaisieRbp.InsClick ;
var lr : LongInt ;
begin
//lr:=GS.Row ; CreateRow(lr) ; GS.Row:=lr ;
if BCreateRow.Enabled=FALSE then Exit ;
lr:=GS.Row+1 ; CreateRow(lr) ; GS.Row:=lr ;
EnableButtons ;
SetGridOptions(GS.Row) ;
GS.Invalidate ;
end ;

procedure TFSaisieRbp.GuiClick ;
var  R : TRect ; Critere : string ;
begin
R:=GS.CellRect(GS.Col, GS.Row) ;
SELGUIDE.Text:='' ; SELGUIDE.Left:=R.Left ; SELGUIDE.Top:=R.Bottom+10 ;
Critere:='GU_TYPE="RLV" AND GU_JOURNAL="'+Jal.GetValue('J_JOURNAL')+'"' ;
LookupList(SELGUIDE,TraduireMemoire('Guide'),'GUIDE','GU_GUIDE','GU_LIBELLE',Critere,'GU_GUIDE',FALSE,0) ;
//SELGUIDE.Text:='%$/PFU' ;
//repeat Application.ProcessMessages until SELGUIDE.Text<>'%$/PFU' ;
if SELGUIDE.Text<>'' then GuideRun ;
end ;

procedure TFSaisieRbp.GuideRun ;
var  ObjGuide : TZGuide ; Ecr : TZF ; i : Integer ;
begin
ObjGuide:=TZGuide.Create('GuideEcr', SELGUIDE.Text) ;
ObjGuide.Load ;
for i:=0 to ObjGuide.Count-1 do
    begin
    if i>0 then CreateRow(GS.Row+i) ;
    Ecr:=ObjReleve.GetRow(GS.Row-GS.FixedRows+i) ; if Ecr=nil then break ;
    Ecr.PutValue('GUIDE',         'X') ;
    Ecr.PutValue('EG_NUMLIGNE',   ObjGuide.GetValue('EG_NUMLIGNE',   i, -1)) ;
    Ecr.PutValue('EG_GENERAL',    ObjGuide.GetValue('EG_GENERAL',    i, -1)) ;
    Ecr.PutValue('EG_AUXILIAIRE', ObjGuide.GetValue('EG_AUXILIAIRE', i, -1)) ;
    Ecr.PutValue('EG_REFINTERNE', ObjGuide.GetValue('EG_REFINTERNE', i, -1)) ;
    Ecr.PutValue('EG_LIBELLE',    ObjGuide.GetValue('EG_LIBELLE',    i, -1)) ;
    Ecr.PutValue('EG_DEBITDEV',   ObjGuide.GetValue('EG_DEBITDEV',   i, -1)) ;
    Ecr.PutValue('EG_CREDITDEV',  ObjGuide.GetValue('EG_CREDITDEV',  i, -1)) ;
    Ecr.PutValue('EG_ARRET',      ObjGuide.GetValue('EG_ARRET',      i, -1)) ;
    bGuideRun:=TRUE ;
    end ;
ObjGuide.Free ;
FlashGuide.Visible:=TRUE ; FlashGuide.Flashing:=TRUE ;
GS.Invalidate ;
end ;

procedure TFSaisieRbp.GuideStop ;
begin
bGuideRun:=FALSE ; FlashGuide.Flashing:=FALSE ; FlashGuide.Visible:=FALSE ;
end ;

procedure TFSaisieRbp.EditClick ;
var R : RMVT ; QEcr : TQuery ; Where : string ;
begin
Where:=' E_JOURNAL="'+Jal.GetValue('J_JOURNAL')+'"'
      +' AND E_EXERCICE="'+QuelExoDT(GetRowDate(GS.Row))+'"'
      +' AND E_DATECOMPTABLE="'+USDateTime(GetRowDate(GS.Row))+'"'
      +' AND E_NUMEROPIECE="'+GS.Cells[SR_GEN, GS.Row]+'"'
      +' AND E_SOCIETE="'+V_PGI.CodeSociete+'"'
      +' AND E_ETABLISSEMENT="'+VH^.EtablisDefaut+'"' ;
QEcr:=OpenSQL('SELECT E_NATUREPIECE FROM ECRITURE WHERE'+Where, TRUE) ;
if QEcr.EOF then begin Ferme(QEcr) ; Exit ; end ;
FillChar(R,Sizeof(R),#0) ;  R.DateC:=GetRowDate(GS.Row) ;
R.Exo:=QuelExoDT(R.DateC) ; R.Simul:='N' ;
R.Jal:=Jal.GetValue('J_JOURNAL') ;
R.Nature:=QEcr.FindField('E_NATUREPIECE').AsString ;
R.Num:=StrToInt(GS.Cells[SR_GEN, GS.Row]) ;
R.Etabl:=VH^.EtablisDefaut ;
Ferme(QEcr) ;
LanceSaisie(nil, taModif, R) ;
end ;

procedure TFSaisieRbp.ToolsClick ;
begin
(*
if SaisieZParams(memParams) then
   begin
   if memParams.bPiece then GS.ColWidths[SF_PIECE]:=20
                       else GS.ColWidths[SF_PIECE]:=0 ;
   SoldeProgressif(GS.Row) ;
   end ;
*)
end ;

//=======================================================
//============= Evénements des boutons Zoom==============
//=======================================================
procedure TFSaisieRbp.POPSPopup(Sender: TObject) ;
begin InitPopUp(Self) ; end ;

end.

