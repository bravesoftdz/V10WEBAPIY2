unit SaisTier ;

//=======================================================
//======== +++++++ ===   LEGENDE    === +++++++ =========
//=======================================================
// ( ) Reste à faire
// (o) Implémenté mais non testé
// (x) Implémenté et testé
//=======================================================
//======== +++++++ === +++ TODO +++ === +++++++ =========
//=======================================================

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  StdCtrls,
  Grids,
  Hctrls,
  Mask,
  ExtCtrls,
  ComCtrls,
  Buttons,
  Hspliter,
  Ent1,
  HCompte,
  HEnt1,
  Hmsgbox,
  HQry,
  Menus,
  HFLabel,
  HSysMenu,
  HTB97,
  HPanel,
{$IFDEF EAGLCLIENT}
{$ELSE}
  db,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
{$IFDEF VER150}
   Variants,
{$ENDIF}
  CPTiers_TOM,   // FicheTiers
  SaisUtil,      // RDevise
  SaisComm,      // ZeroBlanc, FormatMontant
  UiUtil,        // FindInsidePanel
  TZ,            // TZF
  ZPlan, TntStdCtrls, TntExtCtrls, TntGrids
  ; { Mac3 }

procedure SaisieBalAux(EcrRef : TZF; var PlanTiers : TZPlan; NatureGen : string; memSaisie : RDevise; SoldeDAvant, SoldeCAvant: Double) ;

const SB_ETABL  : integer = 0 ;
      SB_NEW    : integer = 1 ;
      SB_COLL   : integer = 2 ;
      SB_TYPE   : integer = 3 ;
      SB_LETTRE : integer = 4 ;
      SB_MODER  : integer = 5 ;
      SB_AUX    : integer = 6 ;
      SB_LIB    : integer = 7 ;
      SB_DEBIT  : integer = 8 ;
      SB_CREDIT : integer = 9 ;
      SB_FIRST  : integer = 6 ;
      SB_LAST   : integer = 9 ;

const RC_BALNONSOLDE      =  0 ;
      RC_BADWRITE         =  1 ;
      RC_BLOQUERESEAU     =  2 ;
      RC_CUMULCLASSE      =  3 ;
      RC_CHARGEPLAN       =  4 ;
      RC_BALANCEAU        =  5 ;
      RC_PERTE            =  6 ;
      RC_BENEFICE         =  7 ;
      RC_GENINTERDIT      =  8 ;
      RC_AUXINTERDIT      =  9 ;
      RC_EXO              = 10 ;
      RC_BALN1CEXISTE     = 11 ;
      RC_NOEXO            = 12 ;
      RC_ANMODIFIE        = 13 ;
      RC_BALN0AEXISTE     = 14 ;
      RC_BALN1AEXISTE     = 15 ;
      RC_ECRSAISIES       = 16 ;
      RC_DATEINCORRECTE   = 17 ;
      RC_AVECAUX          = 18 ;
      RC_CHOIXBALANCE     = 19 ;
      RC_NOBALANCE        = 20 ;
      RC_DELBALANCE       = 21 ;
      RC_NOGENBALANCE     = 22 ;
      RC_ATTEND           = 23 ;

const MAX_ROW             = 500 ;

type
  TFSaisieAux = class(TForm)
    GS              : THGrid;
    PPied: THPanel;
    H_SOLDE         : THLabel;
    SA_TOTALCREDIT  : THNumEdit;
    SA_SOLDE        : THNumEdit;
    HBalance: THMsgBox;
    FindSais        : TFindDialog;
    Bevel2          : TBevel;
    Bevel3          : TBevel;
    Bevel4          : TBevel;
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
    BSolde: TToolbarButton97;
    BChercher: TToolbarButton97;
    ToolbarSep971: TToolbarSep97;
    Sep97: TToolbarSep97;
    HConf: TToolbarButton97;
    ISigneEuro: TImage;
    LSA_SOLDE: THLabel;
    LSA_TOTALCREDIT: THLabel;
    LSA_TOTALDEBIT: THLabel;
    SA_TOTALDEBIT: THNumEdit;
    HLabel2: THLabel;
    FLASHDEVISE: TFlashingLabel;
    BNew: TToolbarButton97;
    LSA_BALANCE: THLabel;
    HLabel1: THLabel;
    Bevel7: TBevel;
    SA_BALANCE: THNumEdit;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GereAffSolde(Sender: TObject);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSEnter(Sender: TObject);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSKeyPress(Sender: TObject; var Key: Char);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure BValideClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BSoldeClick(Sender: TObject);
    procedure GSMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GSExit(Sender: TObject);
    procedure FindSaisFind(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure BNewClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
  private
    EcrRef       : TZF ;
    EcrAux       : TZF ;
    PlanTiers    : TZPlan ;
    memPivot     : RDevise ;
//    memOptions   : ROpt ;
//    memParams    : RPar ;
    bFirst       : Boolean ;
    GridX, GridY : Integer ;
    bReading     : Boolean ;
    bClosing     : Boolean ;
    bSave        : Boolean ;
    NatureGen    : string ;
    SoldeDAvant  : Double ;
    SoldeCAvant  : Double ;
    // Fonctions ressource
    function  GetMessageRC(MessageID : integer) : string ;
    function  PrintMessageRC(MessageID : Integer) : Integer ;
    // Fonctions utilitaires
    procedure InitBalance ;
    procedure PosLesSoldes ;
    procedure InitPivot ;
    procedure InitSaisie ;
    function  CanCloseSaisie : Boolean ;
    procedure CloseSaisie ;
    procedure InitPied ;
    procedure InitGrid ;
    procedure InfosPied ;
    procedure EnableButtons ;
    // Fontions TZF
    function  SetRowTZF(Row : LongInt) : TZF ;
    function  GetRowTZF(Row, RowCur : LongInt) : Boolean ;
    function  GetRow(Row : LongInt) : TZF ;
    function  GetRowRef(Row : LongInt) : TZF ;
    // Calculs
    procedure CalculSoldeBalance ;
    function  IsSoldeBalance : boolean ;
    // Fonctions Grid
    procedure UpdateScrollBars ;
    function  CreateRow(Row : LongInt) : TZF ;
    procedure DeleteRow(Row : LongInt) ;
    function  NextRow(Row : LongInt) : Boolean ;
    function  IsRowValid(Row : integer; var ACol : integer) : Boolean ;
    function  GetGridSens(ACol, ARow : Integer) : Integer ;
    procedure GetCellCanvas(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    procedure PostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    procedure SetGridGrise(ACol, ARow : Integer ; Canvas : TCanvas) ;
    //procedure SetGridSep(ACol, ARow : Integer ; Canvas : TCanvas ; bHaut : Boolean) ;
    //procedure SetGridBold(ACol, ARow : Integer ; Canvas : TCanvas) ;
    //procedure SetGridEditing(bOn : Boolean) ;
    function  IsRowOk(Row : LongInt) : Boolean ;
    // SQL
    procedure ReadBalance ;
    procedure PutGrid(Row : LongInt) ;
    function  GetNature(Nature : string) : Boolean ;
    // Click boutons
    procedure SearchClick ;
    procedure ValClick ;
    procedure DelClick ;
    procedure NewClick ;
    procedure ByeClick ;
    procedure SoldeClick(Row : longint) ;
  public
    nbLignes : LongInt ;
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  HStatus;


//=======================================================
//======== Point d'entrée dans la saisie balance ========
//=======================================================
procedure SaisieBalAux(EcrRef : TZF; var PlanTiers : TZPlan; NatureGen : string; memSaisie : RDEVISE; SoldeDAvant, SoldeCAvant: Double) ;
var Balance : TFSaisieAux ; PP : THPanel ;
begin
if PlanTiers.Count=0 then Exit ;
Balance:=TFSaisieAux.Create(Application) ;
Balance.EcrRef:=EcrRef ;
Balance.PlanTiers:=PlanTiers ;
Balance.NatureGen:=NatureGen ;
Balance.memPivot:=memSaisie ;
Balance.SoldeDAvant:=SoldeDAvant ;
Balance.SoldeCAvant:=SoldeCAvant ;
PP:=FindInsidePanel ;
if PP=nil then
   begin
   try
     Balance.ShowModal ;
     PlanTiers:=Balance.PlanTiers ;
   finally
     Balance.Free ;
   end ;
  Screen.Cursor:=SyncrDefault ;
  end else
  begin
  InitInside(Balance, PP) ;
  Balance.Show ;
  end ;
end ;

//=======================================================
//================= Fonctions Ressource =================
//=======================================================
function TFSaisieAux.GetMessageRC(MessageID : Integer) : string ;
begin
Result:=HBalance.Mess[MessageID] ;
end ;

function TFSaisieAux.PrintMessageRC(MessageID : Integer) : Integer ;
begin
Result:=HBalance.Execute(MessageID, Caption, '') ;
end ;

//=======================================================
//================ Evénements de la Form ================
//=======================================================
procedure TFSaisieAux.FormCreate(Sender: TObject);
begin
bClosing:=FALSE ; bReading:=FALSE ; bSave:=TRUE ;
InitGrid ;
end;

procedure TFSaisieAux.FormShow(Sender: TObject);
begin
InitPivot ;
PosLesSoldes ;
LookLesDocks(Self) ;
// Copie de EcrRef sans Détail
EcrAux:=TZF.Create('HISTOBAL', nil, -1) ;
EcrAux.Dupliquer(EcrRef, FALSE, TRUE) ;
ReadBalance ;
GS.Enabled:=TRUE ; GS.SetFocus ; GSEnter(nil) ;
end;

procedure TFSaisieAux.FormClose(Sender: TObject;
  var Action: TCloseAction);
var Ecr : TZF ; i : Integer ;  
begin
if bReading then begin Action:=caNone ; Exit ; end ;
CloseSaisie ;
// CleanUp !!!
if bSave then
  begin
  while EcrRef.Detail.Count>0 do EcrRef.Detail[0].Free ;
  for i:=0 to EcrAux.Detail.Count-1 do
    begin
    if not IsRowOk(i+GS.FixedRows) then Continue ;
    Ecr:=TZF.Create('HISTOBAL', EcrRef, -1) ;
    Ecr.Assign(EcrAux.Detail[i]) ;
    end ;
  end ;
EcrAux.Free ;
if Parent is THPanel then Action:=caFree ;
if FindSais.Handle<>0 then FindSais.CloseDialog ;
bClosing:=TRUE ;
end;

procedure TFSaisieAux.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
CanClose:=CanCloseSaisie ;
//if CanClose then CloseSaisie ;
end;

procedure TFSaisieAux.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var Vide : Boolean ;
begin
if not GS.SynEnabled then begin Key:=0 ; Exit ; end ;
if not (Screen.ActiveControl=GS) then Exit ;
Vide:=(Shift=[]) ;
case Key of
  VK_RETURN : if (Vide) then Key:=VK_TAB ;
     VK_END : if Shift=[ssCtrl] then
                 begin
                 Key:=0 ; 
                 GS.Row:=nbLignes ; GS.Col:=SB_CREDIT ;
                 PostMessage(GS.Handle, WM_KEYDOWN, VK_TAB,  0) ;
                 PostMessage(GS.Handle, WM_KEYDOWN, VK_LEFT, 0) ;
                 end ;
  VK_INSERT : if (Vide) then begin Key:=0; (*InsClick ;*) end ;
  VK_DELETE : if Shift=[ssCtrl] then begin Key:=0 ; DelClick ; end ;
      VK_F6 : if (Vide) then begin Key:=0 ; SoldeClick(-1) ; end ;
  VK_ESCAPE,
     VK_F10 : if (Vide) then
                 begin
                 Key:=0 ;
                 if FindSais.Handle<>0 then begin FindSais.CloseDialog ; Exit ; end ;
                 if bReading then Exit ; ValClick ;
                 end ;
  end ;
end;

//=======================================================
//================ Evénements des contrôles =============
//=======================================================
procedure TFSaisieAux.GereAffSolde(Sender: TObject);
var Nam : string ; c : THLabel ;
begin
Nam:=THNumEdit(Sender).Name ; Nam:='L'+Nam ;
c:=THLabel(FindComponent(Nam)) ;
if c<>nil then c.Caption:=THNumEdit(Sender).Text ;
end;

//=======================================================
//================ Fonctions utilitaires ================
//=======================================================
procedure TFSaisieAux.InitBalance ;
begin
CloseSaisie ;
InitPivot ;
InitGrid ;
PosLesSoldes ;
InitSaisie ;
end ;

procedure TFSaisieAux.InitSaisie ;
begin
nbLignes:=0 ;
InitPied ;
end ;

function TFSaisieAux.CanCloseSaisie : Boolean ;
begin
Result:=TRUE ;
if bReading then begin Result:=FALSE ; Exit ; end ;
//if not IsSoldeBalance then
//   begin PrintMessageRC(RC_BALNONSOLDE) ; Result:=FALSE ; end ;
end ;

procedure TFSaisieAux.CloseSaisie ;
begin
end ;

procedure TFSaisieAux.InitPied ;
begin
InfosPied ;
ZeroBlanc(PPied) ;
end ;

procedure TFSaisieAux.PosLesSoldes ;
begin
LSA_BALANCE.SetBounds(SA_BALANCE.Left,SA_BALANCE.Top,SA_BALANCE.Width,SA_BALANCE.Height) ;
LSA_SOLDE.SetBounds(SA_SOLDE.Left,SA_SOLDE.Top,SA_SOLDE.Width,SA_SOLDE.Height) ;
LSA_TOTALDEBIT.SetBounds(SA_TOTALDEBIT.Left,SA_TOTALDEBIT.Top,SA_TOTALDEBIT.Width,SA_TOTALDEBIT.Height) ;
LSA_TOTALCREDIT.SetBounds(SA_TOTALCREDIT.Left,SA_TOTALCREDIT.Top,SA_TOTALCREDIT.Width,SA_TOTALCREDIT.Height) ;
SA_SOLDE.Visible:=FALSE ; SA_BALANCE.Visible:=FALSE ;
SA_TOTALDEBIT.Visible:=FALSE ; SA_TOTALCREDIT.Visible:=FALSE ;
ChangeMask(SA_TOTALDEBIT,  memPivot.Decimale, memPivot.Symbole) ;
ChangeMask(SA_TOTALCREDIT, memPivot.Decimale, memPivot.Symbole) ;
ChangeMask(SA_SOLDE,       memPivot.Decimale, memPivot.Symbole) ;
ChangeMask(SA_BALANCE,     memPivot.Decimale, memPivot.Symbole) ;
end ;

procedure TFSaisieAux.InitGrid ;
begin
bFirst:=TRUE ;
// Avant le VidePile pour bien placer le curseur
GS.Row:=GS.FixedRows ; GS.Col:=SB_DEBIT ;
GS.VidePile(FALSE)  ;
GS.RowCount:=MAX_ROW ; GS.TypeSais:=tsFolio ;
GS.Enabled:=FALSE ; UpdateScrollBars ;
GS.ColWidths[SB_ETABL]:=0 ;
GS.ColWidths[SB_NEW]:=0 ;
GS.ColWidths[SB_COLL]:=0 ;
GS.ColWidths[SB_TYPE]:=0 ;
GS.ColWidths[SB_LETTRE]:=0 ;
GS.ColWidths[SB_MODER]:=0 ;
GS.ColAligns[SB_DEBIT]:=taRightJustify ;
GS.ColAligns[SB_CREDIT]:=taRightJustify ;
GS.GetCellCanvas:=GetCellCanvas ;
GS.PostDrawCell:=PostDrawCell ;
GS.Refresh ;
end ;

procedure TFSaisieAux.InitPivot ;
begin
//memPivot.Code:=V_PGI.DevisePivot ;
//if memPivot.Code<>'' then GetInfosDevise(memPivot) ;
//memPivot.Taux:=GetTaux(memPivot.Code,memPivot.DateTaux,EncodeDate(memOptions.Year, memOptions.Month, 1)) ;
end ;

procedure TFSaisieAux.EnableButtons ;
var bEnabled : Boolean ;
begin
BNew.Enabled:=TRUE ;
// La balance peut être validée ?
bEnabled:=IsSoldeBalance ;
// Entête
//HB_DATE1.Enabled:=bEnabled ;
BValide.Enabled:=TRUE ; //bEnabled ;
// Boutons
BSolde.Enabled:=bEnabled ;
// Bouton rechercher
BChercher.Enabled:=(nbLignes<>0) ;
end ;

procedure TFSaisieAux.InfosPied ;
begin
end ;

(*
procedure TFSaisieAux.SetGridEditing(bOn : Boolean) ;
begin
if bOn then
   begin
   GS.Options:=GS.Options-[GoRowSelect] ;
   GS.Options:=GS.Options+[GoEditing,GoTabs,GoAlwaysShowEditor] ;
   end else
   begin
   GS.Options:=GS.Options-[GoEditing,GoTabs,GoAlwaysShowEditor] ;
   GS.Options:=GS.Options+[GoRowSelect] ;
   end ;
end ;*)

//=======================================================
//=================== Gestion des TZF ===================
//=======================================================
function TFSaisieAux.GetRow(Row : LongInt) : TZF ;
begin
Result:=nil ; if Row>EcrAux.Detail.Count-1 then Exit ; Result:=TZF(EcrAux.Detail[Row]) ;
end ;

function TFSaisieAux.GetRowRef(Row : LongInt) : TZF ;
begin
Result:=nil ; if Row>EcrRef.Detail.Count-1 then Exit ; Result:=TZF(EcrRef.Detail[Row]) ;
end ;

function TFSaisieAux.SetRowTZF(Row : longint) : TZF ;
var Ecr : TZF ;
begin
Ecr:=GetRow(Row-GS.FixedRows) ;
if Ecr=nil then begin Result:=nil ; Exit ; end ;
// Entête / Commun
if Ecr.GetValue('HB_DEVISE')='' then Ecr.PutValue('HB_DEVISE', memPivot.Code) ;
Ecr.SetMontants(Valeur(GS.Cells[SB_DEBIT, Row]), Valeur(GS.Cells[SB_CREDIT, Row]),
                memPivot, FALSE) ;
// Corps
Ecr.PutValue('HB_COMPTE2', GS.Cells[SB_AUX,    Row]) ;
// Champs sup
Ecr.PutValue('LIBELLE',    GS.Cells[SB_LIB,    Row]) ;
Ecr.PutValue('LETTRABLE',  GS.Cells[SB_LETTRE, Row]) ;
Ecr.PutValue('MODEREGL',   GS.Cells[SB_MODER,  Row]) ;
Result:=Ecr ;
end ;

function TFSaisieAux.GetRowTZF(Row, RowCur : LongInt) : Boolean ;
var Ecr : TZF ; i : integer ; NumCompte : string ; EcrDebit, EcrCredit : Double ;
begin
if RowCur<0 then Ecr:=GetRow(Row) else Ecr:=GetRowRef(RowCur) ;
if Ecr=nil then begin Result:=FALSE ; Exit ; end ;
// Ajustement
Row:=Row+GS.FixedRows ;
// Entête / Commun
// Corps
NumCompte:=Ecr.GetValue('HB_COMPTE2') ;
GS.Cells[SB_AUX, Row]:=NumCompte ; i:=PlanTiers.FindCompte(NumCompte) ;
GS.Cells[SB_LIB, Row]:=PlanTiers.GetValue('T_LIBELLE', i) ;
GS.Cells[SB_LETTRE, Row]:=PlanTiers.GetValue('T_LETTRABLE', i) ;
if VarType(PlanTiers.GetValue('T_MODEREGLE', i))<>VarNull
  then GS.Cells[SB_MODER, Row]:=PlanTiers.GetValue('T_MODEREGLE', i)
  else GS.Cells[SB_MODER, Row]:='' ;
EcrDebit:=Arrondi(Ecr.GetValue('HB_DEBIT'), memPivot.Decimale) ;
EcrCredit:=Arrondi(Ecr.GetValue('HB_CREDIT'), memPivot.Decimale) ;
if EcrDebit<>0 then GS.Cells[SB_DEBIT, Row]:=StrFMontant(EcrDebit,15,memPivot.Decimale,'',TRUE) ;
if EcrCredit<>0 then GS.Cells[SB_CREDIT, Row]:=StrFMontant(EcrCredit,15,memPivot.Decimale,'',TRUE) ;
Result:=TRUE ;
end ;

//=======================================================
//==================== Fonctions SQL ====================
//=======================================================
procedure TFSaisieAux.PutGrid(Row : LongInt) ;
var NumCompte : string ; i, iFind : Integer ;
begin
NumCompte:='' ; NumCompte:=GS.Cells[SB_AUX, Row] ; iFind:=-1 ;
for i:=0 to EcrRef.Detail.Count-1 do
  if EcrRef.Detail[i].GetValue('HB_COMPTE2')=NumCompte then
    begin iFind:=i ; Break ; end ;
if iFind>=0 then GetRowTZF(Row-1, iFind) ;
end ;

function TFSaisieAux.GetNature(Nature : string) : Boolean ;
begin
Result:=FALSE ;
if (NatureGen='COC') and ((Nature='CLI') or (Nature='AUD') or (Nature='DIV')) then Result:=TRUE ;
if (NatureGen='COF') and ((Nature='FOU') or (Nature='AUC') or (Nature='DIV')) then Result:=TRUE ;
if (NatureGen='COS') and (Nature='SAL') then Result:=TRUE ;
if (NatureGen='COD') and (Nature<>'NCP') then Result:=TRUE ;
end ;

procedure TFSaisieAux.ReadBalance ;
var k, kRef, kRow : integer ;
    //bFind : Boolean ;
begin
bReading:=TRUE ;
//bFind:=FALSE ;
InitBalance ; kRow:=1 ;
GS.RowCount:=PlanTiers.Count+10 ;
InitMove(PlanTiers.Count, GetMessageRC(RC_CHARGEPLAN)) ;
for k:=1 to PlanTiers.Count do
  begin
  MoveCur(FALSE) ; kRef:=k-1 ;
  //bFind:=FALSE ;
  if (PlanTiers.GetValue('T_AUXILIAIRE', kRef)=VH^.TiersDefCli) and ((NatureGen<>'COC')) then Continue ;
  if (PlanTiers.GetValue('T_AUXILIAIRE', kRef)=VH^.TiersDefFou) and ((NatureGen<>'COF')) then Continue ;
  if (PlanTiers.GetValue('T_AUXILIAIRE', kRef)=VH^.TiersDefSal) and ((NatureGen<>'COS')) then Continue ;
  if (PlanTiers.GetValue('T_AUXILIAIRE', kRef)=VH^.TiersDefDiv) and ((NatureGen<>'COD')) then Continue ;
  if not GetNature(PlanTiers.GetValue('T_NATUREAUXI', kRef)) then Continue;
  (*
  if not GetNature(PlanTiers.GetValue('T_NATUREAUXI', kRef)) then
  begin
    if bFind then
      Break
    else
      Continue ;
  end;
  //bFind:=TRUE ; *)

  CreateRow(kRow) ;
  GS.Cells[SB_AUX,    kRow]:=PlanTiers.GetValue('T_AUXILIAIRE', kRef) ;
  GS.Cells[SB_LIB,    kRow]:=PlanTiers.GetValue('T_LIBELLE',    kRef) ;
  GS.Cells[SB_LETTRE, kRow]:=PlanTiers.GetValue('T_LETTRABLE',  kRef) ;
  GS.Cells[SB_MODER,  kRow]:=PlanTiers.GetValue('T_MODEREGLE',  kRef) ;
  PutGrid(kRow) ;
  SetRowTZF(kRow) ;
  kRow:=kRow+1 ;
  end ;
FiniMove ;
GS.RowCount:=kRow+1 ;
GS.Invalidate ;
bReading:=FALSE ;
end ;

//=======================================================
//================ Saisie complémentaires ===============
//=======================================================

//=======================================================
//================= Fonctions de calcul =================
//=======================================================
procedure TFSaisieAux.CalculSoldeBalance ;
var TotalD, TotalC, SoldeBalD, SoldeBalC : Double ; i : integer ;
begin
if GS.Row>nbLignes then Exit ;
TotalD:=0; TotalC:=0 ; SoldeBalD:=SoldeDAvant ; SoldeBalC:=SoldeCAvant ;
for i:=GS.FixedRows to nbLignes do
    begin
    // Totaux des comptes progressif
    if i<=GS.Row then
       begin
       TotalD:=TotalD+Valeur(GS.Cells[SB_DEBIT,  i]) ;
       TotalC:=TotalC+Valeur(GS.Cells[SB_CREDIT, i]) ;
       end ;
    // Calcul du solde de la balance
    SoldeBalD:=SoldeBalD+Valeur(GS.Cells[SB_DEBIT,   i]) ;
    SoldeBalC:=SoldeBalC+Valeur(GS.Cells[SB_CREDIT,  i]) ;
    end ;
SA_TOTALDEBIT.Value:=TotalD ;
SA_TOTALCREDIT.Value:=TotalC ;
AfficheLeSolde(SA_SOLDE, TotalD, TotalC) ;
AfficheLeSolde(SA_BALANCE, SoldeBalD, SoldeBalC) ;
end ;

function TFSaisieAux.IsSoldeBalance : boolean ;
begin
Result:=Arrondi(SA_BALANCE.Value, memPivot.Decimale)=0 ;
end ;

//=======================================================
//=================== Gestion du Grid ===================
//=======================================================
function TFSaisieAux.GetGridSens(ACol, ARow : integer) : integer;
begin
// Sens de déplacement dans le Grid
if (GS.Row=ARow) then
   begin
   if (GS.Col>ACol) then Result:=1 else Result:=-1;
   end else if (GS.Row>ARow) then Result:=1 else Result:=-1;
end;

(*
procedure TFSaisieAux.SetGridSep(ACol, ARow : integer ; Canvas : TCanvas; bHaut : boolean) ;
var R : TRect ;
begin
Canvas.Brush.Color := clRed ;
Canvas.Brush.Style := bsSolid ;
Canvas.Pen.Color   := clRed ;
Canvas.Pen.Mode    := pmCopy ;
Canvas.Pen.Style   := psSolid ;
Canvas.Pen.Width   := 1 ;
R:=GS.CellRect(ACol, ARow) ;
if bHaut then begin Canvas.MoveTo(R.Left, R.Top) ; Canvas.LineTo(R.Right+1, R.Top) end
         else begin Canvas.MoveTo(R.Left, R.Bottom-1) ; Canvas.LineTo(R.Right+1, R.Bottom-1) end ;
end ;*)

procedure TFSaisieAux.SetGridGrise(ACol, ARow : integer ; Canvas : TCanvas) ;
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

(*
procedure TFSaisieAux.SetGridBold(ACol, ARow : integer ; Canvas : TCanvas) ;
var R : TRect ; Text : array[0..255] of Char ;
begin
Canvas.Font.Style:=[fsBold]+[fsItalic] ;
Canvas.Font.Size:=15 ;
R:=GS.CellRect(ACol, ARow) ;
//Canvas.Rectangle(R.Left, R.Top, R.Right, R.Bottom) ;
StrPCopy(Text, GS.Cells[ACol, ARow]);
ExtTextOut(Canvas.Handle, R.Left+2, R.Top+2,
           ETO_OPAQUE or ETO_CLIPPED, @Rect, Text, StrLen(Text), nil) ;
end ; *)

procedure TFSaisieAux.GetCellCanvas(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
//var NumCompte : string ; k : Integer ;
begin
//if Plan=nil then Exit ;
//NumCompte:=GS.Cells[SB_GEN, ARow] ; k:=Plan.FindCompte(NumCompte) ;
//if (k>=0) and (Plan.GetValue('G_COLLECTIF', k)='X')
if (bReading) and (ARow>(GS.Row+GS.VisibleRowCount)) then Exit ;
if GS.Cells[SB_COLL, ARow]='X' then Canvas.Font.Style:=[fsBold] ;
if GS.Cells[SB_COLL, ARow]='O' then Canvas.Font.Style:=[fsBold] ;
end ;

procedure TFSaisieAux.PostDrawCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
var bGrise : Boolean ;
begin
if (ARow<GS.FixedRows) or (ARow>nbLignes+1) then Exit ;
if (bReading) and (ARow>(GS.Row+GS.VisibleRowCount)) then Exit ;
bGrise:=FALSE ;
if GS.Cells[SB_COLL, ARow]='X' then bGrise:=TRUE ;
if bGrise then SetGridGrise(ACol, ARow, Canvas);
end ;

procedure TFSaisieAux.UpdateScrollBars ;
begin
//if nbLignes>GS.VisibleRowCount-1 then GS.ScrollBars:=ssBoth else GS.ScrollBars:=ssHorizontal ;
end ;

function TFSaisieAux.CreateRow(Row : LongInt) : TZF ;
var //bInsert : Boolean ;
    RowRef : LongInt ;
begin
  //RowRef:=Row-1 ; //bInsert:=FALSE ;
  if Row<=nbLignes then
  begin
    //bInsert:=TRUE ;
    GS.InsertRow(Row) ;
    //if Row=1 then RowRef:=Row+1 ;
  end ;
  Inc(nbLignes) ;
  UpdateScrollBars ;
  if GS.RowCount-2<nbLignes then
    GS.RowCount:=GS.RowCount+1 ;

  // Etablissement par défaut
  GS.Cells[SB_ETABL,Row]:=VH^.EtablisDefaut ;
  // Ligne créée en saisie
  GS.Cells[SB_NEW,Row]:='O' ;
  if Row-GS.FixedRows>EcrAux.Detail.Count-1 then
    RowRef:=-1
  else
    RowRef:=Row-GS.FixedRows ;

  Result:=TZF.Create('HISTOBAL', EcrAux, RowRef) ;
end ;

procedure TFSaisieAux.DeleteRow(Row : longint) ;
begin
//ObjBalance.DeleteRow(Row-GS.FixedRows) ;
GS.DeleteRow(Row) ; Dec(nbLignes) ; UpdateScrollBars ;
if GS.RowCount>MAX_ROW then GS.RowCount:=GS.RowCount-1 ;
//SetOrdre ;
CalculSoldeBalance ;
if (Row=GS.FixedRows) and (nbLignes=0) then NextRow(Row) ;
end ;

function TFSaisieAux.NextRow(Row : longint) : boolean ;
begin
if Row>nbLignes(*+1*) then begin Result:=FALSE ; Exit ; end ;
//if Row>nbLignes then CreateRow(Row) ;
Result:=TRUE ;
end ;

// En saisie, on considère que la ligne est valide si le compte général est renseigné
function TFSaisieAux.IsRowValid(Row : integer; var ACol : integer) : boolean ;
begin
Result:=TRUE ;
if GS.Cells[SB_AUX, Row]='' then begin Result:=FALSE ; ACol:=SB_AUX ; end ;
end ;

function TFSaisieAux.IsRowOk(Row : longint) : boolean ;
begin
Result:=TRUE ;
if GS.Cells[SB_AUX, Row]='' then Result:=FALSE ;
if (GS.Cells[SB_DEBIT, Row]='') and (GS.Cells[SB_CREDIT, Row]='') then Result:=FALSE ;
end ;

//=======================================================
//================= Evénements du Grid ==================
//=======================================================
procedure TFSaisieAux.GSEnter(Sender: TObject);
begin
if bFirst then
   begin
   NextRow(GS.Row) ; CalculSoldeBalance ; 
   GS.Col:=SB_DEBIT ;
   EnableButtons ;
   PostMessage(GS.Handle, WM_KEYDOWN, VK_TAB,  0) ;
   PostMessage(GS.Handle, WM_KEYDOWN, VK_LEFT, 0) ;
   end ;
bFirst:=FALSE ;
end;

procedure TFSaisieAux.GSCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var lc, lr, sens : integer ;
begin
// Bouton Elipsis
GS.ElipsisButton:=FALSE ;
if goRowSelect in GS.Options then Exit ;
lc:=GS.Col ; lr:=GS.Row ; sens:=GetGridSens(ACol, ARow) ;
// Vérifier que l'on est pas en dehors des zones de saisies autorisées
if lr>nbLignes then begin Cancel:=TRUE ; Exit ; end ;
// Attention l'ordre des tests est important
if (ACol=SB_DEBIT) and (sens=-1) and (lr<GS.FixedRows) then
   begin Cancel:=TRUE ; Exit ; end ;
if (ACol=SB_CREDIT) and (sens=1) and (lr>nbLignes) then
   begin Cancel:=TRUE ; Exit ; end ;
if ((lc<>SB_DEBIT) and (lc<>SB_CREDIT)) then
   begin
   if sens=1 then begin if lc>SB_DEBIT  then ACol:=SB_CREDIT else ACol:=SB_DEBIT ;  end
             else begin if lc>SB_CREDIT then ACol:=SB_DEBIT  else ACol:=SB_CREDIT ; end ;
   ARow:=lr ;
   Cancel:=TRUE ;
   Exit ;
   end ;
if (GS.Cells[SB_COLL, lr]='X') then
   begin
   ARow:=lr+sens ;
   Cancel:=TRUE ; Exit ;
   end ;
end;

procedure TFSaisieAux.GSCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
// Formatage du montant
if (ACol=SB_DEBIT) or (ACol=SB_CREDIT) then
   begin
   FormatMontant(GS, ACol, ARow, memPivot.Decimale) ;
   CalculSoldeBalance ;
   end ;
end;

procedure TFSaisieAux.GSRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
var ACol : integer ;
begin
if not IsRowValid(Ou, ACol) then
   begin
   if (Ou=nbLignes) and (GetGridSens(GS.Col, Ou)=-1) then DeleteRow(Ou)
                                                     else Cancel:=TRUE ;
   end else
   begin
   SetRowTZF(Ou) ;
   if GS.Row>Ou then Cancel:=not NextRow(GS.Row) ;
   end ;
end;

procedure TFSaisieAux.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
EnableButtons ;
InfosPied ;
GS.Invalidate ;
end;

procedure TFSaisieAux.GSExit(Sender: TObject) ;
var Cancel, Chg : Boolean ; ACol, ARow : LongInt ;
begin
ACol:=GS.Col ; ARow:=GS.Row ; Cancel:=FALSE ; Chg:=FALSE ;
GSCellExit(GS, ACol, ARow, Cancel) ; GSRowExit(GS, ARow, Cancel, Chg) ;
GS.Col:=ACol ; GS.Row:=ARow ;
end;

procedure TFSaisieAux.GSKeyPress(Sender: TObject; var Key: Char);
begin
if not GS.SynEnabled then Key:=#0 else
   begin
//   if (Key='+') and ((GS.Col=SF_NAT) or (GS.Col=SF_GEN) or (GS.Col=SF_AUX)) then Key:=#0 ;
//   if (Key=' ') and (GS.Col=SF_NAT) then Key:=#0 ;
//   if (Key<>' ') and (GS.Col=SF_NAT) then Key:=#0 ;
   end ;
end;

procedure TFSaisieAux.GSMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin GridX:=X ; GridY:=Y ; end ;

//=======================================================
//=============== Evénements des boutons ================
//=======================================================
procedure TFSaisieAux.BChercherClick(Sender: TObject);
begin if bReading then Exit ; SearchClick ; end ;

procedure TFSaisieAux.BAbandonClick(Sender: TObject);
begin if bReading then Exit ; ByeClick ; end ;

procedure TFSaisieAux.BValideClick(Sender: TObject);
begin if bReading then Exit ; ValClick ; end ;

procedure TFSaisieAux.BNewClick(Sender: TObject);
begin if bReading then Exit ; NewClick ; end ;

procedure TFSaisieAux.BSoldeClick(Sender: TObject);
begin if bReading then Exit ; SoldeClick(-1) ; end ;

procedure TFSaisieAux.SearchClick ;
begin if bReading then Exit ; FindSais.Execute ; end ;

procedure TFSaisieAux.FindSaisFind(Sender: TObject);
var FindFirst : Boolean ;
begin FindFirst:=TRUE ; Rechercher(GS, FindSais, FindFirst) ; GS.Col:=SB_DEBIT ; end ;

procedure TFSaisieAux.ByeClick ;
begin bSave:=FALSE ; if not IsInside(Self) then Close ; end ;

procedure TFSaisieAux.ValClick ;
var i : Integer ;
begin
CalculSoldeBalance ;
EnableButtons ;
if not BValide.Enabled then Exit ;
for i:=1 to nbLignes do SetRowTZF(i) ;
if Parent is THPanel then CloseInsidePanel(Self) else Close ;
end;

procedure TFSaisieAux.DelClick ;
var lr, ltop : longint ;
begin
lr:=GS.Row ; ltop:=GS.TopRow ; DeleteRow(lr) ;
if lr>nbLignes then GS.Row:=nbLignes else GS.Row:=lr ;
GS.TopRow:=ltop ; GS.ElipsisButton:=FALSE ;
end ;

procedure TFSaisieAux.NewClick ;
var NumAux : string ; EcrTmp, T: TZF ; i: Integer ; iRow: LongInt ;
begin
if bClosing then Exit ;
if not ExJaiLeDroitConcept(TConcept(ccAuxCreat), FALSE) then
   begin PrintMessageRC(RC_AUXINTERDIT) ; Exit ; end ;
NumAux:=GS.Cells[SB_AUX, GS.Row] ; iRow:=-1 ;
FicheTiers(nil, '', '', taCreatOne, 0) ;
EcrTmp:=TZF.Create('HISTOBAL', nil, -1) ;
EcrTmp.Dupliquer(EcrAux, TRUE, TRUE) ;
EcrAux.Free ;
EcrAux:=TZF.Create('HISTOBAL', nil, -1) ;
EcrAux.Dupliquer(EcrRef, FALSE, TRUE) ;
PlanTiers.Free ;
PlanTiers:=TZPlan.Create ;
PlanTiers.Load(TRUE, TRUE, FALSE) ;
ReadBalance ;
for i:=0 to EcrAux.Detail.Count-1 do
  begin
  if EcrAux.Detail[i].GetValue('HB_COMPTE2')=NumAux then iRow:=i ;
  T:=TZF(EcrTmp.FindFirst(['HB_COMPTE2'], [EcrAux.Detail[i].GetValue('HB_COMPTE2')], FALSE)) ;
  if T<>nil then
    begin
    EcrAux.Detail[i].PutValue('HB_DEBIT',  T.GetValue('HB_DEBIT')) ;
    EcrAux.Detail[i].PutValue('HB_CREDIT', T.GetValue('HB_CREDIT')) ;
    GetRowTZF(i, -1) ;
    end ;
  end ;
EcrTmp.Free ;
if iRow>0 then GS.Row:=iRow+1 ;
GS.Enabled:=TRUE ; GS.SetFocus ; GSEnter(nil) ;
end ;

procedure TFSaisieAux.SoldeClick(Row : longint) ;
var lr : longint ; Solde, Debit, Credit : double ; bSoldeDebit : Boolean ;
begin
if Row<0 then lr:=GS.Row else lr:=Row ;
Solde:=SA_BALANCE.Value ; bSoldeDebit:=SA_BALANCE.Debit ;
Debit:=Valeur(GS.Cells[SB_DEBIT,lr]) ; Credit:=Valeur(GS.Cells[SB_CREDIT,lr]) ;
// Partie débit
if (Debit<>0) and (bSoldeDebit) then
   if (Debit-Solde)<0 then
      begin
      GS.Cells[SB_CREDIT,lr]:=StrFMontant(Abs(Debit-Solde),15,memPivot.Decimale,'',TRUE) ;
      GS.Cells[SB_DEBIT,lr]:='' ;
      GS.Col:=SB_CREDIT ;
      end else GS.Cells[SB_DEBIT,lr]:=StrFMontant(Debit-Solde,15,memPivot.Decimale,'',TRUE) ;
if (Debit<>0) and (not bSoldeDebit) then
   GS.Cells[SB_DEBIT,lr]:=StrFMontant(Debit+Solde,15,memPivot.Decimale,'',TRUE) ;
// Partie crédit
if (Credit<>0) and (bSoldeDebit) then
   GS.Cells[SB_CREDIT,lr]:=StrFMontant(Credit+Solde,15,memPivot.Decimale,'',TRUE) ;
if (Credit<>0) and (not bSoldeDebit) then
   if (Credit-Solde)<0 then
      begin
      GS.Cells[SB_DEBIT,lr]:=StrFMontant(Abs(Credit-Solde),15,memPivot.Decimale,'',TRUE) ;
      GS.Cells[SB_CREDIT,lr]:='' ;
      GS.Col:=SB_DEBIT ;
      end else GS.Cells[SB_CREDIT,lr]:=StrFMontant(Credit-Solde,15,memPivot.Decimale,'',TRUE) ;
if (Debit=0) and (Credit=0) then
   begin
   if bSoldeDebit then GS.Cells[SB_CREDIT,lr]:=StrFMontant(Solde,15,memPivot.Decimale,'',TRUE)
                  else GS.Cells[SB_DEBIT,lr]:=StrFMontant(Solde,15,memPivot.Decimale,'',TRUE) ;
   end ;
CalculSoldeBalance ;
EnableButtons ;
end ;

//=======================================================
//=========== Evénements des boutons d'entête ===========
//=======================================================

end.

