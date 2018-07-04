unit eSaisBud;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Forms, Dialogs,
  Hcompte, StdCtrls, Hctrls, ExtCtrls, Menus, Grids, Buttons, HSysMenu, HTB97, HPanel, HMsgBox,
  RepartBU_TOM, // ParamRepartBUDGET
  BudGene_TOM,  // FicheBudGene
  BudSect_TOM,  // FicheBudSect
  BudJal_TOM,   // FicheBudjal
  UTob,
{$IFDEF EAGLCLIENT}
{$ELSE}
  DB, 
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  HEnt1,    // TActionFiche, String3, V_PGI
  Ent1,     // PurgePopup, LookLesDocks, InitPopup, ListePeriode, PositionneEtabUser, PopZoom97
  SaisUtil, // RMVt, TOBM, TModeSaisBud
  BudComp,  // SaisieBudComp
  BudVent,  // X_BV, VentilBud
  CleRepBU, // GetCleRepart
  ed_tools, // VideListe
  UiUtil,   // FindInsidePanel, InitInside
  SaisComm; // BudgetToIdent, WhereEcriture

procedure SaisieBudget ( Action : TActionFiche ; M : RMVT ; VisuConso : boolean ) ;
Function  TrouveSaisBud(Q : TQuery ; Var M : RMVT ) : Boolean ;
Function  TrouveEtLanceSaisBud(Q : TQuery ; TypeAction : TActionFiche) : Boolean ;
procedure VisuConsoBudget ( tsb : Char ; UnJal : String) ;
function ModifieRequeteGrille(SQL: String): String; {b FP FQ16008-FQ16048}

type
  TFSaisBud = class(TForm)
    PEntete: THPanel;
    H_PERFIN: THLabel;
    BRazCol: THBitBtn;
    BRazLigne: THBitBtn;
    BRazAll: THBitBtn;
    BPasteAll: THBitBtn;
    BZoomJournal: THBitBtn;
    BZoomBudget: THBitBtn;
    BZoomSection: THBitBtn;
    POPS: TPopupMenu;
    POPZ: TPopupMenu;
    HM: THMsgBox;
    FindSais: TFindDialog;
    HDiv: THMsgBox;
    PG: TPanel;
    GB: THGrid;
    GDroit: THGrid;
    Panel5: THPanel;
    GBas: THGrid;
    TG: TPanel;
    PD: THPanel;
    SD: TScrollBar;
    SB: TScrollBar;
    HMTrad: THSystemMenu;
    H_JOURNAL: THLabel;
    H_NATUREBUD: THLabel;
    H_AXE: THLabel;
    H_BUDSECT: THLabel;
    H_BUDGENE: THLabel;
    BE_BUDSECT: THCpteEdit;
    BE_NATUREBUD: THValComboBox;
    BE_BUDJAL: THValComboBox;
    H_PERDEB: THLabel;
    PERDEB: THValComboBox;
    H_NUMEROPIECE: THLabel;
    BE_NUMEROPIECE: THLabel;
    BevelNum: TBevel;
    BE_BUDGENE: THCpteEdit;
    BE_AXE: THValComboBox;
    BE_EXERCICE: THValComboBox;
    PERFIN: THValComboBox;
    BE_EXERCICE_: THValComboBox;
    H_ETABLISSEMENT: THLabel;
    BE_ETABLISSEMENT: THValComboBox;
    BRepart: THBitBtn;
    DockTop: TDock97;
    DockRight: TDock97;
    DockLeft: TDock97;
    DockBottom: TDock97;
    Outils97: TToolbar97;
    BVentil: TToolbarButton97;
    BComplement: TToolbarButton97;
    BChercher: TToolbarButton97;
    BMenuZoom: TToolbarButton97;
    BAction: TToolbarButton97;
    BCreerRepart: THBitBtn;
    ToolbarSep972: TToolbarSep97;
    Valide97: TToolbar97;
    BValide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    BCopyCol: TToolbarButton97;
    BPasteCol: TToolbarButton97;
    ToolbarSep973: TToolbarSep97;
    BZoomSousPlan: TToolbarButton97;
    EXOSINTER: THValComboBox;
    ExoCalcul: THValComboBox;
    BSwapSaisie: THBitBtn;
    HLabel3: THLabel;
    VUE: THValComboBox;
    HLabel4: THLabel;
    RESOL: THValComboBox;
    CSens: TComboBox;
    H_CSens: THLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormShow(Sender: TObject);

    procedure POPSPopup(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure BRazColClick(Sender: TObject);
    procedure BRazLigneClick(Sender: TObject);
    procedure BRazAllClick(Sender: TObject);
    procedure BCopyColClick(Sender: TObject);
    procedure BPasteColClick(Sender: TObject);
    procedure BPasteAllClick(Sender: TObject);
    procedure BRepartClick(Sender: TObject);
    procedure BCreerRepartClick(Sender: TObject);
    procedure BZoomBudgetClick(Sender: TObject);
    procedure BZoomSectionClick(Sender: TObject);
    procedure BZoomJournalClick(Sender: TObject);
    procedure BComplementClick(Sender: TObject);
    procedure BVentilClick(Sender: TObject);
    procedure BE_EXERCICEChange(Sender: TObject);
    procedure BE_EXERCICE_Change(Sender: TObject);
    procedure BE_BUDJALChange(Sender: TObject);
    procedure BE_AXEChange(Sender: TObject);
    procedure VUEChange(Sender: TObject);
    procedure RESOLChange(Sender: TObject);
    procedure BE_BUDSECTExit(Sender: TObject);
    procedure BE_BUDGENEExit(Sender: TObject);
    procedure FindSaisFind(Sender: TObject);
    procedure GBCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure GBCellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure GBTopLeftChanged(Sender: TObject);
    procedure SBChange(Sender: TObject);
    procedure SDChange(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure GBEnter(Sender: TObject);
    procedure GBExit(Sender: TObject);
    procedure BE_ETABLISSEMENTChange(Sender: TObject);
    procedure BActionMouseEnter(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure Combo97DockChanged(Sender: TObject);
    procedure BZoomSousPlanClick(Sender: TObject);
    procedure BSwapSaisieClick(Sender: TObject);
    procedure CSensChange(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    ScrollEnCours,SensCorrect,OkControle,MonoCol : Boolean ;
    WidthInit     : Integer ;
    HeightInit,CurCopy : Integer ;
    CurCoef       : Double ;
    ColCopy       : TStrings ;
    LesComptes,LesSections  : String ;
    LesLibComptes,LesLibSections : String ;
    NbPer,NbLigOrig,CurDec,Deci  : integer ;
    OkD,GeneCharge,DateCoher,PasOkVUE : boolean ;
    StCellCur,GeneSouche    : String ;
    NowFutur,OldModif       : TDateTime ;
    GeneAttente,SectAttente : String ;
    GeneSens                : String3 ;
    CatBudget               : String ;
    SPCat                   : String ;
    DateDebExos             : Array[1..10] of TDateTime ;

    {JP 24/08/05 : FQ 16501 et 16502 : On désactive la posibilité de modifier le compte ou
                   la section en saisie croisée dès lors que la saisie est commencée => je
                   Transforme le "champ" PieceModifiee en property}
    FPieceModifiee : Boolean;
    {Setter de la property PieceModifiee}
    procedure SetPieceModifiee (Value : Boolean);
    {Active / désactive la section / compte}
    procedure GereCodeImplicite(Actif : Boolean);
    {FIN JP 24/08/05 : FQ 16501 et 16502}

    Procedure Synchro (DepuisGrid : Boolean)  ;
    Procedure AjusteAscenceurs ;
    function  QuellePer ( DD : TDateTime ) : integer ;
    Procedure TrouveCase ( O : TOBM ; Var ACol,ARow : integer ) ;
    Procedure ShowCase ( ACol,ARow : integer ) ;
    function  SommeCase ( ACol,ARow : integer ) : double ;
    procedure SommeTout ;
    procedure SommeLigne ( ARow : integer ) ;
    procedure SommeCol ( ACol : integer ) ;
    Procedure FormateZone ( ACol,ARow : integer ) ;
    procedure GereCell ( ACol,ARow : integer ) ;
    function  ColToDate ( Col : integer ) : TDateTime ;
    Function  Ecraser ( ACol : integer ) : boolean ;
    Function  ColVide ( ACol : integer ) : boolean ;
// Objets
    Procedure CreerObjetCase ( ACol,ARow : integer ) ;
    Procedure FreeObjetCase ( ACol,ARow : integer ) ;
    Procedure FreeCopy ;
    Procedure AddOBMCase ( ACol,ARow : integer ; O : TOBM ) ;
    Function  GetL ( ACol,ARow : integer ) : TList ;
    procedure CreerNewOBM ( ACol,ARow : integer ) ;
    procedure TransformeGrille ;
    Function  DuplicImage ( T : TList ; ACol : integer ) : TList ;
    procedure CollerColonne ( ACol : integer ) ;
    procedure RepartBudget ( ColRef : integer ; LesTaux : String ; Courant : boolean ) ;
// Entête, Chargements
(*    Procedure ChargeLesLib(St : String ; OnCpt : Boolean ; Var LesLib : String); *)
    procedure ChargeInfosJournal ;
    procedure RempliExosInter ;
    Procedure RemplitCase ( O : TOBM ) ;
    Procedure ChargeLesComptes ;
    Procedure ChargeLesPeriodes ;
    Procedure ChargeLesMouvements ;
    procedure DefautEntete ;
    procedure LectureSeule ;
    procedure GereEnabled ;
    procedure GereComboVue ;
    procedure GereModeSaisie ;
// Outils
    procedure ClickCherche ;
    procedure ClickComplement ;
    procedure ClickVentil ;
    procedure ClickCopyCol ;
    procedure ClickPasteCol ;
    procedure ClickPasteAll ;
    procedure ClickSwapSaisie ;
    procedure ClickRepart ;
    procedure ClickCreerRepart ;
    procedure ClickJournal ;
    procedure ClickValide ( OkFerme : boolean ) ;
    procedure ClickZoomBudget ;
    procedure ClickZoomSection ;
    procedure ClickRazCL(Acol,ARow : Integer ; RazAll : boolean ) ;
// Validation
    Function  ControleLeSens : Boolean ;
    function  SensCoher ( O : TOBM ; TBud,TSect : TStrings ) : boolean ;
    function  PieceCorrecte : boolean ;
    procedure ValideLaPiece ;
    procedure DetruitAncien ;
    procedure ValideNumero ;
    procedure ValideCells ;
    Function  JalOk : Boolean ;
    Function  NatureOk : Boolean ;
    Function  GeneAttOk : Boolean ;
    Function  SectAttOk : Boolean ;
// Vues
    procedure ShowBudget ;
    procedure ChangeImplicite ;
    Procedure AttribLeTitre ;
// Form
    procedure CloseSaisie ;
    procedure InitSaisie ;
    procedure CreateSaisie ;
    procedure CloseFen ;
  public
    M : RMVT ;
    Action,ActionComp : TActionFiche ;
    FindFirst,VisuConso,LastCourant : Boolean ;
    NumeroPiece  : Longint ;
    ModeSaisBud  : TModeSaisBud ;
    {JP 24/08/05 : FQ 16501 et 16502 : On désactive la posibilité de modifier le compte ou
                   la section en saisie croisée dès lors que la saisie est commencée => je
                   Transforme le "champ" PieceModifiee en property}
    property PieceModifiee : Boolean read FPieceModifiee write SetPieceModifiee;
  end;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPTypeCons,
  {$ENDIF MODENT1}
  HZoomSP; // CatExiste, ChoisirValCat

procedure VisuConsoBudget ( tsb : Char ; UnJal : String) ;
Var M  : RMVT ;
begin
FillChar(M,Sizeof(M),#0) ;
M.CodeD:=V_PGI.DevisePivot ;
M.Simul:='N' ; M.TypeSaisie:=tsB ; M.Jal:=UnJal ;
SaisieBudget(taConsult,M,True) ;
end;

Function TrouveSaisBud(Q : TQuery ; Var M : RMVT ) : Boolean ;
Var Q1 : TQuery ;
    Trouv : boolean ;
begin
TrouveSaisBud:=FALSE ;
if (Q.EOF) And (Q.Bof) then Exit ;
Q1:=OpenSQL('Select * from BUDECR where BE_BUDJAL="'+Q.FindField('BE_BUDJAL').AsString+'"'
          +' AND BE_NATUREBUD="'+Q.FindField('BE_NATUREBUD').AsString+'"'
          +' AND BE_QUALIFPIECE="'+Q.FindField('BE_QUALIFPIECE').AsString+'"'
          +' AND BE_NUMEROPIECE='+Q.FindField('BE_NUMEROPIECE').AsString,True) ;
Trouv:=Not Q1.EOF ; if Trouv then M:=BudgetToIdent(Q1) ;
Ferme(Q1) ;
TrouveSaisBud:=Trouv ;
END ;

Function TrouveEtLanceSaisBud(Q : TQuery ; TypeAction : TActionFiche) : Boolean ;
Var M : RMVT ;
begin
Result:=TrouveSaisBud(Q,M) ;
if Result then SaisieBudget(TypeAction,M,False) ;
END ;

procedure SaisieBudget ( Action : TActionFiche ; M : RMVT ; VisuConso : boolean ) ;
Var X  : TFSaisBud ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
X:=TFSaisBud.Create(Application) ;
X.M:=M ; X.Action:=Action ; X.ActionComp:=Action ; X.VisuConso:=VisuConso ;
if PP=Nil then
   BEGIN
    Try
     X.ShowModal ;
    Finally
     X.Free ;
    End ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;


{============================= FORM =============================}
procedure TFSaisBud.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var OkG,Vide : boolean ;
begin
if Not GB.SynEnabled then BEGIN Key:=0 ; Exit ; END ;
OkG:=(Screen.ActiveControl=GB) ; Vide:=(Shift=[]) ;
Case Key of
  VK_ESCAPE : if Vide then CloseFen ;
  VK_RETURN : if ((OkG) and (Vide)) then KEY:=VK_TAB ;
  VK_F5     : if OkG then BEGIN Key:=0 ; if Vide then ClickZoomBudget else if Shift=[ssAlt] then ClickZoomSection ; END ;
  VK_F8     : BEGIN Key:=0 ; ClickSwapSaisie ; END ;
  VK_F10    : if Vide then BEGIN Key:=0 ; ClickValide(True) ; END ;
{AA}     65 : if ((OkG) and (Shift=[ssAlt])) then BEGIN Key:=0 ; ClickVentil ; END ;
{AC}     67 : if ((OkG) and (Shift=[ssAlt])) then BEGIN Key:=0 ; ClickComplement ; END ;
{^F}     70 : if Shift=[ssCtrl] then BEGIN Key:=0 ; ClickCherche ; END ;
{AJ}     74 : if Shift=[ssAlt] then BEGIN Key:=0 ; ClickJournal ; END ;
{AR}     82 : if Shift=[ssAlt] then BEGIN Key:=0 ; ClickRepart ; END ;
  END ;
end;

procedure TFSaisBud.InitSaisie ;
Var i : integer ;
BEGIN
GB.VidePile(True) ; GB.ColCount:=50 ; GB.TypeSais:=tsBudget ;
GB.ColWidths[0]:=90 ; GBas.ColCount:=GB.ColCount ;
for i:=1 to GB.ColCount-1 do BEGIN GB.ColAligns[i]:=taRightJustify ; GB.ColWidths[i]:=70 ; END ;
for i:=1 to GBas.ColCount-1 do GBas.ColAligns[i]:=taRightJustify ;
GDroit.ColAligns[0]:=taRightJustify ;
ModeSaisBud:=msbGene ; OkControle:=True ;
WidthInit := Width ; HeightInit := Height ;
Deci:=V_PGI.OkDecV ; MonoCol:=False ;
FillChar(DateDebExos,Sizeof(DateDebexos),#0) ;
END ;

procedure TFSaisBud.CreateSaisie ;
BEGIN
GeneSouche:='' ; DateCoher:=True ;
NbPer:=0 ; PieceModifiee:=False ; NbLigOrig:=0 ;
ScrollEnCours:=FALSE ;
PasOkVUE:=False ; SensCorrect:=True ; GeneSens:='M' ;
CurCopy:=0 ; ColCopy:=TStringList.Create ; ColCopy.Add('') ; LastCourant:=True ;
END ;

procedure TFSaisBud.FormCreate(Sender: TObject);
begin
InitSaisie ;
CreateSaisie ;
RegLoadToolbarPos(Self,'SaisieBudget') ;

end;

procedure TFSaisBud.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
if Action=taConsult then Exit ;
if Not PieceModifiee then Exit ;
if HM.Execute(0,'','')<>mrYes then CanClose:=False ;
end;


procedure TFSaisBud.CloseSaisie ;
BEGIN
GB.VidePile(True) ; ColCopy.Free ;
PurgePopup(POPS) ; PurgePopup(POPZ) ;
END ;

procedure TFSaisBud.FormClose(Sender: TObject; var Action: TCloseAction);
begin
CloseSaisie ;
RegSaveToolbarPos(Self,'SaisieBudget') ;
if Parent is THPanel then
  Action:=caFree ;
end;

procedure TFSaisBud.ShowBudget ;
BEGIN
GereModeSaisie ;
if Not GeneCharge then BEGIN ChargeLesPeriodes ; ChargeLesComptes ; END ;
if ((Action<>taCreat) and (Not VisuConso)) then ChargeLesMouvements ;
PieceModifiee:=False ;
END ;

procedure TFSaisBud.FormShow(Sender: TObject);
Var OkM : boolean ;
begin
if Action=taCreat then BE_BUDJAL.DataType:='TTBUDJALSAIS' ; 
LookLesDocks(Self) ;
GeneCharge:=True ; VUE.Enabled:=False ;
OkM:=((M.Valide) and (Action=taModif)) ;
if OkM then Action:=taConsult ;
OkD:=True ; RESOL.Value:='F' ;
AttribLeTitre ;
DefautEntete ;
ShowBudget ;
if OkM then HM.Execute(2,'','') ;
GeneCharge:=False ;
if Action=taModif then BEGIN GB.SetFocus ; GBEnter(Nil) ; END ;
end;

{============================== POP ============================}
procedure TFSaisBud.POPSPopup(Sender: TObject);
begin
InitPopUp(Self) ;
end;

{================================ BOUTONS =================================}
procedure TFSaisBud.BAbandonClick(Sender: TObject);
begin
CloseFen ;
end;

procedure TFSaisBud.BZoomBudgetClick(Sender: TObject);
begin ClickZoomBudget ; end;

procedure TFSaisBud.BZoomSectionClick(Sender: TObject);
begin ClickZoomSection ; end;

procedure TFSaisBud.BZoomJournalClick(Sender: TObject);
begin ClickJournal ; end;

procedure TFSaisBud.BChercherClick(Sender: TObject);
begin ClickCherche ; end;

procedure TFSaisBud.BCreerRepartClick(Sender: TObject);
begin ClickCreerRepart ; end;

procedure TFSaisBud.BComplementClick(Sender: TObject);
begin ClickComplement ; end;

procedure TFSaisBud.BCopyColClick(Sender: TObject);
begin ClickCopyCol ; end;

procedure TFSaisBud.BPasteColClick(Sender: TObject);
begin ClickPasteCol ; end;

procedure TFSaisBud.BSwapSaisieClick(Sender: TObject);
begin ClickSwapSaisie ; end;

procedure TFSaisBud.BRepartClick(Sender: TObject);
begin ClickRepart ; end;

procedure TFSaisBud.BPasteAllClick(Sender: TObject);
begin ClickPasteAll ; end;

procedure TFSaisBud.BVentilClick(Sender: TObject);
begin ClickVentil ; end;

procedure TFSaisBud.BRazColClick(Sender: TObject);
begin ClickRazCL(GB.Col,0,False) ; end;

procedure TFSaisBud.BRazLigneClick(Sender: TObject);
begin ClickRazCL(0,GB.Row,False) ; end;

procedure TFSaisBud.BRazAllClick(Sender: TObject);
begin ClickRazCL(0,0,True) ; end;

procedure TFSaisBud.FindSaisFind(Sender: TObject);
begin Rechercher(GB,FindSais,FindFirst) ; end;

procedure TFSaisBud.ClickComplement ;
Var T : TList ;
BEGIN
if Not BComplement.Enabled then Exit ;
T:=GetL(GB.Col,GB.Row) ; if T=Nil then Exit ;
if PasOkVUE then ActionComp:=taConsult ;
if SaisieBudComp(T,ActionComp) then PieceModifiee:=True;
END ;

Function TFSaisBud.DuplicImage ( T : TList ; ACol : integer ) : TList ;
Var Exo : String ;
    DD     : TDateTime ;
    L      : TList ;
    k      : integer ;
    OS,OD  : TOBM ;
BEGIN
Result:=Nil ; if T=Nil then Exit ;
DD:=ColToDate(ACol) ; Exo:=QuelExoDT(DD) ;
L:=TList.Create ;
for k:=0 to T.Count-1 do
    BEGIN
    OS:=TOBM(T[k]) ; if OS=Nil then Continue ;
    OD:=Nil ; EgaliseOBM(OS,OD) ;
    OD.PutMvt('BE_DATECOMPTABLE',DD) ;
    OD.PutMvt('BE_EXERCICE',Exo) ;
    L.Add(OD) ;
    END ;
Result:=L ;
END ;

procedure TFSaisBud.ClickCopyCol ;
Var i    : integer ;
    T,T2   : TList ;
    St     : String ;
BEGIN
if Action=taConsult then Exit ;
if Not BAction.Enabled then Exit ;
CurCopy:=GB.Col ; FreeCopy ;
for i:=1 to GB.RowCount-1 do
    BEGIN
    T:=GetL(GB.Col,i) ;
    if T=Nil then ColCopy.Add('') else
       BEGIN
       T2:=DuplicImage(T,GB.Col) ; St:=GB.Cells[GB.Col,i] ;
       ColCopy.AddObject(St,T2) ;
       END ;
    END ;
PieceModifiee:=True ;
END ;

Function TFSaisBud.ColVide ( ACol : integer ) : boolean ;
Var i : integer ;
BEGIN
ColVide:=True ;
for i:=1 to GB.RowCount-1 do if GetL(ACol,i)<>Nil then BEGIN ColVide:=False ; Break ; END ;
END ;

Function TFSaisBud.Ecraser ( ACol : integer ) : boolean ;
BEGIN
Ecraser:=True ;
if Not ColVide(ACol) then
   BEGIN
   if HM.Execute(3,'','')<>mrYes then Ecraser:=False ;
   END ;
END ;

procedure TFSaisBud.RepartBudget ( ColRef : integer ; LesTaux : String ; Courant : boolean ) ;
Var ARow,ACol : integer ;
    Montant,X,Total   : Double ;
    St,StC    : String ;
    T,T2      : TList ;
BEGIN
for ARow:=1 to GB.RowCount-1 do if ((ARow=GB.Row) or (Not Courant)) then
    BEGIN
    Montant:=Valeur(GB.Cells[ColRef,ARow]) ; if Arrondi(Montant,CurDec)=0 then Continue ;
    St:=LesTaux ; Total:=0 ;
    for ACol:=1 to GB.ColCount-1 do
        BEGIN
        StC:=ReadTokenSt(St) ;
        if StC<>'' then
           BEGIN
           T:=GetL(ColRef,ARow) ;
           if ((T<>Nil) and (ACol<>ColRef)) then
              BEGIN
              FreeObjetCase(ACol,ARow) ;
              T2:=DuplicImage(T,ACol) ;
              {JP 25/08/05 : FQ 16380 : Ajout du Trim pour éviter certains espaces intempestifs}
              GB.Cells[ACol,ARow]:= Trim(GB.Cells[ColRef,ARow]);
              GB.Objects[ACol,ARow]:= T2 ;
              END ;
           if ACol=GB.ColCount-1 then X:=Arrondi(Montant-Total,CurDEC) else X:=Arrondi(Montant*Valeur(StC)/100.0,CurDec) ;
           {JP 25/08/05 : FQ 16380 : Ajout du Trim pour éviter certains espaces intempestifs}
           GB.Cells[ACol,ARow]:=Trim(StrS(X,CurDec));
           FormateZone(ACol,ARow) ; CreerObjetCase(ACol,ARow) ;
           GereCell(ACol,ARow) ; Total:=Total+X ;
           END ;
        SommeCol(ACol) ;
        END ;
    SommeLigne(ARow) ;
    END ;
SommeTout ;
PieceModifiee:=True ;
END ;

procedure TFSaisBud.CollerColonne ( ACol : integer ) ;
Var i : integer ;
    T,TC : TList ;
BEGIN
for i:=1 to GB.RowCount-1 do FreeObjetCase(ACol,i) ;
for i:=1 to GB.RowCount-1 do
    BEGIN
    TC:=TList(ColCopy.Objects[i]) ;
    if ((ColCopy[i]='') or (TC=Nil)) then Continue ;
    T:=DuplicImage(TC,ACol) ;
    {JP 25/08/05 : FQ 16380 : Ajout du Trim pour éviter certains espaces intempestifs}
    GB.Cells[ACol,i]:= Trim(ColCopy[i]) ;
    GB.Objects[ACol,i]:=T ;
    END ;
PieceModifiee:=True ;
END ;

procedure TFSaisBud.ClickCreerRepart ;
BEGIN
if Action=taConsult then Exit ;
if Not BAction.Enabled then Exit ;
if Not BCreerRepart.Enabled then Exit ;
ParamRepartBUDGET('',True);
END ;

procedure TFSaisBud.ClickRepart ;
Var CleRep  : String ;
    ColRef  : integer ;
    Courant : boolean ;
    LesTaux : String ;
    Q       : TQuery ;
BEGIN
if Action=taConsult then Exit ;
if Not BAction.Enabled then Exit ;
if Not BRepart.Enabled then Exit ;
CleRep:='' ; ColRef:=0 ; LesTaux:='' ; Courant:=LastCourant ;
GetCleRepart(BE_BUDJAL.Value,GB.Rows[0],CleRep,ColRef,Courant);
if ((CleRep='') or (ColRef<=0)) then Exit ; LastCourant:=Courant ;
Q:=OpenSQL('Select BR_TAUX from REPARTBU Where BR_REPARPERIODE="'+CleRep+'"',True) ;
if Not Q.EOF then LesTaux:=Q.Fields[0].AsString else LesTaux:='' ;
Ferme(Q) ;
if LesTaux<>'' then RepartBudget(ColRef,LesTaux,Courant) ;
END ;

procedure TFSaisBud.ClickSwapSaisie ;
Var i,j : integer ;
    DD  : TDateTime ;
    Cacher : boolean ;
BEGIN
MonoCol:=Not MonoCol ;
if MonoCol then
   BEGIN
   for i:=2 to GB.ColCount-1 do
       BEGIN
       Cacher:=True ; DD:=ColToDate(i) ;
       for j:=1 to 10 do if ((DateDebExos[j]<>0) and (DateDebExos[j]=DD)) then Cacher:=False ;
       if Cacher then GB.ColWidths[i]:=0 ;
       END ;
   GB.Options:=GB.Options-[goColSizing] ;
   END else
   BEGIN
   for i:=2 to GB.ColCount-1 do if GB.ColWidths[i]=0 then GB.ColWidths[i]:=GB.ColWidths[1] ;
   GB.Options:=GB.Options+[goColSizing] ;
   END ;
END ;

procedure TFSaisBud.ClickPasteAll ;
Var i,ACol : integer ;
BEGIN
if Action=taConsult then Exit ;
if Not BAction.Enabled then Exit ;
if Not BPasteAll.Enabled then Exit ;
if CurCopy<=0 then Exit ;
for ACol:=1 to GB.ColCount-1 do if ACol<>CurCopy then
    BEGIN
    if ColVide(ACol) then BEGIN CollerColonne(ACol) ; SommeCol(ACol) ; END ;
    END ;
for i:=1 to GB.RowCount-1 do SommeLigne(GB.Row) ;
SommeTout ;
PieceModifiee:=True ;
END ;

procedure TFSaisBud.ClickPasteCol ;
Var i : integer ;
BEGIN
if Action=taConsult then Exit ;
if Not BAction.Enabled then Exit ;
if Not BPasteCol.Enabled then Exit ;
if ((CurCopy<=0) or (CurCopy=GB.Col)) then Exit ;
if Not Ecraser(GB.Col) then Exit ;
CollerColonne(GB.Col) ;
for i:=1 to GB.RowCount-1 do SommeLigne(i) ;
SommeCol(GB.Col) ; SommeTout ;
PieceModifiee:=True ;
END ;

procedure TFSaisBud.ClickVentil ;
Var T : TList ;
    X : X_BV ;
BEGIN
if Not BVentil.Enabled then Exit ;
if ModeSaisBud in [msbGeneSect,msbSectGene] then Exit ;
T:=GetL(GB.Col,GB.Row) ; if T=Nil then Exit ;
if SommeCase(GB.Col,GB.Row)=0 then Exit ;
FillChar(X,Sizeof(X),#0) ;
X.T:=T ; X.Action:=Action ; X.ModeSaisBud:=ModeSaisBud ; X.Cpt:=GB.Cells[0,GB.Row] ;
X.LaPer:=GB.Cells[GB.Col,0] ;
X.OkD:=OkD ; X.CurDec:=CurDec ; X.CurCoef:=CurCoef ; X.Jal:=BE_BUDJAL.Value ;
if ModeSaisBud=msbGene then X.StC:=LesSections else X.StC:=LesComptes ;
X.CatJal:=CatBudget ; X.GeneAttente:=GeneAttente ; X.SectAttente:=SectAttente ;
if VentilBud(X) then
   BEGIN
   PieceModifiee:=True ; ShowCase(GB.Col,GB.Row) ;
   SommeCol(GB.Col) ; SommeLigne(GB.Row) ; SommeTout ;
   END ;
END ;

procedure TFSaisBud.ClickRazCL(Acol,ARow : Integer ; RazAll : Boolean) ;
Var j,ic,ir : Integer ;
    Li : TList ;
    O : TOBM ;
    OkOk : Boolean ;
BEGIN
if Action=taConsult then Exit ;
if Not BAction.Enabled then Exit ;
if RazAll then if HM.Execute(1,'','')<>mrYes then Exit ;
for ir:=1 to GB.RowCount-1 do for ic:=1 to GB.ColCount-1 do
   if (ir=ARow) or (ic=ACol) or RazAll then
   BEGIN
   OkOk:=False ;
   Li:=GetL(ic,ir) ; if Li=Nil then Continue ;
   for j:=0 to Li.Count-1 do
      BEGIN
      O:=TOBM(Li[j]) ;
      if OkD then
         BEGIN
         O.PutMvt('BE_DEBIT',0) ;
         if O.GetDebitSaisi<>0 then OkOk:=True ;
         END else
         BEGIN
         O.PutMvt('BE_CREDIT',0) ;
         if O.GetCreditSaisi<>0 then OkOk:=True ;
         END ;
      END ;
   if Not OkOk then
      BEGIN
      VideListe(Li) ; Li.Free ; GB.Objects[ic,ir]:=Nil ; GB.Cells[ic,ir]:='' ;
      END else
      BEGIN
      {JP 25/08/05 : FQ 16380 : Ajout du Trim pour éviter certains espaces intempestifs}
      GB.Cells[ic,ir]:=Trim(StrS(0,CurDec));
      END ;
   if ((ACol>0) or (RazAll)) then SommeLigne(ir) ;
   if ((ARow>0) or (RazAll)) then SommeCol(ic) ;
   END ;
if (ACol > 0) and (Not RazAll) then SommeCol(Acol) ;
if (ARow > 0) and (Not RazAll) then SommeLigne(ARow) ;
SommeTout ;
PieceModifiee:=True ;
END ;

procedure TFSaisBud.ClickCherche ;
BEGIN
if Not BChercher.Enabled then Exit ;
if GB.RowCount<=1 then Exit ;
FindFirst:=True ; FindSais.Execute ;
END ;

procedure TFSaisBud.ClickZoomBudget ;
Var St : String ;
BEGIN
if Not BZoomBudget.Enabled then Exit ;
St:='' ;
Case ModeSaisBud of
   msbGene,msbGeneSect : St:=GB.Cells[0,GB.Row] ;
               msbSect : St:=GeneAttente ;
           msbSectGene : BEGIN St:=BE_BUDGENE.Text ; if St='' then St:=GeneAttente ; END ;
   END ;
 if St<>'' then FicheBudGene(Nil,'',St,taConsult,0);
END ;

procedure TFSaisBud.ClickZoomSection ;
var
  St : String ;
begin
  if Not BZoomSection.Enabled then Exit;
  St:='';
  Case ModeSaisBud of
    msbSect,msbSectGene : St:=GB.Cells[0,GB.Row] ;
    msbGene : St:=SectAttente ;
    msbGeneSect : begin
      St := BE_BUDSECT.Text;
      if St='' then St:=SectAttente;
    end;
  end;
  if St<>'' then FicheBudSect(Nil,BE_AXE.Value,St,taConsult,0);
END ;

procedure TFSaisBud.ClickJournal ;
BEGIN
if Not BZoomJournal.Enabled then Exit ;
if BE_BUDJAL.Value='' then Exit ;
FicheBudjal(Nil,BE_AXE.Value,BE_BUDJAL.Value,taConsult,0);
END ;

{================================ VALIDATION =================================}
procedure TFSaisBud.BValideClick(Sender: TObject);
begin ClickValide(True) ; end;

function TFSaisBud.PieceCorrecte : boolean ;
BEGIN
Result:=False ;
if Not JalOk then Exit ;
if Not NatureOk then Exit ;
if ModeSaisBud=msbGeneSect then if Not SectAttOk then Exit ;
if ModeSaisBud=msbSectGene then if Not GeneAttOk then Exit ;
Result:=True ;
END ;

procedure TFSaisBud.DetruitAncien ;
Var SQL : String ;
    Nb  : integer ;
BEGIN
if DateCoher then
   BEGIN
   SQL:='DELETE FROM BUDECR WHERE '+WhereEcriture(tsBudget,M,False)+' AND BE_DATEMODIF="'+UsTime(OldModif)+'"' ;
   END else
   BEGIN
   SQL:='DELETE FROM BUDECR WHERE '+WhereEcriture(tsBudget,M,False) ;
   END ;
Nb:=ExecuteSQL(SQL) ;
if Nb<>NbLigOrig then V_PGI.IoError:=oeSaisie ;
END ;

function TFSaisBud.SensCoher ( O : TOBM ; TBud,TSect : TStrings ) : boolean ;
Var Bud,Sect : String17 ;
    SB,SS    : String3 ;
    OkDD,OkCC  : Boolean ;
    St,SCpte,SSens : String ;
    i        : integer ;
BEGIN
Result:=True ;
Bud:=O.GetMvt('BE_BUDGENE') ; Sect:=O.GetMvt('BE_BUDSECT') ;
if ((Bud='') or (Sect='')) then Exit ;
SB:='M' ; SS:='M' ;
for i:=0 to TBud.Count-1 do
    BEGIN
    St:=TBud[i] ; SCpte:=ReadtokenSt(St) ; SSens:=ReadTokenSt(St) ;
    if SCpte=Bud then BEGIN SB:=SSens ; Break ; END ;
    END ;
for i:=0 to TSect.Count-1 do
    BEGIN
    St:=TSect[i] ; SCpte:=ReadtokenSt(St) ; SSens:=ReadTokenSt(St) ;
    if SCpte=Sect then BEGIN SS:=SSens ; Break ; END ;
    END ;
OkDD:=O.GetDebitSaisi<>0 ; OkCC:=O.GetCreditSaisi<>0 ;
if OkDD then if ((SB='C') or (SS='C')) then Result:=False ;
if OkCC then if ((SB='D') or (SS='D')) then Result:=False ;
END ;

Function TFSaisBud.ControleLeSens : Boolean ;
Var ACol,ARow   : integer ;
    T           : TList ;
    O           : TOBM ;
    i           : integer ;
    TBud,TSect  : TStrings ;
    St,StC,StS  : String ;
    Q           : TQuery ;
BEGIN
Result:=True ; SensCorrect:=True ;
if Not OkControle then Exit ;
{Chargement des infos}
TBud:=TStringList.Create ; TSect:=TStringList.Create ;
if ModeSaisBud=msbSectGene then St:=BE_BUDGENE.Text+';' else St:=LesComptes ;
Repeat
 StC:=ReadtokenSt(St) ;
 if StC<>'' then
    BEGIN
    Q:=OpenSQL('Select BG_SENS from BUDGENE Where BG_BUDGENE="'+StC+'"',True) ;
    if Not Q.EOF then BEGIN StS:=Q.Fields[0].AsString ; StS:=StC+';'+StS+';' ; TBud.Add(StS) ; END ;
    Ferme(Q) ;
    END ;
Until ((St='') or (StC='')) ;
if ModeSaisBud=msbGeneSect then St:=BE_BUDSECT.Text+';' else St:=LesSections ;
Repeat
 StC:=ReadtokenSt(St) ;
 if StC<>'' then
    BEGIN
    Q:=OpenSQL('Select BS_SENS from BUDSECT Where BS_BUDSECT="'+StC+'"',True) ;
    if Not Q.EOF then BEGIN StS:=Q.Fields[0].AsString ; StS:=StC+';'+StS+';' ; TSect.Add(StS) ; END ;
    Ferme(Q) ;
    END ;
Until ((St='') or (StC='')) ;
{Lecture de la saisie}
for ACol:=1 to GB.ColCount-1 do for ARow:=1 to GB.RowCount-1 do
    BEGIN
    T:=GetL(ACol,ARow) ; if T=Nil then Continue ;
    for i:=0 to T.Count-1 do
        BEGIN
        O:=TOBM(T[i]) ; if O=Nil then Continue ;
        if ((O.GetDebitSaisi<>0) or (O.GetCreditSaisi<>0)) then
           BEGIN
           if SensCorrect then if Not SensCoher(O,TBud,TSect) then BEGIN SensCorrect:=False ; Break ; END ;
           END ;
        END ;
    if Not SensCorrect then Break ;
    END ;
if Not SensCorrect then
   BEGIN
   if HM.Execute(8,'','')<>mrYes then Result:=False ;
   END ;
TBud.Free ; TSect.Free ;
END ;

procedure TFSaisBud.ValideCells ;
Var ACol,ARow,i : integer ;
    T           : TList ;
    O           : TOBM ;
  TBudEcr : Tob;
BEGIN
  for ACol:=1 to GB.ColCount-1 do for ARow:=1 to GB.RowCount-1 do begin
    T:=GetL(ACol,ARow);
    if T=Nil then Continue;
    for i:=0 to T.Count-1 do begin
      O:=TOBM(T[i]) ; if O=Nil then Continue ;
      if ((O.GetDebitSaisi<>0) or (O.GetCreditSaisi<>0)) then begin
        TBudEcr := TOB.Create('BUDECR', nil, -1) ;
        TBudEcr.InitValeurs ;
        O.PutMvt('BE_NATUREBUD',BE_NATUREBUD.Value) ;
        TBudEcr.Dupliquer(O, False, True);
        TBudEcr.PutValue('BE_BLOCNOTE', O.M.Text);
        TBudEcr.PutValue('BE_ETABLISSEMENT', BE_ETABLISSEMENT.Value);
        TBudEcr.PutValue('BE_DATEMODIF', NowFutur);
        TBudEcr.PutValue('BE_DATECREATION', Date);
        TBudEcr.PutValue('BE_TYPESAISIE', VUE.Value);
        TBudEcr.PutValue('BE_RESOLUTION', RESOL.Value);
        TBudEcr.InsertDB(nil) ;
        TBudEcr.Free ;
      end;
    end;
  end;
  AvertirTable('BUDECR');
end;

procedure TFSaisBud.CloseFen ;
begin
  Close ;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end ;

procedure TFSaisBud.ValideNumero ;
BEGIN
SetIncNum(EcrBud,GeneSouche,NumeroPiece,0) ;
END ;

procedure TFSaisBud.ValideLaPiece ;
BEGIN
if Action=taCreat then ValideNumero ;
if Action=taModif then DetruitAncien ;
if V_PGI.IoError=oeOk then ValideCells ;
END ;

procedure TFSaisBud.ClickValide ( OkFerme : boolean ) ;
Var io : TIoErr ;
BEGIN
if Action=taConsult then BEGIN if OkFerme then CloseFen ; Exit ; END ;
if Not PieceModifiee then BEGIN if OkFerme then CloseFen ; Exit ; END ;
if Not PieceCorrecte then Exit ;
if Not ControleLeSens then Exit ;
NowFutur:=NowH ;
io:=Transactions(ValideLaPiece,3) ;
Case io of
   oeOk : ;
   oeunknown : MessageAlerte(HDiv.Mess[5]) ;
   oeSaisie : MessageAlerte(HDiv.Mess[4]) ;
   END ;
PieceModifiee:=False ;
if ((Action=taCreat) {and (ModeSaisBud in [msbGeneSect,msbSectGene])} and   {FP FQ16153 A faire quelque soit le type de saisie}
    (io=oeOk) and (OkFerme) and (Not VisuConso)) then
    BEGIN
    CloseSaisie ; CreateSaisie ;
    BE_BUDJALChange(Nil) ; BE_BUDJAL.Enabled:=True ; BE_BUDJAL.SetFocus ;
    END else
    BEGIN
    if OkFerme then CloseFen else if Action=taCreat then
       BEGIN
       Action:=taModif ; {Remplir le M comme pour une modif}
       M.Num:=NumeroPiece ; M.Jal:=BE_BUDJAL.Value ;
       M.Nature:=BE_NATUREBUD.Value ; M.Axe:=BE_AXE.Value ;
       END ;
    END ;
END ;

Function TFSaisBud.JalOk : Boolean ;
BEGIN
Result:=False ;
if BE_BUDJAL.Value='' then BEGIN BE_BUDJAL.SetFocus ; HM.Execute(4,'','') ; Exit ;  END ;
Result:=True ;
END ;

Function TFSaisBud.NatureOk : Boolean ;
BEGIN
Result:=False ;
if BE_NATUREBUD.Value='' then BEGIN BE_NATUREBUD.SetFocus ; HM.Execute(5,'','') ; Exit ;  END ;
Result:=True ;
END ;

{JP 19/04/07 : FQ 19920 : Ajout du CanFocus : en fait il est inutile de faire un Focus
               car la zone est Grisé !!!!
{---------------------------------------------------------------------------------------}
function TFSaisBud.SectAttOk : Boolean ;
{---------------------------------------------------------------------------------------}
begin
  Result:=False ;
  if BE_BUDSECT.ExisteH<=0 then begin
    if BE_BUDSECT.CanFocus then BE_BUDSECT.SetFocus;
    HM.Execute(6,'','');
    Exit;
  end;

  if Pos(BE_BUDSECT.Text + ';', LesSections) <= 0 then begin
    if BE_BUDSECT.CanFocus then BE_BUDSECT.SetFocus;
    HM.Execute(6,'','');
    Exit;
  end ;
  Result:=True ;
end;

{---------------------------------------------------------------------------------------}
function TFSaisBud.GeneAttOk : Boolean ;
{---------------------------------------------------------------------------------------}
begin
  Result:=False ;
  if BE_BUDGENE.ExisteH <= 0 then begin
    if BE_BUDGENE.CanFocus then BE_BUDGENE.SetFocus;
    HM.Execute(7,Caption,'');
    Exit;
  end;

  if Pos(BE_BUDGENE.Text + ';', LesComptes) <= 0 then begin
    if BE_BUDGENE.CanFocus then BE_BUDGENE.SetFocus;
    HM.Execute(7, Caption, '');
    Exit;
  end;

  Result:=True;
end ;

{======================= POSITIONNEMENTS, AFFICHAGES ========================}
Procedure TFSaisBud.FormateZone ( ACol,ARow : integer ) ;
Var X : Double ;
BEGIN
X:=Valeur(GB.Cells[ACol,ARow]) ;
{JP 25/08/05 : FQ 16380 : Ajout du Trim pour éviter certains espaces intempestifs}
GB.Cells[ACol,ARow]:= Trim(StrS(X,CurDec));
END ;

function TFSaisBud.ColToDate ( Col : integer ) : TDateTime ;
var
  ind : integer ;
begin
  if Col <= PERDEB.Items.Count - PERDEB.ItemIndex then begin
    ind := PERDEB.ItemIndex + Col - 1;
    Result := StrToDate(PERDEB.Values[ind]);
  end
  else begin
    {Le budget est sur deux exercices qui se suivent ...}
    if ExosInter.Values.Count <= 0 then begin
      {JP 26/12/07 : FQ 22110 : Si le budget est à cheval sur deux exercices, c'est à dire que
                     PERDEB et PERFIN on un contenu différent
       Exemple : Budget sur exo 2007 et 2008 avec 2007 (01/01/07 - > 31/01/07) et 2008 (01/02/08 et 31/12/08)
                 mais le budget est sur 2008 (01/01/08 - 31/12/08).
                 Alors PERDEB.ItemIndex = 12, PERDEB.Items.Count = 13

       Donc si on passe ici, c'est que la colonne correspond à une période du deuxième exercice, il faut donc
       déduire de l'indice de la colonne les périodes du premier exercice}
      Ind := Col + PERDEB.ItemIndex - PERDEB.Items.Count - 1;
      Result := StrToDate(PERFIN.Values[ind]);
    end

    {... Le budget est sur plus de deux exercices}
    else begin
      if Col > (PERDEB.Items.Count - PERDEB.ItemIndex) + ExosInter.Values.Count then begin
        Ind:=Col-(PERDEB.Items.Count-PERDEB.ItemIndex)-EXOSINTER.Items.Count-1 ;
        Result:=StrToDate(PERFIN.Values[ind]) ;
      end
      else begin
         Ind:=Col-(PERDEB.Items.Count-PERDEB.ItemIndex)-1 ;
         Result:=StrToDate(ExosInter.Values[ind]) ;
      end;
    end;
  end;
end;

function TFSaisBud.QuellePer ( DD : TDateTime ) : integer ;
Var ii : integer ;
BEGIN
Result:=1 ;
ii:=PERDEB.Values.IndexOf(DateToStr(DD)) ;
if ii>=0 then
   BEGIN
   QuellePer:=ii-PERDEB.ItemIndex+1 ;
   END else
   BEGIN
   if ExosInter.Values.Count<=0 then
      BEGIN
      ii:=PERFIN.Values.IndexOf(DateToStr(DD)) ;
      if ii>=0 then QuellePer:=(ii+1)+(PERDEB.Items.Count-PERDEB.ItemIndex) ;
      END else
      BEGIN
      ii:=EXOSINTER.Values.IndexOf(DateToStr(DD)) ;
      if ii>=0 then QuellePer:=(PERDEB.Items.Count-PERDEB.ItemIndex)+(ii+1) else
         BEGIN
         ii:=PERFIN.Values.IndexOf(DateToStr(DD)) ;
         if ii>=0 then QuellePer:=ii+(PERDEB.Items.Count-PERDEB.ItemIndex)+EXOSINTER.Items.Count+1 ;
         END ;
      END ;
   END ;
END ;

Procedure TFSaisBud.TrouveCase ( O : TOBM ; Var ACol,ARow : integer ) ;
BEGIN
ACol:=0 ; ARow:=0 ;
if ModeSaisBud in [msbGene,msbGeneSect] then
   BEGIN
   ARow:=GB.Cols[0].IndexOf(O.GetMvt('BE_BUDGENE')) ;
   END else
   BEGIN
   ARow:=GB.Cols[0].IndexOf(O.GetMvt('BE_BUDSECT')) ;
   END ;
ACol:=QuellePer(O.GetMvt('BE_DATECOMPTABLE')) ;
END ;

Procedure TFSaisBud.ShowCase ( ACol,ARow : integer ) ;
Var X : Double ;
BEGIN
X:=SommeCase(ACol,ARow) ;
{JP 25/08/05 : FQ 16380 : Ajout du Trim pour éviter certains espaces intempestifs}
GB.Cells[ACol,ARow] := Trim(StrS(X / CurCoef, CurDec));
END ;

procedure TFSaisBud.SommeTout ;
Var i : integer ;
    X : Double ;
BEGIN
X:=0 ;
for i:=1 to GBas.ColCount-1 do X:=X+Valeur(GBas.Cells[i,0]) ;
TG.Caption:=StrS(X,CurDec) ;
END ;

procedure TFSaisBud.SommeLigne ( ARow : integer ) ;
Var i : integer ;
    X : Double ;
BEGIN
X:=0 ;
for i:=1 to GB.ColCount-1 do X:=X+Valeur(GB.Cells[i,ARow]) ;
GDroit.Cells[0,ARow]:=StrS(X,CurDec) ;
END ;

procedure TFSaisBud.SommeCol ( ACol : integer ) ;
Var i : integer ;
    X : double ;
BEGIN
X:=0 ;
for i:=1 to GB.RowCount-1 do X:=X+Valeur(GB.Cells[ACol,i]) ;
GBas.Cells[ACol,0]:=StrS(X,CurDec) ;
END ;

Function TFSaisBud.SommeCase ( ACol,ARow : integer ) : double ;
Var X : Double ;
    i : integer ;
    O : TOBM ;
    T : TList ;
BEGIN
X:=0 ; T:=GetL(ACol,ARow) ;
if T<>Nil then
   BEGIN
   for i:=0 to T.Count-1 do
       BEGIN
       O:=TOBM(T[i]) ;
       if OkD then X:=X+O.GetDebitSaisi else X:=X+O.GetCreditSaisi ;
       END ;
   END ;
Result:=Arrondi(X,Deci) ;
END ;

Procedure TFSaisBud.RemplitCase ( O : TOBM ) ;
Var ACol,ARow : integer ;
BEGIN
TrouveCase(O,ACol,ARow) ; if ((ACol<=0) or (ARow<=0)) then Exit ;
CreerObjetCase(ACol,ARow) ;
AddOBMCase(ACol,ARow,O) ;
ShowCase(ACol,ARow) ;
END ;

{============================== OBJETS ====================================}
Function TFSaisBud.GetL ( ACol,ARow : integer ) : TList ;
BEGIN
Result:=TList(GB.Objects[ACol,ARow]) ;
END ;

procedure TFSaisBud.CreerNewOBM ( ACol,ARow : integer ) ;
Var O : TOBM ;
    DD : TDateTime ;
BEGIN
O:=TOBM.Create(EcrBud,'',True) ;
O.PutMvt('BE_BUDJAL',BE_BUDJAL.Value)  ; O.PutMvt('BE_NATUREBUD',BE_NATUREBUD.Value) ;
O.PutMvt('BE_NUMEROPIECE',NumeroPiece) ; O.PutMvt('BE_QUALIFPIECE',M.Simul) ;
O.PutMvt('BE_AXE',BE_AXE.Value)        ; O.PutMvt('BE_ETABLISSEMENT',BE_ETABLISSEMENT.Value) ;
DD:=ColToDate(ACol) ;
O.PutMvt('BE_DATECOMPTABLE',DD) ;
O.PutMvt('BE_EXERCICE',QuelExoDT(DD)) ;
Case ModeSaisBud of
     msbGene : BEGIN
               O.PutMvt('BE_BUDGENE',GB.Cells[0,ARow]) ;
               O.PutMvt('BE_BUDSECT',SectAttente) ;
               END ;
 msbGeneSect : BEGIN
               O.PutMvt('BE_BUDGENE',GB.Cells[0,ARow]) ;
               O.PutMvt('BE_BUDSECT',BE_BUDSECT.Text) ;
               END ;
     msbSect : BEGIN
               O.PutMvt('BE_BUDGENE',GeneAttente) ;
               O.PutMvt('BE_BUDSECT',GB.Cells[0,ARow]) ;
               END ;
 msbSectGene : BEGIN
               O.PutMvt('BE_BUDGENE',BE_BUDGENE.Text) ;
               O.PutMvt('BE_BUDSECT',GB.Cells[0,ARow]) ;
               END ;
   END ;
AddOBMCase(ACol,ARow,O) ;
END ;

Procedure TFSaisBud.FreeCopy ;
Var i : integer ;
    T : TList ;
BEGIN
for i:=0 to ColCopy.Count-1 do
    BEGIN
    T:=TList(ColCopy.Objects[i]) ;
    if T<>Nil then BEGIN VideListe(T) ; T.Free ; ColCopy.Objects[i]:=Nil ; END ;
    END ;
ColCopy.Clear ; ColCopy.Add('') ;
END ;

Procedure TFSaisBud.FreeObjetCase ( ACol,ARow : integer ) ;
Var T : TList ;
BEGIN
T:=GetL(ACol,ARow) ; if T=Nil then Exit ;
VideListe(T) ; T.Free ; GB.Objects[ACol,ARow]:=Nil ;
GB.Cells[ACol,ARow]:='' ;
END ;

Procedure TFSaisBud.CreerObjetCase ( ACol,ARow : integer ) ;
Var T : TList ;
BEGIN
{JP 25/08/05 : FQ 16380 : Ajout du Trim pour éviter certains espaces intempestifs}
if GB.Cells[ACol,ARow]='' then GB.Cells[ACol,ARow] := Trim(StrS(0,CurDec));
T:=GetL(ACol,ARow) ; if T<>Nil then Exit ;
T:=TList.Create ; GB.Objects[ACol,ARow]:=T ;
END ;

Procedure TFSaisBud.AddOBMCase ( ACol,ARow : integer ; O : TOBM ) ;
Var T : TList ;
BEGIN
if O=Nil then Exit ;
T:=GetL(ACol,ARow) ; T.Add(O) ;
END ;

{============================== CHARGEMENTS ===============================}
Procedure TFSaisBud.ChargeLesMouvements ;
Var Q : TQuery ;
    SQL : String ;
    O   : TOBM ;
    Premier : boolean ;
    ACol,ARow : integer ;
{$IFDEF EAGLCLIENT}
    BlocNote : TStringList;
{$ENDIF}
BEGIN
if Not VisuConso then
   BEGIN
   SQL:='Select * from BUDECR Where BE_BUDJAL="'+BE_BUDJAL.Value+'" AND BE_NATUREBUD="'+BE_NATUREBUD.Value+'"'
       +' AND BE_QUALIFPIECE="'+M.Simul+'" AND BE_NUMEROPIECE='+IntToStr(M.Num) ;
   END else
   BEGIN
   SQL:='Select * from BUDECR Where BE_BUDJAL="'+BE_BUDJAL.Value+'" AND BE_QUALIFPIECE="'+M.Simul+'"' ;
   if BE_ETABLISSEMENT.Value<>'' then SQL:=SQL+' AND BE_ETABLISSEMENT="'+BE_ETABLISSEMENT.Value+'"' ;
   END ;
Q:=OpenSQL(SQL,True) ; Premier:=True ; NbLigOrig:=0 ;
While Not Q.EOF do
   BEGIN
   O:=TOBM.Create(EcrBud,'',False) ; O.ChargeMvt(Q) ;
{$IFDEF EAGLCLIENT}
    BlocNote := TStringList.Create;
    BlocNote.Text := Q.FindField('BE_BLOCNOTE').AsString;
	  if not (Q.FindField('BE_BLOCNOTE').AsString = '') then  O.M.Assign(BlocNote);
    FreeAndNil(BlocNote) ;
{$ELSE}
   if Not IsFieldNull(Q,'BE_BLOCNOTE') then O.M.Assign(TMemoField(Q.FindField('BE_BLOCNOTE'))) ;
{$ENDIF}

   if Premier then
      BEGIN
      OldModif:=O.GetMvt('BE_DATEMODIF') ;
      RESOL.Value:=O.GetMvt('BE_RESOLUTION') ;
      if ModeSaisBud=msbGeneSect then BE_BUDSECT.Text:=O.GetMvt('BE_BUDSECT') else
      if ModeSaisBud=msbSectGene then BE_BUDGENE.Text:=O.GetMvt('BE_BUDGENE') ;
      END else
      BEGIN
      if ((Action=taModif) and (O.GetMvt('BE_DATEMODIF')<>OldModif)) then DateCoher:=False ;
      END ;
   RemplitCase(O) ; Inc(NbLigOrig) ;
   Q.Next ; Premier:=False ;
   END ;
Ferme(Q) ;
for ACol:=1 to GB.ColCount-1 do SommeCol(ACol) ;
for ARow:=1 to GB.RowCount-1 do SommeLigne(ARow) ;
SommeTout ;
END ;

Procedure TFSaisBud.ChargeLesPeriodes ;
Var Cur : integer ;
    St  : String ;
    DD  : TDateTime ;
BEGIN
NbPer:=0 ;
if BE_EXERCICE.Value=BE_EXERCICE_.Value then
   BEGIN
   Cur:=PERDEB.ItemIndex ;
   While Cur<=PERFIN.ItemIndex do
      BEGIN
      DD:=StrToDate(PERDEB.Values[Cur]) ;
      St:=FormatDateTime('mmmm yy',DD) ;
      {JP 25/08/05 : FQ 16380 : Ajout du Trim pour éviter certains espaces intempestifs}
      GB.Cells[Cur-PERDEB.ItemIndex+1,0]:= Trim(FirstMajuscule(St));
      Inc(Cur) ; Inc(NbPer) ;
      END ;
   END else
   BEGIN
   for Cur:=PERDEB.ItemIndex to PERDEB.Items.Count-1 do
       BEGIN
       DD:=StrToDate(PERDEB.Values[Cur]) ;
       St:=FormatDateTime('mmmm yy',DD) ;
       {JP 25/08/05 : FQ 16380 : Ajout du Trim pour éviter certains espaces intempestifs}
       GB.Cells[Cur-PERDEB.ItemIndex+1,0]:= Trim(FirstMajuscule(St)) ;
       Inc(NbPer) ;
       END ;
   for Cur:=0 to EXOSINTER.Items.Count-1 do
       BEGIN
       DD:=StrToDate(EXOSINTER.Values[Cur]) ;
       St:=FormatDateTime('mmmm yy',DD) ;
       {JP 25/08/05 : FQ 16380 : Ajout du Trim pour éviter certains espaces intempestifs}
       GB.Cells[Cur-PERDEB.ItemIndex+1+PERDEB.Items.Count,0]:= Trim(FirstMajuscule(St));
       Inc(NbPer) ;
       END ;
   for Cur:=0 to PERFIN.ItemIndex do
       BEGIN
       DD:=StrToDate(PERFIN.Values[Cur]) ;
       St:=FormatDateTime('mmmm yy',DD) ;
       {JP 25/08/05 : FQ 16380 : Ajout du Trim pour éviter certains espaces intempestifs}
       GB.Cells[Cur+(PERDEB.Items.Count-PERDEB.ItemIndex+1)+EXOSINTER.Items.Count,0]:= Trim(FirstMajuscule(St));
       Inc(NbPer) ;
       END ;
   END ;
GB.ColCount:=1+NbPer ;
GDroit.Cells[0,0]:=HDiv.Mess[0] ;
GBas.ColCount:=GB.ColCount ;
For Cur:=0 to GB.ColCount - 1 do GBas.ColWidths[cur] := GB.ColWidths[cur] ;
AjusteAscenceurs ;
END ;

Procedure TFSaisBud.ChargeLesComptes ;
Var Nb : integer ;
    StC,St : String ;
    OkS    : boolean ;
BEGIN
OkS:=ModeSaisBud in [msbSect,msbSectGene] ;
if OkS then St:=LesSections else St:=LesComptes ;
Repeat
 StC:=ReadTokenSt(St) ;
 if StC<>'' then               
    BEGIN
    {JP 25/08/05 : FQ 16380 : Ajout du Trim pour éviter certains espaces intempestifs}
    GB.Cells[0,GB.RowCount-1]:= Trim(StC);
    GB.RowCount:=GB.RowCount+1 ;
    END ;
Until ((St='') or (StC='')) ;
if OkS then GB.Cells[0,0]:=HDiv.Mess[1] else GB.Cells[0,0]:=HDiv.Mess[2] ;
GB.RowCount:=GB.RowCount-1 ;
GBas.Cells[0,0]:=HDiv.Mess[3] ;
GDroit.RowCount:=GB.RowCount ;
For nb:=0 to GB.RowCount - 1 do GDroit.RowHeights[nb] := GB.RowHeights[nb] ;
AjusteAscenceurs ;
END ;

{================================ ENTETE ==================================}
procedure TFSaisBud.BE_EXERCICEChange(Sender: TObject);
begin
if BE_EXERCICE.Value='' then Exit ;
ListePeriode(BE_EXERCICE.Value,PERDEB.Items,PERDEB.Values,True) ;
end;

procedure TFSaisBud.BE_EXERCICE_Change(Sender: TObject);
begin
if BE_EXERCICE_.Value='' then Exit ;
ListePeriode(BE_EXERCICE_.Value,PERFIN.Items,PERFIN.Values,True) ;
end;

procedure TFSaisBud.RempliExosInter ;
Var Q : TQuery ;
    FinPremier,DebDernier : TDateTime ;
    SQL  : String ;
    Exo  : String3 ;
    i,Nb : integer ;
BEGIN
FinPremier:=0 ; DebDernier:=0 ; Nb:=0 ;
ExosInter.Items.Clear ; ExoCalcul.Values.Clear ;
ExoCalcul.Items.Clear ; ExosInter.Values.Clear ;
Q:=OpenSQL('SELECT EX_DATEFIN, EX_DATEDEBUT From EXERCICE Where EX_EXERCICE="'+BE_EXERCICE.Value+'"',True) ;
if Not Q.EOF then
   BEGIN
   FinPremier:=Q.Fields[0].AsDateTime ;
   DateDebExos[1]:=Q.Fields[1].AsDateTime ;
   END ;
Ferme(Q) ;
Q:=OpenSQL('SELECT EX_DATEDEBUT From EXERCICE Where EX_EXERCICE="'+BE_EXERCICE_.Value+'"',True) ;
if Not Q.EOF then
   BEGIN
   DebDernier:=Q.Fields[0].AsDateTime ;
   DateDebExos[2]:=Q.Fields[0].AsDateTime ;
   END ;
Ferme(Q) ;
if ((FinPremier=0) or (DebDernier=0)) then Exit ;
if DebDernier-FinPremier<=1 then Exit ;
SQL:='Select EX_EXERCICE, EX_DATEDEBUT from EXERCICE Where EX_EXERCICE<>"'+BE_EXERCICE.Value+'"'
    +' AND EX_EXERCICE<>"'+BE_EXERCICE_.Value+'" AND EX_DATEDEBUT>"'+USDateTime(FinPremier)+'"'
    +' AND EX_DATEFIN<"'+UsDateTime(DebDernier)+'"'
    +' ORDER BY EX_DATEDEBUT' ;
Q:=OpenSQL(SQL,True) ;
While Not Q.EOF do
   BEGIN
   Exo:=Q.Fields[0].AsString ; ListePeriode(Exo,ExoCalcul.Items,ExoCalcul.Values,True) ;
   Inc(Nb) ; DateDebExos[Nb+2]:=Q.Fields[1].AsDateTime ;
   for i:=0 to ExoCalcul.Items.Count-1 do
       BEGIN
       ExosInter.Items.Add(ExoCalcul.Items[i]) ;
       ExosInter.Values.Add(ExoCalcul.Values[i]) ;
       END ;
   Q.Next ;
   END ;
Ferme(Q) ;
END ;

(*Procedure TFSaisBud.ChargeLesLib(St : String ; OnCpt : Boolean ; Var LesLib : String);
Var St1,St2,SQL : String ;
    Q : TQuery ;
BEGIN
If LesLib<>'' Then Exit ;
St1:=St ; LesLib:='' ;
While St1<>'' Do
  BEGIN
  St2:=ReadTokenSt(St1) ;
  If OnCpt Then SQL:='SELECT BG_LIBELLE FROM BUDGENE WHERE BG_BUDGENE="'+St2+'" '
           Else SQL:='SELECT BS_LIBELLE FROM BUDSECT WHERE BS_BUDSECT="'+St2+'" ' ;
  Q:=OpenSQL(SQL,TRUE);
  If Not Q.Eof Then LesLib:=LesLib+Q.Fields[0].AsString+';' Else LesLib:=LesLib+St2+';' ;
  Ferme(Q) ;
  END ;
END ;*)

procedure TFSaisBud.ChargeInfosJournal ;
Var Q  : TQuery ;
    MM : String17 ;
    i  : integer ;
BEGIN
MM:='' ; GeneSouche:='' ; GeneSens:='M' ;
Q:=OpenSQL('Select * from BUDJAL Where BJ_BUDJAL="'+BE_BUDJAL.Value+'"',True) ;
if Not Q.EOF then
   BEGIN
   BE_EXERCICE.Value:=Q.FindField('BJ_EXODEB').AsString ;
   BE_EXERCICE_.Value:=Q.FindField('BJ_EXOFIN').AsString ;
   PERDEB.Value:=DateToStr(DebutDeMois(Q.FindField('BJ_PERDEB').AsDateTime)) ;
   PERFIN.Value:=DateToStr(DebutDeMois(Q.FindField('BJ_PERFIN').AsDateTime)) ;
   RempliExosInter ;
   BE_AXE.Value:=Q.FindField('BJ_AXE').AsString ;
   CatBudget:=Q.FindField('BJ_CATEGORIE').AsString ;
   SPCat:=Q.FindField('BJ_SOUSPLAN').AsString ;
   GeneSens:=Q.FindField('BJ_SENS').AsString ;
   OkControle:=(Q.FindField('BJ_CONTROLESENS').AsString='X') ;
   GeneAttente:=Q.FindField('BJ_GENEATTENTE').AsString ; SectAttente:=Q.FindField('BJ_SECTATTENTE').AsString ;
   if Q.FindField('BJ_BUDGENES2').AsString<>'' then LesComptes:=Q.FindField('BJ_BUDGENES').AsString+Q.FindField('BJ_BUDGENES2').AsString+GeneAttente+';'
                                               else LesComptes:=Q.FindField('BJ_BUDGENES').AsString+GeneAttente+';' ;
   if Q.FindField('BJ_BUDSECTS2').AsString<>'' then LesSections:=Q.FindField('BJ_BUDSECTS').AsString+Q.FindField('BJ_BUDSECTS2').AsString+SectAttente+';'
                                               else LesSections:=Q.FindField('BJ_BUDSECTS').AsString+SectAttente+';' ;
   for i:=Length(LesComptes) Downto 1 do if LesComptes[i]=#0 then Delete(LesComptes,i,1) ;
   for i:=Length(LesSections) Downto 1 do if LesSections[i]=#0 then Delete(LesSections,i,1) ;
   if Action=taCreat then
      BEGIN
      if M.Simul<>'N' then GeneSouche:=Q.FindField('BJ_COMPTEURSIMUL').AsString else GeneSouche:=Q.FindField('BJ_COMPTEURNORMAL').AsString ;
      NumeroPiece:=GetNum(EcrBud,GeneSouche,MM,0) ;
      if ((BZoomSousPlan.Visible) and (BE_BUDSECT.Text<>'') and (ModeSaisBud=msbGeneSect))
         then else BE_BUDSECT.Text:=SectAttente ;
      BE_BUDGENE.Text:=GeneAttente ;
      if NumeroPiece>0 then BE_NUMEROPIECE.Caption:=IntToStr(NumeroPiece) ;
      END ;
   (*
   LesLibComptes:=ChargeLesLib(LesComptes,TRUE) ;
   LesLibSections:=ChargeLesLib(LesSections,FALSE) ;
   *)
   END ;
Ferme(Q) ;
if GeneSens='C' then CSens.ItemIndex:=1 else CSens.ItemIndex:=0 ;
CSensChange(Nil) ;
END ;

procedure TFSaisBud.TransformeGrille ;
Var i,j,k : integer ;
    T     : TList ;
    O     : TOBM ;
    XD,XC : Double ;
BEGIN
for i:=1 to GB.RowCount-1 do for j:=1 to GB.ColCount-1 do
    BEGIN
    T:=GetL(j,i) ; if T=Nil then Continue ;
    for k:=0 to T.Count-1 do
        BEGIN
        O:=TOBM(T[k]) ; if O=Nil then Continue ;
        XD:=O.GetDebitSaisi ; XC:=O.GetCreditSaisi ;
        XD:=Arrondi(XD/CurCoef,CurDec)*CurCoef ; XC:=Arrondi(XC/CurCoef,CurDec)*CurCoef ;
        O.SetMontantsBUD(XD,XC) ;
        ShowCase(j,i) ;
        END ;
    END ;
for i:=1 to GB.RowCount-1 do SommeLigne(i) ;
for j:=1 to GB.ColCount-1 do SommeCol(j) ;
SommeTout ;
END ;

procedure TFSaisBud.RESOLChange(Sender: TObject);
begin
RESOL.Tag:=RESOL.ItemIndex ;
if RESOL.Value='C' then BEGIN CurDec:=Deci ; CurCoef:=1.0 ; END else
if RESOL.Value='F' then BEGIN CurDec:=0 ; CurCoef:=1.0 ; END else
if RESOL.Value='K' then BEGIN CurDec:=0 ; CurCoef:=1000.0 ; END else
if RESOL.Value='M' then BEGIN CurDec:=0 ; CurCoef:=1000000.0 ; END ;
if Action=taConsult then Exit ;
if GeneCharge then Exit ;
if PgiAsk(TraduireMemoire('En changeant l''unité, vous risquez de perdre la précision de votre saisie.') + #13#13 +
          TraduireMemoire('Souhaitez vous poursuivre ?'), Caption) = mrYes then
  TransformeGrille
else begin
  RESOL.OnChange := nil;
  RESOL.ItemIndex := RESOL.Tag;
  RESOL.OnChange := RESOLChange;
end;
end;

procedure TFSaisBud.VUEChange(Sender: TObject);
Var ms : TModeSaisBud ;
begin
ms:=msbGene ; VUE.Tag:=VUE.ItemIndex ;
if VUE.Value='' then Exit ;
if VUE.Value='G'  then ms:=msbGene else
if VUE.Value='S'  then ms:=msbSect else
if VUE.Value='GS' then ms:=msbGeneSect else
if VUE.Value='SG' then ms:=msbSectGene ;
if ms=ModeSaisBud then Exit ;
if Not GeneCharge then
   BEGIN
   ClickValide(False) ; GB.VidePile(True) ;
   ModeSaisBud:=ms ; ShowBudget ; PasOkVUE:=True ;
   END ;
ModeSaisBud:=ms ; CurCopy:=0 ;
AttribLeTitre ; 
end;

Procedure TFSaisBud.AttribLeTitre ;
Var Num : integer ;
BEGIN
Num:=10 ;
if Action=taCreat then
   BEGIN
   Case ModeSaisBud of
      msbGene     : BEGIN Num:=10 ; HelpContext:=15211000 ; END ;
      msbSect     : BEGIN Num:=14 ; HelpContext:=15213000 ; END ;
      msbGeneSect : BEGIN Num:=15 ; HelpContext:=15217000 ; END ;
      msbSectGene : BEGIN Num:=16 ; HelpContext:=15219000 ; END ;
      END ;
   END else if Action=taModif then
   BEGIN
   Num:=9 ; HelpContext:=15230100 ;
   END else
   BEGIN
   if Not VisuConso then
      BEGIN
      Num:=8 ; HelpContext:=15221100 ;
      END else
      BEGIN
      if ModeSaisBud in [msbGene,msbGeneSect] then
         BEGIN
         Num:=11 ; HelpContext:=15225000 ;
         END else
         BEGIN
         Num:=12 ; HelpContext:=15227000 ;
         END ;
      END ;
   END ;
Caption:=HDiv.Mess[Num] ;
UpdateCaption(Self) ;
END ;

procedure TFSaisBud.ChangeImplicite ;
Var ic,ir,k : integer ;
    T       : TList ;
    O       : TOBM ;
BEGIN
for ic:=1 to GB.ColCount-1 do for ir:=1 to GB.RowCount-1 do
    BEGIN
    T:=GetL(ic,ir) ; if T=Nil then Continue ;
    for k:=0 to T.Count-1 do
        BEGIN
        O:=TOBM(T[k]) ; if O=Nil then Continue ;
        if ModeSaisBud=msbGeneSect then O.PutMvt('BE_BUDSECT',BE_BUDSECT.Text) else
        if ModeSaisBud=msbSectGene then O.PutMvt('BE_BUDGENE',BE_BUDGENE.Text) ;
        END ;
    END ;
PieceModifiee:=True ;
END ;

procedure TFSaisBud.BE_BUDSECTExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;
  if Action = taConsult then Exit ;
  if GeneCharge then Exit ;
  if ModeSaisBud <> msbGeneSect then Exit ;
  ChangeImplicite ;

  {JP 24/08/05 : FQ 16502 : Active la section qui a été désactivée dans ChangeImplicite}
  GereCodeImplicite(True);
end;

procedure TFSaisBud.BE_BUDGENEExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;
  if Action = taConsult then Exit ;
  if GeneCharge then Exit ;
  if ModeSaisBud <> msbSectGene then Exit ;
  ChangeImplicite ;
  {JP 24/08/05 : FQ 16501 : Active le compte qui a été désactivée dans ChangeImplicite}
  GereCodeImplicite(True);
end;

procedure TFSaisBud.BE_BUDJALChange(Sender: TObject);
Var ACol,ARow : integer ;
begin
if BE_BUDJAL.Value='' then Exit ;
VUE.Enabled:=((BE_BUDJAL.Value<>'') and (Not VisuConso) and (Action<>taConsult)) ;
if ModeSaisBud in [msbGeneSect,msbSectGene] then VUE.Enabled:=False ;
GB.VidePile(True) ;
ChargeInfosJournal ;
ChargeLesPeriodes ;
ChargeLesComptes ;
if Action=taCreat then
   BEGIN
   for ACol:=1 to GB.ColCount-1 do SommeCol(ACol) ;
   for ARow:=1 to GB.RowCount-1 do SommeLigne(ARow) ;
   SommeTout ;
   GB.Enabled:=True ;
   END else if ((Action=taConsult) and (VisuConso)) then
   BEGIN
   ChargeLesMouvements ;
   END ;
CurCopy:=0 ; GB.Invalidate ;
end;

procedure TFSaisBud.BE_AXEChange(Sender: TObject);
begin
if BE_AXE.Value='' then Exit ;
if BE_AXE.Value='A1' then BE_BUDSECT.ZoomTable:=tzBudSec1 else
if BE_AXE.Value='A2' then BE_BUDSECT.ZoomTable:=tzBudSec2 else
if BE_AXE.Value='A3' then BE_BUDSECT.ZoomTable:=tzBudSec3 else
if BE_AXE.Value='A4' then BE_BUDSECT.ZoomTable:=tzBudSec4 else
if BE_AXE.Value='A5' then BE_BUDSECT.ZoomTable:=tzBudSec5 ;
end;


procedure TFSaisBud.GereModeSaisie ;
BEGIN
Case ModeSaisBud of
     msbGene : BEGIN
               BE_BUDGENE.Visible:=False ; H_BUDGENE.Visible:=False ;
               BE_BUDSECT.Visible:=False ; H_BUDSECT.Visible:=False ;
               BE_AXE.Visible:=True      ; H_AXE.Visible:=True ;
               BVentil.Hint:=HDiv.Mess[6] ;
               END ;
     msbSect : BEGIN
               BE_BUDGENE.Visible:=False ; H_BUDGENE.Visible:=False ;
               BE_BUDSECT.Visible:=False ; H_BUDSECT.Visible:=False ;
               BE_AXE.Visible:=True      ; H_AXE.Visible:=True ;
               BVentil.Hint:=HDiv.Mess[7] ;
               END ;
 msbGeneSect : BEGIN
               BE_BUDGENE.Visible:=False ; H_BUDGENE.Visible:=False ;
               BE_BUDSECT.Visible:=True  ; H_BUDSECT.Visible:=True ;
               BE_AXE.Visible:=False     ; H_AXE.Visible:=False ;
               BVentil.Hint:=HDiv.Mess[6] ;
               END ;
 msbSectGene : BEGIN
               BE_BUDSECT.Visible:=False ; H_BUDSECT.Visible:=False ;
               BE_BUDGENE.Visible:=True  ; H_BUDGENE.Visible:=True  ;
               BE_AXE.Visible:=False     ; H_AXE.Visible:=False ;
               BVentil.Hint:=HDiv.Mess[7] ;
               END ;
   END ;
END ;

procedure TFSaisBud.GereComboVue ;
BEGIN
if Action=taConsult then BEGIN VUE.Enabled:=FALSE ; RESOL.Enabled:=False ; Exit ; END ;
if ModeSaisBud in [msbGeneSect,msbSectGene] then VUE.Enabled:=False else
   BEGIN
   VUE.Items.Delete(3) ; VUE.Values.Delete(3) ;
   VUE.Items.Delete(2) ; VUE.Values.Delete(2) ;
   END ;
END ;

procedure TFSaisBud.LectureSeule ;
Var i : integer ;
BEGIN
if Not VisuConso then
   BEGIN
   for i:=0 to PEntete.ControlCount-1 do
       if PEntete.Controls[i] is THValComboBox then THValComboBox(PEntete.Controls[i]).Enabled:=False ;
   END else
   BEGIN
   BE_BUDJAL.Enabled:=True ; BE_ETABLISSEMENT.Enabled:=True ;
   PositionneEtabUser(BE_ETABLISSEMENT, False); // 15087
   BevelNum.Visible:=False ; H_NUMEROPIECE.Visible:=False ; BE_NUMEROPIECE.Visible:=False ;
   BE_NATUREBUD.Visible:=False ; H_NATUREBUD.Visible:=False ;
   H_ETABLISSEMENT.Left:=H_NATUREBUD.Left ;
   BE_ETABLISSEMENT.Left:=BE_NATUREBUD.Left ;
   BE_ETABLISSEMENT.Width:=BE_NATUREBUD.Width ;
   PERFIN.Left:=PERDEB.Left ; PERFIN.Top:=BE_ETABLISSEMENT.Top ;
   H_PERFIN.Left:=H_PERDEB.Left+H_PERDEB.Width-H_PERFIN.Width-1 ; H_PERFIN.Top:=H_NATUREBUD.Top ;
   CSens.Top:=PERDEB.Top ;
   END ;
END ;

procedure TFSaisBud.DefautEntete ;
BEGIN
CatBudget:='' ; LesLibComptes:='' ; LesLibSections:='' ;
VUE.Value:=M.TypeSaisie ;
GereComboVue ; BZoomSousPlan.Visible:=FALSE ;
BE_BUDJAL.Value:=M.JAL ; BE_NATUREBUD.Value:=M.Nature ; BE_ETABLISSEMENT.Value:=M.Etabl ;
Case Action of
   taCreat : BEGIN
             BE_NumeroPiece.Caption:='' ; BE_NATUREBUD.Value:='INI' ;
             GB.Enabled:=False ;
             If CatExiste And (Not VisuConso) And (M.TypeSaisie='GS') Then BZoomSousPlan.Visible:=TRUE ;
             PositionneEtabUser(BE_ETABLISSEMENT, False); // 15087
             END ;
   taModif : BEGIN
             BE_BUDJAL.Enabled:=False ; BE_NATUREBUD.Enabled:=False ;

             {JP 24/08/05 : FQ 16501 et 16502 : Désactive la section / compte}
             GereCodeImplicite(False);

             BE_NumeroPiece.Caption:=InttoStr(M.Num) ; NumeroPiece:=M.Num ;
             GB.Enabled:=True ; GB.SetFocus ;
             END ;
 taConsult : BEGIN
             LectureSeule ;
             BE_NumeroPiece.Caption:=InttoStr(M.Num) ; NumeroPiece:=M.Num ;
             GB.Enabled:=True ; BAction.Enabled:=False ;
             if VisuConso then BEGIN BE_ETABLISSEMENT.Vide:=True ; BE_ETABLISSEMENT.Reload ; END ;
             PositionneEtabUser(BE_ETABLISSEMENT, False); // 15087
             END ;
   END ;
AffecteGrid(GB,Action) ;
END ;

procedure TFSaisBud.GereEnabled ;
Var ACol,ARow : integer ;
    T         : TList ;
BEGIN
ACol:=GB.Col ; ARow:=GB.Row ;
T:=GetL(ACol,ARow) ;
BVentil.Enabled:=((T<>Nil) and (ModeSaisBud in [msbGene,msbSect])) ;
BComplement.Enabled:=(T<>Nil) ;
BAction.Enabled:=(Action<>taConsult) ;
BPasteCol.Enabled:=((Action<>taConsult) and (CurCopy>0)) ;
BPasteAll.Enabled:=((Action<>taConsult) and (CurCopy>0)) ;
BCopyCol.Enabled:=(Action<>taConsult) ;
END ;

{================================== GRID ============================================}
procedure TFSaisBud.GBCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
begin
GereEnabled ;
if Action=taConsult then Exit ;
if GB.ColWidths[GB.Col]=0 then
   BEGIN
   Cancel:=True ;
   if GB.Col<GB.ColCount-1 then ACol:=GB.Col+1 else ACol:=1 ;
   END ;
StCellCur:=GB.Cells[GB.Col,GB.Row] ;
end;

procedure TFSaisBud.GereCell ( ACol,ARow : integer ) ;
Var T : TList ;
    i : integer ;
    OldX,NewX,Mt,Tot,XD,XC : Double ;
    O         : TOBM ;
BEGIN
T:=GetL(ACol,ARow) ; if T.Count<=0 then CreerNewOBM(ACol,ARow) ;
OldX:=SommeCase(ACol,ARow) ; NewX:=Valeur(GB.Cells[ACol,ARow])*CurCoef ;
if Arrondi(NewX-OldX,Deci)=0 then Exit ;
T:=GetL(ACol,ARow) ; Tot:=0 ;
for i:=0 to T.Count-1 do
    BEGIN
    O:=TOBM(T[i]) ;
    if OldX=0 then
       BEGIN
       if i=0 then
          BEGIN
          if OkD then BEGIN XD:=NewX ; XC:=O.GetCreditSaisi ; END
                 else BEGIN XC:=NewX ; XD:=O.GetDebitSaisi  ; END ;
          END else
          BEGIN
          if OkD then BEGIN XD:=0 ; XC:=O.GetCreditSaisi ; END
                 else BEGIN XC:=0 ; XD:=O.GetDebitSaisi  ; END ;
          END ;
       END else
       BEGIN
       if OkD then
          BEGIN
          if i=T.Count-1 then Mt:=NewX-Tot else
             BEGIN
             Mt:=O.GetDebitSaisi ; Mt:=Arrondi(Mt*NewX/OldX,CurDec) ; Tot:=Tot+Mt ;
             END ;
          XD:=Mt ; XC:=O.GetCreditSaisi ;
          END else
          BEGIN
          if i=T.Count-1 then Mt:=NewX-Tot else
             BEGIN
             Mt:=O.GetCreditSaisi ; Mt:=Arrondi(Mt*NewX/OldX,CurDec) ; Tot:=Tot+Mt ;
             END ;
          XC:=Mt ; XD:=O.GetDebitSaisi ;
          END ;
       END ;
    O.SetMontantsBUD(XD,XC) ;
    END ;
END ;

procedure TFSaisBud.GBCellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
Var St : String ;
    T  : TList ;
    Wh : string; {JP 23/08/05 : FQ 16372}
    sCX_JAL: string; {Thl FQ 17542 07/06/06}
begin
if csDestroying in ComponentState then Exit ;
if Action=taConsult then Exit ;
{Pas de changement -- > Exit}
St:=GB.Cells[ACol,ARow] ; if St=StCellCur then Exit ;
{Pas de montant ni d'objet --> Exit}
T:=GetL(ACol,ARow) ; if ((T=Nil) and (St='')) then Exit ;

  {JP 23/08/05 : FQ 16372 : Vérification des croisements autorisés}
  if (ModeSaisBud = msbSectGene) or (ModeSaisBud = msbGeneSect) then begin
    {Constitution de la requête sur les croisements. Ajout du Trim, car certaines zones de la grille ont un ' ' en entête}
         if ModeSaisBud = msbSectGene then Wh := 'CX_COMPTE = "' + Trim(BE_BUDGENE.Text) + '" AND CX_SECTION = "' + Trim(GB.Cells[0, ARow]) + '"'
    else if ModeSaisBud = msbGeneSect then Wh := 'CX_COMPTE = "' + Trim(GB.Cells[0, ARow]) + '" AND CX_SECTION = "' + Trim(BE_BUDSECT.Text) + '"';
    {b Thl FQ 17542 07/06/06}
    sCX_JAL := BE_BUDJAL.Value;
    if CatBudget<>'' then sCX_JAL := CatBudget;
    {On vérifie si le croisement existe et donc s'il est autorisé}
    //if not ExisteSQL('SELECT CX_INFO FROM CROISCPT WHERE CX_TYPE = "BUD" AND CX_JAL = "' + BE_BUDJAL.Value + '" AND ' + Wh) then begin
    if not ExisteSQL('SELECT CX_INFO FROM CROISCPT WHERE CX_TYPE = "BUD" AND CX_JAL = "' + sCX_JAL + '" AND ' + Wh) then begin
    {e Thl FQ 17542}
      HShowMessage('0;' + Caption + ';Votre saisie ne respecte pas les croisements autorisés.'#13 +
                                    'Veuillez modifier votre saisie;W;O;O;O', '', '');
      Cancel := True;
      Exit;
    end;
  end;

{Objet ou Saisie --> Gérer la cellule}
PieceModifiee:=True ;
FormateZone(ACol,ARow) ; CreerObjetCase(ACol,ARow) ;
GereCell(ACol,ARow) ;
{Sommations}
SommeLigne(ARow) ; SommeCol(ACol) ; SommeTout ; Synchro(TRUE) ;
end;

procedure TFSaisBud.GBTopLeftChanged(Sender: TObject);
begin Synchro(TRUE) ; end;

procedure TFSaisBud.SBChange(Sender: TObject);
begin Synchro(FALSE) ; end;

procedure TFSaisBud.SDChange(Sender: TObject);
begin Synchro(FALSE) ; end;

procedure TFSaisBud.FormResize(Sender: TObject);
BEGIN
if Width<WidthInit then Width:=WidthInit ;
if Height<HeightInit then Height:=HeightInit ;
SD.Top:=0  ; SD.Height:=PD.Height ;
SB.Left:=1 ; SB.Width:=GB.Width-2 ;
GBas.Width:=GB.Width ; TG.Left:=GBas.Width ;
end;

Procedure TFSaisBud.Synchro (DepuisGrid : Boolean)  ;
BEGIN
if ScrollEnCours then exit ;
ScrollEnCours:=TRUE ;
if DepuisGrid then BEGIN SD.Position:=GB.Row ; SB.Position:=GB.Col ; END
              else BEGIN GB.Row:=SD.Position ; GB.Col:=SB.Position ; END ;
GBas.LeftCol:=GB.LeftCol  ; GBas.Col:=GB.Col ;
GDroit.TopRow:=GB.TopRow  ; GDroit.Row:=GB.Row ;
ScrollEnCours:=FALSE ;
END ;

// Ajuste les ascenceurs
Procedure TFSaisBud.AjusteAscenceurs ;
BEGIN
SD.Max:=GB.RowCount-1 ; SB.Max:=GB.ColCount-1 ;
END ;

procedure TFSaisBud.GBEnter(Sender: TObject);
Var ACol,ARow : longint ;
    Cancel : boolean ;
begin
  ACol:=GB.Col ; ARow:=GB.Row ; Cancel:=False ;
  GBCellEnter(Nil,ACol,ARow,Cancel) ;
  if ((Action<>taConsult) and (BE_BUDJAL.Enabled) and (BE_BUDJAL.Value <> '')) then begin
    BE_BUDJAL.Enabled := False;
    {JP 24/08/05 : FQ 16501 et 16502 : Désactive la section / compte}
    GereCodeImplicite(False);
  end;
end;

procedure TFSaisBud.GBExit(Sender: TObject);
Var ACol,ARow : longint ;
    Cancel : boolean ;
begin
if csDestroying in ComponentState then Exit ;
ACol:=GB.Col ; ARow:=GB.Row ; Cancel:=False ;
GBCellExit(Nil,ACol,ARow,Cancel) ;
end;

procedure TFSaisBud.BE_ETABLISSEMENTChange(Sender: TObject);
begin
if BE_BUDJAL.Value='' then Exit ;
if ((BE_ETABLISSEMENT.Value='') and (Not VisuConso)) then Exit ;
if ((BE_BUDJAL.Enabled) and (Not PieceModifiee)) then BE_BUDJALChange(Nil) ;
end;

procedure TFSaisBud.BActionMouseEnter(Sender: TObject);
begin
PopZoom97(BAction,POPZ) ; 
end;

procedure TFSaisBud.BMenuZoomMouseEnter(Sender: TObject);
begin
PopZoom97(BMenuZoom,POPZ) ;
end;

procedure TFSaisBud.Combo97DockChanged(Sender: TObject);
begin
VUE.ItemIndex:=VUE.Tag ;
RESOL.ItemIndex:=RESOL.Tag ;
end;

procedure TFSaisBud.BZoomSousPlanClick(Sender: TObject);
Var CodeCat,StSPJ,Jal,Sect : String ;
begin
CodeCat:=CatBudget ; StSPJ:=SPCat ; Jal:=BE_BUDJAL.Value ; Sect:=BE_BUDSECT.Text ;
If ChoisirValCat(CodeCat,StSPJ,Jal,Sect,FALSE,taModif) Then
   BEGIN
   BE_BUDJAL.Value:=Jal ; BE_BUDSECT.Text:=Sect ;
   END ;
end;

procedure TFSaisBud.CSensChange(Sender: TObject);
Var Old : boolean ;
    ACol,ARow : integer ;
begin
Old:=OkD ;
OkD:=(CSens.ItemIndex=0) ;
if OkD<>Old then
   BEGIN
   for ACol:=1 to GB.ColCount-1 do
       BEGIN
       for ARow:=1 to GB.RowCount-1 do if GetL(ACol,ARow)<>Nil then ShowCase(ACol,ARow) ;
       SommeCol(ACol) ;
       END ;
   for ARow:=1 to GB.RowCount-1 do SommeLigne(ARow) ;
   SommeTout ; CurCopy:=0 ;
   END ;
end;


procedure TFSaisBud.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

{JP 24/08/05 : FQ 16501 et 16502 : On désactive la posibilité de modifier le compte ou
               la section en saisie croisée dès lors que la saisie est commencée => je
               Transforme le "champ" PieceModifiee en property afin de gérer les zones
               en même temps que le boolean;
{---------------------------------------------------------------------------------------}
procedure TFSaisBud.SetPieceModifiee(Value: Boolean);
{---------------------------------------------------------------------------------------}
begin
  FPieceModifiee := Value;
  {Active / désactive la section / compte}
  GereCodeImplicite(not Value);
end;

{Active / désactive la section / compte
{---------------------------------------------------------------------------------------}
procedure TFSaisBud.GereCodeImplicite(Actif: Boolean);
{---------------------------------------------------------------------------------------}
begin
  {JP 23/08/05 : FQ 16501 : On désactive la possiblité de modifier le compte}
  if (ModeSaisBud = msbSectGene) then
    BE_BudGene.Enabled := Actif
  {JP 23/08/05 : FQ 16502 : On désactive la possiblité de modifier la section}
  else if (ModeSaisBud = msbGeneSect) then
    BE_BudSect.Enabled := Actif;
end;

{b FP FQ16008-FQ16048-FQ16049-FQ16057}
{---------------------------------------------------------------------------------------}
function ModifieRequeteGrille(SQL: String): String;
{---------------------------------------------------------------------------------------}
var

  NewSQL:       String;
  SelectClause: String;
  GroupClause:  String;
  OrderClause:  String;
  FromClause:   String;
  St:           String;
  St1:          String;
  ChampDates:   TStringList;
  i:            Integer;
begin
  {Transforme la requête SQL pour le pas tenir compte des secondes pour les champs dates.
   L'utilisation da le clause distinct affiche plusieurs mouvements budgétaires au lieu d'un seul}
  ChampDates := TStringList.Create;
  try
  ChampDates.Add('BE_DATECREATION');
  ChampDates.Add('BE_DATEMODIF');

  SelectClause := Copy(SQL, 1, Pos('FROM', SQL)-1);
  SelectClause := StringReplace(SelectClause, 'DISTINCT ', '', [rfIgnoreCase]);

  if Pos('ORDER BY', SQL) > 0 then begin
    FromClause  := Copy(SQL, Pos('FROM', SQL), Pos('ORDER BY', SQL)-Pos('FROM', SQL));
    OrderClause := Copy(SQL, Pos('ORDER BY', SQL), Length(SQL));
  end
  else begin
    FromClause  := Copy(SQL, Pos('FROM', SQL), Length(SQL));
    OrderClause := '';
  end;

  {Retire les champs date provenant de la clause Select dans le group by}
  GroupClause := StringReplace(SelectClause, 'SELECT ', '', [rfIgnoreCase]);
  GroupClause := StringReplace(GroupClause, 'DISTINCT ', '', [rfIgnoreCase]);
  St          := StringReplace(GroupClause, ', ', ';', [rfReplaceAll]);
  GroupClause := '';          
  St1 := Trim(ReadTokenSt(St));
  while Trim(St1) <> '' do begin
    if ChampDates.IndexOf(St1) = -1 then
      GroupClause := GroupClause+', '+St1;
    St1 := Trim(ReadTokenSt(St));
  end;
  GroupClause := 'GROUP BY '+Copy(GroupClause, 2, Length(GroupClause)); {Retire la première virgule}

  {Ajoute la fonction MAX pour les champs date}
  for i:=0 to ChampDates.Count-1 do
    SelectClause := StringReplace(SelectClause, ChampDates[i], 'MAX('+ChampDates[i]+') '+ChampDates[i], [rfIgnoreCase]);

  NewSQL := SelectClause+' '+FromClause+' '+GroupClause+' '+OrderClause;
  Result := NewSQL;
  finally
    ChampDates.Free;
  end;
end;
{e FP FQ16008-FQ16048-FQ16049-FQ16057}



end.
