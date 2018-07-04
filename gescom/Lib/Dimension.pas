unit Dimension;

interface

uses Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
     HPanel, UiUtil, HEnt1, StdCtrls, Hctrls, ExtCtrls, HTB97, Grids, hmsgbox, UTOB,
{$IFDEF EAGLCLIENT}
     eTablette,
{$ELSE}
     Tablette, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} PrintDBG,
{$ENDIF}
     Menus, HPop97,M3FP, TntStdCtrls, TntGrids;

Procedure ParamDimension ;

type
  TFDimension = class(TForm)
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    FAutoSave: TCheckBox;
    FListe: THGrid;
    Msg: THMsgBox;
    FGrilles: THValComboBox;
    BInsLigne: TToolbarButton97;
    BDelLigne: TToolbarButton97;
    BInsLigne2: TToolbarButton97;
    BDelLigne2: TToolbarButton97;
    bColonne: TToolbarButton97;
    bLigne: TToolbarButton97;
    TAjoutDim: TToolbarButton97;
    Recherche: TToolbarButton97;
    FindDialog: TFindDialog;
    FGrillesLibres: THValComboBox;
    Panel1: TPanel;
    HLabel1: THLabel;
    FTypeDim: THValComboBox;
    BDecaleApresClick: TToolbarButton97;
    BDecaleAvantClick: TToolbarButton97;
    BDecaleAvantClick2: TToolbarButton97;
    BDecaleApresClick2: TToolbarButton97;
    PopLigCol: TPopupMenu;
    InsereTaille: TMenuItem;
    Supprimetaille: TMenuItem;
    N1: TMenuItem;
    TriAsc: TMenuItem;
    TriDesc: TMenuItem;
    bTriAsc2: TToolbarButton97;
    bTriAsc: TToolbarButton97;
    bTriDesc2: TToolbarButton97;
    bTriDesc: TToolbarButton97;
    cbControleDoublon: TCheckBox;
    procedure BChercheClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeSetEditText(Sender: TObject; ACol, ARow: Integer; const Value: String);
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure bColonneClick(Sender: TObject);
    procedure FTypeDimChange(Sender: TObject);
    procedure InsereTailleClick(Sender: TObject);
    procedure SupprimetailleClick(Sender: TObject);
    procedure TAjoutDimClick(Sender: TObject);
    procedure DecaleApresTailleClick(Sender: TObject);
    procedure DecaleAvantTailleClick(Sender: TObject);
    procedure TriClick(Sender: TObject; AscDesc: String);
    procedure BInsLigneClick(Sender: TObject);
    procedure BDelLigneClick(Sender: TObject);
    procedure FListeCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure RechercheClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure FListeDimMoved(Sender: TObject; FromIndex, ToIndex: Integer);
    procedure FindDialogFind(Sender: TObject);
    procedure BDecaleApres(Sender: TObject);
    procedure BDecaleAvant(Sender: TObject);
    procedure TriAscClick(Sender: TObject);
    procedure TriDescClick(Sender: TObject);
    procedure FListeCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure HelpBtnClick(Sender: TObject);
  private { Déclarations privées }
    FirstFind,FClosing : boolean;
    Modifie : Boolean ;
    TailleEnLigne : Boolean ;
    PresentationLigne : Boolean ;
    PresentationModifie : Boolean ;
    DimensionMoved : Boolean ;
    CurTypeDim : String ;
    CurTabletteDim : String ;
    TOB_DIM : TOB ;
    Procedure GetCellCanvas(ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) ;
    procedure SauveTrans ;
    procedure SauveGrille ;
    procedure Sauve ;
    Procedure Charge ;
    Function GetMaxCode(Ou : Integer) : String ;
    function IsLastTailleVide (Dim : integer) : Boolean ;
    function OnSauve : Integer ;
    procedure Incremente(Var Code : String);
    function BooleanToString (Ligne : Boolean) : String;
    function TestArticleTailleExiste (Col , Row : integer) : Boolean;
    function TestLibelleTailleRenseigne (Col , Row : integer) : Boolean;
    procedure ControlePhoneme(Col,Row : integer) ;
    function TestSortieCell (Col , Row : integer) : Boolean;
    procedure CtrlTailleVide ;
    procedure ModifBoutonDeplacement (PresentationLigne : boolean);
  public  { Déclarations publiques }
  end;

const TexteMessage: array[1..5] of string = (
    {1}   'Un ou plusieurs articles utilisent cette dimension. Vous devez renseigner un libellé'
    {2}  ,'Impossible de supprimer cette dimension qui est utilisée par un article'
    {3}  ,'Il existe déjà un libellé "' // + libellé phonème
    {4}  ,'" positionné en rang ' // + position phonème
    {5}  ,'Attention'
);

implementation

{$R *.DFM}

Procedure ParamDimension ;
var  FDimension: TFDimension;
     PP : THPanel ;
begin
  FDimension:=TFDimension.Create(Application) ;
  PP:=FindInsidePanel ;
  if PP=Nil then
  begin
    Try
      FDimension.ShowModal ;
    Finally
      FDimension.Free ;
    End ;
    Screen.Cursor:=SyncrDefault ;
  end else
  begin
   InitInside(FDimension,PP) ;
   FDimension.Show ;
  end ;
END ;

procedure TFDimension.BChercheClick(Sender: TObject);
begin
  Charge ;
end;

procedure TFDimension.ModifBoutonDeplacement (PresentationLigne : boolean);
begin
// Boutons de déplacement
BDecaleApresClick.Visible := PresentationLigne;
BDecaleAvantClick.Visible := PresentationLigne;
BDecaleApresClick2.Visible := Not PresentationLigne;
BDecaleAvantClick2.Visible := Not PresentationLigne;
// Boutons d'insersion/suppression
BInsLigne.Visible := PresentationLigne;
BDelLigne.Visible := PresentationLigne;
BInsLigne2.Visible := Not PresentationLigne;
BDelLigne2.Visible := Not PresentationLigne;
// Boutons de tri
BTriAsc.Visible := PresentationLigne;
BTriDesc.Visible := PresentationLigne;
BTriAsc2.Visible := Not PresentationLigne;
BTriDesc2.Visible := Not PresentationLigne;
end;

procedure TFDimension.Charge;
Var i,j,Posgrille,PosTaille,MaxDim : Integer ;
    T : TOB ;
    Q : TQuery;
    st : String;
begin
  if (CurTypeDim=FTypeDim.Value) and (Not PresentationModifie) then exit ;
  Case OnSauve of
    mrYes : Sauve ;
    mrCancel : BEGIN FTypeDim.Value:=CurTypeDim ; Exit ; END ;
  end;
  CurTypeDim:=FTypeDim.Value ;
  CurTabletteDim:='GG'+CurtypeDim[3];
  if TOB_DIM<>NIL then BEGIN TOB_DIM.Free ; TOB_DIM:=Nil ; END ;
  For i:=0 to FListe.RowCount-1 do For j:=0 to FListe.ColCount-1 do
  BEGIN
   FListe.Cells[j,i]:='' ;
   FListe.Objects[j,i]:=Nil ;
  END ;
  FListe.ColCount:=1 ;
  FListe.RowCount:=1 ;
  Fliste.Col:=0;
  Fliste.Row:=0;
  FListe.VidePile(FALSE) ;
  FGrilles.Datatype:='GCGRILLEDIM'+IntToStr(FTypeDim.ItemIndex+1) ;
  FGrilles.Reload ;
  if (FGrilles.Items.Count<=0) or (Trim(FGrilles.Items[0])='') then
  BEGIN
    Msg.execute(1,caption,'') ;
    FListe.Enabled:=FALSE ;
    Exit ;
  END else
  BEGIN
    FGrillesLibres.Datatype:='GCGRILLELIBREDIM'+IntToStr(FTypeDim.ItemIndex+1) ;
    AvertirTable(FGrillesLibres.Datatype) ;
    FGrillesLibres.Reload ;
  END ;

// Recherche de l'affichage en ligne ou colonne
  PresentationLigne := TRUE;
  Q:=OpenSQL('SELECT CC_LIBRE FROM CHOIXCOD WHERE CC_TYPE = "DIM" and CC_CODE = "'+CurTypeDim+'"',True,-1,'',true);
  if Not Q.EOF then
  BEGIN
    St:=Q.FindField('CC_LIBRE').AsString ;
    if (st<>'') and (st[1]='C') then PresentationLigne := FALSE;
  END ;
  Ferme(Q) ;
  ModifBoutonDeplacement (PresentationLigne);
  bColonne.Down:=Not PresentationLigne;
  bLigne.Down:=PresentationLigne ;

  FListe.Enabled:=TRUE ;
  TOB_DIM:=TOB.Create('Les dimensions',Nil,-1) ;
  TOB_DIM.LoadDetailDB('DIMENSION','"'+CurTypeDim+'"','GDI_TYPEDIM,GDI_GRILLEDIM,GDI_CODEDIM,GDI_DIMORLI',Nil,FALSE,TRUE) ;
//Taille En Ligne    attention au RowSizing et ColSizing à inverser   idem rowmoving

  if PresentationLigne then
  BEGIN
    FListe.ColCount:=4 ;
    FListe.RowCount:=FGrilles.Items.Count+1 ;
    Fliste.FixedCols:=2;
    Fliste.FixedRows:=1;
    Fliste.Options:=Fliste.Options + [goRowMoving] - [goColMoving];
    For i:=1 to FListe.RowCount-1 do
    BEGIN
      FListe.Cells[0,i]:=FGrillesLibres.Values[i-1] ;
      For j:=1 to FListe.RowCount-1 do if FGrillesLibres.Values[i-1]=FGrilles.Values[j-1] then FListe.Cells[1,i]:=FGrilles.Items[j-1] ;
      FListe.Objects[0,i]:=nil ;       // Stockage si modification de la dimension
    END ;
  END else
  BEGIN
    FListe.ColCount:=FGrilles.Items.Count+1 ;
    FListe.RowCount:=4 ;
    Fliste.FixedCols:=1;
    Fliste.FixedRows:=2;
    Fliste.Options:=Fliste.Options - [goRowMoving] + [goColMoving];
    For i:=1 to FListe.ColCount-1 do
    BEGIN
      FListe.Cells[i,0]:=FGrillesLibres.Values[i-1] ;
      For j:=1 to FListe.ColCount-1 do if FGrillesLibres.Values[i-1]=FGrilles.Values[j-1] then FListe.Cells[i,1]:=FGrilles.Items[j-1] ;
      FListe.Objects[i,0]:=nil ;
    END ;
  END ;
  PosGrille:=1 ;
  For i:=0 to TOB_DIM.Detail.Count-1 do
  BEGIN
    T:=TOB_DIM.Detail[i] ;
    FGrilles.Value:=T.GetValue('GDI_GRILLEDIM') ;
    For j:=0 to FGrilles.Items.Count-1 do if FGrilles.Value=FGrillesLibres.Values[j] then PosGrille:=j+1 ;
    PosTaille:=T.GetValue('GDI_RANG')+1 ;
    if PresentationLigne then
    BEGIN
      if PosTaille>=FListe.ColCount then FListe.ColCount:=PosTaille+1 ;
      FListe.Cells[PosTaille,PosGrille]:=T.GetValue('GDI_LIBELLE') ;
      FListe.Objects[PosTaille,PosGrille]:=T ;
    END
    else BEGIN
      if PosTaille>=FListe.RowCount then FListe.RowCount:=PosTaille+1 ;
         FListe.Cells[PosGrille,PosTaille]:=T.GetValue('GDI_LIBELLE') ;
         FListe.Objects[PosGrille,PosTaille]:=T ;
    END ;
  END ;
  if PresentationLigne then
  BEGIN
   FListe.Cells[1,0]:=TraduireMemoire ('Grille/Taille');
   For i:=2 to FListe.ColCount-1 do FListe.Cells[i,0]:=IntToStr(i-1) ;
   Fliste.DefaultColWidth:=50;
   Fliste.ColWidths[0]:=40;
   Fliste.Col:=2;
   Fliste.Row:=1;
   For i:=0 to FListe.RowCount-1 do     // Prise en compte du libelle colonne
   BEGIN
     MaxDim:=FListe.Canvas.textWidth(FListe.Cells[1,i]);
     if MaxDim>Fliste.ColWidths[1] then Fliste.ColWidths[1]:=MaxDim+6;
   END ;
  END
  else BEGIN
    FListe.Cells[0,1]:=TraduireMemoire ('Grille/Taille');
    For i:=2 to FListe.RowCount-1 do FListe.Cells[0,i]:=IntToStr(i-1) ;
    Fliste.Col:=1;
    Fliste.Row:=2;
    Fliste.DefaultColWidth:=80;
    Fliste.ColWidths[0]:=60;
  END ;
  FListe.Cells[0,0]:=TraduireMemoire ('Code');
  For i:=0 to FListe.ColCount-1 do FListe.ColAligns[i]:=taCenter ;
  CtrlTailleVide ;
  Fliste.SetFocus;
  Modifie:=FALSE ;
  DimensionMoved:=FALSE;
end;

Function TFDimension.GetMaxCode(Ou : Integer) : String ;
Var T : TOB ;
    Row,Col : Integer ;
    MaxCode : String ;
BEGIN
MaxCode:='' ;
if PresentationLigne then
   BEGIN
      For Col:=2 to FListe.ColCount-1 do
      BEGIN
      T:=TOB(FListe.Objects[Col,Ou]) ;
      if (T<>NIL) and (T.GetValue('GDI_CODEDIM')>MaxCode) then MaxCode:=T.GetValue('GDI_CODEDIM') ;
      END
   END else
   BEGIN
      For Row:=2 to FListe.RowCount-1 do
      BEGIN
      T:=TOB(FListe.Objects[Ou,Row]) ;
      if (T<>NIL) and (T.GetValue('GDI_CODEDIM')>MaxCode) then MaxCode:=T.GetValue('GDI_CODEDIM') ;
      END
   END ;
Result:=MaxCode ;
END ;

procedure TFDimension.SauveTrans;
Var T : TOB ;
    Code,MaxCode,DimOrli : String ;
    Col,Row   : Integer ;
    Der_rang  : Integer ;
    Cour_rang : Integer ;
begin
if Not modifie then Exit ;
Modifie:=FALSE ;
if PresentationLigne then
   BEGIN
   For Row:=1 to FListe.RowCount-1 do if FListe.Objects[0,Row]<>Nil then
      BEGIN
      Der_rang:=1 ;
      ExecuteSQL('DELETE FROM DIMENSION WHERE GDI_TYPEDIM="'+CurTypeDim+'"'+' AND GDI_GRILLEDIM="'+FListe.Cells[0,Row]+'"') ;
      MaxCode:=GetMaxCode(Row) ;
      For Col:=2 to FListe.ColCount-1 do
      if Trim(FListe.Cells[Col,Row])<>'' then
         BEGIN
         T:=TOB(FListe.Objects[Col,Row]) ;
         if T<>NIL then
            BEGIN
            Code:=T.GetValue('GDI_CODEDIM');
            DimOrli:=T.GetValue('GDI_DIMORLI');
            END else
            BEGIN
            Incremente(MaxCode) ;
            Code:=MaxCode ;
            DimOrli:='';
            END ;
         Cour_rang:=Col-1;
         if (Cour_rang <> Der_rang) then
            begin
            Cour_rang:=Der_rang;
            PresentationModifie:=TRUE;
            end;
         Der_rang:=Der_rang+1;
         ExecuteSQL('INSERT INTO DIMENSION '+
                    '(GDI_TYPEDIM,GDI_GRILLEDIM,GDI_CODEDIM,GDI_LIBELLE,GDI_DIMORLI,GDI_RANG) '+' VALUES '+
                    '("'+CurTypeDim+'","'+FListe.Cells[0,Row]+'","'+Code+'","'+FListe.Cells[Col,Row]+'","'+DimOrli+'",'+IntToStr(Cour_rang)+')') ;
         END ;
      END ;
   END else
   BEGIN
   For Col:=1 to FListe.ColCount-1 do if FListe.Objects[Col,0]<>Nil then
      BEGIN
      Der_rang:=1 ;
      ExecuteSQL('DELETE FROM DIMENSION WHERE GDI_TYPEDIM="'+CurTypeDim+'"'+' AND GDI_GRILLEDIM="'+FListe.Cells[Col,0]+'"') ;
      MaxCode:=GetMaxCode(Col) ;
      For Row:=2 to FListe.RowCount-1 do
      if Trim(FListe.Cells[Col,Row])<>'' then
         BEGIN
         T:=TOB(FListe.Objects[Col,Row]) ;
         if T<>NIL then
            BEGIN
            Code:=T.GetValue('GDI_CODEDIM');
            DimOrli:=T.GetValue('GDI_DIMORLI');
            END
            ELSE BEGIN
            Incremente(MaxCode) ;
            Code:=MaxCode ;
            DimOrli:='';
            END ;
         Cour_rang:=Row-1;
         if (Cour_rang <> Der_rang) then
            begin
            Cour_rang:=Der_rang;
            PresentationModifie:=TRUE;
            end;
         Der_rang:=Der_rang+1;
         ExecuteSQL('INSERT INTO DIMENSION '+
                    '(GDI_TYPEDIM,GDI_GRILLEDIM,GDI_CODEDIM,GDI_LIBELLE,GDI_DIMORLI,GDI_RANG) '+' VALUES '+
                    '("'+CurTypeDim+'","'+FListe.Cells[Col,0]+'","'+Code+'","'+FListe.Cells[Col,Row]+'","'+DimOrli+'",'+IntToStr(Cour_rang)+')') ;
         END ;
      END ;
   END ;
//AvertirTable(FGrillesLibres.Datatype) ;
end;

procedure TFDimension.SauveGrille;
Var Pos : Integer ;
    Q : TQuery;
    st : String;
begin
if Not DimensionMoved then Exit ;
DimensionMoved:=FALSE ;
if PresentationLigne then
   BEGIN
   For Pos:=1 to FListe.RowCount-1 do
      BEGIN
      Q:=OpenSQL('SELECT CC_LIBRE FROM CHOIXCOD WHERE CC_TYPE = "'+CurTabletteDim+'" and CC_CODE = "'+FListe.cells[0,Pos]+'"',False) ;
      if Not Q.EOF then
         BEGIN
         Q.Edit ;
         st:=Q.FindField('CC_LIBRE').AsString;
         Q.FindField('CC_LIBRE').AsString:=FormatFloat ('0000', Pos) ;
         Q.Post ;
         END ;
      Ferme(Q) ;
      END ;
   END else
   BEGIN
   For Pos:=1 to FListe.ColCount-1 do
      BEGIN
      Q:=OpenSQL('SELECT CC_LIBRE FROM CHOIXCOD WHERE CC_TYPE = "'+CurTabletteDim+'" and CC_CODE = "'+FListe.cells[Pos,0]+'"',False) ;
      if Not Q.EOF then
         BEGIN
         Q.Edit ;
         st:=Q.FindField('CC_LIBRE').AsString;
         Q.FindField('CC_LIBRE').AsString:=FormatFloat ('0000', Pos) ;
         Q.Post ;
         END ;
      Ferme(Q) ;
      END ;
   END ;
AvertirTable(FGrillesLibres.Datatype) ;
end;

procedure TFDimension.Sauve ;
var i : integer;
BEGIN
Transactions(SauveTrans,1) ;
if V_PGI.IOError<>oeOk
   then Msg.execute(4,caption,'')
   else BEGIN
   if PresentationLigne
      then for i:=1 to FListe.RowCount-1 do FListe.Objects[0,i]:=TObject (0)
      else for i:=1 to FListe.ColCount-1 do FListe.Objects[i,0]:=TObject (0) ;
   FListe.Invalidate ;
   END ;
Transactions(SauveGrille,1) ;
if V_PGI.IOError<>oeOk then Msg.execute(5,caption,'') ;
END ;

procedure TFDimension.FormShow(Sender: TObject);
begin
FListe.GetCellCanvas:=GetCellCanvas ;
FTypeDim.ItemIndex:=0 ;
BChercheClick(Nil) ;
end;

procedure TFDimension.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
Case OnSauve of
   mrYes : Sauve ;
   mrNo : ;
   mrCancel : Canclose:=FALSE ;
   END ;
end;

procedure TFDimension.FormDestroy(Sender: TObject);
begin
TOB_DIM.Free ;
end;

procedure TFDimension.FormCreate(Sender: TObject);
begin
TOB_DIM:=NIL ; TailleEnLigne:=TRUE ;
FClosing:=False ; 
end;


procedure TFDimension.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if IsInside(Self) then action:=caFree ;
FClosing:=True ; 
end;

procedure TFDimension.FListeSetEditText(Sender: TObject; ACol,ARow: Integer; const Value: String);
Var T : TOB ;
begin
T:=TOB (FListe.Objects[ACol,ARow]);
if ((T=nil) and (Value<>'')) or ((T<>Nil) and (value<>T.GetValue('GDI_LIBELLE'))) then
   BEGIN
   modifie:=TRUE;
   if PresentationLigne then FListe.Objects[0,ARow]:=TObject (1)
                        else FListe.Objects[ACol,0]:=TObject (1) ;
   FListe.Invalidate ;
   END
end;

procedure TFDimension.Incremente(var Code: String);
Var i : integer ;
    CL : String ;
BEGIN
CL:=uppercase(Trim(Code)) ;
if Length(CL)<3 then Code:='001' else
   BEGIN
   i:=3 ; While CL[i]='Z' do BEGIN CL[i]:='0' ; Dec(i) ; END ;
   if Ord(CL[i])=57 then CL[i]:='A' else CL[i]:=Succ(CL[i]) ;
   Code:=CL ;
   END ;
END ;

procedure TFDimension.BValiderClick(Sender: TObject);
begin
if (Not TestSortieCell(Fliste.Col, Fliste.Row)) then exit ;
Sauve ;
Charge;
PresentationModifie:=FALSE;
end;

procedure TFDimension.BFermeClick(Sender: TObject);
begin
Close ;
if FClosing and IsInside(Self) then THPanel(parent).CloseInside ;
end;

function TFDimension.OnSauve: Integer;
var reponse : integer;
begin

// MODIF LM 25/05/2000   BUG
if (Not TestSortieCell (Fliste.Col, Fliste.Row)) then
begin
  Reponse:=Msg.Execute(6,caption,'');
  if (reponse=mrNo) then result:=mrCancel
  else result:=mrNo;
end
else if Modifie then Result:=Msg.Execute(0,caption,'') else Result:=mrNo ;
end;

function TFDimension.BooleanToString (Ligne : Boolean) : String;
begin
if Ligne then Result:='L' else Result:='C' ;
end;

procedure TFDimension.bColonneClick(Sender: TObject);
begin

// MODIF LM 25/05/2000
if (Not TestSortieCell(Fliste.Col,Fliste.Row)) then exit ;

Case OnSauve of
   mrYes : Sauve ;
   mrNo : Modifie:=FALSE ;
   mrCancel :
      begin
      bColonne.Down:=Not PresentationLigne;
      bLigne.Down:=PresentationLigne ;
      exit ;
      end;
   END ;

// UPDATE CHOIXCOD SET CC_LIBRE = "X" WHERE CC_TYPE = "DIM" and CC_CODE = "DI1"
ExecuteSQL( 'UPDATE CHOIXCOD SET CC_LIBRE = "'+BooleanToString(Not PresentationLigne)+
            '" WHERE CC_TYPE = "DIM" and CC_CODE = "'+CurTypeDim+ '"') ;
PresentationModifie:=TRUE;
Charge ;
PresentationModifie:=FALSE;
ModifBoutonDeplacement(PresentationLigne) ;
end;

procedure TFDimension.FTypeDimChange(Sender: TObject);
begin
Charge;
end;

function TFDimension.IsLastTailleVide (Dim : integer) : Boolean ;
var i : integer ;
BEGIN
result:=TRUE;
if PresentationLigne then
   BEGIN
   for i:=1 to FListe.RowCount-1 do
      if FListe.Cells[FListe.ColCount-Dim,i]<>'' then BEGIN result:=FALSE ; Exit ; END ;
   END else
   BEGIN
   for i:=1 to FListe.ColCount-1 do
      if FListe.Cells[i,FListe.RowCount-Dim]<>'' then BEGIN result:=FALSE ; Exit ; END ;
   END ;
END ;

procedure TFDimension.CtrlTailleVide ;
var AjouteTaille, SupprimeTaille : boolean;
    DeuxTaillesVides : Boolean;
BEGIN
AjouteTaille:=FALSE;
SupprimeTaille:=FALSE;
if IsLastTailleVide(1) then
   BEGIN
   DeuxTaillesVides:=IsLastTailleVide(2);
   if PresentationLigne then
      BEGIN
      if (FListe.Col=FListe.ColCount-1)
         then AjouteTaille:=TRUE
         else if DeuxTaillesVides and (FListe.Col<FListe.ColCount-2)then SupprimeTaille:=TRUE;
      END else
      BEGIN
      if (FListe.Row=FListe.RowCount-1)
         then AjouteTaille:=TRUE
         else if DeuxTaillesVides and (FListe.Row<FListe.RowCount-2) then SupprimeTaille:=TRUE;
      END ;
   END else AjouteTaille:=TRUE;

if SupprimeTaille=TRUE then
   BEGIN
   if PresentationLigne then
      BEGIN
      if FListe.ColCount>4 then FListe.ColCount:=FListe.ColCount-1
      END else
      BEGIN
      if FListe.RowCount>4 then FListe.RowCount:=FListe.RowCount-1 ;
      END ;
   END ;
If AjouteTaille=TRUE then
   BEGIN
   if PresentationLigne then
      BEGIN
      FListe.ColCount:=FListe.ColCount+1 ;
      FListe.ColAligns[FListe.ColCount-1]:=taCenter ;
      FListe.Cells[FListe.ColCount-1,0]:=IntToStr(FListe.ColCount-2) ;
      END else
      BEGIN
      FListe.RowCount:=FListe.RowCount+1 ;
      FListe.Cells[0,FListe.RowCount-1]:=IntToStr(FListe.RowCount-2) ;
      END ;
   END ;
END ;

procedure TFDimension.FListeCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
CtrlTailleVide ;
end;

procedure TFDimension.InsereTailleClick(Sender: TObject);
var i : integer;
begin

// MODIF LM 25/05/2000
if (Not TestSortieCell (Fliste.Col, Fliste.Row)) then exit ;

if PresentationLigne then
begin
  for i:=FListe.ColCount-2 downto FListe.Col do
  begin
      FListe.Cells[i+1,FListe.Row]:=FListe.Cells[i,FListe.Row] ;
      FListe.Objects[i+1,FListe.Row]:=FListe.Objects[i,FListe.Row] ;
  end
end else
begin
  for i:=FListe.RowCount-2 downto FListe.Row do
  begin
      FListe.Cells[FListe.Col,i+1]:=FListe.Cells[FListe.Col,i] ;
      FListe.Objects[FListe.Col,i+1]:=FListe.Objects[FListe.Col,i] ;
   end;
end;

FListe.Cells[FListe.Col,FListe.Row]:='';
FListe.Objects[FListe.Col,FListe.Row]:=Nil;
CtrlTailleVide ;
Modifie:=TRUE;
if PresentationLigne then FListe.Objects[0,FListe.Row]:=TObject (1)
                     else FListe.Objects[FListe.Col,0]:=TObject (1) ;
FListe.Invalidate ;
end;

// MODIF LM 25/05/2000
// Teste si le libellé de la taille courante est renseigné
//       si la taille est utilisée par un article
// Signale s'il existe un doublon phonétique du libellé modifié
// MODIF DC 12/12/2002
// Contrôle de doublon effectué ssi cbControleDoublon coché (option non mémorisée)
function TFDimension.TestSortieCell (Col , Row : integer) : Boolean;
var bOk : boolean ;
begin
bOk:=TestLibelleTailleRenseigne(Col,Row) ;
if bOk and cbControleDoublon.Checked then ControlePhoneme(Col,Row) ;
Result:=bOk ;
end;

// Test si la taille est utilisé par un article
function TFDimension.TestArticleTailleExiste (Col , Row : integer) : Boolean;
var T : TOB ;
    GrilleDim, CodeDim: string;
begin
result:=False ;
T:=TOB(Fliste.Objects[Col,Row]) ;
if T<>Nil then
    BEGIN
    GrilleDim:=T.GetValue('GDI_GRILLEDIM');
    CodeDim:=T.GetValue('GDI_CODEDIM');

    if ExisteSQL ('Select GA_ARTICLE From ARTICLE Where GA_GRILLEDIM1="'+GrilleDim+'" and GA_CODEDIM1="'+CodeDim+'"') then result:=True
    else if ExisteSQL ('Select GA_ARTICLE From ARTICLE Where GA_GRILLEDIM2="'+GrilleDim+'" and GA_CODEDIM2="'+CodeDim+'"') then result:=True
    else if ExisteSQL ('Select GA_ARTICLE From ARTICLE Where GA_GRILLEDIM3="'+GrilleDim+'" and GA_CODEDIM3="'+CodeDim+'"') then result:=True
    else if ExisteSQL ('Select GA_ARTICLE From ARTICLE Where GA_GRILLEDIM4="'+GrilleDim+'" and GA_CODEDIM4="'+CodeDim+'"') then result:=True
    else if ExisteSQL ('Select GA_ARTICLE From ARTICLE Where GA_GRILLEDIM5="'+GrilleDim+'" and GA_CODEDIM5="'+CodeDim+'"') then result:=True ;
    END ;
end;

// Test si le libellé de la taille est renseigné
function TFDimension.TestLibelleTailleRenseigne (Col , Row : integer): Boolean;
var T : TOB ;
    Libelle : string;

begin
T := TOB(Fliste.Objects [Col, Row]);
if T<>NIL then
  begin
  Libelle:=FListe.Cells[Col,Row];
  if (Trim(Libelle)='') and (TestArticleTailleExiste(Col,Row)=TRUE) then
    begin
    PGIBox(TexteMessage[1],TexteMessage[5]);
    result:=FALSE;
    exit;
    end;
  end;
result:=TRUE;
end;

procedure TFDimension.ControlePhoneme(Col,Row: integer) ;
var ind,posDouble : integer ;
    libelle : string ;
begin
libelle:=phoneme(FListe.Cells[Col,Row]) ;
if libelle='' then exit ;
posDouble:=-1 ;
if PresentationLigne then
    BEGIN
    ind:=FListe.ColCount ;
    while (ind>1) and (posDouble=-1) do
        if (ind<>Col) and (phoneme(FListe.Cells[ind,Row])=libelle)
            then posDouble:=ind else dec(ind) ;
    END else
    BEGIN
    ind:=FListe.RowCount ;
    while (ind>1) and (posDouble=-1) do
        if (ind<>Row) and (phoneme(FListe.Cells[Col,ind])=libelle)
            then posDouble:=ind else dec(ind) ;
    END ;
if posDouble<>-1 then  // Doublon trouvé
    BEGIN
    if PresentationLigne then libelle:=FListe.Cells[posDouble,Row]
                         else libelle:=FListe.Cells[Col,posDouble] ;
    PGIBox(TexteMessage[3]+libelle+TexteMessage[4]+IntToStr(posDouble-1),TexteMessage[5]) ;
    END ;
end;

procedure TFDimension.SupprimetailleClick(Sender: TObject);
var i : integer ;
begin
// Test si la taille peut être supprimée
if (TestArticleTailleExiste(Fliste.Col,Fliste.Row))
   then BEGIN PGIBox(TexteMessage[2],TexteMessage[5]) ; exit ; END ;

if PresentationLigne then
   BEGIN
   for i:=FListe.Col to FListe.ColCount-2 do
      BEGIN
      FListe.Cells[i,FListe.Row]:=FListe.Cells[i+1,FListe.Row] ;
      FListe.Objects[i,FListe.Row]:=FListe.Objects[i+1,FListe.Row] ;
      END ;
   FListe.Cells[FListe.ColCount-1,FListe.Row]:='';
   FListe.Objects[FListe.ColCount-1,FListe.Row]:=Nil;
   END else
   BEGIN
   for i:=FListe.Row to FListe.RowCount-2 do
      BEGIN
      FListe.Cells[FListe.Col, i]:=FListe.Cells[FListe.Col, i+1] ;
      FListe.Objects[FListe.Col, i]:=FListe.Objects[FListe.Col, i+1] ;
      END ;
   FListe.Cells[FListe.Col,FListe.RowCount-1]:='';
   FListe.Objects[FListe.Col,FListe.RowCount-1]:=Nil;
   END ;
CtrlTailleVide ;
Modifie:=TRUE;
if PresentationLigne then FListe.Objects[0,FListe.Row]:=TObject (1)
                     else FListe.Objects[FListe.Col,0]:=TObject (1) ;
FListe.Invalidate ;
end;

procedure TFDimension.TAjoutDimClick(Sender: TObject);
begin

// MODIF LM 25/05/2000
if (TestSortieCell (Fliste.Col, Fliste.Row)=FALSE) then exit;

Case OnSauve of
   mrYes : Sauve ;
   mrCancel : Exit ;
   END ;

ParamTable(FGrilles.Datatype,taCreat,0,Nil,3) ;
AvertirTable(FGrilleslibres.Datatype) ;
PresentationModifie:=TRUE;
DimensionMoved:=TRUE;
Charge ;
//FListe.InsertCol(3) ; //Attention decale tout y compris les objects
end;

procedure TFDimension.DecaleApresTailleClick(Sender: TObject);
var T : TOB;
    st1 : string;
begin

if PresentationLigne then
begin
  st1 := Fliste.Cells [Fliste.Col, Fliste.Row];
  T := TOB(Fliste.Objects [Fliste.Col, Fliste.row]);

  FListe.Cells[Fliste.Col,FListe.Row]:=FListe.Cells[Fliste.Col+1,FListe.Row] ;
  FListe.Objects[Fliste.Col,FListe.Row]:=FListe.Objects[Fliste.Col+1,FListe.Row] ;
  FListe.Cells[Fliste.Col+1,FListe.Row]:=st1;
  FListe.Objects[Fliste.Col+1,FListe.Row]:=T;
  Fliste.Col:=Fliste.Col+1;
end else
begin
  st1 := Fliste.Cells [Fliste.Col, Fliste.Row];
  T := TOB(Fliste.Objects [Fliste.Col, Fliste.row]);

  FListe.Cells[Fliste.Col,FListe.Row]:=FListe.Cells[Fliste.Col,FListe.Row+1] ;
  FListe.Objects[Fliste.Col,FListe.Row]:=FListe.Objects[Fliste.Col,FListe.Row+1] ;
  FListe.Cells[Fliste.Col,FListe.Row+1]:=st1;
  FListe.Objects[Fliste.Col,FListe.Row+1]:=T;
  FListe.Row:=Fliste.Row+1;
end;

CtrlTailleVide ;
Modifie:=TRUE;
if PresentationLigne then FListe.Objects[0,FListe.Row]:=TObject (1)
                     else FListe.Objects[FListe.Col,0]:=TObject (1) ;
FListe.Invalidate ;
end;

procedure TFDimension.DecaleAvantTailleClick(Sender: TObject);
var T : TOB;
    st1 : string;
begin

if PresentationLigne then
begin
  if (Fliste.Col > 2) then
  begin
    st1 := Fliste.Cells [Fliste.Col, Fliste.Row];
    T := TOB(Fliste.Objects [Fliste.Col, Fliste.row]);

    FListe.Cells[Fliste.Col,FListe.Row]:=FListe.Cells[Fliste.Col-1,FListe.Row] ;
    FListe.Objects[Fliste.Col,FListe.Row]:=FListe.Objects[Fliste.Col-1,FListe.Row] ;
    FListe.Cells[Fliste.Col-1,FListe.Row]:=st1;
    FListe.Objects[Fliste.Col-1,FListe.Row]:=T;
    Fliste.Col:=Fliste.Col-1;
  end;
end else
begin
  if (Fliste.Row > 2) then
  begin
    st1 := Fliste.Cells [Fliste.Col, Fliste.Row];
    T := TOB(Fliste.Objects [Fliste.Col, Fliste.row]);

    FListe.Cells[Fliste.Col,FListe.Row]:=FListe.Cells[Fliste.Col,FListe.Row-1] ;
    FListe.Objects[Fliste.Col,FListe.Row]:=FListe.Objects[Fliste.Col,FListe.Row-1] ;
    FListe.Cells[Fliste.Col,FListe.Row-1]:=st1;
    FListe.Objects[Fliste.Col,FListe.Row-1]:=T;
    FListe.Row:=Fliste.Row-1;
  end;
end;

CtrlTailleVide ;
Modifie:=TRUE;
if PresentationLigne then FListe.Objects[0,FListe.Row]:=TObject (1)
                     else FListe.Objects[FListe.Col,0]:=TObject (1) ;
FListe.Invalidate ;
end;

procedure TFDimension.TriClick(Sender: TObject; AscDesc: String);
var Row,Col,PosTaille : integer ;
    TobGrille,TobItem : TOB ;
    GrilleDim,stTri : string ;
begin
if PresentationLigne then
    BEGIN
    Row:=FListe.Row ;
    TobGrille:=TOB.Create('Les dimensions',nil,-1) ;
    GrilleDim:=FListe.Cells[0,Row] ;
    stTri:='GDI_TYPEDIM,GDI_GRILLEDIM,GDI_LIBELLE '+AscDesc ;
    TobGrille.LoadDetailDB('DIMENSION','"'+CurTypeDim+'";"'+GrilleDim+'"',stTri,nil,False,True) ;
    // Maj de la grille par les tobs rechargées
    PosTaille:=Fliste.FixedCols ;
    While TobGrille.Detail.Count>0 do
        BEGIN
        // Supprime les tobs de la grille rechargée
        if TOB(FListe.Objects[PosTaille,Row])<>nil then TOB(FListe.Objects[PosTaille,Row]).Free ;
        FListe.Cells[PosTaille,Row]:='' ;

        TobItem:=TobGrille.Detail[0] ;
        TobItem.ChangeParent(TOB_DIM,-1) ; // Rattachement de la tob fille à TOB_DIM
        FListe.Cells[PosTaille,Row]:=TobItem.GetValue('GDI_LIBELLE') ;
        FListe.Objects[PosTaille,Row]:=TobItem ;
        inc(PosTaille) ;
        END ;
    END else
    BEGIN
    Col:=Fliste.Col ;
    TobGrille:=TOB.Create('Les dimensions',nil,-1) ;
    GrilleDim:=FListe.Cells[Col,0] ;
    stTri:='GDI_TYPEDIM,GDI_GRILLEDIM,GDI_LIBELLE '+AscDesc ;
    TobGrille.LoadDetailDB('DIMENSION','"'+CurTypeDim+'";"'+GrilleDim+'"',stTri,nil,False,True) ;
    // Maj de la grille par les tobs rechargées
    PosTaille:=Fliste.FixedRows ;
    While TobGrille.Detail.Count>0 do
        BEGIN
        // Supprime les tobs de la grille rechargée
        if TOB(FListe.Objects[Col,PosTaille])<>nil then TOB(FListe.Objects[Col,PosTaille]).Free ;
        FListe.Cells[Col,PosTaille]:='' ;

        TobItem:=TobGrille.Detail[0] ;
        TobItem.ChangeParent(TOB_DIM,-1) ; // Rattachement de la tob fille à TOB_DIM
        FListe.Cells[Col,PosTaille]:=TobItem.GetValue('GDI_LIBELLE') ;
        FListe.Objects[Col,PosTaille]:=TobItem ;
        inc(PosTaille) ;
        END ;
    END ;

TobGrille.Free ;
Modifie:=True ;
if PresentationLigne then FListe.Objects[0,FListe.Row]:=TObject (1)
                     else FListe.Objects[FListe.Col,0]:=TObject (1) ;
FListe.Invalidate ;
end;

procedure TFDimension.BInsLigneClick(Sender: TObject);
begin
InsereTailleClick(Self);
end;

procedure TFDimension.BDelLigneClick(Sender: TObject);
begin
SupprimetailleClick(Self);
end;

procedure TFDimension.GetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas;
  AState: TGridDrawState);
Var OkModifiee : boolean ;
    MaxDim : integer;
begin
if PresentationLigne then OkModifiee:=(Integer(FListe.Objects[0,ARow])=1)
                     else OkModifiee:=(Integer(FListe.Objects[ACol,0])=1) ;
if (gdFixed In Astate) and (OkModifiee) then Canvas.Font.Style:=[fsBold]
                                        else Canvas.Font.Style:=[] ;
if PresentationLigne and OkModifiee then
   BEGIN
   MaxDim:=FListe.Canvas.textWidth(FListe.Cells[1,ARow]);
   if MaxDim>Fliste.ColWidths[1] then Fliste.ColWidths[1]:=MaxDim+6 ;
   END ;
end;

procedure TFDimension.RechercheClick(Sender: TObject);
begin
// MODIF LM 25/05/2000
if (TestSortieCell (Fliste.Col, Fliste.Row)=FALSE) then exit;
FirstFind:=true; FindDialog.Execute ;
end;

procedure TFDimension.BImprimerClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
// AFAIREEAGL
{$ELSE}
PrintDBGrid (FListe, FTypeDim, 'Liste des dimensions', '');
{$ENDIF}
end;

procedure TFDimension.FListeDimMoved(Sender: TObject; FromIndex,
  ToIndex: Integer);
begin
DimensionMoved:=TRUE;
Modifie:=TRUE;
end;

procedure TFDimension.FindDialogFind(Sender: TObject);
begin
Rechercher(FListe,FindDialog, FirstFind);
end;

procedure TFDimension.BDecaleApres(Sender: TObject);
begin
DecaleApresTailleClick(Self);
end;

procedure TFDimension.BDecaleAvant(Sender: TObject);
begin
DecaleAvantTailleClick(Self);
end;

procedure TFDimension.TriAscClick(Sender: TObject);
begin
TriClick(Self,'asc') ;
end;

procedure TFDimension.TriDescClick(Sender: TObject);
begin
TriClick(Self,'desc') ;
end;

procedure TFDimension.FListeCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
cancel:=Not TestSortieCell(ACol,ARow) ;
end;

Procedure AGLParamDimension ( Parms : array of variant ; nb : integer ) ;
BEGIN
ParamDimension ;
END ;

procedure TFDimension.HelpBtnClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;


Initialization
RegisterAglProc('ParamDimension',False,0,AGLParamDimension) ;

end.
