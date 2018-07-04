unit choisir;

interface

uses
  Windows, Classes, Controls, Forms, HTB97, StdCtrls,
  Hctrls, HPanel, Hent1, Grids, Buttons, HSysMenu,
  sysutils, LookUp //LM20070315
{$IFDEF UNUSED}
  , ExtCtrls
{$ENDIF}
  ;

type TArrayofBoolean = array of boolean;

function ChoisirDansLaListe(Titre: string; Liste: TStrings; DefautIndex: Integer; MultiSelect: Boolean; var Selection: TArrayofBoolean; Info: string = ''; VoirDel: Boolean = false): integer;
function ChoisirParTablette(TT, Plus: string; var code: string): boolean; //LM20070315
function Tablette2TStringList(TT, Plus: string; var Liste: HTStringList; CodeLibOnly: boolean = true): boolean;

type
  TfChoisir = class(TForm)
    hmTrad: THSystemMenu;
    pnlCorps: THPanel;
    Grid: THGrid;
    pnlBottom: THPanel;
    BOuvrir: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BAide: TToolbarButton97;
    LblInfo: TLabel;
    BDelete: TBitBtn;
    procedure BOuvrirClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BDeleteClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure GridDblClick(Sender: TObject);
    procedure GridKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Déclarations privées }
    AValider, ASupprimer: boolean;
    procedure Affiche(Liste: TStrings; DefautIndex: Integer);
    procedure InitMultiSelection(Selection: TArrayofBoolean);
    procedure LitMultiSection(var Selection: TArrayofBoolean);
    function GetRow: integer;
  public
    property Row: integer read getRow;
  end;


implementation
uses uSatUtil;

{$R *.DFM}

function Tablette2TStringList(TT, Plus: string; var Liste: HTStringList; CodeLibOnly: boolean = true): boolean;
var i, x: integer;
//rq : que code + libelle
begin
  result := false;
  RemplirListe(TT, Plus);

  for i := 1 to High(V_PGI.DECombos) do
  begin
    if V_PGI.DECombos[i].TT = TT then
    begin
      //V_PGI.DECombos[i].Valeurs ; //Token = Tab ou #9
      if CodeLibOnly then
        for x := 0 to V_PGI.DECombos[i].Valeurs.count - 1 do
          Liste.add(gtfs(V_PGI.DECombos[i].Valeurs[x], #9, 1) + #9 +
            gtfs(V_PGI.DECombos[i].Valeurs[x], #9, 2))
      else
        Liste := V_PGI.DECombos[i].Valeurs;

      result := true;
      break;
    end;
  end;
end;


///////////////////////////////////////////////////////////////////////////////

function ChoisirDansLaListe(Titre: string; Liste: TStrings; DefautIndex: Integer; MultiSelect: Boolean; var Selection: TArrayofBoolean; Info: string = ''; VoirDel: Boolean = false): integer;
var f: TfChoisir;
begin
  f := TfChoisir.create(Application);
  f.Caption := Titre;
  f.LblInfo.caption := Info;
  f.Affiche(Liste, DefautIndex);
  f.bDelete.visible := VoirDel (*pour le moment...  :*) and (not MultiSelect);
  if MultiSelect then f.InitMultiSelection(Selection);
  f.ShowModal;
  if (f.AValider) or (f.ASupprimer) then
  begin
    if not MultiSelect then
    begin
      result := f.Row;
      if f.ASupprimer then result := -result;
    end
    else
    begin
      result := 0; //indique le l'utilisateur à valider
      f.LitMultiSection(Selection);
    end;

  end
  else
    result := -1;
  f.free;
end;

function NbCarInString(St, Car: string): integer;
var i, n: Integer;
begin
  Result := 0; n := 0; i := 0;
  repeat
    i := i + n;
    n := pos(car, Copy(st, 1 + i, length(St)));
    if n > 0 then inc(Result) else break;
  until n = 0;
end;

//////////////////////////////////////////////////////////////////////////////////

procedure AjouteligneListeGrid(Ligne: string; Grid: THGrid);
var j, nb, mxw: Integer;
  St: string;
begin
  Ligne := Ligne + #9;
  nb := nbCarInString(Ligne, #9);
  if grid.ColCount < Nb then
    for j := Grid.ColCount + 1 to nb do
    begin
      Grid.ColCount := Grid.ColCount + 1;
      Grid.ColWidths[Grid.ColCount - 1] := 1;
    end;
  for j := 1 to Nb do
  begin
    St := gtfs(ligne, #9, j);
    mxw := Grid.Canvas.TextWidth(St) + 5;
    if mxw > Grid.ColWidths[j - 1] then Grid.ColWidths[j - 1] := mxw;
    Grid.Cells[j - 1, Grid.RowCount - 1] := St;
  end;
  Grid.RowCount := Grid.RowCount + 1;
end;

///////////////////////////////////////////////////////////////////////////////

procedure tfChoisir.Affiche(Liste: TStrings; DefautIndex: Integer);
//XMG 08/11/02 début
var i: integer;
begin
  if Liste.Count = 0 then exit;
  for i := 0 to Liste.Count - 1 do
    AjouteligneListeGrid(Liste[i], Grid);
//XMG 11/12/02 début
  grid.Row := Grid.FixedRows + DefautIndex * ord((DefautIndex > -1) and (DefautIndex <= Grid.RowCount + Grid.FixedRows - 1));
  if Grid.RowCount > Grid.FixedRows then Grid.RowCOunt := Grid.RowCount - 1;
//XMG 11/12/02 fin
end;

///////////////////////////////////////////////////////////////////////////////

procedure TfChoisir.InitMultiSelection(Selection: TArrayofBoolean);
var i: integer;
begin
  Grid.MultiSelect := TRUE;
  Grid.ClearSelected;
  for i := Low(Selection) to High(Selection) do
    grid.FlipSelection(i + Grid.FixedRows);
end;

///////////////////////////////////////////////////////////////////////////////

procedure TfChoisir.LitMultiSection(var Selection: TArrayofBoolean);
var i: integer;
begin
  for i := Low(Selection) to High(Selection) do
    Selection[i] := false;
//XMG 08/10/02 début
  if Length(Selection) < Grid.RowCount - Grid.FixedRows - 1 then Setlength(Selection, Grid.RowCount - Grid.FixedRows - 1) else
    if Length(Selection) > Grid.RowCount - Grid.FixedRows - 1 then Selection := copy(Selection, 0, Grid.RowCount - Grid.FixedRows - 1);
//XMG 08/10/02 fin
  for i := Grid.FixedRows to Grid.RowCount - 1 do
    Selection[i - Grid.FixedRows] := grid.IsSelected(i);
end;

///////////////////////////////////////////////////////////////////////////////

procedure TfChoisir.BOuvrirClick(Sender: TObject);
begin
  AValider := True;
  Close;
end;

///////////////////////////////////////////////////////////////////////////////

procedure TfChoisir.BAnnulerClick(Sender: TObject);
begin
  AValider := false; ASupprimer := FALSE;
  Close;
end;

///////////////////////////////////////////////////////////////////////////////

procedure TfChoisir.FormShow(Sender: TObject);
var i, j, t: integer;
begin
  j := 0;
  for i := 0 to Grid.ColCount - 1 do j := j + Grid.ColWidths[i];
  if j + 4 > Width then Width := j + 4;
  j := Grid.width div Grid.ColCount;
  t := Grid.Width;
  for i := 0 to Grid.ColCOunt - 1 do
  begin
    if i = Grid.ColCOunt - 1 then j := t;
    if j > Grid.ColWidths[i] then Grid.ColWidths[i] := j;
    t := t - Grid.Colwidths[i];
  end;
//if Grid.RowCount>Grid.FixedRows then Grid.RowCOunt:=Grid.RowCount-1 ; //XMG 11/12/02
  AValider := false;
end;

///////////////////////////////////////////////////////////////////////////////

procedure TfChoisir.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (key = VK_VALIDE) or (key = VK_RETURN) then BOuvrirClick(nil);
end;
//////////////////////////////////////////////////////////////////////////////////

procedure TfChoisir.BDeleteClick(Sender: TObject);
begin
  ASupprimer := true;
  close;
end;
//////////////////////////////////////////////////////////////////////////////////

procedure TfChoisir.FormCreate(Sender: TObject);
begin
  Grid.RowCount := Grid.FixedRows + 1;
  Grid.ColCount := Grid.FixedCols;
  Grid.RowHeights[Grid.FixedRows - 1] := 5;
  Grid.ColWidths[0] := 1;
end;
//////////////////////////////////////////////////////////////////////////////////

function TfChoisir.GetRow: integer;
begin
  Result := -1;
  if AValider or Asupprimer then Result := Grid.Row - Grid.FixedRows;
end;
//////////////////////////////////////////////////////////////////////////////////

procedure TfChoisir.GridDblClick(Sender: TObject);
begin
  if not Grid.MultiSelect then BOuvrirCLick(nil);
end;
//////////////////////////////////////////////////////////////////////////////////

procedure TfChoisir.GridKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if not Grid.MultiSelect then if (key = VK_RETURN) then BOuvrirCLick(nil);
end;



function ChoisirParTablette(TT, Plus: string; var code: string): boolean;
//LM20070315
var p: integer;
  StTable, StCode, StWhere, StPrefixe, StLib, sql, stOrder, Titre: string;
  //Local: boolean;
  pt: TPoint;
  ctrl: TEdit;
begin
  TT := Trim(UpperCase(TT));
  GetCorrespType(TT, StTable, StCode, StWhere, StPrefixe, StLib);

  ChangeWhereCombo(StWhere);
  SQL := 'Select ' + StCode + ', ' + StLib + ' FROM ' + StTable;
  ChangeWhereTT(StWhere, Plus, false);
  if Pos(' ORDER BY ', StWhere) <= 0 then
  begin
    P := Pos('DISTINCT', StCode);
    if P > 0 then
      StOrder := Copy(StCode, p + 9, 50)
    else
      StOrder := StCode;
  end
  else
    StOrder := StCode;

  (*local := (ctrl = nil);
  if local then*)
  begin
    ctrl := TEdit.create(screen.ActiveForm);
    ctrl.visible := false;
    ctrl.Parent := screen.ActiveForm;
    GetCursorPos(pt);
    pt := screen.ActiveForm.ScreenToClient(pt);
    ctrl.left := pt.X;
    ctrl.top := pt.Y;
  end;

  p := TTToNum(TT);
  if p >= 0 then Titre := V_PGI.DECombos[p].libelle;

  result := LookupList(ctrl, Titre, stTable, stCode, stLib, stWhere, stOrder,
    (Pos('|', stLib) = 0), 0, '', tlDefault, 1); //20070420

  if result then code := TEdit(ctrl).text;
  //if local then
    ctrl.free
end;

end.
