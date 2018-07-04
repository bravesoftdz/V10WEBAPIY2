unit ListeDeSaisie;

interface

uses
  Forms, Windows, Menus, Dialogs, HSysMenu, Grids, Hctrls, HTB97, Controls, ExtCtrls,
  HPanel, Classes, StdCtrls,
{$IFDEF EAGLCLIENT}
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  M3FP, UTOB, Spin, Mask, hmsgbox,Dicobtp, TntExtCtrls, TntGrids;

(*
Tablette : GCSAISPIECE
  CO_LIBELLE : Titre de colonne
  CO_ABREGE :  Nom du champ sans le préfixe GL_
  CO_LIBRE :   Paramètres
                  Dnn :    Champ ajouté par défaut dans une nouvelle liste ;
                           nn = n° d'ordre d'ajout
                  O :      Champ obligatoire
                  F :      Colonne figée (non déplacable, doit être en tête de liste)
                  E=xxxx : Affiche xxxx dans la ligne exemple
                  Ax :     Alignement, x = G (à gauche), C (centré) ou D (à droite)
                  Vn :     Nombre de décimales = n
*)


// Passer le code de la liste à ouvrir au lancement de la fiche, ou chaine vide
//   pour lancer avec la liste par défaut.
procedure EntreeListeSaisie(CodeListe : String);

const
  MaxColonnes = 10;

type
  // Différents types de champs qu'on peut rencontrer dans la base
  TTypeChamp = (tcInteger, tcDouble, tcRate, tcBoolean, tcCombo, tcDate, tcVarchar, tcBlob, tcOther);

  TColonne = record      // Contient tous les paramètres d'une colonne
    Code : String;             // Code dans la tablette
    NomChamp : String;
    Titre : String;
    Exemple : String;
    Largeur : Integer;
    TypeChamp : TTypeChamp;
    NbDecim : Integer;
    Align : TAlignment;
    SepMilliers : Boolean;
    Obligatoire : Boolean;
    Figee : Boolean;
  end;

  TColonneList = class     // Classe qui gère la liste de colonnes
    FGrid : THGrid;
    FHMTrad : THSystemMenu;
    FColumns : Array of TColonne;
    FAvailFields : TOB;

    procedure SlideCols(BgnIndex, EndIndex, Sens : integer);
    function GetColCount : integer;
    function GetCol(index : integer) : TColonne;
    function GetCurrentCol : TColonne;
    function TypeFromType(StType : String) : TTypeChamp;
    function GenerateDetail : String;
    procedure LoadDetail(Dtl : String);

  public
    constructor Create(Grid : THGrid; SM : THSystemMenu; Fields : TOB);
    destructor Destroy; override;

    procedure ApplyToGrid;
    procedure ReAjustWidths;
    procedure GetGridWidths;
    function WidthHasChanged : boolean;

    procedure ClearList;
    procedure AddColumn(ColTOB : TOB; Index : integer; ApplyNow : Boolean = true);
    procedure MoveColumn(FromIndex, ToIndex : integer);
    procedure DelColumn(ColIndex : integer);
    function ColumnExists(Code : String) : boolean;
    function SkipFiged(ColNb : integer) : integer;
    procedure SetColAlign(ColNb : integer; Align : TAlignment);
    procedure SetColNbDec(ColNb, NbDec : integer);

    property Count : integer read GetColCount;
    property Columns[index : integer] : TColonne read GetCol;
    property CurrentColumn : TColonne read GetCurrentCol;
    property Detail : String read GenerateDetail write LoadDetail;

  end;

  TFListeSaisie = class(TForm)
    ToolWindow971: TToolWindow97;
    Dock971: TDock97;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    HMTrad: THSystemMenu;
    G_Liste: THGrid;
    FindLigne: TFindDialog;
    TDetailPrix: TToolWindow97;
    FieldsList: TListBox;
    tbLoad: TToolbarButton97;
    tbNew: TToolbarButton97;
    Dock972: TDock97;
    ToolWindow972: TToolWindow97;
    tbDelCol: TToolbarButton97;
    tbLeft: TToolbarButton97;
    tbCenter: TToolbarButton97;
    tbRight: TToolbarButton97;
    SpinDec: TSpinEdit;
    PEntete: THPanel;
    MsgCaptions: THMsgBox;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure G_ListeCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_ListeColumnMoved(Sender: TObject; FromIndex, ToIndex: Integer);
    procedure G_ListeMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure G_ListeDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure G_ListeDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure FieldsListMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure FieldsListEndDrag(Sender, Target: TObject; X, Y: Integer);
    procedure tbDelColClick(Sender: TObject);
    procedure tbLeftClick(Sender: TObject);
    procedure tbCenterClick(Sender: TObject);
    procedure tbRightClick(Sender: TObject);
    procedure SpinDecChange(Sender: TObject);
    procedure tbNewClick(Sender: TObject);
    procedure tbLoadClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);

  private
    FOldColX : integer;
    FWorkList : TColonneList;
    TOBAvailableFields : TOB;

    function AFInList(Code : String) : integer;
    procedure DrawPositionLine(ColX : integer);
    procedure RefreshCurrColAligns;
    procedure SetTitle;

  public
    CodeListe, LibListe : String;

    procedure LoadAvailableFields;
    procedure RefreshFieldsList;
    procedure AutoAddDefault;

    procedure NewListe;
    procedure LoadListe(Code : String);
    procedure SaveListe(Code,Lib : String);
    procedure BeforeDelete(Q : TQuery; var CanDelete : boolean);
    function CheckOverwrite(OldCode, OldLib, Code, Lib : String) : boolean;

  end;

const
   TitreRow = 0;
   ChampRow = 1;
   ExempleRow = 3;

implementation
uses HEnt1, UIUtil, SysUtils, Graphics, Choix, LookUp, OpenSaveDlg;

{$R *.DFM}

// Procédure d'ouverture de la fiche de paramétrage de liste de saisie
procedure EntreeListeSaisie(CodeListe : String);
var FF : TFListeSaisie;
    PPANEL : THPanel;
begin
SourisSablier;
FF := TFListeSaisie.Create(Application);
FF.CodeListe := CodeListe;

PPANEL := FindInsidePanel;
if PPANEL = nil then
  begin
  try
    FF.ShowModal;
  finally
    FF.Free;
  end;
  SourisNormale;
  end else
  begin
  InitInside(FF, PPANEL);
  FF.Show;
  end;
end;

// Procédure pour l'appel de la fiche depuis le script
procedure AGLEntreeListeSaisie(Parms : Array of Variant; Nb : Integer);
begin
EntreeListeSaisie(Parms[0]);
end;


//----------------- CLASSE LISTE DE COLONNES -----------------------------------

constructor TColonneList.Create(Grid : THGrid; SM : THSystemMenu; Fields : TOB);
begin
SetLength(FColumns,0);
FGrid := Grid;
FHMTrad := SM;
FAvailFields := Fields;
end;

destructor TColonneList.Destroy;
begin
SetLength(FColumns,0);
inherited;
end;


// Déplace les colonnes comprises entre [BgnIndex,EndIndex] vers [BgnIndex+Sens,EndIndex+Sens]
procedure TColonneList.SlideCols(BgnIndex, EndIndex, Sens : integer);
var i : integer;
begin
if Sens < 0 then for i := BgnIndex to EndIndex do FColumns[i+Sens] := FColumns[i]
            else for i := EndIndex downto BgnIndex do FColumns[i+Sens] := FColumns[i];
end;

// Retourne le nombre de colonnes
function TColonneList.GetColCount : integer;
begin
result := High(FColumns)+1;
end;

// Retourne la colonne d'index index
function TColonneList.GetCol(index : integer) : TColonne;
begin
result := FColumns[index];
end;

// Retourne la colonne sélectionnée sur le grid
function TColonneList.GetCurrentCol : TColonne;
begin
if FGrid = nil then result := FColumns[0]
               else result := FColumns[FGrid.Col];
end;

// Convertit un nom de type tel qu'on le trouve dans DECHAMPS en type TTypeChamp
function TColonneList.TypeFromType(StType : String) : TTypeChamp;
begin
if StType = 'INTEGER' then result := tcInteger else
if StType = 'DOUBLE' then result := tcDouble else
if StType = 'RATE' then result := tcRate else
if StType = 'BOOLEAN' then result := tcBoolean else
if StType = 'COMBO' then result := tcCombo else
if StType = 'DATE' then result := tcDate else
if Copy(StType,1,6) = 'VARCHAR' then result := tcVarchar else
if StType = 'BLOB' then result := tcBlob
                   else result := tcOther;
end;

// Génère le détail de la liste au format "decla"
function TColonneList.GenerateDetail : String;
var i, MaxWidth : integer;
    Fieldz, Captions, Widths, Aligns, Cols : String;
    Al, SM, Ob : Char;
begin
Fieldz := ''; Captions := ''; Widths := ''; Aligns := ''; Cols := '';
MaxWidth := 0;
for i := 0 to Count-1 do
 with FColumns[i] do
  if MaxWidth < Largeur then MaxWidth := Largeur;

for i := 0 to Count-1 do
 with FColumns[i] do
   begin
   Fieldz := Fieldz + NomChamp + ';';
   Captions := Captions + Titre + ';';
   Widths := Widths + inttostr((Largeur*100) div MaxWidth) + ';';

   case Align of
    taLeftJustify : Al := 'G';
    taCenter : Al := 'C';
    taRightJustify : Al := 'D';
   else
    Al := 'G';
   end;
   if SepMilliers then SM := '/' else SM := '.';
   if Obligatoire then Ob := 'O' else Ob := ' ';
   Aligns := Aligns + Al + SM + IntToStr(NbDecim) + Ob + ' ---;';

   Cols := Cols + inttostr(i+1) + ';';
   end;
if ctxScot in V_PGI.PGIContexte then Captions := TraduitGA(Captions);
with TStringList.Create do
  begin
  Clear;
  Add('LIGNE');
  Add(Fieldz);
  Add('');
  Add('');
  Add(Captions);
  Add(Widths);
  Add(Aligns);
  Add('');
  Add(Cols);
  result := Text;
  Free;
  end;
end;

// Construit une liste à partir du détail au format "decla"
// Les captions ne sont pas conservés
procedure TColonneList.LoadDetail(Dtl : String);
var S, Fieldz, Widths, Aligns : String;
    ColTOB : TOB;
begin
with TStringList.Create do
  begin
  Text := Dtl;
  Fieldz := Strings[1];
  Widths := Strings[5];
  Aligns := Strings[6];
  Free;
  end;

ClearList;

while Fieldz <> '' do
  begin
  S := ReadTokenSt(Fieldz);
  if Copy(s, 1, 3) = 'GL_' then 
    ColTOB := FAvailFields.FindFirst(['CO_ABREGE'], [Copy(S,4,Length(S))], false)
  else if Copy(s, 1, 4) = 'GLC_' then
    ColTOB := FAvailFields.FindFirst(['CO_ABREGE'], [S], false);
  if ColTOB = nil then
    begin
    ReadTokenSt(Widths);
    ReadTokenSt(Aligns);
    Continue;
    end;

  AddColumn(ColTOB, -1, false);
  with FColumns[High(FColumns)] do
    begin
    Largeur := StrToInt(ReadTokenSt(Widths));
    S := ReadTokenSt(Aligns);
    case S[1] of
      'G' : Align := taLeftJustify;
      'C' : Align := taCenter;
      'D' : Align := taRightJustify;
    end;
    NbDecim := StrToInt(S[3]);
    end;
  end;

ApplyToGrid;
end;


// Applique la définition de la liste au grid associé
procedure TColonneList.ApplyToGrid;
var i : integer;
begin
if FGrid = nil then exit;
with FGrid do
  begin
  ColCount := Count;
  for i := 0 to Count-1 do
    begin
    Cells[i,TitreRow] := FColumns[i].Titre;
    Cells[i,ChampRow] := '['+FColumns[i].NomChamp+']';
    if FColumns[i].Exemple = ''
      then Cells[i,ExempleRow] := '' else
      case FColumns[i].TypeChamp of
       tcInteger, tcBoolean, tcCombo, tcDate, tcVarchar, tcBlob, tcOther : Cells[i,ExempleRow] := FColumns[i].Exemple;
       tcDouble, tcRate : Cells[i,ExempleRow] := Format('%.'+inttostr(FColumns[i].NbDecim)+'f',[StrToFloat(FColumns[i].Exemple)]);
      end;
    ColWidths[i] := FColumns[i].Largeur;
    FGrid.ColAligns[i] := FColumns[i].Align;
    end;
  end;
ReAjustWidths;
end;

// Réajuste les largeurs de colonnes à la largeur du grid et mémorise les nouvelles largeurs
procedure TColonneList.ReAjustWidths;
begin
if FGrid = nil then exit;
if FHMTrad <> nil then
  begin FHMTrad.ResizeGridColumns(FGrid);
        GetGridWidths; end;
end;

// Mémorise dans la liste les largeurs des colonnes du grid
procedure TColonneList.GetGridWidths;
var i : integer;
begin
if FGrid = nil then exit;
for i := 0 to Count-1 do
 FColumns[i].Largeur := FGrid.ColWidths[i];
end;

// Retourne true si une colonne sur le grid a changé de largeur p/r à la liste
function TColonneList.WidthHasChanged : boolean;
var i : integer;
begin
result := false;
if FGrid = nil then exit;

for i := 0 to Count-1 do
 if FColumns[i].Largeur <> FGrid.ColWidths[i] then
   begin
   result := true;
   break;
   end;
end;


// Réinitialise la liste à vide
procedure TColonneList.ClearList;
begin
SetLength(FColumns,0);
end;

// Ajoute une colonne à la liste ; ColTOB = ligne de la tablette chargée dans LoadAvailableFields ;
//    Index = endroit où insérer ou -1 pour insérer à la fin ; ApplyNow = effectuer l'ApplyToGrid?
procedure TColonneList.AddColumn(ColTOB : TOB; Index : integer; ApplyNow : Boolean = true);
var Width : integer;
    TagList, Tag : String;
begin
if Count+1 > MaxColonnes then exit;

SetLength(FColumns, Count+1);
if Index = -1 then begin
                   Index := Count-1;
                   Width := FColumns[Index-1].Largeur;
                   if Width = 0 then Width := 100;
                   end
              else begin
                   Width := FColumns[Index].Largeur;
                   SlideCols(Index, Count-2, 1);
                   end;
with FColumns[Index] do
  begin
  Code := ColTOB.GetValue('CO_CODE');
  if pos('_', ColTOB.GetValue('CO_ABREGE')) = 0 then 
    NomChamp := 'GL_' + ColTOB.GetValue('CO_ABREGE')
  else
    NomChamp := ColTOB.GetValue('CO_ABREGE');
  Titre := ColTOB.GetValue('CO_LIBELLE');
  Exemple := '';
  Largeur := Width;
  Obligatoire := false;
  Figee := false;
  TypeChamp := TypeFromType(ColTOB.GetValue('DH_TYPECHAMP'));

  case TypeChamp of
    tcInteger : begin Align := taRightJustify; NbDecim := 0; SepMilliers := true; end;
    tcDouble, tcRate : begin Align := taRightJustify; NbDecim := 2; SepMilliers := true; end;
    tcVarchar, tcCombo, tcOther : begin Align := taLeftJustify; NbDecim := 0; SepMilliers := false; end;
    tcBlob, tcDate, tcBoolean : begin Align := taCenter; NbDecim := 0; SepMilliers := false; end;
   end;

  TagList := ColTOB.GetValue('CO_LIBRE');
  while TagList <> '' do
    begin
    Tag := ReadTokenSt(TagList);
    if Tag = 'O' then Obligatoire := true else
    if Tag = 'F' then Figee := true else
    if Copy(Tag,1,2) = 'E=' then Exemple := Copy(Tag,3,Length(Tag)) else
    if Tag[1] = 'A' then
      case Tag[2] of
       'G' : Align := taLeftJustify;
       'C' : Align := taCenter;
       'D' : Align := taRightJustify;
      end else
    if Tag[1] = 'V' then NbDecim := strtoint(Tag[2]);
    end;
  end;

if ApplyNow then ApplyToGrid;
end;

// Déplace une colonne de FromIndex vers ToIndex
procedure TColonneList.MoveColumn(FromIndex, ToIndex : integer);
var tmp : TColonne;
begin
if FromIndex = ToIndex then exit;

tmp := FColumns[FromIndex];
if FromIndex < ToIndex then SlideCols(FromIndex+1,ToIndex,-1)
                       else SlideCols(ToIndex,FromIndex-1,1);
FColumns[ToIndex] := tmp;
end;

// Supprime la colonne d'index ColIndex
procedure TColonneList.DelColumn(ColIndex : integer);
begin
if FColumns[ColIndex].Obligatoire then exit;
if (ColIndex < 0) or (ColIndex > Count-1) then exit;
SlideCols(ColIndex+1, Count-1, -1);
SetLength(FColumns, Count-1);

ApplyToGrid;
end;

// Retourne true si la colonne (dont le code tablette est Code) existe dans la liste
function TColonneList.ColumnExists(Code : String) : boolean;
var i : integer;
begin
result := false;
for i := 0 to Count-1 do
 if FColumns[i].Code = Code then
   begin
   result := true;
   break;
   end;
end;

// Si la colonne d'index ColNb est figée, retourne l'index de colonne suivant
//   immédiatement la derniere colonne figée.
// Retourne ColNb si ColNb n'est pas figée.
function TColonneList.SkipFiged(ColNb : integer) : integer;
var i : integer;
begin
result := ColNb;
if result = -1 then exit;
for i := ColNb to Count-1 do
 if FColumns[i].Figee then result := i+1;
end;

// Modifie l'alignement d'une colonne
procedure TColonneList.SetColAlign(ColNb : integer; Align : TAlignment);
begin
FColumns[ColNb].Align := Align;
if FGrid <> nil then
  begin
  FGrid.ColAligns[ColNb] := Align;
  ApplyToGrid;
  end;
end;

// Modifie le nombre de décimales d'une colonne
procedure TColonneList.SetColNbDec(ColNb, NbDec : integer);
begin
if FColumns[ColNb].NbDecim = NbDec then exit;
FColumns[ColNb].NbDecim := NbDec;
if FGrid <> nil then ApplyToGrid;
end;


//---------------- FICHE PRINCIPALE --------------------------------------------

procedure TFListeSaisie.FormCreate(Sender: TObject);
begin
FOldColX := Low(integer);  // Ca vaut ça quand on est pas en train de draguer
TOBAvailableFields := TOB.Create('no table',nil,-1);
LoadAvailableFields;
FWorkList := TColonneList.Create(G_Liste, HMTrad, TOBAvailableFields);
end;

procedure TFListeSaisie.FormDestroy(Sender: TObject);
begin
FWorkList.Free;
TOBAvailableFields.Free;
end;

procedure TFListeSaisie.FormShow(Sender: TObject);
begin
with G_Liste do begin RowHeights[0] := 18; RowHeights[2] := 4; end;
if CodeListe = '' then NewListe
                  else LoadListe(CodeListe);
SourisNormale;
end;

procedure TFListeSaisie.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if IsInside(Self) then Action := caFree;
end;


procedure TFListeSaisie.G_ListeCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
RefreshCurrColAligns;
end;

// Après déplacement d'une colonne sur le grid, déplace la colonne dans la liste
//  en veillant à ne pas la déplacer avant la fin des colonnes figées, et à ne pas
//  déplacer de colonne figée.
procedure TFListeSaisie.G_ListeColumnMoved(Sender: TObject; FromIndex, ToIndex: Integer);
var SkipIndex : integer;
    b : boolean;
begin
SkipIndex := FWorkList.SkipFiged(ToIndex);
if (FromIndex < SkipIndex) and (SkipIndex <> ToIndex) then dec(SkipIndex);
b := (SkipIndex <> ToIndex);

if FWorkList.Columns[FromIndex].Figee then b := true
                                      else FWorkList.MoveColumn(FromIndex,SkipIndex);
if b then FWorkList.ApplyToGrid; // On rafraichit le grid que si y'a besoin
end;

// Détection de changement d'une largeur de colonne quand on relache le bouton souris
procedure TFListeSaisie.G_ListeMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
if FWorkList.WidthHasChanged then FWorkList.ReAjustWidths;
end;

// Affichage de la barre à l'endroit d'insertion quand on est en train de draguer
procedure TFListeSaisie.G_ListeDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
var ACol, ARow : longint;
    ColX : integer;
begin
Accept := Source is TListBox;
if Accept then
  begin
  with THGrid(Sender) do
    begin
    MouseToCell(X, 0, ACol, ARow);
    ACol := FWorkList.SkipFiged(ACol); // Ne pas insérer avant la fin des colonnes figées
    if ACol = FWorkList.Count then ACol := -1;
    if ACol = -1 then ColX := CellRect(ColCount-1,0).Right
                 else ColX := CellRect(ACol,0).Left;

    if FOldColX <> ColX then
      begin
      if FOldColX > Low(integer) then DrawPositionLine(FOldColX); // Réafficher la ligne à l'ancienne position = l'effacer (XOR)
      DrawPositionLine(ColX);
      FOldColX := ColX;
      end;
    end;
  end;
end;

// Fin du draguage : droppage de colonne
procedure TFListeSaisie.G_ListeDragDrop(Sender, Source: TObject; X, Y: Integer);
var ACol, ARow : longint;
begin
if (Sender is THGrid) and (Source is TListBox) then
  with TListBox(Source) do
    begin
    if FOldColX > Low(integer) then DrawPositionLine(FOldColX);
    FOldColX := Low(integer);

    if ItemIndex > -1 then
      if FWorkList.Count+1 > MaxColonnes
        then PGIInfo(MsgCaptions.Mess[0]+inttostr(MaxColonnes)+'.'#13+MsgCaptions.Mess[1],Caption)
        else begin
             THGrid(Sender).MouseToCell(X, 0, ACol, ARow);
             ACol := FWorkList.SkipFiged(ACol);
             if ACol = FWorkList.Count then ACol := -1;
             FWorkList.AddColumn(TOB(Items.Objects[ItemIndex]), ACol);
             RefreshFieldsList;
             end;
    end;
end;

// Début du draguage à partir de la liste de champs
procedure TFListeSaisie.FieldsListMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
if (Button = mbLeft) and (FieldsList.ItemIndex > -1) then FieldsList.BeginDrag(false);
end;

// Fin du draguage : effacement de la barre d'insertion
procedure TFListeSaisie.FieldsListEndDrag(Sender, Target: TObject; X, Y: Integer);
begin
if FOldColX > Low(integer) then DrawPositionLine(FOldColX);
FOldColX := Low(integer);
end;

procedure TFListeSaisie.tbDelColClick(Sender: TObject);
begin
FWorkList.DelColumn(G_Liste.Col);
RefreshFieldsList;
end;

procedure TFListeSaisie.tbLeftClick(Sender: TObject);
begin
FWorkList.SetColAlign(G_Liste.Col, taLeftJustify);
end;

procedure TFListeSaisie.tbCenterClick(Sender: TObject);
begin
FWorkList.SetColAlign(G_Liste.Col, taCenter);
end;

procedure TFListeSaisie.tbRightClick(Sender: TObject);
begin
FWorkList.SetColAlign(G_Liste.Col, taRightJustify);
end;

procedure TFListeSaisie.SpinDecChange(Sender: TObject);
begin
FWorkList.SetColNbDec(G_Liste.Col, SpinDec.Value);
end;

procedure TFListeSaisie.tbNewClick(Sender: TObject);
begin
if PGIAsk(MsgCaptions.Mess[2], Caption) = mrYes
  then NewListe;
end;

procedure TFListeSaisie.tbLoadClick(Sender: TObject);
var Code, Lib : String;
begin
if ctxAffaire in V_PGI.PGIContexte then
   BEGIN
   if PGIOpenSave(MsgCaptions.Mess[5], 'LISTE', 'LI_LISTE', 'LI_LIBELLE', 'LI_LISTE LIKE "AFSAISIEF%"', '', Code, Lib, BeforeDelete, false)
   and (PGIAsk(MsgCaptions.Mess[2], Caption) = mrYes) then LoadListe(Code);
   END
else
   BEGIN
   if PGIOpenSave(MsgCaptions.Mess[5], 'LISTE', 'LI_LISTE', 'LI_LIBELLE', 'LI_LISTE LIKE "GCSAISIE%"', '', Code, Lib, BeforeDelete, false)
   and (PGIAsk(MsgCaptions.Mess[2], Caption) = mrYes) then LoadListe(Code);
   END;
end;

procedure TFListeSaisie.BValiderClick(Sender: TObject);
var CL,LL : String;
begin
CL := CodeListe; LL := LibListe;
if ctxAffaire in V_PGI.PGIContexte then
   BEGIN
   if PGIOpenSave(MsgCaptions.Mess[6], 'LISTE', 'LI_LISTE', 'LI_LIBELLE', 'LI_LISTE LIKE "AFSAISIEF%"', 'AFSAISIEF', CL, LL, BeforeDelete, true)
   and CheckOverwrite(CodeListe, LibListe, CL, LL) then SaveListe(CL, LL);
   END
else
   BEGIN
   if PGIOpenSave(MsgCaptions.Mess[6], 'LISTE', 'LI_LISTE', 'LI_LIBELLE', 'LI_LISTE LIKE "GCSAISIE%"', 'GCSAISIE', CL, LL, BeforeDelete, true)
   and CheckOverwrite(CodeListe, LibListe, CL, LL) then SaveListe(CL, LL);
   END;
end;


// Retourne l'index d'un champ dans la liste de champs (code tablette)
function TFListeSaisie.AFInList(Code : String) : integer;
begin
result := FieldsList.Items.IndexOfObject(TOBAvailableFields.FindFirst(['CO_CODE'], [Code], false));
end;

// Affiche la barre d'insertion en XORPut
procedure TFListeSaisie.DrawPositionLine(ColX : integer);
var oc : TColor;
    om : TPenMode;
begin
with G_Liste do
  begin
  oc := Canvas.Pen.Color; Canvas.Pen.Color := clLtGray;
  om := Canvas.Pen.Mode; Canvas.Pen.Mode := pmXor;
  Canvas.Pen.Width := 5;
  Canvas.PolyLine([Point(ColX,0),Point(ColX,ClientHeight)]);
  Canvas.Pen.Mode := om;
  Canvas.Pen.Color := oc;
  end;
end;

// Actualise les boutons d'alignement et le nb de décimales
procedure TFListeSaisie.RefreshCurrColAligns;
begin
  case FWorkList.CurrentColumn.Align of
    taLeftJustify : tbLeft.Down := true;
    taCenter : tbCenter.Down := true;
    taRightJustify : tbRight.Down := true;
  end;
  SpinDec.Value := FWorkList.CurrentColumn.NbDecim;
  SpinDec.Visible := FWorkList.CurrentColumn.TypeChamp in [tcDouble, tcRate];
  tbDelCol.Enabled := not FWorkList.CurrentColumn.Obligatoire;
end;

// Titre de la fenetre
procedure TFListeSaisie.SetTitle;
begin
PEntete.Caption := CodeListe + ' - ' + LibListe;
end;


// Charge les champs disponibles à partir de la tablette -> dans une TOB
procedure TFListeSaisie.LoadAvailableFields;
var Q : TQuery;
    TypeTablette : string;
begin
if ctxAffaire in V_PGI.PGIContexte then TypeTablette := 'ASP'
                                   else TypeTablette := 'GSP';
(*
Q := OpenSQL('SELECT CO_CODE, CO_LIBELLE, CO_ABREGE, CO_LIBRE, DH_TYPECHAMP '+
             'FROM COMMUN LEFT JOIN DECHAMPS ON DH_NOMCHAMP="GL_"||CO_ABREGE '+
             'WHERE CO_TYPE="' + TypeTablette + '"', true);
*)
Q := OpenSQL('SELECT CO_CODE, CO_LIBELLE, CO_ABREGE, CO_LIBRE, DH_TYPECHAMP ' +
             'FROM COMMUN LEFT JOIN DECHAMPS ON DH_NOMCHAMP=IIF(SUBSTRING(CO_ABREGE,1,4)="GLC_", CO_ABREGE, "GL_"||CO_ABREGE) ' +
             'WHERE CO_TYPE="' + TypeTablette + '"', true,-1, '', True);
TOBAvailableFields.LoadDetailDB('COMMUN','','',Q, false);
Ferme(Q);
end;

// Met à jour la liste de champs en y enlevant les champs présents dans la liste de colonnes
//    et en y ajoutant ceux qui n'y sont plus
procedure TFListeSaisie.RefreshFieldsList;
var i,inl : integer;
    FCode,FLib : String;
    FCurTOB : TOB;
begin
for i := 0 to TOBAvailableFields.Detail.Count-1 do
  begin
  FCurTOB := TOBAvailableFields.Detail[i];
  FCode := FCurTOB.GetValue('CO_CODE');
  FLib := FCurTOB.GetValue('CO_LIBELLE');
  inl := AFInList(FCode);
  if FWorkList.ColumnExists(FCode) and (inl > -1) then FieldsList.Items.Delete(inl);
  if (not FWorkList.ColumnExists(FCode)) and (inl = -1) then FieldsList.Items.AddObject(FLib,FCurTOB);
  end;
RefreshCurrColAligns;
end;

// Ajoute tous les champs possédant un tag Dxx dans le bon ordre
procedure TFListeSaisie.AutoAddDefault;
var i : integer;
    FCurTOB : TOB;
    TagList, Tag : String;
    DefCols : Array of TOB;
begin
SetLength(DefCols, TOBAvailableFields.Detail.Count);
for i := Low(DefCols) to High(DefCols) do DefCols[i] := nil;

for i := 0 to TOBAvailableFields.Detail.Count-1 do
  begin
  FCurTOB := TOBAvailableFields.Detail[i];

  TagList := FCurTOB.GetValue('CO_LIBRE');
  while TagList <> '' do
    begin
    Tag := ReadTokenSt(TagList);
    if Tag[1] = 'D' then DefCols[StrToInt(Copy(Tag,2,Length(Tag)))] := FCurTOB;
    end;
  end;

for i := Low(DefCols) to High(DefCols) do
  if DefCols[i] <> nil then FWorkList.AddColumn(TOB(DefCols[i]), -1);

SetLength(DefCols, 0);
end;


// Crée une nouvelle liste avec les colonnes par défaut
procedure TFListeSaisie.NewListe;
begin
if ctxAffaire in V_PGI.PGIContexte then CodeListe := 'AFSAISIEFNEW'
                                   else CodeListe := 'GCSAISIENEW';
LibListe := 'Nouvelle liste de saisie';
FWorkList.ClearList;
AutoAddDefault;
RefreshFieldsList;
SetTitle;
end;

// Charge une liste à partir de son code
procedure TFListeSaisie.LoadListe(Code : String);
var Q : TQuery;
begin
Q := OpenSQL('SELECT LI_DATA, LI_LIBELLE FROM LISTE WHERE LI_LISTE="'+Code+'"', true,-1, '', True);
if not Q.EOF then
  begin
  FWorkList.Detail := Q.FindField('LI_DATA').AsString;
  CodeListe := Code;
  LibListe := Q.FindField('LI_LIBELLE').AsString;
  end;
Ferme(Q);
RefreshFieldsList;
SetTitle;
end;

// Sauve la liste courante sous le Code & Libellé en param.
procedure TFListeSaisie.SaveListe(Code,Lib : String);
//var SaveTOB : TOB;
begin
Code := UpperCase(Code);
with TOB.Create('LISTE', nil, -1) do
  begin
  AddChampSup('LI_LISTE', false); PutValue('LI_LISTE', Code);
  AddChampSup('LI_UTILISATEUR', false); PutValue('LI_UTILISATEUR', '---');
  AddChampSup('LI_LIBELLE', false); PutValue('LI_LIBELLE', Lib);
  AddChampSup('LI_SOCIETE', false); PutValue('LI_SOCIETE', V_PGI.CodeSociete);
  AddChampSup('LI_NUMOK', false); PutValue('LI_NUMOK', 'X');
  AddChampSup('LI_TRIOK', false); PutValue('LI_TRIOK', 'X');
  AddChampSup('LI_LANGUE', false); PutValue('LI_LANGUE', V_PGI.LanguePrinc);
  AddChampSup('LI_DATA', false); PutValue('LI_DATA', FWorkList.Detail);

  InsertOrUpdateDB;

  Free;
  end;
CodeListe := Code;
LibListe := Lib;
SetTitle;
end;

// Exécuté avant effacement d'une liste pour demander confirmation
procedure TFListeSaisie.BeforeDelete(Q : TQuery; var CanDelete : boolean);
begin
CanDelete := (PGIAsk(MsgCaptions.Mess[4]+Q.FindField('LI_LISTE').AsString+' ?', Caption) = mrYes);
end;

// Vérifie si la liste existe déjà ou quoi, renvoir True si peut enregistrer
function TFListeSaisie.CheckOverwrite(OldCode, OldLib, Code, Lib : String) : boolean;
begin
result := (OldCode = Code)
          or (not ExisteSQL('SELECT LI_LISTE FROM LISTE WHERE LI_LISTE="'+Code+{'" AND LI_LIBELLE<>"'+Lib+}'"'))
          or (PGIAsk(MsgCaptions.Mess[3], Caption) = mrYes);
end;


procedure TFListeSaisie.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

Initialization
RegisterAGLProc('EntreeListeSaisie',false,1,AGLEntreeListeSaisie);
end.
