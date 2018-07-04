{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 24/04/2003
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit ParamSaisEff;

interface

uses
  Forms, Windows, Menus, Dialogs, HSysMenu, Grids, Hctrls, HTB97, Controls, ExtCtrls,
  HPanel, Classes, StdCtrls, M3FP, UTOB, Spin, Mask, hmsgbox;

(*
Tablette :
  CO_LIBELLE : Titre de colonne
  CO_ABREGE :  Nom du champ sans le préfixe E_
  CO_LIBRE :   Paramètres
                  Dnn :    Champ ajouté par défaut dans une nouvelle liste ;
                           nn = n° d'ordre d'ajout
                  O :      Champ obligatoire
                  F :      Colonne figée (non déplacable, doit être en tête de liste)
                  FF :     Colonne figée en fin de liste (déplacable, mais doit être en tête de liste)
                  E=xxxx : Affiche xxxx dans la ligne exemple
                  Ax :     Alignement, x = G (à gauche), C (centré) ou D (à droite)
                  Vn :     Nombre de décimales = n
*)


// Passer le code de la liste à ouvrir au lancement de la fiche, ou chaine vide
//   pour lancer avec la liste par défaut.
procedure ParamSaisieListeEff(CCM : String);

const
  MaxColonnes = 16;

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
    FigeeFinListe : Boolean;
    Visible : boolean;
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

  TFParamSaisEff = class(TForm)
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
    HM: THMsgBox;
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

  private
    FOldColX : integer;
    FWorkList : TColonneList;
    TOBAvailableFields : TOB;

    function AFInList(Code : String) : integer;
    procedure DrawPositionLine(ColX : integer);
    procedure RefreshCurrColAligns;

  public
    CodeListe, LibListe : String;

    procedure LoadAvailableFields;
    procedure RefreshFieldsList;
    procedure AutoAddDefault;

    procedure NewListe;
    procedure LoadListe(Code : String);
    procedure SaveListe(Code,Lib : String);
    function CheckOverwrite(Code, Lib : String) : boolean;

  end;


implementation
uses HEnt1, UIUtil,
{$IFDEF EAGLCLIENT}

{$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  SysUtils, Graphics, Choix, LookUp;

{$R *.DFM}

Var CodeCommun : String;

// Procédure d'ouverture de la fiche de paramétrage de liste de saisie
procedure ParamSaisieListeEff(CCM : String);
var FF : TFParamSaisEff;
    PPANEL : THPanel;
begin
SourisSablier;
CodeCommun := CCM ;
FF := TFParamSaisEff.Create(Application);
//FF.CodeListe := CodeListe ;

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
(*
procedure AGLEntreeListeSaisie(Parms : Array of Variant; Nb : Integer);
begin
EntreeListeSaisie(Parms[0]);
end;
*)

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
    Al, SM, Ob, Vi : Char;
begin
Fieldz := ''; Captions := ''; Widths := ''; Aligns := ''; Cols := '';
MaxWidth := 0;
for i := 0 to Count-1 do
 with FColumns[i] do
  if MaxWidth < Largeur then MaxWidth := Largeur;

for i := 0 to Count-1 do
 with FColumns[i] do
   begin
   //SG6 30/12/2004
   if (NomChamp = 'E_EXERCICE') or (NomChamp='E_NUMEROPIECE') or (NomChamp = 'E_DATECOMPTABLE') then Visible := False;

   Fieldz := Fieldz + NomChamp + ';';
   Captions := Captions + Titre + ';';
   // 13865
   // BPY le 01/12/2004
   if ((Obligatoire) and (Largeur = 0) and (Visible)) then Largeur := 35;
   if (not visible) then Largeur := -1;
   // Fin BPY

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
   if Visible then Vi := '-' else Vi := 'X'; 
   Aligns := Aligns + Al + SM + IntToStr(NbDecim) + Ob + ' ' + Vi + '--;';

   Cols := Cols + inttostr(i+1) + ';';
   end;

with TStringList.Create do
  begin
  Clear;
  Add('ECRITURE');
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
var
	S, Fieldz, Widths, Aligns : String;
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
    	// recup du champs ...
	    S := ReadTokenSt(Fieldz);
    	ColTOB := FAvailFields.FindFirst(['CO_ABREGE'], [Copy(S,3,Length(S)-2)], false);
        // si pas de champs on continue
	    if ColTOB = nil then
	    begin
    		ReadTokenSt(Widths);
		    ReadTokenSt(Aligns);
	    	Continue;
	    end;
        // sinon ....

	    AddColumn(ColTOB, -1, false);
	    with FColumns[High(FColumns)] do
	    begin
		    // largeur
		    Largeur := StrToInt(ReadTokenSt(Widths));
		    // autre option :
		    S := ReadTokenSt(Aligns);
		    //                - alignement
		    case S[1] of
			    'G' : Align := taLeftJustify;
			    'C' : Align := taCenter;
			    'D' : Align := taRightJustify;
		    end;
		    //                - visibilité
		    // BPY le 01/12/2004
            Obligatoire := false;
			if (Length(S) >= 4) then Obligatoire := (S[4] = 'O');
            Visible := true;
		    if (Length(S) >= 6) then Visible := (S[6] = '-');
		    // fin BPY
		    //                - nb de decimal
		    NbDecim := 0;
		    if ((Length(S) >= 3) and (S[3] in ['0'..'9'])) then NbDecim := StrToInt(S[3]);
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
    Cells[i,0] := FColumns[i].Titre;
    Cells[i,1] := '['+FColumns[i].NomChamp+']';
    if FColumns[i].Exemple = ''
      then Cells[i,2] := '' else
      case FColumns[i].TypeChamp of
       tcInteger, tcBoolean, tcCombo, tcDate, tcVarchar, tcBlob, tcOther : Cells[i,2] := FColumns[i].Exemple;
       tcDouble, tcRate : Cells[i,2] := Format('%.'+inttostr(FColumns[i].NbDecim)+'f',[StrToFloat(FColumns[i].Exemple)]);
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
  NomChamp := 'E_' + ColTOB.GetValue('CO_ABREGE');
  Titre := ColTOB.GetValue('CO_LIBELLE');
  Exemple := '';
  Largeur := Width;
  Obligatoire := false;
  Figee := false; FigeeFinListe:=FALSE ;
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
    If Tag = 'FF' Then FigeeFinListe:=TRUE Else
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

procedure TFParamSaisEff.FormCreate(Sender: TObject);
begin
FOldColX := Low(integer);  // Ca vaut ça quand on est pas en train de draguer
TOBAvailableFields := TOB.Create('no table',nil,-1);
LoadAvailableFields;
FWorkList := TColonneList.Create(G_Liste, HMTrad, TOBAvailableFields);
end;

procedure TFParamSaisEff.FormDestroy(Sender: TObject);
begin
FWorkList.Free;
TOBAvailableFields.Free;
end;

procedure TFParamSaisEff.FormShow(Sender: TObject);
begin
with G_Liste do RowHeights[0] := 18;
If CodeCommun='CSF' Then
  BEGIN
  CodeListe := 'CPSAISIEEFFA';
  LibListe := 'Liste saisie des effets en retour d''acceptation';
  END Else If CodeCommun='CSG' Then
  BEGIN
  CodeListe := 'CPSAISIEEFFB';
  LibListe := 'Liste saisie des chèques';
  END Else If CodeCommun='CSH' Then
  BEGIN
  CodeListe := 'CPSAISIEEFFC';
  LibListe := 'Liste saisie des cartes bleues';
  END Else If CodeCommun='CSI' Then
  BEGIN
  CodeListe := 'CPSAISIEEFFD';
  LibListe := 'Liste saisie trésorerie client';
  END Else If CodeCommun='CSC' Then
  BEGIN
  CodeListe := 'CPSAISIEEFFE';
  LibListe := 'Liste saisie trésorerie fournisseur';
  END ;
if CodeListe = '' then NewListe
                  else LoadListe(CodeListe);
SourisNormale;
end;

procedure TFParamSaisEff.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if IsInside(Self) then Action := caFree;
end;


procedure TFParamSaisEff.G_ListeCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
RefreshCurrColAligns;
end;

// Après déplacement d'une colonne sur le grid, déplace la colonne dans la liste
//  en veillant à ne pas la déplacer avant la fin des colonnes figées, et à ne pas
//  déplacer de colonne figée.
procedure TFParamSaisEff.G_ListeColumnMoved(Sender: TObject; FromIndex, ToIndex: Integer);
var SkipIndex : integer;
    b : boolean;
begin
SkipIndex := FWorkList.SkipFiged(ToIndex);
if (FromIndex < SkipIndex) and (SkipIndex <> ToIndex) then dec(SkipIndex);
b := (SkipIndex <> ToIndex);

if FWorkList.Columns[FromIndex].Figee Or FWorkList.Columns[FromIndex].FigeeFinListe Or 
   FWorkList.Columns[SkipIndex].FigeeFinListe then b := true
                                      else FWorkList.MoveColumn(FromIndex,SkipIndex);
if b then FWorkList.ApplyToGrid; // On rafraichit le grid que si y'a besoin
end;

// Détection de changement d'une largeur de colonne quand on relache le bouton souris
procedure TFParamSaisEff.G_ListeMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
if FWorkList.WidthHasChanged then FWorkList.ReAjustWidths;
end;

// Affichage de la barre à l'endroit d'insertion quand on est en train de draguer
procedure TFParamSaisEff.G_ListeDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
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
    if FWorkList.Columns[ACol].Figee Or FWorkList.Columns[ACol].FigeeFinListe then BEGIN Accept:=FALSE ; Exit ; END ;
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
procedure TFParamSaisEff.G_ListeDragDrop(Sender, Source: TObject; X, Y: Integer);
var ACol, ARow : longint;
begin
if (Sender is THGrid) and (Source is TListBox) then
  with TListBox(Source) do
    begin
    if FOldColX > Low(integer) then DrawPositionLine(FOldColX);
    FOldColX := Low(integer);

    if ItemIndex > -1 then
      if FWorkList.Count+1 > MaxColonnes
        then PGIInfo(HM.Mess[0]+inttostr(MaxColonnes)+'.'#13+HM.Mess[1],Caption)
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
procedure TFParamSaisEff.FieldsListMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
if (Button = mbLeft) and (FieldsList.ItemIndex > -1) then FieldsList.BeginDrag(false);
end;

// Fin du draguage : effacement de la barre d'insertion
procedure TFParamSaisEff.FieldsListEndDrag(Sender, Target: TObject; X, Y: Integer);
begin
if FOldColX > Low(integer) then DrawPositionLine(FOldColX);
FOldColX := Low(integer);
end;

procedure TFParamSaisEff.tbDelColClick(Sender: TObject);
begin
FWorkList.DelColumn(G_Liste.Col);
RefreshFieldsList;
end;

procedure TFParamSaisEff.tbLeftClick(Sender: TObject);
begin
FWorkList.SetColAlign(G_Liste.Col, taLeftJustify);
end;

procedure TFParamSaisEff.tbCenterClick(Sender: TObject);
begin
FWorkList.SetColAlign(G_Liste.Col, taCenter);
end;

procedure TFParamSaisEff.tbRightClick(Sender: TObject);
begin
FWorkList.SetColAlign(G_Liste.Col, taRightJustify);
end;

procedure TFParamSaisEff.SpinDecChange(Sender: TObject);
begin
FWorkList.SetColNbDec(G_Liste.Col, SpinDec.Value);
end;

procedure TFParamSaisEff.tbNewClick(Sender: TObject);
begin
if PGIAsk(HM.Mess[2], Caption) = mrYes
  then NewListe;
end;

procedure TFParamSaisEff.tbLoadClick(Sender: TObject);
//var Code : String;
begin
(*
Code := Choisir(HM.Mess[4],'LISTE','LI_LIBELLE','LI_LISTE','LI_LISTE LIKE "GCSAISIE%"','');
if PGIAsk(HM.Mess[2], Caption) = mrYes
  then LoadListe(Code);
*)
end;

procedure TFParamSaisEff.BValiderClick(Sender: TObject);
var CL,LL : String;
begin
CL := CodeListe; LL := LibListe;
If PGIAsk(HM.Mess[3],Caption) = mrYes then SaveListe(CL, LL);
end;


// Retourne l'index d'un champ dans la liste de champs (code tablette)
function TFParamSaisEff.AFInList(Code : String) : integer;
begin
result := FieldsList.Items.IndexOfObject(TOBAvailableFields.FindFirst(['CO_CODE'], [Code], false));
end;

// Affiche la barre d'insertion en XORPut
procedure TFParamSaisEff.DrawPositionLine(ColX : integer);
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
procedure TFParamSaisEff.RefreshCurrColAligns;
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

// Charge les champs disponibles à partir de la tablette -> dans une TOB
procedure TFParamSaisEff.LoadAvailableFields;
var Q : TQuery;
begin
Q := OpenSQL('SELECT CO_CODE, CO_LIBELLE, CO_ABREGE, CO_LIBRE, DH_TYPECHAMP '+
             'FROM COMMUN LEFT JOIN DECHAMPS ON DH_NOMCHAMP="E_"||CO_ABREGE '+
             'WHERE CO_TYPE="'+CodeCommun+'"', true);
TOBAvailableFields.LoadDetailDB('COMMUN','','',Q, false);
Ferme(Q);
end;

// Met à jour la liste de champs en y enlevant les champs présents dans la liste de colonnes
//    et en y ajoutant ceux qui n'y sont plus
procedure TFParamSaisEff.RefreshFieldsList;
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
procedure TFParamSaisEff.AutoAddDefault;
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
procedure TFParamSaisEff.NewListe;
begin
If CodeCommun='CSF' Then
  BEGIN
  CodeListe := 'CPSAISIEEFFA';
  LibListe := 'Liste saisie des effets en retour d''acceptation';
  END Else If CodeCommun='CSG' Then
  BEGIN
  CodeListe := 'CPSAISIEEFFB';
  LibListe := 'Liste saisie des chèques';
  END Else If CodeCommun='CSH' Then
  BEGIN
  CodeListe := 'CPSAISIEEFFC';
  LibListe := 'Liste saisie des cartes bleues';
  END Else If CodeCommun='CSI' Then
  BEGIN
  CodeListe := 'CPSAISIEEFFD';
  LibListe := 'Liste saisie trésorerie client';
  END Else If CodeCommun='CSC' Then
  BEGIN
  CodeListe := 'CPSAISIEEFFE';
  LibListe := 'Liste saisie trésorerie fournisseur';
  END ;
FWorkList.ClearList;
AutoAddDefault;
RefreshFieldsList;
end;

// Charge une liste à partir de son code
procedure TFParamSaisEff.LoadListe(Code : String);
var Q : TQuery;
begin
Q := OpenSQL('SELECT LI_DATA, LI_LIBELLE FROM LISTE WHERE LI_LISTE="'+Code+'"', true);
if not Q.EOF then
  begin
  FWorkList.Detail := Q.FindField('LI_DATA').AsString;
  CodeListe := Code;
  LibListe := Q.FindField('LI_LIBELLE').AsString;
  end;
Ferme(Q);
RefreshFieldsList;
end;

// Sauve la liste courante sous le Code & Libellé en param.
procedure TFParamSaisEff.SaveListe(Code,Lib : String);
var SaveTOB : TOB;
begin
Code := UpperCase(Code);
SaveTob := TOB.Create('LISTE', nil, -1);
with SaveTob do
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
end;

// Vérifie si la liste existe déjà ou quoi
function TFParamSaisEff.CheckOverwrite(Code, Lib : String) : boolean;
begin
result := (not ExisteSQL('SELECT LI_LISTE FROM LISTE WHERE LI_LISTE="'+Code+'" AND LI_LIBELLE<>"'+Lib+'"'))
          or (PGIAsk(HM.Mess[3], Caption) = mrYes);
end;


Initialization
//RegisterAGLProc('EntreeListeSaisie',false,1,AGLEntreeListeSaisie);
end.
