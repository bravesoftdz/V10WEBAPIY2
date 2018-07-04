{***********UNITE*************************************************
Auteur  ...... : TJ
Créé le ...... : 16/01/2007
Modifié le ... :   /  /
Description .. : Class pour la fenêtre Outil et saisie en grille de la fiche de
Suite ........ : révision : Détermination du résultat fiscal (DRF)
Mots clefs ... :
*****************************************************************}
unit OutilsDRF;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
{$IFNDEF DBXPRESS}dbtables, {$ELSE}uDbxDataSet, {$ENDIF}
  mul,
{$ELSE}
  eMul,
{$ENDIF}
  uTob,
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  Types, //Trect
  hpanel,
  vierge,
  ent1, //VH^, ...
  Windows, //Formkeydown
  HTB97, //Ttoolwindows97, .....
  Graphics, //[FsBold]...
  Grids, //propriétés grille
{$IFDEF AMORTISSEMENT}
  ImOutGen, //CalculPMValue
{$ENDIF}
  UTOF;

  // objet pour le passage vers ETafi
type TzDRF = class
  public
    TobZDRF: Tob;
    LeFormulaire: string;
    Lexo: string;

    constructor Create(Formulaire: string; Exercice: string);
    destructor Destroy; override;
    function GetValue(LeChamp: string): Variant;
    function GetDetail(LeChamp: string): Tob;

  end;


  //Objet pour les outils du formulaire DRF
type TDRFToolBox = class(TToolWindow97)
  private
    PiedGrille: integer; // ligne des totaux

    function CreeBouton(LePanel: ThPanel; Gly: TBitmap; Lenom: string): TToolbarButton97;
    function CreeLabel(Lenom: string): ThLabel;
    function CreeEdit(Lenom: string): THEdit;
    function CreePanel: THPanel;
    function CreeGrille: THGrid;
    function CreeLaCoche: THCheckbox;
    procedure ActivLabel(LeLaBel: THLabel; LeCaption: string; LeTop: integer; LeLeft: integer; LeHeight: integer = 50; LeWidth: integer = 100);
    procedure ActivGrille(LaGrille: THGrid; NbCol: integer = 6; NbRow: integer = 1; LeTop: integer = 60; LeLeft: integer = 0; LeHeight: integer = 200; LeWidth: integer = 100);
    procedure ActivEdit(Ledit: THEdit; LeHint: string; LeTop: integer; LeLeft: integer);
    procedure ActivBouton(LeButton: TToolbarButton97; LeHint: string; LeTop: integer; LeLeft: integer);
    procedure ActivCoche(LaCoche: THCheckbox; LeTitre: string; LeTop: integer = 10; LeLeft: integer = 10);
{$IFDEF AMORTISSEMENT}
    procedure ChargeImo(TobV: TOB);
{$ENDIF}
    function RegimeSocieteMere: Extended;
    procedure RepriseSaisieUnique(LaTob: Tob; LaNat: string; LeCom: integer = -1; LeMt1: integer = -1);

    procedure LaGrilleAjouteLigne(Index: integer);
    procedure LaGrilleSupprimeLigne(Index: integer);
    procedure LaGrilleOnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure LaGrilleAjouteLigneOnClick(Sender: Tobject);
    procedure LaGrilleSupprimeLigneOnClick(Sender: Tobject);
    procedure LaGrilleOnRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure LaGrilleOnMouseWheelDown(Sender: Tobject; Shift: TShiftState; MousePos: TPoint; var Handled: boolean);
    procedure LaGrillePostDrawCell(Acol: integer; ARow: integer; Canvas: TCanvas; Astate: TGridDrawState);


    procedure BDetail_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure B2058A_XL_MOA_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure B2058A_WC_MOA_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure B2058A_XA_MOA_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure B2058B_PCP_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure B2058A_WN_MOA_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure B2058A_WM_MOA_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure B2058A_VA_MOA_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure B2058AB_AH_MOA_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure B2058AB_AG_MOA_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure B2058AB_AF_MOA_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure BJetonPresence_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure BDetail_OnCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure BJetonPresence_LaCocheOnClick(Sender: Tobject);
    procedure CessionImmo_OnCellexit(Sender: Tobject; var ACol, ARow: Integer; var Cancel: Boolean);

    procedure B2058A_XL_MOA_CalculLigne(Index: integer);
    procedure B2058A_WC_MOA_CalculLigne(Index: integer);
    procedure B2058A_XA_MOA_CalculLigne(Index: integer);
    procedure B2058B_PCP_CalculLigne(Index: integer);
    procedure B2058A_WN_MOA_CalculLigne(Index: integer);
    procedure B2058A_WM_MOA_CalculLigne(Index: integer);
    procedure B2058A_VA_MOA_CalculLigne(Index: integer);
    procedure B2058AB_AH_MOA_CalculLigne(Index: integer);
    procedure B2058AB_AG_MOA_CalculLigne(Index: integer);
    procedure B2058AB_AF_MOA_CalculLigne(Index: integer);
    procedure BJetonPresence_CalculLigne;
    procedure BDetail_CalculLigne(Index: integer);
    procedure CessionImmo_CalculLigne(Index: integer);

  public
    LaGrille: THGrid;
    LeLabel1: THLabel;
    LeLabel2: THLabel;
    LeLabel3: THLabel;
    LeLabel4: THLabel;
    LeLabel5: THLabel;
    LeLabel6: THLabel;
    LeLabel7: THLabel;
    LePanel: THPanel;
    LEdit1: THEdit;
    LEdit2: THEdit;
    BValide: TToolbarButton97;
    BAnnule: TToolbarButton97;
    BAjouteLg: TToolbarButton97;
    BSupprimeLg: TToolbarButton97;
    LaCoche: THCheckbox;
//    THImmo              : THalleyWindow;
    ValeurPVLT: extended;
    RemConjointMin: Extended;
    RemConjointMax: Extended;
    TheTool: string;
    TheTag: integer;
    BaseJeton: Extended;

    constructor Create(mParent: TWinControl; G1, G2, G3, G4: TBitmap); reintroduce; overload; { FQ 20531 BVE 08.06.07 }
    destructor Destroy; override;

    procedure TypeToolBox(LeTool: string; LaTob: Tob; LeTitre: string = ''; Borne1: extended = 0; Borne2: extended = 0; LeTag: integer = 6);

  published
    procedure FormOnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);

  end;

procedure ChargeTobDRF(TobGene: TOB; Exo: string);


implementation

uses TntStdCtrls;

{ DRFToolBox }

constructor TDRFToolBox.Create(mParent: TWinControl; G1, G2, G3, G4: TBitmap);
begin
  inherited Create(mParent);
  Parent := mParent;
  Visible := True;
  Caption := 'Fenêtre Outils';
  Left := 350;
  Top := 260;
  Height := 300;
  Width := 550;
  BorderStyle := bsSingle;
  Name := 'LeTOOL';
  self.OnAllKeyDown := FormOnKeyDown;

  LaGrille := CreeGrille;
  LeLabel1 := CreeLabel('Label1');
  LeLabel2 := CreeLabel('Label2');
  LeLabel3 := CreeLabel('Label3');
  LeLabel4 := CreeLabel('Label4');
  LeLabel5 := CreeLabel('Label5');
  LeLabel6 := CreeLabel('Label6');
  LeLabel7 := CreeLabel('Label7');
  LEdit1 := CreeEdit('Edit1');
  LEdit2 := CreeEdit('Edit54');
  LePanel := CreePanel;
  BValide := CreeBouton(LePanel, G1, 'Bt1');
  BAnnule := CreeBouton(LePanel, G2, 'Bt2');
  BAjouteLg := CreeBouton(LePanel, G3, 'Bt3', );
  BSupprimeLg := CreeBouton(LePanel, G4, 'Bt4');
  LaCoche := CreeLaCoche;

  LaGrille.OnRowEnter := LaGrilleOnRowEnter;
  LaGrille.OnKeyDown := LaGrilleOnKeyDown;
  LaGrille.OnMouseWheelDown := LaGrilleOnMouseWheelDown;
  LaGrille.PostDrawCell := LaGrillePostDrawCell;
  BAjouteLg.OnClick := LaGrilleAjouteLigneOnClick;
  BSupprimeLg.OnClick := LaGrilleSupprimeLigneOnClick;
end;

procedure TDRFToolBox.FormOnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_ESCAPE:
      begin
        Key := 0;
        BAnnule.onClick(Sender);
      end;
    VK_F5:
      begin
        Key := 0;
        BValide.onClick(Sender);
      end;
    VK_F10:
      begin
        Key := 0;
        BValide.OnClick(Sender);
      end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 16/01/2007
Modifié le ... :   /  /
Description .. : Creation d'un bouton de type Toolbar
Mots clefs ... :
*****************************************************************}

function TDRFToolBox.CreeBouton(LePanel: THPanel; Gly: TBitmap; Lenom: string): TToolbarButton97;
var
  LeBouton: TToolbarButton97;

begin
  Result := nil;
  LeBouton := TToolbarButton97.Create(self);
  if not Assigned(LeBouton) then
    exit;
  LeBouton.Parent := LePanel;
  LeBouton.Name := Lenom;
  LeBouton.Visible := False;
  LeBouton.Glyph := Gly;
  LeBouton.Height := 27;
  LeBouton.Width := 28;
  LeBouton.Flat := False;
  LeBouton.Color := clBtnFace;
  LeBouton.Opaque := True;
  Result := LeBouton;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 16/01/2007
Modifié le ... :   /  /
Description .. : création d'un zone de saisie
Mots clefs ... :
*****************************************************************}

function TDRFToolBox.CreeEdit(Lenom: string): THEdit;
var
  LeChamp: THEdit;

begin
  Result := nil;
  LeChamp := THEdit.Create(self);
  if not Assigned(LeChamp) then
    exit;
  LeChamp.Parent := Self;
  LeChamp.Name := Lenom;
  Lechamp.Visible := False;
  LeChamp.Height := 21;
  LeChamp.Width := 91;
  LeChamp.OpeType := otReel;
  { FQ 21150 BVE 20.07.07 }
  LeChamp.Text := '';
  { END FQ 21150 }
  Result := LeChamp;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 16/01/2007
Modifié le ... :   /  /
Description .. : création de la grille de saisie
Mots clefs ... :
*****************************************************************}

function TDRFToolBox.CreeGrille: THGrid;
var
  Grill: THGrid;

begin
  Result := nil;
  Grill := THGrid.Create(self);
  if not Assigned(Grill) then
    exit;
  Grill.Name := 'Lagrille';
  Grill.Parent := Self;
  Grill.Visible := False;
  Grill.Align := alNone;
  Grill.FixedCols := 1;
  Grill.FixedRows := 0;
  Grill.DefaultRowHeight := 18;
  Grill.CalcInCell := True; // pour permettre le calcul dans les cellules
  Grill.Options := Grill.Options - [goRowSelect] + [goEditing] + [goTabs];
  Result := Grill;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 16/01/2007
Modifié le ... :   /  /
Description .. : Création de libellé fixe
Mots clefs ... :
*****************************************************************}

function TDRFToolBox.CreeLabel(Lenom: string): ThLabel;
var
  LeChamp: THLabel;

begin
  Result := nil;
  LeChamp := THLabel.Create(self);
  if not Assigned(LeChamp) then
    exit;
  LeChamp.Parent := Self;
  LeChamp.Name := Lenom;
  LeChamp.Visible := False;
  LeChamp.AutoSize := False;
  LeChamp.Font.Style := [fsBold];
  LeChamp.Alignment := taCenter;
  LeChamp.WordWrap := True;
  Result := LeChamp;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 16/01/2007
Modifié le ... :   /  /
Description .. : Création du panel pour stocker les boutons
Mots clefs ... :
*****************************************************************}

function TDRFToolBox.CreePanel: THPanel;
var
  LePanell: THPanel;

begin
  Result := nil;
  LePanell := THPanel.Create(self);
  if not Assigned(LePanell) then
    exit;
  LePanell.parent := Self;
  LePanell.Name := 'Lepanel';
  LePanell.Caption := '';
  LePanell.Visible := False;
  LePanell.Align := alBottom;
  LePanell.Height := 35;
  LePanell.ParentColor := False;
  LePanell.Color := clBtnFace;
  Result := LePanell;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 30/01/2007
Modifié le ... :   /  /
Description .. : Création d'une case à cocher
Mots clefs ... :
*****************************************************************}

function TDRFToolBox.CreeLaCoche: THCheckbox;
var
  LeChamp: THCheckbox;

begin
  Result := nil;
  LeChamp := THCheckbox.Create(self);
  if not Assigned(LeChamp) then
    exit;
  LeChamp.Parent := Self;
  LeChamp.Name := 'Lacoche';
  LeChamp.Visible := False;
  LeChamp.Checked := False;
  LeChamp.Caption := 'Case à Cocher';
  LeChamp.Width := 300;
  Result := LeChamp;


end;


destructor TDRFToolBox.Destroy;
begin
{  FreeAndNil(Lagrille);
  FreeAndNil(LeLabel1);
  FreeAndNil(LeLabel2);
  FreeAndNil(LeLabel3);
  FreeAndNil(LeLabel4);
  FreeAndNil(LeLabel5);
  FreeAndNil(LeLabel6);
  FreeAndNil(LeLabel7);
  FreeAndNil(LEdit1);
  FreeAndNil(LEdit2);
  FreeAndNil(BValide);
  FreeAndNil(BAnnule);
  FreeAndNil(BAjouteLg);
  FreeAndNil(BSupprimeLg);
  FreeAndNil(LePanel);
  FreeAndNil(LaCoche);  }

  inherited Destroy;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : Active le label pour l'outil
Mots clefs ... :
*****************************************************************}

procedure TDRFToolBox.ActivLabel(LeLaBel: THLabel; LeCaption: string; LeTop, LeLeft, LeHeight, LeWidth: integer);
begin
  with LeLaBel do
  begin
    Visible := True;
    Caption := LeCaption;
    Top := LeTop;
    Left := LeLeft;
    Height := LeHeight;
    Width := LeWidth;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : active la grille pour l'outil
Mots clefs ... :
*****************************************************************}

procedure TDRFToolBox.ActivGrille(LaGrille: THGrid; NbCol, NbRow, LeTop, LeLeft, LeHeight, LeWidth: integer);
var
  i: integer;

begin
  with LaGrille do
  begin
    Visible := True;
    Top := LeTop;
    Left := LeLeft;
    Height := LeHeight;
    Width := LeWidth;
    ColCount := NbCol;
    RowCount := NbRow;

    ColEditables[0] := False;
    ColWidths[0] := 10;
    ColWidths[1] := 100;

    for i := 2 to NbCol - 1 do
    begin
      ColWidths[i] := 100;
      ColTypes[i] := 'K';
      ColFormats[i] := 'R';
      ColAligns[i] := taRightJustify;
    end;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : Active les edits dans l'outil
Mots clefs ... :
*****************************************************************}

procedure TDRFToolBox.ActivEdit(Ledit: THEdit; LeHint: string; LeTop, LeLeft: integer);
begin
  with Ledit do
  begin
    Visible := True;
    Top := LeTop;
    Left := LeLeft;
    Hint := LeHint;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : Active les boutons dans l'outil
Mots clefs ... :
*****************************************************************}

procedure TDRFToolBox.ActivBouton(LeButton: TToolbarButton97; LeHint: string; LeTop, LeLeft: integer);
begin
  with LeButton do
  begin
    Visible := True;
    Top := LeTop;
    Left := LeLeft;
    Hint := LeHint;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 30/01/2007
Modifié le ... :   /  /
Description .. : Active la case à cocher dans l'outil
Mots clefs ... :
*****************************************************************}

procedure TDRFToolBox.ActivCoche(LaCoche: THCheckbox; LeTitre: string; LeTop, LeLeft: integer);
begin
  with LaCoche do
  begin
    Visible := True;
    Top := LeTop;
    Left := LeLeft;
    Hint := LeTitre;
    Caption := LeTitre;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 12/01/2007
Modifié le ... : 12/01/2007
Description .. : Ajout d'une ligne dans la grille de la fenêtre outil.
Suite ........ : On décale la dernière ligne de la grille qui contient les
Suite ........ : totaux.
Suite ........ : Insertion à la ligne Index
Mots clefs ... :
*****************************************************************}

procedure TDRFToolBox.LaGrilleAjouteLigneOnClick(Sender: Tobject);
begin
  LaGrilleAjouteLigne(LaGrille.Row + 1);
end;

procedure TDRFToolBox.LaGrilleAjouteLigne(Index: integer);
var
  i: integer;
  j: integer;
  NbLg: integer;

begin
  LaGrille.RowCount := LaGrille.RowCount + 1;
  NbLg := LaGrille.RowCount;

  for i := NbLg - 1 downto Index do
  begin
    for j := 1 to Lagrille.ColCount - 1 do //la première colonne est TOUJOURS fixe
    begin
      LaGrille.Cells[j, i + 1] := LaGrille.Cells[j, i];
      LaGrille.Cells[j, i] := '';
    end;
  end;
  LaGrille.Row := Index;
  LaGrille.Col := 1;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 12/01/2007
Modifié le ... :   /  /
Description .. : Suppression d'une ligne de la grille
Mots clefs ... :
*****************************************************************}

procedure TDRFToolBox.LaGrilleSupprimeLigneOnClick(Sender: Tobject);
begin
  LaGrilleSupprimeLigne(LaGrille.Row);
end;

procedure TDRFToolBox.LaGrilleSupprimeLigne(Index: integer);
var
  i: integer;
  j: integer;


begin
  if (LaGrille.RowCount) <= (PiedGrille + 1) then
  begin
    for i := 1 to LaGrille.ColCount - 1 do
    begin
      LaGrille.Cells[i, 0] := '';
    end;
    exit;
  end;

  for i := Index + 1 to LaGrille.RowCount - 1 do
  begin
    for j := 1 to LaGrille.ColCount - 1 do
    begin
      LaGrille.Cells[j, i - 1] := LaGrille.Cells[j, i];
    end;
  end;
  LaGrille.RowCount := LaGrille.RowCount - 1;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 12/01/2007
Modifié le ... :   /  /
Description .. : Gère l'entrée sur les lignes de la grille. la dernière ligne est
Suite ........ : réservé aux totaux. On va donc la protéger
Mots clefs ... :
*****************************************************************}

procedure TDRFToolBox.LaGrilleOnRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  if LaGrille.Row >= LaGrille.RowCount - PiedGrille then
    LaGrille.Row := LaGrille.RowCount - (PiedGrille + 1);

end;

procedure TDRFToolBox.LaGrilleOnMouseWheelDown(Sender: Tobject; Shift: TShiftState; MousePos: TPoint; var Handled: boolean);
begin
  if LaGrille.Row > LaGrille.RowCount - (PiedGrille + 1) then
    LaGrille.Row := LaGrille.RowCount - (PiedGrille + 1);

end;


procedure TDRFToolBox.LaGrillePostDrawCell(Acol, ARow: integer; Canvas: TCanvas; Astate: TGridDrawState);
var
  R: TRect;
  XPos: integer;
  LaCol: integer;

begin
  if (ARow >= LaGrille.RowCount - PiedGrille) and (TheTool <> 'JETONPRESENCE') then
  begin
    R := LaGrille.CellRect(ACol, ARow);


    // Définition de la achure de la cellule
    Canvas.Brush.Color := LaGrille.FixedColor;
    Canvas.Brush.Style := bsSolid;
    Canvas.Pen.Color := LaGrille.FixedColor;
    Canvas.Pen.Mode := pmCopy;
    Canvas.Pen.Style := psClear;
    Canvas.Pen.Width := 1;
    // on achure la cellule
    Canvas.Rectangle(R.Left, R.top, R.Right + 1, R.Bottom + 1);

    // On met en gras la ligne des totaux
    Canvas.Font.Style := Canvas.Font.Style + [fsBold];
    XPos := R.Right - Canvas.TextWidth(LaGrille.Cells[ACol, ARow]) - 3;
    Canvas.TextRect(R, XPos, R.Top + 3, LaGrille.Cells[Acol, Arow]);


  end;

  if TheTool = 'SAISIEGRILLE' then
  begin
    LaCol := 0;
    if TheTag = 4 then
      LaCol := 3
    else if TheTag = 5 then
      LaCol := 2;

    if (ACol = LaCol) and (LaCol <> 0) then
    begin
      R := LaGrille.CellRect(ACol, ARow);
      // Définition de la achure de la cellule
      Canvas.Brush.Color := LaGrille.FixedColor;
      Canvas.Brush.Style := bsSolid;
      Canvas.Pen.Color := LaGrille.FixedColor;
      Canvas.Pen.Mode := pmCopy;
      Canvas.Pen.Style := psClear;
      Canvas.Pen.Width := 1;
      // on achure la cellule
      Canvas.Rectangle(R.Left, R.top, R.Right + 1, R.Bottom + 1);
    end;
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 12/01/2007
Modifié le ... :   /  /
Description .. : Gestion des touches dans la grille de la fenêtre outils
Mots clefs ... :
*****************************************************************}

procedure TDRFToolBox.LaGrilleOnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
{$IFDEF AMORTISSEMENT}
var
  LaVal: extended;
  Idx: integer;
  i: integer;
  TobImo: Tob;
  TobV: Tob;
{$ENDIF}

begin
  case Key of
    VK_INSERT:
      begin
        if (Shift = []) then
          LaGrilleAjouteLigne(LaGrille.Row + 1);
        if (Shift = [ssShift]) then
          LaGrilleAjouteLigne(LaGrille.Row);
      end;

    VK_TAB:
      begin
        if (Shift = []) then
        begin
          if TheTool <> 'SAISIEGRILLE' then
          begin
            if LaGrille.Col = LaGrille.ColCount - 1 then
            begin
              Key := 0;
              if LaGrille.Row = LaGrille.RowCount - 1 - PiedGrille then
                LaGrilleAjouteLigne(LaGrille.Row + 1)
              else
              begin
                { FQ 20531 BVE 27.07.07 }
                LaGrille.Row := LaGrille.Row + 1;
                LaGrille.Col := 1;
                { END FQ 20531 }
              end;
            end;
          end
          else
          begin
            if (TheTag = 4) and (LaGrille.Col = 2) then
            begin
              Key := 0;
              if LaGrille.Row = LaGrille.RowCount - 1 - PiedGrille then
                LaGrilleAjouteLigne(LaGrille.Row + 1)
              else
              begin
                { FQ 20531 BVE 11.06.07 }
                LaGrille.Row := LaGrille.Row + 1;
                LaGrille.Col := 1;
                { END FQ 20531 }
              end;
            end;
            if (TheTag = 5) and (LaGrille.Col = 3) then
            begin
              Key := 0;
              if LaGrille.Row = LaGrille.RowCount - 1 - PiedGrille then
                LaGrilleAjouteLigne(LaGrille.Row + 1)
              else
              begin
                { FQ 20531 BVE 11.06.07 }
                LaGrille.Row := LaGrille.Row + 1;
                LaGrille.Col := 1;
                { END FQ 20531 }
              end;
            end;
          end;
        end;
      end;

    VK_DELETE:
      begin
        if (Shift = [ssCtrl]) then
        begin
          if LaGrille.Row < (LaGrille.RowCount - PiedGrille) then
            LaGrilleSupprimeLigne(LaGrille.Row);
        end;
      end;
    (*VK_F5 :
      begin
         { FQ 20531 BVE 27.07.07 }
         Self.BValide.OnClick(Sender);
         { END FQ 20531 }
      end;   *)
    VK_F6:
      begin
        if TheTool = '2058A_XA_MOA' then // régime société mère
        begin
          if LaGrille.Col = 2 then // colonne 'montant des produits encaissés'
          begin
            LaGrille.Cells[2, LaGrille.Row] := STRFMONTANT(RegimeSocieteMere, 15, V_PGI.OkDecV, '', True);
          end;
        end;

        if TheTool = '2058A_VA_MOA' then //imposition PVLT
        begin
          if (LaGrille.Col = 2) and (ValeurPVLT <> 0) then
            LaGrille.Cells[2, LaGrille.Row] := STRFMONTANT(VALEUR(LaGrille.Cells[2, LaGrille.RowCount - 1]), 15, V_PGI.OkDecV, '', True);
        end;

        if TheTool = 'CESSIONIMO' then // Cession des immobilisations
        begin
{$IFDEF AMORTISSEMENT}
          Idx := 0;
          TobImo := Tob.Create('les immos', nil, -1);
          ChargeImo(TOBImo);
          try
            for i := 0 to TobImo.Detail.Count - 1 do
            begin
              LaGrille.RowCount := Idx + 1;
              if Idx = 0 then
                LaGrille.Cells[1, 0] := 'Total';

              LaGrilleAjouteLigne(Idx);
              TobV := TobImo.Detail[i];
              LaGrille.Cells[1, Idx] := TobV.GetValue('LIBELLE');
              LaGrille.Cells[2, Idx] := TobV.GetValue('MONTANTCES');
              LaVal := TobV.GetValue('PLT') - TobV.GetValue('MLT');
              LaGrille.Cells[3, Idx] := STRFMONTANT(LaVal, 15, V_PGI.OkDecV, '', True);
              LaVal := TobV.GetValue('PCT') - TobV.GetValue('MCT');
              LaGrille.Cells[4, Idx] := STRFMONTANT(LaVal, 15, V_PGI.OkDecV, '', True);

              CessionImmo_CalculLigne(Idx);

              Inc(Idx);
            end;
          finally
            TobImo.Free;
          end;
{$ENDIF}
        end;
      end;
(*
    VK_ESCAPE :
      begin
        if (Shift = []) then
          close;
      end;
      *)
  end;

end;




{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 16/01/2007
Modifié le ... :   /  /
Description .. : Formatte la form selon le type d'outils ou de saisie en grille
Mots clefs ... :
*****************************************************************}

procedure TDRFToolBox.TypeToolBox(LeTool: string; LaTob: Tob; LeTitre: string; Borne1, Borne2: extended; LeTag: integer);
var
  TobV: Tob;
  Idx: integer;
  i: integer;


begin
  TheTool := LeTool;

  // Suivi des déficits   2058A_XL_MOA
  if LeTool = '2058A_XL_MOA' then
  begin
    Caption := 'Suivi des déficits';
    PiedGrille := 1;
    Height := 410;

    ActivLabel(LeLabel1, 'Explication', 6, 15);
    ActivLabel(LeLabel2, 'Déficit restant à reporter au titre de l''exercice précédent', 6, 115);
    ActivLabel(LeLabel3, 'Imputés sur les PVLT', 6, 215);
    ActivLabel(LeLabel4, 'Imputés sur les résultats de l''exercice', 6, 315);
    ActivLabel(LeLaBel5, 'Déficits reportables', 6, 415);

    ActivGrille(LaGrille, 6, 1, 60, 0, 200, Width - 10);
    LaGrille.ColEditables[5] := False;
    LaGrille.Cells[1, 0] := 'Total';

    ActivLabel(LeLabel6, 'Déficit de l''exercice', 270, 150, 13, 250);
    LeLabel6.Alignment := taRightJustify;

    ActivLabel(LeLabel7, 'Total des déficits restant à reporter', 300, 150, 13, 250);
    LeLabel7.Alignment := taRightJustify;

    ActivEdit(LEdit1, 'Reprise du montant du déficit éventuel de l''exercice 2058A_XO_MOA', 270, 415);
    ActivEdit(LEdit2, 'Total des déficits restant à reporter', 300, 415);

    LePanel.Visible := True;
    ActivBouton(BValide, 'Valider la saisie', 2, 460);
    ActivBouton(BAnnule, 'Annuler la saisie', 2, 500);
    ActivBouton(BAjouteLg, 'Ajouter une ligne', 2, 250);
    ActivBouton(BSupprimeLg, 'Supprimer une ligne', 2, 290);


    // recherche des éléments
    if Assigned(LaTob) then
    begin
      TobV := LaTob.FindFirst(['NATURE'], ['2058A_XO_MOA'], False);
      if TobV <> nil then
        LEdit1.Text := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True)
      else
        LEdit1.Text := STRFMONTANT(0, 15, V_PGI.OkDecV, '', True);

      Idx := 0;
      TobV := LaTob.FindFirst(['NATURE'], ['2058A_XL_MOA'], False);
      while TobV <> nil do
      begin
        if TobV.GetValue('INDICE') <> '0' then // on ne prend pas le premier, c'est le total général
        begin
          LaGrille.RowCount := Idx + 1;
          LaGrilleAjouteLigne(Idx);
          LaGrille.Cells[1, Idx] := TobV.GetValue('COMMENTAIRE');
          LaGrille.Cells[2, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[3, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT2')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[4, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT3')), 15, V_PGI.OkDecV, '', True);
          B2058A_XL_MOA_CalculLigne(Idx);
          Inc(Idx);
        end;
        TobV := LaTob.FindNext(['NATURE'], ['2058A_XL_MOA'], False);
      end;
    end;

    if LaGrille.RowCount = PiedGrille then
      LaGrilleAjouteLigne(0);

    //les propriétés
    LaGrille.OnCellExit := B2058A_XL_MOA_OnCellexit;

  end
  //==============================================================================================
  //ToolBox pour rémunération du conjoint de l'exploitant
  else if LeTool = '2058A_WC_MOA' then
  begin
    Caption := 'Rémunération du conjoint de l''exploitant';
    PiedGrille := 1;
    Height := 330;
    RemConjointMin := Borne1;
    RemConjointMax := Borne2;

    ActivLabel(LeLabel1, 'Explication', 6, 15, 50, 200);
    ActivLabel(LeLabel2, 'Rémunération du conjoint', 6, 215);
    ActivLabel(LeLabel3, 'Part déductible', 6, 315);
    ActivLabel(LeLabel4, 'A réintégrer', 6, 415);

    ActivGrille(LaGrille, 5, 1, 60, 0, 200, Width - 10);
    LaGrille.ColWidths[1] := 200;
    LaGrille.ColEditables[4] := False;
    LaGrille.Cells[1, 0] := 'Total';

    LePanel.Visible := True;
    ActivBouton(BValide, 'Valider la saisie', 2, 460);
    ActivBouton(BAnnule, 'Annuler la saisie', 2, 500);
    ActivBouton(BAjouteLg, 'Ajouter une ligne', 2, 250);
    ActivBouton(BSupprimeLg, 'Supprimer une ligne', 2, 290);

      // recherche des éléments
    Idx := 0;
    if Assigned(LaTob) then
    begin
      TobV := LaTob.FindFirst(['NATURE'], ['2058A_WC_MOA'], False);
      while TobV <> nil do
      begin
        if TobV.GetValue('INDICE') <> '0' then // on ne prend pas le premier, c'est le total général
        begin
          LaGrille.RowCount := Idx + 1;
          LaGrilleAjouteLigne(Idx);
          LaGrille.Cells[1, Idx] := TobV.GetValue('COMMENTAIRE');
          LaGrille.Cells[2, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[3, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT2')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[4, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT3')), 15, V_PGI.OkDecV, '', True);
          B2058A_WC_MOA_CalculLigne(Idx);
          Inc(Idx);
        end;
        TobV := LaTob.FindNext(['NATURE'], ['2058A_WC_MOA'], False);
      end;
    end;

    if LaGrille.RowCount = PiedGrille then
      LaGrilleAjouteLigne(0);

    //les propriétés
    LaGrille.OnCellExit := B2058A_WC_MOA_OnCellexit;

  end
  //==============================================================================================
  //ToolBox pour Régime des sociétés mères et des filiales
  else if LeTool = '2058A_XA_MOA' then
  begin
    Caption := 'Régime des sociétés mères et des filiales';
    PiedGrille := 1;
    Height := 330;

    ActivLabel(LeLabel1, 'Explication', 6, 15, 50, 200);
    ActivLabel(LeLabel2, 'Montant des produits encaissés', 6, 215);
    ActivLabel(LeLabel3, 'Quote-parts pour frais', 6, 315);
    ActivLabel(LeLabel4, 'A déduire', 6, 415);

    ActivGrille(LaGrille, 5, 1, 60, 0, 200, Width - 10);
    LaGrille.ColWidths[1] := 200;
    LaGrille.ColEditables[3] := False;
    LaGrille.ColEditables[4] := False;
    LaGrille.Cells[1, 0] := 'Total';

    LePanel.Visible := True;
    ActivBouton(BValide, 'Valider la saisie', 2, 460);
    ActivBouton(BAnnule, 'Annuler la saisie', 2, 500);
    ActivBouton(BAjouteLg, 'Ajouter une ligne', 2, 250);
    ActivBouton(BSupprimeLg, 'Supprimer une ligne', 2, 290);

    // recherche des éléments
    Idx := 0;
    if Assigned(LaTob) then
    begin
      TobV := LaTob.FindFirst(['NATURE'], ['2058A_XA_MOA'], False);
      while TobV <> nil do
      begin
        if TobV.GetValue('INDICE') <> '0' then // on ne prend pas le premier, c'est le total général
        begin
          LaGrille.RowCount := Idx + 1;
          LaGrilleAjouteLigne(Idx);
          LaGrille.Cells[1, Idx] := TobV.GetValue('COMMENTAIRE');
          LaGrille.Cells[2, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
          B2058A_XA_MOA_CalculLigne(Idx);
          Inc(Idx);
        end;
        TobV := LaTob.FindNext(['NATURE'], ['2058A_XA_MOA'], False);
      end;
    end;

    if LaGrille.RowCount = PiedGrille then
      LaGrilleAjouteLigne(0);

    //les propriétés
    LaGrille.OnCellExit := B2058A_XA_MOA_OnCellexit;

  end
  //==============================================================================================
  //ToolBox pour provisions pour risques et charges
  else if LeTool = '2058B_CA_MOA' then
  begin
    Caption := 'Provisions pour risques et charges';
    PiedGrille := 1;
    Height := 330;
    Width := 450;

    ActivLabel(LeLabel1, 'Explication', 6, 15, 50, 200);
    ActivLabel(LeLabel2, 'Dotations', 6, 215);
    ActivLabel(LeLabel3, 'Reprises', 6, 315);

    ActivGrille(LaGrille, 4, 1, 60, 0, 200, Width - 10);
    LaGrille.ColWidths[1] := 200;
    LaGrille.Cells[1, 0] := 'Total';

    LePanel.Visible := True;
    ActivBouton(BValide, 'Valider la saisie', 2, 360);
    ActivBouton(BAnnule, 'Annuler la saisie', 2, 400);
    ActivBouton(BAjouteLg, 'Ajouter une ligne', 2, 150);
    ActivBouton(BSupprimeLg, 'Supprimer une ligne', 2, 190);

    // recherche des éléments
    Idx := 0;
    if Assigned(LaTob) then
    begin
      for i := 0 to LaTob.Detail.Count - 1 do
      begin
        TobV := LaTob.Detail[i];     
        { FQ 21621 BVE 10.10.07 }
        if (TobV.GetValue('NATURE') = '2058B_CL_MOA') and (TobV.GetValue('INDICE') <> 0) then
        begin
          LaGrille.RowCount := Idx + 1;
          LaGrilleAjouteLigne(Idx);
          LaGrille.Cells[1, Idx] := TobV.GetValue('COMMENTAIRE');
          LaGrille.Cells[2, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[3, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT2')), 15, V_PGI.OkDecV, '', True);
          B2058B_PCP_CalculLigne(Idx);
          Inc(Idx);
        end;
        { END FQ 21621 }
      end;
    end;

    if LaGrille.RowCount = PiedGrille then
      LaGrilleAjouteLigne(0);

    //les propriétés
    LaGrille.OnCellExit := B2058B_PCP_OnCellexit;

  end
  //==============================================================================================
  //ToolBox pour provisions pour dépréciation
  else if LeTool = '2058B_CD_MOA' then
  begin
    Caption := 'Provisions pour dépréciation';
    PiedGrille := 1;
    Height := 330;
    Width := 450;

    ActivLabel(LeLabel1, 'Explication', 6, 15, 50, 200);
    ActivLabel(LeLabel2, 'Dotations', 6, 215);
    ActivLabel(LeLabel3, 'Reprises', 6, 315);

    ActivGrille(LaGrille, 4, 1, 60, 0, 200, Width - 10);
    LaGrille.ColWidths[1] := 200;
    LaGrille.Cells[1, 0] := 'Total';

    LePanel.Visible := True;
    ActivBouton(BValide, 'Valider la saisie', 2, 360);
    ActivBouton(BAnnule, 'Annuler la saisie', 2, 400);
    ActivBouton(BAjouteLg, 'Ajouter une ligne', 2, 150);
    ActivBouton(BSupprimeLg, 'Supprimer une ligne', 2, 190);


    // recherche des éléments
    Idx := 0;
    if Assigned(LaTob) then
    begin
      for i := 0 to LaTob.Detail.Count - 1 do
      begin
        TobV := LaTob.Detail[i];        
        { FQ 21621 BVE 10.10.07 }
        if (TobV.GetValue('NATURE') = '2058B_CM_MOA') and (TobV.GetValue('INDICE') <> 0) then
        begin
          LaGrille.RowCount := Idx + 1;
          LaGrilleAjouteLigne(Idx);
          LaGrille.Cells[1, Idx] := TobV.GetValue('COMMENTAIRE');
          LaGrille.Cells[2, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[3, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT2')), 15, V_PGI.OkDecV, '', True);
          B2058B_PCP_CalculLigne(Idx);
          Inc(Idx);
        end;
        { END FQ 21621 }
      end;
    end;

    if LaGrille.RowCount = PiedGrille then
      LaGrilleAjouteLigne(0);

    //les propriétés
    LaGrille.OnCellExit := B2058B_PCP_OnCellexit;

  end
  else if LeTool = '2058B_CG_MOA' then
  begin
    Caption := 'Charges à payer';
    PiedGrille := 1;
    Height := 330;
    Width := 450;

    ActivLabel(LeLabel1, 'Explication', 6, 15, 50, 200);
    ActivLabel(LeLabel2, 'Dotations', 6, 215);
    ActivLabel(LeLabel3, 'Reprises', 6, 315);

    ActivGrille(LaGrille, 4, 1, 60, 0, 200, Width - 10);
    LaGrille.ColWidths[1] := 200;
    LaGrille.Cells[1, 0] := 'Total';

    LePanel.Visible := True;
    ActivBouton(BValide, 'Valider la saisie', 2, 360);
    ActivBouton(BAnnule, 'Annuler la saisie', 2, 400);
    ActivBouton(BAjouteLg, 'Ajouter une ligne', 2, 150);
    ActivBouton(BSupprimeLg, 'Supprimer une ligne', 2, 190);

    // recherche des éléments
    Idx := 0;
    if Assigned(LaTob) then
    begin
      for i := 0 to LaTob.Detail.Count - 1 do
      begin
        TobV := LaTob.Detail[i];       
        { FQ 21621 BVE 10.10.07 }
        if (TobV.GetValue('NATURE') = '2058B_CN_MOA') and (TobV.GetValue('INDICE') <> 0) then
        begin
          LaGrille.RowCount := Idx + 1;
          LaGrilleAjouteLigne(Idx);
          LaGrille.Cells[1, Idx] := TobV.GetValue('COMMENTAIRE');
          LaGrille.Cells[2, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[3, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT2')), 15, V_PGI.OkDecV, '', True);
          B2058B_PCP_CalculLigne(Idx);
          Inc(Idx);
        end;
        { END FQ 21621 }
      end;
    end;

    if LaGrille.RowCount = PiedGrille then
      LaGrilleAjouteLigne(0);

    //les propriétés
    LaGrille.OnRowEnter := LaGrilleOnRowEnter;
    LaGrille.OnCellExit := B2058B_PCP_OnCellexit;
    LaGrille.OnKeyDown := LaGrilleOnKeyDown;
    BAjouteLg.OnClick := LaGrilleAjouteLigneOnClick;
    BSupprimeLg.OnClick := LaGrilleSupprimeLigneOnClick;

  end
  else if LeTool = '2058B_ZV_MOA' then
  begin
    Caption := 'Indemnités congès à payer';
    PiedGrille := 1;
    Height := 330;
    Width := 450;

    ActivLabel(LeLabel1, 'Explication', 6, 15, 50, 200);
    ActivLabel(LeLabel2, 'Dotations', 6, 215);
    ActivLabel(LeLabel3, 'Reprises', 6, 315);

    ActivGrille(LaGrille, 4, 1, 60, 0, 200, Width - 10);
    LaGrille.ColWidths[1] := 200;
    LaGrille.Cells[1, 0] := 'Total';

    LePanel.Visible := True;
    ActivBouton(BValide, 'Valider la saisie', 2, 360);
    ActivBouton(BAnnule, 'Annuler la saisie', 2, 400);
    ActivBouton(BAjouteLg, 'Ajouter une ligne', 2, 150);
    ActivBouton(BSupprimeLg, 'Supprimer une ligne', 2, 190);

    // recherche des éléments
    Idx := 0;
    if Assigned(LaTob) then
    begin
      for i := 0 to LaTob.Detail.Count - 1 do
      begin
        TobV := LaTob.Detail[i];         
        { FQ 21621 BVE 10.10.07 }
        if (TobV.GetValue('NATURE') = '2058B_ZV_MOA') and (TobV.GetValue('INDICE') <> 0) then
        begin
          LaGrille.RowCount := Idx + 1;
          LaGrilleAjouteLigne(Idx);
          LaGrille.Cells[1, Idx] := TobV.GetValue('COMMENTAIRE');
          LaGrille.Cells[2, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[3, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT2')), 15, V_PGI.OkDecV, '', True);
          B2058B_PCP_CalculLigne(Idx);
          Inc(Idx);
        end;
        { END FQ 21621 }
      end;
    end;

    if LaGrille.RowCount = PiedGrille then
      LaGrilleAjouteLigne(0);

    //les propriétés
    LaGrille.OnCellExit := B2058B_PCP_OnCellexit;

  end
  //==============================================================================================
  //ToolBox pour suivi des PCVT différés
  else if LeTool = '2058A_WN_MOA' then
  begin
    Caption := 'Suivi des PVCT différées';
    PiedGrille := 1;
    Height := 330;

    ActivLabel(LeLabel1, 'Explication', 6, 15, 50);
    ActivLabel(LeLabel2, 'Montant antérieur', 6, 115);
    ActivLabel(LeLabel3, 'Montant rapporté au résultat de l''exercice', 6, 215);
    ActivLabel(LeLabel4, 'Fraction différée', 6, 315);
    ActivLabel(LeLabel5, 'Montant restant à réintégrer', 6, 415);

    ActivGrille(LaGrille, 6, 1, 60, 0, 200, Width - 10);
    LaGrille.ColEditables[5] := False;
    LaGrille.Cells[1, 0] := 'Total';

    LePanel.Visible := True;
    ActivBouton(BValide, 'Valider la saisie', 2, 460);
    ActivBouton(BAnnule, 'Annuler la saisie', 2, 500);
    ActivBouton(BAjouteLg, 'Ajouter une ligne', 2, 250);
    ActivBouton(BSupprimeLg, 'Supprimer une ligne', 2, 290);

    // recherche des éléments
    Idx := 0;
    if Assigned(LaTob) then
    begin
      TobV := LaTob.FindFirst(['NATURE'], ['2058A_WN_MOA'], False);
      while TobV <> nil do
      begin
        if TobV.GetValue('INDICE') <> '0' then // on ne prend pas le premier, c'est le total général
        begin
          LaGrille.RowCount := Idx + 1;
          LaGrilleAjouteLigne(Idx);
          LaGrille.Cells[1, Idx] := TobV.GetValue('COMMENTAIRE');
          LaGrille.Cells[2, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[3, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT2')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[4, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT3')), 15, V_PGI.OkDecV, '', True);
          B2058A_WN_MOA_CalculLigne(Idx);
          Inc(Idx);
        end;
        TobV := LaTob.FindNext(['NATURE'], ['2058A_WN_MOA'], False);
      end;
    end;

    if LaGrille.RowCount = PiedGrille then
      LaGrilleAjouteLigne(0);

    //les propriétés
    LaGrille.OnCellExit := B2058A_WN_MOA_OnCellexit;

  end
  else if LeTool = '2058A_WM_MOA' then
  begin
    Caption := 'Suivi des MVLT';
    PiedGrille := 1;
    Height := 330;
    Width := 650;

    ActivLabel(LeLabel1, 'Explication', 6, 15, 50);
    ActivLabel(LeLabel2, 'Montant antérieur', 6, 115);
    ActivLabel(LeLabel3, 'Imputation sur les plus-values à long terme', 6, 215);
    ActivLabel(LeLabel4, 'Imputations sur le résultat de l''exercice', 6, 315);
    ActivLabel(LeLabel5, 'Moins-values de l''exercice', 6, 415);
    ActivLabel(LeLabel6, 'Solde des moins-values à reporter', 6, 515);

    ActivGrille(LaGrille, 7, 1, 60, 0, 200, Width - 10);
    LaGrille.ColEditables[6] := False;
    LaGrille.Cells[1, 0] := 'Total';

    LePanel.Visible := True;
    ActivBouton(BValide, 'Valider la saisie', 2, 560);
    ActivBouton(BAnnule, 'Annuler la saisie', 2, 600);
    ActivBouton(BAjouteLg, 'Ajouter une ligne', 2, 250);
    ActivBouton(BSupprimeLg, 'Supprimer une ligne', 2, 290);


    // recherche des éléments
    Idx := 0;
    if Assigned(LaTob) then
    begin
      TobV := LaTob.FindFirst(['NATURE'], ['2058A_WM_MOA'], False);
      while TobV <> nil do
      begin
        if TobV.GetValue('INDICE') <> '0' then // on ne prend pas le premier, c'est le total général
        begin
          LaGrille.RowCount := Idx + 1;
          LaGrilleAjouteLigne(Idx);
          LaGrille.Cells[1, Idx] := TobV.GetValue('COMMENTAIRE');
          LaGrille.Cells[2, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[3, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT2')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[4, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT3')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[5, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT4')), 15, V_PGI.OkDecV, '', True);
          B2058A_WM_MOA_CalculLigne(Idx);
          Inc(Idx);
        end;
        TobV := LaTob.FindNext(['NATURE'], ['2058A_WM_MOA'], False);
      end;
    end;

    if LaGrille.RowCount = PiedGrille then
      LaGrilleAjouteLigne(0);

    //les propriétés
    LaGrille.OnCellExit := B2058A_WM_MOA_OnCellexit;

  end
  else if LeTool = '2058A_VA_MOA' then
  begin
    Caption := 'Imposition des PVLT';
    PiedGrille := 3;
    Height := 330;
    Width := 600;
    ValeurPVLT := Borne1;

    ActivLabel(LeLabel1, 'Explication', 6, 15, 50);
    ActivLabel(LeLabel2, 'Montant PVLT', 6, 115, 50, 75);
    ActivLabel(LeLabel3, 'Imputées sur MVLT antérieures', 6, 190, 50, 75);
    ActivLabel(LeLabel4, 'Imputées sur les déficits antérieurs', 6, 265, 50, 75);
    ActivLabel(LeLabel5, 'Imposées au taux de 8%', 6, 340, 50, 75);
    ActivLabel(LeLabel6, 'Imputées sur le déficit de l''exercice', 6, 415, 50, 75);
    ActivLabel(LeLabel7, 'Imposées à un taux réduit', 6, 490, 50, 75);

    ActivGrille(LaGrille, 8, PiedGrille, 60, 0, 200, Width - 10);
    LaGrille.ColWidths[1] := 100;
    LaGrille.ColWidths[2] := 75;
    LaGrille.ColWidths[3] := 75;
    LaGrille.ColWidths[4] := 75;
    LaGrille.ColWidths[5] := 75;
    LaGrille.ColWidths[6] := 75;
    LaGrille.ColWidths[7] := 75;
    LaGrille.ColEditables[7] := False;
    LaGrille.Cells[1, 0] := 'Total';
    LaGrille.Cells[1, 1] := 'Pour Contrôle';
    LaGrille.Cells[1, 2] := 'Ecart';

    LePanel.Visible := True;
    ActivBouton(BValide, 'Valider la saisie', 2, 510);
    ActivBouton(BAnnule, 'Annuler la saisie', 2, 550);
    ActivBouton(BAjouteLg, 'Ajouter une ligne', 2, 250);
    ActivBouton(BSupprimeLg, 'Supprimer une ligne', 2, 290);

    // recherche des éléments
    if Assigned(LaTob) then
    begin
      TobV := LaTob.FindFirst(['NATURE', 'INDICE'], ['2058A_WW_MOA', 0], False);
      if TobV <> nil then
        LaGrille.Cells[3, LaGrille.RowCount - 2] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
      { FQ 21152 BVE 20.07.07
      TobV                := LaTob.FindFirst(['NATURE', 'INDICE'], ['2058A_WX_MOA', 0], False); }
      TobV := LaTob.FindFirst(['NATURE', 'INDICE'], ['2058A_XB_MOA', 0], False);
      { END FQ 21152 }
      if TobV <> nil then
        LaGrille.Cells[4, LaGrille.RowCount - 2] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
      LaGrille.Cells[2, LaGrille.RowCount - 2] := STRFMONTANT(ValeurPVLT, 15, V_PGI.OkDecV, '', True);

      Idx := 0;
      TobV := LaTob.FindFirst(['NATURE'], ['2058A_VA_MOA'], False);
      while TobV <> nil do
      begin
        if TobV.GetValue('INDICE') <> '0' then // on ne prend pas le premier, c'est le total général
        begin
          LaGrille.RowCount := PiedGrille + Idx;
          LaGrilleAjouteLigne(Idx);
          LaGrille.Cells[1, Idx] := TobV.GetValue('COMMENTAIRE');
          LaGrille.Cells[2, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[3, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT2')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[4, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT3')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[5, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT4')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[6, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT5')), 15, V_PGI.OkDecV, '', True);
          B2058A_VA_MOA_CalculLigne(Idx);
          Inc(Idx);
        end;
        TobV := LaTob.FindNext(['NATURE'], ['2058A_VA_MOA'], False);
      end;
    end;

    if LaGrille.RowCount = PiedGrille then
    begin
      LaGrilleAjouteLigne(0);
      B2058A_VA_MOA_CalculLigne(0);
    end;

    //les propriétés
    LaGrille.OnCellExit := B2058A_VA_MOA_OnCellexit;

  end
  //==============================================================================================
  //ToolBox pour suivi des déficits (intégration fiscale)
  else if LeTool = '2058AB_AH_MOA' then
  begin
    Caption := 'Suivi des déficits (intégration fiscale)';
    PiedGrille := 1;
    Height := 410;

    ActivLabel(LeLabel1, 'Explication', 6, 15);
    ActivLabel(LeLabel2, 'Montant antérieur', 6, 115);
    ActivLabel(LeLabel3, 'Imputés sur les PVLT', 6, 215);
    ActivLabel(LeLabel4, 'Imputés sur les résultats de l''exercice', 6, 315);
    ActivLabel(LeLaBel5, 'Montant de l''exercice', 6, 415);

    ActivGrille(LaGrille, 6, 1, 60, 0, 200, Width - 10);
    LaGrille.ColEditables[5] := False;
    LaGrille.Cells[1, 0] := 'Total';

    ActivLabel(LeLabel6, 'Déficit de l''exercice', 270, 150, 13, 250);
    LeLabel6.Alignment := taRightJustify;

    ActivLabel(LeLabel7, 'Montant du déficit à reporter', 300, 150, 13, 250);
    LeLabel7.Alignment := taRightJustify;

    ActivEdit(LEdit1, 'Reprise du montant du déficit éventuel de l''exercice 2058AB_AQ_MOA', 270, 415);
    ActivEdit(LEdit2, 'Montant du déficit à reporter', 300, 415);

    LePanel.Visible := True;
    ActivBouton(BValide, 'Valider la saisie', 2, 460);
    ActivBouton(BAnnule, 'Annuler la saisie', 2, 500);
    ActivBouton(BAjouteLg, 'Ajouter une ligne', 2, 250);
    ActivBouton(BSupprimeLg, 'Supprimer une ligne', 2, 290);


    // recherche des éléments
    TobV := LaTob.FindFirst(['NATURE'], ['2058AB_AQ_MOA'], False);
    if TobV <> nil then
      LEdit1.Text := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True)
    else
      LEdit1.Text := STRFMONTANT(0, 15, V_PGI.OkDecV, '', True);

    TobV := LaTob.FindFirst(['NATURE'], ['2058BB_EH_MOA'], False);
    if TobV <> nil then
      LEdit1.Text := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True)
    else
      LEdit1.Text := STRFMONTANT(0, 15, V_PGI.OkDecV, '', True);

    Idx := 0;
    if Assigned(LaTob) then
    begin
      TobV := LaTob.FindFirst(['NATURE'], ['2058AB_AH_MOA'], False);
      while TobV <> nil do
      begin
        if TobV.GetValue('INDICE') <> '0' then // on ne prend pas le premier, c'est le total général
        begin
          LaGrille.RowCount := Idx + 1;
          LaGrilleAjouteLigne(Idx);
          LaGrille.Cells[1, Idx] := TobV.GetValue('COMMENTAIRE');
          LaGrille.Cells[2, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[3, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT2')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[4, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT3')), 15, V_PGI.OkDecV, '', True);
          B2058AB_AH_MOA_CalculLigne(Idx);
          Inc(Idx);
        end;
        TobV := LaTob.FindNext(['NATURE'], ['2058AB_AH_MOA'], False);
      end;
    end;

    if LaGrille.RowCount = PiedGrille then
      LaGrilleAjouteLigne(0);

    //les propriétés
    LaGrille.OnCellExit := B2058AB_AH_MOA_OnCellexit;

  end
  //==============================================================================================
  //ToolBox pour suivi des MVLT (intégration fiscale)
  else if LeTool = '2058AB_AG_MOA' then
  begin
    Caption := 'Suivi des MVLT (Intégration fiscale)';
    PiedGrille := 1;
    Height := 330;
    Width := 650;

    ActivLabel(LeLabel1, 'Explication', 6, 15, 50);
    ActivLabel(LeLabel2, 'Montant antérieur', 6, 115);
    ActivLabel(LeLabel3, 'PVLT imputées sur MVLT antérieures', 6, 215);
    ActivLabel(LeLabel4, 'MVLT imputées sur le résultat de l''exercice', 6, 315);
    ActivLabel(LeLabel5, 'MVLT de l''exercice', 6, 415);
    ActivLabel(LeLabel6, 'Montant de fin d''exercice', 6, 515);

    ActivGrille(LaGrille, 7, 1, 60, 0, 200, Width - 10);
    LaGrille.ColEditables[6] := False;
    LaGrille.Cells[1, 0] := 'Total';

    LePanel.Visible := True;
    ActivBouton(BValide, 'Valider la saisie', 2, 560);
    ActivBouton(BAnnule, 'Annuler la saisie', 2, 600);
    ActivBouton(BAjouteLg, 'Ajouter une ligne', 2, 250);
    ActivBouton(BSupprimeLg, 'Supprimer une ligne', 2, 290);


    // recherche des éléments
    Idx := 0;
    if Assigned(LaTob) then
    begin
      TobV := LaTob.FindFirst(['NATURE'], ['2058AB_AG_MOA'], False);
      while TobV <> nil do
      begin
        if TobV.GetValue('INDICE') <> '0' then // on ne prend pas le premier, c'est le total général
        begin
          LaGrille.RowCount := Idx + 1;
          LaGrilleAjouteLigne(Idx);
          LaGrille.Cells[1, Idx] := TobV.GetValue('COMMENTAIRE');
          LaGrille.Cells[2, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[3, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT2')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[4, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT3')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[5, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT4')), 15, V_PGI.OkDecV, '', True);
          B2058AB_AG_MOA_CalculLigne(Idx);
          Inc(Idx);
        end;
        TobV := LaTob.FindNext(['NATURE'], ['2058AB_AG_MOA'], False);
      end;
    end;

    if LaGrille.RowCount = PiedGrille then
      LaGrilleAjouteLigne(0);

    //les propriétés
    LaGrille.OnCellExit := B2058AB_AG_MOA_OnCellexit;

  end
  //==============================================================================================
  //ToolBox pour Jeton de présence
  else if LeTool = 'JETONPRESENCE' then
  begin
    Caption := 'Jetons de présence';
    Width := 550;
    Height := 160;
    BaseJeton := Borne1;

    ActivLabel(LeLabel1, 'Jetons versés', 6, 15);
    ActivLabel(LeLabel2, 'Quote-part', 6, 115);
    ActivLabel(LeLabel3, 'Base', 6, 215);
    ActivLabel(LeLabel4, 'Membres du conseil', 6, 315);
    ActivLabel(LeLabel5, 'Total', 6, 415);

    ActivGrille(LaGrille, 6, 1, 40, 0, 30, Width - 10);
    LaGrille.ColWidths[1] := 100;
    LaGrille.ColTypes[1] := 'K';
    LaGrille.ColFormats[1] := 'R';
    LaGrille.ColAligns[1] := taRightJustify;

    LaGrille.ColEditables[2] := False;
    LaGrille.ColEditables[5] := False;

    ActivCoche(LaCoche, 'Entreprise employant moins de 5 personnes', 80);

    LePanel.Visible := True;
    ActivBouton(BValide, 'Valider la saisie', 2, 450);
    ActivBouton(BAnnule, 'Annuler la saisie', 2, 490);

    if Assigned(LaTob) then
    begin
      TobV := LaTob.FindFirst(['NATURE', 'INDICE'], ['JETONPRESENCE', 1], False);
      if Assigned(TobV) then
      begin
        LaGrille.Cells[1, 0] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
        LaGrille.Cells[2, 0] := STRFMONTANT(VALEUR(TobV.GetValue('MT2')), 15, V_PGI.OkDecV, '', True);
        LaGrille.Cells[3, 0] := STRFMONTANT(VALEUR(TobV.GetValue('MT3')), 15, V_PGI.OkDecV, '', True);
        LaGrille.Cells[4, 0] := STRFMONTANT(VALEUR(TobV.GetValue('MT4')), 15, V_PGI.OkDecV, '', True);
        if Copy(TobV.GetValue('COMMENTAIRE'), 1, 1) = 'X' then
          LaCoche.Checked := True;
      end
      else
      begin
        LaGrille.Cells[2, 0] := '5';
      end;
    end;
    BJetonPresence_CalculLigne;

    LaGrille.OnCellExit := BJetonPresence_OnCellexit;
    LaCoche.OnClick := BJetonPresence_LaCocheOnClick;

  end
  //==============================================================================================
  //ToolBox pour Cession des immobilisations
  else if LeTool = 'CESSIONIMO' then
  begin
    Caption := 'Cession des immobilisations';
    PiedGrille := 1;
    Height := 330;

    ActivLabel(LeLabel1, 'Explication', 6, 15, 50, 200);
    ActivLabel(LeLabel2, 'PVLT Total', 6, 215, 50);
    ActivLabel(LeLabel3, 'Long terme', 6, 315, 50);
    ActivLabel(LeLabel4, 'Court terme', 6, 415, 50);

    ActivGrille(LaGrille, 5, 1, 60, 0, 200, Width - 10);
    LaGrille.ColWidths[1] := 200;
    LAgrille.Cells[1, 0] := 'Total';

    LePanel.Visible := True;
    ActivBouton(BValide, 'Valider la saisie', 2, 460);
    ActivBouton(BAnnule, 'Annuler la saisie', 2, 500);
    ActivBouton(BAjouteLg, 'Ajouter une ligne', 2, 250);
    ActivBouton(BSupprimeLg, 'Supprimer une ligne', 2, 290);

    // recherche des éléments
    Idx := 0;
    if Assigned(LaTob) then
    begin
      TobV := LaTob.FindFirst(['NATURE'], ['CESSIONIMO'], False);
      while TobV <> nil do
      begin
        if TobV.GetValue('INDICE') <> '0' then // on ne prend pas le premier, c'est le total général
        begin
          LaGrille.RowCount := Idx + 1;
          LaGrilleAjouteLigne(Idx);
          LaGrille.Cells[1, Idx] := TobV.GetValue('COMMENTAIRE');
          LaGrille.Cells[2, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[3, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT2')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[4, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT3')), 15, V_PGI.OkDecV, '', True);
          CessionImmo_CalculLigne(Idx);
          Inc(Idx);
        end;
        TobV := LaTob.FindNext(['NATURE'], ['CESSIONIMO'], False);
      end;
    end;

    if LaGrille.RowCount = PiedGrille then
      LaGrilleAjouteLigne(0);


    //les propriétés
    LaGrille.OnCellExit := CessionImmo_OnCellexit;


  end
  //==============================================================================================
  //ToolBox pour imposition des PVLT (intégration fiscale)
  else if LeTool = '2058AB_AF_MOA' then
  begin
    Caption := 'Imposition des PVLT (intégration fiscale)';
    PiedGrille := 3;
    Height := 330;
    Width := 600;
    ValeurPVLT := Borne1;

    ActivLabel(LeLabel1, 'Explication', 6, 15, 50);
    ActivLabel(LeLabel2, 'Montant PVLT', 6, 115, 50, 75);
    ActivLabel(LeLabel3, 'Imputées sur MVLT antérieures', 6, 190, 50, 75);
    ActivLabel(LeLabel4, 'Imputées sur les déficits antérieurs', 6, 265, 50, 75);
    ActivLabel(LeLabel5, 'Imposées au taux de 8%', 6, 340, 50, 75);
    ActivLabel(LeLabel6, 'Imputées sur le déficit de l''exercice', 6, 415, 50, 75);
    ActivLabel(LeLabel7, 'Imposées à un taux réduit', 6, 490, 50, 75);

    ActivGrille(LaGrille, 8, PiedGrille, 60, 0, 200, Width - 10);
    LaGrille.ColWidths[1] := 100;
    LaGrille.ColWidths[2] := 75;
    LaGrille.ColWidths[3] := 75;
    LaGrille.ColWidths[4] := 75;
    LaGrille.ColWidths[5] := 75;
    LaGrille.ColWidths[6] := 75;
    LaGrille.ColWidths[7] := 75;
    LaGrille.ColEditables[7] := False;
    LaGrille.Cells[1, 0] := 'Total';
    LaGrille.Cells[1, 1] := 'Pour Contrôle';
    LaGrille.Cells[1, 2] := 'Ecart';

    LePanel.Visible := True;
    ActivBouton(BValide, 'Valider la saisie', 2, 510);
    ActivBouton(BAnnule, 'Annuler la saisie', 2, 550);
    ActivBouton(BAjouteLg, 'Ajouter une ligne', 2, 250);
    ActivBouton(BSupprimeLg, 'Supprimer une ligne', 2, 290);

    // recherche des éléments
    if Assigned(LaTob) then
    begin
      TobV := LaTob.FindFirst(['NATURE', 'INDICE'], ['2058AB_AG_MOA', 0], False);
      if TobV <> nil then
        LaGrille.Cells[3, LaGrille.RowCount - 2] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
      TobV := LaTob.FindFirst(['NATURE', 'INDICE'], ['2058AB_AH_MOA', 0], False);
      if TobV <> nil then
        LaGrille.Cells[4, LaGrille.RowCount - 2] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);

      Idx := 0;
      TobV := LaTob.FindFirst(['NATURE'], ['2058AB_AF_MOA'], False);
      while TobV <> nil do
      begin
        if TobV.GetValue('INDICE') <> '0' then // on ne prend pas le premier, c'est le total général
        begin
          LaGrille.RowCount := PiedGrille + Idx;
          LaGrilleAjouteLigne(Idx);
          LaGrille.Cells[1, Idx] := TobV.GetValue('COMMENTAIRE');
          LaGrille.Cells[2, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[3, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT2')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[4, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT3')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[5, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT4')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[6, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT5')), 15, V_PGI.OkDecV, '', True);
          B2058AB_AF_MOA_CalculLigne(Idx);
          Inc(Idx);
        end;
        TobV := LaTob.FindNext(['NATURE'], ['2058AB_AF_MOA'], False);
      end;
    end;

    if LaGrille.RowCount = PiedGrille then
    begin
      LaGrilleAjouteLigne(0);
      B2058AB_AF_MOA_CalculLigne(0);
    end;

    //les propriétés
    LaGrille.OnCellExit := B2058AB_AF_MOA_OnCellexit;




  end
  else
  //==============================================================================================
  //ToolBox pour saisie détaillée
  //LeTitre = le nom de la saisie multiple
  //LeTag = Sens de l'ecriture 4 -> réintégration 5 -> déduction 6 -> les deux
  begin
    Caption := 'Saisie détaillée : ' + LeTitre;
    PiedGrille := 1;
    Height := 330;
    TheTool := 'SAISIEGRILLE';
    TheTag := LeTag;


    ActivLabel(LeLabel1, 'Explication', 6, 15, 50, 300);
    ActivLabel(LeLabel2, 'Réintégration', 6, 315, 50);
    ActivLabel(LeLabel3, 'Déduction', 6, 415, 50);

    ActivGrille(LaGrille, 4, 1, 60, 0, 200, Width - 10);
    LaGrille.Cells[1, 0] := 'Total';
    LaGrille.ColWidths[1] := 300;
    if Letag = 4 then
      LaGrille.ColEditables[3] := False
    else if LeTag = 5 then
      LaGrille.ColEditables[2] := False;


    LePanel.Visible := True;
    ActivBouton(BValide, 'Valider la saisie', 2, 460);
    ActivBouton(BAnnule, 'Annuler la saisie', 2, 500);
    ActivBouton(BAjouteLg, 'Ajouter une ligne', 2, 250);
    ActivBouton(BSupprimeLg, 'Supprimer une ligne', 2, 290);

    // recherche des éléments
    Idx := 0;
    if Assigned(LaTob) then
    begin
      TobV := LaTob.FindFirst(['NATURE'], [LeTool], False);
      while TobV <> nil do
      begin
        if TobV.GetValue('INDICE') <> '0' then // on ne prend pas le premier, c'est le total général
        begin
          LaGrille.RowCount := Idx + 1;
          LaGrilleAjouteLigne(Idx);
          LaGrille.Cells[1, Idx] := TobV.GetValue('COMMENTAIRE');
  (*        if LeTag = 4 then
            LaGrille.Cells[2, Idx]  := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True)
          else if LeTag = 5 then
            LaGrille.Cells[3, Idx]  := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
  *)
          LaGrille.Cells[2, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MONTANT')), 15, V_PGI.OkDecV, '', True);
          LaGrille.Cells[3, Idx] := STRFMONTANT(VALEUR(TobV.GetValue('MT2')), 15, V_PGI.OkDecV, '', True);

          BDetail_CalculLigne(Idx);
          Inc(Idx);
        end;
        TobV := LaTob.FindNext(['NATURE'], [LeTool], False);
      end;
    end;

    if (idx = 0) then
    begin
      TobV := LaTob.FindFirst(['NATURE', 'INDICE'], [LeTool, 0], False);
      if TobV <> nil then
      begin
        if LeTag = 4 then
          RepriseSaisieUnique(LaTob, LeTool, 1, 2)
        else
          RepriseSaisieUnique(LaTob, Letool, 1, 3);
        BDetail_CalculLigne(0);
      end;
    end;

    if LaGrille.RowCount = PiedGrille then
      LaGrilleAjouteLigne(0);

    //les propriétés
    LaGrille.OnCellExit := BDetail_OnCellexit;
    LaGrille.OnCellEnter := BDetail_OnCellEnter;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 01/03/2007
Modifié le ... :   /  /
Description .. : Recherche dans la tob s'il n'y a qu'un seul enregistrement
Mots clefs ... :
*****************************************************************}

procedure TDRFToolBox.RepriseSaisieUnique(LaTob: Tob; LaNat: string; LeCom, LeMt1: integer);
var
  TobV: Tob;

begin
  TobV := LaTob.FindFirst(['NATURE', 'INDICE'], [LaNat, 0], False);
  if (Assigned(TobV)) and ((TobV.GetValue('COMMENTAIRE') <> '') or (TobV.GetValue('MONTANT') <> 0)) then
  begin
    if LeCom <> -1 then
      LaGrille.Cells[LeCom, 0] := TobV.GetValue('COMMENTAIRE');
    if LeMt1 <> -1 then
      LaGrille.Cells[LeMt1, 0] := TobV.GetValue('MONTANT');
    LaGrilleAjouteLigne(1);
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 12/01/2007
Modifié le ... :   /  /
Description .. : événement sur sortie de cellule pour outil suivi déficits
Mots clefs ... :
*****************************************************************}

procedure TDRFToolBox.B2058A_XL_MOA_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if csDestroying in ComponentState then Exit;
  B2058A_XL_MOA_CalculLigne(Arow);
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 19/01/2007
Modifié le ... :   /  /
Description .. : événement sur sortie de cellule pour saisie en détail
Mots clefs ... :
*****************************************************************}

procedure TDRFToolBox.BDetail_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if csDestroying in ComponentState then Exit;
  BDetail_CalculLigne(ARow);
end;

procedure TDRFToolBox.BDetail_OnCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if csDestroying in ComponentState then Exit;

  if (TheTag = 4) and (LaGrille.Col = 3) then
  begin
    if LaGrille.row = LaGrille.RowCount - 2 then
    begin
      LaGrille.Col := ACol;
      LaGrille.Row := ARow;
      LaGrille.SetFocus;
    end
    else
    begin
      LaGrille.Col := 1;
      LaGrille.Row := ARow + 1;
      LaGrille.SetFocus;
    end;
  end;
  if (TheTag = 5) and (LaGrille.Col = 2) then
  begin
    if Acol = 1 then
    begin
      LaGrille.Col := 3;
      LaGrille.Row := ARow;
      LaGrille.SetFocus;
    end
    else
    begin
      LaGrille.Col := 1;
      LaGrille.Row := ARow;
      LaGrille.SetFocus;
    end;

  end;

  if LaGrille.Row = LaGrille.RowCount - 1 then
  begin
    LaGrille.Col := ACol;
    LaGrille.Row := ARow;
    LaGrille.SetFocus;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 16/01/2007
Modifié le ... :   /  /
Description .. : Calcul de la ligne pour outils 2058A_XL_MOA
Suite ........ : Suivi des déficits
Mots clefs ... :
*****************************************************************}

procedure TDRFToolBox.B2058A_XL_MOA_CalculLigne(Index: integer);
var
  V05_1: extended;
  V05_2: extended;
  V05_3: extended;
  V05_4: extended;
  V05_5: extended;
  i: integer;
  Tot051: Extended;
  Tot052: Extended;
  Tot053: Extended;
  Tot054: Extended;
  Tot056: Extended;

begin
  V05_1 := VALEUR(LaGrille.Cells[2, Index]);
  V05_2 := VALEUR(LaGrille.Cells[3, Index]);
  V05_3 := VALEUR(LaGrille.Cells[4, Index]);
  V05_5 := VALEUR(LEdit1.Text);

  if (V05_2 + V05_3) > V05_1 then
  begin
    PGIInfo('La somme des montants imputés ne peut pas être suppérieur au montant antérieur');
    LaGrille.Col := 2;
    LaGrille.Row := Index;
    LaGrille.SetFocus;
    exit;
  end;

  V05_4 := V05_1 - V05_2 - V05_3;
  LaGrille.Cells[1, Index] := Copy(LaGrille.Cells[1, Index], 1, 35);
  LaGrille.Cells[2, Index] := STRFMONTANT(V05_1, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Index] := STRFMONTANT(V05_2, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Index] := STRFMONTANT(V05_3, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[5, Index] := STRFMONTANT(V05_4, 15, V_PGI.OkDecV, '', True);

  Tot051 := 0;
  Tot052 := 0;
  Tot053 := 0;
  Tot054 := 0;

  for i := 0 to LaGrille.RowCount - 2 do
  begin
    Tot051 := Tot051 + VALEUR(LaGrille.Cells[2, i]);
    Tot052 := Tot052 + VALEUR(LaGrille.Cells[3, i]);
    Tot053 := Tot053 + VALEUR(LaGrille.Cells[4, i]);
    Tot054 := Tot054 + VALEUR(LaGrille.Cells[5, i]);
  end;
  LaGrille.Cells[2, Lagrille.RowCount - 1] := STRFMONTANT(Tot051, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Lagrille.RowCount - 1] := STRFMONTANT(Tot052, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Lagrille.RowCount - 1] := STRFMONTANT(Tot053, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[5, Lagrille.RowCount - 1] := STRFMONTANT(Tot054, 15, V_PGI.OkDecV, '', True);

  Tot056 := Tot054 + V05_5;
  LEdit2.Text := STRFMONTANT(Tot056, 15, V_PGI.OkDecV, '', True);
end;


procedure TDRFToolBox.BDetail_CalculLigne(Index: integer);
var
  VRei: extended;
  VDed: extended;
  TotRei: extended;
  TotDed: extended;
  i: integer;

begin
  VRei := VALEUR(LaGrille.Cells[2, Index]);
  VDed := VALEUR(LaGrille.Cells[3, Index]);

  LaGrille.Cells[2, Index] := STRFMONTANT(VRei, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Index] := STRFMONTANT(VDed, 15, V_PGI.OkDecV, '', True);

  TotRei := 0;
  TotDed := 0;

  for i := 0 to LaGrille.RowCount - (1 + PiedGrille) do
  begin
    TotRei := TotRei + VALEUR(LaGrille.Cells[2, i]);
    TotDed := TotDed + VALEUR(LaGrille.Cells[3, i]);
  end;
  LaGrille.Cells[2, LaGrille.RowCount - PiedGrille] := STRFMONTANT(TotRei, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, LaGrille.RowCount - PiedGrille] := STRFMONTANT(TotDed, 15, V_PGI.OkDecV, '', True);
end;


procedure TDRFToolBox.B2058A_WC_MOA_CalculLigne(Index: integer);
var
  V01_1: extended;
  V01_2: extended;
  V01_3: extended;
  i: integer;
  Tot011: Extended;
  Tot012: Extended;
  Tot013: Extended;

begin
  V01_1 := VALEUR(LaGrille.Cells[2, Index]);
  V01_2 := VALEUR(LaGrille.Cells[3, Index]);

  V01_3 := V01_1 - V01_2;
  if V01_3 < 0 then
  begin
    PGIInfo('Le montant à réintégrer ne peut être négatif');
    LaGrille.Col := 2;
    LaGrille.Row := Index;
    LaGrille.Setfocus;
    exit;
  end;

  LaGrille.Cells[1, Index] := Copy(LaGrille.Cells[1, Index], 1, 35);
  LaGrille.Cells[2, Index] := STRFMONTANT(V01_1, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Index] := STRFMONTANT(V01_2, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Index] := STRFMONTANT(V01_3, 15, V_PGI.OkDecV, '', True);

  Tot011 := 0;
  Tot012 := 0;
  Tot013 := 0;

  for i := 0 to LaGrille.RowCount - 2 do
  begin
    Tot011 := Tot011 + VALEUR(LaGrille.Cells[2, i]);
    Tot012 := Tot012 + VALEUR(LaGrille.Cells[3, i]);
    Tot013 := Tot013 + VALEUR(LaGrille.Cells[4, i]);
  end;
  LaGrille.Cells[2, Lagrille.RowCount - 1] := STRFMONTANT(Tot011, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Lagrille.RowCount - 1] := STRFMONTANT(Tot012, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Lagrille.RowCount - 1] := STRFMONTANT(Tot013, 15, V_PGI.OkDecV, '', True);

end;

procedure TDRFToolBox.B2058A_WC_MOA_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  V01_2: extended;

begin
  if csDestroying in ComponentState then Exit;
  V01_2 := VALEUR(LaGrille.Cells[3, Arow]);
  if (Acol = 3) and ((V01_2 < RemConjointMin) or (V01_2 > RemConjointMax)) then
  begin
    PGIInfo('Le montant doit être au minimun de ' + FloatToStr(RemConjointMin) + '€ et au maiximun de ' + FloatToStr(RemConjointMax) + '€');
    LaGrille.Col := Acol;
    LaGrille.Row := Arow;
    LaGrille.SetFocus;
    exit;
  end;

  B2058A_WC_MOA_CalculLigne(ARow);

end;



procedure TDRFToolBox.B2058A_XA_MOA_CalculLigne(Index: integer);
var
  V02_1: extended;
  V02_2: extended;
  V02_3: extended;
  i: integer;
  Tot021: Extended;
  Tot022: Extended;
  Tot023: Extended;

begin
  V02_1 := VALEUR(LaGrille.Cells[2, Index]);

  V02_2 := V02_1 * 5 / 100;
  V02_3 := V02_1 - V02_2;

  LaGrille.Cells[1, Index] := Copy(LaGrille.Cells[1, Index], 1, 35);
  LaGrille.Cells[2, Index] := STRFMONTANT(V02_1, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Index] := STRFMONTANT(V02_2, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Index] := STRFMONTANT(V02_3, 15, V_PGI.OkDecV, '', True);

  Tot021 := 0;
  Tot022 := 0;
  Tot023 := 0;

  for i := 0 to LaGrille.RowCount - 2 do
  begin
    Tot021 := Tot021 + VALEUR(LaGrille.Cells[2, i]);
    Tot022 := Tot022 + VALEUR(LaGrille.Cells[3, i]);
    Tot023 := Tot023 + VALEUR(LaGrille.Cells[4, i]);
  end;
  LaGrille.Cells[2, Lagrille.RowCount - 1] := STRFMONTANT(Tot021, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Lagrille.RowCount - 1] := STRFMONTANT(Tot022, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Lagrille.RowCount - 1] := STRFMONTANT(Tot023, 15, V_PGI.OkDecV, '', True);


end;

procedure TDRFToolBox.B2058A_XA_MOA_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if csDestroying in ComponentState then Exit;
  B2058A_XA_MOA_CalculLigne(ARow);
end;


procedure TDRFToolBox.B2058B_PCP_CalculLigne(Index: integer);
var
  V01: extended;
  V02: extended;
  i: integer;
  Tot01: Extended;
  Tot02: Extended;

begin
  V01 := VALEUR(LaGrille.Cells[2, Index]);
  V02 := VALEUR(LaGrille.Cells[3, Index]);


  LaGrille.Cells[1, Index] := Copy(LaGrille.Cells[1, Index], 1, 35);
  LaGrille.Cells[2, Index] := STRFMONTANT(V01, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Index] := STRFMONTANT(V02, 15, V_PGI.OkDecV, '', True);

  Tot01 := 0;
  Tot02 := 0;

  for i := 0 to LaGrille.RowCount - 2 do
  begin
    Tot01 := Tot01 + VALEUR(LaGrille.Cells[2, i]);
    Tot02 := Tot02 + VALEUR(LaGrille.Cells[3, i]);
  end;
  LaGrille.Cells[2, Lagrille.RowCount - 1] := STRFMONTANT(Tot01, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Lagrille.RowCount - 1] := STRFMONTANT(Tot02, 15, V_PGI.OkDecV, '', True);

end;

procedure TDRFToolBox.B2058B_PCP_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
(*
  if (VALEUR(LaGrille.Cells[2, ARow]) <> 0) and (VALEUR(LaGrille.Cells[3, ARow]) <> 0) then
  begin
    PGIInfo('Il ne peut y avoir une dotation et une reprise en même temps');
    LaGrille.Col        := ACol;
    LaGrille.Row        := ARow;
    LaGrille.SetFocus;
    exit;
  end;
*)
  if csDestroying in ComponentState then Exit;
  B2058B_PCP_CalculLigne(ARow);
end;

procedure TDRFToolBox.B2058A_WN_MOA_CalculLigne(Index: integer);
var
  V03_1: extended;
  V03_2: extended;
  V03_3: extended;
  V03_4: extended;
  i: integer;
  Tot031: Extended;
  Tot032: Extended;
  Tot033: Extended;
  Tot034: Extended;

begin
  V03_1 := VALEUR(LaGrille.Cells[2, Index]);
  V03_2 := VALEUR(LaGrille.Cells[3, Index]);
  V03_3 := VALEUR(LaGrille.Cells[4, Index]);

  if V03_2 > V03_1 then
  begin
    PGIInfo('Le montant rapporté est supérieur au montant antérieur');
    LaGrille.Col := 3;
    LaGrille.Row := Index;
    LaGrille.SetFocus;
    Exit;
  end;

  V03_4 := V03_1 + V03_2 - V03_3;

  LaGrille.Cells[1, Index] := Copy(LaGrille.Cells[1, Index], 1, 35);
  LaGrille.Cells[2, Index] := STRFMONTANT(V03_1, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Index] := STRFMONTANT(V03_2, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Index] := STRFMONTANT(V03_3, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[5, Index] := STRFMONTANT(V03_4, 15, V_PGI.OkDecV, '', True);

  Tot031 := 0;
  Tot032 := 0;
  Tot033 := 0;
  Tot034 := 0;

  for i := 0 to LaGrille.RowCount - 2 do
  begin
    Tot031 := Tot031 + VALEUR(LaGrille.Cells[2, i]);
    Tot032 := Tot032 + VALEUR(LaGrille.Cells[3, i]);
    Tot033 := Tot033 + VALEUR(LaGrille.Cells[4, i]);
    Tot034 := Tot034 + VALEUR(LaGrille.Cells[5, i]);
  end;
  LaGrille.Cells[2, Lagrille.RowCount - 1] := STRFMONTANT(Tot031, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Lagrille.RowCount - 1] := STRFMONTANT(Tot032, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Lagrille.RowCount - 1] := STRFMONTANT(Tot033, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[5, Lagrille.RowCount - 1] := STRFMONTANT(Tot034, 15, V_PGI.OkDecV, '', True);

end;

procedure TDRFToolBox.B2058A_WN_MOA_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if csDestroying in ComponentState then Exit;
  B2058A_WN_MOA_CalculLigne(ARow);

end;

procedure TDRFToolBox.B2058A_WM_MOA_CalculLigne(Index: integer);
var
  V04_1: extended;
  V04_2: extended;
  V04_3: extended;
  V04_4: extended;
  V04_5: extended;
  i: integer;
  Tot041: Extended;
  Tot042: Extended;
  Tot043: Extended;
  Tot044: Extended;
  Tot045: Extended;

begin
  V04_1 := VALEUR(LaGrille.Cells[2, Index]);
  V04_2 := VALEUR(LaGrille.Cells[3, Index]);
  V04_3 := VALEUR(LaGrille.Cells[4, Index]);
  V04_4 := VALEUR(LaGrille.Cells[5, Index]);

  if (V04_2 + V04_3) > V04_1 then
  begin
    PGIInfo('La somme des imputations est supérieure au montant antérieur');
    LaGrille.Col := 3;
    LaGrille.Row := Index;
    LaGrille.SetFocus;
    Exit;
  end;
  V04_5 := V04_1 - V04_2 + V04_4 - V04_3;
  LaGrille.Cells[1, Index] := Copy(LaGrille.Cells[1, Index], 1, 35);
  LaGrille.Cells[2, Index] := STRFMONTANT(V04_1, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Index] := STRFMONTANT(V04_2, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Index] := STRFMONTANT(V04_3, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[5, Index] := STRFMONTANT(V04_4, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[6, Index] := STRFMONTANT(V04_5, 15, V_PGI.OkDecV, '', True);

  Tot041 := 0;
  Tot042 := 0;
  Tot043 := 0;
  Tot044 := 0;
  Tot045 := 0;

  for i := 0 to LaGrille.RowCount - 2 do
  begin
    Tot041 := Tot041 + VALEUR(LaGrille.Cells[2, i]);
    Tot042 := Tot042 + VALEUR(LaGrille.Cells[3, i]);
    Tot043 := Tot043 + VALEUR(LaGrille.Cells[4, i]);
    Tot044 := Tot044 + VALEUR(LaGrille.Cells[5, i]);
    { FQ 21551 BVE 01.10.07
    Tot045 := Tot045 + VALEUR(LaGrille.Cells[5, i]); }
    Tot045 := Tot045 + VALEUR(LaGrille.Cells[6, i]);
    { END FQ 21551 }
  end;

  LaGrille.Cells[2, Lagrille.RowCount - 1] := STRFMONTANT(Tot041, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Lagrille.RowCount - 1] := STRFMONTANT(Tot042, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Lagrille.RowCount - 1] := STRFMONTANT(Tot043, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[5, Lagrille.RowCount - 1] := STRFMONTANT(Tot044, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[6, Lagrille.RowCount - 1] := STRFMONTANT(Tot045, 15, V_PGI.OkDecV, '', True);


end;

procedure TDRFToolBox.B2058A_WM_MOA_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if csDestroying in ComponentState then Exit;
  B2058A_WM_MOA_CalculLigne(ARow);
end;

procedure TDRFToolBox.B2058A_VA_MOA_CalculLigne(Index: integer);
var
  V06_8: extended;
  V06_9: extended;
  V06_10: extended;
  V06_11: extended;
  V06_12: extended;
  V06_13: extended;
  i: integer;
  Tot068: Extended;
  Tot069: Extended;
  Tot0610: Extended;
  Tot0611: Extended;
  Tot0612: Extended;
  Tot0613: Extended;
  PC8: Extended;
  PC9: Extended;
  PC10: Extended;
  Ecart8: extended;
  Ecart9: extended;
  Ecart10: extended;

begin
  V06_8 := VALEUR(LaGrille.Cells[2, Index]);
  V06_9 := VALEUR(LaGrille.Cells[3, Index]);
  V06_10 := VALEUR(LaGrille.Cells[4, Index]);
  V06_11 := VALEUR(LaGrille.Cells[5, Index]);
  V06_12 := VALEUR(LaGrille.Cells[6, Index]);
  V06_13 := V06_8 - (V06_9 + V06_10 + V06_11 + V06_12);

  LaGrille.Cells[1, Index] := Copy(LaGrille.Cells[1, Index], 1, 35);
  LaGrille.Cells[2, Index] := STRFMONTANT(V06_8, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Index] := STRFMONTANT(V06_9, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Index] := STRFMONTANT(V06_10, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[5, Index] := STRFMONTANT(V06_11, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[6, Index] := STRFMONTANT(V06_12, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[7, Index] := STRFMONTANT(V06_13, 15, V_PGI.OkDecV, '', True);

  Tot068 := 0;
  Tot069 := 0;
  Tot0610 := 0;
  Tot0611 := 0;
  Tot0612 := 0;
  Tot0613 := 0;

  for i := 0 to LaGrille.RowCount - 4 do
  begin
    Tot068 := Tot068 + VALEUR(LaGrille.Cells[2, i]);
    Tot069 := Tot069 + VALEUR(LaGrille.Cells[3, i]);
    Tot0610 := Tot0610 + VALEUR(LaGrille.Cells[4, i]);
    Tot0611 := Tot0611 + VALEUR(LaGrille.Cells[5, i]);
    Tot0612 := Tot0612 + VALEUR(LaGrille.Cells[6, i]);
    Tot0613 := Tot0613 + VALEUR(LaGrille.Cells[7, i]);
  end;

  LaGrille.Cells[2, Lagrille.RowCount - 3] := STRFMONTANT(Tot068, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Lagrille.RowCount - 3] := STRFMONTANT(Tot069, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Lagrille.RowCount - 3] := STRFMONTANT(Tot0610, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[5, Lagrille.RowCount - 3] := STRFMONTANT(Tot0611, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[6, Lagrille.RowCount - 3] := STRFMONTANT(Tot0612, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[7, Lagrille.RowCount - 3] := STRFMONTANT(Tot0613, 15, V_PGI.OkDecV, '', True);

  PC8 := VALEUR(LaGrille.Cells[2, Lagrille.RowCount - 2]);
  PC9 := VALEUR(LaGrille.Cells[3, Lagrille.RowCount - 2]);
  PC10 := VALEUR(LaGrille.Cells[4, Lagrille.RowCount - 2]);
  Ecart8 := Tot068 - PC8;
  Ecart9 := Tot069 - PC9;
  Ecart10 := Tot0610 - PC10;
  LaGrille.Cells[2, Lagrille.RowCount - 1] := STRFMONTANT(Ecart8, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Lagrille.RowCount - 1] := STRFMONTANT(Ecart9, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Lagrille.RowCount - 1] := STRFMONTANT(Ecart10, 15, V_PGI.OkDecV, '', True);



end;

procedure TDRFToolBox.B2058A_VA_MOA_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if csDestroying in ComponentState then Exit;
  B2058A_VA_MOA_CalculLigne(Arow);
end;








function TDRFToolBox.RegimeSocieteMere: Extended;
var
  StrSql: string;
  Tob761: Tob;
  Tot761: Extended;

begin
  StrSql := 'SELECT G_GENERAL, G_TOTALDEBIT, G_TOTALCREDIT, G_TOTALCREDIT-G_TOTALDEBIT RES, "X" RECH ' +
    'FROM GENERAUX WHERE G_GENERAL like "761%"';

  Tob761 := Tob.Create('CG 761', nil, -1);
  Tob761.LoadDetailFromSQL(StrSql);
  Tot761 := Tob761.Somme('RES', ['RECH'], ['X'], True);

  Result := Tot761;
  Tob761.Free;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 26/01/2007
Modifié le ... : 29/01/2007
Description .. : Chargement des infos contenu dans le Blob vers une Tob
Suite ........ : Paramètres :
Suite ........ : - Exo -> code de l'exercice
Mots clefs ... :
*****************************************************************}

procedure ChargeTobDRF(TobGene: TOB; Exo: string);
var
  StrSql: string;
  Q: TQuery;
  StrStream: TStringStream;
  Entete: Boolean;
  Encoding: string;
  TobTable: Tob;

begin
  TobTable := TOB.Create('CREVBLOCNOTE', nil, -1);
  try
    StrSql := 'SELECT * FROM CREVBLOCNOTE WHERE CBN_EXERCICE="' + Exo + '" AND CBN_NATURE="DRF" AND CBN_CODE="BIC"';
    Q := OpenSQL(StrSql, True);
    TobTable.SelectDB('', Q);
    Ferme(Q);
    if TobTable.GetString('CBN_EXERCICE') = '' then
      TobTable.SetString('CBN_EXERCICE', Exo);

    StrStream := TStringStream.Create(TobTable.GetString('CBN_BLOCNOTE'));
    try
      if StrStream.Size > 0 then
      begin
        StrStream.Seek(0, soFromBeginning);
        TobGene.LoadFromXMLStream(StrStream, Entete, Encoding);
      end;
    finally
      strStream.Free;
    end;
  finally
    TobTable.Free;
  end;
end;


procedure TDRFToolBox.B2058AB_AH_MOA_CalculLigne(Index: integer);
var
  V08_1: extended;
  V08_2: extended;
  V08_3: extended;
  V08_4: extended;
  V08_5: extended;
  i: integer;
  Tot081: Extended;
  Tot082: Extended;
  Tot083: Extended;
  Tot084: Extended;
  Tot086: Extended;

begin
  V08_1 := VALEUR(LaGrille.Cells[2, Index]);
  V08_2 := VALEUR(LaGrille.Cells[3, Index]);
  V08_3 := VALEUR(LaGrille.Cells[4, Index]);
  V08_5 := VALEUR(LEdit1.Text);

  if (V08_2 + V08_3) > V08_1 then
  begin
    PGIInfo('La somme des montants imputés ne peut pas être suppérieur au montant antérieur');
    LaGrille.Col := 2;
    LaGrille.Row := Index;
    LaGrille.SetFocus;
    exit;
  end;

  V08_4 := V08_1 - V08_2 - V08_3;
  LaGrille.Cells[1, Index] := Copy(LaGrille.Cells[1, Index], 1, 35);
  LaGrille.Cells[2, Index] := STRFMONTANT(V08_1, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Index] := STRFMONTANT(V08_2, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Index] := STRFMONTANT(V08_3, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[5, Index] := STRFMONTANT(V08_4, 15, V_PGI.OkDecV, '', True);

  Tot081 := 0;
  Tot082 := 0;
  Tot083 := 0;
  Tot084 := 0;

  for i := 0 to LaGrille.RowCount - 2 do
  begin
    Tot081 := Tot081 + VALEUR(LaGrille.Cells[2, i]);
    Tot082 := Tot082 + VALEUR(LaGrille.Cells[3, i]);
    Tot083 := Tot083 + VALEUR(LaGrille.Cells[4, i]);
    Tot084 := Tot084 + VALEUR(LaGrille.Cells[5, i]);
  end;
  LaGrille.Cells[2, Lagrille.RowCount - 1] := STRFMONTANT(Tot081, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Lagrille.RowCount - 1] := STRFMONTANT(Tot082, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Lagrille.RowCount - 1] := STRFMONTANT(Tot083, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[5, Lagrille.RowCount - 1] := STRFMONTANT(Tot084, 15, V_PGI.OkDecV, '', True);

  Tot086 := Tot084 + V08_5;
  LEdit2.Text := STRFMONTANT(Tot086, 15, V_PGI.OkDecV, '', True);
end;

procedure TDRFToolBox.B2058AB_AH_MOA_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if csDestroying in ComponentState then Exit;
  B2058AB_AH_MOA_CalculLigne(ARow);
end;

procedure TDRFToolBox.B2058AB_AG_MOA_CalculLigne(Index: integer);
var
  V09_1: extended;
  V09_2: extended;
  V09_3: extended;
  V09_4: extended;
  V09_5: extended;
  i: integer;
  Tot091: Extended;
  Tot092: Extended;
  Tot093: Extended;
  Tot094: Extended;
  Tot095: Extended;

begin
  V09_1 := VALEUR(LaGrille.Cells[2, Index]);
  V09_2 := VALEUR(LaGrille.Cells[3, Index]);
  V09_3 := VALEUR(LaGrille.Cells[4, Index]);
  V09_4 := VALEUR(LaGrille.Cells[5, Index]);

  if (V09_2 + V09_3) > V09_1 then
  begin
    PGIInfo('La somme des imputations est supérieure au montant antérieur');
    LaGrille.Col := 3;
    LaGrille.Row := Index;
    LaGrille.SetFocus;
    Exit;
  end;
  V09_5 := V09_1 - V09_2 + V09_4 - V09_3;
  LaGrille.Cells[1, Index] := Copy(LaGrille.Cells[1, Index], 1, 35);
  LaGrille.Cells[2, Index] := STRFMONTANT(V09_1, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Index] := STRFMONTANT(V09_2, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Index] := STRFMONTANT(V09_3, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[5, Index] := STRFMONTANT(V09_4, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[6, Index] := STRFMONTANT(V09_5, 15, V_PGI.OkDecV, '', True);

  Tot091 := 0;
  Tot092 := 0;
  Tot093 := 0;
  Tot094 := 0;
  Tot095 := 0;

  for i := 0 to LaGrille.RowCount - 2 do
  begin
    Tot091 := Tot091 + VALEUR(LaGrille.Cells[2, i]);
    Tot092 := Tot092 + VALEUR(LaGrille.Cells[3, i]);
    Tot093 := Tot093 + VALEUR(LaGrille.Cells[4, i]);
    Tot094 := Tot094 + VALEUR(LaGrille.Cells[5, i]);
    Tot095 := Tot095 + VALEUR(LaGrille.Cells[5, i]);
  end;

  LaGrille.Cells[2, Lagrille.RowCount - 1] := STRFMONTANT(Tot091, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Lagrille.RowCount - 1] := STRFMONTANT(Tot092, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Lagrille.RowCount - 1] := STRFMONTANT(Tot093, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[5, Lagrille.RowCount - 1] := STRFMONTANT(Tot094, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[6, Lagrille.RowCount - 1] := STRFMONTANT(Tot095, 15, V_PGI.OkDecV, '', True);


end;

procedure TDRFToolBox.B2058AB_AG_MOA_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if csDestroying in ComponentState then Exit;
  B2058AB_AG_MOA_CalculLigne(ARow);
end;




procedure TDRFToolBox.BJetonPresence_LaCocheOnClick(Sender: Tobject);
begin
  if csDestroying in ComponentState then Exit;
  if LaCoche.Checked then
  begin
    LaGrille.Cells[2, 0] := '100';
    LaGrille.Cells[3, 0] := STRFMONTANT(BaseJeton, 15, V_PGI.OkDecV, '', True);
  end
  else
  begin
    LaGrille.Cells[2, 0] := '5';
    LaGrille.Cells[3, 0] := STRFMONTANT(0, 15, V_PGI.OkDecV, '', True);
  end;
  BJetonPresence_CalculLigne;

end;

procedure TDRFToolBox.BJetonPresence_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if csDestroying in ComponentState then Exit;
  BJetonPresence_CalculLigne;
end;

procedure TDRFToolBox.BJetonPresence_CalculLigne;
var
  Tob653: Tob;
  StrSql: string;
  Val653: Extended;
  Val07_3: Extended;
  Val07_4: Extended;
  Val07_5: Extended;
  TotJeton: Extended;

begin

  if VALEUR(LaGrille.Cells[1, 0]) = 0 then
  begin
    StrSql := 'SELECT G_GENERAL, G_TOTALDEBIT, G_TOTALCREDIT, G_TOTALDEBIT-G_TOTALCREDIT RES, "X" RECH  ' +
      'FROM GENERAUX WHERE G_GENERAL LIKE "653%"';
    Tob653 := Tob.Create('CG 653', nil, -1);
    Tob653.LoadDetailFromSQL(StrSql);
    Val653 := Tob653.Somme('RES', ['RECH'], ['X'], True);
    Tob653.Free;
    LaGrille.Cells[1, 0] := STRFMONTANT(Val653, 15, V_PGI.OkDecV, '', True);
  end
  else
    Val653 := VALEUR(LaGrille.Cells[1, 0]);

  Val07_3 := VALEUR(LaGrille.Cells[2, 0]);
  Val07_4 := VALEUR(LaGrille.Cells[3, 0]);
  Val07_5 := VALEUR(LaGrille.Cells[4, 0]);
  TotJeton := Val653 - (Val07_3 / 100 * Val07_4 * Val07_5);

  LaGrille.Cells[2, 0] := STRFMONTANT(Val07_3, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, 0] := STRFMONTANT(Val07_4, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, 0] := STRFMONTANT(Val07_5, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[5, 0] := STRFMONTANT(TotJeton, 15, V_PGI.OkDecV, '', True);

end;




procedure TDRFToolBox.CessionImmo_CalculLigne(Index: integer);
var
  ValTot: extended;
  ValCT: extended;
  ValLT: extended;
  TotG: extended;
  TotCT: extended;
  TotLT: extended;
  i: integer;

begin
  ValTot := VALEUR(LaGrille.cells[2, Index]);
  ValLT := VALEUR(LaGrille.Cells[3, Index]);
  ValCT := ValTot - ValLT;

  LaGrille.Cells[2, Index] := STRFMONTANT(ValTot, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Index] := STRFMONTANT(ValLT, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Index] := STRFMONTANT(ValCT, 15, V_PGI.OkDecV, '', True);

  TotG := 0;
  TotCT := 0;
  TotLT := 0;

  for i := 0 to LaGrille.RowCount - 2 do
  begin
    TotG := TotG + VALEUR(LaGrille.Cells[2, i]);
    TotLT := TotLT + VALEUR(LaGrille.Cells[3, i]);
    TotCT := TotG - TotLT;
  end;

  i := LaGrille.RowCount - 1;
  LaGrille.Cells[2, i] := STRFMONTANT(TotG, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, i] := STRFMONTANT(TotLT, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, i] := STRFMONTANT(TotCT, 15, V_PGI.OkDecV, '', True);


end;

procedure TDRFToolBox.CessionImmo_OnCellexit(Sender: Tobject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if csDestroying in ComponentState then Exit;
  CessionImmo_CalculLigne(ARow);
end;

{$IFDEF AMORTISSEMENT}
{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 12/04/2007
Modifié le ... :   /  /
Description .. : Retourne dans une tob les immo avec les valeurs à court et
Suite ........ : long termes
Mots clefs ... : IMMO, PMV
*****************************************************************}

procedure TDRFToolBox.ChargeImo(TobV: TOB);
var
  StrSql: string;
//  TobV                  : Tob;
  TobV2: Tob;
  Q: TQuery;
  PMV: TImPMValue;

begin
  StrSql := 'SELECT I_LIBELLE,I_OPECESSION,I_DATEPIECEA, I_DATEAMORT, I_METHODEECO,I_REGLECESSION, ' +
    'IL_PVALUE,IL_DATEOP,IL_CUMANTCESECO,IL_DOTCESSECO,IL_MONTANTCES ' +
    'FROM IMMOLOG LEFT JOIN IMMO ON IL_IMMO=I_IMMO WHERE IL_TYPEOP LIKE "CE%" ' +
    'AND IL_DATEOP>="' + USDateTime(VH^.Encours.Deb) + '"' +
    'AND IL_DATEOP<="' + USDatetime(VH^.Encours.Fin) + '"';

  Q := OpenSQL(StrSql, True);
  while not Q.Eof do
  begin
    PMV := CalculPMValue(Q.FindField('IL_DATEOP').AsDateTime,
      Q.FindField('I_DATEPIECEA').AsDateTime, Q.FindField('I_DATEAMORT').AsDateTime,
      Q.FindField('IL_PVALUE').AsFloat, Q.FindField('IL_CUMANTCESECO').AsFloat,
      Q.FindField('IL_DOTCESSECO').AsFloat, Q.FindField('I_METHODEECO').AsString,
      Q.FindField('I_REGLECESSION').AsString);

    TobV2 := Tob.Create('Data Immo', TobV, -1);
    TobV2.AddChampSupValeur('LIBELLE', Q.FindField('I_LIBELLE').AsString);
    TobV2.AddChampSupValeur('MONTANTCES', Q.FindField('IL_MONTANTCES').AsFloat);
    TobV2.AddChampSupValeur('PCT', PMV.PCT);
    TobV2.AddChampSupValeur('MCT', PMV.MCT);
    TobV2.AddChampSupValeur('PLT', PMV.PLT);
    TobV2.AddChampSupValeur('MLT', PMV.MLT);
    Q.Next;
  end;
  Ferme(Q);
end;
{$ENDIF}

procedure TDRFToolBox.B2058AB_AF_MOA_CalculLigne(Index: integer);
var
  V09_8: extended;
  V09_9: extended;
  V09_10: extended;
  V09_11: extended;
  V09_12: extended;
  V09_13: extended;
  i: integer;
  Tot098: Extended;
  Tot099: Extended;
  Tot0910: Extended;
  Tot0911: Extended;
  Tot0912: Extended;
  Tot0913: Extended;
  PC8: Extended;
  PC9: Extended;
  PC10: Extended;
  Ecart8: extended;
  Ecart9: extended;
  Ecart10: extended;

begin
  V09_8 := VALEUR(LaGrille.Cells[2, Index]);
  V09_9 := VALEUR(LaGrille.Cells[3, Index]);
  V09_10 := VALEUR(LaGrille.Cells[4, Index]);
  V09_11 := VALEUR(LaGrille.Cells[5, Index]);
  V09_12 := VALEUR(LaGrille.Cells[6, Index]);
  V09_13 := V09_8 - (V09_9 + V09_10 + V09_11 + V09_12);

  LaGrille.Cells[1, Index] := Copy(LaGrille.Cells[1, Index], 1, 35);
  LaGrille.Cells[2, Index] := STRFMONTANT(V09_8, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Index] := STRFMONTANT(V09_9, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Index] := STRFMONTANT(V09_10, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[5, Index] := STRFMONTANT(V09_11, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[6, Index] := STRFMONTANT(V09_12, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[7, Index] := STRFMONTANT(V09_13, 15, V_PGI.OkDecV, '', True);

  Tot098 := 0;
  Tot099 := 0;
  Tot0910 := 0;
  Tot0911 := 0;
  Tot0912 := 0;
  Tot0913 := 0;

  for i := 0 to LaGrille.RowCount - 4 do
  begin
    Tot098 := Tot098 + VALEUR(LaGrille.Cells[2, i]);
    Tot099 := Tot099 + VALEUR(LaGrille.Cells[3, i]);
    Tot0910 := Tot0910 + VALEUR(LaGrille.Cells[4, i]);
    Tot0911 := Tot0911 + VALEUR(LaGrille.Cells[5, i]);
    Tot0912 := Tot0912 + VALEUR(LaGrille.Cells[6, i]);
    Tot0913 := Tot0913 + VALEUR(LaGrille.Cells[7, i]);
  end;

  LaGrille.Cells[2, Lagrille.RowCount - 3] := STRFMONTANT(Tot098, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Lagrille.RowCount - 3] := STRFMONTANT(Tot099, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Lagrille.RowCount - 3] := STRFMONTANT(Tot0910, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[5, Lagrille.RowCount - 3] := STRFMONTANT(Tot0911, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[6, Lagrille.RowCount - 3] := STRFMONTANT(Tot0912, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[7, Lagrille.RowCount - 3] := STRFMONTANT(Tot0913, 15, V_PGI.OkDecV, '', True);

  PC8 := VALEUR(LaGrille.Cells[2, Lagrille.RowCount - 2]);
  PC9 := VALEUR(LaGrille.Cells[3, Lagrille.RowCount - 2]);
  PC10 := VALEUR(LaGrille.Cells[4, Lagrille.RowCount - 2]);
  Ecart8 := Tot098 - PC8;
  Ecart9 := Tot099 - PC9;
  Ecart10 := Tot0910 - PC10;
  LaGrille.Cells[2, Lagrille.RowCount - 1] := STRFMONTANT(Ecart8, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[3, Lagrille.RowCount - 1] := STRFMONTANT(Ecart9, 15, V_PGI.OkDecV, '', True);
  LaGrille.Cells[4, Lagrille.RowCount - 1] := STRFMONTANT(Ecart10, 15, V_PGI.OkDecV, '', True);


end;

procedure TDRFToolBox.B2058AB_AF_MOA_OnCellexit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if csDestroying in ComponentState then Exit;
  B2058AB_AF_MOA_CalculLigne(ARow);
end;





{ TzDRF }

{***********A.G.L.*************************************************
Auteur  ...... : TJ
Créé le ...... : 08/02/2007
Modifié le ... :   /  /
Description .. : Création de l'objet ZDrf
Suite ........ : paramètres :
Suite ........ :   - Formulaire (ex 2058A)
Suite ........ :   - Exercice (ex 006)
Mots clefs ... :
*****************************************************************}

constructor TzDRF.Create(Formulaire, Exercice: string);
var
  TobR: Tob;
  TobA: Tob;
  TobGene: Tob;
  i: integer;
begin
  LeFormulaire := Formulaire;
  Lexo := Exercice;
  TobZDRF := Tob.Create('La tob zdrf', nil, -1);
  TobA := nil;
  TOBGene := TOB.Create('DATA_DRF', nil, -1);
  try
    ChargeTobDRF(TobGene, Lexo);
    TobGene.detail.Sort('NATURE;INDICE');

    if Assigned(TobGene) then
    begin
      if LeFormulaire = '2058A' then
      begin
        for i := 0 to TobGene.detail.count - 1 do
        begin
          TobR := TobGene.detail[i];
          if (TobR.GetString('CODE1') = '2058A') or (Pos('2058A_', TobR.GetString('RESULTAT')) > 0) then
          begin
            TobA := Tob.Create('La data', TobZDRF, -1);
            TobA.AddChampSupValeur('NATURE', TobR.GetValue('NATURE'));
            TobA.AddChampSupValeur('INDICE', TobR.GetValue('INDICE'));
            TobA.AddChampSupValeur('CODE1', TobR.GetValue('CODE1'));
            TobA.AddChampSupValeur('CODE2', TobR.GetValue('CODE2'));
            TobA.AddChampSupValeur('COMMENTAIRE', TobR.GetValue('COMMENTAIRE'));
            TobA.AddChampSupValeur('MONTANT', VALEUR(TobR.GetValue('MONTANT')));
            TobA.AddChampSupValeur('MT2', VALEUR(TobR.GetValue('MT2')));
            TobA.AddChampSupValeur('MT3', VALEUR(TobR.GetValue('MT3')));
            TobA.AddChampSupValeur('MT4', VALEUR(TobR.GetValue('MT4')));
            TobA.AddChampSupValeur('MT5', VALEUR(TobR.GetValue('MT5')));
            TobA.AddChampSupValeur('RESULTAT', TobR.GetValue('RESULTAT'));
          end;
        end;
      end
      else
      begin
        TobR := TobGene.FindFirst(['CODE1'], [LeFormulaire], False);
        while Assigned(TobR) do
        begin
          TobA := Tob.Create('La data', TobZDRF, -1);
          TobA.AddChampSupValeur('NATURE', TobR.GetValue('NATURE'));
          TobA.AddChampSupValeur('INDICE', TobR.GetValue('INDICE'));
          TobA.AddChampSupValeur('CODE1', TobR.GetValue('CODE1'));
          TobA.AddChampSupValeur('CODE2', TobR.GetValue('CODE2'));
          TobA.AddChampSupValeur('COMMENTAIRE', TobR.GetValue('COMMENTAIRE'));
          TobA.AddChampSupValeur('MONTANT', VALEUR(TobR.GetValue('MONTANT')));
          TobA.AddChampSupValeur('MT2', VALEUR(TobR.GetValue('MT2')));
          TobA.AddChampSupValeur('MT3', VALEUR(TobR.GetValue('MT3')));
          TobA.AddChampSupValeur('MT4', VALEUR(TobR.GetValue('MT4')));
          TobA.AddChampSupValeur('MT5', VALEUR(TobR.GetValue('MT5')));
          TobA.AddChampSupValeur('RESULTAT', TobR.GetValue('RESULTAT'));
          TobR := TobGene.FindNext(['CODE1'], [LeFormulaire], False);
        end;
      end;
    end;
  finally
    TobGene.Free;
  end;
end;


destructor TzDRF.Destroy;
begin

  inherited;
  TobZDRF.Free;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/02/2007
Modifié le ... :   /  /
Description .. : Retourne une Tob avec le détail d'une nature
Mots clefs ... :
*****************************************************************}

function TzDRF.GetDetail(LeChamp: string): Tob;
var
  TobD: Tob;
  TobR: Tob;
  TobA: Tob;

begin
  Result := nil;
  TobD := Tob.Create('Les details', nil, -1);
  if LeChamp = 'WQ_MOA' then
    TobR := TobZDRF.FindFirst(['RESULTAT'], ['P2058A_WQ_MOA'], False)
  else if LeChamp = 'XG_MOA' then
    TobR := TobZDRF.FindFirst(['RESULTAT'], ['P2058A_XG_MOA'], False)
  else
    TobR := TobZDRF.FindFirst(['CODE2'], [LeChamp], False);
  while Assigned(TobR) do
  begin
    if TobR.GetValue('INDICE') <> 0 then
    begin
      TobA := Tob.Create('Le detail', TobD, -1);
      TobA.AddChampSupValeur('NUM', TobR.GetValue('INDICE'));
      TobA.AddChampSupValeur('COMMENTAIRE', TobR.GetValue('COMMENTAIRE'));
      TobA.AddChampSupValeur('MONTANT', TobR.GetValue('MONTANT'));
      TobA.AddChampSupValeur('MT2', TobR.GetValue('MT2'));
      TobA.AddChampSupValeur('MT3', TobR.GetValue('MT3'));
      TobA.AddChampSupValeur('MT4', TobR.GetValue('MT4'));
      TobA.AddChampSupValeur('MT5', TobR.GetValue('MT5'));
      TobA.AddChampSupValeur('RESULTAT', TobR.GetValue('RESULTAT'));
    end;
    if LeChamp = 'WQ_MOA' then
      TobR := TobZDRF.FindNext(['RESULTAT'], ['P2058A_WQ_MOA'], False)
    else if LeChamp = 'XG_MOA' then
      TobR := TobZDRF.FindNext(['RESULTAT'], ['P2058A_XG_MOA'], False)
    else
      TobR := TobZDRF.FindNext(['CODE2'], [LeChamp], False);
  end;

  if TobD.Detail.Count {- 1} > 0 then //il peut y avoir une seule ligne de détail
    Result := TobD;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/02/2007
Modifié le ... :   /  /
Description .. : Retourne la valeur pour une nature
Mots clefs ... :
*****************************************************************}

function TzDRF.GetValue(LeChamp: string): Variant;
var
  LaValeur: extended;
  TobC: Tob;

begin
  Result := 0;
  TobC := TobZDRF.FindFirst(['CODE2', 'INDICE'], [LeChamp, 0], False);
  if Assigned(TobC) then
  begin
    LaValeur := TobC.GetValue('MONTANT');
    Result := LaValeur;
  end;
end;



end.

