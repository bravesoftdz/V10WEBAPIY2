unit uTOFTradTablette;

interface

uses StdCtrls, Controls, Classes, db, forms, sysutils, dbTables, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOF, UTOB;

type
  TOF_TradTablette = class(TOF)
  private
    FTobTablette : TOB; // Contient la tablette des tablettes traductibles (avec CO_LIBRE)
    FCurrentData : TOB; // Données de la tablette choisie
    FCurrentTrad : TOB; // Séléction de TradTablette pour la tablette choisie
    FDisplayData : TOB; // Données destinées au grid
    FLibCol, FAbrCol : integer;
    FGrid : THGrid;
    FDimension : boolean;
    FCurrentLangue : string;
    FModified : boolean;

    procedure LoadCurrentData(Code : String);
    procedure SaveCurrentTrad;
    function AddDisplayDataChild : TOB;
    procedure PutDisplayData;
    procedure PutDisplayTrad(Code : String);
    procedure GetDisplayTrad(Code : String);
    procedure DisplayGrid;
    procedure AnalyseColumns;

  public
    procedure OnNew; override;      // Bouton nouveau
    procedure OnDelete; override;   // Bouton detruire
    procedure OnUpdate; override;   // Mouette verte
    procedure OnLoad; override;     // Fin de FormShow
    procedure OnArgument(S : String); override;  // Début de FormShow
    procedure OnClose; override;

    procedure OnTabletteChange(Sender : TObject);
    procedure OnLangueChange(Sender : TObject);
    procedure OnGridCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure OnGridCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);

  end;

implementation
uses Grids;

procedure TOF_TradTablette.OnNew;
begin
  inherited;
end;

procedure TOF_TradTablette.OnDelete;
begin
  inherited;
end;

procedure TOF_TradTablette.OnUpdate;
begin
  GetDisplayTrad(FCurrentLangue);
  SaveCurrentTrad;
  inherited;
end;

procedure TOF_TradTablette.OnLoad;
begin
  inherited;
end;

procedure TOF_TradTablette.OnArgument(S : String);
var Q : TQuery;
begin
  FCurrentData := TOB.Create('', nil, -1);
  FCurrentTrad := TOB.Create('', nil, -1);
  FDisplayData := TOB.Create('', nil, -1);
  FTobTablette := TOB.Create('COMMUN', nil, -1);

  Q := OpenSQL('SELECT CO_CODE, CO_ABREGE, CO_LIBRE FROM COMMUN WHERE CO_TYPE="TRT"', true);
  FTobTablette.LoadDetailDB('COMMUN', '', '', Q, false);
  Ferme(Q);

  THValComboBox(Ecran.FindComponent('CBTABLETTE')).OnChange := OnTabletteChange;
  THValComboBox(Ecran.FindComponent('CBLANGUE')).OnChange := OnLangueChange;
  FGrid := THGrid(Ecran.FindComponent('GRID'));
  FGrid.OnCellEnter := OnGridCellEnter;
  FGrid.OnCellExit := OnGridCellExit;

  DisplayGrid;
  FGrid.Titres.Add('(TYPE);(CODE);(LIB);(LIBTRAD);(ABR);(ABRTRAD)');

  AnalyseColumns;

  inherited;
end;

procedure TOF_TradTablette.OnClose;
begin
  FTobTablette.Free;
  FCurrentData.Free;
  FCurrentTrad.Free;
  FDisplayData.Free;
  inherited;
end;


procedure TOF_TradTablette.OnTabletteChange(Sender : TObject);
begin
  if FModified and (PGIAsk('La saisie a été modifiée, voulez-vous l''enregistrer ?', Ecran.Caption) = mrYes)
    then OnUpdate; // Mouette verte

  LoadCurrentData(THValComboBox(Sender).Value);
  PutDisplayData;
  PutDisplayTrad(THValComboBox(Ecran.FindComponent('CBLANGUE')).Value);
  DisplayGrid;
end;

procedure TOF_TradTablette.OnLangueChange(Sender : TObject);
begin
  GetDisplayTrad(FCurrentLangue);

  FCurrentLangue := THValComboBox(Sender).Value;
  PutDisplayTrad(FCurrentLangue);
  DisplayGrid;
end;

procedure TOF_TradTablette.OnGridCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if (FGrid.Objects[0, FGrid.Row] <> nil) and ((FGrid.Col = FLibCol) or ((not FDimension) and (FGrid.Col = FAbrCol)))
    then FGrid.Options := FGrid.Options + [goEditing, goAlwaysShowEditor]
    else begin FGrid.EditorMode := false;
               FGrid.Options := FGrid.Options - [goEditing, goAlwaysShowEditor]; end;
end;

procedure TOF_TradTablette.OnGridCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var Old, Champ : String;
begin
  if csDestroying in FGrid.ComponentState then exit;
  if (FGrid.Objects[0, ARow] <> nil) and ((ACol = FLibCol) or ((not FDimension) and (ACol = FAbrCol))) then
  begin
    if ACol = FLibCol then Champ := '(LIBTRAD)'
                      else Champ := '(ABRTRAD)';
    Old := TOB(FGrid.Objects[0, ARow]).GetValue(Champ);
    TOB(FGrid.Objects[0, ARow]).PutValue(Champ, FGrid.Cells[ACol, ARow]);
    if Old <> FGrid.Cells[ACol, ARow] then FModified := true;
  end;
end;


procedure TOF_TradTablette.LoadCurrentData(Code : String);
var Pfx, Dim, S, Rq : String;
    T : TOB;
    Q : TQuery;
begin
  T := FTobTablette.FindFirst(['CO_CODE'], [Code], false);
  if T = nil then exit;

  Pfx := T.GetValue('CO_ABREGE');

  FDimension := (Pfx = 'GDI');
  if FDimension then Dim := 'DIM'
                else Dim := '';

  Rq := 'SELECT * FROM ' + PrefixeToTable(Pfx);
  S := T.GetValue('CO_LIBRE');
  if S <> '' then
  begin
    Rq := Rq + ' WHERE ';
    repeat
      Rq := Rq + Pfx+'_TYPE'+Dim+'="'+ReadTokenSt(S)+'"';
      if S <> '' then Rq := Rq + ' OR ';
    until S = '';
  end;

  Q := OpenSQL(Rq, true);
  FCurrentData.LoadDetailDB(PrefixeToTable(Pfx), '', '', Q, false);
  Ferme(Q);


  Rq := 'SELECT * FROM TRADTABLETTE';
  S := T.GetValue('CO_LIBRE');
  if S <> '' then
  begin
    Rq := Rq + ' WHERE ';
    repeat
      Rq := Rq + 'YTT_TYPE="'+ReadTokenSt(S)+'"';
      if S <> '' then Rq := Rq + ' OR ';
    until S = '';
  end;

  Q := OpenSQL(Rq, true);
  FCurrentTrad.LoadDetailDB('TRADTABLETTE', '', '', Q, false);
  Ferme(Q);

  FModified := false;
end;

procedure TOF_TradTablette.SaveCurrentTrad;
begin
  FCurrentTrad.InsertOrUpdateDB(true);
  FModified := false;
end;

function TOF_TradTablette.AddDisplayDataChild : TOB;
begin
  result := TOB.Create('', FDisplayData, -1);
  result.AddChampSup('(TYPE)', false);
  result.AddChampSup('(CODE)', false);
  result.AddChampSup('(LIB)', false);
  result.AddChampSup('(LIBTRAD)', false);
  result.AddChampSup('(ABR)', false);
  result.AddChampSup('(ABRTRAD)', false);
end;

procedure TOF_TradTablette.PutDisplayData;
var i : integer;
    Pfx : String;
    T : TOB;
begin
  FDisplayData.ClearDetail;
  for i := 0 to FCurrentData.Detail.Count-1 do
   with AddDisplayDataChild do
   begin
     T := FCurrentData.Detail[i];
     Pfx := TableToPrefixe(T.NomTable);
     if FDimension then
     begin
       PutValue('(TYPE)', T.GetValue('GDI_TYPEDIM'));
       PutValue('(CODE)', T.GetValue('GDI_GRILLEDIM') +';'+ T.GetValue('GDI_CODEDIM'));
     end else
     begin
       PutValue('(TYPE)', T.GetValue(Pfx+'_TYPE'));
       PutValue('(CODE)', T.GetValue(Pfx+'_CODE'));
       PutValue('(ABR)', T.GetValue(Pfx+'_ABREGE'));
     end;
     PutValue('(LIB)', T.GetValue(Pfx+'_LIBELLE'));
   end;
end;

procedure TOF_TradTablette.PutDisplayTrad(Code : String);
var i : integer;
    T : TOB;
begin
  for i := 0 to FDisplayData.Detail.Count-1 do
   with FDisplayData.Detail[i] do
   begin
     T := FCurrentTrad.FindFirst(['YTT_LANGUE', 'YTT_TYPE', 'YTT_CODE'],
                                 [Code, GetValue('(TYPE)'), GetValue('(CODE)')], false);
     if T <> nil then
     begin
       PutValue('(LIBTRAD)', T.GetValue('YTT_LIBELLE'));
       PutValue('(ABRTRAD)', T.GetValue('YTT_ABREGE'));
     end else
     begin
       PutValue('(LIBTRAD)', '');
       PutValue('(ABRTRAD)', '');
     end;
   end;
end;

procedure TOF_TradTablette.GetDisplayTrad(Code : String);
var i : integer;
    T : TOB;
begin
  if Code = '' then exit;

  for i := 0 to FDisplayData.Detail.Count-1 do
   with FDisplayData.Detail[i] do
   begin
     if (GetValue('(LIBTRAD)') = '') and (GetValue('(ABRTRAD)') = '') then Continue;

     T := FCurrentTrad.FindFirst(['YTT_LANGUE', 'YTT_TYPE', 'YTT_CODE'],
                                 [Code, GetValue('(TYPE)'), GetValue('(CODE)')], false);
     if T = nil then
     begin
       T := TOB.Create('TRADTABLETTE', FCurrentTrad, -1);
       T.PutValue('YTT_LANGUE', Code);
       T.PutValue('YTT_TYPE', GetValue('(TYPE)'));
       T.PutValue('YTT_CODE', GetValue('(CODE)'));
     end;

     T.PutValue('YTT_LIBELLE', GetValue('(LIBTRAD)'));
     T.PutValue('YTT_ABREGE', GetValue('(ABRTRAD)'));
   end;
end;

procedure TOF_TradTablette.DisplayGrid;
var C, R : integer;
    Can : boolean;
begin
  FDisplayData.PutGridDetail(FGrid, false, false, '', true);
  OnGridCellEnter(FGrid, C, R, Can);
end;


procedure TOF_TradTablette.AnalyseColumns;
var NomCol, LesCols : String;
    col : integer;
begin
  LesCols := FGrid.Titres[0];

  FLibCol := -1;
  FAbrCol := -1;

  for col := 0 to FGrid.ColCount-1 do
  begin
    NomCol := Uppercase(Trim(ReadTokenSt(LesCols)));
    if NomCol = '(LIBTRAD)' then FLibCol := col;
    if NomCol = '(ABRTRAD)' then FAbrCol := col;
  end;
end;


initialization
  RegisterClasses([TOF_TradTablette]);
end.

