{***********UNITE*************************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : Spécifique mode: Tarif détail
Suite ........ : Gestion des types de tarif
Mots clefs ... : TARIF;TYPE
*****************************************************************}
unit UTofMBOTarifType;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls
  , HCtrls, HEnt1, HMsgBox, UTOF, vierge, UTOB, AglInit, LookUp, EntGC, SaisUtil, graphics, paramsoc,
  {$IFDEF EAGLCLIENT}
  Maineagl,
  {$ELSE}
  dbTables, Fe_Main,
  {$ENDIF}
  grids, windows, TarifUtil, M3FP, HDB;

type
  TOF_MBOTarifType = class(TOF)
  private
    GridTyp: THGRID;
    TobTyp: Tob;
    NewLigne: Boolean;
    NatureType: string;
    Col_Mov, colNature, colCode, colLibelle, colDevise, colCoef, colExemple: Integer;
    LesColonnes, EtablisRef: string;
    TabEtabMem: array[1..99] of string;

    procedure GSElipsisClick(Sender: TObject);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSEnter(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ChargePeriode;
    procedure CalculExemple;
    procedure FormateColSaisie(ACol, ARow: Longint);

    function VerifCombo(Champ, Code: string): boolean;
    function VerifEnreg: Boolean;

    function TypeUtiliseDansTarif: boolean;

    function EstRempli(Lig: integer): Boolean;
    procedure ColTriangle(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure CodeGras(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);

  public
    Action: TActionFiche;
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
    procedure OnUpdate; override;
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnCancel; override;
    procedure OnClose; override;

  end;

const MaxType = 99;

implementation

procedure TOF_MBOTarifType.OnArgument(Arguments: string);
var St, S, NomCol, ChampMul, ValMul, Critere, StAction: string;
  i: Integer;
begin
  inherited;
  St := Arguments;
  Action := taModif;
  repeat
    Critere := Trim(ReadTokenSt(Arguments));
    if Critere <> '' then
    begin
      i := pos('=', Critere);
      if i <> 0 then
      begin
        ChampMul := copy(Critere, 1, i - 1);
        ValMul := copy(Critere, i + 1, length(Critere));
        if ChampMul = 'ACTION' then StAction := ValMul;
        if ChampMul = 'TYPE' then NatureType := ValMul;
      end;
    end;
  until Critere = '';
  if StAction = 'CREATION' then Action := taCreat
  else if StAction = 'MODIFICATION' then Action := taModif
  else if StAction = 'CONSULTATION' then Action := taConsult;

  LesColonnes := 'MOV;GFT_NATURETYPE;GFT_CODETYPE;GFT_LIBELLE;GFT_DEVISE;GFT_COEF;EXEMPLE';
  GRIDTYP := THGRID(GetControl('GRIDTYP'));
  GRIDTYP.OnElipsisClick := GSElipsisClick;
  GRIDTYP.OnDblClick := GSElipsisClick;
  GRIDTYP.OnCellEnter := GSCellEnter;
  GRIDTYP.OnCellExit := GSCellExit;
  GRIDTYP.OnRowExit := GSRowExit;
  GRIDTYP.OnRowEnter := GSRowEnter;
  GRIDTYP.OnEnter := GSEnter;
  GRIDTYP.GetCellCanvas := CodeGras;
  GRIDTYP.PostDrawCell := ColTriangle;
  GRIDTYP.ColCount := 1;
  i := 0;
  S := LesColonnes;
  Col_Mov := -1;
  colNature := -1;
  colCode := -1;
  colLibelle := -1;
  colDevise := -1;
  colCoef := -1;
  colExemple := -1;
  repeat
    NomCol := ReadTokenSt(S);
    if NomCol <> '' then
    begin
      if NomCol = 'MOV' then
      begin
        if i <> 0 then GRIDTYP.ColCount := GRIDTYP.ColCount + 1;
        Col_Mov := i;
        GRIDTYP.ColWidths[Col_Mov] := 10;
      end
      else if NomCol = 'GFT_NATURETYPE' then
      begin
        if i <> 0 then GRIDTYP.ColCount := GRIDTYP.ColCount + 1;
        colNature := i;
      end
      else if NomCol = 'GFT_CODETYPE' then
      begin
        if i <> 0 then GRIDTYP.ColCount := GRIDTYP.ColCount + 1;
        colCode := i;
        GRIDTYP.ColWidths[colCode] := 70;
        GRIDTYP.ColLengths[colCode] := 6;
        GRIDTYP.ColFormats[colCode] := 'UPPER';
        GRIDTYP.ColAligns[colCode] := taCenter;
      end
      else if NomCol = 'GFT_LIBELLE' then
      begin
        if i <> 0 then GRIDTYP.ColCount := GRIDTYP.ColCount + 1;
        colLibelle := i;
        GRIDTYP.ColWidths[colLibelle] := 110;
      end
      else if NomCol = 'GFT_DEVISE' then
      begin
        if i <> 0 then GRIDTYP.ColCount := GRIDTYP.ColCount + 1;
        colDevise := i;
        GRIDTYP.ColFormats[colDevise] := 'UPPER';
        GRIDTYP.ColWidths[colDevise] := 45;
        GRIDTYP.ColAligns[colDevise] := taCenter;
      end
      else if NomCol = 'GFT_COEF' then
      begin
        if i <> 0 then GRIDTYP.ColCount := GRIDTYP.ColCount + 1;
        colCoef := i;
        GRIDTYP.ColTypes[colCoef] := 'R';
        GRIDTYP.ColFormats[colCoef] := '#0.#######';
        GRIDTYP.ColWidths[colCoef] := 45;
        GRIDTYP.ColAligns[colCoef] := taRightJustify;
      end
      else if NomCol = 'EXEMPLE' then
      begin
        if i <> 0 then GRIDTYP.ColCount := GRIDTYP.ColCount + 1;
        colExemple := i;
        GRIDTYP.ColWidths[colExemple] := 70;
        GRIDTYP.ColLengths[colExemple] := -1;
      end;
      Inc(i);
    end;
  until ((St = '') or (NomCol = ''));
  if Col_Mov <> -1 then GRIDTYP.FixedCols := 1;
  GridTyp.Cells[colExemple, 0] := 'Exemple';
  AffecteGrid(GRIDTYP, Action);
  GRIDTYP.ColWidths[colNature] := 0;
  TFVierge(Ecran).Hmtrad.ResizeGridColumns(GRIDTYP);
  GRIDTYP.Col := ColLibelle;
  TFVierge(Ecran).OnKeyDown := FormKeyDown;
end;

procedure TOF_MBOTarifType.OnLoad;
begin
  inherited;
  if NatureType = 'VTE' then TFVierge(Ecran).Caption := 'Type de tarifs de vente' else TFVierge(Ecran).Caption := 'Type de tarifs d''achat';
  UpdateCaption(Ecran);
  ChargePeriode;
  // Permet d'afficher en gras le code de la première ligne du Grid
  GRIDTYP.SetFocus;
end;

procedure TOF_MBOTarifType.OnUpdate;
begin
  inherited;
  SetControlEnabled('BInsert', True);
  SetControlEnabled('BDelete', True);
  SetControlVisible('ETAB', False);
  if NewLigne then GridTyp.Cells[ColNature, GridTyp.rowcount - 1] := NatureType;
  TobTyp.GetGridDetail(GRIDTYP, GridTyp.rowcount - 1, 'TARIFTYPMODE', LesColonnes);
  if not VerifEnreg then exit;
  TobTyp.InsertOrUpdateDB(False);
  TobTyp.SetAllModifie(false);
  NewLigne := False;
  CalculExemple;
end;

procedure TOF_MBOTarifType.OnNew;
begin
  inherited;
  NewLigne := True;
  GRIDTYP.InsertRow(GRIDTYP.RowCount);
  GRIDTYP.Row := GRIDTYP.RowCount - 1;
  GRIDTYP.Col := ColCode;
  GRIDTYP.SetFocus;
  GRIDTYP.Refresh;
  SetControlEnabled('BInsert', False);
  SetControlEnabled('BDelete', False);
end;

procedure TOF_MBOTarifType.OnDelete;
begin
  inherited;
  if GRIDTYP.Row <= 0 then Exit;
  if GRIDTYP.RowCount < 2 then Exit;
  if GRIDTYP.RowCount = 2 then GRIDTYP.RowCount := 3;
  if TypeUtiliseDansTarif then
  begin
    PGIInfo('Un tarif de ce type est utilisé,Vous ne pouvez pas le supprimer !', Ecran.Caption);
    exit;
  end else
  begin
    if HShowMessage('1;Type de tarif;Confirmer vous la suppression de l''enregistrement ?;Q;YNC;N;C;', '', '') <> mrYes then exit;
    ExecuteSQL('DELETE FROM TARIFTYPMODE WHERE GFT_CODETYPE="' + GRIDTYP.Cells[ColCode, GRIDTYP.Row] + '" AND GFT_NATURETYPE="' + NatureType + '"');
    GRIDTYP.DeleteRow(GRIDTYP.row);
  end;
end;

procedure TOF_MBOTarifType.OnCancel;
begin
  inherited;
  SetControlEnabled('BInsert', True);
  SetControlEnabled('BDelete', True);
  if GRIDTYP.RowCount - 1 > TobTyp.Detail.count then
    if NewLigne then
    begin
      GRIDTYP.DeleteRow(GridTyp.RowCount);
      NewLigne := False;
    end;
  OnLoad;
end;

procedure TOF_MBOTarifType.OnClose;
var i: Integer;
  TOBTypeSav: TOB;
  Compare: integer;
begin
  inherited;
  Compare := 0;
  TOBTypeSav := TOB.Create('_TARIFTYPMODE', nil, -1);
  TOBTypeSav.Dupliquer(TobTyp, True, True);
  LastError := 0;
  if GridTyp.rowcount > 1 then TobTyp.GetGridDetail(GridTyp, GridTyp.RowCount, 'TARIFTYPMODE', LesColonnes);
  if TOBTypeSav.Detail.Count = TobTyp.Detail.count then
  begin
    for i := 0 to TobTyp.Detail.count - 1 do
    begin
      if TobTyp.Detail.count > 1 then
        Compare := CompareTOB(TobTyp.Detail[i], TOBTypeSav.Detail[i], LesColonnes);
      if Compare <> 0 then break;
    end;
  end else Compare := 10;
  if Compare <> 0 then
  begin
    if PGIAsk('Confirmez-vous l''abandon des modifications?', Ecran.Caption) = mrNo then
    begin
      LastError := -1;
      TOBTypeSav.free;
      Exit;
    end else
    begin
      LastError := 0;
    end;
  end;
  TOBTypeSav.free;
  TobTyp.free;
  TobTyp := nil;
end;

function TOF_MBOTarifType.VerifEnreg: Boolean;
var ARow, i: Integer;
  Coef: Double;
begin
  Result := True;
  for i := 1 to GridTyp.RowCount - 1 do
  begin
    Arow := i;
    if (NewLigne) and (ARow = GridTyp.RowCount - 1) then
    begin
      if Arow > MaxType then exit;
      if VerifCombo('CODETYP', GridTyp.Cells[ColCode, ARow]) then
      begin
        PGIBox('Le type de tarif que vous avez saisi existe déjà. Vous devez le modifier', Ecran.Caption);
        Result := False;
        GridTyp.Col := colCode;
        GridTyp.Row := ARow;
        exit;
      end;
    end;
    Result := True;
    if VerifCombo('DEVISE', GridTyp.Cells[ColDevise, ARow]) then
    begin
      PGIBox('Le code devise n''existe pas. Vous devez le modifier', Ecran.Caption);
      Result := False;
      GridTyp.Col := colDevise;
      GridTyp.Row := ARow;
      exit;
    end;
    if (GRIDTYP.Cells[ColCode, Arow] = '') then
    begin
      PGIBox('Le code est obligatoire', Ecran.Caption);
      GRIDTYP.Col := ColCode;
      GridTyp.Row := ARow;
      Result := False;
      Exit;
    end;
    if (GRIDTYP.Cells[ColLibelle, Arow] = '') then
    begin
      PGIBox('Le libelle est obligatoire', Ecran.Caption);
      GRIDTYP.Col := ColLibelle;
      GridTyp.Row := ARow;
      Result := False;
      Exit;
    end;
    if (GRIDTYP.Cells[ColDevise, Arow] = '') then
    begin
      PGIBox('La devise est obligatoire', Ecran.Caption);
      GRIDTYP.Col := Coldevise;
      GridTyp.Row := ARow;
      Result := False;
      Exit;
    end;
    // Maj table tarif mode
    Coef := StrToFloat(GridTyp.Cells[colCoef, ARow]);
    ExecuteSQL('Update TarifMode SET GFM_COEF="' + StrFPoint(Coef) + '",' +
      'GFM_DEVISE="' + GridTyp.Cells[colDevise, ARow] + '"' +
      'where GFM_TYPETARIF="' + GridTyp.Cells[ColCode, ARow] + '" AND GFM_NATURETYPE="' + NatureType + '"');
  end;
  PGIInfo('Ces modifications seront prises en compte sur les nouveaux articles', TFVierge(Ecran).Caption);
end;

function TOF_MBOTarifType.TypeUtiliseDansTarif: boolean;
var CodeType: string;
  Q: TQuery;
begin
  Result := False;
  CodeType := GRIDTYP.Cells[ColCode, GRIDTYP.Row];
  Q := OpenSQL('Select GFM_TARFMODE from TARIFMODE Where GFM_TYPETARIF="' + CodeType + '" And GFM_NATURETYPE="' + NatureType + '"', True);
  while not Q.EOF do
  begin
    if ExisteSQL('Select GF_Tarif from Tarif where GF_TARFMODE="' + Q.FindField('GFM_TARFMODE').AsString + '"') then
    begin
      Result := True;
      ferme(Q);
      Exit;
    end
    else Q.next;
  end;
  ferme(Q);
end;

function TOF_MBOTarifType.EstRempli(Lig: integer): Boolean;
begin
  Result := False;
  if (GRIDTYP.Cells[ColCode, Lig] <> '') then result := true;
end;

function TOF_MBOTarifType.VerifCombo(Champ, Code: string): boolean;
begin
  Result := True;
  if Champ = 'CODETYP' then
  begin
    Result := ExisteSQL('Select GFT_CODETYPE From TARIFTYPMODE Where GFT_CODETYPE="' + Code + '" AND GFT_NATURETYPE="' + NatureType + '"');
  end else
    if Champ = 'DEVISE' then
  begin
    Result := not ExisteSQL('Select D_DEVISE From DEVISE Where D_DEVISE="' + Code + '"');
  end;
end;

{==============================================================================================}
{=============================== Evenement du Grid ========================================}
{==============================================================================================}

procedure TOF_MBOTarifType.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var Cancel, Chg: Boolean;
begin
  Cancel := False;
  Chg := False;
  case Key of
    VK_F5: if (Screen.ActiveControl = GRIDTYP) then GSElipsisClick(Sender);
    VK_DOWN: if (Screen.ActiveControl = GRIDTYP) then
      begin
        if ((NewLigne) and (GridTyp.Row = GridTyp.RowCount - 1)) and EstRempli(GRIDTYP.Row) then OnUpdate
        else if EstRempli(GRIDTYP.Row + 1) then GSRowEnter(GRIDTYP, GRIDTYP.Row + 1, Cancel, Chg)
        else GRIDTYP.InsertRow(GRIDTYP.Row + 1);
      end;
    VK_TAB: if (Screen.ActiveControl = GRIDTYP) then
        if (GRIDTYP.Col = ColCoef) and (NewLigne) and (EstRempli(GRIDTYP.row)) then
        begin
          OnUpdate;
          if not NewLigne then GSRowEnter(GRIDTYP, GRIDTYP.Row + 1, Cancel, Chg);
        end;
  end;

end;

procedure TOF_MBOTarifType.GSElipsisClick(Sender: TObject);
var DEVISE: THCritMaskEdit;
  Coord: TRect;
begin
  inherited;
  if (GRIDTYP.Col = colDevise) then
  begin
    Coord := GRIDTYP.CellRect(GRIDTYP.Col, GRIDTYP.Row);
    DEVISE := THCritMaskEdit.Create(GRIDTYP);
    DEVISE.Parent := GRIDTYP;
    DEVISE.Top := Coord.Top;
    DEVISE.left := Coord.Left;
    DEVISE.Width := 3;
    DEVISE.Visible := False;
    DEVISE.OpeType := otString;
    DEVISE.DATATYPE := 'TTDEVISE';
    LookUpCombo(DEVISE);
    if DEVISE.Text <> '' then GRIDTYP.Cells[GRIDTYP.Col, GRIDTYP.Row] := DEVISE.Text;
    DEVISE.Destroy;
  end;
end;

procedure TOF_MBOTarifType.GSEnter(Sender: TObject);
var ACol, ARow: integer;
  Temp: Boolean;
begin
  if Action = taConsult then Exit;
  ACol := GRIDTYP.Col;
  ARow := GRIDTYP.Row;
  if not EstRempli(ARow) then
  begin
    GRIDTYP.Col := ColCode;
    NewLigne := True;
  end;
  GRIDTYP.InvalidateRow(ARow);
  GSCellEnter(GRIDTYP, ACol, ARow, Temp);
end;

procedure TOF_MBOTarifType.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  if Action = taConsult then Exit;
  if not EstRempli(ou) then
  begin
    GRIDTYP.DeleteRow(ou);
    SetControlEnabled('BInsert', True);
    SetControlEnabled('BDelete', True);
  end;
  GRIDTYP.InvalidateRow(Ou);
end;

procedure TOF_MBOTarifType.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  if Action = taConsult then Exit;
  if not EstRempli(Ou) then
  begin
    GRIDTYP.Col := ColCode;
    NewLigne := True;
  end
  else NewLigne := False;
  GRIDTYP.InvalidateRow(Ou);
end;

procedure TOF_MBOTarifType.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if Action = taConsult then Exit;
  GRIDTYP.ElipsisButton := (GRIDTYP.Col = colDevise);
  if GRIDTYP.Col = ColCode then
  begin
    if (GRIDTYP.Cells[ColCode, GRIDTYP.Row] <> '') and (not newligne) then
    begin
      GRIDTYP.Col := colLibelle;
    end;
  end;
  if GRIDTYP.Col = ColDevise then
  begin
    if (GRIDTYP.Cells[ColDevise, GRIDTYP.Row] <> '') and (not newligne) then
    begin
      GRIDTYP.Col := colCoef;
    end;
  end;
end;

procedure TOF_MBOTarifType.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if Action = taConsult then Exit;
  FormateColSaisie(Acol, Arow);
  if ACol = ColCode then
    if (GRIDTYP.Cells[ColCode, Arow] <> '') and (NewLigne) and (GRIDTYP.Cells[ColCoef, Arow] = '') then
    begin
      GRIDTYP.Cells[ColDevise, Arow] := GetParamSoc('SO_DEVISEPRINC');
      GRIDTYP.Cells[ColCoef, Arow] := '1';
    end;
  if Acol = colCoef then CalculExemple;
end;

{==============================================================================================}
{=============================== Actions liées au grid ========================================}
{==============================================================================================}

procedure TOF_MBOTarifType.FormateColSaisie(ACol, ARow: Longint);
var st, Stc: string;
begin
  St := GRIDTYP.Cells[ACol, ARow];
  StC := St;
  if ACol = 0 then StC := Trim(St);
  GRIDTYP.Cells[ACol, ARow] := StC;
end;

{==============================================================================================}
{=============================== Mide en forme du grid ========================================}
{==============================================================================================}

procedure TOF_MBOTarifType.ColTriangle(ACol, ARow: Longint; Canvas: TCanvas;
  AState: TGridDrawState);
var Triangle: array[0..2] of TPoint;
  Arect: Trect;
begin
  if Arow < GRIDTYP.Fixedrows then exit;
  if (gdFixed in AState) and (ACol = Col_Mov) then
  begin
    Arect := GRIDTYP.CellRect(Acol, Arow);
    Canvas.Brush.Color := GRIDTYP.FixedColor;
    Canvas.FillRect(ARect);
    if (ARow = GRIDTYP.row) and (GridTyp.RowHeights[ARow] <> 0) then
    begin
      Canvas.Brush.Color := clBlack;
      Canvas.Pen.Color := clBlack;
      Triangle[1].X := ((ARect.Left + ARect.Right) div 2);
      Triangle[1].Y := ((ARect.Top + ARect.Bottom) div 2);
      Triangle[0].X := Triangle[1].X - 5;
      Triangle[0].Y := Triangle[1].Y - 5;
      Triangle[2].X := Triangle[1].X - 5;
      Triangle[2].Y := Triangle[1].Y + 5;
      if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle);
    end;
  end;
end;

procedure TOF_MBOTarifType.CodeGras(ACol, ARow: Longint; Canvas: TCanvas;
  AState: TGridDrawState);
begin
  if (ACol = colCode) and (ARow > 0) then
  begin
    Canvas.Font.Style := [fsBold];
  end;
end;

{==============================================================================================}
{=============================== Chargement des types de tarifs ====================================}
{==============================================================================================}

procedure TOF_MBOTarifType.ChargePeriode;
var Q: TQuery;
  i: Integer;
begin
  TobTyp := TOB.Create('_TARIFTYPMODE', nil, -1);
  Q := OpenSql('Select * from TARIFTYPMODE where GFT_NATURETYPE="' + NatureType + '"', True);
  if not Q.EOF then TobTyp.LoadDetailDB('TARIFTYPMODE', '', '', Q, false, true);
  if TobTyp.detail.count = 0 then Tob.create('TARIFTYPMODE', TobTyp, -1);
  TobTyp.PutGridDetail(GridTyp, True, True, LesColonnes, True);
  TobTyp.GetGridDetail(GridTyp, GridTyp.rowcount - 1, 'TARIFTYPMODE', LesColonnes);
  TobTyp.SetAllModifie(false);
  ferme(Q);
  for i := 0 to TobTyp.detail.count - 1 do
  begin
    if i > MaxType then exit;
    TabEtabMem[i + 1] := TobTyp.detail[i].GetValue('GFT_ETABLISREF');
  end;
  CalculExemple;
end;

//Calcul de la colonne exemple

procedure TOF_MBOTarifType.CalculExemple;
var i: Integer;
  Prix: Double;
begin
  for i := 0 to TobTyp.detail.count - 1 do
  begin
    if i > MaxType then exit;
    Prix := 10 * TobTyp.detail[i].GetValue('GFT_COEF');
    GridTyp.Cells[colExemple, i + 1] := '10 ' + V_PGI.DevisePivot + ' = ' + FloatToStr(Prix) + ' ' + TobTyp.detail[i].GetValue('GFT_DEVISE');
  end;
end;

// Modification des dates par établissement

procedure AGLModifEtabRef(parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then TOTOF := TFVierge(F).LaTof else exit;
  if (TOTOF is TOF_MBOTarifType) then
  begin
    if string(Parms[1]) = '' then TOF_MBOTarifType(TOTOF).EtablisRef := 'x'
    else TOF_MBOTarifType(TOTOF).EtablisRef := string(Parms[1])
  end
  else exit;
end;

initialization
  registerclasses([TOF_MBOTarifType]);
  RegisterAglProc('ModifEtabRef', TRUE, 1, AGLModifEtabRef);
end.
