{***********UNITE*************************************************
Auteur  ...... : O. TARCY
Créé le ...... : 10/01/2001
Modifié le ... : 23/07/2001
Description .. : Source TOF de la FICHE : REMISEBANQUE
Suite ........ : Sélection des chèques à remettre en banque
Mots clefs ... : TOF;UTOFMFOREMISEBANQUE;FO
*****************************************************************}
unit UTOFMFOREMISEBANQUE;

interface
uses
  Classes, Controls, ComCtrls, sysutils, forms, Windows, Messages,
  HCtrls, HFLabel, HMsgBox, HEnt1, Vierge, HStatus, M3FP,
  {$IFNDEF EAGLCLIENT}
  dbTables,
  {$ENDIF}
  UTOB, UTOF ;

type
  TOF_MFOREMISEBANQUE = class(TOF)
  private
    sCaisse, sNaturePiece, Fmt, sDevise: string;
    inbSelected, inbtot: integer;
    TOBSelected, TOBPieces: TOB;
    dtotselected: double;
    gCHQDIFF, gCHQDIFFHE, gCHQ: THGrid;
    GPK_CAISSE: THValComboBox;
    PAGE_CHQ: TPageControl;
    CHQDIFF: TFlashingLabel;
  public
    procedure OnLoad; override;
    procedure OnClose; override;
    procedure OnArgument(S: string); override;
    procedure OnDblClick(Sender: TObject);
    procedure MiseEnForme(grid: THGrid);
    procedure InitTotaux(grid: THGrid; inb: integer; dtot: double);
    procedure OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure OnMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure SelectOne;
    procedure FOSelectALLCHQ;
    procedure FODeselectALLCHQ(Delete: Boolean);
    procedure FOImpression;
    procedure ValideLaRemise;
    function FORecupSelection(grid: THGrid): boolean;
    procedure ReLoad;
    procedure GPK_CAISSEChange(Sender: TObject);
  end;

function SumCol(grid: THGrid; iCol: integer): string;

implementation
uses
  EntGC, FactUtil, FactTOB, FactPiece, FOUtil, FODefi;

const DATEP = 0; //les colonnes des grids
  NUM = 1;
  DATECH = 2;
  NUMECH = 3;
  CLT = 4;
  NOM = 5;
  MTT = 6;
  MTTDEV = 7;
  DEV = 8;
  BQ = 9;

procedure TOF_MFOREMISEBANQUE.FOImpression;
var where: string;
  b1, b2, b3: boolean;
begin
  b1 := False;
  b2 := False;
  b3 := False;
  where := '';
  TOBSelected := TOB.Create('', nil, -1);
  TOBPieces := TOB.Create('', nil, -1);

  // des chèques ont-ils été selectionnés ?
  SetActiveTabSheet('TABCHQDIFF');
  if StrToInt(GetControlText('NBSELGDCHQDIFF')) > 0 then b1 := FORecupSelection(gCHQDIFF);
  SetActiveTabSheet('TABCHQ');
  if StrToInt(GetControlText('NBSELGDCHQ')) > 0 then b2 := FORecupSelection(gCHQ);
  SetActiveTabSheet('TABCHQDIFFHE');
  if StrToInt(GetControlText('NBSELGDCHQDIFFHE')) > 0 then b3 := FORecupSelection(gCHQDIFFHE);

  // s'il y au moins une sélection
  if not ((b1 = False) and (b2 = False) and (b3 = False)) then
  begin
    if PGIAsk('Etes-vous certain des éléments à remettre en banque ?', Ecran.Caption) = mrYes then
    begin
      if b1 = True then
      begin
        SetActiveTabSheet('TABCHQDIFF');
        SetFocusControl('GDCHQDIFF');
        FODeselectALLCHQ(True);
      end;
      if b2 = True then
      begin
        SetActiveTabSheet('TABCHQ');
        SetFocusControl('GDCHQ');
        FODeselectALLCHQ(True);
      end;
      if b3 = True then
      begin
        SetActiveTabSheet('TABCHQDIFFHE');
        SetFocusControl('GDCHQDIFFHE');
        FODeselectALLCHQ(True);
      end;
      if Transactions(ValideLaRemise, 10) <> oeOK then
      begin
        PGIBox('La remise en banque n''est pas enregistrée !', Ecran.Caption);
      end else
      begin
        where := '(MP_TYPEMODEPAIE = "' + TYPEPAIECHEQUE + '" OR MP_TYPEMODEPAIE = "' + TYPEPAIECHQDIFF + '") AND ' +
          'GPE_DATEREMBNQ ="' + USDateTime(Now) + '" AND GPE_CHQDIFUTIL="X" AND ' +
          'GPE_NATUREPIECEG="' + sNaturePiece + '"' ;
      end;
    end else
    begin
      if b3 = True then
      begin
        SetActiveTabSheet('TABCHQDIFFHE');
        SetFocusControl('GDCHQDIFFHE');
        FODeselectALLCHQ(True);
      end;
      if b2 = True then
      begin
        SetActiveTabSheet('TABCHQ');
        SetFocusControl('GDCHQ');
        FODeselectALLCHQ(True);
      end;
      if b1 = True then
      begin
        SetActiveTabSheet('TABCHQDIFF');
        SetFocusControl('GDCHQDIFF');
        FODeselectALLCHQ(True);
      end;
    end;
  end else
    PGIInfo('Pas d''impression car aucun chèque n''a été sélectionné !', Ecran.Caption);
  TOBSelected.Free;
  TOBPieces.Free;
  if where <> '' then
  begin
    if (V_PGI.WinVersion <> '95') and (V_PGI.WinVersion <> '98') then
      FOLanceImprimeLP(efoRemiseBq, where, False, nil);
  end;
  ReLoad;
end;

// Enregistrement de la remise en banque

procedure TOF_MFOREMISEBANQUE.ValideLaRemise;
begin
  if V_PGI.IoError = oeOk then TOBSelected.UpdateDB(True);
  if V_PGI.IoError = oeOk then TOBPieces.UpdateDB;
end;

procedure TOF_MFOREMISEBANQUE.OnClose;
begin
  inherited;
  CHQDIFF.Destroy;
end;

procedure TOF_MFOREMISEBANQUE.OnMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if (ssCtrl in Shift) and (Button = mbLeft) then SelectOne;
end;

procedure TOF_MFOREMISEBANQUE.MiseEnForme(grid: THGrid);
begin
  Fmt := StrfMask(V_PGI.OkDecV, '', TRUE);
  grid.ColAligns[NUM] := taRightJustify;
  grid.ColAligns[MTT] := taRightJustify;
  grid.ColAligns[MTTDEV] := taRightJustify;
  grid.ColAligns[DATECH] := taCenter;
  grid.ColAligns[NUMECH] := taCenter;
  grid.ColAligns[DATEP] := taCenter;
  grid.ColAligns[DEV] := taCenter;
  grid.ColAligns[BQ] := taCenter;
  grid.ColAligns[CLT] := taLeftJustify;
  grid.ColAligns[NOM] := taLeftJustify;
  grid.ColFormats[MTT] := Fmt;
  grid.ColFormats[MTTDEV] := Fmt;
  grid.OnKeyDown := OnKeyDown;
  grid.OnMouseDown := OnMouseDown;
  grid.ColWidths[DATECH] := 58;
  grid.ColWidths[NUM] := 48;
  grid.ColWidths[CLT] := 90;
  grid.ColWidths[NOM] := 90;
  grid.ColWidths[MTT] := 90;
  grid.ColWidths[MTTDEV] := 90;
  grid.ColWidths[DEV] := 42;
  grid.ColWidths[BQ] := 60;
  grid.ColWidths[NUMECH] := 60;
  grid.ColWidths[DATEP] := 70;
  grid.RowCount := 2;
  if (Ecran is TFVierge) then TFVierge(Ecran).HMTrad.ResizeGridColumns(grid);
end;

procedure TOF_MFOREMISEBANQUE.OnLoad;
begin
  inherited;
  ReLoad;
end;

procedure TOF_MFOREMISEBANQUE.ReLoad;
var Q1: TQuery;
  sDate, sSouche: string;
  iRow, inb, i, ichqdiff, ichq, ichqdiffhe: Integer;
  dtot: double;
  c1, c2, c3: TControl;
begin
  inherited;
  MiseEnForme(gCHQDIFF);
  MiseEnForme(gCHQ);
  MiseEnForme(gCHQDIFFHE);
  c1 := GetControl('CHQ');
  c2 := GetControl('CHQDIFFHE');
  c3 := GetControl('MAIN_GROUPBOX');
  CHQDIFF := TFlashingLabel.Create(c3);
  CHQDIFF.parent := TWinControl(c3);
  CHQDIFF.Flashing := True;
  CHQDIFF.Color := $00282CA4;
  CHQDIFF.Left := c1.Left;
  CHQDIFF.Top := c1.Top - (c2.Top - c1.Top);
  CHQDIFF.Caption := TraduireMemoire('Chèques différés à échéance');
  CHQDIFF.Font.Size := 10;
  CHQDIFF.Font.Style := [];
  CHQDIFF.AutoSize := False;
  CHQDIFF.Interval := 750;
  gCHQDIFF.OnDblClick := OnDblClick;
  gCHQ.OnDblClick := OnDblClick;
  gCHQDIFFHE.OnDblClick := OnDblClick;
  inb := 0;
  dtot := 0;
  ichqdiff := 0;
  SetControlText('GPK_CAISSE', sCaisse);
  sCaisse := GetControlText('GPK_CAISSE');
  sDate := USDateTime(Now);
  sSouche := GetSoucheG(sNaturePiece, VH_GC.TOBPCaisse.GetValue('GPK_ETABLISSEMENT'), '');
  Q1 := OpenSQL('SELECT GPE_DATEPIECE, GPE_DATEECHE, GPE_NUMERO, GP_TIERS, T_LIBELLE, GPE_MONTANTECHE, GPE_MONTANTENCAIS, GPE_DEVISE, GPE_NUMECHE '
    + 'FROM GCREGLEMENTFO '
    + 'WHERE (GPE_DATEECHE <= "' + sDate + '" AND MP_TYPEMODEPAIE="' + TYPEPAIECHQDIFF + '" AND '
    + 'GPE_CAISSE="' + sCaisse + '" AND GPE_CHQDIFUTIL="-" AND GPE_NATUREPIECEG="' + sNaturePiece + '" '
    + 'AND GPE_SOUCHE="' + sSouche + '") '
    + 'ORDER BY GPE_DATEECHE ,GPE_DATEPIECE ,GPE_NATUREPIECEG ,GPE_NUMERO', True);
  while not Q1.EOF do
  begin
    Inc(ichqdiff);
    Q1.Next;
  end;
  Q1.First;
  InitMove(ichqdiff, TraduireMemoire('Chargement des chèques différés à échéance'));
  ichqdiff := 0;
  while not Q1.EOF do
  begin
    gCHQDIFF.Cells[DATECH, gCHQDIFF.RowCount - 1] := Q1.FindField('GPE_DATEECHE').AsString;
    gCHQDIFF.Cells[NUM, gCHQDIFF.RowCount - 1] := Q1.FindField('GPE_NUMERO').AsString;
    gCHQDIFF.Cells[CLT, gCHQDIFF.RowCount - 1] := Q1.FindField('GP_TIERS').AsString;
    gCHQDIFF.Cells[NOM, gCHQDIFF.RowCount - 1] := Q1.FindField('T_LIBELLE').AsString;
    gCHQDIFF.Cells[MTT, gCHQDIFF.RowCount - 1] := Q1.FindField('GPE_MONTANTECHE').AsString;
    gCHQDIFF.Cells[MTT, gCHQDIFF.RowCount - 1] := FormatFloat(gCHQDIFF.ColFormats[MTT], Valeur(gCHQDIFF.Cells[MTT, gCHQDIFF.RowCount - 1]));
    gCHQDIFF.Cells[MTTDEV, gCHQDIFF.RowCount - 1] := Q1.FindField('GPE_MONTANTENCAIS').AsString;
    gCHQDIFF.Cells[MTTDEV, gCHQDIFF.RowCount - 1] := FormatFloat(gCHQDIFF.ColFormats[MTTDEV], Valeur(gCHQDIFF.Cells[MTTDEV, gCHQDIFF.RowCount - 1]));
    gCHQDIFF.Cells[DEV, gCHQDIFF.RowCount - 1] := Q1.FindField('GPE_DEVISE').AsString;
    gCHQDIFF.Cells[BQ, gCHQDIFF.RowCount - 1] := 'NON';
    gCHQDIFF.Cells[NUMECH, gCHQDIFF.RowCount - 1] := Q1.FindField('GPE_NUMECHE').AsString;
    gCHQDIFF.Cells[DATEP, gCHQDIFF.RowCount - 1] := Q1.FindField('GPE_DATEPIECE').AsString;
    Inc(inb);
    dtot := dtot + Valeur(gCHQDIFF.Cells[MTT, gCHQDIFF.RowCount - 1]);
    gCHQDIFF.InsertRow(gCHQDIFF.RowCount);
    MoveCur(FALSE);
    Inc(ichqdiff);
    Q1.Next;
  end;
  Ferme(Q1);
  FiniMove;
  InitTotaux(gCHQDIFF, inb, dtot);
  dtotselected := dtot;
  inbselected := inb;
  inb := 0;
  dtot := 0;
  ichq := 0;

  Q1 := OpenSQL('SELECT GPE_DATEPIECE, GPE_DATEECHE, GPE_NUMERO, GP_TIERS, T_LIBELLE, GPE_MONTANTECHE, GPE_MONTANTENCAIS, GPE_DEVISE, GPE_NUMECHE '
    + 'FROM GCREGLEMENTFO '
    + 'WHERE (MP_TYPEMODEPAIE = "' + TYPEPAIECHEQUE + '" AND '
    + 'GPE_CAISSE = "' + sCaisse + '" AND GPE_CHQDIFUTIL="-" AND GPE_NATUREPIECEG="' + sNaturePiece + '" '
    + 'AND GPE_SOUCHE="' + sSouche + '") '
    + 'ORDER BY GPE_DATEECHE ,GPE_DATEPIECE ,GPE_NATUREPIECEG ,GPE_NUMERO ', True);
  while not Q1.EOF do
  begin
    Inc(ichq);
    Q1.Next;
  end;
  Q1.First;
  InitMove(ichq, TraduireMemoire('Chargement des chèques'));
  ichq := 0;
  while not Q1.EOF do
  begin
    GCHQ.Cells[DATECH, gchq.rowcount - 1] := Q1.FindField('GPE_DATEECHE').AsString;
    GCHQ.Cells[NUM, gchq.rowcount - 1] := Q1.FindField('GPE_NUMERO').AsString;
    GCHQ.Cells[CLT, gchq.rowcount - 1] := Q1.FindField('GP_TIERS').AsString;
    GCHQ.Cells[NOM, gchq.rowcount - 1] := Q1.FindField('T_LIBELLE').AsString;
    GCHQ.Cells[MTT, gchq.rowcount - 1] := Q1.FindField('GPE_MONTANTECHE').AsString;
    GCHQ.Cells[MTT, gchq.rowcount - 1] := FormatFloat(GCHQ.ColFormats[MTT], Valeur(GCHQ.Cells[MTT, gchq.rowcount - 1]));
    GCHQ.Cells[MTTDEV, gchq.rowcount - 1] := Q1.FindField('GPE_MONTANTENCAIS').AsString;
    GCHQ.Cells[MTTDEV, gchq.rowcount - 1] := FormatFloat(GCHQ.ColFormats[MTTDEV], Valeur(GCHQ.Cells[MTTDEV, gchq.rowcount - 1]));
    GCHQ.Cells[DEV, gchq.rowcount - 1] := Q1.FindField('GPE_DEVISE').AsString;
    GCHQ.Cells[BQ, gchq.rowcount - 1] := 'NON';
    GCHQ.Cells[NUMECH, gchq.rowcount - 1] := Q1.FindField('GPE_NUMECHE').AsString;
    GCHQ.Cells[DATEP, gchq.rowcount - 1] := Q1.FindField('GPE_DATEPIECE').AsString;
    Inc(inb);
    dtot := dtot + Valeur(gCHQ.Cells[MTT, gchq.rowcount - 1]);
    GCHQ.InsertRow(GCHQ.RowCount);
    MoveCur(FALSE);
    Inc(ichq);
    Q1.Next;
  end;
  Ferme(Q1);
  FiniMove;
  InitTotaux(gCHQ, inb, dtot);
  iRow := 0;
  dtotselected := dtotselected + dtot;
  inbselected := inbselected + inb;
  inb := 0;
  dtot := 0;
  ichqdiffhe := 0;

  Q1 := OpenSQL('SELECT GPE_DATEPIECE, GPE_DATEECHE, GPE_NUMERO, GP_TIERS, T_LIBELLE, GPE_MONTANTECHE, GPE_MONTANTENCAIS, GPE_DEVISE, GPE_NUMECHE '
    + 'FROM GCREGLEMENTFO '
    + 'WHERE (GPE_DATEECHE > "' + sDate + '" AND MP_TYPEMODEPAIE = "' + TYPEPAIECHQDIFF + '" AND '
    + 'GPE_CAISSE = "' + sCaisse + '" AND GPE_CHQDIFUTIL="-" AND GPE_NATUREPIECEG="' + sNaturePiece + '" '
    + 'AND GPE_SOUCHE="' + sSouche + '") '
    + 'ORDER BY GPE_DATEECHE ,GPE_DATEPIECE ,GPE_NATUREPIECEG ,GPE_NUMERO ', True);
  while not Q1.EOF do
  begin
    Inc(ichqdiffhe);
    Q1.Next;
  end;
  Q1.First;
  InitMove(ichqdiffhe, TraduireMemoire('Chargement des chèques hors échéance'));
  ichqdiffhe := 0;
  while not Q1.EOF do
  begin
    GCHQDIFFHE.Cells[DATECH, GCHQDIFFHE.RowCount - 1] := Q1.FindField('GPE_DATEECHE').AsString;
    GCHQDIFFHE.Cells[NUM, GCHQDIFFHE.RowCount - 1] := Q1.FindField('GPE_NUMERO').AsString;
    GCHQDIFFHE.Cells[CLT, GCHQDIFFHE.RowCount - 1] := Q1.FindField('GP_TIERS').AsString;
    GCHQDIFFHE.Cells[NOM, GCHQDIFFHE.RowCount - 1] := Q1.FindField('T_LIBELLE').AsString;
    GCHQDIFFHE.Cells[MTT, GCHQDIFFHE.RowCount - 1] := Q1.FindField('GPE_MONTANTECHE').AsString;
    GCHQDIFFHE.Cells[MTT, GCHQDIFFHE.RowCount - 1] := FormatFloat(GCHQDIFFHE.ColFormats[MTT], Valeur(GCHQDIFFHE.Cells[MTT, GCHQDIFFHE.RowCount - 1]));
    GCHQDIFFHE.Cells[MTTDEV, GCHQDIFFHE.RowCount - 1] := Q1.FindField('GPE_MONTANTENCAIS').AsString;
    GCHQDIFFHE.Cells[MTTDEV, GCHQDIFFHE.RowCount - 1] := FormatFloat(GCHQDIFFHE.ColFormats[MTTDEV], Valeur(GCHQDIFFHE.Cells[MTTDEV, GCHQDIFFHE.RowCount - 1]));
    GCHQDIFFHE.Cells[DEV, GCHQDIFFHE.RowCount - 1] := Q1.FindField('GPE_DEVISE').AsString;
    GCHQDIFFHE.Cells[BQ, GCHQDIFFHE.RowCount - 1] := 'NON';
    GCHQDIFFHE.Cells[NUMECH, GCHQDIFFHE.RowCount - 1] := Q1.FindField('GPE_NUMECHE').AsString;
    GCHQDIFFHE.Cells[DATEP, GCHQDIFFHE.RowCount - 1] := Q1.FindField('GPE_DATEPIECE').AsString;
    GCHQDIFFHE.InsertRow(GCHQDIFFHE.RowCount);
    Inc(iRow);
    MoveCur(FALSE);
    Inc(ichqdiffhe);
    Q1.Next;
  end;
  Ferme(Q1);
  FiniMove;
  InitTotaux(gCHQDIFFHE, inb, dtot);
  for i := 0 to 13 do
    SetControlText('DEV' + IntToStr(i), sDevise);
  SetControlText('NBSEL', '0');
  SetControlText('MTTSEL', FormatFloat(Fmt, 0));
  SetControlText('NBTOT', IntToStr(inbselected + iRow));
  inbtot := inbselected + iRow;
  SetControlText('MTTTOT', FormatFloat(Fmt, dtotselected + Valeur(GetControlText('MTTTOTGDCHQDIFFHE'))));

  {  SetActiveTabSheet('TABCHQDIFF') ;
    if gCHQDIFF.Enabled = True then SetFocusControl('GDCHQDIFF') ;
    FOSelectAllCHQ ;
    SetActiveTabSheet('TABCHQ') ;
    if gCHQ.Enabled = True then SetFocusControl('GDCHQ') ;
    FOSelectAllCHQ ;
    SetActiveTabSheet('TABCHQDIFFHE') ;
    if gCHQDIFFHE.Enabled = True then SetFocusControl('GDCHQDIFFHE') ;  }

  if ichqdiffhe = 0 then SetControlVisible('TABCHQDIFFHE', False) else
  begin
    SetActiveTabSheet('TABCHQDIFFHE');
    if gCHQDIFFHE.Enabled = True then SetFocusControl('GDCHQDIFFHE');
  end;

  if ichq = 0 then SetControlVisible('TABCHQ', False) else
  begin
    SetActiveTabSheet('TABCHQ');
    if gCHQ.Enabled = True then SetFocusControl('GDCHQ');
    FOSelectAllCHQ;
  end;

  if ichqdiff = 0 then SetControlVisible('TABCHQDIFF', False) else
  begin
    SetActiveTabSheet('TABCHQDIFF');
    if gCHQDIFF.Enabled = True then SetFocusControl('GDCHQDIFF');
    FOSelectAllCHQ;
  end;

  if (ichqdiffhe = 0) and (ichq = 0) and (ichqdiff = 0) then
  begin
    SetControlEnabled('BVALIDATION', False);
    SetControlEnabled('BVALIDATION2', False);
  end;
end;

procedure TOF_MFOREMISEBANQUE.OnArgument(S: string);
var x: integer;
  ChampMul, ValMul, Critere: string;
begin
  inherited;
  if V_PGI.TenueEuro = True then sDevise := SIGLEEURO else sDevise := V_PGI.DevisePivot;
  sCaisse := '';
  repeat
    Critere := Trim(ReadTokenSt(S));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'CAISSE' then sCaisse := ValMul;
        if ChampMul = 'NATUREPIECE' then sNaturePiece := ValMul;
      end;
    end;
  until Critere = '';
  gCHQDIFF := THGrid(GetControl('GDCHQDIFF'));
  gCHQ := THGrid(GetControl('GDCHQ'));
  gCHQDIFFHE := THGrid(GetControl('GDCHQDIFFHE'));
  GPK_CAISSE := THValComboBox(GetControl('GPK_CAISSE'));
  GPK_CAISSE.OnChange := GPK_CAISSEChange;
  GPK_CAISSE.Enabled := False;
  PAGE_CHQ := TPageControl(GetControl('PAGE_CHQ'));
end;

function TOF_MFOREMISEBANQUE.FORecupSelection(grid: THGrid): boolean;
var iRow: integer;
  sSQL: string;
  CleDoc: R_CleDoc;
  TOBS, TOBP: TOB;
  Q: TQuery;
begin
  result := False;
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  CleDoc.NaturePiece := sNaturePiece;
  CleDoc.Souche := GetSoucheG(CleDoc.NaturePiece, VH_GC.TOBPCaisse.GetValue('GPK_ETABLISSEMENT'), '');
  CleDoc.Indice := 0;
  CleDoc.NumLigne := 1;
  InitMove(grid.RowCount - 1 - grid.FixedRows, TraduireMemoire('Traitement'));
  for iRow := grid.FixedRows to grid.RowCount - 1 do
    if grid.Cells[BQ, iRow] = 'OUI' then
    begin
      MoveCur(False);
      if result = False then result := True;
      CleDoc.DatePiece := StrToDate(grid.Cells[DATEP, iRow]);
      CleDoc.NumeroPiece := StrToInt(grid.Cells[NUM, iRow]);
      sSQL := 'SELECT * FROM PIEDECHE WHERE ' + WherePiece(CleDoc, ttdEche, False) + ' and GPE_NUMECHE=' + grid.Cells[NUMECH, iRow];
      Q := OpenSQL(sSQL, True);
      if not Q.EOF then
      begin
        TOBS := TOB.Create('PIEDECHE', TOBSelected, -1);
        TOBS.SelectDB('', Q);
        TOBS.PutValue('GPE_CHQDIFUTIL', 'X');
        TOBS.PutValue('GPE_DATEREMBNQ', Date);
      end;
      Ferme(Q);
      // mise à jour de la date de modification de la pièce
      sSQL := 'SELECT * FROM PIECE WHERE ' + WherePiece(CleDoc, ttdPiece, False);
      Q := OpenSQL(sSQL, True);
      if not Q.EOF then
      begin
        TOBP := TOB.Create('PIECE', TOBPieces, -1);
        TOBP.SelectDB('', Q);
        TOBP.UpdateDateModif;
      end;
      Ferme(Q);
    end;
  FiniMove;
end;

//Calcul de la somme d'une colonne d'un grid

function SumCol(grid: THGrid; iCol: integer): string;
var iRow: integer;
  sumcol: double;
begin
  sumcol := 0;
  for iRow := grid.FixedRows to grid.RowCount - 1 do
    sumcol := sumcol + Valeur(grid.cells[iCol, iRow]);
  result := FormatFloat(grid.ColFormats[MTT], sumcol);
end;

procedure TOF_MFOREMISEBANQUE.SelectOne;
var grid: THGrid;
  c: TWinControl;
begin
  c := Screen.ActiveControl;
  if c is THGrid then
  begin
    grid := THGrid(c);
    if grid.Cells[BQ, grid.Row] = 'NON' then
    begin
      grid.Cells[BQ, grid.Row] := 'OUI';
      SetControlText('NBSEL' + grid.Name, IntToStr(StrToInt(GetControlText('NBSEL' + grid.Name)) + 1));
      SetControlText('NBSEL', IntToStr(StrToInt(GetControlText('NBSEL')) + 1));
      SetControlText('MTTSEL' + grid.Name, FormatFloat(Fmt, Valeur(GetControlText('MTTSEL' + grid.Name)) + Valeur(grid.Cells[MTT, grid.Row])));
      SetControlText('MTTSEL', FormatFloat(Fmt, Valeur(GetControlText('MTTSEL')) + Valeur(grid.Cells[MTT, grid.Row])));
    end else
    begin
      grid.Cells[BQ, grid.Row] := 'NON';
      SetControlText('NBSEL' + grid.Name, IntToStr(StrToInt(GetControlText('NBSEL' + grid.Name)) - 1));
      SetControlText('NBSEL', IntToStr(StrToInt(GetControlText('NBSEL')) - 1));
      SetControlText('MTTSEL' + grid.Name, FormatFloat(Fmt, Valeur(GetControlText('MTTSEL' + grid.Name)) - Valeur(grid.Cells[MTT, grid.Row])));
      SetControlText('MTTSEL', FormatFloat(Fmt, Valeur(GetControlText('MTTSEL')) - Valeur(grid.Cells[MTT, grid.Row])));
    end;
    SetControlText('NBSEL' + grid.Name + '_', GetControlText('NBSEL' + grid.Name));
    SetControlText('MTTSEL' + grid.Name + '_', GetControlText('MTTSEL' + grid.Name));
  end;
end;

//Actualisation des champs calculés

procedure TOF_MFOREMISEBANQUE.OnDblClick(Sender: TObject);
var grid: THGrid;
begin
  grid := THGrid(Sender);
  SendMessage(Grid.Handle, WM_KEYDOWN, VK_SPACE, 0);
end;

procedure TOF_MFOREMISEBANQUE.InitTotaux(grid: THGrid; inb: integer; dtot: double);
begin
  if (grid.RowCount > 2) then
  begin
    grid.DeleteRow(grid.RowCount - 1);
    SetControlText('MTTTOT' + grid.Name, SumCol(grid, MTT));
    SetControlText('NBTOT' + grid.Name, IntToStr(grid.RowCount - grid.FixedRows));
  end else
  begin
    SetControlText('MTTTOT' + grid.Name, '0' + V_PGI.SepDecimal + '00');
    SetControlText('NBTOT' + grid.Name, '0');
  end;
  SetControlText('NBSEL' + grid.Name, '0');
  SetControlText('MTTSEL' + grid.Name, '0' + V_PGI.SepDecimal + '00');
  SetControlText('MTTTOT' + grid.Name + '_', GetControlText('MTTTOT' + grid.Name));
  SetControlText('NBTOT' + grid.Name + '_', GetControlText('NBTOT' + grid.Name));
  SetControlText('NBSEL' + grid.Name + '_', GetControlText('NBSEL' + grid.Name));
  SetControlText('MTTSEL' + grid.Name + '_', GetControlText('MTTSEL' + grid.Name));
end;

procedure TOF_MFOREMISEBANQUE.FOSelectALLCHQ;
var iRow: integer;
  grid: THGrid;
  c: TWinControl;
begin
  c := Screen.ActiveControl;
  if c is THGrid then
  begin
    grid := THGrid(c);
    grid.row := grid.FixedRows;
    for iRow := grid.FixedRows to grid.RowCount - 1 do
    begin
      if grid.Cells[BQ, iRow] = 'NON' then SendMessage(grid.Handle, WM_KEYDOWN, VK_SPACE, 0);
      if iRow <> grid.RowCount - 1 then Grid.Row := Grid.Row + 1;
    end;
    grid.row := grid.FixedRows;
  end;
end;

procedure TOF_MFOREMISEBANQUE.FODeselectALLCHQ(Delete: Boolean);
var iRow, iCol: integer;
  grid: THGrid;
  c: TWinControl;
begin
  c := Screen.ActiveControl;
  if c is THGrid then
  begin
    grid := THGrid(c);
    grid.row := grid.FixedRows;
    for iRow := grid.FixedRows to grid.RowCount - 1 do
    begin
      if grid.Cells[BQ, iRow] = 'OUI' then
      begin
        SendMessage(grid.Handle, WM_KEYDOWN, VK_SPACE, 0);
        if Delete = True then
          for iCol := DATEP to BQ do grid.cells[iCol, iRow] := ''; // le chèque sélectionné est "effacé"
      end;
      if iRow <> grid.RowCount - 1 then Grid.Row := Grid.Row + 1;
    end;
    grid.row := grid.FixedRows;
  end;
end;

procedure TOF_MFOREMISEBANQUE.OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_SPACE: SelectOne;
    VK_F10: FOImpression;
  end;
end;

procedure TOF_MFOREMISEBANQUE.GPK_CAISSEChange(Sender: TObject);
var St: string;
begin
  St := GetControlText('GPK_CAISSE');
  if sCaisse <> St then
  begin
    sCaisse := St;
    if gCHQDIFF.Enabled = True then
    begin
      SetFocusControl('GDCHQDIFF');
      FODeselectALLCHQ(True);
    end;
    gCHQDIFF.VidePile(False);
    if gCHQDIFFHE.Enabled = True then
    begin
      SetFocusControl('GDCHQDIFFHE');
      FODeselectALLCHQ(True);
    end;
    gCHQDIFFHE.VidePile(False);
    if gCHQ.Enabled = True then
    begin
      SetFocusControl('GDCHQ');
      FODeselectALLCHQ(True);
    end;
    gCHQ.VidePile(False);
    PAGE_CHQ.ActivePageIndex := -1;
    SetControlVisible('TABCHQDIFFHE', False);
    SetControlVisible('TABCHQ', False);
    SetControlVisible('TABCHQDIFF', False);
    SetControlEnabled('BVALIDATION', True);
    SetControlEnabled('BVALIDATION2', True);
    Reload;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : FOUp : Déplacer la focalisation vers le haut depuis le script
Suite ........ : d'une fiche
Suite ........ :  - Parms[0] = Fiche
Mots clefs ... : FO
*****************************************************************}

procedure FOAGLUp(Parms: array of variant; nb: Integer);
var F: TForm;
  grid: THGrid;
  c: TWinControl;
begin
  F := TForm(Longint(Parms[0]));
  if not (F is TFVierge) then exit;
  c := Screen.ActiveControl;
  if c is THGrid then
  begin
    grid := THGrid(c);
    SendMessage(Grid.Handle, WM_KEYDOWN, VK_UP, 0);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : FODown : Déplacer la focalisation vers le bas depuis le
Suite ........ : script d'une fiche
Suite ........ :  - Parms[0] = Fiche
Mots clefs ... : FO
*****************************************************************}

procedure FOAGLDown(Parms: array of variant; nb: Integer);
var F: TForm;
  grid: THGrid;
  c: TWinControl;
begin
  F := TForm(Longint(Parms[0]));
  if not (F is TFVierge) then exit;
  c := Screen.ActiveControl;
  if c is THGrid then
  begin
    grid := THGrid(c);
    SendMessage(Grid.Handle, WM_KEYDOWN, VK_DOWN, 0);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : FOSelect = Sélectionne un chèque depuis le script d'une
Suite ........ : fiche
Suite ........ :  - Parms[0] = Fiche
Mots clefs ... : FO
*****************************************************************}

procedure FOAGLSelect(Parms: array of variant; nb: Integer);
var F: TForm;
  grid: THGrid;
  c: TWinControl;
begin
  F := TForm(Longint(Parms[0]));
  if not (F is TFVierge) then exit;
  c := Screen.ActiveControl;
  if c is THGrid then
  begin
    grid := THGrid(c);
    SendMessage(Grid.Handle, WM_KEYDOWN, VK_SPACE, 0);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : FOSelectALLCHQ = Sélectionne tout le grid des CHQ
Suite ........ : depuis le script d'une fiche
Suite ........ :  - Parms[0] = Fiche
Mots clefs ... : FO
*****************************************************************}

procedure FOAGLSelectAllCHQ(Parms: array of variant; nb: Integer);
var F: TForm;
  MaTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then MaTOF := TFVierge(F).LaTOF else exit;
  if (MaTOF is TOF_MFOREMISEBANQUE) then TOF_MFOREMISEBANQUE(MaTOF).FOSelectAllCHQ else Exit;
end;

{***********A.G.L.***********************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : FODeselectALLCHQ = Désélectionne tout le grid des CHQ
Suite ........ : depuis le script d'une fiche
Suite ........ :  - Parms[0] = Fiche
Suite ........ :  - Parms[1] = le chèque sélectionné est "effacé"
Mots clefs ... : FO
*****************************************************************}

procedure FOAGLDeselectAllCHQ(Parms: array of variant; nb: Integer);
var F: TForm;
  MaTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then MaTOF := TFVierge(F).LaTOF else exit;
  if (MaTOF is TOF_MFOREMISEBANQUE) then TOF_MFOREMISEBANQUE(MaTOF).FODeselectAllCHQ(Boolean(Parms[1])) else Exit;
end;

{***********A.G.L.***********************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : FOImpression = Impression du bordereau de remise en
Suite ........ : banque des chèques depuis le script d'une fiche
Suite ........ :  - Parms[0] = Fiche
Mots clefs ... : FO
*****************************************************************}

procedure FOAGLImpression(Parms: array of variant; nb: Integer);
var F: TForm;
  MaTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then MaTOF := TFVierge(F).LaTOF else exit;
  if (MaTOF is TOF_MFOREMISEBANQUE) then TOF_MFOREMISEBANQUE(MaTOF).FOImpression else Exit;
end;

initialization
  Registerclasses([TOF_MFOREMISEBANQUE]);
  RegisterAglProc('FOUp', TRUE, 0, FOAGLUp);
  RegisterAglProc('FODown', TRUE, 0, FOAGLDown);
  RegisterAglProc('FOSelect', TRUE, 0, FOAGLSelect);
  RegisterAglProc('FOSelectALLCHQ', TRUE, 0, FOAGLSelectALLCHQ);
  RegisterAglProc('FODeselectALLCHQ', TRUE, 1, FOAGLDeselectALLCHQ);
  RegisterAglProc('FOImpression', TRUE, 0, FOAGLImpression);
end.
