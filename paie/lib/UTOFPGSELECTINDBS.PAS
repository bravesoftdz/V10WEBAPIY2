{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 15/04/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGSelectIndBs ()
Mots clefs ... : TOF;PGSelectIndBs
*****************************************************************}
unit UTofPGSelectIndBs;

interface

uses StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} mul,
  {$ELSE}
  eMul,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, uTob, HTB97, Vierge, HSysMenu;

type
  TOF_PGSelectIndBs = class(TOF)
    procedure OnArgument(S: string); override;
  private
    GrilleParam, GrilleSeclect : THGrid;
    procedure RemplirLesGrilles;
    procedure GRechercheOnDblClick(Sender: TObject);
    procedure DeposeCompetSal;
    procedure DeposeCompet;
    procedure AfficherSelection(Sender: TObject);
    procedure MiseEnFormeGrille;
    procedure GrilleMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure ChangeOnglet(Sender: TObject);
    procedure TEST_DEPOSE_OBJET(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure DEPOSE_OBJET(Sender, Source: TObject; X, Y: Integer);
  end;

implementation

procedure TOF_PGSelectIndBs.OnArgument(S: string);
begin
  inherited;
  MiseEnFormeGrille;
  RemplirLesGrilles;
  GrilleParam := THGrid(GetControl('GCOMPETENCES'));
  GrilleSelect := THGrid(GetControl('GCOMPSAL'));
  GrilleParam.OnMouseDown := GrilleMouseDown;
  GrilleSelect.OnMouseDown := GrilleMouseDown;
  GrilleParam.OnDragDrop := DEPOSE_OBJET;
  GrilleParam.OnDragOver := TEST_DEPOSE_OBJET;
  GrilleSelect.OnDragDrop := DEPOSE_OBJET;
  GrilleSelect.OnDragOver := TEST_DEPOSE_OBJET;
end;

procedure TOF_PGSelectIndBs.MiseEnFormeGrille;
var
  GrilleEmploi, GrilleSalarie: THGrid;
begin
  GrilleEmploi := THGrid(GetControl('GEMPLOI'));
  if GrilleEmploi <> nil then
  begin
    GrilleEmploi.ColFormats[0] := 'CB=PGRHCOMPETENCES';
    GrilleEmploi.ColFormats[1] := '# ##0.00';
    GrilleEmploi.ColAligns[1] := taRightJustify;
    GrilleEmploi.ColFormats[2] := '# ##0.00';
    GrilleEmploi.ColAligns[2] := taRightJustify;
    GrilleEmploi.ColFormats[3] := 'CB=PGOUINONGRAPHIQUE';
  end;
  GrilleSalarie := THGrid(GetControl('GSALARIE'));
  if GrilleSalarie <> nil then
  begin
    GrilleSalarie.ColFormats[0] := 'CB=PGRHCOMPETENCES';
    GrilleSalarie.ColFormats[1] := '# ##0.00';
    GrilleSalarie.ColAligns[1] := taRightJustify;
  end;
end;


procedure TOF_PGSelectIndBs.RemplirLesGrilles;
var
  Q: TQuery;
  TobCompetences: Tob;
  HMTrad: THSystemMenu;
  Grille: THGrid;
  i: Integer;
begin
  Grille := THGrid(GetControl('GCOMPSAL'));
  for i := 1 to Grille.RowCount - 1 do
  begin
    Grille.Rows[i].Clear;
  end;
  Q := OpenSQL('SELECT PCH_COMPETENCE,PCO_LIBELLE,PCH_DEGREMAITRISE FROM RHCOMPETRESSOURCE ' +
    'LEFT JOIN RHCOMPETENCES ON PCO_COMPETENCE=PCH_COMPETENCE ' +
    'WHERE PCH_TYPERESSOURCE="SAL" AND PCH_SALARIE="' + LeSalarie + '"', True);
  TobCompetences := Tob.Create('LesCompetences', nil, -1);
  TobCompetences.LoadDetailDB('LesCompetences', '', '', Q, False);
  Ferme(Q);
  TobCompetences.PutGridDetail(Grille, False, True, '', False);
  if TobCompetences.Detail.count > 0 then Grille.RowCount := TobCompetences.Detail.count + 1
  else Grille.RowCount := 2;
  TobCompetences.Free;
  HMTrad.ResizeGridColumns(THGrid(GetControl('GCOMPSAL')));
  Grille := THGrid(GetControl('GCOMPETENCES'));
  for i := 1 to Grille.RowCount - 1 do
  begin
    Grille.Rows[i].Clear;
  end;
  Q := OpenSQL('SELECT PCO_COMPETENCE,PCO_LIBELLE FROM RHCOMPETENCES', True);
  TobCompetences := Tob.Create('LesCompetences', nil, -1);
  TobCompetences.LoadDetailDB('LesCompetences', '', '', Q, False);
  Ferme(Q);
  TobCompetences.PutGridDetail(Grille, False, True, '', False);
  if TobCompetences.Detail.count > 0 then Grille.RowCount := TobCompetences.Detail.count + 1
  else Grille.RowCount := 2;
  TobCompetences.Free;
  HMTrad.ResizeGridColumns(THGrid(GetControl('GCOMPETENCES')));
end;

procedure TOF_PGSelectIndBs.DeposeCompet;
var
  i, NouvelleLigne: Integer;
begin
  for i := 1 to GrilleComp.RowCount - 1 do
  begin
    if GrilleComp.IsSelected(i) then
    begin
      if (GrilleRecherche.RowCount = 2) and (GrilleRecherche.cellValues[0, 1] = '') then GrilleRecherche.RowCount := 1;
      GrilleRecherche.RowCount := GrilleRecherche.RowCount + 1;
      GrilleRecherche.FixedRows := 1;
      NouvelleLigne := GrilleRecherche.RowCount - 1;
      GrilleRecherche.CellValues[0, NouvelleLigne] := GrilleComp.CellValues[0, i];
      GrilleRecherche.CellValues[1, NouvelleLigne] := GrilleComp.CellValues[1, i];
      GrilleRecherche.CellValues[2, NouvelleLigne] := '0';
    end;
  end;
  GrilleComp.ClearSelected;
end;

procedure TOF_PGSelectIndBs.DeposeCompetSal;
var
  i, NouvelleLigne: Integer;
begin
  for i := 1 to GrilleCompSal.RowCount - 1 do
  begin
    if GrilleCompSal.IsSelected(i) then
    begin
      if (GrilleRecherche.RowCount = 2) and (GrilleRecherche.cellValues[0, 1] = '') then GrilleRecherche.RowCount := 1;
      GrilleRecherche.RowCount := GrilleRecherche.RowCount + 1;
      GrilleRecherche.FixedRows := 1;
      NouvelleLigne := GrilleRecherche.RowCount - 1;
      GrilleRecherche.CellValues[0, NouvelleLigne] := GrilleCompSal.CellValues[0, i];
      GrilleRecherche.CellValues[1, NouvelleLigne] := GrilleCompSal.CellValues[1, i];
      GrilleRecherche.CellValues[2, NouvelleLigne] := GrilleCompSal.CellValues[2, i];
    end;
  end;
  GrilleCompSal.ClearSelected;
end;

procedure TOF_PGSelectIndBs.GRechercheOnDblClick(Sender: TObject);
var
  i: Integer;
begin
  i := GrilleRecherche.Row;
  if (i = 1) and (GrilleRecherche.RowCount = 2) then GrilleRecherche.Rows[1].Clear
  else GrilleRecherche.DeleteRow(i);
end;

procedure TOF_PGSelectIndBs.AfficherSelection(Sender: TObject);
var
  i, c, Ligne: Integer;
  TobComp: Tob;
  Q: TQuery;
begin
  for i := 1 to GrilleRecherche.RowCount - 1 do
  begin
    GrilleRecherche.CellValues[3, i] := 'NON';
    GrilleRecherche.CellValues[4, i] := '0';
  end;
  Ligne := GrilleForm.Row;
  Q := OpenSQL('SELECT POS_COMPETENCE,POS_DEGREMAITRISE FROM STAGEOBJECTIF WHERE POS_CODESTAGE="' + GrilleForm.cellValues[0, Ligne] + '"', True);
  TobComp := Tob.Create('LesComp', nil, -1);
  TobComp.LoadDetailDB('LesComp', '', '', Q, False);
  Ferme(Q);
  for c := 0 to TobComp.Detail.Count - 1 do
  begin
    for i := 1 to GrilleRecherche.RowCount - 1 do
    begin
      if GrilleRecherche.CellValues[0, i] = TobComp.Detail[c].GetValue('POS_COMPETENCE') then
      begin
        GrilleRecherche.CellValues[3, i] := 'OUI';
        GrilleRecherche.CellValues[4, i] := FloatToStr(TobComp.Detail[c].GetValue('POS_DEGREMAITRISE'));
      end;
    end;
  end;
  TobComp.Free;
end;

procedure TOF_PGSelectIndBs.ChangeOnglet(Sender: TObject);
var
  GEmploi: THGrid;
  i, NouvelleLigne: Integer;
begin
  for i := 1 to GrilleRecherche.RowCount - 1 do
  begin
    GrilleRecherche.Rows[i].Clear;
  end;
  GrilleRecherche.RowCount := 2;
  GEmploi := THGrid(GetControl('GEMPLOI'));
  for i := 1 to GEmploi.RowCount - 1 do
  begin
    if GEmploi.Cellvalues[3, i] = 'NON' then
    begin
      if (GrilleRecherche.RowCount = 2) and (GrilleRecherche.cellValues[0, 1] = '') then GrilleRecherche.RowCount := 1;
      GrilleRecherche.RowCount := GrilleRecherche.RowCount + 1;
      GrilleRecherche.FixedRows := 1;
      NouvelleLigne := GrilleRecherche.RowCount - 1;
      GrilleRecherche.CellValues[0, NouvelleLigne] := GEmploi.CellValues[0, i];
      GrilleRecherche.CellValues[1, NouvelleLigne] := RechDom('PGRHCOMPETENCES', GEmploi.CellValues[0, i], False);
      GrilleRecherche.CellValues[2, NouvelleLigne] := GEmploi.CellValues[1, i];
    end;
  end;
  if TTabSheet(Sender).Name = 'PCOMPETENCES' then ChercherFormations;
end;

procedure TOF_PGSelectIndBs.GrilleMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if not (ssCtrl in Shift) then
  begin
    if Sender = GrilleComp then GrilleComp.BeginDrag(TRUE, 10);
    if Sender = GrilleCompSal then GrilleCompSal.BeginDrag(TRUE, 10);
  end;
end;

procedure TOF_PGSelectIndBs.DEPOSE_OBJET(Sender, Source: TObject; X, Y: Integer);
begin
  if Source = GrilleComp then DeposeCompet;
  if Source = GrilleCompSal then DeposeCompetSal;
end;

procedure TOF_PGSelectIndBs.TEST_DEPOSE_OBJET(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
  Accept := FALSE;
  if (Sender is THgrid) then Accept := TRUE;
end;

initialization
  registerclasses([TOF_PGSelectIndBs]);
end.

