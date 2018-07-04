{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Unit de visualisation des cumuls pdt la consultation du
Suite ........ : bulletin
Mots clefs ... : PAIE;PGBULLETIN
*****************************************************************}
{
PT1 : 31/03/2006 PH V_65 FQ 12225 rajout colonne total
}
unit UTofCUMUL_BUL;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, Grids,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DBGrids,
{$ELSE}

{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, UTOB, Vierge, P5Util, AGLInit;
type
  TOF_PGCUMUL_BUL = class(TOF)
  private
    procedure PasDeSaisieRow(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure PasDeSaisieCell(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
  public
    procedure OnLoad; override;
    procedure OnArgument(Arguments: string); override;
  end;

implementation

procedure TOF_PGCUMUL_BUL.OnArgument(Arguments: string);
var
  Grille: THGrid;
begin
  inherited;
  Grille := THGrid(GetControl('GRILLECUMUL'));
  if Grille <> nil then
  begin
    Grille.OnRowEnter := PasDeSaisieRow;
    Grille.OnCellEnter := PasDeSaisieCell;
    Grille.CacheEdit;
    Grille.Options := Grille.Options - [GoEditing, GoTabs, GoAlwaysShowEditor];
    Grille.Options := Grille.Options + [GoRowSelect];
  end
  else Grille.Enabled := FALSE;
end;

procedure TOF_PGCUMUL_BUL.OnLoad;
var
  FichCumBul: TFVierge;
  Grille: THGrid;
  I, j: Integer;
  T1: TOB;
begin
  inherited;
  j := 1;
  Grille := NIL;
  if LaTOB <> nil then
  begin
    FichCumBul := TFVierge(Ecran);
    if FichCumBul <> nil then
    begin
      with FichCumBul do
      begin
        for i := 0 to ComponentCount - 1 do
        begin
          if Components[i] is THGrid then
          begin
            Grille := THGrid(Components[i]);
          end;
        end;
      end;
      if Grille <> nil then
      begin
        Grille.RowCount := LaTOB.Detail.Count + 1; // Affectation du nombre de lignes de la grille
        for I := 0 to LaTOB.Detail.Count - 1 do
        begin
          T1 := LaTOB.Detail[I];
          Grille.Cells[0, j] := T1.GetValue('PCL_CUMULPAIE');
          Grille.Cells[1, j] := T1.GetValue('PCL_LIBELLE');
          Grille.Cells[2, j] := DoubleToCell(T1.GetValue('MONTANT'), 2);
          Grille.Cells[3, j] := DoubleToCell(T1.GetValue('ENCOURS'), 2);
          Grille.Cells[4, j] := DoubleToCell(T1.GetValue('MONTANT')+T1.GetValue('ENCOURS'), 2);
          j := j + 1;
        end;
      end;
    end; // si fiche existe
  end; // si TOB non null
end;

procedure TOF_PGCUMUL_BUL.PasDeSaisieCell(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
  exit;
end;

procedure TOF_PGCUMUL_BUL.PasDeSaisieRow(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  exit;
end;

initialization
  registerclasses([TOF_PGCUMUL_BUL]);
end.

