{***********UNITE*************************************************
Auteur  ...... : FC
Créé le ...... : 23/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : UTOFPGHISTOCUMULS ()
                 Historique des cumuls depuis le menu Zoom de la saisie du bulletin
Mots clefs ... : TOF;UTOFPGHISTOCUMULS
*****************************************************************}
Unit UTOFPGHISTOCUMULS ;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Fe_Main, DBGrids,
{$ELSE}
  MaineAgl,
{$ENDIF}
  Grids, HCtrls, HEnt1, HMsgBox, UTOF, UTOB, Vierge, P5Util, P5Def, AGLInit;
type
  TOF_UTOFPGHISTOCUMULS = class(TOF)
  private
    CodeSal, Etab, DateD, DateF: string;
    DD, DF: TDateTime;
    Grille: THGrid;
    VCbxLRub: THValComboBox;
    procedure RechHisto(Sender: TObject);
    procedure PasDeSaisieRow(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure PasDeSaisieCell(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
  public
    procedure OnArgument(Arguments: string); override;
  end;

implementation

procedure TOF_UTOFPGHISTOCUMULS.OnArgument(Arguments: string);
var
  i: Integer;
  T1: TOB;
  Btn: TToolbarButton97;
  st, Nom: string;
  stCumul:String;
begin
  inherited;
  st := Trim(Arguments);
  CodeSal := ReadTokenSt(st); // Recup code Salarie
  Etab := ReadTokenSt(st); // Recup Code Etablissement
  DateD := ReadTokenSt(st); // Recup Date debut et date de fin de la session de paie
  DateF := ReadTokenSt(st);
  if st <> '' then Nom := ReadTokenSt(st);
  if Nom <> '' then
  begin
    Ecran.Caption := Ecran.Caption + Nom;
    UpdateCaption(Ecran);
  end;
  RendDateExerSocial(StrToDate(DateD), StrToDate(DateF), DD, DF, TRUE);
  Btn := TToolbarButton97(GetControl('RECHHISTO'));
  if Btn <> nil then Btn.OnClick := RechHisto;

  Grille := THGrid(GetControl('GRILLECUMULS'));
  if Grille <> nil then
  begin
    Grille.OnRowEnter := PasDeSaisieRow;
    Grille.OnCellEnter := PasDeSaisieCell;
    Grille.CacheEdit;
    Grille.Options := Grille.Options - [GoEditing, GoTabs, GoAlwaysShowEditor];
    Grille.Options := Grille.Options + [GoRowSelect];
  end
  else Grille.Enabled := FALSE;
  VCbxLRub := THValComboBox(GetControl('VALCBXRUB'));
  if VCbxLRub <> nil then
  begin
    VCbxLRub.clear;
    stCumul := '';

    for I := 0 to LaTOB.Detail.Count - 1 do
    begin
      T1 := LaTOB.Detail[I];
      if (stCumul <> T1.GetValue('PCL_CUMULPAIE')) then
      begin
        VCbxLRub.Items.Add(T1.GetValue('PCL_CUMULPAIE') + ' ' + T1.GetValue('PCL_LIBELLE'));
        VCbxLRub.Values.Add(T1.GetValue('PCL_CUMULPAIE'));
        stCumul := T1.GetValue('PCL_CUMULPAIE');
      end;
    end;
    VCbxLRub.ItemIndex := 0;
  end;
end;

procedure TOF_UTOFPGHISTOCUMULS.PasDeSaisieCell(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  exit;
end;

procedure TOF_UTOFPGHISTOCUMULS.PasDeSaisieRow(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  exit;
end;

procedure TOF_UTOFPGHISTOCUMULS.RechHisto(Sender: TObject);
var
  Rub: string;
  i, j, NbMt: Integer;
  T2: TOB;
  CC: Double;
  Lbl4: THLabel;
begin
  Rub := VCbxLRub.Value;
  CC := 0;
  Grille.ClearSelected;
  Grille.RowCount := 2;
  NbMt := 2;
  j := 1;
  for I := 0 to LaTOB.Detail.Count - 1 do
  begin
    T2 := LaTOB.Detail[I];
    if (T2.GetValue('PCL_CUMULPAIE') = Rub) then
    begin
      Grille.Cells[0, j] := T2.GetValue('PHC_DATEDEBUT'); // DateDebut
      Grille.Cells[1, j] := T2.GetValue('PHC_DATEFIN'); // DateFin
      Grille.Cells[2, j] := DoubleToCell(T2.GetValue('MONTANT'), NbMt); // Montant
      CC := CC + T2.GetValue('MONTANT');
      j := j + 1;
      Grille.Cells[0, j] := '';
      Grille.Cells[1, j] := '';
      Grille.Cells[2, j] := '';
      Grille.RowCount := Grille.RowCount + 1;
    end;
  end;

  Lbl4 := THLabel(GetControl('Lbl4'));
  if Lbl4 <> nil then Lbl4.Caption := DoubleToCell(CC, NbMt);
end;

initialization
  registerclasses([TOF_UTOFPGHISTOCUMULS]);
end.

