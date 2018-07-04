{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Unit de consultation des rémunérations des paies
Suite ........ : précédantes dans la saisie du bulletin
Mots clefs ... : PAIE;PGBULLETIN
*****************************************************************}
{
PT1 : 31/03/2006 PH V_65 FQ 12568 Affichage Code+libelle dans la tablette
PT2 : 31/03/2006 PH V_65 FQ 12225 Prise en compte paie validée du mois

}
unit UTofRem_Bul;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Fe_Main, DBGrids,
{$ELSE}
  MaineAgl,
{$ENDIF}
  uPaieRemunerations,
  Grids, HCtrls, HEnt1, HMsgBox, UTOF, UTOB, Vierge, P5Util, P5Def, AGLInit;
type
  TOF_PGRem_Bul = class(TOF)
  private
    CodeSal, Etab, DateD, DateF: string;
    DD, DF: TDateTime;
    Grille: THGrid;
    VCbxLRub: THValComboBox;
    procedure RechHisto(Sender: TObject);
    procedure PasDeSaisieRow(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure PasDeSaisieCell(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    //       procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
  public
    procedure OnArgument(Arguments: string); override;
  end;

implementation
{ Jamais utilisée
procedure TOF_PGRem_BUL.GetCellCanvas(Acol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
if ARow = Grille.RowCount - 2 then
 begin
 Grille.RowHeights[ARow] := 2;
 exit;
 end;
if ARow <> Grille.RowCount - 1 then exit; // On ne fait rien car on n'est pas positionné sur la dernière ligne affichant les cumuls
Grille.canvas.Font.Height:=13;
if ACol=1 then Grille.canvas.Font.Style:=Grille.canvas.Font.Style+[fsItalic]+[fsBold]
 else Grille.canvas.Font.Style:=Grille.canvas.Font.Style+[fsItalic];
Grille.Canvas.Brush.Color:=Grille.FixedColor ;
Grille.Canvas.Font.Color:=Grille.Font.Color ;
end;
}

procedure TOF_PGRem_BUL.OnArgument(Arguments: string);
var
  i: Integer;
  T1: TOB;
  Btn: TToolbarButton97;
  st, Nom: string;
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
  RendDateExerSocial(StrToDate(DateD), StrToDate(DateF), DD, DF, TRUE); // PT2
  Btn := TToolbarButton97(GetControl('RECHHISTO'));
  if Btn <> nil then Btn.OnClick := RechHisto;

  Grille := THGrid(GetControl('GRILLEREM'));
  if Grille <> nil then
  begin
    Grille.OnRowEnter := PasDeSaisieRow;
    Grille.OnCellEnter := PasDeSaisieCell;
    Grille.CacheEdit;
    Grille.Options := Grille.Options - [GoEditing, GoTabs, GoAlwaysShowEditor];
    Grille.Options := Grille.Options + [GoRowSelect];
  end
  else Grille.Enabled := FALSE;
  VCbxLRub := THValComboBox(GetControl('VALCBXRUB')); // Controle sur la combo contenant les codes des rubriques de bases
  if VCbxLRub <> nil then
  begin // il convient de récupèrer LaTob pour alimenter les combos Au maxi 5 à 6 Rubriques
    VCbxLRub.clear;
    for I := 0 to LaTOB.Detail.Count - 1 do
    begin
      T1 := LaTOB.Detail[I];
      VCbxLRub.Items.Add(T1.GetValue('CODEBASE')+' '+T1.GetValue('LIBELBASE')); // PT1
      VCbxLRub.Values.Add(T1.GetValue('CODEBASE'));
    end;
    VCbxLRub.ItemIndex := 0;
  end;
  // Design canvas de cellule ==> Grille.GetCellCanvas:=GetCellCanvas;
end;

procedure TOF_PGRem_BUL.PasDeSaisieCell(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  exit;
end;

procedure TOF_PGRem_BUL.PasDeSaisieRow(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  exit;
end;

procedure TOF_PGRem_BUL.RechHisto(Sender: TObject);
var
  Q: TQuery;
  St, Rub: string;
  i, NbDec, NbTaux, NbCoeff, NbMt: Integer;
  T2: TOB;
  CC: Double;
  Lbl4: THLabel;
begin
  Rub := VCbxLRub.Value;
  CC := 0;
  St := 'SELECT PHB_DATEDEBUT,PHB_DATEFIN,PHB_BASEREM,PHB_TAUXREM,PHB_COEFFREM,PHB_MTREM' +
    ' FROM HISTOBULLETIN WHERE PHB_SALARIE="' + CodeSal + '" AND PHB_NATURERUB="AAA" AND PHB_RUBRIQUE="' + Rub +
    '" AND PHB_DATEDEBUT>="' + UsDateTime(DD) + '" AND PHB_DATEFIN<="' + UsDateTime(DF) + '" ORDER BY PHB_DATEFIN';
  Q := OpenSQL(St, TRUE);
  T2 := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [Rub], TRUE);
  // A mettre les bonnes valeurs du nbre de décimales des taux si besoin
  NbDec := 2;
  NbTaux := 2;
  NbCoeff := 2;
  NbMt := 2;
  Grille.RowCount := 2;
  if T2 <> nil then
  begin
    NbDec := T2.GetValue('PRM_DECBASE');
    NbTaux := T2.GetValue('PRM_DECTAUX');
    NbCoeff := T2.GetValue('PRM_DECCOEFF');
    NbMt := T2.GetValue('PRM_DECMONTANT');
  end;
  i := 1;
  if Q.EOF then
  begin
    for i := 0 to 4 do Grille.Cells[i, 1] := '';
    Grille.Cells[5, 1] := 'Aucune rémunération';
    FERME (Q);
    Exit;
  end;

  while not Q.EOF do
  begin
    Grille.Cells[0, i] := DateToStr(Q.FindField('PHB_DATEDEBUT').AsFloat); // DateDebut
    Grille.Cells[1, i] := DateToStr(Q.FindField('PHB_DATEFIN').AsFloat); // DateFin
    CC := CC + Q.FindField('PHB_MTREM').AsFloat;
    Grille.Cells[2, i] := DoubleToCell(Q.FindField('PHB_BASEREM').AsFloat, NbDec); // Base
    Grille.Cells[3, i] := DoubleToCell(Q.FindField('PHB_TAUXREM').AsFloat, NbTaux); // Taux
    Grille.Cells[4, i] := DoubleToCell(Q.FindField('PHB_COEFFREM').AsFloat, NbCoeff); // Coeff
    Grille.Cells[5, i] := DoubleToCell(Q.FindField('PHB_MTREM').AsFloat, NbMt); // Montant
    Q.Next;
    i := i + 1;
    Grille.RowCount := Grille.RowCount + 1;
  end;
  Ferme(Q);
  Lbl4 := THLabel(GetControl('Lbl4'));
  if Lbl4 <> nil then Lbl4.Caption := DoubleToCell(CC, NbMt);
end;

initialization
  registerclasses([TOF_PGREM_BUL]);
end.

