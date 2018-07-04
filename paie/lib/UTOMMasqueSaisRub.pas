{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 04/09/2001
Modifié le ... :   /  /
Description .. : TOM saisie des masques de la saisie par rubrique
Mots clefs ... : PAIE;SAISRUB
*****************************************************************}
{
PT1   : 09/07/2002 VG V585 Modification de la requête pour que la rubrique
                           sélectionnée soit la rubrique du dossier et non la
                           première de la liste
PT2   : 13/10/2003 JL V_42 Ajout lancement édition fiche pour CWAS
PT3   : 14/10/2003 SB V_42 FQ 10816 Si suppr. ligne de masque => suppr. valeurs dans la base
PT4   : 07/01/2004 PH V_50 Prise en compte des rémunérations de type elt permanent dans la saisie par rubrique
PT5   : 20/08/2004 PH V_50 FQ 11517 gestion des combo dans la grille suite modifs AGL
PT6   : 08/11/2004 PH V_50 FQ 11756 Prise en compte PT4 dans le cas du multi dossier
PT7   : 17/05/2004 PH V_60 FQ 12034 Niveau de parenthèse pour chargement des rémunérations
PT8   : 10/06/2005 SB V_65 CWAS : Chargement tablette compatibilite CWAS
PT9   : 10/04/2004 PH V_65 Rajout saisie des aides en ligne pour chaque colonne du masque

}
unit UTOMMasqueSaisRub;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, Spin,
  {$IFDEF EAGLCLIENT}
  UtileAGL, eFiche,
  {$ELSE}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fiche, HDB, DBCtrls, EdtREtat,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOM, UTOB, HTB97, LookUp,
  PgOutils2, Graphics, Grids, EntPaie, P5Util;

type
  TOM_MASQUESAISRUB = class(TOM)
  private
    StCellCur: string; // string contenant le contenu de la cellule en cours de traitement
    {$IFNDEF EAGLCLIENT}
    Nbre: THDbSpinEdit;
    VCP, VCC: THDBValComboBox;
    {$ELSE}
    Nbre: TSpinEdit;
    VCP, VCC: THValComboBox;
    {$ENDIF}
    TOB_Rem: TOB; // Tob des remunerations ayant un element variable defini
    Grille: THGrid; // grille de saisie de infos de la colonne
    StSql: string; // Chaine SQL de selection des remunerations
    procedure ContenuGrille(AvecAffectation: Boolean);
    procedure PostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure DessineCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GrilleClick(Sender: TObject);
    procedure ImprimerMasque(Sender: TObject); //PT2
  public
    procedure OnArgument(stArgument: string); override;
    procedure OnChangeField(F: TField); override;
    procedure OnUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnClose; override;
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;

  end;

implementation
{ TOM_MASQUESAISRUB }

// Affectation du contenu de la grille en fonction des donnees de la table
procedure TOM_MASQUESAISRUB.ContenuGrille(AvecAffectation: Boolean);
var
  i: integer;
  No: string;
begin
  if AvecAffectation = TRUE then
  begin
    for i := 1 to nbre.Value do
    begin
      No := IntToStr(i);
      Grille.Cells[0, i] := No;
      Grille.Cells[1, i] := GetField('PMR_COL' + No);
      Grille.Cells[2, i] := GetField('PMR_LIBCOL' + No);
      Grille.Cells[3, i] := RechDom('PGALIMSAISIERUB', GetField('PMR_TYPECOL' + No), FALSE);
//      Grille.CellValues[3, i] := GetField('PMR_TYPECOL' + No); // PT5
      Grille.Cells[4, i] := RechDom('PGOUINON', GetField('PMR_REPORTCOL' + No), FALSE);
//      Grille.CellValues[4, i] := GetField('PMR_REPORTCOL' + No); // PT5
    end;
  end;
  // DEBPT9
  for i := 1 to nbre.Value do
  begin
    SetControlVisible ('PMR_AIDECOL'+IntToStr (i), TRUE);
    SetControlVisible ('TPMR_AIDECOL'+IntToStr (i), TRUE);
  end;
  // FIN PT9
  for i := nbre.Value + 1 to nbre.MaxValue do
  begin
    No := IntToStr(i);
    Grille.Cells[0, i] := No;
    Grille.Cells[1, i] := '';
    Grille.Cells[2, i] := '';
    Grille.Cells[3, i] := '';
//    Grille.CellValues[3, i] := ''; // PT5
    Grille.Cells[4, i] := '';
//    Grille.CellValues[4, i] := ''; // PT5
    { DEB PT3 }
    if GetField('PMR_COL' + No) <> '' then SetField('PMR_COL' + No, '');
    if GetField('PMR_LIBCOL' + No) <> '' then SetField('PMR_LIBCOL' + No, '');
    if GetField('PMR_TYPECOL' + No) <> '' then SetField('PMR_TYPECOL' + No, '');
    if GetField('PMR_REPORTCOL' + No) <> '' then SetField('PMR_REPORTCOL' + No, '');
    SetControlVisible ('PMR_AIDECOL'+IntToStr (i), FALSE); // PT9
    SetControlVisible ('TPMR_AIDECOL'+IntToStr (i), FALSE); // PT9
    { FIN PT3 }
  end;
end;

procedure TOM_MASQUESAISRUB.GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if Grille.Col > 1 then Grille.ElipsisButton := FALSE;
  if Grille.Col = 1 then Grille.ElipsisButton := TRUE;
  if (Grille.col = 2) and (Grille.Cells[1, Grille.Row] = '') then cancel := TRUE;
  if Grille.Row > Nbre.Value then cancel := TRUE;
  if not Cancel then StCellCur := Grille.Cells[Grille.Col, Grille.Row];
  if not (DS.State in [dsInsert, dsEdit]) then DS.edit;
end;

procedure TOM_MASQUESAISRUB.GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  T: tob;
  Erreur: Integer;
  CasB, CasC, CasT, CasM: Boolean;
  St, st1, Rub: string;
begin
  CasB := FALSE;
  CasC := FALSE;
  CasT := FALSE;
  CasM := FALSE;
  Grille.ElipsisButton := FALSE;
  if Grille.Cells[ACol, ARow] = StCellCur then Exit;
  if (ACol = 1) or (ACol = 3) then
  begin
    St := Grille.CellValues[Acol, ARow];
    Rub := Grille.Cells[1, ARow];
    T := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [Rub], FALSE);
    if (DS.State in [dsInsert]) and (ACol = 1) then
    begin
      if T <> nil then Grille.Cells[2, Grille.Row] := Copy(T.GetValue('PRM_LIBELLE'), 1, 12)
      else Grille.Cells[2, Grille.Row] := '';
    end;
    if (ACol = 1) and (T = nil) and (Rub <> '') then
    begin
      PGIBOX('Attention, la rubrique ' + Rub + ' n''existe pas', Ecran.Caption);
      cancel := TRUE;
    end;
    if (ACol = 1) and (Rub = '') and (ARow <= nbre.Value) then
    begin
      PGIBOX('Attention, vous n''avez pas saisi de rubrique', Ecran.Caption);
      cancel := TRUE;
    end;

    if (ACol = 3) and (T <> nil) then
    begin
      Erreur := 0;
      if (T.GetValue('PRM_TYPEBASE') = '00') or (T.GetValue('PRM_TYPEBASE') = '01') then CasB := TRUE;
      if (T.GetValue('PRM_TYPEMONTANT') = '00') or (T.GetValue('PRM_TYPEMONTANT') = '01') then CasM := TRUE;
      if (T.GetValue('PRM_TYPETAUX') = '00') or (T.GetValue('PRM_TYPETAUX') = '01') then CasT := TRUE;
      if (T.GetValue('PRM_TYPECOEFF') = '00') or (T.GetValue('PRM_TYPECOEFF') = '01') then CasC := TRUE;
      if (CasB = FALSE) and (st = 'BAS') then Erreur := 1;
      if (CasM = FALSE) and (st = 'MON') then Erreur := 2;
      if (CasT = FALSE) and (st = 'TAU') then Erreur := 3;
      if (CasC = FALSE) and (st = 'COE') then Erreur := 4;
      if Erreur <> 0 then
      begin
        st1 := 'Pour la rémunération ' + Rub + ' ,vous pouvez renseigner ';
        if CasB = TRUE then st1 := st1 + 'Base, ';
        if CasM = TRUE then st1 := st1 + 'Montant, ';
        if CasT = TRUE then st1 := st1 + 'Taux, ';
        if CasC = TRUE then st1 := st1 + 'Coefficient, ';
        PGIBOX(St1, 'Erreur de paramètrage');
        cancel := TRUE;
      end;
    end;
  end; // fin 3eme colone
end;

procedure TOM_MASQUESAISRUB.GrilleClick(Sender: TObject);
begin
  // PT4
  // PT6 Mauvaise requête dans le cas du multi, rajout des ()
  StSql := '##PRM_PREDEFINI## ((PRM_TYPEBASE="00" OR PRM_TYPETAUX="00" OR PRM_TYPECOEFF="00" OR PRM_TYPEMONTANT="00") ';
  StSql := StSql + ' OR (PRM_TYPEBASE="01" OR PRM_TYPETAUX="01" OR PRM_TYPECOEFF="01" OR PRM_TYPEMONTANT="01")) ';
  LookUpList(Grille, 'Rémunérations', 'REMUNERATION', 'PRM_RUBRIQUE', 'PRM_LIBELLE', StSql, 'PRM_RUBRIQUE', TRUE, -1);
end;

procedure TOM_MASQUESAISRUB.OnArgument(stArgument: string);
var
  Q: TQuery;
  LBL: THLabel;
  BImprimer: TToolBarButton97;
begin
  inherited;
  Grille := THGrid(GetControl('GRILLE'));
  if Grille = nil then exit;
  {$IFNDEF EAGLCLIENT}
  VCP := THDBValComboBox(GetControl('PMR_PROFIL'));
  VCC := THDBValComboBox(GetControl('PMR_CRITEREORG'));
  Nbre := THDBSpinEdit(GetControl('PMR_NBRECOL'));
  {$ELSE}
  VCP := THValComboBox(GetControl('PMR_PROFIL'));
  VCC := THValComboBox(GetControl('PMR_CRITEREORG'));
  Nbre := TSpinEdit(GetControl('PMR_NBRECOL'));
  {$ENDIF}
  Grille.PostDrawCell := PostDrawCell;
  Grille.GetCellCanvas := DessineCell;
  Grille.OnElipsisClick := GrilleClick;
  Grille.OnCellExit := GrilleCellexit;
  Grille.OnCellEnter := GrilleCellenter;
  Grille.ColAligns[0] := taCenter;
  Grille.ColAligns[2] := taLeftJustify;
  Grille.ColAligns[1] := taCenter;
  Grille.ColAligns[3] := taCenter;
  Grille.ColAligns[4] := taCenter;
  Grille.ColFormats[3] := 'CB=PGALIMSAISIERUB';
  Grille.ColFormats[4] := 'CB=PGOUINON';
  TOB_Rem := TOB.Create('Les Rémunerations Saisissables', nil, -1);
  //PT1 StSql := 'SELECT * FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_TYPEBASE="00" OR PRM_TYPETAUX="00" OR PRM_TYPECOEFF="00" OR PRM_TYPEMONTANT="00" ORDER BY PRM_RUBRIQUE';
  StSql := 'SELECT *' +
    ' FROM REMUNERATION WHERE' +
    // PT7 rajout parenthése en début en fin de requete
    ' ##PRM_PREDEFINI## ((PRM_TYPEBASE="00" OR PRM_TYPETAUX="00" OR' +
    ' PRM_TYPECOEFF="00" OR PRM_TYPEMONTANT="00")' +
    // PT4
  ' OR (PRM_TYPEBASE="01" OR PRM_TYPETAUX="01" OR PRM_TYPECOEFF="01" OR PRM_TYPEMONTANT="01"))' +
    // PT4
  ' ORDER BY PRM_RUBRIQUE';
  //FIN PT1
  Q := OpenSql(StSql, TRUE);
  TOB_Rem.LoadDetailDB('REMUNERATION', '', '', Q, False);
  Ferme(Q);
  LBL := THLabel(GetControl('TPMR_CRITEREORG'));
  if VH_Paie.PGCritSaisRub = 'OR1' then
  begin
    LBL.Caption := VH_Paie.PGLibelleOrgStat1;
    VCC.DataType := 'PGTRAVAILN1';
  end;
  if VH_Paie.PGCritSaisRub = 'OR2' then
  begin
    LBL.Caption := VH_Paie.PGLibelleOrgStat2;
    VCC.DataType := 'PGTRAVAILN2';
  end;
  if VH_Paie.PGCritSaisRub = 'OR3' then
  begin
    LBL.Caption := VH_Paie.PGLibelleOrgStat3;
    VCC.DataType := 'PGTRAVAILN3';
  end;
  if VH_Paie.PGCritSaisRub = 'OR4' then
  begin
    LBL.Caption := VH_Paie.PGLibelleOrgStat4;
    VCC.DataType := 'PGTRAVAILN4';
  end;
  if VH_Paie.PGCritSaisRub = 'STA' then
  begin
    LBL.Caption := VH_Paie.PGLibCodeStat;
    VCC.DataType := 'PGCODESTAT';
  end;

  BImprimer := TToolBarButton97(GetControl('BIMPRIMER'));
  if BImprimer <> nil then
  begin
    BImprimer.OnClick := ImprimerMasque; //PT2
    BImprimer.Visible := True;
  end;
end;

procedure TOM_MASQUESAISRUB.OnChangeField(F: TField);
begin
  inherited;
  if VCP <> nil then if VCP.Value = '' then VCP.ItemIndex := 0;
  if VCC <> nil then if VCC.Value = '' then VCC.ItemIndex := 0;
  if F.FieldName = 'PMR_NBRECOL' then ContenuGrille(FALSE);
end;

procedure TOM_MASQUESAISRUB.OnClose;
begin
  inherited;
  TOB_Rem.Free;
  TOB_Rem := nil;
end;

procedure TOM_MASQUESAISRUB.OnLoadRecord;
var
  Ordre: TSpinEdit;
  {$IFNDEF EAGLCLIENT}
  P_Ordre: THDBValComboBox;
  {$ELSE}
  P_Ordre: THValComboBox;
  {$ENDIF}
begin
  inherited;
  if VCP <> nil then if VCP.Value = '' then VCP.ItemIndex := 0;
  if VCC <> nil then if VCC.Value = '' then VCC.ItemIndex := 0;

  SetControlEnabled('NUMORDRE', (DS.State in [dsInsert])); // Spin Edit visualisant le thdbspinedit
  Ordre := TSpinEdit(GetControl('NUMORDRE'));
  {$IFNDEF EAGLCLIENT}
  P_Ordre := THDBValComboBox(GetControl('PMR_ORDRE'));
  {$ELSE}
  P_Ordre := THValComboBox(GetControl('PMR_ORDRE'));
  {$ENDIF}
  if GetField('PRM_BOUGE') = 'X'
    then SetField('PRM_BOUGE', '-') else
    SetField('PRM_BOUGE', 'X');
  UpdateChamp('PRM_BOUGE');
  if (Ordre <> nil) and (P_Ordre <> nil) then
    if GetField('PMR_ORDRE') <> '' then
      Ordre.Value := StrToInt(GetField('PMR_ORDRE'));
  ContenuGrille(TRUE);
end;

procedure TOM_MASQUESAISRUB.OnUpdateRecord;
var
  I: integer;
  No: string;
  Ordre: TSpinEdit;
  {$IFNDEF EAGLCLIENT}
  P_Ordre: THDBValComboBox;
  {$ELSE}
  P_Ordre: THValComboBox;
  {$ENDIF}
  Erreur: Integer;
  CasB, CasC, CasT, CasM: Boolean;
  Rub, St1, st: string;
  T: TOB;
begin
  inherited;
  CasB := FALSE;
  CasC := FALSE;
  CasT := FALSE;
  CasM := FALSE;
  LastError := 0;
  for i := 1 to nbre.Value do
  begin
    No := IntToStr(i);
    if Grille.CellValues[1, i] = '' then
    begin
      LastErrorMsg := 'Vous devez renseigner une rémunération de la colonne' + No;
      LastError := 1;
      SetFocusControl('PMR_NBRECOL');
      exit;
    end;
    if Grille.CellValues[2, i] = '' then
    begin
      LastErrorMsg := 'Vous devez renseigner un libellé de la colonne' + No;
      LastError := 1;
      SetFocusControl('PMR_NBRECOL');
      exit;
    end;
    if Grille.CellValues[3, i] = '' then
    begin
      LastErrorMsg := 'Vous devez renseigner le type d''alimentation de la colonne' + No;
      LastError := 1;
      SetFocusControl('PMR_NBRECOL');
      exit;
    end;
    if Grille.CellValues[4, i] = '' then
    begin
      LastErrorMsg := 'Vous devez renseigner la zone à reporter de la colonne' + No;
      LastError := 2;
      SetFocusControl('PMR_NBRECOL');
      exit;
    end;
  end;
  for i := 1 to nbre.Value do
  begin
    No := IntToStr(i);
    St := Grille.CellValues[3, i];
    Rub := Grille.Cells[1, i];
    T := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [Rub], FALSE);
    if (T = nil) then
    begin
      PGIBOX('Attention, la rubrique ' + Rub + ' ne comporte pas d''éléments saisissable', Ecran.Caption);
      LastError := 3;
      SetFocusControl('PMR_NBRECOL');
      exit;
    end;
    if T <> nil then
    begin
      Erreur := 0;
      // PT4
      if (T.GetValue('PRM_TYPEBASE') = '00') or (T.GetValue('PRM_TYPEBASE') = '01') then CasB := TRUE;
      if (T.GetValue('PRM_TYPEMONTANT') = '00') or (T.GetValue('PRM_TYPEMONTANT') = '01') then CasM := TRUE;
      if (T.GetValue('PRM_TYPETAUX') = '00') or (T.GetValue('PRM_TYPETAUX') = '01') then CasT := TRUE;
      if (T.GetValue('PRM_TYPECOEFF') = '00') or (T.GetValue('PRM_TYPECOEFF') = '01') then CasC := TRUE;
      // FIN PT4
      if (CasB = FALSE) and (st = 'BAS') then Erreur := 1;
      if (CasM = FALSE) and (st = 'MON') then Erreur := 2;
      if (CasT = FALSE) and (st = 'TAU') then Erreur := 3;
      if (CasC = FALSE) and (st = 'COE') then Erreur := 4;
      if Erreur <> 0 then
      begin
        st1 := 'Pour la rémunération ' + Rub + ' ,vous pouvez renseigner ';
        if CasB = TRUE then st1 := st1 + 'Base, ';
        if CasM = TRUE then st1 := st1 + 'Montant, ';
        if CasT = TRUE then st1 := st1 + 'Taux, ';
        if CasC = TRUE then st1 := st1 + 'Coefficient, ';
        PGIBOX(St1, 'Erreur de paramètrage');
        LastError := 2;
        SetFocusControl('PMR_NBRECOL');
        exit;
      end;
    end;
  end;
  for i := 1 to nbre.Value do
  begin
    No := IntToStr(i);
    SetField('PMR_COL' + No, Grille.Cells[1, i]);
    SetField('PMR_LIBCOL' + No, Grille.Cells[2, i]);
    SetField('PMR_TYPECOL' + No, Grille.CellValues[3, i]);
    SetField('PMR_REPORTCOL' + No, Grille.CellValues[4, i]);
  end;
  if (DS.State in [dsInsert]) then
  begin // Uniquement en creation
    Ordre := TSpinEdit(GetControl('NUMORDRE'));
    {$IFNDEF EAGLCLIENT}
    P_Ordre := THDBValComboBox(GetControl('PMR_ORDRE'));
    {$ELSE}
    P_Ordre := THValComboBox(GetControl('PMR_ORDRE'));
    {$ENDIF}
    if (Ordre <> nil) and (P_Ordre <> nil) then
      SetField('PMR_ORDRE', IntToStr(Ordre.Value));
  end;
  //Rechargement des tablettes
  if (LastError = 0) and (Getfield('PMR_ORDRE') <> '') and (Getfield('PMR_LIBELLE') <> '') then
     {$IFNDEF EAGLCLIENT}
     ChargementTablette(TFFiche(Ecran).TableName, '');
     {$ELSE}
     ChargementTablette(TableToPrefixe(TFFiche(Ecran).TableName), ''); //PT8
     {$ENDIF}
end;


procedure TOM_MASQUESAISRUB.OnNewRecord;
var
  QQ: TQuery;
  IMax: integer;
  St: string;
  Ordre: TSpinEdit;
begin
  inherited;
  QQ := OpenSQL('SELECT MAX(PMR_ORDRE) FROM MASQUESAISRUB', TRUE);
  if not QQ.EOF then
  begin
    St := QQ.Fields[0].AsString;
    if St <> '' then
    begin
      imax := StrToInt(st);
      imax := imax + 1;
    end
    else iMax := 1;
  end
  else
  begin
    iMax := 1;
  end;
  Ferme(QQ);
  Ordre := TSpinEdit(GetControl('NUMORDRE'));
  if Ordre <> nil
    then Ordre.Value := IMax;
  SetField('PMR_ORDRE', IntToStr(IMax));
  SetField('PMR_NBRECOL', 7);
end;

procedure TOM_MASQUESAISRUB.PostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  if ARow = 0 then exit;
  if ACol = 0 then exit;
  if ARow <= Nbre.Value then GridDeGriseCell(Grille, Acol, Arow, Canvas)
  else GridGriseCell(Grille, Acol, Arow, Canvas);
end;

procedure TOM_MASQUESAISRUB.DessineCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
begin
  exit;
  if ARow <= Nbre.Value then Canvas.Brush.Style := bsClear
  else Canvas.Brush.Style := bsBDiagonal;
end;



procedure TOM_MASQUESAISRUB.OnDeleteRecord;
begin
  inherited;
  {$IFNDEF EAGLCLIENT}
  ChargementTablette(TFFiche(Ecran).TableName, '');
  {$ELSE}
  ChargementTablette(TableToPrefixe(TFFiche(Ecran).TableName), ''); //PT8
  {$ENDIF}
end;

procedure TOM_MASQUESAISRUB.ImprimerMasque(Sender: TObject); //PT2
var
  St, StPages: string;
  Pages: TPageControl;
begin
  Pages := TPageControl(GetControl('PAGES'));
  StPages := '';
  {$IFDEF EAGLCLIENT}
  StPages := AglGetCriteres(Pages, FALSE);
  {$ENDIF}
  St := 'SELECT * FROM MASQUESAISRUB WHERE PMR_ORDRE="' + GetField('PMR_ORDRE') + '"';
  LanceEtat('E', 'PAY', 'PMS', True, False, False, Pages, St, '', False, 0, StPages);
end;

initialization
  registerclasses([TOM_MASQUESAISRUB]);
end.

